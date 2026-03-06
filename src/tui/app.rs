use anyhow::Result;
use crossterm::event::{self, Event, KeyCode, KeyEventKind, KeyModifiers, MouseButton, MouseEventKind};
use ratatui::layout::{Constraint, Layout, Rect};
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, List, ListItem, ListState, Paragraph};
use ratatui::{DefaultTerminal, Frame};
use std::collections::HashMap;
use std::time::Duration;
use unicode_width::UnicodeWidthStr;

use crate::api::Api;
use crate::auth::Auth;
use crate::cache::{Cache, CachedConv};
use crate::html;
use crate::store;
use crate::types::{Conversation, ConvKind, Message};

const COLOR_HEADER: Color = Color::Rgb(137, 180, 250);     // #89b4fa blue

#[derive(Debug, Clone, PartialEq)]
enum Focus {
    Sidebar,
    Messages,
    Input,
    ThreadSearch,
    SidebarFilter,
}

/// Sidebar item: either a section header or a conversation/email entry
#[derive(Debug, Clone)]
enum SidebarItem {
    Header(String),
    Conv(usize),  // index into self.conversations
    Email(usize), // index into self.emails
}

pub struct App {
    auth: Auth,
    api: Api,
    store: store::Store,
    focus: Focus,
    exit: bool,

    // Sidebar
    conversations: Vec<Conversation>,
    sidebar_items: Vec<SidebarItem>,
    sidebar_state: ListState,
    section_starts: Vec<usize>, // sidebar_items index where each section header is
    collapsed_sections: std::collections::HashSet<String>, // collapsed section names
    saved_collapsed: Option<std::collections::HashSet<String>>, // saved state during search
    favourites: Vec<String>,    // favourite conv/email IDs
    sidebar_filter: String,      // filter sidebar by name

    // Messages
    current_conv_id: Option<String>,
    current_conv_topic: String,
    messages: Vec<Message>,
    local_plain_text: Option<String>, // pre-formatted text from local store (for fast preview)
    scroll_offset: usize,
    rendered_line_count: usize,
    view_height: usize, // actual visible height for page scroll

    // Layout areas (cached for mouse hit-testing)
    sidebar_area: Rect,
    msg_area: Rect,

    // Input
    input_buffer: String,

    // Status
    status: String,

    // Identity
    my_name: String,

    // Message content cache
    cached_snippets: HashMap<String, Vec<String>>,

    // Polling
    tick_count: u32,
    last_message_ids: HashMap<String, String>,
    has_new_messages: HashMap<String, bool>,

    // Track conversations read locally (to suppress stale unread from API)
    read_locally: HashMap<String, bool>,

    // Render cache
    cached_rendered_lines: Vec<Line<'static>>,
    render_dirty: bool,

    // New message indicator
    has_new_below: bool,

    // Scroll acceleration
    last_scroll_key: Option<KeyCode>,
    scroll_repeat_count: u32,
    last_scroll_time: std::time::Instant,

    // In-thread search
    thread_search_query: String,
    thread_search_input: String,

    // Text selection (character-level in message panel)
    select_start: Option<(usize, usize)>, // (line, col) in cached_rendered_lines
    select_end: Option<(usize, usize)>,
    selecting: bool, // mouse drag in progress

    // Emails (Microsoft Graph)
    email_folders: Vec<(String, String, Vec<serde_json::Value>)>, // (folder_name, folder_id, emails)
    emails: Vec<serde_json::Value>,
    current_email_id: Option<String>,
    current_email_body: Option<String>, // HTML body of selected email
    current_email_subject: String,
    email_loaded: bool,

    // TV search navigation
    pending_tv_nav: Option<String>,    // conv_id to navigate to after tv returns
    pending_tv_search: Option<String>, // matched text to highlight + scroll to

    // Demo mode (no API calls)
    demo: bool,
}

impl App {
    pub async fn new() -> Result<Self> {
        let auth = Auth::new()?;
        let api = Api::new(&auth.region());
        let store = store::Store::new(auth.config_dir())?;

        let mut app = Self {
            auth,
            api,
            store,
            focus: Focus::Sidebar,
            exit: false,
            conversations: Vec::new(),
            sidebar_items: Vec::new(),
            sidebar_state: ListState::default(),
            section_starts: Vec::new(),
            collapsed_sections: ["Direct Messages", "Channels", "Meetings", "Emails"]
                .iter().map(|s| s.to_string()).collect(),
            saved_collapsed: None,
            favourites: Vec::new(),
            sidebar_filter: String::new(),
            current_conv_id: None,
            current_conv_topic: String::new(),
            messages: Vec::new(),
            local_plain_text: None,
            scroll_offset: 0,
            rendered_line_count: 0,
            view_height: 20,
            sidebar_area: Rect::default(),
            msg_area: Rect::default(),
            input_buffer: String::new(),
            status: "Loading...".to_string(),
            my_name: String::new(),
            cached_snippets: HashMap::new(),
            tick_count: 0,
            last_message_ids: HashMap::new(),
            has_new_messages: HashMap::new(),
            read_locally: HashMap::new(),
            thread_search_query: String::new(),
            thread_search_input: String::new(),
            select_start: None,
            select_end: None,
            selecting: false,
            email_folders: Vec::new(),
            emails: Vec::new(),
            current_email_id: None,
            current_email_body: None,
            current_email_subject: String::new(),
            email_loaded: false,
            cached_rendered_lines: Vec::new(),
            render_dirty: true,
            has_new_below: false,
            last_scroll_key: None,
            scroll_repeat_count: 0,
            last_scroll_time: std::time::Instant::now(),
            pending_tv_nav: None,
            pending_tv_search: None,
            demo: false,
        };

        // Try loading from cache first for instant startup, then always refresh from API
        app.load_from_cache();
        if !app.conversations.is_empty() {
            // Mark cached as read (stale unread state)
            for conv in &app.conversations {
                app.read_locally.insert(conv.id.clone(), true);
            }
            app.status = format!("{} conversations (cached, refreshing...)", app.conversations.len());
        }
        // Always load fresh from API (updates cache if it existed, or first load if no cache)
        app.load_conversations().await;
        for conv in &app.conversations {
            app.read_locally.entry(conv.id.clone()).or_insert(true);
        }

        // Load favourites from store
        if let Ok(favs) = app.store.load_favourites() {
            app.favourites = favs;
        }
        app.rebuild_sidebar();

        Ok(app)
    }

    pub fn new_demo() -> Result<Self> {
        let auth = Auth::new()?;
        let api = Api::new(&auth.region());
        let store = store::Store::new(auth.config_dir())?;

        // Create dummy data
        store.create_dummy_data()?;

        // Build dummy Conversation structs from store index
        let dummy_channels = [
            ("test_general", "general", "channel"),
            ("test_engineering", "engineering", "channel"),
            ("test_random", "random", "channel"),
        ];
        let conversations: Vec<Conversation> = dummy_channels
            .iter()
            .map(|(id, name, _kind)| {
                use crate::types::*;
                Conversation {
                    id: id.to_string(),
                    conv_type: None,
                    thread_properties: Some(ThreadProperties {
                        topic: Some(name.to_string()),
                        last_join_at: None,
                        member_count: None,
                        product_thread_type: None,
                    }),
                    last_message: None,
                    version: Some(1),
                    properties: None,
                    member_names: vec![
                        "Alice".to_string(),
                        "Bob".to_string(),
                        "Charlie".to_string(),
                    ],
                }
            })
            .collect();

        let mut app = Self {
            auth,
            api,
            store,
            focus: Focus::Sidebar,
            exit: false,
            conversations,
            sidebar_items: Vec::new(),
            sidebar_state: ListState::default(),
            section_starts: Vec::new(),
            collapsed_sections: ["Direct Messages", "Channels", "Meetings", "Emails"]
                .iter().map(|s| s.to_string()).collect(),
            saved_collapsed: None,
            favourites: Vec::new(),
            sidebar_filter: String::new(),
            current_conv_id: None,
            current_conv_topic: String::new(),
            messages: Vec::new(),
            local_plain_text: None,
            scroll_offset: 0,
            rendered_line_count: 0,
            view_height: 20,
            sidebar_area: Rect::default(),
            msg_area: Rect::default(),
            input_buffer: String::new(),
            status: "Demo mode".to_string(),
            my_name: "You".to_string(),
            cached_snippets: HashMap::new(),
            tick_count: 0,
            last_message_ids: HashMap::new(),
            has_new_messages: HashMap::new(),
            read_locally: HashMap::new(),
            thread_search_query: String::new(),
            thread_search_input: String::new(),
            select_start: None,
            select_end: None,
            selecting: false,
            email_folders: Vec::new(),
            emails: Vec::new(),
            current_email_id: None,
            current_email_body: None,
            current_email_subject: String::new(),
            email_loaded: false,
            cached_rendered_lines: Vec::new(),
            render_dirty: true,
            has_new_below: false,
            last_scroll_key: None,
            scroll_repeat_count: 0,
            last_scroll_time: std::time::Instant::now(),
            pending_tv_nav: None,
            pending_tv_search: None,
            demo: true,
        };

        app.rebuild_sidebar();
        // Select first non-header item
        if app.sidebar_items.len() > 1 {
            app.sidebar_state.select(Some(1));
        }

        Ok(app)
    }

    fn is_token_error(msg: &str) -> bool {
        msg.contains("expired") || msg.contains("invalid_grant") || msg.contains("refresh token")
    }

    fn suspend_tui() {
        use std::io::Write;
        let _ = crossterm::execute!(
            std::io::stdout(),
            crossterm::event::DisableMouseCapture,
            crossterm::terminal::LeaveAlternateScreen,
        );
        let _ = crossterm::terminal::disable_raw_mode();
        let _ = std::io::stdout().flush();
    }

    fn resume_tui(terminal: &mut DefaultTerminal) {
        use std::io::Write;
        let _ = crossterm::terminal::enable_raw_mode();
        let _ = crossterm::execute!(
            std::io::stdout(),
            crossterm::terminal::EnterAlternateScreen,
            crossterm::event::EnableMouseCapture,
        );
        let _ = std::io::stdout().flush();
        let _ = terminal.clear();
    }

    /// Extract section name from header title like "▶ Direct Messages (42)"
    fn extract_section_name(title: &str) -> &str {
        let name = title.trim_start_matches(|c: char| !c.is_alphabetic());
        name.split('(').next().unwrap_or("").trim()
    }

    pub async fn run(&mut self, terminal: &mut DefaultTerminal) -> Result<()> {
        let mut needs_draw = true;
        let mut last_tick = std::time::Instant::now();

        while !self.exit {
            if needs_draw {
                terminal.draw(|frame| self.draw(frame))?;
                needs_draw = false;
            }

            // Short poll for snappy input, drain all pending events
            if event::poll(Duration::from_millis(10))? {
                loop {
                    match event::read()? {
                        Event::Key(key) => {
                            if key.kind == KeyEventKind::Press {
                                self.handle_key(key.code, key.modifiers, terminal).await;
                                needs_draw = true;
                            }
                        }
                        Event::Mouse(mouse) => {
                            self.handle_mouse(mouse).await;
                            needs_draw = true;
                        }
                        Event::Resize(_, _) => {
                            self.render_dirty = true;
                            needs_draw = true;
                        }
                        _ => {}
                    }
                    // Drain remaining events without blocking
                    if !event::poll(Duration::from_millis(0))? {
                        break;
                    }
                }
            }

            // Tick-based background tasks (~every 100ms) - skip in demo mode
            if !self.demo && last_tick.elapsed() >= Duration::from_millis(100) {
                last_tick = std::time::Instant::now();
                self.tick_count += 1;
                // Load emails on second tick (deferred from startup)
                if self.tick_count == 2 && self.email_folders.is_empty() {
                    self.load_email_folders().await;
                    needs_draw = true;
                }
                // Refresh conversations on tick 5 (~0.5s) to catch messages since cache
                if self.tick_count == 5 && !self.conversations.is_empty() {
                    self.poll_new_messages().await;
                    needs_draw = true;
                }
                // Poll every ~15 seconds (150 ticks)
                if self.tick_count >= 150 {
                    self.tick_count = 0;
                    self.poll_new_messages().await;
                    needs_draw = true;
                }
            }
        }
        Ok(())
    }

    fn scroll_amount(&mut self, key: KeyCode) -> usize {
        let now = std::time::Instant::now();
        if self.last_scroll_key == Some(key) && now.duration_since(self.last_scroll_time).as_millis() < 200 {
            self.scroll_repeat_count += 1;
        } else {
            self.scroll_repeat_count = 0;
        }
        self.last_scroll_key = Some(key);
        self.last_scroll_time = now;
        if self.scroll_repeat_count >= 3 { 3 } else { 1 }
    }

    // --- Drawing ---

    fn draw(&mut self, frame: &mut Frame) {
        let area = frame.area();

        // Split: sidebar (25%) | main (75%)
        let [sidebar_area, main_area] = Layout::horizontal([
            Constraint::Percentage(25),
            Constraint::Percentage(75),
        ]).areas(area);

        self.sidebar_area = sidebar_area;
        self.draw_sidebar(frame, sidebar_area);
        self.draw_main(frame, main_area);
    }

    fn draw_sidebar(&mut self, frame: &mut Frame, area: Rect) {
        let [list_area, status_area] = Layout::vertical([
            Constraint::Min(1),
            Constraint::Length(1),
        ]).areas(area);

        let mut items: Vec<ListItem> = Vec::new();

        for sidebar_item in &self.sidebar_items {
            match sidebar_item {
                SidebarItem::Header(title) => {
                    items.push(ListItem::new(Line::from(
                        Span::styled(
                            format!(" {}", title),
                            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD),
                        ),
                    )));
                }
                SidebarItem::Conv(conv_idx) => {
                    let conv = &self.conversations[*conv_idx];
                    let name = conv.display_name(&self.my_name);
                    let kind = conv.kind();
                    let prefix = match kind {
                        ConvKind::Channel => "#",
                        ConvKind::Chat => "@",
                        ConvKind::Meeting => "M",
                        _ => " ",
                    };
                    let is_current = self.current_conv_id.as_deref() == Some(&conv.id);
                    let read_local = self.read_locally.get(&conv.id).copied().unwrap_or(false);
                    let unread = if read_local {
                        self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                    } else {
                        conv.is_unread()
                            || self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                    };

                    let marker = if unread { "\u{25cf} " } else { "  " }; // filled circle for unread
                    let line = if is_current {
                        Line::from(vec![
                            Span::styled(
                                format!("{}{} ", marker, prefix),
                                Style::default().fg(Color::Cyan),
                            ),
                            Span::styled(
                                name,
                                Style::default().fg(Color::Cyan).add_modifier(Modifier::BOLD),
                            ),
                        ])
                    } else if unread {
                        Line::from(vec![
                            Span::styled(
                                format!("{}{} ", marker, prefix),
                                Style::default().fg(Color::Rgb(166, 227, 161)).add_modifier(Modifier::BOLD),
                            ),
                            Span::styled(
                                name,
                                Style::default().fg(Color::White).add_modifier(Modifier::BOLD),
                            ),
                        ])
                    } else {
                        Line::from(vec![
                            Span::styled(
                                format!("{}{} ", marker, prefix),
                                Style::default().fg(Color::DarkGray),
                            ),
                            Span::raw(name),
                        ])
                    };
                    items.push(ListItem::new(line));
                }
                SidebarItem::Email(email_idx) => {
                    if let Some(email) = self.emails.get(*email_idx) {
                        let subject = email.get("subject")
                            .and_then(|v| v.as_str())
                            .unwrap_or("(no subject)");
                        let from = email.get("from")
                            .and_then(|v| v.get("emailAddress"))
                            .and_then(|v| v.get("name"))
                            .and_then(|v| v.as_str())
                            .unwrap_or("?");
                        let is_read = email.get("isRead")
                            .and_then(|v| v.as_bool())
                            .unwrap_or(true);
                        let email_id = email.get("id")
                            .and_then(|v| v.as_str())
                            .unwrap_or("");
                        let is_current = self.current_email_id.as_deref() == Some(email_id);

                        // Truncate subject for sidebar
                        let label = if subject.len() > 30 {
                            format!("{}..", &subject[..28])
                        } else {
                            subject.to_string()
                        };

                        let email_marker = if !is_read { "\u{25cf} " } else { "  " };
                        let line = if is_current {
                            Line::from(vec![
                                Span::styled(email_marker, Style::default().fg(Color::Cyan)),
                                Span::styled(
                                    format!("{} ", from),
                                    Style::default().fg(Color::Cyan).add_modifier(Modifier::BOLD),
                                ),
                                Span::styled(
                                    label,
                                    Style::default().fg(Color::Cyan),
                                ),
                            ])
                        } else if !is_read {
                            Line::from(vec![
                                Span::styled(email_marker, Style::default().fg(Color::Rgb(166, 227, 161)).add_modifier(Modifier::BOLD)),
                                Span::styled(
                                    format!("{} ", from),
                                    Style::default().fg(Color::White).add_modifier(Modifier::BOLD),
                                ),
                                Span::styled(
                                    label,
                                    Style::default().fg(Color::White).add_modifier(Modifier::BOLD),
                                ),
                            ])
                        } else {
                            Line::from(vec![
                                Span::styled(email_marker, Style::default().fg(Color::DarkGray)),
                                Span::styled(
                                    format!("{} ", from),
                                    Style::default().fg(Color::DarkGray),
                                ),
                                Span::raw(label),
                            ])
                        };
                        items.push(ListItem::new(line));
                    }
                }
            }
        }

        let border_color = if self.focus == Focus::Sidebar || self.focus == Focus::SidebarFilter {
            Color::Cyan
        } else {
            Color::DarkGray
        };

        let list = List::new(items)
            .block(
                Block::default()
                    .title(" teamsh ")
                    .borders(Borders::ALL)
                    .border_style(Style::default().fg(border_color)),
            )
            .highlight_style(
                Style::default()
                    .bg(Color::DarkGray)
                    .add_modifier(Modifier::BOLD),
            );

        frame.render_stateful_widget(list, list_area, &mut self.sidebar_state);

        let status_text = if self.focus == Focus::Sidebar || self.focus == Focus::SidebarFilter {
            self.status.clone()
        } else {
            String::new()
        };
        let status = Paragraph::new(
            Line::from(Span::styled(status_text, Style::default().fg(Color::DarkGray))),
        );
        frame.render_widget(status, status_area);
    }

    fn draw_main(&mut self, frame: &mut Frame, area: Rect) {
        if self.current_conv_id.is_none() && self.current_email_id.is_none() {
            // No conversation or email open - show welcome
            let block = Block::default()
                .borders(Borders::ALL)
                .border_style(Style::default().fg(Color::DarkGray));
            let welcome = Paragraph::new(vec![
                Line::from(""),
                Line::from("  Select a conversation or email from the sidebar"),
                Line::from(""),
                Line::from(
                    Span::styled("  j/k to navigate, Enter to open", Style::default().fg(Color::DarkGray))
                ),
            ])
            .block(block);
            frame.render_widget(welcome, area);
            return;
        }

        let is_email = self.current_email_id.is_some();
        let input_lines = if is_email {
            0
        } else if self.input_buffer.is_empty() {
            5 // minimum: 3 lines + borders
        } else {
            // Calculate wrapped line count using msg_area width as reference
            let wrap_width = self.msg_area.width.saturating_sub(2).max(20) as usize;
            let mut total_lines: usize = 0;
            for line in self.input_buffer.split('\n') {
                let w: usize = unicode_width::UnicodeWidthStr::width(line);
                total_lines += (w / wrap_width).max(0) + 1;
            }
            (total_lines as u16 + 2).min(10) // +2 for borders, max 10
        };
        let [header_area, msg_area, input_area, help_area] = Layout::vertical([
            Constraint::Length(1),
            Constraint::Min(1),
            Constraint::Length(input_lines),
            Constraint::Length(1),
        ])
        .areas(area);

        // Header
        let mut header_spans = vec![
            Span::styled(" ", Style::default()),
            Span::styled(
                &self.current_conv_topic,
                Style::default()
                    .fg(COLOR_HEADER)
                    .add_modifier(Modifier::BOLD),
            ),
        ];
        let header = Paragraph::new(Line::from(header_spans));
        frame.render_widget(header, header_area);

        // Messages - manual wrapping for correct scroll
        let msg_border_color = if self.focus == Focus::Messages {
            COLOR_HEADER
        } else {
            Color::DarkGray
        };
        let msg_block = Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(msg_border_color));

        let inner_width = msg_area.width.saturating_sub(2) as usize; // borders
        let view_height = msg_area.height.saturating_sub(2) as usize;
        self.view_height = view_height;
        self.msg_area = msg_area;

        let wrapped_lines = if self.render_dirty {
            let lines = if self.current_email_id.is_some() {
                self.render_email(inner_width)
            } else {
                self.render_messages(inner_width)
            };
            self.cached_rendered_lines = lines.clone();
            self.render_dirty = false;
            lines
        } else {
            self.cached_rendered_lines.clone()
        };
        self.rendered_line_count = wrapped_lines.len();

        // Cap scroll
        let max_scroll = self.rendered_line_count.saturating_sub(view_height);
        if self.scroll_offset > max_scroll {
            self.scroll_offset = max_scroll;
        }

        // Slice visible lines, highlighting selection
        let sel = self.selection_normalized();
        let sel_bg = Color::Rgb(49, 50, 68); // Catppuccin surface0
        let visible: Vec<Line> = wrapped_lines
            .into_iter()
            .enumerate()
            .skip(self.scroll_offset)
            .take(view_height)
            .map(|(i, line)| {
                let Some((start, end)) = sel else { return line; };
                if i < start.0 || i > end.0 { return line; }

                // Determine selected char range on this line
                let line_len: usize = line.spans.iter().map(|s| s.content.len()).sum();
                let sel_from = if i == start.0 { start.1.min(line_len) } else { 0 };
                let sel_to = if i == end.0 { end.1.min(line_len) } else { line_len };
                if sel_from >= sel_to && !(start.0 != end.0 && i > start.0 && i < end.0) {
                    return line;
                }

                // Walk spans and split at selection boundaries
                let mut new_spans: Vec<Span<'static>> = Vec::new();
                let mut pos: usize = 0;
                for span in line.spans {
                    let slen = span.content.len();
                    let s_start = pos;
                    let s_end = pos + slen;
                    pos = s_end;

                    if s_end <= sel_from || s_start >= sel_to {
                        // Entirely outside selection
                        new_spans.push(span);
                    } else {
                        // Partially or fully inside selection
                        let clip_from = sel_from.saturating_sub(s_start);
                        let clip_to = (sel_to - s_start).min(slen);
                        if clip_from > 0 {
                            new_spans.push(Span::styled(
                                span.content[..clip_from].to_string(), span.style));
                        }
                        new_spans.push(Span::styled(
                            span.content[clip_from..clip_to].to_string(),
                            span.style.bg(sel_bg)));
                        if clip_to < slen {
                            new_spans.push(Span::styled(
                                span.content[clip_to..].to_string(), span.style));
                        }
                    }
                }
                Line::from(new_spans)
            })
            .collect();

        let messages = Paragraph::new(visible).block(msg_block);
        frame.render_widget(messages, msg_area);

        // Scroll indicator
        if self.rendered_line_count > view_height {
            let pct = if max_scroll > 0 {
                (self.scroll_offset * 100) / max_scroll
            } else {
                100
            };
            let indicator = format!(" {}% ", pct);
            let ind_area = Rect::new(
                msg_area.x + msg_area.width - indicator.len() as u16 - 1,
                msg_area.y,
                indicator.len() as u16,
                1,
            );
            frame.render_widget(
                Paragraph::new(Span::styled(
                    indicator,
                    Style::default().fg(Color::DarkGray),
                )),
                ind_area,
            );
        }

        // New messages indicator
        if self.has_new_below {
            let label = " New messages (G to jump) ";
            let label_area = Rect::new(
                msg_area.x + (msg_area.width / 2).saturating_sub(label.len() as u16 / 2),
                msg_area.y + msg_area.height - 1,
                label.len() as u16,
                1,
            );
            frame.render_widget(
                Paragraph::new(Span::styled(label, Style::default().fg(Color::Black).bg(Color::Yellow))),
                label_area,
            );
        }

        // Input
        let input_border = if self.focus == Focus::Input {
            COLOR_HEADER
        } else {
            Color::DarkGray
        };
        let input = Paragraph::new(self.input_buffer.as_str())
            .wrap(ratatui::widgets::Wrap { trim: false })
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .title(" Message (\\:newline) ")
                    .border_style(Style::default().fg(input_border)),
            );
        frame.render_widget(input, input_area);

        if self.focus == Focus::Input {
            // Calculate cursor row/col accounting for word wrap
            let inner_w = (input_area.width.saturating_sub(2)).max(1) as usize;
            let mut row: u16 = 0;
            let mut col: u16 = 0;
            for line in self.input_buffer.split('\n') {
                let w = unicode_width::UnicodeWidthStr::width(line);
                let wrapped_rows = (w / inner_w) as u16;
                row += wrapped_rows;
                col = (w % inner_w) as u16;
            }
            // Adjust for the lines before the last logical line
            let newline_count = self.input_buffer.chars().filter(|&c| c == '\n').count() as u16;
            // row already includes wrap rows for all lines; add newline_count for the line breaks themselves
            // Actually: split('\n') gives us each line. For each line we counted wrapped_rows.
            // But we iterated ALL lines, so row has the total wrap rows. We need to add (num_lines - 1) for the \n breaks.
            row += newline_count;
            // But the last iteration set col correctly for the last line.
            // row is cumulative of all wrap rows + newlines. Subtract the last line's wrap rows since
            // we only want rows ABOVE the cursor line.
            // Simpler: just recompute
            let mut cursor_row: u16 = 0;
            for line in self.input_buffer.split('\n') {
                let w = unicode_width::UnicodeWidthStr::width(line);
                cursor_row += (w / inner_w) as u16 + 1; // +1 for the line itself
            }
            cursor_row -= 1; // cursor is on the last row, not after it
            let last_line = self.input_buffer.rsplit('\n').next().unwrap_or("");
            let last_w = unicode_width::UnicodeWidthStr::width(last_line);
            let cursor_col = (last_w % inner_w) as u16;

            let cx = input_area.x + 1 + cursor_col;
            let cy = (input_area.y + 1 + cursor_row).min(input_area.y + input_area.height - 2);
            frame.set_cursor_position((cx, cy));
        }

        // Help / search bar
        if self.focus == Focus::SidebarFilter {
            let search_line = Line::from(vec![
                Span::styled(" s/", Style::default().fg(Color::Rgb(249, 226, 175))),
                Span::styled(self.sidebar_filter.clone(), Style::default().fg(Color::White)),
                Span::styled("  Enter:next  Esc:done  \u{2191}/\u{2193}:nav", Style::default().fg(Color::DarkGray)),
            ]);
            frame.render_widget(Paragraph::new(search_line), help_area);
            let cursor_x = help_area.x + 3 + self.sidebar_filter.len() as u16;
            frame.set_cursor_position((cursor_x, help_area.y));
        } else if self.focus == Focus::ThreadSearch {
            let search_line = Line::from(vec![
                Span::styled(" /", Style::default().fg(Color::Rgb(249, 226, 175))),
                Span::styled(self.thread_search_input.clone(), Style::default().fg(Color::White)),
                Span::styled("  Enter:search  Esc:clear", Style::default().fg(Color::DarkGray)),
            ]);
            frame.render_widget(Paragraph::new(search_line), help_area);
            let cursor_x = help_area.x + 2 + self.thread_search_input.width() as u16;
            let cursor_x = cursor_x.min(help_area.x + help_area.width - 1);
            frame.set_cursor_position((cursor_x, help_area.y));
        } else {
            let help = match self.focus {
                Focus::Input => " Enter:send  \\+Enter:newline  Esc:cancel ",
                Focus::Messages => {
                    if !self.thread_search_query.is_empty() {
                        " j/k:scroll  g/G:top/end  s:search  n/N:next/prev  /:tv  i:compose  Esc:back "
                    } else {
                        " j/k:scroll  g/G:top/end  s:search  /:tv  i:compose  r:refresh  Esc:back "
                    }
                }
                _ => {
                    if !self.sidebar_filter.is_empty() {
                        " j/k:nav  n/N:next/prev  s:new search  Tab:section  Enter:open  Space:fav  q:quit "
                    } else {
                        " j/k:nav  Tab:section  Enter:open  Space:fav  s:search  /:tv  r:refresh  L:login  q:quit "
                    }
                }
            };
            let help_widget = Paragraph::new(
                Line::from(help).style(Style::default().fg(Color::DarkGray)),
            );
            frame.render_widget(help_widget, help_area);
        }
    }

    /// Render messages into pre-wrapped lines for correct scroll calculation.
    fn render_messages(&mut self, width: usize) -> Vec<Line<'static>> {
        if width == 0 {
            return Vec::new();
        }

        // If we have local plain text (from store files), use it directly
        if self.messages.is_empty() {
            if let Some(ref text) = self.local_plain_text {
                return style_lines(text, width, &self.thread_search_query);
            }
        }

        let my_name = self.my_name.clone();

        // Build plain text in the same format as stored .txt files
        let mut plain = String::new();
        for m in &self.messages {
            let msgtype = m.messagetype.as_deref().unwrap_or("");
            if msgtype != "RichText/Html" && msgtype != "Text" {
                continue;
            }

            let sender = m
                .imdisplayname
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or(if my_name.is_empty() { "You" } else { &my_name });
            let content = m.content.as_deref().unwrap_or("");
            let timestamp_raw = m.timestamp.as_deref().unwrap_or("");
            let text = html::strip_html(content);

            // Use the same format as stored files: "Name  Mar 05 14:32"
            // Channel name is shown in the panel header, not repeated per message
            plain.push_str(&store::format_message(sender, timestamp_raw, "", &text));

            // Reactions
            let reactions = m
                .properties
                .as_ref()
                .and_then(|p| p.get("emotions"))
                .and_then(|e| {
                    if let Some(s) = e.as_str() {
                        serde_json::from_str::<Vec<serde_json::Value>>(s).ok()
                    } else {
                        e.as_array().cloned()
                    }
                });
            if let Some(reactions) = reactions {
                let reaction_str: String = reactions
                    .iter()
                    .filter_map(|r| {
                        let key = r.get("key")?.as_str()?;
                        let count = r.get("users")?.as_array()?.len();
                        if count == 0 {
                            return None;
                        }
                        let emoji = html::teams_emoji(key);
                        Some(format!("{}{}", emoji, count))
                    })
                    .collect::<Vec<_>>()
                    .join(" ");
                if !reaction_str.is_empty() {
                    plain.push_str(&format!("  {}\n", reaction_str));
                }
            }

            plain.push('\n');
        }

        style_lines(&plain, width, &self.thread_search_query)
    }

    /// Render an email body into wrapped lines for the main panel.
    fn render_email(&self, width: usize) -> Vec<Line<'static>> {
        if width == 0 {
            return Vec::new();
        }

        let mut lines: Vec<Line<'static>> = Vec::new();

        // Email header info
        if let Some(email_idx) = self.emails.iter().position(|e| {
            e.get("id").and_then(|v| v.as_str()) == self.current_email_id.as_deref()
        }) {
            let email = &self.emails[email_idx];
            let from = email.get("from")
                .and_then(|v| v.get("emailAddress"))
                .map(|ea| {
                    let name = ea.get("name").and_then(|v| v.as_str()).unwrap_or("?");
                    let addr = ea.get("address").and_then(|v| v.as_str()).unwrap_or("");
                    format!("{} <{}>", name, addr)
                })
                .unwrap_or_else(|| "?".to_string());
            let date = email.get("receivedDateTime")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .get(..16)
                .unwrap_or("?");

            // Build email as plain text for bat
            let mut plain = format!("From: {}\nDate: {}\n\n", from, date);
            let body_html = self.current_email_body.as_deref().unwrap_or("Loading...");
            let body_text = html::strip_html(body_html);
            plain.push_str(&body_text);

            return style_lines(&plain, width, &self.thread_search_query);
        }

        lines
    }

    // --- Key handling ---

    async fn handle_key(&mut self, key: KeyCode, modifiers: KeyModifiers, terminal: &mut DefaultTerminal) {
        match &self.focus {
            Focus::Sidebar => match key {
                KeyCode::Char('q') => self.exit = true,
                KeyCode::Char('j') | KeyCode::Down => {
                    self.sidebar_next();
                    self.preview_selected().await;
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    self.sidebar_prev();
                    self.preview_selected().await;
                }
                KeyCode::Char('r') => {
                    self.load_conversations().await;
                    self.load_email_folders().await;
                }
                KeyCode::Char('L') => {
                    self.spawn_login(terminal).await;
                }
                KeyCode::Char('/') => {
                    self.spawn_tv(terminal);
                    if let Some(conv_id) = self.pending_tv_nav.take() {
                        let search = self.pending_tv_search.take();
                        self.navigate_to_conv(&conv_id, search).await;
                    }
                }
                KeyCode::Right => {
                    if self.current_conv_id.is_some() || self.current_email_id.is_some() {
                        self.focus = Focus::Messages;
                    }
                }
                KeyCode::Tab => {
                    if !self.section_starts.is_empty() {
                        let selected = self.sidebar_state.selected().unwrap_or(0);
                        let mut current_sec = 0;
                        for (i, &start) in self.section_starts.iter().enumerate() {
                            if selected >= start {
                                current_sec = i;
                            }
                        }
                        let next_sec = (current_sec + 1) % self.section_starts.len();
                        let start = self.section_starts[next_sec];
                        self.sidebar_state.select(Some(start));
                    }
                }
                KeyCode::BackTab => {
                    if !self.section_starts.is_empty() {
                        let selected = self.sidebar_state.selected().unwrap_or(0);
                        let mut current_sec = 0;
                        for (i, &start) in self.section_starts.iter().enumerate() {
                            if selected >= start {
                                current_sec = i;
                            }
                        }
                        let prev_sec = if current_sec == 0 {
                            self.section_starts.len() - 1
                        } else {
                            current_sec - 1
                        };
                        let start = self.section_starts[prev_sec];
                        self.sidebar_state.select(Some(start));
                    }
                }
                KeyCode::Char(' ') => {
                    self.toggle_favourite();
                }
                KeyCode::Char('i') => {
                    if self.current_conv_id.is_some() {
                        self.focus = Focus::Input;
                    }
                }
                KeyCode::Enter => {
                    if let Some(selected) = self.sidebar_state.selected() {
                        if let Some(SidebarItem::Header(title)) = self.sidebar_items.get(selected) {
                            let title = title.clone();
                            self.toggle_section(&title);
                        } else {
                            self.open_conversation().await;
                            self.focus = Focus::Messages;
                        }
                    }
                }
                KeyCode::Char('s') | KeyCode::Char('f') => {
                    // Restore previous collapsed state before starting new search
                    if let Some(saved) = self.saved_collapsed.take() {
                        self.collapsed_sections = saved;
                    }
                    self.sidebar_filter.clear();
                    // Save collapsed state and expand all sections for search
                    self.saved_collapsed = Some(self.collapsed_sections.clone());
                    self.collapsed_sections.clear();
                    self.rebuild_sidebar();
                    self.focus = Focus::SidebarFilter;
                }
                KeyCode::Char('n') => {
                    self.jump_sidebar_to_filter_next();
                    self.preview_selected().await;
                }
                KeyCode::Char('N') => {
                    self.jump_sidebar_to_filter_prev();
                    self.preview_selected().await;
                }
                _ => {}
            },
            Focus::SidebarFilter => match key {
                KeyCode::Enter => {
                    // Accept search, go to n/N mode, preview selected
                    self.focus = Focus::Sidebar;
                    self.preview_selected().await;
                }
                KeyCode::Down => {
                    self.jump_sidebar_to_filter_next();
                    self.preview_selected().await;
                }
                KeyCode::Up => {
                    self.jump_sidebar_to_filter_prev();
                    self.preview_selected().await;
                }
                KeyCode::Esc => {
                    // Cancel search, restore collapsed state, clear filter
                    if let Some(saved) = self.saved_collapsed.take() {
                        self.collapsed_sections = saved;
                        self.rebuild_sidebar();
                    }
                    self.sidebar_filter.clear();
                    self.focus = Focus::Sidebar;
                }
                KeyCode::Backspace => {
                    self.sidebar_filter.pop();
                    self.jump_sidebar_to_filter();
                }
                KeyCode::Char(c) => {
                    self.sidebar_filter.push(c);
                    self.jump_sidebar_to_filter();
                }
                _ => {}
            },
            Focus::Messages => match key {
                KeyCode::Esc => {
                    if self.select_start.is_some() {
                        self.clear_selection();
                    } else {
                        self.focus = Focus::Sidebar;
                    }
                }
                KeyCode::Left | KeyCode::Char('h') => {
                    self.focus = Focus::Sidebar;
                    self.clear_selection();
                }
                KeyCode::Tab => {
                    self.focus = Focus::Sidebar;
                }
                KeyCode::Char('i') => {
                    self.focus = Focus::Input;
                }
                KeyCode::Char('j') | KeyCode::Down => {
                    let amount = self.scroll_amount(key);
                    self.scroll_offset = self.scroll_offset.saturating_add(amount);
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    let amount = self.scroll_amount(key);
                    self.scroll_offset = self.scroll_offset.saturating_sub(amount);
                }
                KeyCode::Char('d') if modifiers.contains(KeyModifiers::CONTROL) => {
                    let half = self.view_height / 2;
                    self.scroll_offset = self.scroll_offset.saturating_add(half.max(1));
                }
                KeyCode::Char('u') if modifiers.contains(KeyModifiers::CONTROL) => {
                    let half = self.view_height / 2;
                    self.scroll_offset = self.scroll_offset.saturating_sub(half.max(1));
                }
                KeyCode::PageDown => {
                    self.scroll_offset = self.scroll_offset.saturating_add(self.view_height.max(1));
                }
                KeyCode::PageUp => {
                    self.scroll_offset = self.scroll_offset.saturating_sub(self.view_height.max(1));
                }
                KeyCode::Char('G') => {
                    self.scroll_offset = self.rendered_line_count;
                    self.has_new_below = false;
                }
                KeyCode::Char('g') => {
                    self.scroll_offset = 0;
                }
                KeyCode::Char('r') => self.load_messages().await,
                KeyCode::Char('/') => {
                    self.spawn_tv(terminal);
                    if let Some(conv_id) = self.pending_tv_nav.take() {
                        let search = self.pending_tv_search.take();
                        self.navigate_to_conv(&conv_id, search).await;
                    }
                }
                KeyCode::Char('s') | KeyCode::Char('f') => {
                    self.thread_search_input = self.thread_search_query.clone();
                    self.focus = Focus::ThreadSearch;
                }
                KeyCode::Char('n') => {
                    if !self.thread_search_query.is_empty() {
                        self.jump_to_search_match(true);
                    }
                }
                KeyCode::Char('N') => {
                    if !self.thread_search_query.is_empty() {
                        self.jump_to_search_match(false);
                    }
                }
                _ => {}
            },
            Focus::ThreadSearch => match key {
                KeyCode::Enter => {
                    self.thread_search_query = self.thread_search_input.clone();
                    self.render_dirty = true;
                    self.focus = Focus::Messages;
                    // Jump to first match
                    if !self.thread_search_query.is_empty() {
                        self.scroll_offset = 0;
                        self.jump_to_search_match(true);
                    }
                }
                KeyCode::Esc => {
                    self.thread_search_query.clear();
                    self.thread_search_input.clear();
                    self.render_dirty = true;
                    self.focus = Focus::Messages;
                }
                KeyCode::Backspace => {
                    self.thread_search_input.pop();
                    // Live search as you type
                    self.thread_search_query = self.thread_search_input.clone();
                    self.render_dirty = true;
                }
                KeyCode::Char(c) => {
                    self.thread_search_input.push(c);
                    // Live search as you type
                    self.thread_search_query = self.thread_search_input.clone();
                    self.render_dirty = true;
                }
                _ => {}
            },
            Focus::Input => match key {
                KeyCode::Esc => {
                    self.focus = Focus::Messages;
                }
                KeyCode::Enter => {
                    // If last char is \, replace it with a newline
                    if self.input_buffer.ends_with('\\') {
                        self.input_buffer.pop();
                        self.input_buffer.push('\n');
                    } else {
                        self.send_message().await;
                    }
                }
                KeyCode::Backspace => {
                    self.input_buffer.pop();
                }
                KeyCode::Char(c) => {
                    self.input_buffer.push(c);
                }
                _ => {}
            },
        }
    }

    // --- Mouse handling (selection in message panel, clicks in sidebar) ---

    async fn handle_mouse(&mut self, mouse: crossterm::event::MouseEvent) {
        let col = mouse.column;
        let row = mouse.row;
        let area = crossterm::terminal::size().unwrap_or((80, 24));
        // Ignore mouse events with coordinates outside terminal (common in zellij)
        if col >= area.0 || row >= area.1 {
            return;
        }

        match mouse.kind {
            MouseEventKind::ScrollUp => {
                if self.in_area(col, row, self.msg_area) {
                    self.scroll_offset = self.scroll_offset.saturating_sub(3);
                } else if self.in_area(col, row, self.sidebar_area) {
                    self.sidebar_prev();
                }
            }
            MouseEventKind::ScrollDown => {
                if self.in_area(col, row, self.msg_area) {
                    let max = self.rendered_line_count.saturating_sub(self.view_height);
                    self.scroll_offset = self.scroll_offset.saturating_add(3).min(max);
                } else if self.in_area(col, row, self.sidebar_area) {
                    self.sidebar_next();
                }
            }
            MouseEventKind::Down(MouseButton::Left) => {
                if self.in_area(col, row, self.msg_area) {
                    self.focus = Focus::Messages;
                    let pos = self.mouse_to_pos(col, row);
                    self.select_start = Some(pos);
                    self.select_end = Some(pos);
                    self.selecting = true;
                } else if self.in_area(col, row, self.sidebar_area) {
                    self.focus = Focus::Sidebar;
                    self.clear_selection();
                    let visible_row = row.saturating_sub(self.sidebar_area.y).saturating_sub(1) as usize;
                    let list_offset = self.sidebar_state.offset();
                    let item_idx = list_offset + visible_row;
                    if item_idx < self.sidebar_items.len() {
                        match &self.sidebar_items[item_idx] {
                            SidebarItem::Header(title) if !title.starts_with("  ") => {
                                let title = title.clone();
                                self.toggle_section(&title);
                            }
                            SidebarItem::Conv(_) | SidebarItem::Email(_) => {
                                self.sidebar_state.select(Some(item_idx));
                                self.preview_selected().await;
                            }
                            _ => {}
                        }
                    }
                } else {
                    self.clear_selection();
                }
            }
            MouseEventKind::Drag(MouseButton::Left) => {
                if self.selecting && self.in_area(col, row, self.msg_area) {
                    self.select_end = Some(self.mouse_to_pos(col, row));
                }
            }
            MouseEventKind::Up(MouseButton::Left) => {
                if self.selecting {
                    self.selecting = false;
                    self.copy_selection_to_clipboard();
                }
            }
            _ => {}
        }
    }

    fn in_area(&self, col: u16, row: u16, area: Rect) -> bool {
        col >= area.x && col < area.x + area.width && row >= area.y && row < area.y + area.height
    }

    /// Convert mouse (col, row) to (line_index, char_col) in cached_rendered_lines.
    fn mouse_to_pos(&self, col: u16, row: u16) -> (usize, usize) {
        let row_in_view = row.max(self.msg_area.y).saturating_sub(self.msg_area.y).saturating_sub(1) as usize;
        let line_idx = self.scroll_offset + row_in_view;
        // Content starts after left border (1 char)
        let char_col = col.max(self.msg_area.x).saturating_sub(self.msg_area.x).saturating_sub(1) as usize;
        (line_idx, char_col)
    }

    fn clear_selection(&mut self) {
        self.select_start = None;
        self.select_end = None;
        self.selecting = false;
    }

    /// Normalize selection to (start, end) where start <= end in reading order.
    fn selection_normalized(&self) -> Option<((usize, usize), (usize, usize))> {
        match (self.select_start, self.select_end) {
            (Some(a), Some(b)) => {
                let (start, end) = if a.0 < b.0 || (a.0 == b.0 && a.1 <= b.1) {
                    (a, b)
                } else {
                    (b, a)
                };
                Some((start, end))
            }
            _ => None,
        }
    }

    /// Extract the plain text of a line.
    fn line_text(line: &Line) -> String {
        line.spans.iter().map(|s| s.content.as_ref()).collect()
    }

    /// Copy selected text to clipboard.
    fn copy_selection_to_clipboard(&mut self) {
        let (start, end) = match self.selection_normalized() {
            Some(r) => r,
            None => return,
        };

        let mut text = String::new();
        for i in start.0..=end.0 {
            if let Some(line) = self.cached_rendered_lines.get(i) {
                let full = Self::line_text(line);
                if start.0 == end.0 {
                    // Single line: slice from start.col to end.col
                    let from = start.1.min(full.len());
                    let to = end.1.min(full.len());
                    text.push_str(&full[from..to]);
                } else if i == start.0 {
                    let from = start.1.min(full.len());
                    text.push_str(&full[from..]);
                    text.push('\n');
                } else if i == end.0 {
                    let to = end.1.min(full.len());
                    text.push_str(&full[..to]);
                } else {
                    text.push_str(&full);
                    text.push('\n');
                }
            }
        }

        if text.is_empty() {
            return;
        }

        // Use OSC 52 escape sequence to set terminal clipboard
        use std::io::Write;
        use base64::Engine;
        let encoded = base64::engine::general_purpose::STANDARD.encode(text.as_bytes());
        let _ = write!(std::io::stdout(), "\x1b]52;c;{}\x07", encoded);
        let _ = std::io::stdout().flush();
        self.status = format!("Copied: {}", if text.len() > 40 { &text[..40] } else { &text });
    }

    // --- Sidebar navigation ---

    fn sidebar_next(&mut self) {
        let total = self.sidebar_items.len();
        if total == 0 { return; }
        let current = self.sidebar_state.selected().unwrap_or(0);
        // Find next item (skip sub-folder headers like "  Inbox (5)")
        for i in (current + 1)..total {
            match &self.sidebar_items[i] {
                SidebarItem::Header(t) if t.starts_with("  ") => continue, // skip sub-headers
                _ => {
                    self.sidebar_state.select(Some(i));
                    return;
                }
            }
        }
    }

    fn sidebar_prev(&mut self) {
        let current = self.sidebar_state.selected().unwrap_or(0);
        // Find prev item (skip sub-folder headers)
        for i in (0..current).rev() {
            match &self.sidebar_items[i] {
                SidebarItem::Header(t) if t.starts_with("  ") => continue,
                _ => {
                    self.sidebar_state.select(Some(i));
                    return;
                }
            }
        }
    }

    /// Jump sidebar selection to the first item matching the filter
    fn jump_sidebar_to_filter(&mut self) {
        if self.sidebar_filter.is_empty() { return; }
        let query = self.sidebar_filter.to_lowercase();
        for (i, item) in self.sidebar_items.iter().enumerate() {
            if self.sidebar_item_matches(item, &query) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
    }

    /// Jump to next matching sidebar item after current selection
    fn jump_sidebar_to_filter_next(&mut self) {
        if self.sidebar_filter.is_empty() { return; }
        let query = self.sidebar_filter.to_lowercase();
        let current = self.sidebar_state.selected().unwrap_or(0);
        for i in (current + 1)..self.sidebar_items.len() {
            if self.sidebar_item_matches(&self.sidebar_items[i], &query) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
        // Wrap around
        for i in 0..=current {
            if self.sidebar_item_matches(&self.sidebar_items[i], &query) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
    }

    /// Jump to prev matching sidebar item before current selection
    fn jump_sidebar_to_filter_prev(&mut self) {
        if self.sidebar_filter.is_empty() { return; }
        let query = self.sidebar_filter.to_lowercase();
        let current = self.sidebar_state.selected().unwrap_or(0);
        for i in (0..current).rev() {
            if self.sidebar_item_matches(&self.sidebar_items[i], &query) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
        // Wrap around
        for i in (current..self.sidebar_items.len()).rev() {
            if self.sidebar_item_matches(&self.sidebar_items[i], &query) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
    }

    fn sidebar_item_matches(&self, item: &SidebarItem, query: &str) -> bool {
        match item {
            SidebarItem::Conv(idx) => {
                self.conversations[*idx].display_name(&self.my_name).to_lowercase().contains(query)
            }
            SidebarItem::Email(idx) => {
                if let Some(email) = self.emails.get(*idx) {
                    let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("");
                    let from = email.get("from")
                        .and_then(|v| v.get("emailAddress"))
                        .and_then(|v| v.get("name"))
                        .and_then(|v| v.as_str())
                        .unwrap_or("");
                    subject.to_lowercase().contains(query) || from.to_lowercase().contains(query)
                } else {
                    false
                }
            }
            SidebarItem::Header(_) => false,
        }
    }

    fn toggle_favourite(&mut self) {
        if let Some(selected) = self.sidebar_state.selected() {
            let id = match &self.sidebar_items[selected] {
                SidebarItem::Conv(idx) => Some(self.conversations[*idx].id.clone()),
                SidebarItem::Email(idx) => self.emails.get(*idx)
                    .and_then(|e| e.get("id").and_then(|v| v.as_str()))
                    .map(|s| s.to_string()),
                _ => None,
            };
            if let Some(id) = id {
                if let Some(pos) = self.favourites.iter().position(|f| f == &id) {
                    self.favourites.remove(pos);
                } else {
                    self.favourites.push(id);
                }
                let _ = self.store.save_favourites(&self.favourites);
                self.rebuild_sidebar();
            }
        }
    }

    /// Toggle collapse/expand for a section header.
    /// Extracts the section name (before the arrow/count) and toggles it.
    fn toggle_section(&mut self, title: &str) {
        let name = Self::extract_section_name(title);
        if name.is_empty() { return; }
        let name = name.to_string();
        if self.collapsed_sections.contains(&name) {
            self.collapsed_sections.remove(&name);
        } else {
            self.collapsed_sections.insert(name.clone());
        }
        self.rebuild_sidebar();
        // Restore selection to the toggled section header
        for (i, item) in self.sidebar_items.iter().enumerate() {
            if let SidebarItem::Header(h) = item {
                if Self::extract_section_name(h) == name {
                    self.sidebar_state.select(Some(i));
                    break;
                }
            }
        }
    }

    fn selected_conversation_idx(&self) -> Option<usize> {
        let selected = self.sidebar_state.selected()?;
        match self.sidebar_items.get(selected)? {
            SidebarItem::Conv(idx) => Some(*idx),
            _ => None,
        }
    }

    /// Rebuild sidebar_items from conversations, grouped by section.
    /// Sections: Favourites, Activity, DMs, Channels, Meetings, Emails.
    fn rebuild_sidebar(&mut self) {
        // Remember currently selected conv ID so we can restore it
        let prev_selected_id = self.selected_conversation_idx()
            .map(|i| self.conversations[i].id.clone());

        let mut dms: Vec<usize> = Vec::new();
        let mut channels: Vec<usize> = Vec::new();
        let mut meetings: Vec<usize> = Vec::new();

        for (i, conv) in self.conversations.iter().enumerate() {
            match conv.kind() {
                ConvKind::Chat => dms.push(i),
                ConvKind::Channel => channels.push(i),
                ConvKind::Meeting => meetings.push(i),
                ConvKind::System => {}
            }
        }

        let mut items: Vec<SidebarItem> = Vec::new();
        let mut section_starts: Vec<usize> = Vec::new();

        // Helper: collect favourite conv indices and email indices
        let fav_conv_indices: Vec<usize> = self.conversations.iter().enumerate()
            .filter(|(_, c)| self.favourites.contains(&c.id))
            .map(|(i, _)| i)
            .collect();
        let fav_email_indices: Vec<usize> = self.emails.iter().enumerate()
            .filter(|(_, e)| {
                e.get("id").and_then(|v| v.as_str())
                    .map(|id| self.favourites.contains(&id.to_string()))
                    .unwrap_or(false)
            })
            .map(|(i, _)| i)
            .collect();

        // Helper: check if a section name is collapsed
        let is_collapsed = |name: &str| -> bool {
            self.collapsed_sections.iter().any(|s| name.starts_with(s))
        };

        // 1. Favourites (always show header)
        let fav_count = fav_conv_indices.len() + fav_email_indices.len();
        let fav_collapsed = is_collapsed("Favourites");
        section_starts.push(items.len());
        items.push(SidebarItem::Header(format!(
            "{} Favourites ({})",
            if fav_collapsed { "\u{25b6}" } else { "\u{25bc}" },
            fav_count
        )));
        if !fav_collapsed {
            for idx in &fav_conv_indices {
                items.push(SidebarItem::Conv(*idx));
            }
            for idx in &fav_email_indices {
                items.push(SidebarItem::Email(*idx));
            }
        }

        // 2. Activity - top 10 most recent conversations (sorted by version), excluding favourites
        let mut activity: Vec<(usize, u64)> = self.conversations.iter().enumerate()
            .filter(|(_, c)| {
                !matches!(c.kind(), ConvKind::System) && !self.favourites.contains(&c.id)
            })
            .map(|(i, c)| (i, c.version.unwrap_or(0)))
            .collect();
        activity.sort_by(|a, b| b.1.cmp(&a.1));
        activity.truncate(10);
        if !activity.is_empty() {
            let act_collapsed = is_collapsed("Activity");
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!(
                "{} Activity ({})",
                if act_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                activity.len()
            )));
            if !act_collapsed {
                for (idx, _) in &activity {
                    items.push(SidebarItem::Conv(*idx));
                }
            }
        }

        // 3. DMs
        if !dms.is_empty() {
            let dm_collapsed = is_collapsed("Direct Messages");
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!(
                "{} Direct Messages ({})",
                if dm_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                dms.len()
            )));
            if !dm_collapsed {
                for idx in dms {
                    items.push(SidebarItem::Conv(idx));
                }
            }
        }

        // 4. Channels
        if !channels.is_empty() {
            let ch_collapsed = is_collapsed("Channels");
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!(
                "{} Channels ({})",
                if ch_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                channels.len()
            )));
            if !ch_collapsed {
                for idx in channels {
                    items.push(SidebarItem::Conv(idx));
                }
            }
        }

        // 5. Meetings
        if !meetings.is_empty() {
            let mt_collapsed = is_collapsed("Meetings");
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!(
                "{} Meetings ({})",
                if mt_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                meetings.len()
            )));
            if !mt_collapsed {
                for idx in meetings {
                    items.push(SidebarItem::Conv(idx));
                }
            }
        }

        // 6. Emails - always show header
        {
            let em_collapsed = is_collapsed("Emails");
            let email_label = if !self.email_folders.is_empty() {
                format!("{} Emails ({})",
                    if em_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                    self.email_folders.len())
            } else if !self.emails.is_empty() {
                format!("{} Emails ({})",
                    if em_collapsed { "\u{25b6}" } else { "\u{25bc}" },
                    self.emails.len())
            } else if self.email_loaded {
                format!("{} Emails (failed)",
                    if em_collapsed { "\u{25b6}" } else { "\u{25bc}" })
            } else {
                format!("{} Emails (loading...)",
                    if em_collapsed { "\u{25b6}" } else { "\u{25bc}" })
            };
            section_starts.push(items.len());
            items.push(SidebarItem::Header(email_label));
            if !em_collapsed {
                if !self.email_folders.is_empty() {
                    for (folder_name, _, folder_emails) in &self.email_folders {
                        items.push(SidebarItem::Header(format!("  {} ({})", folder_name, folder_emails.len())));
                        for email in folder_emails {
                            let email_id = email.get("id").and_then(|v| v.as_str()).unwrap_or("");
                            if let Some(idx) = self.emails.iter().position(|e| {
                                e.get("id").and_then(|v| v.as_str()) == Some(email_id)
                            }) {
                                items.push(SidebarItem::Email(idx));
                            }
                        }
                    }
                } else if !self.emails.is_empty() {
                    for idx in 0..self.emails.len() {
                        items.push(SidebarItem::Email(idx));
                    }
                }
            }
        }

        self.sidebar_items = items;
        self.section_starts = section_starts;

        // Restore selection by conv/email ID, or select first item
        let mut restored = false;
        if let Some(prev_id) = prev_selected_id {
            for (i, item) in self.sidebar_items.iter().enumerate() {
                if let SidebarItem::Conv(idx) = item {
                    if self.conversations[*idx].id == prev_id {
                        self.sidebar_state.select(Some(i));
                        restored = true;
                        break;
                    }
                }
            }
        }
        if !restored {
            // Try to restore email selection
            if let Some(email_id) = &self.current_email_id {
                for (i, item) in self.sidebar_items.iter().enumerate() {
                    if let SidebarItem::Email(idx) = item {
                        if let Some(email) = self.emails.get(*idx) {
                            if email.get("id").and_then(|v| v.as_str()) == Some(email_id) {
                                self.sidebar_state.select(Some(i));
                                restored = true;
                                break;
                            }
                        }
                    }
                }
            }
        }
        if !restored {
            for (i, item) in self.sidebar_items.iter().enumerate() {
                if matches!(item, SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                    self.sidebar_state.select(Some(i));
                    break;
                }
            }
        }
    }

    // --- Actions ---

    /// Auto-preview: load messages/email for the currently selected sidebar item
    /// Preview conversation using local cached files (no API call).
    /// Fast enough for j/k navigation.
    async fn preview_selected(&mut self) {
        let selected = self.sidebar_state.selected();
        if selected.is_none() { return; }
        let selected = selected.unwrap();

        match self.sidebar_items.get(selected) {
            Some(SidebarItem::Conv(idx)) => {
                let idx = *idx;
                let conv = &self.conversations[idx];
                let id = conv.id.clone();
                let topic = conv.display_name(&self.my_name);
                if self.current_conv_id.as_deref() != Some(&id) {
                    self.has_new_messages.insert(id.clone(), false);
                    self.read_locally.insert(id.clone(), true);
                    self.current_email_id = None;
                    self.current_email_body = None;

                    self.current_conv_id = Some(id.clone());
                    self.current_conv_topic = topic;
                    self.messages.clear();
                    self.local_plain_text = None;
                    // Try local store files first (instant)
                    let mut has_local = false;
                    if let Ok(entries) = self.store.load_messages(&id) {
                        if !entries.is_empty() {
                            let mut plain = String::new();
                            for (_filename, content) in &entries {
                                plain.push_str(content);
                                if !content.ends_with('\n') {
                                    plain.push('\n');
                                }
                                plain.push('\n');
                            }
                            self.local_plain_text = Some(plain);
                            has_local = true;
                        }
                    }
                    // Fall back to API if no local data
                    if !has_local {
                        self.load_messages().await;
                    }
                    self.render_dirty = true;
                    self.scroll_offset = usize::MAX;
                }
            }
            Some(SidebarItem::Email(idx)) => {
                let idx = *idx;
                self.preview_email(idx).await;
            }
            _ => {}
        }
    }

    async fn open_conversation(&mut self) {
        let selected = self.sidebar_state.selected();
        if selected.is_none() { return; }
        let selected = selected.unwrap();

        match self.sidebar_items.get(selected) {
            Some(SidebarItem::Conv(idx)) => {
                let idx = *idx;
                let conv = &self.conversations[idx];
                let id = conv.id.clone();
                let topic = conv.display_name(&self.my_name);
                self.has_new_messages.insert(id.clone(), false);
                self.read_locally.insert(id.clone(), true);
                self.current_email_id = None;
                self.current_email_body = None;
                self.current_conv_id = Some(id);
                self.current_conv_topic = topic;
                self.thread_search_query.clear();
                self.thread_search_input.clear();
                self.local_plain_text = None;
                self.load_messages().await;
                self.scroll_offset = usize::MAX;
            }
            Some(SidebarItem::Email(idx)) => {
                let idx = *idx;
                self.preview_email(idx).await;
            }
            _ => {}
        }
    }

    /// Navigate to a conversation by ID (used by tv search result)
    async fn navigate_to_conv(&mut self, conv_id: &str, search: Option<String>) {
        for (_i, conv) in self.conversations.iter().enumerate() {
            if conv.id == conv_id {
                let topic = conv.display_name(&self.my_name);
                self.has_new_messages.insert(conv_id.to_string(), false);
                self.read_locally.insert(conv_id.to_string(), true);
                self.current_email_id = None;
                self.current_email_body = None;
                self.current_conv_id = Some(conv_id.to_string());
                self.current_conv_topic = topic;
                self.messages.clear();
                self.local_plain_text = None;
                self.load_messages().await;

                if let Some(query) = search {
                    // Set search highlight and pre-render so we can jump
                    self.thread_search_query = query.clone();
                    self.thread_search_input = query;
                    // Pre-render lines so jump_to_search_match has data
                    let width = self.msg_area.width.saturating_sub(2) as usize;
                    if width > 0 {
                        self.cached_rendered_lines = self.render_messages(width);
                    }
                    self.render_dirty = true;
                    self.scroll_offset = 0;
                    self.jump_to_search_match(true);
                } else {
                    self.thread_search_query.clear();
                    self.thread_search_input.clear();
                    self.scroll_offset = usize::MAX;
                }
                self.focus = Focus::Messages;
                break;
            }
        }
    }

    /// Load from disk cache for instant startup
    fn load_from_cache(&mut self) {
        let cache = Cache::load(self.auth.config_dir());
        if cache.conversations.is_empty() {
            return;
        }
        self.my_name = cache.my_name;
        self.cached_snippets = cache.snippets;

        // Reconstruct Conversation structs from cached data
        self.conversations = cache
            .conversations
            .into_iter()
            .map(|cc| {
                use crate::types::*;
                Conversation {
                    id: cc.id,
                    conv_type: None,
                    thread_properties: if cc.topic.is_empty() || cc.topic == "(no topic)" {
                        None
                    } else {
                        Some(ThreadProperties {
                            topic: Some(cc.topic),
                            last_join_at: None,
                            member_count: None,
                            product_thread_type: match cc.kind.as_str() {
                                "Chat" => Some("chat".to_string()),
                                _ => None,
                            },
                        })
                    },
                    last_message: Some(LastMessage {
                        id: cc.last_message_id,
                        from_display_name: None,
                        from_given_name: None,
                        imdisplayname: None,
                    }),
                    version: Some(cc.version),
                    properties: cc.consumptionhorizon.map(|h| ConvProperties {
                        consumptionhorizon: Some(h),
                        lastimreceivedtime: None,
                    }),
                    member_names: cc.member_names,
                }
            })
            .collect();

        // Seed last_message_ids so polling can detect new messages
        for conv in &self.conversations {
            if let Some(lm) = &conv.last_message {
                if let Some(id) = &lm.id {
                    self.last_message_ids.insert(conv.id.clone(), id.clone());
                }
            }
        }
    }

    /// Save current state to disk cache
    fn save_to_cache(&self) {
        let cached_convs: Vec<CachedConv> = self
            .conversations
            .iter()
            .filter(|c| !matches!(c.kind(), ConvKind::System))
            .map(|c| CachedConv {
                id: c.id.clone(),
                topic: c.topic().to_string(),
                member_names: c.member_names.clone(),
                version: c.version.unwrap_or(0),
                kind: format!("{:?}", c.kind()),
                display_name: c.display_name(&self.my_name),
                last_message_id: c
                    .last_message
                    .as_ref()
                    .and_then(|lm| lm.id.clone()),
                consumptionhorizon: c
                    .properties
                    .as_ref()
                    .and_then(|p| p.consumptionhorizon.clone()),
            })
            .collect();

        let cache = Cache {
            my_name: self.my_name.clone(),
            conversations: cached_convs,
            snippets: self.cached_snippets.clone(),
            last_refresh: std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.as_secs())
                .unwrap_or(0),
        };

        let _ = cache.save(self.auth.config_dir());
    }

    async fn preview_email(&mut self, idx: usize) {
        let email = match self.emails.get(idx) {
            Some(e) => e,
            None => return,
        };
        let email_id = email.get("id").and_then(|v| v.as_str()).unwrap_or("").to_string();
        if self.current_email_id.as_deref() == Some(&email_id) {
            return;
        }
        let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("(no subject)").to_string();
        self.current_conv_id = None;
        self.messages.clear();
        self.current_email_id = Some(email_id.clone());
        self.current_email_subject = subject.clone();
        self.current_conv_topic = subject;
        self.current_email_body = None;
        self.scroll_offset = 0;

        // Fetch full email body
        match self.api.get_email(&mut self.auth, &email_id).await {
            Ok(data) => {
                let body_html = data.get("body")
                    .and_then(|b| b.get("content"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("")
                    .to_string();
                self.current_email_body = Some(body_html);
                self.render_dirty = true;
            }
            Err(e) => {
                self.status = format!("Email error: {}", e);
                self.current_email_body = Some(format!("Error loading email: {}", e));
                self.render_dirty = true;
            }
        }
    }

    async fn load_email_folders(&mut self) {
        self.status = "Loading email folders...".to_string();
        let result = match self.api.list_mail_folders(&mut self.auth).await {
            Ok(folders) => Ok(folders),
            Err(_) => {
                // Retry with fresh token
                self.auth.clear_graph_token();
                self.api.list_mail_folders(&mut self.auth).await
            }
        };
        match result {
            Ok(folders) => {
                let mut loaded = Vec::new();
                for folder in &folders {
                    let name = folder.get("displayName").and_then(|v| v.as_str()).unwrap_or("?").to_string();
                    let id = folder.get("id").and_then(|v| v.as_str()).unwrap_or("").to_string();
                    let total = folder.get("totalItemCount").and_then(|v| v.as_u64()).unwrap_or(0);
                    if total == 0 || id.is_empty() {
                        continue;
                    }
                    match self.api.list_emails(&mut self.auth, &id, 15).await {
                        Ok(emails) => {
                            if !emails.is_empty() {
                                loaded.push((name, id, emails));
                            }
                        }
                        Err(_) => {}
                    }
                }
                self.email_folders = loaded;
                // Flatten into self.emails for backward compat
                self.emails = self.email_folders.iter()
                    .flat_map(|(_, _, emails)| emails.clone())
                    .collect();
                self.email_loaded = true;
                self.rebuild_sidebar();
                self.status = format!("{} conversations, {} email folders", self.conversations.len(), self.email_folders.len());
            }
            Err(e) => {
                self.status = format!("Email folders failed: {} (r to retry)", e);
                self.email_loaded = true;
                self.rebuild_sidebar();
            }
        }
    }

    async fn load_conversations(&mut self) {
        self.status = "Loading conversations...".to_string();
        // Clear polling state on full refresh
        self.has_new_messages.clear();

        match self.api.list_conversations(&mut self.auth, 100).await {
            Ok(resp) => {
                let mut convs = resp.conversations;
                // Sort by version (most recent activity first)
                convs.sort_by(|a, b| b.version.unwrap_or(0).cmp(&a.version.unwrap_or(0)));

                // Build lookup from old conversations for member names
                let old_member_names: HashMap<String, Vec<String>> = self
                    .conversations
                    .iter()
                    .filter(|c| !c.member_names.is_empty())
                    .map(|c| (c.id.clone(), c.member_names.clone()))
                    .collect();
                let old_snippets = std::mem::take(&mut self.cached_snippets);

                let mut name_counts: HashMap<String, usize> = HashMap::new();
                for conv in convs.iter_mut() {
                    if matches!(conv.kind(), ConvKind::System) {
                        continue;
                    }

                    let needs_names = conv.topic() == "(no topic)";

                    // Check if we already have cached data for this conv
                    let last_msg_id = conv
                        .last_message
                        .as_ref()
                        .and_then(|lm| lm.id.as_deref())
                        .unwrap_or("");
                    let old_msg_id = self
                        .last_message_ids
                        .get(&conv.id)
                        .map(|s| s.as_str())
                        .unwrap_or("");

                    let has_cached = old_snippets.contains_key(&conv.id);
                    let unchanged = !last_msg_id.is_empty()
                        && !old_msg_id.is_empty()
                        && last_msg_id == old_msg_id;

                    // Restore member names from previous data
                    if needs_names {
                        if let Some(names) = old_member_names.get(&conv.id) {
                            conv.member_names = names.clone();
                        }
                    }

                    if has_cached && unchanged {
                        // Reuse cached snippets - no need to re-fetch
                        if let Some(snippets) = old_snippets.get(&conv.id) {
                            self.cached_snippets
                                .insert(conv.id.clone(), snippets.clone());
                        }
                        continue;
                    }

                    let fetch_count = if needs_names { 10 } else { 5 };
                    if let Ok(msg_resp) = self
                        .api
                        .get_messages(&mut self.auth, &conv.id, fetch_count)
                        .await
                    {
                        if needs_names {
                            let mut names: Vec<String> = Vec::new();
                            for m in &msg_resp.messages {
                                if let Some(name) = &m.imdisplayname {
                                    if !name.is_empty() && !names.contains(name) {
                                        names.push(name.clone());
                                        *name_counts.entry(name.clone()).or_insert(0) += 1;
                                    }
                                }
                            }
                            if !names.is_empty() {
                                conv.member_names = names;
                            }
                        }

                        let snippets: Vec<String> = msg_resp
                            .messages
                            .iter()
                            .filter_map(|m| m.content.as_ref().map(|c| html::strip_tags_only(c)))
                            .filter(|s| !s.is_empty())
                            .collect();
                        if !snippets.is_empty() {
                            self.cached_snippets.insert(conv.id.clone(), snippets);
                        }
                    }
                }

                // Detect own name
                if self.my_name.is_empty() {
                    if let Some((name, _)) = name_counts.iter().max_by_key(|(_, c)| *c) {
                        self.my_name = name.clone();
                    }
                }

                // Record last message IDs
                for conv in &convs {
                    if let Some(lm) = &conv.last_message {
                        if let Some(id) = &lm.id {
                            self.last_message_ids.insert(conv.id.clone(), id.clone());
                        }
                    }
                }

                self.conversations = convs;
                self.rebuild_sidebar();
                self.status = format!("{} conversations", self.conversations.len());

                // Save to cache for next startup
                self.save_to_cache();

                // Write to local file store
                let mut conv_index = Vec::new();
                for conv in &self.conversations {
                    let kind = conv.kind();
                    if matches!(kind, ConvKind::System) { continue; }
                    let meta = store::ConvMeta {
                        name: conv.display_name(&self.my_name),
                        kind: format!("{:?}", kind),
                        members: conv.member_names.clone(),
                        unread: conv.is_unread(),
                        version: conv.version.unwrap_or(0),
                        last_message_id: conv.last_message.as_ref().and_then(|lm| lm.id.clone()),
                        consumptionhorizon: conv.properties.as_ref().and_then(|p| p.consumptionhorizon.clone()),
                    };
                    let _ = self.store.save_conv_meta(&conv.id, &meta);
                    conv_index.push(store::ConvIndex {
                        id: conv.id.clone(),
                        name: conv.display_name(&self.my_name),
                        kind: format!("{:?}", kind),
                        last_activity: conv.version.unwrap_or(0),
                        unread: conv.is_unread(),
                    });
                }
                let index = store::Index {
                    my_name: self.my_name.clone(),
                    conversations: conv_index,
                    email_folders: Vec::new(), // emails synced separately
                };
                let _ = self.store.save_index(&index);
            }
            Err(e) => {
                let msg = e.to_string();
                if Self::is_token_error(&msg) {
                    self.status = "Token expired! Press L to re-login (r to retry)".to_string();
                } else {
                    self.status = format!("Error: {} (r to retry)", e);
                }
            }
        }
    }

    /// Suspend TUI, spawn tv for search, navigate to selected conversation
    fn spawn_tv(&mut self, terminal: &mut DefaultTerminal) {
        use std::process::Command;

        let data_dir = self.auth.config_dir().join("data");
        let data_path = data_dir.to_string_lossy();

        Self::suspend_tui();

        let preview_script_path = self.auth.config_dir().join("tv-preview.sh");
        let preview_cmd = format!("bash {} '{{}}'", preview_script_path.to_string_lossy());

        let source_cmd = format!(
            "rg . --no-heading --line-number --color=never --sortr=path {}",
            data_path
        );

        // Run tv with all stdio inherited (proven to work).
        // Capture selection by redirecting stdout to temp file in a shell wrapper.
        let outfile = std::env::temp_dir().join("teamsh-tv-out.txt");
        let _ = std::fs::remove_file(&outfile);

        let shell_cmd = format!(
            "tv --source-command '{}' --preview-command '{}' --preview-word-wrap > '{}'",
            source_cmd.replace('\'', "'\\''"),
            preview_cmd.replace('\'', "'\\''"),
            outfile.to_string_lossy(),
        );

        let _ = Command::new("bash")
            .arg("-c")
            .arg(&shell_cmd)
            .stdin(std::process::Stdio::inherit())
            .stdout(std::process::Stdio::inherit())
            .stderr(std::process::Stdio::inherit())
            .status();

        let selected = std::fs::read_to_string(&outfile)
            .unwrap_or_default()
            .trim()
            .to_string();
        let _ = std::fs::remove_file(&outfile);

        Self::resume_tui(terminal);
        self.render_dirty = true;

        // Navigate to selected conversation if tv returned a result
        // Format: /path/data/conversations/{safe_id}/messages/{msg}.txt:line_num:matched_text
        if selected.is_empty() {
            self.status = "tv: no selection returned".to_string();
        } else {
            self.status = format!("tv: {}", &selected[..selected.len().min(60)]);
        }
        if !selected.is_empty() {
            let conv_prefix = format!("{}/conversations/", data_path);
            if let Some(rest) = selected.strip_prefix(&conv_prefix) {
                if let Some(safe_id) = rest.split('/').next() {
                    // Extract matched text (after .txt:line:)
                    let matched_text = rest
                        .find(".txt:")
                        .and_then(|pos| {
                            let after_txt = &rest[pos + 5..]; // skip ".txt:"
                            // skip line number
                            after_txt.find(':').map(|p| after_txt[p + 1..].trim().to_string())
                        })
                        .unwrap_or_default();

                    for (sidebar_idx, item) in self.sidebar_items.iter().enumerate() {
                        if let SidebarItem::Conv(conv_idx) = item {
                            let conv_id = &self.conversations[*conv_idx].id;
                            if store::safe_filename(conv_id) == safe_id {
                                self.sidebar_state.select(Some(sidebar_idx));
                                self.pending_tv_nav = Some(conv_id.clone());
                                // Set search query so matched text is highlighted
                                if !matched_text.is_empty() {
                                    self.pending_tv_search = Some(matched_text);
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }
    }

    /// Suspend TUI, run device code login for Teams, resume TUI
    async fn spawn_login(&mut self, terminal: &mut DefaultTerminal) {
        Self::suspend_tui();

        println!("\n--- Teams Re-Login ---");
        println!("1. Open https://teams.cloud.microsoft in browser");
        println!("2. DevTools (F12) > Network > filter: login.microsoftonline.com");
        println!("3. Find POST to oauth2/v2.0/token > Response > copy refresh_token");
        println!();
        println!("Paste refresh token:");

        let mut rt_input = String::new();
        if std::io::stdin().read_line(&mut rt_input).is_ok() {
            let rt_input = rt_input.trim();
            if !rt_input.is_empty() {
                std::fs::write(self.auth.config_dir().join("refresh_token"), rt_input).ok();
                self.auth.clear_access_token();
                match self.auth.access_token().await {
                    Ok(_) => println!("Login OK! Resuming..."),
                    Err(e) => println!("Token invalid: {}", e),
                }
            } else {
                println!("Skipped.");
            }
        }

        std::thread::sleep(std::time::Duration::from_secs(1));

        Self::resume_tui(terminal);
        self.render_dirty = true;

        // Refresh data
        self.load_conversations().await;
    }

    /// Jump scroll to the next (or previous) line matching the search query.
    fn jump_to_search_match(&mut self, forward: bool) {
        if self.thread_search_query.is_empty() || self.cached_rendered_lines.is_empty() {
            return;
        }
        let query_lower = self.thread_search_query.to_lowercase();
        let total = self.cached_rendered_lines.len();
        let start = if forward {
            self.scroll_offset + 1
        } else {
            self.scroll_offset.saturating_sub(1)
        };

        let range: Box<dyn Iterator<Item = usize>> = if forward {
            Box::new((start..total).chain(0..start))
        } else {
            Box::new((0..=start).rev().chain((start + 1..total).rev()))
        };

        for i in range {
            let line_text: String = self.cached_rendered_lines[i]
                .spans
                .iter()
                .map(|s| s.content.as_ref())
                .collect();
            if line_text.to_lowercase().contains(&query_lower) {
                // Position match near top of viewport
                self.scroll_offset = i.saturating_sub(2);
                return;
            }
        }
    }

    async fn load_messages(&mut self) {
        if let Some(conv_id) = &self.current_conv_id.clone() {
            // In demo mode, load from local store files only
            if self.demo {
                if let Ok(entries) = self.store.load_messages(conv_id) {
                    if !entries.is_empty() {
                        let mut plain = String::new();
                        for (_filename, content) in &entries {
                            plain.push_str(content);
                            if !content.ends_with('\n') { plain.push('\n'); }
                            plain.push('\n');
                        }
                        self.local_plain_text = Some(plain);
                        self.messages.clear();
                        self.render_dirty = true;
                    }
                }
                return;
            }
            match self.api.get_messages(&mut self.auth, conv_id, 50).await {
                Ok(resp) => {
                    let mut msgs: Vec<Message> = resp
                        .messages
                        .into_iter()
                        .filter(|m| {
                            matches!(
                                m.messagetype.as_deref(),
                                Some("RichText/Html") | Some("Text")
                            )
                        })
                        .collect();
                    msgs.sort_by(|a, b| a.timestamp.cmp(&b.timestamp));

                    // Update snippet cache
                    let snippets: Vec<String> = msgs
                        .iter()
                        .filter_map(|m| m.content.as_ref().map(|c| html::strip_tags_only(c)))
                        .filter(|s| !s.is_empty())
                        .collect();
                    if !snippets.is_empty() {
                        self.cached_snippets.insert(conv_id.clone(), snippets);
                    }
                    self.messages = msgs;
                    self.render_dirty = true;
                    // Write messages to local files
                    if let Some(conv_id) = &self.current_conv_id {
                        for m in &self.messages {
                            let msg_id = m.id.as_deref().unwrap_or("unknown");
                            let sender = m.imdisplayname.as_deref()
                                .filter(|s| !s.is_empty())
                                .unwrap_or(if self.my_name.is_empty() { "You" } else { &self.my_name });
                            let timestamp = m.timestamp.as_deref().unwrap_or("");
                            let content = m.content.as_deref().unwrap_or("");
                            let _ = self.store.save_message(conv_id, msg_id, timestamp, sender, "", content);
                        }
                    }
                }
                Err(e) => {
                    self.status = format!("Error: {}", e);
                }
            }
        }
    }

    async fn send_message(&mut self) {
        let text = self.input_buffer.trim().to_string();
        if text.is_empty() {
            return;
        }
        if self.demo {
            self.status = "Demo mode: sending disabled".to_string();
            return;
        }
        // Convert newlines to <br> for Teams HTML
        let html_text = text.replace('\n', "<br>");
        if let Some(conv_id) = &self.current_conv_id.clone() {
            match self.api.send_message(&mut self.auth, conv_id, &html_text).await {
                Ok(()) => {
                    self.input_buffer.clear();
                    self.focus = Focus::Messages;
                    self.load_messages().await;
                    self.scroll_offset = usize::MAX; // scroll to bottom
                }
                Err(e) => {
                    self.status = format!("Send error: {}", e);
                }
            }
        }
    }

    // --- Polling ---

    async fn poll_new_messages(&mut self) {
        match self.api.list_conversations(&mut self.auth, 100).await {
            Ok(resp) => {
                let mut found_new = false;
                let mut current_conv_has_new = false;
                for conv in &resp.conversations {
                    if let Some(lm) = &conv.last_message {
                        if let Some(id) = &lm.id {
                            if let Some(old_id) = self.last_message_ids.get(&conv.id) {
                                if id != old_id {
                                    self.has_new_messages.insert(conv.id.clone(), true);
                                    found_new = true;
                                    if self.current_conv_id.as_deref() == Some(&conv.id) {
                                        current_conv_has_new = true;
                                    }
                                }
                            }
                            self.last_message_ids.insert(conv.id.clone(), id.clone());
                        }
                    }
                }

                // Update conversation list with fresh data
                let mut convs = resp.conversations;
                convs.sort_by(|a, b| b.version.unwrap_or(0).cmp(&a.version.unwrap_or(0)));
                // Preserve member names from old conversations
                let old_names: HashMap<String, Vec<String>> = self.conversations.iter()
                    .filter(|c| !c.member_names.is_empty())
                    .map(|c| (c.id.clone(), c.member_names.clone()))
                    .collect();
                for conv in convs.iter_mut() {
                    if conv.member_names.is_empty() {
                        if let Some(names) = old_names.get(&conv.id) {
                            conv.member_names = names.clone();
                        }
                    }
                }
                self.conversations = convs;
                self.rebuild_sidebar();

                // Auto-reload messages for the currently open conversation
                if current_conv_has_new {
                    let at_bottom = self.scroll_offset >= self.rendered_line_count.saturating_sub(self.view_height + 5);
                    self.render_dirty = true;
                    self.load_messages().await;
                    if at_bottom {
                        self.scroll_offset = usize::MAX; // stay at bottom
                    } else {
                        self.has_new_below = true;
                    }
                }

                if found_new {
                    print!("\x07");
                }
            }
            Err(e) => {
                let msg = e.to_string();
                if Self::is_token_error(&msg) {
                    self.status = "Token expired! Press L to re-login".to_string();
                }
            }
        }
    }
}

/// Wrap a text string to fit within `width` columns, respecting Unicode width.
/// Pipe plain text through bat for syntax highlighting, then parse ANSI into ratatui Lines.
/// If bat is unavailable, returns plain unstyled lines. Wraps to width.
/// If `search_query` is non-empty, highlights matching text with a background color.
/// Style plain text into ratatui Lines with syntax highlighting and word wrap.
/// Replaces the old bat subprocess approach for instant rendering.
fn style_lines(plain: &str, width: usize, search_query: &str) -> Vec<Line<'static>> {
    use unicode_width::UnicodeWidthStr;

    let style_default = Style::default().fg(Color::Rgb(205, 214, 244)); // #cdd6f4 text
    let style_sender = Style::default().fg(Color::Rgb(166, 227, 161)).add_modifier(Modifier::BOLD); // #a6e3a1
    let style_timestamp = Style::default().fg(Color::Rgb(137, 180, 250)); // #89b4fa
    let style_mention = Style::default().fg(Color::Rgb(249, 226, 175)).add_modifier(Modifier::BOLD); // #f9e2af
    let style_quote = Style::default().fg(Color::Rgb(147, 153, 178)).add_modifier(Modifier::ITALIC); // #9399b2
    let style_link = Style::default().fg(Color::Rgb(137, 180, 250)).add_modifier(Modifier::UNDERLINED);
    let style_keyword = Style::default().fg(Color::Rgb(203, 166, 247)).add_modifier(Modifier::BOLD); // #cba6f7
    let style_channel = Style::default().fg(Color::Rgb(148, 226, 213)); // #94e2d5

    let mut lines: Vec<Line<'static>> = Vec::new();

    for text_line in plain.lines() {
        let spans = style_line(text_line, style_default, style_sender, style_timestamp,
                               style_mention, style_quote, style_link, style_keyword, style_channel);

        // Wrap to width: collect all chars with their style, then chunk
        let display_width: usize = spans.iter().map(|s| s.content.width()).sum();
        if width > 0 && display_width > width {
            // Flatten spans into (char, style) pairs, then re-chunk by width
            let mut chars: Vec<(char, Style)> = Vec::new();
            for span in &spans {
                for ch in span.content.chars() {
                    chars.push((ch, span.style));
                }
            }
            let mut line_spans: Vec<Span<'static>> = Vec::new();
            let mut current_text = String::new();
            let mut current_style = style_default;
            let mut current_width = 0;
            for (ch, style) in chars {
                let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(0);
                if current_width + cw > width && current_width > 0 {
                    if !current_text.is_empty() {
                        line_spans.push(Span::styled(std::mem::take(&mut current_text), current_style));
                    }
                    lines.push(Line::from(std::mem::take(&mut line_spans)));
                    current_width = 0;
                }
                if style != current_style && !current_text.is_empty() {
                    line_spans.push(Span::styled(std::mem::take(&mut current_text), current_style));
                }
                current_style = style;
                current_text.push(ch);
                current_width += cw;
            }
            if !current_text.is_empty() {
                line_spans.push(Span::styled(current_text, current_style));
            }
            if !line_spans.is_empty() {
                lines.push(Line::from(line_spans));
            }
        } else {
            lines.push(Line::from(spans));
        }
    }

    // Apply search highlighting
    if !search_query.is_empty() {
        let query_lower = search_query.to_lowercase();
        let hl_style = Style::default()
            .bg(Color::Rgb(243, 139, 168))  // Catppuccin red bg
            .fg(Color::Rgb(17, 17, 27))      // crust fg
            .add_modifier(ratatui::style::Modifier::BOLD);
        for line in &mut lines {
            let full_text: String = line.spans.iter().map(|s| s.content.as_ref()).collect();
            if !full_text.to_lowercase().contains(&query_lower) { continue; }
            let mut new_spans: Vec<Span<'static>> = Vec::new();
            for span in &line.spans {
                let span_text = span.content.as_ref();
                let span_lower = span_text.to_lowercase();
                if let Some(pos) = span_lower.find(&query_lower) {
                    if pos > 0 {
                        new_spans.push(Span::styled(span_text[..pos].to_string(), span.style));
                    }
                    let end = (pos + search_query.len()).min(span_text.len());
                    new_spans.push(Span::styled(span_text[pos..end].to_string(), hl_style));
                    if end < span_text.len() {
                        new_spans.push(Span::styled(span_text[end..].to_string(), span.style));
                    }
                } else {
                    new_spans.push(span.clone());
                }
            }
            *line = Line::from(new_spans);
        }
    }

    lines
}

/// Style a single line of text based on its content pattern.
fn style_line(
    text: &str,
    style_default: Style,
    style_sender: Style,
    style_timestamp: Style,
    style_mention: Style,
    style_quote: Style,
    style_link: Style,
    style_keyword: Style,
    style_channel: Style,
) -> Vec<Span<'static>> {
    // Header line: "Name  2026 Mar 05 14:32  " (name first, then double-space, then timestamp)
    // Regex-like: starts with non-space, has "  YYYY Mon DD HH:MM"
    if let Some(ts_start) = find_timestamp(text) {
        let name = &text[..ts_start].trim_end();
        let ts_text = text.get(ts_start..ts_start + 17).unwrap_or("");
        let rest = text.get(ts_start + 17..).unwrap_or("");
        let mut spans = vec![
            Span::styled(name.to_string(), style_sender),
            Span::styled("  ".to_string(), style_default),
            Span::styled(ts_text.to_string(), style_timestamp),
        ];
        if !rest.is_empty() {
            spans.push(Span::styled(rest.to_string(), style_default));
        }
        return spans;
    }

    // Quoted text (starts with optional whitespace then >)
    let trimmed = text.trim_start();
    if trimmed.starts_with('>') {
        return vec![Span::styled(text.to_string(), style_quote)];
    }

    // Email headers
    if text.starts_with("From:") || text.starts_with("Date:") || text.starts_with("Subject:") {
        if let Some(colon_pos) = text.find(": ") {
            return vec![
                Span::styled(text[..colon_pos + 1].to_string(), style_keyword),
                Span::styled(text[colon_pos + 1..].to_string(), style_default),
            ];
        }
    }

    // Body line - highlight @mentions, #[channels], URLs
    style_body(text, style_default, style_mention, style_link, style_channel)
}

/// Find the byte offset of a timestamp pattern "YYYY Mon DD HH:MM" preceded by double-space.
fn find_timestamp(text: &str) -> Option<usize> {
    // Look for "  DDDD Www DD HH:MM" pattern
    let bytes = text.as_bytes();
    let len = bytes.len();
    if len < 19 { return None; }
    for i in 2..len.saturating_sub(16) {
        if bytes[i - 1] == b' ' && bytes[i - 2] == b' '
            && bytes[i].is_ascii_digit() && bytes[i + 1].is_ascii_digit()
            && bytes[i + 2].is_ascii_digit() && bytes[i + 3].is_ascii_digit()
            && bytes[i + 4] == b' '
            && bytes[i + 5].is_ascii_uppercase()
            && bytes[i + 8] == b' '
            && bytes[i + 11] == b' '
            && bytes[i + 14] == b':'
        {
            return Some(i);
        }
    }
    None
}

/// Style body text with @mentions, #[channels], and URLs highlighted.
fn style_body(
    text: &str,
    style_default: Style,
    style_mention: Style,
    style_link: Style,
    style_channel: Style,
) -> Vec<Span<'static>> {
    let mut spans: Vec<Span<'static>> = Vec::new();
    let mut pos = 0;
    let bytes = text.as_bytes();
    let len = bytes.len();

    while pos < len {
        if bytes[pos] == b'@' && pos + 1 < len && bytes[pos + 1].is_ascii_uppercase() {
            // @Mention - consume word
            if pos > 0 {
                let prev = &text[..pos];
                if !prev.is_empty() {
                    // Flush previous default text (only from last flushed position)
                }
            }
            let start = pos;
            pos += 1;
            while pos < len && (bytes[pos].is_ascii_alphanumeric() || bytes[pos] == b'_' || bytes[pos] == b'-') {
                pos += 1;
            }
            spans.push(Span::styled(text[start..pos].to_string(), style_mention));
            continue;
        }

        if bytes[pos] == b'#' && pos + 1 < len && bytes[pos + 1] == b'[' {
            // #[Channel Name]
            if let Some(close) = text[pos + 2..].find(']') {
                let end = pos + 2 + close + 1;
                spans.push(Span::styled(text[pos..end].to_string(), style_channel));
                pos = end;
                continue;
            }
        }

        if bytes[pos] == b'h' && text[pos..].starts_with("http") {
            // URL
            let start = pos;
            while pos < len && !bytes[pos].is_ascii_whitespace()
                && bytes[pos] != b')' && bytes[pos] != b'>' && bytes[pos] != b']'
            {
                pos += 1;
            }
            spans.push(Span::styled(text[start..pos].to_string(), style_link));
            continue;
        }

        // Regular character - accumulate
        let start = pos;
        while pos < len
            && !(bytes[pos] == b'@' && pos + 1 < len && bytes[pos + 1].is_ascii_uppercase())
            && !(bytes[pos] == b'#' && pos + 1 < len && bytes[pos + 1] == b'[')
            && !(bytes[pos] == b'h' && text[pos..].starts_with("http"))
        {
            pos += 1;
        }
        spans.push(Span::styled(text[start..pos].to_string(), style_default));
    }

    if spans.is_empty() {
        spans.push(Span::styled(text.to_string(), style_default));
    }
    spans
}

