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
use crate::types::{Conversation, ConvKind, Message};

const SENDER_COLORS: [Color; 8] = [
    Color::Green,
    Color::Cyan,
    Color::Magenta,
    Color::Yellow,
    Color::Blue,
    Color::Red,
    Color::LightGreen,
    Color::LightCyan,
];

#[derive(Debug, Clone, PartialEq)]
enum Focus {
    Sidebar,
    Messages,
    Input,
    Search,
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
    focus: Focus,
    exit: bool,

    // Sidebar
    conversations: Vec<Conversation>,
    sidebar_items: Vec<SidebarItem>,
    sidebar_state: ListState,
    section_starts: Vec<usize>, // sidebar_items index where each section header is
    favourites: Vec<String>,    // favourite conv/email IDs

    // Messages
    current_conv_id: Option<String>,
    current_conv_topic: String,
    messages: Vec<Message>,
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
    sender_colors: HashMap<String, Color>,
    next_color_idx: usize,

    // Search
    search_active: bool,
    search_query: String,
    search_results: Vec<usize>,       // conversation indices
    search_email_results: Vec<usize>, // email indices
    search_list_state: ListState,
    search_people_results: Vec<(String, String)>,
    search_highlight: String, // active search term to highlight in message view

    // Message content cache for search
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

    // Emails (Microsoft Graph)
    email_folders: Vec<(String, String, Vec<serde_json::Value>)>, // (folder_name, folder_id, emails)
    emails: Vec<serde_json::Value>,
    current_email_id: Option<String>,
    current_email_body: Option<String>, // HTML body of selected email
    current_email_subject: String,
    email_loaded: bool,
}

impl App {
    pub async fn new() -> Result<Self> {
        let auth = Auth::new()?;
        let api = Api::new(&auth.region());

        let mut app = Self {
            auth,
            api,
            focus: Focus::Sidebar,
            exit: false,
            conversations: Vec::new(),
            sidebar_items: Vec::new(),
            sidebar_state: ListState::default(),
            section_starts: Vec::new(),
            favourites: Vec::new(),
            current_conv_id: None,
            current_conv_topic: String::new(),
            messages: Vec::new(),
            scroll_offset: 0,
            rendered_line_count: 0,
            view_height: 20,
            sidebar_area: Rect::default(),
            msg_area: Rect::default(),
            input_buffer: String::new(),
            status: "Loading...".to_string(),
            my_name: String::new(),
            sender_colors: HashMap::new(),
            next_color_idx: 0,
            search_active: false,
            search_query: String::new(),
            search_results: Vec::new(),
            search_list_state: ListState::default(),
            search_email_results: Vec::new(),
            search_people_results: Vec::new(),
            search_highlight: String::new(),
            cached_snippets: HashMap::new(),
            tick_count: 0,
            last_message_ids: HashMap::new(),
            has_new_messages: HashMap::new(),
            read_locally: HashMap::new(),
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
        };

        // Try loading from cache first for instant startup
        app.load_from_cache();
        if app.conversations.is_empty() {
            app.load_conversations().await;
            // Mark all as read on first load (user has already seen everything)
            for conv in &app.conversations {
                app.read_locally.insert(conv.id.clone(), true);
            }
        } else {
            // From cache - mark all as read (stale unread state)
            for conv in &app.conversations {
                app.read_locally.insert(conv.id.clone(), true);
            }
            app.status = format!("{} conversations (cached)", app.conversations.len());
        }

        // Load emails in background (don't block startup)
        app.load_email_folders().await;

        // Load favourites from store
        if let Ok(store) = crate::store::Store::new(app.auth.config_dir()) {
            if let Ok(favs) = store.load_favourites() {
                app.favourites = favs;
            }
        }

        Ok(app)
    }

    pub async fn run(&mut self, terminal: &mut DefaultTerminal) -> Result<()> {
        while !self.exit {
            terminal.draw(|frame| self.draw(frame))?;
            if event::poll(Duration::from_millis(100))? {
                match event::read()? {
                    Event::Key(key) => {
                        if key.kind == KeyEventKind::Press {
                            self.handle_key(key.code, key.modifiers).await;
                        }
                    }
                    Event::Mouse(mouse) => {
                        self.handle_mouse(mouse).await;
                    }
                    Event::Resize(_, _) => {
                        self.render_dirty = true;
                    }
                    _ => {}
                }
            }

            self.tick_count += 1;
            // Poll every ~15 seconds (150 ticks * 100ms)
            if self.tick_count >= 150 {
                self.tick_count = 0;
                self.poll_new_messages().await;
            }
        }
        Ok(())
    }

    fn color_for_sender(&mut self, name: &str) -> Color {
        if let Some(&color) = self.sender_colors.get(name) {
            return color;
        }
        let color = SENDER_COLORS[self.next_color_idx % SENDER_COLORS.len()];
        self.next_color_idx += 1;
        self.sender_colors.insert(name.to_string(), color);
        color
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
        if self.search_active {
            self.draw_search(frame);
            return;
        }

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
                    let kind = conv.kind();
                    let prefix = match kind {
                        ConvKind::Channel => "#",
                        ConvKind::Chat => "@",
                        ConvKind::Meeting => "M",
                        _ => " ",
                    };
                    let name = conv.display_name(&self.my_name);
                    let is_current = self.current_conv_id.as_deref() == Some(&conv.id);
                    let read_local = self.read_locally.get(&conv.id).copied().unwrap_or(false);
                    let unread = if read_local {
                        self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                    } else {
                        conv.is_unread()
                            || self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                    };

                    let line = if is_current {
                        Line::from(vec![
                            Span::styled(
                                format!("  {} ", prefix),
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
                                format!("  {} ", prefix),
                                Style::default().fg(Color::White).add_modifier(Modifier::BOLD),
                            ),
                            Span::styled(
                                name,
                                Style::default().fg(Color::White).add_modifier(Modifier::BOLD),
                            ),
                        ])
                    } else {
                        Line::from(vec![
                            Span::styled(
                                format!("  {} ", prefix),
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

                        let line = if is_current {
                            Line::from(vec![
                                Span::styled("  ", Style::default()),
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
                                Span::styled("  ", Style::default()),
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
                                Span::styled("  ", Style::default().fg(Color::DarkGray)),
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

        let border_color = if self.focus == Focus::Sidebar {
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

        let status_text = if self.focus == Focus::Sidebar {
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
        let [header_area, msg_area, input_area, help_area] = Layout::vertical([
            Constraint::Length(1),
            Constraint::Min(1),
            Constraint::Length(if is_email { 0 } else { 3 }),
            Constraint::Length(1),
        ])
        .areas(area);

        // Header
        let mut header_spans = vec![
            Span::styled(" ", Style::default()),
            Span::styled(
                &self.current_conv_topic,
                Style::default()
                    .fg(Color::Cyan)
                    .add_modifier(Modifier::BOLD),
            ),
        ];
        if !self.search_highlight.is_empty() {
            header_spans.push(Span::styled(
                format!("  [search: {}]", self.search_highlight),
                Style::default().fg(Color::Yellow),
            ));
        }
        let header = Paragraph::new(Line::from(header_spans));
        frame.render_widget(header, header_area);

        // Messages - manual wrapping for correct scroll
        let msg_border_color = if self.focus == Focus::Messages {
            Color::Cyan
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

        // Slice visible lines
        let visible: Vec<Line> = wrapped_lines
            .into_iter()
            .skip(self.scroll_offset)
            .take(view_height)
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
            Color::Cyan
        } else {
            Color::DarkGray
        };
        let input = Paragraph::new(self.input_buffer.as_str()).block(
            Block::default()
                .borders(Borders::ALL)
                .title(" Message ")
                .border_style(Style::default().fg(input_border)),
        );
        frame.render_widget(input, input_area);

        if self.focus == Focus::Input {
            let cursor_x = input_area.x + 1 + self.input_buffer.width() as u16;
            let cursor_x = cursor_x.min(input_area.x + input_area.width - 2);
            frame.set_cursor_position((cursor_x, input_area.y + 1));
        }

        // Help
        let help = match self.focus {
            Focus::Input => " Enter:send  Esc:cancel ",
            Focus::Messages => " j/k:scroll  PgUp/Dn  G:end  g:top  h:sidebar  i:compose  /:search  r:refresh ",
            _ => " Tab:section  j/k:nav  l:messages  /:search  f:fav  r:refresh  e:emails  q:quit ",
        };
        let help_widget = Paragraph::new(
            Line::from(help).style(Style::default().fg(Color::DarkGray)),
        );
        frame.render_widget(help_widget, help_area);
    }

    /// Render messages into pre-wrapped lines for correct scroll calculation.
    fn render_messages(&mut self, width: usize) -> Vec<Line<'static>> {
        if width == 0 {
            return Vec::new();
        }

        // Pre-assign colors for all senders to avoid borrow issues
        let my_name = self.my_name.clone();
        let senders: Vec<String> = self
            .messages
            .iter()
            .filter_map(|m| m.imdisplayname.clone())
            .collect();
        for s in &senders {
            self.color_for_sender(s);
        }

        let mut lines: Vec<Line<'static>> = Vec::new();

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
            let time = m
                .timestamp
                .as_deref()
                .unwrap_or("")
                .get(11..16)
                .unwrap_or("??:??");
            let sender_color = self.sender_colors.get(sender).copied().unwrap_or(Color::Green);
            let is_me = !my_name.is_empty() && sender == my_name;

            // Sender line
            lines.push(Line::from(vec![
                Span::styled(
                    format!("{} ", time),
                    Style::default().fg(Color::DarkGray),
                ),
                Span::styled(
                    sender.to_string(),
                    Style::default()
                        .fg(if is_me { Color::Yellow } else { sender_color })
                        .add_modifier(Modifier::BOLD),
                ),
            ]));

            // Rich content lines with formatting
            let rich_lines = html::strip_html_rich(content);
            let highlight = &self.search_highlight;
            for rich_line in &rich_lines {
                let mut spans = vec![Span::raw("  ".to_string())];
                spans.extend(html::rich_to_spans(rich_line));

                // Apply search highlighting if active
                if !highlight.is_empty() {
                    spans = apply_search_highlight(spans, highlight);
                }

                let line_text: String = spans.iter().map(|s| s.content.as_ref()).collect();
                if line_text.len() > width && width > 4 {
                    let style = if rich_line.iter().any(|s| s.quote) {
                        Style::default().fg(Color::DarkGray)
                    } else if rich_line.iter().any(|s| s.bold) {
                        Style::default().add_modifier(Modifier::BOLD)
                    } else {
                        Style::default()
                    };
                    for wrapped in wrap_text(&format!("  {}", line_text.trim()), width) {
                        lines.push(Line::from(Span::styled(wrapped, style)));
                    }
                } else {
                    lines.push(Line::from(spans));
                }
            }

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
                    lines.push(Line::from(Span::styled(
                        format!("  {}", reaction_str),
                        Style::default().fg(Color::DarkGray),
                    )));
                }
            }

            lines.push(Line::from(""));
        }

        lines
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

            lines.push(Line::from(vec![
                Span::styled("From: ", Style::default().fg(Color::DarkGray)),
                Span::styled(from, Style::default().fg(Color::Green).add_modifier(Modifier::BOLD)),
            ]));
            lines.push(Line::from(vec![
                Span::styled("Date: ", Style::default().fg(Color::DarkGray)),
                Span::styled(date.to_string(), Style::default().fg(Color::DarkGray)),
            ]));
            lines.push(Line::from(""));
        }

        // Email body
        let body_html = self.current_email_body.as_deref().unwrap_or("Loading...");
        let rich_lines = html::strip_html_rich(body_html);
        for rich_line in &rich_lines {
            let mut spans = html::rich_to_spans(rich_line);

            if !self.search_highlight.is_empty() {
                spans = apply_search_highlight(spans, &self.search_highlight);
            }

            let line_text: String = spans.iter().map(|s| s.content.as_ref()).collect();
            if line_text.len() > width && width > 4 {
                let style = if rich_line.iter().any(|s| s.bold) {
                    Style::default().add_modifier(Modifier::BOLD)
                } else {
                    Style::default()
                };
                for wrapped in wrap_text(&line_text, width) {
                    lines.push(Line::from(Span::styled(wrapped, style)));
                }
            } else {
                lines.push(Line::from(spans));
            }
        }

        lines
    }

    fn draw_search(&mut self, frame: &mut Frame) {
        let area = frame.area();
        let [input_area, results_area, status_area] = Layout::vertical([
            Constraint::Length(3),
            Constraint::Min(1),
            Constraint::Length(1),
        ])
        .areas(area);

        let input = Paragraph::new(self.search_query.as_str()).block(
            Block::default()
                .borders(Borders::ALL)
                .title(" Search ")
                .border_style(Style::default().fg(Color::Cyan)),
        );
        frame.render_widget(input, input_area);
        frame.set_cursor_position((
            input_area.x + 1 + self.search_query.width() as u16,
            input_area.y + 1,
        ));

        let mut items: Vec<ListItem> = Vec::new();

        for &idx in &self.search_results {
            if let Some(conv) = self.conversations.get(idx) {
                let kind = conv.kind();
                if matches!(kind, ConvKind::System) {
                    continue;
                }
                let prefix = match kind {
                    ConvKind::Channel => "# ",
                    ConvKind::Chat => "@ ",
                    ConvKind::Meeting => "M ",
                    ConvKind::System => continue,
                };
                let name = conv.display_name(&self.my_name);
                let read_local = self.read_locally.get(&conv.id).copied().unwrap_or(false);
                let unread = if read_local {
                    self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                } else {
                    conv.is_unread()
                        || self.has_new_messages.get(&conv.id).copied().unwrap_or(false)
                };
                let style = if unread {
                    Style::default()
                        .fg(Color::White)
                        .add_modifier(Modifier::BOLD)
                } else {
                    Style::default()
                };
                items.push(ListItem::new(
                    Line::from(format!(" {}{}", prefix, name)).style(style),
                ));
            }
        }

        // Email search results
        if !self.search_email_results.is_empty() {
            items.push(ListItem::new(
                Line::from(" Emails")
                    .style(Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)),
            ));
            for &idx in &self.search_email_results {
                if let Some(email) = self.emails.get(idx) {
                    let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("(no subject)");
                    let from = email.get("from")
                        .and_then(|v| v.get("emailAddress"))
                        .and_then(|v| v.get("name"))
                        .and_then(|v| v.as_str())
                        .unwrap_or("?");
                    items.push(ListItem::new(
                        Line::from(format!("   {} - {}", from, subject)),
                    ));
                }
            }
        }

        if !self.search_people_results.is_empty() {
            if !items.is_empty() {
                items.push(ListItem::new(""));
            }
            items.push(ListItem::new(
                Line::from(" People")
                    .style(Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)),
            ));
            for (name, email) in &self.search_people_results {
                let label = if email.is_empty() {
                    format!("   {}", name)
                } else {
                    format!("   {} ({})", name, email)
                };
                items.push(ListItem::new(label));
            }
        }

        let total = self.search_results.len() + self.search_people_results.len();
        let list = List::new(items)
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .title(format!(" {} results ", total)),
            )
            .highlight_style(
                Style::default()
                    .bg(Color::DarkGray)
                    .add_modifier(Modifier::BOLD),
            );
        frame.render_stateful_widget(list, results_area, &mut self.search_list_state);

        let status = Paragraph::new(
            Line::from(" Type to filter  Enter:open  Esc:cancel  j/k:nav ")
                .style(Style::default().fg(Color::DarkGray)),
        );
        frame.render_widget(status, status_area);
    }

    // --- Key handling ---

    async fn handle_key(&mut self, key: KeyCode, modifiers: KeyModifiers) {
        if self.search_active {
            self.handle_search_key(key).await;
            return;
        }

        match &self.focus {
            Focus::Sidebar => match key {
                KeyCode::Char('q') => self.exit = true,
                KeyCode::Char('j') | KeyCode::Down => {
                    self.sidebar_next();
                    self.search_highlight.clear();
                    self.preview_selected().await;
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    self.sidebar_prev();
                    self.search_highlight.clear();
                    self.preview_selected().await;
                }
                KeyCode::Char('r') => {
                    self.load_conversations().await;
                    self.load_email_folders().await;
                }
                KeyCode::Char('e') => self.load_email_folders().await,
                KeyCode::Char('/') => {
                    if let Some((path, line)) = self.spawn_tv() {
                        self.navigate_to_tv_selection(&path, line).await;
                    }
                }
                KeyCode::Right | KeyCode::Char('l') => {
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
                        for i in (start + 1)..self.sidebar_items.len() {
                            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                                self.sidebar_state.select(Some(i));
                                self.search_highlight.clear();
                                self.preview_selected().await;
                                break;
                            }
                        }
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
                        for i in (start + 1)..self.sidebar_items.len() {
                            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                                self.sidebar_state.select(Some(i));
                                self.search_highlight.clear();
                                self.preview_selected().await;
                                break;
                            }
                        }
                    }
                }
                KeyCode::Char('f') => {
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
                            if let Ok(store) = crate::store::Store::new(self.auth.config_dir()) {
                                let _ = store.save_favourites(&self.favourites);
                            }
                            self.rebuild_sidebar();
                        }
                    }
                }
                KeyCode::Char('i') => {
                    if self.current_conv_id.is_some() {
                        self.focus = Focus::Input;
                    }
                }
                KeyCode::Enter => {
                    self.open_conversation().await;
                    self.focus = Focus::Messages;
                }
                _ => {}
            },
            Focus::Messages => match key {
                KeyCode::Esc => {
                    if !self.search_highlight.is_empty() {
                        self.search_highlight.clear();
                    } else {
                        self.focus = Focus::Sidebar;
                    }
                }
                KeyCode::Left | KeyCode::Char('h') => {
                    self.focus = Focus::Sidebar;
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
                    // Jump to bottom
                    self.scroll_offset = self.rendered_line_count;
                    self.has_new_below = false;
                }
                KeyCode::Char('g') => {
                    self.scroll_offset = 0;
                }
                KeyCode::Char('r') => self.load_messages().await,
                KeyCode::Char('/') => {
                    if let Some((path, line)) = self.spawn_tv() {
                        self.navigate_to_tv_selection(&path, line).await;
                    }
                }
                _ => {}
            },
            Focus::Input => match key {
                KeyCode::Esc => {
                    self.focus = Focus::Messages;
                }
                KeyCode::Enter => {
                    self.send_message().await;
                }
                KeyCode::Backspace => {
                    self.input_buffer.pop();
                }
                KeyCode::Char(c) => {
                    self.input_buffer.push(c);
                }
                _ => {}
            },
            Focus::Search => {
                self.handle_search_key(key).await;
            }
        }
    }

    async fn handle_mouse(&mut self, mouse: crossterm::event::MouseEvent) {
        let col = mouse.column;
        let row = mouse.row;

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
                    self.scroll_offset = self.scroll_offset.saturating_add(3);
                } else if self.in_area(col, row, self.sidebar_area) {
                    self.sidebar_next();
                }
            }
            MouseEventKind::Down(MouseButton::Left) => {
                if self.search_active {
                    return;
                }
                if self.in_area(col, row, self.sidebar_area) {
                    self.focus = Focus::Sidebar;
                    // Calculate which sidebar item was clicked
                    // sidebar list has 1-line border top + 1-line status bottom
                    let visible_row = (row - self.sidebar_area.y).saturating_sub(1) as usize;
                    let list_offset = self.sidebar_state.offset();
                    let item_idx = list_offset + visible_row;
                    if item_idx < self.sidebar_items.len() {
                        if matches!(self.sidebar_items[item_idx], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                            self.sidebar_state.select(Some(item_idx));
                            self.search_highlight.clear();
                            self.preview_selected().await;
                        }
                    }
                } else if self.in_area(col, row, self.msg_area) {
                    self.focus = Focus::Messages;
                }
            }
            _ => {}
        }
    }

    fn in_area(&self, col: u16, row: u16, area: Rect) -> bool {
        col >= area.x && col < area.x + area.width && row >= area.y && row < area.y + area.height
    }

    async fn handle_search_key(&mut self, key: KeyCode) {
        match key {
            KeyCode::Esc => {
                self.close_search();
                self.focus = Focus::Sidebar;
            }
            KeyCode::Enter => {
                self.open_search_result().await;
            }
            KeyCode::Down | KeyCode::Tab => {
                let total = self.search_total_items();
                if total > 0 {
                    let i = self
                        .search_list_state
                        .selected()
                        .map(|i| (i + 1).min(total - 1))
                        .unwrap_or(0);
                    self.search_list_state.select(Some(i));
                }
            }
            KeyCode::Up => {
                let i = self
                    .search_list_state
                    .selected()
                    .map(|i| i.saturating_sub(1))
                    .unwrap_or(0);
                self.search_list_state.select(Some(i));
            }
            KeyCode::Backspace => {
                self.search_query.pop();
                self.update_search_results();
                if self.search_query.len() >= 3 {
                    self.remote_search().await;
                } else {
                    self.search_people_results.clear();
                }
            }
            KeyCode::Char(c) => {
                self.search_query.push(c);
                self.update_search_results();
                if self.search_query.len() >= 3 {
                    self.remote_search().await;
                }
            }
            _ => {}
        }
    }

    // --- Sidebar navigation ---

    fn sidebar_next(&mut self) {
        let total = self.sidebar_items.len();
        if total == 0 { return; }
        let current = self.sidebar_state.selected().unwrap_or(0);
        // Find next selectable item (skip headers)
        for i in (current + 1)..total {
            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                self.sidebar_state.select(Some(i));
                return;
            }
        }
    }

    fn sidebar_prev(&mut self) {
        let current = self.sidebar_state.selected().unwrap_or(0);
        // Find prev selectable item (skip headers)
        for i in (0..current).rev() {
            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                self.sidebar_state.select(Some(i));
                return;
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

        // 1. Favourites
        if !fav_conv_indices.is_empty() || !fav_email_indices.is_empty() {
            section_starts.push(items.len());
            let count = fav_conv_indices.len() + fav_email_indices.len();
            items.push(SidebarItem::Header(format!("Favourites ({})", count)));
            for idx in &fav_conv_indices {
                items.push(SidebarItem::Conv(*idx));
            }
            for idx in &fav_email_indices {
                items.push(SidebarItem::Email(*idx));
            }
        }

        // 2. Activity - conversations with new messages (up to 10)
        let mut activity: Vec<usize> = self.conversations.iter().enumerate()
            .filter(|(_, c)| {
                self.has_new_messages.get(&c.id).copied().unwrap_or(false)
                    && !self.favourites.contains(&c.id)
            })
            .map(|(i, _)| i)
            .take(10)
            .collect();
        // Also include unread (not read locally) that aren't favourites
        if activity.len() < 10 {
            for (i, conv) in self.conversations.iter().enumerate() {
                if activity.len() >= 10 { break; }
                let read_local = self.read_locally.get(&conv.id).copied().unwrap_or(false);
                if !read_local && conv.is_unread() && !self.favourites.contains(&conv.id) && !activity.contains(&i) {
                    activity.push(i);
                }
            }
        }
        if !activity.is_empty() {
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!("Activity ({})", activity.len())));
            for idx in activity {
                items.push(SidebarItem::Conv(idx));
            }
        }

        // 3. DMs
        if !dms.is_empty() {
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!("Direct Messages ({})", dms.len())));
            for idx in dms {
                items.push(SidebarItem::Conv(idx));
            }
        }

        // 4. Channels
        if !channels.is_empty() {
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!("Channels ({})", channels.len())));
            for idx in channels {
                items.push(SidebarItem::Conv(idx));
            }
        }

        // 5. Meetings
        if !meetings.is_empty() {
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!("Meetings ({})", meetings.len())));
            for idx in meetings {
                items.push(SidebarItem::Conv(idx));
            }
        }

        // 6. Emails - grouped by folder
        if !self.email_folders.is_empty() {
            section_starts.push(items.len());
            items.push(SidebarItem::Header("Emails".to_string()));
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
            // Fallback: flat list if email_folders not loaded yet
            section_starts.push(items.len());
            items.push(SidebarItem::Header(format!("Emails ({})", self.emails.len())));
            for idx in 0..self.emails.len() {
                items.push(SidebarItem::Email(idx));
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
                    self.current_conv_id = Some(id);
                    self.current_conv_topic = topic;
                    self.messages.clear();
                    self.load_messages().await;
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

        self.rebuild_sidebar();
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
        match self.api.list_mail_folders(&mut self.auth).await {
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
                self.status = format!("Email folders failed: {} (press 'e' to retry)", e);
                self.email_loaded = true;
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
            }
            Err(e) => {
                self.status = format!("Error: {}", e);
            }
        }
    }

    /// Suspend TUI, spawn tv for search, return selected file path and line number
    fn spawn_tv(&self) -> Option<(String, Option<usize>)> {
        use std::process::Command;

        let data_dir = self.auth.config_dir().join("data");
        let data_path = data_dir.to_string_lossy();

        // Suspend TUI
        let _ = crossterm::execute!(
            std::io::stdout(),
            crossterm::event::DisableMouseCapture,
            crossterm::terminal::LeaveAlternateScreen,
        );
        let _ = crossterm::terminal::disable_raw_mode();

        // Spawn tv with rg as source
        let source_cmd = format!(
            "rg . --no-heading --line-number --color=never {}",
            data_path
        );
        let result = Command::new("tv")
            .arg("--source-command")
            .arg(&source_cmd)
            .arg("--preview-command")
            .arg("teamsh preview '{}'")
            .stdin(std::process::Stdio::inherit())
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::inherit())
            .output();

        // Resume TUI
        let _ = crossterm::terminal::enable_raw_mode();
        let _ = crossterm::execute!(
            std::io::stdout(),
            crossterm::terminal::EnterAlternateScreen,
            crossterm::event::EnableMouseCapture,
        );

        match result {
            Ok(output) => {
                let selection = String::from_utf8_lossy(&output.stdout).trim().to_string();
                if selection.is_empty() {
                    return None;
                }
                // Parse: path/to/file.txt:line_number:matched_text
                let parts: Vec<&str> = selection.splitn(3, ':').collect();
                let file_path = parts[0].to_string();
                let line_num = parts.get(1).and_then(|s| s.parse::<usize>().ok());
                Some((file_path, line_num))
            }
            Err(_e) => {
                // tv not found or failed
                None
            }
        }
    }

    async fn navigate_to_tv_selection(&mut self, file_path: &str, _line_num: Option<usize>) {
        let parts: Vec<&str> = file_path.split('/').collect();

        // Look for "conversations" in path
        if let Some(pos) = parts.iter().position(|&p| p == "conversations") {
            if let Some(conv_id_safe) = parts.get(pos + 1) {
                let conv_id_safe = *conv_id_safe;
                // Find conversation by matching safe_filename version of ID
                for (i, conv) in self.conversations.iter().enumerate() {
                    let safe = conv.id.replace('/', "_").replace('\\', "_").replace(':', "_")
                        .replace('?', "_").replace('*', "_").replace('"', "_")
                        .replace('<', "_").replace('>', "_").replace('|', "_");
                    if safe == conv_id_safe || conv.id == conv_id_safe {
                        let id = conv.id.clone();
                        let topic = conv.display_name(&self.my_name);
                        self.current_email_id = None;
                        self.current_email_body = None;
                        self.current_conv_id = Some(id.clone());
                        self.current_conv_topic = topic;
                        self.has_new_messages.insert(id.clone(), false);
                        self.read_locally.insert(id, true);
                        self.load_messages().await;
                        self.scroll_offset = usize::MAX;
                        self.focus = Focus::Messages;

                        // Select in sidebar
                        for (si, item) in self.sidebar_items.iter().enumerate() {
                            if let SidebarItem::Conv(idx) = item {
                                if *idx == i {
                                    self.sidebar_state.select(Some(si));
                                    break;
                                }
                            }
                        }
                        return;
                    }
                }
            }
        }

        // Look for "emails" in path
        if let Some(pos) = parts.iter().position(|&p| p == "emails") {
            if parts.get(pos + 2).is_some() {
                let email_file = parts[pos + 2];
                let email_id_safe = email_file.trim_end_matches(".txt");
                for (i, email) in self.emails.iter().enumerate() {
                    let eid = email.get("id").and_then(|v| v.as_str()).unwrap_or("");
                    let safe = eid.replace('/', "_").replace('\\', "_").replace(':', "_")
                        .replace('?', "_").replace('*', "_").replace('"', "_")
                        .replace('<', "_").replace('>', "_").replace('|', "_");
                    if safe == email_id_safe || eid == email_id_safe {
                        self.preview_email(i).await;
                        self.focus = Focus::Messages;
                        return;
                    }
                }
            }
        }

        self.status = "Could not find item from tv selection".to_string();
    }

    async fn load_messages(&mut self) {
        if let Some(conv_id) = &self.current_conv_id.clone() {
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
        if let Some(conv_id) = &self.current_conv_id.clone() {
            match self.api.send_message(&mut self.auth, conv_id, &text).await {
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

    // --- Search ---

    fn update_search_results(&mut self) {
        use nucleo_matcher::{Matcher, Config, Utf32Str};
        use nucleo_matcher::pattern::{Pattern, CaseMatching, Normalization, AtomKind};

        let query = &self.search_query;

        if query.is_empty() {
            // Show all non-system conversations and all emails
            self.search_results = self.conversations.iter().enumerate()
                .filter(|(_, conv)| !matches!(conv.kind(), ConvKind::System))
                .map(|(i, _)| i)
                .collect();
            self.search_email_results = (0..self.emails.len()).collect();
        } else {
            let mut matcher = Matcher::new(Config::DEFAULT);
            let pattern = Pattern::new(query, CaseMatching::Smart, Normalization::Smart, AtomKind::Fuzzy);
            let mut buf = Vec::new();

            // Fuzzy match conversations - score each individually
            let mut conv_scored: Vec<(usize, u32)> = Vec::new();
            for (i, conv) in self.conversations.iter().enumerate() {
                if matches!(conv.kind(), ConvKind::System) { continue; }
                let mut text = conv.display_name(&self.my_name);
                let topic = conv.topic();
                if topic != "(no topic)" {
                    text = format!("{} {}", text, topic);
                }
                for member in &conv.member_names {
                    text = format!("{} {}", text, member);
                }
                if let Some(snippets) = self.cached_snippets.get(&conv.id) {
                    for s in snippets.iter().take(3) {
                        text = format!("{} {}", text, s);
                    }
                }
                let haystack = Utf32Str::new(&text, &mut buf);
                if let Some(score) = pattern.score(haystack, &mut matcher) {
                    conv_scored.push((i, score));
                }
            }
            conv_scored.sort_by(|a, b| b.1.cmp(&a.1));
            self.search_results = conv_scored.into_iter().map(|(i, _)| i).collect();

            // Fuzzy match emails
            let mut email_scored: Vec<(usize, u32)> = Vec::new();
            for (i, email) in self.emails.iter().enumerate() {
                let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("");
                let from = email.get("from")
                    .and_then(|v| v.get("emailAddress"))
                    .and_then(|v| v.get("name"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                let preview = email.get("bodyPreview").and_then(|v| v.as_str()).unwrap_or("");
                let text = format!("{} {} {}", subject, from, preview);
                let haystack = Utf32Str::new(&text, &mut buf);
                if let Some(score) = pattern.score(haystack, &mut matcher) {
                    email_scored.push((i, score));
                }
            }
            email_scored.sort_by(|a, b| b.1.cmp(&a.1));
            self.search_email_results = email_scored.into_iter().map(|(i, _)| i).collect();
        }

        let total = self.search_results.len() + self.search_email_results.len();
        if total > 0 {
            self.search_list_state.select(Some(0));
        } else {
            self.search_list_state.select(None);
        }
    }

    fn search_total_items(&self) -> usize {
        let mut count = self.search_results.len();
        if !self.search_email_results.is_empty() {
            count += 1 + self.search_email_results.len(); // +1 for header
        }
        if !self.search_people_results.is_empty() {
            count += 2 + self.search_people_results.len();
        }
        count
    }

    async fn remote_search(&mut self) {
        if let Ok(results) = self
            .api
            .search_people(&mut self.auth, &self.search_query)
            .await
        {
            self.search_people_results = results;
        }
    }

    async fn open_search_result(&mut self) {
        let selected_idx = self.search_list_state.selected();
        if selected_idx.is_none() { return; }
        let sel = selected_idx.unwrap();

        // Determine what was selected: conversation, email header, email, or people
        let conv_count = self.search_results.len();
        let _email_header_offset = conv_count;
        let email_start = if self.search_email_results.is_empty() { conv_count } else { conv_count + 1 };
        let email_end = email_start + self.search_email_results.len();

        if sel < conv_count {
            // Conversation result
            if let Some(&conv_idx) = self.search_results.get(sel) {
                if let Some(conv) = self.conversations.get(conv_idx) {
                    let id = conv.id.clone();
                    let topic = conv.display_name(&self.my_name);
                    self.has_new_messages.insert(id.clone(), false);
                    self.read_locally.insert(id.clone(), true);
                    self.current_email_id = None;
                    self.current_email_body = None;
                    self.current_conv_id = Some(id);
                    self.current_conv_topic = topic;
                    self.search_highlight = self.search_query.clone();
                    self.close_search();
                    self.focus = Focus::Messages;
                    self.load_messages().await;
                    self.scroll_offset = usize::MAX;
                }
            }
        } else if sel >= email_start && sel < email_end {
            // Email result
            let email_idx_pos = sel - email_start;
            if let Some(&email_idx) = self.search_email_results.get(email_idx_pos) {
                self.search_highlight = self.search_query.clone();
                self.close_search();
                self.focus = Focus::Messages;
                self.preview_email(email_idx).await;
            }
        }
        // Headers and people results are not openable
    }

    fn close_search(&mut self) {
        self.search_active = false;
        self.search_query.clear();
        self.search_results.clear();
        self.search_email_results.clear();
        self.search_people_results.clear();
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
            Err(_) => {}
        }
    }
}

/// Split spans to highlight search matches with a bright background.
fn apply_search_highlight(spans: Vec<Span<'static>>, query: &str) -> Vec<Span<'static>> {
    let query_lower = query.to_lowercase();
    let query_lower_len = query_lower.len();
    let mut result = Vec::new();
    for span in spans {
        let text = span.content.to_string();
        let text_lower = text.to_lowercase();
        if !text_lower.contains(&query_lower) {
            result.push(span);
            continue;
        }
        // Map byte positions in lowercase back to original text
        // For ASCII this is 1:1, for Unicode we iterate both in sync
        let base_style = span.style;
        let hl_style = base_style.bg(Color::Yellow).fg(Color::Black);
        let mut pos = 0; // byte position in text_lower
        loop {
            if let Some(idx) = text_lower[pos..].find(&query_lower) {
                let match_start = pos + idx;
                let match_end = match_start + query_lower_len;
                // Find corresponding byte positions in original text
                let orig_start = byte_pos_in_original(&text, match_start);
                let orig_end = byte_pos_in_original(&text, match_end);
                let before_start = byte_pos_in_original(&text, pos);
                if orig_start > before_start {
                    result.push(Span::styled(text[before_start..orig_start].to_string(), base_style));
                }
                result.push(Span::styled(text[orig_start..orig_end].to_string(), hl_style));
                pos = match_end;
            } else {
                let orig_pos = byte_pos_in_original(&text, pos);
                if orig_pos < text.len() {
                    result.push(Span::styled(text[orig_pos..].to_string(), base_style));
                }
                break;
            }
        }
    }
    result
}

/// Map a byte position in the lowercased string to the corresponding byte position
/// in the original string. For ASCII text (which covers most Teams messages) these are identical.
fn byte_pos_in_original(original: &str, lower_pos: usize) -> usize {
    lower_pos.min(original.len())
}

/// Wrap a text string to fit within `width` columns, respecting Unicode width.
fn wrap_text(text: &str, width: usize) -> Vec<String> {
    if width == 0 {
        return vec![text.to_string()];
    }
    let mut lines = Vec::new();
    let mut current = String::new();
    let mut current_width = 0;

    for ch in text.chars() {
        let ch_width = if ch == '\t' { 4 } else { UnicodeWidthStr::width(ch.encode_utf8(&mut [0; 4])) };
        if current_width + ch_width > width && current_width > 0 {
            lines.push(current.clone());
            current.clear();
            current_width = 0;
        }
        current.push(ch);
        current_width += ch_width;
    }
    if !current.is_empty() || lines.is_empty() {
        lines.push(current);
    }
    lines
}
