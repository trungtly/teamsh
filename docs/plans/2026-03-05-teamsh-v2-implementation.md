# teamsh v2 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Redesign teamsh with local file storage, tv integration, sectioned sidebar, and smooth scrolling.

**Architecture:** Keep existing ratatui TUI for viewing/replying. Add a `store` module for local file-based storage. Add CLI subcommands (`list`, `preview`, `sync`) for tv integration. `/` key shells out to tv subprocess. Sidebar refactored into sections (Favourites, Activity, DMs, Channels, Meetings, Emails) with Tab navigation.

**Tech Stack:** Rust, ratatui, crossterm, tv (external binary), rg (external binary), Microsoft Graph API

---

### Task 1: Create the `store` module -- data structures and directory setup

**Files:**
- Create: `src/store.rs`
- Modify: `src/main.rs:1` (add `mod store;`)

**Step 1: Create `src/store.rs` with types and directory init**

```rust
use anyhow::Result;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

use crate::html;

/// Index entry for a conversation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConvIndex {
    pub id: String,
    pub name: String,
    pub kind: String, // "Chat", "Channel", "Meeting"
    pub last_activity: u64,
    pub unread: bool,
}

/// Index entry for an email folder
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EmailFolderIndex {
    pub name: String,
    pub id: String,
    pub count: usize,
}

/// Full index stored at data/index.json
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct Index {
    pub my_name: String,
    pub conversations: Vec<ConvIndex>,
    pub email_folders: Vec<EmailFolderIndex>,
}

/// Per-conversation metadata stored at conversations/{id}/meta.json
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConvMeta {
    pub name: String,
    pub kind: String,
    pub members: Vec<String>,
    pub unread: bool,
    pub version: u64,
    pub last_message_id: Option<String>,
    pub consumptionhorizon: Option<String>,
}

/// Local file store
pub struct Store {
    data_dir: PathBuf,
}

impl Store {
    pub fn new(config_dir: &Path) -> Result<Self> {
        let data_dir = config_dir.join("data");
        fs::create_dir_all(data_dir.join("conversations"))?;
        fs::create_dir_all(data_dir.join("emails"))?;
        Ok(Self { data_dir })
    }

    pub fn data_dir(&self) -> &Path {
        &self.data_dir
    }

    // --- Index ---

    pub fn load_index(&self) -> Index {
        let path = self.data_dir.join("index.json");
        fs::read_to_string(&path)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_default()
    }

    pub fn save_index(&self, index: &Index) -> Result<()> {
        let path = self.data_dir.join("index.json");
        let json = serde_json::to_string_pretty(index)?;
        fs::write(&path, json)?;
        Ok(())
    }

    // --- Favourites ---

    pub fn load_favourites(&self) -> Vec<String> {
        let path = self.data_dir.join("favourites.json");
        fs::read_to_string(&path)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_default()
    }

    pub fn save_favourites(&self, favs: &[String]) -> Result<()> {
        let path = self.data_dir.join("favourites.json");
        let json = serde_json::to_string_pretty(favs)?;
        fs::write(&path, json)?;
        Ok(())
    }

    // --- Conversation metadata ---

    pub fn save_conv_meta(&self, conv_id: &str, meta: &ConvMeta) -> Result<()> {
        let dir = self.data_dir.join("conversations").join(safe_filename(conv_id));
        fs::create_dir_all(&dir)?;
        let json = serde_json::to_string_pretty(meta)?;
        fs::write(dir.join("meta.json"), json)?;
        Ok(())
    }

    pub fn load_conv_meta(&self, conv_id: &str) -> Option<ConvMeta> {
        let path = self.data_dir.join("conversations").join(safe_filename(conv_id)).join("meta.json");
        fs::read_to_string(&path)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
    }

    // --- Messages ---

    pub fn save_message(&self, conv_id: &str, msg_id: &str, timestamp: &str, sender: &str, content_html: &str) -> Result<()> {
        let dir = self.data_dir.join("conversations").join(safe_filename(conv_id)).join("messages");
        fs::create_dir_all(&dir)?;
        let time = timestamp.get(11..16).unwrap_or("??:??");
        let text = html::strip_html(content_html);
        let body = format!("{} {}\n  {}\n", time, sender, text);
        fs::write(dir.join(format!("{}.txt", safe_filename(msg_id))), body)?;
        Ok(())
    }

    pub fn load_messages(&self, conv_id: &str) -> Vec<(String, String)> {
        let dir = self.data_dir.join("conversations").join(safe_filename(conv_id)).join("messages");
        let mut entries: Vec<(String, String)> = Vec::new();
        if let Ok(read_dir) = fs::read_dir(&dir) {
            for entry in read_dir.flatten() {
                let path = entry.path();
                if path.extension().and_then(|e| e.to_str()) == Some("txt") {
                    let name = path.file_stem().unwrap_or_default().to_string_lossy().to_string();
                    let content = fs::read_to_string(&path).unwrap_or_default();
                    entries.push((name, content));
                }
            }
        }
        entries.sort_by(|a, b| a.0.cmp(&b.0));
        entries
    }

    // --- Emails ---

    pub fn save_email(&self, folder_name: &str, email_id: &str, from: &str, date: &str, subject: &str, body_html: &str) -> Result<()> {
        let dir = self.data_dir.join("emails").join(safe_filename(folder_name));
        fs::create_dir_all(&dir)?;
        let body_text = html::strip_html(body_html);
        let content = format!("From: {}\nDate: {}\nSubject: {}\n\n{}\n", from, date, subject, body_text);
        fs::write(dir.join(format!("{}.txt", safe_filename(email_id))), content)?;
        Ok(())
    }

    pub fn list_email_files(&self, folder_name: &str) -> Vec<PathBuf> {
        let dir = self.data_dir.join("emails").join(safe_filename(folder_name));
        let mut files = Vec::new();
        if let Ok(read_dir) = fs::read_dir(&dir) {
            for entry in read_dir.flatten() {
                let path = entry.path();
                if path.extension().and_then(|e| e.to_str()) == Some("txt") {
                    files.push(path);
                }
            }
        }
        files.sort();
        files
    }
}

/// Sanitize an ID for use as a filename (replace problematic chars)
fn safe_filename(id: &str) -> String {
    id.replace('/', "_")
        .replace('\\', "_")
        .replace(':', "_")
        .replace('?', "_")
        .replace('*', "_")
        .replace('"', "_")
        .replace('<', "_")
        .replace('>', "_")
        .replace('|', "_")
}
```

**Step 2: Add `mod store;` to `src/main.rs`**

Add `mod store;` after `mod cache;` on line 3.

**Step 3: Build and verify**

Run: `cargo build 2>&1 | grep error`
Expected: No errors (warnings OK)

**Step 4: Commit**

```bash
git add src/store.rs src/main.rs
git commit -m "Add store module for local file storage"
```

---

### Task 2: Add `list_mail_folders` API method

**Files:**
- Modify: `src/api.rs` (add `list_mail_folders` method after `list_emails`)

**Step 1: Add the method**

Add after `list_emails` method (~line 233):

```rust
    /// List all mail folders
    pub async fn list_mail_folders(&self, auth: &mut Auth) -> Result<Vec<serde_json::Value>> {
        let token = auth.graph_token().await?;
        let url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=50";
        let resp = self.client
            .get(url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await?;
        let status = resp.status();
        let text = resp.text().await?;
        if !status.is_success() {
            anyhow::bail!("List mail folders failed ({}): {}", status, &text[..text.len().min(300)]);
        }
        let data: serde_json::Value = serde_json::from_str(&text)?;
        let folders = data.get("value")
            .and_then(|v| v.as_array())
            .cloned()
            .unwrap_or_default();
        Ok(folders)
    }
```

**Step 2: Build and verify**

Run: `cargo build 2>&1 | grep error`
Expected: No errors

**Step 3: Commit**

```bash
git add src/api.rs
git commit -m "Add list_mail_folders API method"
```

---

### Task 3: Add `teamsh sync` CLI subcommand

This is the core sync logic: fetch conversations + messages + email folders from API, write to local files.

**Files:**
- Modify: `src/cli.rs` (add `Sync` command)
- Modify: `src/main.rs` (add `cmd_sync` function and wire it up)

**Step 1: Add `Sync` to CLI**

In `src/cli.rs`, add to the `Commands` enum after `Tui`:

```rust
    /// Sync conversations and emails to local files
    Sync,
```

**Step 2: Add `cmd_sync` to `src/main.rs`**

Add the match arm in `main()`:
```rust
        Some(Commands::Sync) => {
            cmd_sync().await?;
        }
```

Add the function:

```rust
async fn cmd_sync() -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    let store = store::Store::new(auth.config_dir())?;

    // Sync conversations
    eprintln!("Syncing conversations...");
    let resp = api.list_conversations(&mut auth, 100).await?;
    let mut convs = resp.conversations;
    convs.sort_by(|a, b| b.version.unwrap_or(0).cmp(&a.version.unwrap_or(0)));

    // Detect own name from message senders
    let mut name_counts: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    let mut my_name = String::new();

    let mut conv_index = Vec::new();
    let mut synced = 0;
    for conv in &convs {
        let kind = conv.kind();
        if matches!(kind, types::ConvKind::System) {
            continue;
        }

        let kind_str = format!("{:?}", kind);
        let display = conv.display_name(&my_name);

        // Save metadata
        let meta = store::ConvMeta {
            name: display.clone(),
            kind: kind_str.clone(),
            members: conv.member_names.clone(),
            unread: conv.is_unread(),
            version: conv.version.unwrap_or(0),
            last_message_id: conv.last_message.as_ref().and_then(|lm| lm.id.clone()),
            consumptionhorizon: conv.properties.as_ref().and_then(|p| p.consumptionhorizon.clone()),
        };
        let _ = store.save_conv_meta(&conv.id, &meta);

        // Fetch and save messages
        if let Ok(msg_resp) = api.get_messages(&mut auth, &conv.id, 30).await {
            for m in &msg_resp.messages {
                if !matches!(m.messagetype.as_deref(), Some("RichText/Html") | Some("Text")) {
                    continue;
                }
                let msg_id = m.id.as_deref().unwrap_or("unknown");
                let sender = m.imdisplayname.as_deref().unwrap_or("?");
                let timestamp = m.timestamp.as_deref().unwrap_or("");
                let content = m.content.as_deref().unwrap_or("");

                if !sender.is_empty() && sender != "?" {
                    *name_counts.entry(sender.to_string()).or_insert(0) += 1;
                }

                let _ = store.save_message(&conv.id, msg_id, timestamp, sender, content);
            }
        }

        conv_index.push(store::ConvIndex {
            id: conv.id.clone(),
            name: display,
            kind: kind_str,
            last_activity: conv.version.unwrap_or(0),
            unread: conv.is_unread(),
        });

        synced += 1;
    }

    // Detect own name (most frequent sender)
    if let Some((name, _)) = name_counts.iter().max_by_key(|(_, c)| *c) {
        my_name = name.clone();
    }

    // Sync email folders
    eprintln!("Syncing email folders...");
    let mut folder_index = Vec::new();
    match api.list_mail_folders(&mut auth).await {
        Ok(folders) => {
            for folder in &folders {
                let folder_name = folder.get("displayName").and_then(|v| v.as_str()).unwrap_or("Unknown");
                let folder_id = folder.get("id").and_then(|v| v.as_str()).unwrap_or("");
                let total = folder.get("totalItemCount").and_then(|v| v.as_u64()).unwrap_or(0) as usize;

                if total == 0 {
                    continue;
                }

                // Fetch emails for this folder (up to 25)
                match api.list_emails(&mut auth, folder_id, 25).await {
                    Ok(emails) => {
                        for email in &emails {
                            let email_id = email.get("id").and_then(|v| v.as_str()).unwrap_or("");
                            let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("(no subject)");
                            let from = email.get("from")
                                .and_then(|v| v.get("emailAddress"))
                                .map(|ea| {
                                    let name = ea.get("name").and_then(|v| v.as_str()).unwrap_or("?");
                                    let addr = ea.get("address").and_then(|v| v.as_str()).unwrap_or("");
                                    format!("{} <{}>", name, addr)
                                })
                                .unwrap_or_else(|| "?".to_string());
                            let date = email.get("receivedDateTime").and_then(|v| v.as_str()).unwrap_or("");
                            let preview = email.get("bodyPreview").and_then(|v| v.as_str()).unwrap_or("");

                            let _ = store.save_email(folder_name, email_id, &from, date, subject, preview);
                        }

                        folder_index.push(store::EmailFolderIndex {
                            name: folder_name.to_string(),
                            id: folder_id.to_string(),
                            count: emails.len(),
                        });
                    }
                    Err(e) => eprintln!("  Error fetching {}: {}", folder_name, e),
                }
            }
        }
        Err(e) => eprintln!("Error listing folders: {}", e),
    }

    // Save index
    let index = store::Index {
        my_name,
        conversations: conv_index,
        email_folders: folder_index,
    };
    store.save_index(&index)?;

    eprintln!("Synced {} conversations, {} email folders", synced, index.email_folders.len());
    Ok(())
}
```

**Step 3: Build and test**

Run: `cargo build 2>&1 | grep error`
Expected: No errors

Run: `cargo run -- sync 2>&1`
Expected: Output like "Syncing conversations... Syncing email folders... Synced N conversations, M email folders"

Verify files exist:
Run: `ls ~/.config/teamsh/data/ && ls ~/.config/teamsh/data/conversations/ | head -5`

**Step 4: Commit**

```bash
git add src/cli.rs src/main.rs
git commit -m "Add teamsh sync command for local file storage"
```

---

### Task 4: Add `teamsh list` and `teamsh preview` CLI subcommands

These are what tv calls to get data.

**Files:**
- Modify: `src/cli.rs` (add `List` and `Preview` commands)
- Modify: `src/main.rs` (add `cmd_list` and `cmd_preview` functions)

**Step 1: Add to CLI**

In `src/cli.rs`, add to `Commands` enum:

```rust
    /// List conversations and emails (for tv integration)
    List,
    /// Preview a conversation or email (for tv integration)
    Preview {
        /// File path from tv selection (e.g. conversations/{id}/messages/{msg_id}.txt:14:text)
        path: String,
    },
```

**Step 2: Add `cmd_list` to `src/main.rs`**

```rust
async fn cmd_list() -> Result<()> {
    let auth = auth::Auth::new()?;
    let store = store::Store::new(auth.config_dir())?;
    let index = store.load_index();

    for conv in &index.conversations {
        let prefix = match conv.kind.as_str() {
            "Channel" => "#",
            "Chat" => "@",
            "Meeting" => "M",
            _ => " ",
        };
        let marker = if conv.unread { "*" } else { " " };
        println!("{}{} {}\t{}", marker, prefix, conv.name, conv.id);
    }

    for folder in &index.email_folders {
        for file in store.list_email_files(&folder.name) {
            if let Some(first_lines) = std::fs::read_to_string(&file).ok() {
                let subject = first_lines.lines()
                    .find(|l| l.starts_with("Subject: "))
                    .map(|l| &l[9..])
                    .unwrap_or("(no subject)");
                let from = first_lines.lines()
                    .find(|l| l.starts_with("From: "))
                    .map(|l| &l[6..])
                    .unwrap_or("?");
                println!("  {} - {} [{}]", from, subject, file.display());
            }
        }
    }

    Ok(())
}
```

**Step 3: Add `cmd_preview` to `src/main.rs`**

```rust
async fn cmd_preview(path: &str) -> Result<()> {
    // tv outputs: path/to/file.txt:line_number:matched_text
    // Strip the line number and matched text, keep just the file path
    let file_path = path.split(':').next().unwrap_or(path);

    let auth = auth::Auth::new()?;
    let data_dir = auth.config_dir().join("data");

    // Try as absolute path first, then relative to data dir
    let full_path = if std::path::Path::new(file_path).exists() {
        PathBuf::from(file_path)
    } else {
        data_dir.join(file_path)
    };

    if full_path.exists() {
        let content = std::fs::read_to_string(&full_path)?;
        print!("{}", content);
    } else {
        eprintln!("File not found: {}", full_path.display());
    }

    Ok(())
}
```

**Step 4: Wire up in main()**

```rust
        Some(Commands::List) => {
            cmd_list().await?;
        }
        Some(Commands::Preview { path }) => {
            cmd_preview(path).await?;
        }
```

**Step 5: Build and test**

Run: `cargo build && cargo run -- list | head -10`
Expected: Lines like `*@ Alice, Bob	conv_id_here`

Run: `cargo run -- preview conversations/ 2>&1 | head -5`
(use an actual path from the data dir)

**Step 6: Commit**

```bash
git add src/cli.rs src/main.rs
git commit -m "Add list and preview CLI commands for tv integration"
```

---

### Task 5: Add tv subprocess spawning from TUI

**Files:**
- Modify: `src/tui/app.rs` (add `spawn_tv` method, modify `/` key handler)
- Modify: `src/tui/mod.rs` (pass terminal ownership for suspend/resume)

**Step 1: Add `spawn_tv` method to App in `src/tui/app.rs`**

Add a new method to the `impl App` block:

```rust
    /// Suspend TUI, spawn tv for search, return the selected file path and line
    fn spawn_tv(&self) -> Option<(String, Option<usize>)> {
        use std::process::Command;

        let data_dir = self.auth.config_dir().join("data");
        let data_path = data_dir.to_string_lossy();

        // Exit alternate screen
        let _ = crossterm::execute!(
            std::io::stdout(),
            crossterm::event::DisableMouseCapture,
            crossterm::terminal::LeaveAlternateScreen,
        );
        let _ = crossterm::terminal::disable_raw_mode();

        // Spawn tv with rg over data dir
        let source_cmd = format!(
            "rg . --no-heading --line-number --color=never {}",
            data_path
        );
        let result = Command::new("tv")
            .arg("--source-command")
            .arg(&source_cmd)
            .arg("--preview-command")
            .arg(format!("teamsh preview '{{}}'"))
            .stdin(std::process::Stdio::inherit())
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::inherit())
            .output();

        // Re-enter alternate screen
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
            Err(_) => None,
        }
    }
```

**Step 2: Add method to parse tv selection and navigate**

```rust
    /// Parse a tv selection path and navigate to the conversation/email
    async fn navigate_to_tv_selection(&mut self, file_path: &str, _line_num: Option<usize>) {
        // Extract conv_id from path like: .../conversations/{conv_id}/messages/{msg_id}.txt
        // Or email from: .../emails/{folder}/{email_id}.txt
        let parts: Vec<&str> = file_path.split('/').collect();

        // Find "conversations" or "emails" in path
        if let Some(conv_pos) = parts.iter().position(|&p| p == "conversations") {
            if let Some(conv_id_encoded) = parts.get(conv_pos + 1) {
                // Try to find the conversation by matching the safe_filename version of its ID
                let conv_id_encoded = *conv_id_encoded;
                for (i, conv) in self.conversations.iter().enumerate() {
                    let safe = conv.id.replace('/', "_").replace('\\', "_").replace(':', "_");
                    if safe == conv_id_encoded || conv.id == conv_id_encoded {
                        let id = conv.id.clone();
                        let topic = conv.display_name(&self.my_name);
                        self.current_email_id = None;
                        self.current_email_body = None;
                        self.current_conv_id = Some(id.clone());
                        self.current_conv_topic = topic;
                        self.has_new_messages.insert(id.clone(), false);
                        self.read_locally.insert(id, true);
                        self.load_messages().await;
                        // TODO: scroll to specific message if msg_id available
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
        } else if let Some(email_pos) = parts.iter().position(|&p| p == "emails") {
            // Email: navigate to email view
            if let Some(_folder) = parts.get(email_pos + 1) {
                if let Some(email_file) = parts.get(email_pos + 2) {
                    let email_id_approx = email_file.trim_end_matches(".txt");
                    for (i, email) in self.emails.iter().enumerate() {
                        let eid = email.get("id").and_then(|v| v.as_str()).unwrap_or("");
                        let safe = eid.replace('/', "_").replace('\\', "_").replace(':', "_");
                        if safe == email_id_approx || eid == email_id_approx {
                            self.preview_email(i).await;
                            self.focus = Focus::Messages;
                            return;
                        }
                    }
                }
            }
        }

        self.status = "Could not find item from tv selection".to_string();
    }
```

**Step 3: Replace `/` key handler in sidebar and messages focus**

Replace the `/` key handler in the `Focus::Sidebar` match arm:

```rust
                KeyCode::Char('/') => {
                    if let Some((path, line)) = self.spawn_tv() {
                        self.navigate_to_tv_selection(&path, line).await;
                    }
                }
```

Do the same for the `/` key handler in the `Focus::Messages` match arm.

**Step 4: Build and test**

Run: `cargo build 2>&1 | grep error`
Expected: No errors

Manual test: run `cargo run`, press `/`, verify tv opens with rg results, select one, verify TUI resumes.

**Step 5: Commit**

```bash
git add src/tui/app.rs
git commit -m "Add tv integration: / key spawns tv subprocess"
```

---

### Task 6: Refactor sidebar into sections with Tab navigation

**Files:**
- Modify: `src/tui/app.rs` (refactor sidebar sections, add Tab cycling, add Favourites/Activity)

**Step 1: Add section enum and tracking fields**

Add to App struct:

```rust
    // Sidebar sections
    current_section: usize, // index into SECTION_ORDER
    favourites: Vec<String>, // favourite conv/email IDs
```

Add constants at top of file:

```rust
const SECTIONS: &[&str] = &["Favourites", "Activity", "DMs", "Channels", "Meetings", "Emails"];
```

Initialize in `App::new()`:

```rust
            current_section: 2, // Start at DMs
            favourites: Vec::new(),
```

Load favourites from store in `App::new()` after store init:

```rust
        let store = store::Store::new(app.auth.config_dir())?;
        app.favourites = store.load_favourites();
```

**Step 2: Update `rebuild_sidebar` to include Favourites and Activity sections**

Rewrite `rebuild_sidebar` to:
1. Build Favourites section from `self.favourites` IDs
2. Build Activity section from conversations with `has_new_messages` or recently active
3. Build DMs, Channels, Meetings, Emails sections as before
4. Store section start indices for Tab navigation

Add a field to track section positions:

```rust
    section_starts: Vec<usize>, // sidebar_items index where each section header is
```

In rebuild, record the index of each Header item pushed to `section_starts`.

**Step 3: Add Tab/Shift+Tab handling**

In `Focus::Sidebar` key handler, add:

```rust
                KeyCode::Tab => {
                    // Jump to next section
                    if !self.section_starts.is_empty() {
                        self.current_section = (self.current_section + 1) % self.section_starts.len();
                        let start = self.section_starts[self.current_section];
                        // Select first selectable item after the header
                        for i in (start + 1)..self.sidebar_items.len() {
                            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                                self.sidebar_state.select(Some(i));
                                self.preview_selected().await;
                                break;
                            }
                        }
                    }
                }
                KeyCode::BackTab => {
                    // Jump to previous section
                    if !self.section_starts.is_empty() {
                        if self.current_section == 0 {
                            self.current_section = self.section_starts.len() - 1;
                        } else {
                            self.current_section -= 1;
                        }
                        let start = self.section_starts[self.current_section];
                        for i in (start + 1)..self.sidebar_items.len() {
                            if matches!(self.sidebar_items[i], SidebarItem::Conv(_) | SidebarItem::Email(_)) {
                                self.sidebar_state.select(Some(i));
                                self.preview_selected().await;
                                break;
                            }
                        }
                    }
                }
```

**Step 4: Add `f` key for favourites toggle**

In `Focus::Sidebar` key handler:

```rust
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
                            let store = store::Store::new(self.auth.config_dir()).ok();
                            if let Some(store) = store {
                                let _ = store.save_favourites(&self.favourites);
                            }
                            self.rebuild_sidebar();
                        }
                    }
                }
```

**Step 5: Change left/right arrow to move focus**

Replace the existing `Tab | Right` handler with:

```rust
                KeyCode::Right | KeyCode::Char('l') => {
                    if self.current_conv_id.is_some() || self.current_email_id.is_some() {
                        self.focus = Focus::Messages;
                    }
                }
```

In `Focus::Messages`, add:

```rust
                KeyCode::Left | KeyCode::Char('h') => {
                    self.focus = Focus::Sidebar;
                }
```

**Step 6: Build and test**

Run: `cargo build 2>&1 | grep error`

Manual test: run TUI, verify Tab cycles sections, `f` pins/unpins, left/right moves focus.

**Step 7: Commit**

```bash
git add src/tui/app.rs
git commit -m "Refactor sidebar: sections, Tab navigation, favourites"
```

---

### Task 7: Add email folder support

**Files:**
- Modify: `src/tui/app.rs` (fetch and display email folders instead of flat inbox)
- Modify: `src/api.rs` (already done in Task 2)

**Step 1: Add email folder state to App**

Add fields:

```rust
    email_folders: Vec<(String, String, Vec<serde_json::Value>)>, // (name, id, emails)
```

**Step 2: Replace `load_emails` with `load_email_folders`**

New method that calls `list_mail_folders`, then `list_emails` for each non-empty folder:

```rust
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
                self.rebuild_sidebar();
                self.status = format!("{} conversations, {} email folders", self.conversations.len(), self.email_folders.len());
            }
            Err(e) => {
                self.status = format!("Email folders failed: {}", e);
            }
        }
    }
```

**Step 3: Update `rebuild_sidebar` email section**

Replace the flat email section with folder-grouped display:

```rust
        // Emails - grouped by folder
        if !self.email_folders.is_empty() {
            items.push(SidebarItem::Header("Emails".to_string()));
            for (folder_name, _, folder_emails) in &self.email_folders {
                items.push(SidebarItem::Header(format!("  {} ({})", folder_name, folder_emails.len())));
                // Find indices in self.emails for this folder's emails
                for email in folder_emails {
                    let email_id = email.get("id").and_then(|v| v.as_str()).unwrap_or("");
                    if let Some(idx) = self.emails.iter().position(|e| {
                        e.get("id").and_then(|v| v.as_str()) == Some(email_id)
                    }) {
                        items.push(SidebarItem::Email(idx));
                    }
                }
            }
        }
```

**Step 4: Replace all `load_emails().await` calls with `load_email_folders().await`**

**Step 5: Build and test**

Run: `cargo build && cargo run`
Verify emails show grouped by folder.

**Step 6: Commit**

```bash
git add src/tui/app.rs
git commit -m "Show emails grouped by folder from Graph API"
```

---

### Task 8: Fix scrolling -- render cache

**Files:**
- Modify: `src/tui/app.rs`

**Step 1: Add render cache fields to App**

```rust
    // Render cache
    cached_rendered_lines: Vec<Line<'static>>,
    render_dirty: bool,
```

Initialize in `App::new()`:

```rust
            cached_rendered_lines: Vec::new(),
            render_dirty: true,
```

**Step 2: Set dirty flag when messages change**

In `load_messages`, after `self.messages = msgs;`, add:
```rust
self.render_dirty = true;
```

In `preview_email`, after setting `self.current_email_body`, add:
```rust
self.render_dirty = true;
```

On window resize (add to run loop after event handling):
```rust
Event::Resize(_, _) => {
    self.render_dirty = true;
}
```

**Step 3: Use cache in `draw_main`**

Replace the wrapped_lines computation:

```rust
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
```

**Step 4: Build and test**

Run: `cargo build`
Manual: verify scrolling feels smoother (no re-render on every frame).

**Step 5: Commit**

```bash
git add src/tui/app.rs
git commit -m "Cache rendered lines for smoother scrolling"
```

---

### Task 9: Fix scrolling -- auto-follow and new message indicator

**Files:**
- Modify: `src/tui/app.rs`

**Step 1: Add fields**

```rust
    has_new_below: bool, // true when new messages arrived while scrolled up
```

Initialize: `has_new_below: false,`

**Step 2: In `poll_new_messages`, after reloading messages for current conv:**

```rust
                if current_conv_has_new {
                    let at_bottom = self.scroll_offset >= self.rendered_line_count.saturating_sub(self.view_height + 5);
                    self.render_dirty = true;
                    self.load_messages().await;
                    if at_bottom {
                        self.scroll_offset = usize::MAX;
                    } else {
                        self.has_new_below = true;
                    }
                }
```

**Step 3: Clear indicator when jumping to bottom**

In the `G` key handler:
```rust
                KeyCode::Char('G') => {
                    self.scroll_offset = self.rendered_line_count;
                    self.has_new_below = false;
                }
```

**Step 4: Render indicator in `draw_main`**

After the scroll indicator, add:
```rust
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
```

**Step 5: Build and test**

Run: `cargo build`

**Step 6: Commit**

```bash
git add src/tui/app.rs
git commit -m "Auto-follow new messages, show indicator when scrolled up"
```

---

### Task 10: Fix scrolling -- key acceleration

**Files:**
- Modify: `src/tui/app.rs`

**Step 1: Add repeat tracking fields**

```rust
    last_scroll_key: Option<KeyCode>,
    scroll_repeat_count: u32,
    last_scroll_time: std::time::Instant,
```

Initialize:
```rust
            last_scroll_key: None,
            scroll_repeat_count: 0,
            last_scroll_time: std::time::Instant::now(),
```

**Step 2: Add acceleration helper**

```rust
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
```

**Step 3: Use in j/k handlers**

```rust
                KeyCode::Char('j') | KeyCode::Down => {
                    let amount = self.scroll_amount(key);
                    self.scroll_offset = self.scroll_offset.saturating_add(amount);
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    let amount = self.scroll_amount(key);
                    self.scroll_offset = self.scroll_offset.saturating_sub(amount);
                }
```

**Step 4: Build and test**

Run: `cargo build`
Manual: hold j/k rapidly, verify it accelerates after 3 presses.

**Step 5: Commit**

```bash
git add src/tui/app.rs
git commit -m "Accelerate j/k scrolling on rapid repeat"
```

---

### Task 11: Wire store sync into TUI startup and polling

**Files:**
- Modify: `src/tui/app.rs` (add Store to App, write files during sync)

**Step 1: Add Store to App struct**

```rust
    store: store::Store,
```

Initialize in `App::new()`:
```rust
            store: store::Store::new(auth.config_dir())?,
```

**Step 2: In `load_conversations`, after saving cache, also write to store**

After `self.save_to_cache();`, add:

```rust
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
```

**Step 3: In `load_messages`, write messages to store**

After `self.messages = msgs;`, add:

```rust
                    // Write messages to local files
                    if let Some(conv_id) = &self.current_conv_id {
                        for m in &self.messages {
                            let msg_id = m.id.as_deref().unwrap_or("unknown");
                            let sender = m.imdisplayname.as_deref().unwrap_or("?");
                            let timestamp = m.timestamp.as_deref().unwrap_or("");
                            let content = m.content.as_deref().unwrap_or("");
                            let _ = self.store.save_message(conv_id, msg_id, timestamp, sender, content);
                        }
                    }
```

**Step 4: Build and test**

Run: `cargo build && cargo run`
Open a conversation, then check: `ls ~/.config/teamsh/data/conversations/ | head -5`

**Step 5: Commit**

```bash
git add src/tui/app.rs
git commit -m "Wire store into TUI: write files during sync and message load"
```

---

### Task 12: Update help text and remove old search code

**Files:**
- Modify: `src/tui/app.rs` (update help, remove nucleo search, clean up)

**Step 1: Update help text**

Replace help strings:
```rust
            Focus::Messages => " j/k:scroll  PgUp/Dn  G:end  g:top  h:sidebar  i:compose  /:tv-search  r:refresh ",
            _ => " Tab:section  j/k:nav  l:messages  /:tv-search  f:fav  r:refresh  e:emails  q:quit ",
```

**Step 2: Remove search_active, search_query, search_results fields and all draw_search/handle_search_key code**

Since `/` now spawns tv, the built-in search overlay is no longer needed. Remove:
- `search_active`, `search_query`, `search_results`, `search_email_results`, `search_list_state`, `search_people_results`, `search_highlight` fields
- `draw_search()` method
- `handle_search_key()` method
- `update_search_results()` method
- `search_total_items()` method
- `remote_search()` method
- `open_search_result()` method
- `close_search()` method
- `apply_search_highlight()` free function
- `byte_pos_in_original()` free function
- `Focus::Search` variant
- nucleo-matcher import in update_search_results

Keep `search_highlight` field if you want to support highlight from tv selection (can add later).

**Step 3: Remove nucleo-matcher from Cargo.toml**

Remove `nucleo-matcher = "0.3"` from dependencies.

**Step 4: Build and test**

Run: `cargo build 2>&1 | grep error`
Expected: No errors. May need to fix references to removed fields.

**Step 5: Commit**

```bash
git add src/tui/app.rs Cargo.toml Cargo.lock
git commit -m "Replace built-in search with tv integration, remove nucleo"
```

---

### Task 13: Final integration test and cleanup

**Files:**
- All modified files

**Step 1: Run full sync**

```bash
cargo run -- sync
```

Verify data dir is populated:
```bash
find ~/.config/teamsh/data/ -type f | wc -l
ls ~/.config/teamsh/data/conversations/ | head -5
ls ~/.config/teamsh/data/emails/ | head -5
```

**Step 2: Test list command**

```bash
cargo run -- list | head -20
```

**Step 3: Test tv integration**

```bash
cargo run
# Press /
# Verify tv opens, search works, preview shows content
# Select a result, verify TUI navigates to it
```

**Step 4: Test sidebar**

- Tab cycles through sections
- f pins/unpins
- Left/right moves between sidebar and main panel
- j/k navigates
- Scrolling is smooth, accelerates on repeat

**Step 5: Release build**

```bash
cargo build --release
ls -la target/release/teamsh
```

**Step 6: Commit any remaining fixes**

```bash
git add -A
git commit -m "teamsh v2: local storage, tv integration, sidebar sections, scroll fixes"
```
