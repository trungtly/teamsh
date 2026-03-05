# teamsh Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a dual-mode (CLI + TUI) Microsoft Teams client in Rust that uses OAuth2 refresh tokens to access the Teams chatsvc API.

**Architecture:** Shared core (auth, API client, types, HTML stripping) with two thin frontends: CLI (clap subcommands, JSON/plain output) and TUI (ratatui full-screen views). Auth uses the Teams web SPA client ID with a user-provided refresh token to mint IC3 access tokens.

**Tech Stack:** Rust, ratatui + crossterm, clap (derive), reqwest (async/rustls), tokio, serde

**Project dir:** `~/workspace/repos/trung-notes/shared/teamsh/`

---

### Task 1: Scaffold Rust Project

**Files:**
- Create: `Cargo.toml`
- Create: `src/main.rs`

**Step 1: Initialize cargo project**

Run: `cd ~/workspace/repos/trung-notes/shared/teamsh && cargo init`

**Step 2: Add dependencies to Cargo.toml**

Replace the generated `Cargo.toml` with:

```toml
[package]
name = "teamsh"
version = "0.1.0"
edition = "2021"

[dependencies]
clap = { version = "4", features = ["derive"] }
tokio = { version = "1", features = ["full"] }
reqwest = { version = "0.12", features = ["json", "rustls-tls"], default-features = false }
serde = { version = "1", features = ["derive"] }
serde_json = "1"
dirs = "6"
ratatui = "0.29"
crossterm = "0.28"
anyhow = "1"
```

**Step 3: Write minimal main.rs to verify it compiles**

```rust
fn main() {
    println!("teamsh");
}
```

**Step 4: Verify it builds**

Run: `cd ~/workspace/repos/trung-notes/shared/teamsh && cargo build`
Expected: Compiles successfully (dependencies download on first build)

**Step 5: Commit**

```bash
git add shared/teamsh/Cargo.toml shared/teamsh/src/main.rs
git commit -m "Scaffold teamsh Rust project with dependencies"
```

---

### Task 2: CLI Argument Parsing

**Files:**
- Create: `src/cli.rs`
- Modify: `src/main.rs`

**Step 1: Create src/cli.rs with clap derive structs**

```rust
use clap::{Parser, Subcommand};

#[derive(Parser)]
#[command(name = "teamsh", version, about = "Microsoft Teams from the terminal")]
pub struct Cli {
    #[command(subcommand)]
    pub command: Option<Commands>,

    /// Output format
    #[arg(long, global = true, default_value = "plain")]
    pub format: OutputFormat,
}

#[derive(Subcommand)]
pub enum Commands {
    /// Initialize auth with a refresh token
    Auth {
        #[command(subcommand)]
        action: AuthAction,
    },
    /// List conversations (channels, chats, meetings)
    Chats {
        /// Show only channels
        #[arg(long)]
        channels: bool,
        /// Show only DMs/group chats
        #[arg(long)]
        dms: bool,
    },
    /// Read messages from a conversation
    Messages {
        /// Conversation ID
        conv_id: String,
        /// Number of messages to fetch
        #[arg(long, default_value = "20")]
        last: u32,
        /// Output plain text (strip HTML)
        #[arg(long)]
        plain: bool,
    },
    /// Send a message to a conversation
    Send {
        /// Conversation ID
        conv_id: String,
        /// Message text (reads from stdin if not provided)
        message: Option<String>,
        /// Read message from stdin
        #[arg(long)]
        stdin: bool,
    },
    /// Launch TUI mode
    Tui,
}

#[derive(Subcommand)]
pub enum AuthAction {
    /// Set up refresh token
    Init,
    /// Test current token
    Test,
}

#[derive(Clone, Debug, clap::ValueEnum)]
pub enum OutputFormat {
    Plain,
    Json,
}
```

**Step 2: Update main.rs to parse args and dispatch**

```rust
mod cli;

use clap::Parser;
use cli::{Cli, Commands};

fn main() {
    let cli = Cli::parse();

    match &cli.command {
        Some(Commands::Auth { action }) => {
            println!("auth: {:?}", action);
        }
        Some(Commands::Chats { channels, dms }) => {
            println!("chats: channels={channels}, dms={dms}");
        }
        Some(Commands::Messages { conv_id, last, plain }) => {
            println!("messages: conv={conv_id}, last={last}, plain={plain}");
        }
        Some(Commands::Send { conv_id, message, stdin }) => {
            println!("send: conv={conv_id}, msg={message:?}, stdin={stdin}");
        }
        Some(Commands::Tui) | None => {
            println!("TUI mode (not yet implemented)");
        }
    }
}
```

**Step 3: Verify it builds and help works**

Run: `cargo build && ./target/debug/teamsh --help`
Expected: Shows help with subcommands (auth, chats, messages, send, tui)

Run: `./target/debug/teamsh chats --channels`
Expected: `chats: channels=true, dms=false`

**Step 4: Commit**

```bash
git add shared/teamsh/src/
git commit -m "Add CLI argument parsing with clap derive"
```

---

### Task 3: Auth Module

**Files:**
- Create: `src/auth.rs`
- Modify: `src/main.rs`

**Step 1: Create src/auth.rs**

This module handles:
- Storing/loading refresh token from `~/.config/teamsh/refresh_token`
- Exchanging refresh token for IC3 access token
- Rotating refresh tokens (each refresh returns a new one)

```rust
use anyhow::{Context, Result};
use serde::Deserialize;
use std::fs;
use std::path::PathBuf;

const CLIENT_ID: &str = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const ORIGIN: &str = "https://teams.cloud.microsoft";
const SCOPE: &str = "https://ic3.teams.office.com/.default openid profile offline_access";

#[derive(Deserialize)]
struct TokenResponse {
    access_token: String,
    refresh_token: Option<String>,
    expires_in: u64,
}

#[derive(Deserialize)]
struct TokenError {
    error: String,
    error_description: String,
}

pub struct Auth {
    config_dir: PathBuf,
    tenant_id: String,
    access_token: Option<String>,
}

impl Auth {
    pub fn new() -> Result<Self> {
        let config_dir = dirs::config_dir()
            .context("Could not find config directory")?
            .join("teamsh");
        fs::create_dir_all(&config_dir)?;

        let tenant_id = Self::load_file(&config_dir, "tenant_id")
            .unwrap_or_default();

        Ok(Self {
            config_dir,
            tenant_id,
            access_token: None,
        })
    }

    pub fn config_dir(&self) -> &PathBuf {
        &self.config_dir
    }

    /// Save refresh token and tenant ID during init
    pub fn save_init(&self, refresh_token: &str, tenant_id: &str) -> Result<()> {
        fs::write(self.config_dir.join("refresh_token"), refresh_token.trim())?;
        fs::write(self.config_dir.join("tenant_id"), tenant_id.trim())?;
        Ok(())
    }

    /// Get a valid access token, refreshing if needed
    pub async fn access_token(&mut self) -> Result<&str> {
        if self.access_token.is_none() {
            self.refresh().await?;
        }
        Ok(self.access_token.as_ref().unwrap())
    }

    /// Force refresh the access token
    pub async fn refresh(&mut self) -> Result<()> {
        let refresh_token = Self::load_file(&self.config_dir, "refresh_token")
            .context("No refresh token found. Run: teamsh auth init")?;

        if self.tenant_id.is_empty() {
            self.tenant_id = Self::load_file(&self.config_dir, "tenant_id")
                .context("No tenant ID found. Run: teamsh auth init")?;
        }

        let url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );

        let client = reqwest::Client::new();
        let resp = client
            .post(&url)
            .header("Origin", ORIGIN)
            .form(&[
                ("client_id", CLIENT_ID),
                ("grant_type", "refresh_token"),
                ("refresh_token", &refresh_token),
                ("scope", SCOPE),
            ])
            .send()
            .await?;

        let body = resp.text().await?;

        // Try to parse as success
        if let Ok(token_resp) = serde_json::from_str::<TokenResponse>(&body) {
            self.access_token = Some(token_resp.access_token);

            // Rotate refresh token
            if let Some(new_rt) = &token_resp.refresh_token {
                fs::write(self.config_dir.join("refresh_token"), new_rt)?;
            }

            return Ok(());
        }

        // Parse as error
        let err: TokenError = serde_json::from_str(&body)
            .context("Failed to parse token response")?;
        anyhow::bail!("Token refresh failed: {} - {}", err.error, err.error_description)
    }

    fn load_file(dir: &PathBuf, name: &str) -> Option<String> {
        fs::read_to_string(dir.join(name))
            .ok()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
    }
}
```

**Step 2: Wire auth into main.rs**

Update `main.rs` to add `mod auth;` and implement the `auth init` and `auth test` commands:

```rust
mod auth;
mod cli;

use anyhow::Result;
use clap::Parser;
use cli::{AuthAction, Cli, Commands};

#[tokio::main]
async fn main() -> Result<()> {
    let cli = Cli::parse();

    match &cli.command {
        Some(Commands::Auth { action }) => match action {
            AuthAction::Init => cmd_auth_init().await?,
            AuthAction::Test => cmd_auth_test().await?,
        },
        Some(Commands::Chats { .. }) => {
            println!("chats: not yet implemented");
        }
        Some(Commands::Messages { .. }) => {
            println!("messages: not yet implemented");
        }
        Some(Commands::Send { .. }) => {
            println!("send: not yet implemented");
        }
        Some(Commands::Tui) | None => {
            println!("TUI mode (not yet implemented)");
        }
    }

    Ok(())
}

async fn cmd_auth_init() -> Result<()> {
    println!("teamsh auth setup");
    println!();
    println!("To get your refresh token:");
    println!("1. Open https://teams.cloud.microsoft in browser");
    println!("2. Open DevTools > Network tab");
    println!("3. Filter for: login.microsoftonline.com");
    println!("4. Find a POST to oauth2/v2.0/token");
    println!("5. In Response, copy the refresh_token value");
    println!();

    let mut auth = auth::Auth::new()?;

    // Read tenant ID
    println!("Enter your tenant ID (from the URL path, e.g. f2dbeea5-...):");
    let mut tenant_id = String::new();
    std::io::stdin().read_line(&mut tenant_id)?;
    let tenant_id = tenant_id.trim();
    if tenant_id.is_empty() {
        anyhow::bail!("Tenant ID is required");
    }

    // Read refresh token
    println!("Paste your refresh token:");
    let mut refresh_token = String::new();
    std::io::stdin().read_line(&mut refresh_token)?;
    let refresh_token = refresh_token.trim();
    if refresh_token.is_empty() {
        anyhow::bail!("Refresh token is required");
    }

    auth.save_init(refresh_token, tenant_id)?;
    println!("Saved to {:?}", auth.config_dir());

    // Test it
    println!("Testing token refresh...");
    auth.refresh().await?;
    println!("Success! Token refresh works.");

    Ok(())
}

async fn cmd_auth_test() -> Result<()> {
    let mut auth = auth::Auth::new()?;
    println!("Refreshing token...");
    auth.refresh().await?;
    println!("Token is valid.");
    Ok(())
}
```

**Step 3: Build and test auth test (should fail gracefully without token)**

Run: `cargo build && ./target/debug/teamsh auth test`
Expected: Error message about missing refresh token

**Step 4: Test auth init with real token**

Run: `./target/debug/teamsh auth init`
Enter tenant ID and refresh token when prompted. Should print "Success! Token refresh works."

**Step 5: Commit**

```bash
git add shared/teamsh/src/
git commit -m "Add auth module with OAuth2 refresh token flow"
```

---

### Task 4: API Client and Types

**Files:**
- Create: `src/api.rs`
- Create: `src/types.rs`

**Step 1: Create src/types.rs**

```rust
use serde::Deserialize;

#[derive(Debug, Deserialize)]
pub struct ConversationsResponse {
    pub conversations: Vec<Conversation>,
}

#[derive(Debug, Deserialize)]
pub struct Conversation {
    pub id: String,
    #[serde(rename = "type")]
    pub conv_type: Option<String>,
    #[serde(rename = "threadProperties")]
    pub thread_properties: Option<ThreadProperties>,
}

#[derive(Debug, Deserialize)]
pub struct ThreadProperties {
    pub topic: Option<String>,
    #[serde(rename = "lastjoinat")]
    pub last_join_at: Option<String>,
    #[serde(rename = "memberCount")]
    pub member_count: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct MessagesResponse {
    pub messages: Vec<Message>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct Message {
    pub id: Option<String>,
    #[serde(rename = "originalarrivaltime")]
    pub timestamp: Option<String>,
    pub messagetype: Option<String>,
    pub imdisplayname: Option<String>,
    pub content: Option<String>,
    pub properties: Option<serde_json::Value>,
}

/// Categorized conversation for display
#[derive(Debug, Clone)]
pub enum ConvKind {
    Channel,
    Chat,
    Meeting,
    System,
}

impl Conversation {
    pub fn topic(&self) -> &str {
        self.thread_properties
            .as_ref()
            .and_then(|p| p.topic.as_deref())
            .unwrap_or("(no topic)")
    }

    pub fn kind(&self) -> ConvKind {
        let id = &self.id;
        if id.starts_with("48:") {
            ConvKind::System
        } else if id.contains("meeting_") {
            ConvKind::Meeting
        } else if id.contains("@thread.skype") || id.contains("@thread.tacv2") || id.contains("@thread.v2") {
            // Threads with topic are channels, without are group chats
            // Heuristic: if topic is set and not "(no topic)", it's likely a channel
            let topic = self.topic();
            if topic != "(no topic)" {
                ConvKind::Channel
            } else {
                ConvKind::Chat
            }
        } else {
            ConvKind::Chat
        }
    }
}
```

**Step 2: Create src/api.rs**

```rust
use anyhow::{Context, Result};
use crate::auth::Auth;
use crate::types::{ConversationsResponse, MessagesResponse};

const BASE_URL: &str = "https://teams.cloud.microsoft/api/chatsvc";
const DEFAULT_REGION: &str = "au";

pub struct Api {
    client: reqwest::Client,
    region: String,
}

impl Api {
    pub fn new() -> Self {
        Self {
            client: reqwest::Client::new(),
            region: DEFAULT_REGION.to_string(),
        }
    }

    pub async fn list_conversations(&self, auth: &mut Auth, page_size: u32) -> Result<ConversationsResponse> {
        let url = format!(
            "{}/{}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize={}",
            BASE_URL, self.region, page_size
        );
        self.get(auth, &url).await
    }

    pub async fn get_messages(&self, auth: &mut Auth, conv_id: &str, count: u32) -> Result<MessagesResponse> {
        let encoded_id = urlencoding::encode(conv_id);
        let url = format!(
            "{}/{}/v1/users/ME/conversations/{}/messages?view=msnp24Equivalent&pageSize={}",
            BASE_URL, self.region, encoded_id, count
        );
        self.get(auth, &url).await
    }

    pub async fn send_message(&self, auth: &mut Auth, conv_id: &str, content: &str) -> Result<()> {
        let encoded_id = urlencoding::encode(conv_id);
        let url = format!(
            "{}/{}/v1/users/ME/conversations/{}/messages",
            BASE_URL, self.region, encoded_id
        );
        let token = auth.access_token().await?;
        let body = serde_json::json!({
            "content": content,
            "messagetype": "RichText/Html",
            "contenttype": "text",
        });
        let resp = self.client
            .post(&url)
            .header("Authorization", format!("Bearer {}", token))
            .header("behavioroverride", "redirectAs404")
            .header("x-ms-migration", "True")
            .json(&body)
            .send()
            .await?;

        if !resp.status().is_success() {
            let status = resp.status();
            let text = resp.text().await.unwrap_or_default();
            anyhow::bail!("Send failed ({}): {}", status, text);
        }
        Ok(())
    }

    async fn get<T: serde::de::DeserializeOwned>(&self, auth: &mut Auth, url: &str) -> Result<T> {
        let token = auth.access_token().await?;
        let resp = self.client
            .get(url)
            .header("Authorization", format!("Bearer {}", token))
            .header("behavioroverride", "redirectAs404")
            .header("x-ms-migration", "True")
            .send()
            .await?
            .text()
            .await?;

        serde_json::from_str(&resp)
            .with_context(|| format!("Failed to parse response: {}", &resp[..resp.len().min(200)]))
    }
}
```

**Step 3: Add urlencoding dependency**

Add to Cargo.toml under `[dependencies]`:
```toml
urlencoding = "2"
```

**Step 4: Add mod declarations to main.rs**

Add at top of main.rs:
```rust
mod api;
mod types;
```

**Step 5: Verify it builds**

Run: `cargo build`

**Step 6: Commit**

```bash
git add shared/teamsh/src/ shared/teamsh/Cargo.toml
git commit -m "Add Teams chatsvc API client and types"
```

---

### Task 5: HTML Stripping

**Files:**
- Create: `src/html.rs`
- Modify: `src/main.rs` (add mod)

**Step 1: Create src/html.rs**

Simple HTML tag stripper without pulling in a full HTML parser. Handles the common patterns from Teams messages.

```rust
/// Strip HTML tags and decode common entities to plain text
pub fn strip_html(html: &str) -> String {
    let mut result = String::with_capacity(html.len());
    let mut in_tag = false;
    let mut in_entity = false;
    let mut entity = String::new();

    for ch in html.chars() {
        match ch {
            '<' => {
                in_tag = true;
                // Check if previous content suggests a block element
            }
            '>' if in_tag => {
                in_tag = false;
            }
            '&' if !in_tag => {
                in_entity = true;
                entity.clear();
            }
            ';' if in_entity => {
                in_entity = false;
                match entity.as_str() {
                    "amp" => result.push('&'),
                    "lt" => result.push('<'),
                    "gt" => result.push('>'),
                    "quot" => result.push('"'),
                    "nbsp" => result.push(' '),
                    "apos" => result.push('\''),
                    _ => {
                        result.push('&');
                        result.push_str(&entity);
                        result.push(';');
                    }
                }
            }
            _ if in_tag => {}
            _ if in_entity => {
                entity.push(ch);
            }
            _ => {
                result.push(ch);
            }
        }
    }

    // Collapse multiple whitespace/newlines
    let mut collapsed = String::with_capacity(result.len());
    let mut last_was_space = false;
    for ch in result.trim().chars() {
        if ch.is_whitespace() {
            if !last_was_space {
                collapsed.push(' ');
                last_was_space = true;
            }
        } else {
            collapsed.push(ch);
            last_was_space = false;
        }
    }

    collapsed
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_strip_simple() {
        assert_eq!(strip_html("<p>Hello</p>"), "Hello");
    }

    #[test]
    fn test_strip_entities() {
        assert_eq!(strip_html("A &amp; B"), "A & B");
        assert_eq!(strip_html("&lt;code&gt;"), "<code>");
    }

    #[test]
    fn test_strip_nested() {
        assert_eq!(
            strip_html("<p>Hi <strong>team</strong>, how are you?</p>"),
            "Hi team, how are you?"
        );
    }

    #[test]
    fn test_strip_nbsp() {
        assert_eq!(strip_html("<p>&nbsp;</p>"), "");
    }

    #[test]
    fn test_plain_text_passthrough() {
        assert_eq!(strip_html("no html here"), "no html here");
    }
}
```

**Step 2: Add `mod html;` to main.rs**

**Step 3: Run tests**

Run: `cd ~/workspace/repos/trung-notes/shared/teamsh && cargo test`
Expected: All 5 tests pass

**Step 4: Commit**

```bash
git add shared/teamsh/src/html.rs shared/teamsh/src/main.rs
git commit -m "Add HTML tag stripping for Teams messages"
```

---

### Task 6: CLI Commands (chats, messages, send)

**Files:**
- Modify: `src/main.rs`

**Step 1: Implement all CLI commands in main.rs**

Replace the placeholder match arms with real implementations:

```rust
mod api;
mod auth;
mod cli;
mod html;
mod types;

use anyhow::Result;
use clap::Parser;
use cli::{AuthAction, Cli, Commands, OutputFormat};
use types::ConvKind;

#[tokio::main]
async fn main() -> Result<()> {
    let cli = Cli::parse();

    match &cli.command {
        Some(Commands::Auth { action }) => match action {
            AuthAction::Init => cmd_auth_init().await?,
            AuthAction::Test => cmd_auth_test().await?,
        },
        Some(Commands::Chats { channels, dms }) => {
            cmd_chats(&cli.format, *channels, *dms).await?;
        }
        Some(Commands::Messages { conv_id, last, plain }) => {
            cmd_messages(&cli.format, conv_id, *last, *plain).await?;
        }
        Some(Commands::Send { conv_id, message, stdin }) => {
            cmd_send(conv_id, message.as_deref(), *stdin).await?;
        }
        Some(Commands::Tui) | None => {
            println!("TUI mode (not yet implemented)");
        }
    }

    Ok(())
}

async fn cmd_auth_init() -> Result<()> {
    println!("teamsh auth setup");
    println!();
    println!("To get your refresh token:");
    println!("1. Open https://teams.cloud.microsoft in browser");
    println!("2. Open DevTools > Network tab");
    println!("3. Filter for: login.microsoftonline.com");
    println!("4. Find a POST to oauth2/v2.0/token");
    println!("5. In Response, copy the refresh_token value");
    println!();

    let auth = auth::Auth::new()?;

    println!("Enter your tenant ID (from the URL path, e.g. f2dbeea5-...):");
    let mut tenant_id = String::new();
    std::io::stdin().read_line(&mut tenant_id)?;
    let tenant_id = tenant_id.trim();
    if tenant_id.is_empty() {
        anyhow::bail!("Tenant ID is required");
    }

    println!("Paste your refresh token:");
    let mut refresh_token = String::new();
    std::io::stdin().read_line(&mut refresh_token)?;
    let refresh_token = refresh_token.trim();
    if refresh_token.is_empty() {
        anyhow::bail!("Refresh token is required");
    }

    auth.save_init(refresh_token, tenant_id)?;
    println!("Saved to {:?}", auth.config_dir());

    let mut auth = auth::Auth::new()?;
    println!("Testing token refresh...");
    auth.refresh().await?;
    println!("Auth is working.");

    Ok(())
}

async fn cmd_auth_test() -> Result<()> {
    let mut auth = auth::Auth::new()?;
    println!("Refreshing token...");
    auth.refresh().await?;
    println!("Token is valid.");
    Ok(())
}

async fn cmd_chats(format: &OutputFormat, channels_only: bool, dms_only: bool) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new();
    let resp = api.list_conversations(&mut auth, 100).await?;

    let convs: Vec<_> = resp.conversations.into_iter().filter(|c| {
        let kind = c.kind();
        if channels_only {
            return matches!(kind, ConvKind::Channel);
        }
        if dms_only {
            return matches!(kind, ConvKind::Chat);
        }
        // Default: show channels and chats, skip system
        !matches!(kind, ConvKind::System)
    }).collect();

    match format {
        OutputFormat::Json => {
            let items: Vec<serde_json::Value> = convs.iter().map(|c| {
                serde_json::json!({
                    "id": c.id,
                    "topic": c.topic(),
                    "kind": format!("{:?}", c.kind()),
                })
            }).collect();
            println!("{}", serde_json::to_string_pretty(&items)?);
        }
        OutputFormat::Plain => {
            for c in &convs {
                let kind = match c.kind() {
                    ConvKind::Channel => "#",
                    ConvKind::Chat => "@",
                    ConvKind::Meeting => "M",
                    ConvKind::System => "S",
                };
                println!("{} {} [{}]", kind, c.topic(), c.id);
            }
        }
    }

    Ok(())
}

async fn cmd_messages(format: &OutputFormat, conv_id: &str, last: u32, plain: bool) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new();
    let resp = api.get_messages(&mut auth, conv_id, last).await?;

    // Filter to actual messages (not system/control messages)
    let msgs: Vec<_> = resp.messages.into_iter().filter(|m| {
        matches!(m.messagetype.as_deref(), Some("RichText/Html") | Some("Text"))
    }).collect();

    match format {
        OutputFormat::Json => {
            let items: Vec<serde_json::Value> = msgs.iter().map(|m| {
                let content = m.content.as_deref().unwrap_or("");
                serde_json::json!({
                    "id": m.id,
                    "timestamp": m.timestamp,
                    "sender": m.imdisplayname,
                    "content": if plain { html::strip_html(content) } else { content.to_string() },
                })
            }).collect();
            println!("{}", serde_json::to_string_pretty(&items)?);
        }
        OutputFormat::Plain => {
            for m in msgs.iter().rev() {
                let sender = m.imdisplayname.as_deref().unwrap_or("?");
                let content = m.content.as_deref().unwrap_or("");
                let time = m.timestamp.as_deref().unwrap_or("").get(11..16).unwrap_or("??:??");
                let text = if plain { html::strip_html(content) } else { content.to_string() };
                println!("[{}] {}: {}", time, sender, text);
            }
        }
    }

    Ok(())
}

async fn cmd_send(conv_id: &str, message: Option<&str>, from_stdin: bool) -> Result<()> {
    let text = if from_stdin || message.is_none() {
        let mut buf = String::new();
        std::io::Read::read_to_string(&mut std::io::stdin(), &mut buf)?;
        buf
    } else {
        message.unwrap().to_string()
    };

    if text.trim().is_empty() {
        anyhow::bail!("Empty message");
    }

    let mut auth = auth::Auth::new()?;
    let api = api::Api::new();
    api.send_message(&mut auth, conv_id, text.trim()).await?;
    println!("Sent.");

    Ok(())
}
```

**Step 2: Verify it builds**

Run: `cargo build`

**Step 3: Test CLI commands with real token**

Run:
```bash
./target/debug/teamsh chats
./target/debug/teamsh chats --channels
./target/debug/teamsh messages "19:019b4922f6cc4d278c7cc288038c596d@thread.v2" --last 5 --plain
./target/debug/teamsh chats --format json | head -20
```

**Step 4: Commit**

```bash
git add shared/teamsh/src/
git commit -m "Implement CLI commands: chats, messages, send"
```

---

### Task 7: TUI - App State and Event Loop

**Files:**
- Create: `src/tui/mod.rs`
- Create: `src/tui/app.rs`
- Modify: `src/main.rs`

**Step 1: Create src/tui/mod.rs**

```rust
pub mod app;

use anyhow::Result;
use app::App;

pub async fn run() -> Result<()> {
    let mut terminal = ratatui::init();
    let result = App::new().await?.run(&mut terminal).await;
    ratatui::restore();
    result
}
```

**Step 2: Create src/tui/app.rs**

```rust
use anyhow::Result;
use crossterm::event::{self, Event, KeyCode, KeyEventKind};
use ratatui::buffer::Buffer;
use ratatui::layout::{Constraint, Layout, Rect};
use ratatui::style::{Color, Modifier, Style, Stylize};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, List, ListItem, ListState, Paragraph, Widget, Wrap};
use ratatui::{DefaultTerminal, Frame};
use std::time::Duration;

use crate::api::Api;
use crate::auth::Auth;
use crate::html;
use crate::types::{Conversation, ConvKind, Message};

#[derive(Debug, Clone, PartialEq)]
enum Screen {
    ChatList,
    Conversation,
}

#[derive(Debug, Clone, PartialEq)]
enum InputMode {
    Normal,
    Composing,
}

pub struct App {
    auth: Auth,
    api: Api,
    screen: Screen,
    exit: bool,

    // Chat list
    conversations: Vec<Conversation>,
    chat_list_state: ListState,

    // Conversation view
    current_conv_id: Option<String>,
    current_conv_topic: String,
    messages: Vec<Message>,
    scroll_offset: u16,

    // Input
    input_mode: InputMode,
    input_buffer: String,

    // Status
    status: String,
}

impl App {
    pub async fn new() -> Result<Self> {
        let mut auth = Auth::new()?;
        let api = Api::new();

        let mut app = Self {
            auth,
            api,
            screen: Screen::ChatList,
            exit: false,
            conversations: Vec::new(),
            chat_list_state: ListState::default(),
            current_conv_id: None,
            current_conv_topic: String::new(),
            messages: Vec::new(),
            scroll_offset: 0,
            input_mode: InputMode::Normal,
            input_buffer: String::new(),
            status: "Loading...".to_string(),
        };

        app.load_conversations().await;
        Ok(app)
    }

    pub async fn run(&mut self, terminal: &mut DefaultTerminal) -> Result<()> {
        while !self.exit {
            terminal.draw(|frame| self.draw(frame))?;
            if event::poll(Duration::from_millis(100))? {
                if let Event::Key(key) = event::read()? {
                    if key.kind == KeyEventKind::Press {
                        self.handle_key(key.code).await;
                    }
                }
            }
        }
        Ok(())
    }

    fn draw(&mut self, frame: &mut Frame) {
        match self.screen {
            Screen::ChatList => self.draw_chat_list(frame),
            Screen::Conversation => self.draw_conversation(frame),
        }
    }

    fn draw_chat_list(&mut self, frame: &mut Frame) {
        let area = frame.area();
        let [main_area, status_area] = Layout::vertical([
            Constraint::Min(1),
            Constraint::Length(1),
        ]).areas(area);

        // Build grouped list items
        let mut items: Vec<ListItem> = Vec::new();
        let mut current_kind: Option<String> = None;

        for conv in &self.conversations {
            let kind = conv.kind();
            if matches!(kind, ConvKind::System) {
                continue;
            }

            let kind_label = match &kind {
                ConvKind::Channel => "Channels",
                ConvKind::Chat => "Chats",
                ConvKind::Meeting => "Meetings",
                ConvKind::System => continue,
            };

            if current_kind.as_deref() != Some(kind_label) {
                if current_kind.is_some() {
                    items.push(ListItem::new(""));
                }
                items.push(ListItem::new(
                    Line::from(kind_label.to_string()).style(Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD))
                ));
                current_kind = Some(kind_label.to_string());
            }

            let prefix = match kind {
                ConvKind::Channel => " # ",
                ConvKind::Chat => " @ ",
                ConvKind::Meeting => " M ",
                _ => "   ",
            };
            items.push(ListItem::new(format!("{}{}", prefix, conv.topic())));
        }

        let list = List::new(items)
            .block(Block::default().title(" teamsh ").borders(Borders::ALL))
            .highlight_style(Style::default().bg(Color::DarkGray).add_modifier(Modifier::BOLD));

        frame.render_stateful_widget(list, main_area, &mut self.chat_list_state);

        // Status bar
        let status = Paragraph::new(Line::from(vec![
            Span::raw(" j/k: navigate  Enter: open  q: quit  r: refresh "),
        ]).style(Style::default().fg(Color::DarkGray)));
        frame.render_widget(status, status_area);
    }

    fn draw_conversation(&mut self, frame: &mut Frame) {
        let area = frame.area();
        let [header_area, msg_area, input_area, status_area] = Layout::vertical([
            Constraint::Length(1),
            Constraint::Min(1),
            Constraint::Length(3),
            Constraint::Length(1),
        ]).areas(area);

        // Header
        let header = Paragraph::new(
            Line::from(self.current_conv_topic.clone())
                .style(Style::default().fg(Color::Cyan).add_modifier(Modifier::BOLD))
        );
        frame.render_widget(header, header_area);

        // Messages (reversed - newest at bottom)
        let msg_lines: Vec<Line> = self.messages.iter().rev().flat_map(|m| {
            let sender = m.imdisplayname.as_deref().unwrap_or("?");
            let content = m.content.as_deref().unwrap_or("");
            let time = m.timestamp.as_deref().unwrap_or("").get(11..16).unwrap_or("??:??");
            let text = html::strip_html(content);

            vec![
                Line::from(vec![
                    Span::styled(format!("{} ", time), Style::default().fg(Color::DarkGray)),
                    Span::styled(sender.to_string(), Style::default().fg(Color::Green).add_modifier(Modifier::BOLD)),
                ]),
                Line::from(format!("  {}", text)),
                Line::from(""),
            ]
        }).collect();

        let messages = Paragraph::new(msg_lines)
            .block(Block::default().borders(Borders::ALL))
            .wrap(Wrap { trim: false })
            .scroll((self.scroll_offset, 0));
        frame.render_widget(messages, msg_area);

        // Input
        let input_style = if self.input_mode == InputMode::Composing {
            Style::default().fg(Color::White)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        let input = Paragraph::new(self.input_buffer.as_str())
            .block(Block::default().borders(Borders::ALL).title(" Message ").border_style(input_style))
            .wrap(Wrap { trim: false });
        frame.render_widget(input, input_area);

        // Cursor in input when composing
        if self.input_mode == InputMode::Composing {
            frame.set_cursor_position((
                input_area.x + 1 + self.input_buffer.len() as u16,
                input_area.y + 1,
            ));
        }

        // Status bar
        let help = match self.input_mode {
            InputMode::Normal => " Esc: back  Tab: compose  j/k: scroll  r: refresh ",
            InputMode::Composing => " Enter: send  Esc: cancel ",
        };
        let status = Paragraph::new(Line::from(help).style(Style::default().fg(Color::DarkGray)));
        frame.render_widget(status, status_area);
    }

    async fn handle_key(&mut self, key: KeyCode) {
        match (&self.screen, &self.input_mode) {
            // Chat list - normal mode
            (Screen::ChatList, InputMode::Normal) => match key {
                KeyCode::Char('q') => self.exit = true,
                KeyCode::Char('j') | KeyCode::Down => self.next_chat(),
                KeyCode::Char('k') | KeyCode::Up => self.prev_chat(),
                KeyCode::Char('r') => self.load_conversations().await,
                KeyCode::Enter => self.open_conversation().await,
                _ => {}
            },
            // Conversation - normal mode
            (Screen::Conversation, InputMode::Normal) => match key {
                KeyCode::Esc | KeyCode::Char('q') => {
                    self.screen = Screen::ChatList;
                    self.messages.clear();
                }
                KeyCode::Tab | KeyCode::Char('i') => {
                    self.input_mode = InputMode::Composing;
                }
                KeyCode::Char('j') | KeyCode::Down => {
                    self.scroll_offset = self.scroll_offset.saturating_add(1);
                }
                KeyCode::Char('k') | KeyCode::Up => {
                    self.scroll_offset = self.scroll_offset.saturating_sub(1);
                }
                KeyCode::Char('r') => self.load_messages().await,
                _ => {}
            },
            // Conversation - composing mode
            (Screen::Conversation, InputMode::Composing) => match key {
                KeyCode::Esc => {
                    self.input_mode = InputMode::Normal;
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
            _ => {}
        }
    }

    fn next_chat(&mut self) {
        let total = self.visible_count();
        if total == 0 { return; }
        let i = self.chat_list_state.selected().map(|i| (i + 1).min(total - 1)).unwrap_or(0);
        self.chat_list_state.select(Some(i));
    }

    fn prev_chat(&mut self) {
        let i = self.chat_list_state.selected().map(|i| i.saturating_sub(1)).unwrap_or(0);
        self.chat_list_state.select(Some(i));
    }

    fn visible_count(&self) -> usize {
        // Count list items including headers and spacers
        let mut count = 0;
        let mut current_kind: Option<String> = None;
        for conv in &self.conversations {
            let kind = conv.kind();
            if matches!(kind, ConvKind::System) { continue; }
            let kind_label = match &kind {
                ConvKind::Channel => "Channels",
                ConvKind::Chat => "Chats",
                ConvKind::Meeting => "Meetings",
                _ => continue,
            };
            if current_kind.as_deref() != Some(kind_label) {
                if current_kind.is_some() { count += 1; } // spacer
                count += 1; // header
                current_kind = Some(kind_label.to_string());
            }
            count += 1;
        }
        count
    }

    /// Map a list selection index back to a conversation
    fn selected_conversation(&self) -> Option<&Conversation> {
        let selected = self.chat_list_state.selected()?;
        let mut idx = 0;
        let mut current_kind: Option<String> = None;
        for conv in &self.conversations {
            let kind = conv.kind();
            if matches!(kind, ConvKind::System) { continue; }
            let kind_label = match &kind {
                ConvKind::Channel => "Channels",
                ConvKind::Chat => "Chats",
                ConvKind::Meeting => "Meetings",
                _ => continue,
            };
            if current_kind.as_deref() != Some(kind_label) {
                if current_kind.is_some() { idx += 1; }
                idx += 1;
                current_kind = Some(kind_label.to_string());
            }
            if idx == selected {
                return Some(conv);
            }
            idx += 1;
        }
        None
    }

    async fn open_conversation(&mut self) {
        if let Some(conv) = self.selected_conversation() {
            self.current_conv_id = Some(conv.id.clone());
            self.current_conv_topic = conv.topic().to_string();
            self.screen = Screen::Conversation;
            self.scroll_offset = 0;
            self.load_messages().await;
        }
    }

    async fn load_conversations(&mut self) {
        self.status = "Loading conversations...".to_string();
        match self.api.list_conversations(&mut self.auth, 100).await {
            Ok(resp) => {
                // Sort: channels first, then chats, then meetings
                let mut convs = resp.conversations;
                convs.sort_by_key(|c| match c.kind() {
                    ConvKind::Channel => 0,
                    ConvKind::Chat => 1,
                    ConvKind::Meeting => 2,
                    ConvKind::System => 3,
                });
                self.conversations = convs;
                self.status = format!("{} conversations loaded", self.conversations.len());
                if self.chat_list_state.selected().is_none() && !self.conversations.is_empty() {
                    // Select first actual item (skip header)
                    self.chat_list_state.select(Some(2));
                }
            }
            Err(e) => {
                self.status = format!("Error: {}", e);
            }
        }
    }

    async fn load_messages(&mut self) {
        if let Some(conv_id) = &self.current_conv_id.clone() {
            match self.api.get_messages(&mut self.auth, conv_id, 30).await {
                Ok(resp) => {
                    self.messages = resp.messages.into_iter().filter(|m| {
                        matches!(m.messagetype.as_deref(), Some("RichText/Html") | Some("Text"))
                    }).collect();
                }
                Err(e) => {
                    self.status = format!("Error: {}", e);
                }
            }
        }
    }

    async fn send_message(&mut self) {
        let text = self.input_buffer.trim().to_string();
        if text.is_empty() { return; }
        if let Some(conv_id) = &self.current_conv_id.clone() {
            match self.api.send_message(&mut self.auth, conv_id, &text).await {
                Ok(()) => {
                    self.input_buffer.clear();
                    self.input_mode = InputMode::Normal;
                    // Reload messages to see our sent message
                    self.load_messages().await;
                }
                Err(e) => {
                    self.status = format!("Send error: {}", e);
                }
            }
        }
    }
}
```

**Step 3: Add tui mod to main.rs and wire up the TUI command**

Add `mod tui;` and change the TUI match arm:

```rust
Some(Commands::Tui) | None => {
    tui::run().await?;
}
```

**Step 4: Build and test**

Run: `cargo build && ./target/debug/teamsh`
Expected: TUI launches, shows conversations, navigate with j/k, Enter to open, Esc to go back, q to quit

**Step 5: Commit**

```bash
git add shared/teamsh/src/
git commit -m "Add TUI mode with chat list and conversation views"
```

---

### Task 8: Region Auto-Detection

**Files:**
- Modify: `src/api.rs`
- Modify: `src/auth.rs`

**Step 1: Extract region from refresh token JWT**

In `auth.rs`, add a method to decode the JWT and extract the region. The Skype token embedded in the OAuth flow contains `rgn` field, but we can also detect it from the access token audience. Simpler: make it configurable with a default.

Add to `Auth`:

```rust
pub fn region(&self) -> String {
    Self::load_file(&self.config_dir, "region")
        .unwrap_or_else(|| "au".to_string())
}
```

Update `save_init` to also accept region:

```rust
pub fn save_init(&self, refresh_token: &str, tenant_id: &str, region: &str) -> Result<()> {
    fs::write(self.config_dir.join("refresh_token"), refresh_token.trim())?;
    fs::write(self.config_dir.join("tenant_id"), tenant_id.trim())?;
    fs::write(self.config_dir.join("region"), region.trim())?;
    Ok(())
}
```

**Step 2: Update Api to take region**

In `api.rs`, change `new()`:

```rust
pub fn new(region: &str) -> Self {
    Self {
        client: reqwest::Client::new(),
        region: region.to_string(),
    }
}
```

**Step 3: Update callers in main.rs and tui/app.rs**

```rust
let auth = auth::Auth::new()?;
let api = api::Api::new(&auth.region());
```

**Step 4: Update auth init to ask for region**

Add a region prompt (default "au") in cmd_auth_init.

**Step 5: Build and test**

Run: `cargo build && ./target/debug/teamsh chats`

**Step 6: Commit**

```bash
git add shared/teamsh/src/
git commit -m "Add configurable region for chatsvc API"
```

---

## Summary

| Task | What | Output |
|------|------|--------|
| 1 | Scaffold Rust project | Compiling skeleton |
| 2 | CLI arg parsing | `teamsh --help` works |
| 3 | Auth module | `teamsh auth init/test` works |
| 4 | API client + types | Core API layer |
| 5 | HTML stripping | Tests pass |
| 6 | CLI commands | `teamsh chats`, `teamsh messages`, `teamsh send` work |
| 7 | TUI mode | Full interactive TUI |
| 8 | Region config | Multi-region support |

After Task 6 you have a **fully working CLI tool** you can pipe to LLMs.
After Task 7 you have a **working TUI** for interactive use.
