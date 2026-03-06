mod api;
mod auth;
mod cache;
mod cli;
mod store;
mod html;
mod tui;
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
        Some(Commands::Search { query }) => {
            cmd_search(query).await?;
        }
        Some(Commands::Green { hours, keep }) => {
            cmd_green(*hours, *keep).await?;
        }
        Some(Commands::Emails { last }) => {
            cmd_emails(*last).await?;
        }
        Some(Commands::Sync) => {
            cmd_sync().await?;
        }
        Some(Commands::List) => {
            cmd_list().await?;
        }
        Some(Commands::Preview { path }) => {
            cmd_preview(path).await?;
        }
        Some(Commands::Tui { demo }) if *demo => {
            tui::run_demo().await?;
        }
        Some(Commands::Tui { .. }) | None => {
            tui::run().await?;
        }
    }

    Ok(())
}

async fn cmd_auth_init() -> Result<()> {
    println!("teamsh auth setup");
    println!();

    let mut auth = auth::Auth::new()?;

    // Check if tenant_id already exists
    let existing_tenant = std::fs::read_to_string(auth.config_dir().join("tenant_id"))
        .ok()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty());

    let tenant_id = if let Some(ref t) = existing_tenant {
        println!("Using existing tenant: {}", t);
        t.clone()
    } else {
        println!("Enter your tenant ID (from the URL path, e.g. f2dbeea5-...):");
        println!("(leave empty to auto-detect from login)");
        let mut tid = String::new();
        std::io::stdin().read_line(&mut tid)?;
        tid.trim().to_string()
    };

    // Set region if not already set
    let region_path = auth.config_dir().join("region");
    if !region_path.exists() {
        println!("Enter your region (default: au):");
        let mut region = String::new();
        std::io::stdin().read_line(&mut region)?;
        let region = region.trim();
        let region = if region.is_empty() { "au" } else { region };
        std::fs::write(&region_path, region)?;
    }

    // Login 1: Teams chat — must grab refresh token from browser (SPA limitation)
    println!("=== Login 1/2: Teams (chat & messages) ===");
    println!();
    println!("The Teams API requires a browser token (SPA app, no device code).");
    println!("1. Open https://teams.cloud.microsoft in your browser");
    println!("2. Open DevTools (F12) > Network tab");
    println!("3. Filter for: login.microsoftonline.com");
    println!("4. Find a POST to oauth2/v2.0/token");
    println!("5. In Response tab, copy the refresh_token value");
    println!();
    println!("Paste refresh token (or press Enter to skip if already valid):");
    let mut rt_input = String::new();
    std::io::stdin().read_line(&mut rt_input)?;
    let rt_input = rt_input.trim();
    if !rt_input.is_empty() {
        std::fs::write(auth.config_dir().join("refresh_token"), rt_input)?;
        // Verify
        let mut test_auth = auth::Auth::new()?;
        test_auth.refresh_teams().await?;
        println!("Teams token OK!");
    } else {
        println!("Skipped (using existing token).");
    }

    // Login 2: Graph emails — device code flow (90 day token)
    println!("\n=== Login 2/2: Emails (device code, ~90 day token) ===");
    println!("Skip with Ctrl+C if already valid.");
    match auth.login_graph(&tenant_id).await {
        Ok(()) => {}
        Err(e) => println!("Skipped: {}", e),
    }

    println!("\nDone! Teams token: ~24h. Email token: ~90 days.");

    Ok(())
}

async fn cmd_auth_test() -> Result<()> {
    let mut auth = auth::Auth::new()?;
    println!("Refreshing token...");
    auth.refresh_teams().await?;
    println!("Token is valid.");
    Ok(())
}

async fn cmd_chats(format: &OutputFormat, channels_only: bool, dms_only: bool) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    let resp = api.list_conversations(&mut auth, 100).await?;

    let convs: Vec<_> = resp.conversations.into_iter().filter(|c| {
        let kind = c.kind();
        if channels_only {
            return matches!(kind, ConvKind::Channel);
        }
        if dms_only {
            return matches!(kind, ConvKind::Chat);
        }
        !matches!(kind, ConvKind::System)
    }).collect();

    match format {
        OutputFormat::Json => {
            let items: Vec<serde_json::Value> = convs.iter().map(|c| {
                serde_json::json!({
                    "id": c.id,
                    "topic": c.display_name(""),
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
                println!("{} {} [{}]", kind, c.display_name(""), c.id);
            }
        }
    }

    Ok(())
}

async fn cmd_messages(format: &OutputFormat, conv_id: &str, last: u32, plain: bool) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    let resp = api.get_messages(&mut auth, conv_id, last).await?;

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

async fn cmd_search(query: &str) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    println!("Searching for: {}", query);
    match api.search_people(&mut auth, query).await {
        Ok(results) => {
            println!("Found {} results:", results.len());
            for (name, email) in &results {
                println!("  {} ({})", name, email);
            }
        }
        Err(e) => {
            println!("Error: {}", e);
        }
    }
    Ok(())
}

async fn cmd_green(hours: u64, keep: bool) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());

    println!("Setting presence to Available for {} hours...", hours);
    api.set_available(&mut auth, hours).await?;
    println!("Status set to Available (green)");

    if keep {
        println!("Keeping alive (Ctrl+C to stop)...");
        loop {
            tokio::time::sleep(std::time::Duration::from_secs(240)).await;
            if let Err(e) = api.report_activity(&mut auth).await {
                eprintln!("Activity report failed: {}", e);
            } else {
                println!("  activity reported");
            }
        }
    }

    Ok(())
}

async fn cmd_emails(last: u32) -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    println!("Fetching {} emails from inbox...", last);
    match api.list_emails(&mut auth, "inbox", last).await {
        Ok(emails) => {
            println!("Got {} emails:", emails.len());
            for email in &emails {
                let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("?");
                let from = email.get("from")
                    .and_then(|v| v.get("emailAddress"))
                    .and_then(|v| v.get("name"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("?");
                let is_read = email.get("isRead").and_then(|v| v.as_bool()).unwrap_or(true);
                let marker = if is_read { " " } else { "*" };
                println!("{} {} - {}", marker, from, subject);
            }
        }
        Err(e) => {
            eprintln!("Error: {}", e);
        }
    }
    Ok(())
}

async fn cmd_sync() -> Result<()> {
    let mut auth = auth::Auth::new()?;
    let api = api::Api::new(&auth.region());
    let store = store::Store::new(auth.config_dir())?;

    // --- Conversations ---
    eprintln!("Fetching conversations...");
    let resp = api.list_conversations(&mut auth, 100).await?;
    let mut convs = resp.conversations;
    convs.sort_by(|a, b| b.version.cmp(&a.version));

    // Detect own name: most frequent sender across last messages
    let mut name_counts: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    for c in &convs {
        if let Some(lm) = &c.last_message {
            if let Some(name) = lm.imdisplayname.as_ref().or(lm.from_display_name.as_ref()) {
                if !name.is_empty() {
                    *name_counts.entry(name.clone()).or_default() += 1;
                }
            }
        }
    }
    let my_name = name_counts
        .into_iter()
        .max_by_key(|(_, count)| *count)
        .map(|(name, _)| name)
        .unwrap_or_default();

    let mut conv_indices: Vec<store::ConvIndex> = Vec::new();
    let mut msg_count: usize = 0;

    for conv in &convs {
        let kind = conv.kind();
        if matches!(kind, ConvKind::System) {
            continue;
        }

        let name = conv.display_name(&my_name);
        let kind_str = format!("{:?}", kind).to_lowercase();

        let meta = store::ConvMeta {
            name: name.clone(),
            kind: kind_str.clone(),
            members: conv.member_names.clone(),
            unread: conv.is_unread(),
            version: conv.version.unwrap_or(0),
            last_message_id: conv.last_message.as_ref().and_then(|lm| lm.id.clone()),
            consumptionhorizon: conv.properties.as_ref().and_then(|p| p.consumptionhorizon.clone()),
        };
        store.save_conv_meta(&conv.id, &meta)?;

        // Fetch messages
        match api.get_messages(&mut auth, &conv.id, 30).await {
            Ok(msg_resp) => {
                for m in &msg_resp.messages {
                    if !matches!(m.messagetype.as_deref(), Some("RichText/Html") | Some("Text")) {
                        continue;
                    }
                    let mid = m.id.as_deref().unwrap_or("unknown");
                    let ts = m.timestamp.as_deref().unwrap_or("");
                    let sender = m.imdisplayname.as_deref()
                        .filter(|s| !s.is_empty())
                        .unwrap_or(if my_name.is_empty() { "?" } else { &my_name });
                    let content = m.content.as_deref().unwrap_or("");
                    store.save_message(&conv.id, mid, ts, sender, "", content)?;
                    msg_count += 1;
                }
            }
            Err(e) => {
                eprintln!("  Warning: messages for '{}': {}", name, e);
            }
        }

        conv_indices.push(store::ConvIndex {
            id: conv.id.clone(),
            name,
            kind: kind_str,
            last_activity: conv.version.unwrap_or(0),
            unread: conv.is_unread(),
        });
    }

    // --- Emails ---
    eprintln!("Fetching mail folders...");
    let folders = api.list_mail_folders(&mut auth).await?;
    let mut email_folder_indices: Vec<store::EmailFolderIndex> = Vec::new();
    let mut email_count: usize = 0;

    for folder in &folders {
        let folder_name = folder.get("displayName").and_then(|v| v.as_str()).unwrap_or("Unknown");
        let folder_id = folder.get("id").and_then(|v| v.as_str()).unwrap_or("");
        let total = folder.get("totalItemCount").and_then(|v| v.as_u64()).unwrap_or(0);

        if total == 0 || folder_id.is_empty() {
            continue;
        }

        match api.list_emails(&mut auth, folder_id, 25).await {
            Ok(emails) => {
                for email in &emails {
                    let eid = email.get("id").and_then(|v| v.as_str()).unwrap_or("unknown");
                    let subject = email.get("subject").and_then(|v| v.as_str()).unwrap_or("(no subject)");
                    let from = email
                        .get("from")
                        .and_then(|v| v.get("emailAddress"))
                        .map(|ea| {
                            let name = ea.get("name").and_then(|v| v.as_str()).unwrap_or("");
                            let addr = ea.get("address").and_then(|v| v.as_str()).unwrap_or("");
                            if name.is_empty() { addr.to_string() } else { format!("{} <{}>", name, addr) }
                        })
                        .unwrap_or_else(|| "?".to_string());
                    let date = email.get("receivedDateTime").and_then(|v| v.as_str()).unwrap_or("");
                    let body = email.get("bodyPreview").and_then(|v| v.as_str()).unwrap_or("");
                    store.save_email(folder_name, eid, &from, date, subject, body)?;
                    email_count += 1;
                }

                email_folder_indices.push(store::EmailFolderIndex {
                    name: folder_name.to_string(),
                    id: folder_id.to_string(),
                    count: emails.len(),
                });
            }
            Err(e) => {
                eprintln!("  Warning: emails for '{}': {}", folder_name, e);
            }
        }
    }

    // --- Save index ---
    let index = store::Index {
        my_name: my_name.clone(),
        conversations: conv_indices,
        email_folders: email_folder_indices,
    };
    store.save_index(&index)?;

    eprintln!(
        "Synced: {} conversations, {} messages, {} email folders, {} emails",
        index.conversations.len(),
        msg_count,
        index.email_folders.len(),
        email_count,
    );
    eprintln!("Data dir: {:?}", store.data_dir());

    Ok(())
}

async fn cmd_list() -> Result<()> {
    let auth = auth::Auth::new()?;
    let store = store::Store::new(auth.config_dir())?;
    let index = store.load_index()?;

    // Conversations
    for conv in &index.conversations {
        let marker = if conv.unread { "*" } else { " " };
        let prefix = match conv.kind.as_str() {
            "channel" => "#",
            "chat" => "@",
            "meeting" => "M",
            _ => "?",
        };
        println!("{}{} {}\t{}", marker, prefix, conv.name, conv.id);
    }

    // Emails
    for folder in &index.email_folders {
        let files = store.list_email_files(&folder.name)?;
        for filepath in &files {
            let content = std::fs::read_to_string(filepath)?;
            let mut from = "";
            let mut subject = "";
            for line in content.lines() {
                if let Some(v) = line.strip_prefix("From: ") {
                    from = v;
                } else if let Some(v) = line.strip_prefix("Subject: ") {
                    subject = v;
                }
                if !from.is_empty() && !subject.is_empty() {
                    break;
                }
            }
            println!("  {} - {} [{}]", from, subject, filepath.display());
        }
    }

    Ok(())
}

async fn cmd_preview(path: &str) -> Result<()> {
    // tv passes rg output lines like: /full/path/file.txt:14:matched text
    // Extract file path (everything before first colon that follows .txt or similar)
    let file_path = extract_file_path(path);
    let file_path = file_path.trim();

    let p = std::path::PathBuf::from(file_path);
    let resolved = if p.exists() {
        p
    } else {
        let auth = auth::Auth::new()?;
        let data_path = auth.config_dir().join("data").join(file_path);
        if data_path.exists() {
            data_path
        } else {
            eprintln!("File not found: {}", file_path);
            return Ok(());
        }
    };

    // If this is a message file inside conversations/{id}/messages/,
    // show all messages from that conversation for context
    if let Some(parent) = resolved.parent() {
        if parent.file_name().map(|n| n == "messages").unwrap_or(false) {
            // Show all messages in this conversation
            let mut files: Vec<_> = std::fs::read_dir(parent)?
                .filter_map(|e| e.ok())
                .map(|e| e.path())
                .filter(|p| p.extension().map(|e| e == "txt").unwrap_or(false))
                .collect();
            files.sort();
            for f in files {
                if let Ok(content) = std::fs::read_to_string(&f) {
                    print!("{}", content);
                }
            }
            return Ok(());
        }
    }

    let content = std::fs::read_to_string(&resolved)?;
    print!("{}", content);
    Ok(())
}

/// Extract file path from rg output line.
/// rg format: /path/to/file.txt:linenum:matched text
/// We find the .txt boundary and take everything before the colon after it.
fn extract_file_path(line: &str) -> &str {
    // Find .txt: pattern which marks end of file path
    if let Some(pos) = line.find(".txt:") {
        &line[..pos + 4] // include .txt
    } else {
        // Fallback: first colon-separated field
        line.split(':').next().unwrap_or(line)
    }
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
    let api = api::Api::new(&auth.region());
    api.send_message(&mut auth, conv_id, text.trim()).await?;
    println!("Sent.");

    Ok(())
}
