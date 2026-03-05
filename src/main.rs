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
        Some(Commands::Tui) | None => {
            tui::run().await?;
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

    println!("Enter your region (default: au):");
    let mut region = String::new();
    std::io::stdin().read_line(&mut region)?;
    let region = region.trim();
    let region = if region.is_empty() { "au" } else { region };

    auth.save_init(refresh_token, tenant_id, region)?;
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
