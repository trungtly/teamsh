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
    /// Search people
    Search {
        /// Search query
        query: String,
    },
    /// Set presence to Available (green) for N hours
    Green {
        /// Duration in hours (default: 8)
        #[arg(default_value = "8")]
        hours: u64,
        /// Keep refreshing presence every 4 minutes
        #[arg(long, short)]
        keep: bool,
    },
    /// List emails from inbox
    Emails {
        /// Number of emails to fetch
        #[arg(long, default_value = "10")]
        last: u32,
    },
    /// Sync conversations and emails to local files
    Sync,
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
