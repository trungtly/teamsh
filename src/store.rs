use std::fs;
use std::path::{Path, PathBuf};

use anyhow::Result;
use serde::{Deserialize, Serialize};

use crate::html::strip_html;

// --- Index types ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConvIndex {
    pub id: String,
    pub name: String,
    pub kind: String,
    pub last_activity: u64,
    pub unread: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EmailFolderIndex {
    pub name: String,
    pub id: String,
    pub count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Index {
    pub my_name: String,
    pub conversations: Vec<ConvIndex>,
    pub email_folders: Vec<EmailFolderIndex>,
}

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

// --- Shared formatting ---

/// Format an ISO 8601 timestamp to "2026 Mar 05 14:32" style.
/// Used by both file storage and TUI rendering for consistency.
pub fn format_timestamp(ts: &str) -> String {
    // Input: "2026-03-05T14:32:00.000Z"
    if ts.len() < 16 {
        return "??:??".to_string();
    }
    let year = ts.get(0..4).unwrap_or("????");
    let month = match ts.get(5..7) {
        Some("01") => "Jan", Some("02") => "Feb", Some("03") => "Mar",
        Some("04") => "Apr", Some("05") => "May", Some("06") => "Jun",
        Some("07") => "Jul", Some("08") => "Aug", Some("09") => "Sep",
        Some("10") => "Oct", Some("11") => "Nov", Some("12") => "Dec",
        _ => "???",
    };
    let day = ts.get(8..10).unwrap_or("??");
    let time = ts.get(11..16).unwrap_or("??:??");
    format!("{} {} {} {}", year, month, day, time)
}

/// Format a message into the standard text format used by both
/// file storage and TUI rendering. This is the single source of truth.
///
/// Output:
/// ```text
/// Alice  Mar 05 14:32  #general
///   Hey @Bob check this out
/// ```
pub fn format_message(sender: &str, timestamp: &str, channel: &str, body: &str) -> String {
    let time = format_timestamp(timestamp);
    let tag = if channel.is_empty() {
        String::new()
    } else {
        format!("  #[{}]", channel)
    };
    let mut out = format!("{}  {}{}  \n", sender, time, tag);
    for line in body.lines() {
        out.push_str(&format!("  {}\n", line));
    }
    out
}

/// Format a message from HTML content (convenience wrapper).
pub fn format_message_html(sender: &str, timestamp: &str, channel: &str, content_html: &str) -> String {
    let text = strip_html(content_html);
    format_message(sender, timestamp, channel, &text)
}

// --- Store ---

pub struct Store {
    data_dir: PathBuf,
}

impl Store {
    pub fn new(config_dir: &Path) -> Result<Self> {
        let data_dir = config_dir.join("data");
        fs::create_dir_all(data_dir.join("conversations"))?;
        fs::create_dir_all(data_dir.join("emails"))?;
        let store = Self { data_dir };
        store.ensure_tv_assets(config_dir)?;
        Ok(store)
    }

    /// Write the tv-preview.sh script and bat syntax to config dirs.
    /// Always overwrites so updates are picked up automatically.
    fn ensure_tv_assets(&self, config_dir: &Path) -> Result<()> {
        // tv preview script
        let script = config_dir.join("tv-preview.sh");
        fs::write(&script, include_str!("tv-preview.sh"))?;
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            fs::set_permissions(&script, fs::Permissions::from_mode(0o755))?;
        }

        // bat custom syntax and theme for Teams messages
        if let Some(home) = dirs::home_dir() {
            let bat_config = home.join(".config/bat");
            let syntax_dir = bat_config.join("syntaxes");
            let theme_dir = bat_config.join("themes");
            fs::create_dir_all(&syntax_dir)?;
            fs::create_dir_all(&theme_dir)?;

            let syntax_file = syntax_dir.join("TeamshMessage.sublime-syntax");
            let theme_file = theme_dir.join("teamsh.tmTheme");
            let new_syntax = include_str!("TeamshMessage.sublime-syntax");
            let new_theme = include_str!("teamsh.tmTheme");

            let syntax_changed = fs::read_to_string(&syntax_file)
                .map(|old| old != new_syntax)
                .unwrap_or(true);
            let theme_changed = fs::read_to_string(&theme_file)
                .map(|old| old != new_theme)
                .unwrap_or(true);

            if syntax_changed || theme_changed {
                fs::write(&syntax_file, new_syntax)?;
                fs::write(&theme_file, new_theme)?;
                let _ = std::process::Command::new("bat")
                    .args(["cache", "--build"])
                    .stdout(std::process::Stdio::null())
                    .stderr(std::process::Stdio::null())
                    .status();
            }
        }

        Ok(())
    }

    pub fn data_dir(&self) -> &Path {
        &self.data_dir
    }

    // --- Index ---

    pub fn load_index(&self) -> Result<Index> {
        let path = self.data_dir.join("index.json");
        if !path.exists() {
            return Ok(Index::default());
        }
        let data = fs::read_to_string(&path)?;
        Ok(serde_json::from_str(&data)?)
    }

    pub fn save_index(&self, index: &Index) -> Result<()> {
        let path = self.data_dir.join("index.json");
        let json = serde_json::to_string_pretty(index)?;
        fs::write(&path, json)?;
        Ok(())
    }

    // --- Favourites ---

    pub fn load_favourites(&self) -> Result<Vec<String>> {
        let path = self.data_dir.join("favourites.json");
        if !path.exists() {
            return Ok(Vec::new());
        }
        let data = fs::read_to_string(&path)?;
        Ok(serde_json::from_str(&data)?)
    }

    pub fn save_favourites(&self, favourites: &[String]) -> Result<()> {
        let path = self.data_dir.join("favourites.json");
        let json = serde_json::to_string_pretty(favourites)?;
        fs::write(&path, json)?;
        Ok(())
    }

    // --- Conversation metadata ---

    pub fn save_conv_meta(&self, conv_id: &str, meta: &ConvMeta) -> Result<()> {
        let dir = self.data_dir.join("conversations").join(safe_filename(conv_id));
        fs::create_dir_all(&dir)?;
        let path = dir.join("meta.json");
        let json = serde_json::to_string_pretty(meta)?;
        fs::write(&path, json)?;
        Ok(())
    }

    pub fn load_conv_meta(&self, conv_id: &str) -> Result<Option<ConvMeta>> {
        let path = self
            .data_dir
            .join("conversations")
            .join(safe_filename(conv_id))
            .join("meta.json");
        if !path.exists() {
            return Ok(None);
        }
        let data = fs::read_to_string(&path)?;
        Ok(Some(serde_json::from_str(&data)?))
    }

    // --- Messages ---

    pub fn save_message(
        &self,
        conv_id: &str,
        msg_id: &str,
        timestamp: &str,
        sender: &str,
        channel: &str,
        content_html: &str,
    ) -> Result<()> {
        let dir = self
            .data_dir
            .join("conversations")
            .join(safe_filename(conv_id))
            .join("messages");
        fs::create_dir_all(&dir)?;

        let filename = format!("{}.txt", safe_filename(msg_id));
        let path = dir.join(&filename);

        let content = format_message_html(sender, timestamp, channel, content_html);
        fs::write(&path, content)?;
        Ok(())
    }

    pub fn load_messages(&self, conv_id: &str) -> Result<Vec<(String, String)>> {
        let dir = self
            .data_dir
            .join("conversations")
            .join(safe_filename(conv_id))
            .join("messages");
        if !dir.exists() {
            return Ok(Vec::new());
        }

        let mut entries: Vec<(String, String)> = Vec::new();
        for entry in fs::read_dir(&dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension().map_or(false, |e| e == "txt") {
                let filename = entry.file_name().to_string_lossy().to_string();
                let content = fs::read_to_string(&path)?;
                entries.push((filename, content));
            }
        }
        entries.sort_by(|a, b| a.0.cmp(&b.0));
        Ok(entries)
    }

    // --- Emails ---

    pub fn save_email(
        &self,
        folder_name: &str,
        email_id: &str,
        from: &str,
        date: &str,
        subject: &str,
        body_html: &str,
    ) -> Result<()> {
        let dir = self.data_dir.join("emails").join(safe_filename(folder_name));
        fs::create_dir_all(&dir)?;

        let filename = format!("{}.txt", safe_filename(email_id));
        let path = dir.join(&filename);

        let body = strip_html(body_html);
        let content = format!("From: {}\nDate: {}\nSubject: {}\n\n{}\n", from, date, subject, body);
        fs::write(&path, content)?;
        Ok(())
    }

    pub fn list_email_files(&self, folder_name: &str) -> Result<Vec<PathBuf>> {
        let dir = self.data_dir.join("emails").join(safe_filename(folder_name));
        if !dir.exists() {
            return Ok(Vec::new());
        }

        let mut files: Vec<PathBuf> = Vec::new();
        for entry in fs::read_dir(&dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension().map_or(false, |e| e == "txt") {
                files.push(path);
            }
        }
        files.sort();
        Ok(files)
    }

    // --- Dummy test data ---

    /// Create a dummy dataset for testing tv search with @people and #channel tags.
    /// Uses fake names only. Placed under data/conversations/ like real data.
    pub fn create_dummy_data(&self) -> Result<()> {
        let base_ts = 1700000000000u64; // base timestamp for filenames
        let channels = [
            ("test_general", "general"),
            ("test_engineering", "engineering"),
            ("test_random", "random"),
        ];

        let messages: Vec<(&str, &str, &str, &str, u64)> = vec![
            // (conv_id, sender, channel, body, offset_ms)
            ("test_general", "Alice", "general", "Hey team, standup in 5 mins", 1000),
            ("test_general", "Bob", "general", "Sure, I'll be there. @Alice did you see the new PR?", 2000),
            ("test_general", "Charlie", "general", "Good morning everyone! @Alice @Bob ready when you are", 3000),
            ("test_general", "Alice", "general", "Let's go. @Charlie can you share your screen?", 4000),
            ("test_engineering", "Diana", "engineering", "Deployed v2.1 to staging. @Bob please review the API changes", 5000),
            ("test_engineering", "Bob", "engineering", "On it. The new endpoint looks good. @Diana one question about auth", 6000),
            ("test_engineering", "Eve", "engineering", "I found a bug in the search indexer. Creating a ticket now", 7000),
            ("test_engineering", "Diana", "engineering", "@Eve can you check if it's related to the migration we did yesterday?", 8000),
            ("test_engineering", "Frank", "engineering", "The CI pipeline is green again. @Eve @Diana the fix for #1234 is merged", 9000),
            ("test_random", "Charlie", "random", "Anyone want coffee? Going to the kitchen", 10000),
            ("test_random", "Eve", "random", "Yes please! @Charlie large latte for me", 11000),
            ("test_random", "Alice", "random", "Team lunch today? @Bob @Charlie @Diana @Eve @Frank", 12000),
            ("test_random", "Frank", "random", "I'm in! @Alice where are we going?", 13000),
            ("test_random", "Bob", "random", "How about the new place on 5th? @Alice @Frank", 14000),
        ];

        // Create conv directories with meta
        for (conv_id, channel) in &channels {
            let meta = ConvMeta {
                name: channel.to_string(),
                kind: "channel".to_string(),
                members: vec![],
                unread: false,
                version: 1,
                last_message_id: None,
                consumptionhorizon: None,
            };
            self.save_conv_meta(conv_id, &meta)?;
        }

        // Write messages with timestamps spread across dates
        let dates = [
            "2026-03-03T09:15:00.000Z",
            "2026-03-03T09:16:00.000Z",
            "2026-03-03T09:17:00.000Z",
            "2026-03-03T09:18:00.000Z",
            "2026-03-04T11:30:00.000Z",
            "2026-03-04T11:35:00.000Z",
            "2026-03-04T14:20:00.000Z",
            "2026-03-04T14:25:00.000Z",
            "2026-03-04T16:00:00.000Z",
            "2026-03-05T08:45:00.000Z",
            "2026-03-05T08:47:00.000Z",
            "2026-03-05T12:00:00.000Z",
            "2026-03-05T12:05:00.000Z",
            "2026-03-05T12:10:00.000Z",
        ];

        for (i, (conv_id, sender, channel, body, offset)) in messages.iter().enumerate() {
            let msg_id = format!("{}", base_ts + offset);
            let ts = dates.get(i).unwrap_or(&"2026-03-05T12:00:00.000Z");
            // Write using format_message directly (body is already plain text)
            let dir = self
                .data_dir
                .join("conversations")
                .join(safe_filename(conv_id))
                .join("messages");
            fs::create_dir_all(&dir)?;
            let filename = format!("{}.txt", safe_filename(&msg_id));
            let path = dir.join(&filename);
            let content = format_message(sender, ts, channel, body);
            fs::write(&path, content)?;
        }

        Ok(())
    }
}

/// Replace characters that are unsafe for filenames with underscores.
pub fn safe_filename(id: &str) -> String {
    id.chars()
        .map(|c| match c {
            '/' | '\\' | ':' | '?' | '*' | '"' | '<' | '>' | '|' => '_',
            _ => c,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_safe_filename() {
        assert_eq!(safe_filename("abc/def\\ghi:jkl"), "abc_def_ghi_jkl");
        assert_eq!(safe_filename("normal-id_123"), "normal-id_123");
        assert_eq!(safe_filename("a?b*c\"d<e>f|g"), "a_b_c_d_e_f_g");
    }

    #[test]
    fn test_format_timestamp() {
        assert_eq!(format_timestamp("2026-03-05T14:32:00.000Z"), "2026 Mar 05 14:32");
        assert_eq!(format_timestamp("2026-12-25T08:00:00Z"), "2026 Dec 25 08:00");
        assert_eq!(format_timestamp("short"), "??:??");
    }

    #[test]
    fn test_format_message() {
        let msg = format_message("Alice", "2026-03-05T14:32:00.000Z", "general", "Hello @Bob");
        assert!(msg.starts_with("Alice  2026 Mar 05 14:32  #[general]"), "got: {}", msg);
        assert!(msg.contains("  Hello @Bob"), "got: {}", msg);
    }

    #[test]
    fn test_format_message_no_channel() {
        let msg = format_message("Alice", "2026-03-05T14:32:00.000Z", "", "Hello");
        assert!(msg.starts_with("Alice  2026 Mar 05 14:32"), "got: {}", msg);
        assert!(!msg.contains('#'), "should have no channel tag: {}", msg);
    }

    #[test]
    fn test_store_roundtrip() {
        let tmp = std::env::temp_dir().join("teamsh_test_store");
        let _ = fs::remove_dir_all(&tmp);

        let store = Store::new(&tmp).unwrap();

        // Index roundtrip
        let index = Index {
            my_name: "Test User".into(),
            conversations: vec![ConvIndex {
                id: "c1".into(),
                name: "General".into(),
                kind: "channel".into(),
                last_activity: 1000,
                unread: false,
            }],
            email_folders: vec![],
        };
        store.save_index(&index).unwrap();
        let loaded = store.load_index().unwrap();
        assert_eq!(loaded.my_name, "Test User");
        assert_eq!(loaded.conversations.len(), 1);

        // Message roundtrip - now with channel
        store
            .save_message("c1", "msg1", "2024-01-01T10:30:00Z", "Alice", "general", "<p>Hello @Bob</p>")
            .unwrap();
        let msgs = store.load_messages("c1").unwrap();
        assert_eq!(msgs.len(), 1);
        assert!(msgs[0].1.contains("Alice"), "should have sender: {}", msgs[0].1);
        assert!(msgs[0].1.contains("2024 Jan 01 10:30"), "should have full timestamp: {}", msgs[0].1);
        assert!(msgs[0].1.contains("#[general]"), "should have channel: {}", msgs[0].1);
        assert!(msgs[0].1.contains("@Bob"), "should have mention: {}", msgs[0].1);

        // Email roundtrip
        store
            .save_email("inbox", "e1", "bob@test.com", "2024-01-01", "Hi", "<p>Body text</p>")
            .unwrap();
        let files = store.list_email_files("inbox").unwrap();
        assert_eq!(files.len(), 1);

        let _ = fs::remove_dir_all(&tmp);
    }

    #[test]
    fn test_dummy_data() {
        let tmp = std::env::temp_dir().join("teamsh_test_dummy");
        let _ = fs::remove_dir_all(&tmp);

        let store = Store::new(&tmp).unwrap();
        store.create_dummy_data().unwrap();

        // Verify files exist
        let msgs = store.load_messages("test_general").unwrap();
        assert_eq!(msgs.len(), 4, "general should have 4 messages");

        let msgs = store.load_messages("test_engineering").unwrap();
        assert_eq!(msgs.len(), 5, "engineering should have 5 messages");

        let msgs = store.load_messages("test_random").unwrap();
        assert_eq!(msgs.len(), 5, "random should have 5 messages");

        // Verify format - first message
        let first = &msgs[0].1;
        assert!(first.contains("#[random]"), "should have channel tag: {}", first);

        // Verify @mentions are searchable
        let all_text: String = msgs.iter().map(|m| m.1.as_str()).collect();
        assert!(all_text.contains("@Alice"), "should find @Alice");
        assert!(all_text.contains("@Bob"), "should find @Bob");
        assert!(all_text.contains("@Charlie"), "should find @Charlie");

        let _ = fs::remove_dir_all(&tmp);
    }
}
