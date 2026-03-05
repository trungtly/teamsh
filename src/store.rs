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

// --- Store ---

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

        let time = timestamp.get(11..16).unwrap_or(timestamp);
        let text = strip_html(content_html);
        let content = format!("{} {}\n  {}\n", time, sender, text);
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

        // Message roundtrip
        store
            .save_message("c1", "msg1", "2024-01-01T10:30:00Z", "Alice", "<p>Hello</p>")
            .unwrap();
        let msgs = store.load_messages("c1").unwrap();
        assert_eq!(msgs.len(), 1);
        assert!(msgs[0].1.contains("Alice"));
        assert!(msgs[0].1.contains("Hello"));

        // Email roundtrip
        store
            .save_email("inbox", "e1", "bob@test.com", "2024-01-01", "Hi", "<p>Body text</p>")
            .unwrap();
        let files = store.list_email_files("inbox").unwrap();
        assert_eq!(files.len(), 1);

        let _ = fs::remove_dir_all(&tmp);
    }
}
