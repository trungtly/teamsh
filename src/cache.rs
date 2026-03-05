use anyhow::Result;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

/// Cached conversation metadata
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CachedConv {
    pub id: String,
    pub topic: String,
    pub member_names: Vec<String>,
    pub version: u64,
    pub kind: String, // "Channel", "Chat", "Meeting"
    pub display_name: String,
    pub last_message_id: Option<String>,
    pub consumptionhorizon: Option<String>,
}

/// Full cache stored on disk
#[derive(Debug, Default, Serialize, Deserialize)]
pub struct Cache {
    pub my_name: String,
    pub conversations: Vec<CachedConv>,
    pub snippets: HashMap<String, Vec<String>>,
    /// Timestamp of last full refresh
    pub last_refresh: u64,
}

impl Cache {
    fn path(config_dir: &PathBuf) -> PathBuf {
        config_dir.join("cache.json")
    }

    pub fn load(config_dir: &PathBuf) -> Cache {
        let path = Self::path(config_dir);
        fs::read_to_string(&path)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_default()
    }

    pub fn save(&self, config_dir: &PathBuf) -> Result<()> {
        let path = Self::path(config_dir);
        let json = serde_json::to_string(self)?;
        fs::write(&path, json)?;
        Ok(())
    }

}
