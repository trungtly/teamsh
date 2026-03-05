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
    #[serde(rename = "lastMessage")]
    pub last_message: Option<LastMessage>,
    pub version: Option<u64>,
    pub properties: Option<ConvProperties>,
    #[serde(skip)]
    pub member_names: Vec<String>,
}

#[derive(Debug, Deserialize)]
pub struct ConvProperties {
    pub consumptionhorizon: Option<String>,
    pub lastimreceivedtime: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct ThreadProperties {
    pub topic: Option<String>,
    #[serde(rename = "lastjoinat")]
    pub last_join_at: Option<String>,
    #[serde(rename = "memberCount")]
    pub member_count: Option<String>,
    #[serde(rename = "productThreadType")]
    pub product_thread_type: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct LastMessage {
    pub id: Option<String>,
    #[serde(rename = "fromDisplayNameInToken")]
    pub from_display_name: Option<String>,
    #[serde(rename = "fromGivenNameInToken")]
    pub from_given_name: Option<String>,
    pub imdisplayname: Option<String>,
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

#[derive(Debug, Clone, PartialEq)]
pub enum ConvKind {
    Channel,
    Chat,
    Meeting,
    System,
}

impl Conversation {
    /// Display name for the conversation
    pub fn display_name(&self, my_name: &str) -> String {
        // If there's an explicit topic, use it
        let topic = self.thread_properties
            .as_ref()
            .and_then(|p| p.topic.as_deref())
            .unwrap_or("");

        if !topic.is_empty() {
            return topic.to_string();
        }

        // If we have resolved member names, show them (excluding self)
        if !self.member_names.is_empty() {
            let others: Vec<&str> = self.member_names.iter()
                .filter(|n| !my_name.is_empty() && n.as_str() != my_name)
                .map(|s| s.as_str())
                .collect();
            if others.is_empty() {
                // Chat with self
                return format!("{} (you)", my_name);
            }
            return others.join(", ");
        }

        // Fallback to last message sender (but skip if it's our own name)
        if let Some(lm) = &self.last_message {
            let sender = lm.from_display_name.as_deref()
                .or(lm.imdisplayname.as_deref())
                .unwrap_or("");

            if !sender.is_empty() && (my_name.is_empty() || sender != my_name) {
                return sender.to_string();
            }
        }

        if !my_name.is_empty() {
            return format!("{} (you)", my_name);
        }
        "(unnamed chat)".to_string()
    }

    pub fn topic(&self) -> &str {
        self.thread_properties
            .as_ref()
            .and_then(|p| p.topic.as_deref())
            .unwrap_or("(no topic)")
    }

    /// Check if conversation has unread messages by comparing
    /// lastMessage.id with consumptionhorizon's read-up-to ID
    pub fn is_unread(&self) -> bool {
        let last_msg_id = self.last_message.as_ref()
            .and_then(|lm| lm.id.as_deref())
            .unwrap_or("");
        if last_msg_id.is_empty() { return false; }

        let horizon = self.properties.as_ref()
            .and_then(|p| p.consumptionhorizon.as_deref())
            .unwrap_or("");
        if horizon.is_empty() { return false; }

        // Format: timestamp;readUpToMsgId;timestamp
        let read_up_to = horizon.split(';').nth(1).unwrap_or("");
        if read_up_to.is_empty() { return false; }

        // Compare as numbers (message IDs are timestamps)
        let last: u64 = last_msg_id.parse().unwrap_or(0);
        let read: u64 = read_up_to.parse().unwrap_or(0);
        last > read
    }

    pub fn kind(&self) -> ConvKind {
        let id = &self.id;
        if id.starts_with("48:") {
            ConvKind::System
        } else if id.contains("meeting_") {
            ConvKind::Meeting
        } else if id.contains("@thread.skype") || id.contains("@thread.tacv2") || id.contains("@thread.v2") {
            let topic = self.topic();
            if topic != "(no topic)" {
                ConvKind::Channel
            } else {
                // Check productThreadType for better classification
                let ptype = self.thread_properties.as_ref()
                    .and_then(|p| p.product_thread_type.as_deref())
                    .unwrap_or("");
                if ptype == "GroupChat" || ptype == "chat" {
                    ConvKind::Chat
                } else {
                    ConvKind::Chat
                }
            }
        } else {
            ConvKind::Chat
        }
    }
}
