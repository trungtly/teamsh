use anyhow::{Context, Result};
use crate::auth::Auth;
use crate::types::{ConversationsResponse, MessagesResponse};

/// Generate a random UUID v4 string
fn uuid_v4() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let seed = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    // Simple pseudo-random UUID v4
    let mut bytes = [0u8; 16];
    let mut s = seed;
    for b in bytes.iter_mut() {
        s = s.wrapping_mul(6364136223846793005).wrapping_add(1);
        *b = (s >> 33) as u8;
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x40; // version 4
    bytes[8] = (bytes[8] & 0x3f) | 0x80; // variant 1
    format!(
        "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        bytes[0], bytes[1], bytes[2], bytes[3],
        bytes[4], bytes[5], bytes[6], bytes[7],
        bytes[8], bytes[9], bytes[10], bytes[11],
        bytes[12], bytes[13], bytes[14], bytes[15],
    )
}

const BASE_URL: &str = "https://teams.cloud.microsoft/api/chatsvc";

pub struct Api {
    client: reqwest::Client,
    region: String,
}

impl Api {
    pub fn new(region: &str) -> Self {
        Self {
            client: reqwest::Client::new(),
            region: region.to_string(),
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

    /// Search people via Microsoft Graph API
    pub async fn search_people(&self, auth: &mut Auth, query: &str) -> Result<Vec<(String, String)>> {
        let token = auth.graph_token().await?;
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/people?$search=\"{}\"&$top=10",
            urlencoding::encode(query)
        );
        let resp = self.client
            .get(&url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await?;

        let status = resp.status();
        let text = resp.text().await?;

        if !status.is_success() {
            eprintln!("Graph people search error ({}): {}", status, &text[..text.len().min(200)]);
            return Ok(Vec::new());
        }

        let data: serde_json::Value = serde_json::from_str(&text).unwrap_or_default();
        let mut results = Vec::new();
        if let Some(people) = data.get("value").and_then(|v| v.as_array()) {
            for person in people {
                let name = person.get("displayName").and_then(|v| v.as_str()).unwrap_or("").to_string();
                let email = person.get("scoredEmailAddresses")
                    .and_then(|v| v.as_array())
                    .and_then(|a| a.first())
                    .and_then(|v| v.get("address"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("").to_string();
                if !name.is_empty() {
                    results.push((name, email));
                }
            }
        }
        Ok(results)
    }

    /// Register endpoint and set presence to Available (green)
    pub async fn set_available(&self, auth: &mut Auth, hours: u64) -> Result<()> {
        let token = auth.access_token().await?;
        let ups_base = format!("https://teams.cloud.microsoft/ups/{}", self.region);
        let endpoint_id = uuid_v4();

        // Step 1: Register endpoint
        let reg_url = format!("{}/v1/me/endpoints/", ups_base);
        let reg_body = serde_json::json!({
            "id": endpoint_id,
            "availability": "Available",
            "activity": "Available",
            "activityReporting": "Transport",
            "deviceType": "Web",
        });
        let resp = self.client
            .put(&reg_url)
            .header("Authorization", format!("Bearer {}", token))
            .header("behavioroverride", "redirectAs404")
            .header("x-ms-client-user-agent", "Teams-V2-Web")
            .json(&reg_body)
            .send()
            .await?;
        if !resp.status().is_success() {
            let status = resp.status();
            let text = resp.text().await.unwrap_or_default();
            anyhow::bail!("Endpoint register failed ({}): {}", status, &text[..text.len().min(300)]);
        }

        // Step 2: Force availability
        let force_url = format!("{}/v1/me/forceavailability/", ups_base);
        let expiry = chrono::Utc::now() + chrono::Duration::hours(hours as i64);
        let force_body = serde_json::json!({
            "availability": "Available",
            "activity": "Available",
            "expiry": expiry.to_rfc3339(),
        });
        let token = auth.access_token().await?;
        let resp = self.client
            .put(&force_url)
            .header("Authorization", format!("Bearer {}", token))
            .header("behavioroverride", "redirectAs404")
            .header("x-ms-client-user-agent", "Teams-V2-Web")
            .header("x-ms-endpoint-id", &endpoint_id)
            .json(&force_body)
            .send()
            .await?;
        if !resp.status().is_success() {
            let status = resp.status();
            let text = resp.text().await.unwrap_or_default();
            anyhow::bail!("Force availability failed ({}): {}", status, &text[..text.len().min(300)]);
        }
        Ok(())
    }

    /// Report activity to keep presence alive
    pub async fn report_activity(&self, auth: &mut Auth) -> Result<()> {
        let token = auth.access_token().await?;
        let url = format!(
            "https://teams.cloud.microsoft/ups/{}/v1/me/reportmyactivity/",
            self.region
        );
        let resp = self.client
            .post(&url)
            .header("Authorization", format!("Bearer {}", token))
            .header("behavioroverride", "redirectAs404")
            .header("x-ms-client-user-agent", "Teams-V2-Web")
            .send()
            .await?;
        if !resp.status().is_success() {
            let status = resp.status();
            let text = resp.text().await.unwrap_or_default();
            anyhow::bail!("Report activity failed ({}): {}", status, &text[..text.len().min(200)]);
        }
        Ok(())
    }

    // --- Email (Microsoft Graph API) ---

    /// List emails from inbox (or a folder)
    pub async fn list_emails(&self, auth: &mut Auth, folder: &str, top: u32) -> Result<Vec<serde_json::Value>> {
        let token = auth.graph_token().await?;
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/mailFolders/{}/messages?$top={}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,importance&$orderby=receivedDateTime desc",
            urlencoding::encode(folder), top
        );
        let resp = self.client
            .get(&url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await?;
        let status = resp.status();
        let text = resp.text().await?;
        if !status.is_success() {
            anyhow::bail!("List emails failed ({}): {}", status, &text[..text.len().min(300)]);
        }
        let data: serde_json::Value = serde_json::from_str(&text)?;
        let emails = data.get("value")
            .and_then(|v| v.as_array())
            .cloned()
            .unwrap_or_default();
        Ok(emails)
    }

    /// Get a single email's full body
    pub async fn get_email(&self, auth: &mut Auth, message_id: &str) -> Result<serde_json::Value> {
        let token = auth.graph_token().await?;
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/messages/{}?$select=id,subject,from,toRecipients,ccRecipients,body,receivedDateTime,isRead,hasAttachments,importance",
            urlencoding::encode(message_id)
        );
        let resp = self.client
            .get(&url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await?;
        let status = resp.status();
        let text = resp.text().await?;
        if !status.is_success() {
            anyhow::bail!("Get email failed ({}): {}", status, &text[..text.len().min(300)]);
        }
        let data: serde_json::Value = serde_json::from_str(&text)?;
        Ok(data)
    }

    /// Search emails
    pub async fn search_emails(&self, auth: &mut Auth, query: &str, top: u32) -> Result<Vec<serde_json::Value>> {
        let token = auth.graph_token().await?;
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/messages?$search=\"{}\"&$top={}&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview",
            urlencoding::encode(query), top
        );
        let resp = self.client
            .get(&url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await?;
        let status = resp.status();
        let text = resp.text().await?;
        if !status.is_success() {
            anyhow::bail!("Search emails failed ({}): {}", status, &text[..text.len().min(300)]);
        }
        let data: serde_json::Value = serde_json::from_str(&text)?;
        let emails = data.get("value")
            .and_then(|v| v.as_array())
            .cloned()
            .unwrap_or_default();
        Ok(emails)
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
