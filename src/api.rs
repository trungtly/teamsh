use anyhow::{Context, Result};
use crate::auth::Auth;
use crate::types::{ConversationsResponse, MessagesResponse};

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
