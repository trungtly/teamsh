use anyhow::{Context, Result};
use base64::Engine;
use serde::Deserialize;
use std::fs;
use std::path::PathBuf;

// Teams web client (SPA) — required for chat API, 24h refresh token
const TEAMS_CLIENT_ID: &str = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
// Microsoft Office native client — for Graph API (emails), ~90 day refresh token
const GRAPH_CLIENT_ID: &str = "d3590ed6-52b3-4102-aeff-aad2292ab01c";

const ORIGIN: &str = "https://teams.cloud.microsoft";
const TEAMS_SCOPE: &str = "https://ic3.teams.office.com/.default openid profile offline_access";
const GRAPH_SCOPE: &str = "https://graph.microsoft.com/.default openid profile offline_access";

#[derive(Deserialize)]
struct TokenResponse {
    access_token: String,
    refresh_token: Option<String>,
    #[allow(dead_code)]
    expires_in: u64,
}

#[derive(Deserialize)]
struct TokenError {
    error: String,
    error_description: String,
}

#[derive(Deserialize)]
struct DeviceCodeResponse {
    device_code: String,
    #[allow(dead_code)]
    user_code: String,
    verification_uri: String,
    message: String,
    interval: Option<u64>,
}

pub struct Auth {
    config_dir: PathBuf,
    tenant_id: String,
    access_token: Option<String>,
    graph_token: Option<String>,
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
            graph_token: None,
        })
    }

    pub fn config_dir(&self) -> &PathBuf {
        &self.config_dir
    }

    pub fn region(&self) -> String {
        Self::load_file(&self.config_dir, "region")
            .unwrap_or_else(|| "au".to_string())
    }

    /// Get Teams access token (for chat/messages API)
    pub async fn access_token(&mut self) -> Result<&str> {
        if self.access_token.is_none() {
            self.refresh_teams().await?;
        }
        Ok(self.access_token.as_ref().unwrap())
    }

    /// Get Graph access token (for emails via Microsoft Graph)
    pub async fn graph_token(&mut self) -> Result<&str> {
        if self.graph_token.is_none() {
            self.refresh_graph().await?;
        }
        Ok(self.graph_token.as_ref().unwrap())
    }

    pub fn clear_graph_token(&mut self) {
        self.graph_token = None;
    }

    pub fn clear_access_token(&mut self) {
        self.access_token = None;
    }

    /// Refresh Teams token using Teams SPA client (24h refresh token)
    pub async fn refresh_teams(&mut self) -> Result<()> {
        let refresh_token = Self::load_file(&self.config_dir, "refresh_token")
            .context("No refresh token. Run: teamsh auth init")?;
        let tenant_id = self.get_tenant_id()?;

        let url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            tenant_id
        );

        let client = reqwest::Client::new();
        let resp = client
            .post(&url)
            .header("Origin", ORIGIN)
            .form(&[
                ("client_id", TEAMS_CLIENT_ID),
                ("grant_type", "refresh_token"),
                ("refresh_token", &refresh_token),
                ("scope", TEAMS_SCOPE),
            ])
            .send()
            .await?;

        let body = resp.text().await?;

        if let Ok(token_resp) = serde_json::from_str::<TokenResponse>(&body) {
            self.access_token = Some(token_resp.access_token);
            if let Some(new_rt) = &token_resp.refresh_token {
                fs::write(self.config_dir.join("refresh_token"), new_rt)?;
            }
            return Ok(());
        }

        let err: TokenError = serde_json::from_str(&body)
            .context("Failed to parse token response")?;
        anyhow::bail!("{} - {}", err.error, err.error_description)
    }

    /// Refresh Graph token using Office native client (~90 day refresh token)
    async fn refresh_graph(&mut self) -> Result<()> {
        let refresh_token = Self::load_file(&self.config_dir, "graph_refresh_token")
            .context("No email token. Run: teamsh auth init")?;
        let tenant_id = self.get_tenant_id()?;

        let url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            tenant_id
        );
        let client = reqwest::Client::new();
        let resp = client
            .post(&url)
            .form(&[
                ("client_id", GRAPH_CLIENT_ID),
                ("grant_type", "refresh_token"),
                ("refresh_token", &refresh_token),
                ("scope", GRAPH_SCOPE),
            ])
            .send()
            .await?;
        let body = resp.text().await?;
        let token_resp: TokenResponse = serde_json::from_str(&body)
            .context(format!("Graph token failed: {}", &body[..body.len().min(300)]))?;
        self.graph_token = Some(token_resp.access_token);
        if let Some(new_rt) = &token_resp.refresh_token {
            fs::write(self.config_dir.join("graph_refresh_token"), new_rt)?;
        }
        Ok(())
    }

    fn get_tenant_id(&mut self) -> Result<String> {
        if self.tenant_id.is_empty() {
            self.tenant_id = Self::load_file(&self.config_dir, "tenant_id")
                .context("No tenant ID. Run: teamsh auth init")?;
        }
        Ok(self.tenant_id.clone())
    }

    /// Login for Teams chat (SPA client, 24h token)
    pub async fn login_teams(&mut self, tenant: &str) -> Result<()> {
        println!("Teams login (chat & messages — expires daily)");
        self.device_code_login(tenant, TEAMS_CLIENT_ID, TEAMS_SCOPE, "refresh_token").await
    }

    /// Login for Graph/emails (Office native client, ~90 day token)
    pub async fn login_graph(&mut self, tenant: &str) -> Result<()> {
        println!("Graph login (emails — expires ~90 days)");
        self.device_code_login(tenant, GRAPH_CLIENT_ID, GRAPH_SCOPE, "graph_refresh_token").await
    }

    async fn device_code_login(
        &mut self,
        tenant: &str,
        client_id: &str,
        scope: &str,
        token_file: &str,
    ) -> Result<()> {
        let tenant = if tenant.is_empty() { "organizations" } else { tenant };

        let device_url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/devicecode",
            tenant
        );

        let client = reqwest::Client::new();
        let resp = client
            .post(&device_url)
            .form(&[
                ("client_id", client_id),
                ("scope", scope),
            ])
            .send()
            .await?;

        let body = resp.text().await?;
        let device: DeviceCodeResponse = serde_json::from_str(&body)
            .context(format!("Device code failed: {}", &body[..body.len().min(300)]))?;

        println!("\n{}\n", device.message);

        if let Err(_) = open::that(&device.verification_uri) {
            println!("Open the URL above in your browser.");
        }

        let token_url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            tenant
        );
        let interval = device.interval.unwrap_or(5);

        loop {
            tokio::time::sleep(std::time::Duration::from_secs(interval)).await;

            let resp = client
                .post(&token_url)
                .form(&[
                    ("client_id", client_id),
                    ("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
                    ("device_code", &*device.device_code),
                    ("scope", scope),
                ])
                .send()
                .await?;

            let body = resp.text().await?;

            if let Ok(token_resp) = serde_json::from_str::<TokenResponse>(&body) {
                let actual_tenant = Self::extract_tenant_from_jwt(&token_resp.access_token)
                    .unwrap_or_else(|| tenant.to_string());

                self.tenant_id = actual_tenant.clone();
                fs::write(self.config_dir.join("tenant_id"), &actual_tenant)?;

                if let Some(ref rt) = token_resp.refresh_token {
                    fs::write(self.config_dir.join(token_file), rt)?;
                    println!("OK!");
                } else {
                    println!("OK (session only, no refresh token)");
                }

                return Ok(());
            }

            let err: serde_json::Value = serde_json::from_str(&body).unwrap_or_default();
            match err.get("error").and_then(|v| v.as_str()).unwrap_or("") {
                "authorization_pending" => {
                    print!(".");
                    std::io::Write::flush(&mut std::io::stdout())?;
                }
                "slow_down" => {
                    tokio::time::sleep(std::time::Duration::from_secs(5)).await;
                }
                "expired_token" => {
                    anyhow::bail!("Login timed out. Try again.");
                }
                other => {
                    let desc = err.get("error_description")
                        .and_then(|v| v.as_str()).unwrap_or("");
                    anyhow::bail!("Login failed: {} - {}", other, desc);
                }
            }
        }
    }

    fn extract_tenant_from_jwt(token: &str) -> Option<String> {
        let parts: Vec<&str> = token.split('.').collect();
        if parts.len() < 2 { return None; }
        let mut encoded = parts[1].to_string();
        while encoded.len() % 4 != 0 { encoded.push('='); }
        let payload = base64::engine::general_purpose::URL_SAFE
            .decode(&encoded).ok()?;
        let json: serde_json::Value = serde_json::from_slice(&payload).ok()?;
        json.get("tid").and_then(|v| v.as_str()).map(|s| s.to_string())
    }

    fn load_file(dir: &PathBuf, name: &str) -> Option<String> {
        fs::read_to_string(dir.join(name))
            .ok()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
    }
}
