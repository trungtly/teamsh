use anyhow::{Context, Result};
use serde::Deserialize;
use std::fs;
use std::path::PathBuf;

const CLIENT_ID: &str = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const ORIGIN: &str = "https://teams.cloud.microsoft";
const SCOPE: &str = "https://ic3.teams.office.com/.default openid profile offline_access";

#[derive(Deserialize)]
struct TokenResponse {
    access_token: String,
    refresh_token: Option<String>,
    expires_in: u64,
}

#[derive(Deserialize)]
struct TokenError {
    error: String,
    error_description: String,
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

    pub fn save_init(&self, refresh_token: &str, tenant_id: &str, region: &str) -> Result<()> {
        fs::write(self.config_dir.join("refresh_token"), refresh_token.trim())?;
        fs::write(self.config_dir.join("tenant_id"), tenant_id.trim())?;
        fs::write(self.config_dir.join("region"), region.trim())?;
        Ok(())
    }

    pub async fn access_token(&mut self) -> Result<&str> {
        if self.access_token.is_none() {
            self.refresh().await?;
        }
        Ok(self.access_token.as_ref().unwrap())
    }

    pub async fn graph_token(&mut self) -> Result<&str> {
        if self.graph_token.is_none() {
            self.refresh_graph().await?;
        }
        Ok(self.graph_token.as_ref().unwrap())
    }

    async fn refresh_graph(&mut self) -> Result<()> {
        let refresh_token = Self::load_file(&self.config_dir, "refresh_token")
            .context("No refresh token found")?;
        if self.tenant_id.is_empty() {
            self.tenant_id = Self::load_file(&self.config_dir, "tenant_id")
                .context("No tenant ID found")?;
        }
        let url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );
        let client = reqwest::Client::new();
        let resp = client
            .post(&url)
            .header("Origin", ORIGIN)
            .form(&[
                ("client_id", CLIENT_ID),
                ("grant_type", "refresh_token"),
                ("refresh_token", &refresh_token),
                ("scope", "https://graph.microsoft.com/.default openid profile offline_access"),
            ])
            .send()
            .await?;
        let body = resp.text().await?;
        let token_resp: TokenResponse = serde_json::from_str(&body)
            .context("Failed to get graph token")?;
        self.graph_token = Some(token_resp.access_token);
        if let Some(new_rt) = &token_resp.refresh_token {
            fs::write(self.config_dir.join("refresh_token"), new_rt)?;
        }
        Ok(())
    }

    pub async fn refresh(&mut self) -> Result<()> {
        let refresh_token = Self::load_file(&self.config_dir, "refresh_token")
            .context("No refresh token found. Run: teamsh auth init")?;

        if self.tenant_id.is_empty() {
            self.tenant_id = Self::load_file(&self.config_dir, "tenant_id")
                .context("No tenant ID found. Run: teamsh auth init")?;
        }

        let url = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );

        let client = reqwest::Client::new();
        let resp = client
            .post(&url)
            .header("Origin", ORIGIN)
            .form(&[
                ("client_id", CLIENT_ID),
                ("grant_type", "refresh_token"),
                ("refresh_token", &refresh_token),
                ("scope", SCOPE),
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
        anyhow::bail!("Token refresh failed: {} - {}", err.error, err.error_description)
    }

    fn load_file(dir: &PathBuf, name: &str) -> Option<String> {
        fs::read_to_string(dir.join(name))
            .ok()
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
    }
}
