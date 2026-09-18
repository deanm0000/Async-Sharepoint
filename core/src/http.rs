use std::sync::Arc;
use std::time::Duration;

use reqwest::header::{ACCEPT, AUTHORIZATION, CONTENT_LENGTH, CONTENT_RANGE, RANGE, RETRY_AFTER};
use reqwest::{Method, Response, StatusCode};
use serde_json::Value;
use tokio::runtime::Handle;

use crate::auth::{CertificateCredential, TokenManager};
use crate::{Error, Result};

const MAX_RETRIES: usize = 6;

pub struct ClientState {
    pub site_url: String,
    pub web_url: String,
    pub tokens: Arc<TokenManager>,
    http: reqwest::Client,
}

impl ClientState {
    pub fn new(
        site_url: impl Into<String>,
        credential: &CertificateCredential,
    ) -> Result<Arc<Self>> {
        let runtime = Handle::try_current().map_err(|_| {
            Error::Unexpected("SharePointClient::new must run inside a Tokio runtime".to_owned())
        })?;
        Self::new_with_runtime(site_url, credential, &runtime)
    }

    pub fn new_with_runtime(
        site_url: impl Into<String>,
        credential: &CertificateCredential,
        runtime: &Handle,
    ) -> Result<Arc<Self>> {
        let site_url = site_url.into().trim_end_matches('/').to_owned();
        let scope = default_scope(&site_url)
            .ok_or_else(|| Error::InvalidUrl(format!("{site_url} is not an absolute URL")))?;
        let http = build_http_client()?;
        Ok(Arc::new(Self {
            web_url: format!("{site_url}/_api/web"),
            tokens: Arc::new(TokenManager::spawn(
                credential,
                scope,
                http.clone(),
                runtime,
            )),
            site_url,
            http,
        }))
    }

    pub fn with_static_token(
        site_url: impl Into<String>,
        token: impl Into<String>,
    ) -> Result<Arc<Self>> {
        let site_url = site_url.into().trim_end_matches('/').to_owned();
        Ok(Arc::new(Self {
            web_url: format!("{site_url}/_api/web"),
            site_url,
            tokens: Arc::new(TokenManager::static_token(token)),
            http: build_http_client()?,
        }))
    }

    pub fn for_site(&self, site_url: impl Into<String>) -> Arc<Self> {
        let site_url = site_url.into().trim_end_matches('/').to_owned();
        Arc::new(Self {
            web_url: format!("{site_url}/_api/web"),
            site_url,
            tokens: Arc::clone(&self.tokens),
            http: self.http.clone(),
        })
    }

    pub async fn request(
        &self,
        method: Method,
        url: &str,
        params: Option<&[(String, String)]>,
        json_body: Option<&Value>,
        content: Option<Vec<u8>>,
        range: Option<(u64, u64)>,
    ) -> Result<Response> {
        let mut stale = None;
        for attempt in 0..=MAX_RETRIES {
            let (generation, token) = self.tokens.get(stale).await?;
            let mut request = self
                .http
                .request(method.clone(), url)
                .header(AUTHORIZATION, format!("Bearer {token}"))
                .header(ACCEPT, "application/json;odata=nometadata");
            if let Some(params) = params {
                request = request.query(params);
            }
            if let Some(json_body) = json_body {
                request = request.json(json_body);
            }
            if let Some(content) = &content {
                // IIS returns 411 Length Required if Content-Length is missing, including for
                // an empty body; reqwest doesn't reliably send it on its own for zero-length bodies.
                request = request
                    .header(CONTENT_LENGTH, content.len())
                    .body(content.clone());
            }
            if let Some((start, end)) = range {
                request = request.header(RANGE, format!("bytes={start}-{end}"));
            }
            let response = request.send().await?;
            let status = response.status();
            if matches!(status, StatusCode::UNAUTHORIZED | StatusCode::FORBIDDEN) && stale.is_none()
            {
                stale = Some(generation);
                continue;
            }
            if status.is_client_error() || status.is_server_error() {
                return Err(Error::Http {
                    status,
                    url: url.to_owned(),
                    body: response.text().await.unwrap_or_default(),
                });
            }
            if !matches!(
                status,
                StatusCode::TOO_MANY_REQUESTS
                    | StatusCode::SERVICE_UNAVAILABLE
                    | StatusCode::GATEWAY_TIMEOUT
            ) || attempt == MAX_RETRIES
            {
                return response.error_for_status().map_err(Error::from);
            }
            let retry_after = response
                .headers()
                .get(RETRY_AFTER)
                .and_then(|value| value.to_str().ok())
                .and_then(|value| value.parse::<f64>().ok());
            tokio::time::sleep(Duration::from_secs_f64(
                retry_after.unwrap_or_else(|| (2_f64.powi(attempt as i32)).min(30.0))
                    + rand::random_range(0.0..0.5),
            ))
            .await;
        }
        Err(Error::Unexpected(
            "retry loop exited unexpectedly".to_owned(),
        ))
    }

    pub async fn get_json(&self, url: &str, params: Option<&[(String, String)]>) -> Result<Value> {
        Ok(self
            .request(Method::GET, url, params, None, None, None)
            .await?
            .json()
            .await?)
    }
    pub async fn post_json(
        &self,
        url: &str,
        json_body: Option<&Value>,
        content: Option<Vec<u8>>,
    ) -> Result<Value> {
        let response = self
            .request(Method::POST, url, None, json_body, content, None)
            .await?;
        let body = response.bytes().await?;
        Ok(if body.is_empty() {
            Value::Object(Default::default())
        } else {
            serde_json::from_slice(&body)?
        })
    }
    pub async fn get_bytes(&self, url: &str) -> Result<Vec<u8>> {
        Ok(self
            .request(Method::GET, url, None, None, None, None)
            .await?
            .bytes()
            .await?
            .to_vec())
    }

    /// Fetch one byte range (`bytes=start-end`, inclusive) of a file's `/$value` endpoint,
    /// returning the chunk plus the total file size when the server reports `Content-Range`.
    pub async fn get_bytes_range(
        &self,
        url: &str,
        start: u64,
        end: u64,
    ) -> Result<(Vec<u8>, Option<u64>)> {
        let response = self
            .request(Method::GET, url, None, None, None, Some((start, end)))
            .await?;
        let total = response
            .headers()
            .get(CONTENT_RANGE)
            .and_then(|value| value.to_str().ok())
            .and_then(|value| value.rsplit('/').next())
            .and_then(|value| value.parse::<u64>().ok());
        Ok((response.bytes().await?.to_vec(), total))
    }
}

fn build_http_client() -> Result<reqwest::Client> {
    Ok(reqwest::Client::builder()
        .timeout(Duration::from_secs(30))
        .build()?)
}
fn default_scope(site_url: &str) -> Option<String> {
    Some(format!(
        "https://{}/.default",
        url::Url::parse(site_url).ok()?.host_str()?
    ))
}
