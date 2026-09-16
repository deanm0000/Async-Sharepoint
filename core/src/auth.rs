use std::sync::Arc;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use base64::Engine;
use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use jsonwebtoken::{Algorithm, EncodingKey, Header, encode};
use serde::{Deserialize, Serialize};
use tokio::runtime::Handle;
use tokio::sync::{mpsc, watch};
use uuid::Uuid;

use crate::{Error, Result};

const CLIENT_ASSERTION_TYPE: &str = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";
const ASSERTION_LIFETIME: u64 = 600;
const TOKEN_REFRESH_MARGIN: f64 = 300.0;
const MIN_TOKEN_LIFETIME: f64 = 60.0;
const MAX_REFRESH_NAP: Duration = Duration::from_secs(600);
const FORCED_REFRESH_DEBOUNCE: Duration = Duration::from_secs(5);
const TOKEN_WAIT_TIMEOUT: Duration = Duration::from_secs(60);
const MAX_BACKOFF: f64 = 60.0;

#[derive(Serialize)]
struct ClientAssertion<'a> {
    aud: &'a str,
    iss: &'a str,
    sub: &'a str,
    jti: String,
    iat: u64,
    nbf: u64,
    exp: u64,
}

#[derive(Deserialize)]
struct TokenResponse {
    access_token: String,
    expires_in: f64,
}

#[derive(Deserialize)]
struct TokenErrorResponse {
    error: Option<String>,
    error_description: Option<String>,
}

/// Certificate credentials for the Entra ID client-credentials flow.
pub struct CertificateCredential {
    tenant_id: String,
    client_id: String,
    token_endpoint: String,
    encoding_key: EncodingKey,
    x5t: String,
}

impl CertificateCredential {
    pub fn load(
        tenant_id: impl Into<String>,
        client_id: impl Into<String>,
        private_key_path: &str,
        thumbprint: &str,
    ) -> Result<Self> {
        let tenant_id = tenant_id.into();
        let client_id = client_id.into();
        let pem = std::fs::read_to_string(private_key_path).map_err(|error| {
            Error::Credential(format!(
                "could not read the private key at {private_key_path}: {error}"
            ))
        })?;
        let block = private_key_block(&pem).ok_or_else(|| {
            Error::Credential(format!(
                "no PEM private key block found in {private_key_path}"
            ))
        })?;
        let encoding_key = EncodingKey::from_rsa_pem(block.as_bytes()).map_err(|error| {
            Error::Credential(format!(
                "unsupported private key in {private_key_path}: {error}"
            ))
        })?;
        let digest = hex::decode(thumbprint.replace([':', ' '], "")).map_err(|error| {
            Error::Credential(format!("thumbprint must be hex encoded: {error}"))
        })?;
        Ok(Self {
            token_endpoint: format!(
                "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
            ),
            tenant_id,
            client_id,
            encoding_key,
            x5t: URL_SAFE_NO_PAD.encode(digest),
        })
    }

    pub fn tenant_id(&self) -> &str {
        &self.tenant_id
    }
    pub fn client_id(&self) -> &str {
        &self.client_id
    }

    fn assertion(&self) -> std::result::Result<String, jsonwebtoken::errors::Error> {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        let claims = ClientAssertion {
            aud: &self.token_endpoint,
            iss: &self.client_id,
            sub: &self.client_id,
            jti: Uuid::new_v4().to_string(),
            iat: now,
            nbf: now,
            exp: now + ASSERTION_LIFETIME,
        };
        let mut header = Header::new(Algorithm::RS256);
        header.x5t = Some(self.x5t.clone());
        encode(&header, &claims, &self.encoding_key)
    }
}

fn private_key_block(pem: &str) -> Option<String> {
    let mut collected: Option<Vec<&str>> = None;
    for line in pem.lines().map(str::trim) {
        if line.starts_with("-----BEGIN") && line.contains("PRIVATE KEY") {
            collected = Some(vec![line]);
        } else if let Some(lines) = collected.as_mut() {
            lines.push(line);
            if line.starts_with("-----END") && line.contains("PRIVATE KEY") {
                return Some(lines.join("\n"));
            }
        }
    }
    None
}

#[derive(Clone, Default)]
struct TokenSlot {
    generation: u64,
    token: Option<Arc<str>>,
    error: Option<String>,
}
enum Source {
    Static(Arc<str>),
    Refreshed {
        slot: watch::Receiver<TokenSlot>,
        refresh: mpsc::Sender<()>,
    },
}

pub struct TokenManager {
    source: Source,
}

impl TokenManager {
    pub fn static_token(token: impl Into<String>) -> Self {
        let token = token.into();
        Self {
            source: Source::Static(Arc::from(token)),
        }
    }

    pub fn spawn(
        credential: &CertificateCredential,
        scope: String,
        http: reqwest::Client,
        runtime: &Handle,
    ) -> Self {
        let (sender, slot) = watch::channel(TokenSlot::default());
        let (refresh, refresh_rx) = mpsc::channel(1);
        runtime.spawn(refresher(
            http,
            Arc::new(credential.clone_inner()),
            scope,
            sender,
            refresh_rx,
        ));
        Self {
            source: Source::Refreshed { slot, refresh },
        }
    }

    pub async fn get(&self, stale: Option<u64>) -> Result<(u64, Arc<str>)> {
        match &self.source {
            Source::Static(token) => Ok((0, Arc::clone(token))),
            Source::Refreshed { slot, refresh } => {
                if stale.is_some() {
                    let _ = refresh.try_send(());
                }
                tokio::time::timeout(TOKEN_WAIT_TIMEOUT, wait_for_token(slot.clone(), stale))
                    .await
                    .map_err(|_| {
                        Error::Timeout("timed out acquiring a SharePoint access token".to_owned())
                    })?
            }
        }
    }

    pub async fn wait_ready(&self) -> Result<()> {
        self.get(None).await.map(|_| ())
    }
}

impl CertificateCredential {
    fn clone_inner(&self) -> Self {
        Self {
            tenant_id: self.tenant_id.clone(),
            client_id: self.client_id.clone(),
            token_endpoint: self.token_endpoint.clone(),
            encoding_key: self.encoding_key.clone(),
            x5t: self.x5t.clone(),
        }
    }
}

async fn wait_for_token(
    mut slot: watch::Receiver<TokenSlot>,
    stale: Option<u64>,
) -> Result<(u64, Arc<str>)> {
    loop {
        {
            let current = slot.borrow_and_update();
            if stale.is_none_or(|generation| current.generation > generation) {
                if let Some(error) = &current.error
                    && (stale.is_some() || current.token.is_none())
                {
                    return Err(Error::Token(error.clone()));
                }
                if let Some(token) = &current.token {
                    return Ok((current.generation, Arc::clone(token)));
                }
            }
        }
        if slot.changed().await.is_err() {
            return Err(Error::Token("the token refresher stopped".to_owned()));
        }
    }
}

async fn refresher(
    http: reqwest::Client,
    credential: Arc<CertificateCredential>,
    scope: String,
    sender: watch::Sender<TokenSlot>,
    mut refresh_rx: mpsc::Receiver<()>,
) {
    let mut generation = 0;
    let mut failures = 0;
    loop {
        generation += 1;
        let (lifetime, error) = match fetch_token(&http, &credential, &scope).await {
            Ok((token, expires_in)) => {
                failures = 0;
                if sender
                    .send(TokenSlot {
                        generation,
                        token: Some(token),
                        error: None,
                    })
                    .is_err()
                {
                    return;
                }
                (
                    (expires_in - TOKEN_REFRESH_MARGIN).max(MIN_TOKEN_LIFETIME),
                    None,
                )
            }
            Err(error) => {
                failures += 1;
                if sender
                    .send(TokenSlot {
                        generation,
                        token: sender.borrow().token.clone(),
                        error: Some(error.clone()),
                    })
                    .is_err()
                {
                    return;
                }
                (
                    (2_f64.powi(failures.min(6))).min(MAX_BACKOFF) + rand::random_range(0.0..0.5),
                    Some(error),
                )
            }
        };
        let _ = error;
        let fetched_at = Instant::now();
        let deadline = SystemTime::now() + Duration::from_secs_f64(lifetime);
        while let Ok(remaining) = deadline.duration_since(SystemTime::now()) {
            tokio::select! {
                () = tokio::time::sleep(remaining.min(MAX_REFRESH_NAP)) => {}
                request = refresh_rx.recv() => { if request.is_none() { return; } if fetched_at.elapsed() >= FORCED_REFRESH_DEBOUNCE { break; } }
                () = sender.closed() => return,
            }
        }
    }
}

async fn fetch_token(
    http: &reqwest::Client,
    credential: &CertificateCredential,
    scope: &str,
) -> std::result::Result<(Arc<str>, f64), String> {
    let assertion = credential
        .assertion()
        .map_err(|error| format!("could not sign the client assertion: {error}"))?;
    let response = http
        .post(&credential.token_endpoint)
        .form(&[
            ("grant_type", "client_credentials"),
            ("client_id", credential.client_id.as_str()),
            ("scope", scope),
            ("client_assertion_type", CLIENT_ASSERTION_TYPE),
            ("client_assertion", assertion.as_str()),
        ])
        .send()
        .await
        .map_err(|error| format!("the token request failed: {error}"))?;
    let status = response.status();
    let body = response
        .bytes()
        .await
        .map_err(|error| format!("could not read the token response: {error}"))?;
    if !status.is_success() {
        let detail = serde_json::from_slice::<TokenErrorResponse>(&body)
            .ok()
            .and_then(|payload| payload.error_description.or(payload.error))
            .unwrap_or_else(|| String::from_utf8_lossy(&body).into_owned());
        return Err(format!("Entra ID returned {status}: {detail}"));
    }
    let parsed = serde_json::from_slice::<TokenResponse>(&body)
        .map_err(|error| format!("could not parse the token response: {error}"))?;
    Ok((Arc::from(parsed.access_token), parsed.expires_in))
}
