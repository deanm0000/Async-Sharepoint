use std::sync::Arc;
use std::time::Duration;

use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use reqwest::header::{ACCEPT, AUTHORIZATION, RETRY_AFTER};
use reqwest::{Method, Response, StatusCode};
use serde_json::Value;

use crate::auth::TokenManager;

const MAX_RETRIES: usize = 6;

pub struct ClientState {
    pub site_url: String,
    pub web_url: String,
    pub tokens: Arc<TokenManager>,
    http: reqwest::Client,
}

impl ClientState {
    pub fn new(site_url: String, get_token: Py<PyAny>) -> PyResult<Arc<Self>> {
        let http = reqwest::Client::builder()
            .timeout(Duration::from_secs(30))
            .build()
            .map_err(runtime_error)?;
        Ok(Arc::new(Self {
            web_url: format!("{site_url}/_api/web"),
            site_url,
            tokens: Arc::new(TokenManager::new(get_token)),
            http,
        }))
    }

    pub fn for_site(&self, site_url: String) -> Arc<Self> {
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
    ) -> PyResult<Response> {
        let mut forced_refresh = false;
        for attempt in 0..=MAX_RETRIES {
            let token = self.tokens.get(forced_refresh).await?;
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
                request = request.body(content.clone());
            }

            let response = request.send().await.map_err(runtime_error)?;
            let status = response.status();
            if matches!(status, StatusCode::UNAUTHORIZED | StatusCode::FORBIDDEN) && !forced_refresh
            {
                forced_refresh = true;
                continue;
            }
            if !matches!(
                status,
                StatusCode::TOO_MANY_REQUESTS
                    | StatusCode::SERVICE_UNAVAILABLE
                    | StatusCode::GATEWAY_TIMEOUT
            ) || attempt == MAX_RETRIES
            {
                return response.error_for_status().map_err(runtime_error);
            }

            let retry_after = response
                .headers()
                .get(RETRY_AFTER)
                .and_then(|value| value.to_str().ok())
                .and_then(|value| value.parse::<f64>().ok());
            let delay = retry_after.unwrap_or_else(|| (2_f64.powi(attempt as i32)).min(30.0));
            tokio::time::sleep(Duration::from_secs_f64(
                delay + rand::random_range(0.0..0.5),
            ))
            .await;
        }
        Err(PyRuntimeError::new_err("retry loop exited unexpectedly"))
    }

    pub async fn get_json(
        &self,
        url: &str,
        params: Option<&[(String, String)]>,
    ) -> PyResult<Value> {
        self.request(Method::GET, url, params, None, None)
            .await?
            .json()
            .await
            .map_err(runtime_error)
    }

    pub async fn post_json(
        &self,
        url: &str,
        json_body: Option<&Value>,
        content: Option<Vec<u8>>,
    ) -> PyResult<Value> {
        let response = self
            .request(Method::POST, url, None, json_body, content)
            .await?;
        if response.content_length() == Some(0) {
            return Ok(Value::Object(Default::default()));
        }
        let body = response.bytes().await.map_err(runtime_error)?;
        if body.is_empty() {
            return Ok(Value::Object(Default::default()));
        }
        serde_json::from_slice(&body).map_err(runtime_error)
    }

    pub async fn get_bytes(&self, url: &str) -> PyResult<Vec<u8>> {
        self.request(Method::GET, url, None, None, None)
            .await?
            .bytes()
            .await
            .map(|bytes| bytes.to_vec())
            .map_err(runtime_error)
    }
}

pub fn runtime_error(error: impl std::fmt::Display) -> PyErr {
    PyRuntimeError::new_err(error.to_string())
}
