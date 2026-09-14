use std::sync::Arc;
use std::time::{Duration, Instant};

use pyo3::exceptions::{PyRuntimeError, PyTimeoutError};
use pyo3::prelude::*;
use tokio::sync::Mutex;

const TOKEN_REFRESH_MARGIN: f64 = 60.0;
const MIN_TOKEN_LIFETIME: f64 = 60.0;
const TOKEN_REFRESH_TIMEOUT: Duration = Duration::from_secs(30);

struct TokenState {
    value: Option<String>,
    expires_at: Instant,
}

pub struct TokenManager {
    callback: Arc<Py<PyAny>>,
    state: Mutex<TokenState>,
}

impl TokenManager {
    pub fn new(callback: Py<PyAny>) -> Self {
        Self {
            callback: Arc::new(callback),
            state: Mutex::new(TokenState {
                value: None,
                expires_at: Instant::now(),
            }),
        }
    }

    pub async fn get(&self, force: bool) -> PyResult<String> {
        let mut state = tokio::time::timeout(TOKEN_REFRESH_TIMEOUT, self.state.lock())
            .await
            .map_err(|_| PyTimeoutError::new_err("timed out waiting for token refresh lock"))?;
        if !force
            && let Some(value) = &state.value
            && Instant::now() < state.expires_at
        {
            return Ok(value.clone());
        }

        let callback = Arc::clone(&self.callback);
        let token = tokio::time::timeout(
            TOKEN_REFRESH_TIMEOUT,
            tokio::task::spawn_blocking(move || {
                Python::attach(|py| {
                    let payload = callback.call0(py)?;
                    let payload = payload.bind(py);
                    let value = payload.get_item("access_token")?.extract::<String>()?;
                    let lifetime = payload
                        .get_item("expires_in")
                        .ok()
                        .and_then(|item| item.extract::<f64>().ok())
                        .unwrap_or(0.0);
                    Ok::<_, PyErr>((value, lifetime))
                })
            }),
        )
        .await
        .map_err(|_| PyTimeoutError::new_err("timed out acquiring a SharePoint access token"))?
        .map_err(|error| PyRuntimeError::new_err(format!("token worker failed: {error}")))??;

        let lifetime = (token.1 - TOKEN_REFRESH_MARGIN).max(MIN_TOKEN_LIFETIME);
        state.value = Some(token.0.clone());
        state.expires_at = Instant::now() + Duration::from_secs_f64(lifetime);
        Ok(token.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn token_lifetime_has_a_floor() {
        assert_eq!(
            (0.0_f64 - TOKEN_REFRESH_MARGIN).max(MIN_TOKEN_LIFETIME),
            60.0
        );
        assert_eq!(
            (3600.0_f64 - TOKEN_REFRESH_MARGIN).max(MIN_TOKEN_LIFETIME),
            3540.0
        );
    }
}
