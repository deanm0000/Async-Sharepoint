use std::fmt::{Display, Formatter};

/// Errors returned by the SharePoint client.
#[derive(Debug)]
pub enum Error {
    InvalidUrl(String),
    Credential(String),
    Token(String),
    Timeout(String),
    Transport(reqwest::Error),
    Http {
        status: reqwest::StatusCode,
        url: String,
        body: String,
    },
    Json(serde_json::Error),
    Unexpected(String),
}

impl Display for Error {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidUrl(message)
            | Self::Credential(message)
            | Self::Token(message)
            | Self::Timeout(message)
            | Self::Unexpected(message) => formatter.write_str(message),
            Self::Transport(error) => Display::fmt(error, formatter),
            Self::Http { status, url, body } => {
                write!(formatter, "SharePoint API {status} for {url}: {body}")
            }
            Self::Json(error) => Display::fmt(error, formatter),
        }
    }
}

impl std::error::Error for Error {}

impl From<reqwest::Error> for Error {
    fn from(error: reqwest::Error) -> Self {
        Self::Transport(error)
    }
}

impl From<serde_json::Error> for Error {
    fn from(error: serde_json::Error) -> Self {
        Self::Json(error)
    }
}

pub type Result<T> = std::result::Result<T, Error>;

#[cfg(feature = "python")]
impl From<Error> for pyo3::PyErr {
    fn from(error: Error) -> Self {
        pyo3::exceptions::PyRuntimeError::new_err(error.to_string())
    }
}
