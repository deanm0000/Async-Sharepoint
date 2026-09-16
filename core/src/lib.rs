//! Async SharePoint REST client primitives.

mod auth;
mod client;
mod error;
mod http;
pub mod odata;

pub use auth::CertificateCredential;
pub use client::SharePointClient;
pub use error::{Error, Result};
pub use http::ClientState;
