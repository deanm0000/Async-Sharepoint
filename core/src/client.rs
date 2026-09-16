use std::sync::Arc;

use serde_json::{Value, json};

use crate::odata;
use crate::{CertificateCredential, ClientState, Error, Result};

/// High-level asynchronous SharePoint client for Rust applications.
pub struct SharePointClient {
    state: Arc<ClientState>,
}

impl SharePointClient {
    pub fn new(site_url: impl Into<String>, credential: &CertificateCredential) -> Result<Self> {
        Ok(Self {
            state: ClientState::new(site_url, credential)?,
        })
    }

    pub fn from_static_token(
        site_url: impl Into<String>,
        token: impl Into<String>,
    ) -> Result<Self> {
        Ok(Self {
            state: ClientState::with_static_token(site_url, token)?,
        })
    }

    pub fn site_url(&self) -> &str {
        &self.state.site_url
    }

    /// Return the file metadata for a server-relative path or an AllItems browser URL.
    pub async fn get_file(&self, path: &str) -> Result<Value> {
        let path = normalize_path(path)?;
        let url = format!(
            "{}/GetFileByServerRelativePath(DecodedUrl={})",
            self.state.web_url,
            odata::literal(&path)
        );
        match self.state.get_json(&url, None).await {
            Ok(file) => Ok(file),
            Err(Error::Http { status, .. }) if status == reqwest::StatusCode::NOT_FOUND => Err(
                Error::Unexpected(format!("Either {path} does not exist or is a folder")),
            ),
            Err(error) => Err(error),
        }
    }

    /// Download a file from a server-relative path or an AllItems browser URL.
    pub async fn download(&self, path: &str) -> Result<Vec<u8>> {
        let path = normalize_path(path)?;
        let url = format!(
            "{}/GetFileByServerRelativePath(DecodedUrl={})/$value",
            self.state.web_url,
            odata::literal(&path)
        );
        self.state.get_bytes(&url).await
    }

    /// List direct children of `path`, or of the default document library when omitted.
    /// If `path` names a file, returns that single file, matching `ls file` behavior.
    pub async fn ls(&self, path: Option<&str>) -> Result<Vec<Value>> {
        let library = self.default_document_library().await?;
        let list_id = odata::string(&library, "Id").ok_or_else(|| {
            Error::Unexpected("DefaultDocumentLibrary response has no Id".to_owned())
        })?;
        let path = match path {
            Some(path) => normalize_path(path)?,
            None => {
                let root_url = format!(
                    "{}/lists/GetById({})/RootFolder",
                    self.state.web_url,
                    odata::literal(&list_id)
                );
                let root = self.state.get_json(&root_url, None).await?;
                odata::string(&root, "ServerRelativeUrl").ok_or_else(|| {
                    Error::Unexpected(
                        "DefaultDocumentLibrary root has no ServerRelativeUrl".to_owned(),
                    )
                })?
            }
        };
        let list_url = format!(
            "{}/lists/GetById({})",
            self.state.web_url,
            odata::literal(&list_id)
        );
        let view = format!(
            "<View Scope=\"RecursiveAll\"><ViewFields><FieldRef Name=\"FileRef\" /><FieldRef Name=\"FileLeafRef\" /><FieldRef Name=\"FSObjType\" /><FieldRef Name=\"FileDirRef\" /></ViewFields><Query><Where><Eq><FieldRef Name=\"FileDirRef\" /><Value Type=\"Text\">{path}</Value></Eq></Where></Query><RowLimit>500</RowLimit></View>"
        );
        let body = json!({"query": {"ViewXml": view, "FolderServerRelativeUrl": path}});
        match self
            .state
            .post_json(&format!("{list_url}/GetItems"), Some(&body), None)
            .await
        {
            Ok(data) => Ok(data
                .get("value")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default()),
            Err(Error::Http { status, .. })
                if status == reqwest::StatusCode::INTERNAL_SERVER_ERROR =>
            {
                Ok(vec![self.get_file(&path).await?])
            }
            Err(error) => Err(error),
        }
    }

    pub async fn default_document_library(&self) -> Result<Value> {
        self.state
            .get_json(
                &format!("{}/DefaultDocumentLibrary", self.state.web_url),
                None,
            )
            .await
    }
}

fn normalize_path(path: &str) -> Result<String> {
    if path.starts_with("http://") || path.starts_with("https://") {
        return odata::browser_path(path)
            .ok_or_else(|| Error::InvalidUrl("browser URL has no id query parameter".to_owned()));
    }
    Ok(path.to_owned())
}
