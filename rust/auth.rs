use async_sharepoint_core::CertificateCredential as CoreCertificateCredential;
use pyo3::prelude::*;

/// Python wrapper for the core Entra ID certificate credentials.
#[pyclass(module = "async_sharepoint", frozen)]
pub struct CertificateCredential {
    pub(crate) inner: CoreCertificateCredential,
}

#[pymethods]
impl CertificateCredential {
    #[new]
    #[pyo3(signature = (*, tenant_id, client_id, private_key_path, thumbprint))]
    fn new(
        tenant_id: String,
        client_id: String,
        private_key_path: String,
        thumbprint: String,
    ) -> PyResult<Self> {
        Ok(Self {
            inner: CoreCertificateCredential::load(
                tenant_id,
                client_id,
                &private_key_path,
                &thumbprint,
            )?,
        })
    }

    #[getter]
    fn tenant_id(&self) -> &str {
        self.inner.tenant_id()
    }

    #[getter]
    fn client_id(&self) -> &str {
        self.inner.client_id()
    }

    fn __repr__(&self) -> String {
        format!(
            "CertificateCredential(tenant_id={:?}, client_id={:?})",
            self.inner.tenant_id(),
            self.inner.client_id(),
        )
    }
}
