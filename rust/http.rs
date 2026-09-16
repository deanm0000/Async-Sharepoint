pub use async_sharepoint_core::ClientState;
use pyo3::PyErr;
use pyo3::exceptions::PyRuntimeError;

pub fn runtime_error(error: impl std::fmt::Display) -> PyErr {
    PyRuntimeError::new_err(error.to_string())
}
