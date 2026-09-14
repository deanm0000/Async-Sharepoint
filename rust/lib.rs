use std::sync::Arc;
use std::time::Instant;

use percent_encoding::{AsciiSet, CONTROLS, utf8_percent_encode};
use pyo3::exceptions::{PyAttributeError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyModule, PyTuple};
use pyo3_async_runtimes::TaskLocals;
use serde_json::{Value, json};

mod auth;
mod http;
mod odata;

use crate::http::{ClientState, runtime_error};

const UPLOAD_CHUNK_SIZE: usize = 4 * 1024 * 1024;
const SHAREPOINT_PATH_ENCODE: &AsciiSet =
    &CONTROLS.add(b' ').add(b'"').add(b'<').add(b'>').add(b'`');

fn value_dict(py: Python<'_>, value: &Value) -> PyResult<Py<PyDict>> {
    Ok(pythonize::pythonize(py, value)
        .map_err(runtime_error)?
        .cast_into::<PyDict>()?
        .unbind())
}

fn value_object(py: Python<'_>, value: &Value) -> PyResult<Py<PyAny>> {
    Ok(pythonize::pythonize(py, value)
        .map_err(runtime_error)?
        .unbind())
}

fn item_id(value: &Value) -> String {
    match value.get("Id") {
        Some(Value::String(value)) => value.clone(),
        Some(Value::Number(value)) => value.to_string(),
        Some(value) => value.to_string(),
        None => "None".to_owned(),
    }
}

async fn fetch_all(
    state: &ClientState,
    url: String,
    params: Option<Vec<(String, String)>>,
) -> PyResult<Vec<Value>> {
    let mut results = Vec::new();
    let mut next_url = Some(url);
    let mut next_params = params;
    while let Some(url) = next_url {
        println!("Fetching URL: {:?}", url);
        let data = state.get_json(&url, next_params.as_deref()).await?;
        if let Some(values) = data.get("value").and_then(Value::as_array) {
            results.extend(values.iter().cloned());
        }
        next_url = odata::string(&data, "odata.nextLink");
        next_params = None;
    }
    Ok(results)
}

async fn fetch_deferred(
    state: &ClientState,
    url: String,
    params: Option<Vec<(String, String)>>,
    max_wait: f64,
) -> PyResult<(Vec<Value>, Option<String>)> {
    let start = Instant::now();
    let data = state.get_json(&url, params.as_deref()).await?;
    let mut results = data
        .get("value")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    let mut next_url = odata::string(&data, "odata.nextLink");
    println!("Initial next_url: {:?}", next_url);
    while next_url.is_some() && start.elapsed().as_secs_f64() <= max_wait {
        let data = state
            .get_json(next_url.as_deref().expect("next URL checked above"), None)
            .await?;
        if let Some(values) = data.get("value").and_then(Value::as_array) {
            results.extend(values.iter().cloned());
        }
        next_url = odata::string(&data, "odata.nextLink");
        println!("Next next_url: {:?}", next_url);
    }
    Ok((results, next_url))
}

fn properties_getattr(properties: &Py<PyDict>, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
    properties
        .bind(py)
        .get_item(name)?
        .map(Bound::unbind)
        .ok_or_else(|| PyAttributeError::new_err(name.to_owned()))
}

fn merge_properties(py: Python<'_>, properties: &Py<PyDict>, value: &Value) -> PyResult<()> {
    if let Some(values) = value.as_object() {
        for (key, value) in values {
            properties
                .bind(py)
                .set_item(key, value_object(py, value)?)?;
        }
    }
    Ok(())
}

#[pyclass(module = "async_sharepoint")]
struct SPFolder {
    #[pyo3(get, set)]
    client: Option<Py<SharePointClient>>,
    #[pyo3(get, set)]
    server_relative_url: Option<String>,
    properties: Py<PyDict>,
    #[pyo3(get, set)]
    list_url: Option<String>,
    #[pyo3(get, set)]
    item_id: Option<String>,
}

#[pymethods]
impl SPFolder {
    #[new]
    #[pyo3(signature = (client=None, server_relative_url=None, properties=None, list_url=None, item_id=None))]
    fn new(
        py: Python<'_>,
        client: Option<Py<SharePointClient>>,
        server_relative_url: Option<String>,
        properties: Option<Py<PyDict>>,
        list_url: Option<String>,
        item_id: Option<String>,
    ) -> Self {
        Self {
            client,
            server_relative_url,
            properties: properties.unwrap_or_else(|| PyDict::new(py).unbind()),
            list_url,
            item_id,
        }
    }

    #[getter]
    fn properties(&self, py: Python<'_>) -> Py<PyDict> {
        self.properties.clone_ref(py)
    }

    #[setter]
    fn set_properties(&mut self, value: Py<PyDict>) {
        self.properties = value;
    }

    fn __getattr__(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        properties_getattr(&self.properties, py, name)
    }
}

#[pyclass(module = "async_sharepoint")]
struct SPFile {
    client: Py<SharePointClient>,
    #[pyo3(get, set)]
    server_relative_path: Option<String>,
    properties: Py<PyDict>,
    #[pyo3(get, set)]
    list_url: Option<String>,
    #[pyo3(get, set)]
    item_id: Option<String>,
    resolve_task: Option<tokio::task::JoinHandle<PyResult<Option<String>>>>,
}

impl Drop for SPFile {
    fn drop(&mut self) {
        if let Some(handle) = self.resolve_task.take() {
            if !handle.is_finished() {
                handle.abort();
            }
        }
    }
}

fn start_file_resolution(
    py: Python<'_>,
    file: &Py<SPFile>,
) -> PyResult<tokio::task::JoinHandle<PyResult<Option<String>>>> {
    let file = file.bind(py).borrow();
    let state = Arc::clone(&file.client.bind(py).borrow().state);
    let properties = file.properties.clone_ref(py);
    let (file_url, list_item_url) = if let Some(path) = &file.server_relative_path {
        let file_url = format!(
            "{}/GetFileByServerRelativePath(DecodedUrl={})",
            state.web_url,
            odata::literal(path)
        );
        (
            file_url.clone(),
            Some(format!("{file_url}/ListItemAllFields")),
        )
    } else {
        (
            format!(
                "{}/items({})/File",
                file.list_url
                    .as_deref()
                    .ok_or_else(|| PyRuntimeError::new_err("file has no list_url"))?,
                file.item_id
                    .as_deref()
                    .ok_or_else(|| PyRuntimeError::new_err("file has no item_id"))?
            ),
            None,
        )
    };
    Ok(pyo3_async_runtimes::tokio::get_runtime().spawn(async move {
        let file_data = state.get_json(&file_url, None).await?;
        let path = odata::string(&file_data, "ServerRelativeUrl");
        Python::attach(|py| merge_properties(py, &properties, &file_data))?;
        if let Some(url) = list_item_url {
            let params = [(
                "$select".to_owned(),
                "*,ServerRedirectedEmbedUri".to_owned(),
            )];
            let list_item_data = state.get_json(&url, Some(&params)).await?;
            Python::attach(|py| merge_properties(py, &properties, &list_item_data))?;
        }
        Ok(path)
    }))
}

async fn resolve_file(file: &Py<SPFile>) -> PyResult<(Arc<ClientState>, String)> {
    let (state, task) = Python::attach(|py| {
        let mut file = file.bind(py).borrow_mut();
        let state = Arc::clone(&file.client.bind(py).borrow().state);
        Ok::<_, PyErr>((state, file.resolve_task.take()))
    })?;
    if let Some(task) = task {
        let resolved_path = task.await.map_err(runtime_error)??;
        Python::attach(|py| {
            let mut file = file.bind(py).borrow_mut();
            if resolved_path.is_some() {
                file.server_relative_path = resolved_path;
            }
        });
    }
    let path = Python::attach(|py| file.bind(py).borrow().server_relative_path.clone());
    path.map(|path| (state, path)).ok_or_else(|| {
        PyRuntimeError::new_err("file has no server_relative_path even after resolving")
    })
}

#[pymethods]
impl SPFile {
    #[new]
    #[pyo3(signature = (client, server_relative_path=None, properties=None, list_url=None, item_id=None))]
    fn new(
        py: Python<'_>,
        client: Py<SharePointClient>,
        server_relative_path: Option<String>,
        properties: Option<Py<PyDict>>,
        list_url: Option<String>,
        item_id: Option<String>,
    ) -> PyResult<Py<Self>> {
        let file = Py::new(
            py,
            Self {
                client,
                server_relative_path,
                properties: properties.unwrap_or_else(|| PyDict::new(py).unbind()),
                list_url,
                item_id,
                resolve_task: None,
            },
        )?;
        let handle = start_file_resolution(py, &file)?;
        file.bind(py).borrow_mut().resolve_task = Some(handle);
        Ok(file)
    }

    #[getter]
    fn client(&self, py: Python<'_>) -> Py<SharePointClient> {
        self.client.clone_ref(py)
    }

    #[getter]
    fn resolved(&self) -> bool {
        self.resolve_task
            .as_ref()
            .is_none_or(tokio::task::JoinHandle::is_finished)
    }

    #[getter]
    fn properties(&self, py: Python<'_>) -> Py<PyDict> {
        self.properties.clone_ref(py)
    }

    #[setter]
    fn set_properties(&mut self, value: Py<PyDict>) {
        self.properties = value;
    }

    fn __getattr__(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        properties_getattr(&self.properties, py, name)
    }

    fn resolve<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            resolve_file(&slf).await.map(|_| ())
        })
    }

    fn download<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_file(&slf).await?;
            let url = format!(
                "{}/GetFileByServerRelativePath(DecodedUrl={})/$value",
                state.web_url,
                odata::literal(&path)
            );
            let content = state.get_bytes(&url).await?;
            Python::attach(|py| Ok(PyBytes::new(py, &content).unbind()))
        })
    }

    fn get_url(&self, py: Python<'_>) -> PyResult<String> {
        if self.server_relative_path.is_none() {
            return Err(PyRuntimeError::new_err(
                "file has no server_relative_path even after resolving",
            ));
        }
        let uri: String = self
            .properties
            .bind(py)
            .get_item("ServerRedirectedEmbedUri")?
            .ok_or_else(|| PyAttributeError::new_err("ServerRedirectedEmbedUri"))?
            .extract()?;
        let mut parsed = url::Url::parse(&uri).map_err(runtime_error)?;
        let pairs: Vec<(String, String)> = parsed
            .query_pairs()
            .map(|(key, value)| {
                let value = if key == "action" {
                    "default".to_owned()
                } else {
                    value.into_owned()
                };
                (key.into_owned(), value)
            })
            .collect();
        parsed.query_pairs_mut().clear().extend_pairs(pairs);
        Ok(parsed.into())
    }
}

#[pyclass(module = "async_sharepoint")]
struct SPList {
    client: Py<SharePointClient>,
    #[pyo3(get, set)]
    id: String,
    #[pyo3(get, set)]
    title: String,
    properties: Py<PyDict>,
}

#[pymethods]
impl SPList {
    #[getter]
    fn client(&self, py: Python<'_>) -> Py<SharePointClient> {
        self.client.clone_ref(py)
    }

    #[getter]
    fn properties(&self, py: Python<'_>) -> Py<PyDict> {
        self.properties.clone_ref(py)
    }

    #[setter]
    fn set_properties(&mut self, value: Py<PyDict>) {
        self.properties = value;
    }

    fn __getattr__(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        properties_getattr(&self.properties, py, name)
    }

    #[pyo3(signature = (*, caml=None, max_wait=None))]
    fn get_items<'py>(
        &self,
        py: Python<'py>,
        caml: Option<String>,
        max_wait: Option<f64>,
    ) -> PyResult<Bound<'py, PyAny>> {
        get_items_awaitable(
            py,
            self.client.clone_ref(py),
            None,
            Some(self.id.clone()),
            caml,
            max_wait,
        )
    }
}

fn list_from_value(
    py: Python<'_>,
    client: Py<SharePointClient>,
    value: &Value,
) -> PyResult<Py<SPList>> {
    Py::new(
        py,
        SPList {
            client,
            id: odata::string(value, "Id").unwrap_or_default(),
            title: odata::string(value, "Title").unwrap_or_default(),
            properties: value_dict(py, value)?,
        },
    )
}

fn item_from_value(
    py: Python<'_>,
    client: Py<SharePointClient>,
    list_url: String,
    value: &Value,
) -> PyResult<Py<PyAny>> {
    let properties = value_dict(py, value)?;
    let id = item_id(value);
    if let Some(link_url) = value
        .get("Link")
        .and_then(|link| link.get("Url"))
        .and_then(Value::as_str)
    {
        let parent = client.bind(py).borrow();
        let child = SharePointClient {
            state: parent
                .state
                .for_site(link_url.trim_end_matches('/').to_owned()),
            properties,
            item_id: Some(id),
            list_url: Some(list_url),
        };
        return Ok(Py::new(py, child)?.into_any());
    }
    let fsobj_type = value.get("FileSystemObjectType").and_then(Value::as_i64);
    if fsobj_type == Some(1) {
        return Ok(Py::new(
            py,
            SPFolder {
                client: Some(client),
                server_relative_url: odata::string(value, "ServerRelativeUrl")
                    .or_else(|| odata::string(value, "FileRef")),
                properties,
                list_url: Some(list_url),
                item_id: Some(id),
            },
        )?
        .into_any());
    }
    if fsobj_type != Some(0) {
        return Err(PyRuntimeError::new_err(format!(
            "unexpected FileSystemObjectType: {fsobj_type:?}"
        )));
    }
    let file = Py::new(
        py,
        SPFile {
            client,
            server_relative_path: None,
            properties,
            list_url: Some(list_url),
            item_id: Some(id),
            resolve_task: None,
        },
    )?;
    let handle = start_file_resolution(py, &file)?;
    file.bind(py).borrow_mut().resolve_task = Some(handle);
    Ok(file.into_any())
}

fn items_list(
    py: Python<'_>,
    client: &Py<SharePointClient>,
    list_url: &str,
    values: &[Value],
) -> PyResult<Py<PyList>> {
    let items = values
        .iter()
        .map(|value| item_from_value(py, client.clone_ref(py), list_url.to_owned(), value))
        .collect::<PyResult<Vec<_>>>()?;
    Ok(PyList::new(py, items)?.unbind())
}

fn lists_list(
    py: Python<'_>,
    client: &Py<SharePointClient>,
    values: &[Value],
) -> PyResult<Py<PyList>> {
    let lists = values
        .iter()
        .map(|value| list_from_value(py, client.clone_ref(py), value))
        .collect::<PyResult<Vec<_>>>()?;
    Ok(PyList::new(py, lists)?.unbind())
}

fn deferred_items(
    py: Python<'_>,
    locals: TaskLocals,
    state: Arc<ClientState>,
    client: Py<SharePointClient>,
    list_url: String,
    next_url: Option<String>,
) -> PyResult<Py<PyAny>> {
    Ok(
        pyo3_async_runtimes::tokio::future_into_py_with_locals(py, locals, async move {
            let values = match next_url {
                Some(url) => fetch_all(&state, url, None).await?,
                None => Vec::new(),
            };
            Python::attach(|py| items_list(py, &client, &list_url, &values))
        })?
        .unbind(),
    )
}

fn get_items_awaitable<'py>(
    py: Python<'py>,
    client: Py<SharePointClient>,
    title: Option<String>,
    id: Option<String>,
    caml: Option<String>,
    max_wait: Option<f64>,
) -> PyResult<Bound<'py, PyAny>> {
    let state = Arc::clone(&client.bind(py).borrow().state);
    let locals = pyo3_async_runtimes::tokio::get_current_locals(py)?;
    pyo3_async_runtimes::tokio::future_into_py_with_locals(py, locals.clone(), async move {
        if title.is_some() && id.is_some() {
            return Err(PyValueError::new_err("specify only one of title or id"));
        }
        let list_url = if let Some(title) = title {
            format!(
                "{}/lists/GetByTitle({})",
                state.web_url,
                odata::literal(&title)
            )
        } else if let Some(id) = id {
            format!("{}/lists/GetById({})", state.web_url, odata::literal(&id))
        } else {
            return Err(PyValueError::new_err("must specify title or id"));
        };

        let (values, next_url) = if let Some(caml) = caml {
            let body = json!({"query": {"__metadata": {"type": "SP.CamlQuery"}, "ViewXml": caml}});
            let data = state
                .post_json(&format!("{list_url}/GetItems"), Some(&body), None)
                .await?;
            (
                data.get("value")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default(),
                None,
            )
        } else if let Some(max_wait) = max_wait {
            fetch_deferred(&state, format!("{list_url}/items"), None, max_wait).await?
        } else {
            (
                fetch_all(&state, format!("{list_url}/items"), None).await?,
                None,
            )
        };

        Python::attach(|py| {
            let items = items_list(py, &client, &list_url, &values)?;
            if max_wait.is_some() {
                let remaining =
                    deferred_items(py, locals, Arc::clone(&state), client, list_url, next_url)?;
                Ok(PyTuple::new(py, [items.into_any(), remaining])?
                    .into_any()
                    .unbind())
            } else {
                Ok(items.into_any())
            }
        })
    })
}

#[pyclass(module = "async_sharepoint")]
struct SharePointClient {
    state: Arc<ClientState>,
    properties: Py<PyDict>,
    #[pyo3(get, set)]
    item_id: Option<String>,
    #[pyo3(get, set)]
    list_url: Option<String>,
}

#[pymethods]
impl SharePointClient {
    #[new]
    #[pyo3(signature = (site_url, get_token, *, properties=None, item_id=None, list_url=None))]
    fn new(
        py: Python<'_>,
        site_url: &str,
        get_token: Py<PyAny>,
        properties: Option<Py<PyDict>>,
        item_id: Option<String>,
        list_url: Option<String>,
    ) -> PyResult<Self> {
        Ok(Self {
            state: ClientState::new(site_url.trim_end_matches('/').to_owned(), get_token)?,
            properties: properties.unwrap_or_else(|| PyDict::new(py).unbind()),
            item_id,
            list_url,
        })
    }

    #[getter]
    fn site_url(&self) -> &str {
        &self.state.site_url
    }

    #[getter]
    fn properties(&self, py: Python<'_>) -> Py<PyDict> {
        self.properties.clone_ref(py)
    }

    #[setter]
    fn set_properties(&mut self, value: Py<PyDict>) {
        self.properties = value;
    }

    fn __getattr__(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        properties_getattr(&self.properties, py, name)
    }

    fn aclose<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async { Ok(()) })
    }

    fn __aenter__<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move { Ok(slf) })
    }

    fn __aexit__<'py>(
        &self,
        py: Python<'py>,
        _exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async { Ok(()) })
    }

    #[pyo3(signature = (title=None, *, id=None, max_wait=None))]
    fn get<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        title: Option<String>,
        id: Option<String>,
        max_wait: Option<f64>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&slf.bind(py).borrow().state);
        let locals = pyo3_async_runtimes::tokio::get_current_locals(py)?;
        pyo3_async_runtimes::tokio::future_into_py_with_locals(py, locals.clone(), async move {
            if title.is_some() && id.is_some() {
                return Err(PyValueError::new_err("specify only one of title or id"));
            }
            if (title.is_some() || id.is_some()) && max_wait.is_some() {
                return Err(PyValueError::new_err(
                    "max_wait is only supported when fetching all lists",
                ));
            }
            if let Some(title) = title {
                let data = state
                    .get_json(
                        &format!(
                            "{}/lists/GetByTitle({})",
                            state.web_url,
                            odata::literal(&title)
                        ),
                        None,
                    )
                    .await?;
                return Python::attach(|py| Ok(list_from_value(py, slf, &data)?.into_any()));
            }
            if let Some(id) = id {
                let data = state
                    .get_json(
                        &format!("{}/lists/GetById({})", state.web_url, odata::literal(&id)),
                        None,
                    )
                    .await?;
                return Python::attach(|py| Ok(list_from_value(py, slf, &data)?.into_any()));
            }

            let (values, next_url) = if let Some(max_wait) = max_wait {
                fetch_deferred(&state, format!("{}/lists", state.web_url), None, max_wait).await?
            } else {
                (
                    fetch_all(&state, format!("{}/lists", state.web_url), None).await?,
                    None,
                )
            };
            Python::attach(|py| {
                let lists = lists_list(py, &slf, &values)?;
                if max_wait.is_some() {
                    let state = Arc::clone(&state);
                    let client = slf;
                    let remaining = pyo3_async_runtimes::tokio::future_into_py_with_locals(
                        py,
                        locals,
                        async move {
                            let values = match next_url {
                                Some(url) => fetch_all(&state, url, None).await?,
                                None => Vec::new(),
                            };
                            Python::attach(|py| lists_list(py, &client, &values))
                        },
                    )?
                    .unbind();
                    Ok(PyTuple::new(py, [lists.into_any(), remaining])?
                        .into_any()
                        .unbind())
                } else {
                    Ok(lists.into_any())
                }
            })
        })
    }

    fn get_default_document_library<'py>(
        slf: Py<Self>,
        py: Python<'py>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&slf.bind(py).borrow().state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let data = state
                .get_json(&format!("{}/DefaultDocumentLibrary", state.web_url), None)
                .await?;
            Python::attach(|py| list_from_value(py, slf, &data))
        })
    }

    #[pyo3(signature = (title=None, *, id=None, caml=None, max_wait=None))]
    fn get_items<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        title: Option<String>,
        id: Option<String>,
        caml: Option<String>,
        max_wait: Option<f64>,
    ) -> PyResult<Bound<'py, PyAny>> {
        get_items_awaitable(py, slf, title, id, caml, max_wait)
    }

    fn get_file<'py>(slf: Py<Self>, py: Python<'py>, path: String) -> PyResult<Bound<'py, PyAny>> {
        let file = Py::new(
            py,
            SPFile {
                client: slf,
                server_relative_path: Some(path),
                properties: PyDict::new(py).unbind(),
                list_url: None,
                item_id: None,
                resolve_task: None,
            },
        )?;
        let handle = start_file_resolution(py, &file)?;
        file.bind(py).borrow_mut().resolve_task = Some(handle);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            resolve_file(&file).await?;
            Ok(file)
        })
    }

    fn download<'py>(&self, py: Python<'py>, path: String) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let url = format!(
                "{}/GetFileByServerRelativePath(DecodedUrl={})/$value",
                state.web_url,
                odata::literal(&path)
            );
            let content = state.get_bytes(&url).await?;
            Python::attach(|py| Ok(PyBytes::new(py, &content).unbind()))
        })
    }

    #[pyo3(signature = (folder_path, filename, content, *, overwrite=true))]
    fn upload<'py>(
        &self,
        py: Python<'py>,
        folder_path: String,
        filename: String,
        content: &Bound<'_, PyBytes>,
        overwrite: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let content = content.as_bytes().to_vec();
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let folder_url = format!(
                "{}/GetFolderByServerRelativeUrl({})",
                state.web_url,
                odata::literal(&folder_path)
            );
            let add_url = format!(
                "{folder_url}/Files/add(url={},overwrite={})",
                odata::literal(&filename),
                odata::bool_literal(overwrite)
            );
            let data = if content.len() <= UPLOAD_CHUNK_SIZE {
                state.post_json(&add_url, None, Some(content)).await?
            } else {
                state.post_json(&add_url, None, Some(Vec::new())).await?;
                let file_url = format!("{folder_url}/Files({})", odata::literal(&filename));
                let upload_id = uuid::Uuid::new_v4().to_string();
                let mut position = UPLOAD_CHUNK_SIZE;
                state
                    .post_json(
                        &format!(
                            "{file_url}/startUpload(uploadID={})",
                            odata::literal(&upload_id)
                        ),
                        None,
                        Some(content[..UPLOAD_CHUNK_SIZE].to_vec()),
                    )
                    .await?;
                loop {
                    let end = (position + UPLOAD_CHUNK_SIZE).min(content.len());
                    let chunk = content[position..end].to_vec();
                    if end < content.len() {
                        state
                            .post_json(
                                &format!(
                                    "{file_url}/continueUpload(uploadID={},fileOffset={position})",
                                    odata::literal(&upload_id)
                                ),
                                None,
                                Some(chunk),
                            )
                            .await?;
                        position = end;
                    } else {
                        break state
                            .post_json(
                                &format!(
                                    "{file_url}/finishUpload(uploadID={},fileOffset={position})",
                                    odata::literal(&upload_id)
                                ),
                                None,
                                Some(chunk),
                            )
                            .await?;
                    }
                }
            };
            let file = Python::attach(|py| {
                let client = Py::new(
                    py,
                    SharePointClient {
                        state,
                        properties: PyDict::new(py).unbind(),
                        item_id: None,
                        list_url: None,
                    },
                )?;
                let file = Py::new(
                    py,
                    SPFile {
                        client,
                        server_relative_path: odata::string(&data, "ServerRelativeUrl"),
                        properties: value_dict(py, &data)?,
                        list_url: None,
                        item_id: None,
                        resolve_task: None,
                    },
                )?;
                let handle = start_file_resolution(py, &file)?;
                file.bind(py).borrow_mut().resolve_task = Some(handle);
                Ok::<_, PyErr>(file)
            })?;
            resolve_file(&file).await?;
            Ok(file)
        })
    }

    #[pyo3(signature = (path, *, overwrite=false))]
    fn add_folder<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        path: String,
        overwrite: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&slf.bind(py).borrow().state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let url = format!(
                "{}/Folders/AddUsingPath(DecodedUrl={},Overwrite={})",
                state.web_url,
                odata::literal(&path),
                odata::bool_literal(overwrite)
            );
            let data = state.post_json(&url, None, None).await?;
            Python::attach(|py| {
                Py::new(
                    py,
                    SPFolder {
                        client: Some(slf),
                        server_relative_url: odata::string(&data, "ServerRelativeUrl")
                            .or(Some(path)),
                        properties: value_dict(py, &data)?,
                        list_url: None,
                        item_id: None,
                    },
                )
            })
        })
    }

    fn get_current_user<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let data = state
                .get_json(&format!("{}/CurrentUser", state.web_url), None)
                .await?;
            Python::attach(|py| value_dict(py, &data))
        })
    }

    #[pyo3(signature = (path, *, login_name=None))]
    fn get_effective_permissions<'py>(
        &self,
        py: Python<'py>,
        path: String,
        login_name: Option<String>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let login_name = match login_name {
                Some(login_name) => login_name,
                None => state
                    .get_json(&format!("{}/CurrentUser", state.web_url), None)
                    .await?
                    .get("LoginName")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        PyRuntimeError::new_err("CurrentUser response has no LoginName")
                    })?
                    .to_owned(),
            };
            let url = format!(
                "{}/GetFolderByServerRelativeUrl({})/ListItemAllFields/GetUserEffectivePermissions({})",
                state.web_url,
                odata::literal(&path),
                odata::literal(&login_name)
            );
            let data = state.post_json(&url, None, None).await?;
            Python::attach(|py| value_dict(py, &data))
        })
    }

    #[pyo3(signature = (text, *, title=None, id=None, row_limit=6))]
    fn search<'py>(
        &self,
        py: Python<'py>,
        text: String,
        title: Option<String>,
        id: Option<String>,
        row_limit: i64,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            if title.is_none() && id.is_none() {
                return Err(PyValueError::new_err("must specify title or id"));
            }
            if title.is_some() && id.is_some() {
                return Err(PyValueError::new_err("specify only one of title or id"));
            }
            let list_url = if let Some(title) = title {
                format!(
                    "{}/lists/GetByTitle({})",
                    state.web_url,
                    odata::literal(&title)
                )
            } else {
                format!(
                    "{}/lists/GetById({})",
                    state.web_url,
                    odata::literal(id.as_deref().expect("validated above"))
                )
            };
            let list = state.get_json(&list_url, None).await?;
            let list_id = odata::string(&list, "Id").unwrap_or_default();
            let site = state
                .get_json(&format!("{}/_api/site", state.site_url), None)
                .await?;
            let web = state.get_json(&state.web_url, None).await?;
            let root = state
                .get_json(
                    &format!(
                        "{}/lists/GetById({})/RootFolder",
                        state.web_url,
                        odata::literal(&list_id)
                    ),
                    None,
                )
                .await?;
            let site_id = odata::string(&site, "Id").unwrap_or_default();
            let web_id = odata::string(&web, "Id").unwrap_or_default();
            let root_path = odata::string(&root, "ServerRelativeUrl").unwrap_or_default();
            let parsed = url::Url::parse(&state.site_url).map_err(runtime_error)?;
            let origin = parsed.origin().ascii_serialization();
            let absolute_path = format!(
                "{origin}{}",
                utf8_percent_encode(&root_path, SHAREPOINT_PATH_ENCODE)
            );
            let query_template = format!(
                "{{searchTerms}} (siteId:{{{site_id}}} OR siteId:{site_id}) (webId:{{{web_id}}} OR webId:{web_id}) (NormListID:{list_id}) (path:\"{absolute_path}\" OR ParentLink:\"{origin}{root_path}*\") ContentTypeId:0x0*"
            );
            let select_properties = [
                "editorowsuser",
                "authorowsuser",
                "Filename",
                "SPSiteURL",
                "Title",
                "ParentLink",
                "ListItemID",
                "ListID",
                "contentclass",
                "IsDocument",
                "IsContainer",
                "FileExtension",
                "SecondaryFileExtension",
                "OriginalPath",
                "DefaultEncodingURL",
                "ServerRedirectedURL",
                "ServerRedirectedPreviewURL",
                "LastModifiedTime",
                "SharedWithUsersOWSUser",
                "HitHighlightedSummary",
                "ModifierDates",
                "LastModifiedTimeForRetention",
            ];
            let params = vec![
                ("querytext".to_owned(), format!("'({text}*)'")),
                ("querytemplate".to_owned(), format!("'{query_template}'")),
                (
                    "selectproperties".to_owned(),
                    format!("'{}'", select_properties.join(",")),
                ),
                ("SummaryLength".to_owned(), "100".to_owned()),
                ("RowLimit".to_owned(), row_limit.to_string()),
                ("culture".to_owned(), "1033".to_owned()),
                ("BypassResultTypes".to_owned(), "true".to_owned()),
                ("EnableQueryRules".to_owned(), "false".to_owned()),
                ("ProcessBestBets".to_owned(), "false".to_owned()),
                ("ProcessPersonalFavorites".to_owned(), "false".to_owned()),
                ("clienttype".to_owned(), "'sug_SPListInline'".to_owned()),
                (
                    "properties".to_owned(),
                    "'EnableDynamicGroups:true'".to_owned(),
                ),
                ("TrimDuplicates".to_owned(), "false".to_owned()),
            ];
            let payload = state
                .get_json(
                    &format!("{}/_api/search/query", state.site_url),
                    Some(&params),
                )
                .await?;
            let rows = odata::search_rows(&payload);
            Python::attach(|py| {
                let values = rows
                    .iter()
                    .map(|value| value_dict(py, value))
                    .collect::<PyResult<Vec<_>>>()?;
                Ok(PyList::new(py, values)?.unbind())
            })
        })
    }
}

#[pymodule]
fn async_sharepoint(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<SharePointClient>()?;
    module.add_class::<SPFile>()?;
    module.add_class::<SPFolder>()?;
    module.add_class::<SPList>()?;
    module.add("__all__", vec!["SPFile", "SPFolder", "SharePointClient"])?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chunk_size_is_four_mebibytes() {
        assert_eq!(UPLOAD_CHUNK_SIZE, 4_194_304);
    }
}
