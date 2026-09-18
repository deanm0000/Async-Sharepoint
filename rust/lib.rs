use async_sharepoint_core::{Error, odata};
use percent_encoding::{AsciiSet, CONTROLS, utf8_percent_encode};
use pyo3::exceptions::{PyAttributeError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyModule, PyTuple};
use pyo3_async_runtimes::TaskLocals;
use serde_json::{Value, json};
use std::collections::{HashMap, HashSet};
use std::env;
use std::future::Future;
use std::sync::Arc;
use std::time::Instant;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::task::{AbortHandle, JoinSet};

mod auth;
mod http;

use crate::auth::CertificateCredential;
use crate::http::{ClientState, runtime_error};

const UPLOAD_CHUNK_SIZE: usize = 4 * 1024 * 1024;
const DOWNLOAD_CHUNK_SIZE: u64 = 4 * 1024 * 1024;
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

fn browser_path(path: &str) -> PyResult<String> {
    if !path.starts_with("http://") && !path.starts_with("https://") {
        return Ok(path.to_owned());
    }
    let url = url::Url::parse(path).map_err(runtime_error)?;
    url.query_pairs()
        .find(|(key, _)| key == "id")
        .map(|(_, value)| value.into_owned())
        .ok_or_else(|| PyValueError::new_err("browser URL has no id query parameter"))
}

/// Either a server-relative path or a file's unique id (e.g. from a `sourcedoc` browser link).
enum FileLocator {
    Path(String),
    Id(String),
}

/// Resolves a Python `str` or `uuid.UUID` into a `FileLocator`, without round-tripping a
/// `uuid.UUID` through string parsing back into a `uuid::Uuid`.
fn file_locator_from_py(value: &Bound<'_, PyAny>) -> PyResult<FileLocator> {
    let uuid_type = value.py().import("uuid")?.getattr("UUID")?;
    if value.is_instance(&uuid_type)? {
        return Ok(FileLocator::Id(value.str()?.extract()?));
    }
    browser_file_locator(&value.extract::<String>()?)
}

/// Resolve a server-relative path, a bare/braced UniqueId (as found in a
/// "Doc.aspx" `sourcedoc` query parameter), an AllItems browser URL (`?id=`), or a
/// "Doc.aspx" / OneDrive-style browser URL (`?sourcedoc={guid}`).
fn browser_file_locator(path: &str) -> PyResult<FileLocator> {
    let trimmed = path.trim_matches(|c| c == '{' || c == '}');
    if let Ok(id) = trimmed.parse::<uuid::Uuid>() {
        return Ok(FileLocator::Id(id.to_string()));
    }
    if !path.starts_with("http://") && !path.starts_with("https://") {
        return Ok(FileLocator::Path(path.to_owned()));
    }
    let url = url::Url::parse(path).map_err(runtime_error)?;
    if let Some((_, value)) = url.query_pairs().find(|(key, _)| key == "sourcedoc") {
        let id = value.trim_matches(|c| c == '{' || c == '}').to_owned();
        return Ok(FileLocator::Id(id));
    }
    url.query_pairs()
        .find(|(key, _)| key == "id")
        .map(|(_, value)| FileLocator::Path(value.into_owned()))
        .ok_or_else(|| PyValueError::new_err("browser URL has no id or sourcedoc query parameter"))
}

fn browser_url(site_url: &str, path: &str, is_file: bool) -> PyResult<String> {
    let mut url = url::Url::parse(site_url).map_err(runtime_error)?;
    let site_path = url.path().trim_end_matches('/').to_owned();
    let relative = path
        .strip_prefix(&site_path)
        .unwrap_or(path)
        .trim_start_matches('/');
    let library = relative
        .split('/')
        .next()
        .filter(|segment| !segment.is_empty())
        .ok_or_else(|| PyRuntimeError::new_err("path is not inside this SharePoint site"))?;
    url.set_path(&format!("{site_path}/{library}/Forms/AllItems.aspx"));
    if path != format!("{site_path}/{library}") && path != format!("{site_path}/{library}/") {
        let mut pairs = url.query_pairs_mut();
        pairs.append_pair("id", path);
        if is_file {
            let parent = path.rsplit_once('/').map_or(path, |(parent, _)| parent);
            pairs.append_pair("parent", parent);
        }
    }
    Ok(url.into())
}

// SharePoint list item payloads key their id as "ID", not the "Id" used elsewhere.
fn item_dedup_id(value: &Value) -> Option<String> {
    let id = value.get("ID").or_else(|| value.get("Id"))?;
    Some(match id {
        Value::String(id) => id.clone(),
        other => other.to_string(),
    })
}

/// Append `data`'s "value" items to `results`, skipping ones already in `seen`.
/// Returns (items in this page, items newly added).
fn merge_page(
    results: &mut Vec<Value>,
    seen: &mut HashSet<String>,
    data: &Value,
) -> (usize, usize) {
    let Some(values) = data.get("value").and_then(Value::as_array) else {
        return (0, 0);
    };
    let mut added = 0;
    for item in values {
        let is_new = match item_dedup_id(item) {
            Some(id) => seen.insert(id),
            None => true,
        };
        if is_new {
            results.push(item.clone());
            added += 1;
        }
    }
    (values.len(), added)
}

// A guess's actual next id beyond this multiple of its expected next id triggers a regroup.
const REDUNDANCY_THRESHOLD: f64 = 1.5;

struct GuessOutcome {
    guess_id: u64,
    expected_next: u64,
    data: Value,
    next_id: Option<u64>,
    next_url: Option<String>,
}

async fn fetch_guess(
    state: Arc<ClientState>,
    url: String,
    guess_id: u64,
    expected_next: u64,
) -> PyResult<GuessOutcome> {
    let data = state.get_json(&url, None).await?;
    let next_url = odata::string(&data, "odata.nextLink");
    let next_id = next_url.as_deref().and_then(odata::skiptoken_p_id);
    Ok(GuessOutcome {
        guess_id,
        expected_next,
        data,
        next_id,
        next_url,
    })
}

fn batch_min_max(data: &Value) -> (Option<u64>, Option<u64>) {
    let Some(values) = data.get("value").and_then(Value::as_array) else {
        return (None, None);
    };
    values
        .iter()
        .filter_map(|item| item_dedup_id(item)?.parse::<u64>().ok())
        .fold(
            (None, None),
            |(min, max): (Option<u64>, Option<u64>), id| {
                (
                    Some(min.map_or(id, |m| m.min(id))),
                    Some(max.map_or(id, |m| m.max(id))),
                )
            },
        )
}

async fn fetch_all(
    state: Arc<ClientState>,
    url: String,
    params: Option<Vec<(String, String)>>,
) -> PyResult<Vec<Value>> {
    let data = state.get_json(&url, params.as_deref()).await?;
    let mut results = Vec::new();
    let mut seen = HashSet::new();
    let (mut n, _) = merge_page(&mut results, &mut seen, &data);
    let mut next_url = odata::string(&data, "odata.nextLink");
    let max_guesses = env::var("MAX_GUESSES")
        .ok()
        .and_then(|s| s.parse::<usize>().ok())
        .unwrap_or(20);
    // Fall back to plain sequential fetching until we can parse a skiptoken to guess ahead from.
    while let Some(url) = next_url.take() {
        let Some(base_id) = (n > 0).then(|| odata::skiptoken_p_id(&url)).flatten() else {
            let data = state.get_json(&url, None).await?;
            let (page_len, _) = merge_page(&mut results, &mut seen, &data);
            n = page_len;
            next_url = odata::string(&data, "odata.nextLink");
            continue;
        };

        // Continuously fetch the real chain plus speculative guesses ahead of it, capped at
        // GUESS_COUNT concurrent requests, until this pagination chain is exhausted.
        let mut frontier_id = base_id;
        let mut frontier_n = n as u64;
        let mut template = url;
        let mut next_offset: u64 = 0;
        let mut dispatched_max = frontier_id;
        let mut terminal_at: Option<u64> = None;
        // let mut paused_awaiting: Option<u64> = None;
        // let mut paused_awaitings: HashSet<u64> = HashSet::new();
        let mut tasks: JoinSet<PyResult<GuessOutcome>> = JoinSet::new();
        let mut abort_handles: HashMap<u64, AbortHandle> = HashMap::new();

        loop {
            while tasks.len() <= max_guesses
                && terminal_at.is_none_or(|t| frontier_id + frontier_n * next_offset <= t)
            {
                let guess_id = frontier_id + frontier_n * next_offset;
                let guess_url = odata::with_p_id(&template, guess_id);
                let expected_next = guess_id + frontier_n;
                let handle = tasks.spawn(fetch_guess(
                    Arc::clone(&state),
                    guess_url,
                    guess_id,
                    expected_next,
                ));
                abort_handles.insert(guess_id, handle);
                dispatched_max = dispatched_max.max(guess_id);
                next_offset += 1;
            }

            let Some(joined) = tasks.join_next().await else {
                break;
            };
            let outcome = match joined {
                Ok(outcome) => outcome?,
                Err(join_err) if join_err.is_cancelled() => continue,
                Err(join_err) => return Err(runtime_error(join_err)),
            };
            abort_handles.remove(&outcome.guess_id);

            let (page_len, _unique_added) = merge_page(&mut results, &mut seen, &outcome.data);
            let (_min_id, _max_id) = batch_min_max(&outcome.data);
            // println!(
            //     "Guess id={} -> next_url id={:?}, batch min={:?} max={:?}, records={}, unique_added={}",
            //     outcome.guess_id, outcome.next_id, min_id, max_id, page_len, unique_added
            // );

            match outcome.next_id {
                Some(next_id) => {
                    // A small overshoot past the next guess id is normal (ids aren't evenly
                    // spaced) and does NOT mean that guess is redundant; only an anomalously
                    // large jump means this page's range already swallowed other guesses.
                    let expected_next_threshold =
                        outcome.expected_next as f64 + frontier_n as f64 * REDUNDANCY_THRESHOLD;
                    if (next_id as f64) > expected_next_threshold {
                        // Never cancel the guess we're waiting to regroup on.
                        let redundant: Vec<u64> = abort_handles
                            .keys()
                            .copied()
                            .filter(|id| *id < next_id - frontier_n && *id > outcome.guess_id)
                            .collect();
                        // println!(
                        //     "Found next_id {} is greater than {}. Redundant guess ids to abort: {:?}",
                        //     next_id, expected_next_threshold, redundant
                        // );
                        for id in redundant {
                            if let Some(handle) = abort_handles.remove(&id) {
                                handle.abort();
                            }
                        }

                        frontier_id = next_id;
                        frontier_n = page_len.max(1) as u64;
                        template = outcome.next_url.clone().unwrap_or(template);
                        next_offset = 0;
                    }

                    // if paused_awaitings.contains(&outcome.guess_id) {
                    //     paused_awaitings.remove(&outcome.guess_id);
                    //     frontier_id = next_id;
                    //     frontier_n = page_len.max(1) as u64;
                    //     template = outcome.next_url.clone().unwrap_or(template);
                    //     next_offset = 0;
                    // }
                }
                None => {
                    // Nothing exists past this id; drop any guesses further out than it, but
                    // never the one we're waiting to regroup on.
                    let boundary =
                        terminal_at.map_or(outcome.guess_id, |t| t.min(outcome.guess_id));
                    terminal_at = Some(boundary);

                    let redundant: Vec<u64> = abort_handles
                        .keys()
                        .copied()
                        .filter(|id| *id > boundary)
                        .collect();
                    // println!(
                    //     "Setting terminal boundary at {}. Cancelling redundant guesses. {:?}",
                    //     boundary, redundant
                    // );
                    for id in redundant {
                        if let Some(handle) = abort_handles.remove(&id) {
                            handle.abort();
                        }
                    }
                }
            }
        }
    }

    results.sort_by(|a, b| match (item_dedup_id(a), item_dedup_id(b)) {
        (Some(ka), Some(kb)) => match (ka.parse::<u64>(), kb.parse::<u64>()) {
            (Ok(na), Ok(nb)) => na.cmp(&nb),
            _ => ka.cmp(&kb),
        },
        (Some(_), None) => std::cmp::Ordering::Less,
        (None, Some(_)) => std::cmp::Ordering::Greater,
        (None, None) => std::cmp::Ordering::Equal,
    });
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
    while next_url.is_some() && start.elapsed().as_secs_f64() <= max_wait {
        let data = state
            .get_json(next_url.as_deref().expect("next URL checked above"), None)
            .await?;
        if let Some(values) = data.get("value").and_then(Value::as_array) {
            results.extend(values.iter().cloned());
        }
        next_url = odata::string(&data, "odata.nextLink");
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

fn properties_label(properties: &Py<PyDict>, py: Python<'_>) -> PyResult<String> {
    let dict = properties.bind(py);
    let title = dict.get_item("Title")?;
    let truthy_title = match &title {
        Some(value) => value.is_truthy()?,
        None => false,
    };
    let value = if truthy_title {
        title
    } else {
        dict.get_item("Name")?
    };
    match value {
        Some(value) => Ok(value.str()?.to_string()),
        None => Ok("None".to_owned()),
    }
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
    resolve_task: Option<tokio::task::JoinHandle<PyResult<Option<String>>>>,
}

impl Drop for SPFolder {
    fn drop(&mut self) {
        if let Some(handle) = self.resolve_task.take()
            && !handle.is_finished()
        {
            handle.abort();
        }
    }
}

fn start_folder_resolution(
    py: Python<'_>,
    folder: &Py<SPFolder>,
) -> PyResult<tokio::task::JoinHandle<PyResult<Option<String>>>> {
    let folder = folder.bind(py).borrow();
    let client = folder
        .client
        .as_ref()
        .ok_or_else(|| PyRuntimeError::new_err("folder has no client"))?;
    let state = Arc::clone(&client.bind(py).borrow().state);
    let properties = folder.properties.clone_ref(py);
    let folder_url = if let Some(path) = &folder.server_relative_url {
        format!(
            "{}/GetFolderByServerRelativeUrl({})",
            state.web_url,
            odata::literal(path)
        )
    } else {
        format!(
            "{}/items({})/Folder",
            folder
                .list_url
                .as_deref()
                .ok_or_else(|| PyRuntimeError::new_err("folder has no list_url"))?,
            folder
                .item_id
                .as_deref()
                .ok_or_else(|| PyRuntimeError::new_err("folder has no item_id"))?
        )
    };
    Ok(pyo3_async_runtimes::tokio::get_runtime().spawn(async move {
        let folder_data = state.get_json(&folder_url, None).await?;
        let path = odata::string(&folder_data, "ServerRelativeUrl");
        Python::attach(|py| merge_properties(py, &properties, &folder_data))?;
        Ok(path)
    }))
}

async fn resolve_folder(folder: &Py<SPFolder>) -> PyResult<(Arc<ClientState>, String)> {
    let (state, task) = Python::attach(|py| {
        let mut folder = folder.bind(py).borrow_mut();
        let client = folder
            .client
            .as_ref()
            .ok_or_else(|| PyRuntimeError::new_err("folder has no client"))?;
        let state = Arc::clone(&client.bind(py).borrow().state);
        Ok::<_, PyErr>((state, folder.resolve_task.take()))
    })?;
    if let Some(task) = task {
        let resolved_path = task.await.map_err(runtime_error)??;
        Python::attach(|py| {
            let mut folder = folder.bind(py).borrow_mut();
            if resolved_path.is_some() {
                folder.server_relative_url = resolved_path;
            }
        });
    }
    let path = Python::attach(|py| folder.bind(py).borrow().server_relative_url.clone());
    path.map(|path| (state, path)).ok_or_else(|| {
        PyRuntimeError::new_err("folder has no server_relative_url even after resolving")
    })
}

fn ls_awaitable<'py>(
    py: Python<'py>,
    client: Py<SharePointClient>,
    path: Option<String>,
) -> PyResult<Bound<'py, PyAny>> {
    let state = Arc::clone(&client.bind(py).borrow().state);
    pyo3_async_runtimes::tokio::future_into_py(py, ls_items(state, client, path))
}

async fn ls_items(
    state: Arc<ClientState>,
    client: Py<SharePointClient>,
    path: Option<String>,
) -> PyResult<Py<PyAny>> {
    let documents = state
        .get_json(&format!("{}/DefaultDocumentLibrary", state.web_url), None)
        .await
        .map_err(runtime_error)?;
    let id = odata::string(&documents, "Id").unwrap_or_default();
    let root_path = {
        let root = state
            .get_json(
                &format!(
                    "{}/lists/GetById({})/RootFolder",
                    state.web_url,
                    odata::literal(&id)
                ),
                None,
            )
            .await
            .map_err(runtime_error)?;
        odata::string(&root, "ServerRelativeUrl").ok_or_else(|| {
            PyRuntimeError::new_err("DefaultDocumentLibrary root has no ServerRelativeUrl")
        })?
    };
    let path = match path {
        Some(path) => browser_path(&path)?,
        None => root_path.clone(),
    };
    let list_url = format!("{}/lists/GetById({})", state.web_url, odata::literal(&id));
    // SharePoint's CAML FileDirRef equality filter 500s when compared against the
    // document library's own root folder, so scope the root via FolderServerRelativeUrl
    // alone instead of also filtering by FileDirRef.
    let caml = if path == root_path {
        "<View Scope=\"Default\"><ViewFields><FieldRef Name=\"FileRef\" /><FieldRef Name=\"FileLeafRef\" /><FieldRef Name=\"FSObjType\" /><FieldRef Name=\"FileDirRef\" /></ViewFields><RowLimit>500</RowLimit></View>"
                .to_owned()
    } else {
        format!(
            "<View Scope=\"RecursiveAll\"><ViewFields><FieldRef Name=\"FileRef\" /><FieldRef Name=\"FileLeafRef\" /><FieldRef Name=\"FSObjType\" /><FieldRef Name=\"FileDirRef\" /></ViewFields><Query><Where><Eq><FieldRef Name=\"FileDirRef\" /><Value Type=\"Text\">{path}</Value></Eq></Where></Query><RowLimit>500</RowLimit></View>"
        )
    };
    let body = json!({"query": {"ViewXml": caml, "FolderServerRelativeUrl": path}});
    let data = match state
        .post_json(&format!("{list_url}/GetItems"), Some(&body), None)
        .await
    {
        Ok(data) => data,
        Err(error) if error.to_string().contains(" 500 ") => {
            let file_url = format!(
                "{}/GetFileByServerRelativePath(DecodedUrl={})",
                state.web_url,
                odata::literal(&path)
            );
            let file = state
                .get_json(&file_url, None)
                .await
                .map_err(runtime_error)?;
            return Python::attach(|py| {
                let file = Py::new(
                    py,
                    SPFile {
                        client,
                        properties: value_dict(py, &file)?,
                        list_url: Some(list_url),
                        item_id: None,
                        unique_id: None,
                        resolve_task: None,
                    },
                )?;
                Ok(PyList::new(py, [file])?.unbind().into_any())
            });
        }
        Err(error) => return Err(runtime_error(error)),
    };
    let values = data
        .get("value")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    Python::attach(|py| Ok(items_list(py, &client, &list_url, &values)?.into_any()))
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
            resolve_task: None,
        }
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

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let label = if self.resolved() {
            properties_label(&self.properties, py)?
        } else {
            "(unresolved)".to_owned()
        };
        Ok(format!("SPFolder: {label}"))
    }

    fn __str__(&self, py: Python<'_>) -> PyResult<String> {
        self.__repr__(py)
    }

    fn resolve<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            resolve_folder(&slf).await.map(|_| ())
        })
    }

    fn ls<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let (client, path) = {
            let folder = slf.bind(py).borrow();
            let client = folder
                .client
                .as_ref()
                .map(|client| client.clone_ref(py))
                .ok_or_else(|| PyRuntimeError::new_err("folder has no client"))?;
            (client, folder.server_relative_url.clone())
        };
        if let Some(path) = path {
            return ls_awaitable(py, client, Some(path));
        }
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_folder(&slf).await?;
            ls_items(state, client, Some(path)).await
        })
    }

    fn get_url(&self, py: Python<'_>) -> PyResult<String> {
        let path = self
            .server_relative_url
            .as_deref()
            .ok_or_else(|| PyRuntimeError::new_err("folder has no server_relative_url"))?;
        let client = self
            .client
            .as_ref()
            .ok_or_else(|| PyRuntimeError::new_err("folder has no client"))?;
        browser_url(&client.bind(py).borrow().state.site_url, path, false)
    }

    #[pyo3(signature = (*, ignore_missing=false, recursive=false))]
    fn del_folder<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        ignore_missing: bool,
        recursive: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_folder(&slf).await?;
            delete_folder_path(&state, &path, ignore_missing, recursive).await?;
            Python::attach(|py| Ok(py.None()))
        })
    }
}

#[pyclass(module = "async_sharepoint")]
struct SPFile {
    client: Py<SharePointClient>,
    properties: Py<PyDict>,
    #[pyo3(get, set)]
    list_url: Option<String>,
    #[pyo3(get, set)]
    item_id: Option<String>,
    #[pyo3(get, set)]
    unique_id: Option<String>,
    resolve_task: Option<tokio::task::JoinHandle<PyResult<()>>>,
}

impl Drop for SPFile {
    fn drop(&mut self) {
        if let Some(handle) = self.resolve_task.take()
            && !handle.is_finished()
        {
            handle.abort();
        }
    }
}

fn properties_server_relative_path(properties: &Py<PyDict>, py: Python<'_>) -> Option<String> {
    properties
        .bind(py)
        .get_item("ServerRelativeUrl")
        .ok()
        .flatten()
        .and_then(|value| value.extract::<String>().ok())
}

/// Extracts the file GUID used to build an embed/preview URL, preferring the
/// unparsed `UniqueId`, then falling back to the GUID embedded in `ContentTag` or `ETag`
/// (both formatted like `{GUID},...`).
fn properties_embed_guid(properties: &Py<PyDict>, py: Python<'_>) -> PyResult<Option<String>> {
    let dict = properties.bind(py);
    if let Some(value) = dict.get_item("UniqueId")?
        && value.is_truthy()?
    {
        return Ok(Some(value.extract()?));
    }
    for key in ["ContentTag", "ETag"] {
        if let Some(value) = dict.get_item(key)?
            && value.is_truthy()?
        {
            let text: String = value.extract()?;
            if let Some(start) = text.find('{')
                && let Some(end) = text[start..].find('}').map(|end| start + end)
            {
                return Ok(Some(text[start + 1..end].to_owned()));
            }
        }
    }
    Ok(None)
}

fn start_file_resolution(
    py: Python<'_>,
    file: &Py<SPFile>,
) -> PyResult<tokio::task::JoinHandle<PyResult<()>>> {
    let file = file.bind(py).borrow();
    let state = Arc::clone(&file.client.bind(py).borrow().state);
    let properties = file.properties.clone_ref(py);
    let (file_url, list_item_url) =
        if let Some(path) = properties_server_relative_path(&file.properties, py) {
            let file_url = format!(
                "{}/GetFileByServerRelativePath(DecodedUrl={})",
                state.web_url,
                odata::literal(&path)
            );
            (
                file_url.clone(),
                Some(format!("{file_url}/ListItemAllFields")),
            )
        } else if let Some(unique_id) = &file.unique_id {
            let file_url = format!("{}/GetFileById(guid'{}')", state.web_url, unique_id);
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
        Python::attach(|py| merge_properties(py, &properties, &file_data))?;
        if let Some(url) = list_item_url {
            let params = [(
                "$select".to_owned(),
                "*,ServerRedirectedEmbedUri".to_owned(),
            )];
            let list_item_data = state.get_json(&url, Some(&params)).await?;
            Python::attach(|py| merge_properties(py, &properties, &list_item_data))?;
        }
        Ok(())
    }))
}

async fn resolve_file(file: &Py<SPFile>) -> PyResult<(Arc<ClientState>, String)> {
    let (state, task) = Python::attach(|py| {
        let mut file = file.bind(py).borrow_mut();
        let state = Arc::clone(&file.client.bind(py).borrow().state);
        Ok::<_, PyErr>((state, file.resolve_task.take()))
    })?;
    if let Some(task) = task {
        task.await.map_err(runtime_error)??;
    }
    let path = Python::attach(|py| {
        let file = file.bind(py).borrow();
        properties_server_relative_path(&file.properties, py)
    });
    path.map(|path| (state, path)).ok_or_else(|| {
        PyRuntimeError::new_err("file has no server_relative_path even after resolving")
    })
}

#[pymethods]
impl SPFile {
    #[new]
    #[pyo3(signature = (client, server_relative_path=None, properties=None, list_url=None, item_id=None, unique_id=None))]
    fn new(
        py: Python<'_>,
        client: Py<SharePointClient>,
        server_relative_path: Option<String>,
        properties: Option<Py<PyDict>>,
        list_url: Option<String>,
        item_id: Option<String>,
        unique_id: Option<String>,
    ) -> PyResult<Py<Self>> {
        let properties = properties.unwrap_or_else(|| PyDict::new(py).unbind());
        if let Some(path) = server_relative_path {
            properties.bind(py).set_item("ServerRelativeUrl", path)?;
        }
        let file = Py::new(
            py,
            Self {
                client,
                properties,
                list_url,
                item_id,
                unique_id,
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

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let label = if self.resolved() {
            properties_label(&self.properties, py)?
        } else {
            "(unresolved)".to_owned()
        };
        Ok(format!("SPFile:   {label}"))
    }

    fn __str__(&self, py: Python<'_>) -> PyResult<String> {
        self.__repr__(py)
    }

    fn resolve<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            resolve_file(&slf).await.map(|_| ())
        })
    }

    fn download<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_file(&slf).await?;
            let url = download_url(&state, FileLocator::Path(path));
            let content = state.get_bytes(&url).await?;
            Python::attach(|py| Ok(PyBytes::new(py, &content).unbind()))
        })
    }

    /// Async context manager streaming this file's contents in `Range`-request chunks;
    /// use `async with file.download_chunks() as session:` and `await session.get_chunk()`.
    fn download_chunks(slf: Py<Self>, py: Python<'_>) -> PyResult<Py<DownloadChunks>> {
        Py::new(
            py,
            DownloadChunks {
                state: None,
                url: None,
                pending_file: Some(slf),
                position: 0,
                total: None,
                finished: false,
            },
        )
    }

    /// Downloads this file to `local_path`, streaming it in chunks instead of buffering it
    /// in memory. Returns `None`.
    fn download_file<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        local_path: String,
    ) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_file(&slf).await?;
            let url = download_url(&state, FileLocator::Path(path));
            download_url_to_file(&state, &url, &local_path).await?;
            Python::attach(|py| Ok(py.None()))
        })
    }

    fn del_file<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, path) = resolve_file(&slf).await?;
            delete_file_locator(&state, FileLocator::Path(path)).await?;
            Python::attach(|py| Ok(py.None()))
        })
    }

    fn get_url(&self, py: Python<'_>) -> PyResult<String> {
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

    fn browser_url(&self, py: Python<'_>) -> PyResult<String> {
        let path = properties_server_relative_path(&self.properties, py).ok_or_else(|| {
            PyRuntimeError::new_err("file has no server_relative_path even after resolving")
        })?;
        browser_url(&self.client.bind(py).borrow().state.site_url, &path, true)
    }

    /// Returns `ServerRedirectedEmbedUri` if present, otherwise builds a Doc.aspx preview
    /// link from the file's GUID (`UniqueId`, else `ContentTag`/`ETag`).
    fn embed_url(&self, py: Python<'_>) -> PyResult<String> {
        if let Some(value) = self
            .properties
            .bind(py)
            .get_item("ServerRedirectedEmbedUri")?
            && value.is_truthy()?
        {
            return value.extract();
        }
        let guid = properties_embed_guid(&self.properties, py)?.ok_or_else(|| {
            PyRuntimeError::new_err(
                "file has no ServerRedirectedEmbedUri, UniqueId, ContentTag, or ETag to build an embed url",
            )
        })?;
        let site_url = self.client.bind(py).borrow().state.site_url.clone();
        let mut url = url::Url::parse(&site_url).map_err(runtime_error)?;
        url.set_path(&format!(
            "{}/_layouts/15/Doc.aspx",
            url.path().trim_end_matches('/')
        ));
        url.query_pairs_mut()
            .append_pair("sourcedoc", &format!("{{{guid}}}"))
            .append_pair("action", "interactivepreview");
        Ok(url.into())
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

    #[pyo3(signature = (*, caml=None, folder_path=None, max_wait=None))]
    fn get_items<'py>(
        &self,
        py: Python<'py>,
        caml: Option<String>,
        folder_path: Option<String>,
        max_wait: Option<f64>,
    ) -> PyResult<Bound<'py, PyAny>> {
        get_items_awaitable(
            py,
            self.client.clone_ref(py),
            None,
            Some(self.id.clone()),
            caml,
            folder_path,
            max_wait,
        )
    }

    fn get_url(&self, py: Python<'_>) -> PyResult<String> {
        let site_url = &self.client.bind(py).borrow().state.site_url;
        let mut url = url::Url::parse(site_url).map_err(runtime_error)?;
        url.set_path(&format!(
            "{}/{}/Forms/AllItems.aspx",
            url.path().trim_end_matches('/'),
            self.title
        ));
        Ok(url.into())
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
        let folder = Py::new(
            py,
            SPFolder {
                client: Some(client),
                server_relative_url: odata::string(value, "ServerRelativeUrl")
                    .or_else(|| odata::string(value, "FileRef")),
                properties,
                list_url: Some(list_url),
                item_id: Some(id),
                resolve_task: None,
            },
        )?;
        let handle = start_folder_resolution(py, &folder)?;
        folder.bind(py).borrow_mut().resolve_task = Some(handle);
        return Ok(folder.into_any());
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
            properties,
            list_url: Some(list_url),
            item_id: Some(id),
            unique_id: None,
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
                Some(url) => fetch_all(Arc::clone(&state), url, None).await?,
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
    folder_path: Option<String>,
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
            let mut query = json!({"ViewXml": caml});
            if let Some(folder_path) = folder_path {
                query["FolderServerRelativeUrl"] = Value::String(folder_path);
            }
            let body = json!({"query": query});
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
                fetch_all(Arc::clone(&state), format!("{list_url}/items"), None).await?,
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

/// Splits `full_path` (e.g. `/sites/team/Documents/report.txt`) into its folder path and filename.
fn split_full_path(full_path: &str) -> PyResult<(String, String)> {
    match full_path.rsplit_once('/') {
        Some((folder_path, filename)) => Ok((browser_path(folder_path)?, filename.to_owned())),
        None => Err(PyValueError::new_err("Invalid full_path")),
    }
}

fn is_not_found(error: &Error) -> bool {
    matches!(error, Error::Http { status, .. } if status.as_u16() == 404)
}

async fn ensure_folder_path(state: &Arc<ClientState>, folder_path: &str) -> Result<(), Error> {
    let site_path = url::Url::parse(&state.site_url)
        .map_err(|error| Error::InvalidUrl(error.to_string()))?
        .path()
        .trim_end_matches('/')
        .to_owned();
    let relative_path = if site_path.is_empty() {
        folder_path.trim_matches('/')
    } else if folder_path == site_path {
        ""
    } else {
        folder_path
            .strip_prefix(&format!("{site_path}/"))
            .ok_or_else(|| {
                Error::InvalidUrl(format!(
                    "folder path {folder_path} is outside site {site_path}"
                ))
            })?
    };
    let mut current = site_path;
    let mut paths = Vec::new();
    if !current.is_empty() {
        paths.push(current.clone());
    }
    for component in relative_path
        .split('/')
        .filter(|component| !component.is_empty())
    {
        current.push('/');
        current.push_str(component);
        paths.push(current.clone());
    }

    let mut probes = JoinSet::new();
    for (index, path) in paths.iter().cloned().enumerate() {
        let state = Arc::clone(state);
        probes.spawn(async move {
            let url = format!(
                "{}/GetFolderByServerRelativePath(DecodedUrl={})",
                state.web_url,
                odata::literal(&path)
            );
            (index, path, state.get_json(&url, None).await)
        });
    }

    let mut deepest_existing = None;
    while let Some(result) = probes.join_next().await {
        let (index, _path, result) =
            result.map_err(|error| Error::Unexpected(error.to_string()))?;
        match result {
            Ok(_) => {
                deepest_existing =
                    Some(deepest_existing.map_or(index, |current: usize| current.max(index)))
            }
            Err(error) if is_not_found(&error) => {}
            Err(error) => return Err(error),
        }
    }
    let first_missing = deepest_existing.map(|index| index + 1).ok_or_else(|| {
        Error::Unexpected(format!("no existing parent folder found for {folder_path}"))
    })?;
    for path in &paths[first_missing..] {
        let url = format!(
            "{}/Folders/AddUsingPath(DecodedUrl={},Overwrite=true)",
            state.web_url,
            odata::literal(path)
        );
        state.post_json(&url, None, Some(Vec::new())).await?;
    }
    Ok(())
}

async fn with_missing_folder_retry<T, F, Fut>(
    state: &Arc<ClientState>,
    folder_path: &str,
    operation: F,
) -> Result<T, Error>
where
    F: Fn() -> Fut,
    Fut: Future<Output = Result<T, Error>>,
{
    match operation().await {
        Err(error) if is_not_found(&error) => {
            ensure_folder_path(state, folder_path).await?;
            operation().await
        }
        result => result,
    }
}

async fn delete_file_locator(state: &ClientState, locator: FileLocator) -> PyResult<()> {
    let url = match locator {
        FileLocator::Path(path) => format!(
            "{}/GetFileByServerRelativePath(DecodedUrl={})",
            state.web_url,
            odata::literal(&path)
        ),
        FileLocator::Id(id) => format!("{}/GetFileById(guid'{}')", state.web_url, id),
    };
    state.delete(&url).await?;
    Ok(())
}

async fn delete_folder_path(
    state: &Arc<ClientState>,
    path: &str,
    ignore_missing: bool,
    recursive: bool,
) -> PyResult<()> {
    let url = format!(
        "{}/GetFolderByServerRelativePath(DecodedUrl={})",
        state.web_url,
        odata::literal(path)
    );
    let data = match state.get_json(&url, None).await {
        Ok(data) => data,
        Err(error) if ignore_missing && is_not_found(&error) => return Ok(()),
        Err(error) => return Err(error.into()),
    };
    if !recursive {
        if data.get("ItemCount").and_then(Value::as_i64).unwrap_or(0) > 0 {
            return Err(PyRuntimeError::new_err(format!(
                "folder is not empty: {path}; pass recursive=True to delete it"
            )));
        }
        state.delete(&url).await?;
        return Ok(());
    }

    let mut pending = vec![path.to_owned()];
    let mut folders = Vec::new();
    while let Some(folder_path) = pending.pop() {
        let folder_url = format!(
            "{}/GetFolderByServerRelativePath(DecodedUrl={})",
            state.web_url,
            odata::literal(&folder_path)
        );
        let params = Some(vec![("$select".to_owned(), "ServerRelativeUrl".to_owned())]);
        let (files, children) = tokio::try_join!(
            fetch_all(
                Arc::clone(state),
                format!("{folder_url}/Files"),
                params.clone()
            ),
            fetch_all(Arc::clone(state), format!("{folder_url}/Folders"), params)
        )?;
        for file in files {
            let file_path = odata::string(&file, "ServerRelativeUrl").ok_or_else(|| {
                PyRuntimeError::new_err("SharePoint file response has no ServerRelativeUrl")
            })?;
            delete_file_locator(state, FileLocator::Path(file_path)).await?;
        }
        for child in children {
            let child_path = odata::string(&child, "ServerRelativeUrl").ok_or_else(|| {
                PyRuntimeError::new_err("SharePoint folder response has no ServerRelativeUrl")
            })?;
            pending.push(child_path);
        }
        folders.push(folder_url);
    }
    for folder_url in folders.into_iter().rev() {
        state.delete(&folder_url).await?;
    }
    Ok(())
}

fn download_url(state: &ClientState, locator: FileLocator) -> String {
    match locator {
        FileLocator::Path(path) => format!(
            "{}/GetFileByServerRelativePath(DecodedUrl={})/$value",
            state.web_url,
            odata::literal(&path)
        ),
        FileLocator::Id(id) => format!("{}/GetFileById(guid'{}')/$value", state.web_url, id),
    }
}

async fn chunked_upload_add(
    state: &ClientState,
    add_url: &str,
    content: Vec<u8>,
) -> Result<Value, Error> {
    state.post_json(add_url, None, Some(content)).await
}

async fn chunked_upload_start(
    state: &ClientState,
    add_url: &str,
    file_url: &str,
    upload_id: &str,
    chunk: Vec<u8>,
) -> Result<usize, Error> {
    state.post_json(add_url, None, Some(Vec::new())).await?;
    let len = chunk.len();
    state
        .post_json(
            &format!(
                "{file_url}/startUpload(uploadID={})",
                odata::literal(upload_id)
            ),
            None,
            Some(chunk),
        )
        .await?;
    Ok(len)
}

async fn chunked_upload_continue(
    state: &ClientState,
    file_url: &str,
    upload_id: &str,
    position: usize,
    chunk: Vec<u8>,
) -> Result<usize, Error> {
    let len = chunk.len();
    state
        .post_json(
            &format!(
                "{file_url}/continueUpload(uploadID={},fileOffset={position})",
                odata::literal(upload_id)
            ),
            None,
            Some(chunk),
        )
        .await?;
    Ok(position + len)
}

async fn chunked_upload_finish(
    state: &ClientState,
    file_url: &str,
    upload_id: &str,
    position: usize,
    chunk: Vec<u8>,
) -> Result<Value, Error> {
    state
        .post_json(
            &format!(
                "{file_url}/finishUpload(uploadID={},fileOffset={position})",
                odata::literal(upload_id)
            ),
            None,
            Some(chunk),
        )
        .await
}

async fn download_url_to_file(state: &ClientState, url: &str, local_path: &str) -> PyResult<()> {
    let mut file = tokio::fs::File::create(local_path)
        .await
        .map_err(runtime_error)?;
    let mut position: u64 = 0;
    let mut total: Option<u64> = None;
    loop {
        let end = position + DOWNLOAD_CHUNK_SIZE - 1;
        let (data, content_total) = state.get_bytes_range(url, position, end).await?;
        if data.is_empty() {
            break;
        }
        file.write_all(&data).await.map_err(runtime_error)?;
        total = total.or(content_total);
        position += data.len() as u64;
        if (data.len() as u64) < DOWNLOAD_CHUNK_SIZE || total.is_some_and(|value| position >= value)
        {
            break;
        }
    }
    Ok(())
}

async fn upload_file_to_folder(
    state: &Arc<ClientState>,
    folder_path: &str,
    folder_url: &str,
    filename: &str,
    overwrite: bool,
    local_path: &str,
) -> PyResult<()> {
    let add_url = format!(
        "{folder_url}/Files/add(url={},overwrite={})",
        odata::literal(filename),
        odata::bool_literal(overwrite)
    );
    let file_url = format!("{folder_url}/Files({})", odata::literal(filename));
    let upload_id = uuid::Uuid::new_v4().to_string();
    let mut source = tokio::fs::File::open(local_path)
        .await
        .map_err(runtime_error)?;
    let mut buffer = vec![0u8; UPLOAD_CHUNK_SIZE];
    let mut pending: Option<Vec<u8>> = None;
    let mut started = false;
    let mut position = 0usize;
    loop {
        let read = source.read(&mut buffer).await.map_err(runtime_error)?;
        if read == 0 {
            break;
        }
        if let Some(previous) = pending.replace(buffer[..read].to_vec()) {
            if started {
                position =
                    chunked_upload_continue(state, &file_url, &upload_id, position, previous)
                        .await?;
            } else {
                position = with_missing_folder_retry(state, folder_path, || {
                    chunked_upload_start(state, &add_url, &file_url, &upload_id, previous.clone())
                })
                .await?;
                started = true;
            }
        }
    }
    match pending {
        Some(chunk) if started => {
            chunked_upload_finish(state, &file_url, &upload_id, position, chunk).await?;
        }
        Some(chunk) => {
            with_missing_folder_retry(state, folder_path, || {
                chunked_upload_add(state, &add_url, chunk.clone())
            })
            .await?;
        }
        None => {
            with_missing_folder_retry(state, folder_path, || {
                chunked_upload_add(state, &add_url, Vec::new())
            })
            .await?;
        }
    }
    Ok(())
}

/// Async context manager returned by `upload_chunks`; `write()` streams one chunk at a time,
/// buffering the most recent chunk so the final `write` (or `__aexit__`) can be sent as
/// SharePoint's `finishUpload` (or, for a single-chunk upload, a plain `Files/add`).
#[pyclass(module = "async_sharepoint")]
struct UploadChunks {
    state: Arc<ClientState>,
    folder_path: String,
    add_url: String,
    file_url: String,
    upload_id: String,
    position: usize,
    started: bool,
    pending: Option<Vec<u8>>,
    closed: bool,
}

#[pymethods]
impl UploadChunks {
    fn __aenter__<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async move { Ok(slf) })
    }

    fn write<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        chunk: &Bound<'_, PyBytes>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let chunk = chunk.as_bytes().to_vec();
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let (state, folder_path, add_url, file_url, upload_id, position, started, previous) =
                Python::attach(|py| {
                    let mut session = slf.bind(py).borrow_mut();
                    let previous = session.pending.replace(chunk);
                    (
                        Arc::clone(&session.state),
                        session.folder_path.clone(),
                        session.add_url.clone(),
                        session.file_url.clone(),
                        session.upload_id.clone(),
                        session.position,
                        session.started,
                        previous,
                    )
                });
            let Some(previous) = previous else {
                return Ok(());
            };
            let new_position = if started {
                chunked_upload_continue(&state, &file_url, &upload_id, position, previous).await?
            } else {
                with_missing_folder_retry(&state, &folder_path, || {
                    chunked_upload_start(&state, &add_url, &file_url, &upload_id, previous.clone())
                })
                .await?
            };
            Python::attach(|py| {
                let mut session = slf.bind(py).borrow_mut();
                session.started = true;
                session.position = new_position;
            });
            Ok(())
        })
    }

    fn __aexit__<'py>(
        &self,
        py: Python<'py>,
        exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let has_exception = !exc_type.is_none(py);
        let state = Arc::clone(&self.state);
        let folder_path = self.folder_path.clone();
        let add_url = self.add_url.clone();
        let file_url = self.file_url.clone();
        let upload_id = self.upload_id.clone();
        let position = self.position;
        let started = self.started;
        let pending = self.pending.clone();
        let closed = self.closed;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            if closed || has_exception {
                return Ok(false);
            }
            match pending {
                Some(chunk) if started => {
                    chunked_upload_finish(&state, &file_url, &upload_id, position, chunk).await?;
                }
                Some(chunk) => {
                    with_missing_folder_retry(&state, &folder_path, || {
                        chunked_upload_add(&state, &add_url, chunk.clone())
                    })
                    .await?;
                }
                None if !started => {
                    with_missing_folder_retry(&state, &folder_path, || {
                        chunked_upload_add(&state, &add_url, Vec::new())
                    })
                    .await?;
                }
                None => {}
            }
            Ok(false)
        })
    }
}

/// Async context manager returned by `download_chunks`; `get_chunk()` streams one range at a
/// time from SharePoint's `/$value` endpoint (via HTTP `Range` requests) and returns `None`
/// once the file has been fully read.
#[pyclass(module = "async_sharepoint")]
struct DownloadChunks {
    state: Option<Arc<ClientState>>,
    url: Option<String>,
    pending_file: Option<Py<SPFile>>,
    position: u64,
    total: Option<u64>,
    finished: bool,
}

#[pymethods]
impl DownloadChunks {
    fn __aenter__<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let pending_file = slf
            .bind(py)
            .borrow()
            .pending_file
            .as_ref()
            .map(|file| file.clone_ref(py));
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            if let Some(file) = pending_file {
                let (state, path) = resolve_file(&file).await?;
                let url = download_url(&state, FileLocator::Path(path));
                Python::attach(|py| {
                    let mut session = slf.bind(py).borrow_mut();
                    session.state = Some(state);
                    session.url = Some(url);
                    session.pending_file = None;
                });
            }
            Ok(slf)
        })
    }

    fn __aexit__<'py>(
        &self,
        py: Python<'py>,
        _exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        pyo3_async_runtimes::tokio::future_into_py(py, async { Ok(false) })
    }

    fn get_chunk<'py>(slf: Py<Self>, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let (state, url, position, total, finished) = Python::attach(|py| {
            let session = slf.bind(py).borrow();
            (
                session.state.as_ref().map(Arc::clone),
                session.url.clone(),
                session.position,
                session.total,
                session.finished,
            )
        });
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            if finished {
                return Python::attach(|py| Ok(py.None()));
            }
            let state = state.ok_or_else(|| {
                PyRuntimeError::new_err("download_chunks session was not entered")
            })?;
            let url = url.ok_or_else(|| {
                PyRuntimeError::new_err("download_chunks session was not entered")
            })?;
            let end = position + DOWNLOAD_CHUNK_SIZE - 1;
            let (data, content_total) = state.get_bytes_range(&url, position, end).await?;
            let received = data.len() as u64;
            let new_total = total.or(content_total);
            let new_position = position + received;
            let is_finished = received == 0
                || received < DOWNLOAD_CHUNK_SIZE
                || new_total.is_some_and(|value| new_position >= value);
            Python::attach(|py| {
                let mut session = slf.bind(py).borrow_mut();
                session.position = new_position;
                session.total = new_total;
                session.finished = is_finished;
            });
            if data.is_empty() {
                Python::attach(|py| Ok(py.None()))
            } else {
                Python::attach(|py| Ok(PyBytes::new(py, &data).into_any().unbind()))
            }
        })
    }
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
    #[pyo3(signature = (site_url, credential, *, properties=None, item_id=None, list_url=None))]
    fn new(
        py: Python<'_>,
        site_url: &str,
        credential: &CertificateCredential,
        properties: Option<Py<PyDict>>,
        item_id: Option<String>,
        list_url: Option<String>,
    ) -> PyResult<Self> {
        Ok(Self {
            state: ClientState::new_with_runtime(
                site_url.trim_end_matches('/').to_owned(),
                &credential.inner,
                pyo3_async_runtimes::tokio::get_runtime().handle(),
            )?,
            properties: properties.unwrap_or_else(|| PyDict::new(py).unbind()),
            item_id,
            list_url,
        })
    }

    /// Builds a client around a pre-obtained token; it is never refreshed.
    #[staticmethod]
    #[pyo3(signature = (site_url, token, *, properties=None, item_id=None, list_url=None))]
    fn from_static_token(
        py: Python<'_>,
        site_url: &str,
        token: String,
        properties: Option<Py<PyDict>>,
        item_id: Option<String>,
        list_url: Option<String>,
    ) -> PyResult<Self> {
        Ok(Self {
            state: ClientState::with_static_token(
                site_url.trim_end_matches('/').to_owned(),
                token,
            )?,
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
        let state = Arc::clone(&slf.bind(py).borrow().state);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            state.tokens.wait_ready().await?;
            Ok(slf)
        })
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
                    fetch_all(Arc::clone(&state), format!("{}/lists", state.web_url), None).await?,
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
                                Some(url) => fetch_all(Arc::clone(&state), url, None).await?,
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

    #[pyo3(signature = (title=None, *, id=None, caml=None, folder_path=None, max_wait=None))]
    fn get_items<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        title: Option<String>,
        id: Option<String>,
        caml: Option<String>,
        folder_path: Option<String>,
        max_wait: Option<f64>,
    ) -> PyResult<Bound<'py, PyAny>> {
        get_items_awaitable(py, slf, title, id, caml, folder_path, max_wait)
    }

    fn get_file<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        path: Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let display = path.str()?.to_string();
        let locator = file_locator_from_py(&path)?;
        let (server_relative_path, unique_id) = match locator {
            FileLocator::Path(path) => (Some(path), None),
            FileLocator::Id(id) => (None, Some(id)),
        };
        let properties = PyDict::new(py).unbind();
        if let Some(path) = &server_relative_path {
            properties.bind(py).set_item("ServerRelativeUrl", path)?;
        }
        let file = Py::new(
            py,
            SPFile {
                client: slf,
                properties,
                list_url: None,
                item_id: None,
                unique_id,
                resolve_task: None,
            },
        )?;
        let handle = start_file_resolution(py, &file)?;
        file.bind(py).borrow_mut().resolve_task = Some(handle);
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            if let Err(error) = resolve_file(&file).await {
                if error.to_string().contains(" 404 ") {
                    return Err(PyRuntimeError::new_err(format!(
                        "Either {display} does not exist or is a folder"
                    )));
                }
                return Err(error);
            }
            Ok(file)
        })
    }

    fn download<'py>(
        &self,
        py: Python<'py>,
        path: Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let locator = file_locator_from_py(&path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let url = download_url(&state, locator);
            let content = state.get_bytes(&url).await?;
            Python::attach(|py| Ok(PyBytes::new(py, &content).unbind()))
        })
    }

    /// Async context manager streaming a file's contents in `Range`-request chunks;
    /// use `async with client.download_chunks(path) as session:` and
    /// `await session.get_chunk()` until it returns `None`.
    fn download_chunks(
        &self,
        py: Python<'_>,
        path: Bound<'_, PyAny>,
    ) -> PyResult<Py<DownloadChunks>> {
        let locator = file_locator_from_py(&path)?;
        let url = download_url(&self.state, locator);
        Py::new(
            py,
            DownloadChunks {
                state: Some(Arc::clone(&self.state)),
                url: Some(url),
                pending_file: None,
                position: 0,
                total: None,
                finished: false,
            },
        )
    }

    /// Downloads `path` to `local_path`, streaming it in chunks instead of buffering it in
    /// memory. Returns `None`.
    fn download_file<'py>(
        &self,
        py: Python<'py>,
        path: Bound<'py, PyAny>,
        local_path: String,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let locator = file_locator_from_py(&path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let url = download_url(&state, locator);
            download_url_to_file(&state, &url, &local_path).await?;
            Python::attach(|py| Ok(py.None()))
        })
    }

    #[pyo3(signature = (full_path,  content, *, overwrite=true))]
    fn upload<'py>(
        &self,
        py: Python<'py>,
        full_path: String,
        content: &Bound<'_, PyBytes>,
        overwrite: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let (folder_path, filename) = split_full_path(&full_path)?;
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
                with_missing_folder_retry(&state, &folder_path, || {
                    chunked_upload_add(&state, &add_url, content.clone())
                })
                .await?
            } else {
                let file_url = format!("{folder_url}/Files({})", odata::literal(&filename));
                let upload_id = uuid::Uuid::new_v4().to_string();
                let first_chunk = content[..UPLOAD_CHUNK_SIZE].to_vec();
                let mut position = with_missing_folder_retry(&state, &folder_path, || {
                    chunked_upload_start(
                        &state,
                        &add_url,
                        &file_url,
                        &upload_id,
                        first_chunk.clone(),
                    )
                })
                .await?;
                loop {
                    let end = (position + UPLOAD_CHUNK_SIZE).min(content.len());
                    let chunk = content[position..end].to_vec();
                    if end < content.len() {
                        position =
                            chunked_upload_continue(&state, &file_url, &upload_id, position, chunk)
                                .await?;
                    } else {
                        break chunked_upload_finish(
                            &state, &file_url, &upload_id, position, chunk,
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
                        properties: value_dict(py, &data)?,
                        list_url: None,
                        item_id: None,
                        unique_id: None,
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

    /// Async context manager streaming a file upload one chunk at a time; use
    /// `async with client.upload_chunks(full_path, overwrite=...) as session:` and
    /// `await session.write(chunk)` for each chunk.
    #[pyo3(signature = (full_path, *, overwrite=true))]
    fn upload_chunks(
        &self,
        py: Python<'_>,
        full_path: String,
        overwrite: bool,
    ) -> PyResult<Py<UploadChunks>> {
        let (folder_path, filename) = split_full_path(&full_path)?;
        let folder_url = format!(
            "{}/GetFolderByServerRelativeUrl({})",
            self.state.web_url,
            odata::literal(&folder_path)
        );
        let add_url = format!(
            "{folder_url}/Files/add(url={},overwrite={})",
            odata::literal(&filename),
            odata::bool_literal(overwrite)
        );
        let file_url = format!("{folder_url}/Files({})", odata::literal(&filename));
        Py::new(
            py,
            UploadChunks {
                state: Arc::clone(&self.state),
                folder_path,
                add_url,
                file_url,
                upload_id: uuid::Uuid::new_v4().to_string(),
                position: 0,
                started: false,
                pending: None,
                closed: false,
            },
        )
    }

    /// Uploads `local_path` to `full_path`, streaming it in chunks instead of buffering it
    /// in memory. Returns `None`.
    #[pyo3(signature = (full_path, local_path, *, overwrite=true))]
    fn upload_file<'py>(
        &self,
        py: Python<'py>,
        full_path: String,
        local_path: String,
        overwrite: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let (folder_path, filename) = split_full_path(&full_path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let folder_url = format!(
                "{}/GetFolderByServerRelativeUrl({})",
                state.web_url,
                odata::literal(&folder_path)
            );
            upload_file_to_folder(
                &state,
                &folder_path,
                &folder_url,
                &filename,
                overwrite,
                &local_path,
            )
            .await?;
            Python::attach(|py| Ok(py.None()))
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
        let path = browser_path(&path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let url = format!(
                "{}/Folders/AddUsingPath(DecodedUrl={},Overwrite={})",
                state.web_url,
                odata::literal(&path),
                odata::bool_literal(overwrite)
            );
            let data = state.post_json(&url, None, Some(Vec::new())).await?;
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
                        resolve_task: None,
                    },
                )
            })
        })
    }

    fn del_file<'py>(
        &self,
        py: Python<'py>,
        path: Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let locator = file_locator_from_py(&path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            delete_file_locator(&state, locator).await?;
            Python::attach(|py| Ok(py.None()))
        })
    }

    #[pyo3(signature = (path, *, ignore_missing=false, recursive=false))]
    fn del_folder<'py>(
        &self,
        py: Python<'py>,
        path: String,
        ignore_missing: bool,
        recursive: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let state = Arc::clone(&self.state);
        let path = browser_path(&path)?;
        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            delete_folder_path(&state, &path, ignore_missing, recursive).await?;
            Python::attach(|py| Ok(py.None()))
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

    #[pyo3(signature = (path=None))]
    fn ls<'py>(
        slf: Py<Self>,
        py: Python<'py>,
        path: Option<String>,
    ) -> PyResult<Bound<'py, PyAny>> {
        ls_awaitable(py, slf, path)
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
    module.add_class::<CertificateCredential>()?;
    module.add_class::<SharePointClient>()?;
    module.add_class::<SPFile>()?;
    module.add_class::<SPFolder>()?;
    module.add_class::<SPList>()?;
    module.add_class::<UploadChunks>()?;
    module.add_class::<DownloadChunks>()?;
    let sp_item = module
        .getattr("SPFile")?
        .call_method1("__or__", (module.getattr("SPFolder")?,))?
        .call_method1("__or__", (module.getattr("SharePointClient")?,))?;
    module.add("SPItem", sp_item)?;
    module.add(
        "__all__",
        vec![
            "CertificateCredential",
            "SPFile",
            "SPFolder",
            "SPItem",
            "SharePointClient",
        ],
    )?;
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
