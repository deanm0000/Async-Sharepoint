use serde_json::Value;
use url::Url;

/// Extract the `p_ID` value from a `$skiptoken` query param, if present.
pub fn skiptoken_p_id(url_str: &str) -> Option<u64> {
    let url = Url::parse(url_str).ok()?;
    let skiptoken = url.query_pairs().find(|(key, _)| key == "$skiptoken")?.1;
    skiptoken
        .split('&')
        .find_map(|part| part.strip_prefix("p_ID="))?
        .parse()
        .ok()
}

/// Rebuild `url_str` with its `$skiptoken` `p_ID` replaced by `new_id`.
pub fn with_p_id(url_str: &str, new_id: u64) -> String {
    let Ok(mut url) = Url::parse(url_str) else {
        return url_str.to_owned();
    };
    let pairs: Vec<(String, String)> = url
        .query_pairs()
        .map(|(key, value)| {
            let value = if key == "$skiptoken" {
                value
                    .split('&')
                    .map(|part| match part.strip_prefix("p_ID=") {
                        Some(_) => format!("p_ID={new_id}"),
                        None => part.to_owned(),
                    })
                    .collect::<Vec<_>>()
                    .join("&")
            } else {
                value.into_owned()
            };
            (key.into_owned(), value)
        })
        .collect();
    url.query_pairs_mut().clear().extend_pairs(&pairs);
    url.to_string()
}

/// Escape a SharePoint OData string literal.
pub fn literal(value: &str) -> String {
    let escaped = value
        .replace('%', "%25")
        .replace('+', "%2B")
        .replace('#', "%23")
        .replace('&', "%26")
        .replace('\'', "''");
    format!("'{escaped}'")
}

pub const fn bool_literal(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

pub fn string(value: &Value, key: &str) -> Option<String> {
    value.get(key).and_then(Value::as_str).map(str::to_owned)
}

/// Normalize either supported SharePoint search response shape into row objects.
pub fn search_rows(payload: &Value) -> Vec<Value> {
    let result = payload
        .get("PrimaryQueryResult")
        .or_else(|| payload.get("d")?.get("query")?.get("PrimaryQueryResult"));
    let rows = result
        .and_then(|result| result.get("RelevantResults"))
        .and_then(|results| results.get("Table"))
        .and_then(|table| table.get("Rows"))
        .and_then(Value::as_array);
    rows.into_iter()
        .flatten()
        .map(|row| {
            let cells = row
                .get("Cells")
                .and_then(|cells| cells.get("results").or(Some(cells)))
                .and_then(Value::as_array);
            Value::Object(
                cells
                    .into_iter()
                    .flatten()
                    .filter_map(|cell| {
                        Some((
                            cell.get("Key")?.as_str()?.to_owned(),
                            cell.get("Value")?.clone(),
                        ))
                    })
                    .collect(),
            )
        })
        .collect()
}

/// Extract a SharePoint browser `id` query parameter as a server-relative path.
pub fn browser_path(value: &str) -> Option<String> {
    let url = Url::parse(value).ok()?;
    if !matches!(url.scheme(), "http" | "https") {
        return None;
    }
    url.query_pairs()
        .find(|(key, _)| key == "id")
        .map(|(_, path)| path.into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extracts_a_browser_path() {
        assert_eq!(
            browser_path(
                "https://contoso.sharepoint.com/sites/team/Docs/Forms/AllItems.aspx?id=%2Fsites%2Fteam%2FDocs%2Fhello.txt"
            ),
            Some("/sites/team/Docs/hello.txt".to_owned())
        );
    }

    #[test]
    fn leaves_non_browser_paths_alone() {
        assert_eq!(browser_path("/sites/team/Docs"), None);
    }
}
