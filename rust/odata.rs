use url::Url;

use serde_json::Value;

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
            if key != "$skiptoken" {
                return (key.into_owned(), value.into_owned());
            }
            let new_value = value
                .split('&')
                .map(|part| {
                    if part.starts_with("p_ID=") {
                        format!("p_ID={new_id}")
                    } else {
                        part.to_owned()
                    }
                })
                .collect::<Vec<_>>()
                .join("&");
            (key.into_owned(), new_value)
        })
        .collect();
    url.query_pairs_mut().clear().extend_pairs(&pairs);
    url.to_string()
}

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
            let values = cells
                .into_iter()
                .flatten()
                .filter_map(|cell| {
                    Some((
                        cell.get("Key")?.as_str()?.to_owned(),
                        cell.get("Value")?.clone(),
                    ))
                })
                .collect();
            Value::Object(values)
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;

    #[test]
    fn escapes_sharepoint_literals() {
        assert_eq!(literal("A%+#&'B"), "'A%25%2B%23%26''B'");
        assert_eq!(bool_literal(true), "true");
        assert_eq!(bool_literal(false), "false");
    }

    #[test]
    fn parses_and_rewrites_skiptoken_p_id() {
        let url = "https://example.com/items?%24skiptoken=Paged%3dTRUE%26p_ID%3d101";
        assert_eq!(skiptoken_p_id(url), Some(101));

        let rewritten = with_p_id(url, 201);
        assert_eq!(skiptoken_p_id(&rewritten), Some(201));
        assert!(rewritten.contains("Paged%3DTRUE"));
    }

    #[test]
    fn normalizes_both_search_payload_shapes() {
        let table = json!({"RelevantResults": {"Table": {"Rows": [
            {"Cells": {"results": [{"Key": "Title", "Value": "Report"}]}}
        ]}}});
        let direct = json!({"PrimaryQueryResult": table});
        let verbose = json!({"d": {"query": {"PrimaryQueryResult": table}}});
        assert_eq!(search_rows(&direct), vec![json!({"Title": "Report"})]);
        assert_eq!(search_rows(&verbose), vec![json!({"Title": "Report"})]);
    }
}
