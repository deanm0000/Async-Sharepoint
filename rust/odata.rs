use serde_json::Value;

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
