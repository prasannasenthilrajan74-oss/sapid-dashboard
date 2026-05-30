use serde::{Deserialize, Serialize};
use url::Url;

// No whitelists. Any google.com sheets/scripts are accepted.

// ─── Response Types ───────────────────────────────────────────────────────────
#[derive(Serialize, Deserialize, Debug)]
pub struct ApiResponse {
    pub success: bool,
    pub message: Option<String>,
    pub data: Option<serde_json::Value>,
    pub csv: Option<String>,
    pub is_json: Option<bool>,
}

// ─── URL Helpers (mirrors server.js logic) ────────────────────────────────────

/// Check if URL is any Apps Script exec endpoint
fn is_apps_script_url(url_str: &str) -> bool {
    url_str.contains("script.google.com/a/macros/")
        || url_str.contains("script.google.com/macros/s/")
}

/// Check if URL is a domain-restricted Apps Script (/a/macros/)
fn is_domain_script_url(url_str: &str) -> bool {
    url_str.contains("script.google.com/a/macros/")
}

/// Extract Google Sheets file ID from URL
fn extract_sheet_id(url_str: &str) -> Option<String> {
    let re_pattern = "/spreadsheets/d/";
    if let Some(pos) = url_str.find(re_pattern) {
        let after = &url_str[pos + re_pattern.len()..];
        let id: String = after
            .chars()
            .take_while(|c| c.is_alphanumeric() || *c == '-' || *c == '_')
            .collect();
        if !id.is_empty() {
            return Some(id);
        }
    }
    None
}


/// Convert a standard Google Sheets URL to its CSV export equivalent
fn to_csv_export_url(url_str: &str) -> String {
    if url_str.contains("/pub") {
        if url_str.contains("output=csv") {
            return url_str.to_string();
        }
        // Strip query/fragment from /pub and add output=csv
        if let Some(pos) = url_str.find("/pub") {
            return format!("{}/pub?output=csv", &url_str[..pos]);
        }
    }
    // /d/<ID>/... → gviz/tq CSV
    if let Some(sheet_id) = extract_sheet_id(url_str) {
        let gid = if let Some(pos) = url_str.find("gid=") {
            let gid_str = &url_str[pos + 4..];
            gid_str
                .chars()
                .take_while(|c| c.is_ascii_digit())
                .collect()
        } else {
            "0".to_string()
        };
        return format!(
            "https://docs.google.com/spreadsheets/d/{}/gviz/tq?tqx=out:csv&gid={}",
            sheet_id, gid
        );
    }
    url_str.to_string()
}

/// Validate that a URL is a google.com host (SSRF protection)
fn validate_google_host(url_str: &str) -> Result<Url, String> {
    let parsed = Url::parse(url_str).map_err(|_| "Invalid URL format.".to_string())?;
    let host = parsed.host_str().unwrap_or("");
    if !host.ends_with("google.com") {
        return Err("Invalid URL. Host must be google.com.".to_string());
    }
    Ok(parsed)
}

// ─── Tauri Commands ───────────────────────────────────────────────────────────

/// Validate a URL against the whitelist (mirrors GET /api/validate)
/// Used by the frontend for domain-restricted /a/macros/ URLs that
/// the browser fetches directly with the user's Google session.
#[tauri::command]
async fn validate_url(url: String) -> Result<ApiResponse, String> {
    // SSRF protection
    validate_google_host(&url).map_err(|e| e)?;

    Ok(ApiResponse {
        success: true,
        message: None,
        data: None,
        csv: None,
        is_json: None,
    })
}

/// Fetch Google Sheets / Apps Script data (mirrors GET /api/data)
/// Handles both JSON (Apps Script) and CSV (Google Sheets) responses.
/// Domain-restricted /a/macros/ URLs are validated only — not fetched here
/// (they must be fetched by the browser WebView with the user's Google session).
#[tauri::command]
async fn fetch_sheet_data(url: String) -> Result<ApiResponse, String> {
    let default_url = "https://script.google.com/macros/s/AKfycbytaI36PDf09D7O2RicMWEkGn-JXiew3zPL6bc3OLGKTc0klmd0gUj9ZCfdg2JvY9Sb/exec";
    let sheet_url = if url.trim().is_empty() {
        default_url.to_string()
    } else {
        url.trim().to_string()
    };

    // SSRF protection
    validate_google_host(&sheet_url).map_err(|e| e)?;

    let is_apps_script = is_apps_script_url(&sheet_url);
    let is_domain_script = is_domain_script_url(&sheet_url);

    // Domain-restricted URLs: validate whitelist but don't fetch — the browser handles it.
    // Return a special flag so the frontend knows to fetch directly.
    // Domain-restricted URLs: validate whitelist but don't fetch — the browser handles it.
    // Return a special flag so the frontend knows to fetch directly.
    if is_domain_script {
        // Signal the frontend to fetch directly (domain-restricted — needs user's Google session)
        return Ok(ApiResponse {
            success: true,
            message: Some("FETCH_DIRECT".to_string()),
            data: None,
            csv: None,
            is_json: Some(true),
        });
    }

    // Convert to exportable URL for non-Apps-Script sheets
    let target_url = if !is_apps_script {
        to_csv_export_url(&sheet_url)
    } else {
        sheet_url.clone()
    };

    // Fetch from Google
    let client = reqwest::Client::builder()
        .danger_accept_invalid_certs(false)
        .build()
        .map_err(|e| format!("Failed to create HTTP client: {}", e))?;

    let response = client
        .get(&target_url)
        .header("User-Agent", "SkillTrack-Analyzer/1.0")
        .send()
        .await
        .map_err(|e| format!("Failed to fetch sheet data: {}", e))?;

    if !response.status().is_success() {
        return Ok(ApiResponse {
            success: false,
            message: Some(format!(
                "Google responded with status {}",
                response.status()
            )),
            data: None,
            csv: None,
            is_json: None,
        });
    }

    if is_apps_script {
        // Parse as JSON
        let json_data: serde_json::Value = response
            .json()
            .await
            .map_err(|e| format!("Failed to parse JSON response: {}", e))?;

        Ok(ApiResponse {
            success: true,
            message: None,
            data: Some(json_data),
            csv: None,
            is_json: Some(true),
        })
    } else {
        // Return as CSV text
        let csv_text = response
            .text()
            .await
            .map_err(|e| format!("Failed to read CSV response: {}", e))?;

        Ok(ApiResponse {
            success: true,
            message: None,
            data: None,
            csv: Some(csv_text),
            is_json: Some(false),
        })
    }
}

// ─── App Entry ────────────────────────────────────────────────────────────────
#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![fetch_sheet_data, validate_url])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
