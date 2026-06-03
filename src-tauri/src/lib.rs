use serde::{Deserialize, Serialize};
use std::fs::{create_dir_all, OpenOptions};
use std::io::{Read, Write};
use std::time::Duration;
use tauri::Manager;
use url::Url;

const MAX_REMOTE_RESPONSE_BYTES: usize = 10 * 1024 * 1024;
const MAX_LOG_MESSAGE_CHARS: usize = 4_000;
const MAX_LOG_FILE_BYTES: u64 = 1_000_000;

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

/// Checks whether a URL targets an Apps Script execution endpoint so the app can parse it as JSON.
fn is_apps_script_url(url_str: &str) -> bool {
    Url::parse(url_str)
        .ok()
        .and_then(|url| {
            let is_script_host = url.host_str() == Some("script.google.com");
            let path = url.path();
            Some(
                is_script_host
                    && (path.starts_with("/a/macros/") || path.starts_with("/macros/s/"))
                    && path.ends_with("/exec"),
            )
        })
        .unwrap_or(false)
}

/// Checks for Workspace-domain Apps Script URLs that must be fetched by the WebView session.
fn is_domain_script_url(url_str: &str) -> bool {
    Url::parse(url_str)
        .ok()
        .map(|url| {
            url.host_str() == Some("script.google.com") && url.path().starts_with("/a/macros/")
        })
        .unwrap_or(false)
}

/// Extracts the spreadsheet ID from a Google Sheets URL so it can be converted to CSV export.
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

/// Converts supported Google Sheets URLs to CSV export URLs while preserving only the sheet ID and gid.
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
            gid_str.chars().take_while(|c| c.is_ascii_digit()).collect()
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

/// Validates user-provided spreadsheet URLs before any browser or Rust-side network access.
fn validate_allowed_source_url(url_str: &str) -> Result<Url, String> {
    let parsed = Url::parse(url_str).map_err(|_| "Invalid URL format.".to_string())?;
    if parsed.scheme() != "https" {
        return Err("Invalid URL. Only HTTPS links are allowed.".to_string());
    }

    let host = parsed.host_str().unwrap_or("");
    let path = parsed.path();
    let is_supported_sheet = host == "docs.google.com" && path.starts_with("/spreadsheets/d/");
    let is_supported_script = host == "script.google.com"
        && (path.starts_with("/macros/s/") || path.starts_with("/a/macros/"))
        && path.ends_with("/exec");

    if !is_supported_sheet && !is_supported_script {
        return Err(
            "Invalid URL. Use a Google Sheets link or a Google Apps Script /exec link.".to_string(),
        );
    }

    Ok(parsed)
}

/// Validates final Rust-side fetch targets, including Google-owned Apps Script redirect hosts.
fn validate_fetch_target_url(url_str: &str) -> Result<Url, String> {
    let parsed = Url::parse(url_str).map_err(|_| "Invalid fetch URL format.".to_string())?;
    if parsed.scheme() != "https" {
        return Err("Invalid fetch URL. Only HTTPS is allowed.".to_string());
    }

    let host = parsed.host_str().unwrap_or("");
    let is_google_fetch_host = host == "docs.google.com"
        || host == "script.google.com"
        || host == "script.googleusercontent.com";

    if !is_google_fetch_host {
        return Err("Invalid fetch URL. Host is not an approved Google data endpoint.".to_string());
    }

    Ok(parsed)
}

/// Downloads a bounded response body to avoid unbounded memory use from large or malformed replies.
async fn read_limited_body(mut response: reqwest::Response) -> Result<Vec<u8>, String> {
    if let Some(content_length) = response.content_length() {
        if content_length > MAX_REMOTE_RESPONSE_BYTES as u64 {
            return Err("Google response is too large to process safely.".to_string());
        }
    }

    let mut body = Vec::new();
    while let Some(chunk) = response
        .chunk()
        .await
        .map_err(|e| format!("Failed to read response body: {}", e))?
    {
        if body.len() + chunk.len() > MAX_REMOTE_RESPONSE_BYTES {
            return Err("Google response is too large to process safely.".to_string());
        }
        body.extend_from_slice(&chunk);
    }

    Ok(body)
}

/// Removes control characters and caps log size so user-controlled text cannot poison log files.
fn sanitize_log_value(value: &str, max_chars: usize) -> String {
    value
        .chars()
        .filter(|ch| !ch.is_control() || *ch == '\t')
        .take(max_chars)
        .collect()
}

// ─── Tauri Commands ───────────────────────────────────────────────────────────

/// Validate a URL against the whitelist (mirrors GET /api/validate)
/// Used by the frontend for domain-restricted /a/macros/ URLs that
/// the browser fetches directly with the user's Google session.
#[tauri::command]
async fn validate_url(url: String) -> Result<ApiResponse, String> {
    validate_allowed_source_url(&url).map_err(|e| e)?;

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
    let sheet_url = url.trim().to_string();
    if sheet_url.is_empty() {
        return Err("Spreadsheet URL is required.".to_string());
    }

    validate_allowed_source_url(&sheet_url).map_err(|e| e)?;

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
    validate_fetch_target_url(&target_url).map_err(|e| e)?;

    // Fetch from Google
    let client = reqwest::Client::builder()
        .danger_accept_invalid_certs(false)
        .timeout(Duration::from_secs(20))
        .redirect(reqwest::redirect::Policy::custom(|attempt| {
            if attempt.previous().len() > 5 {
                return attempt.error("too many redirects");
            }
            match validate_fetch_target_url(attempt.url().as_str()) {
                Ok(_) => attempt.follow(),
                Err(_) => attempt.error("redirected to an unapproved host"),
            }
        }))
        .build()
        .map_err(|e| format!("Failed to create HTTP client: {}", e))?;

    let response = client
        .get(&target_url)
        .header("User-Agent", "SAPID-license-analyser/1.0")
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

    let body = read_limited_body(response).await?;

    if is_apps_script {
        // Parse as JSON
        let json_data: serde_json::Value = serde_json::from_slice(&body)
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
        let csv_text = String::from_utf8(body)
            .map_err(|_| "Failed to read CSV response as UTF-8 text.".to_string())?;

        Ok(ApiResponse {
            success: true,
            message: None,
            data: None,
            csv: Some(csv_text),
            is_json: Some(false),
        })
    }
}

#[tauri::command]
async fn append_log(app: tauri::AppHandle, level: String, message: String) -> Result<(), String> {
    let log_dir = app.path().app_log_dir().map_err(|e| e.to_string())?;
    create_dir_all(&log_dir).map_err(|e| format!("Failed to create log directory: {}", e))?;

    let log_path = log_dir.join("error.log");
    if log_path.exists() {
        let metadata = std::fs::metadata(&log_path)
            .map_err(|e| format!("Failed to inspect log file: {}", e))?;
        if metadata.len() > MAX_LOG_FILE_BYTES {
            std::fs::remove_file(&log_path)
                .map_err(|e| format!("Failed to rotate oversized log file: {}", e))?;
        }
    }

    let mut file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
        .map_err(|e| format!("Failed to open log file: {}", e))?;

    let now = chrono::Local::now();
    let timestamp = now.format("%Y-%m-%d %H:%M:%S").to_string();
    let safe_level = sanitize_log_value(&level, 20).to_uppercase();
    let safe_message = sanitize_log_value(&message, MAX_LOG_MESSAGE_CHARS);
    let log_line = format!("[{}] [{}] {}\n", timestamp, safe_level, safe_message);

    file.write_all(log_line.as_bytes())
        .map_err(|e| format!("Failed to write to log file: {}", e))?;

    Ok(())
}

#[tauri::command]
async fn read_logs(app: tauri::AppHandle) -> Result<String, String> {
    let log_dir = app.path().app_log_dir().map_err(|e| e.to_string())?;
    let log_path = log_dir.join("error.log");
    if !log_path.exists() {
        return Ok("No logs found.".to_string());
    }

    let mut file = OpenOptions::new()
        .read(true)
        .open(&log_path)
        .map_err(|e| format!("Failed to open log file: {}", e))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| format!("Failed to read log file: {}", e))?;

    let lines: Vec<&str> = content.lines().collect();
    let last_lines = if lines.len() > 50 {
        lines[lines.len() - 50..].join("\n")
    } else {
        lines.join("\n")
    };

    Ok(last_lines)
}

#[tauri::command]
async fn clear_logs(app: tauri::AppHandle) -> Result<(), String> {
    let log_dir = app.path().app_log_dir().map_err(|e| e.to_string())?;
    let log_path = log_dir.join("error.log");
    if log_path.exists() {
        std::fs::remove_file(log_path).map_err(|e| format!("Failed to clear log file: {}", e))?;
    }
    Ok(())
}

#[tauri::command]
async fn open_log_directory(app: tauri::AppHandle) -> Result<(), String> {
    let log_dir = app.path().app_log_dir().map_err(|e| e.to_string())?;
    create_dir_all(&log_dir).map_err(|e| format!("Failed to create log directory: {}", e))?;

    #[cfg(target_os = "windows")]
    {
        std::process::Command::new("explorer")
            .arg(&log_dir)
            .spawn()
            .map_err(|e| format!("Failed to open explorer: {}", e))?;
    }
    #[cfg(not(target_os = "windows"))]
    {
        let log_dir_str = log_dir.to_string_lossy().to_string();
        tauri_plugin_opener::open_path(log_dir_str, None)
            .map_err(|e| format!("Failed to open log directory: {}", e))?;
    }

    Ok(())
}

// ─── App Entry ────────────────────────────────────────────────────────────────
#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            fetch_sheet_data,
            validate_url,
            append_log,
            read_logs,
            clear_logs,
            open_log_directory
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
