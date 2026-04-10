use base64::engine::general_purpose::{STANDARD, URL_SAFE_NO_PAD};
use base64::Engine;
use reqwest::redirect::Policy;
use reqwest::Client;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, State};

const APP_DIR_NAME: &str = "plutus-investment-dashboard";
const GRAPH_SCOPE: &str = "https://graph.microsoft.com/.default";
const GRAPH_SESSION_KEY: &str = "graph_session_v1";
const SHARED_TASKS_KEY: &str = "sharedrive-tasks";
const TEAM_STORE_PATH_FILE: &str = "team-store-path.json";
const UPLOAD_CHUNK_SIZE: usize = 5 * 1024 * 1024;
const DEFAULT_DELEGATED_SCOPES: &[&str] = &[
    "offline_access",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read",
    "Mail.Read",
];
const BUNDLED_CONFIG_JSON: &str = include_str!("../../app/data/config.json");
const BUNDLED_SHAREDRIVE_JSON: &str = include_str!("../../app/data/sharedrive-tasks.json");

#[derive(Debug, Serialize)]
struct StoreResponse<T> {
    ok: bool,
    data: Option<T>,
    error: Option<String>,
}

#[derive(Debug, Default, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct GraphSession {
    #[serde(default)]
    access_token: String,
    #[serde(default)]
    expires_at: i64,
    #[serde(default)]
    refresh_token: String,
}

#[derive(Debug, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct ShareDrivePayload {
    #[serde(default)]
    share_url: String,
    #[serde(default)]
    access_token: String,
    #[serde(default)]
    parent_item_id: String,
    #[serde(default)]
    drive_id: String,
    #[serde(default)]
    item_id: String,
    #[serde(default)]
    file_name: String,
    #[serde(default)]
    content_base64: String,
    #[serde(default)]
    conflict_behavior: String,
}

#[derive(Debug, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct DeviceCodePollPayload {
    #[serde(default)]
    device_code: String,
}

#[derive(Debug, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct OutlookMessagesPayload {
    #[serde(default)]
    access_token: String,
    top: Option<u32>,
    #[serde(default)]
    search: String,
}

struct AppState {
    client: Client,
    graph_session: Mutex<GraphSession>,
}

fn ok_response<T>(data: T) -> StoreResponse<T> {
    StoreResponse {
        ok: true,
        data: Some(data),
        error: None,
    }
}

fn err_response<T>(message: impl Into<String>) -> StoreResponse<T> {
    StoreResponse {
        ok: false,
        data: None,
        error: Some(message.into()),
    }
}

fn stringify_error(error: impl std::fmt::Display) -> String {
    error.to_string()
}

fn auth_log(stage: &str, details: impl AsRef<str>) {
    eprintln!("[plutus-auth] {stage} {}", details.as_ref());
}

fn now_millis() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis() as i64)
        .unwrap_or(0)
}

fn sanitize_key(key: &str) -> String {
    key.chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '_' || c == '-' {
                c
            } else {
                '_'
            }
        })
        .collect()
}

fn app_data_base(app: &AppHandle) -> Result<PathBuf, String> {
    app.path_resolver()
        .app_data_dir()
        .or_else(|| app.path_resolver().app_local_data_dir())
        .ok_or_else(|| "Unable to determine application data directory".to_string())
}

fn runtime_store_dir(app: &AppHandle) -> Result<PathBuf, String> {
    if let Some(custom_dir) = get_custom_store_dir(app)? {
        fs::create_dir_all(&custom_dir).map_err(stringify_error)?;
        return Ok(custom_dir);
    }

    let mut dir = app_data_base(app)?;
    dir.push(APP_DIR_NAME);
    dir.push("runtime-store");
    fs::create_dir_all(&dir).map_err(stringify_error)?;
    Ok(dir)
}

fn editable_data_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let mut dir = app_data_base(app)?;
    dir.push(APP_DIR_NAME);
    dir.push("data");
    fs::create_dir_all(&dir).map_err(stringify_error)?;
    Ok(dir)
}

fn get_store_file(app: &AppHandle, key: &str) -> Result<PathBuf, String> {
    let mut dir = runtime_store_dir(app)?;
    dir.push(format!("{}.json", sanitize_key(key)));
    Ok(dir)
}

fn get_editable_data_file(app: &AppHandle, key: &str) -> Result<PathBuf, String> {
    let mut dir = editable_data_dir(app)?;
    dir.push(format!("{}.json", sanitize_key(key)));
    Ok(dir)
}

fn project_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from(env!("CARGO_MANIFEST_DIR")))
}

fn bundled_data_candidates(app: &AppHandle, file_name: &str) -> Vec<PathBuf> {
    let mut candidates = Vec::new();

    if let Some(resource_dir) = app.path_resolver().resource_dir() {
        candidates.push(resource_dir.join("data").join(file_name));
    }

    let root = project_root();
    candidates.push(root.join("build").join("web").join("data").join(file_name));
    candidates.push(root.join("app").join("data").join(file_name));

    let mut seen = HashSet::new();
    candidates
        .into_iter()
        .filter(|candidate| seen.insert(candidate.clone()))
        .collect()
}

fn bundled_seed_value(key: &str) -> Option<Value> {
    match sanitize_key(key).as_str() {
        "config" => serde_json::from_str(BUNDLED_CONFIG_JSON).ok(),
        "sharedrive-tasks" => serde_json::from_str(BUNDLED_SHAREDRIVE_JSON).ok(),
        _ => None,
    }
}

fn read_json_file(path: &Path) -> Result<Value, String> {
    let raw = fs::read_to_string(path).map_err(stringify_error)?;
    serde_json::from_str(&raw).map_err(stringify_error)
}

fn read_data_json_impl(app: &AppHandle, key: &str) -> Result<Option<Value>, String> {
    let editable_file = get_editable_data_file(app, key)?;
    if editable_file.exists() {
        return read_json_file(&editable_file).map(Some);
    }

    let file_name = format!("{}.json", sanitize_key(key));
    for candidate in bundled_data_candidates(app, &file_name) {
        if candidate.exists() {
            return read_json_file(&candidate).map(Some);
        }
    }

    Ok(bundled_seed_value(key))
}

fn write_data_json_impl(app: &AppHandle, key: &str, value: &Value) -> Result<(), String> {
    let file = get_editable_data_file(app, key)?;
    let payload = serde_json::to_string_pretty(value).map_err(stringify_error)?;
    fs::write(file, payload).map_err(stringify_error)
}

fn read_store_json_impl(app: &AppHandle, key: &str) -> Result<Option<Value>, String> {
    let file = get_store_file(app, key)?;
    if !file.exists() {
        return Ok(None);
    }
    read_json_file(&file).map(Some)
}

fn write_store_json_impl(app: &AppHandle, key: &str, value: &Value) -> Result<(), String> {
    let file = get_store_file(app, key)?;
    let payload = serde_json::to_string_pretty(value).map_err(stringify_error)?;
    fs::write(file, payload).map_err(stringify_error)
}

fn read_team_store_config(app: &AppHandle) -> Result<Option<Value>, String> {
    for candidate in bundled_data_candidates(app, TEAM_STORE_PATH_FILE) {
        if candidate.exists() {
            return read_json_file(&candidate).map(Some);
        }
    }
    Ok(None)
}

fn get_custom_store_dir(app: &AppHandle) -> Result<Option<PathBuf>, String> {
    let Some(payload) = read_team_store_config(app)? else {
        return Ok(None);
    };

    let custom_dir = payload
        .get("storeDir")
        .and_then(Value::as_str)
        .map(str::trim)
        .unwrap_or_default();

    if custom_dir.is_empty() {
        return Ok(None);
    }

    Ok(Some(PathBuf::from(custom_dir)))
}

fn sharedrive_config_value(app: &AppHandle) -> Result<Value, String> {
    Ok(read_data_json_impl(app, SHARED_TASKS_KEY)?.unwrap_or_else(|| json!({})))
}

fn sharedrive_config_string(config: &Value, key: &str) -> String {
    config
        .get(key)
        .and_then(Value::as_str)
        .or_else(|| {
            config
                .get("tasks")
                .and_then(Value::as_object)
                .and_then(|tasks| tasks.get(key))
                .and_then(Value::as_str)
        })
        .map(str::trim)
        .unwrap_or_default()
        .to_string()
}

fn sharedrive_config_scopes(app: &AppHandle) -> Result<Vec<String>, String> {
    if let Ok(env_scopes) = std::env::var("PLUTUS_GRAPH_SCOPES") {
        let scopes: Vec<String> = env_scopes
            .split(|c: char| c.is_whitespace() || c == ',')
            .map(str::trim)
            .filter(|scope| !scope.is_empty())
            .map(ToOwned::to_owned)
            .collect();
        if !scopes.is_empty() {
            return Ok(scopes);
        }
    }

    let config = sharedrive_config_value(app)?;
    let configured = sharedrive_config_string(&config, "graphScopes");
    if !configured.is_empty() {
        let scopes: Vec<String> = configured
            .split(|c: char| c.is_whitespace() || c == ',')
            .map(str::trim)
            .filter(|scope| !scope.is_empty())
            .map(ToOwned::to_owned)
            .collect();
        if !scopes.is_empty() {
            return Ok(scopes);
        }
    }

    Ok(DEFAULT_DELEGATED_SCOPES
        .iter()
        .map(|scope| (*scope).to_string())
        .collect())
}

fn graph_tenant_id(app: &AppHandle) -> Result<String, String> {
    if let Ok(tenant_id) = std::env::var("PLUTUS_AZURE_TENANT_ID") {
        let tenant_id = tenant_id.trim().to_string();
        if !tenant_id.is_empty() {
            return Ok(tenant_id);
        }
    }

    let config = sharedrive_config_value(app)?;
    let tenant_id = sharedrive_config_string(&config, "azureTenantId");
    if !tenant_id.is_empty() {
        return Ok(tenant_id);
    }

    Ok("common".to_string())
}

fn graph_client_id(app: &AppHandle) -> Result<String, String> {
    if let Ok(client_id) = std::env::var("PLUTUS_AZURE_CLIENT_ID") {
        let client_id = client_id.trim().to_string();
        if !client_id.is_empty() {
            return Ok(client_id);
        }
    }

    let config = sharedrive_config_value(app)?;
    Ok(sharedrive_config_string(&config, "azureClientId"))
}

fn graph_client_secret() -> String {
    std::env::var("PLUTUS_AZURE_CLIENT_SECRET")
        .unwrap_or_default()
        .trim()
        .to_string()
}

fn parse_json_text(text: &str) -> Option<Value> {
    serde_json::from_str(text).ok()
}

fn payload_error_message(payload: &Value) -> String {
    payload
        .get("error")
        .and_then(|error| {
            error
                .get("message")
                .or_else(|| error.get("error_description"))
        })
        .and_then(Value::as_str)
        .or_else(|| payload.get("error_description").and_then(Value::as_str))
        .or_else(|| payload.get("error").and_then(Value::as_str))
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| "Request failed".to_string())
}

async fn fetch_json(request: reqwest::RequestBuilder) -> Result<Value, String> {
    let response = request.send().await.map_err(stringify_error)?;
    let status = response.status();
    let text = response.text().await.map_err(stringify_error)?;
    let payload = parse_json_text(&text).unwrap_or(Value::Null);

    if !status.is_success() {
        let message = if payload.is_null() {
            format!("HTTP {}: {}", status.as_u16(), text.trim())
        } else {
            format!(
                "HTTP {}: {}",
                status.as_u16(),
                payload_error_message(&payload)
            )
        };
        return Err(message);
    }

    Ok(payload)
}

async fn fetch_json_with_status(
    request: reqwest::RequestBuilder,
) -> Result<(bool, u16, Value), String> {
    let response = request.send().await.map_err(stringify_error)?;
    let status = response.status();
    let text = response.text().await.map_err(stringify_error)?;
    let payload = parse_json_text(&text).unwrap_or(Value::Null);
    Ok((status.is_success(), status.as_u16(), payload))
}

fn decode_jwt_payload(token: &str) -> Option<Value> {
    let encoded = token.split('.').nth(1)?.trim();
    if encoded.is_empty() {
        return None;
    }

    let decoded = URL_SAFE_NO_PAD.decode(encoded.as_bytes()).ok()?;
    serde_json::from_slice(&decoded).ok()
}

fn token_has_scope(token: &str, required_scope: &str) -> bool {
    let Some(payload) = decode_jwt_payload(token) else {
        return false;
    };

    let scopes = payload
        .get("scp")
        .and_then(Value::as_str)
        .unwrap_or_default();

    scopes
        .split_whitespace()
        .any(|scope| scope.eq_ignore_ascii_case(required_scope))
}

fn read_graph_session(app: &AppHandle, state: &AppState) -> Result<GraphSession, String> {
    let mut session = state.graph_session.lock().map_err(stringify_error)?;

    if session.access_token.is_empty() && session.refresh_token.is_empty() {
        if let Some(value) = read_store_json_impl(app, GRAPH_SESSION_KEY)? {
            if let Ok(parsed) = serde_json::from_value::<GraphSession>(value) {
                *session = parsed;
            }
        }
    }

    Ok(session.clone())
}

fn write_graph_session(
    app: &AppHandle,
    state: &AppState,
    graph_session: &GraphSession,
) -> Result<(), String> {
    {
        let mut session = state.graph_session.lock().map_err(stringify_error)?;
        *session = graph_session.clone();
    }

    let value = serde_json::to_value(graph_session).map_err(stringify_error)?;
    write_store_json_impl(app, GRAPH_SESSION_KEY, &value)
}

fn graph_session_summary(app: &AppHandle, state: &AppState) -> Result<Value, String> {
    let session = read_graph_session(app, state)?;
    Ok(json!({
        "accessToken": session.access_token,
        "expiresAt": session.expires_at,
        "hasRefreshToken": !session.refresh_token.is_empty(),
    }))
}

fn update_graph_tokens(
    app: &AppHandle,
    state: &AppState,
    access_token: String,
    expires_in_seconds: i64,
    refresh_token: Option<String>,
) -> Result<GraphSession, String> {
    let mut session = read_graph_session(app, state)?;
    session.access_token = access_token.trim().to_string();
    session.expires_at = now_millis() + expires_in_seconds.max(0) * 1000;

    if let Some(refresh_token) = refresh_token {
        session.refresh_token = refresh_token.trim().to_string();
    }

    write_graph_session(app, state, &session)?;
    Ok(session)
}

async fn refresh_access_token(
    app: &AppHandle,
    state: &AppState,
    client: &Client,
    refresh_token: &str,
) -> Result<String, String> {
    let client_id = graph_client_id(app)?;
    if client_id.trim().is_empty() {
        return Err("Missing PLUTUS_AZURE_CLIENT_ID for refresh token flow.".to_string());
    }

    let tenant_id = graph_tenant_id(app)?;
    let scopes = sharedrive_config_scopes(app)?.join(" ");
    let token_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
        urlencoding::encode(&tenant_id)
    );

    let body = [
        ("client_id", client_id),
        ("grant_type", "refresh_token".to_string()),
        ("refresh_token", refresh_token.to_string()),
        ("scope", scopes),
    ];

    let request = client
        .post(token_url)
        .header("Content-Type", "application/x-www-form-urlencoded")
        .form(&body);

    let (ok, _status, payload) = fetch_json_with_status(request).await?;
    if !ok {
        let message = payload
            .get("error_description")
            .and_then(Value::as_str)
            .or_else(|| payload.get("error").and_then(Value::as_str))
            .unwrap_or("Refresh token flow failed.");
        return Err(message.to_string());
    }

    let access_token = payload
        .get("access_token")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    let expires_in = payload
        .get("expires_in")
        .and_then(Value::as_i64)
        .unwrap_or(0);
    let new_refresh_token = payload
        .get("refresh_token")
        .and_then(Value::as_str)
        .map(ToOwned::to_owned);

    update_graph_tokens(
        app,
        state,
        access_token.clone(),
        expires_in,
        new_refresh_token,
    )?;
    Ok(access_token)
}

async fn client_credentials_token(_app: &AppHandle, client: &Client) -> Result<String, String> {
    let tenant_id = std::env::var("PLUTUS_AZURE_TENANT_ID")
        .unwrap_or_default()
        .trim()
        .to_string();
    let client_id = std::env::var("PLUTUS_AZURE_CLIENT_ID")
        .unwrap_or_default()
        .trim()
        .to_string();
    let client_secret = graph_client_secret();

    if tenant_id.is_empty() || client_id.is_empty() || client_secret.is_empty() {
        return Ok(String::new());
    }

    let token_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
        urlencoding::encode(&tenant_id)
    );
    let body = [
        ("client_id", client_id),
        ("client_secret", client_secret),
        ("scope", GRAPH_SCOPE.to_string()),
        ("grant_type", "client_credentials".to_string()),
    ];

    let payload = fetch_json(
        client
            .post(token_url)
            .header("Content-Type", "application/x-www-form-urlencoded")
            .form(&body),
    )
    .await?;

    Ok(payload
        .get("access_token")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string())
}

async fn resolve_graph_access_token(
    app: &AppHandle,
    state: &AppState,
    override_token: &str,
) -> Result<String, String> {
    let explicit = override_token.trim();
    if !explicit.is_empty() {
        return Ok(explicit.to_string());
    }

    let session = read_graph_session(app, state)?;
    if !session.access_token.is_empty() && now_millis() < session.expires_at - 60_000 {
        return Ok(session.access_token);
    }

    if let Ok(env_token) = std::env::var("PLUTUS_GRAPH_ACCESS_TOKEN") {
        let env_token = env_token.trim().to_string();
        if !env_token.is_empty() {
            return Ok(env_token);
        }
    }

    if !session.refresh_token.is_empty() {
        match refresh_access_token(app, state, &state.client, &session.refresh_token).await {
            Ok(token) => return Ok(token),
            Err(_) => {
                let mut cleared = session.clone();
                cleared.refresh_token.clear();
                write_graph_session(app, state, &cleared)?;
            }
        }
    }

    let client_token = client_credentials_token(app, &state.client).await?;
    if !client_token.is_empty() {
        return Ok(client_token);
    }

    Err("No Microsoft Graph access token available. Provide one in the page, or set PLUTUS_GRAPH_ACCESS_TOKEN, or set PLUTUS_AZURE_TENANT_ID / PLUTUS_AZURE_CLIENT_ID / PLUTUS_AZURE_CLIENT_SECRET.".to_string())
}

async fn request_device_code_impl(app: &AppHandle, client: &Client) -> Result<Value, String> {
    let client_id = graph_client_id(app)?;
    if client_id.trim().is_empty() {
        return Err("Missing PLUTUS_AZURE_CLIENT_ID for device code flow.".to_string());
    }

    let tenant_id = graph_tenant_id(app)?;
    let scopes = sharedrive_config_scopes(app)?.join(" ");
    let device_code_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/devicecode",
        urlencoding::encode(&tenant_id)
    );

    auth_log(
        "device-code:start",
        format!("tenant_id={tenant_id} client_id={client_id} scopes={scopes}"),
    );

    let payload = fetch_json(
        client
            .post(device_code_url)
            .header("Content-Type", "application/x-www-form-urlencoded")
            .form(&[("client_id", client_id), ("scope", scopes)]),
    )
    .await?;

    auth_log(
        "device-code:issued",
        format!(
            "interval={} expires_in={} verification_uri={}",
            payload.get("interval").and_then(Value::as_i64).unwrap_or(0),
            payload
                .get("expires_in")
                .and_then(Value::as_i64)
                .unwrap_or(0),
            payload
                .get("verification_uri")
                .and_then(Value::as_str)
                .unwrap_or_default()
        ),
    );

    Ok(payload)
}

async fn poll_device_code_impl(
    app: &AppHandle,
    state: &AppState,
    device_code: &str,
) -> Result<Value, String> {
    let client_id = graph_client_id(app)?;
    if client_id.trim().is_empty() {
        return Err("Missing PLUTUS_AZURE_CLIENT_ID for device code flow.".to_string());
    }

    let tenant_id = graph_tenant_id(app)?;
    let token_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
        urlencoding::encode(&tenant_id)
    );

    let request = state
        .client
        .post(token_url)
        .header("Content-Type", "application/x-www-form-urlencoded")
        .form(&[
            ("client_id", client_id),
            (
                "grant_type",
                "urn:ietf:params:oauth:grant-type:device_code".to_string(),
            ),
            ("device_code", device_code.trim().to_string()),
        ]);

    let (ok, _status, payload) = fetch_json_with_status(request).await?;
    if !ok {
        auth_log(
            "device-code:poll-pending",
            format!(
                "error={} description={}",
                payload
                    .get("error")
                    .and_then(Value::as_str)
                    .unwrap_or("authorization_pending"),
                payload
                    .get("error_description")
                    .and_then(Value::as_str)
                    .unwrap_or("Authorization pending.")
            ),
        );
        return Ok(json!({
            "ok": false,
            "error": payload.get("error").and_then(Value::as_str).unwrap_or("authorization_pending"),
            "error_description": payload.get("error_description").and_then(Value::as_str).unwrap_or("Authorization pending."),
            "interval": payload.get("interval").cloned().unwrap_or_else(|| json!(5)),
        }));
    }

    let access_token = payload
        .get("access_token")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    let expires_in = payload
        .get("expires_in")
        .and_then(Value::as_i64)
        .unwrap_or(0);
    let refresh_token = payload
        .get("refresh_token")
        .and_then(Value::as_str)
        .map(ToOwned::to_owned);

    update_graph_tokens(
        app,
        state,
        access_token.clone(),
        expires_in,
        refresh_token.clone(),
    )?;

    auth_log(
        "device-code:poll-success",
        format!(
            "expires_in={} refresh_token={} scope={}",
            expires_in,
            refresh_token
                .as_deref()
                .map(|value| !value.is_empty())
                .unwrap_or(false),
            payload
                .get("scope")
                .and_then(Value::as_str)
                .unwrap_or_default()
        ),
    );

    Ok(json!({
        "ok": true,
        "accessToken": access_token,
        "expiresIn": expires_in,
        "refreshToken": refresh_token.unwrap_or_default(),
        "scope": payload.get("scope").and_then(Value::as_str).unwrap_or_default(),
        "tokenType": payload.get("token_type").and_then(Value::as_str).unwrap_or_default(),
    }))
}

fn to_base64_url(value: &str) -> String {
    URL_SAFE_NO_PAD.encode(value.as_bytes())
}

fn normalize_drive_item(item: &Value) -> Value {
    json!({
        "id": item.get("id").and_then(Value::as_str).unwrap_or_default(),
        "name": item.get("name").and_then(Value::as_str).unwrap_or_default(),
        "webUrl": item.get("webUrl").and_then(Value::as_str).unwrap_or_default(),
        "size": item.get("size").and_then(Value::as_i64),
        "lastModifiedDateTime": item.get("lastModifiedDateTime").and_then(Value::as_str).unwrap_or_default(),
        "isFolder": item.get("folder").is_some(),
        "isFile": item.get("file").is_some(),
        "mimeType": item
            .get("file")
            .and_then(|file| file.get("mimeType"))
            .and_then(Value::as_str)
            .unwrap_or_default(),
        "childCount": item
            .get("folder")
            .and_then(|folder| folder.get("childCount"))
            .and_then(Value::as_i64),
        "parentPath": item
            .get("parentReference")
            .and_then(|reference| reference.get("path"))
            .and_then(Value::as_str)
            .unwrap_or_default(),
        "downloadUrl": item.get("@microsoft.graph.downloadUrl").and_then(Value::as_str).unwrap_or_default(),
    })
}

fn graph_drive_id(item: &Value) -> String {
    item.get("parentReference")
        .and_then(|reference| reference.get("driveId"))
        .and_then(Value::as_str)
        .or_else(|| {
            item.get("remoteItem")
                .and_then(|remote_item| remote_item.get("parentReference"))
                .and_then(|reference| reference.get("driveId"))
                .and_then(Value::as_str)
        })
        .unwrap_or_default()
        .trim()
        .to_string()
}

fn graph_item_id(item: &Value) -> String {
    item.get("id")
        .and_then(Value::as_str)
        .or_else(|| {
            item.get("remoteItem")
                .and_then(|remote_item| remote_item.get("id"))
                .and_then(Value::as_str)
        })
        .unwrap_or_default()
        .trim()
        .to_string()
}

fn graph_parent_item_id(item: &Value) -> String {
    item.get("parentReference")
        .and_then(|reference| reference.get("id"))
        .and_then(Value::as_str)
        .or_else(|| {
            item.get("remoteItem")
                .and_then(|remote_item| remote_item.get("parentReference"))
                .and_then(|reference| reference.get("id"))
                .and_then(Value::as_str)
        })
        .unwrap_or_default()
        .trim()
        .to_string()
}

fn default_drive_item_fields() -> Vec<&'static str> {
    vec![
        "id",
        "name",
        "webUrl",
        "parentReference",
        "folder",
        "file",
        "size",
        "lastModifiedDateTime",
        "remoteItem",
        "@microsoft.graph.downloadUrl",
    ]
}

async fn get_drive_item(
    client: &Client,
    drive_id: &str,
    item_id: &str,
    token: &str,
    select_fields: &[&str],
) -> Result<Value, String> {
    let fields = if select_fields.is_empty() {
        default_drive_item_fields()
    } else {
        select_fields.to_vec()
    };

    let url = format!(
        "https://graph.microsoft.com/v1.0/drives/{}/items/{}?$select={}",
        urlencoding::encode(drive_id),
        urlencoding::encode(item_id),
        fields.join(",")
    );

    fetch_json(
        client
            .get(url)
            .header("Authorization", format!("Bearer {token}"))
            .header("Accept", "application/json"),
    )
    .await
}

async fn get_share_drive_item(
    client: &Client,
    share_url: &str,
    token: &str,
    select_fields: &[&str],
) -> Result<Value, String> {
    let clean_share_url = share_url.trim();
    if clean_share_url.is_empty() {
        return Err("Share URL is required.".to_string());
    }

    let fields = if select_fields.is_empty() {
        default_drive_item_fields()
    } else {
        select_fields.to_vec()
    };

    let encoded_share = format!("u!{}", to_base64_url(clean_share_url));
    let url = format!(
        "https://graph.microsoft.com/v1.0/shares/{encoded_share}/driveItem?$select={}",
        fields.join(",")
    );

    fetch_json(
        client
            .get(url)
            .header("Authorization", format!("Bearer {token}"))
            .header("Accept", "application/json"),
    )
    .await
}

async fn list_drive_item_children(
    client: &Client,
    drive_id: &str,
    item_id: &str,
    token: &str,
) -> Result<Vec<Value>, String> {
    let select = [
        "id",
        "name",
        "webUrl",
        "parentReference",
        "folder",
        "file",
        "size",
        "lastModifiedDateTime",
        "@microsoft.graph.downloadUrl",
    ]
    .join(",");

    let mut next_url = format!(
        "https://graph.microsoft.com/v1.0/drives/{}/items/{}/children?$top=200&$select={select}",
        urlencoding::encode(drive_id),
        urlencoding::encode(item_id),
    );
    let mut items = Vec::new();

    while !next_url.is_empty() {
        let page = fetch_json(
            client
                .get(&next_url)
                .header("Authorization", format!("Bearer {token}"))
                .header("Accept", "application/json"),
        )
        .await?;

        if let Some(values) = page.get("value").and_then(Value::as_array) {
            items.extend(values.iter().cloned());
        }

        next_url = page
            .get("@odata.nextLink")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .to_string();
    }

    Ok(items)
}

async fn list_share_drive_children_impl(
    app: &AppHandle,
    state: &AppState,
    payload: &ShareDrivePayload,
) -> Result<Value, String> {
    let token = resolve_graph_access_token(app, state, &payload.access_token).await?;
    let shared_item = get_share_drive_item(&state.client, &payload.share_url, &token, &[]).await?;
    let drive_id = graph_drive_id(&shared_item);
    if drive_id.is_empty() {
        return Err("Unable to resolve drive for SharePoint item.".to_string());
    }

    let root_item = if shared_item.get("folder").is_some() {
        shared_item.clone()
    } else {
        let parent_id = graph_parent_item_id(&shared_item);
        if parent_id.is_empty() {
            return Err("Unable to resolve parent folder for shared file.".to_string());
        }
        get_drive_item(&state.client, &drive_id, &parent_id, &token, &[]).await?
    };

    let parent_item_id = if payload.parent_item_id.trim().is_empty() {
        graph_item_id(&root_item)
    } else {
        payload.parent_item_id.trim().to_string()
    };

    if parent_item_id.is_empty() {
        return Err("Unable to resolve folder id for SharePoint item.".to_string());
    }

    let items = list_drive_item_children(&state.client, &drive_id, &parent_item_id, &token).await?;
    let normalized_items: Vec<Value> = items.iter().map(normalize_drive_item).collect();

    Ok(json!({
        "root": normalize_drive_item(&root_item),
        "driveId": drive_id,
        "parentItemId": parent_item_id,
        "items": normalized_items,
        "fetchedAt": chrono::Utc::now().to_rfc3339(),
    }))
}

async fn get_share_drive_download_url_impl(
    app: &AppHandle,
    state: &AppState,
    payload: &ShareDrivePayload,
) -> Result<Value, String> {
    let token = resolve_graph_access_token(app, state, &payload.access_token).await?;

    let item = if !payload.drive_id.trim().is_empty() && !payload.item_id.trim().is_empty() {
        let url = format!(
            "https://graph.microsoft.com/v1.0/drives/{}/items/{}?$select=id,name,webUrl,parentReference,@microsoft.graph.downloadUrl",
            urlencoding::encode(payload.drive_id.trim()),
            urlencoding::encode(payload.item_id.trim()),
        );
        fetch_json(
            state
                .client
                .get(url)
                .header("Authorization", format!("Bearer {token}"))
                .header("Accept", "application/json"),
        )
        .await?
    } else {
        get_share_drive_item(
            &state.client,
            &payload.share_url,
            &token,
            &[
                "id",
                "name",
                "webUrl",
                "parentReference",
                "remoteItem",
                "@microsoft.graph.downloadUrl",
            ],
        )
        .await?
    };

    let mut drive_id = if payload.drive_id.trim().is_empty() {
        graph_drive_id(&item)
    } else {
        payload.drive_id.trim().to_string()
    };
    let item_id = if payload.item_id.trim().is_empty() {
        graph_item_id(&item)
    } else {
        payload.item_id.trim().to_string()
    };
    let parent_item_id = graph_parent_item_id(&item);
    let mut download_url = item
        .get("@microsoft.graph.downloadUrl")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();

    if download_url.is_empty() && !drive_id.is_empty() && !item_id.is_empty() {
        let retry_url = format!(
            "https://graph.microsoft.com/v1.0/drives/{}/items/{}?$select=id,name,@microsoft.graph.downloadUrl",
            urlencoding::encode(&drive_id),
            urlencoding::encode(&item_id),
        );

        if let Ok(retry_response) = fetch_json(
            state
                .client
                .get(retry_url)
                .header("Authorization", format!("Bearer {token}"))
                .header("Accept", "application/json"),
        )
        .await
        {
            if let Some(resolved) = retry_response
                .get("@microsoft.graph.downloadUrl")
                .and_then(Value::as_str)
            {
                download_url = resolved.trim().to_string();
                if drive_id.is_empty() {
                    drive_id = graph_drive_id(&retry_response);
                }
            }
        }
    }

    if download_url.is_empty() && !drive_id.is_empty() && !item_id.is_empty() {
        let redirect_client = Client::builder()
            .redirect(Policy::none())
            .build()
            .map_err(stringify_error)?;
        let content_url = format!(
            "https://graph.microsoft.com/v1.0/drives/{}/items/{}/content",
            urlencoding::encode(&drive_id),
            urlencoding::encode(&item_id),
        );
        let response = redirect_client
            .get(content_url)
            .header("Authorization", format!("Bearer {token}"))
            .send()
            .await
            .map_err(stringify_error)?;
        if let Some(location) = response
            .headers()
            .get("location")
            .and_then(|value| value.to_str().ok())
        {
            download_url = location.to_string();
        }
    }

    if download_url.is_empty() {
        return Err("Download URL not available for this item.".to_string());
    }

    Ok(json!({
        "id": item_id,
        "itemId": item_id,
        "name": item.get("name").and_then(Value::as_str).unwrap_or_default(),
        "webUrl": item.get("webUrl").and_then(Value::as_str).unwrap_or_default(),
        "driveId": drive_id,
        "parentItemId": parent_item_id,
        "downloadUrl": download_url,
    }))
}

async fn download_share_drive_file_impl(
    app: &AppHandle,
    state: &AppState,
    payload: &ShareDrivePayload,
) -> Result<Value, String> {
    let token = resolve_graph_access_token(app, state, &payload.access_token).await?;

    let (drive_id, item_id) =
        if !payload.drive_id.trim().is_empty() && !payload.item_id.trim().is_empty() {
            (
                payload.drive_id.trim().to_string(),
                payload.item_id.trim().to_string(),
            )
        } else {
            let item = get_share_drive_item(
                &state.client,
                &payload.share_url,
                &token,
                &[
                    "id",
                    "name",
                    "parentReference",
                    "remoteItem",
                    "@microsoft.graph.downloadUrl",
                ],
            )
            .await?;
            (graph_drive_id(&item), graph_item_id(&item))
        };

    if drive_id.is_empty() || item_id.is_empty() {
        return Err("Unable to resolve sharedrive file for download.".to_string());
    }

    let url = format!(
        "https://graph.microsoft.com/v1.0/drives/{}/items/{}/content",
        urlencoding::encode(&drive_id),
        urlencoding::encode(&item_id),
    );

    let response = state
        .client
        .get(url)
        .header("Authorization", format!("Bearer {token}"))
        .send()
        .await
        .map_err(stringify_error)?;
    let status = response.status();
    let text = response.text().await.map_err(stringify_error)?;

    if !status.is_success() {
        let message = parse_json_text(&text)
            .and_then(|payload| {
                payload
                    .get("error")
                    .and_then(|error| error.get("message"))
                    .and_then(Value::as_str)
                    .map(ToOwned::to_owned)
            })
            .unwrap_or_else(|| format!("Download failed: {}", status));
        return Err(message);
    }

    Ok(json!({
        "driveId": drive_id,
        "itemId": item_id,
        "text": text,
    }))
}

async fn create_upload_session(
    client: &Client,
    drive_id: &str,
    parent_item_id: &str,
    file_name: &str,
    token: &str,
    conflict_behavior: &str,
) -> Result<Value, String> {
    let url = format!(
        "https://graph.microsoft.com/v1.0/drives/{}/items/{}:/{}/createUploadSession",
        urlencoding::encode(drive_id),
        urlencoding::encode(parent_item_id),
        urlencoding::encode(file_name),
    );
    let payload = json!({
        "item": {
            "@microsoft.graph.conflictBehavior": if conflict_behavior.trim().is_empty() {
                "replace"
            } else {
                conflict_behavior.trim()
            }
        }
    });

    fetch_json(
        client
            .post(url)
            .header("Authorization", format!("Bearer {token}"))
            .header("Accept", "application/json")
            .header("Content-Type", "application/json")
            .body(payload.to_string()),
    )
    .await
}

async fn upload_direct_file(
    client: &Client,
    drive_id: &str,
    item_id: &str,
    parent_item_id: &str,
    file_name: &str,
    token: &str,
    conflict_behavior: &str,
    buffer: &[u8],
) -> Result<Value, String> {
    let mut url = if !item_id.trim().is_empty() {
        format!(
            "https://graph.microsoft.com/v1.0/drives/{}/items/{}/content",
            urlencoding::encode(drive_id),
            urlencoding::encode(item_id),
        )
    } else {
        format!(
            "https://graph.microsoft.com/v1.0/drives/{}/items/{}:/{}/content",
            urlencoding::encode(drive_id),
            urlencoding::encode(parent_item_id),
            urlencoding::encode(file_name),
        )
    };

    if item_id.trim().is_empty() && !conflict_behavior.trim().is_empty() {
        url.push_str("?@microsoft.graph.conflictBehavior=");
        url.push_str(&urlencoding::encode(conflict_behavior.trim()));
    }

    fetch_json(
        client
            .put(url)
            .header("Authorization", format!("Bearer {token}"))
            .header("Accept", "application/json")
            .header("Content-Type", "application/octet-stream")
            .body(buffer.to_vec()),
    )
    .await
}

async fn upload_with_session(
    client: &Client,
    upload_url: &str,
    buffer: &[u8],
) -> Result<Option<Value>, String> {
    let total = buffer.len();
    let mut start = 0;

    while start < total {
        let end = std::cmp::min(start + UPLOAD_CHUNK_SIZE, total);
        let chunk = buffer[start..end].to_vec();
        let response = client
            .put(upload_url)
            .header("Content-Type", "application/octet-stream")
            .header("Content-Length", chunk.len().to_string())
            .header(
                "Content-Range",
                format!("bytes {start}-{}/{total}", end - 1),
            )
            .body(chunk)
            .send()
            .await
            .map_err(stringify_error)?;
        let status = response.status();
        let text = response.text().await.map_err(stringify_error)?;
        let payload = parse_json_text(&text).unwrap_or(Value::Null);

        if !status.is_success() {
            let message = if payload.is_null() {
                format!("Upload failed: {}", text.trim())
            } else {
                format!("Upload failed: {}", payload_error_message(&payload))
            };
            return Err(message);
        }

        if end == total {
            return Ok(Some(payload));
        }

        start = end;
    }

    Ok(None)
}

async fn upload_share_drive_file_impl(
    app: &AppHandle,
    state: &AppState,
    payload: &ShareDrivePayload,
) -> Result<Value, String> {
    let share_url = payload.share_url.trim();
    let file_name = payload.file_name.trim();
    if share_url.is_empty() {
        return Err("Share URL is required.".to_string());
    }
    if file_name.is_empty() {
        return Err("File name is required.".to_string());
    }

    let content = STANDARD
        .decode(payload.content_base64.trim().as_bytes())
        .map_err(stringify_error)?;
    if content.is_empty() {
        return Err("Upload content is empty.".to_string());
    }

    let token = resolve_graph_access_token(app, state, &payload.access_token).await?;
    let root_item = get_share_drive_item(
        &state.client,
        share_url,
        &token,
        &[
            "id",
            "name",
            "parentReference",
            "remoteItem",
            "folder",
            "file",
        ],
    )
    .await?;
    let drive_id = graph_drive_id(&root_item);
    if drive_id.is_empty() {
        return Err("Unable to resolve SharePoint drive ID.".to_string());
    }

    let parent_item_id = if !payload.parent_item_id.trim().is_empty() {
        payload.parent_item_id.trim().to_string()
    } else if root_item.get("folder").is_some() {
        graph_item_id(&root_item)
    } else {
        graph_parent_item_id(&root_item)
    };

    if parent_item_id.is_empty() {
        return Err("Target folder ID is required for upload.".to_string());
    }

    let direct_item_id = if root_item.get("file").is_some()
        && root_item
            .get("name")
            .and_then(Value::as_str)
            .map(|name| name.eq_ignore_ascii_case(file_name))
            .unwrap_or(false)
    {
        graph_item_id(&root_item)
    } else {
        String::new()
    };

    let session = match create_upload_session(
        &state.client,
        &drive_id,
        &parent_item_id,
        file_name,
        &token,
        &payload.conflict_behavior,
    )
    .await
    {
        Ok(session) => session,
        Err(error) if error.trim_start().starts_with("HTTP 400:") => {
            let uploaded = upload_direct_file(
                &state.client,
                &drive_id,
                &direct_item_id,
                &parent_item_id,
                file_name,
                &token,
                &payload.conflict_behavior,
                &content,
            )
            .await?;

            return Ok(json!({
                "item": normalize_drive_item(&uploaded),
                "driveId": drive_id,
                "parentItemId": parent_item_id,
            }));
        }
        Err(error) => return Err(error),
    };

    let upload_url = session
        .get("uploadUrl")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    if upload_url.is_empty() {
        return Err("Failed to create upload session.".to_string());
    }

    let uploaded = upload_with_session(&state.client, &upload_url, &content).await?;
    Ok(json!({
        "item": uploaded.map(|value| normalize_drive_item(&value)).unwrap_or(Value::Null),
        "driveId": drive_id,
        "parentItemId": parent_item_id,
    }))
}

fn normalize_email_address(entry: Option<&Value>) -> Value {
    let email_address = entry
        .and_then(|value| value.get("emailAddress"))
        .cloned()
        .unwrap_or_else(|| json!({}));
    json!({
        "name": email_address.get("name").and_then(Value::as_str).unwrap_or_default(),
        "address": email_address.get("address").and_then(Value::as_str).unwrap_or_default(),
    })
}

fn normalize_recipients(recipients: Option<&Value>) -> Vec<Value> {
    recipients
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(|entry| normalize_email_address(Some(entry)))
        .filter(|entry| {
            entry
                .get("address")
                .and_then(Value::as_str)
                .map(|value| !value.is_empty())
                .unwrap_or(false)
        })
        .collect()
}

fn normalize_outlook_message(message: &Value) -> Value {
    json!({
        "id": message.get("id").and_then(Value::as_str).unwrap_or_default(),
        "subject": message.get("subject").and_then(Value::as_str).unwrap_or_default(),
        "bodyPreview": message.get("bodyPreview").and_then(Value::as_str).unwrap_or_default(),
        "receivedDateTime": message.get("receivedDateTime").and_then(Value::as_str).unwrap_or_default(),
        "webLink": message.get("webLink").and_then(Value::as_str).unwrap_or_default(),
        "isRead": message.get("isRead").and_then(Value::as_bool).unwrap_or(false),
        "from": normalize_email_address(message.get("from")),
        "toRecipients": normalize_recipients(message.get("toRecipients")),
        "ccRecipients": normalize_recipients(message.get("ccRecipients")),
    })
}

async fn list_outlook_messages_impl(
    app: &AppHandle,
    state: &AppState,
    payload: &OutlookMessagesPayload,
) -> Result<Value, String> {
    let mut token = resolve_graph_access_token(app, state, &payload.access_token).await?;
    let graph_session = read_graph_session(app, state)?;

    if !token_has_scope(&token, "Mail.Read")
        && payload.access_token.trim().is_empty()
        && !graph_session.refresh_token.is_empty()
    {
        if let Ok(refreshed) =
            refresh_access_token(app, state, &state.client, &graph_session.refresh_token).await
        {
            token = refreshed;
        }
    }

    if !token_has_scope(&token, "Mail.Read") {
        return Err("Microsoft sign-in is missing Mail.Read. Click 'Sign in with Microsoft' again and approve Outlook inbox access.".to_string());
    }

    let safe_top = payload.top.unwrap_or(50).clamp(1, 5000);
    let page_size = safe_top.min(250);
    let search_term = payload.search.trim();
    let mut query = vec![
        ("$top".to_string(), page_size.to_string()),
        (
            "$select".to_string(),
            "id,subject,bodyPreview,receivedDateTime,webLink,isRead,from,toRecipients,ccRecipients"
                .to_string(),
        ),
    ];

    if search_term.is_empty() {
        query.push(("$orderby".to_string(), "receivedDateTime DESC".to_string()));
    } else {
        query.push((
            "$search".to_string(),
            format!("\"{}\"", search_term.replace('"', "\\\"")),
        ));
    }

    let mut next_url = format!(
        "https://graph.microsoft.com/v1.0/me/messages?{}",
        serde_urlencoded::to_string(query).map_err(stringify_error)?
    );
    let mut items = Vec::new();
    let mut seen_ids = HashSet::new();

    while !next_url.is_empty() && items.len() < safe_top as usize {
        let mut request = state
            .client
            .get(&next_url)
            .header("Authorization", format!("Bearer {token}"))
            .header("Accept", "application/json");

        if !search_term.is_empty() {
            request = request.header("ConsistencyLevel", "eventual");
        }

        let payload = fetch_json(request).await?;
        if let Some(values) = payload.get("value").and_then(Value::as_array) {
            for message in values {
                let normalized = normalize_outlook_message(message);
                let id = normalized
                    .get("id")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .trim()
                    .to_string();
                if !id.is_empty() && !seen_ids.insert(id) {
                    continue;
                }
                items.push(normalized);
                if items.len() >= safe_top as usize {
                    break;
                }
            }
        }

        next_url = payload
            .get("@odata.nextLink")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .to_string();
    }

    Ok(json!({
        "items": items,
        "fetchedAt": chrono::Utc::now().to_rfc3339(),
    }))
}

#[tauri::command]
fn read_array_store(app: AppHandle, key: String) -> Result<Option<Vec<Value>>, String> {
    let value = read_store_json_impl(&app, &key)?;
    match value {
        Some(Value::Array(values)) => Ok(Some(values)),
        Some(_) => Err("Store value is not an array".to_string()),
        None => Ok(None),
    }
}

#[tauri::command]
fn write_array_store(app: AppHandle, key: String, values: Vec<Value>) -> Result<bool, String> {
    write_store_json_impl(&app, &key, &Value::Array(values))?;
    Ok(true)
}

#[tauri::command]
fn read_data_json(app: AppHandle, key: String) -> StoreResponse<Value> {
    match read_data_json_impl(&app, &key) {
        Ok(Some(value)) => ok_response(value),
        Ok(None) => err_response("Not found"),
        Err(error) => err_response(error),
    }
}

#[tauri::command]
fn write_data_json(app: AppHandle, key: String, value: Value) -> StoreResponse<Value> {
    match write_data_json_impl(&app, &key, &value) {
        Ok(()) => ok_response(Value::Null),
        Err(error) => err_response(error),
    }
}

#[tauri::command]
async fn list_share_drive_children(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: ShareDrivePayload,
) -> Result<StoreResponse<Value>, String> {
    match list_share_drive_children_impl(&app, state.inner(), &payload).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => Ok(err_response(error)),
    }
}

#[tauri::command]
async fn get_share_drive_download_url(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: ShareDrivePayload,
) -> Result<StoreResponse<Value>, String> {
    match get_share_drive_download_url_impl(&app, state.inner(), &payload).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => Ok(err_response(error)),
    }
}

#[tauri::command]
async fn download_share_drive_file(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: ShareDrivePayload,
) -> Result<StoreResponse<Value>, String> {
    match download_share_drive_file_impl(&app, state.inner(), &payload).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => Ok(err_response(error)),
    }
}

#[tauri::command]
async fn upload_share_drive_file(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: ShareDrivePayload,
) -> Result<StoreResponse<Value>, String> {
    match upload_share_drive_file_impl(&app, state.inner(), &payload).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => Ok(err_response(error)),
    }
}

#[tauri::command]
async fn request_graph_device_code(
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<StoreResponse<Value>, String> {
    match request_device_code_impl(&app, &state.client).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => {
            auth_log("device-code:start-error", &error);
            Ok(err_response(error))
        }
    }
}

#[tauri::command]
async fn poll_graph_device_code(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: DeviceCodePollPayload,
) -> Result<StoreResponse<Value>, String> {
    match poll_device_code_impl(&app, state.inner(), &payload.device_code).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => {
            auth_log("device-code:poll-error", &error);
            Ok(err_response(error))
        }
    }
}

#[tauri::command]
fn get_graph_session(app: AppHandle, state: State<'_, AppState>) -> StoreResponse<Value> {
    match graph_session_summary(&app, state.inner()) {
        Ok(value) => ok_response(value),
        Err(error) => err_response(error),
    }
}

#[tauri::command]
async fn list_outlook_messages(
    app: AppHandle,
    state: State<'_, AppState>,
    payload: OutlookMessagesPayload,
) -> Result<StoreResponse<Value>, String> {
    match list_outlook_messages_impl(&app, state.inner(), &payload).await {
        Ok(value) => Ok(ok_response(value)),
        Err(error) => Ok(err_response(error)),
    }
}

fn main() {
    tauri::Builder::default()
        .manage(AppState {
            client: Client::new(),
            graph_session: Mutex::new(GraphSession::default()),
        })
        .invoke_handler(tauri::generate_handler![
            read_array_store,
            write_array_store,
            read_data_json,
            write_data_json,
            list_share_drive_children,
            get_share_drive_download_url,
            download_share_drive_file,
            upload_share_drive_file,
            request_graph_device_code,
            poll_graph_device_code,
            get_graph_session,
            list_outlook_messages,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
