const { app, BrowserWindow, ipcMain } = require("electron");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const GRAPH_SESSION_KEY = "graph_session_v1";
const DEFAULT_DELEGATED_SCOPES = [
  "offline_access",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "User.Read",
];
const UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024; // 5MB, multiple of 320 KiB.

let sessionAccessToken = "";
let sessionAccessTokenExpiresAt = 0;
let sessionRefreshToken = "";

let cachedSharedriveConfig = null;
let sharedriveConfigLoaded = false;

function loadSharedriveConfig() {
  if (sharedriveConfigLoaded) return cachedSharedriveConfig;
  sharedriveConfigLoaded = true;

  const candidates = [
    path.join(app.getAppPath(), "data", "sharedrive-tasks.js"),
    path.join(app.getAppPath(), "web", "data", "sharedrive-tasks.js"),
  ];

  for (const filePath of candidates) {
    if (!fs.existsSync(filePath)) continue;
    try {
      const source = fs.readFileSync(filePath, "utf8");
      const sandbox = { window: {} };
      vm.runInNewContext(source, sandbox, { filename: filePath });
      const config =
        sandbox && sandbox.window && sandbox.window.SHAREDRIVE_TASKS && typeof sandbox.window.SHAREDRIVE_TASKS === "object"
          ? sandbox.window.SHAREDRIVE_TASKS
          : null;
      if (config) {
        cachedSharedriveConfig = config;
        break;
      }
    } catch {
      // Ignore invalid config and continue to other candidates.
    }
  }

  return cachedSharedriveConfig;
}

function getStoreDir() {
  const teamConfigPath = path.join(app.getAppPath(), "data", "team-store-path.json");
  if (fs.existsSync(teamConfigPath)) {
    try {
      const cfg = JSON.parse(fs.readFileSync(teamConfigPath, "utf8"));
      const customDir = cfg && typeof cfg.storeDir === "string" ? cfg.storeDir.trim() : "";
      if (customDir) return path.resolve(customDir);
    } catch {
      // Ignore invalid team config and fall back to userData
    }
  }

  // Default: OS-level app data directory.
  return path.join(app.getPath("userData"), "runtime-store");
}

function getStorePath(key) {
  const safeKey = String(key || "").replace(/[^a-z0-9_-]/gi, "_");
  return path.join(getStoreDir(), `${safeKey}.json`);
}

function ensureStoreDir() {
  fs.mkdirSync(getStoreDir(), { recursive: true });
}

function readArrayStore(key) {
  ensureStoreDir();
  const filePath = getStorePath(key);
  if (!fs.existsSync(filePath)) return null;
  try {
    const raw = fs.readFileSync(filePath, "utf8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function writeArrayStore(key, values) {
  ensureStoreDir();
  const filePath = getStorePath(key);
  const payload = JSON.stringify(Array.isArray(values) ? values : [], null, 2);
  fs.writeFileSync(filePath, payload, "utf8");
  return true;
}

function readJsonStore(key) {
  ensureStoreDir();
  const filePath = getStorePath(key);
  if (!fs.existsSync(filePath)) return null;
  try {
    const raw = fs.readFileSync(filePath, "utf8");
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function writeJsonStore(key, value) {
  ensureStoreDir();
  const filePath = getStorePath(key);
  const payload = JSON.stringify(value && typeof value === "object" ? value : {}, null, 2);
  fs.writeFileSync(filePath, payload, "utf8");
  return true;
}

function toBase64Url(value) {
  return Buffer.from(String(value || ""), "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  if (!response.ok) {
    const message =
      (payload && payload.error && (payload.error.message || payload.error_description)) ||
      response.statusText ||
      "Request failed";
    throw new Error(`HTTP ${response.status}: ${message}`);
  }

  return payload;
}

async function fetchJsonWithStatus(url, options) {
  const response = await fetch(url, options);
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }
  return { ok: response.ok, status: response.status, payload };
}

function setSessionAccessToken(token, expiresInSeconds) {
  sessionAccessToken = String(token || "").trim();
  const expiresIn = Number(expiresInSeconds || 0);
  sessionAccessTokenExpiresAt = Date.now() + Math.max(expiresIn, 0) * 1000;
  persistGraphSession();
}

function setSessionRefreshToken(token) {
  sessionRefreshToken = String(token || "").trim();
  persistGraphSession();
}

function persistGraphSession() {
  writeJsonStore(GRAPH_SESSION_KEY, {
    accessToken: sessionAccessToken,
    expiresAt: sessionAccessTokenExpiresAt,
    refreshToken: sessionRefreshToken,
  });
}

function loadGraphSession() {
  if (sessionRefreshToken || sessionAccessToken) return;
  const payload = readJsonStore(GRAPH_SESSION_KEY);
  if (!payload) return;
  const refreshToken = String(payload.refreshToken || "").trim();
  if (refreshToken) sessionRefreshToken = refreshToken;
  const accessToken = String(payload.accessToken || "").trim();
  const expiresAt = Number(payload.expiresAt || 0);
  if (accessToken && expiresAt) {
    sessionAccessToken = accessToken;
    sessionAccessTokenExpiresAt = expiresAt;
  }
}

function getTenantId() {
  const envTenant = String(process.env.PLUTUS_AZURE_TENANT_ID || "").trim();
  if (envTenant) return envTenant;
  const config = loadSharedriveConfig();
  const cfgTenant = config && typeof config.azureTenantId === "string" ? config.azureTenantId.trim() : "";
  return cfgTenant || "common";
}

function getClientId() {
  const envClient = String(process.env.PLUTUS_AZURE_CLIENT_ID || "").trim();
  if (envClient) return envClient;
  const config = loadSharedriveConfig();
  return config && typeof config.azureClientId === "string" ? config.azureClientId.trim() : "";
}

function getDelegatedScopes() {
  const raw = String(process.env.PLUTUS_GRAPH_SCOPES || "").trim();
  if (raw) {
    return raw
      .split(/[\s,]+/)
      .map((scope) => scope.trim())
      .filter(Boolean);
  }
  const config = loadSharedriveConfig();
  const cfgScopes = config && typeof config.graphScopes === "string" ? config.graphScopes.trim() : "";
  if (cfgScopes) {
    return cfgScopes
      .split(/[\s,]+/)
      .map((scope) => scope.trim())
      .filter(Boolean);
  }
  return DEFAULT_DELEGATED_SCOPES.slice();
}

async function refreshAccessToken(refreshToken) {
  const clientId = getClientId();
  if (!clientId) {
    throw new Error("Missing PLUTUS_AZURE_CLIENT_ID for refresh token flow.");
  }

  const tenantId = getTenantId();
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    tenantId,
  )}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: String(refreshToken || ""),
    scope: getDelegatedScopes().join(" "),
  });

  const result = await fetchJsonWithStatus(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!result.ok) {
    const payload = result.payload || {};
    const message = payload.error_description || payload.error || "Refresh token flow failed.";
    throw new Error(message);
  }

  const payload = result.payload || {};
  const accessToken = String(payload.access_token || "").trim();
  const expiresIn = Number(payload.expires_in || 0);
  const newRefreshToken = String(payload.refresh_token || "").trim();
  setSessionAccessToken(accessToken, expiresIn);
  if (newRefreshToken) setSessionRefreshToken(newRefreshToken);

  return accessToken;
}

async function getClientCredentialsToken() {
  const tenantId = String(process.env.PLUTUS_AZURE_TENANT_ID || "").trim();
  const clientId = String(process.env.PLUTUS_AZURE_CLIENT_ID || "").trim();
  const clientSecret = String(process.env.PLUTUS_AZURE_CLIENT_SECRET || "").trim();

  if (!tenantId || !clientId || !clientSecret) return "";

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: GRAPH_SCOPE,
    grant_type: "client_credentials",
  });

  const tokenPayload = await fetchJson(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  return String(tokenPayload && tokenPayload.access_token ? tokenPayload.access_token : "").trim();
}

async function resolveGraphAccessToken(overrideToken) {
  const explicit = String(overrideToken || "").trim();
  if (explicit) return explicit;

  if (sessionAccessToken && Date.now() < sessionAccessTokenExpiresAt - 60000) {
    return sessionAccessToken;
  }

  loadGraphSession();

  if (sessionAccessToken && Date.now() < sessionAccessTokenExpiresAt - 60000) {
    return sessionAccessToken;
  }

  const envToken = String(process.env.PLUTUS_GRAPH_ACCESS_TOKEN || "").trim();
  if (envToken) return envToken;

  if (sessionRefreshToken) {
    try {
      return await refreshAccessToken(sessionRefreshToken);
    } catch (error) {
      sessionRefreshToken = "";
      persistGraphSession();
    }
  }

  const clientToken = await getClientCredentialsToken();
  if (clientToken) return clientToken;

  throw new Error(
    "No Microsoft Graph access token available. Provide one in the page, or set PLUTUS_GRAPH_ACCESS_TOKEN, or set PLUTUS_AZURE_TENANT_ID / PLUTUS_AZURE_CLIENT_ID / PLUTUS_AZURE_CLIENT_SECRET.",
  );
}

async function requestDeviceCode() {
  const clientId = getClientId();
  if (!clientId) {
    throw new Error("Missing PLUTUS_AZURE_CLIENT_ID for device code flow.");
  }

  const tenantId = getTenantId();
  const deviceCodeUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    tenantId,
  )}/oauth2/v2.0/devicecode`;
  const scopes = getDelegatedScopes().join(" ");
  const body = new URLSearchParams({
    client_id: clientId,
    scope: scopes,
  });

  return fetchJson(deviceCodeUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
}

async function pollDeviceCode(deviceCode) {
  const clientId = getClientId();
  if (!clientId) {
    throw new Error("Missing PLUTUS_AZURE_CLIENT_ID for device code flow.");
  }

  const tenantId = getTenantId();
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    tenantId,
  )}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "urn:ietf:params:oauth:grant-type:device_code",
    device_code: String(deviceCode || ""),
  });

  const result = await fetchJsonWithStatus(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!result.ok) {
    const payload = result.payload || {};
    return {
      ok: false,
      error: payload.error || "authorization_pending",
      error_description: payload.error_description || "Authorization pending.",
      interval: payload.interval,
    };
  }

  const payload = result.payload || {};
  const accessToken = String(payload.access_token || "").trim();
  const expiresIn = Number(payload.expires_in || 0);
  setSessionAccessToken(accessToken, expiresIn);
  const refreshToken = String(payload.refresh_token || "").trim();
  if (refreshToken) setSessionRefreshToken(refreshToken);

  return {
    ok: true,
    accessToken,
    expiresIn,
    refreshToken: String(payload.refresh_token || ""),
    scope: String(payload.scope || ""),
    tokenType: String(payload.token_type || ""),
  };
}

function normalizeDriveItem(item) {
  return {
    id: item && item.id ? item.id : "",
    name: item && item.name ? item.name : "",
    webUrl: item && item.webUrl ? item.webUrl : "",
    size: item && typeof item.size === "number" ? item.size : null,
    lastModifiedDateTime: item && item.lastModifiedDateTime ? item.lastModifiedDateTime : "",
    isFolder: Boolean(item && item.folder),
    isFile: Boolean(item && item.file),
    mimeType: item && item.file && item.file.mimeType ? item.file.mimeType : "",
    childCount:
      item && item.folder && typeof item.folder.childCount === "number" ? item.folder.childCount : null,
    parentPath:
      item && item.parentReference && typeof item.parentReference.path === "string"
        ? item.parentReference.path
        : "",
    downloadUrl: item && item["@microsoft.graph.downloadUrl"] ? item["@microsoft.graph.downloadUrl"] : "",
  };
}

async function getShareDriveItem(shareUrl, token, selectFields) {
  const cleanShareUrl = String(shareUrl || "").trim();
  if (!cleanShareUrl) throw new Error("Share URL is required.");

  const encodedShare = `u!${toBase64Url(cleanShareUrl)}`;
  const fields = Array.isArray(selectFields) && selectFields.length
    ? selectFields
    : ["id", "name", "webUrl", "parentReference", "folder", "file", "size", "lastModifiedDateTime", "@microsoft.graph.downloadUrl"];
  const params = new URLSearchParams({
    $select: fields.join(","),
  });
  const url = `https://graph.microsoft.com/v1.0/shares/${encodedShare}/driveItem?${params.toString()}`;

  return fetchJson(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });
}

async function listDriveItemChildren(driveId, itemId, token) {
  const fields = [
    "id",
    "name",
    "webUrl",
    "parentReference",
    "folder",
    "file",
    "size",
    "lastModifiedDateTime",
    "@microsoft.graph.downloadUrl",
  ];
  const params = new URLSearchParams({
    $top: "200",
    $select: fields.join(","),
  });
  let nextUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId,
  )}/items/${encodeURIComponent(itemId)}/children?${params.toString()}`;
  const items = [];

  while (nextUrl) {
    const page = await fetchJson(nextUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    const values = Array.isArray(page && page.value) ? page.value : [];
    values.forEach((item) => items.push(item));
    nextUrl = page && typeof page["@odata.nextLink"] === "string" ? page["@odata.nextLink"] : "";
  }

  return items;
}

async function listShareDriveItems({ shareUrl, accessToken }) {
  const token = await resolveGraphAccessToken(accessToken);
  const rootItem = await getShareDriveItem(shareUrl, token);
  const driveId =
    (rootItem && rootItem.parentReference && rootItem.parentReference.driveId) ||
    (rootItem && rootItem.remoteItem && rootItem.remoteItem.parentReference && rootItem.remoteItem.parentReference.driveId) ||
    "";

  let items = [];
  if (rootItem && rootItem.folder && driveId) {
    const children = await listDriveItemChildren(driveId, rootItem.id, token);
    items = children.map(normalizeDriveItem);
  }

  const totalFolders = items.filter((item) => item.isFolder).length;
  const totalFiles = items.filter((item) => item.isFile).length;

  return {
    root: normalizeDriveItem(rootItem),
    driveId,
    items,
    totalItems: items.length,
    totalFolders,
    totalFiles,
    fetchedAt: new Date().toISOString(),
  };
}

async function getShareDriveDownloadUrl({ shareUrl, accessToken, driveId, itemId }) {
  const token = await resolveGraphAccessToken(accessToken);

  let item;
  if (driveId && itemId) {
    const params = new URLSearchParams({
      $select: "id,name,webUrl,parentReference,@microsoft.graph.downloadUrl",
    });
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
      driveId,
    )}/items/${encodeURIComponent(itemId)}?${params.toString()}`;
    item = await fetchJson(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });
  } else {
    item = await getShareDriveItem(shareUrl, token, [
      "id",
      "name",
      "webUrl",
      "parentReference",
      "@microsoft.graph.downloadUrl",
    ]);
  }

  const downloadUrl = item && item["@microsoft.graph.downloadUrl"] ? item["@microsoft.graph.downloadUrl"] : "";
  if (!downloadUrl) {
    throw new Error("Download URL not available for this item.");
  }

  return {
    id: item && item.id ? item.id : "",
    name: item && item.name ? item.name : "",
    webUrl: item && item.webUrl ? item.webUrl : "",
    driveId:
      driveId || (item && item.parentReference && item.parentReference.driveId) || "",
    downloadUrl,
  };
}

async function downloadShareDriveFile({ shareUrl, accessToken, driveId, itemId }) {
  const token = await resolveGraphAccessToken(accessToken);

  let resolvedDriveId = String(driveId || "").trim();
  let resolvedItemId = String(itemId || "").trim();
  if (!resolvedDriveId || !resolvedItemId) {
    const item = await getShareDriveItem(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
    ]);
    resolvedItemId = item && item.id ? item.id : "";
    resolvedDriveId =
      (item && item.parentReference && item.parentReference.driveId) ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.driveId) ||
      "";
  }

  if (!resolvedDriveId || !resolvedItemId) {
    throw new Error("Unable to resolve sharedrive file for download.");
  }

  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    resolvedDriveId,
  )}/items/${encodeURIComponent(resolvedItemId)}/content`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const text = await response.text();
  if (!response.ok) {
    let message = response.statusText || "Download failed.";
    try {
      const payload = text ? JSON.parse(text) : null;
      if (payload && payload.error && payload.error.message) {
        message = payload.error.message;
      }
    } catch {
      // ignore json parse
    }
    throw new Error(`Download failed: ${message}`);
  }

  return {
    driveId: resolvedDriveId,
    itemId: resolvedItemId,
    text,
  };
}

async function createUploadSession({ driveId, parentItemId, fileName, token, conflictBehavior }) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId,
  )}/items/${encodeURIComponent(parentItemId)}:/${encodeURIComponent(fileName)}:/createUploadSession`;
  const payload = {
    item: {
      "@microsoft.graph.conflictBehavior": conflictBehavior || "replace",
    },
  };

  return fetchJson(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
}

async function uploadWithSession(uploadUrl, buffer) {
  const total = buffer.length;
  let start = 0;

  while (start < total) {
    const end = Math.min(start + UPLOAD_CHUNK_SIZE, total);
    const chunk = buffer.subarray(start, end);
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${end - 1}/${total}`,
      },
      body: chunk,
    });

    const text = await response.text();
    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = null;
    }

    if (!response.ok) {
      const message =
        (payload && payload.error && (payload.error.message || payload.error_description)) ||
        response.statusText ||
        "Upload failed";
      throw new Error(`Upload failed: ${message}`);
    }

    if (end === total) {
      return payload;
    }

    start = end;
  }

  return null;
}

async function uploadShareDriveFile({ shareUrl, accessToken, parentItemId, fileName, contentBase64, conflictBehavior }) {
  const token = await resolveGraphAccessToken(accessToken);
  const rootItem = await getShareDriveItem(shareUrl, token, [
    "id",
    "name",
    "parentReference",
    "folder",
    "file",
  ]);

  const driveId =
    (rootItem && rootItem.parentReference && rootItem.parentReference.driveId) ||
    (rootItem && rootItem.remoteItem && rootItem.remoteItem.parentReference && rootItem.remoteItem.parentReference.driveId) ||
    "";
  if (!driveId) throw new Error("Unable to resolve SharePoint drive ID.");

  let targetParentId = String(parentItemId || "").trim();
  if (!targetParentId) {
    if (rootItem && rootItem.folder) {
      targetParentId = rootItem.id;
    } else if (rootItem && rootItem.parentReference && rootItem.parentReference.id) {
      targetParentId = rootItem.parentReference.id;
    }
  }

  if (!targetParentId) throw new Error("Target folder ID is required for upload.");

  const cleanFileName = String(fileName || "").trim();
  if (!cleanFileName) throw new Error("File name is required.");

  const buffer = Buffer.from(String(contentBase64 || ""), "base64");
  if (!buffer.length) throw new Error("Upload content is empty.");

  const session = await createUploadSession({
    driveId,
    parentItemId: targetParentId,
    fileName: cleanFileName,
    token,
    conflictBehavior,
  });

  if (!session || !session.uploadUrl) {
    throw new Error("Failed to create upload session.");
  }

  const uploaded = await uploadWithSession(session.uploadUrl, buffer);
  return {
    item: uploaded ? normalizeDriveItem(uploaded) : null,
    driveId,
    parentItemId: targetParentId,
  };
}

async function listShareDriveFolders({ shareUrl, accessToken }) {
  const cleanShareUrl = String(shareUrl || "").trim();
  if (!cleanShareUrl) {
    throw new Error("Share URL is required.");
  }

  const token = await resolveGraphAccessToken(accessToken);
  const encodedShare = `u!${toBase64Url(cleanShareUrl)}`;
  const base = `https://graph.microsoft.com/v1.0/shares/${encodedShare}/driveItem`;

  const root = await fetchJson(`${base}?$select=id,name,webUrl,parentReference`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  const folders = [];
  let nextUrl = `${base}/children?$top=200&$select=id,name,folder,webUrl,lastModifiedDateTime,size,parentReference`;

  while (nextUrl) {
    const page = await fetchJson(nextUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    const values = Array.isArray(page && page.value) ? page.value : [];
    values.forEach((item) => {
      if (item && item.folder) {
        folders.push({
          id: item.id || "",
          name: item.name || "",
          webUrl: item.webUrl || "",
          childCount:
            item.folder && typeof item.folder.childCount === "number" ? item.folder.childCount : null,
          lastModifiedDateTime: item.lastModifiedDateTime || "",
          size: typeof item.size === "number" ? item.size : null,
          parentPath:
            item.parentReference && typeof item.parentReference.path === "string"
              ? item.parentReference.path
              : "",
        });
      }
    });

    nextUrl = page && typeof page["@odata.nextLink"] === "string" ? page["@odata.nextLink"] : "";
  }

  return {
    root: {
      id: root && root.id ? root.id : "",
      name: root && root.name ? root.name : "",
      webUrl: root && root.webUrl ? root.webUrl : cleanShareUrl,
      parentPath:
        root && root.parentReference && typeof root.parentReference.path === "string"
          ? root.parentReference.path
          : "",
    },
    folders,
    totalFolders: folders.length,
    fetchedAt: new Date().toISOString(),
  };
}

function createMainWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.cjs"),
    },
  });

  win.loadFile(path.join(app.getAppPath(), "index.html"));
}

app.whenReady().then(() => {
  ipcMain.on("plutus:store:read-array", (event, key) => {
    event.returnValue = readArrayStore(key);
  });

  ipcMain.on("plutus:store:write-array", (event, payload) => {
    const key = payload && payload.key;
    const values = payload && payload.values;
    event.returnValue = writeArrayStore(key, values);
  });

  ipcMain.handle("plutus:sharedrive:list-folders", async (_event, payload) => {
    try {
      return {
        ok: true,
        data: await listShareDriveFolders(payload || {}),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to fetch sharedrive folders.",
      };
    }
  });

  ipcMain.handle("plutus:sharedrive:list-items", async (_event, payload) => {
    try {
      return {
        ok: true,
        data: await listShareDriveItems(payload || {}),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to list sharedrive items.",
      };
    }
  });

  ipcMain.handle("plutus:sharedrive:get-download-url", async (_event, payload) => {
    try {
      return {
        ok: true,
        data: await getShareDriveDownloadUrl(payload || {}),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to resolve download URL.",
      };
    }
  });

  ipcMain.handle("plutus:sharedrive:download-file", async (_event, payload) => {
    try {
      return {
        ok: true,
        data: await downloadShareDriveFile(payload || {}),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to download file.",
      };
    }
  });

  ipcMain.handle("plutus:sharedrive:upload-file", async (_event, payload) => {
    try {
      return {
        ok: true,
        data: await uploadShareDriveFile(payload || {}),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to upload file.",
      };
    }
  });

  ipcMain.handle("plutus:graph:device-code", async () => {
    try {
      return {
        ok: true,
        data: await requestDeviceCode(),
      };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to start device code flow.",
      };
    }
  });

  ipcMain.handle("plutus:graph:device-code:poll", async (_event, payload) => {
    try {
      const result = await pollDeviceCode(payload && payload.deviceCode ? payload.deviceCode : "");
      return { ok: true, data: result };
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : "Failed to poll device code.",
      };
    }
  });

  createMainWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
