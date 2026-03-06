const { app, BrowserWindow, ipcMain } = require("electron");
const fs = require("fs");
const path = require("path");

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

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

  const envToken = String(process.env.PLUTUS_GRAPH_ACCESS_TOKEN || "").trim();
  if (envToken) return envToken;

  const clientToken = await getClientCredentialsToken();
  if (clientToken) return clientToken;

  throw new Error(
    "No Microsoft Graph access token available. Provide one in the page, or set PLUTUS_GRAPH_ACCESS_TOKEN, or set PLUTUS_AZURE_TENANT_ID / PLUTUS_AZURE_CLIENT_ID / PLUTUS_AZURE_CLIENT_SECRET.",
  );
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

  createMainWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
