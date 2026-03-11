(function initAppCore(global) {
  const STORAGE_KEYS = {
    deals: "deals_data_v1",
    tasks: "owner_tasks_v1",
  };

  const AUTO_CONTACT_TASK_PREFIX = "auto-contact-status";
  const SHARED_TASKS_DEFAULTS = {
    enabled: false,
    shareUrl: "",
    fileName: "sharedrive-task.json",
    pollIntervalMs: 60000,
    accessToken: "",
    downloadUrl: "",
    parentItemId: "",
    writerLabel: "",
    azureClientId: "",
    azureTenantId: "common",
    graphScopes: "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read",
  };
  const GRAPH_BROWSER_SESSION_KEY = "plutus_graph_session_v1";
  const GRAPH_UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024;
  const NATIVE_HTTP_PLUGIN = "CapacitorHttp";

  let nativeHttpClient = null;
  let nativeHttpChecked = false;

  let tasksCache = null;
  let sharedTasksSyncTimer = null;
  let sharedTasksUploadTimer = null;
  let sharedTasksUploadPending = null;
  let sharedTasksDirty = false;
  const sharedTasksState = {
    started: false,
    inFlight: false,
    uploadInFlight: false,
    lastSyncAt: "",
    lastUploadAt: "",
    lastActionAt: "",
    lastRemoteUpdatedAt: "",
    lastError: "",
  };

  function normalizeValue(value) {
    return String(value || "").trim().toLowerCase();
  }

  function getNativeHttpClient() {
    if (nativeHttpChecked) return nativeHttpClient;
    nativeHttpChecked = true;

    const cap = global.Capacitor;
    if (!cap || typeof cap.getPlatform !== "function") {
      nativeHttpClient = null;
      return null;
    }

    if (cap.getPlatform() === "web") {
      nativeHttpClient = null;
      return null;
    }

    try {
      if (cap.Plugins && cap.Plugins[NATIVE_HTTP_PLUGIN]) {
        nativeHttpClient = { type: "plugin", client: cap.Plugins[NATIVE_HTTP_PLUGIN] };
        return nativeHttpClient;
      }
      if (typeof cap.registerPlugin === "function") {
        const plugin = cap.registerPlugin(NATIVE_HTTP_PLUGIN);
        if (plugin) {
          nativeHttpClient = { type: "plugin", client: plugin };
          return nativeHttpClient;
        }
      }
      if (typeof cap.nativePromise === "function") {
        nativeHttpClient = { type: "nativePromise", client: cap };
        return nativeHttpClient;
      }
    } catch {
      // ignore
    }

    nativeHttpClient = null;

    return nativeHttpClient;
  }

  function normalizeHttpResponse(response) {
    const status = response && typeof response.status === "number" ? response.status : 0;
    const ok = status >= 200 && status < 300;
    const data = response && typeof response.data !== "undefined" ? response.data : null;
    return { ok, status, data };
  }

  function toFormParams(params) {
    if (params instanceof URLSearchParams) {
      return Object.fromEntries(params.entries());
    }
    if (params && typeof params === "object") {
      return params;
    }
    return {};
  }

  async function postFormUrlEncoded(url, params) {
    const body =
      params instanceof URLSearchParams ? params.toString() : new URLSearchParams(params || {}).toString();
    const formParams = toFormParams(params);
    const headers = { "Content-Type": "application/x-www-form-urlencoded" };
    const nativeHttp = getNativeHttpClient();

    if (nativeHttp) {
      try {
        if (nativeHttp.type === "plugin" && nativeHttp.client && typeof nativeHttp.client.post === "function") {
          const response = await nativeHttp.client.post({
            url,
            headers,
            data: formParams,
            responseType: "json",
          });
          return normalizeHttpResponse(response);
        }
        if (nativeHttp.type === "nativePromise" && nativeHttp.client) {
          const response = await nativeHttp.client.nativePromise("CapacitorHttp", "request", {
            url,
            method: "POST",
            headers,
            data: formParams,
            responseType: "json",
          });
          return normalizeHttpResponse(response);
        }
      } catch (error) {
        return {
          ok: false,
          status: 0,
          data: null,
          error: (error && error.message) || "Native HTTP request failed.",
        };
      }
    }

    try {
      const response = await fetch(url, {
        method: "POST",
        headers,
        body,
      });
      let data = null;
      try {
        data = await response.json();
      } catch {
        data = null;
      }
      return { ok: response.ok, status: response.status, data };
    } catch (error) {
      return {
        ok: false,
        status: 0,
        data: null,
        error: (error && error.message) || "Request failed.",
      };
    }
  }

  function cloneArray(value) {
    if (!Array.isArray(value)) return [];
    return JSON.parse(JSON.stringify(value));
  }

  function hasDesktopBridge() {
    return Boolean(
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.readArrayStore === "function" &&
      typeof global.PlutusDesktop.writeArrayStore === "function",
    );
  }

  function hasShareDriveBridge() {
    return Boolean(
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.getShareDriveDownloadUrl === "function" &&
      typeof global.PlutusDesktop.uploadShareDriveFile === "function",
    );
  }

  function getSharedTasksConfig() {
    const raw =
      global.SHAREDRIVE_TASKS ||
      (global.DASHBOARD_CONFIG && global.DASHBOARD_CONFIG.settings && global.DASHBOARD_CONFIG.settings.sharedTasks) ||
      {};

    const config = Object.assign({}, SHARED_TASKS_DEFAULTS, raw || {});
    config.enabled = Boolean(raw && raw.enabled);
    config.shareUrl = String(config.shareUrl || "").trim();
    config.fileName = String(config.fileName || "").trim();
    config.downloadUrl = String(config.downloadUrl || "").trim();
    config.accessToken = String(config.accessToken || "").trim();
    config.parentItemId = String(config.parentItemId || "").trim();
    config.writerLabel = String(config.writerLabel || "").trim();
    config.azureClientId = String(config.azureClientId || "").trim();
    config.azureTenantId = String(config.azureTenantId || "").trim() || "common";
    config.graphScopes = String(config.graphScopes || "").trim() || SHARED_TASKS_DEFAULTS.graphScopes;

    const interval = Number(config.pollIntervalMs);
    config.pollIntervalMs = Number.isFinite(interval) ? interval : SHARED_TASKS_DEFAULTS.pollIntervalMs;

    return config;
  }

  function getBrowserAuthConfig() {
    const config = getSharedTasksConfig();
    if (!config.azureClientId) return null;
    return {
      clientId: config.azureClientId,
      tenantId: config.azureTenantId || "common",
      scopes: config.graphScopes || SHARED_TASKS_DEFAULTS.graphScopes,
    };
  }

  function readArrayFromStorage(key) {
    if (hasDesktopBridge()) {
      try {
        const values = global.PlutusDesktop.readArrayStore(key);
        return Array.isArray(values) ? values : null;
      } catch (error) {
        console.warn(`[AppCore] Desktop read failed for ${key}`, error);
      }
    }

    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : null;
    } catch (error) {
      console.warn(`[AppCore] Failed to read ${key}`, error);
      return null;
    }
  }

  function writeArrayToStorage(key, values) {
    if (hasDesktopBridge()) {
      try {
        global.PlutusDesktop.writeArrayStore(key, Array.isArray(values) ? values : []);
        return;
      } catch (error) {
        console.warn(`[AppCore] Desktop write failed for ${key}`, error);
      }
    }

    try {
      localStorage.setItem(key, JSON.stringify(Array.isArray(values) ? values : []));
    } catch (error) {
      console.warn(`[AppCore] Failed to write ${key}`, error);
    }
  }

  function readBrowserGraphSession() {
    try {
      const raw = localStorage.getItem(GRAPH_BROWSER_SESSION_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" ? parsed : null;
    } catch {
      return null;
    }
  }

  function writeBrowserGraphSession(payload) {
    try {
      localStorage.setItem(GRAPH_BROWSER_SESSION_KEY, JSON.stringify(payload || {}));
      return true;
    } catch {
      return false;
    }
  }

  async function requestDeviceCodeBrowser() {
    const auth = getBrowserAuthConfig();
    if (!auth) throw new Error("Missing Azure client ID for device code flow.");
    const url = `https://login.microsoftonline.com/${encodeURIComponent(auth.tenantId)}/oauth2/v2.0/devicecode`;
    const body = new URLSearchParams({
      client_id: auth.clientId,
      scope: auth.scopes,
    });
    const response = await postFormUrlEncoded(url, body);
    const payload = response.data;
    if (!response.ok) {
      const message = payload && (payload.error_description || payload.error)
        ? (payload.error_description || payload.error)
        : response.error || "Device code flow failed.";
      return { ok: false, error: message };
    }
    return { ok: true, data: payload };
  }

  async function pollDeviceCodeBrowser(deviceCode) {
    const auth = getBrowserAuthConfig();
    if (!auth) throw new Error("Missing Azure client ID for device code flow.");
    const url = `https://login.microsoftonline.com/${encodeURIComponent(auth.tenantId)}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: auth.clientId,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: String(deviceCode || ""),
    });
    const response = await postFormUrlEncoded(url, body);
    const payload = response.data;
    if (!response.ok) {
      return {
        ok: true,
        data: {
          ok: false,
          error: payload && payload.error ? payload.error : "authorization_pending",
          error_description: payload && payload.error_description ? payload.error_description : "Authorization pending.",
          interval: payload && payload.interval ? payload.interval : null,
        },
      };
    }

    const accessToken = String(payload.access_token || "").trim();
    const refreshToken = String(payload.refresh_token || "").trim();
    const expiresIn = Number(payload.expires_in || 0);
    const expiresAt = Date.now() + Math.max(expiresIn, 0) * 1000;
    writeBrowserGraphSession({
      accessToken,
      refreshToken,
      expiresAt,
      scope: payload.scope || "",
      tokenType: payload.token_type || "",
    });

    return {
      ok: true,
      data: {
        ok: true,
        accessToken,
        refreshToken,
        expiresIn,
        scope: payload.scope || "",
        tokenType: payload.token_type || "",
      },
    };
  }

  async function refreshBrowserAccessToken(refreshToken) {
    const auth = getBrowserAuthConfig();
    if (!auth) throw new Error("Missing Azure client ID for refresh token flow.");
    const url = `https://login.microsoftonline.com/${encodeURIComponent(auth.tenantId)}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: auth.clientId,
      grant_type: "refresh_token",
      refresh_token: String(refreshToken || ""),
      scope: auth.scopes,
    });
    const response = await postFormUrlEncoded(url, body);
    const payload = response.data;
    if (!response.ok) {
      const message = payload && (payload.error_description || payload.error)
        ? (payload.error_description || payload.error)
        : response.error || "Refresh token flow failed.";
      throw new Error(message);
    }

    const accessToken = String(payload.access_token || "").trim();
    const newRefreshToken = String(payload.refresh_token || "").trim() || refreshToken;
    const expiresIn = Number(payload.expires_in || 0);
    const expiresAt = Date.now() + Math.max(expiresIn, 0) * 1000;
    writeBrowserGraphSession({
      accessToken,
      refreshToken: newRefreshToken,
      expiresAt,
      scope: payload.scope || "",
      tokenType: payload.token_type || "",
    });
    return accessToken;
  }

  async function resolveBrowserAccessToken() {
    const config = getSharedTasksConfig();
    if (config.accessToken) return config.accessToken;
    const session = readBrowserGraphSession();
    if (session && session.accessToken && session.expiresAt && Date.now() < session.expiresAt - 60000) {
      return session.accessToken;
    }
    if (session && session.refreshToken) {
      return refreshBrowserAccessToken(session.refreshToken);
    }
    throw new Error("Not signed in to Microsoft Graph.");
  }

  async function graphFetchJson(url, token, options) {
    const response = await fetch(url, {
      ...(options || {}),
      headers: {
        ...(options && options.headers ? options.headers : {}),
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
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
        "Request failed";
      throw new Error(message);
    }
    return payload;
  }

  function toBase64Url(value) {
    return btoa(unescape(encodeURIComponent(String(value || ""))))
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/g, "");
  }

  async function getShareDriveItemBrowser(shareUrl, token, selectFields) {
    const cleanShareUrl = String(shareUrl || "").trim();
    if (!cleanShareUrl) throw new Error("Share URL is required.");
    const encodedShare = `u!${toBase64Url(cleanShareUrl)}`;
    const fields = Array.isArray(selectFields) && selectFields.length
      ? selectFields
      : ["id", "name", "webUrl", "parentReference", "remoteItem", "folder", "file", "size", "lastModifiedDateTime"];
    const params = new URLSearchParams({
      $select: fields.join(","),
    });
    const url = `https://graph.microsoft.com/v1.0/shares/${encodedShare}/driveItem?${params.toString()}`;
    return graphFetchJson(url, token);
  }

  async function downloadShareDriveFileBrowser(shareUrl) {
    const token = await resolveBrowserAccessToken();
    const item = await getShareDriveItemBrowser(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
    ]);
    const itemId = item && item.id ? item.id : "";
    const driveId =
      (item && item.parentReference && item.parentReference.driveId) ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.driveId) ||
      "";
    if (!driveId || !itemId) {
      throw new Error("Unable to resolve sharedrive file for download.");
    }
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
      driveId,
    )}/items/${encodeURIComponent(itemId)}/content`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error("Download failed.");
    }
    return { text, driveId, itemId };
  }

  async function createUploadSessionBrowser({ driveId, parentItemId, fileName, token, conflictBehavior }) {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
      driveId,
    )}/items/${encodeURIComponent(parentItemId)}:/${encodeURIComponent(fileName)}:/createUploadSession`;
    const payload = {
      item: {
        "@microsoft.graph.conflictBehavior": conflictBehavior || "replace",
      },
    };
    return graphFetchJson(url, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  }

  function base64ToBytes(base64) {
    const cleaned = String(base64 || "").trim();
    const binary = atob(cleaned);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  async function uploadWithSessionBrowser(uploadUrl, bytes) {
    const total = bytes.length;
    let start = 0;
    while (start < total) {
      const end = Math.min(start + GRAPH_UPLOAD_CHUNK_SIZE, total);
      const chunk = bytes.slice(start, end);
      const response = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(chunk.length),
          "Content-Range": `bytes ${start}-${end - 1}/${total}`,
        },
        body: chunk,
      });
      const text = await response.text();
      if (!response.ok) {
        let message = response.statusText || "Upload failed";
        try {
          const payload = text ? JSON.parse(text) : null;
          if (payload && payload.error && payload.error.message) {
            message = payload.error.message;
          }
        } catch {
          // ignore
        }
        throw new Error(`Upload failed: ${message}`);
      }
      if (end === total) {
        try {
          return text ? JSON.parse(text) : null;
        } catch {
          return null;
        }
      }
      start = end;
    }
    return null;
  }

  async function uploadShareDriveFileBrowser({ shareUrl, parentItemId, fileName, contentBase64 }) {
    const token = await resolveBrowserAccessToken();
    const item = await getShareDriveItemBrowser(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
      "folder",
      "file",
    ]);
    const driveId =
      (item && item.parentReference && item.parentReference.driveId) ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.driveId) ||
      "";
    if (!driveId) throw new Error("Unable to resolve SharePoint drive ID.");

    let targetParentId = String(parentItemId || "").trim();
    if (!targetParentId) {
      if (item && item.folder) {
        targetParentId = item.id;
      } else if (item && item.parentReference && item.parentReference.id) {
        targetParentId = item.parentReference.id;
      }
    }
    if (!targetParentId) throw new Error("Target folder ID is required for upload.");

    const cleanFileName = String(fileName || "").trim();
    if (!cleanFileName) throw new Error("File name is required.");

    const bytes = base64ToBytes(contentBase64);
    if (!bytes.length) throw new Error("Upload content is empty.");

    const session = await createUploadSessionBrowser({
      driveId,
      parentItemId: targetParentId,
      fileName: cleanFileName,
      token,
      conflictBehavior: "replace",
    });
    if (!session || !session.uploadUrl) {
      throw new Error("Failed to create upload session.");
    }

    return uploadWithSessionBrowser(session.uploadUrl, bytes);
  }
  async function fetchJson(url, options) {
    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`Download failed (${response.status})`);
    }

    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      return response.json();
    }

    const text = await response.text();
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch (error) {
      throw new Error("Sharedrive tasks file is not valid JSON.");
    }
  }

  function parseSharedTasksPayload(payload) {
    if (typeof payload === "string") {
      try {
        return parseSharedTasksPayload(JSON.parse(payload));
      } catch {
        return { tasks: [], updatedAt: "" };
      }
    }

    if (Array.isArray(payload)) {
      return { tasks: payload, updatedAt: "" };
    }

    if (payload && Array.isArray(payload.tasks)) {
      return {
        tasks: payload.tasks,
        updatedAt: payload.updatedAt || payload.updated_at || "",
      };
    }

    return { tasks: [], updatedAt: "" };
  }

  function parseJsonText(text) {
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch {
      return null;
    }
  }

  function buildSharedTasksPayload(tasks, config) {
    return {
      version: 1,
      updatedAt: new Date().toISOString(),
      updatedBy: config.writerLabel || "",
      tasks: Array.isArray(tasks) ? tasks : [],
    };
  }

  function publishTasksUpdate(source) {
    try {
      const detail = {
        tasks: cloneArray(tasksCache || []),
        source: source || "local",
        syncedAt: sharedTasksState.lastSyncAt || "",
      };
      global.dispatchEvent(new CustomEvent("appcore:tasks-updated", { detail }));
    } catch (error) {
      console.warn("[AppCore] Failed to dispatch tasks update event", error);
    }
  }

  function publishSharedTasksStatus(stage) {
    try {
      const detail = {
        stage: stage || "idle",
        dirty: sharedTasksDirty,
        started: sharedTasksState.started,
        inFlight: sharedTasksState.inFlight,
        uploadInFlight: sharedTasksState.uploadInFlight,
        lastSyncAt: sharedTasksState.lastSyncAt,
        lastUploadAt: sharedTasksState.lastUploadAt,
        lastRemoteUpdatedAt: sharedTasksState.lastRemoteUpdatedAt,
        lastActionAt: sharedTasksState.lastActionAt,
        lastError: sharedTasksState.lastError,
      };
      global.dispatchEvent(new CustomEvent("appcore:tasks-sync", { detail }));
    } catch (error) {
      console.warn("[AppCore] Failed to dispatch tasks sync event", error);
    }
  }

  function loadDealsData() {
    return readArrayFromStorage(STORAGE_KEYS.deals) || cloneArray(global.DEALS);
  }

  function saveDealsData(deals) {
    writeArrayToStorage(STORAGE_KEYS.deals, deals);
  }

  function ensureSharedTasksSync() {
    const config = getSharedTasksConfig();
    if (!config.enabled || sharedTasksState.started) return;

    sharedTasksState.started = true;
    setTimeout(() => refreshTasksFromShareDrive("init"), 0);

    if (config.pollIntervalMs > 0) {
      sharedTasksSyncTimer = setInterval(
        () => refreshTasksFromShareDrive("poll"),
        config.pollIntervalMs,
      );
    }
  }

  function loadTasksData() {
    const config = getSharedTasksConfig();
    if (!tasksCache) {
      if (config.enabled) {
        tasksCache = [];
      } else {
        tasksCache = readArrayFromStorage(STORAGE_KEYS.tasks) || cloneArray(global.TASKS);
      }
    }

    ensureSharedTasksSync();
    publishSharedTasksStatus(sharedTasksState.started ? "ready" : "idle");
    return cloneArray(tasksCache);
  }

  function saveTasksData(tasks) {
    const config = getSharedTasksConfig();
    tasksCache = cloneArray(tasks);
    if (config.enabled) {
      sharedTasksDirty = true;
      sharedTasksState.lastActionAt = new Date().toISOString();
      publishSharedTasksStatus("dirty");
    }
    if (!config.enabled) {
      writeArrayToStorage(STORAGE_KEYS.tasks, tasksCache);
    }
    publishTasksUpdate("local");
    queueSharedTasksUpload(tasksCache);
  }

  async function refreshTasksFromShareDrive(reason) {
    const config = getSharedTasksConfig();
    if (!config.enabled) return null;
    if (sharedTasksState.inFlight) return null;
    if (sharedTasksDirty) return null;

    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());

    if (!config.downloadUrl && !config.shareUrl) {
      sharedTasksState.lastError = "Sharedrive tasks share URL is missing.";
      return null;
    }

    sharedTasksState.inFlight = true;
    sharedTasksState.lastActionAt = new Date().toISOString();
    publishSharedTasksStatus("syncing");

    try {
      let downloadUrl = config.downloadUrl;
      if (!downloadUrl && canUseDesktop) {
        const result = await global.PlutusDesktop.getShareDriveDownloadUrl({
          shareUrl: config.shareUrl,
          accessToken: config.accessToken,
        });
        if (!result || !result.ok || !result.data || !result.data.downloadUrl) {
          throw new Error((result && result.error) || "Failed to resolve sharedrive download URL.");
        }
        downloadUrl = result.data.downloadUrl;
      }

      let payload = null;
      try {
        if (downloadUrl) {
          payload = await fetchJson(downloadUrl, { cache: "no-store" });
        } else if (canUseBrowser) {
          const result = await downloadShareDriveFileBrowser(config.shareUrl);
          payload = parseJsonText(result && result.text ? result.text : "");
          if (!payload) {
            throw new Error("Sharedrive tasks file is not valid JSON.");
          }
        } else {
          throw new Error("Sharedrive download unavailable.");
        }
      } catch (error) {
        if (canUseDesktop && config.shareUrl && global.PlutusDesktop && typeof global.PlutusDesktop.downloadShareDriveFile === "function") {
          const result = await global.PlutusDesktop.downloadShareDriveFile({
            shareUrl: config.shareUrl,
            accessToken: config.accessToken,
          });
          if (!result || !result.ok || !result.data || typeof result.data.text !== "string") {
            throw new Error((result && result.error) || "Sharedrive download failed.");
          }
          payload = parseJsonText(result.data.text);
          if (!payload) {
            throw new Error("Sharedrive tasks file is not valid JSON.");
          }
        } else if (canUseBrowser && config.shareUrl) {
          const result = await downloadShareDriveFileBrowser(config.shareUrl);
          payload = parseJsonText(result && result.text ? result.text : "");
          if (!payload) {
            throw new Error("Sharedrive tasks file is not valid JSON.");
          }
        } else {
          throw error;
        }
      }
      const parsed = parseSharedTasksPayload(payload);
      if (Array.isArray(parsed.tasks)) {
        tasksCache = cloneArray(parsed.tasks);
        if (!config.enabled) {
          writeArrayToStorage(STORAGE_KEYS.tasks, tasksCache);
        }
        sharedTasksState.lastRemoteUpdatedAt = parsed.updatedAt || "";
        sharedTasksState.lastSyncAt = new Date().toISOString();
        sharedTasksState.lastError = "";
        publishSharedTasksStatus("synced");
        publishTasksUpdate(reason || "sharedrive");
      }
    } catch (error) {
      sharedTasksState.lastError = error instanceof Error ? error.message : "Sharedrive sync failed.";
      console.warn("[AppCore] Sharedrive tasks refresh failed", error);
      publishSharedTasksStatus("sync_failed");
    } finally {
      sharedTasksState.inFlight = false;
    }

    return null;
  }

  function scheduleSharedTasksUpload(tasks) {
    if (sharedTasksUploadTimer) {
      clearTimeout(sharedTasksUploadTimer);
      sharedTasksUploadTimer = null;
    }

    sharedTasksUploadPending = cloneArray(tasks);
    publishSharedTasksStatus("upload_queued");
    sharedTasksUploadTimer = setTimeout(() => {
      sharedTasksUploadTimer = null;
      performSharedTasksUpload(sharedTasksUploadPending);
      sharedTasksUploadPending = null;
    }, 800);
  }

  async function performSharedTasksUpload(tasks) {
    const config = getSharedTasksConfig();
    if (!config.enabled) return;
    if (sharedTasksState.uploadInFlight) return;
    if (!config.shareUrl) {
      sharedTasksState.lastError = "Sharedrive tasks share URL is missing.";
      return;
    }

    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());
    if (!canUseDesktop && !canUseBrowser) {
      sharedTasksState.lastError = "Sharedrive sync unavailable (sign in required).";
      return;
    }

    sharedTasksState.uploadInFlight = true;
    sharedTasksState.lastActionAt = new Date().toISOString();
    publishSharedTasksStatus("uploading");

    try {
      const payload = buildSharedTasksPayload(tasks, config);
      const content = JSON.stringify(payload, null, 2);
      const fileName = config.fileName || SHARED_TASKS_DEFAULTS.fileName;

      if (canUseDesktop) {
        const base64 = toBase64(content);
        const result = await global.PlutusDesktop.uploadShareDriveFile({
          shareUrl: config.shareUrl,
          accessToken: config.accessToken,
          parentItemId: config.parentItemId,
          fileName,
          contentBase64: base64,
          conflictBehavior: "replace",
        });
        if (!result || !result.ok) {
          throw new Error((result && result.error) || "Sharedrive upload failed.");
        }
      } else {
        const base64 = toBase64(content);
        await uploadShareDriveFileBrowser({
          shareUrl: config.shareUrl,
          parentItemId: config.parentItemId,
          fileName,
          contentBase64: base64,
        });
      }

      sharedTasksState.lastSyncAt = new Date().toISOString();
      sharedTasksState.lastUploadAt = sharedTasksState.lastSyncAt;
      sharedTasksState.lastError = "";
      sharedTasksDirty = false;
      publishSharedTasksStatus("uploaded");
    } catch (error) {
      sharedTasksState.lastError = error instanceof Error ? error.message : "Sharedrive upload failed.";
      console.warn("[AppCore] Sharedrive tasks upload failed", error);
      publishSharedTasksStatus("upload_failed");
    } finally {
      sharedTasksState.uploadInFlight = false;
    }
  }

  function queueSharedTasksUpload(tasks) {
    const config = getSharedTasksConfig();
    if (!config.enabled) return;
    scheduleSharedTasksUpload(tasks);
  }

  function toBase64(text) {
    if (typeof global.TextEncoder === "function") {
      const encoder = new TextEncoder();
      const bytes = encoder.encode(text);
      let binary = "";
      bytes.forEach((b) => {
        binary += String.fromCharCode(b);
      });
      return btoa(binary);
    }

    return btoa(unescape(encodeURIComponent(text)));
  }

  function findDealForTask(deals, task) {
    if (!Array.isArray(deals) || !task) return null;
    const rawDealId = task.dealId ?? task.deal ?? task.dealName ?? "";
    const dealKey = normalizeValue(rawDealId);
    if (!dealKey) return null;

    return deals.find((deal) => {
      const id = normalizeValue(deal.id);
      const name = normalizeValue(deal.name);
      const company = normalizeValue(deal.company);
      const dashboardId = normalizeValue(deal.fundraisingDashboardId);
      return id === dealKey || name === dealKey || company === dealKey || dashboardId === dealKey;
    }) || null;
  }

  function isAutoTask(task) {
    const title = normalizeValue(task && task.title);
    const metaSource = normalizeValue(task && task.metaSource);
    const taskId = normalizeValue(task && task.id);
    return (
      title.startsWith("[auto]") ||
      metaSource === "dashboard-contact-status" ||
      taskId.startsWith(`${AUTO_CONTACT_TASK_PREFIX}-`)
    );
  }

  global.AppCore = {
    STORAGE_KEYS,
    AUTO_CONTACT_TASK_PREFIX,
    normalizeValue,
    loadDealsData,
    saveDealsData,
    loadTasksData,
    saveTasksData,
    refreshTasksFromShareDrive,
    requestGraphDeviceCode: requestDeviceCodeBrowser,
    pollGraphDeviceCode: pollDeviceCodeBrowser,
    getSharedTasksStatus: () => ({
      stage: "idle",
      dirty: sharedTasksDirty,
      started: sharedTasksState.started,
      inFlight: sharedTasksState.inFlight,
      uploadInFlight: sharedTasksState.uploadInFlight,
      lastSyncAt: sharedTasksState.lastSyncAt,
      lastUploadAt: sharedTasksState.lastUploadAt,
      lastRemoteUpdatedAt: sharedTasksState.lastRemoteUpdatedAt,
      lastActionAt: sharedTasksState.lastActionAt,
      lastError: sharedTasksState.lastError,
    }),
    findDealForTask,
    isAutoTask,
  };
})(window);
