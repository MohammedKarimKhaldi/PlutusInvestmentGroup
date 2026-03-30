(function initAppCore(global) {
  const APP_CONFIG = global.PlutusAppConfig || {};
  const STORAGE_KEYS = Object.assign({
    deals: "deals_data_v1",
    tasks: "owner_tasks_v1",
    graphSession: "plutus_graph_session_v1",
    dashboardConfig: "plutus_dashboard_config_v1",
  }, APP_CONFIG.storageKeys || {});
  const DATA_FILES = Object.assign({
    config: "config.json",
    deals: "deals.json",
    tasks: "tasks.json",
    sharedTasks: "sharedrive-tasks.json",
  }, APP_CONFIG.dataFiles || {});

  const AUTO_CONTACT_TASK_PREFIX = "auto-contact-status";
  const AUTO_DEAL_READINESS_TASK_PREFIX = "auto-deal-readiness";
  const SHARED_TASKS_DEFAULTS = {
    enabled: false,
    shareUrl: "",
    fileName: DATA_FILES.sharedTasks,
    pollIntervalMs: 60000,
    accessToken: "",
    downloadUrl: "",
    parentItemId: "",
    writerLabel: "",
    azureClientId: "",
    azureTenantId: "common",
    graphScopes: "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read Mail.Read",
  };
  const GRAPH_BROWSER_SESSION_KEY = STORAGE_KEYS.graphSession;
  const GRAPH_UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024;
  const DASHBOARD_CONFIG_KEY = STORAGE_KEYS.dashboardConfig;
  const NATIVE_HTTP_PLUGIN = "CapacitorHttp";
  const SHARED_FILE_DEFAULTS = {
    enabled: false,
    shareUrl: "",
    fileName: "",
    pollIntervalMs: 60000,
    downloadUrl: "",
    parentItemId: "",
    accessToken: "",
  };

  let nativeHttpClient = null;
  let nativeHttpChecked = false;

  let tasksCache = null;
  let dealsCache = null;
  let dashboardConfigRefreshPromise = null;
  let sharedTasksSyncTimer = null;
  let sharedDealsSyncTimer = null;
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

  const sharedDealsState = {
    started: false,
    inFlight: false,
    lastSyncAt: "",
    lastActionAt: "",
    lastRemoteUpdatedAt: "",
    lastError: "",
  };

  function normalizeValue(value) {
    return String(value || "").trim().toLowerCase();
  }

  function parseDealAmount(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const raw = String(value == null ? "" : value).trim();
    if (!raw) return 0;
    const cleaned = raw.replace(/,/g, "").replace(/[^0-9.-]/g, "");
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function getDealRetainerRawValue(deal) {
    if (!deal || typeof deal !== "object") return "";
    if (deal.retainerMonthly != null && String(deal.retainerMonthly).trim()) return String(deal.retainerMonthly).trim();
    if (deal.Retainer != null && String(deal.Retainer).trim()) return String(deal.Retainer).trim();
    return "";
  }

  function getDealRetainerState(deal) {
    const rawValue = getDealRetainerRawValue(deal);
    const amount = parseDealAmount(rawValue);
    const hasRetainer = amount > 0;
    return {
      rawValue,
      amount,
      hasRetainer,
      bucket: hasRetainer ? "with-retainer" : "no-retainer",
      label: hasRetainer ? "With retainer" : "0 / no retainer",
    };
  }

  function hasPositiveRetainer(deal) {
    return getDealRetainerState(deal).hasRetainer;
  }

  function matchesDealRetainerFilter(deal, filterValue) {
    const normalizedFilter = normalizeValue(filterValue);
    if (!normalizedFilter || normalizedFilter === "all") return true;
    const bucket = getDealRetainerState(deal).bucket;
    return normalizedFilter === bucket;
  }

  function compareDealsByRetainerState(left, right) {
    const leftState = getDealRetainerState(left);
    const rightState = getDealRetainerState(right);
    if (leftState.hasRetainer === rightState.hasRetainer) return 0;
    return leftState.hasRetainer ? -1 : 1;
  }

  function sortDealsByRetainerState(deals, fallbackComparator) {
    return (Array.isArray(deals) ? deals.slice() : []).sort((left, right) => {
      const retainerOrder = compareDealsByRetainerState(left, right);
      if (retainerOrder) return retainerOrder;
      return typeof fallbackComparator === "function" ? fallbackComparator(left, right) : 0;
    });
  }

  function getPageUrl(pageId, params) {
    if (APP_CONFIG && typeof APP_CONFIG.buildPageHref === "function") {
      return APP_CONFIG.buildPageHref(pageId, params);
    }
    const query = new URLSearchParams();
    Object.entries(params || {}).forEach(([key, value]) => {
      if (value == null || value === "") return;
      query.set(key, String(value));
    });
    const queryString = query.toString();
    return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
  }

  function testAzureClientId() {
    try {
      const config = getSharedTasksConfig();
      console.log("📋 Azure Client ID Test Results:");
      console.log("✅ Config loaded successfully");
      console.log(`📝 Azure Client ID: ${config.azureClientId}`);
      console.log(`📄 Tenant ID: ${config.azureTenantId}`);
      console.log(`🔗 Share URL: ${config.shareUrl}`);
      console.log(`📄 File Name: ${config.fileName}`);
      
      if (config.azureClientId) {
        console.log("✅ Azure client ID is available");
        return { ok: true, config };
      } else {
        console.log("❌ Azure client ID is missing");
        return { ok: false, error: "Azure client ID not found" };
      }
    } catch (error) {
      console.error("❌ Error loading config:", error);
      return { ok: false, error: error.message };
    }
  }

  function writeJsonToStorage(key, payload) {
    try {
      localStorage.setItem(key, JSON.stringify(payload || {}));
      return true;
    } catch {
      return false;
    }
  }

  function readJsonFromStorage(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" ? parsed : null;
    } catch {
      return null;
    }
  }

  function readDataJson(key) {
    if (global.PlutusDesktop && typeof global.PlutusDesktop.readDataJson === "function") {
      try {
        const result = global.PlutusDesktop.readDataJson(key);
        if (result && result.ok) return result.data;
      } catch {
        // ignore desktop read failures
      }
    }
    return null;
  }

  function writeDataJson(key, payload) {
    if (global.PlutusDesktop && typeof global.PlutusDesktop.writeDataJson === "function") {
      try {
        const result = global.PlutusDesktop.writeDataJson(key, payload);
        return !!(result && result.ok);
      } catch {
        return false;
      }
    }
    return false;
  }

  function readDashboardOverrides() {
    const desktop = readDataJson("config");
    if (desktop && typeof desktop === "object") {
      return desktop;
    }
    return readJsonFromStorage(DASHBOARD_CONFIG_KEY);
  }

  function mergeDashboardConfig(baseConfig) {
    const base = baseConfig && typeof baseConfig === "object" ? baseConfig : {};
    const overrides = readDashboardOverrides() || {};
    const remoteConfigEnabled = getSharedDashboardConfigSyncConfig().enabled;
    const hasRemoteOverrides = Boolean(
      remoteConfigEnabled &&
      overrides &&
      typeof overrides === "object" &&
      (
        (Array.isArray(overrides.dashboards) && overrides.dashboards.length) ||
        (overrides.settings && typeof overrides.settings === "object") ||
        (Array.isArray(overrides.proxies) && overrides.proxies.length)
      ),
    );

    if (hasRemoteOverrides) {
      return {
        dashboards: Array.isArray(overrides.dashboards) ? cloneArray(overrides.dashboards) : [],
        settings: Object.assign({}, overrides.settings || {}),
        proxies: Array.isArray(overrides.proxies) ? overrides.proxies.slice() : (base.proxies || []),
      };
    }

    const baseDashboards = Array.isArray(base.dashboards) ? cloneArray(base.dashboards) : [];
    const overrideDashboards = Array.isArray(overrides.dashboards) ? overrides.dashboards : [];

    const mergedDashboards = baseDashboards.slice();
    overrideDashboards.forEach((entry) => {
      if (!entry || !entry.id) return;
      const idx = mergedDashboards.findIndex(
        (dashboard) => normalizeValue(dashboard.id) === normalizeValue(entry.id),
      );
      if (idx >= 0) {
        const merged = Object.assign({}, mergedDashboards[idx], entry);
        if (mergedDashboards[idx].sheets && entry.sheets) {
          merged.sheets = Object.assign({}, mergedDashboards[idx].sheets, entry.sheets);
        }
        mergedDashboards[idx] = merged;
      } else {
        mergedDashboards.push(entry);
      }
    });

    const settings = Object.assign({}, base.settings || {}, overrides.settings || {});
    return { dashboards: mergedDashboards, settings, proxies: overrides.proxies || base.proxies || [] };
  }

  async function upsertDashboardConfigEntry(entry, options) {
    if (!entry || !entry.id) throw new Error("Dashboard id is required.");
    const remoteConfig = getSharedDashboardConfigSyncConfig();
    if (remoteConfig.enabled) {
      await refreshDashboardConfigFromShareDrive();
    }

    const overrides = readDashboardOverrides() || {};
    const dashboards = Array.isArray(overrides.dashboards) ? overrides.dashboards.slice() : [];
    const idx = dashboards.findIndex(
      (dashboard) => normalizeValue(dashboard.id) === normalizeValue(entry.id),
    );
    if (idx >= 0) {
      dashboards[idx] = Object.assign({}, dashboards[idx], entry);
    } else {
      dashboards.push(entry);
    }

    const settings = Object.assign({}, overrides.settings || {});
    if (options && options.makeDefault) {
      settings.defaultDashboard = entry.id;
    }

    const baseConfig = global.DASHBOARD_CONFIG || { dashboards: [], settings: {}, proxies: [] };
    const payload = {
      dashboards,
      settings,
      proxies: overrides.proxies || baseConfig.proxies || [],
    };
    if (!writeDataJson("config", payload)) {
      writeJsonToStorage(DASHBOARD_CONFIG_KEY, payload);
    }
    const mergedConfig = mergeDashboardConfig(baseConfig);
    if (!remoteConfig.enabled) {
      return Promise.resolve(mergedConfig);
    }

    return uploadJsonToShareDrive(remoteConfig, mergedConfig).then(() => {
      return mergedConfig;
    });
  }

  async function refreshDashboardConfigFromShareDrive() {
    if (dashboardConfigRefreshPromise) return dashboardConfigRefreshPromise;

    dashboardConfigRefreshPromise = (async () => {
      const remoteConfig = getSharedDashboardConfigSyncConfig();
      if (!remoteConfig.enabled) return mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} });

      try {
        const payload = await downloadJsonFromShareDrive(remoteConfig);
        if (payload && !Array.isArray(payload) && typeof payload === "object") {
          if (!writeDataJson("config", payload)) {
            writeJsonToStorage(DASHBOARD_CONFIG_KEY, payload);
          }
          const merged = mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} });
          try {
            global.dispatchEvent(new CustomEvent("appcore:dashboard-config-updated", {
              detail: { config: merged, source: "sharedrive" },
            }));
          } catch {
            // ignore
          }
          return merged;
        }
      } catch (error) {
        console.warn("[AppCore] Shared dashboard config refresh failed", error);
      }

      return mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} });
    })();

    try {
      return await dashboardConfigRefreshPromise;
    } finally {
      dashboardConfigRefreshPromise = null;
    }
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

  function getHeaderValue(headers, key) {
    if (!headers || typeof headers !== "object") return "";
    const matchKey = Object.keys(headers).find((headerKey) => String(headerKey || "").toLowerCase() === key.toLowerCase());
    return matchKey ? String(headers[matchKey] || "") : "";
  }

  function base64ToArrayBuffer(base64) {
    const binary = atob(String(base64 || "").trim());
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes.buffer;
  }

  function normalizeNativeHttpBody(body, headers) {
    if (body == null) return undefined;
    if (body instanceof URLSearchParams) {
      return Object.fromEntries(body.entries());
    }

    const contentType = getHeaderValue(headers, "Content-Type").toLowerCase();
    if (typeof body === "string") {
      if (contentType.includes("application/json")) {
        try {
          return JSON.parse(body);
        } catch {
          return body;
        }
      }
      if (contentType.includes("application/x-www-form-urlencoded")) {
        return Object.fromEntries(new URLSearchParams(body).entries());
      }
      return body;
    }

    return body;
  }

  async function sendNativeHttpRequest(url, options, responseType) {
    const nativeHttp = getNativeHttpClient();
    if (!nativeHttp) return null;

    const normalizedOptions = options || {};
    const request = {
      url,
      method: String(normalizedOptions.method || "GET").toUpperCase(),
      headers: normalizedOptions.headers || {},
      responseType: responseType || "json",
    };
    const data = normalizeNativeHttpBody(normalizedOptions.body, request.headers);
    if (typeof data !== "undefined") {
      request.data = data;
    }
    if (normalizedOptions.dataType) {
      request.dataType = normalizedOptions.dataType;
    }

    try {
      if (nativeHttp.type === "plugin" && nativeHttp.client && typeof nativeHttp.client.request === "function") {
        const response = await nativeHttp.client.request(request);
        return normalizeHttpResponse(response);
      }
      if (nativeHttp.type === "nativePromise" && nativeHttp.client) {
        const response = await nativeHttp.client.nativePromise("CapacitorHttp", "request", request);
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

    return null;
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
    const nativeResponse = await sendNativeHttpRequest(url, {
      method: "POST",
      headers,
      body: formParams,
    }, "json");

    if (nativeResponse) {
      return nativeResponse;
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

  function hasOutlookBridge() {
    return Boolean(
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.listOutlookMessages === "function",
    );
  }

  function hasShareDriveListBridge() {
    return Boolean(
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.listShareDriveChildren === "function",
    );
  }

  function getNestedSharedriveSection(sectionName) {
    const sources = [
      global.SHAREDRIVE_TASKS,
      readDataJson("sharedrive-tasks"),
      readJsonFromStorage("sharedrive-tasks"),
    ];

    for (const source of sources) {
      if (!source || typeof source !== "object") continue;
      if (sectionName && source[sectionName] && typeof source[sectionName] === "object") {
        return source[sectionName];
      }
      if (!sectionName) {
        return source;
      }
    }

    return {};
  }

  function normalizeSharedFileConfig(raw, defaults) {
    const config = Object.assign({}, SHARED_FILE_DEFAULTS, defaults || {}, raw || {});
    config.enabled = Boolean(raw && raw.enabled);
    config.shareUrl = String(config.shareUrl || "").trim();
    config.fileName = String(config.fileName || "").trim();
    config.downloadUrl = String(config.downloadUrl || "").trim();
    config.parentItemId = String(config.parentItemId || "").trim();
    config.accessToken = String(config.accessToken || "").trim();
    config.explicitShareUrl = Boolean(raw && raw.shareUrl);

    const interval = Number(config.pollIntervalMs);
    config.pollIntervalMs = Number.isFinite(interval) ? interval : SHARED_FILE_DEFAULTS.pollIntervalMs;

    return config;
  }

  function getSharedTasksConfig() {
    let raw =
      global.SHAREDRIVE_TASKS ||
      (global.DASHBOARD_CONFIG && global.DASHBOARD_CONFIG.settings && global.DASHBOARD_CONFIG.settings.sharedTasks) ||
      readJsonFromStorage("sharedrive-tasks") ||
      {};

    // Handle nested format if present
    if (raw && raw.tasks && typeof raw.tasks === "object") {
      raw = raw.tasks;
    }

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

  function getSharedDealsConfig() {
    const settingsConfig =
      (global.DASHBOARD_CONFIG && global.DASHBOARD_CONFIG.settings && global.DASHBOARD_CONFIG.settings.sharedDeals) ||
      {};
    const nestedConfig = getNestedSharedriveSection("deals");
    const raw = Object.assign({}, nestedConfig || {}, settingsConfig || {});
    const config = normalizeSharedFileConfig(raw, {
      fileName: DATA_FILES.deals,
    });
    config.enabled = Boolean(
      (settingsConfig && settingsConfig.enabled) ||
      (nestedConfig && nestedConfig.enabled),
    );
    return config;
  }

  function getSharedDashboardConfigSyncConfig() {
    const nestedConfig = getNestedSharedriveSection("config");
    const dealsConfig = getSharedDealsConfig();
    const raw = Object.assign({}, nestedConfig || {});
    const hasExplicitShareUrl = Boolean(raw.shareUrl);
    if (!raw.shareUrl && dealsConfig.shareUrl) {
      raw.shareUrl = dealsConfig.shareUrl;
    }
    if (!raw.parentItemId && dealsConfig.parentItemId) {
      raw.parentItemId = dealsConfig.parentItemId;
    }

    const config = normalizeSharedFileConfig(raw, {
      fileName: DATA_FILES.config,
    });
    config.explicitShareUrl = hasExplicitShareUrl;
    config.enabled = Boolean(
      (nestedConfig && nestedConfig.enabled) ||
      (dealsConfig.enabled && config.shareUrl),
    );
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

  async function readDesktopGraphSession() {
    if (!hasDesktopBridge() || !global.PlutusDesktop || typeof global.PlutusDesktop.getGraphSession !== "function") {
      return null;
    }

    try {
      const result = await global.PlutusDesktop.getGraphSession();
      if (!result || !result.ok) return null;
      const data = result.data;
      return data && typeof data === "object" ? data : null;
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

  async function resolveRendererGraphAccessToken() {
    const config = getSharedTasksConfig();
    if (config.accessToken) return config.accessToken;

    const desktopSession = await readDesktopGraphSession();
    if (
      desktopSession &&
      desktopSession.accessToken &&
      desktopSession.expiresAt &&
      Date.now() < Number(desktopSession.expiresAt) - 60000
    ) {
      return desktopSession.accessToken;
    }

    try {
      return await resolveBrowserAccessToken();
    } catch (error) {
      if (desktopSession && desktopSession.accessToken) {
        return desktopSession.accessToken;
      }
      throw error;
    }
  }

  async function graphFetchJson(url, token, options) {
    const requestOptions = {
      ...(options || {}),
      headers: {
        ...(options && options.headers ? options.headers : {}),
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    };

    const nativeResponse = await sendNativeHttpRequest(url, requestOptions, "json");
    if (nativeResponse) {
      const payload = nativeResponse.data;
      if (!nativeResponse.ok) {
        const message =
          (payload && payload.error && (payload.error.message || payload.error_description)) ||
          nativeResponse.error ||
          "Request failed";
        throw new Error(message);
      }
      return payload;
    }

    const response = await fetch(url, {
      ...requestOptions,
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

  function normalizeGraphEmailAddress(entry) {
    const emailAddress = entry && entry.emailAddress && typeof entry.emailAddress === "object"
      ? entry.emailAddress
      : {};
    return {
      name: String(emailAddress.name || "").trim(),
      address: String(emailAddress.address || "").trim(),
    };
  }

  function normalizeOutlookRecipients(recipients) {
    return (Array.isArray(recipients) ? recipients : [])
      .map((entry) => normalizeGraphEmailAddress(entry))
      .filter((entry) => entry.address);
  }

  function normalizeOutlookMessage(message) {
    return {
      id: String(message && message.id || "").trim(),
      subject: String(message && message.subject || "").trim(),
      bodyPreview: String(message && message.bodyPreview || "").trim(),
      receivedDateTime: String(message && message.receivedDateTime || "").trim(),
      webLink: String(message && message.webLink || "").trim(),
      isRead: Boolean(message && message.isRead),
      from: normalizeGraphEmailAddress(message && message.from),
      toRecipients: normalizeOutlookRecipients(message && message.toRecipients),
      ccRecipients: normalizeOutlookRecipients(message && message.ccRecipients),
    };
  }

  function tokenHasScope(token, requiredScope) {
    const claims = decodeJwtPayload(token);
    const scopes = String(claims && claims.scp || "")
      .split(/\s+/)
      .map((entry) => normalizeValue(entry))
      .filter(Boolean);
    return scopes.includes(normalizeValue(requiredScope));
  }

  async function listOutlookMessagesBrowser(payload) {
    const config = payload && typeof payload === "object" ? payload : {};
    let token = await resolveRendererGraphAccessToken();
    if (!tokenHasScope(token, "Mail.Read")) {
      const session = readBrowserGraphSession();
      if (session && session.refreshToken) {
        try {
          token = await refreshBrowserAccessToken(session.refreshToken);
        } catch {
          // Ignore refresh failure and surface the clearer message below.
        }
      }
    }
    if (!tokenHasScope(token, "Mail.Read")) {
      throw new Error("Microsoft sign-in is missing Mail.Read. Click 'Sign in with Microsoft' again and approve Outlook inbox access.");
    }
    const requestedTop = Number(config.top || 50);
    const safeTop = Number.isFinite(requestedTop)
      ? Math.min(Math.max(Math.floor(requestedTop), 1), 100)
      : 50;
    const searchTerm = String(config.search || "").trim();
    const params = new URLSearchParams({
      $top: String(safeTop),
      $select: [
        "id",
        "subject",
        "bodyPreview",
        "receivedDateTime",
        "webLink",
        "isRead",
        "from",
        "toRecipients",
        "ccRecipients",
      ].join(","),
    });
    if (searchTerm) {
      params.set("$search", `"${searchTerm.replace(/"/g, '\\"')}"`);
    } else {
      params.set("$orderby", "receivedDateTime DESC");
    }
    const url = `https://graph.microsoft.com/v1.0/me/messages?${params.toString()}`;
    const response = await graphFetchJson(url, token, {
      headers: searchTerm ? { ConsistencyLevel: "eventual" } : {},
    });
    return {
      items: (Array.isArray(response && response.value) ? response.value : [])
        .map((message) => normalizeOutlookMessage(message)),
      fetchedAt: new Date().toISOString(),
    };
  }

  async function listOutlookMessages(payload) {
    if (hasOutlookBridge()) {
      const result = await global.PlutusDesktop.listOutlookMessages(payload || {});
      if (!result || !result.ok) {
        throw new Error((result && result.error) || "Failed to load Outlook messages.");
      }
      return result.data || { items: [], fetchedAt: "" };
    }
    return listOutlookMessagesBrowser(payload);
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

  async function listDriveItemChildrenBrowser(driveId, itemId, token) {
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
      const page = await graphFetchJson(nextUrl, token);
      const values = Array.isArray(page && page.value) ? page.value : [];
      values.forEach((item) => items.push(item));
      nextUrl = page && typeof page["@odata.nextLink"] === "string" ? page["@odata.nextLink"] : "";
    }

    return items;
  }

  function normalizeDriveItemBrowser(item) {
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

  async function listShareDriveChildrenBrowser({ shareUrl, parentItemId }) {
    const token = await resolveBrowserAccessToken();
    const rootItem = await getShareDriveItemBrowser(shareUrl, token);
    const driveId =
      (rootItem && rootItem.parentReference && rootItem.parentReference.driveId) ||
      (rootItem &&
        rootItem.remoteItem &&
        rootItem.remoteItem.parentReference &&
        rootItem.remoteItem.parentReference.driveId) ||
      "";
    if (!driveId) {
      throw new Error("Unable to resolve drive for SharePoint item.");
    }

    const targetId = parentItemId || (rootItem && rootItem.id ? rootItem.id : "");
    if (!targetId) {
      throw new Error("Unable to resolve folder id for SharePoint item.");
    }

    const children = await listDriveItemChildrenBrowser(driveId, targetId, token);
    return {
      root: normalizeDriveItemBrowser(rootItem),
      driveId,
      parentItemId: targetId,
      items: children.map(normalizeDriveItemBrowser),
      fetchedAt: new Date().toISOString(),
    };
  }

  async function fetchText(url, options) {
    const nativeResponse = await sendNativeHttpRequest(url, options || {}, "text");
    if (nativeResponse) {
      if (!nativeResponse.ok) {
        throw new Error(nativeResponse.error || `Download failed (${nativeResponse.status || 0})`);
      }
      if (nativeResponse.data && typeof nativeResponse.data === "object") {
        return JSON.stringify(nativeResponse.data, null, 2);
      }
      return String(nativeResponse.data || "");
    }

    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`Download failed (${response.status})`);
    }
    return response.text();
  }

  async function resolveShareDriveDownloadUrlBrowser(shareUrl) {
    const token = await resolveBrowserAccessToken();
    const item = await getShareDriveItemBrowser(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
      "@microsoft.graph.downloadUrl",
    ]);
    const downloadUrl = item && item["@microsoft.graph.downloadUrl"] ? item["@microsoft.graph.downloadUrl"] : "";
    if (!downloadUrl) {
      throw new Error("Unable to resolve sharedrive download URL.");
    }
    return downloadUrl;
  }

  async function downloadShareDriveFileBrowser(shareUrl) {
    const token = await resolveBrowserAccessToken();
    const item = await getShareDriveItemBrowser(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
      "@microsoft.graph.downloadUrl",
    ]);
    const directDownloadUrl = item && item["@microsoft.graph.downloadUrl"] ? item["@microsoft.graph.downloadUrl"] : "";
    if (directDownloadUrl) {
      const text = await fetchText(directDownloadUrl, { cache: "no-store" });
      if (!looksLikeHtmlPayload(text)) {
        return { text, driveId: "", itemId: item && item.id ? item.id : "" };
      }
    }
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
    const nativeResponse = await sendNativeHttpRequest(url, {
      headers: { Authorization: `Bearer ${token}` },
    }, "text");
    if (nativeResponse) {
      if (!nativeResponse.ok) {
        throw new Error(nativeResponse.error || "Download failed.");
      }
      if (typeof nativeResponse.data === "string") {
        return { text: nativeResponse.data, driveId, itemId };
      }
      return { text: JSON.stringify(nativeResponse.data || {}), driveId, itemId };
    }

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

  function bytesToBase64(bytes) {
    let binary = "";
    const chunkSize = 0x8000;
    for (let start = 0; start < bytes.length; start += chunkSize) {
      const slice = bytes.subarray(start, Math.min(start + chunkSize, bytes.length));
      binary += String.fromCharCode.apply(null, slice);
    }
    return btoa(binary);
  }

  async function uploadWithSessionBrowser(uploadUrl, bytes) {
    const total = bytes.length;
    let start = 0;
    while (start < total) {
      const end = Math.min(start + GRAPH_UPLOAD_CHUNK_SIZE, total);
      const chunk = bytes.slice(start, end);
      const nativeResponse = await sendNativeHttpRequest(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Type": "application/octet-stream",
          "Content-Length": String(chunk.length),
          "Content-Range": `bytes ${start}-${end - 1}/${total}`,
        },
        body: bytesToBase64(chunk),
        dataType: "file",
      }, "text");
      if (nativeResponse) {
        if (!nativeResponse.ok) {
          throw new Error(`Upload failed: ${nativeResponse.error || "Unknown error"}`);
        }
        const text = String(nativeResponse.data || "");
        if (end === total) {
          try {
            return text ? JSON.parse(text) : null;
          } catch {
            return null;
          }
        }
        start = end;
        continue;
      }

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

  async function resolveShareDriveFileBrowser({ shareUrl, parentItemId, fileName }) {
    const token = await resolveBrowserAccessToken();
    const item = await getShareDriveItemBrowser(shareUrl, token, [
      "id",
      "name",
      "parentReference",
      "remoteItem",
      "folder",
      "file",
      "@microsoft.graph.downloadUrl",
    ]);
    const driveId =
      (item && item.parentReference && item.parentReference.driveId) ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.driveId) ||
      "";
    const resolvedParentItemId =
      (item && item.parentReference && item.parentReference.id) ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.id) ||
      "";
    if (!driveId) throw new Error("Unable to resolve SharePoint drive ID.");

    const normalizedFileName = normalizeValue(fileName);
    if (!normalizedFileName || (item && item.file && normalizeValue(item.name) === normalizedFileName)) {
      return {
        token,
        driveId,
        itemId: item && item.id ? item.id : "",
        name: item && item.name ? item.name : "",
        parentItemId: resolvedParentItemId,
        downloadUrl: item && item["@microsoft.graph.downloadUrl"] ? item["@microsoft.graph.downloadUrl"] : "",
      };
    }

    const targetParentId =
      String(parentItemId || "").trim() ||
      (item && item.folder && item.id ? item.id : "") ||
      (item && item.parentReference && item.parentReference.id ? item.parentReference.id : "") ||
      (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.id
        ? item.remoteItem.parentReference.id
        : "");
    if (!targetParentId) {
      throw new Error("Unable to resolve target folder for SharePoint file.");
    }

    const children = await listDriveItemChildrenBrowser(driveId, targetParentId, token);
    const child = children.find((entry) => normalizeValue(entry && entry.name) === normalizedFileName);
    if (!child || !child.id) {
      throw new Error(`Sharedrive file not found: ${fileName}`);
    }

    return {
      token,
      driveId,
      itemId: child.id,
      name: child && child.name ? child.name : "",
      parentItemId: targetParentId,
      downloadUrl: child && child["@microsoft.graph.downloadUrl"] ? child["@microsoft.graph.downloadUrl"] : "",
    };
  }

  async function downloadShareDriveJsonBrowser({ shareUrl, parentItemId, fileName }) {
    const resolved = await resolveShareDriveFileBrowser({ shareUrl, parentItemId, fileName });
    if (resolved && resolved.downloadUrl) {
      const directText = await fetchText(resolved.downloadUrl, { cache: "no-store" });
      const directPayload = parseJsonText(directText);
      if (directPayload !== null) {
        return directPayload;
      }
    }
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
      resolved.driveId,
    )}/items/${encodeURIComponent(resolved.itemId)}/content`;
    const nativeResponse = await sendNativeHttpRequest(url, {
      headers: { Authorization: `Bearer ${resolved.token}` },
    }, "text");
    if (nativeResponse) {
      if (!nativeResponse.ok) {
        throw new Error(nativeResponse.error || "Download failed.");
      }
      if (nativeResponse.data && typeof nativeResponse.data === "object") {
        return nativeResponse.data;
      }
      return parseJsonText(String(nativeResponse.data || ""));
    }

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${resolved.token}` },
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error("Download failed.");
    }
    return parseJsonText(text);
  }

  async function downloadShareDriveTextBrowser({ shareUrl, parentItemId, fileName }) {
    const resolved = await resolveShareDriveFileBrowser({ shareUrl, parentItemId, fileName });
    if (resolved && resolved.downloadUrl) {
      const directText = await fetchText(resolved.downloadUrl, { cache: "no-store" });
      if (!looksLikeHtmlPayload(directText)) {
        return directText;
      }
    }
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
      resolved.driveId,
    )}/items/${encodeURIComponent(resolved.itemId)}/content`;
    const nativeResponse = await sendNativeHttpRequest(url, {
      headers: { Authorization: `Bearer ${resolved.token}` },
    }, "text");
    if (nativeResponse) {
      if (!nativeResponse.ok) {
        throw new Error(nativeResponse.error || "Download failed.");
      }
      if (typeof nativeResponse.data === "string") return nativeResponse.data;
      return JSON.stringify(nativeResponse.data, null, 2);
    }

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${resolved.token}` },
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error("Download failed.");
    }
    return text;
  }

  async function fetchJson(url, options) {
    const text = await fetchText(url, options);
    if (!text) return null;
    const parsed = parseJsonText(text);
    if (parsed !== null) return parsed;
    throw new Error(buildInvalidJsonError(text));
  }

  async function downloadBinary(url, options) {
    const nativeResponse = await sendNativeHttpRequest(url, options || {}, "arraybuffer");
    if (nativeResponse) {
      if (!nativeResponse.ok) {
        throw new Error(nativeResponse.error || "Download failed.");
      }
      if (nativeResponse.data instanceof ArrayBuffer) return nativeResponse.data;
      if (ArrayBuffer.isView(nativeResponse.data)) return nativeResponse.data.buffer;
      if (typeof nativeResponse.data === "string") return base64ToArrayBuffer(nativeResponse.data);
      throw new Error("Binary download returned an unsupported payload.");
    }

    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`Download failed (${response.status})`);
    }
    return response.arrayBuffer();
  }

  const USER_MAP = {
    "1": "Chris",
    "2": "MJ"
  };

  const TYPE_MAP = {
    "1": "Document",
    "2": "Meeting",
    "3": "Research",
    "4": "Internal"
  };

  function mapSharedTaskPayload(task) {
    if (!task) return task;
    const mapped = { ...task };
    
    if (mapped.owner && USER_MAP[String(mapped.owner)]) {
      mapped.owner = USER_MAP[String(mapped.owner)];
    }
    
    if (mapped.type && TYPE_MAP[String(mapped.type)]) {
      mapped.type = TYPE_MAP[String(mapped.type)];
    }
    
    return mapped;
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
      return { tasks: payload.map(mapSharedTaskPayload), updatedAt: "" };
    }

    if (payload && Array.isArray(payload.tasks)) {
      return {
        tasks: payload.tasks.map(mapSharedTaskPayload),
        updatedAt: payload.updatedAt || payload.updated_at || "",
      };
    }

    return { tasks: [], updatedAt: "" };
  }

  function sanitizeJsonText(text) {
    let value = String(text || "");
    if (value.charCodeAt(0) === 0xfeff) {
      value = value.slice(1);
    }
    value = value.trim();
    if (value.startsWith(")]}',")) {
      value = value.replace(/^\)\]\}',?\s*/, "");
    }
    return value;
  }

  function looksLikeHtmlPayload(text) {
    const sample = sanitizeJsonText(text).slice(0, 256).toLowerCase();
    return sample.startsWith("<!doctype html") || sample.startsWith("<html") || sample.includes("<head") || sample.includes("<body");
  }

  function summarizePayloadPreview(text, maxLen) {
    const cleaned = sanitizeJsonText(text).replace(/\s+/g, " ");
    return cleaned.slice(0, Math.max(40, maxLen || 180));
  }

  function buildInvalidJsonError(text) {
    const preview = summarizePayloadPreview(text, 180);
    if (looksLikeHtmlPayload(text)) {
      return `Sharedrive tasks file is not valid JSON (received HTML; likely auth/session/permissions). Preview: ${preview}`;
    }
    return `Sharedrive tasks file is not valid JSON. Preview: ${preview}`;
  }

  function parseJsonText(text) {
    if (!text) return null;
    try {
      return JSON.parse(sanitizeJsonText(text));
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

  async function uploadJsonToShareDrive(config, payload) {
    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());
    if (!config || !config.enabled) return null;
    if (!config.shareUrl) {
      throw new Error("Sharedrive share URL is missing.");
    }
    if (!canUseDesktop && !canUseBrowser) {
      throw new Error("Sharedrive sync unavailable (sign in required).");
    }

    const content = JSON.stringify(payload, null, 2);
    const fileName = String(config.fileName || "").trim();
    const contentBase64 = toBase64(content);

    if (canUseDesktop) {
      const result = await global.PlutusDesktop.uploadShareDriveFile({
        shareUrl: config.shareUrl,
        accessToken: config.accessToken,
        parentItemId: config.parentItemId,
        fileName,
        contentBase64,
        conflictBehavior: "replace",
      });
      if (!result || !result.ok) {
        throw new Error((result && result.error) || "Sharedrive upload failed.");
      }
      return result.data || null;
    }

    return uploadShareDriveFileBrowser({
      shareUrl: config.shareUrl,
      parentItemId: config.parentItemId,
      fileName,
      contentBase64,
    });
  }

  async function downloadJsonFromShareDrive(config) {
    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());
    if (!config || !config.enabled || !config.shareUrl) return null;

    if (canDesktopDirectDownload(config)) {
      const result = await global.PlutusDesktop.downloadShareDriveFile({
        shareUrl: config.shareUrl,
        accessToken: config.accessToken,
      });
      if (!result || !result.ok || !result.data || typeof result.data.text !== "string") {
        throw new Error((result && result.error) || "Sharedrive download failed.");
      }
      return parseJsonText(result.data.text);
    }

    if (canUseDesktop && hasShareDriveListBridge() && config.parentItemId) {
      const listing = await global.PlutusDesktop.listShareDriveChildren({
        shareUrl: config.shareUrl,
        accessToken: config.accessToken,
        parentItemId: config.parentItemId,
      });
      const items = listing && listing.ok && listing.data && Array.isArray(listing.data.items)
        ? listing.data.items
        : [];
      const target = items.find((item) => normalizeValue(item && item.name) === normalizeValue(config.fileName));
      if (target && target.id && listing.data && listing.data.driveId) {
        const result = await global.PlutusDesktop.downloadShareDriveFile({
          shareUrl: config.shareUrl,
          accessToken: config.accessToken,
          driveId: listing.data.driveId,
          itemId: target.id,
        });
        if (!result || !result.ok || !result.data || typeof result.data.text !== "string") {
          throw new Error((result && result.error) || "Sharedrive download failed.");
        }
        return parseJsonText(result.data.text);
      }
    }

    if (canUseBrowser) {
      try {
        return await downloadShareDriveJsonBrowser({
          shareUrl: config.shareUrl,
          parentItemId: config.parentItemId,
          fileName: config.fileName,
        });
      } catch {
        // fall through
      }
    }

    return null;
  }

  async function downloadTextFromShareDrive(config) {
    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());
    if (!config || !config.enabled || !config.shareUrl) return "";

    if (canDesktopDirectDownload(config)) {
      const result = await global.PlutusDesktop.downloadShareDriveFile({
        shareUrl: config.shareUrl,
        accessToken: config.accessToken,
      });
      if (!result || !result.ok || !result.data || typeof result.data.text !== "string") {
        throw new Error((result && result.error) || "Sharedrive download failed.");
      }
      return result.data.text;
    }

    if (canUseDesktop && hasShareDriveListBridge() && config.parentItemId) {
      const listing = await global.PlutusDesktop.listShareDriveChildren({
        shareUrl: config.shareUrl,
        accessToken: config.accessToken,
        parentItemId: config.parentItemId,
      });
      const items = listing && listing.ok && listing.data && Array.isArray(listing.data.items)
        ? listing.data.items
        : [];
      const target = items.find((item) => normalizeValue(item && item.name) === normalizeValue(config.fileName));
      if (target && target.id && listing.data && listing.data.driveId) {
        const result = await global.PlutusDesktop.downloadShareDriveFile({
          shareUrl: config.shareUrl,
          accessToken: config.accessToken,
          driveId: listing.data.driveId,
          itemId: target.id,
        });
        if (!result || !result.ok || !result.data || typeof result.data.text !== "string") {
          throw new Error((result && result.error) || "Sharedrive download failed.");
        }
        return result.data.text;
      }
    }

    if (canUseBrowser) {
      try {
        return await downloadShareDriveTextBrowser({
          shareUrl: config.shareUrl,
          parentItemId: config.parentItemId,
          fileName: config.fileName,
        });
      } catch {
        // fall through
      }
    }

    return "";
  }

  async function inspectSharedJsonSource(kind) {
    const normalizedKind = normalizeValue(kind);
    const config =
      normalizedKind === "tasks"
        ? getSharedTasksConfig()
        : normalizedKind === "deals"
          ? getSharedDealsConfig()
          : normalizedKind === "config"
            ? getSharedDashboardConfigSyncConfig()
            : null;

    if (!config) {
      return { ok: false, kind: normalizedKind, error: "Unknown shared file type." };
    }

    const summary = {
      ok: false,
      kind: normalizedKind,
      enabled: Boolean(config.enabled),
      fileName: config.fileName || "",
      shareUrl: config.shareUrl || "",
      parentItemId: config.parentItemId || "",
      preview: "",
      parsed: false,
      parsedType: "",
      itemCount: 0,
      error: "",
    };

    if (!config.enabled) {
      summary.error = "Shared sync is not enabled for this file.";
      return summary;
    }
    if (!config.shareUrl) {
      summary.error = "Sharedrive share URL is missing.";
      return summary;
    }

    try {
      const text = await downloadTextFromShareDrive(config);
      summary.preview = String(text || "").slice(0, 1200);
      const parsed = parseJsonText(text);
      summary.parsed = Boolean(parsed);
      if (Array.isArray(parsed)) {
        summary.parsedType = "array";
        summary.itemCount = parsed.length;
      } else if (parsed && typeof parsed === "object") {
        summary.parsedType = "object";
        if (Array.isArray(parsed.tasks)) summary.itemCount = parsed.tasks.length;
      }
      summary.ok = Boolean(parsed);
      if (!summary.ok) {
        summary.error = "Downloaded content is not valid JSON.";
      }
      return summary;
    } catch (error) {
      summary.error = error instanceof Error ? error.message : "Sharedrive inspection failed.";
      return summary;
    }
  }

  function canDesktopDirectDownload(config) {
    return Boolean(
      hasShareDriveBridge() &&
      config &&
      config.shareUrl &&
      config.explicitShareUrl,
    );
  }

  function publishTasksUpdate(source) {
    try {
      const detail = {
        tasks: getResolvedTasksData(),
        source: source || "local",
        syncedAt: sharedTasksState.lastSyncAt || "",
      };
      global.dispatchEvent(new CustomEvent("appcore:tasks-updated", { detail }));
    } catch (error) {
      console.warn("[AppCore] Failed to dispatch tasks update event", error);
    }
  }

  function publishDealsUpdate(source) {
    try {
      const detail = {
        deals: cloneArray(dealsCache || []),
        source: source || "local",
        syncedAt: sharedDealsState.lastSyncAt || "",
      };
      global.dispatchEvent(new CustomEvent("appcore:deals-updated", { detail }));
    } catch (error) {
      console.warn("[AppCore] Failed to dispatch deals update event", error);
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

  function publishSharedDealsStatus(stage) {
    try {
      const detail = {
        stage: stage || "idle",
        started: sharedDealsState.started,
        inFlight: sharedDealsState.inFlight,
        lastSyncAt: sharedDealsState.lastSyncAt,
        lastRemoteUpdatedAt: sharedDealsState.lastRemoteUpdatedAt,
        lastActionAt: sharedDealsState.lastActionAt,
        lastError: sharedDealsState.lastError,
      };
      global.dispatchEvent(new CustomEvent("appcore:deals-sync", { detail }));
    } catch (error) {
      console.warn("[AppCore] Failed to dispatch deals sync event", error);
    }
  }

  function ensureDealsCacheInitialized() {
    if (dealsCache) return;
    const config = getSharedDealsConfig();
    const desktopDeals = readDataJson("deals");
    if (Array.isArray(desktopDeals)) {
      dealsCache = cloneArray(desktopDeals);
      return;
    }

    const storedDeals = readArrayFromStorage(STORAGE_KEYS.deals);
    if (Array.isArray(storedDeals) && storedDeals.length) {
      dealsCache = cloneArray(storedDeals);
      return;
    }

    const bundledDeals = Array.isArray(global.DEALS) ? cloneArray(global.DEALS) : [];
    if (bundledDeals.length || !config.enabled) {
      dealsCache = bundledDeals;
      return;
    }

    dealsCache = [];
  }

  function ensureTasksCacheInitialized() {
    if (tasksCache) return;
    const desktopTasks = readDataJson("tasks");
    console.log("[AppCore] Local desktopTasks:", desktopTasks);
    tasksCache = Array.isArray(desktopTasks)
      ? cloneArray(desktopTasks)
      : (readArrayFromStorage(STORAGE_KEYS.tasks) || cloneArray(global.TASKS));
    console.log("[AppCore] Initialized tasksCache:", tasksCache);
  }

  function normalizeDealLegalLinksForTasks(deal) {
    const source =
      deal && Array.isArray(deal.legalLinks)
        ? deal.legalLinks
        : deal && Array.isArray(deal.legalAspects)
          ? deal.legalAspects
          : [];

    return source
      .map((entry, index) => {
        if (typeof entry === "string") {
          const url = String(entry || "").trim();
          return url ? { title: `Legal link ${index + 1}`, url } : null;
        }
        if (!entry || typeof entry !== "object") return null;
        const title = String(entry.title || entry.label || entry.name || "").trim();
        const url = String(entry.url || entry.href || entry.link || "").trim();
        if (!title && !url) return null;
        return {
          title: title || (url ? `Legal link ${index + 1}` : ""),
          url,
        };
      })
      .filter(Boolean);
  }

  function toSafeExternalTaskUrl(value) {
    const raw = String(value || "").trim();
    if (!raw) return "";
    try {
      const parsed = new URL(raw);
      const protocol = String(parsed.protocol || "").toLowerCase();
      if (protocol !== "http:" && protocol !== "https:") return "";
      return parsed.toString();
    } catch {
      return "";
    }
  }

  function stripComputedAutoTasks(tasks) {
    return (Array.isArray(tasks) ? tasks : []).filter((task) => {
      const metaSource = normalizeValue(task && task.metaSource);
      const taskId = normalizeValue(task && task.id);
      return !(
        metaSource === "deal-readiness-status" ||
        taskId.startsWith(`${AUTO_DEAL_READINESS_TASK_PREFIX}-`)
      );
    });
  }

  function normalizeDealLifecycleStatus(value) {
    const normalized = normalizeValue(value);
    if (normalized === "finished") return "finished";
    if (normalized === "closed") return "closed";
    return "active";
  }

  function buildAutoDealReadinessTasks(deals) {
    const createdAt = new Date().toISOString().slice(0, 10);
    return (Array.isArray(deals) ? deals : [])
      .flatMap((deal) => {
        if (!deal || typeof deal !== "object" || !String(deal.id || "").trim()) return [];
        if (normalizeDealLifecycleStatus(deal.lifecycleStatus || deal.dealStatus) !== "active") return [];

        const dealId = String(deal.id || "").trim();
        const owner = String(deal.seniorOwner || deal.owner || "System").trim() || "System";
        const companyLabel = String(deal.company || deal.name || dealId).trim();
        const legalLinks = normalizeDealLegalLinksForTasks(deal);
        const hasLegalDocument = legalLinks.some((entry) => Boolean(toSafeExternalTaskUrl(entry.url)));
        const hasDeck = Boolean(toSafeExternalTaskUrl(deal.deckUrl));
        const hasDashboard = Boolean(String(deal.fundraisingDashboardId || "").trim());

        const requirements = [
          {
            category: "legal",
            missing: !hasLegalDocument,
            title: `[Auto] Add legal document for ${companyLabel}`,
            type: "Legal setup",
            notes: `Auto-generated on ${createdAt}. Attach at least one legal document link to this deal.`,
          },
          {
            category: "deck",
            missing: !hasDeck,
            title: `[Auto] Add deck for ${companyLabel}`,
            type: "Deck setup",
            notes: `Auto-generated on ${createdAt}. Link a deck PDF to this deal.`,
          },
          {
            category: "dashboard",
            missing: !hasDashboard,
            title: `[Auto] Attach dashboard for ${companyLabel}`,
            type: "Dashboard setup",
            notes: `Auto-generated on ${createdAt}. Attach a fundraising dashboard to this deal.`,
          },
        ];

        return requirements
          .filter((entry) => entry.missing)
          .map((entry) => ({
            id: `${AUTO_DEAL_READINESS_TASK_PREFIX}-${dealId}-${entry.category}`,
            owner,
            dealId,
            title: entry.title,
            type: entry.type,
            status: "in progress",
            dueDate: "",
            notes: entry.notes,
            metaSource: "deal-readiness-status",
            metaCategory: entry.category,
          }));
      });
  }

  function getResolvedTasksData() {
    ensureTasksCacheInitialized();
    ensureDealsCacheInitialized();
    const baseTasks = stripComputedAutoTasks(tasksCache);
    const autoDealReadinessTasks = buildAutoDealReadinessTasks(dealsCache);
    return cloneArray(baseTasks.concat(autoDealReadinessTasks));
  }

  function loadDealsData() {
    const config = getSharedDealsConfig();
    ensureDealsCacheInitialized();

    if (config.enabled) {
      ensureSharedDealsSync();
    }
    publishSharedDealsStatus(sharedDealsState.started ? "ready" : "idle");
    return sortDealsByRetainerState(cloneArray(dealsCache), (left, right) => {
      const leftLabel = normalizeValue(left && (left.company || left.name || left.id));
      const rightLabel = normalizeValue(right && (right.company || right.name || right.id));
      return leftLabel.localeCompare(rightLabel);
    });
  }

  function saveDealsData(deals) {
    const config = getSharedDealsConfig();
    dealsCache = cloneArray(deals);
    if (!config.enabled) {
      writeArrayToStorage(STORAGE_KEYS.deals, dealsCache);
    }
    writeDataJson("deals", Array.isArray(dealsCache) ? dealsCache : []);
    publishDealsUpdate("local");
    publishTasksUpdate("deals-local");
    if (!config.enabled) {
      return Promise.resolve();
    }
    return performSharedDealsUpload(dealsCache);
  }

  async function performSharedDealsUpload(deals) {
    const config = getSharedDealsConfig();
    if (!config.enabled) return;
    try {
      sharedDealsState.lastActionAt = new Date().toISOString();
      publishSharedDealsStatus("uploading");
      await uploadJsonToShareDrive(config, Array.isArray(deals) ? deals : []);
      sharedDealsState.lastSyncAt = new Date().toISOString();
      sharedDealsState.lastError = "";
      publishSharedDealsStatus("uploaded");
    } catch (error) {
      sharedDealsState.lastError = error instanceof Error ? error.message : "Deals upload failed.";
      publishSharedDealsStatus("upload_failed");
      throw error;
    }
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

  function ensureSharedDealsSync() {
    const config = getSharedDealsConfig();
    if (!config.enabled || sharedDealsState.started) return;

    sharedDealsState.started = true;
    setTimeout(() => refreshDealsFromShareDrive("init"), 0);

    if (config.pollIntervalMs > 0) {
      sharedDealsSyncTimer = setInterval(
        () => refreshDealsFromShareDrive("poll"),
        config.pollIntervalMs,
      );
    }
  }

  function loadTasksData() {
    console.log("[AppCore] loadTasksData called");
    const config = getSharedTasksConfig();
    console.log("[AppCore] Tasks config:", config);
    ensureTasksCacheInitialized();

    if (config.enabled) {
      console.log("[AppCore] Triggering task sync from loadTasksData");
      ensureSharedTasksSync();
    }
    publishSharedTasksStatus(sharedTasksState.started ? "ready" : "idle");
    return getResolvedTasksData();
  }

  function saveTasksData(tasks) {
    const config = getSharedTasksConfig();
    tasksCache = stripComputedAutoTasks(tasks);
    if (config.enabled) {
      sharedTasksDirty = true;
      sharedTasksState.lastActionAt = new Date().toISOString();
      publishSharedTasksStatus("dirty");
    }
    if (!config.enabled) {
      writeArrayToStorage(STORAGE_KEYS.tasks, tasksCache);
      writeDataJson("tasks", tasksCache);
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
        console.log("[AppCore] getShareDriveDownloadUrl result:", result);
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
          const rawText = result && result.text ? result.text : "";
          payload = parseJsonText(rawText);
          if (!payload) {
            throw new Error(buildInvalidJsonError(rawText));
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
          const rawText = result.data.text;
          payload = parseJsonText(rawText);
          if (!payload) {
            throw new Error(buildInvalidJsonError(rawText));
          }
        } else if (canUseBrowser && config.shareUrl) {
          const result = await downloadShareDriveFileBrowser(config.shareUrl);
          const rawText = result && result.text ? result.text : "";
          payload = parseJsonText(rawText);
          if (!payload) {
            throw new Error(buildInvalidJsonError(rawText));
          }
        } else {
          throw error;
        }
      }
      const parsed = parseSharedTasksPayload(payload);
      console.log("[AppCore] Parsed shared tasks:", parsed);
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
      } else {
        console.warn("[AppCore] Parsed tasks is not an array");
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

  async function refreshDealsFromShareDrive(reason) {
    const config = getSharedDealsConfig();
    if (!config.enabled) return null;
    if (sharedDealsState.inFlight) return null;

    const canUseDesktop = hasShareDriveBridge();
    const canUseBrowser = Boolean(getBrowserAuthConfig());

    if (!config.downloadUrl && !config.shareUrl) {
      sharedDealsState.lastError = "Sharedrive deals URL is missing.";
      return null;
    }

    sharedDealsState.inFlight = true;
    sharedDealsState.lastActionAt = new Date().toISOString();
    publishSharedDealsStatus("syncing");

    try {
      let payload = null;
      if (config.downloadUrl) {
        payload = await fetchJson(config.downloadUrl, { cache: "no-store" });
      }
      if (payload == null && config.shareUrl) {
        payload = await downloadJsonFromShareDrive(config);
      }

      if (Array.isArray(payload)) {
        dealsCache = cloneArray(payload);
        if (!config.enabled) {
          writeArrayToStorage(STORAGE_KEYS.deals, dealsCache);
        }
        writeDataJson("deals", dealsCache);
        sharedDealsState.lastSyncAt = new Date().toISOString();
        sharedDealsState.lastError = "";
        publishSharedDealsStatus("synced");
        publishDealsUpdate(reason || "sharedrive");
        publishTasksUpdate(reason || "sharedrive");
      }
    } catch (error) {
      sharedDealsState.lastError = error instanceof Error ? error.message : "Deals sync failed.";
      console.warn("[AppCore] Sharedrive deals refresh failed", error);
      publishSharedDealsStatus("sync_failed");
    } finally {
      sharedDealsState.inFlight = false;
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
    return findDealByReference(deals, rawDealId);
  }

  function getDealReferenceKeys(deal) {
    if (!deal || typeof deal !== "object") return [];

    const references = [
      deal.id,
      deal.name,
      deal.company,
      deal.fundraisingDashboardId,
      deal.company && deal.name ? `${deal.company} ${deal.name}` : "",
      deal.name && deal.company ? `${deal.name} ${deal.company}` : "",
    ];

    return references
      .map((value) => normalizeValue(value))
      .filter(Boolean);
  }

  function findDealByReference(deals, reference) {
    if (!Array.isArray(deals)) return null;

    const references = Array.isArray(reference) ? reference : [reference];
    const keys = references
      .map((value) => normalizeValue(value))
      .filter(Boolean);

    if (!keys.length) return null;

    return deals.find((deal) => {
      const dealKeys = getDealReferenceKeys(deal);
      return keys.some((key) => dealKeys.includes(key));
    }) || null;
  }

  function getDashboardById(config, dashboardId) {
    const key = normalizeValue(dashboardId);
    if (!key || !config || !Array.isArray(config.dashboards)) return null;
    return config.dashboards.find((dashboard) => normalizeValue(dashboard.id) === key) || null;
  }

  function getDashboardForDeal(deal, config) {
    if (!deal) return null;
    const resolvedConfig = config || global.DASHBOARD_CONFIG || {};
    return getDashboardById(resolvedConfig, deal.fundraisingDashboardId);
  }

  function decodeJwtPayload(token) {
    const encoded = String(token || "").split(".")[1] || "";
    if (!encoded) return null;

    try {
      const base64 = encoded.replace(/-/g, "+").replace(/_/g, "/");
      const padded = base64.padEnd(base64.length + ((4 - (base64.length % 4)) % 4), "=");
      const json = atob(padded);
      return JSON.parse(json);
    } catch {
      return null;
    }
  }

  function getUserDirectoryMap(config) {
    const resolvedConfig = config || mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} });
    const rawDirectory =
      resolvedConfig &&
      resolvedConfig.settings &&
      resolvedConfig.settings.userDirectory &&
      typeof resolvedConfig.settings.userDirectory === "object"
        ? resolvedConfig.settings.userDirectory
        : {};

    return Object.entries(rawDirectory).reduce((accumulator, entry) => {
      const email = normalizeValue(entry[0]);
      const alias = String(entry[1] || "").trim();
      if (email && alias) accumulator[email] = alias;
      return accumulator;
    }, {});
  }

  function getConnectedEmailFromClaims(claims) {
    if (!claims || typeof claims !== "object") return "";
    return String(
      claims.preferred_username ||
      claims.upn ||
      claims.unique_name ||
      claims.email ||
      claims.mail ||
      "",
    ).trim();
  }

  function buildFallbackAlias(email, displayName) {
    if (displayName) return displayName;
    const localPart = String(email || "").split("@")[0] || "";
    if (!localPart) return "Unknown user";
    return localPart
      .split(/[._-]+/)
      .filter(Boolean)
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
      .join(" ");
  }

  async function getCurrentConnectedPerson() {
    const config = mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} });
    const directory = getUserDirectoryMap(config);
    const desktopSession = await readDesktopGraphSession();
    const browserSession = readBrowserGraphSession();
    const session = desktopSession || browserSession || null;

    if (!session) {
      return {
        connected: false,
        alias: "No user connected",
        email: "",
        displayName: "",
        source: "",
      };
    }

    const claims = decodeJwtPayload(session.accessToken);
    const email = normalizeValue(getConnectedEmailFromClaims(claims));
    const displayName = String((claims && claims.name) || "").trim();
    const alias = directory[email] || buildFallbackAlias(email, displayName);

    return {
      connected: Boolean(email || displayName),
      alias,
      email,
      displayName,
      source: desktopSession ? "desktop" : "browser",
    };
  }

  function getAllowedEmailsForPage(pageId) {
    const config = global.PlutusAppConfig;
    const page =
      config && typeof config.getPage === "function"
        ? config.getPage(pageId)
        : (config && config.pages && config.pages[pageId]) || null;
    const allowedEmails = page && Array.isArray(page.allowedEmails) ? page.allowedEmails : [];
    return Array.from(new Set(allowedEmails.map((entry) => normalizeValue(entry)).filter(Boolean)));
  }

  async function getPageAccessStatus(pageId) {
    const allowedEmails = getAllowedEmailsForPage(pageId);
    if (!allowedEmails.length) {
      return {
        pageId,
        restricted: false,
        allowed: true,
        allowedEmails: [],
        connected: false,
        email: "",
        person: null,
      };
    }

    const person = await getCurrentConnectedPerson();
    const email = normalizeValue(person && person.email);
    return {
      pageId,
      restricted: true,
      allowed: Boolean(email && allowedEmails.includes(email)),
      allowedEmails,
      connected: Boolean(person && person.connected),
      email,
      person,
    };
  }

  async function resolveShareDriveDownloadUrl(shareUrl, accessToken) {
    const cleanShareUrl = String(shareUrl || "").trim();
    if (!cleanShareUrl) return "";

    if (
      hasShareDriveBridge() &&
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.getShareDriveDownloadUrl === "function"
    ) {
      const result = await global.PlutusDesktop.getShareDriveDownloadUrl({
        shareUrl: cleanShareUrl,
        accessToken: String(accessToken || "").trim(),
      });
      if (!result || !result.ok || !result.data || !result.data.downloadUrl) {
        throw new Error((result && result.error) || "Failed to resolve sharedrive download URL.");
      }
      return result.data.downloadUrl;
    }

    if (getBrowserAuthConfig()) {
      return resolveShareDriveDownloadUrlBrowser(cleanShareUrl);
    }

    return "";
  }

  async function resolveShareDriveFile({ shareUrl, parentItemId, fileName, accessToken }) {
    const cleanShareUrl = String(shareUrl || "").trim();
    if (!cleanShareUrl) {
      throw new Error("Sharedrive share URL is required.");
    }

    if (
      hasShareDriveBridge() &&
      global.PlutusDesktop &&
      typeof global.PlutusDesktop.getShareDriveDownloadUrl === "function"
    ) {
      const result = await global.PlutusDesktop.getShareDriveDownloadUrl({
        shareUrl: cleanShareUrl,
        accessToken: String(accessToken || "").trim(),
      });
      if (!result || !result.ok || !result.data) {
        throw new Error((result && result.error) || "Failed to resolve sharedrive file.");
      }
      const data = result.data || {};
      return {
        id: String(data.itemId || data.id || "").trim(),
        itemId: String(data.itemId || data.id || "").trim(),
        name: String(data.name || fileName || "").trim(),
        driveId: String(data.driveId || "").trim(),
        parentItemId: String(data.parentItemId || parentItemId || "").trim(),
        downloadUrl: String(data.downloadUrl || "").trim(),
      };
    }

    if (getBrowserAuthConfig()) {
      const resolved = await resolveShareDriveFileBrowser({
        shareUrl: cleanShareUrl,
        parentItemId,
        fileName,
      });
      return {
        id: String(resolved.itemId || "").trim(),
        itemId: String(resolved.itemId || "").trim(),
        name: String(resolved.name || fileName || "").trim(),
        driveId: String(resolved.driveId || "").trim(),
        parentItemId: String(resolved.parentItemId || parentItemId || "").trim(),
        downloadUrl: String(resolved.downloadUrl || "").trim(),
      };
    }

    throw new Error("Sharedrive sync unavailable (sign in required).");
  }

  async function uploadShareDriveFile(payload) {
    const config = payload && typeof payload === "object" ? payload : {};
    const shareUrl = String(config.shareUrl || "").trim();
    const fileName = String(config.fileName || "").trim();
    const contentBase64 = String(config.contentBase64 || "").trim();

    if (!shareUrl) throw new Error("Sharedrive share URL is required.");
    if (!fileName) throw new Error("File name is required.");
    if (!contentBase64) throw new Error("Upload content is empty.");

    if (hasShareDriveBridge()) {
      const result = await global.PlutusDesktop.uploadShareDriveFile({
        shareUrl,
        accessToken: String(config.accessToken || "").trim(),
        parentItemId: String(config.parentItemId || "").trim(),
        fileName,
        contentBase64,
        conflictBehavior: String(config.conflictBehavior || "replace").trim() || "replace",
      });
      if (!result || !result.ok) {
        throw new Error((result && result.error) || "Sharedrive upload failed.");
      }
      return result.data || null;
    }

    if (getBrowserAuthConfig()) {
      return uploadShareDriveFileBrowser({
        shareUrl,
        parentItemId: String(config.parentItemId || "").trim(),
        fileName,
        contentBase64,
      });
    }

    throw new Error("Sharedrive sync unavailable (sign in required).");
  }

  function ensureSidebarUserCard() {
    if (!global.document) return null;
    const sidebar = global.document.querySelector(".sidebar-nav");
    if (!sidebar) return null;

    let card = sidebar.querySelector("[data-sidebar-user-card]");
    if (card) return card;

    card = global.document.createElement("section");
    card.className = "sidebar-user-card";
    card.setAttribute("data-sidebar-user-card", "true");
    card.innerHTML = `
      <div class="sidebar-user-label">Connected user</div>
      <div class="sidebar-user-name">Checking connection...</div>
      <div class="sidebar-user-email">Waiting for session data</div>
      <div class="sidebar-user-status">Sharedrive sign-in</div>
    `;
    sidebar.appendChild(card);
    return card;
  }

  async function renderSidebarUserCard() {
    const card = ensureSidebarUserCard();
    if (!card) return;

    const info = await getCurrentConnectedPerson();
    const nameEl = card.querySelector(".sidebar-user-name");
    const emailEl = card.querySelector(".sidebar-user-email");
    const statusEl = card.querySelector(".sidebar-user-status");
    if (!nameEl || !emailEl || !statusEl) return;

    nameEl.textContent = info.alias || "Unknown user";
    emailEl.textContent = info.email || "No mapped email available";
    statusEl.textContent = info.connected
      ? `Signed in${info.source ? ` via ${info.source}` : ""}`
      : "Sign in from Workspace connection";
    card.classList.toggle("is-connected", Boolean(info.connected));
  }

  function initializeSidebarUserCard() {
    if (!global.document) return;
    renderSidebarUserCard();
    global.addEventListener("appcore:graph-session-updated", renderSidebarUserCard);
  }

  function isAutoTask(task) {
    const title = normalizeValue(task && task.title);
    const metaSource = normalizeValue(task && task.metaSource);
    const taskId = normalizeValue(task && task.id);
    return (
      title.startsWith("[auto]") ||
      metaSource === "dashboard-contact-status" ||
      metaSource === "deal-readiness-status" ||
      taskId.startsWith(`${AUTO_CONTACT_TASK_PREFIX}-`) ||
      taskId.startsWith(`${AUTO_DEAL_READINESS_TASK_PREFIX}-`)
    );
  }

  global.AppCore = {
    STORAGE_KEYS,
    AUTO_CONTACT_TASK_PREFIX,
    normalizeValue,
    parseDealAmount,
    getDealRetainerRawValue,
    getDealRetainerState,
    hasPositiveRetainer,
    matchesDealRetainerFilter,
    compareDealsByRetainerState,
    sortDealsByRetainerState,
    getPageUrl,
    loadDealsData,
    saveDealsData,
    loadTasksData,
    saveTasksData,
    refreshTasksFromShareDrive,
    refreshDealsFromShareDrive,
    requestGraphDeviceCode: requestDeviceCodeBrowser,
    pollGraphDeviceCode: pollDeviceCodeBrowser,
    listShareDriveChildren: listShareDriveChildrenBrowser,
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
    getSharedDealsStatus: () => ({
      stage: "idle",
      started: sharedDealsState.started,
      inFlight: sharedDealsState.inFlight,
      lastSyncAt: sharedDealsState.lastSyncAt,
      lastRemoteUpdatedAt: sharedDealsState.lastRemoteUpdatedAt,
      lastActionAt: sharedDealsState.lastActionAt,
      lastError: sharedDealsState.lastError,
    }),
    refreshDashboardConfigFromShareDrive,
    getDashboardConfig: () =>
      mergeDashboardConfig(global.DASHBOARD_CONFIG || { dashboards: [], settings: {} }),
    upsertDashboardConfigEntry,
    findDealForTask,
    findDealByReference,
    getDashboardById,
    getDashboardForDeal,
    isAutoTask,
    resolveShareDriveFile,
    resolveShareDriveDownloadUrl,
    downloadBinary,
    uploadShareDriveFile,
    inspectSharedJsonSource,
    listOutlookMessages,
    getCurrentConnectedPerson,
    getPageAccessStatus,
    refreshConnectedPersonUi: renderSidebarUserCard,
  };

  if (global.document) {
    if (global.document.readyState === "loading") {
      global.document.addEventListener("DOMContentLoaded", initializeSidebarUserCard, { once: true });
    } else {
      initializeSidebarUserCard();
    }
  }
})(window);
