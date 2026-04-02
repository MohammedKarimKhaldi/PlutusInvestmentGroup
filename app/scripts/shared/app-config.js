(function initPlutusAppConfig(global) {
  const PAGE_MAP = {
    "investor-dashboard": {
      file: "investor-dashboard.html",
      label: "Investor dashboard",
      showInNav: true,
    },
    "deals-overview": {
      file: "deals-overview.html",
      label: "Deals overview",
      showInNav: true,
    },
    "deal-ownership": {
      file: "deal-ownership.html",
      label: "Deal ownership",
      showInNav: true,
    },
    accounting: {
      file: "accounting.html",
      label: "Accounting",
      showInNav: true,
      allowedEmails: [
        "fl@plutus-investment.com",
        "mj@plutus-investment.com",
        "mk@plutus-investment.com"
      ],
    },
    "legal-management": {
      file: "legal-management.html",
      label: "Legal",
      showInNav: true,
    },
    "tasks-management": {
      file: "tasks-management.html",
      label: "Tasks by owner",
      showInNav: true,
    },
    "owner-tasks": {
      file: "owner-tasks.html",
      label: "Owner tasks",
      navPageId: "tasks-management",
    },
    "deal-details": {
      file: "deal-details.html",
      label: "Deal details",
      navPageId: "deals-overview",
    },
    "sharedrive-folders": {
      file: "sharedrive-folders.html",
      label: "Workspace connection",
      showInNav: true,
    },
    "outlook-investor-sync": {
      file: "outlook-investor-sync.html",
      label: "Outlook investors",
      showInNav: true,
    },
  };

  const storageKeys = {
    deals: "deals_data_v1",
    tasks: "owner_tasks_v1",
    sharedriveGate: "sharedrive_connected_v1",
    graphSession: "plutus_graph_session_v1",
    dashboardConfig: "plutus_dashboard_config_v1",
  };

  const dataFiles = {
    config: "config.json",
    deals: "deals.json",
    tasks: "tasks.json",
    sharedTasks: "sharedrive-tasks.json",
  };

  function getPage(pageId) {
    return PAGE_MAP[pageId] || null;
  }

  function buildPageHref(pageId, params) {
    const page = getPage(pageId);
    const baseHref = page ? page.file : `${pageId}.html`;
    const query = new URLSearchParams();

    Object.entries(params || {}).forEach(([key, value]) => {
      if (value == null || value === "") return;
      query.set(key, String(value));
    });

    const queryString = query.toString();
    return queryString ? `${baseHref}?${queryString}` : baseHref;
  }

  function getCurrentPageId() {
    if (!global.document || !global.document.body) return "";
    const explicit = String(global.document.body.dataset.pageId || "").trim();
    if (explicit) return explicit;

    const currentFile = String((global.location && global.location.pathname) || "")
      .split("/")
      .pop()
      .trim();

    const match = Object.entries(PAGE_MAP).find(([, page]) => page.file === currentFile);
    return match ? match[0] : "";
  }

  function getCurrentNavPageId() {
    const pageId = getCurrentPageId();
    const page = getPage(pageId);
    return (page && page.navPageId) || pageId;
  }

  function localReadDataJson(key) {
    try {
      const raw = localStorage.getItem(String(key || "").trim());
      if (!raw) return { ok: false, data: null, error: "not_found" };
      return { ok: true, data: JSON.parse(raw), error: null };
    } catch (error) {
      return { ok: false, data: null, error: String(error) };
    }
  }

  function localWriteDataJson(key, payload) {
    try {
      localStorage.setItem(String(key || "").trim(), JSON.stringify(payload || {}));
      return { ok: true };
    } catch (error) {
      return { ok: false, error: String(error) };
    }
  }

  function localReadArrayStore(key) {
    try {
      const raw = localStorage.getItem(String(key || "").trim());
      if (!raw) return null;
      const data = JSON.parse(raw);
      return Array.isArray(data) ? data : null;
    } catch {
      return null;
    }
  }

  function localWriteArrayStore(key, values) {
    try {
      localStorage.setItem(String(key || "").trim(), JSON.stringify(Array.isArray(values) ? values : []));
      return true;
    } catch {
      return false;
    }
  }

  const isTauri = Boolean(global.__TAURI__ && global.__TAURI__.tauri && typeof global.__TAURI__.tauri.invoke === "function");
  const tauriInvoke = isTauri ? (cmd, payload) => global.__TAURI__.tauri.invoke(cmd, payload) : null;

  function sanitizeTauriPayload(value) {
    if (Array.isArray(value)) {
      return value
        .map((entry) => sanitizeTauriPayload(entry))
        .filter((entry) => typeof entry !== "undefined");
    }

    if (value && typeof value === "object") {
      return Object.entries(value).reduce((result, [key, entry]) => {
        if (entry == null) return result;
        const sanitized = sanitizeTauriPayload(entry);
        if (typeof sanitized !== "undefined") result[key] = sanitized;
        return result;
      }, {});
    }

    if (typeof value === "undefined") return undefined;
    return value;
  }

  function tauriInvokeWithPayload(cmd, payload) {
    return tauriInvoke
      ? tauriInvoke(cmd, { payload: sanitizeTauriPayload(payload || {}) })
      : Promise.resolve({ ok: false, error: "Tauri unavailable" });
  }

  function syncLocalDataJsonFromTauri(key) {
    if (!isTauri || !tauriInvoke) return;
    tauriInvoke("read_data_json", { key })
      .then((result) => {
        if (result && result.ok) localWriteDataJson(key, result.data);
      })
      .catch(() => {});
  }

  function syncArrayStoreFromTauri(key) {
    if (!isTauri || !tauriInvoke) return;
    tauriInvoke("read_array_store", { key })
      .then((values) => {
        if (Array.isArray(values)) localWriteArrayStore(key, values);
      })
      .catch(() => {});
  }

  const desktopBridge = {
    readDataJson(key) {
      const result = localReadDataJson(key);
      syncLocalDataJsonFromTauri(key);
      return result;
    },
    writeDataJson(key, value) {
      const localResult = localWriteDataJson(key, value);
      if (isTauri && tauriInvoke) {
        tauriInvoke("write_data_json", { key, value }).catch(() => {});
      }
      return localResult;
    },
    readArrayStore(key) {
      const local = localReadArrayStore(key);
      syncArrayStoreFromTauri(key);
      return local;
    },
    writeArrayStore(key, values) {
      const local = localWriteArrayStore(key, values);
      if (isTauri && tauriInvoke) {
        tauriInvoke("write_array_store", { key, values }).catch(() => {});
      }
      return local;
    },
  };

  if (isTauri && tauriInvoke) {
    Object.assign(desktopBridge, {
      async listShareDriveChildren(payload) {
        return tauriInvokeWithPayload("list_share_drive_children", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      async getShareDriveDownloadUrl(payload) {
        return tauriInvokeWithPayload("get_share_drive_download_url", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      async downloadShareDriveFile(payload) {
        return tauriInvokeWithPayload("download_share_drive_file", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      async uploadShareDriveFile(payload) {
        return tauriInvokeWithPayload("upload_share_drive_file", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      async requestGraphDeviceCode() {
        return tauriInvoke("request_graph_device_code").catch((err) => ({ ok: false, error: String(err) }));
      },
      async pollGraphDeviceCode(payload) {
        return tauriInvokeWithPayload("poll_graph_device_code", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      async getGraphSession() {
        return tauriInvoke("get_graph_session").catch((err) => ({ ok: false, error: String(err) }));
      },
      async listOutlookMessages(payload) {
        return tauriInvokeWithPayload("list_outlook_messages", payload).catch((err) => ({ ok: false, error: String(err) }));
      },
      isTauri: true,
    });

    ["config", "deals", "tasks", "sharedrive-tasks"].forEach(syncLocalDataJsonFromTauri);
    [storageKeys.deals, storageKeys.tasks].forEach(syncArrayStoreFromTauri);
  }

  global.PlutusDesktop = Object.assign({}, global.PlutusDesktop || {}, desktopBridge);

  global.PlutusAppConfig = {
    entryPageId: "sharedrive-folders",
    storageKeys,
    dataFiles,
    pages: PAGE_MAP,
    getPage,
    getPageIds() {
      return Object.keys(PAGE_MAP);
    },
    getNavPages() {
      return Object.entries(PAGE_MAP)
        .filter(([, page]) => page.showInNav)
        .map(([pageId, page]) => ({ pageId, ...page }));
    },
    buildPageHref,
    getDataPath(dataKey) {
      const fileName = dataFiles[dataKey] || dataKey;
      return `../data/${fileName}`;
    },
    getCurrentPageId,
    getCurrentNavPageId,
  };
})(window);
