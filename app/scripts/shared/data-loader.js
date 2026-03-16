(() => {
  const appConfig = window.PlutusAppConfig || {};
  const dataFiles = Object.assign({
    config: "config.json",
    sharedTasks: "sharedrive-tasks.json",
    deals: "deals.json",
    tasks: "tasks.json",
  }, appConfig.dataFiles || {});

  function readDesktopJson(key) {
    if (window.PlutusDesktop && typeof window.PlutusDesktop.readDataJson === "function") {
      try {
        const result = window.PlutusDesktop.readDataJson(key);
        if (result && result.ok) return result.data;
      } catch {
        // ignore
      }
    }
    return null;
  }

  function readJsonSync(path) {
    try {
      const request = new XMLHttpRequest();
      request.open("GET", path, false);
      request.send(null);
      const hasResponseBody = String(request.responseText || "").trim().length > 0;
      if ((request.status >= 200 && request.status < 300) || (request.status === 0 && hasResponseBody)) {
        return JSON.parse(request.responseText);
      }
    } catch {
      // ignore
    }
    return null;
  }

  function ensureArray(value) {
    return Array.isArray(value) ? value : [];
  }

  function ensureObject(value) {
    return value && typeof value === "object" ? value : {};
  }

  function mergeConfig(baseConfig, overrideConfig) {
    const base = ensureObject(baseConfig);
    const override = ensureObject(overrideConfig);
    const mergedDashboards = Array.isArray(override.dashboards)
      ? ensureArray(override.dashboards)
      : ensureArray(base.dashboards);

    return {
      dashboards: mergedDashboards,
      settings: Object.assign({}, ensureObject(base.settings), ensureObject(override.settings)),
      proxies: override.proxies || base.proxies || [],
    };
  }

  const bootstrapConfigJson =
    readJsonSync(appConfig.getDataPath ? appConfig.getDataPath("config") : `../data/${dataFiles.config}`) ||
    {};
  const desktopConfigJson = readDesktopJson("config") || {};
  const sharedriveJson =
    readDesktopJson("sharedrive-tasks") ||
    readJsonSync(appConfig.getDataPath ? appConfig.getDataPath("sharedTasks") : `../data/${dataFiles.sharedTasks}`) ||
    {};
  const configJson = mergeConfig(bootstrapConfigJson, desktopConfigJson);

  const isFileProtocol =
    window.location && String(window.location.protocol || "").toLowerCase() === "file:";

  const sharedTasksEnabled = Boolean(
    sharedriveJson &&
      sharedriveJson.tasks &&
      sharedriveJson.tasks.enabled,
  );

  const sharedDealsEnabled = Boolean(
    configJson &&
      configJson.settings &&
      configJson.settings.sharedDeals &&
      configJson.settings.sharedDeals.enabled,
  );

  const dealsJson =
    readDesktopJson("deals") ||
    (sharedDealsEnabled || isFileProtocol
      ? null
      : readJsonSync(appConfig.getDataPath ? appConfig.getDataPath("deals") : `../data/${dataFiles.deals}`)) ||
    [];
  const tasksJson =
    readDesktopJson("tasks") ||
    (sharedTasksEnabled || isFileProtocol
      ? null
      : readJsonSync(appConfig.getDataPath ? appConfig.getDataPath("tasks") : `../data/${dataFiles.tasks}`)) ||
    [];

  const dashboards = ensureArray(configJson.dashboards);
  const settings = ensureObject(configJson.settings);
  const proxyBases = ensureArray(configJson.proxies);

  window.DASHBOARD_CONFIG = {
    dashboards,
    settings,
    proxies: proxyBases,
  };
  window.DASHBOARD_PROXIES = proxyBases.map(
    (base) => (url) => String(base || "") + encodeURIComponent(url),
  );

  window.DEALS = ensureArray(dealsJson);
  window.TASKS = ensureArray(tasksJson);
  window.SHAREDRIVE_TASKS = ensureObject(sharedriveJson);
})();
