(() => {
  const appConfig = window.PlutusAppConfig || {};
  const storageKeys = Object.assign({
    sharedriveGate: "sharedrive_connected_v1",
    graphSession: "plutus_graph_session_v1",
  }, appConfig.storageKeys || {});
  const SHAREDRIVE_GATE_KEY = storageKeys.sharedriveGate;
  const GRAPH_SESSION_KEY = storageKeys.graphSession;

  function hasGraphSession() {
    try {
      const raw = localStorage.getItem(GRAPH_SESSION_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      if (!payload || typeof payload !== "object") return false;
      if (payload.accessToken) return true;
      if (payload.refreshToken) return true;
      return false;
    } catch {
      return false;
    }
  }

  function hasSharedriveConnection() {
    if (localStorage.getItem(SHAREDRIVE_GATE_KEY) === "true") return true;
    if (hasGraphSession()) {
      localStorage.setItem(SHAREDRIVE_GATE_KEY, "true");
      return true;
    }
    return false;
  }

  const currentPath = window.location.pathname || "";
  const gatePageHref = appConfig.buildPageHref
    ? appConfig.buildPageHref(appConfig.entryPageId || "sharedrive-folders")
    : "sharedrive-folders.html";

  if (currentPath.endsWith(gatePageHref)) return;

  if (!hasSharedriveConnection()) {
    window.location.replace(gatePageHref);
  }
})();
