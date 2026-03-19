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
      label: "Sharedrive folders",
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
