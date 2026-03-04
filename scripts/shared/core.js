(function initAppCore(global) {
  const STORAGE_KEYS = {
    deals: "deals_data_v1",
    tasks: "owner_tasks_v1",
  };

  const AUTO_CONTACT_TASK_PREFIX = "auto-contact-status";

  function normalizeValue(value) {
    return String(value || "").trim().toLowerCase();
  }

  function cloneArray(value) {
    if (!Array.isArray(value)) return [];
    return JSON.parse(JSON.stringify(value));
  }

  function readArrayFromStorage(key) {
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
    try {
      localStorage.setItem(key, JSON.stringify(Array.isArray(values) ? values : []));
    } catch (error) {
      console.warn(`[AppCore] Failed to write ${key}`, error);
    }
  }

  function loadDealsData() {
    return readArrayFromStorage(STORAGE_KEYS.deals) || cloneArray(global.DEALS);
  }

  function saveDealsData(deals) {
    writeArrayToStorage(STORAGE_KEYS.deals, deals);
  }

  function loadTasksData() {
    return readArrayFromStorage(STORAGE_KEYS.tasks) || cloneArray(global.TASKS);
  }

  function saveTasksData(tasks) {
    writeArrayToStorage(STORAGE_KEYS.tasks, tasks);
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
    findDealForTask,
    isAutoTask,
  };
})(window);
