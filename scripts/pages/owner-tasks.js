const AppCore = window.AppCore;
const TASKS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.tasks) || "owner_tasks_v1";
let dealsData = [];
const AUTO_CONTACT_TASK_PREFIX = (AppCore && AppCore.AUTO_CONTACT_TASK_PREFIX) || "auto-contact-status";

function normalizeValue(value) {
  if (AppCore) return AppCore.normalizeValue(value);
  return String(value || "").trim().toLowerCase();
}

function loadAllTasks() {
  return AppCore ? AppCore.loadTasksData() : (Array.isArray(TASKS) ? JSON.parse(JSON.stringify(TASKS)) : []);
}

function saveAllTasks(allTasks) {
  if (AppCore) {
    AppCore.saveTasksData(allTasks);
    return;
  }
  try {
    localStorage.setItem(TASKS_STORAGE_KEY, JSON.stringify(allTasks));
  } catch (e) {
    console.warn("Failed to save tasks to storage", e);
  }
}

function loadDealsData() {
  dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
}

function findDealForTask(task) {
  if (AppCore) return AppCore.findDealForTask(dealsData, task);
  if (!Array.isArray(dealsData) || !task) return null;
  const rawDealId = task.dealId ?? task.deal ?? task.dealName ?? "";
  const dealKey = normalizeValue(rawDealId);
  if (!dealKey) return null;
  return dealsData.find((deal) => {
    const id = normalizeValue(deal.id);
    const name = normalizeValue(deal.name);
    const company = normalizeValue(deal.company);
    const dashboardId = normalizeValue(deal.fundraisingDashboardId);
    return id === dealKey || name === dealKey || company === dealKey || dashboardId === dealKey;
  }) || null;
}

function statusClass(status) {
  const s = String(status || "").toLowerCase();
  if (s === "in progress") return "badge-status-in-progress";
  if (s === "waiting") return "badge-status-waiting";
  if (s === "done") return "badge-status-done";
  return "";
}

function formatDate(value) {
  if (!value) return "No due date";
  const dt = new Date(value);
  if (isNaN(dt.getTime())) return String(value);
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
}

function formatTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
}

function truncateMessage(value, maxLength) {
  const text = String(value || "");
  if (!text) return "";
  if (text.length <= maxLength) return text;
  return `${text.slice(0, Math.max(0, maxLength - 1))}…`;
}

function applySharedriveStatus(detail) {
  const pill = document.getElementById("sharedrive-status-pill");
  const textEl = document.getElementById("sharedrive-status-text");
  const dotEl = document.getElementById("sharedrive-status-dot");
  if (!pill || !textEl || !dotEl || !detail) return;

  let label = "Sharedrive: idle";
  let color = "#94a3b8";
  let glow = "0 0 8px rgba(148, 163, 184, 0.6)";
  pill.removeAttribute("title");

  if (detail.lastError) {
    const shortError = truncateMessage(detail.lastError, 60);
    label = shortError ? `Sharedrive: error - ${shortError}` : "Sharedrive: error";
    color = "#f43f5e";
    glow = "0 0 8px rgba(244, 63, 94, 0.7)";
    pill.title = detail.lastError;
  } else if (detail.stage === "uploading") {
    label = "Sharedrive: uploading…";
    color = "#f59e0b";
    glow = "0 0 8px rgba(245, 158, 11, 0.7)";
  } else if (detail.stage === "upload_queued" || detail.dirty) {
    label = "Sharedrive: queued";
    color = "#f59e0b";
    glow = "0 0 8px rgba(245, 158, 11, 0.7)";
  } else if (detail.stage === "synced" || detail.stage === "uploaded") {
    const time = formatTime(detail.lastUploadAt || detail.lastSyncAt);
    label = time ? `Sharedrive: synced ${time}` : "Sharedrive: synced";
    color = "#22c55e";
    glow = "0 0 8px rgba(34, 197, 94, 0.9)";
  } else if (detail.stage === "syncing") {
    label = "Sharedrive: syncing…";
    color = "#38bdf8";
    glow = "0 0 8px rgba(56, 189, 248, 0.7)";
  }

  textEl.textContent = label;
  dotEl.style.background = color;
  dotEl.style.boxShadow = glow;
}

function isTaskOverdue(task, today) {
  const status = normalizeValue(task.status);
  if (!task.dueDate || status === "done") return false;
  const due = new Date(task.dueDate);
  if (isNaN(due.getTime())) return false;
  due.setHours(0, 0, 0, 0);
  return due < today;
}

function isAutoTask(task) {
  if (AppCore) return AppCore.isAutoTask(task);
  const title = normalizeValue(task && task.title);
  return title.startsWith("[auto]");
}

function renderOwnerPage() {
  const params = new URLSearchParams(window.location.search);
  const ownerParam = params.get("owner");
  const ownerKey = normalizeValue(ownerParam);

  const crumbOwner = document.getElementById("crumb-owner");
  const ownerTitle = document.getElementById("owner-title");
  const ownerSubtitle = document.getElementById("owner-subtitle");
  const ownerMetaRow = document.getElementById("owner-meta-row");
  const ownerTaskList = document.getElementById("owner-task-list");
  const ownerEmpty = document.getElementById("owner-empty");
  const tasksCount = document.getElementById("tasks-count");
  const searchInput = document.getElementById("task-search");
  const filterButtons = Array.from(document.querySelectorAll(".person-filter"));
  const layoutButtons = Array.from(document.querySelectorAll(".person-layout"));
  const editModal = document.getElementById("task-edit-modal");
  const editForm = document.getElementById("task-edit-form");
  const editTitleInput = document.getElementById("edit-task-title");
  const editStatusInput = document.getElementById("edit-task-status");
  const editDueInput = document.getElementById("edit-task-due");
  const editTypeInput = document.getElementById("edit-task-type");
  const editNotesInput = document.getElementById("edit-task-notes");
  let editingTaskId = null;

  if (!ownerParam) {
    crumbOwner.textContent = "Unknown";
    ownerTitle.textContent = "Owner not specified";
    ownerSubtitle.textContent = "Open this page from Tasks by owner.";
    ownerMetaRow.innerHTML = "";
    ownerTaskList.innerHTML = "";
    ownerEmpty.style.display = "block";
    tasksCount.textContent = "0 shown";
    return;
  }

  let allTasks = loadAllTasks();
  loadDealsData();
  if (AppCore) {
    window.addEventListener("appcore:tasks-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.tasks)) {
        allTasks = event.detail.tasks;
        renderTaskList();
      }
    });
    window.addEventListener("appcore:tasks-sync", (event) => {
      if (event && event.detail) {
        applySharedriveStatus(event.detail);
      }
    });
    if (typeof AppCore.getSharedTasksStatus === "function") {
      applySharedriveStatus(AppCore.getSharedTasksStatus());
    }
  }
  const getOwnerTasks = () => allTasks.filter((task) => normalizeValue(task.owner) === ownerKey);
  const baseTasks = getOwnerTasks();
  const displayOwner = baseTasks[0]?.owner || ownerParam;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  crumbOwner.textContent = displayOwner;
  ownerTitle.textContent = displayOwner;
  const renderOwnerSummary = (ownerTasks) => {
    const openCount = ownerTasks.filter((task) => normalizeValue(task.status) !== "done").length;
    const overdueCount = ownerTasks.filter((task) => isTaskOverdue(task, today)).length;

    ownerSubtitle.textContent = `${openCount} open task${openCount === 1 ? "" : "s"} · ${overdueCount} overdue`;
    ownerMetaRow.innerHTML = [
      `<div class="owner-chip"><strong>${ownerTasks.length}</strong> total</div>`,
      `<div class="owner-chip"><strong>${openCount}</strong> open</div>`,
      `<div class="owner-chip"><strong>${overdueCount}</strong> overdue</div>`,
    ].join("");
  };

  let currentView = "all";
  let currentLayout = "grouped";
  const collapsedGroups = new Set();

  function openTaskEditor(task) {
    if (!editModal || !editForm || !task) return;
    editingTaskId = task.id;
    editTitleInput.value = task.title || "";
    editStatusInput.value = normalizeValue(task.status) || "in progress";
    editDueInput.value = task.dueDate || "";
    editTypeInput.value = task.type || "";
    editNotesInput.value = task.notes || "";
    editModal.classList.add("open");
    editModal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
    editTitleInput.focus();
    editTitleInput.select();
  }

  function closeTaskEditor() {
    if (!editModal || !editForm) return;
    editingTaskId = null;
    editForm.reset();
    editModal.classList.remove("open");
    editModal.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
  }

  function applyFilters(tasks) {
    const keyword = normalizeValue(searchInput?.value || "");

    return tasks.filter((task) => {
      const status = normalizeValue(task.status);
      const overdue = isTaskOverdue(task, today);

      if (currentView === "open" && status === "done") return false;
      if (currentView === "done" && status !== "done") return false;
      if (currentView === "overdue" && !overdue) return false;

      if (!keyword) return true;

      const deal = findDealForTask(task);
      const haystack = [
        task.title,
        task.notes,
        task.type,
        task.status,
        task.dueDate,
        deal?.name,
        deal?.company,
      ]
        .map(normalizeValue)
        .join(" ");

      return haystack.includes(keyword);
    });
  }

  function renderTaskList() {
    const ownerTasks = getOwnerTasks();
    renderOwnerSummary(ownerTasks);

    const tasks = [...ownerTasks].sort((a, b) => {
      if (!a.dueDate && !b.dueDate) return 0;
      if (!a.dueDate) return 1;
      if (!b.dueDate) return -1;
      return new Date(a.dueDate) - new Date(b.dueDate);
    });

    const filtered = applyFilters(tasks);
    const groupedCount = new Set(
      filtered.map((task) => {
        const deal = findDealForTask(task);
        return deal?.company || "General";
      }),
    ).size;
    tasksCount.textContent = `${filtered.length} shown${currentLayout === "grouped" ? ` · ${groupedCount} companies` : ""}`;

    if (!filtered.length) {
      ownerTaskList.innerHTML = "";
      ownerEmpty.style.display = "block";
      return;
    }

    ownerEmpty.style.display = "none";

    const renderTaskItem = (task) => {
        const status = String(task.status || "").toLowerCase();
        const statusLabel = status ? status.charAt(0).toUpperCase() + status.slice(1) : "In progress";
        const deal = findDealForTask(task);
        const dealLabel = deal ? (deal.name || deal.company || "View deal") : "General";
        const dealLink = deal && deal.id ? `./deal-details.html?id=${encodeURIComponent(deal.id)}` : "";
        const dashboardLink = deal && deal.fundraisingDashboardId
          ? `./investor-dashboard.html?dashboard=${encodeURIComponent(deal.fundraisingDashboardId)}`
          : "";
        const overdue = isTaskOverdue(task, today);
        const editable = !isAutoTask(task);

        return `
          <div class="task-item">
            <div class="task-title-row">
              <span class="task-title">${task.title || "Task"}</span>
              <div class="task-badges">
                <span class="badge ${statusClass(status)}">${statusLabel}</span>
                ${task.type ? `<span class="badge">${task.type}</span>` : ""}
              </div>
            </div>
            <div class="task-secondary">
              <span>${formatDate(task.dueDate)}${overdue ? ' · <span class="overdue-label">Overdue</span>' : ""}</span>
              <span>${deal ? (deal.company || "") : ""}</span>
              ${dealLink ? `<a class="task-link" href="${dealLink}">${dealLabel}</a>` : ""}
            </div>
            <div class="task-actions">
              ${editable ? `<button class="task-btn task-btn-edit" type="button" data-action="edit-task" data-task-id="${task.id}">Edit</button>` : `<span class="task-lock">Auto</span>`}
              ${dealLink ? `<a class="task-btn" href="${dealLink}">Open deal</a>` : ""}
              ${dashboardLink ? `<a class="task-btn" href="${dashboardLink}">Open dashboard</a>` : ""}
            </div>
            ${task.notes ? `<div class="task-notes">${task.notes}</div>` : ""}
          </div>
        `;
      };

    if (currentLayout === "list") {
      ownerTaskList.innerHTML = filtered.map(renderTaskItem).join("");
      return;
    }

    const groups = new Map();
    filtered.forEach((task) => {
      const deal = findDealForTask(task);
      const company = deal?.company || "General";
      if (!groups.has(company)) groups.set(company, []);
      groups.get(company).push(task);
    });

    const groupKeyFor = (name) => normalizeValue(name) || "general";

    const groupsHtml = [...groups.entries()]
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([company, companyTasks]) => {
        const doneCount = companyTasks.filter((task) => normalizeValue(task.status) === "done").length;
        const progress = companyTasks.length ? Math.round((doneCount / companyTasks.length) * 100) : 0;
        const items = companyTasks.map(renderTaskItem).join("");
        const groupKey = groupKeyFor(company);
        const isCollapsed = collapsedGroups.has(groupKey);

        return `
          <section class="company-group${isCollapsed ? " collapsed" : ""}" data-company-key="${groupKey}">
            <div class="company-header">
              <div>
                <div class="company-title company-title-toggle" data-company-key="${groupKey}" role="button" tabindex="0" aria-expanded="${isCollapsed ? "false" : "true"}">
                  ${company}
                </div>
                <div class="company-meta">${companyTasks.length} task${companyTasks.length === 1 ? "" : "s"} · ${doneCount} done</div>
              </div>
              <div class="company-progress">
                <div class="company-progress-top">
                  <span>Advancement</span>
                  <span>${progress}%</span>
                </div>
                <div class="company-progress-bar">
                  <div class="company-progress-fill" style="width:${progress}%;"></div>
                </div>
              </div>
            </div>
            <div class="company-task-list">${items}</div>
          </section>
        `;
      })
      .join("");

    ownerTaskList.innerHTML = groupsHtml;
  }

  filterButtons.forEach((button) => {
    button.addEventListener("click", () => {
      filterButtons.forEach((btn) => btn.classList.remove("active"));
      button.classList.add("active");
      currentView = button.getAttribute("data-view") || "all";
      renderTaskList();
    });
  });

  if (searchInput) {
    searchInput.addEventListener("input", renderTaskList);
  }

  ownerTaskList.addEventListener("click", (event) => {
    const editButton = event.target.closest('[data-action="edit-task"]');
    if (editButton) {
      const taskId = editButton.getAttribute("data-task-id");
      const taskIndex = allTasks.findIndex((task) => String(task.id) === String(taskId));
      if (taskIndex < 0) return;

      const originalTask = allTasks[taskIndex];
      if (isAutoTask(originalTask)) {
        window.alert("Auto tasks are synced from dashboard and cannot be edited here.");
        return;
      }
      openTaskEditor(originalTask);
      return;
    }

    const title = event.target.closest(".company-title-toggle");
    if (!title) return;
    const groupKey = title.getAttribute("data-company-key");
    const groupEl = ownerTaskList.querySelector(`.company-group[data-company-key="${groupKey}"]`);
    if (!groupEl || !groupKey) return;

    const isCollapsed = groupEl.classList.toggle("collapsed");
    if (isCollapsed) collapsedGroups.add(groupKey);
    else collapsedGroups.delete(groupKey);
    title.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  });

  ownerTaskList.addEventListener("keydown", (event) => {
    const title = event.target.closest(".company-title-toggle");
    if (!title) return;
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    title.click();
  });

  if (editModal) {
    editModal.addEventListener("click", (event) => {
      const closeTarget = event.target.closest('[data-action="close-task-edit"]');
      if (closeTarget) {
        closeTaskEditor();
      }
    });
  }

  if (editForm) {
    editForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!editingTaskId) return;

      const taskIndex = allTasks.findIndex((task) => String(task.id) === String(editingTaskId));
      if (taskIndex < 0) {
        closeTaskEditor();
        return;
      }

      const originalTask = allTasks[taskIndex];
      if (isAutoTask(originalTask)) {
        window.alert("Auto tasks are synced from dashboard and cannot be edited here.");
        closeTaskEditor();
        return;
      }

      const cleanTitle = (editTitleInput.value || "").trim();
      const cleanStatus = normalizeValue(editStatusInput.value || "in progress");
      const cleanDueDate = (editDueInput.value || "").trim();
      if (!cleanTitle) {
        window.alert("Task title cannot be empty.");
        return;
      }
      if (!["in progress", "waiting", "done"].includes(cleanStatus)) {
        window.alert("Invalid status. Use: in progress, waiting, or done.");
        return;
      }
      if (cleanDueDate) {
        const parsedDate = new Date(cleanDueDate);
        if (isNaN(parsedDate.getTime())) {
          window.alert("Invalid due date format.");
          return;
        }
      }

      allTasks[taskIndex] = {
        ...originalTask,
        title: cleanTitle,
        status: cleanStatus,
        dueDate: cleanDueDate,
        type: (editTypeInput.value || "").trim(),
        notes: (editNotesInput.value || "").trim(),
      };
      saveAllTasks(allTasks);
      closeTaskEditor();
      renderTaskList();
    });
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && editModal && editModal.classList.contains("open")) {
      closeTaskEditor();
    }
  });

  layoutButtons.forEach((button) => {
    button.addEventListener("click", () => {
      layoutButtons.forEach((btn) => btn.classList.remove("active"));
      button.classList.add("active");
      currentLayout = button.getAttribute("data-layout") || "grouped";
      renderTaskList();
    });
  });

  renderTaskList();
}

document.addEventListener("DOMContentLoaded", renderOwnerPage);
