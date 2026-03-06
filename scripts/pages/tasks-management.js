    const AppCore = window.AppCore;
    const TASKS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.tasks) || "owner_tasks_v1";
    let dealsData = [];
    let tasks = [];

    function loadDealsData() {
      dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
    }

    function loadTasksData() {
      tasks = AppCore ? AppCore.loadTasksData() : (Array.isArray(TASKS) ? JSON.parse(JSON.stringify(TASKS)) : []);
    }

    function saveTasksData() {
      if (AppCore) {
        AppCore.saveTasksData(tasks);
        return;
      }
      try {
        localStorage.setItem(TASKS_STORAGE_KEY, JSON.stringify(tasks));
      } catch (e) {
        console.warn("Failed to save tasks to storage", e);
      }
    }

    function statusClass(status) {
      const s = String(status || "").toLowerCase();
      if (s === "in progress") return "badge-status-in-progress";
      if (s === "waiting") return "badge-status-waiting";
      if (s === "done") return "badge-status-done";
      return "";
    }

    function formatDate(d) {
      if (!d) return "No due date";
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return d;
      return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
    }

    function normalizeValue(value) {
      if (AppCore) return AppCore.normalizeValue(value);
      return String(value || "").trim().toLowerCase();
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

    function renderTasks() {
      const grid = document.getElementById("owners-grid");
      const footerMeta = document.getElementById("footer-meta");
      const kpis = document.getElementById("management-kpis");
      const ownerSearch = document.getElementById("owner-search");
      const statusFilter = document.getElementById("status-filter");
      const ownerSort = document.getElementById("owner-sort");
      const groupBySelect = document.getElementById("group-by");
      const titleOnlyToggle = document.getElementById("title-only");
      if (!Array.isArray(tasks)) return;

      const groups = {};
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const searchTerm = normalizeValue(ownerSearch ? ownerSearch.value : "");
      const filterValue = statusFilter ? String(statusFilter.value || "all").toLowerCase() : "all";
      const groupBy = groupBySelect ? String(groupBySelect.value || "owner").toLowerCase() : "owner";
      const titleOnly = Boolean(titleOnlyToggle && titleOnlyToggle.checked);

      const isTaskOverdue = (task) => {
        const status = String(task.status || "").toLowerCase();
        if (!task.dueDate || status === "done") return false;
        const d = new Date(task.dueDate);
        if (isNaN(d.getTime())) return false;
        d.setHours(0, 0, 0, 0);
        return d < today;
      };

      const resolveGroupMeta = (task, deal) => {
        const owner = task.owner || "Unassigned";
        if (groupBy === "deal") {
          const label = deal ? (deal.name || deal.company || "Unnamed deal") : (task.dealId || "General");
          return {
            key: `deal:${normalizeValue(deal ? (deal.id || label) : label)}`,
            label,
            href: deal && deal.id ? `deal-details.html?id=${encodeURIComponent(deal.id)}` : ""
          };
        }

        if (groupBy === "status") {
          const status = String(task.status || "in progress").toLowerCase();
          const label = status.charAt(0).toUpperCase() + status.slice(1);
          return { key: `status:${status}`, label, href: "" };
        }

        if (groupBy === "type") {
          const type = String(task.type || "").trim() || "Unspecified";
          return { key: `type:${normalizeValue(type)}`, label: type, href: "" };
        }

        return {
          key: `owner:${normalizeValue(owner)}`,
          label: owner,
          href: `owner-tasks.html?owner=${encodeURIComponent(owner)}`
        };
      };

      tasks.forEach(task => {
        if (filterValue !== "all") {
          const taskStatus = String(task.status || "").toLowerCase();
          if (filterValue === "overdue") {
            if (!isTaskOverdue(task)) return;
          } else if (taskStatus !== filterValue) {
            return;
          }
        }

        const deal = findDealForTask(task);
        if (searchTerm) {
          const haystack = [
            task.owner,
            task.title,
            task.type,
            task.notes,
            deal && deal.name,
            deal && deal.company
          ].map(normalizeValue).join(" ");
          if (!haystack.includes(searchTerm)) return;
        }

        const groupMeta = resolveGroupMeta(task, deal);
        if (!groups[groupMeta.key]) {
          groups[groupMeta.key] = {
            label: groupMeta.label,
            href: groupMeta.href,
            entries: []
          };
        }
        groups[groupMeta.key].entries.push({ task, deal });
      });

      Object.keys(groups).forEach((groupKey) => {
        groups[groupKey].entries.sort((a, b) => {
          if (!a.task.dueDate && !b.task.dueDate) return 0;
          if (!a.task.dueDate) return 1;
          if (!b.task.dueDate) return -1;
          return new Date(a.task.dueDate) - new Date(b.task.dueDate);
        });
      });

      const groupItems = Object.values(groups);
      grid.innerHTML = "";

      let overdueTotal = 0;
      let waitingTotal = 0;
      let doneTotal = 0;
      let totalShown = 0;

      const renderedCards = groupItems.map((groupItem) => {
        const entries = groupItem.entries;
        const tasksInGroup = entries.map((entry) => entry.task);
        totalShown += tasksInGroup.length;

        const doneCount = tasksInGroup.filter(t => String(t.status || "").toLowerCase() === "done").length;
        const waitingCount = tasksInGroup.filter(t => String(t.status || "").toLowerCase() === "waiting").length;
        const openCount = tasksInGroup.length - doneCount;
        const overdueCount = tasksInGroup.filter(isTaskOverdue).length;
        const completionPct = tasksInGroup.length ? Math.round((doneCount / tasksInGroup.length) * 100) : 0;

        waitingTotal += waitingCount;
        doneTotal += doneCount;
        overdueTotal += overdueCount;

        const listItems = entries.map(({ task, deal }) => {
          const status = String(task.status || "").toLowerCase();
          const statusLabel = status ? status.charAt(0).toUpperCase() + status.slice(1) : "In progress";
          const dealLabel = deal ? (deal.name || deal.company || "View deal") : "General";
          const dealLink = deal && deal.id ? `./deal-details.html?id=${encodeURIComponent(deal.id)}` : "";
          const overdue = isTaskOverdue(task);

          if (titleOnly) {
            return `
              <li class="task-item task-item-title-only">
                <div class="task-title-row">
                  <span class="task-title">${task.title || "Task"}</span>
                  <a class="task-link" href="javascript:void(0)" onclick="deleteTask('${task.id}')">Delete</a>
                </div>
              </li>
            `;
          }

          return `
            <li class="task-item">
              <div class="task-title-row">
                <span class="task-title">${task.title || "Task"}</span>
                <div class="task-badges">
                  <span class="badge ${statusClass(status)}">${statusLabel}</span>
                  ${task.type ? `<span class="badge">${task.type}</span>` : ""}
                </div>
              </div>
              <div class="task-secondary">
                <span>${formatDate(task.dueDate)}${overdue ? ' · <span class="overdue-label">Overdue</span>' : ''}</span>
                <span>${deal ? (deal.company || "") : ""}</span>
                ${dealLink ? `<a class="task-link" href="${dealLink}">${dealLabel}</a>` : ""}
                <a class="task-link" href="javascript:void(0)" onclick="deleteTask('${task.id}')">Delete</a>
              </div>
              ${task.notes ? `<div class="task-notes">${task.notes}</div>` : ""}
            </li>
          `;
        }).join("");

        const titleContent = groupItem.href
          ? `<a class="owner-name" href="${groupItem.href}">${groupItem.label}</a>`
          : `<span class="owner-name">${groupItem.label}</span>`;

        return {
          label: groupItem.label,
          tasksCount: tasksInGroup.length,
          doneCount,
          waitingCount,
          openCount,
          overdueCount,
          completionPct,
          html: `
          <div class="owner-header">
            <div>
              ${titleContent}
              <div class="owner-meta">${tasksInGroup.length} task${tasksInGroup.length === 1 ? "" : "s"} · ${openCount} open</div>
              <div class="owner-progress">
                <div class="owner-progress-row">
                  <span>Completion</span>
                  <span>${completionPct}%</span>
                </div>
                <div class="owner-progress-bar">
                  <div class="owner-progress-fill" style="width:${completionPct}%;"></div>
                </div>
              </div>
            </div>
            <div class="owner-chips">
              <span class="owner-chip">${openCount} open</span>
              <span class="owner-chip">${waitingCount} waiting</span>
              <span class="owner-chip">${doneCount} done</span>
            </div>
          </div>
          <ul class="task-list">
            ${listItems}
          </ul>
        `
        };
      });

      const sortValue = ownerSort ? String(ownerSort.value || "workload").toLowerCase() : "workload";
      renderedCards.sort((a, b) => {
        if (sortValue === "name") return a.label.localeCompare(b.label);
        if (sortValue === "overdue") return b.overdueCount - a.overdueCount || b.openCount - a.openCount || a.label.localeCompare(b.label);
        if (sortValue === "completion") return a.completionPct - b.completionPct || b.openCount - a.openCount || a.label.localeCompare(b.label);
        return b.openCount - a.openCount || b.tasksCount - a.tasksCount || a.label.localeCompare(b.label);
      });

      renderedCards.forEach((cardData) => {
        const card = document.createElement("div");
        card.className = "owner-card";
        card.innerHTML = cardData.html;
        grid.appendChild(card);
      });

      if (!renderedCards.length) {
        grid.innerHTML = `<div class="empty-state">No tasks found for this filter. Try another status or search term.</div>`;
      }

      const openTotal = totalShown - doneTotal;
      footerMeta.textContent =
        `${openTotal} open · ${waitingTotal} waiting · ${overdueTotal} overdue · ${totalShown} shown.`;

      if (kpis) {
        kpis.innerHTML = `
          <div class="management-kpi">
            <div class="management-kpi-label">Groups</div>
            <div class="management-kpi-value">${groupItems.length}</div>
          </div>
          <div class="management-kpi">
            <div class="management-kpi-label">Tasks Shown</div>
            <div class="management-kpi-value">${totalShown}</div>
          </div>
          <div class="management-kpi">
            <div class="management-kpi-label">Open</div>
            <div class="management-kpi-value">${openTotal}</div>
          </div>
          <div class="management-kpi">
            <div class="management-kpi-label">Waiting</div>
            <div class="management-kpi-value">${waitingTotal}</div>
          </div>
          <div class="management-kpi">
            <div class="management-kpi-label">Overdue</div>
            <div class="management-kpi-value">${overdueTotal}</div>
          </div>
          <div class="management-kpi">
            <div class="management-kpi-label">Done</div>
            <div class="management-kpi-value">${doneTotal}</div>
          </div>
        `;
      }
    }

    function setupForm() {
      const form = document.getElementById("task-form");
      const ownerInput = document.getElementById("task-owner");
      const dealSelect = document.getElementById("task-deal");
      const titleInput = document.getElementById("task-title");
      const typeInput = document.getElementById("task-type");
      const statusSelect = document.getElementById("task-status");
      const dueInput = document.getElementById("task-due");
      const notesInput = document.getElementById("task-notes");

      if (Array.isArray(dealsData)) {
        dealsData.forEach(d => {
          const opt = document.createElement("option");
          opt.value = d.id;
          opt.textContent = d.name || d.company || d.id;
          dealSelect.appendChild(opt);
        });
      }

      form.addEventListener("submit", (e) => {
        e.preventDefault();
        const owner = ownerInput.value.trim();
        const title = titleInput.value.trim();
        if (!owner || !title) return;

        const newTask = {
          id: `t-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          owner,
          dealId: dealSelect.value || null,
          title,
          type: typeInput.value.trim() || "",
          status: statusSelect.value || "in progress",
          dueDate: dueInput.value || "",
          notes: notesInput.value.trim() || "",
        };

        tasks.push(newTask);
        saveTasksData();
        renderTasks();

        ownerInput.value = "";
        titleInput.value = "";
        typeInput.value = "";
        notesInput.value = "";
      });
    }

    function deleteTask(taskId) {
      tasks = tasks.filter(t => t.id !== taskId);
      saveTasksData();
      renderTasks();
    }

    document.addEventListener("DOMContentLoaded", () => {
      loadDealsData();
      loadTasksData();
      setupForm();
      const ownerSearch = document.getElementById("owner-search");
      const statusFilter = document.getElementById("status-filter");
      const ownerSort = document.getElementById("owner-sort");
      const groupBySelect = document.getElementById("group-by");
      const titleOnlyToggle = document.getElementById("title-only");
      if (ownerSearch) ownerSearch.addEventListener("input", renderTasks);
      if (statusFilter) statusFilter.addEventListener("change", renderTasks);
      if (ownerSort) ownerSort.addEventListener("change", renderTasks);
      if (groupBySelect) groupBySelect.addEventListener("change", renderTasks);
      if (titleOnlyToggle) titleOnlyToggle.addEventListener("change", renderTasks);
      renderTasks();
    });
