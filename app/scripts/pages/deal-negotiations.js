(function initDealNegotiations() {
  const AppCore = window.AppCore;
  const NEGOTIATION_STATUS_LABELS = {
    reviewing: "Reviewing",
    engagement_to_send: "Engagement letter to send",
    engagement_sent: "Engagement letter sent",
    signed_back: "Signed back",
  };

  let dealsData = [];
  let tasksData = [];
  let negotiationFilters = {
    search: "",
    status: "all",
    owner: "all",
    readiness: "all",
  };

  function normalizeValue(value) {
    if (AppCore && typeof AppCore.normalizeValue === "function") {
      return AppCore.normalizeValue(value);
    }
    return String(value || "").trim().toLowerCase();
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function buildPageUrl(pageId, params) {
    if (AppCore && typeof AppCore.getPageUrl === "function") {
      return AppCore.getPageUrl(pageId, params);
    }
    if (window.PlutusAppConfig && typeof window.PlutusAppConfig.buildPageHref === "function") {
      return window.PlutusAppConfig.buildPageHref(pageId, params);
    }
    const query = new URLSearchParams(params || {});
    const queryString = query.toString();
    return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
  }

  function loadDealsData() {
    dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(window.DEALS) ? JSON.parse(JSON.stringify(window.DEALS)) : []);
  }

  function loadTasksData() {
    tasksData = AppCore ? AppCore.loadTasksData() : (Array.isArray(window.TASKS) ? JSON.parse(JSON.stringify(window.TASKS)) : []);
  }

  function normalizeDealLifecycleStatus(value) {
    const normalized = normalizeValue(value);
    if (normalized === "finished") return "finished";
    if (normalized === "closed") return "closed";
    return "active";
  }

  function normalizeDealPipelineStatus(value) {
    return normalizeValue(value) === "negotiation" ? "negotiation" : "pipeline";
  }

  function normalizeNegotiationStatus(value) {
    const normalized = normalizeValue(value);
    if (normalized === "engagement_to_send") return "engagement_to_send";
    if (normalized === "engagement_sent") return "engagement_sent";
    if (normalized === "signed_back") return "signed_back";
    return "reviewing";
  }

  function isNegotiationDeal(deal) {
    if (!deal || typeof deal !== "object") return false;
    if (normalizeDealLifecycleStatus(deal.lifecycleStatus || deal.dealStatus) !== "active") return false;
    return normalizeDealPipelineStatus(deal.pipelineStatus) === "negotiation";
  }

  function buildUniqueList(values) {
    const seen = new Set();
    return (Array.isArray(values) ? values : []).filter((entry) => {
      const value = String(entry || "").trim();
      const key = normalizeValue(value);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function getDealOwners(deal) {
    const subOwners = Array.isArray(deal && deal.subOwners)
      ? deal.subOwners
      : typeof (deal && deal.subOwners) === "string"
        ? deal.subOwners.split(/[\n,;]+/)
        : [];
    return buildUniqueList([
      deal && (deal.seniorOwner || deal.owner),
      deal && deal.juniorOwner,
      ...subOwners,
    ]);
  }

  function normalizeDealContacts(deal) {
    const source = deal && Array.isArray(deal.contacts) ? deal.contacts : [];
    const contacts = source
      .map((entry) => {
        if (!entry || typeof entry !== "object") return null;
        const name = String(entry.name || "").trim();
        const title = String(entry.title || "").trim();
        const email = String(entry.email || "").trim();
        const isPrimary = Boolean(entry.isPrimary);
        if (!name && !title && !email) return null;
        return { name, title, email, isPrimary };
      })
      .filter(Boolean);

    if (!contacts.length) {
      const fallbackName = String((deal && deal.mainContactName) || "").trim();
      const fallbackTitle = String((deal && deal.mainContactTitle) || "").trim();
      const fallbackEmail = String((deal && (deal.mainContactEmail || deal.email)) || "").trim();
      if (fallbackName || fallbackTitle || fallbackEmail) {
        contacts.push({
          name: fallbackName,
          title: fallbackTitle,
          email: fallbackEmail,
          isPrimary: true,
        });
      }
    }

    if (contacts.length && !contacts.some((entry) => entry.isPrimary)) {
      contacts[0].isPrimary = true;
    }

    return contacts;
  }

  function getPrimaryContact(deal) {
    const contacts = normalizeDealContacts(deal);
    return contacts.find((entry) => entry.isPrimary) || contacts[0] || null;
  }

  function normalizeDealLegalLinks(deal) {
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
          title: title || `Legal link ${index + 1}`,
          url,
        };
      })
      .filter(Boolean);
  }

  function toSafeExternalUrl(value) {
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

  function isEngagementLetterLink(link) {
    const haystack = `${link && link.title || ""} ${link && link.url || ""}`;
    return normalizeValue(haystack).includes("engagement");
  }

  function isTaskDone(task) {
    return normalizeValue(task && task.status) === "done";
  }

  function isNegotiationTask(task) {
    const haystack = `${task && task.title || ""} ${task && task.type || ""} ${task && task.notes || ""}`;
    const normalized = normalizeValue(haystack);
    return (
      normalized.includes("engagement") ||
      normalized.includes("letter") ||
      normalized.includes("sign") ||
      normalized.includes("signature") ||
      normalized.includes("contract") ||
      normalized.includes("review") ||
      normalized.includes("legal")
    );
  }

  function formatDate(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
  }

  function buildTaskSummary(deal) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const tasks = (Array.isArray(tasksData) ? tasksData : [])
      .filter((task) => normalizeValue(task && task.dealId) === normalizeValue(deal && deal.id))
      .filter(isNegotiationTask);

    const done = tasks.filter(isTaskDone).length;
    const overdue = tasks.filter((task) => {
      if (isTaskDone(task) || !task || !task.dueDate) return false;
      const due = new Date(task.dueDate);
      if (Number.isNaN(due.getTime())) return false;
      due.setHours(0, 0, 0, 0);
      return due < today;
    }).length;

    const nextDue = tasks
      .filter((task) => !isTaskDone(task) && task && task.dueDate)
      .map((task) => {
        const due = new Date(task.dueDate);
        return Number.isNaN(due.getTime()) ? null : { due, raw: task.dueDate };
      })
      .filter(Boolean)
      .sort((left, right) => left.due.getTime() - right.due.getTime())[0];

    return {
      total: tasks.length,
      done,
      open: Math.max(tasks.length - done, 0),
      overdue,
      nextDue: nextDue ? nextDue.raw : "",
    };
  }

  function buildNegotiationSnapshot(deal) {
    const owners = getDealOwners(deal);
    const primaryContact = getPrimaryContact(deal);
    const legalLinks = normalizeDealLegalLinks(deal).filter((entry) => Boolean(toSafeExternalUrl(entry.url)));
    const engagementLetters = legalLinks.filter(isEngagementLetterLink);
    const status = normalizeNegotiationStatus(deal && deal.negotiationStatus);
    const taskSummary = buildTaskSummary(deal);

    let readinessTone = "working";
    let readinessLabel = "Under review";

    if (taskSummary.overdue > 0) {
      readinessTone = "warning";
      readinessLabel = `${taskSummary.overdue} overdue task${taskSummary.overdue === 1 ? "" : "s"}`;
    } else if (status !== "reviewing" && engagementLetters.length === 0) {
      readinessTone = "warning";
      readinessLabel = "Engagement letter link missing";
    } else if (status === "signed_back") {
      readinessTone = "ready";
      readinessLabel = "Ready for pipeline";
    } else if (status === "engagement_sent") {
      readinessLabel = "Waiting for signed letter";
    } else if (status === "engagement_to_send") {
      readinessLabel = "Letter ready to send";
    }

    return {
      deal,
      status,
      owners,
      primaryContact,
      legalLinks,
      engagementLetters,
      taskSummary,
      readinessTone,
      readinessLabel,
    };
  }

  function getReadinessRank(snapshot) {
    if (snapshot.readinessTone === "warning") return 0;
    if (snapshot.readinessTone === "working") return 1;
    return 2;
  }

  function compareSnapshots(left, right) {
    const readinessRank = getReadinessRank(left) - getReadinessRank(right);
    if (readinessRank) return readinessRank;
    if (left.taskSummary.overdue !== right.taskSummary.overdue) {
      return right.taskSummary.overdue - left.taskSummary.overdue;
    }
    const leftLabel = normalizeValue(left.deal && (left.deal.company || left.deal.name || left.deal.id));
    const rightLabel = normalizeValue(right.deal && (right.deal.company || right.deal.name || right.deal.id));
    return leftLabel.localeCompare(rightLabel);
  }

  function buildAllSnapshots() {
    return (Array.isArray(dealsData) ? dealsData : [])
      .filter(isNegotiationDeal)
      .map(buildNegotiationSnapshot)
      .sort(compareSnapshots);
  }

  function matchesReadinessFilter(snapshot, filterValue) {
    const filter = normalizeValue(filterValue);
    if (!filter || filter === "all") return true;
    if (filter === "attention") return snapshot.readinessTone === "warning";
    if (filter === "ready") {
      return snapshot.status === "signed_back"
        && snapshot.engagementLetters.length > 0
        && snapshot.taskSummary.overdue === 0;
    }
    if (filter === "missing-letter") return snapshot.engagementLetters.length === 0;
    if (filter === "overdue-tasks") return snapshot.taskSummary.overdue > 0;
    return true;
  }

  function getVisibleSnapshots() {
    const query = normalizeValue(negotiationFilters.search);
    return buildAllSnapshots()
      .filter((snapshot) => {
        if (negotiationFilters.status !== "all" && snapshot.status !== normalizeValue(negotiationFilters.status)) {
          return false;
        }
        if (negotiationFilters.owner !== "all") {
          const hasOwner = snapshot.owners.some((entry) => normalizeValue(entry) === normalizeValue(negotiationFilters.owner));
          if (!hasOwner) return false;
        }
        if (!matchesReadinessFilter(snapshot, negotiationFilters.readiness)) return false;

        if (!query) return true;
        const haystack = [
          snapshot.deal && snapshot.deal.name,
          snapshot.deal && snapshot.deal.company,
          snapshot.deal && snapshot.deal.id,
          snapshot.primaryContact && snapshot.primaryContact.name,
          snapshot.primaryContact && snapshot.primaryContact.email,
          ...snapshot.owners,
        ].map((value) => normalizeValue(value)).join(" ");
        return haystack.includes(query);
      })
      .sort(compareSnapshots);
  }

  function populateOwnerFilter(allSnapshots) {
    const select = document.getElementById("negotiation-owner-filter");
    if (!select) return;

    const currentValue = String(negotiationFilters.owner || "all").trim() || "all";
    const owners = buildUniqueList(
      (Array.isArray(allSnapshots) ? allSnapshots : []).flatMap((snapshot) => snapshot.owners),
    ).sort((left, right) => left.localeCompare(right));

    select.innerHTML = '<option value="all">All owners</option>';
    owners.forEach((owner) => {
      const option = document.createElement("option");
      option.value = owner;
      option.textContent = owner;
      select.appendChild(option);
    });

    const hasCurrent = currentValue === "all" || owners.some((owner) => normalizeValue(owner) === normalizeValue(currentValue));
    if (!hasCurrent) negotiationFilters.owner = "all";
    select.value = hasCurrent ? currentValue : "all";
  }

  function renderMetaRow(allSnapshots, visibleSnapshots) {
    const row = document.getElementById("negotiation-meta-row");
    if (!row) return;

    const snapshots = Array.isArray(visibleSnapshots) ? visibleSnapshots : [];
    const reviewingCount = snapshots.filter((snapshot) => snapshot.status === "reviewing").length;
    const toSendCount = snapshots.filter((snapshot) => snapshot.status === "engagement_to_send").length;
    const sentCount = snapshots.filter((snapshot) => snapshot.status === "engagement_sent").length;
    const signedCount = snapshots.filter((snapshot) => snapshot.status === "signed_back").length;
    const attentionCount = snapshots.filter((snapshot) => snapshot.readinessTone === "warning").length;

    row.innerHTML = [
      `<div class="chip"><strong>${snapshots.length}</strong> shown</div>`,
      `<div class="chip"><strong>${Array.isArray(allSnapshots) ? allSnapshots.length : 0}</strong> negotiation deals total</div>`,
      `<div class="chip"><strong>${reviewingCount}</strong> reviewing</div>`,
      `<div class="chip"><strong>${toSendCount}</strong> letters to send</div>`,
      `<div class="chip"><strong>${sentCount}</strong> letters sent</div>`,
      `<div class="chip"><strong>${signedCount}</strong> signed back</div>`,
      `<div class="chip"><strong>${attentionCount}</strong> need attention</div>`,
    ].join("");
  }

  function renderSummary(allSnapshots, visibleSnapshots) {
    const summary = document.getElementById("negotiation-filter-summary");
    const footerMeta = document.getElementById("negotiation-footer-meta");
    if (!summary || !footerMeta) return;

    const total = Array.isArray(allSnapshots) ? allSnapshots.length : 0;
    const visible = Array.isArray(visibleSnapshots) ? visibleSnapshots.length : 0;
    summary.textContent = total === visible
      ? `Showing all ${visible} negotiation deal${visible === 1 ? "" : "s"}.`
      : `Showing ${visible} of ${total} negotiation deal${total === 1 ? "" : "s"}.`;
    footerMeta.textContent = total
      ? `${total} active pre-pipeline deal${total === 1 ? "" : "s"} currently being reviewed or negotiated.`
      : "No active negotiation deals at the moment.";
  }

  function renderStatusCell(snapshot) {
    const statusClass = snapshot.status.replace(/_/g, "-");
    return `<span class="negotiation-status-pill is-${statusClass}">${escapeHtml(NEGOTIATION_STATUS_LABELS[snapshot.status] || "Reviewing")}</span>`;
  }

  function renderContactCell(snapshot) {
    const contact = snapshot.primaryContact;
    if (!contact) {
      return '<div class="negotiation-contact-stack"><span class="negotiation-subline">No main contact saved yet.</span></div>';
    }

    const titleLine = [contact.title, contact.email].filter(Boolean).join(" · ");
    return `
      <div class="negotiation-contact-stack">
        <span class="negotiation-contact-name">${escapeHtml(contact.name || contact.email || "Main contact")}</span>
        <span class="negotiation-subline">${escapeHtml(titleLine || "Contact saved on the deal.")}</span>
        ${contact.email ? `<a class="negotiation-contact-email" href="mailto:${escapeHtml(contact.email)}">${escapeHtml(contact.email)}</a>` : ""}
      </div>
    `;
  }

  function renderLetterCell(snapshot) {
    const linked = snapshot.engagementLetters.length > 0;
    return `
      <div class="negotiation-letter-cell">
        <div class="negotiation-pill-row">
          <span class="negotiation-letter-pill ${linked ? "is-linked" : "is-missing"}">${linked ? `${snapshot.engagementLetters.length} linked` : "Not linked yet"}</span>
        </div>
        <div class="negotiation-subline">${snapshot.legalLinks.length} legal link${snapshot.legalLinks.length === 1 ? "" : "s"} saved on this deal.</div>
      </div>
    `;
  }

  function renderTaskCell(snapshot) {
    const completion = snapshot.taskSummary.total
      ? Math.round((snapshot.taskSummary.done / snapshot.taskSummary.total) * 100)
      : 0;
    const nextDueLabel = snapshot.taskSummary.nextDue ? formatDate(snapshot.taskSummary.nextDue) : "No due date";
    return `
      <div class="negotiation-task-cell">
        <div class="negotiation-task-progress">
          <div class="negotiation-task-topline">
            <span>${snapshot.taskSummary.done}/${snapshot.taskSummary.total} done</span>
            <span>${completion}%</span>
          </div>
          <div class="negotiation-task-bar">
            <div class="negotiation-task-fill" style="width:${completion}%;"></div>
          </div>
        </div>
        <div class="negotiation-task-meta">
          ${snapshot.taskSummary.total
            ? `${snapshot.taskSummary.open} open · ${snapshot.taskSummary.overdue} overdue · next due ${escapeHtml(nextDueLabel)}`
            : "No negotiation tasks linked yet."}
        </div>
      </div>
    `;
  }

  function renderRow(snapshot) {
    const deal = snapshot.deal;
    const dealHref = buildPageUrl("deal-details", { id: deal.id });
    const legalHref = buildPageUrl("legal-management", { id: deal.id });
    const ownerHtml = snapshot.owners.length
      ? snapshot.owners.map((owner) => `<span class="negotiation-owner-pill">${escapeHtml(owner)}</span>`).join("")
      : '<span class="negotiation-subline">No owners assigned</span>';

    return `
      <tr>
        <td class="name-cell">
          <div class="negotiation-deal-cell">
            <a class="negotiation-deal-link" href="${dealHref}">${escapeHtml(String(deal.name || "Untitled deal"))}</a>
            <div class="negotiation-subline">${escapeHtml([deal.company || "Company", deal.id || ""].filter(Boolean).join(" · "))}</div>
            <div class="negotiation-subline">${escapeHtml([deal.fundingStage, deal.location].filter(Boolean).join(" · ") || "No funding stage or location saved yet.")}</div>
          </div>
        </td>
        <td>${renderStatusCell(snapshot)}</td>
        <td><div class="negotiation-owner-list">${ownerHtml}</div></td>
        <td>${renderContactCell(snapshot)}</td>
        <td>${renderLetterCell(snapshot)}</td>
        <td>${renderTaskCell(snapshot)}</td>
        <td><span class="negotiation-readiness-pill is-${snapshot.readinessTone}">${escapeHtml(snapshot.readinessLabel)}</span></td>
        <td>
          <div class="negotiation-action-cluster">
            <a class="action-link" href="${dealHref}">Open deal</a>
            <a class="action-link" href="${legalHref}">Legal</a>
          </div>
        </td>
      </tr>
    `;
  }

  function renderTable(visibleSnapshots) {
    const body = document.getElementById("negotiation-body");
    if (!body) return;

    if (!visibleSnapshots.length) {
      body.innerHTML = '<tr><td colspan="8">No active negotiation deals match this view.</td></tr>';
      return;
    }

    body.innerHTML = visibleSnapshots.map(renderRow).join("");
  }

  function applyFiltersToUi() {
    const searchInput = document.getElementById("negotiation-search");
    const statusFilter = document.getElementById("negotiation-status-filter");
    const readinessFilter = document.getElementById("negotiation-readiness-filter");

    if (searchInput) searchInput.value = negotiationFilters.search;
    if (statusFilter) statusFilter.value = negotiationFilters.status;
    if (readinessFilter) readinessFilter.value = negotiationFilters.readiness;
  }

  function renderNegotiationsPage() {
    const allSnapshots = buildAllSnapshots();
    populateOwnerFilter(allSnapshots);
    applyFiltersToUi();
    const visibleSnapshots = getVisibleSnapshots();
    renderMetaRow(allSnapshots, visibleSnapshots);
    renderSummary(allSnapshots, visibleSnapshots);
    renderTable(visibleSnapshots);
  }

  function setupFilters() {
    const searchInput = document.getElementById("negotiation-search");
    const statusFilter = document.getElementById("negotiation-status-filter");
    const ownerFilter = document.getElementById("negotiation-owner-filter");
    const readinessFilter = document.getElementById("negotiation-readiness-filter");
    const resetBtn = document.getElementById("btn-reset-negotiation-filters");

    if (searchInput) {
      searchInput.addEventListener("input", (event) => {
        negotiationFilters.search = String(event.target && event.target.value || "");
        renderNegotiationsPage();
      });
    }

    [statusFilter, ownerFilter, readinessFilter].forEach((control) => {
      if (!control) return;
      control.addEventListener("change", (event) => {
        const key = String(event.target && event.target.id || "");
        const value = String(event.target && event.target.value || "all");
        if (key === "negotiation-status-filter") negotiationFilters.status = value;
        if (key === "negotiation-owner-filter") negotiationFilters.owner = value;
        if (key === "negotiation-readiness-filter") negotiationFilters.readiness = value;
        renderNegotiationsPage();
      });
    });

    if (resetBtn) {
      resetBtn.addEventListener("click", () => {
        negotiationFilters = {
          search: "",
          status: "all",
          owner: "all",
          readiness: "all",
        };
        renderNegotiationsPage();
      });
    }
  }

  function setupLiveUpdates() {
    if (!AppCore) return;

    window.addEventListener("appcore:deals-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.deals)) {
        dealsData = event.detail.deals.slice();
        renderNegotiationsPage();
      }
    });

    window.addEventListener("appcore:tasks-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.tasks)) {
        tasksData = event.detail.tasks.slice();
        renderNegotiationsPage();
      }
    });
  }

  function initializePage() {
    loadDealsData();
    loadTasksData();
    setupFilters();
    setupLiveUpdates();
    renderNegotiationsPage();
  }

  document.addEventListener("DOMContentLoaded", initializePage);
})();
