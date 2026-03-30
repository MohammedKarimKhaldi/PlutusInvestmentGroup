(function initOutlookInvestorSyncPage() {
  const AppCore = window.AppCore;
  const CONTACT_STATUS_OPTIONS = [
    { value: "not-contacted", label: "Not contacted" },
    { value: "queued", label: "Queued" },
    { value: "contacted", label: "Contacted" },
    { value: "follow-up", label: "Follow up" },
    { value: "replied", label: "Replied" },
    { value: "invested", label: "Invested" },
    { value: "passed", label: "Passed" },
  ];
  const CONTACT_STATUS_LABELS = CONTACT_STATUS_OPTIONS.reduce((accumulator, option) => {
    accumulator[option.value] = option.label;
    return accumulator;
  }, {});
  const state = {
    deals: [],
    messages: [],
    selectedDealId: "",
    selectedMessageId: "",
    selectedRecipientEmails: new Set(),
    manualRecipients: [],
    messageFilter: "",
    hasSearchedMessages: false,
    deviceCodeState: null,
    deviceCodeTimer: null,
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

  function formatDateTime(value) {
    if (!value) return "Unknown time";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleString("en-GB", {
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function setStatus(elementId, text) {
    const element = document.getElementById(elementId);
    if (element) element.textContent = text;
  }

  function clearDeviceCodeTimer() {
    if (state.deviceCodeTimer) {
      window.clearTimeout(state.deviceCodeTimer);
      state.deviceCodeTimer = null;
    }
  }

  function setDeviceCodeBox(visible, payload) {
    const box = document.getElementById("device-code-box");
    if (!box) return;
    box.hidden = !visible;
    if (!visible) return;

    const codeEl = document.getElementById("device-code-value");
    const linkEl = document.getElementById("device-code-link");
    const messageEl = document.getElementById("device-code-message");
    const statusEl = document.getElementById("device-code-status");

    if (codeEl) codeEl.textContent = payload && payload.userCode ? payload.userCode : "—";
    if (linkEl && payload && payload.verificationUri) {
      linkEl.href = payload.verificationUri;
      linkEl.textContent = payload.verificationUri;
    }
    if (messageEl) {
      messageEl.textContent = payload && payload.message
        ? payload.message
        : "Follow the prompt to finish signing in.";
    }
    if (statusEl) statusEl.textContent = "Waiting for authorization…";
  }

  function getRequestedDealId() {
    const params = new URLSearchParams(window.location.search);
    return String(params.get("id") || "").trim();
  }

  function updateSelectedDealInUrl() {
    const url = new URL(window.location.href);
    if (state.selectedDealId) {
      url.searchParams.set("id", state.selectedDealId);
    } else {
      url.searchParams.delete("id");
    }
    window.history.replaceState({}, "", url.toString());
  }

  function loadDealsData() {
    const loadedDeals = AppCore && typeof AppCore.loadDealsData === "function"
      ? AppCore.loadDealsData()
      : (Array.isArray(window.DEALS) ? JSON.parse(JSON.stringify(window.DEALS)) : []);
    state.deals = sortDealsByRetainerState(loadedDeals);
  }

  function saveDealsData() {
    if (AppCore && typeof AppCore.saveDealsData === "function") {
      return AppCore.saveDealsData(state.deals);
    }
    return Promise.resolve();
  }

  function getSelectedDeal() {
    return state.deals.find((deal) => normalizeValue(deal && deal.id) === normalizeValue(state.selectedDealId)) || null;
  }

  function getDealRetainerState(deal) {
    if (AppCore && typeof AppCore.getDealRetainerState === "function") {
      return AppCore.getDealRetainerState(deal);
    }
    const rawValue = String(deal && (deal.retainerMonthly != null ? deal.retainerMonthly : deal && deal.Retainer) || "").trim();
    const cleaned = rawValue.replace(/,/g, "").replace(/[^0-9.-]/g, "");
    const amount = Number(cleaned);
    const hasRetainer = Number.isFinite(amount) && amount > 0;
    return {
      rawValue,
      amount: hasRetainer ? amount : 0,
      hasRetainer,
      bucket: hasRetainer ? "with-retainer" : "no-retainer",
      label: hasRetainer ? "With retainer" : "0 / no retainer",
    };
  }

  function sortDealsByRetainerState(deals) {
    if (AppCore && typeof AppCore.sortDealsByRetainerState === "function") {
      return AppCore.sortDealsByRetainerState(
        deals,
        (left, right) => normalizeValue(left && (left.company || left.name || left.id)).localeCompare(
          normalizeValue(right && (right.company || right.name || right.id)),
        ),
      );
    }
    return (Array.isArray(deals) ? deals.slice() : []).sort((left, right) => {
      const leftState = getDealRetainerState(left);
      const rightState = getDealRetainerState(right);
      if (leftState.hasRetainer !== rightState.hasRetainer) {
        return leftState.hasRetainer ? -1 : 1;
      }
      return normalizeValue(left && (left.company || left.name || left.id)).localeCompare(
        normalizeValue(right && (right.company || right.name || right.id)),
      );
    });
  }

  function normalizeContactStatus(value) {
    const normalized = normalizeValue(value);
    return CONTACT_STATUS_LABELS[normalized] ? normalized : "not-contacted";
  }

  function normalizeDealInvestorContacts(deal) {
    const source = deal && Array.isArray(deal.investorContacts) ? deal.investorContacts : [];
    const normalized = source
      .map((entry, index) => {
        if (!entry || typeof entry !== "object") return null;
        const email = String(entry.email || "").trim();
        if (!email) return null;
        return {
          id: String(entry.id || `investor-${index}`).trim(),
          name: String(entry.name || "").trim(),
          email,
          contactStatus: normalizeContactStatus(entry.contactStatus),
          source: String(entry.source || "").trim(),
          sourceMessageId: String(entry.sourceMessageId || "").trim(),
          sourceMessageSubject: String(entry.sourceMessageSubject || "").trim(),
          sourceMessageWebLink: String(entry.sourceMessageWebLink || "").trim(),
          sourceReceivedAt: String(entry.sourceReceivedAt || "").trim(),
          sourceRecipientType: String(entry.sourceRecipientType || "").trim(),
          sourceFromName: String(entry.sourceFromName || "").trim(),
          sourceFromEmail: String(entry.sourceFromEmail || "").trim(),
          notes: String(entry.notes || "").trim(),
          addedAt: String(entry.addedAt || "").trim(),
          updatedAt: String(entry.updatedAt || "").trim(),
        };
      })
      .filter(Boolean)
      .sort((left, right) => {
        const leftValue = Date.parse(left.updatedAt || left.addedAt || "") || 0;
        const rightValue = Date.parse(right.updatedAt || right.addedAt || "") || 0;
        if (leftValue !== rightValue) return rightValue - leftValue;
        return String(left.name || left.email).localeCompare(String(right.name || right.email));
      });

    if (deal && typeof deal === "object") {
      deal.investorContacts = normalized;
    }

    return normalized;
  }

  function getStatusBadgeClass(status) {
    return `investor-status-badge status-${normalizeContactStatus(status)}`;
  }

  function getSelectedMessage() {
    return state.messages.find((message) => String(message.id || "").trim() === String(state.selectedMessageId || "").trim()) || null;
  }

  function getFilteredMessages() {
    return Array.isArray(state.messages) ? state.messages.slice() : [];
  }

  function buildMessageOptionLabel(message) {
    const subject = String(message && message.subject || "").trim() || "(No subject)";
    const sender = String(
      message && message.from && (message.from.name || message.from.address) || "Unknown sender",
    ).trim();
    const receivedAt = formatDateTime(message && message.receivedDateTime);
    return `${subject} · ${sender} · ${receivedAt}`;
  }

  function refreshSelectedRecipientsForCurrentMessage() {
    const selectedEmails = new Set(
      state.manualRecipients
        .map((recipient) => normalizeValue(recipient.email))
        .filter(Boolean),
    );
    getMessageRecipients(getSelectedMessage()).forEach((recipient) => {
      const key = normalizeValue(recipient.email);
      if (key) selectedEmails.add(key);
    });
    state.selectedRecipientEmails = selectedEmails;
  }

  function ensureSelectedMessageIsVisible(filteredMessages) {
    const visible = Array.isArray(filteredMessages) ? filteredMessages : [];
    if (!visible.length) {
      state.selectedMessageId = "";
      refreshSelectedRecipientsForCurrentMessage();
      return;
    }

    const selectedStillVisible = visible.some(
      (message) => String(message.id || "").trim() === String(state.selectedMessageId || "").trim(),
    );
    if (selectedStillVisible) return;

    state.selectedMessageId = String(visible[0].id || "").trim();
    refreshSelectedRecipientsForCurrentMessage();
  }

  function getMessageRecipients(message) {
    const deduped = new Map();
    [
      ...((message && Array.isArray(message.toRecipients)) ? message.toRecipients.map((entry) => ({ ...entry, kind: "to" })) : []),
      ...((message && Array.isArray(message.ccRecipients)) ? message.ccRecipients.map((entry) => ({ ...entry, kind: "cc" })) : []),
    ].forEach((recipient) => {
      const email = String(recipient && recipient.address || "").trim();
      const key = normalizeValue(email);
      if (!key) return;
      const existing = deduped.get(key);
      deduped.set(key, {
        name: String(recipient && recipient.name || (existing && existing.name) || "").trim(),
        email,
        kind: existing && existing.kind === "to" ? "to" : (recipient.kind || "to"),
      });
    });
    return Array.from(deduped.values());
  }

  function getAllRecipientCandidates() {
    const recipients = new Map();

    state.manualRecipients.forEach((recipient) => {
      const key = normalizeValue(recipient.email);
      if (!key) return;
      recipients.set(key, {
        name: String(recipient.name || "").trim(),
        email: String(recipient.email || "").trim(),
        kind: "manual",
      });
    });

    getMessageRecipients(getSelectedMessage()).forEach((recipient) => {
      const key = normalizeValue(recipient.email);
      if (!key) return;
      const existing = recipients.get(key);
      recipients.set(key, {
        name: existing && existing.name ? existing.name : recipient.name,
        email: recipient.email,
        kind: recipient.kind,
      });
    });

    return Array.from(recipients.values());
  }

  function getSelectedRecipients() {
    return getAllRecipientCandidates().filter((recipient) => state.selectedRecipientEmails.has(normalizeValue(recipient.email)));
  }

  function renderConnectionSummary(person) {
    const nameEl = document.getElementById("connected-user-name");
    const emailEl = document.getElementById("connected-user-email");
    const summaryEl = document.getElementById("connection-summary");
    const statusEl = document.getElementById("connection-status");

    if (nameEl) nameEl.textContent = person && person.alias ? person.alias : "No user connected";
    if (emailEl) {
      emailEl.textContent = person && person.email
        ? `${person.email}${person.source ? ` · ${person.source}` : ""}`
        : "Sign in from this page or Workspace connection.";
    }
    if (summaryEl) {
      summaryEl.textContent = person && person.connected
        ? "Microsoft Graph session is ready for Outlook mailbox reads."
        : "Sign in with Microsoft to load recent Outlook emails.";
    }
    if (statusEl) {
      statusEl.textContent = person && person.connected ? "Connected" : "Not connected";
    }
  }

  function renderDealInvestorMetrics() {
    const container = document.getElementById("deal-investor-metrics");
    const deal = getSelectedDeal();
    if (!container) return;

    if (!deal) {
      container.innerHTML = "";
      return;
    }

    const investors = normalizeDealInvestorContacts(deal);
    const counts = CONTACT_STATUS_OPTIONS.reduce((accumulator, option) => {
      accumulator[option.value] = 0;
      return accumulator;
    }, {});
    investors.forEach((entry) => {
      counts[normalizeContactStatus(entry.contactStatus)] += 1;
    });

    const chips = [
      { label: "Tracked", value: investors.length },
      { label: "Contacted", value: (counts.contacted || 0) + (counts["follow-up"] || 0) },
      { label: "Replied", value: counts.replied || 0 },
      { label: "Invested", value: counts.invested || 0 },
      { label: "Passed", value: counts.passed || 0 },
    ];

    container.innerHTML = chips.map((chip) => `
      <div class="chip">
        <strong>${escapeHtml(String(chip.value))}</strong> ${escapeHtml(chip.label)}
      </div>
    `).join("");
  }

  function renderDealInvestors() {
    const deal = getSelectedDeal();
    const listEl = document.getElementById("deal-investors-list");
    const emptyEl = document.getElementById("deal-investors-empty");
    if (!listEl || !emptyEl) return;

    if (!deal) {
      emptyEl.hidden = false;
      emptyEl.textContent = "Select a deal to review saved investors.";
      listEl.innerHTML = "";
      return;
    }

    const investors = normalizeDealInvestorContacts(deal);
    emptyEl.hidden = investors.length > 0;
    if (!investors.length) {
      emptyEl.textContent = "No investor contacts saved on this deal yet.";
      listEl.innerHTML = "";
      return;
    }

    listEl.innerHTML = investors.map((investor) => `
      <div class="deal-investor-card" data-investor-email="${escapeHtml(investor.email)}">
        <div class="deal-investor-card-top">
          <div>
            <div class="deal-investor-name">${escapeHtml(investor.name || investor.email)}</div>
            <a class="deal-investor-email" href="mailto:${escapeHtml(investor.email)}">${escapeHtml(investor.email)}</a>
          </div>
          <span class="${escapeHtml(getStatusBadgeClass(investor.contactStatus))}">
            ${escapeHtml(CONTACT_STATUS_LABELS[normalizeContactStatus(investor.contactStatus)])}
          </span>
        </div>
        <div class="deal-investor-meta">
          ${investor.sourceMessageSubject ? `Source: ${escapeHtml(investor.sourceMessageSubject)}` : "Manual entry"}
          ${investor.sourceFromEmail ? ` · From ${escapeHtml(investor.sourceFromName || investor.sourceFromEmail)}` : ""}
          ${investor.sourceReceivedAt ? ` · ${escapeHtml(formatDateTime(investor.sourceReceivedAt))}` : ""}
        </div>
        ${investor.notes ? `<div class="deal-investor-meta">${escapeHtml(investor.notes)}</div>` : ""}
        <div class="deal-investor-actions">
          <select data-investor-status="${escapeHtml(investor.email)}">
            ${CONTACT_STATUS_OPTIONS.map((option) => `
              <option value="${escapeHtml(option.value)}"${normalizeContactStatus(investor.contactStatus) === option.value ? " selected" : ""}>
                ${escapeHtml(option.label)}
              </option>
            `).join("")}
          </select>
          ${investor.sourceMessageWebLink ? `<a class="btn" href="${escapeHtml(investor.sourceMessageWebLink)}" target="_blank" rel="noopener noreferrer">Open source email</a>` : ""}
          <button class="btn" type="button" data-remove-investor="${escapeHtml(investor.email)}">Remove</button>
        </div>
      </div>
    `).join("");
  }

  function renderDealSelect() {
    const select = document.getElementById("deal-select");
    const selectedDealLabel = document.getElementById("selected-deal-label");
    const selectedDealMeta = document.getElementById("selected-deal-meta");
    const openDealBtn = document.getElementById("open-deal-btn");
    const saveBtn = document.getElementById("save-selection-btn");
    if (!select) return;

    const requestedId = getRequestedDealId();
    const hasCurrentSelection = state.deals.some((deal) => normalizeValue(deal.id) === normalizeValue(state.selectedDealId));
    if (!hasCurrentSelection) {
      const requestedDeal = state.deals.find((deal) => normalizeValue(deal.id) === normalizeValue(requestedId));
      state.selectedDealId = requestedDeal ? requestedDeal.id : (state.deals[0] && state.deals[0].id ? state.deals[0].id : "");
    }

    select.innerHTML = state.deals.length
      ? [
        { key: "with-retainer", label: "With retainer" },
        { key: "no-retainer", label: "0 / no retainer" },
      ].map((group) => {
        const deals = state.deals.filter((deal) => getDealRetainerState(deal).bucket === group.key);
        if (!deals.length) return "";
        return `
          <optgroup label="${escapeHtml(group.label)}">
            ${deals.map((deal) => {
              const id = String(deal.id || "").trim();
              const label = `${deal.company || deal.name || id}${deal.name && deal.company && normalizeValue(deal.name) !== normalizeValue(deal.company) ? ` · ${deal.name}` : ""}`;
              const selected = normalizeValue(id) === normalizeValue(state.selectedDealId) ? " selected" : "";
              return `<option value="${escapeHtml(id)}"${selected}>${escapeHtml(label)}</option>`;
            }).join("")}
          </optgroup>
        `;
      }).join("")
      : '<option value="">No deals available</option>';

    const selectedDeal = getSelectedDeal();
    if (selectedDealLabel) {
      selectedDealLabel.textContent = selectedDeal
        ? String(selectedDeal.company || selectedDeal.name || selectedDeal.id || "Selected deal")
        : "No deal selected";
    }
    if (selectedDealMeta) {
      const selectedRetainerState = getDealRetainerState(selectedDeal);
      selectedDealMeta.textContent = selectedDeal
        ? `${selectedDeal.name || selectedDeal.company || selectedDeal.id}${selectedDeal.stage ? ` · ${selectedDeal.stage}` : ""} · ${selectedRetainerState.label}`
        : (state.deals.length
          ? "Choose a deal below to store investor contacts."
          : "No deals are loaded yet. Open Deals overview or let the shared deals sync finish.");
    }
    if (saveBtn) saveBtn.disabled = !selectedDeal;
    if (openDealBtn) {
      if (selectedDeal && selectedDeal.id) {
        openDealBtn.hidden = false;
        openDealBtn.setAttribute("href", buildPageUrl("deal-details", { id: selectedDeal.id }));
      } else {
        openDealBtn.hidden = true;
        openDealBtn.removeAttribute("href");
      }
    }

    updateSelectedDealInUrl();
    renderDealInvestorMetrics();
    renderDealInvestors();
  }

  function renderMessages() {
    const listEl = document.getElementById("message-list");
    const emptyEl = document.getElementById("message-empty");
    if (!listEl || !emptyEl) return;

    const filteredMessages = getFilteredMessages();
    ensureSelectedMessageIsVisible(filteredMessages);
    renderMessagePicker(filteredMessages);

    emptyEl.hidden = filteredMessages.length > 0;
    if (!filteredMessages.length) {
      emptyEl.textContent = !state.hasSearchedMessages
        ? "Search Outlook to load matching emails."
        : (state.messages.length
          ? "No loaded emails match that search."
          : "No emails were returned for that Outlook search.");
    }

    listEl.innerHTML = filteredMessages.map((message) => {
      const isActive = String(message.id || "") === String(state.selectedMessageId || "");
      const totalRecipients =
        (Array.isArray(message.toRecipients) ? message.toRecipients.length : 0) +
        (Array.isArray(message.ccRecipients) ? message.ccRecipients.length : 0);
      return `
        <button class="message-card${isActive ? " is-active" : ""}" type="button" data-message-id="${escapeHtml(message.id || "")}">
          <div class="message-card-top">
            <div>
              <div class="message-card-subject">${escapeHtml(message.subject || "(No subject)")}</div>
              <div class="message-card-meta">
                From ${escapeHtml((message.from && (message.from.name || message.from.address)) || "Unknown sender")}
                · ${escapeHtml(formatDateTime(message.receivedDateTime))}
              </div>
            </div>
            <span class="recipient-pill">${escapeHtml(String(totalRecipients))} recipients</span>
          </div>
          <div class="message-card-preview">${escapeHtml(message.bodyPreview || "No preview available.")}</div>
        </button>
      `;
    }).join("");
  }

  function renderMessagePicker(filteredMessages) {
    const selectEl = document.getElementById("message-select");
    const summaryEl = document.getElementById("message-results-summary");
    const copyEl = document.getElementById("message-selection-copy");
    const visibleMessages = Array.isArray(filteredMessages) ? filteredMessages : [];
    const currentMessage = getSelectedMessage();

    if (summaryEl) {
      if (!state.hasSearchedMessages) {
        summaryEl.textContent = "Search Outlook to begin";
      } else if (!state.messages.length) {
        summaryEl.textContent = "No emails found";
      } else if (state.messageFilter) {
        summaryEl.textContent = `${visibleMessages.length} match${visibleMessages.length === 1 ? "" : "es"} for "${state.messageFilter}"`;
      } else {
        summaryEl.textContent = `${visibleMessages.length} emails loaded`;
      }
    }

    if (selectEl) {
      selectEl.disabled = visibleMessages.length === 0;
      selectEl.innerHTML = visibleMessages.length
        ? visibleMessages.map((message) => `
          <option value="${escapeHtml(message.id || "")}"${String(message.id || "") === String(state.selectedMessageId || "") ? " selected" : ""}>
            ${escapeHtml(buildMessageOptionLabel(message))}
          </option>
        `).join("")
        : '<option value="">No email matches the current filter</option>';
    }

    if (copyEl) {
      if (!state.hasSearchedMessages) {
        copyEl.textContent = "Type your Outlook search first, click Search Outlook, then choose the email you want to use.";
      } else if (!visibleMessages.length) {
        copyEl.textContent = "Try a different Outlook search to load matching emails.";
      } else if (currentMessage) {
        copyEl.textContent = `Using "${currentMessage.subject || "(No subject)"}" as the source email for recipient import.`;
      } else {
        copyEl.textContent = "Choose the source email you want to use for recipient import.";
      }
    }
  }

  function renderManualRecipients() {
    const listEl = document.getElementById("manual-recipient-list");
    const countEl = document.getElementById("recipient-count");
    if (!listEl || !countEl) return;

    listEl.innerHTML = state.manualRecipients.map((recipient) => {
      const key = normalizeValue(recipient.email);
      const checked = state.selectedRecipientEmails.has(key) ? " checked" : "";
      return `
        <div class="manual-recipient-card">
          <label>
            <input type="checkbox" data-recipient-email="${escapeHtml(recipient.email)}"${checked} />
            <span>
              <div class="deal-investor-name">${escapeHtml(recipient.name || recipient.email)}</div>
              <div class="manual-recipient-meta">${escapeHtml(recipient.email)} · Added manually</div>
            </span>
          </label>
          <div class="deal-investor-actions">
            <button class="btn" type="button" data-remove-manual-email="${escapeHtml(recipient.email)}">Remove</button>
          </div>
        </div>
      `;
    }).join("");

    countEl.textContent = `${getSelectedRecipients().length} selected`;
  }

  function renderSelectedMessage() {
    const message = getSelectedMessage();
    const titleEl = document.getElementById("selected-message-title");
    const metaEl = document.getElementById("selected-message-meta");
    const recipientList = document.getElementById("recipient-list");
    const emptyEl = document.getElementById("recipient-empty");
    const countEl = document.getElementById("recipient-count");

    if (!titleEl || !metaEl || !recipientList || !emptyEl || !countEl) return;

    const recipients = getMessageRecipients(message);
    titleEl.textContent = message ? (message.subject || "(No subject)") : "No email selected";
    metaEl.innerHTML = message
      ? `
        ${escapeHtml((message.from && (message.from.name || message.from.address)) || "Unknown sender")}
        · ${escapeHtml(formatDateTime(message.receivedDateTime))}
        ${message.webLink ? ` · <a class="deal-investor-source-link" href="${escapeHtml(message.webLink)}" target="_blank" rel="noopener noreferrer">Open in Outlook</a>` : ""}
      `
      : "Pick an email on the left to review its recipients.";

    emptyEl.hidden = recipients.length > 0;
    if (!recipients.length) {
      emptyEl.textContent = message
        ? "This email has no To or Cc recipients to import."
        : "Select an email to see recipients.";
    }

    recipientList.innerHTML = recipients.map((recipient) => {
      const key = normalizeValue(recipient.email);
      const checked = state.selectedRecipientEmails.has(key) ? " checked" : "";
      return `
        <div class="recipient-card">
          <label>
            <input type="checkbox" data-recipient-email="${escapeHtml(recipient.email)}"${checked} />
            <span>
              <div class="deal-investor-name">${escapeHtml(recipient.name || recipient.email)}</div>
              <div class="recipient-card-meta">
                ${escapeHtml(recipient.email)}
                · <span class="recipient-pill">${escapeHtml(recipient.kind.toUpperCase())}</span>
              </div>
            </span>
          </label>
        </div>
      `;
    }).join("");

    countEl.textContent = `${getSelectedRecipients().length} selected`;
    renderManualRecipients();
  }

  async function refreshConnectionState() {
    if (!AppCore || typeof AppCore.getCurrentConnectedPerson !== "function") {
      renderConnectionSummary(null);
      return null;
    }
    const person = await AppCore.getCurrentConnectedPerson();
    renderConnectionSummary(person);
    return person;
  }

  function setSelectedMessage(messageId) {
    state.selectedMessageId = String(messageId || "").trim();
    refreshSelectedRecipientsForCurrentMessage();
    renderMessages();
    renderSelectedMessage();
  }

  async function loadMessages(options = {}) {
    const searchInput = document.getElementById("message-search");
    const inputValue = String(searchInput && searchInput.value || "").trim();
    const searchTerm = options && options.reuseLastSearch
      ? (inputValue || String(state.messageFilter || "").trim())
      : inputValue;
    if (searchInput && !inputValue && searchTerm) {
      searchInput.value = searchTerm;
    }
    if (!searchTerm) {
      state.hasSearchedMessages = false;
      state.messageFilter = "";
      state.messages = [];
      state.selectedMessageId = "";
      refreshSelectedRecipientsForCurrentMessage();
      renderMessages();
      renderSelectedMessage();
      setStatus("messages-status", "Enter a search to load emails");
      return;
    }

    setStatus("messages-status", "Loading…");
    try {
      state.hasSearchedMessages = true;
      state.messageFilter = searchTerm;
      const result = AppCore && typeof AppCore.listOutlookMessages === "function"
        ? await AppCore.listOutlookMessages({ top: 60, search: searchTerm })
        : { items: [] };
      state.messages = Array.isArray(result && result.items) ? result.items : [];
      if (!state.selectedMessageId && state.messages.length) {
        setSelectedMessage(state.messages[0].id);
      } else {
        renderMessages();
        renderSelectedMessage();
      }
      setStatus("messages-status", `${state.messages.length} match${state.messages.length === 1 ? "" : "es"} loaded`);
    } catch (error) {
      state.hasSearchedMessages = true;
      state.messages = [];
      state.selectedMessageId = "";
      renderMessages();
      renderSelectedMessage();
      setStatus("messages-status", error instanceof Error ? error.message : "Failed to load emails");
    }
  }

  async function persistSelectedRecipients() {
    const deal = getSelectedDeal();
    const statusEl = document.getElementById("assignment-status");
    const notesInput = document.getElementById("assignment-notes");
    const contactStatusSelect = document.getElementById("contact-status-select");
    if (!statusEl || !notesInput || !contactStatusSelect) return;

    if (!deal) {
      statusEl.textContent = "Choose a deal first.";
      return;
    }

    const recipients = getSelectedRecipients();
    if (!recipients.length) {
      statusEl.textContent = "Select at least one recipient.";
      return;
    }

    const selectedMessage = getSelectedMessage();
    const nextStatus = normalizeContactStatus(contactStatusSelect.value);
    const note = String(notesInput.value || "").trim();
    const now = new Date().toISOString();
    const byEmail = new Map(
      normalizeDealInvestorContacts(deal).map((entry) => [normalizeValue(entry.email), entry]),
    );

    recipients.forEach((recipient, index) => {
      const email = String(recipient.email || "").trim();
      const key = normalizeValue(email);
      if (!key) return;
      const existing = byEmail.get(key);
      byEmail.set(key, {
        id: existing && existing.id ? existing.id : `investor-${Date.now()}-${index}`,
        name: String(recipient.name || (existing && existing.name) || "").trim(),
        email,
        contactStatus: nextStatus,
        source: selectedMessage ? "outlook-message" : "manual-entry",
        sourceMessageId: selectedMessage ? String(selectedMessage.id || "") : String((existing && existing.sourceMessageId) || ""),
        sourceMessageSubject: selectedMessage ? String(selectedMessage.subject || "") : String((existing && existing.sourceMessageSubject) || ""),
        sourceMessageWebLink: selectedMessage ? String(selectedMessage.webLink || "") : String((existing && existing.sourceMessageWebLink) || ""),
        sourceReceivedAt: selectedMessage ? String(selectedMessage.receivedDateTime || "") : String((existing && existing.sourceReceivedAt) || ""),
        sourceRecipientType: String(recipient.kind || (existing && existing.sourceRecipientType) || "manual"),
        sourceFromName: selectedMessage && selectedMessage.from ? String(selectedMessage.from.name || "") : String((existing && existing.sourceFromName) || ""),
        sourceFromEmail: selectedMessage && selectedMessage.from ? String(selectedMessage.from.address || "") : String((existing && existing.sourceFromEmail) || ""),
        notes: note || String((existing && existing.notes) || ""),
        addedAt: String((existing && existing.addedAt) || now),
        updatedAt: now,
      });
    });

    deal.investorContacts = Array.from(byEmail.values());

    statusEl.textContent = "Saving…";
    try {
      await saveDealsData();
      renderDealInvestorMetrics();
      renderDealInvestors();
      statusEl.textContent = `Saved ${recipients.length} investor${recipients.length === 1 ? "" : "s"}`;
      window.setTimeout(() => {
        statusEl.textContent = "";
      }, 1800);
    } catch (error) {
      statusEl.textContent = error instanceof Error ? error.message : "Save failed";
    }
  }

  async function updateInvestorStatus(email, nextStatus) {
    const deal = getSelectedDeal();
    if (!deal) return;
    const investors = normalizeDealInvestorContacts(deal);
    const target = investors.find((entry) => normalizeValue(entry.email) === normalizeValue(email));
    if (!target) return;
    target.contactStatus = normalizeContactStatus(nextStatus);
    target.updatedAt = new Date().toISOString();
    deal.investorContacts = investors;
    await saveDealsData();
    renderDealInvestorMetrics();
    renderDealInvestors();
  }

  async function removeInvestor(email) {
    const deal = getSelectedDeal();
    if (!deal) return;
    const investors = normalizeDealInvestorContacts(deal);
    deal.investorContacts = investors.filter((entry) => normalizeValue(entry.email) !== normalizeValue(email));
    await saveDealsData();
    renderDealInvestorMetrics();
    renderDealInvestors();
  }

  function addManualRecipient(name, email) {
    const cleanEmail = String(email || "").trim();
    const key = normalizeValue(cleanEmail);
    if (!cleanEmail || !key.includes("@")) {
      throw new Error("Enter a valid email address.");
    }

    const cleanName = String(name || "").trim();
    const existingIndex = state.manualRecipients.findIndex((entry) => normalizeValue(entry.email) === key);
    if (existingIndex >= 0) {
      state.manualRecipients[existingIndex] = {
        name: cleanName || state.manualRecipients[existingIndex].name,
        email: cleanEmail,
      };
    } else {
      state.manualRecipients.push({ name: cleanName, email: cleanEmail });
    }
    state.selectedRecipientEmails.add(key);
    renderManualRecipients();
  }

  function startDeviceCodePolling(useDesktop) {
    clearDeviceCodeTimer();

    const poll = async () => {
      if (!state.deviceCodeState || !state.deviceCodeState.deviceCode) return;

      try {
        const result = useDesktop && window.PlutusDesktop && typeof window.PlutusDesktop.pollGraphDeviceCode === "function"
          ? await window.PlutusDesktop.pollGraphDeviceCode({ deviceCode: state.deviceCodeState.deviceCode })
          : await AppCore.pollGraphDeviceCode(state.deviceCodeState.deviceCode);

        const payload = result && result.ok && result.data ? result.data : null;
        if (!payload) {
          throw new Error((result && result.error) || "Sign-in polling failed.");
        }

        if (payload.ok === false) {
          const statusEl = document.getElementById("device-code-status");
          if (statusEl) {
            statusEl.textContent = payload.error === "authorization_pending"
              ? "Waiting for authorization…"
              : (payload.error_description || payload.error || "Authorization pending.");
          }
          state.deviceCodeTimer = window.setTimeout(poll, Math.max(Number(payload.interval || state.deviceCodeState.interval || 5), 2) * 1000);
          return;
        }

        clearDeviceCodeTimer();
        setDeviceCodeBox(false);
        window.dispatchEvent(new CustomEvent("appcore:graph-session-updated"));
        await refreshConnectionState();
        if (state.hasSearchedMessages && state.messageFilter) {
          await loadMessages({ reuseLastSearch: true });
        }
      } catch (error) {
        clearDeviceCodeTimer();
        const statusEl = document.getElementById("device-code-status");
        if (statusEl) {
          statusEl.textContent = error instanceof Error ? error.message : "Sign-in polling failed.";
        }
      }
    };

    state.deviceCodeTimer = window.setTimeout(poll, Math.max(Number(state.deviceCodeState.interval || 5), 2) * 1000);
  }

  async function startMicrosoftSignIn() {
    const useDesktop = Boolean(window.PlutusDesktop && typeof window.PlutusDesktop.requestGraphDeviceCode === "function");
    try {
      const result = useDesktop
        ? await window.PlutusDesktop.requestGraphDeviceCode()
        : await AppCore.requestGraphDeviceCode();
      const payload = result && result.ok && result.data ? result.data : null;
      if (!payload) {
        throw new Error((result && result.error) || "Unable to start Microsoft sign-in.");
      }

      state.deviceCodeState = {
        deviceCode: payload.device_code,
        interval: payload.interval,
      };
      setDeviceCodeBox(true, {
        userCode: payload.user_code,
        verificationUri: payload.verification_uri,
        message: payload.message,
      });
      startDeviceCodePolling(useDesktop);
    } catch (error) {
      setStatus("connection-status", error instanceof Error ? error.message : "Microsoft sign-in failed");
    }
  }

  function bindEvents() {
    const messageList = document.getElementById("message-list");
    const recipientList = document.getElementById("recipient-list");
    const manualRecipientList = document.getElementById("manual-recipient-list");
    const manualRecipientForm = document.getElementById("manual-recipient-form");
    const dealSelect = document.getElementById("deal-select");
    const dealInvestorsList = document.getElementById("deal-investors-list");
    const signInBtn = document.getElementById("sign-in-btn");
    const reloadMessagesBtn = document.getElementById("reload-messages-btn");
    const searchMessagesBtn = document.getElementById("search-messages-btn");
    const saveSelectionBtn = document.getElementById("save-selection-btn");
    const messageSearch = document.getElementById("message-search");
    const messageSelect = document.getElementById("message-select");

    if (signInBtn) {
      signInBtn.addEventListener("click", () => {
        startMicrosoftSignIn();
      });
    }

    if (reloadMessagesBtn) {
      reloadMessagesBtn.addEventListener("click", () => {
        loadMessages({ reuseLastSearch: true });
      });
    }

    if (searchMessagesBtn) {
      searchMessagesBtn.addEventListener("click", () => {
        loadMessages();
      });
    }

    if (messageSearch) {
      messageSearch.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") return;
        event.preventDefault();
        loadMessages();
      });
    }

    if (messageSelect) {
      messageSelect.addEventListener("change", (event) => {
        setSelectedMessage(event.target.value);
      });
    }

    if (messageList) {
      messageList.addEventListener("click", (event) => {
        const button = event.target.closest("[data-message-id]");
        if (!button) return;
        setSelectedMessage(button.getAttribute("data-message-id"));
      });
    }

    function handleRecipientToggle(event) {
      const checkbox = event.target.closest('input[data-recipient-email]');
      if (!checkbox) return;
      const key = normalizeValue(checkbox.getAttribute("data-recipient-email"));
      if (!key) return;
      if (checkbox.checked) {
        state.selectedRecipientEmails.add(key);
      } else {
        state.selectedRecipientEmails.delete(key);
      }
      renderSelectedMessage();
    }

    if (recipientList) {
      recipientList.addEventListener("change", handleRecipientToggle);
    }

    if (manualRecipientList) {
      manualRecipientList.addEventListener("change", handleRecipientToggle);
      manualRecipientList.addEventListener("click", (event) => {
        const button = event.target.closest("[data-remove-manual-email]");
        if (!button) return;
        const email = String(button.getAttribute("data-remove-manual-email") || "").trim();
        state.manualRecipients = state.manualRecipients.filter((entry) => normalizeValue(entry.email) !== normalizeValue(email));
        state.selectedRecipientEmails.delete(normalizeValue(email));
        renderSelectedMessage();
      });
    }

    if (manualRecipientForm) {
      manualRecipientForm.addEventListener("submit", (event) => {
        event.preventDefault();
        const nameInput = document.getElementById("manual-recipient-name");
        const emailInput = document.getElementById("manual-recipient-email");
        try {
          addManualRecipient(
            nameInput ? nameInput.value : "",
            emailInput ? emailInput.value : "",
          );
          if (nameInput) nameInput.value = "";
          if (emailInput) emailInput.value = "";
          setStatus("assignment-status", "");
        } catch (error) {
          setStatus("assignment-status", error instanceof Error ? error.message : "Failed to add email");
        }
      });
    }

    if (dealSelect) {
      dealSelect.addEventListener("change", (event) => {
        state.selectedDealId = String(event.target.value || "").trim();
        renderDealSelect();
      });
    }

    if (saveSelectionBtn) {
      saveSelectionBtn.addEventListener("click", () => {
        persistSelectedRecipients();
      });
    }

    if (dealInvestorsList) {
      dealInvestorsList.addEventListener("change", async (event) => {
        const select = event.target.closest("[data-investor-status]");
        if (!select) return;
        await updateInvestorStatus(select.getAttribute("data-investor-status"), select.value);
      });

      dealInvestorsList.addEventListener("click", async (event) => {
        const button = event.target.closest("[data-remove-investor]");
        if (!button) return;
        await removeInvestor(button.getAttribute("data-remove-investor"));
      });
    }

    window.addEventListener("appcore:graph-session-updated", async () => {
      const person = await refreshConnectionState();
      if (person && person.connected && state.hasSearchedMessages && state.messageFilter) {
        await loadMessages({ reuseLastSearch: true });
      }
    });

    window.addEventListener("appcore:deals-updated", (event) => {
      if (!event || !event.detail || !Array.isArray(event.detail.deals)) return;
      state.deals = event.detail.deals;
      renderDealSelect();
    });
  }

  async function initializePage() {
    bindEvents();

    if (AppCore && typeof AppCore.refreshDealsFromShareDrive === "function") {
      try {
        await AppCore.refreshDealsFromShareDrive("outlook-investor-sync");
      } catch {
        // Keep local deals if the remote sync is unavailable.
      }
    }

    loadDealsData();
    state.selectedDealId = getRequestedDealId();
    renderDealSelect();
    renderMessages();
    renderSelectedMessage();
    await refreshConnectionState();
  }

  document.addEventListener("DOMContentLoaded", initializePage);
})();
