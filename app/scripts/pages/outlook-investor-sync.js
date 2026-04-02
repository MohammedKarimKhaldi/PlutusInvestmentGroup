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
    recipientSourceMode: "message",
    messageFilter: "",
    hasSearchedMessages: false,
    deviceCodeState: null,
    deviceCodeTimer: null,
    connectedPerson: null,
    autoSelectedDealId: "",
  };

  function normalizeValue(value) {
    if (AppCore && typeof AppCore.normalizeValue === "function") {
      return AppCore.normalizeValue(value);
    }
    return String(value || "").trim().toLowerCase();
  }

  function isExcludedRecipientEmail(email) {
    return normalizeValue(email).endsWith("@plutus-investment.com");
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function toTitleCase(value) {
    return String(value || "")
      .split(/([\s'-]+)/)
      .map((part) => {
        if (!part || /^[\s'-]+$/.test(part)) return part;
        return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
      })
      .join("");
  }

  function deriveContactNameParts(name, email) {
    const cleanName = String(name || "").replace(/\s+/g, " ").trim();
    if (cleanName) {
      if (cleanName.includes(",")) {
        const [lastChunk, ...restChunks] = cleanName.split(",");
        const firstName = toTitleCase(restChunks.join(" ").trim());
        const lastName = toTitleCase(lastChunk.trim());
        const fullName = [firstName, lastName].filter(Boolean).join(" ").trim();
        return {
          name: fullName || cleanName,
          firstName,
          lastName,
        };
      }

      const parts = cleanName.split(/\s+/).filter(Boolean);
      const prefixes = new Set(["mr", "mrs", "ms", "dr", "miss", "sir"]);
      while (parts.length > 2 && prefixes.has(normalizeValue(parts[0]).replace(/\./g, ""))) {
        parts.shift();
      }
      const firstName = toTitleCase(parts[0] || "");
      const lastName = toTitleCase(parts.slice(1).join(" "));
      return {
        name: cleanName,
        firstName,
        lastName,
      };
    }

    const localPart = String(email || "").split("@")[0] || "";
    const derivedParts = localPart
      .replace(/[._-]+/g, " ")
      .split(/\s+/)
      .map((part) => part.trim())
      .filter(Boolean);
    const firstName = toTitleCase(derivedParts[0] || "");
    const lastName = toTitleCase(derivedParts.slice(1).join(" "));
    return {
      name: [firstName, lastName].filter(Boolean).join(" ").trim() || String(email || "").trim(),
      firstName,
      lastName,
    };
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

  function getTimeValue(value) {
    const timestamp = Date.parse(String(value || ""));
    return Number.isFinite(timestamp) ? timestamp : 0;
  }

  function describeRelativeTime(value) {
    const timestamp = getTimeValue(value);
    if (!timestamp) return "No recent email";
    const diffMs = Date.now() - timestamp;
    const diffDays = Math.floor(diffMs / (24 * 60 * 60 * 1000));
    if (diffDays <= 0) return "Today";
    if (diffDays === 1) return "1 day ago";
    if (diffDays < 7) return `${diffDays} days ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) === 1 ? "" : "s"} ago`;
    return formatDateTime(value);
  }

  function sanitizeFileNamePart(value) {
    return String(value || "")
      .trim()
      .replace(/[\\/:*?"<>|]+/g, " ")
      .replace(/\s+/g, "-")
      .replace(/-+/g, "-")
      .replace(/^-+|-+$/g, "")
      .toLowerCase();
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

  function findDealsMatchingSearch(term) {
    const normalizedTerm = normalizeValue(term);
    if (!normalizedTerm) return [];
    const parts = normalizedTerm.split(/\s+/).filter(Boolean);
    return state.deals.filter((deal) => {
      const haystack = normalizeValue([
        deal && deal.company,
        deal && deal.name,
        deal && deal.id,
      ].join(" "));
      return parts.every((part) => haystack.includes(part));
    });
  }

  function syncMatchedDealsForSearch(term) {
    const matches = findDealsMatchingSearch(term);
    state.autoSelectedDealId = "";

    if (matches.length === 1 && normalizeValue(matches[0].id) !== normalizeValue(state.selectedDealId)) {
      state.selectedDealId = matches[0].id;
      state.autoSelectedDealId = matches[0].id;
    }

    return matches;
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
        const nameParts = deriveContactNameParts(entry.name, email);
        return {
          id: String(entry.id || `investor-${index}`).trim(),
          name: String(entry.name || nameParts.name || "").trim(),
          firstName: String(entry.firstName || nameParts.firstName || "").trim(),
          lastName: String(entry.lastName || nameParts.lastName || "").trim(),
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

  function sortMessagesNewestFirst(messages) {
    return (Array.isArray(messages) ? messages.slice() : []).sort((left, right) => {
      const leftTime = getTimeValue(left && left.receivedDateTime);
      const rightTime = getTimeValue(right && right.receivedDateTime);
      if (leftTime !== rightTime) return rightTime - leftTime;
      return String(left && left.subject || "").localeCompare(String(right && right.subject || ""));
    });
  }

  function getFilteredMessages() {
    return Array.isArray(state.messages) ? state.messages.slice() : [];
  }

  function getTitleMatchedMessages() {
    const titleTerm = normalizeValue(state.messageFilter);
    const messages = getFilteredMessages();
    if (!titleTerm) return messages;
    const subjectMatches = messages.filter((message) => normalizeValue(message && message.subject).includes(titleTerm));
    return subjectMatches.length ? subjectMatches : messages;
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
    getActiveOutlookRecipients().forEach((recipient) => {
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
      if (!key || isExcludedRecipientEmail(email)) return;
      const existing = deduped.get(key);
      const nameParts = deriveContactNameParts(
        String(recipient && recipient.name || (existing && existing.name) || "").trim(),
        email,
      );
      deduped.set(key, {
        name: nameParts.name,
        firstName: String((existing && existing.firstName) || nameParts.firstName || "").trim(),
        lastName: String((existing && existing.lastName) || nameParts.lastName || "").trim(),
        email,
        kind: existing && existing.kind === "to" ? "to" : (recipient.kind || "to"),
      });
    });
    return Array.from(deduped.values());
  }

  function getBulkRecipientsFromMatchedMessages() {
    const deduped = new Map();

    getTitleMatchedMessages().forEach((message) => {
      getMessageRecipients(message).forEach((recipient) => {
        const key = normalizeValue(recipient.email);
        if (!key) return;
        const existing = deduped.get(key);
        const latestReceivedAt = existing && getTimeValue(existing.latestReceivedAt) > getTimeValue(message.receivedDateTime)
          ? existing.latestReceivedAt
          : String(message.receivedDateTime || "").trim();
        deduped.set(key, {
          name: String((existing && existing.name) || recipient.name || recipient.email).trim(),
          firstName: String((existing && existing.firstName) || recipient.firstName || "").trim(),
          lastName: String((existing && existing.lastName) || recipient.lastName || "").trim(),
          email: String(recipient.email || "").trim(),
          kind: existing && existing.kind === "to" ? "to" : String(recipient.kind || "to"),
          matchCount: Number(existing && existing.matchCount || 0) + 1,
          latestReceivedAt,
          latestSubject: latestReceivedAt === String(message.receivedDateTime || "").trim()
            ? String(message.subject || "").trim()
            : String(existing && existing.latestSubject || "").trim(),
        });
      });
    });

    return Array.from(deduped.values()).sort((left, right) => {
      if ((left.matchCount || 0) !== (right.matchCount || 0)) {
        return (right.matchCount || 0) - (left.matchCount || 0);
      }
      return String(left.name || left.email).localeCompare(String(right.name || right.email));
    });
  }

  function getActiveOutlookRecipients() {
    return state.recipientSourceMode === "search"
      ? getBulkRecipientsFromMatchedMessages()
      : getMessageRecipients(getSelectedMessage());
  }

  function getAllRecipientCandidates() {
    const recipients = new Map();

    state.manualRecipients.forEach((recipient) => {
      const key = normalizeValue(recipient.email);
      if (!key || isExcludedRecipientEmail(recipient.email)) return;
      recipients.set(key, {
        name: String(recipient.name || "").trim(),
        firstName: String(recipient.firstName || "").trim(),
        lastName: String(recipient.lastName || "").trim(),
        email: String(recipient.email || "").trim(),
        kind: "manual",
      });
    });

    getActiveOutlookRecipients().forEach((recipient) => {
      const key = normalizeValue(recipient.email);
      if (!key) return;
      const existing = recipients.get(key);
      recipients.set(key, {
        name: existing && existing.name ? existing.name : recipient.name,
        firstName: String((existing && existing.firstName) || recipient.firstName || "").trim(),
        lastName: String((existing && existing.lastName) || recipient.lastName || "").trim(),
        email: recipient.email,
        kind: recipient.kind,
        matchCount: Number(recipient.matchCount || 0),
        latestReceivedAt: String(recipient.latestReceivedAt || "").trim(),
        latestSubject: String(recipient.latestSubject || "").trim(),
      });
    });

    return Array.from(recipients.values());
  }

  function getTrackedInvestorByEmail(email) {
    const deal = getSelectedDeal();
    if (!deal) return null;
    return normalizeDealInvestorContacts(deal).find((entry) => normalizeValue(entry.email) === normalizeValue(email)) || null;
  }

  function countTrackedRecipientsForMessage(message) {
    return getMessageRecipients(message).filter((recipient) => Boolean(getTrackedInvestorByEmail(recipient.email))).length;
  }

  function renderTrackedInvestorHint(email) {
    const trackedInvestor = getTrackedInvestorByEmail(email);
    if (!trackedInvestor) return "";
    const status = normalizeContactStatus(trackedInvestor.contactStatus);
    return `
      <div class="recipient-card-meta">
        Already tracked on this deal
        · <span class="${escapeHtml(getStatusBadgeClass(status))}">
          ${escapeHtml(CONTACT_STATUS_LABELS[status])}
        </span>
      </div>
    `;
  }

  function buildCompanyStatusSummary() {
    const messages = sortMessagesNewestFirst(state.messages);
    const selectedDeal = getSelectedDeal();
    const latestMessage = messages[0] || null;
    const unreadCount = messages.filter((message) => !message.isRead).length;
    const connectedEmail = normalizeValue(state.connectedPerson && state.connectedPerson.email);
    const contactMap = new Map();

    messages.forEach((message) => {
      const from = message && message.from ? message.from : null;
      const fromEmail = normalizeValue(from && from.address);
      if (fromEmail && fromEmail !== connectedEmail && !isExcludedRecipientEmail(from && from.address)) {
        contactMap.set(fromEmail, {
          name: String(from && (from.name || from.address) || "").trim(),
          email: String(from && from.address || "").trim(),
        });
      }

      getMessageRecipients(message).forEach((recipient) => {
        const email = normalizeValue(recipient.email);
        if (!email || email === connectedEmail || isExcludedRecipientEmail(recipient.email)) return;
        if (!contactMap.has(email)) {
          contactMap.set(email, {
            name: String(recipient.name || recipient.email || "").trim(),
            email: String(recipient.email || "").trim(),
          });
        }
      });
    });

    const trackedOnDeal = selectedDeal
      ? Array.from(contactMap.values()).filter((contact) => Boolean(getTrackedInvestorByEmail(contact.email))).length
      : 0;

    let tone = "idle";
    let label = "Idle";
    let note = "Search Outlook by company name to surface the latest mail status.";
    if (state.hasSearchedMessages && !messages.length) {
      tone = "muted";
      label = "No activity";
      note = "No Outlook emails matched this company search.";
    } else if (messages.length) {
      const latestAgeDays = Math.floor((Date.now() - getTimeValue(latestMessage && latestMessage.receivedDateTime)) / (24 * 60 * 60 * 1000));
      if (unreadCount > 0) {
        tone = "warning";
        label = "Needs review";
        note = `${unreadCount} unread email${unreadCount === 1 ? "" : "s"} matched this company search.`;
      } else if (latestAgeDays <= 7) {
        tone = "live";
        label = "Active";
        note = `Latest Outlook activity was ${describeRelativeTime(latestMessage.receivedDateTime)}.`;
      } else {
        tone = "muted";
        label = "Quiet";
        note = `Latest Outlook activity was ${describeRelativeTime(latestMessage.receivedDateTime)}.`;
      }
    }

    return {
      tone,
      label,
      note,
      totalMessages: messages.length,
      unreadCount,
      latestMessage,
      relevantContacts: contactMap.size,
      trackedOnDeal,
      matchedDeals: findDealsMatchingSearch(state.messageFilter),
    };
  }

  function getSelectedRecipients() {
    return getAllRecipientCandidates().filter((recipient) => state.selectedRecipientEmails.has(normalizeValue(recipient.email)));
  }

  function getExportRecipients() {
    return getActiveOutlookRecipients()
      .filter((recipient) => recipient && recipient.email && !isExcludedRecipientEmail(recipient.email))
      .map((recipient) => ({
        firstName: String(recipient.firstName || "").trim(),
        lastName: String(recipient.lastName || "").trim(),
        name: String(recipient.name || "").trim(),
        email: String(recipient.email || "").trim(),
        recipientType: String(recipient.kind || "").trim().toUpperCase(),
        matchCount: Number(recipient.matchCount || 0),
        latestSubject: String(recipient.latestSubject || "").trim(),
        latestReceivedAt: String(recipient.latestReceivedAt || "").trim(),
      }));
  }

  function renderRecipientExportAction() {
    const exportBtn = document.getElementById("export-recipients-btn");
    const statusEl = document.getElementById("recipient-export-status");
    if (!exportBtn) return;

    const exportRows = getExportRecipients();
    const label = state.recipientSourceMode === "search" ? "Export title list" : "Export Excel";
    exportBtn.textContent = label;
    exportBtn.disabled = exportRows.length === 0;
    if (statusEl && statusEl.textContent !== "Exporting…") {
      statusEl.textContent = exportRows.length ? `${exportRows.length} ready for Excel` : "";
    }
  }

  function exportRecipientsToExcel() {
    const exportBtn = document.getElementById("export-recipients-btn");
    const statusEl = document.getElementById("recipient-export-status");
    const rows = getExportRecipients();

    if (typeof window.XLSX === "undefined") {
      if (statusEl) statusEl.textContent = "Excel export is unavailable right now.";
      return;
    }

    if (!rows.length) {
      if (statusEl) statusEl.textContent = "No recipients available to export.";
      return;
    }

    const selectedDeal = getSelectedDeal();
    const sheetRows = rows.map((recipient) => ({
      "First Name": recipient.firstName,
      "Last Name": recipient.lastName,
      Name: recipient.name,
      Email: recipient.email,
      "Recipient Type": recipient.recipientType,
      "Seen In Matching Emails": recipient.matchCount || "",
      "Latest Email Subject": recipient.latestSubject,
      "Latest Email Date": recipient.latestReceivedAt ? formatDateTime(recipient.latestReceivedAt) : "",
      Search: state.messageFilter || "",
      Deal: selectedDeal ? String(selectedDeal.company || selectedDeal.name || selectedDeal.id || "").trim() : "",
    }));

    const workbook = window.XLSX.utils.book_new();
    const worksheet = window.XLSX.utils.json_to_sheet(sheetRows);
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "Recipients");

    const fileStemBase = state.recipientSourceMode === "search"
      ? (state.messageFilter || "outlook-title-recipients")
      : ((getSelectedMessage() && getSelectedMessage().subject) || state.messageFilter || "outlook-recipients");
    const fileStem = sanitizeFileNamePart(fileStemBase) || "outlook-recipients";
    const filename = `${fileStem}.xlsx`;

    if (exportBtn) exportBtn.disabled = true;
    if (statusEl) statusEl.textContent = "Exporting…";

    try {
      window.XLSX.writeFile(workbook, filename);
      if (statusEl) statusEl.textContent = `Exported ${rows.length} recipients to ${filename}`;
    } catch (error) {
      if (statusEl) {
        statusEl.textContent = error instanceof Error ? error.message : "Failed to export Excel file.";
      }
    } finally {
      window.setTimeout(() => {
        if (statusEl && statusEl.textContent.includes("Exported")) {
          statusEl.textContent = `${rows.length} ready for Excel`;
        }
        renderRecipientExportAction();
      }, 1600);
    }
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
        ? "Microsoft Graph session is ready for company-level Outlook lookups."
        : "Sign in with Microsoft to search Outlook by company and review mailbox status.";
    }
    if (statusEl) {
      statusEl.textContent = person && person.connected ? "Connected" : "Not connected";
    }
  }

  function renderCompanyStatus() {
    const pillEl = document.getElementById("company-status-pill");
    const copyEl = document.getElementById("company-status-copy");
    const matchEl = document.getElementById("company-deal-match");
    const gridEl = document.getElementById("company-status-grid");
    if (!pillEl || !copyEl || !matchEl || !gridEl) return;

    const summary = buildCompanyStatusSummary();
    pillEl.className = `company-status-pill is-${summary.tone}`;
    pillEl.textContent = summary.label;
    copyEl.textContent = summary.note;

    if (!state.hasSearchedMessages || !state.messageFilter) {
      matchEl.textContent = "No linked deal detected yet.";
    } else if (!summary.matchedDeals.length) {
      matchEl.textContent = "No deal matched this company search. You can still review emails and save contacts manually.";
    } else if (summary.matchedDeals.length === 1) {
      const match = summary.matchedDeals[0];
      const wasAutoSelected = normalizeValue(match.id) === normalizeValue(state.autoSelectedDealId);
      matchEl.textContent = wasAutoSelected
        ? `Matched deal ${match.company || match.name || match.id} and selected it automatically.`
        : `Matched deal ${match.company || match.name || match.id}.`;
    } else {
      const currentMatch = summary.matchedDeals.find((deal) => normalizeValue(deal.id) === normalizeValue(state.selectedDealId));
      matchEl.textContent = currentMatch
        ? `${summary.matchedDeals.length} deals match this company search. ${currentMatch.company || currentMatch.name || currentMatch.id} is currently selected.`
        : `${summary.matchedDeals.length} deals match this company search. Keep the right deal selected before saving contacts.`;
    }

    const cards = [
      {
        label: "Mail status",
        value: summary.label,
        note: summary.note,
      },
      {
        label: "Matching emails",
        value: String(summary.totalMessages),
        note: summary.unreadCount
          ? `${summary.unreadCount} unread need review`
          : "No unread emails in these matches",
      },
      {
        label: "Latest activity",
        value: summary.latestMessage ? describeRelativeTime(summary.latestMessage.receivedDateTime) : "No recent email",
        note: summary.latestMessage
          ? `${summary.latestMessage.subject || "(No subject)"} · ${formatDateTime(summary.latestMessage.receivedDateTime)}`
          : "Search Outlook to see the most recent email touchpoint",
      },
      {
        label: "Relevant contacts",
        value: String(summary.relevantContacts),
        note: getSelectedDeal()
          ? `${summary.trackedOnDeal} already tracked on the selected deal`
          : "Select a deal to compare against saved investor contacts",
      },
    ];

    gridEl.innerHTML = cards.map((card) => `
      <div class="company-status-card">
        <div class="company-status-label">${escapeHtml(card.label)}</div>
        <div class="company-status-value">${escapeHtml(card.value)}</div>
        <div class="company-status-note">${escapeHtml(card.note)}</div>
      </div>
    `).join("");
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
    const useDealSearchBtn = document.getElementById("search-selected-deal-btn");
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
    if (useDealSearchBtn) {
      const dealSearchLabel = selectedDeal
        ? `Use ${selectedDeal.company || selectedDeal.name || "selected deal"}`
        : "Use selected deal company";
      useDealSearchBtn.textContent = dealSearchLabel;
      useDealSearchBtn.disabled = !selectedDeal;
    }
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
    renderMessages();
    renderSelectedMessage();
    renderCompanyStatus();
    renderRecipientExportAction();
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
      const totalRecipients = getMessageRecipients(message).length;
      const trackedRecipients = countTrackedRecipientsForMessage(message);
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
            <div class="message-card-badges">
              <span class="recipient-pill">${escapeHtml(String(totalRecipients))} recipients</span>
              ${!message.isRead ? '<span class="message-state-pill is-unread">Unread</span>' : ""}
              ${trackedRecipients ? `<span class="message-state-pill is-linked">${escapeHtml(String(trackedRecipients))} tracked</span>` : ""}
            </div>
          </div>
          <div class="message-card-preview">${escapeHtml(message.bodyPreview || "No preview available.")}</div>
        </button>
      `;
    }).join("");

    renderRecipientExportAction();
  }

  function renderMessagePicker(filteredMessages) {
    const selectEl = document.getElementById("message-select");
    const summaryEl = document.getElementById("message-results-summary");
    const copyEl = document.getElementById("message-selection-copy");
    const visibleMessages = Array.isArray(filteredMessages) ? filteredMessages : [];
    const currentMessage = getSelectedMessage();

    if (summaryEl) {
      if (!state.hasSearchedMessages) {
        summaryEl.textContent = "Search a company to begin";
      } else if (!state.messages.length) {
        summaryEl.textContent = `No Outlook emails found for "${state.messageFilter}"`;
      } else if (state.messageFilter) {
        summaryEl.textContent = `${visibleMessages.length} matching email${visibleMessages.length === 1 ? "" : "s"} for "${state.messageFilter}"`;
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
        copyEl.textContent = "Type a company name, check Outlook status, and the latest matching email will be picked automatically.";
      } else if (!visibleMessages.length) {
        copyEl.textContent = "Try a different company name or keyword to pull matching Outlook emails.";
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
              ${(recipient.firstName || recipient.lastName) ? `<div class="manual-recipient-meta">First name: ${escapeHtml(recipient.firstName || "—")} · Last name: ${escapeHtml(recipient.lastName || "—")}</div>` : ""}
              ${renderTrackedInvestorHint(recipient.email)}
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

  function renderRecipientSourceSwitcher() {
    const selectedBtn = document.getElementById("use-selected-message-btn");
    const searchBtn = document.getElementById("use-search-recipient-list-btn");
    if (!selectedBtn || !searchBtn) return;

    const hasSelectedMessage = Boolean(getSelectedMessage());
    const hasBulkRecipients = getBulkRecipientsFromMatchedMessages().length > 0;
    const useSearchMode = state.recipientSourceMode === "search";

    selectedBtn.disabled = !hasSelectedMessage;
    searchBtn.disabled = !hasBulkRecipients;
    selectedBtn.className = `btn${!useSearchMode ? " btn-primary" : ""}`;
    searchBtn.className = `btn${useSearchMode ? " btn-primary" : ""}`;
  }

  function setRecipientSourceMode(mode) {
    state.recipientSourceMode = mode === "search" ? "search" : "message";
    refreshSelectedRecipientsForCurrentMessage();
    renderSelectedMessage();
  }

  function renderSelectedMessage() {
    const message = getSelectedMessage();
    const titleEl = document.getElementById("selected-message-title");
    const metaEl = document.getElementById("selected-message-meta");
    const recipientList = document.getElementById("recipient-list");
    const emptyEl = document.getElementById("recipient-empty");
    const countEl = document.getElementById("recipient-count");

    if (!titleEl || !metaEl || !recipientList || !emptyEl || !countEl) return;

    const useSearchMode = state.recipientSourceMode === "search";
    const recipients = useSearchMode ? getBulkRecipientsFromMatchedMessages() : getMessageRecipients(message);
    const titleMatchedMessages = getTitleMatchedMessages();
    titleEl.textContent = useSearchMode
      ? `All emails with title "${state.messageFilter || "current search"}"`
      : (message ? (message.subject || "(No subject)") : "No email selected");
    metaEl.innerHTML = useSearchMode
      ? (
        recipients.length
          ? `${escapeHtml(String(recipients.length))} unique recipient${recipients.length === 1 ? "" : "s"} found across ${escapeHtml(String(titleMatchedMessages.length))} email${titleMatchedMessages.length === 1 ? "" : "s"} whose subject matches this title.`
          : "Run an Outlook search first to collect recipients from all emails with this title."
      )
      : (
        message
          ? `
            ${escapeHtml((message.from && (message.from.name || message.from.address)) || "Unknown sender")}
            · ${escapeHtml(formatDateTime(message.receivedDateTime))}
            ${message.webLink ? ` · <a class="deal-investor-source-link" href="${escapeHtml(message.webLink)}" target="_blank" rel="noopener noreferrer">Open in Outlook</a>` : ""}
          `
          : "Pick an email on the left to review its recipients."
      );

    emptyEl.hidden = recipients.length > 0;
    if (!recipients.length) {
      emptyEl.textContent = useSearchMode
        ? "No recipients were found across the current matching emails."
        : (message
          ? "This email has no To or Cc recipients to import."
          : "Select an email to see recipients.");
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
              ${(recipient.firstName || recipient.lastName) ? `<div class="recipient-card-meta">First name: ${escapeHtml(recipient.firstName || "—")} · Last name: ${escapeHtml(recipient.lastName || "—")}</div>` : ""}
              ${useSearchMode ? `<div class="recipient-card-meta">Seen on ${escapeHtml(String(recipient.matchCount || 0))} matching email${recipient.matchCount === 1 ? "" : "s"}${recipient.latestSubject ? ` · Latest: ${escapeHtml(recipient.latestSubject)}` : ""}</div>` : ""}
              ${renderTrackedInvestorHint(recipient.email)}
            </span>
          </label>
        </div>
      `;
    }).join("");

    renderRecipientSourceSwitcher();
    countEl.textContent = `${getSelectedRecipients().length} selected`;
    renderManualRecipients();
    renderRecipientExportAction();
  }

  async function refreshConnectionState() {
    if (!AppCore || typeof AppCore.getCurrentConnectedPerson !== "function") {
      state.connectedPerson = null;
      renderConnectionSummary(null);
      renderCompanyStatus();
      return null;
    }
    const person = await AppCore.getCurrentConnectedPerson();
    state.connectedPerson = person || null;
    renderConnectionSummary(person);
    renderCompanyStatus();
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
      state.autoSelectedDealId = "";
      refreshSelectedRecipientsForCurrentMessage();
      renderMessages();
      renderSelectedMessage();
      renderCompanyStatus();
      setStatus("messages-status", "Enter a search to load emails");
      return;
    }

    setStatus("messages-status", "Loading…");
    try {
      state.hasSearchedMessages = true;
      state.messageFilter = searchTerm;
      syncMatchedDealsForSearch(searchTerm);
      renderDealSelect();
      const result = AppCore && typeof AppCore.listOutlookMessages === "function"
        ? await AppCore.listOutlookMessages({ top: 5000, search: searchTerm })
        : { items: [] };
      state.messages = sortMessagesNewestFirst(Array.isArray(result && result.items) ? result.items : []);
      if (!state.selectedMessageId && state.messages.length) {
        setSelectedMessage(state.messages[0].id);
      } else {
        renderMessages();
        renderSelectedMessage();
      }
      renderCompanyStatus();
      setStatus("messages-status", `${state.messages.length} match${state.messages.length === 1 ? "" : "es"} loaded`);
    } catch (error) {
      state.hasSearchedMessages = true;
      state.messages = [];
      state.selectedMessageId = "";
      renderMessages();
      renderSelectedMessage();
      renderCompanyStatus();
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
    const isSearchAggregateSource = state.recipientSourceMode === "search";
    const byEmail = new Map(
      normalizeDealInvestorContacts(deal).map((entry) => [normalizeValue(entry.email), entry]),
    );

    recipients.forEach((recipient, index) => {
      const email = String(recipient.email || "").trim();
      const key = normalizeValue(email);
      if (!key) return;
      const existing = byEmail.get(key);
      const nameParts = deriveContactNameParts(recipient.name || (existing && existing.name) || "", email);
      byEmail.set(key, {
        id: existing && existing.id ? existing.id : `investor-${Date.now()}-${index}`,
        name: String(recipient.name || (existing && existing.name) || nameParts.name || "").trim(),
        firstName: String(recipient.firstName || (existing && existing.firstName) || nameParts.firstName || "").trim(),
        lastName: String(recipient.lastName || (existing && existing.lastName) || nameParts.lastName || "").trim(),
        email,
        contactStatus: nextStatus,
        source: isSearchAggregateSource
          ? "outlook-search-aggregate"
          : (selectedMessage ? "outlook-message" : "manual-entry"),
        sourceMessageId: isSearchAggregateSource
          ? String((existing && existing.sourceMessageId) || "")
          : (selectedMessage ? String(selectedMessage.id || "") : String((existing && existing.sourceMessageId) || "")),
        sourceMessageSubject: isSearchAggregateSource
          ? String(recipient.latestSubject || state.messageFilter || (existing && existing.sourceMessageSubject) || "")
          : (selectedMessage ? String(selectedMessage.subject || "") : String((existing && existing.sourceMessageSubject) || "")),
        sourceMessageWebLink: isSearchAggregateSource
          ? String((existing && existing.sourceMessageWebLink) || "")
          : (selectedMessage ? String(selectedMessage.webLink || "") : String((existing && existing.sourceMessageWebLink) || "")),
        sourceReceivedAt: isSearchAggregateSource
          ? String(recipient.latestReceivedAt || (existing && existing.sourceReceivedAt) || "")
          : (selectedMessage ? String(selectedMessage.receivedDateTime || "") : String((existing && existing.sourceReceivedAt) || "")),
        sourceRecipientType: String(recipient.kind || (existing && existing.sourceRecipientType) || "manual"),
        sourceFromName: isSearchAggregateSource
          ? String((existing && existing.sourceFromName) || "")
          : (selectedMessage && selectedMessage.from ? String(selectedMessage.from.name || "") : String((existing && existing.sourceFromName) || "")),
        sourceFromEmail: isSearchAggregateSource
          ? String((existing && existing.sourceFromEmail) || "")
          : (selectedMessage && selectedMessage.from ? String(selectedMessage.from.address || "") : String((existing && existing.sourceFromEmail) || "")),
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
    if (isExcludedRecipientEmail(cleanEmail)) {
      throw new Error("Plutus internal email addresses are excluded from this list.");
    }

    const cleanName = String(name || "").trim();
    const nameParts = deriveContactNameParts(cleanName, cleanEmail);
    const existingIndex = state.manualRecipients.findIndex((entry) => normalizeValue(entry.email) === key);
    if (existingIndex >= 0) {
      state.manualRecipients[existingIndex] = {
        name: cleanName || state.manualRecipients[existingIndex].name || nameParts.name,
        firstName: nameParts.firstName || state.manualRecipients[existingIndex].firstName || "",
        lastName: nameParts.lastName || state.manualRecipients[existingIndex].lastName || "",
        email: cleanEmail,
      };
    } else {
      state.manualRecipients.push({
        name: cleanName || nameParts.name,
        firstName: nameParts.firstName,
        lastName: nameParts.lastName,
        email: cleanEmail,
      });
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
    const searchSelectedDealBtn = document.getElementById("search-selected-deal-btn");
    const saveSelectionBtn = document.getElementById("save-selection-btn");
    const messageSearch = document.getElementById("message-search");
    const messageSelect = document.getElementById("message-select");
    const useSelectedMessageBtn = document.getElementById("use-selected-message-btn");
    const useSearchRecipientListBtn = document.getElementById("use-search-recipient-list-btn");
    const exportRecipientsBtn = document.getElementById("export-recipients-btn");

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

    if (searchSelectedDealBtn) {
      searchSelectedDealBtn.addEventListener("click", () => {
        const deal = getSelectedDeal();
        if (!deal) return;
        const searchTerm = String(deal.company || deal.name || "").trim();
        if (!searchTerm) return;
        if (messageSearch) messageSearch.value = searchTerm;
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
        state.recipientSourceMode = "message";
        setSelectedMessage(event.target.value);
      });
    }

    if (messageList) {
      messageList.addEventListener("click", (event) => {
        const button = event.target.closest("[data-message-id]");
        if (!button) return;
        state.recipientSourceMode = "message";
        setSelectedMessage(button.getAttribute("data-message-id"));
      });
    }

    if (useSelectedMessageBtn) {
      useSelectedMessageBtn.addEventListener("click", () => {
        setRecipientSourceMode("message");
      });
    }

    if (useSearchRecipientListBtn) {
      useSearchRecipientListBtn.addEventListener("click", () => {
        setRecipientSourceMode("search");
      });
    }

    if (exportRecipientsBtn) {
      exportRecipientsBtn.addEventListener("click", () => {
        exportRecipientsToExcel();
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
        state.autoSelectedDealId = "";
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
      state.deals = sortDealsByRetainerState(event.detail.deals);
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
