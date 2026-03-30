(function initDealIntegrity(global) {
  const AppCore = global.AppCore;
  const DEFAULT_CURRENCY = "GBP";
  const KNOWN_INVOICE_STATUSES = new Set(["draft", "prepared", "sent", "part_paid", "paid", "cancelled"]);

  function normalizeValue(value) {
    if (AppCore && typeof AppCore.normalizeValue === "function") {
      return AppCore.normalizeValue(value);
    }
    return String(value || "").trim().toLowerCase();
  }

  function normalizeDateInput(value) {
    const raw = String(value || "").trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : "";
  }

  function normalizeCurrencyCode(value, fallback = DEFAULT_CURRENCY) {
    const raw = String(value || "").trim().toUpperCase();
    const cleaned = raw.replace(/[^A-Z]/g, "").slice(0, 3);
    return cleaned || fallback;
  }

  function parseAmount(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const raw = String(value == null ? "" : value).trim();
    if (!raw) return 0;
    const cleaned = raw.replace(/,/g, "").replace(/[^0-9.-]/g, "");
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function getDealCurrency(deal) {
    if (!deal || typeof deal !== "object") return DEFAULT_CURRENCY;
    return normalizeCurrencyCode(deal.currency, DEFAULT_CURRENCY);
  }

  function getRetainerMonthly(deal) {
    if (!deal || typeof deal !== "object") return "";
    if (deal.retainerMonthly != null && String(deal.retainerMonthly).trim()) return String(deal.retainerMonthly).trim();
    if (deal.Retainer != null && String(deal.Retainer).trim()) return String(deal.Retainer).trim();
    return "";
  }

  function hasPositiveRetainer(deal) {
    return parseAmount(getRetainerMonthly(deal)) > 0;
  }

  function normalizePaymentIntervalMonths(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) return 1;
    return Math.min(12, Math.max(1, Math.round(parsed)));
  }

  function getPaymentIntervalMonths(deal) {
    if (!deal || typeof deal !== "object") return 1;
    return normalizePaymentIntervalMonths(deal.retainerIntervalMonths || deal.paymentIntervalMonths || 1);
  }

  function getRetainerNextPaymentDate(deal) {
    if (!deal || typeof deal !== "object") return "";
    return normalizeDateInput(deal.retainerNextPaymentDate || deal.nextPaymentDate || "");
  }

  function getPaymentDay(deal) {
    if (!deal || typeof deal !== "object") return "";
    const raw = deal.retainerPaymentDay != null ? String(deal.retainerPaymentDay).trim() : "";
    if (!raw) return "";
    const parsed = Number(raw);
    if (!Number.isFinite(parsed)) return "";
    return String(Math.min(31, Math.max(1, Math.round(parsed))));
  }

  function computeNextExpectedDateObject(dayValue) {
    const day = Number(dayValue);
    if (!Number.isFinite(day) || day < 1 || day > 31) return null;

    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();
    const today = now.getDate();

    const monthDays = new Date(year, month + 1, 0).getDate();
    const targetDayThisMonth = Math.min(day, monthDays);
    if (targetDayThisMonth >= today) {
      return new Date(year, month, targetDayThisMonth);
    }

    const nextMonthDays = new Date(year, month + 2, 0).getDate();
    const targetDayNextMonth = Math.min(day, nextMonthDays);
    return new Date(year, month + 1, targetDayNextMonth);
  }

  function addMonths(date, monthCount) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
    const target = new Date(date.getFullYear(), date.getMonth() + monthCount, 1);
    const monthDays = new Date(target.getFullYear(), target.getMonth() + 1, 0).getDate();
    target.setDate(Math.min(date.getDate(), monthDays));
    return target;
  }

  function getNextScheduledPaymentDateObject(deal) {
    const intervalMonths = getPaymentIntervalMonths(deal);
    const explicitNextDate = getRetainerNextPaymentDate(deal);

    if (explicitNextDate) {
      const explicitDate = new Date(explicitNextDate);
      if (!Number.isNaN(explicitDate.getTime())) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        let nextDate = new Date(explicitDate.getFullYear(), explicitDate.getMonth(), explicitDate.getDate());
        while (nextDate < today) {
          const advanced = addMonths(nextDate, intervalMonths);
          if (!advanced) break;
          nextDate = advanced;
        }
        return nextDate;
      }
    }

    if (intervalMonths > 1) return null;
    return computeNextExpectedDateObject(getPaymentDay(deal));
  }

  function normalizeDealContacts(deal, options = {}) {
    const keepEmpty = Boolean(options.keepEmpty);
    const source = deal && Array.isArray(deal.contacts) ? deal.contacts : [];
    const normalized = source
      .map((entry) => {
        if (!entry || typeof entry !== "object") return null;
        const name = String(entry.name || "").trim();
        const title = String(entry.title || "").trim();
        const email = String(entry.email || "").trim();
        const isPrimary = Boolean(entry.isPrimary);
        if (!keepEmpty && !name && !title && !email) return null;
        return { name, title, email, isPrimary };
      })
      .filter(Boolean);

    if (!normalized.length) {
      const fallbackEmail = String((deal && (deal.mainContactEmail || deal.email)) || "").trim();
      const fallbackName = String((deal && deal.mainContactName) || "").trim();
      const fallbackTitle = String((deal && deal.mainContactTitle) || "").trim();
      if (fallbackEmail || fallbackName || fallbackTitle) {
        normalized.push({
          name: fallbackName,
          title: fallbackTitle,
          email: fallbackEmail,
          isPrimary: true,
        });
      }
    }

    if (normalized.length && !normalized.some((entry) => entry.isPrimary)) {
      normalized[0].isPrimary = true;
    }

    return normalized;
  }

  function getPrimaryDealContact(deal) {
    const contacts = normalizeDealContacts(deal);
    return contacts.find((entry) => entry.isPrimary) || contacts[0] || null;
  }

  function normalizeInvoiceStatus(value) {
    const raw = String(value || "").trim().toLowerCase();
    return KNOWN_INVOICE_STATUSES.has(raw) ? raw : "draft";
  }

  function isPdfReference(value) {
    return /\.pdf(?:$|[?#])/i.test(String(value || "").trim());
  }

  function hasInvoiceRecordPayload(entry) {
    if (!entry || typeof entry !== "object") return false;
    if (String(entry.url || "").trim()) return true;
    return Boolean(
      String(entry.invoiceNumber || "").trim() ||
      normalizeDateInput(entry.invoiceDate) ||
      normalizeDateInput(entry.dueDate) ||
      String(entry.description || "").trim() ||
      entry.readyToSend
    );
  }

  function normalizeDealInvoices(deal) {
    const source = deal && Array.isArray(deal.invoices) ? deal.invoices : [];
    return source
      .map((entry) => {
        if (typeof entry === "string") {
          const raw = entry.trim();
          if (!raw) return null;
          return {
            name: raw.split("/").pop() || "Invoice.pdf",
            url: raw,
            parentPath: "",
            addedAt: "",
            status: "draft",
            sentDate: "",
            dueDate: "",
            paidDate: "",
            invoiceNumber: "",
            invoiceDate: "",
            description: "",
            clientName: "",
            addressLine1: "",
            addressLine2: "",
            country: "",
            currency: getDealCurrency(deal),
            amount: "",
            vatRate: "",
            autoGenerated: false,
            readyToSend: false,
            billingMonthKey: "",
            generatedAt: "",
          };
        }

        if (!entry || typeof entry !== "object") return null;
        const url = String(entry.url || entry.webUrl || "").trim();
        const name = String(entry.name || entry.fileName || "").trim() || (url.split("/").pop() || "Invoice.pdf");
        const status = normalizeInvoiceStatus(entry.status);
        return {
          name,
          url,
          parentPath: String(entry.parentPath || "").trim(),
          addedAt: String(entry.addedAt || "").trim(),
          status,
          sentDate: normalizeDateInput(entry.sentDate || entry.invoiceDate || entry.sentAt),
          dueDate: normalizeDateInput(entry.dueDate),
          paidDate: normalizeDateInput(entry.paidDate || entry.payDate || entry.paidAt),
          invoiceNumber: String(entry.invoiceNumber || "").trim(),
          invoiceDate: normalizeDateInput(entry.invoiceDate),
          description: String(entry.description || "").trim(),
          clientName: String(entry.clientName || entry.client || "").trim(),
          addressLine1: String(entry.addressLine1 || "").trim(),
          addressLine2: String(entry.addressLine2 || "").trim(),
          country: String(entry.country || "").trim(),
          currency: normalizeCurrencyCode(entry.currency || getDealCurrency(deal), getDealCurrency(deal)),
          amount: String(entry.amount == null ? "" : entry.amount).trim(),
          vatRate: String(entry.vatRate == null ? "" : entry.vatRate).trim(),
          autoGenerated: Boolean(entry.autoGenerated),
          readyToSend: Boolean(entry.readyToSend || status === "prepared"),
          billingMonthKey: String(entry.billingMonthKey || "").trim(),
          generatedAt: String(entry.generatedAt || entry.addedAt || "").trim(),
        };
      })
      .filter((entry) => {
        if (!entry) return false;
        if (entry.url) return isPdfReference(entry.url) || isPdfReference(entry.name);
        return hasInvoiceRecordPayload(entry);
      });
  }

  function getInvoiceTimestamp(invoice) {
    const candidates = [
      invoice && invoice.paidDate,
      invoice && invoice.dueDate,
      invoice && invoice.sentDate,
      invoice && invoice.invoiceDate,
      invoice && invoice.generatedAt,
      invoice && invoice.addedAt,
    ];
    for (let index = 0; index < candidates.length; index += 1) {
      const parsed = Date.parse(candidates[index] || "");
      if (Number.isFinite(parsed)) return parsed;
    }
    return 0;
  }

  function getLatestInvoice(deal) {
    const invoices = normalizeDealInvoices(deal);
    return invoices
      .slice()
      .sort((left, right) => getInvoiceTimestamp(right) - getInvoiceTimestamp(left))[0] || null;
  }

  function isInvoiceOverdue(invoice) {
    if (!invoice || normalizeInvoiceStatus(invoice.status) === "paid" || normalizeInvoiceStatus(invoice.status) === "cancelled") {
      return false;
    }
    if (invoice.paidDate) return false;
    const dueDate = normalizeDateInput(invoice.dueDate);
    if (!dueDate) return false;
    const today = new Date().toISOString().slice(0, 10);
    return dueDate < today;
  }

  function getInvoiceDisplayStatus(invoice) {
    if (isInvoiceOverdue(invoice)) return "overdue";
    return normalizeInvoiceStatus(invoice && invoice.status);
  }

  function getInvoiceSummaryForDeal(deal) {
    const invoices = normalizeDealInvoices(deal);
    const summary = {
      total: invoices.length,
      outstanding: 0,
      paid: 0,
      overdue: 0,
      sent: 0,
      latest: null,
    };

    invoices.forEach((invoice) => {
      const status = normalizeInvoiceStatus(invoice.status);
      if (isInvoiceOverdue(invoice)) summary.overdue += 1;
      if (status === "paid") summary.paid += 1;
      if (status === "sent" || status === "part_paid") summary.sent += 1;
      if (["prepared", "sent", "part_paid"].includes(status)) summary.outstanding += 1;
    });

    summary.latest = getLatestInvoice(deal);
    return summary;
  }

  function normalizeDealLegalLinks(deal, options = {}) {
    const keepEmpty = Boolean(options.keepEmpty);
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
          if (!keepEmpty && !url) return null;
          return {
            title: url ? `Legal link ${index + 1}` : "",
            url,
          };
        }
        if (!entry || typeof entry !== "object") return null;
        const title = String(entry.title || entry.label || entry.name || "").trim();
        const url = String(entry.url || entry.href || entry.link || "").trim();
        if (!keepEmpty && !title && !url) return null;
        return {
          title: title || (url ? `Legal link ${index + 1}` : ""),
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

  function getLegalLinkStats(deal) {
    const links = normalizeDealLegalLinks(deal);
    const validCount = links.filter((entry) => Boolean(toSafeExternalUrl(entry.url))).length;
    return {
      total: links.length,
      valid: validCount,
      invalid: Math.max(links.length - validCount, 0),
    };
  }

  function getSubOwners(deal) {
    const source = deal && deal.subOwners;
    const values = Array.isArray(source)
      ? source
      : typeof source === "string"
        ? source.split(/[\n,;]+/)
        : [];
    return Array.from(new Set(values.map((entry) => String(entry || "").trim()).filter(Boolean)));
  }

  function getDealOwners(deal) {
    return Array.from(new Set([
      String(deal && (deal.seniorOwner || deal.owner) || "").trim(),
      String(deal && deal.juniorOwner || "").trim(),
      ...getSubOwners(deal),
    ].filter(Boolean)));
  }

  function hasDashboardLinked(deal) {
    return Boolean(String(deal && deal.fundraisingDashboardId || "").trim());
  }

  function hasDeckLinked(deal) {
    return Boolean(toSafeExternalUrl(deal && deal.deckUrl));
  }

  function buildDealIntegrityReport(deal, options = {}) {
    if (!deal || typeof deal !== "object") {
      return {
        issues: [],
        topIssues: [],
        counts: { total: 0, ownership: 0, accounting: 0, legal: 0, setup: 0 },
        statuses: { ownership: "unknown", accounting: "unknown", legal: "unknown", setup: "unknown" },
        fullyLinked: false,
        accounting: { hasRetainer: false, invoices: [], invoiceSummary: null, primaryContact: null, nextPaymentDate: null },
        legal: { links: [], stats: { total: 0, valid: 0, invalid: 0 } },
      };
    }

    const ownershipAvailable = Boolean(options.ownershipAvailable);
    const ownershipEntry = options.ownershipEntry || null;
    const issues = [];
    const addIssue = (scope, code, title, detail) => {
      issues.push({ scope, code, title, detail });
    };

    const owners = getDealOwners(deal);
    if (!owners.length) {
      addIssue("ownership", "missing_owners", "No deal owner assigned", "Add a senior, junior, or sub-owner so the deal has clear accountability.");
    }

    if (ownershipAvailable) {
      if (!ownershipEntry) {
        addIssue("ownership", "missing_ownership_link", "Missing from ownership dashboard", "This deal is not linked to any staffing row in the ownership dashboard.");
      } else if (ownershipEntry.linkStatus === "missing-staffing") {
        addIssue("ownership", "missing_staffing", "Needs staffing coverage", "The deal exists in overview but has no active staffing coverage linked yet.");
      } else if (ownershipEntry.linkStatus === "unlinked") {
        addIssue("ownership", "unlinked_staffing", "Unlinked staffing", "A staffing cluster exists, but it is not connected cleanly to this deal.");
      }
    }

    const primaryContact = getPrimaryDealContact(deal);
    const invoices = normalizeDealInvoices(deal);
    const invoiceSummary = getInvoiceSummaryForDeal(deal);
    const nextPaymentDate = getNextScheduledPaymentDateObject(deal);
    const hasRetainer = hasPositiveRetainer(deal);
    if (hasRetainer && !primaryContact) {
      addIssue("accounting", "missing_contact", "Missing billing contact", "Add a main accounting contact so invoices and payment follow-up have an owner.");
    }
    if (hasRetainer && !nextPaymentDate) {
      addIssue("accounting", "missing_schedule", "Missing payment schedule", "A retainer is set, but no next payment timing is configured.");
    }
    if (hasRetainer && !invoices.length) {
      addIssue("accounting", "missing_invoices", "No invoice history", "A retainer exists, but no invoice records have been added yet.");
    }
    if (invoiceSummary.overdue > 0) {
      addIssue(
        "accounting",
        "overdue_invoices",
        `${invoiceSummary.overdue} overdue invoice${invoiceSummary.overdue === 1 ? "" : "s"}`,
        "One or more invoices are past due and still marked unpaid."
      );
    }

    const legalLinks = normalizeDealLegalLinks(deal);
    const legalStats = getLegalLinkStats(deal);
    if (!legalStats.total) {
      addIssue("legal", "missing_legal", "No legal documents linked", "Attach at least one legal file or folder link so the deal package is complete.");
    } else if (legalStats.invalid > 0) {
      addIssue(
        "legal",
        "invalid_legal",
        `${legalStats.invalid} invalid legal link${legalStats.invalid === 1 ? "" : "s"}`,
        "One or more legal links are not valid http/https URLs."
      );
    }

    if (!hasDashboardLinked(deal)) {
      addIssue("setup", "missing_dashboard", "Dashboard not linked", "Attach the fundraising dashboard so investor activity is connected.");
    }
    if (!hasDeckLinked(deal)) {
      addIssue("setup", "missing_deck", "Deck not linked", "Add a deck link so the deal materials stay complete.");
    }

    const counts = ["ownership", "accounting", "legal", "setup"].reduce((accumulator, scope) => {
      accumulator[scope] = issues.filter((issue) => issue.scope === scope).length;
      return accumulator;
    }, { total: issues.length });
    counts.total = issues.length;

    const statuses = {
      ownership: ownershipAvailable ? (counts.ownership ? "attention" : "good") : (counts.ownership ? "attention" : "unknown"),
      accounting: counts.accounting ? "attention" : "good",
      legal: counts.legal ? "attention" : "good",
      setup: counts.setup ? "attention" : "good",
    };

    return {
      issues,
      topIssues: issues.slice(0, 3),
      counts,
      statuses,
      fullyLinked: issues.length === 0,
      accounting: {
        hasRetainer,
        invoices,
        invoiceSummary,
        latestInvoice: invoiceSummary.latest,
        primaryContact,
        nextPaymentDate,
      },
      legal: {
        links: legalLinks,
        stats: legalStats,
      },
    };
  }

  global.DealIntegrity = {
    normalizeValue,
    normalizeDateInput,
    normalizeCurrencyCode,
    parseAmount,
    getDealCurrency,
    getRetainerMonthly,
    hasPositiveRetainer,
    getPaymentIntervalMonths,
    getNextScheduledPaymentDateObject,
    normalizeDealContacts,
    getPrimaryDealContact,
    normalizeInvoiceStatus,
    normalizeDealInvoices,
    getLatestInvoice,
    isInvoiceOverdue,
    getInvoiceDisplayStatus,
    getInvoiceSummaryForDeal,
    normalizeDealLegalLinks,
    toSafeExternalUrl,
    getLegalLinkStats,
    getSubOwners,
    getDealOwners,
    hasDashboardLinked,
    hasDeckLinked,
    buildDealIntegrityReport,
  };
})(window);
