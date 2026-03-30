(function initSharedriveFoldersPage() {
  const AppCore = window.AppCore;
  const appConfig = window.PlutusAppConfig || {};
  const storageKeys = Object.assign({
    sharedriveGate: "sharedrive_connected_v1",
  }, appConfig.storageKeys || {});
  const AUTH_DEBUG_KEY = "sharedrive_auth_debug_v1";
  const STORAGE_KEY_URL = "sharedrive_url_v2";
  const SHAREDRIVE_GATE_KEY = storageKeys.sharedriveGate;
  const DEFAULT_SHARE_URL =
    String(
      (window.SHAREDRIVE_TASKS &&
        window.SHAREDRIVE_TASKS.tasks &&
        window.SHAREDRIVE_TASKS.tasks.shareUrl) ||
      (window.SHAREDRIVE_TASKS && window.SHAREDRIVE_TASKS.shareUrl) ||
      "",
    ).trim();

  let deviceCodeState = null;
  let deviceCodeTimer = null;
  let emptyDefaultText = "No folders found at this location.";
  const connectionUiState = {
    status: "Idle",
    person: null,
    meta: null,
  };

  function hasDesktopBridge() {
    return Boolean(
      window.PlutusDesktop &&
        typeof window.PlutusDesktop.listShareDriveChildren === "function" &&
        typeof window.PlutusDesktop.uploadShareDriveFile === "function",
    );
  }

  function hasBrowserAuth() {
    return Boolean(
      AppCore &&
        typeof AppCore.requestGraphDeviceCode === "function" &&
        typeof AppCore.pollGraphDeviceCode === "function",
    );
  }

  function bytesToReadable(bytes) {
    if (typeof bytes !== "number" || !Number.isFinite(bytes) || bytes < 0) return "-";
    const sizes = ["B", "KB", "MB", "GB", "TB"];
    if (bytes === 0) return "0 B";
    const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), sizes.length - 1);
    const value = bytes / Math.pow(1024, i);
    return `${value.toFixed(value >= 10 ? 0 : 1)} ${sizes[i]}`;
  }

  function formatDateTime(value) {
    if (!value) return "-";
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

  function getDebugPanel() {
    return document.getElementById("connection-debug-panel");
  }

  function setDebugPanelOpen(isOpen) {
    const panel = getDebugPanel();
    if (panel) panel.open = Boolean(isOpen);
  }

  function getConfiguredShareUrl() {
    const urlInput = document.getElementById("share-url");
    const currentValue = String(urlInput && urlInput.value ? urlInput.value : "").trim();
    if (currentValue) return currentValue;
    if (treeState && treeState.shareUrl) return String(treeState.shareUrl).trim();
    try {
      return String(localStorage.getItem(STORAGE_KEY_URL) || DEFAULT_SHARE_URL || "").trim();
    } catch {
      return String(DEFAULT_SHARE_URL || "").trim();
    }
  }

  function setStatus(text) {
    const status = document.getElementById("folders-status");
    if (status) status.textContent = text;
    connectionUiState.status = text || "Idle";
    renderConnectionSummary();
  }

  function setError(message) {
    const errorEl = document.getElementById("folders-error");
    if (!errorEl) return;
    if (!message) {
      errorEl.hidden = true;
      errorEl.textContent = "";
      return;
    }
    errorEl.hidden = false;
    errorEl.textContent = message;
  }

  function setUploadStatus(message, isError) {
    const status = document.getElementById("upload-status");
    if (!status) return;
    status.textContent = message;
    status.style.color = isError ? "#fecdd3" : "";
  }

  function setDiagnosticsStatus(message, isError) {
    const status = document.getElementById("diagnostics-status");
    if (!status) return;
    status.textContent = message || "Idle";
    status.style.color = isError ? "#fecdd3" : "";
  }

  function setDiagnosticsOutput(text) {
    const output = document.getElementById("diagnostics-output");
    if (output) output.textContent = text || "";
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
      messageEl.textContent = payload && payload.message ? payload.message : "Follow the prompt to finish signing in.";
    }
    if (statusEl) statusEl.textContent = "Waiting for authorization...";
  }

  function setAuthDebug(message, isError) {
    const box = document.getElementById("auth-debug-box");
    if (!box) return;
    if (!message) {
      box.hidden = true;
      box.textContent = "";
      return;
    }
    box.hidden = false;
    box.textContent = message;
    box.style.color = isError ? "#fecdd3" : "";
    if (isError) setDebugPanelOpen(true);
  }

  function renderMeta(data) {
    const metaRow = document.getElementById("sharedrive-meta-row");
    if (!metaRow) return;
    connectionUiState.meta = data || null;
    renderConnectionSummary();

    const rootName = data && data.root && data.root.name ? data.root.name : "-";
    const totalFolders = data && typeof data.totalFolders === "number" ? data.totalFolders : 0;
    const totalFiles = data && typeof data.totalFiles === "number" ? data.totalFiles : 0;
    const fetchedAt = data && data.fetchedAt ? formatDateTime(data.fetchedAt) : "-";

    metaRow.innerHTML = "";
    [
      { label: "Workspace root", value: rootName },
      { label: "Visible folders", value: totalFolders },
      { label: "Visible files", value: totalFiles },
      { label: "Last refresh", value: fetchedAt },
    ].forEach((chipData) => {
      const chip = document.createElement("div");
      chip.className = "chip";
      chip.innerHTML = `<strong>${chipData.value}</strong> ${chipData.label}`;
      metaRow.appendChild(chip);
    });
  }

  function getConnectionTone(text) {
    const normalized = String(text || "").trim().toLowerCase();
    if (!normalized || normalized === "idle") return "idle";
    if (
      normalized.includes("failed") ||
      normalized.includes("error") ||
      normalized.includes("unavailable") ||
      normalized.includes("required") ||
      normalized.includes("expired")
    ) {
      return "error";
    }
    if (
      normalized.includes("loading") ||
      normalized.includes("starting") ||
      normalized.includes("awaiting") ||
      normalized.includes("sign-in") ||
      normalized.includes("uploading") ||
      normalized.includes("reading")
    ) {
      return "working";
    }
    if (
      normalized.includes("connected") ||
      normalized.includes("signed in") ||
      normalized.includes("complete") ||
      normalized.includes("valid")
    ) {
      return "success";
    }
    return "idle";
  }

  function setConnectionStepState(elementId, nextState) {
    const step = document.getElementById(elementId);
    if (!step) return;
    step.classList.toggle("is-active", nextState === "active");
    step.classList.toggle("is-complete", nextState === "complete");
  }

  function renderConnectionExperienceState({ person, meta, tone, configuredShareUrl }) {
    const body = document.body;
    const hasAccount = Boolean(person && person.connected);
    const hasWorkspace = Boolean(meta && meta.root && meta.root.name);
    if (body) {
      body.dataset.accountConnected = hasAccount ? "true" : "false";
      body.dataset.workspaceConnected = hasWorkspace ? "true" : "false";
      body.dataset.workspaceConfigured = configuredShareUrl ? "true" : "false";
      body.dataset.connectionTone = tone;
    }

    setConnectionStepState(
      "connection-step-account",
      hasAccount ? "complete" : (tone === "working" ? "active" : "idle"),
    );
    setConnectionStepState(
      "connection-step-workspace",
      hasWorkspace ? "complete" : (hasAccount ? "active" : "idle"),
    );
    setConnectionStepState(
      "connection-step-tools",
      hasWorkspace ? "complete" : "idle",
    );
  }

  function renderConnectionSummary(personOverride) {
    if (typeof personOverride !== "undefined") {
      connectionUiState.person = personOverride;
    }

    const pillEl = document.getElementById("connection-state-pill");
    const titleEl = document.getElementById("connection-state-title");
    const copyEl = document.getElementById("connection-state-copy");
    const nameEl = document.getElementById("connection-user-name");
    const emailEl = document.getElementById("connection-user-email");
    const metaEl = document.getElementById("connection-user-meta");
    const workspaceNameEl = document.getElementById("connection-workspace-name");
    const workspaceNoteEl = document.getElementById("connection-workspace-note");
    if (!pillEl || !titleEl || !copyEl || !nameEl || !emailEl || !metaEl || !workspaceNameEl || !workspaceNoteEl) return;

    const person = connectionUiState.person;
    const meta = connectionUiState.meta;
    const status = String(connectionUiState.status || "Idle").trim() || "Idle";
    const tone = getConnectionTone(status);
    const configuredShareUrl = getConfiguredShareUrl();
    renderConnectionExperienceState({
      person,
      meta,
      tone,
      configuredShareUrl,
    });
    pillEl.className = `connection-pill is-${tone}`;
    pillEl.textContent = status;

    if (person && person.connected) {
      nameEl.textContent = person.alias || person.displayName || "Connected account";
      emailEl.textContent = person.email || "Microsoft session active";
      metaEl.textContent = person.source ? `Signed in via ${person.source}` : "Microsoft session active";
    } else {
      nameEl.textContent = "No Microsoft account connected";
      emailEl.textContent = "Sign in to load the workspace with your session.";
      metaEl.textContent = "No workspace connected yet.";
    }

    if (meta && meta.root && meta.root.name) {
      workspaceNameEl.textContent = meta.root.name;
      workspaceNoteEl.textContent = meta.fetchedAt
        ? `Connected workspace last refreshed on ${formatDateTime(meta.fetchedAt)}.`
        : "Connected workspace is ready.";
    } else if (configuredShareUrl) {
      workspaceNameEl.textContent = "Saved workspace link ready";
      workspaceNoteEl.textContent = "The raw ShareDrive link stays hidden unless you open Advanced and debug.";
    } else {
      workspaceNameEl.textContent = "Workspace link needed";
      workspaceNoteEl.textContent = "Open Advanced and debug if you need to set or change the workspace link.";
    }

    if (tone === "error") {
      titleEl.textContent = "Connection needs attention";
      copyEl.textContent = status;
      return;
    }

    if (meta && meta.root && meta.root.name) {
      titleEl.textContent = `Workspace connected to ${meta.root.name}`;
      copyEl.textContent = meta.fetchedAt
        ? `The sharedrive tree is loaded and was last refreshed on ${formatDateTime(meta.fetchedAt)}.`
        : "The sharedrive tree is loaded and ready to browse.";
      return;
    }

    if (tone === "working") {
      titleEl.textContent = "Connection in progress";
      copyEl.textContent = "Complete Microsoft sign-in if prompted, then connect the workspace to load folders.";
      return;
    }

    if (person && person.connected) {
      titleEl.textContent = "Microsoft account connected";
      copyEl.textContent = configuredShareUrl
        ? "Your session is ready. Load the workspace to browse folders."
        : "Your session is ready. Open Advanced and debug if you need to add the workspace link.";
      return;
    }

    titleEl.textContent = "Ready to connect your workspace";
    copyEl.textContent = "Sign in with Microsoft first, then load the workspace the app should use.";
  }

  async function refreshConnectionSummary() {
    if (!AppCore || typeof AppCore.getCurrentConnectedPerson !== "function") {
      renderConnectionSummary(null);
      return null;
    }
    try {
      const person = await AppCore.getCurrentConnectedPerson();
      renderConnectionSummary(person);
      return person;
    } catch {
      renderConnectionSummary(null);
      return null;
    }
  }

  const treeState = {
    shareUrl: "",
    accessToken: "",
    root: null,
    loaded: new Set(),
  };
  const searchState = {
    query: "",
  };

  function sortItems(items) {
    return items
      .slice()
      .sort((a, b) => {
        if (Boolean(a.isFolder) !== Boolean(b.isFolder)) return a.isFolder ? -1 : 1;
        return String(a.name || "").localeCompare(String(b.name || ""));
      });
  }

  function buildMetaCounts(items) {
    const totalFolders = items.filter((item) => item.isFolder).length;
    const totalFiles = items.filter((item) => item.isFile).length;
    return { totalFolders, totalFiles };
  }

  function createTreeNode(item) {
    const node = document.createElement("div");
    node.className = "tree-node";
    node.dataset.itemId = item.id || "";
    node.dataset.isFolder = item.isFolder ? "1" : "0";
    node.dataset.name = String(item.name || "").toLowerCase();

    const row = document.createElement("div");
    row.className = "tree-row";

    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "tree-toggle";
    toggle.textContent = item.isFolder ? "▸" : "";
    toggle.disabled = !item.isFolder;
    toggle.setAttribute("aria-expanded", "false");
    row.appendChild(toggle);

    const name = document.createElement(item.webUrl ? "a" : "span");
    name.className = "tree-name";
    name.textContent = item.name || "(Unnamed item)";
    if (item.webUrl) {
      name.href = item.webUrl;
      name.target = "_blank";
      name.rel = "noopener noreferrer";
    }
    row.appendChild(name);

    const meta = document.createElement("div");
    meta.className = "tree-meta";
    const typeLabel = item.isFolder ? "Folder" : "File";
    const childLabel =
      item.isFolder && item.childCount != null ? `Items: ${item.childCount}` : "Items: -";
    meta.innerHTML = `
      <span>Type: ${typeLabel}</span>
      <span>${childLabel}</span>
      <span>Size: ${bytesToReadable(item.size)}</span>
      <span>Modified: ${formatDateTime(item.lastModifiedDateTime)}</span>
    `;
    row.appendChild(meta);

    node.appendChild(row);

    if (item.isFolder) {
      const children = document.createElement("div");
      children.className = "tree-children";
      children.hidden = true;
      node.appendChild(children);
    }

    return node;
  }

  function applyFolderSearch(query) {
    const list = document.getElementById("folders-list");
    const empty = document.getElementById("folders-empty");
    if (!list) return;
    const normalized = String(query || "").trim().toLowerCase();
    searchState.query = normalized;

    const nodes = Array.from(list.querySelectorAll(".tree-node"));
    if (!normalized) {
      if (empty) empty.textContent = emptyDefaultText;
      nodes.forEach((node) => {
        node.hidden = false;
        const toggle = node.querySelector(".tree-toggle");
        const children = node.querySelector(".tree-children");
        if (children && !node.classList.contains("is-open")) {
          children.hidden = true;
          if (toggle) toggle.setAttribute("aria-expanded", "false");
        }
      });
      if (empty) {
        empty.hidden = nodes.length > 0;
      }
      return;
    }

    const visible = new Set();
    nodes.forEach((node) => {
      if (node.dataset.isFolder !== "1") return;
      if ((node.dataset.name || "").includes(normalized)) {
        visible.add(node);
        let parent = node.parentElement ? node.parentElement.closest(".tree-node") : null;
        while (parent) {
          visible.add(parent);
          parent = parent.parentElement ? parent.parentElement.closest(".tree-node") : null;
        }
      }
    });

    nodes.forEach((node) => {
      if (node.dataset.isFolder !== "1") {
        node.hidden = true;
        return;
      }
      if (visible.has(node)) {
        node.hidden = false;
        const toggle = node.querySelector(".tree-toggle");
        const children = node.querySelector(".tree-children");
        if (children) {
          children.hidden = false;
          node.classList.add("is-open");
          if (toggle) toggle.setAttribute("aria-expanded", "true");
        }
      } else {
        node.hidden = true;
      }
    });

    const hasMatch = visible.size > 0;
    if (empty) {
      empty.textContent = hasMatch ? emptyDefaultText : "No folders match that search.";
      empty.hidden = hasMatch;
    }
  }

  function renderTree(items, container, emptyEl) {
    container.innerHTML = "";
    const sorted = sortItems(items);
    if (emptyEl) emptyEl.hidden = sorted.length > 0;
    if (!sorted.length) return;
    sorted.forEach((item) => container.appendChild(createTreeNode(item)));
  }

  async function requestChildren(parentItemId) {
    const payload = {
      shareUrl: treeState.shareUrl,
      accessToken: treeState.accessToken,
      parentItemId,
    };
    if (window.PlutusDesktop && typeof window.PlutusDesktop.listShareDriveChildren === "function") {
      try {
        return await window.PlutusDesktop.listShareDriveChildren(payload);
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : "Failed to load items." };
      }
    }
    if (AppCore && typeof AppCore.listShareDriveChildren === "function") {
      try {
        const result = await AppCore.listShareDriveChildren(payload);
        if (result && typeof result === "object" && Object.prototype.hasOwnProperty.call(result, "ok")) {
          return result;
        }
        return { ok: true, data: result };
      } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : "Failed to load items." };
      }
    }
    return { ok: false, error: "Listing items is unavailable on this platform." };
  }

  async function expandFolder(node) {
    const itemId = node.dataset.itemId;
    const children = node.querySelector(".tree-children");
    const toggle = node.querySelector(".tree-toggle");
    if (!itemId || !children || !toggle) return;

    node.classList.add("is-open");
    toggle.setAttribute("aria-expanded", "true");
    children.hidden = false;

    if (treeState.loaded.has(itemId)) return;

    children.innerHTML = `<div class="tree-loading">Loading...</div>`;
    node.dataset.loading = "true";

    const result = await requestChildren(itemId);
    node.dataset.loading = "false";

    if (!result || !result.ok) {
      children.innerHTML = `<div class="tree-error">${(result && result.error) || "Failed to load items."}</div>`;
      return;
    }

    treeState.loaded.add(itemId);
    const items = result.data && Array.isArray(result.data.items) ? result.data.items : [];
    if (!items.length) {
      children.innerHTML = `<div class="tree-empty">No items</div>`;
      return;
    }

    children.innerHTML = "";
    renderTree(items, children);
  }

  function collapseFolder(node) {
    const children = node.querySelector(".tree-children");
    const toggle = node.querySelector(".tree-toggle");
    if (!children || !toggle) return;
    node.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
    children.hidden = true;
  }

  async function fetchItems() {
    const urlInput = document.getElementById("share-url");
    const tokenInput = document.getElementById("access-token");
    const rememberUrlInput = document.getElementById("remember-url");
    const uploadFolderInput = document.getElementById("upload-folder-id");
    const list = document.getElementById("folders-list");
    const empty = document.getElementById("folders-empty");

    const shareUrl = String(urlInput && urlInput.value ? urlInput.value : "").trim();
    const accessToken = String(tokenInput && tokenInput.value ? tokenInput.value : "").trim();
    const rememberUrl = Boolean(rememberUrlInput && rememberUrlInput.checked);

    if (!shareUrl) {
      setStatus("Share URL required");
      setError("Share URL is required.");
      setDebugPanelOpen(true);
      return;
    }

    if (rememberUrl) {
      localStorage.setItem(STORAGE_KEY_URL, shareUrl);
    } else {
      localStorage.removeItem(STORAGE_KEY_URL);
    }

    setError("");
    setStatus("Loading workspace...");

    if (!hasDesktopBridge() && !(AppCore && typeof AppCore.listShareDriveChildren === "function")) {
      setStatus("Unavailable");
      setError(
        hasBrowserAuth()
          ? "Listing items requires the desktop bridge. Sign in to use mobile task sync."
          : "Desktop bridge not available. Open this page in the Electron app (`npm run desktop`).",
      );
      setDebugPanelOpen(true);
      return;
    }

    treeState.shareUrl = shareUrl;
    treeState.accessToken = accessToken;
    treeState.root = null;
    treeState.loaded.clear();

    const result = await requestChildren(null);
    if (!result || !result.ok) {
      setStatus("Failed");
      setError((result && result.error) || "Failed to load items.");
      setDebugPanelOpen(true);
      renderMeta(null);
      if (list) list.innerHTML = "";
      if (empty) empty.hidden = false;
      return;
    }

    setStatus("Workspace connected");
    setError("");
    setDebugPanelOpen(false);
    try {
      localStorage.setItem(SHAREDRIVE_GATE_KEY, "true");
    } catch {
      // ignore storage errors
    }
    const data = result.data || {};
    treeState.root = data.root || null;
    const items = Array.isArray(data.items) ? data.items : [];
    const counts = buildMetaCounts(items);
    renderMeta({
      root: data.root,
      totalFolders: counts.totalFolders,
      totalFiles: counts.totalFiles,
      fetchedAt: data.fetchedAt,
    });

    if (list) {
      renderTree(items, list, empty);
      if (searchState.query) {
        applyFolderSearch(searchState.query);
      }
    }

    if (data.parentItemId) {
      treeState.loaded.add(data.parentItemId);
    }

    if (uploadFolderInput && data.root && data.root.id) {
      if (!uploadFolderInput.value) uploadFolderInput.value = data.root.id;
    }
  }

  function clearDeviceCodeTimer() {
    if (deviceCodeTimer) {
      clearTimeout(deviceCodeTimer);
      deviceCodeTimer = null;
    }
  }

  async function startDeviceCodeFlow() {
    const useDesktop = window.PlutusDesktop && typeof window.PlutusDesktop.requestGraphDeviceCode === "function";
    const useBrowser = hasBrowserAuth();
    if (!useDesktop && !useBrowser) {
      setStatus("Sign-in unavailable");
      setError("Device code flow is unavailable. Update the Electron app to enable it.");
      setDebugPanelOpen(true);
      return;
    }

    setError("");
    setAuthDebug("");
    setStatus("Starting sign-in...");

    const result = useDesktop
      ? await window.PlutusDesktop.requestGraphDeviceCode()
      : await AppCore.requestGraphDeviceCode();
    if (!result || !result.ok) {
      setStatus("Failed to start sign-in");
      setError((result && result.error) || "Failed to start device code flow.");
      setAuthDebug(
        `Auth debug: bridge=${useDesktop ? "desktop" : "browser"} result=${JSON.stringify(result || {})}`,
        true,
      );
      return;
    }

    const data = result.data || {};
    deviceCodeState = {
      deviceCode: data.device_code,
      interval: Number(data.interval || 5),
      expiresAt: Date.now() + Number(data.expires_in || 0) * 1000,
    };

    setDeviceCodeBox(true, {
      userCode: data.user_code,
      verificationUri: data.verification_uri || data.verification_uri_complete,
      message: data.message,
    });
    setStatus("Awaiting sign-in...");
    setAuthDebug("");

    clearDeviceCodeTimer();
    deviceCodeTimer = setTimeout(pollDeviceCodeFlow, deviceCodeState.interval * 1000);
  }

  async function pollDeviceCodeFlow() {
    if (!deviceCodeState || !deviceCodeState.deviceCode) return;

    if (Date.now() > deviceCodeState.expiresAt) {
      setStatus("Device code expired");
      setDeviceCodeBox(false);
      deviceCodeState = null;
      return;
    }

    const result = deviceCodeState && deviceCodeState.deviceCode
      ? (
        (window.PlutusDesktop && typeof window.PlutusDesktop.pollGraphDeviceCode === "function")
          ? await window.PlutusDesktop.pollGraphDeviceCode({ deviceCode: deviceCodeState.deviceCode })
          : await AppCore.pollGraphDeviceCode(deviceCodeState.deviceCode)
      )
      : null;

    if (!result || !result.ok) {
      setError((result && result.error) || "Failed to poll device code.");
      setStatus("Sign-in failed");
      setDebugPanelOpen(true);
      setDeviceCodeBox(false);
      deviceCodeState = null;
      return;
    }

    const payload = result.data || {};
    if (payload.ok) {
      const tokenInput = document.getElementById("access-token");
      if (tokenInput) tokenInput.value = payload.accessToken || "";
      setStatus("Signed in");
      try {
        localStorage.setItem(SHAREDRIVE_GATE_KEY, "true");
      } catch {
        // ignore storage errors
      }
      window.dispatchEvent(new CustomEvent("appcore:graph-session-updated"));
      setDeviceCodeBox(false);
      deviceCodeState = null;
      await refreshConnectionSummary();
      if (getConfiguredShareUrl()) {
        await fetchItems();
      } else {
        setDebugPanelOpen(true);
      }
      return;
    }

    if (payload.error === "authorization_pending") {
      clearDeviceCodeTimer();
      deviceCodeTimer = setTimeout(pollDeviceCodeFlow, deviceCodeState.interval * 1000);
      return;
    }

    if (payload.error === "slow_down") {
      deviceCodeState.interval = Math.max(deviceCodeState.interval + 5, 10);
      clearDeviceCodeTimer();
      deviceCodeTimer = setTimeout(pollDeviceCodeFlow, deviceCodeState.interval * 1000);
      return;
    }

    setError(payload.error_description || "Sign-in failed.");
    setStatus("Sign-in failed");
    setDebugPanelOpen(true);
    setDeviceCodeBox(false);
    deviceCodeState = null;
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error("Failed to read file."));
      reader.readAsDataURL(file);
    });
  }

  async function uploadFile() {
    const urlInput = document.getElementById("share-url");
    const tokenInput = document.getElementById("access-token");
    const folderInput = document.getElementById("upload-folder-id");
    const fileInput = document.getElementById("upload-file");
    const conflictInput = document.getElementById("upload-conflict");

    const shareUrl = String(urlInput && urlInput.value ? urlInput.value : "").trim();
    const accessToken = String(tokenInput && tokenInput.value ? tokenInput.value : "").trim();
    const parentItemId = String(folderInput && folderInput.value ? folderInput.value : "").trim();
    const conflictBehavior = String(conflictInput && conflictInput.value ? conflictInput.value : "").trim();
    const file = fileInput && fileInput.files ? fileInput.files[0] : null;

    if (!shareUrl) {
      setUploadStatus("Share URL is required.", true);
      return;
    }

    if (!file) {
      setUploadStatus("Choose a file to upload.", true);
      return;
    }

    if (!hasDesktopBridge()) {
      setUploadStatus("Desktop bridge not available.", true);
      return;
    }

    setUploadStatus("Reading file...");

    let dataUrl;
    try {
      dataUrl = await readFileAsDataUrl(file);
    } catch (err) {
      setUploadStatus("Failed to read file.", true);
      return;
    }

    const base64 = String(dataUrl || "").split(",")[1] || "";
    if (!base64) {
      setUploadStatus("File data is empty.", true);
      return;
    }

    setUploadStatus("Uploading...");
    const result = await window.PlutusDesktop.uploadShareDriveFile({
      shareUrl,
      accessToken,
      parentItemId,
      fileName: file.name,
      contentBase64: base64,
      conflictBehavior,
    });

    if (!result || !result.ok) {
      setUploadStatus((result && result.error) || "Upload failed.", true);
      return;
    }

    setUploadStatus("Upload complete.");
    await fetchItems();
  }

  async function inspectSharedJson(kind) {
    if (!AppCore || typeof AppCore.inspectSharedJsonSource !== "function") {
      setDiagnosticsStatus("Unavailable", true);
      setDiagnosticsOutput("Shared diagnostics are unavailable on this platform.");
      return;
    }

    setDiagnosticsStatus(`Inspecting ${kind}...`);
    setDiagnosticsOutput("Running sharedrive diagnostics...");

    const result = await AppCore.inspectSharedJsonSource(kind);
    const lines = [
      `kind: ${result.kind || kind}`,
      `enabled: ${result.enabled ? "yes" : "no"}`,
      `fileName: ${result.fileName || "-"}`,
      `shareUrl: ${result.shareUrl || "-"}`,
      `parentItemId: ${result.parentItemId || "-"}`,
      `parsed: ${result.parsed ? "yes" : "no"}`,
      `parsedType: ${result.parsedType || "-"}`,
      `itemCount: ${typeof result.itemCount === "number" ? result.itemCount : 0}`,
      `error: ${result.error || "-"}`,
      "",
      "preview:",
      result.preview || "(empty response)",
    ];

    setDiagnosticsStatus(result.ok ? "Valid JSON" : "Problem found", !result.ok);
    setDiagnosticsOutput(lines.join("\n"));
  }

  document.addEventListener("DOMContentLoaded", () => {
    const urlInput = document.getElementById("share-url");
    const fetchButton = document.getElementById("fetch-folders-btn");
    const deviceCodeButton = document.getElementById("device-code-btn");
    const uploadButton = document.getElementById("upload-btn");
    const debugTasksButton = document.getElementById("debug-tasks-btn");
    const debugDealsButton = document.getElementById("debug-deals-btn");
    const debugConfigButton = document.getElementById("debug-config-btn");
    const list = document.getElementById("folders-list");
    const searchInput = document.getElementById("folder-search");
    const empty = document.getElementById("folders-empty");

    const storedUrl = localStorage.getItem(STORAGE_KEY_URL);
    if (urlInput) {
      urlInput.value = storedUrl || DEFAULT_SHARE_URL;
      urlInput.addEventListener("input", () => {
        renderConnectionSummary();
      });
    }
    if (!(storedUrl || DEFAULT_SHARE_URL)) {
      setDebugPanelOpen(true);
    }

    if (fetchButton) fetchButton.addEventListener("click", fetchItems);
    if (deviceCodeButton) deviceCodeButton.addEventListener("click", startDeviceCodeFlow);
    if (uploadButton) uploadButton.addEventListener("click", uploadFile);
    if (debugTasksButton) debugTasksButton.addEventListener("click", () => inspectSharedJson("tasks"));
    if (debugDealsButton) debugDealsButton.addEventListener("click", () => inspectSharedJson("deals"));
    if (debugConfigButton) debugConfigButton.addEventListener("click", () => inspectSharedJson("config"));
    if (empty && empty.textContent) emptyDefaultText = empty.textContent;
    if (searchInput) {
      searchInput.addEventListener("input", (event) => {
        applyFolderSearch(event.target && event.target.value ? event.target.value : "");
      });
    }
    if (list) {
      list.addEventListener("click", (event) => {
        const toggle = event.target && event.target.closest ? event.target.closest(".tree-toggle") : null;
        if (!toggle) return;
        const node = toggle.closest(".tree-node");
        if (!node || node.dataset.isFolder !== "1") return;
        if (node.classList.contains("is-open")) {
          collapseFolder(node);
        } else {
          expandFolder(node);
        }
      });
    }

    renderMeta(null);
    setStatus("Idle");
    setUploadStatus("Idle");
    setDiagnosticsStatus("Idle");
    refreshConnectionSummary();
    window.addEventListener("appcore:graph-session-updated", () => {
      refreshConnectionSummary();
    });
  });
})();
