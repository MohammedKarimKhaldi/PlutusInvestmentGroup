(function initSharedriveFoldersPage() {
  const AppCore = window.AppCore;
  const AUTH_DEBUG_KEY = "sharedrive_auth_debug_v1";
  const STORAGE_KEY_URL = "sharedrive_url_v2";
  const DEFAULT_SHARE_URL =
    "https://netorgft9359049-my.sharepoint.com/personal/mj_plutus-investment_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fmj%5Fplutus%2Dinvestment%5Fcom%2FDocuments%2FPlutus%20Investment%20Group%2FDeals&viewid=bf589b2c%2D3d69%2D4eb9%2D8cf4%2Dd755f4682b4d";

  let deviceCodeState = null;
  let deviceCodeTimer = null;

  function hasDesktopBridge() {
    return Boolean(
      window.PlutusDesktop &&
        typeof window.PlutusDesktop.listShareDriveItems === "function" &&
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

  function setStatus(text) {
    const status = document.getElementById("folders-status");
    if (status) status.textContent = text;
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
    if (statusEl) statusEl.textContent = "Waiting for authorization…";
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
  }

  function renderMeta(data) {
    const metaRow = document.getElementById("sharedrive-meta-row");
    if (!metaRow) return;

    const rootName = data && data.root && data.root.name ? data.root.name : "-";
    const totalFolders = data && typeof data.totalFolders === "number" ? data.totalFolders : 0;
    const totalFiles = data && typeof data.totalFiles === "number" ? data.totalFiles : 0;
    const fetchedAt = data && data.fetchedAt ? formatDateTime(data.fetchedAt) : "-";

    metaRow.innerHTML = "";
    [
      { label: "Root", value: rootName },
      { label: "Folders", value: totalFolders },
      { label: "Files", value: totalFiles },
      { label: "Fetched", value: fetchedAt },
    ].forEach((chipData) => {
      const chip = document.createElement("div");
      chip.className = "chip";
      chip.innerHTML = `<strong>${chipData.value}</strong> ${chipData.label}`;
      metaRow.appendChild(chip);
    });
  }

  function renderItems(data) {
    const list = document.getElementById("folders-list");
    const empty = document.getElementById("folders-empty");
    if (!list || !empty) return;

    const items = data && Array.isArray(data.items) ? data.items : [];
    list.innerHTML = "";

    if (!items.length) {
      empty.hidden = false;
      return;
    }

    empty.hidden = true;
    items
      .slice()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
      .forEach((item) => {
        const card = document.createElement("a");
        card.className = "item-card";
        card.href = item.webUrl || "javascript:void(0)";
        card.target = "_blank";
        card.rel = "noopener noreferrer";
        const typeLabel = item.isFolder ? "Folder" : "File";
        const childLabel =
          item.isFolder && item.childCount != null ? `Items: ${item.childCount}` : "Items: -";
        card.innerHTML = `
          <div class="item-name">${item.name || "(Unnamed item)"}</div>
          <div class="item-meta">
            <span>Type: ${typeLabel}</span>
            <span>${childLabel}</span>
            <span>Size: ${bytesToReadable(item.size)}</span>
            <span>Modified: ${formatDateTime(item.lastModifiedDateTime)}</span>
          </div>
        `;
        list.appendChild(card);
      });
  }

  async function fetchItems() {
    const urlInput = document.getElementById("share-url");
    const tokenInput = document.getElementById("access-token");
    const rememberUrlInput = document.getElementById("remember-url");
    const uploadFolderInput = document.getElementById("upload-folder-id");

    const shareUrl = String(urlInput && urlInput.value ? urlInput.value : "").trim();
    const accessToken = String(tokenInput && tokenInput.value ? tokenInput.value : "").trim();
    const rememberUrl = Boolean(rememberUrlInput && rememberUrlInput.checked);

    if (!shareUrl) {
      setError("Share URL is required.");
      return;
    }

    if (rememberUrl) {
      localStorage.setItem(STORAGE_KEY_URL, shareUrl);
    } else {
      localStorage.removeItem(STORAGE_KEY_URL);
    }

    setError("");
    setStatus("Loading...");

    if (!hasDesktopBridge()) {
      setStatus("Unavailable");
      setError(
        hasBrowserAuth()
          ? "Listing items requires the Electron app. Use Sign in with Microsoft for mobile task sync."
          : "Desktop bridge not available. Open this page in the Electron app (`npm run desktop`).",
      );
      return;
    }

    const result = await window.PlutusDesktop.listShareDriveItems({ shareUrl, accessToken });
    if (!result || !result.ok) {
      setStatus("Failed");
      setError((result && result.error) || "Failed to load items.");
      renderMeta(null);
      renderItems(null);
      return;
    }

    setStatus("Connected");
    setError("");
    renderMeta(result.data);
    renderItems(result.data);

    if (uploadFolderInput && result.data && result.data.root && result.data.root.id) {
      if (!uploadFolderInput.value) uploadFolderInput.value = result.data.root.id;
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
      setError("Device code flow is unavailable. Update the Electron app to enable it.");
      return;
    }

    setError("");
    setAuthDebug("");
    setStatus("Starting sign-in...");

    const result = useDesktop
      ? await window.PlutusDesktop.requestGraphDeviceCode()
      : await AppCore.requestGraphDeviceCode();
    if (!result || !result.ok) {
      setStatus("Idle");
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
    setAuthDebug(
      `Auth debug: bridge=${useDesktop ? "desktop" : "browser"} client_id=${(window.SHAREDRIVE_TASKS && window.SHAREDRIVE_TASKS.azureClientId) || "n/a"} tenant=${(window.SHAREDRIVE_TASKS && window.SHAREDRIVE_TASKS.azureTenantId) || "common"}`,
      false,
    );

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
      setDeviceCodeBox(false);
      deviceCodeState = null;
      return;
    }

    const payload = result.data || {};
    if (payload.ok) {
      const tokenInput = document.getElementById("access-token");
      if (tokenInput) tokenInput.value = payload.accessToken || "";
      setStatus("Signed in");
      setDeviceCodeBox(false);
      deviceCodeState = null;
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

  document.addEventListener("DOMContentLoaded", () => {
    const urlInput = document.getElementById("share-url");
    const fetchButton = document.getElementById("fetch-folders-btn");
    const deviceCodeButton = document.getElementById("device-code-btn");
    const uploadButton = document.getElementById("upload-btn");

    const storedUrl = localStorage.getItem(STORAGE_KEY_URL);
    if (urlInput) {
      urlInput.value = storedUrl || DEFAULT_SHARE_URL;
    }

    if (fetchButton) fetchButton.addEventListener("click", fetchItems);
    if (deviceCodeButton) deviceCodeButton.addEventListener("click", startDeviceCodeFlow);
    if (uploadButton) uploadButton.addEventListener("click", uploadFile);

    renderMeta(null);
    setStatus("Idle");
    setUploadStatus("Idle");
  });
})();
