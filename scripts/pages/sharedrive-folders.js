(function initSharedriveFoldersPage() {
  const STORAGE_KEY_URL = "sharedrive_url_v1";
  const DEFAULT_SHARE_URL =
    "https://netorgft9359049-my.sharepoint.com/personal/mj_plutus-investment_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fmj%5Fplutus%2Dinvestment%5Fcom%2FDocuments%2FPlutus%20Investment%20Group%2FDeals&viewid=bf589b2c%2D3d69%2D4eb9%2D8cf4%2Dd755f4682b4d";

  function hasDesktopBridge() {
    return Boolean(window.PlutusDesktop && typeof window.PlutusDesktop.listShareDriveFolders === "function");
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

  function renderMeta(data) {
    const metaRow = document.getElementById("sharedrive-meta-row");
    if (!metaRow) return;

    const rootName = data && data.root && data.root.name ? data.root.name : "-";
    const total = data && typeof data.totalFolders === "number" ? data.totalFolders : 0;
    const fetchedAt = data && data.fetchedAt ? formatDateTime(data.fetchedAt) : "-";

    metaRow.innerHTML = "";
    [
      { label: "Root", value: rootName },
      { label: "Folders", value: total },
      { label: "Fetched", value: fetchedAt },
    ].forEach((chipData) => {
      const chip = document.createElement("div");
      chip.className = "chip";
      chip.innerHTML = `<strong>${chipData.value}</strong> ${chipData.label}`;
      metaRow.appendChild(chip);
    });
  }

  function renderFolders(data) {
    const list = document.getElementById("folders-list");
    const empty = document.getElementById("folders-empty");
    if (!list || !empty) return;

    const folders = data && Array.isArray(data.folders) ? data.folders : [];
    list.innerHTML = "";

    if (!folders.length) {
      empty.hidden = false;
      return;
    }

    empty.hidden = true;
    folders
      .slice()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
      .forEach((folder) => {
        const card = document.createElement("a");
        card.className = "folder-card";
        card.href = folder.webUrl || "javascript:void(0)";
        card.target = "_blank";
        card.rel = "noopener noreferrer";
        card.innerHTML = `
          <div class="folder-name">${folder.name || "(Unnamed folder)"}</div>
          <div class="folder-meta">
            <span>Items: ${folder.childCount == null ? "-" : folder.childCount}</span>
            <span>Size: ${bytesToReadable(folder.size)}</span>
            <span>Modified: ${formatDateTime(folder.lastModifiedDateTime)}</span>
          </div>
        `;
        list.appendChild(card);
      });
  }

  async function fetchFolders() {
    const urlInput = document.getElementById("share-url");
    const tokenInput = document.getElementById("access-token");
    const rememberUrlInput = document.getElementById("remember-url");

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
      setError("Desktop bridge not available. Open this page in the Electron app (`npm run desktop`).");
      return;
    }

    const result = await window.PlutusDesktop.listShareDriveFolders({ shareUrl, accessToken });
    if (!result || !result.ok) {
      setStatus("Failed");
      setError((result && result.error) || "Failed to load folders.");
      renderMeta(null);
      renderFolders(null);
      return;
    }

    setStatus("Connected");
    setError("");
    renderMeta(result.data);
    renderFolders(result.data);
  }

  document.addEventListener("DOMContentLoaded", () => {
    const urlInput = document.getElementById("share-url");
    const fetchButton = document.getElementById("fetch-folders-btn");

    const storedUrl = localStorage.getItem(STORAGE_KEY_URL);
    if (urlInput) {
      urlInput.value = storedUrl || DEFAULT_SHARE_URL;
    }

    if (fetchButton) fetchButton.addEventListener("click", fetchFolders);

    renderMeta(null);
    setStatus("Idle");
  });
})();
