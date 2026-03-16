#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");

const {
  sourceAppDir,
  dataFiles,
} = require("../config/project-paths.cjs");

function readJson(filePath) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return null;
  }
}

function toBase64Url(value) {
  return Buffer.from(String(value || ""), "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function getStoreDirCandidates() {
  const teamConfig = readJson(path.join(sourceAppDir, "data", dataFiles.teamStorePath));
  const candidates = [];
  if (teamConfig && typeof teamConfig.storeDir === "string" && teamConfig.storeDir.trim()) {
    candidates.push(teamConfig.storeDir.trim());
  }

  const productName = "Plutus Investment Dashboard";
  const homeDir = os.homedir();
  candidates.push(path.join(homeDir, "Library", "Application Support", productName, "runtime-store"));
  candidates.push(path.join(homeDir, ".config", productName, "runtime-store"));
  if (process.env.APPDATA) {
    candidates.push(path.join(process.env.APPDATA, productName, "runtime-store"));
  }
  return candidates;
}

function resolveAccessToken() {
  if (process.env.PLUTUS_GRAPH_ACCESS_TOKEN) {
    return process.env.PLUTUS_GRAPH_ACCESS_TOKEN.trim();
  }

  for (const storeDir of getStoreDirCandidates()) {
    const session = readJson(path.join(storeDir, "graph_session_v1.json"));
    if (session && typeof session.accessToken === "string" && session.accessToken.trim()) {
      return session.accessToken.trim();
    }
  }

  return "";
}

function getSharedConfig(kind) {
  const dashboardConfig = readJson(path.join(sourceAppDir, "data", dataFiles.config)) || {};
  const sharedConfig = readJson(path.join(sourceAppDir, "data", dataFiles.sharedTasks)) || {};
  const sharedTasks = sharedConfig.tasks || {};
  const sharedDealsNested = sharedConfig.deals || {};
  const sharedConfigNested = sharedConfig.config || {};
  const settings = dashboardConfig.settings || {};

  if (kind === "tasks") {
    return {
      enabled: Boolean(sharedTasks.enabled),
      shareUrl: String(sharedTasks.shareUrl || "").trim(),
      parentItemId: String(sharedTasks.parentItemId || "").trim(),
      fileName: String(sharedTasks.fileName || dataFiles.sharedTasks).trim(),
    };
  }

  if (kind === "deals") {
    return {
      enabled: Boolean(sharedDealsNested.enabled || (settings.sharedDeals && settings.sharedDeals.enabled)),
      shareUrl: String(
        sharedDealsNested.shareUrl || (settings.sharedDeals && settings.sharedDeals.shareUrl) || "",
      ).trim(),
      parentItemId: String(sharedDealsNested.parentItemId || "").trim(),
      fileName: String(sharedDealsNested.fileName || dataFiles.deals).trim(),
    };
  }

  if (kind === "config") {
    const dealsShareUrl = String(
      sharedDealsNested.shareUrl || (settings.sharedDeals && settings.sharedDeals.shareUrl) || "",
    ).trim();
    return {
      enabled: Boolean(sharedConfigNested.enabled || dealsShareUrl),
      shareUrl: String(sharedConfigNested.shareUrl || dealsShareUrl || "").trim(),
      parentItemId: String(sharedConfigNested.parentItemId || sharedDealsNested.parentItemId || "").trim(),
      fileName: String(sharedConfigNested.fileName || dataFiles.config).trim(),
    };
  }

  return null;
}

async function fetchJson(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }
  if (!response.ok) {
    const message =
      (payload && payload.error && (payload.error.message || payload.error_description)) ||
      response.statusText ||
      "Request failed";
    throw new Error(message);
  }
  return payload;
}

async function getShareDriveItem(shareUrl, token, selectFields) {
  const encodedShare = `u!${toBase64Url(shareUrl)}`;
  const fields = Array.isArray(selectFields) && selectFields.length
    ? selectFields
    : ["id", "name", "parentReference", "remoteItem", "folder", "file"];
  const params = new URLSearchParams({ $select: fields.join(",") });
  return fetchJson(`https://graph.microsoft.com/v1.0/shares/${encodedShare}/driveItem?${params.toString()}`, token);
}

async function listDriveChildren(driveId, itemId, token) {
  const params = new URLSearchParams({
    $top: "200",
    $select: "id,name,parentReference,file,folder",
  });
  return fetchJson(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?${params.toString()}`,
    token,
  );
}

async function resolveFile(rootConfig, token) {
  const item = await getShareDriveItem(rootConfig.shareUrl, token);
  const driveId =
    (item && item.parentReference && item.parentReference.driveId) ||
    (item && item.remoteItem && item.remoteItem.parentReference && item.remoteItem.parentReference.driveId) ||
    "";
  if (!driveId) throw new Error("Unable to resolve drive ID.");

  if (item && item.file && String(item.name || "").trim().toLowerCase() === rootConfig.fileName.toLowerCase()) {
    return { driveId, itemId: item.id, itemName: item.name || rootConfig.fileName };
  }

  const parentItemId =
    rootConfig.parentItemId ||
    (item && item.folder && item.id) ||
    (item && item.parentReference && item.parentReference.id) ||
    "";
  if (!parentItemId) throw new Error("Unable to resolve parent folder.");

  const listing = await listDriveChildren(driveId, parentItemId, token);
  const children = Array.isArray(listing && listing.value) ? listing.value : [];
  const target = children.find((entry) => String(entry && entry.name || "").trim().toLowerCase() === rootConfig.fileName.toLowerCase());
  if (!target || !target.id) {
    throw new Error(`Shared file not found: ${rootConfig.fileName}`);
  }
  return { driveId, itemId: target.id, itemName: target.name || rootConfig.fileName };
}

async function downloadFileText(rootConfig, token) {
  const resolved = await resolveFile(rootConfig, token);
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(resolved.driveId)}/items/${encodeURIComponent(resolved.itemId)}/content`,
    {
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Download failed (${response.status})`);
  }
  return { text, itemName: resolved.itemName };
}

async function main() {
  const kind = String(process.argv[2] || "tasks").trim().toLowerCase();
  const config = getSharedConfig(kind);
  if (!config) {
    console.error(`Unknown kind: ${kind}`);
    process.exit(1);
  }

  console.log(`kind: ${kind}`);
  console.log(`enabled: ${config.enabled ? "yes" : "no"}`);
  console.log(`fileName: ${config.fileName || "-"}`);
  console.log(`shareUrl: ${config.shareUrl || "-"}`);
  console.log(`parentItemId: ${config.parentItemId || "-"}`);

  if (!config.enabled || !config.shareUrl) {
    process.exit(1);
  }

  const token = resolveAccessToken();
  if (!token) {
    console.error("\nNo Graph access token found.");
    console.error("Use PLUTUS_GRAPH_ACCESS_TOKEN=... or sign in once in the desktop app.");
    process.exit(1);
  }

  const { text, itemName } = await downloadFileText(config, token);
  console.log(`resolvedItem: ${itemName}`);
  console.log(`contentLength: ${text.length}`);

  try {
    const parsed = JSON.parse(text);
    const itemCount = Array.isArray(parsed) ? parsed.length : Array.isArray(parsed.tasks) ? parsed.tasks.length : 0;
    console.log(`parsed: yes`);
    console.log(`parsedType: ${Array.isArray(parsed) ? "array" : typeof parsed}`);
    console.log(`itemCount: ${itemCount}`);
  } catch (error) {
    console.log("parsed: no");
    console.log(`parseError: ${error instanceof Error ? error.message : "Invalid JSON"}`);
  }

  console.log("\npreview:");
  console.log(text.slice(0, 1200) || "(empty response)");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
