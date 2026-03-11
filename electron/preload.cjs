const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("PlutusDesktop", {
  readArrayStore(key) {
    return ipcRenderer.sendSync("plutus:store:read-array", key);
  },
  writeArrayStore(key, values) {
    return ipcRenderer.sendSync("plutus:store:write-array", { key, values });
  },
  listShareDriveFolders(payload) {
    return ipcRenderer.invoke("plutus:sharedrive:list-folders", payload || {});
  },
  listShareDriveItems(payload) {
    return ipcRenderer.invoke("plutus:sharedrive:list-items", payload || {});
  },
  getShareDriveDownloadUrl(payload) {
    return ipcRenderer.invoke("plutus:sharedrive:get-download-url", payload || {});
  },
  downloadShareDriveFile(payload) {
    return ipcRenderer.invoke("plutus:sharedrive:download-file", payload || {});
  },
  uploadShareDriveFile(payload) {
    return ipcRenderer.invoke("plutus:sharedrive:upload-file", payload || {});
  },
  requestGraphDeviceCode() {
    return ipcRenderer.invoke("plutus:graph:device-code");
  },
  pollGraphDeviceCode(payload) {
    return ipcRenderer.invoke("plutus:graph:device-code:poll", payload || {});
  },
});
