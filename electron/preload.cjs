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
});
