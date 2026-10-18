// Preload do dropdown de contas e dos diálogos (página local).
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("teams", {
  getState: () => ipcRenderer.invoke("get-state"),
  onState: (listener) => ipcRenderer.on("state", (_e, state) => listener(state)),
  onOpenMenu: (listener) => ipcRenderer.on("open-menu", () => listener()),
  setDialogOpen: (open) => ipcRenderer.send("dialog-open", open),
  setMenuOpen: (open) => ipcRenderer.send("menu-open", open),
  activate: (id) => ipcRenderer.send("activate", id),
  accountAction: (action) => ipcRenderer.send("account-action", action),
  addAccount: (account) => ipcRenderer.invoke("add-account", account),
  updateAccount: (account) => ipcRenderer.invoke("update-account", account),
  removeAccount: (id) => ipcRenderer.invoke("remove-account", id),
  skipLoading: () => ipcRenderer.send("skip-loading"),
  previewSound: (id) => ipcRenderer.send("preview-sound", id),
  testMeetingAlert: () => ipcRenderer.send("test-meeting-alert"),
  updateSettings: (settings) => ipcRenderer.invoke("update-settings", settings),
});
