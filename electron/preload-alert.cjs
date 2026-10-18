// Preload for the fullscreen meeting alert window.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("meetingAlert", {
  onAlert: (listener) => ipcRenderer.on("alert", (_e, data) => listener(data)),
  action: (action) => ipcRenderer.send("alert-action", action),
});
