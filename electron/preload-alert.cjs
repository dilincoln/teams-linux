// Preload da janela de aviso de reunião em tela cheia.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("meetingAlert", {
  onAlert: (listener) => ipcRenderer.on("alert", (_e, data) => listener(data)),
  action: (action) => ipcRenderer.send("alert-action", action),
});
