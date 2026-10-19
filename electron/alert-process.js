// Fullscreen meeting alert, one window per monitor. Runs as its own process, started by the app
// with --ozone-platform=x11: on Wayland an app can't choose the monitor of its windows (they all
// open on the same one), while under XWayland each window opens fullscreen on its monitor.
//
// Talks to the app through JSON lines: stdin receives { type: "alert", data }, stdout sends
// { type: "action", action } ("join" | "snooze" | "dismiss"). Any action closes every window.
import { app, BrowserWindow, ipcMain, screen } from "electron";
import path from "node:path";
import readline from "node:readline";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

// Its own profile: the app's one is locked by the app.
app.setPath("userData", path.join(app.getPath("userData"), "meeting-alert"));

const windows = [];
let data = null;
let closing = false;

const send = (message) => process.stdout.write(`${JSON.stringify(message)}\n`);

function finish(action) {
  if (closing) return;
  closing = true;
  send({ type: "action", action });
  for (const win of windows) if (!win.isDestroyed()) win.destroy();
  app.quit();
}

function showAlert() {
  if (!data || !windows.length || !windows.every((win) => win.isVisible())) return;
  for (const win of windows) win.webContents.send("alert", data);
  // Keyboard focus (Enter/Esc) on the monitor where the mouse is.
  const cursor = screen.getDisplayNearestPoint(screen.getCursorScreenPoint()).id;
  (windows.find((win) => win.displayId === cursor) ?? windows[0]).focus();
}

function open() {
  for (const display of screen.getAllDisplays()) {
    const win = new BrowserWindow({
      ...display.bounds,
      fullscreen: true,
      frame: false,
      alwaysOnTop: true,
      skipTaskbar: true,
      show: false,
      icon: path.join(here, "assets", "icon.png"),
      title: "Meeting — Teams",
      backgroundColor: "#0d0d12",
      webPreferences: {
        preload: path.join(here, "preload-alert.cjs"),
        contextIsolation: true,
        sandbox: true,
      },
    });
    win.displayId = display.id;
    win.setMenu(null);
    win.loadFile(path.join(here, "ui", "alert.html"));
    win.once("ready-to-show", () => {
      win.showInactive();
      showAlert();
    });
    // Closed some other way (Alt+F4): same as Dismiss, and the other monitors close too.
    win.on("close", () => finish("dismiss"));
    windows.push(win);
  }
}

ipcMain.on("alert-action", (e, action) => {
  if (windows.some((win) => win.webContents === e.sender)) finish(action);
});

readline.createInterface({ input: process.stdin }).on("line", (line) => {
  let message;
  try {
    message = JSON.parse(line);
  } catch {
    return;
  }
  if (message.type !== "alert") return;
  data = message.data;
  showAlert();
});
// The app went away: so does the alert.
process.stdin.on("end", () => finish("dismiss"));

app.whenReady().then(open);
