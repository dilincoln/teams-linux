// Preload for the Teams pages (sandbox + contextIsolation).
// Exposes only the notification request to the page; clicks come back via postMessage.
const { contextBridge, ipcRenderer } = require("electron");

const CLICK_MESSAGE = "teams-linux:notification-click";
const KEEP_ACTIVE_MESSAGE = "teams-linux:keep-active";
const CALL_MESSAGE = "teams-linux:call";

// A meeting window (its own window, outside the main one) only runs the meeting: no account
// button, next meeting or calendar. The main process ignores its notifications and unread.
let meetingWindow = ipcRenderer.sendSync("is-meeting-window") === true;

ipcRenderer.on("notification-click", (_event, id) => {
  window.postMessage({ type: CLICK_MESSAGE, id }, "*");
});
ipcRenderer.on("keep-active", (_event, on) => {
  window.postMessage({ type: KEEP_ACTIVE_MESSAGE, on: !!on }, "*");
});
// The page was moved into a meeting window, with the call going on.
ipcRenderer.on("meeting-window", () => {
  meetingWindow = true;
  switcher?.remove();
  nextMeetingEl?.remove();
});
// From the page (see below): whether a call is connected.
window.addEventListener("message", (e) => {
  if (e.source === window && e.data?.type === CALL_MESSAGE) ipcRenderer.send("call-state", !!e.data.active);
});

contextBridge.exposeInMainWorld("__teamsLinux", {
  notify: (id, title, body) => ipcRenderer.send("notify", { id, title: String(title), body: String(body) }),
});

// Account button pinned to the top-left corner, in place of Teams' app launcher and sidebar
// toggle (the CSS that hides them comes from the main process).
const TEAMS_HOSTS = /(^|\.)(teams\.microsoft\.com|teams\.cloud\.microsoft|teams\.live\.com)$/;
let switcherState = null; // { name, color, others }
let switcher = null;

function el(tag, className) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

function buildSwitcher() {
  const root = el("div");
  root.id = "teams-linux-switcher";
  root.setAttribute("role", "button");
  root.setAttribute("aria-haspopup", "menu");
  root.tabIndex = 0;
  root.title = "Switch account";

  const avatar = el("span", "tls-avatar");
  avatar.append(el("span", "tls-initial"));
  // Bell with the other accounts' unread count (Teams already shows the open account's).
  const bell = el("span", "tls-bell");
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  for (const d of ["M8 1.6a4 4 0 0 0-4 4v2.6l-1.2 2.2a.6.6 0 0 0 .5.9h9.4a.6.6 0 0 0 .5-.9L12 8.2V5.6a4 4 0 0 0-4-4z", "M6.3 13a1.8 1.8 0 0 0 3.4 0"]) {
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", "currentColor");
    path.setAttribute("stroke-width", "1.3");
    path.setAttribute("stroke-linejoin", "round");
    svg.append(path);
  }
  bell.append(svg, el("span", "tls-badge"));
  root.append(avatar, el("span", "tls-name"), bell);

  const open = () => ipcRenderer.send("open-accounts-menu");
  root.addEventListener("click", open);
  root.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    e.preventDefault();
    open();
  });
  return root;
}

function renderSwitcher() {
  if (!switcher || !switcherState) return;
  const { name, color, others } = switcherState;
  switcher.querySelector(".tls-initial").textContent = (name.trim()[0] || "?").toUpperCase();
  switcher.querySelector(".tls-avatar").style.background = color;
  switcher.querySelector(".tls-name").textContent = name;
  const badge = switcher.querySelector(".tls-badge");
  badge.hidden = others === 0;
  badge.textContent = others > 99 ? "99+" : String(others);
  switcher.querySelector(".tls-bell").classList.toggle("tls-active", others > 0);
  switcher.title = others > 0 ? `Switch account — ${others} unread in other accounts` : "Switch account";
  // Same text color as the Teams header buttons (follows the light/dark theme).
  const reference = document.querySelector('[data-tid="waffle-open-button"]');
  if (reference) switcher.style.color = getComputedStyle(reference).color;
}

function mountSwitcher() {
  if (meetingWindow || !TEAMS_HOSTS.test(location.hostname) || !document.documentElement) return;
  switcher ??= buildSwitcher();
  // Outside <body>, so Teams' React doesn't remove it.
  if (!switcher.isConnected) document.documentElement.append(switcher);
  renderSwitcher();
}

ipcRenderer.on("switcher-state", (_event, state) => {
  switcherState = state;
  mountSwitcher();
});
window.addEventListener("DOMContentLoaded", () => {
  mountSwitcher();
  setInterval(mountSwitcher, 3000);
});

// Tells the main process when Teams has finished loading (accounts are opened one by one
// at startup). Outside Teams (sign-in) there is nothing to wait for.
// The navigation bar only exists once the app is mounted; the "all rendered" marker doesn't
// always show up when the account loads behind another one.
const TEAMS_READY_SELECTOR = '[data-tid="all-phased-rendering-complete"], [data-tid="app-layout-area--nav"]';
window.addEventListener("DOMContentLoaded", () => {
  if (meetingWindow) return;
  if (!TEAMS_HOSTS.test(location.hostname)) {
    ipcRenderer.send("account-ready");
    return;
  }
  const timer = setInterval(() => {
    if (!document.querySelector(TEAMS_READY_SELECTOR)) return;
    clearInterval(timer);
    ipcRenderer.send("account-ready");
  }, 500);
});

// ---- Calendar: upcoming meetings from Microsoft Graph ----
// Uses the calendar token Teams itself keeps (MSAL cache in localStorage); Teams renews it
// on its own. Events go to the main process, which merges the accounts, shows the next
// meeting and fires the fullscreen alert.
const CALENDAR_EVERY = 5 * 60_000;
const CALENDAR_RETRY = 60_000;

function calendarToken() {
  const now = Date.now() / 1000;
  let best = null;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (!/accesstoken/i.test(key)) continue;
    let entry;
    try {
      entry = JSON.parse(localStorage.getItem(key));
    } catch {
      continue;
    }
    const target = String(entry?.target || "");
    if (!/graph\.microsoft\.com/i.test(target) || !/Calendars\.Read/i.test(target)) continue;
    const expires = Number(entry.expiresOn);
    if (typeof entry.secret !== "string" || !(expires > now + 60)) continue;
    if (!best || expires > best.expires) best = { secret: entry.secret, expires };
  }
  return best?.secret ?? null;
}

// Graph returns "2026-10-19T13:00:00.0000000" (UTC, with 7 decimal places).
const graphTime = (value) => Date.parse(`${String(value).slice(0, 19)}Z`);

async function fetchCalendar() {
  const token = calendarToken();
  if (!token) return false;
  const now = Date.now();
  const params = new URLSearchParams({
    startDateTime: new Date(now - 12 * 3600_000).toISOString(),
    endDateTime: new Date(now + 36 * 3600_000).toISOString(),
    $select: "id,subject,start,end,isAllDay,isCancelled,responseStatus,onlineMeeting",
    $orderby: "start/dateTime",
    $top: "100",
  });
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/calendarView?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.timezone="UTC"' },
    });
    if (!res.ok) return false;
    const { value = [] } = await res.json();
    const events = value
      .filter((e) => !e.isAllDay && !e.isCancelled && e.responseStatus?.response !== "declined")
      .map((e) => ({
        id: e.id,
        subject: e.subject || "(no title)",
        start: graphTime(e.start.dateTime),
        end: graphTime(e.end.dateTime),
        joinUrl: e.onlineMeeting?.joinUrl || null,
      }));
    ipcRenderer.send("calendar-events", events);
    return true;
  } catch {
    return false;
  }
}

async function calendarLoop() {
  if (meetingWindow) return;
  const ok = await fetchCalendar();
  setTimeout(calendarLoop, ok ? CALENDAR_EVERY : CALENDAR_RETRY);
}

window.addEventListener("DOMContentLoaded", () => {
  if (TEAMS_HOSTS.test(location.hostname)) setTimeout(calendarLoop, 5000);
});
ipcRenderer.on("calendar-refresh", () => fetchCalendar());

// ---- Next meeting in the header, left of the Teams buttons in the top-right corner ----
const NEXT_MEETING_GAP = 8;
let nextMeeting = null; // { subject, start, end, color, account } (from all accounts)
let nextMeetingEl = null;

function relativeTime(start, end) {
  const now = Date.now();
  if (now >= start) return now < end ? "now" : null;
  const minutes = Math.ceil((start - now) / 60000);
  if (minutes < 60) return `in ${minutes} min`;
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return m ? `in ${h} h ${m} min` : `in ${h} h`;
}

const clock = (ms) => new Date(ms).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

function renderNextMeeting() {
  if (meetingWindow || !TEAMS_HOSTS.test(location.hostname) || !document.documentElement) return;
  if (!nextMeetingEl) {
    nextMeetingEl = el("div");
    nextMeetingEl.id = "teams-linux-next-meeting";
    nextMeetingEl.append(el("span", "tlm-dot"), el("span", "tlm-title"), el("span", "tlm-when"));
  }
  if (!nextMeetingEl.isConnected) document.documentElement.append(nextMeetingEl);

  const when = nextMeeting && relativeTime(nextMeeting.start, nextMeeting.end);
  nextMeetingEl.hidden = !when;
  if (!when) return;
  const { subject, start, end, color, account } = nextMeeting;
  const dot = nextMeetingEl.querySelector(".tlm-dot");
  dot.hidden = !account; // only with more than one account
  dot.style.background = color || "";
  nextMeetingEl.querySelector(".tlm-title").textContent = subject;
  nextMeetingEl.querySelector(".tlm-when").textContent = when;
  nextMeetingEl.title = `${clock(start)}–${clock(end)}${account ? ` · ${account}` : ""}`;
  const minutes = (start - Date.now()) / 60000;
  nextMeetingEl.classList.toggle("tlm-now", when === "now");
  nextMeetingEl.classList.toggle("tlm-soon", when !== "now" && minutes <= 10);

  const corner = document.querySelector('[data-tid="more-options-header"]');
  const right = corner ? window.innerWidth - corner.getBoundingClientRect().left + NEXT_MEETING_GAP : NEXT_MEETING_GAP;
  nextMeetingEl.style.right = `${Math.round(right)}px`;
  if (switcherState) nextMeetingEl.style.color = switcher?.style.color || "";
}

ipcRenderer.on("next-meeting", (_event, meeting) => {
  nextMeeting = meeting;
  renderNextMeeting();
});
window.addEventListener("DOMContentLoaded", () => setInterval(renderNextMeeting, 15_000));
window.addEventListener("resize", () => renderNextMeeting());

// Replaces window.Notification (permission always granted) with a version that goes through
// the main process, so a click opens the window on the right account.
contextBridge.executeInMainWorld({
  args: [CLICK_MESSAGE, KEEP_ACTIVE_MESSAGE, CALL_MESSAGE],
  func: (clickMessage, keepActiveMessage, callMessage) => {
    const bridge = window.__teamsLinux;
    const live = new Map();
    let seq = 0;

    class TeamsNotification extends EventTarget {
      constructor(title, options = {}) {
        super();
        this.title = String(title ?? "");
        this.body = String(options.body ?? "");
        this.tag = String(options.tag ?? "");
        this.icon = String(options.icon ?? "");
        this.data = options.data ?? null;
        this.silent = !!options.silent;
        this.onclick = this.onclose = this.onshow = this.onerror = null;
        this._id = ++seq;
        live.set(this._id, this);
        bridge.notify(this._id, this.title, this.body);
        setTimeout(() => this._fire("show"), 0);
      }
      static get permission() { return "granted"; }
      static requestPermission(callback) {
        if (typeof callback === "function") callback("granted");
        return Promise.resolve("granted");
      }
      static get maxActions() { return 0; }
      close() {
        if (live.delete(this._id)) this._fire("close");
      }
      _fire(type) {
        const event = new Event(type);
        const handler = this["on" + type];
        if (typeof handler === "function") handler.call(this, event);
        this.dispatchEvent(event);
      }
    }

    // "Keep active": the page always reports itself visible (a hidden window or a background
    // account doesn't count as inactivity). Focus stays real, because Teams uses it to decide
    // whether to show a notification.
    let keepActive = false;
    const hidden = Object.getOwnPropertyDescriptor(Document.prototype, "hidden");
    const visibilityState = Object.getOwnPropertyDescriptor(Document.prototype, "visibilityState");
    Object.defineProperty(Document.prototype, "hidden", {
      ...hidden,
      get() { return keepActive ? false : hidden.get.call(this); },
    });
    Object.defineProperty(Document.prototype, "visibilityState", {
      ...visibilityState,
      get() { return keepActive ? "visible" : visibilityState.get.call(this); },
    });
    window.addEventListener("visibilitychange", (e) => { if (keepActive) e.stopImmediatePropagation(); }, true);

    window.addEventListener("message", (e) => {
      if (e.source !== window) return;
      if (e.data?.type === keepActiveMessage) {
        keepActive = e.data.on;
        return;
      }
      if (e.data?.type !== clickMessage) return;
      const n = live.get(e.data.id);
      if (!n) return;
      window.focus();
      n._fire("click");
      n.close();
    });

    // Calls: Teams only opens WebRTC connections for calls and meetings. "Active" while any of
    // them is connected (not while it rings); over once they have all closed.
    const NativePeerConnection = window.RTCPeerConnection;
    if (NativePeerConnection) {
      const connections = new Set();
      let active = false;
      const report = () => {
        const now = [...connections].some((pc) => pc.connectionState === "connected");
        if (now === active) return;
        active = now;
        window.postMessage({ type: callMessage, active }, "*");
      };
      class PeerConnection extends NativePeerConnection {
        constructor(...args) {
          super(...args);
          connections.add(this);
          this.addEventListener("connectionstatechange", () => {
            if (this.connectionState === "closed" || this.connectionState === "failed") connections.delete(this);
            report();
          });
        }
        // close() fires no connectionstatechange.
        close() {
          super.close();
          connections.delete(this);
          report();
        }
      }
      window.RTCPeerConnection = PeerConnection;
      if (window.webkitRTCPeerConnection) window.webkitRTCPeerConnection = PeerConnection;
    }

    // Dead acute + c: on Wayland, without an input method running, Chromium composes accents
    // with its own built-in (en_US) table and gives ć, ignoring the system's ç (~/.XCompose,
    // GTK_IM_MODULE=cedilla). The composition can't be cancelled, so the character is swapped
    // right after it.
    const CEDILLA = { "ć": "ç", "Ć": "Ç" };

    // The Teams message box is a CKEditor 5: it keeps its own model of the text and ignores
    // DOM edits, so the swap goes through its model.
    function swapInEditor(editor, typed, replacement) {
      const { model } = editor;
      model.change((writer) => {
        const selection = model.document.selection;
        const end = selection.getFirstPosition();
        if (!selection.isCollapsed || !end || end.offset < 1) return;
        const range = writer.createRange(end.getShiftedBy(-1), end);
        const text = [...range.getItems()].map((item) => item.data ?? "").join("");
        if (text !== typed) return;
        model.insertContent(writer.createText(replacement, selection.getAttributes()), range);
      });
    }

    function swapInPage(field, typed, replacement) {
      if (field instanceof HTMLTextAreaElement || field instanceof HTMLInputElement) {
        const end = field.selectionEnd;
        if (field.selectionStart !== end || field.value.slice(end - 1, end) !== typed) return;
        field.setSelectionRange(end - 1, end);
      } else {
        const selection = getSelection();
        if (!selection.rangeCount || !selection.isCollapsed) return;
        selection.modify("extend", "backward", "character");
        if (selection.toString() !== typed) return selection.collapseToEnd();
      }
      document.execCommand("insertText", false, replacement);
    }

    document.addEventListener("compositionend", (e) => {
      const typed = e.data;
      const replacement = CEDILLA[typed];
      if (!replacement) return;
      setTimeout(() => {
        const field = document.activeElement;
        const editor = field?.closest?.(".ck-editor__editable")?.ckeditorInstance;
        if (editor) swapInEditor(editor, typed, replacement);
        else swapInPage(field, typed, replacement);
      }, 0);
    }, true);

    Object.defineProperty(window, "Notification", { value: TeamsNotification, writable: true, configurable: true });
    if (window.ServiceWorkerRegistration) {
      ServiceWorkerRegistration.prototype.showNotification = function (title, options) {
        new TeamsNotification(title, options);
        return Promise.resolve();
      };
    }
    return null;
  },
});
