// Preload das páginas do Teams (sandbox + contextIsolation).
// Expõe à página só o pedido de notificação; o clique volta por postMessage.
const { contextBridge, ipcRenderer } = require("electron");

const CLICK_MESSAGE = "teams-linux:notification-click";
const KEEP_ACTIVE_MESSAGE = "teams-linux:keep-active";

ipcRenderer.on("notification-click", (_event, id) => {
  window.postMessage({ type: CLICK_MESSAGE, id }, "*");
});
ipcRenderer.on("keep-active", (_event, on) => {
  window.postMessage({ type: KEEP_ACTIVE_MESSAGE, on: !!on }, "*");
});

contextBridge.exposeInMainWorld("__teamsLinux", {
  notify: (id, title, body) => ipcRenderer.send("notify", { id, title: String(title), body: String(body) }),
});

// Botão de contas fixo no canto superior esquerdo, no lugar do lançador de apps e do
// botão de recolher a barra do Teams (o CSS que os esconde vem do processo principal).
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
  root.title = "Trocar de conta";

  const avatar = el("span", "tls-avatar");
  avatar.append(el("span", "tls-initial"));
  // Sino com as não lidas das outras contas (a da conta aberta o Teams já mostra).
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
  switcher.title = others > 0 ? `Trocar de conta — ${others} não lida(s) em outras contas` : "Trocar de conta";
  // Mesma cor de texto dos botões do cabeçalho do Teams (acompanha o tema claro/escuro).
  const reference = document.querySelector('[data-tid="waffle-open-button"]');
  if (reference) switcher.style.color = getComputedStyle(reference).color;
}

function mountSwitcher() {
  if (!TEAMS_HOSTS.test(location.hostname) || !document.documentElement) return;
  switcher ??= buildSwitcher();
  // Fora do <body>, para o React do Teams não remover.
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

// Avisa o processo principal quando o Teams terminou de carregar (as contas são abertas
// uma a uma na inicialização). Fora do Teams (login) não há o que esperar.
// A barra de navegação só existe com o app montado; o marcador de "tudo renderizado" nem
// sempre aparece quando a conta carrega atrás de outra.
const TEAMS_READY_SELECTOR = '[data-tid="all-phased-rendering-complete"], [data-tid="app-layout-area--nav"]';
window.addEventListener("DOMContentLoaded", () => {
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

// ---- Agenda: próximas reuniões pelo Microsoft Graph ----
// Usa o token de agenda que o próprio Teams guarda (cache do MSAL no localStorage); o
// Teams o renova sozinho. Os eventos vão para o processo principal, que junta as contas,
// mostra a próxima reunião e dispara o aviso em tela cheia.
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

// O Graph devolve "2026-10-19T13:00:00.0000000" (UTC, com 7 casas decimais).
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
        subject: e.subject || "(sem título)",
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
  const ok = await fetchCalendar();
  setTimeout(calendarLoop, ok ? CALENDAR_EVERY : CALENDAR_RETRY);
}

window.addEventListener("DOMContentLoaded", () => {
  if (TEAMS_HOSTS.test(location.hostname)) setTimeout(calendarLoop, 5000);
});
ipcRenderer.on("calendar-refresh", () => fetchCalendar());

// ---- Próxima reunião no cabeçalho, à esquerda dos botões do canto direito do Teams ----
const NEXT_MEETING_GAP = 50;
let nextMeeting = null; // { subject, start, end, color, account } (de todas as contas)
let nextMeetingEl = null;

function relativeTime(start, end) {
  const now = Date.now();
  if (now >= start) return now < end ? "agora" : null;
  const minutes = Math.ceil((start - now) / 60000);
  if (minutes < 60) return `em ${minutes} min`;
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return m ? `em ${h} h ${m} min` : `em ${h} h`;
}

const clock = (ms) => new Date(ms).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

function renderNextMeeting() {
  if (!TEAMS_HOSTS.test(location.hostname) || !document.documentElement) return;
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
  dot.hidden = !account; // só com mais de uma conta
  dot.style.background = color || "";
  nextMeetingEl.querySelector(".tlm-title").textContent = subject;
  nextMeetingEl.querySelector(".tlm-when").textContent = when;
  nextMeetingEl.title = `${clock(start)}–${clock(end)}${account ? ` · ${account}` : ""}`;
  const minutes = (start - Date.now()) / 60000;
  nextMeetingEl.classList.toggle("tlm-now", when === "agora");
  nextMeetingEl.classList.toggle("tlm-soon", when !== "agora" && minutes <= 10);

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

// Substitui window.Notification (sempre com permissão concedida) por uma versão que
// passa pelo processo principal, para o clique abrir a janela na conta certa.
contextBridge.executeInMainWorld({
  args: [CLICK_MESSAGE, KEEP_ACTIVE_MESSAGE],
  func: (clickMessage, keepActiveMessage) => {
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

    // "Manter ativo": a página se diz sempre visível (janela escondida ou conta em segundo
    // plano não contam como inatividade). O foco continua real, porque é por ele que o
    // Teams decide mostrar notificação.
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
