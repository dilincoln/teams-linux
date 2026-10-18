import {
  app,
  BaseWindow,
  BrowserWindow,
  WebContentsView,
  Menu,
  Tray,
  Notification,
  nativeImage,
  session,
  shell,
  ipcMain,
  desktopCapturer,
} from "electron";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const ICON = path.join(here, "assets", "icon.png");

const KIND_URL = {
  work: "https://teams.microsoft.com/v2/",
  personal: "https://teams.live.com/v2/",
};

// Esconde o lançador de apps e o botão de recolher a barra do Teams; no lugar deles fica
// o botão de contas que o preload põe na página (e, à direita, a próxima reunião). A janela não tem moldura, então o
// cabeçalho do Teams vira a área de arrastar (menos o que é clicável).
const PAGE_CSS = `
[data-tid="waffle-open-button"], [data-tid="hamburger-button"] { visibility: hidden !important; }
[data-tid="title-bar"] { -webkit-app-region: drag; }
[data-tid="title-bar"] :is(button, a, input, textarea, select, [role="button"], [role="search"], [role="combobox"], [role="menuitem"], [tabindex]),
#teams-linux-switcher { -webkit-app-region: no-drag; }
#teams-linux-switcher {
  position: fixed; top: 4px; left: 4px; z-index: 2147483647;
  display: flex; align-items: center; gap: 8px; box-sizing: border-box;
  max-width: 160px; height: 40px; padding: 0 10px 0 7px; border-radius: 6px;
  font: 600 13px/1 "Segoe UI", system-ui, "Noto Sans", sans-serif;
  background: transparent; cursor: pointer; user-select: none;
}
#teams-linux-switcher:hover { background: rgba(128, 128, 128, 0.2); }
#teams-linux-switcher:focus-visible { outline: 2px solid #7f85f5; outline-offset: -2px; }
#teams-linux-switcher .tls-avatar {
  flex: none; display: grid; place-items: center;
  width: 26px; height: 26px; border-radius: 50%; color: #fff; font-size: 12px;
}
#teams-linux-switcher .tls-name { flex: 0 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#teams-linux-switcher .tls-bell { position: relative; flex: none; display: grid; place-items: center; opacity: 0.6; }
#teams-linux-switcher .tls-bell.tls-active { opacity: 1; }
#teams-linux-switcher .tls-badge {
  position: absolute; top: -7px; left: 9px; box-sizing: border-box;
  min-width: 16px; height: 16px; padding: 0 4px; border-radius: 8px;
  background: #d92c2c; color: #fff; font-size: 10px; font-weight: 700; line-height: 16px; text-align: center;
}
#teams-linux-switcher .tls-badge[hidden] { display: none; }
#teams-linux-next-meeting {
  position: fixed; top: 4px; right: 135px; z-index: 2147483647;
  display: flex; align-items: center; gap: 7px; box-sizing: border-box;
  max-width: min(420px, calc(50vw - 320px)); height: 40px; padding: 0 4px;
  font: 400 13px/1 "Segoe UI", system-ui, "Noto Sans", sans-serif;
  -webkit-app-region: no-drag; cursor: default; user-select: none;
}
#teams-linux-next-meeting[hidden], #teams-linux-next-meeting .tlm-dot[hidden] { display: none; }
#teams-linux-next-meeting .tlm-dot { flex: none; width: 8px; height: 8px; border-radius: 50%; }
#teams-linux-next-meeting .tlm-title { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 600; }
#teams-linux-next-meeting .tlm-when { flex: none; opacity: 0.7; }
#teams-linux-next-meeting.tlm-soon .tlm-when { color: #f7b23b; opacity: 1; font-weight: 600; }
#teams-linux-next-meeting.tlm-now .tlm-when { color: #ff6b6b; opacity: 1; font-weight: 600; }
`;

// Popups que podem abrir dentro do app (login etc.); o resto vai para o navegador.
const IN_APP_HOSTS = [
  "login.microsoftonline.com",
  "login.microsoft.com",
  "login.live.com",
  "account.live.com",
  "teams.microsoft.com",
  "teams.cloud.microsoft",
  "teams.live.com",
];

if (!app.requestSingleInstanceLock()) {
  app.quit();
  process.exit(0);
}

const startHidden = process.argv.includes("--minimized") || process.argv.includes("--hidden");

// ---------------------------------------------------------------------------
// Configuração
// ---------------------------------------------------------------------------

const configPath = () => path.join(app.getPath("userData"), "config.json");

const DEFAULT_SETTINGS = {
  reminderSeconds: 0, // repetir notificação de não lidas a cada N segundos (0 = desligado)
  keepActive: false, // impedir que o Teams fique "Ausente" por inatividade
  sound: true, // tocar som nas notificações (o KDE não toca para apps sem .notifyrc)
  soundName: "ocean/message-new-instant",
  meetingAlert: 5, // aviso em tela cheia N minutos antes da reunião (null = desligado)
  showNextMeeting: true, // próxima reunião no cabeçalho do Teams
};
const MEETING_ALERT_OPTIONS = [null, 0, 1, 2, 5, 10, 15, 30];

// Sons do sistema oferecidos nas configurações: "tema/nome" em /usr/share/sounds/<tema>/stereo.
const SOUNDS = [
  ["ocean/message-new-instant", "Mensagem (Ocean)"],
  ["ocean/message-contact-in", "Contato (Ocean)"],
  ["ocean/message-attention", "Atenção (Ocean)"],
  ["ocean/message-highlight", "Destaque (Ocean)"],
  ["ocean/message-new-email", "E-mail (Ocean)"],
  ["ocean/bell", "Sino (Ocean)"],
  ["freedesktop/message-new-instant", "Mensagem (Freedesktop)"],
  ["freedesktop/message", "Aviso (Freedesktop)"],
  ["freedesktop/bell", "Sino (Freedesktop)"],
  ["freedesktop/complete", "Concluído (Freedesktop)"],
]
  .map(([id, label]) => {
    const [theme, name] = id.split("/");
    return { id, label, file: `/usr/share/sounds/${theme}/stereo/${name}.oga` };
  })
  .filter((sound) => fs.existsSync(sound.file));

function loadConfig() {
  let saved = {};
  try {
    saved = JSON.parse(fs.readFileSync(configPath(), "utf8"));
  } catch {}
  return { accounts: [], active: null, bounds: null, ...saved, settings: { ...DEFAULT_SETTINGS, ...saved.settings } };
}

function saveConfig() {
  const file = configPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(`${file}.tmp`, JSON.stringify(config, null, 2));
  fs.renameSync(`${file}.tmp`, file);
}

const newId = () => Date.now().toString(16) + Math.random().toString(16).slice(2, 6);
const partitionOf = (id) => `persist:account-${id}`;

/** O Teams põe a contagem de não lidas no título: "(3) Chat | Microsoft Teams". */
const parseUnread = (title) => {
  const m = /^\s*\((\d+)\+?\)/.exec(title || "");
  return m ? Number(m[1]) : 0;
};

// ---------------------------------------------------------------------------
// Estado
// ---------------------------------------------------------------------------

let config;
let win;
let ui; // WebContentsView local: dropdown de contas + diálogos
let tray;
const views = new Map(); // id da conta -> WebContentsView
const unread = new Map();
let dialogOpen = false;
let menuOpen = false; // dropdown de contas aberto: a view local cobre a janela (transparente)
let quitting = false;
let hideAfterLoading = startHidden; // aberto com --minimized: mostra só o carregamento

// Conta escondida não renderiza e o Teams dela fica parado na tela de abertura (sem
// notificações). Por isso cada conta é "aquecida": fica visível, atrás da tela de
// carregamento ou da conta ativa, até o Teams terminar de carregar. Uma por vez.
const READY_TIMEOUT = 60_000;
let loading = false; // tela de carregamento da abertura
const ready = new Set(); // contas com o Teams carregado
const timedOut = new Set(); // contas que não terminaram dentro do limite
const readyWaiters = new Map(); // id -> [resolve]
const warmQueue = [];
let warming = null; // conta visível atrás da ativa enquanto carrega
let warmRunning = false;

const account = (id) => config.accounts.find((a) => a.id === id);
const activeAccount = () => account(config.active);

function snapshot() {
  return {
    accounts: config.accounts,
    active: config.active,
    unread: Object.fromEntries(unread),
    settings: config.settings,
    sounds: SOUNDS.map(({ id, label }) => ({ id, label })),
    loading: loading ? { ready: [...ready], timedOut: [...timedOut], current: warming } : null,
  };
}

function emitState() {
  ui?.webContents.send("state", snapshot());
  for (const id of views.keys()) sendSwitcherState(id);
}

/** Estado do botão de contas dentro da página: conta da view e não lidas das outras. */
function sendSwitcherState(id) {
  const acct = account(id);
  if (!acct) return;
  const others = config.accounts.filter((a) => a.id !== id).reduce((sum, a) => sum + (unread.get(a.id) ?? 0), 0);
  views.get(id)?.webContents.send("switcher-state", { name: acct.name, color: acct.color, others });
}

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

function layout() {
  if (!win) return;
  const { width, height } = win.getContentBounds();
  for (const view of views.values()) view.setBounds({ x: 0, y: 0, width, height });

  ui.setBounds({ x: 0, y: 0, width, height });
  // Fora do carregamento, do dropdown e dos diálogos a view local sai de cima da página.
  ui.setVisible(loading || dialogOpen || menuOpen || config.accounts.length === 0);
}

function applyVisibility() {
  const showing = dialogOpen ? null : config.active;
  for (const [id, view] of views) view.setVisible(id === showing || id === warming);
  // A conta aquecendo fica atrás da ativa, e a view local sempre por cima de todas.
  if (warming && showing && warming !== showing && views.has(showing)) {
    win.contentView.addChildView(views.get(showing));
    win.contentView.addChildView(ui);
  }
  if (showing && views.has(showing) && !loading) views.get(showing).webContents.focus();
  layout();
}

// ---------------------------------------------------------------------------
// Carregamento das contas
// ---------------------------------------------------------------------------

function markReady(id) {
  timedOut.delete(id);
  if (!ready.has(id)) {
    ready.add(id);
    emitState();
  }
  for (const resolve of readyWaiters.get(id) ?? []) resolve(true);
  readyWaiters.delete(id);
}

function waitReady(id, ms) {
  if (ready.has(id)) return Promise.resolve(true);
  return new Promise((resolve) => {
    const timer = setTimeout(() => resolve(false), ms);
    const waiters = readyWaiters.get(id) ?? [];
    waiters.push((ok) => {
      clearTimeout(timer);
      resolve(ok);
    });
    readyWaiters.set(id, waiters);
  });
}

function warm(id) {
  if (warming === id || warmQueue.includes(id)) return;
  warmQueue.push(id);
  if (!warmRunning) runWarmQueue();
}

async function runWarmQueue() {
  warmRunning = true;
  while (warmQueue.length) {
    const id = warmQueue.shift();
    if (!views.has(id) || ready.has(id)) continue;
    warming = id;
    applyVisibility();
    emitState();
    if (!(await waitReady(id, READY_TIMEOUT)) && views.has(id)) timedOut.add(id);
    warming = null;
  }
  warmRunning = false;
  if (loading) finishLoading();
  else applyVisibility();
  emitState();
}

/** Abertura: aquece as outras contas e por último a ativa (que já carrega junto, por baixo). */
function startLoading() {
  if (!config.accounts.length) return;
  loading = true;
  for (const acct of config.accounts) if (acct.id !== config.active) warm(acct.id);
  if (config.active) warm(config.active);
  applyVisibility();
  emitState();
}

function finishLoading() {
  if (!loading) return;
  loading = false;
  if (hideAfterLoading) win.hide();
  hideAfterLoading = false;
  applyVisibility();
  emitState();
}

function refreshIndicators() {
  const total = [...unread.values()].reduce((a, b) => a + b, 0);
  const name = activeAccount()?.name;
  const title = `${total > 0 ? `(${total}) ` : ""}Teams${name ? ` — ${name}` : ""}`;
  win?.setTitle(title);
  if (tray) {
    tray.setToolTip(total > 0 ? `Teams — ${total} não lida(s)` : "Teams");
    tray.setImage(total > 0 ? badgedIcon() : trayIcon());
  }
  app.setBadgeCount(total);
}

// ---------------------------------------------------------------------------
// Ícones da bandeja
// ---------------------------------------------------------------------------

let trayIconCache;
let badgedIconCache;

function trayIcon() {
  trayIconCache ??= nativeImage.createFromPath(ICON).resize({ width: 32, height: 32 });
  return trayIconCache;
}

/** Ícone com um ponto vermelho no canto superior direito. */
function badgedIcon() {
  if (badgedIconCache) return badgedIconCache;
  const base = trayIcon();
  const { width, height } = base.getSize();
  const bitmap = Buffer.from(base.toBitmap()); // BGRA
  const r = Math.min(width, height) * 0.24;
  const cx = width - r - 1;
  const cy = r + 1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (Math.hypot(x + 0.5 - cx, y + 0.5 - cy) <= r) {
        const i = (y * width + x) * 4;
        bitmap[i] = 0x2c; // B
        bitmap[i + 1] = 0x2c; // G
        bitmap[i + 2] = 0xd9; // R
        bitmap[i + 3] = 0xff;
      }
    }
  }
  badgedIconCache = nativeImage.createFromBitmap(bitmap, { width, height });
  return badgedIconCache;
}

// ---------------------------------------------------------------------------
// Contas
// ---------------------------------------------------------------------------

function chromeUserAgent(ua) {
  // Sem "Electron/x" e o nome do app, o Teams trata como Chrome comum.
  return ua.replace(/\s(Electron|teams-linux|Teams)\/\S+/g, "");
}

function setupSession(ses) {
  if (ses.__teamsConfigured) return;
  ses.__teamsConfigured = true;

  ses.setUserAgent(chromeUserAgent(ses.getUserAgent()));

  const allowed = new Set(["notifications", "media", "clipboard-read", "clipboard-sanitized-write", "display-capture", "fullscreen", "speaker-selection"]);
  ses.setPermissionRequestHandler((_wc, permission, callback) => callback(allowed.has(permission)));
  ses.setPermissionCheckHandler((_wc, permission) => allowed.has(permission));

  // Compartilhamento de tela: no Wayland o getSources abre o seletor do sistema (portal).
  ses.setDisplayMediaRequestHandler(
    (_request, callback) => {
      desktopCapturer
        .getSources({ types: ["screen", "window"] })
        .then((sources) => callback(sources.length ? { video: sources[0] } : {}))
        .catch(() => callback({}));
    },
    { useSystemPicker: true },
  );

  ses.on("will-download", (_event, item) => {
    item.setSavePath(uniquePath(path.join(app.getPath("downloads"), item.getFilename())));
    item.once("done", (_e, state) => {
      new Notification({
        title: state === "completed" ? "Download concluído" : "Falha no download",
        body: path.basename(item.getSavePath()),
        icon: ICON,
      }).show();
    });
  });
}

function uniquePath(file) {
  if (!fs.existsSync(file)) return file;
  const { dir, name, ext } = path.parse(file);
  for (let i = 1; ; i++) {
    const candidate = path.join(dir, `${name} (${i})${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
}

function opensInApp(url) {
  try {
    const { protocol, hostname } = new URL(url);
    if (protocol === "about:") return true;
    return IN_APP_HOSTS.some((h) => hostname === h || hostname.endsWith(`.${h}`));
  } catch {
    return false;
  }
}

function createAccountView(acct) {
  const ses = session.fromPartition(partitionOf(acct.id));
  setupSession(ses);

  const view = new WebContentsView({
    webPreferences: {
      session: ses,
      preload: path.join(here, "preload-account.cjs"),
      contextIsolation: true,
      sandbox: true,
      // As páginas da Microsoft pedem Segoe UI/Helvetica; sem elas o padrão é serifado.
      defaultFontFamily: { standard: "Noto Sans", sansSerif: "Noto Sans" },
      // Contas em segundo plano sem throttling, para notificações e presença em dia.
      backgroundThrottling: false,
    },
  });
  const wc = view.webContents;

  wc.on("page-title-updated", (_e, title) => setUnread(acct.id, parseUnread(title)));
  wc.on("dom-ready", () => {
    wc.insertCSS(PAGE_CSS).catch(() => {});
    sendSwitcherState(acct.id);
    wc.send("keep-active", config.settings.keepActive);
    wc.send("next-meeting", lastNextMeeting);
  });
  // Página nova (recarga, redirecionamento): o Teams dela precisa carregar de novo.
  wc.on("did-start-navigation", (details) => {
    if (!details.isMainFrame || details.isSameDocument) return;
    ready.delete(acct.id);
    if (acct.id !== config.active) warm(acct.id);
  });
  wc.setWindowOpenHandler(({ url }) => {
    if (opensInApp(url)) return { action: "allow" };
    shell.openExternal(url);
    return { action: "deny" };
  });
  wc.loadURL(KIND_URL[acct.kind] ?? KIND_URL.work);

  view.setVisible(false);
  win.contentView.addChildView(view);
  // O dropdown/diálogos precisam continuar por cima das contas.
  win.contentView.addChildView(ui);
  views.set(acct.id, view);
  return view;
}

function setUnread(id, count) {
  if ((unread.get(id) ?? 0) === count) return;
  unread.set(id, count);
  if (count === 0) lastNotification.delete(id);
  emitState();
  refreshIndicators();
}

function activate(id) {
  if (!account(id)) return;
  const previous = config.active;
  config.active = id;
  if (previous && previous !== id && views.has(previous) && !ready.has(previous)) warm(previous);
  saveConfig();
  applyVisibility();
  emitState();
  refreshIndicators();
}

function showWindow() {
  if (!win) return;
  hideAfterLoading = false;
  if (win.isMinimized()) win.restore();
  win.show();
  win.focus();
}

function focusAccount(id) {
  showWindow();
  activate(id);
}

function accountAction(action) {
  const view = config.active && views.get(config.active);
  if (action === "reload") view?.webContents.reload();
  else if (action === "devtools") view?.webContents.openDevTools({ mode: "detach" });
}

// ---------------------------------------------------------------------------
// IPC do dropdown/diálogos (só o webview local tem esse preload)
// ---------------------------------------------------------------------------

function fromUi(event) {
  return event.sender === ui?.webContents;
}

ipcMain.handle("get-state", (e) => (fromUi(e) ? snapshot() : null));

// Clique no botão de contas dentro da página: abre o dropdown ancorado à esquerda.
ipcMain.on("open-accounts-menu", (e) => {
  if (dialogOpen || ![...views.values()].some((v) => v.webContents === e.sender)) return;
  ui.webContents.send("open-menu");
});

ipcMain.on("dialog-open", (e, open) => {
  if (!fromUi(e)) return;
  dialogOpen = open;
  applyVisibility();
  if (open) ui.webContents.focus();
});

// O preload avisa quando o Teams da conta terminou de carregar.
ipcMain.on("account-ready", (e) => {
  const entry = [...views].find(([, v]) => v.webContents === e.sender);
  if (entry) markReady(entry[0]);
});

ipcMain.on("skip-loading", (e) => {
  // As contas que faltam continuam carregando atrás da ativa.
  if (fromUi(e)) finishLoading();
});

ipcMain.on("menu-open", (e, open) => {
  if (!fromUi(e)) return;
  menuOpen = !!open;
  layout();
  if (menuOpen) ui.webContents.focus();
});

ipcMain.on("activate", (e, id) => {
  if (fromUi(e)) activate(id);
});

ipcMain.on("account-action", (e, action) => {
  if (fromUi(e)) accountAction(action);
});

ipcMain.handle("add-account", (e, { name, color, kind }) => {
  if (!fromUi(e)) return;
  const acct = { id: newId(), name: String(name).trim(), color, kind: kind === "personal" ? "personal" : "work" };
  config.accounts.push(acct);
  saveConfig();
  createAccountView(acct);
  dialogOpen = false;
  activate(acct.id);
  return acct;
});

ipcMain.handle("update-account", (e, { id, name, color }) => {
  if (!fromUi(e)) return;
  const acct = account(id);
  if (!acct) return;
  acct.name = String(name).trim();
  acct.color = color;
  saveConfig();
  emitState();
  refreshIndicators();
});

ipcMain.handle("remove-account", async (e, id) => {
  if (!fromUi(e)) return;
  const view = views.get(id);
  if (view) {
    win.contentView.removeChildView(view);
    view.webContents.close();
    views.delete(id);
  }
  config.accounts = config.accounts.filter((a) => a.id !== id);
  unread.delete(id);
  ready.delete(id);
  timedOut.delete(id);
  calendars.delete(id);
  refreshMeetings();
  for (const resolve of readyWaiters.get(id) ?? []) resolve(false);
  readyWaiters.delete(id);
  if (config.active === id) config.active = config.accounts[0]?.id ?? null;
  saveConfig();
  // Apaga cookies, cache e armazenamento da conta.
  const ses = session.fromPartition(partitionOf(id));
  await ses.clearStorageData().catch(() => {});
  await ses.clearCache().catch(() => {});
  applyVisibility();
  emitState();
  refreshIndicators();
});

// Notificações das contas (o preload substitui window.Notification da página).
const liveNotifications = new Set(); // sem referência, o clique pode se perder no GC
const lastNotification = new Map(); // id da conta -> { title, body } (para os lembretes)

let lastSoundAt = 0;

/** Toca um som do sistema; várias notificações juntas soam uma vez só. */
function playSound(id = config.settings.soundName, { force = false } = {}) {
  const sound = SOUNDS.find((s) => s.id === id) ?? SOUNDS[0];
  if (!sound || (!force && Date.now() - lastSoundAt < 1000)) return;
  lastSoundAt = Date.now();
  const play = (cmd, fallback) =>
    spawn(cmd, [sound.file], { stdio: "ignore" })
      .on("error", () => fallback?.())
      .unref();
  play("pw-play", () => play("paplay"));
}

function showNotification(accountId, title, body, onClick) {
  const name = account(accountId)?.name;
  const notification = new Notification({
    title: config.accounts.length > 1 && name ? `${title} · ${name}` : title,
    body,
    icon: ICON,
  });
  liveNotifications.add(notification);
  notification.on("click", () => {
    focusAccount(accountId);
    onClick?.();
  });
  notification.on("close", () => liveNotifications.delete(notification));
  notification.show();
  if (config.settings.sound) playSound();
}

ipcMain.on("notify", (e, { id: notificationId, title, body }) => {
  const entry = [...views].find(([, v]) => v.webContents === e.sender);
  if (!entry) return;
  const [accountId] = entry;
  lastNotification.set(accountId, { title, body });
  showNotification(accountId, title, body, () => e.sender.send("notification-click", notificationId));
});

// ---------------------------------------------------------------------------
// Configurações
// ---------------------------------------------------------------------------

const KEEP_ACTIVE_EVERY = 60_000;
let reminderTimer = null;
let keepActiveTimer = null;

/** Enquanto houver não lidas, repete a notificação (menos da conta que está na tela). */
function remindUnread() {
  const looking = win?.isVisible() && win.isFocused();
  for (const acct of config.accounts) {
    const count = unread.get(acct.id) ?? 0;
    if (count === 0 || (looking && acct.id === config.active)) continue;
    const last = lastNotification.get(acct.id);
    const title = count === 1 ? "1 mensagem não lida" : `${count} mensagens não lidas`;
    showNotification(acct.id, title, last ? `${last.title}: ${last.body}` : "Você tem mensagens não lidas no Teams");
  }
}

/**
 * Movimento de mouse real (do navegador, não sintético) no cabeçalho de cada conta: o
 * Teams conta como atividade e não muda o status para Ausente.
 */
function nudgeAccounts() {
  for (const view of views.values()) {
    for (const x of [250, 251]) view.webContents.sendInputEvent({ type: "mouseMove", x, y: 24 });
  }
}

function applySettings() {
  const { reminderSeconds, keepActive } = config.settings;
  clearInterval(reminderTimer);
  reminderTimer = reminderSeconds > 0 ? setInterval(remindUnread, reminderSeconds * 1000) : null;
  clearInterval(keepActiveTimer);
  keepActiveTimer = keepActive ? setInterval(nudgeAccounts, KEEP_ACTIVE_EVERY) : null;
  for (const view of views.values()) view.webContents.send("keep-active", keepActive);
}

ipcMain.on("preview-sound", (e, id) => {
  if (fromUi(e)) playSound(id, { force: true });
});

ipcMain.handle("update-settings", (e, settings) => {
  if (!fromUi(e)) return;
  const seconds = Math.round(Number(settings?.reminderSeconds));
  config.settings = {
    reminderSeconds: Number.isFinite(seconds) && seconds > 0 ? Math.max(seconds, 10) : 0,
    keepActive: !!settings?.keepActive,
    sound: !!settings?.sound,
    soundName: SOUNDS.some((s) => s.id === settings?.soundName) ? settings.soundName : DEFAULT_SETTINGS.soundName,
    meetingAlert: MEETING_ALERT_OPTIONS.includes(settings?.meetingAlert) ? settings.meetingAlert : DEFAULT_SETTINGS.meetingAlert,
    showNextMeeting: !!settings?.showNextMeeting,
  };
  saveConfig();
  applySettings();
  refreshMeetings(true);
  emitState();
  return config.settings;
});

// ---------------------------------------------------------------------------
// Reuniões: próxima no cabeçalho e aviso em tela cheia
// ---------------------------------------------------------------------------

const calendars = new Map(); // id da conta -> eventos do Graph (vindos do preload)
const alerted = new Set(); // reuniões já avisadas ("conta:id:início")
const snoozed = new Map(); // reunião -> horário para avisar de novo
let lastNextMeeting = null;
let alertWin = null;
let currentAlert = null;

const meetingKey = (m) => `${m.accountId}:${m.id}:${m.start}`;

ipcMain.on("calendar-events", (e, events) => {
  const entry = [...views].find(([, v]) => v.webContents === e.sender);
  if (!entry || !Array.isArray(events)) return;
  const clean = events
    .filter((ev) => Number.isFinite(ev?.start) && Number.isFinite(ev?.end))
    .map((ev) => ({
      id: String(ev.id),
      subject: String(ev.subject).slice(0, 300),
      start: ev.start,
      end: ev.end,
      joinUrl: typeof ev.joinUrl === "string" && /^https:\/\//.test(ev.joinUrl) ? ev.joinUrl : null,
    }));
  calendars.set(entry[0], clean);
  refreshMeetings();
});

/** Reuniões de todas as contas que ainda não terminaram, da mais próxima para a mais distante. */
function upcomingMeetings() {
  const now = Date.now();
  const all = [];
  for (const [accountId, events] of calendars) {
    if (!account(accountId)) continue;
    for (const ev of events) if (ev.end > now) all.push({ ...ev, accountId });
  }
  return all.sort((a, b) => a.start - b.start);
}

function describe(meeting) {
  const acct = account(meeting.accountId);
  const multiple = config.accounts.length > 1;
  return {
    subject: meeting.subject,
    start: meeting.start,
    end: meeting.end,
    color: multiple ? acct?.color : null,
    account: multiple ? acct?.name : null,
  };
}

function refreshMeetings(force = false) {
  const now = Date.now();
  // Já começou há mais de 10 min: o cabeçalho passa para a seguinte.
  const next = upcomingMeetings().find((m) => m.start > now - 10 * 60_000);
  const shown = config.settings.showNextMeeting && next ? describe(next) : null;
  if (force || JSON.stringify(shown) !== JSON.stringify(lastNextMeeting)) {
    lastNextMeeting = shown;
    for (const view of views.values()) view.webContents.send("next-meeting", shown);
  }
  checkMeetingAlerts();
}

function checkMeetingAlerts() {
  const lead = config.settings.meetingAlert;
  if (lead === null) return;
  const now = Date.now();
  for (const meeting of upcomingMeetings()) {
    const key = meetingKey(meeting);
    const due = snoozed.get(key) ?? (alerted.has(key) ? Infinity : meeting.start - lead * 60_000);
    // Reunião que já começou há mais de 2 min (ex.: ao abrir o app) não gera aviso.
    if (now < due || now > meeting.start + 2 * 60_000) continue;
    alerted.add(key);
    snoozed.delete(key);
    showMeetingAlert(meeting);
    return;
  }
}

function showMeetingAlert(meeting) {
  currentAlert = meeting;
  const data = { ...describe(meeting), canJoin: !!meeting.joinUrl, canSnooze: Date.now() < meeting.start - 30_000 };
  if (config.settings.sound) playSound(undefined, { force: true });
  if (alertWin) {
    alertWin.webContents.send("alert", data);
    alertWin.show();
    alertWin.focus();
    return;
  }
  alertWin = new BrowserWindow({
    fullscreen: true,
    frame: false,
    alwaysOnTop: true,
    skipTaskbar: true,
    show: false,
    icon: ICON,
    title: "Reunião — Teams",
    backgroundColor: "#0d0d12",
    webPreferences: {
      preload: path.join(here, "preload-alert.cjs"),
      contextIsolation: true,
      sandbox: true,
    },
  });
  alertWin.setMenu(null);
  alertWin.loadFile(path.join(here, "ui", "alert.html"));
  alertWin.once("ready-to-show", () => {
    alertWin.webContents.send("alert", data);
    alertWin.show();
    alertWin.focus();
  });
  alertWin.on("closed", () => {
    alertWin = null;
    currentAlert = null;
  });
}

ipcMain.on("alert-action", (e, action) => {
  if (!alertWin || e.sender !== alertWin.webContents) return;
  const meeting = currentAlert;
  alertWin.close();
  if (!meeting) return;
  if (action === "join") {
    focusAccount(meeting.accountId);
    if (meeting.joinUrl) views.get(meeting.accountId)?.webContents.loadURL(meeting.joinUrl);
  } else if (action === "snooze" && Date.now() < meeting.start) {
    snoozed.set(meetingKey(meeting), meeting.start);
  }
});

// Pelas configurações: avisa a próxima reunião agora, para ver como fica.
ipcMain.on("test-meeting-alert", (e) => {
  if (!fromUi(e)) return;
  const now = Date.now();
  const next = upcomingMeetings().find((m) => m.start > now);
  showMeetingAlert(
    next ?? { id: "teste", accountId: config.active, subject: "Reunião de exemplo", start: now + 5 * 60_000, end: now + 35 * 60_000, joinUrl: null },
  );
});

setInterval(refreshMeetings, 15_000);

// ---------------------------------------------------------------------------
// Janela, bandeja e ciclo de vida
// ---------------------------------------------------------------------------

function createWindow() {
  const bounds = config.bounds ?? { width: 1280, height: 820 };
  win = new BaseWindow({
    ...bounds,
    minWidth: 720,
    minHeight: 480,
    frame: false,
    show: false,
    icon: ICON,
    title: "Teams",
    backgroundColor: "#1f1f1f",
  });
  if (config.maximized) win.maximize();

  ui = new WebContentsView({
    webPreferences: {
      preload: path.join(here, "preload-ui.cjs"),
      contextIsolation: true,
      sandbox: true,
    },
  });
  ui.setBackgroundColor("#00000000");
  ui.webContents.loadFile(path.join(here, "ui", "index.html"));
  win.contentView.addChildView(ui);

  for (const acct of config.accounts) createAccountView(acct);

  const persistBounds = () => {
    config.maximized = win.isMaximized();
    if (!config.maximized && !win.isMinimized()) config.bounds = win.getBounds();
    saveConfig();
  };
  win.on("resize", layout);
  win.on("resized", persistBounds);
  win.on("moved", persistBounds);
  win.on("maximize", () => { layout(); persistBounds(); });
  win.on("unmaximize", () => { layout(); persistBounds(); });
  win.on("close", (event) => {
    // Fechar só esconde: as contas continuam rodando para notificar.
    if (!quitting && tray) {
      event.preventDefault();
      win.hide();
    }
  });

  applyVisibility();
  refreshIndicators();
  startLoading();
  // Mesmo com --minimized a janela aparece durante o carregamento: conta escondida não carrega.
  if (!startHidden || loading) win.show();
}

function createTray() {
  tray = new Tray(trayIcon());
  tray.setToolTip("Teams");
  tray.setContextMenu(
    Menu.buildFromTemplate([
      { label: "Mostrar Teams", click: showWindow },
      { type: "separator" },
      { label: "Sair", click: () => { quitting = true; app.quit(); } },
    ]),
  );
  tray.on("click", () => (win.isVisible() && win.isFocused() ? win.hide() : showWindow()));
}

app.on("second-instance", showWindow);
app.on("before-quit", () => { quitting = true; });
app.on("window-all-closed", () => app.quit());

app.whenReady().then(() => {
  app.userAgentFallback = chromeUserAgent(app.userAgentFallback);
  config = loadConfig();
  if (!account(config.active)) config.active = config.accounts[0]?.id ?? null;
  createWindow();
  applySettings();
  try {
    createTray();
  } catch (err) {
    console.error("bandeja indisponível:", err);
    tray = null;
  }
});
