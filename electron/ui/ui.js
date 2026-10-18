const COLORS = ["#5b5fc7", "#0f6cbd", "#038387", "#107c10", "#ca5010", "#c239b3", "#d13438", "#69797e"];

const $ = (sel) => document.querySelector(sel);
const overlay = $("#overlay");
const form = $("#account-form");
const removeCard = $("#remove-card");
const settingsCard = $("#settings-card");

let state = { accounts: [], active: null, unread: {}, settings: {}, sounds: [] };
// { mode: "add" | "edit" | "remove" | "settings", id?: string }
let dialog = null;
let selectedColor = COLORS[0];

// ---- Dropdown de contas (aberto pelo botão no canto da página do Teams) ----

const menuBackdrop = $("#menu-backdrop");
const menu = $("#menu");
let menuOpen = false;

function menuItem({ label, onSelect, disabled = false, danger = false, check = false, color, count = 0 }) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = danger ? "menu-item danger" : "menu-item";
  b.setAttribute("role", "menuitem");
  b.disabled = disabled;

  if (color) {
    const mark = document.createElement("span");
    mark.className = "check";
    mark.textContent = check ? "✓" : "";
    const dot = document.createElement("span");
    dot.className = "dot";
    dot.style.background = color;
    b.append(mark, dot);
  }
  const text = document.createElement("span");
  text.className = "label";
  text.textContent = label;
  b.append(text);
  if (count > 0) {
    const badge = document.createElement("span");
    badge.className = "badge";
    badge.textContent = count > 99 ? "99+" : String(count);
    b.append(badge);
  }

  b.addEventListener("click", () => {
    closeMenu();
    onSelect();
  });
  return b;
}

function separator() {
  const hr = document.createElement("div");
  hr.className = "menu-sep";
  return hr;
}

function renderMenu() {
  const active = state.active;
  const items = state.accounts.map((a) =>
    menuItem({
      label: a.name,
      color: a.color,
      check: a.id === active,
      count: state.unread[a.id] || 0,
      onSelect: () => window.teams.activate(a.id),
    }),
  );
  if (items.length) items.push(separator());
  items.push(
    menuItem({ label: "Adicionar conta…", onSelect: () => openDialog({ mode: "add" }) }),
    menuItem({ label: "Editar conta…", disabled: !active, onSelect: () => openDialog({ mode: "edit", id: active }) }),
    menuItem({ label: "Recarregar", disabled: !active, onSelect: () => window.teams.accountAction("reload") }),
    menuItem({ label: "Ferramentas de desenvolvedor", disabled: !active, onSelect: () => window.teams.accountAction("devtools") }),
    separator(),
    menuItem({ label: "Configurações…", onSelect: () => openDialog({ mode: "settings" }) }),
    separator(),
    menuItem({ label: "Remover conta…", danger: true, disabled: !active, onSelect: () => openDialog({ mode: "remove", id: active }) }),
  );
  menu.replaceChildren(...items);
}

function openMenu() {
  menuOpen = true;
  renderMenu();
  menuBackdrop.hidden = false;
  window.teams.setMenuOpen(true);
  menu.querySelector(".menu-item:not([disabled])")?.focus();
}

function closeMenu() {
  if (!menuOpen) return;
  menuOpen = false;
  menuBackdrop.hidden = true;
  window.teams.setMenuOpen(false);
}

// O fundo cobre a janela inteira (inclusive o botão de contas): clicar nele fecha.
menuBackdrop.addEventListener("mousedown", (e) => {
  if (e.target === menuBackdrop) closeMenu();
});

menu.addEventListener("keydown", (e) => {
  if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
  e.preventDefault();
  const items = [...menu.querySelectorAll(".menu-item:not([disabled])")];
  const i = items.indexOf(document.activeElement);
  const next = e.key === "ArrowDown" ? (i + 1) % items.length : (i - 1 + items.length) % items.length;
  items[next]?.focus();
});

// ---- Diálogos ----

function renderColors() {
  $("#colors").replaceChildren(
    ...COLORS.map((color) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "swatch";
      b.style.background = color;
      b.setAttribute("aria-label", color);
      b.setAttribute("aria-pressed", String(color === selectedColor));
      b.addEventListener("click", () => {
        selectedColor = color;
        renderColors();
      });
      return b;
    }),
  );
}

function openDialog(next) {
  dialog = next;
  const account = state.accounts.find((a) => a.id === next.id);
  const firstRun = state.accounts.length === 0;

  form.hidden = next.mode === "remove" || next.mode === "settings";
  removeCard.hidden = next.mode !== "remove";
  settingsCard.hidden = next.mode !== "settings";

  if (next.mode === "settings") {
    $("#reminder-seconds").value = String(state.settings.reminderSeconds);
    $("#keep-active").checked = state.settings.keepActive;
    $("#sound").checked = state.settings.sound;
    $("#sound-name").replaceChildren(
      ...state.sounds.map(({ id, label }) => {
        const option = document.createElement("option");
        option.value = id;
        option.textContent = label;
        return option;
      }),
    );
    $("#sound-name").value = state.settings.soundName;
    $("#meeting-alert").value = state.settings.meetingAlert === null ? "off" : String(state.settings.meetingAlert);
    $("#show-next-meeting").checked = state.settings.showNextMeeting;
    syncSoundControls();
  } else if (next.mode === "remove") {
    $("#remove-name").textContent = account?.name ?? "";
  } else {
    const editing = next.mode === "edit";
    $("#form-title").textContent = editing ? "Editar conta" : firstRun ? "Bem-vindo ao Teams" : "Adicionar conta";
    $("#form-intro").hidden = editing;
    $("#kind-field").hidden = editing;
    $("#submit").textContent = editing ? "Salvar" : "Adicionar";
    $("#cancel").hidden = firstRun;
    $("#name").value = editing ? account?.name ?? "" : "";
    form.elements.kind.value = "work";
    selectedColor = editing ? account?.color ?? COLORS[0] : COLORS[state.accounts.length % COLORS.length];
    renderColors();
  }

  overlay.hidden = false;
  window.teams.setDialogOpen(true);
  const focusTarget = { remove: "#remove-cancel", settings: "#reminder-seconds" }[next.mode] ?? "#name";
  setTimeout(() => $(focusTarget).focus(), 0);
}

function closeDialog() {
  if (!dialog || state.accounts.length === 0) return;
  dialog = null;
  overlay.hidden = true;
  window.teams.setDialogOpen(false);
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const name = $("#name").value.trim();
  if (!name) return;
  if (dialog?.mode === "edit") {
    await window.teams.updateAccount({ id: dialog.id, name, color: selectedColor });
    closeDialog();
  } else {
    // add-account já fecha o diálogo no processo principal e ativa a conta nova.
    await window.teams.addAccount({ name, color: selectedColor, kind: form.elements.kind.value });
    dialog = null;
    overlay.hidden = true;
  }
});

function syncSoundControls() {
  const off = !$("#sound").checked;
  $("#sound-name").disabled = off;
  $("#sound-preview").disabled = off;
}

$("#sound").addEventListener("change", syncSoundControls);
$("#sound-preview").addEventListener("click", () => window.teams.previewSound($("#sound-name").value));
$("#meeting-alert-test").addEventListener("click", () => window.teams.testMeetingAlert());

settingsCard.addEventListener("submit", async (e) => {
  e.preventDefault();
  await window.teams.updateSettings({
    reminderSeconds: Number($("#reminder-seconds").value) || 0,
    keepActive: $("#keep-active").checked,
    sound: $("#sound").checked,
    soundName: $("#sound-name").value,
    meetingAlert: $("#meeting-alert").value === "off" ? null : Number($("#meeting-alert").value),
    showNextMeeting: $("#show-next-meeting").checked,
  });
  closeDialog();
});

$("#cancel").addEventListener("click", closeDialog);
$("#settings-cancel").addEventListener("click", closeDialog);
$("#remove-cancel").addEventListener("click", closeDialog);
$("#remove-confirm").addEventListener("click", async () => {
  const id = dialog?.id;
  dialog = null;
  overlay.hidden = true;
  window.teams.setDialogOpen(false);
  if (id) await window.teams.removeAccount(id);
});

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (menuOpen) closeMenu();
  else closeDialog();
});

// ---- Carregamento das contas na abertura ----

function renderLoading() {
  const { ready, timedOut, current } = state.loading;
  $("#loading-list").replaceChildren(
    ...state.accounts.map((a) => {
      const li = document.createElement("li");
      const dot = document.createElement("span");
      dot.className = "dot";
      dot.style.background = a.color;
      const name = document.createElement("span");
      name.className = "name";
      name.textContent = a.name;
      const status = document.createElement("span");
      status.className = "status";
      if (ready.includes(a.id)) {
        status.classList.add("done");
        status.textContent = "✓ Pronta";
      } else if (timedOut.includes(a.id)) {
        status.classList.add("failed");
        status.textContent = "Não respondeu";
      } else if (a.id === current || a.id === state.active) {
        // A ativa carrega desde o início, por baixo desta tela.
        const spinner = document.createElement("span");
        spinner.className = "spinner";
        status.append(spinner, "Carregando…");
      } else {
        status.textContent = "Na fila";
      }
      li.append(dot, name, status);
      return li;
    }),
  );
}

$("#loading-skip").addEventListener("click", () => window.teams.skipLoading());

// ---- Estado ----

function render() {
  $("#loading").hidden = !state.loading;
  if (state.loading) renderLoading();
  if (menuOpen) renderMenu();
  // Sem contas não há o que mostrar: abre direto o cadastro.
  if (state.accounts.length === 0 && !dialog) openDialog({ mode: "add" });
}

window.teams.onState((s) => {
  state = s;
  render();
});
window.teams.onOpenMenu(() => {
  if (dialog) return;
  if (menuOpen) closeMenu();
  else openMenu();
});
window.teams.getState().then((s) => {
  state = s;
  render();
});
