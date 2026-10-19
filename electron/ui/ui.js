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

// ---- Account dropdown (opened by the button in the corner of the Teams page) ----

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
    menuItem({ label: "Add account…", onSelect: () => openDialog({ mode: "add" }) }),
    menuItem({ label: "Edit account…", disabled: !active, onSelect: () => openDialog({ mode: "edit", id: active }) }),
    menuItem({ label: "Reload", disabled: !active, onSelect: () => window.teams.accountAction("reload") }),
    menuItem({ label: "Developer tools", disabled: !active, onSelect: () => window.teams.accountAction("devtools") }),
    separator(),
    menuItem({ label: "Settings…", onSelect: () => openDialog({ mode: "settings" }) }),
    separator(),
    menuItem({ label: "Remove account…", danger: true, disabled: !active, onSelect: () => openDialog({ mode: "remove", id: active }) }),
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

// The backdrop covers the whole window (account button included): clicking it closes.
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

// ---- Dialogs ----

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
    setPressed("#keep-active-days", "day", state.settings.keepActiveDays);
    ranges = state.settings.keepActiveRanges.map((r) => ({ ...r }));
    renderRanges();
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
    setPressed("#meeting-alerts", "minutes", state.settings.meetingAlerts);
    $("#show-next-meeting").checked = state.settings.showNextMeeting;
    syncSoundControls();
    syncKeepActiveControls();
    selectTab("notifications");
  } else if (next.mode === "remove") {
    $("#remove-name").textContent = account?.name ?? "";
  } else {
    const editing = next.mode === "edit";
    $("#form-title").textContent = editing ? "Edit account" : firstRun ? "Welcome to Teams" : "Add account";
    $("#form-intro").hidden = editing;
    $("#kind-field").hidden = editing;
    $("#submit").textContent = editing ? "Save" : "Add";
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
    // add-account already closes the dialog in the main process and activates the new account.
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

// ---- Settings ----

const TABS = ["notifications", "misc"];

function selectTab(name) {
  for (const tab of TABS) {
    const selected = tab === name;
    $(`#tab-${tab}`).setAttribute("aria-selected", String(selected));
    $(`#tab-${tab}`).tabIndex = selected ? 0 : -1;
    $(`#panel-${tab}`).hidden = !selected;
  }
}

for (const tab of TABS) $(`#tab-${tab}`).addEventListener("click", () => selectTab(tab));

$(".tabs").addEventListener("keydown", (e) => {
  if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
  const i = TABS.findIndex((tab) => $(`#tab-${tab}`).getAttribute("aria-selected") === "true");
  const next = TABS[(i + (e.key === "ArrowRight" ? 1 : TABS.length - 1)) % TABS.length];
  selectTab(next);
  $(`#tab-${next}`).focus();
});

/** Toggle buttons ("chips"): the pressed ones are those whose data-<key> is in `values`. */
function setPressed(group, key, values) {
  for (const chip of $(group).querySelectorAll(".chip")) {
    chip.setAttribute("aria-pressed", String(values.includes(Number(chip.dataset[key]))));
  }
}

const pressed = (group, key) =>
  [...$(group).querySelectorAll('.chip[aria-pressed="true"]')].map((chip) => Number(chip.dataset[key]));

for (const group of ["#meeting-alerts", "#keep-active-days"]) {
  $(group).addEventListener("click", (e) => {
    const chip = e.target.closest(".chip");
    if (!chip || chip.disabled) return;
    chip.setAttribute("aria-pressed", String(chip.getAttribute("aria-pressed") !== "true"));
  });
}

// "Keep active" time ranges being edited: [{ from: "08:00", to: "12:00" }, …]
let ranges = [];

function timeInput(value, label, onChange) {
  const input = document.createElement("input");
  input.type = "time";
  input.required = true;
  input.value = value;
  input.setAttribute("aria-label", label);
  input.addEventListener("change", () => onChange(input.value));
  return input;
}

function renderRanges() {
  const disabled = !$("#keep-active").checked;
  $("#keep-active-ranges").replaceChildren(
    ...ranges.map((range, i) => {
      const li = document.createElement("li");
      const dash = document.createElement("span");
      dash.textContent = "–";
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "btn icon";
      remove.textContent = "✕";
      remove.setAttribute("aria-label", "Remove range");
      remove.addEventListener("click", () => {
        ranges.splice(i, 1);
        renderRanges();
      });
      const from = timeInput(range.from, "From", (v) => (range.from = v));
      const to = timeInput(range.to, "To", (v) => (range.to = v));
      for (const el of [from, to, remove]) el.disabled = disabled;
      li.append(from, dash, to, remove);
      return li;
    }),
  );
  $("#ranges-hint").textContent = ranges.length
    ? "A range that ends before it starts goes past midnight."
    : "All day. Add ranges to limit it, e.g. 08:00–12:00 and 13:00–18:00.";
}

$("#add-range").addEventListener("click", () => {
  // Continues from the last range: 08:00–12:00, then 13:00–18:00 by default.
  const last = ranges.at(-1);
  ranges.push(last ? { from: shiftTime(last.to, 60), to: shiftTime(last.to, 6 * 60) } : { from: "08:00", to: "12:00" });
  renderRanges();
  $("#keep-active-ranges li:last-child input")?.focus();
});

function shiftTime(time, minutes) {
  const total = Math.min(Number(time.slice(0, 2)) * 60 + Number(time.slice(3)) + minutes, 23 * 60 + 59);
  return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
}

function syncKeepActiveControls() {
  const off = !$("#keep-active").checked;
  $("#keep-active-schedule").classList.toggle("disabled", off);
  for (const el of $("#keep-active-schedule").querySelectorAll("button, input")) el.disabled = off;
}

$("#keep-active").addEventListener("change", syncKeepActiveControls);
$("#sound").addEventListener("change", syncSoundControls);
$("#sound-preview").addEventListener("click", () => window.teams.previewSound($("#sound-name").value));
$("#meeting-alert-test").addEventListener("click", () => window.teams.testMeetingAlert());

settingsCard.addEventListener("submit", async (e) => {
  e.preventDefault();
  await window.teams.updateSettings({
    reminderSeconds: Number($("#reminder-seconds").value) || 0,
    keepActive: $("#keep-active").checked,
    keepActiveDays: pressed("#keep-active-days", "day"),
    keepActiveRanges: ranges.filter((r) => r.from && r.to),
    sound: $("#sound").checked,
    soundName: $("#sound-name").value,
    meetingAlerts: pressed("#meeting-alerts", "minutes"),
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

// Shift+F11 three times in a row, with Settings open, shows or hides the "keep active" section.
const REVEAL_PRESSES = 3;
const REVEAL_WINDOW = 1500;
let revealPresses = [];

document.addEventListener("keydown", (e) => {
  if (e.code !== "F11" || !e.shiftKey || e.repeat || dialog?.mode !== "settings") return;
  e.preventDefault();
  const now = Date.now();
  revealPresses = [...revealPresses.filter((t) => now - t < REVEAL_WINDOW), now];
  if (revealPresses.length < REVEAL_PRESSES) return;
  revealPresses = [];
  const section = $("#keep-active-section");
  section.hidden = !section.hidden;
  if (!section.hidden) selectTab("misc");
});

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (menuOpen) closeMenu();
  else closeDialog();
});

// ---- Account loading at startup ----

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
        status.textContent = "✓ Ready";
      } else if (timedOut.includes(a.id)) {
        status.classList.add("failed");
        status.textContent = "No response";
      } else if (a.id === current || a.id === state.active) {
        // The active one loads from the start, underneath this screen.
        const spinner = document.createElement("span");
        spinner.className = "spinner";
        status.append(spinner, "Loading…");
      } else {
        status.textContent = "Queued";
      }
      li.append(dot, name, status);
      return li;
    }),
  );
}

$("#loading-skip").addEventListener("click", () => window.teams.skipLoading());

// ---- State ----

function render() {
  $("#loading").hidden = !state.loading;
  if (state.loading) renderLoading();
  if (menuOpen) renderMenu();
  // No accounts means nothing to show: open the add dialog right away.
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
