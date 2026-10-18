const $ = (sel) => document.querySelector(sel);
let meeting = null;

const clock = (ms) => new Date(ms).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

function whenText() {
  const now = Date.now();
  const seconds = Math.round((meeting.start - now) / 1000);
  if (seconds > 59) {
    const minutes = Math.ceil(seconds / 60);
    return minutes < 60 ? `Starts in ${minutes} min` : `Starts in ${Math.floor(minutes / 60)} h ${minutes % 60} min`;
  }
  if (seconds > 0) return `Starts in ${seconds} s`;
  const late = Math.floor(-seconds / 60);
  return late < 1 ? "Starting now" : `Started ${late} min ago`;
}

function tick() {
  if (!meeting) return;
  const when = $("#when");
  when.textContent = whenText();
  when.classList.toggle("now", Date.now() >= meeting.start - 60_000);
  // Once the meeting starts, "remind me at start" no longer makes sense.
  if (Date.now() >= meeting.start) $("#snooze").hidden = true;
}

window.meetingAlert.onAlert((data) => {
  meeting = data;
  $("#subject").textContent = data.subject;
  $("#time").textContent = `${clock(data.start)} – ${clock(data.end)}`;
  $("#account").hidden = !data.account;
  $("#account-name").textContent = data.account ?? "";
  $("#account-dot").style.background = data.color ?? "";
  $("#join").hidden = !data.canJoin;
  $("#snooze").hidden = !data.canSnooze;
  tick();
  ($("#join").hidden ? $("#dismiss") : $("#join")).focus();
});

$("#join").addEventListener("click", () => window.meetingAlert.action("join"));
$("#snooze").addEventListener("click", () => window.meetingAlert.action("snooze"));
$("#dismiss").addEventListener("click", () => window.meetingAlert.action("dismiss"));

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") window.meetingAlert.action("dismiss");
  else if (e.key === "Enter" && !$("#join").hidden && document.activeElement?.tagName !== "BUTTON") window.meetingAlert.action("join");
});

setInterval(tick, 1000);
