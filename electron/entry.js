// Entry point: the app itself or, with --meeting-alert, the process that shows the fullscreen
// meeting alert on every monitor (see alert-process.js).
if (process.argv.includes("--meeting-alert")) await import("./alert-process.js");
else await import("./main.js");
