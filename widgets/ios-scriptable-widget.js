// Graficoo Focus — home-screen widget for iPhone / iPad (free "Scriptable" app from the App Store).
// 1) Scriptable → + → paste this file → set SYNC_URL, SYNC_KEY, APP_URL below → name it "Focus".
// 2) Home screen → long-press → + → Scriptable → pick a size (small / medium / large) → Edit widget → Script: Focus.
// Shows the timer (ends at …) and today's tasks from the sync sheet; tapping it opens the app.
const SYNC_URL = 'https://script.google.com/macros/s/XXXX/exec'; // same link as in the app settings
const SYNC_KEY = 'CHANGE-ME';                                     // same secret word
const APP_URL = 'https://example.github.io/focus/';               // the installed app link

const RED = new Color('#EE2E24'), BG = new Color('#0A0A0A'), INK = Color.white(), MUTED = new Color('#9A9A9E');
const FONT = (s, w) => (w ? Font.boldSystemFont(s) : Font.systemFont(s));

async function load() {
  try { const r = new Request(`${SYNC_URL}?key=${encodeURIComponent(SYNC_KEY)}`); r.timeoutInterval = 15; return await r.loadJSON(); }
  catch (e) { return { ok: false }; }
}
function timeText(ms) { const d = new Date(ms); return d.toLocaleTimeString('ar-EG', { hour: 'numeric', minute: '2-digit' }); }

const data = await load();
const size = config.widgetFamily || 'medium';
const w = new ListWidget();
w.backgroundColor = BG; w.url = APP_URL; w.setPadding(14, 14, 12, 14);
w.refreshAfterDate = new Date(Date.now() + 10 * 60 * 1000);

const head = w.addStack(); head.centerAlignContent();
const brand = head.addText('FOCUS'); brand.font = FONT(13, true); brand.textColor = INK;
head.addSpacer();
const t = data.ok && data.timer;
if (t && t.running && t.endAt > Date.now()) {
  const label = t.mode === 'timer' ? 'تايمر' : t.phase === 'focus' ? 'تركيز' : 'راحة';
  const s = head.addText(`${label} · يخلص ${timeText(t.endAt)}`); s.font = FONT(11, true); s.textColor = RED;
} else {
  const s = head.addText(data.ok ? `${data.tasks.length} مهام النهارده` : 'مفيش اتصال'); s.font = FONT(11); s.textColor = MUTED;
}
w.addSpacer(8);
const bar = w.addStack(); bar.size = new Size(0, 2); bar.backgroundColor = RED; bar.cornerRadius = 1;
w.addSpacer(8);

const max = size === 'small' ? 3 : size === 'large' ? 9 : 4;
const tasks = (data.ok ? data.tasks : []).slice(0, max);
if (!tasks.length) { const e = w.addText(data.ok ? 'يومك فاضي 👌' : 'افتح التطبيق وزامن'); e.font = FONT(13); e.textColor = MUTED; }
for (const k of tasks) {
  const row = w.addStack(); row.centerAlignContent(); row.layoutHorizontally();
  row.addSpacer(); // right-aligned rows (Arabic)
  const tx = row.addText(k.title); tx.font = FONT(size === 'small' ? 12 : 13, k.prio > 0); tx.textColor = INK; tx.lineLimit = 1; tx.rightAlignText();
  row.addSpacer(6);
  const dot = row.addText('●'); dot.font = FONT(9); dot.textColor = k.prio === 2 ? RED : k.color ? new Color(k.color) : MUTED;
  w.addSpacer(5);
}
w.addSpacer();
if (config.runsInWidget) Script.setWidget(w); else await w.presentMedium();
Script.complete();
