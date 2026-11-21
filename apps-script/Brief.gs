/**
 * Graficoo Focus — the morning brief (second file in the "Graficoo Focus Sync" Apps Script project).
 *
 * Every morning (installBriefTrigger → 8:00 Cairo) buildBrief():
 *   1. reads the Graficoo app sheet (projects, money, clients) and this project's Focus data (tasks, notes, habits);
 *   2. turns Graficoo work into Focus tasks with fixed ids (late jobs, money to collect, quotes to follow up),
 *      and closes them by itself when the job is delivered / paid;
 *   3. adds a growth idea when none is open (notes → tasks are done by the daily Claude run with
 *      Graficoo Brief/notes-tool.js, or by hand with «حوّل لمهام» in Focus);
 *   4. saves everything as a 'brief' record (id = the date). The Focus app shows it as «الملخص» (tick → done
 *      everywhere) and the PC sender renders it to an image + PDF and sends it on WhatsApp.
 *
 * Script property (optional): GRAFICOO_SHEET_ID — default the Graficoo_App_Data sheet.
 * Uses sheet_() and merge_() from Code.gs.
 */
const BRIEF_GRAFICOO_SHEET = '1fkySEg2GtUVtwTi5dtE0bI8cNXC76p2vW1Q2xBLbvyc';
const BRIEF_APP = 'https://ismaeloo04.github.io/graficoo-app/';
const BRIEF_FOCUS = 'https://ismaeloo04.github.io/focus-app/';
const BRIEF_QUOTES = ['مش لازم تعمل كل حاجة — اعمل الـ 3 الأولانيين، والباقي مكسب.', 'الفلوس اللي برا محتاجة مكالمة، مش استنى.', 'حاجة صغيرة كل يوم أحسن من حاجة كبيرة كل شهر.', 'خلّص الأصعب الأول وانت لسه فايق.', 'اقفل الشغل في ميعاده — بكرة محتاجك فايق.'];

/** Run once from the editor: daily trigger at 8:00 (removes older brief triggers first). */
function installBriefTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'buildBrief') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('buildBrief').timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone(TZ).create();
  Logger.log('Brief trigger set for 8:00 every day');
}

function buildBrief() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const nowMs = Date.now();
  const step = function (m) { Logger.log((Date.now() - nowMs) + ' ms — ' + m); };
  const day = function (d, n) { const x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + n); return Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); };
  const daysBetween = function (a, b) { return Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 864e5); };
  const fmt = function (n) { return Math.round(Number(n) || 0).toLocaleString('en-US'); };

  /* ---------- Focus data ---------- */
  const F = { task: {}, note: {}, habit: {}, project: {}, brief: {} };
  const sh = sheet_();
  const last = sh.getLastRow();
  (last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : []).forEach(function (r) {
    if (!F[r[0]]) return;
    try { const o = JSON.parse(r[4]); if (!o.deleted) F[r[0]][String(r[1])] = o; } catch (e) { /* skip bad row */ }
  });
  step('Focus data read (' + Object.keys(F.task).length + ' tasks, ' + Object.keys(F.note).length + ' notes)');
  const out = [];   // items to write back: {kind, o}
  const put = function (kind, o) { o.updatedAt = nowMs; out.push({ kind: kind, o: o }); F[kind][o.id] = o; };

  /* ---------- Graficoo data ---------- */
  const gss = SpreadsheetApp.openById(props.getProperty('GRAFICOO_SHEET_ID') || BRIEF_GRAFICOO_SHEET);
  const table = function (name) {
    const s = gss.getSheetByName(name);
    if (!s || s.getLastRow() < 2) return [];
    const head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    return s.getRange(2, 1, s.getLastRow() - 1, head.length).getValues().map(function (r) {
      const o = {}; head.forEach(function (h, j) { if (h) o[h] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd') : r[j]; }); return o;
    }).filter(function (o) { return o.id; });
  };
  step('Graficoo sheet opened');
  const P = table('Projects');
  step('Projects read (' + P.length + ')');
  const C = table('Clients');
  step('Clients read (' + C.length + ')');
  const settings = gss.getSheetByName('App Settings') ? gss.getSheetByName('App Settings').getDataRange().getValues() : [];
  let followDays = 3; settings.forEach(function (r) { if (r[0] === 'followDays' && Number(r[1])) followDays = Number(r[1]); });
  const clientById = {}; C.forEach(function (c) { clientById[c.id] = c; });
  const clientOf = function (p) { const id = String(p.client || '').split('|')[0].trim(); return clientById[id] || null; };
  const wa = function (p, text) {
    const c = clientOf(p); let d = String((c && (c.whatsapp || c.mobile)) || '').replace(/\D/g, '');
    if (d.indexOf('00') === 0) d = d.slice(2);
    if (d.length === 11 && d[0] === '0') d = '20' + d.slice(1);
    return 'https://wa.me/' + d + '?text=' + encodeURIComponent(text);
  };
  const stageNo = function (p) { return parseInt(String(p.stage), 10) || 0; };
  const stageAr = { 1: 'عرض سعر', 2: 'اتفاق بمقدم', 3: 'تصميم', 4: 'موافقة العميل', 5: 'طباعة / تنفيذ', 6: 'جاهز للتسليم', 7: 'اتسلم' };
  const live = P.filter(function (p) { return p.stage !== 'Cancelled'; });
  const open = live.filter(function (p) { const s = stageNo(p); return s >= 2 && s <= 6; });

  /* ---------- Graficoo → Focus tasks (fixed ids, closed by themselves when the reason is gone) ---------- */
  const want = {};   // key → task fields
  open.forEach(function (p) {
    const s = stageNo(p), late = p.delivery && p.delivery < today ? daysBetween(p.delivery, today) : 0, due = Number(p.due) || 0;
    if (!(late || p.delivery === today || s === 6)) return;
    const who = p.clientName || String(p.client || '').split('|')[1] || '';
    const title = (s === 6 ? 'سلّم ' : 'خلّص ') + p.desc + (who ? ' — ' + String(who).trim() : '') + (s === 6 && due ? ' وحصّل ' + fmt(due) + 'ج' : '');
    want['b-job-' + p.id] = { title: title, prio: late ? 2 : 1, score: (s === 6 ? 50 : 30) + late, brief: {
      sec: 'proj', meta: p.id + ' · ' + (stageAr[s] || '') + (p.delivery ? ' — ميعاده ' + (late ? 'كان ' : '') + p.delivery : '') + (due ? ' · متبقي ' + fmt(due) + 'ج' : ''),
      tag: late ? ['bad', 'متأخر ' + late + ' يوم'] : ['warn', p.delivery === today ? 'تسليم النهارده' : 'جاهز'],
      links: [['افتح المشروع', BRIEF_APP]] } };
  });
  live.filter(function (p) { return stageNo(p) === 7 && Number(p.due) > 0; }).forEach(function (p) {
    const who = String(p.clientName || String(p.client || '').split('|')[1] || 'العميل').trim();
    const since = p.delivery && p.delivery <= today ? daysBetween(p.delivery, today) : 0;
    want['b-pay-' + p.id] = { title: 'حصّل ' + fmt(p.due) + 'ج من ' + who, prio: 2, score: 60 + since, brief: {
      sec: 'money', meta: p.id + ' · ' + p.desc + (since ? ' — اتسلم من ' + since + ' يوم' : ''), tag: ['bad', 'فلوس'],
      links: [['رسالة واتساب جاهزة', wa(p, 'أهلاً ' + who + '، بفكّر حضرتك بباقي حساب «' + p.desc + '» ' + fmt(p.due) + 'ج. ممكن تحوّله على إنستاباي أو فودافون كاش. شكراً 🙏')]] } };
  });
  live.filter(function (p) { return stageNo(p) === 1 && p.sentAt; }).forEach(function (p) {
    const from = String(p.lastFollow || p.sentAt).slice(0, 10), d = daysBetween(from, today);
    if (d < followDays) return;
    const who = String(p.clientName || String(p.client || '').split('|')[1] || 'العميل').trim();
    want['b-follow-' + p.id + '-' + from] = { title: 'تابع عرض سعر «' + p.desc + '» مع ' + who, prio: 1, score: 20 + d, brief: {
      sec: 'proj', meta: (p.quoteNo || p.id) + ' · اتبعت من ' + d + ' يوم ومفيش رد', tag: ['info', 'متابعة'],
      links: [['رسالة متابعة', wa(p, 'صباح الخير ' + who + '، حضرتك لحقت تبص على عرض «' + p.desc + '»؟ لو في أي تعديل أنا موجود.')]] } };
  });
  Object.keys(F.task).forEach(function (id) {   // jobs that are delivered / paid now: close their open tasks
    const t = F.task[id];
    if (/^b-(job|pay|follow)-/.test(id) && !t.done && !want[id]) { t.done = true; t.doneAt = nowMs; t.notes = (t.notes ? t.notes + '\n' : '') + 'اتقفلت لوحدها من جرافيكوو'; put('task', t); }
  });
  Object.keys(want).forEach(function (id) {
    const w = want[id], t = F.task[id];
    if (t && t.done) return;   // you ticked it: it stays done
    const o = t || { id: id, notes: '', projectId: null, repeat: '', est: 0, done: false, createdAt: nowMs };
    o.title = w.title; o.prio = w.prio; o.due = today; o.brief = w.brief; o.score = w.score;
    put('task', o);
  });

  /* ---------- notes → tasks happen outside this job: the daily Claude run (notes-tool.js) or «حوّل لمهام» in Focus ----------
     Growth idea when none is open: call a few clients who have been quiet. */
  const openTasks = Object.keys(F.task).map(function (k) { return F.task[k]; }).filter(function (t) { return !t.done && !t.deleted; });
  const growthOpen = openTasks.filter(function (t) { return t.brief && t.brief.sec === 'growth'; });
  const dormant = C.filter(function (c) { return /Dormant|At Risk/i.test(String(c.activity)); });
  if (!growthOpen.length && dormant.length) {
    const pick = dormant.slice((new Date().getDate() * 3) % Math.max(1, dormant.length - 2)).slice(0, 3);
    put('task', { id: 'b-grow-' + today + '-0', title: 'كلّم 3 عملاء بقالهم فترة: ' + pick.map(function (c) { return c.name; }).join('، '), notes: '', projectId: null, due: today, prio: 0, repeat: '', est: 0, done: false, createdAt: nowMs,
      brief: { sec: 'growth', meta: 'عرض جديد أو سؤال عن احتياجاتهم', tag: ['info', 'عملاء'] } });
  }

  /* ---------- the brief record ---------- */
  const all = Object.keys(F.task).map(function (k) { return F.task[k]; });
  const mine = all.filter(function (t) { return !t.done && !t.deleted && !t.parentId && t.due && t.due <= today && !(t.brief && t.brief.sec); });
  const briefTasks = all.filter(function (t) { return t.brief && t.brief.sec && !t.deleted && (!t.done || (t.doneAt && nowMs - t.doneAt < 864e5 && t.due === today)) && (t.due <= today || t.brief.sec === 'growth' && !t.done); });
  const card = function (t) { return { id: t.id, title: t.title, meta: (t.brief && t.brief.meta) || (t.due < today ? 'من ' + daysBetween(t.due, today) + ' يوم' : ''), tag: (t.brief && t.brief.tag) || (t.prio === 2 ? ['bad', 'عاجل'] : t.prio === 1 ? ['warn', 'مهم'] : null), links: (t.brief && t.brief.links) || [] }; };
  const ranked = briefTasks.filter(function (t) { return !t.done && (t.brief.sec === 'proj' || t.brief.sec === 'money'); }).sort(function (a, b) { return (b.score || 0) - (a.score || 0); });
  const top = ranked.slice(0, 3).map(function (t) { return t.id; });
  const sec = function (s) { return briefTasks.filter(function (t) { return t.brief.sec === s && top.indexOf(t.id) < 0; }).map(card); };
  const money = live.filter(function (p) { return Number(p.due) > 0; }).sort(function (a, b) { return stageNo(b) - stageNo(a); }).map(function (p) {
    const s = stageNo(p);
    return { name: String(p.clientName || '').trim(), desc: p.desc, amount: Number(p.due), when: s === 7 ? 'اتسلم' + (p.delivery ? ' ' + p.delivery : '') : s === 6 ? 'عند التسليم (جاهز)' : 'الباقي عند التسليم', lvl: s === 7 ? 'bad' : s === 6 ? 'warn' : 'muted' };
  });
  const moneyNow = money.filter(function (m) { return m.lvl !== 'muted'; }).reduce(function (a, m) { return a + m.amount; }, 0);
  const yday = day(today, -1), wd = new Date(today + 'T12:00:00').getDay();
  const habits = Object.keys(F.habit).map(function (k) { return F.habit[k]; }).filter(function (h) { return !h.days || !h.days.length || h.days.indexOf(wd) >= 0; })
    .sort(function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); })
    .map(function (h) { return { id: h.id, name: h.name, target: h.target || 1, yday: (h.log && h.log[yday]) || 0, color: h.color || '' }; });
  const late = open.filter(function (p) { return p.delivery && p.delivery < today; }).length;
  const brief = {
    id: today, date: today, createdAt: nowMs, ai: false,
    stats: [[String(open.length), 'مشاريع مفتوحة', ''], [String(late), 'تسليمات متأخرة', late ? 'r' : ''], [fmt(money.reduce(function (a, m) { return a + m.amount; }, 0)) + 'ج', 'فلوس برا', 'r'],
      [String(live.filter(function (p) { return stageNo(p) === 1; }).length), 'عروض مستنية رد', '']],
    top: top.map(function (id) { return card(F.task[id]); }),
    sections: [['مهامك النهارده', mine.sort(function (a, b) { return (b.prio || 0) - (a.prio || 0); }).slice(0, 8).map(card)], ['المشاريع والعملاء', sec('proj')], ['من ملاحظاتك', sec('note')], ['تنمية جرافيكوو', sec('growth')]],
    moneyRows: money.slice(0, 8), moneyNow: moneyNow, habits: habits,
    plan: [], quote: BRIEF_QUOTES[new Date().getDate() % BRIEF_QUOTES.length],
    links: { focus: BRIEF_FOCUS + '?brief=' + today, app: BRIEF_APP },
  };
  // money tasks also appear in the money section (as tickable rows above the table)
  brief.sections.splice(2, 0, ['فلوس محتاجة تتحصّل', sec('money')]);
  put('brief', brief);
  step('brief built');
  focusWrite_(out);
  Logger.log('Brief ' + today + ': ' + out.length + ' items written');
  return brief;
}

/** Writes items into the Focus sheet the same way a device sync does (newest updatedAt wins). */
function focusWrite_(list) {
  if (!list.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    const props = PropertiesService.getScriptProperties();
    const now = Math.max(Date.now(), Number(props.getProperty('LAST_NOW') || 0) + 1);
    props.setProperty('LAST_NOW', String(now));
    const seen = {};
    const items = list.reverse().filter(function (x) { const k = x.kind + ':' + x.o.id; if (seen[k]) return false; seen[k] = 1; return true; })
      .map(function (x) { return { kind: x.kind, id: String(x.o.id), u: x.o.updatedAt, d: JSON.stringify(x.o) }; });
    const res = merge_(rows, items, now, now);
    if (res.changed && rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
    if (res.appends.length) sh.getRange(last + 1, 1, res.appends.length, 5).setValues(res.appends);
  } finally { lock.releaseLock(); }
}
