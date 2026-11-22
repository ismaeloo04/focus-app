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

/** Run once from the editor: daily triggers — build at 8:00, email at ~9:45 (removes older ones first). */
function installBriefTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'buildBrief') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('buildBrief').timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone(TZ).create();
  Logger.log('Brief trigger set for 8:00 every day');
  installBriefEmailTrigger();   // and the email at ~9:45
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

/* ===================== the brief by email (free, runs on Google — no PC, no WhatsApp) =====================
   installBriefEmailTrigger() once: sendBriefEmail every day around 9:45 (after the Claude notes run at ~9:16).
   Script property BRIEF_EMAIL = where it goes (default: this Google account). testBriefEmail() sends it now. */
function installBriefEmailTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendBriefEmail') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendBriefEmail').timeBased().everyDays(1).atHour(9).nearMinute(45).inTimezone(TZ).create();
  Logger.log('Brief email trigger set for ~9:45 every day → ' + briefEmailTo_());
}
function testBriefEmail() { sendBriefEmail(true); }
function briefEmailTo_() { return PropertiesService.getScriptProperties().getProperty('BRIEF_EMAIL') || Session.getEffectiveUser().getEmail(); }

function sendBriefEmail(force) {
  force = force === true;
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!force && props.getProperty('BRIEF_EMAILED') === today) return;
  let b = null;
  const sh = sheet_(), last = sh.getLastRow();
  (last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : []).forEach(function (r) {
    if (r[0] === 'brief' && String(r[1]) === today) { try { b = JSON.parse(r[4]); } catch (e) { /* ignore */ } }
  });
  if (!b) b = buildBrief();   // the 8:00 run did not happen: make it now
  const html = briefHtml_(b);
  const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName('ملخص ' + today + '.pdf');
  const top = (b.top || []).length;
  MailApp.sendEmail({
    to: briefEmailTo_(), name: 'Graficoo — ملخص الصبح',
    subject: '☀️ ملخص ' + briefDateAr_(today) + (top ? ' — ' + top + ' حاجات مهمة' : '') + (b.moneyNow ? ' · مستحق ' + Number(b.moneyNow).toLocaleString('en-US') + 'ج' : ''),
    htmlBody: html, attachments: [pdf],
  });
  props.setProperty('BRIEF_EMAILED', today);
  Logger.log('Brief email sent to ' + briefEmailTo_());
}

function briefDateAr_(d) {
  const x = new Date(d + 'T12:00:00');
  const days = ['الأحد', 'الاتنين', 'التلات', 'الأربع', 'الخميس', 'الجمعة', 'السبت'];
  const months = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  return days[x.getDay()] + ' ' + x.getDate() + ' ' + months[x.getMonth()];
}

/** The brief as email-safe HTML (inline styles, tables) — the same HTML becomes the PDF. */
function briefHtml_(b) {
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
  const money = function (n) { return Math.round(Number(n) || 0).toLocaleString('en-US') + 'ج'; };
  const RED = '#EE2E24', INK = '#141414', G1 = '#F3F3F3', G2 = '#E6E6E6', MUT = '#6B6B6B';
  const F = "font-family:Tahoma,'Segoe UI',Arial,sans-serif;";
  const focus = (b.links && b.links.focus) || BRIEF_FOCUS;
  const tagCss = { bad: 'background:' + RED + ';color:#fff;', warn: 'background:#FFE7E5;color:#B3160E;', info: 'background:' + INK + ';color:#fff;', muted: 'background:' + G1 + ';color:' + MUT + ';' };
  const chip = function (label, href, wa) {
    return '<a href="' + esc(href) + '" style="display:inline-block;margin:6px 0 0 6px;padding:5px 12px;border-radius:14px;font-size:12px;font-weight:bold;text-decoration:none;'
      + (wa ? 'background:#E7F7EC;color:#0E7A39;' : 'background:' + G1 + ';color:' + INK + ';') + '">' + esc(label) + '</a>';
  };
  const row = function (x, n) {
    const tag = x.tag ? ' <span style="display:inline-block;font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px;' + (tagCss[x.tag[0]] || tagCss.muted) + '">' + esc(x.tag[1]) + '</span>' : '';
    const links = (x.links || []).map(function (l) { return chip(l[0], l[1], /wa\.me/.test(l[1])); }).join('');
    return '<tr><td style="width:26px;vertical-align:top;padding:12px 0;border-bottom:1px solid ' + G2 + ';"><div style="width:16px;height:16px;border:2px solid ' + INK + ';border-radius:5px;"></div></td>'
      + '<td style="vertical-align:top;padding:12px 0;border-bottom:1px solid ' + G2 + ';">'
      + '<div style="font-weight:bold;font-size:15px;line-height:1.5;">' + (n ? '<span style="color:' + RED + ';">' + n + '. </span>' : '') + esc(x.title) + tag + '</div>'
      + (x.meta ? '<div style="color:' + MUT + ';font-size:13px;line-height:1.5;margin-top:2px;">' + esc(x.meta) + '</div>' : '') + (links ? '<div>' + links + '</div>' : '') + '</td></tr>';
  };
  const card = function (list, strong) {
    return list.length ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:' + (strong ? '2px solid ' + INK : '1px solid ' + G2) + ';border-radius:14px;padding:2px 14px;border-collapse:separate;">'
      + list.map(function (x, i) { return row(x, strong ? i + 1 : 0); }).join('') + '</table>' : '';
  };
  const h2 = function (t) { return '<div style="font-size:18px;font-weight:bold;margin:22px 0 10px;border-right:5px solid ' + RED + ';padding-right:10px;">' + esc(t) + '</div>'; };
  let body = '';
  if ((b.top || []).length) body += h2('أهم 3 حاجات النهارده') + card(b.top, true);
  (b.sections || []).forEach(function (s) {
    const title = s[0], list = s[1] || [];
    let extra = '';
    if (title === 'فلوس محتاجة تتحصّل' && (b.moneyRows || []).length) {
      extra = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;border:1px solid ' + G2 + ';border-radius:14px;padding:2px 14px;border-collapse:separate;">'
        + b.moneyRows.map(function (m) {
          return '<tr><td style="padding:10px 0;border-bottom:1px solid ' + G2 + ';"><b>' + esc(m.name) + '</b><div style="color:' + MUT + ';font-size:12px;">' + esc(m.desc) + '</div></td>'
            + '<td style="padding:10px 6px;border-bottom:1px solid ' + G2 + ';font-weight:bold;font-size:16px;white-space:nowrap;color:' + (m.lvl === 'bad' ? RED : m.lvl === 'muted' ? MUT : INK) + ';">' + money(m.amount) + '</td>'
            + '<td style="padding:10px 0;border-bottom:1px solid ' + G2 + ';color:' + MUT + ';font-size:12px;text-align:left;">' + esc(m.when) + '</td></tr>';
        }).join('') + '</table>'
        + (b.moneyNow ? '<div style="margin-top:8px;background:' + INK + ';color:#fff;border-radius:10px;padding:10px 14px;font-weight:bold;">المستحق دلوقتي (متسلم أو جاهز): ' + money(b.moneyNow) + '</div>' : '');
    }
    if (list.length || extra) body += h2(title) + card(list, false) + extra;
  });
  if ((b.habits || []).length) {
    body += h2('انت وصحتك') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' + b.habits.slice(0, 4).map(function (h) {
      return '<td style="vertical-align:top;padding:4px;"><div style="background:' + G1 + ';border-radius:12px;padding:10px;"><b>' + esc(h.name) + '</b><div style="color:' + MUT + ';font-size:12px;margin-top:3px;">'
        + (h.target > 1 ? 'الهدف ' + h.target + ' — امبارح ' + h.yday : h.yday ? 'امبارح ✓' : 'امبارح لأ') + '</div></div></td>';
    }).join('') + '</tr></table>';
  }
  if ((b.plan || []).length) {
    body += h2('شكل اليوم المقترح') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ' + G2 + ';border-radius:14px;padding:2px 14px;border-collapse:separate;">'
      + b.plan.map(function (p) { return '<tr><td style="padding:8px 0;color:' + RED + ';font-weight:bold;width:60px;direction:ltr;text-align:right;border-bottom:1px dashed ' + G2 + ';">' + esc(p[0]) + '</td><td style="padding:8px 10px;font-weight:bold;border-bottom:1px dashed ' + G2 + ';">' + esc(p[1]) + '</td></tr>'; }).join('') + '</table>';
  }
  if (b.quote) body += '<div style="margin-top:20px;background:' + G1 + ';border-right:5px solid ' + RED + ';border-radius:12px;padding:14px 16px;font-weight:bold;line-height:1.7;">«' + esc(b.quote) + '»</div>';
  const stats = (b.stats || []).map(function (s) {
    return '<td style="padding:4px;"><div style="background:#fff;border-radius:12px;padding:10px;"><div style="font-size:20px;font-weight:bold;white-space:nowrap;color:' + (s[2] ? RED : INK) + ';">' + esc(s[0]) + '</div><div style="color:' + MUT + ';font-size:12px;">' + esc(s[1]) + '</div></div></td>';
  }).join('');
  return '<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>@page{size:A4;margin:12mm}</style></head>'
    + '<body dir="rtl" style="margin:0;background:#fff;color:' + INK + ';' + F + '">'
    + '<div dir="rtl" style="max-width:640px;margin:0 auto;text-align:right;' + F + '">'
    + '<div style="background:' + INK + ';color:#fff;padding:24px 22px;border-radius:0 0 16px 16px;">'
    + '<div style="font-weight:bold;letter-spacing:1px;color:#bbb;font-size:13px;">GRAFICOO STUDIO</div>'
    + '<div style="font-size:26px;font-weight:bold;margin-top:10px;">صباح الخير يا اسماعيلوو ده يومك ☀️</div>'
    + '<div style="color:#bbb;margin-top:4px;">' + briefDateAr_(b.date) + '</div>'
    + '<a href="' + esc(focus) + '" style="display:inline-block;margin-top:14px;background:' + RED + ';color:#fff;font-weight:bold;padding:9px 16px;border-radius:18px;text-decoration:none;">افتح الملخص في Focus وعلّم اللي خلص ←</a></div>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + G1 + ';margin-top:10px;border-radius:14px;"><tr>' + stats + '</tr></table>'
    + '<div style="padding:0 4px;">' + body + '</div>'
    + '<div style="margin-top:22px;background:' + INK + ';color:#ccc;padding:16px 18px;border-radius:14px;font-size:12.5px;line-height:1.8;">'
    + '<b style="color:#fff;">تعلّم اللي خلص:</b> افتح الملخص في Focus ودوس على المربع.<br>'
    + '<b style="color:#fff;">ملاحظاتك:</b> اكتب اللي عاوز تعمله في ملاحظة على Focus — Claude بيحوّلها لمهام كل صبح، أو دوس «حوّل لمهام» في الملاحظة نفسها.</div>'
    + '</div></body></html>';
}
