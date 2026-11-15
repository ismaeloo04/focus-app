/**
 * Graficoo Focus — sync backend (Google Apps Script web app).
 *
 * Setup (once):
 *   1. script.google.com → New project → paste this file as Code.gs (appsscript.json is optional).
 *   2. Change SYNC_KEY below to your own secret word (the same word goes in the app: الإعدادات → المزامنة).
 *   3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → copy the /exec link.
 * The data lives in a Google Sheet named "Graficoo Focus Data" in your Drive (created on the first sync).
 *
 * Protocol: POST text/plain JSON {key, since, items:[{kind,id,u,d}]}
 *   → {ok, now, items:[{kind,id,u,d}]}  (everything written after `since`).
 * Every record keeps the newest version by its updatedAt (u); deletes are records with deleted:true.
 */
const SYNC_KEY = 'CHANGE-ME';
const SHEET_NAME = 'Items';
const TZ = 'Africa/Cairo';
const HEADERS = ['kind', 'id', 'updatedAt', 'syncedAt', 'json'];

function doPost(e) {
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, err: 'bad-json' }); }
  if (SYNC_KEY === 'CHANGE-ME') return json_({ ok: false, err: 'setup' }); // refuse until the owner sets a real password
  if (!req.key || req.key !== SYNC_KEY) return json_({ ok: false, err: 'key' });
  // {fn:'translate', to:'en'|'ar', texts:[…]} → the user's own text translated by Google (source language detected)
  if (req.fn === 'translate') {
    const to = req.to === 'ar' ? 'ar' : 'en';
    const out = (req.texts || []).slice(0, 100).map(function (t) {
      t = String(t == null ? '' : t);
      if (!t.trim()) return t;
      try { return LanguageApp.translate(t, '', to); } catch (err) { return t; }
    });
    return json_({ ok: true, out: out });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    const props = PropertiesService.getScriptProperties();
    const now = Math.max(Date.now(), Number(props.getProperty('LAST_NOW') || 0) + 1);
    props.setProperty('LAST_NOW', String(now));
    const res = merge_(rows, req.items || [], Number(req.since || 0), now);
    if (res.changed && rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
    if (res.appends.length) sh.getRange(last + 1, 1, res.appends.length, 5).setValues(res.appends);
    return json_({ ok: true, now: now, items: res.out });
  } finally {
    lock.releaseLock();
  }
}

/** Pure merge (tested outside Apps Script): updates `rows` in place, returns new rows and what the client must receive. */
function merge_(rows, items, since, now) {
  const index = {};
  rows.forEach(function (r, i) { index[r[0] + ':' + r[1]] = i; });
  const appends = [];
  let changed = false;
  items.forEach(function (it) {
    if (!it || !it.kind || it.id === undefined || typeof it.d !== 'string' || it.d.length > 45000) return;
    const k = it.kind + ':' + it.id, u = Number(it.u || 0);
    if (index[k] === undefined) { index[k] = 'new:' + appends.length; appends.push([it.kind, String(it.id), u, now, it.d]); return; }
    const at = index[k];
    const row = typeof at === 'string' ? appends[Number(at.slice(4))] : rows[at];
    if (u > Number(row[2])) { row[2] = u; row[3] = now; row[4] = it.d; if (typeof at !== 'string') changed = true; }
  });
  const out = [];
  rows.concat(appends).forEach(function (r) {
    if (Number(r[3]) > since) out.push({ kind: r[0], id: String(r[1]), u: Number(r[2]), d: r[4] });
  });
  return { changed: changed, appends: appends, out: out };
}

/** GET ?key=… → today's open tasks and the timer, for home-screen widget apps (e.g. Scriptable on iPhone/iPad). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (SYNC_KEY === 'CHANGE-ME' || p.key !== SYNC_KEY) return ContentService.createTextOutput('Graficoo Focus sync is running.');
  const last = sheet_().getLastRow();
  const rows = last > 1 ? sheet_().getRange(2, 1, last - 1, 5).getValues() : [];
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const tasks = [], projects = {};
  let timer = null;
  rows.forEach(function (r) {
    let o; try { o = JSON.parse(r[4]); } catch (err) { return; }
    if (o.deleted) return;
    if (r[0] === 'project') projects[o.id] = o;
    if (r[0] === 'timer') timer = o;
    if (r[0] === 'task' && !o.done && o.due && o.due <= today) tasks.push(o);
  });
  tasks.sort(function (a, b) { return (b.prio || 0) - (a.prio || 0) || (a.due < b.due ? -1 : 1); });
  return json_({
    ok: true, today: today,
    tasks: tasks.map(function (t) { const pr = projects[t.projectId]; return { id: t.id, title: t.title, due: t.due, prio: t.prio || 0, project: pr ? pr.name : '', color: pr ? pr.color : '' }; }),
    timer: timer && { mode: timer.mode, phase: timer.phase, running: timer.running, endAt: timer.endAt, left: timer.left, task: timer.taskId },
  });
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Graficoo Focus Data');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
    sh.getRange(1, 1, 1, 5).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create the data sheet and grant permissions before deploying. */
function setup() {
  Logger.log('Data sheet: ' + sheet_().getParent().getUrl());
}
