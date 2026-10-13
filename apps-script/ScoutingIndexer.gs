/*****************************************************************
 * ALPHARETTA SCOUTING: INDEXER + WEBSITE API   v7
 * Runs in script.google.com, signed in as alpharettadebatedocs@gmail.com.
 * Replaces v5/v6 completely (keep the RfdForm file as it is).
 *
 * The pipeline:
 *   Gmail chains --(every 10 min)--> weekly feed sheets "Chain Index <Monday>"
 *                                    + every .docx saved to Drive, viewable by link
 *   Claude (Mon + Thu) reads untagged feed rows, tags them, and shares
 *   "scouting-upload-*.json" files with this account
 *   --(every 10 min)--> this script moves them into "Alpharetta Scouting Data"
 *   Website (GitHub Pages) --> doGet() below serves docs, card outlines, and
 *   RFD notes from the Google Form as JSON.
 *
 * AFTER PASTING v7 (two things):
 *   1. Pick "setUp" in the function menu and press Run (say yes to the permissions).
 *      It finishes the one-time jobs by itself in the background over the next few hours.
 *   2. Deploy > New deployment > gear icon > Web app > Execute as "Me" > Who has access
 *      "Anyone" > Deploy. Copy the URL that ends in /exec.
 *****************************************************************/

const SHARE_WITH   = 'akshan.srivastava@gmail.com';
const SINCE        = '2026/08/01';
const ROOT_NAME    = 'Chain Index 2026-27';
const TZ           = 'America/New_York';
const MAX_MS       = 4.5 * 60 * 1000;          // hard stop per run
const BACKFILL_RUN_MS = 2 * 60 * 1000;         // backfill time allowed per run
const BACKFILL_DAY_MS = 45 * 60 * 1000;        // backfill time allowed per day (Google allows 90 min total)
const DOCX         = 'has:attachment (filename:docx OR filename:doc)';

const FEED_MAX_ROWS = 800;                     // start a "#2" sheet after this many rows in a week
const OUTLINE_MAX   = 7000;
const BODY_MAX      = 600;
const LINK_SHARE_DOCS = true;                  // make every saved doc viewable by anyone with the link
const GMAIL = 'https://mail.google.com/mail/?authuser=alpharettadebatedocs@gmail.com#all/';
const BACKLOG_FROM = new Date(2026, 9, 5);     // Oct 5 2026: earlier rows are already on the site

const HEADERS = ['key', 'date', 'round_subject', 'sender', 'filename',
                 'file_url', 'file_id', 'email_body', 'outline', 'thread_url', 'tagged'];
const DATA_NAME = 'Alpharetta Scouting Data';
const DATA_COLS = ['key', 'date', 'tourney', 'round', 'aff', 'neg', 'team', 'side', 'speaker', 'speech',
                   'tag', 'positions', 'notes', 'impacts', 'alerts', 'isNew', 'filename', 'file_url', 'thread'];
const JSON_COLS = ['positions', 'notes', 'impacts', 'alerts', 'isNew'];

/** Run once after pasting. Makes sure the triggers exist and finishes the one-time jobs in the background. */
function setUp() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'indexChains'))
    ScriptApp.newTrigger('indexChains').timeBased().everyMinutes(10).create();
  if (!triggers.some(t => t.getHandlerFunction() === 'finishSetup'))
    ScriptApp.newTrigger('finishSetup').timeBased().everyMinutes(10).create();
  dataSheet_(getRoot_());
  finishSetup();
  Logger.log('Set up. The old rows and doc sharing finish on their own in the background. Now deploy the web app.');
}

/** Background helper: moves the old rows, then shares the old docs, then removes itself. */
function finishSetup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('BACKLOG_DONE')) { moveBacklog(2.5 * 60 * 1000); return; }
  if (!p.getProperty('SHARE_DONE')) { shareOldDocs(2.5 * 60 * 1000); return; }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'finishSetup').forEach(t => ScriptApp.deleteTrigger(t));
}

/** Only needed on a brand-new install (v5 already created the trigger). */
function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('indexChains').timeBased().everyMinutes(10).create();
  indexChains();
}

/** Optional: run by hand to see where things are. */
function status() {
  const p = PropertiesService.getScriptProperties();
  Logger.log('Backfill has reached: ' + (p.getProperty('BACKFILL_DONE') ? 'DONE (back to ' + SINCE + ')' : (p.getProperty('BACKFILL_DAY') || 'not started')));
  Logger.log('Backfill minutes used today: ' + Math.round(Number(p.getProperty(usedKey_()) || 0) / 60000));
  const wk = weekKey_(new Date());
  Logger.log('This week\'s feed: Chain Index ' + wk + ' (part ' + (p.getProperty('FEED_PART_' + wk) || 1) + ')');
  const ds = dataSheet_(getRoot_());
  Logger.log('Tagged docs in ' + DATA_NAME + ': ' + Math.max(ds.getLastRow() - 1, 0));
}

function indexChains() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;                     // a run is already going
  try { run_(); } finally { lock.releaseLock(); }
}

/** Run by hand to pull in Claude's uploads right away instead of waiting 10 minutes. */
function ingestNow() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) { Logger.log('Busy, try again in a minute.'); return; }
  try { Logger.log(ingestUploads_(load_()) + ' docs added.'); } finally { lock.releaseLock(); }
}

function usedKey_() { return 'BF_USED_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd'); }

function run_() {
  const ctx = load_();

  // 1. Live rounds: everything from the last 2 days, every run.
  if (!processQuery_(ctx, DOCX + ' newer_than:2d')) return;

  // 2. Claude's tagged uploads -> data sheet.
  try { ingestUploads_(ctx); } catch (e) { Logger.log('ingest: ' + e.message); }

  // 3. Backfill: one day at a time, newest to oldest, resuming where it stopped.
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('BACKFILL_DONE')) return;
  if (!props.getProperty('SHARE_DONE')) return;          // one-time setup jobs go first
  const used = Number(props.getProperty(usedKey_()) || 0);
  if (used >= BACKFILL_DAY_MS) return;
  const bfStart = Date.now();
  ctx.deadline = Math.min(ctx.start + MAX_MS, bfStart + Math.min(BACKFILL_RUN_MS, BACKFILL_DAY_MS - used));
  try { backfill_(ctx, props); }
  finally { props.setProperty(usedKey_(), String(used + (Date.now() - bfStart))); }
}

function backfill_(ctx, props) {
  let day = props.getProperty('BACKFILL_DAY')
          ? parseDay_(props.getProperty('BACKFILL_DAY'))
          : addDays_(new Date(), -1);
  const stop = parseDay_(SINCE);

  while (day >= stop) {
    const q = DOCX + ' after:' + fmt_(addDays_(day, -1)) + ' before:' + fmt_(addDays_(day, 1));
    if (!processQuery_(ctx, q)) {                      // ran out of time mid-day
      props.setProperty('BACKFILL_DAY', fmt_(day));
      return;
    }
    day = addDays_(day, -1);
    props.setProperty('BACKFILL_DAY', fmt_(day));
  }
  props.setProperty('BACKFILL_DONE', '1');
}

/** Returns true when every thread in the query is finished, false if time ran out. */
function processQuery_(ctx, q) {
  for (let page = 0; ; page += 100) {
    const threads = GmailApp.search(q, page, 100);
    if (!threads.length) return true;

    for (const t of threads) {
      const tid = t.getId();
      const count = t.getMessageCount();
      if (ctx.seen.get(tid) === count) continue;
      if (timeUp_(ctx)) return false;

      const subject = t.getFirstMessageSubject().replace(/^((re|fwd?):\s*)+/i, '').trim();
      let finished = true;

      for (const m of t.getMessages()) {
        if (ctx.doneMsgs.has(m.getId())) continue;
        if (timeUp_(ctx)) { finished = false; break; }

        const atts = m.getAttachments({ includeInlineImages: false });
        for (let i = 0; i < atts.length; i++) {
          const key = m.getId() + '_' + i;
          const name = atts[i].getName();
          if (ctx.done.has(key) || !/\.docx?$/i.test(name)) continue;

          const blob = atts[i].copyBlob();
          const fname = subject.slice(0, 120) || '(no subject)';
          const folder = ctx.folders[fname] || (ctx.folders[fname] = getFolder_(ctx.root, fname));
          const file = folder.createFile(blob).setName(name);
          if (LINK_SHARE_DOCS) {
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
          }
          let outline;
          try { outline = outline_(blob); } catch (e) { outline = 'ERROR: ' + e.message; }

          writeRows_(ctx, [trimRow_([key, m.getDate(), subject, m.getFrom(), name,
            file.getUrl(), file.getId(), body_(m), outline, GMAIL + tid])], true);
          ctx.done.add(key);
        }
        ctx.doneMsgs.add(m.getId());
      }

      if (finished) { ctx.tsh.appendRow([tid, count]); ctx.seen.set(tid, count); }
      else return false;
    }
  }
}

/* ---------- weekly feed sheets ---------- */

/** The sheet new rows go into: "Chain Index <Monday>", then "#2", "#3"... once a sheet is full. */
function feed_(ctx) {
  if (ctx.feed && ctx.feed.rows < FEED_MAX_ROWS) return ctx.feed;
  const props = PropertiesService.getScriptProperties();
  const week = weekKey_(new Date());
  let part = Number(props.getProperty('FEED_PART_' + week) || 1);
  for (;;) {
    const id = props.getProperty('FEED_' + week + '_' + part);
    let ss;
    if (id) ss = SpreadsheetApp.openById(id);
    else {
      ss = SpreadsheetApp.create('Chain Index ' + week + (part > 1 ? ' #' + part : ''));
      DriveApp.getFileById(ss.getId()).moveTo(ctx.root);
      const sh0 = ss.getSheets()[0];
      sh0.appendRow(HEADERS);
      sh0.setFrozenRows(1);
      props.setProperty('FEED_' + week + '_' + part, ss.getId());
      props.setProperty('FEED_PART_' + week, String(part));
    }
    const sh = ss.getSheets()[0];
    const rows = Math.max(sh.getLastRow() - 1, 0);
    if (rows < FEED_MAX_ROWS) return (ctx.feed = { sh, rows });
    part++;
  }
}

function feedIds_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return Object.keys(props).filter(k => /^FEED_\d{4}-\d\d-\d\d_\d+$/.test(k)).sort().map(k => props[k]);
}

function writeRows_(ctx, rows, addKeys) {
  rows = rows.map(r => r.slice(0, 10).concat([ctx.tagged.has(String(r[0])) ? 'y' : '']));
  let i = 0;
  while (i < rows.length) {
    const f = feed_(ctx);
    const chunk = rows.slice(i, i + (FEED_MAX_ROWS - f.rows));
    f.sh.getRange(f.sh.getLastRow() + 1, 1, chunk.length, HEADERS.length).setValues(chunk);
    f.rows += chunk.length;
    i += chunk.length;
  }
  if (addKeys) ctx.keys.getRange(ctx.keys.getLastRow() + 1, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
}

function trimRow_(r) {
  r = r.slice(0, 10);
  r[7] = String(r[7] || '').slice(0, BODY_MAX);
  r[8] = String(r[8] || '').slice(0, OUTLINE_MAX);
  const m = String(r[9] || '').match(/#all\/([0-9a-f]+)/i);
  if (m) r[9] = GMAIL + m[1];
  return r;
}

/* ---------- Claude's uploads -> data sheet ---------- */

function dataSheet_(root) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DATA_ID');
  if (id) return SpreadsheetApp.openById(id).getSheets()[0];
  const ss = SpreadsheetApp.create(DATA_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(root);
  const sh = ss.getSheets()[0];
  sh.getRange(1, 1, 1000, DATA_COLS.length).setNumberFormat('@');   // keep dates/rounds as plain text
  sh.getRange(1, 1, 1, DATA_COLS.length).setValues([DATA_COLS]);
  sh.setFrozenRows(1);
  props.setProperty('DATA_ID', ss.getId());
  return sh;
}

/** Moves every scouting-upload-*.json file into the data sheet and marks those rows tagged.
 *  Claude uploads from its own Drive account and shares each file with this one, so after
 *  reading a file we trash it if we own it, otherwise rename it to scouting-done-*. */
function ingestUploads_(ctx) {
  const me = Session.getEffectiveUser().getEmail();
  const files = DriveApp.searchFiles("title contains 'scouting-upload-' and trashed = false");
  let added = 0;
  const newKeys = [];
  while (files.hasNext()) {
    const f = files.next();
    if (!/^scouting-upload-/.test(f.getName()) || /ERROR/.test(f.getName())) continue;
    let items;
    try { items = JSON.parse(f.getBlob().getDataAsString('UTF-8')); if (!Array.isArray(items)) items = [items]; }
    catch (e) { try { f.setName('scouting-upload-ERROR-' + f.getName()); } catch (x) {} continue; }
    const rows = [];
    for (const d of items) {
      const key = String(d.id || d.key || '');
      if (!key || ctx.tagged.has(key)) continue;
      const row = DATA_COLS.map(c => {
        if (c === 'key') return key;
        if (c === 'speech' && d.skip) return 'SKIP';
        if (c === 'speech' && d.snapshot) return 'SNAPSHOT';
        const v = d[c];
        if (v == null) return '';
        return JSON_COLS.includes(c) ? JSON.stringify(v) : String(v);
      });
      rows.push(row);
      ctx.tagged.add(key);
      newKeys.push(key);
    }
    if (rows.length) {
      const sh = ctx.data;
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, DATA_COLS.length).setNumberFormat('@').setValues(rows);
      added += rows.length;
    }
    let owner = '';
    try { owner = f.getOwner().getEmail(); } catch (e) {}
    if (owner === me) f.setTrashed(true);
    else f.setName(f.getName().replace(/^scouting-upload-/, 'scouting-done-'));
  }
  if (newKeys.length) markTagged_(new Set(newKeys));
  return added;
}

function markTagged_(keys) {
  for (const id of feedIds_()) {
    const sh = SpreadsheetApp.openById(id).getSheets()[0];
    const n = sh.getLastRow() - 1;
    if (n < 1) continue;
    if (sh.getRange(1, 11).getValue() !== 'tagged') sh.getRange(1, 11).setValue('tagged');
    const ks = sh.getRange(2, 1, n, 1).getValues();
    const tg = sh.getRange(2, 11, n, 1).getValues();
    let changed = false;
    for (let i = 0; i < n; i++) if (keys.has(String(ks[i][0])) && tg[i][0] !== 'y') { tg[i][0] = 'y'; changed = true; }
    if (changed) sh.getRange(2, 11, n, 1).setValues(tg);
  }
}

/* ---------- website API ---------- */

/** GET <web app url>?what=data      -> { updated, docs: [...], rfds: [...] }
 *  GET <web app url>?what=outlines  -> { outlines: [[key, outlineText], ...] } */
function doGet(e) {
  const what = (e && e.parameter && e.parameter.what) || 'data';
  let out;
  try { out = what === 'outlines' ? { outlines: apiOutlines_() } : apiData_(); }
  catch (err) { out = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiData_() {
  const sh = dataSheet_(getRoot_());
  const docs = [];
  let updated = '';
  if (sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, DATA_COLS.length).getDisplayValues();
    for (const r of vals) {
      const d = {};
      DATA_COLS.forEach((c, i) => {
        const v = r[i];
        if (v === '') return;
        if (JSON_COLS.includes(c)) { try { d[c] = JSON.parse(v); } catch (x) {} }
        else d[c === 'key' ? 'id' : c] = v;
      });
      if (d.speech === 'SKIP' || d.speech === 'SNAPSHOT') continue;
      docs.push(d);
      if ((d.date || '') > updated) updated = d.date;
    }
  }
  return { updated, docs, rfds: apiRfds_() };
}

function apiRfds_() {
  const id = PropertiesService.getScriptProperties().getProperty('RFD_SHEET_ID');
  if (!id) return [];
  const sh = SpreadsheetApp.openById(id).getSheets()[0];
  if (sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(String);
  const map = { 'Tournament': 'tourney', 'Round': 'round', 'Aff team': 'aff', 'Neg team': 'neg', 'Judge': 'judge',
                'Decision': 'decision', 'Your name': 'by', '1AC': 'ac', '1NC': 'nc', 'Block': 'block', '1AR': 'ar1',
                '2NR': 'nr', '2AR': 'ar', 'Thoughts': 'thoughts', 'Judge feedback / RFD': 'rfd' };
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i], o = {};
    head.forEach((h, j) => {
      const v = r[j];
      if (v === '' || v == null) return;
      if (h === 'Timestamp' && v instanceof Date) { o.at = Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss"); o.id = 'form-' + v.getTime(); }
      else if (map[h]) o[map[h]] = String(v).trim();
    });
    if (o.decision === 'Not disclosed') delete o.decision;
    if (Object.keys(o).length > 2) out.push(o);
  }
  return out;
}

function apiOutlines_() {
  const tagged = new Set();
  const sh = dataSheet_(getRoot_());
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 10).getDisplayValues()
    .forEach(r => { if (r[9] !== 'SKIP' && r[9] !== 'SNAPSHOT') tagged.add(r[0]); });
  const out = [];
  for (const id of feedIds_()) {
    const f = SpreadsheetApp.openById(id).getSheets()[0];
    if (f.getLastRow() < 2) continue;
    const ks = f.getRange(2, 1, f.getLastRow() - 1, 1).getValues();
    const os = f.getRange(2, 9, f.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ks.length; i++) if (tagged.has(String(ks[i][0])) && os[i][0]) out.push([String(ks[i][0]), String(os[i][0])]);
  }
  return out;
}

/* ---------- one-time jobs ---------- */

/** ONE-TIME: copy rows dated Oct 5 or later from the old big sheet into the weekly sheets. */
function moveBacklog(budgetMs) {
  budgetMs = Number(budgetMs) || 4.5 * 60 * 1000;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) { Logger.log('The indexer is running right now. Try again in a few minutes.'); return; }
  try {
    const ctx = load_();
    const sh = ctx.sheet, last = sh.getLastRow();
    const props = PropertiesService.getScriptProperties();
    let r = Number(props.getProperty('BACKLOG_ROW') || 2);
    const t0 = Date.now();
    let moved = Number(props.getProperty('BACKLOG_MOVED') || 0);
    while (r <= last) {
      if (Date.now() - t0 > budgetMs) {
        props.setProperty('BACKLOG_ROW', String(r));
        props.setProperty('BACKLOG_MOVED', String(moved));
        Logger.log('Paused at row ' + r + ' of ' + last + ' (' + moved + ' rows moved). Run moveBacklog again.');
        return;
      }
      const n = Math.min(200, last - r + 1);
      const vals = sh.getRange(r, 1, n, 10).getValues();
      const out = vals.filter(v => v[1] instanceof Date && v[1] >= BACKLOG_FROM && v[5] && v[5] !== 'SKIPPED').map(trimRow_);
      if (out.length) { writeRows_(ctx, out, false); moved += out.length; }
      r += n;
    }
    props.setProperty('BACKLOG_ROW', String(r));
    props.setProperty('BACKLOG_MOVED', String(moved));
    ctx.ss.rename('Scouting Ledger (keep, do not delete)');
    props.setProperty('BACKLOG_DONE', '1');
    Logger.log('Backlog moved: ' + moved + ' rows. Old sheet renamed to Scouting Ledger. All done.');
  } finally { lock.releaseLock(); }
}

/** ONE-TIME: make docs saved before v6 viewable by anyone with the link. Rerun until "All done". */
function shareOldDocs(budgetMs) {
  budgetMs = Number(budgetMs) || 5 * 60 * 1000;
  const t0 = Date.now();
  const types = "(mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/msword') and 'me' in owners";
  let it;
  try { it = DriveApp.searchFiles(types + " and visibility = 'limited'"); }
  catch (e) { it = DriveApp.searchFiles(types); }
  let n = 0;
  while (it.hasNext() && Date.now() - t0 < budgetMs) {
    const f = it.next();
    try {
      if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); n++;
      }
    } catch (e) {}
  }
  const more = it.hasNext();
  if (!more) PropertiesService.getScriptProperties().setProperty('SHARE_DONE', '1');
  Logger.log(n + ' docs shared. ' + (more ? 'More left.' : 'All done.'));
}

/* ---------- reading a .docx directly ---------- */

function outline_(blob) {
  const parts = Utilities.unzip(blob.setContentType('application/zip'));
  let doc = '', sty = '';
  parts.forEach(p => {
    const n = p.getName();
    if (n === 'word/document.xml') doc = p.getDataAsString('UTF-8');
    else if (n === 'word/styles.xml') sty = p.getDataAsString('UTF-8');
  });
  return doc ? outlineFromXml_(doc, sty) : '';
}

function outlineFromXml_(doc, sty) {
  const lvl = styleLevels_(sty);
  const pre = ['# ', '## ', '### ', '- '];
  const out = [];
  let wantCite = false;
  const paras = doc.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  for (const p of paras) {
    const sm = p.match(/<w:pStyle w:val="([^"]+)"/);
    let L = sm ? (lvl[sm[1]] || 0) : 0;
    if (!L) { const o = p.match(/<w:outlineLvl w:val="(\d)"/); if (o && +o[1] < 4) L = +o[1] + 1; }
    const txt = decode_((p.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || [])
                  .map(s => s.replace(/<[^>]+>/g, '')).join('')).replace(/\s+/g, ' ').trim();
    if (!txt) continue;
    if (L) {
      out.push(pre[L - 1] + txt.slice(0, L === 4 ? 150 : 200));
      wantCite = (L === 4);
    } else if (wantCite) {
      out[out.length - 1] += '  |  ' + txt.slice(0, 60);
      wantCite = false;
    }
  }
  return out.join('\n');
}

function styleLevels_(sty) {
  const map = {};
  (sty.match(/<w:style [\s\S]*?<\/w:style>/g) || []).forEach(s => {
    if (!/w:type="paragraph"/.test(s)) return;
    const id = (s.match(/w:styleId="([^"]+)"/) || [])[1];
    if (!id) return;
    const name = ((s.match(/<w:name w:val="([^"]+)"/) || [])[1] || '').toLowerCase();
    let L = 0;
    if (/heading 1\b|\bpocket\b/.test(name)) L = 1;
    else if (/heading 2\b|\bhat\b/.test(name)) L = 2;
    else if (/heading 3\b|\bblock\b(?! text)/.test(name)) L = 3;
    else if (/heading 4\b|\btag\b/.test(name)) L = 4;
    else { const o = s.match(/<w:outlineLvl w:val="(\d)"/); if (o && +o[1] < 4) L = +o[1] + 1; }
    if (L) map[id] = L;
  });
  ['Heading1', 'Heading2', 'Heading3', 'Heading4'].forEach((id, i) => { if (!map[id]) map[id] = i + 1; });
  return map;
}

function decode_(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
          .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n))
          .replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCharCode(parseInt(n, 16)))
          .replace(/&amp;/g, '&');
}

/* ---------- helpers ---------- */

function load_() {
  const root = getRoot_();
  const ss = getSpreadsheet_(root);                    // the ledger (old "Chain Index")
  const sheet = ss.getSheets()[0];
  const tsh = ss.getSheetByName('threads') || ss.insertSheet('threads');
  const keys = ss.getSheetByName('keys') || ss.insertSheet('keys');
  const done = new Set(sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().flat().map(String));
  if (keys.getLastRow() > 0) keys.getRange(1, 1, keys.getLastRow(), 1).getValues().forEach(r => done.add(String(r[0])));
  const doneMsgs = new Set([...done].map(k => k.split('_')[0]));
  const seen = new Map();
  if (tsh.getLastRow() > 0) tsh.getRange(1, 1, tsh.getLastRow(), 2).getValues().forEach(r => seen.set(String(r[0]), Number(r[1])));
  const data = dataSheet_(root);
  const tagged = new Set();
  if (data.getLastRow() > 1) data.getRange(2, 1, data.getLastRow() - 1, 1).getDisplayValues().forEach(r => tagged.add(String(r[0])));
  const start = Date.now();
  return { start, deadline: start + MAX_MS, root, ss, sheet, tsh, keys, done, doneMsgs, seen, data, tagged, folders: {}, feed: null };
}

function timeUp_(ctx) { return Date.now() > ctx.deadline; }
function fmt_(d) { return Utilities.formatDate(d, TZ, 'yyyy/MM/dd'); }
function parseDay_(s) { const [y, m, d] = s.split('/').map(Number); return new Date(y, m - 1, d, 12); }
function addDays_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function weekKey_(d) {                                  // Monday of this week, e.g. 2026-10-12
  const day = parseDay_(fmt_(d));
  return Utilities.formatDate(addDays_(day, -((day.getDay() + 6) % 7)), TZ, 'yyyy-MM-dd');
}

function body_(m) {
  return m.getPlainBody()
    .split('\n').filter(l => !/^\s*>/.test(l)).join('\n')
    .split(/\n\s*On [\s\S]{0,250}?wrote:/)[0]
    .split(/\n--\s*\nYou received this message/)[0]
    .split(/\n\s*--\s*\n/)[0]
    .trim().slice(0, BODY_MAX);
}

function getRoot_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_ID');
  if (id) return DriveApp.getFolderById(id);
  const root = DriveApp.createFolder(ROOT_NAME);
  root.addViewer(SHARE_WITH);
  props.setProperty('ROOT_ID', root.getId());
  return root;
}

function getSpreadsheet_(root) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Scouting Ledger');
  DriveApp.getFileById(ss.getId()).moveTo(root);
  const sh = ss.getSheets()[0];
  sh.appendRow(HEADERS.slice(0, 10));
  sh.setFrozenRows(1);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
