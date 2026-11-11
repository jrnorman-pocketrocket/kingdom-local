/**
 * Kingdom Local: Google Sheet backend (v2.0).
 *
 * HOW TO INSTALL
 * 1. In the Google Sheet: Extensions > Apps Script. Select all the old code, delete it, paste this file, save.
 * 2. Project Settings (cog icon): set Time zone to (GMT+10:00) Brisbane.
 * 3. Back in the editor, pick "setup" in the function dropdown and press Run. Approve the permissions.
 *    This adds the new columns, gives every listing an ID, and switches on the email alerts.
 * 4. Pick "testEmail" and press Run. You should get an email within a minute. If not, check spam.
 * 5. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *    (Same deployment, so the website URL doesn't change.)
 *
 * WHAT IT DOES
 * - New signup: emails NOTIFY_TO straight away, with a link to the exact row.
 * - Every morning at 7am: one reminder email if any signups are still waiting (silent if none).
 * - When you tick Approved: the listing goes live within seconds, and the business gets a
 *   "you're live" email with a ready-to-post share message (only once, and only if they gave an email).
 * - Phone and email are no longer sent to the website in bulk. Each one is fetched only when a
 *   visitor taps "Show phone" or "Show email", with a cap so nobody can harvest the lot.
 * - Signups that come in faster than a human can type, duplicates, and floods are dropped.
 */

// ===== SETTINGS =====
const NOTIFY_TO = 'james@pocketrocketsolutions.com.au'; // where signup alerts go
const SITE_URL = 'https://kingdomlocal.net.au/';
const SEND_LIVE_EMAIL = true;     // email businesses when their listing is approved
const DAILY_REMINDER_HOUR = 7;    // 7am, in the time zone set in Project Settings
const MAX_SIGNUPS_PER_HOUR = 30;  // anything beyond this is treated as a spam run
const MAX_REVEALS_PER_MINUTE = 60; // phone/email taps across the whole site
// ====================

const SHEET_NAME = 'Listings';

// [header in the sheet, key used by the website]. Matched by header name, so order in the sheet doesn't matter.
const COLUMNS = [
  ['Submitted', null],
  ['Business name', 'name'],
  ['What you do', 'what'],
  ['Category', 'category'],
  ['Suburb', 'area'],
  ['Service area', 'serves'],
  ['Website', 'website'],
  ['Phone', 'phone'],
  ['Public email', 'email'],
  ['Facebook / Instagram', 'social'],
  ['Contact name (private)', 'contact'],
  ['Approved', null],
  ['Google Business Profile', 'google'],
  ['ID', null],
  ['Live email sent', null]
];
const PRIVATE = ['contact'];          // never leaves the sheet
const ON_TAP = ['phone', 'email'];    // only sent one listing at a time, when tapped
const FEED_KEY = 'listings-v2';

// ---------- one-off setup ----------

function setup() {
  const s = sheet_();
  ensureIds_(s);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onApprove', 'dailyReminder'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onApprove').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(DAILY_REMINDER_HOUR).create();
  // Existing approved listings shouldn't get a surprise "you're live" email.
  const v = s.sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (isTicked_(v[i][s.col['Approved'] - 1]) && !v[i][s.col['Live email sent'] - 1]) {
      s.sh.getRange(i + 1, s.col['Live email sent']).setValue('(approved before v2)');
    }
  }
  clearFeed_();
}

function testEmail() {
  MailApp.sendEmail(NOTIFY_TO, 'Kingdom Local: alerts are working',
    'This is a test. You\'ll get an email like this whenever a business signs up.');
}

// ---------- helpers ----------

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const header = sh.getLastRow() ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  let changed = false;
  COLUMNS.forEach(([title]) => { if (header.indexOf(title) < 0) { header.push(title); changed = true; } });
  if (changed) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const col = {};
  header.forEach((h, i) => { if (h) col[h] = i + 1; });
  return { sh: sh, col: col, width: header.length };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isTicked_(v) { return v === true || /^(y|yes|true)$/i.test(String(v).trim()); }

function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function ensureIds_(s) {
  const n = s.sh.getLastRow() - 1;
  if (n < 1) return;
  const rng = s.sh.getRange(2, s.col['ID'], n, 1);
  const ids = rng.getValues();
  let changed = false;
  ids.forEach(r => { if (!r[0]) { r[0] = newId_(); changed = true; } });
  if (changed) rng.setValues(ids);
}

function clearFeed_() { CacheService.getScriptCache().removeAll([FEED_KEY]); }

// Counts events in a rolling window; returns false once the limit is hit.
function underLimit_(key, limit, seconds) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key) || 0);
  if (n >= limit) return false;
  cache.put(key, String(n + 1), seconds);
  return true;
}

function rowLink_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getUrl() + '#gid=' + ss.getSheetByName(SHEET_NAME).getSheetId() + '&range=A' + row;
}

// ---------- website reads ----------

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.c) return contact_(String(p.c));

  const cache = CacheService.getScriptCache();
  const hit = cache.get(FEED_KEY);
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);

  const s = sheet_();
  const values = s.sh.getDataRange().getValues();
  const listings = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const id = String(r[s.col['ID'] - 1] || '');
    if (!id || !isTicked_(r[s.col['Approved'] - 1])) continue;
    const o = { id: id };
    COLUMNS.forEach(([title, key]) => {
      if (!key || PRIVATE.indexOf(key) >= 0) return;
      const v = String(r[s.col[title] - 1] || '').trim();
      if (!v) return;
      if (ON_TAP.indexOf(key) >= 0) o[key === 'phone' ? 'p' : 'm'] = 1; // just "has one"
      else o[key] = v;
    });
    if (o.name) listings.push(o);
  }
  const body = JSON.stringify({ listings: listings });
  try { cache.put(FEED_KEY, body, 300); } catch (err) { /* too big to cache, fine */ }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// One listing's phone and email, only when a visitor taps to show them.
function contact_(id) {
  if (!/^[a-f0-9]{10}$/.test(id)) return json_({ ok: false });
  if (!underLimit_('reveals', MAX_REVEALS_PER_MINUTE, 60)) {
    return json_({ ok: false, error: 'Lots of people are looking right now. Please try again in a minute.' });
  }
  const s = sheet_();
  const found = s.sh.getRange(2, s.col['ID'], Math.max(s.sh.getLastRow() - 1, 1), 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  if (!found) return json_({ ok: false });
  const r = s.sh.getRange(found.getRow(), 1, 1, s.width).getValues()[0];
  if (!isTicked_(r[s.col['Approved'] - 1])) return json_({ ok: false });
  return json_({
    ok: true,
    phone: String(r[s.col['Phone'] - 1] || '').trim(),
    email: String(r[s.col['Public email'] - 1] || '').trim()
  });
}

// ---------- website signups ----------

function doPost(e) {
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    // Bots: filled the hidden field, or submitted faster than a person could type. Pretend it worked.
    if (d.hp) return json_({ ok: true });
    if (typeof d.t === 'number' && d.t < 3000) return json_({ ok: true });

    const clean = v => {
      let t = String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, 300);
      if (/^[=+\-@]/.test(t)) t = "'" + t; // stop spreadsheet formula injection
      return t;
    };
    if (!clean(d.name) || !clean(d.what) || !clean(d.area) || !clean(d.contact) || d.consent !== true) {
      return json_({ ok: false, error: 'Please fill in the required fields and tick the box.' });
    }

    // Same business sent twice in 10 minutes (double tap, impatient refresh): keep the first.
    const dupKey = 'dup-' + clean(d.name).toLowerCase().replace(/\W+/g, '').slice(0, 60);
    const cache = CacheService.getScriptCache();
    if (cache.get(dupKey)) return json_({ ok: true });

    if (!underLimit_('signups', MAX_SIGNUPS_PER_HOUR, 3600)) {
      return json_({ ok: false, error: 'We\'re getting a lot of signups right now. Please try again in an hour.' });
    }
    cache.put(dupKey, '1', 600);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    const s = sheet_();
    const row = new Array(s.width).fill('');
    COLUMNS.forEach(([title, key]) => { if (key) row[s.col[title] - 1] = clean(d[key]); });
    const pc = s.col['Phone'] - 1; // keep the leading 0 on mobiles
    if (row[pc] && row[pc].charAt(0) !== "'") row[pc] = "'" + row[pc];
    row[s.col['Submitted'] - 1] = new Date();
    row[s.col['ID'] - 1] = newId_();
    s.sh.appendRow(row);
    const rowNum = s.sh.getLastRow();
    s.sh.getRange(rowNum, s.col['Approved']).insertCheckboxes();
    lock.releaseLock();

    // The signup is saved, so an email hiccup must never show the business an error.
    try {
      MailApp.sendEmail({
        to: NOTIFY_TO,
        subject: 'New Kingdom Local signup: ' + clean(d.name),
        body: [
          clean(d.name) + ' (' + clean(d.area) + ') wants to be listed.',
          '',
          'What they do: ' + clean(d.what),
          'Category: ' + clean(d.category),
          'Website: ' + (clean(d.website) || '-'),
          'Social: ' + (clean(d.social) || '-'),
          'Submitted by: ' + clean(d.contact),
          '',
          'Tick Approved to publish it (opens the right row):',
          rowLink_(rowNum)
        ].join('\n')
      });
    } catch (mailErr) { console.error(mailErr); }

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong. Please try again.' });
  }
}

// ---------- approvals ----------

// Runs on every edit in the sheet (installed by setup). Acts only when Approved is ticked.
function onApprove(e) {
  if (!e || !e.range) return;
  const s = sheet_();
  const rng = e.range;
  if (rng.getSheet().getName() !== SHEET_NAME) return;
  const aCol = s.col['Approved'];
  if (aCol < rng.getColumn() || aCol > rng.getLastColumn()) return;
  clearFeed_(); // ticked or unticked, show the change on the site straight away

  if (!SEND_LIVE_EMAIL) return;
  for (let row = Math.max(rng.getRow(), 2); row <= rng.getLastRow(); row++) {
    const r = s.sh.getRange(row, 1, 1, s.width).getValues()[0];
    if (!isTicked_(r[aCol - 1]) || r[s.col['Live email sent'] - 1]) continue;
    if (!r[s.col['ID'] - 1]) s.sh.getRange(row, s.col['ID']).setValue(newId_());
    const to = String(r[s.col['Public email'] - 1] || '').trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) {
      s.sh.getRange(row, s.col['Live email sent']).setValue('(no email)');
      continue;
    }
    try {
      sendLiveEmail_(to, String(r[s.col['Contact name (private)'] - 1] || '').trim(), String(r[s.col['Business name'] - 1] || '').trim());
      s.sh.getRange(row, s.col['Live email sent']).setValue(new Date());
    } catch (err) { console.error(err); }
  }
}

function sendLiveEmail_(to, person, business) {
  const first = (person.split(/\s+/)[0] || 'there');
  const link = SITE_URL + '?q=' + encodeURIComponent(business);
  const post = 'We\'re now listed on Kingdom Local, a free directory of Christian-owned businesses on the Sunshine Coast. ' +
    'If you\'re a Christian business owner, add yourself too, it takes 2 minutes and it\'s free: ' + SITE_URL;
  const text = [
    'Hi ' + first + ',',
    '',
    business + ' is now live on Kingdom Local: ' + link,
    '',
    'The directory grows when business owners invite each other. If you have a minute, here\'s a post you can copy onto your Facebook or Instagram:',
    '',
    post,
    '',
    'There\'s also a "Find us on Kingdom Local" badge for your website or socials: ' + SITE_URL + 'badge.png',
    '',
    'Need anything changed? Just reply to this email.',
    '',
    'Blessings,',
    'James',
    'Kingdom Local'
  ].join('\n');
  const esc = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const html =
    '<p>Hi ' + esc(first) + ',</p>' +
    '<p><b>' + esc(business) + '</b> is now live on Kingdom Local. <a href="' + link + '">See your listing</a>.</p>' +
    '<p>The directory grows when business owners invite each other. If you have a minute, here\'s a post you can copy onto your Facebook or Instagram:</p>' +
    '<blockquote style="border-left:3px solid #f2b632;margin:0;padding:8px 12px;background:#f6f8f4">' + esc(post) + '</blockquote>' +
    '<p>You can also add our badge to your website or socials:<br><a href="' + SITE_URL + 'badge.png"><img src="' + SITE_URL +
    'badge.png" alt="Find us on Kingdom Local" width="300"></a></p>' +
    '<p>Need anything changed? Just reply to this email.</p><p>Blessings,<br>James<br>Kingdom Local</p>';
  MailApp.sendEmail({ to: to, subject: 'You\'re live on Kingdom Local', body: text, htmlBody: html, name: 'Kingdom Local', replyTo: NOTIFY_TO });
}

// ---------- morning reminder ----------

function dailyReminder() {
  const s = sheet_();
  const v = s.sh.getDataRange().getValues();
  const cutoff = Date.now() - 21 * 864e5; // ignore anything older than 3 weeks (you've decided to leave those)
  const waiting = [];
  let firstRow = 0;
  for (let i = 1; i < v.length; i++) {
    const when = v[i][s.col['Submitted'] - 1];
    if (isTicked_(v[i][s.col['Approved'] - 1]) || !(when instanceof Date) || when.getTime() < cutoff) continue;
    waiting.push('- ' + v[i][s.col['Business name'] - 1] + ' (' + v[i][s.col['Suburb'] - 1] + ')');
    if (!firstRow) firstRow = i + 1;
  }
  if (!waiting.length) return;
  MailApp.sendEmail(NOTIFY_TO, 'Kingdom Local: ' + waiting.length + ' waiting for approval',
    waiting.join('\n') + '\n\nOpen the sheet:\n' + rowLink_(firstRow));
}
