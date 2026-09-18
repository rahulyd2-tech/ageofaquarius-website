/**
 * Age of Aquarius Group — Lead capture backend (Google Apps Script Web App).
 *
 * Deploy: Extensions > Apps Script on a new Google Sheet, paste this file,
 * run setup() once, then Deploy > New deployment > Web app
 * (Execute as: Me — Who has access: Anyone with the link).
 * Put the /exec URL and the token into the website's LEAD_ENDPOINT / LEAD_TOKEN.
 *
 * Script Properties used (Project Settings > Script Properties):
 *   TOKEN       shared secret the website sends with every request (setup() creates one)
 *   ALERT_EMAIL where lead alerts are emailed          (default hello@ageofaquariusdevelopers.in)
 *   WA_TOKEN    Meta WhatsApp Cloud API permanent token   (optional, phase 2)
 *   WA_PHONE_ID Cloud API phone number ID                 (optional, phase 2)
 *   WA_TO       owner's number in E.164, e.g. 918956552480 (optional, phase 2)
 */

var SHEET_LEADS = 'Leads';
var SHEET_ACTIVITY = 'Activity';
var SHEET_CONFIG = 'Config';
var SHEET_BLOCKED = 'Blocked';

var LEAD_COLUMNS = [
  'lead_id', 'created_at', 'lang', 'track', 'name', 'mobile', 'whatsapp_ok', 'email', 'company',
  'role', 'location', 'area_value', 'area_unit', 'need', 'timeline', 'budget_band', 'org_type',
  'sites', 'users', 'stage', 'notes', 'consent', 'score', 'band', 'status', 'owner',
  'next_action_at', 'last_contact_at', 'outcome', 'value_est', 'internal_notes',
  'page', 'referrer', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'utm_content',
  'gclid', 'device', 'fill_seconds', 'wa_status'
];

var ACTIVITY_COLUMNS = ['ts', 'lead_id', 'actor', 'type', 'summary', 'next_step'];

var BLOCKED_COLUMNS = ['ts', 'reason', 'name', 'mobile', 'email', 'track', 'lang', 'notes', 'client_id', 'page', 'referrer', 'device', 'fill_seconds', 'raw'];

var DEFAULT_CONFIG = [
  ['key', 'value', 'note'],
  ['owner', 'Rahul', 'Owner of every new lead'],
  ['sla_hot_hours', '2', 'First response target for Hot leads'],
  ['sla_warm_hours', '24', 'First response target for Warm leads'],
  ['sla_cold_hours', '72', 'First response target for Cold leads'],
  ['w_track_package', '25', 'Score weight — Full Package'],
  ['w_track_land', '20', 'Score weight — Land Development'],
  ['w_track_platform', '20', 'Score weight — Technology Platform'],
  ['w_area_large', '20', 'Score weight — more than 10 acres'],
  ['w_area_mid', '15', 'Score weight — 2 to 10 acres'],
  ['w_area_small', '8', 'Score weight — under 2 acres'],
  ['w_time_0_3', '20', 'Score weight — timeline under 3 months'],
  ['w_time_3_6', '12', 'Score weight — timeline 3 to 6 months'],
  ['w_time_6_plus', '5', 'Score weight — timeline over 6 months'],
  ['w_role_owner', '15', 'Score weight — owner or decision maker'],
  ['w_completeness', '10', 'Score weight — form filled fully'],
  ['w_contactable', '10', 'Score weight — valid Indian mobile'],
  ['band_hot_min', '70', 'Score at or above this is Hot'],
  ['band_warm_min', '40', 'Score at or above this is Warm'],
  ['digest_hour', '9', 'Hour (IST) for the daily digest email'],
  ['spam_min_seconds', '5', 'Reject enquiries filled faster than this many seconds'],
  ['spam_max_per_client_hour', '3', 'Max enquiries from one browser per hour'],
  ['spam_max_per_hour', '40', 'Max enquiries site-wide per hour'],
  ['spam_dupe_minutes', '20', 'Same mobile inside this window is treated as the same enquiry'],
  ['spam_max_links', '1', 'Links allowed inside the message'],
  ['spam_blocklist', 'seo service,backlink,rank your website,guest post,crypto,bitcoin,casino,viagra,escort,forex,porn,web design service,increase traffic', 'Comma-separated words that block an enquiry']
];

/* ------------------------------------------------------------------ setup */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_LEADS, LEAD_COLUMNS);
  ensureSheet(ss, SHEET_ACTIVITY, ACTIVITY_COLUMNS);
  ensureConfig(ss);
  formatLeads(ss.getSheetByName(SHEET_LEADS));

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  }
  if (!props.getProperty('ALERT_EMAIL')) {
    props.setProperty('ALERT_EMAIL', 'hello@ageofaquariusdevelopers.in');
  }
  installDigestTrigger();

  Logger.log('TOKEN: ' + props.getProperty('TOKEN'));
  Logger.log('Alert email: ' + props.getProperty('ALERT_EMAIL'));
  Logger.log('Now: Deploy > New deployment > Web app (Execute as Me, Anyone with the link).');
}

function ensureSheet(ss, name, columns) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, columns.length).setValues([columns]);
    SpreadsheetApp.flush();
  } else if (sh.getLastColumn() < columns.length) {
    // A later version of this script added columns — repair the header row.
    sh.getRange(1, 1, 1, columns.length).setValues([columns]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, columns.length).setFontWeight('bold');
  return sh;
}

function ensureConfig(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONFIG);
    sh.getRange(1, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  } else {
    // Top up any keys added in a later version of this script.
    var have = {};
    sh.getDataRange().getValues().forEach(function (r) { if (r[0]) have[String(r[0])] = true; });
    var add = DEFAULT_CONFIG.slice(1).filter(function (r) { return !have[r[0]]; });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  }
  return sh;
}

function formatLeads(sh) {
  var statusCol = LEAD_COLUMNS.indexOf('status') + 1;
  var range = sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Qualified', 'Site visit', 'Demo', 'Proposal', 'Won', 'Lost', 'Nurture'], true)
    .setAllowInvalid(false).build();
  range.setDataValidation(rule);
}

function config() {
  var sh = ensureConfig(SpreadsheetApp.getActiveSpreadsheet());
  var values = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) out[String(values[i][0])] = values[i][1];
  }
  return out;
}

function num(cfg, key, fallback) {
  var v = Number(cfg[key]);
  return isNaN(v) ? fallback : v;
}

/* ----------------------------------------------------------------- web app */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var props = PropertiesService.getScriptProperties();
    if (body.token !== props.getProperty('TOKEN')) return json({ ok: false, error: 'bad token' });

    if (body.action === 'lead') return json(createLead(body.payload || {}));
    if (body.action === 'ping') return json({ ok: true, data: { pong: true } });
    return json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, data: { service: 'aoa-leads' } });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* -------------------------------------------------------------- lead intake */

function createLead(p) {
  var cfg = config();
  var mobile = normaliseMobile(p.mobile);
  if (!p.name || !mobile) return { ok: false, error: 'name and mobile are required' };

  var gate = spamCheck(p, mobile, cfg);
  if (gate.block) {
    logBlocked(p, gate.reason, mobile);
    if (gate.silent) return { ok: true, data: { lead_id: 'ignored' } };
    return { ok: false, error: gate.retry
      ? 'Too many enquiries just now. Please try again in a few minutes, or WhatsApp us.'
      : 'We could not accept this enquiry. Please WhatsApp us on +91 89565 52480.' };
  }

  // Same mobile inside the dedupe window: hand back the first reference, no second row.
  var cache = CacheService.getScriptCache();
  var dupeKey = 'aoa_m_' + mobile.replace(/\D/g, '');
  var prior = cache.get(dupeKey);
  if (prior) return { ok: true, data: { lead_id: prior, duplicate: true } };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet(ss, SHEET_LEADS, LEAD_COLUMNS);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var now = new Date();
    var leadId = nextLeadId(sh, now);
    var scored = scoreLead(p, cfg, mobile);
    var slaHours = scored.band === 'Hot' ? num(cfg, 'sla_hot_hours', 2)
      : scored.band === 'Warm' ? num(cfg, 'sla_warm_hours', 24)
        : num(cfg, 'sla_cold_hours', 72);

    var row = {
      lead_id: leadId,
      created_at: now,
      lang: p.lang || 'en',
      track: p.track || '',
      name: p.name,
      mobile: mobile,
      whatsapp_ok: p.whatsapp_ok ? 'yes' : 'no',
      email: p.email || '',
      company: p.company || '',
      role: p.role || '',
      location: p.location || '',
      area_value: p.area_value || '',
      area_unit: p.area_unit || '',
      need: p.need || '',
      timeline: p.timeline || '',
      budget_band: p.budget_band || '',
      org_type: p.org_type || '',
      sites: p.sites || '',
      users: p.users || '',
      stage: p.stage || '',
      notes: p.notes || '',
      consent: p.consent ? 'yes' : 'no',
      score: scored.score,
      band: scored.band,
      status: 'New',
      internal_notes: gate.warn || '',
      owner: cfg.owner || 'Rahul',
      next_action_at: new Date(now.getTime() + slaHours * 3600 * 1000),
      page: p.page || '',
      referrer: p.referrer || '',
      utm_source: p.utm_source || '',
      utm_medium: p.utm_medium || '',
      utm_campaign: p.utm_campaign || '',
      utm_term: p.utm_term || '',
      utm_content: p.utm_content || '',
      gclid: p.gclid || '',
      device: p.device || '',
      fill_seconds: p.fill_seconds || '',
      wa_status: ''
    };

    var values = LEAD_COLUMNS.map(function (c) {
      var v = row[c] === undefined ? '' : row[c];
      // Keep '+91 ...' and anything formula-like as plain text in the sheet.
      if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
      return v;
    });
    sh.appendRow(values);
    var rowIndex = sh.getLastRow();

    logActivity(leadId, 'system', 'form', 'Lead captured from ' + (row.page || 'website') + ' (' + row.track + ', ' + row.band + ')', 'First response by ' + row.next_action_at);

    var waStatus = notify(row, cfg, rowIndex);
    if (waStatus) {
      sh.getRange(rowIndex, LEAD_COLUMNS.indexOf('wa_status') + 1).setValue(waStatus);
    }

    cache.put(dupeKey, leadId, Math.max(60, num(cfg, 'spam_dupe_minutes', 20) * 60));
    if (gate.bump) gate.bump();

    return { ok: true, data: { lead_id: leadId, band: scored.band } };
  } finally {
    lock.releaseLock();
  }
}

function nextLeadId(sh, now) {
  var stamp = Utilities.formatDate(now, 'Asia/Kolkata', 'yyMMdd');
  var prefix = 'AOA-' + stamp + '-';
  var last = sh.getLastRow();
  var count = 0;
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).indexOf(prefix) === 0) count++;
    }
  }
  return prefix + ('0' + (count + 1)).slice(-2);
}

/* ----------------------------------------------------------------- anti-spam */

/**
 * Layered gate: hidden fields, timing, content, and rate limits.
 * Returns { block, silent, retry, reason, bump } — silent blocks look like a
 * success to the sender (bots learn nothing); the rest come back as an error.
 */
function spamCheck(p, mobile, cfg) {
  var name = String(p.name || '').trim();
  var notes = String(p.notes || '');

  var secs = Number(p.fill_seconds);
  var warn = '';

  // 1. Honeypot — a hidden field no human can reach. Browser autofill does
  // sometimes fill hidden fields anyway, so a trip only blocks when the form was
  // also filled inhumanly fast; otherwise the lead is kept and flagged.
  if (p.lf_ref2 || p.company_website || p.fax_number) {
    if (!secs || secs < 20) return { block: true, silent: true, reason: 'honeypot' };
    warn = 'Hidden field was filled (browser autofill?) — kept for review';
  }

  // 2. Timing — a real person cannot complete three steps in a few seconds.
  if (secs && secs < num(cfg, 'spam_min_seconds', 5)) {
    return { block: true, silent: true, reason: 'submitted in ' + secs + 's' };
  }

  // 3. Name sanity.
  if (name.length < 2 || name.length > 80) return { block: true, reason: 'name length' };
  if (!/[A-Za-z\u0900-\u097F]/.test(name)) return { block: true, reason: 'name has no letters' };
  if (/https?:\/\/|www\.|<[a-z]/i.test(name)) return { block: true, reason: 'markup or link in name' };

  // 4. Message content.
  if (notes.length > 1500) return { block: true, reason: 'message too long' };
  var links = (notes.match(/https?:\/\/|www\./gi) || []).length;
  if (links > num(cfg, 'spam_max_links', 1)) return { block: true, reason: links + ' links in message' };
  if (/[\u0400-\u04FF\u4E00-\u9FFF]/.test(name + ' ' + notes)) return { block: true, reason: 'unexpected script' };

  // 5. Word blocklist (Config sheet, editable).
  var hay = (name + ' ' + notes + ' ' + (p.email || '') + ' ' + (p.location || '')).toLowerCase();
  var words = String(cfg['spam_blocklist'] || '').split(',');
  for (var i = 0; i < words.length; i++) {
    var w = words[i].trim().toLowerCase();
    if (w && hay.indexOf(w) !== -1) return { block: true, reason: 'blocked word: ' + w };
  }

  // 6. Email, when given, has to look like one.
  if (p.email && !/^[^@\s]+@[^@\s.]+\.[^@\s]{2,}$/.test(String(p.email).trim())) {
    return { block: true, reason: 'email looks invalid' };
  }

  // 7. Placeholder mobile numbers.
  var d = String(mobile).replace(/\D/g, '').slice(-10);
  if (/^(\d)\1{9}$/.test(d) || d === '9876543210' || d === '1234567890') {
    return { block: true, reason: 'placeholder mobile' };
  }

  // 8. Rate limits — per browser and site-wide, one hour each.
  var cache = CacheService.getScriptCache();
  var cid = String(p.client_id || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 40);
  var ckey = cid ? 'aoa_c_' + cid : '';
  var gkey = 'aoa_g_' + Math.floor(new Date().getTime() / 3600000);
  var cCount = ckey ? Number(cache.get(ckey) || 0) : 0;
  var gCount = Number(cache.get(gkey) || 0);
  if (ckey && cCount >= num(cfg, 'spam_max_per_client_hour', 3)) {
    return { block: true, reason: 'browser rate limit (' + cCount + '/hour)' };
  }
  if (gCount >= num(cfg, 'spam_max_per_hour', 40)) {
    return { block: true, retry: true, reason: 'site rate limit (' + gCount + '/hour)' };
  }

  return {
    block: false,
    warn: warn,
    bump: function () {
      if (ckey) cache.put(ckey, String(cCount + 1), 3600);
      cache.put(gkey, String(gCount + 1), 3600);
    }
  };
}

/** Blocked attempts are logged, never lost — check this tab for false positives. */
function logBlocked(p, reason, mobile) {
  try {
    var sh = ensureSheet(SpreadsheetApp.getActiveSpreadsheet(), SHEET_BLOCKED, BLOCKED_COLUMNS);
    sh.appendRow([
      new Date(), reason || '', String(p.name || '').slice(0, 120), "'" + (mobile || String(p.mobile || '')),
      String(p.email || '').slice(0, 120), p.track || '', p.lang || '', String(p.notes || '').slice(0, 500),
      p.client_id || '', p.page || '', p.referrer || '', p.device || '', p.fill_seconds || '',
      JSON.stringify(p).slice(0, 5000)
    ]);
  } catch (err) {
    Logger.log('logBlocked failed: ' + err);
  }
}

function normaliseMobile(raw) {
  var digits = String(raw || '').replace(/\D/g, '');
  if (digits.length === 12 && digits.indexOf('91') === 0) digits = digits.slice(2);
  if (digits.length === 11 && digits.charAt(0) === '0') digits = digits.slice(1);
  if (!/^[6-9]\d{9}$/.test(digits)) return '';
  return '+91 ' + digits;
}

/* ------------------------------------------------------------------ scoring */

function scoreLead(p, cfg, mobile) {
  var s = 0;
  var track = p.track || '';
  if (track === 'package') s += num(cfg, 'w_track_package', 25);
  else if (track === 'land') s += num(cfg, 'w_track_land', 20);
  else if (track === 'platform') s += num(cfg, 'w_track_platform', 20);

  var acres = toAcres(p.area_value, p.area_unit);
  if (acres > 10) s += num(cfg, 'w_area_large', 20);
  else if (acres >= 2) s += num(cfg, 'w_area_mid', 15);
  else if (acres > 0) s += num(cfg, 'w_area_small', 8);
  else if (track === 'platform') s += num(cfg, 'w_area_mid', 15); // platform leads have no land

  var t = String(p.timeline || '');
  if (t === '0_3') s += num(cfg, 'w_time_0_3', 20);
  else if (t === '3_6') s += num(cfg, 'w_time_3_6', 12);
  else if (t) s += num(cfg, 'w_time_6_plus', 5);

  var role = String(p.role || p.org_type || '');
  if (role === 'owner' || role === 'decision_maker' || role === 'developer') s += num(cfg, 'w_role_owner', 15);

  var filled = ['name', 'mobile', 'email', 'location', 'timeline', 'notes'].filter(function (k) { return p[k]; }).length;
  if (filled >= 5) s += num(cfg, 'w_completeness', 10);
  if (mobile) s += num(cfg, 'w_contactable', 10);

  s = Math.max(0, Math.min(100, Math.round(s)));
  var band = s >= num(cfg, 'band_hot_min', 70) ? 'Hot' : s >= num(cfg, 'band_warm_min', 40) ? 'Warm' : 'Cold';
  return { score: s, band: band };
}

function toAcres(value, unit) {
  var v = Number(value);
  if (!v || isNaN(v)) return 0;
  if (unit === 'ha' || unit === 'hectare') return v * 2.4711;
  if (unit === 'guntha') return v / 40;
  if (unit === 'sqft') return v / 43560;
  return v; // acres
}

/* ------------------------------------------------------------ notifications */

function notify(row, cfg, rowIndex) {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('ALERT_EMAIL');
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' +
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS).getSheetId() + '&range=A' + rowIndex;

  if (email) {
    var subject = '[' + row.band + '] New ' + trackLabel(row.track) + ' lead — ' + row.name;
    var lines = [
      row.band + ' lead · score ' + row.score + ' · ' + row.lead_id,
      '',
      'Name: ' + row.name,
      'Mobile: ' + row.mobile + (row.whatsapp_ok === 'yes' ? ' (WhatsApp ok)' : ''),
      'Email: ' + (row.email || '—'),
      'Track: ' + trackLabel(row.track),
      'Language: ' + row.lang,
      'Location: ' + (row.location || '—'),
      'Area: ' + (row.area_value ? row.area_value + ' ' + row.area_unit : '—'),
      'Need: ' + (row.need || row.stage || '—'),
      'Timeline: ' + (row.timeline || '—'),
      'Budget: ' + (row.budget_band || 'not shared'),
      'Notes: ' + (row.notes || '—'),
      '',
      'Source: ' + (row.utm_source || 'direct') + ' / ' + (row.utm_medium || '—') + ' · page ' + (row.page || '—'),
      'Respond by: ' + row.next_action_at,
      '',
      'Call: tel:' + row.mobile.replace(/\s/g, ''),
      'WhatsApp: https://wa.me/' + row.mobile.replace(/\D/g, ''),
      'Sheet row: ' + sheetUrl
    ];
    MailApp.sendEmail({ to: email, subject: subject, body: lines.join('\n') });
  }

  return sendWhatsApp(row);
}

function trackLabel(track) {
  return track === 'land' ? 'Land Development'
    : track === 'platform' ? 'Technology Platform'
      : track === 'package' ? 'Full Package' : 'General';
}

/**
 * Phase 2 — Meta WhatsApp Cloud API. Stays dormant until WA_TOKEN,
 * WA_PHONE_ID and WA_TO are set in Script Properties, so phase 1 runs
 * happily on email alone.
 */
function sendWhatsApp(row) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('WA_TOKEN');
  var phoneId = props.getProperty('WA_PHONE_ID');
  var to = props.getProperty('WA_TO');
  if (!token || !phoneId || !to) return 'skipped (not configured)';

  var alert = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'template',
    template: {
      name: 'aoa_lead_alert',
      language: { code: 'en' },
      components: [{
        type: 'body',
        parameters: [
          { type: 'text', text: row.name },
          { type: 'text', text: trackLabel(row.track) },
          { type: 'text', text: (row.location || '—') + (row.area_value ? ' · ' + row.area_value + ' ' + row.area_unit : '') },
          { type: 'text', text: row.timeline || '—' },
          { type: 'text', text: row.band + ' (' + row.score + ')' },
          { type: 'text', text: row.lead_id }
        ]
      }]
    }
  };

  var status = post('https://graph.facebook.com/v20.0/' + phoneId + '/messages', token, alert);

  if (row.whatsapp_ok === 'yes') {
    var ack = {
      messaging_product: 'whatsapp',
      to: row.mobile.replace(/\D/g, ''),
      type: 'template',
      template: {
        name: 'aoa_lead_ack',
        language: { code: row.lang === 'mr' ? 'mr' : row.lang === 'hi' ? 'hi' : 'en' },
        components: [{ type: 'body', parameters: [{ type: 'text', text: row.name }] }]
      }
    };
    status += ' | ack ' + post('https://graph.facebook.com/v20.0/' + phoneId + '/messages', token, ack);
  }
  return status;
}

function post(url, token, payload) {
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return res.getResponseCode() === 200 ? 'sent' : 'failed ' + res.getResponseCode();
  } catch (err) {
    return 'failed ' + err;
  }
}

/* -------------------------------------------------------------- activity log */

function logActivity(leadId, actor, type, summary, nextStep) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet(ss, SHEET_ACTIVITY, ACTIVITY_COLUMNS);
  sh.appendRow([new Date(), leadId, actor, type, summary, nextStep || '']);
}

/* ---------------------------------------------------------------- reporting */

function installDigestTrigger() {
  var cfg = config();
  var hour = num(cfg, 'digest_hour', 9);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(hour).everyDays(1).inTimezone('Asia/Kolkata').create();
}

function dailyDigest() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('ALERT_EMAIL');
  if (!email) return;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS);
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, LEAD_COLUMNS.length).getValues();
  var idx = {};
  LEAD_COLUMNS.forEach(function (c, i) { idx[c] = i; });

  var now = new Date();
  var since = new Date(now.getTime() - 24 * 3600 * 1000);
  var fresh = [], overdue = [];
  rows.forEach(function (r) {
    var created = r[idx.created_at] instanceof Date ? r[idx.created_at] : new Date(r[idx.created_at]);
    var status = String(r[idx.status] || '');
    var due = r[idx.next_action_at] instanceof Date ? r[idx.next_action_at] : null;
    if (created > since) fresh.push(r);
    if (due && due < now && ['New', 'Contacted', 'Qualified'].indexOf(status) > -1) overdue.push(r);
  });

  var line = function (r) {
    return '· ' + r[idx.lead_id] + ' — ' + r[idx.name] + ' (' + trackLabel(r[idx.track]) + ', ' + r[idx.band] + ') ' + r[idx.mobile];
  };
  var body = [
    'Leads in the last 24 hours: ' + fresh.length,
    fresh.map(line).join('\n') || '  none',
    '',
    'Overdue follow-ups: ' + overdue.length,
    overdue.map(line).join('\n') || '  none',
    '',
    SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].join('\n');

  MailApp.sendEmail({ to: email, subject: 'Age of Aquarius — leads digest (' + fresh.length + ' new, ' + overdue.length + ' overdue)', body: body });
}

/* --------------------------------------------------------------- self-test */

function testLead() {
  var out = createLead({
    lang: 'mr', track: 'land', name: 'Test Owner', mobile: '9890307622', whatsapp_ok: true,
    email: 'test@example.com', role: 'owner', location: 'Saswad, Purandar', area_value: '12',
    area_unit: 'acre', need: 'development', timeline: '0_3', notes: 'Test lead from testLead()',
    consent: true, page: '/', device: 'test', fill_seconds: 42
  });
  Logger.log(JSON.stringify(out));
}
