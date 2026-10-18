function getReviewContactId_() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('REVIEW_CONTACT_ID');
  if (cached) return cached;

  const locationId = props.getProperty('GHL_LOCATION_ID');
  const owner = String(getConfig_('OWNER_EMAIL'));
  const parts = owner.split('@');
  const email = parts[0].split('+')[0] + '+ardatareview@' + parts[1]; // unique alias, hindi tatama sa ibang contact

  const res = ghl_('POST', '/contacts/upsert', { locationId: locationId, firstName: 'AR Data Review', email: email });
  props.setProperty('REVIEW_CONTACT_ID', res.contact.id);
  return res.contact.id;
}

function flagDataIssues_() {
  const props = PropertiesService.getScriptProperties();
  const data = SpreadsheetApp.getActive().getSheetByName('ExceptionLog').getDataRange().getValues();
  const invoicesById = getInvoicesById_();
  const newIssues = [];

  for (let i = 1; i < data.length; i++) {
    const invoiceId = data[i][1];
    const codes = String(data[i][2] || '');
    if (!invoiceId || codes.indexOf('NO_CONTACT_CHANNEL') === -1) continue;

    const alertKey = 'DATA_ISSUE_ALERTED_' + invoiceId;
    if (props.getProperty(alertKey)) continue; // isang beses lang kada issue

    const inv = invoicesById[invoiceId];
    newIssues.push({ invoiceId: invoiceId, clientId: inv ? inv.client_id : invoiceId, alertKey: alertKey });
  }
  if (newIssues.length === 0) return;

  const reviewId = getReviewContactId_();
  const summary = newIssues.map(x => 'NO_CONTACT_CHANNEL: ' + x.invoiceId + ' (' + x.clientId + ')').join('; ');

  ghlUpdateContactFields_(reviewId, { ar_relationship_notes: summary });

  // Tanggalin muna ang tag bago idagdag ulit, para laging mag-fire ang "Tag Added" kahit may luma nang tag
  try { ghl_('DELETE', '/contacts/' + reviewId + '/tags', { tags: ['ar-data-issue'] }); } catch (e) { }
  ghl_('POST', '/contacts/' + reviewId + '/tags', { tags: ['ar-data-issue'] });

  newIssues.forEach(x => props.setProperty(x.alertKey, new Date().toISOString()));
  logSystemException_('DATA_ISSUE', 'DATA_ISSUE_ALERT_SENT', summary);
}

function reconcileCounts_() {
  const props = PropertiesService.getScriptProperties();
  const pipelineId = props.getProperty('PIPELINE_ID');
  const locationId = props.getProperty('GHL_LOCATION_ID');

  // Sheet side: galing sa Dashboard (Computed ang source)
  const dash = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  const sheetCount = Number(dash.getRange('B8').getValue()) || 0;
  const sheetValue = Number(dash.getRange('B5').getValue()) || 0;

  // GHL side: open opportunities sa AR Collections pipeline (may pagination)
  let ghlCount = 0, ghlValue = 0, page = 1, guard = 0;
  while (guard++ < 20) {
    const res = ghl_('GET', '/opportunities/search?location_id=' + locationId +
      '&pipeline_id=' + pipelineId + '&status=open&limit=100&page=' + page);
    const opps = (res && res.opportunities) || [];
    opps.forEach(o => { ghlCount++; ghlValue += Number(o.monetaryValue) || 0; });
    if (opps.length < 100) break;
    page++;
  }

  const ok = (sheetCount === ghlCount) && (Math.abs(sheetValue - ghlValue) < 1);
  const status = ok ? 'OK' : 'MISMATCH';

  dash.getRange('A41:B41').setValues([['Reconciliation', status]]);
  dash.getRange('A42:B42').setValues([['Sheet: open count / value', sheetCount + ' / ' + sheetValue]]);
  dash.getRange('A43:B43').setValues([['GHL: open count / value', ghlCount + ' / ' + ghlValue]]);

  if (!ok) {
    logSystemException_('RECONCILE', 'COUNT_MISMATCH',
      'Sheet ' + sheetCount + '/' + sheetValue + ' vs GHL ' + ghlCount + '/' + ghlValue);
  }
  return { ok: ok, sheetCount: sheetCount, sheetValue: sheetValue, ghlCount: ghlCount, ghlValue: ghlValue };
}

function stampLastSynced_() {
  const dash = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  const tz = Session.getScriptTimeZone();
  dash.getRange('A40:B40').setValues([['Last synced at', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')]]);
}

function sendWeeklyDigest_() {
  const dash = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  const ownerEmail = getConfig_('OWNER_EMAIL');
  const peso = n => 'PHP ' + Number(n || 0).toLocaleString('en-US');

  // As-of date: kayang magbasa ng Date, serial number, o text
  const asOfRaw = dash.getRange('B2').getValue();
  let asOf;
  if (asOfRaw instanceof Date) {
    asOf = Utilities.formatDate(asOfRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (typeof asOfRaw === 'number') {
    asOf = Utilities.formatDate(new Date(Math.round((asOfRaw - 25569) * 86400000)), 'UTC', 'yyyy-MM-dd');
  } else {
    asOf = String(asOfRaw);
  }

  const outstanding = dash.getRange('B5').getValue();
  const overdue     = dash.getRange('B6').getValue();
  const overduePct  = dash.getRange('B7').getDisplayValue();
  const openCount   = dash.getRange('B8').getValue();
  const overdueCnt  = dash.getRange('B9').getValue();
  const collected   = dash.getRange('B10').getValue();
  const needHuman   = dash.getRange('B11').getValue();
  const dso         = dash.getRange('B38').getDisplayValue();

  // A13:C17 = limang buckets lang (walang Total row). Columns: bucket | amount | count
  const aging = dash.getRange('A13:C17').getDisplayValues()
    .map(r => '  ' + r[0] + ': ' + r[2] + ' invoices, PHP ' + Number(r[1]).toLocaleString('en-US')).join('\n');

  const recon  = dash.getRange('B41').getValue();
  const synced = dash.getRange('B40').getDisplayValue();

  const body =
    'WEEKLY AR DIGEST (as of ' + asOf + ')\n\n' +
    'Total outstanding: ' + peso(outstanding) + ' (' + openCount + ' open invoices)\n' +
    'Overdue: ' + peso(overdue) + ' (' + overdueCnt + ' invoices, ' + overduePct + ')\n' +
    'Collected last 7 days: ' + peso(collected) + '\n' +
    'Waiting for human action: ' + needHuman + '\n' +
    'DSO (approx): ' + dso + ' days\n\n' +
    'AGING:\n' + aging + '\n\n' +
    'Data check: reconciliation ' + recon + ' | last synced ' + synced;

  MailApp.sendEmail({ to: ownerEmail, subject: 'Weekly AR Digest - ' + asOf, body: body });
  logSystemException_('DIGEST', 'WEEKLY_DIGEST_SENT', asOf);
}

function installWeeklyDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyDigest_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyDigest_')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).nearMinute(0).create();
}

function snapshotAiCost() {
  const ss = SpreadsheetApp.getActive();
  let out = ss.getSheetByName('CostAudit_Raw');
  if (!out) out = ss.insertSheet('CostAudit_Raw');
  out.clear();

  // 1) Daily counters mula sa Script Properties (AI_COUNT_yyyyMMdd, CACHE_HIT_yyyyMMdd)
  const props = PropertiesService.getScriptProperties().getProperties();
  const byDay = {};
  Object.keys(props).forEach(k => {
    const m = k.match(/^(AI_COUNT|CACHE_HIT)_(\d{8})$/);
    if (!m) return;
    byDay[m[2]] = byDay[m[2]] || { ai: 0, cache: 0 };
    if (m[1] === 'AI_COUNT') byDay[m[2]].ai = Number(props[k]) || 0;
    else byDay[m[2]].cache = Number(props[k]) || 0;
  });

  const days = Object.keys(byDay).sort();
  const rows = [['day', 'ai_calls', 'cache_hits']];
  let totalAi = 0, totalCache = 0;
  days.forEach(d => {
    rows.push([d, byDay[d].ai, byDay[d].cache]);
    totalAi += byDay[d].ai; totalCache += byDay[d].cache;
  });
  rows.push(['TOTAL', totalAi, totalCache]);
  out.getRange(1, 1, rows.length, 3).setValues(rows);

  // 2) Bilang ng source (AI / FALLBACK / TEMPLATE) sa CommLog, column D
  const countSources = (sheetName, colIndex) => {
    const s = ss.getSheetByName(sheetName);
    const counts = {};
    if (!s) return counts;
    s.getDataRange().getValues().slice(1).forEach(r => {
      const src = String(r[colIndex] || '').trim();
      if (src) counts[src] = (counts[src] || 0) + 1;
    });
    return counts;
  };

  const comm = countSources('CommLog', 3);
  const startRow = rows.length + 2;
  const commRows = [['CommLog source', 'count']].concat(Object.keys(comm).map(k => [k, comm[k]]));
  out.getRange(startRow, 1, commRows.length, 2).setValues(commRows);

  Logger.log('Days: ' + days.length + ' | Total AI calls: ' + totalAi + ' | Cache hits: ' + totalCache);
  Logger.log('CommLog sources: ' + JSON.stringify(comm));
}

function resetSheets_() {
  const ss = SpreadsheetApp.getActive();
  // HUWAG isama: Invoices, Clients, Config, Computed, Dashboard, MessageCache, TestLog, CostAudit_Raw
  ['DunningState', 'DunningPlan', 'CommLog', 'SyncLog', 'ExceptionLog'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) { Logger.log('Wala: ' + name); return; }
    const last = s.getLastRow();
    if (last > 1) s.getRange(2, 1, last - 1, s.getLastColumn()).clearContent();
    Logger.log('Nalinis: ' + name + ' (' + Math.max(last - 1, 0) + ' rows)');
  });
}

function listScriptPropertyKeys_() {
  // Mga pangalan lang ang ila-log, hindi ang values (may secrets)
  Logger.log(Object.keys(PropertiesService.getScriptProperties().getProperties()).sort().join('\n'));
}

function resetAlertKeys() {
  const p = PropertiesService.getScriptProperties();
const prefixes = ['PLAN_LIMBO_ALERTED_', 'WATCHDOG_ALERTED_', 'DATA_ISSUE_ALERTED_',
                  'SEND_UNCONFIRMED_ALERTED_', 'AI_FALLBACK_STREAK_'];  let n = 0;
  Object.keys(p.getProperties()).forEach(k => {
    if (prefixes.some(x => k.indexOf(x) === 0)) { p.deleteProperty(k); n++; }
  });
  Logger.log('Binura ang ' + n + ' alert keys');
}
function createFinalTestLog_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('TestLog_Final');
  if (!s) s = ss.insertSheet('TestLog_Final');
  // Safety: huwag burahin ang naitalang resulta (23/23 evidence)
  if (s.getLastRow() > 1) throw new Error('TestLog_Final may laman na -- hindi ito buburahin. Burahin muna nang mano-mano kung talagang sinadya.');
  s.clear();
  s.getRange(1, 1, 1, 7).setValues([['test_id', 'date', 'setup', 'expected', 'actual', 'pass_fail', 'fix_notes']]);
  const tests = [
    ['M1', 'Invoice paid during sequence', 'Chasing stops'],
    ['M2', 'Client requests payment plan', 'Recognized and routed'],
    ['M3', 'Unresponsive through all stages', 'Handoff with context'],
    ['M4', 'Very small invoice', 'Proportionate tone/logic'],
    ['M5', 'Very large invoice', 'Human intervention'],
    ['A1', 'Not yet overdue', 'Not Yet Due, no message'],
    ['A2', 'Becomes overdue', 'Moves to S1 at right day'],
    ['A3', 'Already overdue when imported', 'Backlog handoff, no S1-S4 burst'],
    ['A4', 'Invoice status missing', 'Exception logged, not synced'],
    ['A5', 'Client contact info missing', 'NO_CONTACT_CHANNEL, data issue alert'],
    ['A6', 'Duplicate invoice record', 'One opportunity only'],
    ['A7', 'Duplicate workflow trigger', 'No double send'],
    ['A8', 'Invalid invoice amount', 'Exception logged, not synced'],
    ['A9', 'Invalid due date', 'Exception logged, not synced'],
    ['A10', 'Client with multiple outstanding invoices', '1 contact, 2 opportunities, one message per tick'],
    ['A11', 'Pays one, still owes another', 'Only paid one stops'],
    ['A12', 'Client replies with unrelated info', 'Classified other, routed to human'],
    ['A13', "Client asks why they're being contacted", 'Routed to human, no auto-argue'],
    ['A14', 'AI-generated message fails/unusable', 'Fallback template, logged'],
    ['A15', 'Gemini API unavailable', 'Fallback, no crash'],
    ['A16', 'AI response exceeds expected length', 'Rejected, fallback'],
    ['A17', 'Human escalation notification fails', 'Gmail watchdog alert'],
    ['A18', 'Payment status update delayed', 'Dashboard shows last synced at']
  ];
  s.getRange(2, 1, tests.length, 3).setValues(tests.map(t => [t[0], '', t[1]]));
  s.getRange(2, 4, tests.length, 1).setValues(tests.map(t => [t[2]]));
  s.setFrozenRows(1);
}