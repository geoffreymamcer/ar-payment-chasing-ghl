
function ghl_(method, path, body) {
  const p = PropertiesService.getScriptProperties();
  const opts = {
    method, muteHttpExceptions: true, contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + p.getProperty('GHL_TOKEN'),
      Version: '2021-07-28',
      Accept: 'application/json'
    }
  };
  if (body) opts.payload = JSON.stringify(body);

  for (let i = 0; i < 3; i++) {
    const res = UrlFetchApp.fetch('https://services.leadconnectorhq.com' + path, opts);
    const code = res.getResponseCode();
    if (code === 429 || code >= 500) { Utilities.sleep(1000 * (i + 1)); continue; }
    if (code >= 200 && code < 300) return JSON.parse(res.getContentText() || '{}');
    throw new Error('GHL ' + code + ': ' + res.getContentText());
  }
  throw new Error('GHL_RETRIES_EXHAUSTED');
}

function bootstrapIds() {
  const locationId = PropertiesService.getScriptProperties().getProperty('GHL_LOCATION_ID');

  // (a) Pipelines + stages
  const pipelineRes = ghl_('GET', '/opportunities/pipelines?locationId=' + locationId);
  const pipeline = pipelineRes.pipelines.find(p => p.name === 'AR Collections');
  if (!pipeline) throw new Error('PIPELINE_NOT_FOUND: AR Collections');

  const stageMap = {};
  pipeline.stages.forEach(stage => { stageMap[stage.name] = stage.id; });

  // (b) Custom field IDs (contact + opportunity, lahat sabay)
  const fieldRes = ghl_('GET', '/locations/' + locationId + '/customFields?model=all');
  const fieldMap = {};
  fieldRes.customFields.forEach(f => {
    // fieldKey ay hugis "contact.ar_client_id" o "opportunity.ar_invoice_id"
    // -- kukunin lang natin ang bahagi PAGKATAPOS ng tuldok
    const parts = f.fieldKey.split('.');
    const shortKey = parts[parts.length - 1];
    fieldMap[shortKey] = f.id;
  });

  const props = PropertiesService.getScriptProperties();
  props.setProperty('PIPELINE_ID', pipeline.id);
  props.setProperty('STAGE_MAP', JSON.stringify(stageMap));
  props.setProperty('FIELD_MAP', JSON.stringify(fieldMap));

  Logger.log('Pipeline ID: ' + pipeline.id);
  Logger.log('Stage Map (' + Object.keys(stageMap).length + ' stages): ' + JSON.stringify(stageMap, null, 2));
  Logger.log('Field Map (' + Object.keys(fieldMap).length + ' fields): ' + JSON.stringify(fieldMap, null, 2));
}

function upsertContact_(client) {
  if (!client.contact_email && !client.contact_phone) {
    logSystemException_(client.client_id, 'NO_CONTACT_CHANNEL', 'N/A');
    return null;
  }

  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const locationId = PropertiesService.getScriptProperties().getProperty('GHL_LOCATION_ID');

  const payload = {
    locationId: locationId,
    firstName: client.contact_first_name,
    email: client.contact_email || undefined,
    phone: client.contact_phone || undefined,
    customFields: [
      { id: fieldMap['ar_client_id'], field_value: client.client_id },
      { id: fieldMap['ar_client_tier'], field_value: client.client_tier },
      { id: fieldMap['ar_payment_behavior'], field_value: client.payment_behavior },
      { id: fieldMap['ar_relationship_notes'], field_value: client.relationship_notes || '' }
    ]
  };

  const res = ghl_('POST', '/contacts/upsert', payload);
  return res.contact.id;
}

function upsertOpportunity_(inv, contactId) {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const pipelineId = PropertiesService.getScriptProperties().getProperty('PIPELINE_ID');
  const locationId = PropertiesService.getScriptProperties().getProperty('GHL_LOCATION_ID');

  // Huwag i-sync ang DRAFT invoices
  if (inv.invoice_status === 'DRAFT') {
    logSystemException_(inv.invoice_id, 'SKIPPED_DRAFT', 'N/A');
    return { action: 'SKIPPED', detail: 'DRAFT status' };
  }

  // (a) Hanapin muna ang existing opportunity via ar_invoice_id
  const searchRes = ghl_('GET', '/opportunities/search?location_id=' + locationId +
    '&pipeline_id=' + pipelineId + '&contact_id=' + contactId);

  const existing = (searchRes.opportunities || []).find(opp =>
    getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_id']) === inv.invoice_id
  );

  const stageDecision = determineStage_(inv, stageMap);
  const stageId = stageDecision.stageId;
  const customFieldsPayload = buildOpportunityCustomFields_(inv, fieldMap, stageDecision.holdReason);
  const opportunityName = inv.invoice_number + ' · ' + inv.client_name + ' · PHP ' + formatPeso_(inv.balance_due);

  // Sync ay may-ari LANG ng pre-dunning stages (Not Yet Due, Overdue – Grace).
  // S1-S4 / Handoff / Plan / Dispute = dunning engine at reply handler.
  // Paid / Voided = reconcilePayments.
  const syncOwned = [stageMap['Not Yet Due'], stageMap['Overdue – Grace']].filter(Boolean);
  const isTerminal = [stageMap['Paid'], stageMap['Voided / Closed']].includes(stageId);

  let status = 'open';
  if (inv.invoice_status === 'PAID') status = 'won';
  if (inv.invoice_status === 'VOIDED') status = 'lost';

  if (existing) {
    // (b) Meron na -> update ng data; stage ay ginagalaw LANG kung sync-owned ang kasalukuyan AT ang target
    const updatePayload = {
      name: opportunityName,
      status: status,
      monetaryValue: inv.balance_due,
      customFields: customFieldsPayload
    };
    if (syncOwned.includes(existing.pipelineStageId) && syncOwned.includes(stageId)) {
      updatePayload.pipelineStageId = stageId;
    }
    ghl_('PUT', '/opportunities/' + existing.id, updatePayload);
    return { action: 'UPDATED', opportunityId: existing.id };
  } else {
    // (c) Wala pa -> gumawa ng bago. Kung overdue na, magsimula sa Grace;
    // ang dunning engine (parehong runSync) ang maglilipat sa tamang S-stage o Handoff.
    const createStageId = (isTerminal || syncOwned.includes(stageId)) ? stageId : stageMap['Overdue – Grace'];
    const createRes = ghl_('POST', '/opportunities/', {
      locationId: locationId,
      pipelineId: pipelineId,
      contactId: contactId,
      name: opportunityName,
      pipelineStageId: createStageId,
      status: status,
      monetaryValue: inv.balance_due,
      customFields: customFieldsPayload
    });
    return { action: 'CREATED', opportunityId: createRes.opportunity.id };
  }
}

function buildOpportunityCustomFields_(inv, fieldMap) {
  return [
    { id: fieldMap['ar_invoice_id'], field_value: inv.invoice_id },
    { id: fieldMap['ar_invoice_number'], field_value: inv.invoice_number },
    { id: fieldMap['ar_invoice_amount'], field_value: inv.invoice_amount },
    { id: fieldMap['ar_amount_paid'], field_value: inv.amount_paid },
    { id: fieldMap['ar_balance_due'], field_value: inv.balance_due },
    { id: fieldMap['ar_due_date'], field_value: inv.due_date },
    { id: fieldMap['ar_days_overdue'], field_value: inv.days_overdue },
    { id: fieldMap['ar_aging_bucket'], field_value: inv.aging_bucket },
    { id: fieldMap['ar_invoice_status'], field_value: inv.invoice_status },
    { id: fieldMap['ar_payment_status'], field_value: inv.payment_status },
    { id: fieldMap['ar_payment_link'], field_value: inv.payment_link }
  ];
}

function determineStage_(inv, stageMap) {
  if (inv.invoice_status === 'PAID') return { stageId: stageMap['Paid'], holdReason: '' };
  if (inv.invoice_status === 'VOIDED') return { stageId: stageMap['Voided / Closed'], holdReason: '' };

  const d = inv.days_overdue;
  const s1Days = Number(getConfig_('S1_DAYS'));
  const s2Days = Number(getConfig_('S2_DAYS'));
  const s3Days = Number(getConfig_('S3_DAYS'));
  const s4Days = Number(getConfig_('S4_DAYS'));
  const handoffDays = Number(getConfig_('HANDOFF_DAYS'));

  if (d >= handoffDays) return { stageId: stageMap['Human Handoff'], holdReason: 'Backlog import' };
  if (d <= 0) return { stageId: stageMap['Not Yet Due'], holdReason: '' };
  if (d < s1Days) return { stageId: stageMap['Overdue – Grace'], holdReason: '' };
  if (d < s2Days) return { stageId: stageMap['Stage 1 – Friendly Reminder'], holdReason: '' };
  if (d < s3Days) return { stageId: stageMap['Stage 2 – Direct Reminder'], holdReason: '' };
  if (d < s4Days) return { stageId: stageMap['Stage 3 – Firm Notice'], holdReason: '' };
  return { stageId: stageMap['Stage 4 – Resolution Offer'], holdReason: '' };
}

function getComputedRows_() {
  const data = SpreadsheetApp.getActive().getSheetByName('Computed').getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(r => r.invoice_id); // laktawan ang mga blangkong rows
}

function getInvoicesById_() {
  const data = SpreadsheetApp.getActive().getSheetByName('Invoices').getDataRange().getValues();
  const headers = data[0];
  const map = {};
  data.slice(1).forEach(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    // "first occurrence wins" -- pareho ng policy natin sa Day 2 para sa duplicates (T14)
    if (obj.invoice_id && !map[obj.invoice_id]) map[obj.invoice_id] = obj;
  });
  return map;
}

function getClientsById_() {
  const data = SpreadsheetApp.getActive().getSheetByName('Clients').getDataRange().getValues();
  const headers = data[0];
  const map = {};
  data.slice(1).forEach(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    if (obj.client_id) map[obj.client_id] = obj;
  });
  return map;
}

// Ito na ang totoong laman -- WALANG sariling lock
function syncInvoicesToGhlCore_() {
  const syncLogSheet = SpreadsheetApp.getActive().getSheetByName('SyncLog');
  const computedRows = getComputedRows_();
  const invoicesById = getInvoicesById_();
  const clientsById = getClientsById_();
  let errorCount = 0, firstError = '';

  computedRows.forEach(crow => {
    try {
      const rawInv = invoicesById[crow.invoice_id];
      if (!rawInv) { syncLogSheet.appendRow([new Date(), crow.invoice_id, 'ERROR', 'Raw invoice row not found']); return; }
      const client = clientsById[crow.client_id];
      if (!client) { syncLogSheet.appendRow([new Date(), crow.invoice_id, 'ERROR', 'Client not found: ' + crow.client_id]); return; }

      const contactId = upsertContact_(client);
      if (!contactId) { syncLogSheet.appendRow([new Date(), crow.invoice_id, 'SKIPPED', 'NO_CONTACT_CHANNEL: ' + crow.client_id]); return; }

      const invForOpp = {
        invoice_id: crow.invoice_id, invoice_number: rawInv.invoice_number,
        invoice_amount: Number(rawInv.invoice_amount), amount_paid: Number(rawInv.amount_paid) || 0,
        balance_due: crow.balance_due,
        due_date: Utilities.formatDate(rawInv.due_date, 'Asia/Manila', 'yyyy-MM-dd'),
        days_overdue: crow.days_overdue, aging_bucket: crow.aging_bucket,
        invoice_status: rawInv.invoice_status, payment_status: crow.payment_status,
        payment_link: rawInv.payment_link, client_name: client.client_name
      };

      const result = upsertOpportunity_(invForOpp, contactId);
      syncLogSheet.appendRow([new Date(), crow.invoice_id, result.action, JSON.stringify(result)]);
    } catch (e) {
      errorCount++;
      if (!firstError) firstError = crow.invoice_id + ': ' + e.message;
      syncLogSheet.appendRow([new Date(), crow.invoice_id, 'ERROR', e.message]);
    }
  });

  if (errorCount > 0) {
    logSystemException_('SYNC', 'SYNC_ERRORS', errorCount + ' rows. First: ' + firstError);
    MailApp.sendEmail({
      to: getConfig_('OWNER_EMAIL'),
      subject: 'ALERT: AR sync errors (' + errorCount + ' invoices)',
      body: errorCount + ' invoice ang hindi na-sync sa GHL.\n\nUnang error: ' + firstError +
        '\n\nKaraniwang sanhi: 401 = expired o maling GHL token; 422 = maling payload. Tingnan ang SyncLog.'
    });
  }
}

// Public wrapper -- MAY sariling lock, para sa standalone/manual na pagtakbo
function syncInvoicesToGhl() {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(30000);
  if (!gotLock) { Logger.log('May isa nang tumatakbong sync -- aabort muna.'); return; }
  try {
    syncInvoicesToGhlCore_();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function runSync() {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(30000);
  if (!gotLock) {
    Logger.log('May isa nang tumatakbong runSync -- aabort muna (D6-T6 scenario).');
    return;
  }

  const startTime = new Date();
  let summary = { engine: false, sync: false, reconcile: false, rollups: false, error: null };

  try {
    runEngine();
    summary.engine = true;

    syncInvoicesToGhlCore_();
    summary.sync = true;

    reconcilePayments();
    summary.reconcile = true;

    contactRollups_();
    summary.rollups = true;

    processReplies();
    flagDataIssues_();
    planLimboWatchdog_();

    runDunningTick({ dryRun: false });
    reconcileCounts_();
    stampLastSynced_();
    SpreadsheetApp.flush();
  } catch (e) {
    summary.error = e.message;
    logSystemException_('RUNSYNC', 'RUNSYNC_FAILED', e.message);
    MailApp.sendEmail({
      to: getConfig_('OWNER_EMAIL'),
      subject: 'ALERT: runSync failed',
      body: 'Huminto ang runSync dahil sa error:\n\n' + e.message +
        '\n\nAng mga hakbang pagkatapos ng error ay hindi na tumakbo. Tingnan ang SyncLog at ExceptionLog.'
    });
  } finally {
    lock.releaseLock();
  }

  const syncLogSheet = SpreadsheetApp.getActive().getSheetByName('SyncLog');
  syncLogSheet.appendRow([new Date(), 'RUNSYNC_SUMMARY', summary.error ? 'ERROR' : 'OK', JSON.stringify(summary)]);

  Logger.log('runSync tapos sa ' + ((new Date() - startTime) / 1000) + ' segundo. Summary: ' + JSON.stringify(summary));
}

function getCustomFieldValue_(customFields, fieldId) {
  const f = (customFields || []).find(cf => cf.id === fieldId);
  if (!f) return undefined;
  if (f.fieldValueString !== undefined) return f.fieldValueString;
  if (f.fieldValueNumber !== undefined) return f.fieldValueNumber;
  if (f.value !== undefined) return f.value;
  return undefined;
}

function getOpenOpportunities_() {
  const locationId = PropertiesService.getScriptProperties().getProperty('GHL_LOCATION_ID');
  const pipelineId = PropertiesService.getScriptProperties().getProperty('PIPELINE_ID');
  let all = [];
  let url = '/opportunities/search?location_id=' + locationId + '&pipeline_id=' + pipelineId + '&limit=100';

  while (url) {
    const res = ghl_('GET', url);
    all = all.concat(res.opportunities || []);
    url = res.meta && res.meta.nextPageUrl
      ? res.meta.nextPageUrl.replace('https://services.leadconnectorhq.com', '')
      : null;
  }
  return all;
}

function setPaidNotify_(contactId, invoiceId, fieldMap, invoiceNumber) {
  const contactRes = ghl_('GET', '/contacts/' + contactId);
  const contact = contactRes.contact;
  const existingKey = getCustomFieldValue_(contact.customFields, fieldMap['ar_msg_key']);
  const existingAllowed = getCustomFieldValue_(contact.customFields, fieldMap['ar_send_allowed']);

  if (existingAllowed === 'yes' && existingKey && existingKey.indexOf(invoiceId) !== 0) {
    logSystemException_(invoiceId, 'PAID_NOTIFY_DEFERRED', 'Ibang token naka-queue: ' + existingKey);
    return;
  }

  ghl_('PUT', '/contacts/' + contactId, {
    customFields: [
      { id: fieldMap['ar_msg_key'], field_value: invoiceId + ':PAID' },
      { id: fieldMap['ar_send_allowed'], field_value: 'yes' },
      { id: fieldMap['ar_msg_invoice_number'], field_value: invoiceNumber }
    ]
  });
}

function reconcilePayments() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const computedRows = getComputedRows_();
  const computedById = {};
  computedRows.forEach(r => computedById[r.invoice_id] = r);
  const invoicesById = getInvoicesById_();

  const openOpps = getOpenOpportunities_();
  const syncLogSheet = SpreadsheetApp.getActive().getSheetByName('SyncLog');

  openOpps.forEach(opp => {
    const invoiceId = getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_id']);
    if (!invoiceId) return; // hindi natin invoice opportunity, laktawan

    const rawInv = invoicesById[invoiceId];
    const computed = computedById[invoiceId];
    if (!rawInv || !computed) return; // wala nang katumbas sa Sheet (edge case, laktawan muna)

    try {
      // RULE 1: VOIDED (first match wins -- pinaka-priority)
      if (rawInv.invoice_status === 'VOIDED') {
        ghl_('PUT', '/opportunities/' + opp.id, {
          status: 'lost',
          pipelineStageId: stageMap['Voided / Closed'],
          customFields: [{ id: fieldMap['ar_hold_reason'], field_value: 'VOIDED' }]
        });
        syncLogSheet.appendRow([new Date(), invoiceId, 'VOID_CLOSED', opp.id]);
        return;
      }

      // RULE 2: PAID (amount_due <= 0)
      if (computed.payment_status === 'PAID') {
        if (opp.status === 'won' && opp.pipelineStageId === stageMap['Paid']) {
          return; // wala nang gagawin, tapos na ito dati
        }

        const reminderStages = ['Stage 1 – Friendly Reminder', 'Stage 2 – Direct Reminder',
          'Stage 3 – Firm Notice', 'Stage 4 – Resolution Offer', 'Human Handoff',
          'Payment Plan Requested', 'Payment Plan Active'];
        const receivedReminder = reminderStages.some(name => stageMap[name] === opp.pipelineStageId);

        cleanupQueuedToken_(opp.contactId, invoiceId, fieldMap);
        if (receivedReminder) {
          setPaidNotify_(opp.contactId, invoiceId, fieldMap, rawInv.invoice_number);
        }

        ghl_('PUT', '/opportunities/' + opp.id, {
          status: 'won',
          pipelineStageId: stageMap['Paid'],
          customFields: [
            { id: fieldMap['ar_payment_status'], field_value: 'PAID' },
            { id: fieldMap['ar_balance_due'], field_value: 0 }
          ]
        });

        syncLogSheet.appendRow([new Date(), invoiceId, 'PAID_CLOSED', opp.id]);
        return;
      }

      // RULE 3 (checked before generic "walang pagbabago"): REGRESSION
      // Kung nakasaad na "won" na sa GHL, pero ngayon ay hindi na PAID ang Sheet -- iyon ay pagbabalik
      if (opp.status === 'won' && computed.payment_status !== 'PAID') {
        logSystemException_(invoiceId, 'STATUS_REGRESSION', 'GHL=won pero Sheet=' + computed.payment_status);
        syncLogSheet.appendRow([new Date(), invoiceId, 'REGRESSION_FLAGGED', opp.id]);
        return; // HUWAG awtomatikong i-reopen -- kailangan ng human review
      }

      // RULE 4: PARTIAL payment
      if (computed.payment_status === 'PARTIAL') {
        ghl_('PUT', '/opportunities/' + opp.id, {
          monetaryValue: computed.balance_due,
          customFields: [
            { id: fieldMap['ar_payment_status'], field_value: 'PARTIAL' },
            { id: fieldMap['ar_balance_due'], field_value: computed.balance_due }
          ]
        });
        syncLogSheet.appendRow([new Date(), invoiceId, 'PARTIAL_UPDATED', opp.id]);
        return;
      }

      // Walang pagbabago kailangan (UNPAID pa rin, walang regression)
    } catch (e) {
      syncLogSheet.appendRow([new Date(), invoiceId, 'ERROR', e.message]);
    }
  });
}

function cleanupQueuedToken_(contactId, invoiceId, fieldMap) {
  const contactRes = ghl_('GET', '/contacts/' + contactId);
  const contact = contactRes.contact;
  const msgKey = getCustomFieldValue_(contact.customFields, fieldMap['ar_msg_key']);
  const sendAllowed = getCustomFieldValue_(contact.customFields, fieldMap['ar_send_allowed']);

  // TANDAAN: i-check natin kung ang naka-queue ay PARA SA INVOICE NA ITO langcontactRollups_
  // (huwag galawin ang token ng IBANG invoice ng parehong contact)
  if (msgKey && msgKey.indexOf(invoiceId) === 0 && sendAllowed === 'yes') {
    ghl_('PUT', '/contacts/' + contactId, {
      customFields: [{ id: fieldMap['ar_send_allowed'], field_value: 'no' }]
    });
  }
}

function setPaidNotify(contactId, invoiceId, fieldMap) {
  const contactRes = ghl_('GET', '/contacts/' + contactId);
  const contact = contactRes.contact;
  const existingKey = getCustomFieldValue_(contact.customFields, fieldMap['ar_msg_key']);
  const existingAllowed = getCustomFieldValue_(contact.customFields, fieldMap['ar_send_allowed']);

  // Kung may IBANG invoice na naka-queue at hindi pa na-consume, huwag i-clobber -- i-log na lang
  if (existingAllowed === 'yes' && existingKey && existingKey.indexOf(invoiceId) !== 0) {
    logSystemException_(invoiceId, 'PAID_NOTIFY_DEFERRED', 'Ibang token naka-queue: ' + existingKey);
    return;
  }

  ghl_('PUT', '/contacts/' + contactId, {
    customFields: [
      { id: fieldMap['ar_msg_key'], field_value: invoiceId + ':PAID' },
      { id: fieldMap['ar_send_allowed'], field_value: 'yes' }
    ]
  });
}

function contactRollups_() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const computedRows = getComputedRows_();
  const invoicesById = getInvoicesById_();
  const clientsById = getClientsById_();

  // Group by client_id: total outstanding + bilang ng open invoices
  // (base mismo sa Sheet -- source of truth, hindi sa GHL)
  const rollups = {};
  computedRows.forEach(r => {
    const rawInv = invoicesById[r.invoice_id];
    if (!rawInv || rawInv.invoice_status === 'VOIDED' || r.payment_status === 'PAID') return;
    if (!rollups[r.client_id]) rollups[r.client_id] = { total: 0, count: 0 };
    rollups[r.client_id].total += r.balance_due;
    rollups[r.client_id].count += 1;
  });

  Object.keys(rollups).forEach(clientId => {
    const client = clientsById[clientId];
    if (!client || (!client.contact_email && !client.contact_phone)) return;
    const contactId = upsertContact_(client); // idempotent -- babalik lang ng existing ID
    if (!contactId) return;
    ghl_('PUT', '/contacts/' + contactId, {
      customFields: [
        { id: fieldMap['ar_total_outstanding'], field_value: rollups[clientId].total },
        { id: fieldMap['ar_open_invoice_count'], field_value: rollups[clientId].count }
      ]
    });
  });
}


function debugSearchOpportunities() {
  const locationId = PropertiesService.getScriptProperties().getProperty('GHL_LOCATION_ID');
  const pipelineId = PropertiesService.getScriptProperties().getProperty('PIPELINE_ID');
  const contactId = 'YOUR_TEST_CONTACT_ID'; // test contact ID (placeholder)

  const searchRes = ghl_('GET', '/opportunities/search?location_id=' + locationId +
    '&pipeline_id=' + pipelineId + '&contact_id=' + contactId);

  Logger.log('Buong search response:');
  Logger.log(JSON.stringify(searchRes, null, 2));
}

function testD5T9() {
  const props = PropertiesService.getScriptProperties();
  const realToken = props.getProperty('GHL_TOKEN'); // i-save muna ang totoong token

  props.setProperty('GHL_TOKEN', 'INTENTIONALLY_BROKEN_TOKEN'); // sirain muna
  try {
    syncInvoicesToGhl(); // dapat malinaw ang error sa SyncLog, walang crash na hindi nahuli
  } finally {
    props.setProperty('GHL_TOKEN', realToken); // IBALIK ang totoong token, kahit mag-error
  }
}

function testUpsertContact_() {
  const client = {
    client_id: 'CL-01', contact_first_name: 'Juan', contact_email: 'client01@example.com',
    contact_phone: '', client_tier: 'Regular', payment_behavior: 'On-time', relationship_notes: 'Test note'
  };
  const contactId = upsertContact_(client);
  Logger.log('Contact ID: ' + contactId);
}

function debugAiCounter() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyyMMdd');
  Logger.log('AI_COUNT today: ' + props.getProperty('AI_COUNT_' + day));
}