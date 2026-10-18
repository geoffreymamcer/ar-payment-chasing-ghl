function getDunningStateByInvoice_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('DunningState');
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    map[data[i][0]] = {
      invoice_id: data[i][0], client_id: data[i][1],
      current_stage: Number(data[i][2]) || 0,
      queued_key: data[i][3] || '',
      last_queued_at: data[i][4] ? new Date(data[i][4]) : null
    };
  }
  return map;
}

function getContactLastDunningSent_(clientId, stateByInvoice) {
  let latest = null;
  Object.values(stateByInvoice).forEach(s => {
    if (s.client_id === clientId && s.last_queued_at) {
      if (!latest || s.last_queued_at > latest) latest = s.last_queued_at;
    }
  });
  return latest;
}

function upsertDunningState_(invoiceId, clientId, newStage, queuedKey, timestamp) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('DunningState');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === invoiceId) {
      sheet.getRange(i + 1, 3, 1, 3).setValues([[newStage, queuedKey, timestamp]]);
      return;
    }
  }
  sheet.appendRow([invoiceId, clientId, newStage, queuedKey, timestamp]);
}

function computeTargetStage_(daysOverdue, config) {
  const s1 = Number(config.S1_DAYS), s2 = Number(config.S2_DAYS);
  const s3 = Number(config.S3_DAYS), s4 = Number(config.S4_DAYS);
  const handoff = Number(config.HANDOFF_DAYS);

  if (daysOverdue >= handoff) return 5; // 5 = HANDOFF, hindi na numbered stage
  if (daysOverdue >= s4) return 4;
  if (daysOverdue >= s3) return 3;
  if (daysOverdue >= s2) return 2;
  if (daysOverdue >= s1) return 1;
  return 0; // Grace o hindi pa due
}

function contactHasBlockingCondition_(clientId, contactId) {
  if (!contactId) return false; // walang contact pa (hindi dapat mangyari dito, pero ligtas)
  const contactRes = ghl_('GET', '/contacts/' + contactId);
  const tags = contactRes.contact.tags || [];
  const blockingTags = ['ar-paused-reply', 'ar-intent-dispute', 'ar-intent-payment-plan'];
  return blockingTags.some(tag => tags.includes(tag));
}
function decideDunningAction_(computedRow, rawInv, client, stateByInvoice, contactsQueuedThisTick, config, today, contactId, oppStageId, stageMap) {
  const invoiceId = computedRow.invoice_id;
  const clientId = computedRow.client_id;
  const daysOverdue = computedRow.days_overdue;
  const state = stateByInvoice[invoiceId]; // undefined kung "first-time import"
  const currentStage = state ? state.current_stage : 0;
  const isFirstImport = !state;

  // Rule 1: exception/invalid -> HOLD
  // (Hindi na kailangan dito -- ang mga invalid rows ay na-filter na ng Day 2 engine,
  //  hindi na sila umaabot sa Computed tab, kaya hindi na sila umaabot sa function na ito.)

  // Rule 2: hindi na open (paid/void)
  if (computedRow.payment_status === 'PAID' || rawInv.invoice_status === 'VOIDED') {
    return { action: 'NONE', target: currentStage, reason: 'Invoice not open' };
  }

  // Rule 3: contact-level blocking condition
  if (contactHasBlockingCondition_(clientId, contactId)) {
    return { action: 'NONE', target: currentStage, reason: 'Contact paused/dispute/plan' };
  }

  // Rule 3b: opportunity-level hold (Payment Plan o Disputed) -- kahit nabura na ang tag
  const holdStages = [stageMap['Payment Plan Requested'], stageMap['Payment Plan Active'], stageMap['Disputed / On Hold']];
  if (oppStageId && holdStages.includes(oppStageId)) {
    return { action: 'NONE', target: currentStage, reason: 'Opportunity in Plan/Dispute stage' };
  }


  // Rule 4: hindi pa due
  if (daysOverdue <= 0) {
    return { action: 'NONE', target: 0, reason: 'Not yet due' };
  }

  // Rule 5: Grace period (1-2 days, wala pang S1 threshold)
  const s1Days = Number(config.S1_DAYS);
  if (daysOverdue < s1Days) {
    return { action: 'NONE', target: 0, reason: 'Grace period' };
  }

  // Rule 6: Backlog rule -- UNANG BESES lang ito ma-import, at agad na sobra-sobra na sa overdue
  const handoffDays = Number(config.HANDOFF_DAYS);
  if (isFirstImport && daysOverdue > handoffDays) {
    return { action: 'HANDOFF', target: 5, reason: 'Backlog import' };
  }

  // Rule 7: Tier/relationship overrides
  const s3Days = Number(config.S3_DAYS);
  let target = computeTargetStage_(daysOverdue, config); // baseline mula rule 8

  const invoiceTier = computedRow.tier; // SMALL/STANDARD/LARGE mula Computed tab (Day 2)
  const isHighValue = (invoiceTier === 'LARGE' || client.client_tier === 'Key Account');

  if (isHighValue && daysOverdue >= s3Days) {
    return { action: 'HANDOFF', target: 5, reason: 'High-value' };
  }
  if (invoiceTier === 'SMALL' && daysOverdue >= s3Days) {
    if (daysOverdue >= handoffDays) {
      return { action: 'LOW_VALUE_REVIEW', target: target, reason: 'Low-value backlog' };
    }
    target = Math.min(target, 3);
  }

  // Rule 9: idempotency -- kung meron na o lampas na sa target, wala nang gagawin
  if (currentStage >= target) {
    return { action: 'NONE', target: target, reason: 'Already at or past target stage' };
  }

  // Rule 10: per-contact cooldown
  const cooldownHours = Number(config.COOLDOWN_HOURS);
  const lastSent = getContactLastDunningSent_(clientId, stateByInvoice);
  if (lastSent) {
    const hoursSince = (today - lastSent) / (1000 * 60 * 60);
    if (hoursSince < cooldownHours) {
      return { action: 'DEFER', target: target, reason: 'Cooldown active (' + hoursSince.toFixed(1) + 'h)' };
    }
  }

  // Rule 11: per-contact serialization -- isa lang ang QUEUE kada contact kada tick
  if (contactsQueuedThisTick.has(clientId)) {
    return { action: 'DEFER', target: target, reason: 'Another invoice already queued this tick' };
  }

  // Rule 12: pre-flight re-check -- basahin ULIT ang Sheet, baka nagbago simula nung na-Compute
  const freshInvoicesById = getInvoicesById_();
  const freshInv = freshInvoicesById[invoiceId];
  if (!freshInv || freshInv.invoice_status === 'VOIDED' ||
    (Number(freshInv.invoice_amount) - (Number(freshInv.amount_paid) || 0)) <= 0) {
    return { action: 'NONE', target: target, reason: 'PREFLIGHT_STOP' };
  }

  // Rule 13: lahat ng checks pumasa -- i-queue na, PERO kung ang target ay HANDOFF-level (5),
  // dapat maging HANDOFF ang action, hindi QUEUE_STAGE -- walang "message" para sa "stage 5".
  contactsQueuedThisTick.add(clientId);
  if (target === 5) {
    return { action: 'HANDOFF', target: 5, reason: 'Aged past handoff threshold' };
  }
  return { action: 'QUEUE_STAGE', target: target, reason: 'Threshold reached' };
}

function getConfigObject_() {
  const rows = SpreadsheetApp.getActive().getSheetByName('Config').getDataRange().getValues();
  const config = {};
  rows.forEach(([k, v]) => config[k] = v);
  return config;
}

function runDunningTick(options) {
  const dryRun = !options || options.dryRun !== false; // default: true (dry run)

  const config = getConfigObject_();
  const today = today_();
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const computedRows = getComputedRows_();
  const invoicesById = getInvoicesById_();
  const clientsById = getClientsById_();
  const stateByInvoice = getDunningStateByInvoice_();
  const planSheet = SpreadsheetApp.getActive().getSheetByName('DunningPlan');

  if (!dryRun) planMonitor_(); // i-check muna ang mga plans bago ang normal dunning decisions

  // I-sort: pinaka-maagang due_date muna ("mas matanda" ang invoice), para sa serialization (rule 11)
  const sorted = computedRows.slice().sort((a, b) => {
    const invA = invoicesById[a.invoice_id], invB = invoicesById[b.invoice_id];
    if (!invA || !invB) return 0;
    return invA.due_date - invB.due_date;
  });

  const contactsQueuedThisTick = new Set();

  sorted.forEach(row => {
    const rawInv = invoicesById[row.invoice_id];
    const client = clientsById[row.client_id];
    if (!rawInv || !client) return;

    const state = stateByInvoice[row.invoice_id];
    const currentStage = state ? state.current_stage : 0;

    const oppInfoForCheck = findOpportunityForInvoice_(row.invoice_id);
    const contactId = oppInfoForCheck ? oppInfoForCheck.contactId : null;
    const oppStageId = oppInfoForCheck ? oppInfoForCheck.stageId : null;

    const decision = decideDunningAction_(row, rawInv, client, stateByInvoice, contactsQueuedThisTick, config, today, contactId, oppStageId, stageMap);

    planSheet.appendRow([
      new Date(), row.invoice_id, row.tier, row.days_overdue,
      currentStage, decision.action, decision.target, decision.reason
    ]);

    if (decision.action === 'QUEUE_STAGE' || decision.action === 'HANDOFF') {
      const newStage = decision.target;
      const key = row.invoice_id + ':S' + newStage;
      upsertDunningState_(row.invoice_id, row.client_id, newStage, key, today);
      stateByInvoice[row.invoice_id] = {
        invoice_id: row.invoice_id, client_id: row.client_id,
        current_stage: newStage, queued_key: key, last_queued_at: today
      };
    }

    if (!dryRun && decision.action === 'QUEUE_STAGE') {
      const oppInfo = oppInfoForCheck || findOpportunityForInvoice_(row.invoice_id);
      if (!oppInfo) {
        logSystemException_(row.invoice_id, 'OPPORTUNITY_NOT_FOUND_FOR_QUEUE', 'N/A');
      } else {
        const msg = buildMessage_(row, rawInv, client, decision.target, stateByInvoice, computedRows);
        const streak = trackAiOutcome_(row.invoice_id, msg.source);

        const isKeyAccount = client.client_tier === 'Key Account';
        const isHighStage = decision.target === 3 || decision.target === 4;
        const consecutiveFallback = (msg.source === 'FALLBACK') && streak >= 2;

        if (isKeyAccount && isHighStage && consecutiveFallback) {
          const summary = buildHandoffSummary_(row, rawInv, client, 'AI_FAILURE_ESCALATION', stateByInvoice, computedRows);
          const invForHandoff = {
            invoice_id: row.invoice_id, invoice_number: rawInv.invoice_number,
            contactId: oppInfo.contactId, oppId: oppInfo.oppId, queued_key: oppInfo.queued_key
          };
          queueHandoff_(invForHandoff, 'AI_FAILURE_ESCALATION', summary);
        } else {
          const invForQueue = {
            invoice_id: row.invoice_id, invoice_number: rawInv.invoice_number,
            contactId: oppInfo.contactId, oppId: oppInfo.oppId, queued_key: oppInfo.queued_key
          };
          queueStage_(invForQueue, decision.target, msg);
        }
      }
    }

    if (!dryRun && decision.action === 'HANDOFF') {
      const oppInfo = oppInfoForCheck || findOpportunityForInvoice_(row.invoice_id);
      if (!oppInfo) {
        logSystemException_(row.invoice_id, 'OPPORTUNITY_NOT_FOUND_FOR_HANDOFF', 'N/A');
      } else {
        const reasonCode = mapToHandoffReason_(decision.reason);
        const summary = buildHandoffSummary_(row, rawInv, client, reasonCode, stateByInvoice, computedRows);
        const invForHandoff = {
          invoice_id: row.invoice_id, invoice_number: rawInv.invoice_number,
          contactId: oppInfo.contactId, oppId: oppInfo.oppId, queued_key: oppInfo.queued_key
        };
        queueHandoff_(invForHandoff, reasonCode, summary);
      }
    }

    if (!dryRun && decision.action === 'LOW_VALUE_REVIEW') {
      const oppInfo = oppInfoForCheck || findOpportunityForInvoice_(row.invoice_id);
      if (oppInfo) {
        const invForLowValue = {
          invoice_id: row.invoice_id, invoice_number: rawInv.invoice_number,
          oppId: oppInfo.oppId, queued_key: oppInfo.queued_key
        };
        queueLowValueReview_(invForLowValue, 'LOW_VALUE_BACKLOG');
      }
    }
  });

  SpreadsheetApp.flush();
}

function findOpportunityForInvoice_(invoiceId) {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const allOpps = getOpenOpportunities_();
  const opp = allOpps.find(o => getCustomFieldValue_(o.customFields, fieldMap['ar_invoice_id']) === invoiceId);
  if (!opp) return null;
  return {
    oppId: opp.id,
    contactId: opp.contactId,
    queued_key: getCustomFieldValue_(opp.customFields, fieldMap['ar_stage_queued_key']) || '',
    stageId: opp.pipelineStageId
  };
}

function ghlUpdateContactFields_(contactId, fieldsObj) {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const customFields = Object.keys(fieldsObj).map(k => ({ id: fieldMap[k], field_value: fieldsObj[k] }));
  ghl_('PUT', '/contacts/' + contactId, { customFields });
}

function ghlUpdateOpportunity_(oppId, opts) {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const payload = {};
  if (opts.stageId) payload.pipelineStageId = opts.stageId;
  if (opts.fields) {
    payload.customFields = Object.keys(opts.fields).map(k => ({ id: fieldMap[k], field_value: opts.fields[k] }));
  }
  ghl_('PUT', '/opportunities/' + oppId, payload);
}

function logComm_(inv, stage, source, status) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('CommLog');
  const key = inv.invoice_id + ':S' + stage;
  sheet.appendRow([new Date(), inv.invoice_id, stage, source, status, key]);
}

const DUNNING_STAGE_NAMES_ = {
  1: 'Stage 1 – Friendly Reminder',
  2: 'Stage 2 – Direct Reminder',
  3: 'Stage 3 – Firm Notice',
  4: 'Stage 4 – Resolution Offer'
};

function queueStage_(inv, stage, msg) {
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const key = inv.invoice_id + ':S' + stage;

  Logger.log('queueStage_ called for ' + inv.invoice_id + ' | computed key=' + key + ' | inv.queued_key=' + JSON.stringify(inv.queued_key));

  if (inv.queued_key === key) {
    Logger.log('SKIP_DUPLICATE triggered for ' + inv.invoice_id);
    return 'SKIP_DUPLICATE';
  }

  Logger.log('Proceeding to write for ' + inv.invoice_id);

  // 1) WRITE DATA FIRST
  ghlUpdateContactFields_(inv.contactId, {
    ar_msg_subject: msg.subject, ar_msg_email_body: msg.email, ar_msg_sms_body: msg.sms,
    ar_msg_invoice_number: inv.invoice_number, ar_msg_key: key,
    ar_msg_source: msg.source, ar_send_allowed: 'yes'
  });

  Logger.log('Contact fields written for ' + inv.invoice_id);

  // 2) THEN MOVE STAGE (ito ang nagpapaandar sa workflow)
  ghlUpdateOpportunity_(inv.oppId, {
    stageId: stageMap[DUNNING_STAGE_NAMES_[stage]],
    fields: { ar_dunning_stage: String(stage), ar_stage_queued_key: key }
  });

  Logger.log('Opportunity updated for ' + inv.invoice_id);

  logComm_(inv, stage, msg.source, 'QUEUED');

  Logger.log('CommLog written for ' + inv.invoice_id);

  return 'QUEUED';
}

function buildMsgCtx_(row, rawInv, client, targetStage, stateByInvoice, computedRows) {
  // Bilangin ang ibang OPEN invoices ng parehong client (hindi kasama itong isa)
  const others = computedRows.filter(r =>
    r.client_id === row.client_id && r.invoice_id !== row.invoice_id &&
    r.payment_status !== 'PAID' && Number(r.days_overdue) > 0
  );
  const otherOverdueCount = others.length;
  const combinedTotal = Number(row.balance_due) + others.reduce((s, r) => s + Number(r.balance_due || 0), 0);

  const state = stateByInvoice[row.invoice_id];
  const priorStage = state ? state.current_stage : null;
  const priorSentAt = state ? state.last_queued_at : null;
  const daysSinceLast = priorSentAt ? Math.round((today_() - priorSentAt) / 86400000) : null;
  const respondByDate = new Date(today_());
  respondByDate.setDate(respondByDate.getDate() + 7);

  return {
    stage: targetStage,
    first_name: client.contact_first_name,
    invoice_number: rawInv.invoice_number,
    amount: row.balance_due,
    due_date: Utilities.formatDate(rawInv.due_date, 'Asia/Manila', 'yyyy-MM-dd'),
    days_overdue: row.days_overdue,
    other_overdue_count: otherOverdueCount,
    combined_total: combinedTotal,
    relationship: {
      tier: client.client_tier,
      behavior: client.payment_behavior,
      notes: client.relationship_notes || ''
    },
    prior: {
      last_stage_sent: priorStage,
      days_since_last: daysSinceLast,
      last_reply_intent: null // Day 12 pa ito gagawin
    },
    pay_link: rawInv.payment_link,
    sign_off: 'BDS Team',
    respond_by_date: Utilities.formatDate(respondByDate, 'Asia/Manila', 'yyyy-MM-dd'),

  };
}

function runDunningTickLive() {
  runDunningTick({ dryRun: false });
}


function buildHandoffSummary_(row, rawInv, client, reasonCode, stateByInvoice, computedRows) {
  const state = stateByInvoice[row.invoice_id];
  const timeline = [];

  // Timeline mula CommLog -- lahat ng naipadalang stages para dito
  const commLogSheet = SpreadsheetApp.getActive().getSheetByName('CommLog');
  const commData = commLogSheet.getDataRange().getValues();
  for (let i = 1; i < commData.length; i++) {
    if (commData[i][1] === row.invoice_id) {
      timeline.push(Utilities.formatDate(commData[i][0], 'Asia/Manila', 'yyyy-MM-dd HH:mm') +
        ' - Stage ' + commData[i][2] + ' (' + commData[i][3] + ')');
    }
  }

  // Ibang open invoices ng parehong client
  const otherInvoices = computedRows
    .filter(r => r.client_id === row.client_id && r.invoice_id !== row.invoice_id && r.payment_status !== 'PAID')
    .map(r => r.invoice_id + ' (PHP ' + formatPeso_(r.balance_due) + ', ' + r.days_overdue + 'd overdue)');

  const lines = [
    'HANDOFF REASON: ' + reasonCode,
    'Invoice: ' + rawInv.invoice_number + ' | Amount: PHP ' + formatPeso_(row.balance_due) + ' | Days overdue: ' + row.days_overdue,
    'Client: ' + client.client_name + ' (' + client.client_tier + ', ' + client.payment_behavior + ')',
    'Relationship notes: ' + (client.relationship_notes || 'N/A'),
    '',
    'TIMELINE:',
    timeline.length > 0 ? timeline.join('\n') : '(walang naitalang comms)',
    '',
    'OTHER OPEN INVOICES: ' + (otherInvoices.length > 0 ? otherInvoices.join(', ') : 'wala'),
    '',
    'LAST REPLY: ' + 'wala pa' // Day 12 pa ito gagawin
  ];

  return lines.join('\n');
}

function mapToHandoffReason_(internalReason) {
  const map = {
    'Backlog import': 'BACKLOG_IMPORT',
    'High-value': 'HIGH_VALUE',
    'Aged past handoff threshold': 'NO_RESPONSE'
  };
  return map[internalReason] || 'OTHER_REPLY';
}

function queueHandoff_(inv, reasonCode, summary) {
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const key = inv.invoice_id + ':HANDOFF';

  if (inv.queued_key === key) return 'SKIP_DUPLICATE';

  ghlUpdateContactFields_(inv.contactId, {
    ar_msg_key: key,
    ar_send_allowed: 'no'
  });

  // Siguraduhin na-tag ang contact bilang "active" kahit Backlog/High-value direktang
  // pumunta dito nang hindi dumaan sa normal na S1-S4 (kung saan karaniwan natin
  // idinadagdag ang tag na ito) -- kailangan natin ito para gumana ang WF-06 reply detection.
  ghl_('POST', '/contacts/' + inv.contactId + '/tags', { tags: ['ar-active'] });

  ghlUpdateOpportunity_(inv.oppId, {
    stageId: stageMap['Human Handoff'],
    fields: {
      ar_dunning_stage: '5',
      ar_stage_queued_key: key,
      ar_hold_reason: reasonCode,
      ar_handoff_summary: summary
    }
  });

  logComm_(inv, 'HANDOFF', reasonCode, 'QUEUED');
  return 'QUEUED';
}

function handoffWatchdog() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const ownerEmail = getConfig_('OWNER_EMAIL');
  const thresholdHours = Number(getConfig_('WATCHDOG_THRESHOLD_HOURS')) || 24;
  const props = PropertiesService.getScriptProperties();

  const allOpps = getOpenOpportunities_();
  const handoffOpps = allOpps.filter(o => o.pipelineStageId === stageMap['Human Handoff'] && o.status === 'open');

  handoffOpps.forEach(opp => {
    // Ang LOW_VALUE_BACKLOG ay sa weekly digest napupunta, hindi urgent alert (M4 policy)
    if (getCustomFieldValue_(opp.customFields, fieldMap['ar_hold_reason']) === 'LOW_VALUE_BACKLOG') return;

    const hoursInStage = (new Date() - new Date(opp.lastStageChangeAt)) / (1000 * 60 * 60);
    if (hoursInStage < thresholdHours) return;

    const contactRes = ghl_('GET', '/contacts/' + opp.contactId);
    const tags = contactRes.contact.tags || [];
    if (tags.includes('ar-handoff-ack')) return;

    const alertKey = 'WATCHDOG_ALERTED_' + opp.id;
    if (props.getProperty(alertKey)) return; // isang beses lang

    const invoiceId = getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_id']);
    const invoiceNumber = getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_number']);
    const holdReason = getCustomFieldValue_(opp.customFields, fieldMap['ar_hold_reason']);

    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'WATCHDOG: Unacknowledged Handoff - ' + invoiceNumber,
      body: 'Invoice ' + invoiceNumber + ' (' + invoiceId + ') ay nasa Human Handoff nang ' +
        hoursInStage.toFixed(1) + ' oras, walang acknowledgment pa.\n\nReason: ' + holdReason +
        '\n\nPakitingnan sa GHL dashboard.'
    });

    props.setProperty(alertKey, new Date().toISOString());
    logSystemException_(invoiceId, 'WATCHDOG_ALERT_SENT', 'Hours in handoff: ' + hoursInStage.toFixed(1));
  });
}

function queueLowValueReview_(inv, reasonCode) {
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const key = inv.invoice_id + ':LOWVALUE';

  if (inv.queued_key === key) return 'SKIP_DUPLICATE';

  ghlUpdateOpportunity_(inv.oppId, {
    stageId: stageMap['Human Handoff'], // parehong stage, pero walang task/notification
    fields: {
      ar_dunning_stage: '5',
      ar_stage_queued_key: key,
      ar_hold_reason: 'LOW_VALUE_BACKLOG'
    }
  });

  logComm_(inv, 'LOWVALUE', 'LOW_VALUE_BACKLOG', 'QUEUED');
  return 'QUEUED';
}

function sendLowValueDigest_() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const ownerEmail = getConfig_('OWNER_EMAIL');
  const allOpps = getOpenOpportunities_();

  const lowValueOpps = allOpps.filter(o =>
    getCustomFieldValue_(o.customFields, fieldMap['ar_hold_reason']) === 'LOW_VALUE_BACKLOG' && o.status === 'open'
  );

  if (lowValueOpps.length === 0) return;

  const lines = lowValueOpps.map(o => {
    const invNum = getCustomFieldValue_(o.customFields, fieldMap['ar_invoice_number']);
    const balance = getCustomFieldValue_(o.customFields, fieldMap['ar_balance_due']);
    const days = getCustomFieldValue_(o.customFields, fieldMap['ar_days_overdue']);
    return invNum + ' - PHP ' + balance + ' - ' + days + ' days overdue';
  });

  MailApp.sendEmail({
    to: ownerEmail,
    subject: 'Weekly Low-Value Backlog Digest (' + lowValueOpps.length + ' invoices)',
    body: 'Mga small invoices na overdue nang matagal, hindi na awtomatikong ginagalaw:\n\n' + lines.join('\n')
  });
}

function confirmSends() {
  const commLogSheet = SpreadsheetApp.getActive().getSheetByName('CommLog');
  const data = commLogSheet.getDataRange().getValues();
  const now = new Date();
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const ownerEmail = getConfig_('OWNER_EMAIL');
  const confirmRaw = getConfig_('SEND_CONFIRM_HOURS');
  const confirmHours = (confirmRaw === '' || confirmRaw === null) ? 2 : Number(confirmRaw);
  const invoicesById = getInvoicesById_();
  const props = PropertiesService.getScriptProperties();

  // Iisa lang ang ar_last_sent_confirm kada contact, kaya ang PINAKAHULING padala lang ang masusuri.
  const latestRowByClient = {};
  for (let i = 1; i < data.length; i++) {
    const [timestamp, invoiceId, stage, , status] = data[i];
    if (status !== 'QUEUED' || stage === 'HANDOFF' || stage === 'LOWVALUE') continue;
    const inv = invoicesById[invoiceId];
    if (!inv) continue;
    const prev = latestRowByClient[inv.client_id];
    if (prev === undefined || timestamp > data[prev][0]) latestRowByClient[inv.client_id] = i;
  }

  Object.keys(latestRowByClient).forEach(clientId => {
    const [timestamp, invoiceId, , , , key] = data[latestRowByClient[clientId]];
    const inv = invoicesById[invoiceId];

    // Bayad na -> wala nang saysay ang alert para sa reminder
    if ((Number(inv.invoice_amount) - (Number(inv.amount_paid) || 0)) <= 0) return;

    const hoursSinceQueued = (now - timestamp) / (1000 * 60 * 60);
    if (hoursSinceQueued < confirmHours) return; // baka nasa Wait pa lang

    const oppInfo = findOpportunityForInvoice_(invoiceId);
    if (!oppInfo) return;

    const contactRes = ghl_('GET', '/contacts/' + oppInfo.contactId);
    const lastConfirm = getCustomFieldValue_(contactRes.contact.customFields, fieldMap['ar_last_sent_confirm']);
    if (lastConfirm === key) return; // na-confirm, ayos

    const alertKey = 'SEND_UNCONFIRMED_ALERTED_' + key;
    if (props.getProperty(alertKey)) return; // isang beses lang

    logSystemException_(invoiceId, 'SEND_UNCONFIRMED',
      'Key: ' + key + ', last confirm: ' + (lastConfirm || 'none') + ', hours since queued: ' + hoursSinceQueued.toFixed(1));
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'ALERT: Unconfirmed Send - ' + invoiceId,
      body: 'Ang invoice ' + invoiceId + ' (key: ' + key + ') ay naka-QUEUE na mula ' + hoursSinceQueued.toFixed(1) +
        ' oras, pero hindi pa na-confirm na naipadala (last confirm: ' + (lastConfirm || 'none') + ').\n\n' +
        'Posibleng silent failure: naka-draft o may error ang workflow. Tingnan ang Enrollment History.'
    });
    props.setProperty(alertKey, new Date().toISOString());
  });
}

function testSetPaidNotify() {
  const contactId = 'YOUR_TEST_CONTACT_ID'; // test contact ID (placeholder)
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  setPaidNotify_(contactId, 'INV-1011', fieldMap, 'INV-1011');
  Logger.log('Tapos na, i-check ang contact fields.');
}

function clearWatchdogAlert(oppId) {
  const props = PropertiesService.getScriptProperties();
  const key = 'WATCHDOG_ALERTED_' + oppId;
  const existing = props.getProperty(key);
  Logger.log('Existing value: ' + existing);
  props.deleteProperty(key);
  Logger.log('Cleared.');
}


function debugWatchdogState() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const props = PropertiesService.getScriptProperties();

  const allOpps = getOpenOpportunities_();
  const handoffOpps = allOpps.filter(o => o.pipelineStageId === stageMap['Human Handoff']);

  handoffOpps.forEach(opp => {
    const invoiceId = getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_id']);
    const hoursInStage = (new Date() - new Date(opp.lastStageChangeAt)) / (1000 * 60 * 60);
    const alertKey = 'WATCHDOG_ALERTED_' + opp.id;
    const alreadyAlerted = props.getProperty(alertKey);

    Logger.log('--- ' + invoiceId + ' (oppId: ' + opp.id + ') ---');
    Logger.log('Hours in stage: ' + hoursInStage.toFixed(1));
    Logger.log('Already alerted (script property): ' + alreadyAlerted);
    Logger.log('Opp status: ' + opp.status);

    const contactRes = ghl_('GET', '/contacts/' + opp.contactId);
    const tags = contactRes.contact.tags || [];
    Logger.log('Tags: ' + JSON.stringify(tags));
    Logger.log('Has ar-handoff-ack: ' + tags.includes('ar-handoff-ack'));
  });
}

function buildMessage_(row, rawInv, client, stage, stateByInvoice, computedRows) {
  Logger.log('buildMessage_ called for ' + row.invoice_id + ' stage=' + stage);
  const ctx = buildMsgCtx_(row, rawInv, client, stage, stateByInvoice, computedRows);
  Logger.log('ctx built, respond_by_date=' + ctx.respond_by_date);
  const fallback = fallbackTemplate_(ctx);
  Logger.log('fallback built');

  if (stage === 1) {
    return { subject: fallback.subject, email: fallback.email_body, sms: fallback.sms_body, source: 'TEMPLATE' };
  }

  const mode = getConfig_('AI_MODE');
  Logger.log('AI_MODE=' + mode);
  if (mode === 'OFF') {
    return { subject: fallback.subject, email: fallback.email_body, sms: fallback.sms_body, source: 'FALLBACK' };
  }

  Logger.log('Calling generateMessage_...');
  const result = generateMessage_(ctx);
  Logger.log('generateMessage_ returned, source=' + result.source);

  if (mode === 'SHADOW') {
    logShadow_(row.invoice_id, stage, result, fallback);
    return { subject: fallback.subject, email: fallback.email_body, sms: fallback.sms_body, source: 'TEMPLATE' };
  }

  return result;
}

function logShadow_(invoiceId, stage, aiResult, fallbackResult) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('CommLog');
  const key = invoiceId + ':S' + stage + ':SHADOW';
  sheet.appendRow([new Date(), invoiceId, stage, 'AI_SHADOW', 'LOGGED', key]);
}

function trackAiOutcome_(invoiceId, source) {
  const props = PropertiesService.getScriptProperties();
  const key = 'AI_FALLBACK_STREAK_' + invoiceId;

  if (source === 'AI') {
    props.deleteProperty(key); // successful AI -- i-reset ang streak
    return 0;
  }

  // FALLBACK o TEMPLATE -- dagdagan ang streak
  const current = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(current));
  return current;
}

function clearAiStreak() {
  PropertiesService.getScriptProperties().deleteProperty('AI_FALLBACK_STREAK_INV-1016');
  Logger.log('Cleared.');
}

function testAiFailureEscalation() {
  const invoicesById = getInvoicesById_();
  const clientsById = getClientsById_();
  const computedRows = getComputedRows_();
  const stateByInvoice = getDunningStateByInvoice_();

  const rawInv = invoicesById['INV-1016'];
  const client = clientsById['CL-16']; // Key Account
  const row = computedRows.find(r => r.invoice_id === 'INV-1016');

  if (!row) { Logger.log('INV-1016 wala sa Computed -- i-run muna ang runEngine()'); return; }

  // I-simulate: 2 sunud-sunod na calls sa stage 3, parehong dapat mabigo dahil sira ang API key
  for (let i = 1; i <= 2; i++) {
    const msg = buildMessage_(row, rawInv, client, 3, stateByInvoice, computedRows);
    const streak = trackAiOutcome_('INV-1016', msg.source);
    Logger.log('Call ' + i + ': source=' + msg.source + ', streak=' + streak);

    const isKeyAccount = client.client_tier === 'Key Account';
    const isHighStage = true; // stage 3
    const consecutiveFallback = (msg.source === 'FALLBACK') && streak >= 2;
    Logger.log('isKeyAccount=' + isKeyAccount + ', consecutiveFallback=' + consecutiveFallback);

    if (isKeyAccount && isHighStage && consecutiveFallback) {
      Logger.log('*** DAPAT MAG-HANDOFF DITO ***');
    }
  }
}

function planMonitor_() {
  const fieldMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('FIELD_MAP'));
  const stageMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('STAGE_MAP'));
  const today = today_();
  const invoicesById = getInvoicesById_();
  const computedRows = getComputedRows_();
  const clientsById = getClientsById_();
  const stateByInvoice = getDunningStateByInvoice_();

  const allOpps = getOpenOpportunities_();
  const activeOpps = allOpps.filter(o => o.pipelineStageId === stageMap['Payment Plan Active'] && o.status === 'open');

  activeOpps.forEach(opp => {
    const invoiceId = getCustomFieldValue_(opp.customFields, fieldMap['ar_invoice_id']);
    const nextDueStr = getCustomFieldValue_(opp.customFields, fieldMap['ar_plan_next_due']);
    const expectedPaid = Number(getCustomFieldValue_(opp.customFields, fieldMap['ar_plan_expected_paid']) || 0);

    if (!invoiceId || !nextDueStr) return; // hindi pa kumpleto ang human approval setup

    const rawInv = invoicesById[invoiceId];
    if (!rawInv) return;

    const amountPaid = Number(rawInv.amount_paid) || 0;
    const invoiceAmount = Number(rawInv.invoice_amount) || 0;
    const balanceDue = invoiceAmount - amountPaid;

    if (balanceDue <= 0) return; // fully paid -- reconcilePayments_ na ang bahala

    const nextDue = new Date(nextDueStr + 'T00:00:00+08:00');
    const graceEnd = new Date(nextDue.getTime() + 3 * 86400000); // +3 araw na grace

    if (today <= graceEnd) return; // nasa loob pa ng grace, walang gagawin

    if (amountPaid >= expectedPaid) return; // on-track, maghintay na lang ng susunod na update ng tao

    // Lagpas na sa grace AT kulang pa rin ang bayad -> HANDOFF PLAN_MISSED
    const key = invoiceId + ':PLAN_MISSED';
    const existingKey = getCustomFieldValue_(opp.customFields, fieldMap['ar_stage_queued_key']);
    if (existingKey === key) return; // idempotent guard, naka-handoff na dati

    const row = computedRows.find(r => r.invoice_id === invoiceId);
    const client = clientsById[rawInv.client_id];
    let summary = 'HANDOFF REASON: PLAN_MISSED';
    if (row && client) {
      summary = buildHandoffSummary_(row, rawInv, client, 'PLAN_MISSED', stateByInvoice, computedRows);
    }
    summary += '\n\nPAYMENT PLAN: next due ' + nextDueStr +
      ' | expected paid PHP ' + formatPeso_(expectedPaid) +
      ' | actual paid PHP ' + formatPeso_(amountPaid);

    ghl_('PUT', '/opportunities/' + opp.id, {
      pipelineStageId: stageMap['Human Handoff'],
      customFields: [
        { id: fieldMap['ar_hold_reason'], field_value: 'PLAN_MISSED' },
        { id: fieldMap['ar_plan_status'], field_value: 'MISSED' },
        { id: fieldMap['ar_stage_queued_key'], field_value: key },
        { id: fieldMap['ar_handoff_summary'], field_value: summary },
      ]
    });
    ghlUpdateContactFields_(opp.contactId, { ar_send_allowed: 'no' });

    logSystemException_(invoiceId, 'PLAN_MISSED_HANDOFF',
      'next_due=' + nextDueStr + ', expected=' + expectedPaid + ', paid=' + amountPaid);
  });
}