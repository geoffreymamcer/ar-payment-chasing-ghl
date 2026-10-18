/** Day 11 step 6 / D11-T6: 10 fixed scenarios. Auto-check (a)(b)(c)(e) sa code; (d)(f)(g)(h)(i) manual review.
 *  HINDI gumagamit ng cache -- laging bagong AI call, para totoong nasusubok ang kasalukuyang prompt. */
function promptRegression() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('PromptRegression');
  if (!sheet) {
    sheet = ss.insertSheet('PromptRegression');
    sheet.appendRow(['date', 'prompt_version', 'scenario', 'auto_pass', 'human_pass', 'notes',
                     'subject', 'email_body', 'sms_body']);
    sheet.setFrozenRows(1);
  }
  const version = getConfig_('PROMPT_VERSION');

  const base = o => Object.assign({
    first_name: 'Client', pay_link: 'https://paymentlinksample.com', sign_off: 'BDS Team',
    other_overdue_count: 0, respond_by_date: '2026-10-26',
    relationship: { tier: 'Regular', behavior: 'On time', notes: '' },
    prior: { last_stage_sent: null, days_since_last: null, last_reply_intent: null }
  }, o);
  const late = { tier: 'Regular', behavior: 'Habitually late', notes: '' };

  const scenarios = [
    ['R01 T04-S2 (Habitually late)', base({ stage: 2, first_name: 'Paolo', invoice_number: 'INV-1004', amount: 42000,
      due_date: '2026-09-23', days_overdue: 12, relationship: late, prior: { last_stage_sent: 1, days_since_last: 7, last_reply_intent: null } })],
    ['R02 T09-S3 (On time)', base({ stage: 3, first_name: 'Rico', invoice_number: 'INV-1010', amount: 52000,
      due_date: '2026-10-02', days_overdue: 17, prior: { last_stage_sent: 2, days_since_last: 7, last_reply_intent: null } })],
    ['R03 T16-S3 (Key Account)', base({ stage: 3, first_name: 'Carmela', invoice_number: 'INV-1016', amount: 95000,
      due_date: '2026-09-18', days_overdue: 17, relationship: { tier: 'Key Account', behavior: 'On time', notes: 'Long-time partner since 2022' },
      prior: { last_stage_sent: 2, days_since_last: 7, last_reply_intent: null } })],
    ['R04 T06-S3 (SMALL 1,800)', base({ stage: 3, first_name: 'Jo', invoice_number: 'INV-1006', amount: 1800,
      due_date: '2026-09-23', days_overdue: 19, prior: { last_stage_sent: 2, days_since_last: 7, last_reply_intent: null } })],
    ['R05 T07-S2 (LARGE 280,000)', base({ stage: 2, first_name: 'Miguel', invoice_number: 'INV-1007', amount: 280000,
      due_date: '2026-09-30', days_overdue: 12, prior: { last_stage_sent: 1, days_since_last: 7, last_reply_intent: null } })],
    ['R06 T09-S4 (On time)', base({ stage: 4, first_name: 'Rico', invoice_number: 'INV-1010', amount: 52000,
      due_date: '2026-10-02', days_overdue: 24, prior: { last_stage_sent: 3, days_since_last: 7, last_reply_intent: null } })],
    ['R07 Why-reply follow-up S3', base({ stage: 3, first_name: 'Paolo', invoice_number: 'INV-1004', amount: 42000,
      due_date: '2026-09-23', days_overdue: 19, relationship: late, prior: { last_stage_sent: 2, days_since_last: 7, last_reply_intent: 'why' } })],
    ['R08 Habitually late S3', base({ stage: 3, first_name: 'Paolo', invoice_number: 'INV-1004', amount: 42000,
      due_date: '2026-09-23', days_overdue: 19, relationship: late, prior: { last_stage_sent: 2, days_since_last: 7, last_reply_intent: null } })],
    ['R09 New client S2', base({ stage: 2, first_name: 'Bea', invoice_number: 'INV-1005', amount: 60000,
      due_date: '2026-09-23', days_overdue: 12, relationship: { tier: 'Regular', behavior: 'New', notes: '' },
      prior: { last_stage_sent: 1, days_since_last: 7, last_reply_intent: null } })],
    ['R10 Multi-invoice S2 (other=1)', base({ stage: 2, first_name: 'Ana', invoice_number: 'INV-1008', amount: 32000,
      due_date: '2026-10-02', days_overdue: 10, other_overdue_count: 1, combined_total: 50000,
      prior: { last_stage_sent: 1, days_since_last: 7, last_reply_intent: null } })]
  ];

  scenarios.forEach(([name, ctx]) => {
    const notes = [];
    let auto = 'FAIL', out = { subject: '', email_body: '', sms_body: '' };

    if (!aiBudgetOk_()) { // binibilang sa AI_COUNT at sumusunod sa AI_DAILY_CAP
      sheet.appendRow([new Date(), version, name, 'SKIPPED', '', 'AI cap reached']);
      return;
    }
    try {
      out = validateAi_(callGemini_(buildPrompt_(ctx)), ctx); // (a) invoice/amount/petsa, (b) link, (c) walang banta, (e) haba
      const all = out.subject + ' ' + out.email_body + ' ' + out.sms_body;
      if (all.indexOf(String(ctx.days_overdue)) === -1) notes.push('MISSING_DAYS_OVERDUE');
      if (all.indexOf(ctx.due_date) === -1) notes.push('MISSING_DUE_DATE');
      if (/₱/.test(all)) notes.push('PESO_SIGN_USED');
      if (ctx.stage === 3 && all.indexOf(ctx.respond_by_date) === -1) notes.push('S3_MISSING_RESPOND_BY');
      if (ctx.stage === 4 && !/PLAN/.test(all)) notes.push('S4_MISSING_PLAN');
      if (ctx.other_overdue_count > 0 && all.replace(/,/g, '').indexOf(String(ctx.combined_total)) === -1) notes.push('MISSING_COMBINED_TOTAL');
      if (ctx.other_overdue_count === 0 && /other overdue/i.test(all)) notes.push('MENTIONS_OTHER_WHEN_NONE');
      auto = notes.length === 0 ? 'PASS' : 'FAIL';
    } catch (e) {
      notes.push('VALIDATOR:' + e.message);
    }
    sheet.appendRow([new Date(), version, name, auto, '', notes.join('; '), out.subject, out.email_body, out.sms_body]);
  });
}