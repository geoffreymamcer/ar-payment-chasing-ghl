/** ============ FALLBACK TEMPLATES (deterministic, pre-approved) ============ **/

function fallbackTemplate_(ctx) {
  const isKeyAccount = ctx.relationship.tier === 'Key Account';
  const amountStr = formatPeso_(ctx.amount);
  const greeting = isKeyAccount ? `Hi ${ctx.first_name}, hope you're doing well!` : `Hi ${ctx.first_name},`;

  let subject, email_body, sms_body;

  if (ctx.stage === 1) {
    subject = `Friendly reminder: Invoice ${ctx.invoice_number}`;
    email_body = `${greeting} Just a quick heads-up that invoice ${ctx.invoice_number} (PHP ${amountStr}) was due last ${ctx.due_date}. No worries if it slipped your mind — here's the payment link: ${ctx.pay_link}. Let us know if you have any questions. Thank you! — ${ctx.sign_off}`;
    sms_body = `Hi ${ctx.first_name}, friendly reminder: invoice ${ctx.invoice_number} (PHP ${amountStr}) was due ${ctx.due_date}. Pay here: ${ctx.pay_link} — ${ctx.sign_off}`;
  } else if (ctx.stage === 2) {
    subject = `Following up: Invoice ${ctx.invoice_number}`;
    email_body = `${greeting} Following up po sa invoice ${ctx.invoice_number} (PHP ${amountStr}) na due noong ${ctx.due_date}. Baka nakalampas lang po; ito ang link para sa payment: ${ctx.pay_link}. Kung may concern, reply lang po dito. Salamat! — ${ctx.sign_off}`;
    sms_body = `Hi ${ctx.first_name}, following up on invoice ${ctx.invoice_number} (PHP ${amountStr}), due ${ctx.due_date}, now ${ctx.days_overdue} days overdue. Pay: ${ctx.pay_link} — ${ctx.sign_off}`;
  } else if (ctx.stage === 3) {
    subject = `Action needed: Invoice ${ctx.invoice_number} is ${ctx.days_overdue} days overdue`;
    email_body = `${greeting} Invoice ${ctx.invoice_number} (PHP ${amountStr}) is now ${ctx.days_overdue} days overdue since ${ctx.due_date}. Kindly settle this or reply to us within 3 days so we can coordinate. Payment link: ${ctx.pay_link}. We appreciate your prompt attention to this. — ${ctx.sign_off}`;
    sms_body = `Hi ${ctx.first_name}, invoice ${ctx.invoice_number} (PHP ${amountStr}) is ${ctx.days_overdue} days overdue. Please settle or reply within 3 days. Pay: ${ctx.pay_link} — ${ctx.sign_off}`;
  } else {
    subject = `Please respond: Invoice ${ctx.invoice_number}`;
    email_body = `${greeting} Invoice ${ctx.invoice_number} (PHP ${amountStr}) remains unpaid, now ${ctx.days_overdue} days overdue since ${ctx.due_date}. We understand things come up — if you'd like to discuss a payment plan, simply reply "PLAN" and we'll coordinate with you. Otherwise, here's the payment link: ${ctx.pay_link}. — ${ctx.sign_off}`;
    sms_body = `Hi ${ctx.first_name}, invoice ${ctx.invoice_number} (PHP ${amountStr}) is ${ctx.days_overdue} days overdue. Reply PLAN to discuss options, or pay: ${ctx.pay_link} — ${ctx.sign_off}`;
  }

    if (ctx.other_overdue_count > 0 && ctx.combined_total) {
    const otherLine = `Note: you also have ${ctx.other_overdue_count} other overdue invoice(s) with us, for a combined total overdue of PHP ${formatPeso_(ctx.combined_total)}.`;
    email_body = email_body.replace(` — ${ctx.sign_off}`, ` ${otherLine} — ${ctx.sign_off}`);
  }

  return { subject, email_body, sms_body };
}

/** ============ MANUAL TEST HARNESS ============ **/

function testAi() {
  const samples = [
    {
      stage: 2, first_name: 'Maria', invoice_number: 'INV-1002', amount: 30000,
      due_date: '2026-10-05', days_overdue: 15, other_overdue_count: 0,
      relationship: { tier: 'Regular', behavior: 'On-time', notes: '' },
      prior: { last_stage_sent: 1, days_since_last: 12, last_reply_intent: null },
      pay_link: 'https://pay.example.com/BDS/INV-1002', sign_off: 'BDS Team'
    },
    {
      stage: 1, first_name: 'Juan', invoice_number: 'INV-1001', amount: 45000,
      due_date: '2026-10-20', days_overdue: 0, other_overdue_count: 0,
      relationship: { tier: 'Key Account', behavior: 'On-time', notes: 'Long-time client, prefers email' },
      prior: { last_stage_sent: null, days_since_last: null, last_reply_intent: null },
      pay_link: 'https://pay.example.com/BDS/INV-1001', sign_off: 'BDS Team'
    },
    {
      stage: 4, first_name: 'Pedro', invoice_number: 'INV-1003', amount: 38000,
      due_date: '2026-09-30', days_overdue: 24, other_overdue_count: 1,
      relationship: { tier: 'Regular', behavior: 'Sometimes late', notes: '' },
      prior: { last_stage_sent: 3, days_since_last: 7, last_reply_intent: 'no_reply' },
      pay_link: 'https://pay.example.com/BDS/INV-1003', sign_off: 'BDS Team'
    }
  ];

  samples.forEach((ctx, i) => {
    Logger.log('--- Sample ' + (i + 1) + ' (stage ' + ctx.stage + ') ---');
    const result = generateMessage_(ctx);
    Logger.log(JSON.stringify(result, null, 2));
  });
}

/** ============ D3 MOCK-BASED TESTS ============ **/

// Helper: gumawa ng bagong context tuwing tumatawag, may random suffix sa invoice_number
// para HINDI tumama sa cache ng mga naunang test natin.
function freshTestCtx_(suffix) {
  suffix = suffix + '-' + Date.now();
  return {
    stage: 2, first_name: 'TestUser', invoice_number: 'INV-TEST-' + suffix, amount: 12345,
    due_date: '2026-10-01', days_overdue: 10, other_overdue_count: 0,
    relationship: { tier: 'Regular', behavior: 'On-time', notes: '' },
    prior: { last_stage_sent: 1, days_since_last: 5, last_reply_intent: null },
    pay_link: 'https://pay.example.com/BDS/INV-TEST-' + suffix, sign_off: 'BDS Team'
  };
}

// D3-T2: Gemini unavailable (maling API key) -> dapat AI_FALLBACK_USED, walang crash
function testD3T2() {
  const props = PropertiesService.getScriptProperties();
  const realKey = props.getProperty('GEMINI_API_KEY'); // i-save muna ang totoong key

  props.setProperty('GEMINI_API_KEY', 'INTENTIONALLY_BROKEN_KEY'); // sirain muna
  try {
    const ctx = freshTestCtx_('T2');
    const result = generateMessage_(ctx);
    Logger.log('D3-T2 result: ' + JSON.stringify(result));
    Logger.log('D3-T2 PASS check: source should be FALLBACK -> ' + (result.source === 'FALLBACK'));
  } finally {
    props.setProperty('GEMINI_API_KEY', realKey); // IBALIK ang totoong key, kahit mag-error
  }
}

// D3-T3: Unusable output (non-JSON) -> dapat mag-retry tapos fallback
function testD3T3() {
  const mockNonJson = function(prompt) {
    return 'Sure! Here you go'; // hindi JSON
  };
  const ctx = freshTestCtx_('T3');
  const result = generateMessage_(ctx, mockNonJson);
  Logger.log('D3-T3 result: ' + JSON.stringify(result));
  Logger.log('D3-T3 PASS check: source should be FALLBACK -> ' + (result.source === 'FALLBACK'));
}

// D3-T4: Exceeds length (300-word email) -> dapat EMAIL_TOO_LONG sa log, fallback
function testD3T4() {
  const ctx = freshTestCtx_('T4');
  const longEmail = 'word '.repeat(300).trim();
  const mockLongEmail = function(prompt) {
    return JSON.stringify({
      subject: 'Test',
      email_body: longEmail,
      sms_body: 'Invoice ' + ctx.invoice_number + ' PHP ' + ctx.amount + ' link: ' + ctx.pay_link
    });
  };
  const result = generateMessage_(ctx, mockLongEmail);
  Logger.log('D3-T4 result: ' + JSON.stringify(result));
  Logger.log('D3-T4 PASS check: source should be FALLBACK -> ' + (result.source === 'FALLBACK'));
}

// D3-T5: Hallucination guard (may "legal action" o maling amount) -> dapat ma-reject
function testD3T5() {
  const ctx = freshTestCtx_('T5');
  const mockBadContent = function(prompt) {
    return JSON.stringify({
      subject: 'Urgent',
      email_body: 'Invoice ' + ctx.invoice_number + ' PHP ' + ctx.amount +
        ' due ' + ctx.due_date + '. Pay now or we will pursue legal action. ' + ctx.pay_link,
      sms_body: 'Pay PHP ' + ctx.amount + ' now: ' + ctx.pay_link
    });
  };
  const result = generateMessage_(ctx, mockBadContent);
  Logger.log('D3-T5 result: ' + JSON.stringify(result));
  Logger.log('D3-T5 PASS check: source should be FALLBACK -> ' + (result.source === 'FALLBACK'));
}

// D3-T7: Cost cap (AI_DAILY_CAP=2, tawagin 3x) -> ika-3 dapat FALLBACK dahil sa cap
function testD3T7() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyyMMdd');
  const counterKey = 'AI_COUNT_' + day;
  const realCount = Number(props.getProperty(counterKey) || 0); // i-save ang totoong bilang (para sa cost audit)
  props.deleteProperty(counterKey);

  const configSheet = SpreadsheetApp.getActive().getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  let capRow = -1, realCap = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'AI_DAILY_CAP') { capRow = i; realCap = data[i][1]; break; }
  }
  configSheet.getRange(capRow + 1, 2).setValue(2);

  try {
    for (let i = 1; i <= 3; i++) {
      const ctx = freshTestCtx_('T7-' + i);
      const result = generateMessage_(ctx);
      Logger.log('D3-T7 call ' + i + ': source=' + result.source);
    }
    Logger.log('D3-T7 PASS check: ika-3 na call dapat FALLBACK dahil sa cap');
  } finally {
    configSheet.getRange(capRow + 1, 2).setValue(realCap);
    // Ibalik ang counter: totoong bilang + mga tawag na ginawa ng test na ito
    const usedInTest = Number(props.getProperty(counterKey) || 0);
    props.setProperty(counterKey, String(realCount + usedInTest));
  }
}