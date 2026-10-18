/** ============ VALIDATOR (the gatekeeper) ============ **/

const BLACKLISTED_WORDS_ = ['lawsuit', 'legal action', 'blacklist', 'court', 'penalty', 'demand letter', 'sue'];

function countWords_(text) {
  return text.trim().split(/\s+/).filter(Boolean).length;
}

function formatPeso_(amount) {
  return Number(amount).toLocaleString('en-PH');
}

function extractAmounts_(text) {
  const matches = text.match(/(?:PHP|₱)\s?[\d,]+(?:\.\d+)?/gi) || [];
  return matches.map(m => m.replace(/[^\d.]/g, ''));
}

function extractDates_(text) {
  const iso = text.match(/\d{4}-\d{2}-\d{2}/g) || [];
  const longForm = text.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b/gi) || [];
  return iso.concat(longForm);
}

function validateAi_(out, ctx) {
  let parsed;
  try {
    parsed = JSON.parse(out);
  } catch (e) {
    throw new Error('INVALID_JSON');
  }

  if (typeof parsed.subject !== 'string' || typeof parsed.email_body !== 'string' || typeof parsed.sms_body !== 'string') {
    throw new Error('MISSING_REQUIRED_KEYS');
  }

  if (countWords_(parsed.email_body) > 120) throw new Error('EMAIL_TOO_LONG');
  if (parsed.sms_body.length > 320) throw new Error('SMS_TOO_LONG');

  const combined = parsed.subject + ' ' + parsed.email_body + ' ' + parsed.sms_body;

  if (combined.indexOf(ctx.invoice_number) === -1) throw new Error('MISSING_INVOICE_NUMBER');
  const normalizedCombined = combined.replace(/,/g, '');
  const normalizedAmount = String(ctx.amount);
  if (normalizedCombined.indexOf(normalizedAmount) === -1) throw new Error('MISSING_AMOUNT');
  if (combined.indexOf(ctx.pay_link) === -1) throw new Error('MISSING_PAY_LINK');

  for (const word of BLACKLISTED_WORDS_) {
    const pattern = new RegExp('\\b' + word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
    if (pattern.test(combined)) throw new Error('BLACKLISTED_WORD:' + word);
  }

  const foundAmounts = extractAmounts_(combined);
  const allowedAmounts = [Number(normalizedAmount)];
  if (ctx.combined_total) allowedAmounts.push(Number(ctx.combined_total));
  for (const amt of foundAmounts) {
    if (!allowedAmounts.includes(Number(amt.replace(/,/g, '')))) {
      throw new Error('UNAUTHORIZED_AMOUNT:' + amt);
    }
  }

  const foundDates = extractDates_(combined);
  const allowedDates = [ctx.due_date, ctx.respond_by_date].filter(Boolean);
  for (const d of foundDates) {
    if (!allowedDates.includes(d)) throw new Error('UNAUTHORIZED_DATE:' + d);
  }

  return parsed; // valid: the caller can trust it
}
