/**
 * Upperform Email Assistant — shared pure logic.
 *
 * Every function here is inlined verbatim into the n8n Code nodes by
 * scripts/build-workflows.mjs. It therefore must stay dependency-free and
 * must not touch Node APIs, the network, or the filesystem.
 */

/* ------------------------------------------------------------------ *
 * HTML -> text
 * ------------------------------------------------------------------ */

/** @type {Record<string, string>} */
const NAMED_ENTITIES = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  rsquo: '’',
  lsquo: '‘',
  rdquo: '”',
  ldquo: '“',
  mdash: '—',
  ndash: '–',
  hellip: '…',
  bull: '•',
  reg: '®',
  copy: '©',
  trade: '™',
  middot: '·',
};

/**
 * Decode the HTML entities that actually show up in Outlook mail bodies.
 * @param {string} input
 * @returns {string}
 */
function decodeEntities(input) {
  return String(input)
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex) =>
      String.fromCodePoint(parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_m, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&([a-zA-Z]+);/g, (m, name) => {
      const key = String(name).toLowerCase();
      const hit = NAMED_ENTITIES[key];
      return hit === undefined ? m : hit;
    });
}

/**
 * Convert an HTML mail body to readable plain text.
 * @param {string} html
 * @returns {string}
 */
function htmlToText(html) {
  if (!html) return '';
  let out = String(html);

  out = out.replace(/<!--[\s\S]*?-->/g, ' ');
  out = out.replace(/<(script|style|head)\b[\s\S]*?<\/\1>/gi, ' ');
  out = out.replace(/<\/?(o:p|xml)[^>]*>/gi, ' ');

  out = out.replace(/<br\s*\/?>/gi, '\n');
  out = out.replace(/<\/(p|div|tr|h[1-6]|blockquote|table|ul|ol)\s*>/gi, '\n');
  out = out.replace(/<li[^>]*>/gi, '\n- ');
  out = out.replace(/<\/(td|th)\s*>/gi, '\t');
  out = out.replace(/<hr[^>]*>/gi, '\n----------\n');

  out = out.replace(/<[^>]+>/g, '');
  out = decodeEntities(out);

  out = out.replace(/\r\n?/g, '\n');
  out = out.replace(/ /g, ' ');
  out = out.replace(/[ \t]+\n/g, '\n');
  out = out.replace(/[ \t]{2,}/g, ' ');
  out = out.replace(/\n{3,}/g, '\n\n');

  return out.trim();
}

/* ------------------------------------------------------------------ *
 * Quoted-history trimming
 * ------------------------------------------------------------------ */

const SIMPLE_BOUNDARIES = [
  /^\s*-{2,}\s*original message\s*-{2,}\s*$/i,
  /^\s*-{2,}\s*forwarded message\s*-{2,}\s*$/i,
  /^\s*begin forwarded message:?\s*$/i,
  /^\s*_{10,}\s*$/,
  /^\s*on\b.{4,160}\bwrote:\s*$/i,
  /^\s*.{0,80}\bwrote on\b.{4,80}:\s*$/i,
];

const HEADER_FROM = /^\s*(?:>\s*)?from:\s*\S/i;
const HEADER_COMPANION = /^\s*(?:>\s*)?(sent|date|to|cc|subject|reply-to):/i;

/**
 * True when `lines[i]` opens an Outlook-style quoted header block
 * (`From:` followed closely by `Sent:`/`To:`/`Subject:`).
 * @param {string[]} lines
 * @param {number} i
 * @returns {boolean}
 */
function isQuotedHeaderBlock(lines, i) {
  const line = lines[i];
  if (line === undefined || !HEADER_FROM.test(line)) return false;
  let companions = 0;
  for (let j = i + 1; j <= i + 5 && j < lines.length; j += 1) {
    const next = lines[j];
    if (next === undefined) break;
    if (HEADER_COMPANION.test(next)) companions += 1;
  }
  return companions >= 2;
}

/**
 * True when `lines[i]` starts a run of `>` quoted lines.
 * @param {string[]} lines
 * @param {number} i
 * @returns {boolean}
 */
function isQuoteRun(lines, i) {
  let run = 0;
  for (let j = i; j < lines.length && j < i + 3; j += 1) {
    const next = lines[j];
    if (next === undefined) break;
    if (/^\s*>/.test(next)) run += 1;
    else break;
  }
  return run >= 2;
}

/**
 * Indices of every line that begins a quoted/forwarded section.
 * @param {string[]} lines
 * @returns {number[]}
 */
function findQuoteBoundaries(lines) {
  /** @type {number[]} */
  const found = [];
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (line === undefined) continue;
    const simple = SIMPLE_BOUNDARIES.some((re) => re.test(line));
    if (simple || isQuotedHeaderBlock(lines, i) || isQuoteRun(lines, i)) {
      found.push(i);
    }
  }
  return found;
}

const PASS_THROUGH_HEAD =
  /^(?:fyi|fya|fwd?|see below|please see below|pls see below|below|forwarding|passing (?:this )?along|for your (?:review|awareness|visibility|information|records)|thoughts|please review|pls review|any (?:questions|thoughts)|per below|as below|-+)[\s.!,:?-]*$/i;

/**
 * Collapse a candidate head to the sentence(s) that carry meaning, dropping a
 * trailing one-to-three word signature line ("Chuck", "Thanks Q").
 * @param {string} head
 * @returns {string}
 */
function normalizeHead(head) {
  const lines = String(head || '')
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);
  while (lines.length > 1) {
    const last = lines[lines.length - 1];
    if (last === undefined) break;
    const words = last.split(/\s+/);
    if (words.length <= 3 && !/[.?!:]$/.test(last)) lines.pop();
    else break;
  }
  return lines.join(' ').replace(/\s+/g, ' ').trim();
}

/**
 * True when the text above the quote is only a pointer at the quote
 * ("FYI", "see below") rather than a self-contained message. Those forwards
 * carry all their meaning below the fold.
 * @param {string} head
 * @returns {boolean}
 */
function isPassThroughHead(head) {
  const collapsed = normalizeHead(head);
  if (collapsed.length === 0) return true;
  if (collapsed.length > 60) return false;
  return PASS_THROUGH_HEAD.test(collapsed);
}

/**
 * Keep only the newest message in a reply/forward chain.
 *
 * A bare forward ("FYI" + 12 quoted messages) carries all of its meaning in
 * the quoted part, so when the head is only a pointer we extend past the next
 * boundary rather than handing the classifier a meaningless string. A short
 * but self-contained head ("Sounds good.") is kept as-is — that IS the message.
 *
 * @param {string} text
 * @param {{ maxExtensions?: number }} [options]
 * @returns {{ text: string, trimmed: boolean, boundaries: number, extended: number }}
 */
function stripQuotedHistory(text, options) {
  const opts = options || {};
  const maxExtensions = opts.maxExtensions === undefined ? 2 : opts.maxExtensions;

  const source = String(text || '');
  const lines = source.split('\n');
  const boundaries = findQuoteBoundaries(lines);

  if (boundaries.length === 0) {
    return { text: source.trim(), trimmed: false, boundaries: 0, extended: 0 };
  }

  let cutIndex = 0;
  let extended = 0;
  let head = '';

  while (cutIndex < boundaries.length) {
    const boundary = boundaries[cutIndex];
    if (boundary === undefined) break;
    head = lines.slice(0, boundary).join('\n').trim();
    if (!isPassThroughHead(head)) break;
    if (extended >= maxExtensions) break;
    cutIndex += 1;
    extended += 1;
  }

  if (cutIndex >= boundaries.length) {
    return {
      text: source.trim(),
      trimmed: false,
      boundaries: boundaries.length,
      extended,
    };
  }

  const cleaned = head.replace(/\n{3,}/g, '\n\n').trim();
  return {
    text: cleaned.length > 0 ? cleaned : source.trim(),
    trimmed: cleaned.length > 0,
    boundaries: boundaries.length,
    extended,
  };
}

/**
 * Drop the boilerplate legal footer utilities append to outbound mail.
 * @param {string} text
 * @returns {string}
 */
function stripLegalFooter(text) {
  const markers = [
    /^\s*confidentiality\s+notice/im,
    /^\s*this\s+(e-?mail|message)\s+(and\s+any\s+attachments\s+)?(is|are|may\s+be)\s+(confidential|intended)/im,
    /^\s*the\s+information\s+contained\s+in\s+this\s+(e-?mail|message)\s+is\s+confidential/im,
    /^\s*disclaimer:\s*$/im,
  ];
  let cut = String(text || '');
  for (const marker of markers) {
    const match = marker.exec(cut);
    if (match && match.index > 20) cut = cut.slice(0, match.index);
  }
  return cut.trim();
}

/**
 * @param {string} text
 * @param {number} maxChars
 * @returns {{ text: string, truncated: boolean }}
 */
function truncate(text, maxChars) {
  const source = String(text || '');
  if (source.length <= maxChars) return { text: source, truncated: false };
  return {
    text: source.slice(0, maxChars) + '\n\n[...truncated...]',
    truncated: true,
  };
}

/* ------------------------------------------------------------------ *
 * Graph message normalisation
 * ------------------------------------------------------------------ */

/**
 * @param {any} recipients
 * @returns {string[]}
 */
function recipientAddresses(recipients) {
  if (!Array.isArray(recipients)) return [];
  /** @type {string[]} */
  const out = [];
  for (const entry of recipients) {
    const address =
      entry && entry.emailAddress && entry.emailAddress.address
        ? String(entry.emailAddress.address)
        : '';
    if (address) out.push(address.toLowerCase());
  }
  return out;
}

/**
 * Turn a raw Graph message into the shape every downstream node expects.
 * @param {any} msg
 * @param {{ maxChars?: number }} [options]
 */
function normalizeGraphMessage(msg, options) {
  const opts = options || {};
  const maxChars = opts.maxChars === undefined ? 6000 : opts.maxChars;

  const body = msg && msg.body ? msg.body : {};
  const contentType = String(body.contentType || '').toLowerCase();
  const rawContent = String(body.content || msg.bodyPreview || '');

  const plain =
    contentType === 'html'
      ? htmlToText(rawContent)
      : rawContent.replace(/\r\n?/g, '\n').trim();
  const withoutFooter = stripLegalFooter(plain);
  const stripped = stripQuotedHistory(withoutFooter);
  const limited = truncate(stripped.text, maxChars);

  const from = msg && msg.from && msg.from.emailAddress ? msg.from.emailAddress : {};
  const senderEmail = String(from.address || '').toLowerCase();
  const senderName = String(from.name || '').trim();

  return {
    messageId: String((msg && msg.id) || ''),
    conversationId: String((msg && msg.conversationId) || ''),
    internetMessageId: String((msg && msg.internetMessageId) || ''),
    receivedAt: String((msg && msg.receivedDateTime) || ''),
    senderEmail,
    senderName,
    subject: String((msg && msg.subject) || '(no subject)'),
    toRecipients: recipientAddresses(msg && msg.toRecipients),
    ccRecipients: recipientAddresses(msg && msg.ccRecipients),
    hasAttachments: Boolean(msg && msg.hasAttachments),
    webLink: String((msg && msg.webLink) || ''),
    categories: Array.isArray(msg && msg.categories) ? msg.categories.slice() : [],
    bodyContentType: contentType || 'text',
    cleanText: limited.text,
    // The whole visible thread, quoted history included, for the deterministic
    // guards only. A fee ask one level down the chain is still a fee ask —
    // scanning just the newest message misses it. Never sent to the model.
    threadText: truncate(withoutFooter, 12000).text,
    quoteTrimmed: stripped.trimmed,
    quoteBoundaries: stripped.boundaries,
    truncated: limited.truncated,
  };
}

/**
 * The user-turn payload handed to the classifier.
 * @param {ReturnType<typeof normalizeGraphMessage>} n
 * @returns {string}
 */
function buildClassifierInput(n) {
  const lines = [
    'SUBJECT: ' + n.subject,
    'FROM: ' + (n.senderName ? n.senderName + ' <' + n.senderEmail + '>' : n.senderEmail),
    'TO: ' + (n.toRecipients.length ? n.toRecipients.join(', ') : '(none)'),
    'CC: ' + (n.ccRecipients.length ? n.ccRecipients.join(', ') : '(none)'),
    'RECEIVED: ' + n.receivedAt,
    'HAS_ATTACHMENTS: ' + (n.hasAttachments ? 'yes' : 'no'),
    'FORWARD_CHAIN_DEPTH: ' + n.quoteBoundaries,
    '',
    '--- EMAIL BODY (newest message only) ---',
    n.cleanText || '(empty body)',
  ];
  return lines.join('\n');
}

/* ------------------------------------------------------------------ *
 * Classification
 * ------------------------------------------------------------------ */

const VALID_LABELS = ['opportunity', 'admin', 'noise', 'flag_for_yancey'];
const VALID_URGENCY = ['today', 'this_week', 'fyi'];

/**
 * Pull a JSON object out of whatever wrapper the OpenAI node returns.
 * @param {any} value
 * @returns {any}
 */
function parseModelJson(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'object' && !Array.isArray(value)) {
    if (typeof value.label === 'string') return value;
    const candidateKeys = [
      'output_parsed',
      'parsed',
      'output',
      'message',
      'content',
      'text',
      'output_text',
      'response',
      'data',
    ];
    for (const key of candidateKeys) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        const nested = parseModelJson(value[key]);
        if (nested) return nested;
      }
    }
    return null;
  }
  if (Array.isArray(value)) {
    for (const entry of value) {
      const nested = parseModelJson(entry);
      if (nested) return nested;
    }
    return null;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim().replace(/^```(?:json)?\s*/i, '').replace(/```$/, '');
    const start = trimmed.indexOf('{');
    const end = trimmed.lastIndexOf('}');
    if (start === -1 || end <= start) return null;
    try {
      return JSON.parse(trimmed.slice(start, end + 1));
    } catch (err) {
      return null;
    }
  }
  return null;
}

/**
 * Pull plain generated text out of whatever wrapper the OpenAI node returns.
 * @param {any} value
 * @param {number} [depth]
 * @returns {string}
 */
function extractModelText(value, depth) {
  const level = depth === undefined ? 0 : depth;
  if (level > 6 || value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) {
    /** @type {string[]} */
    const parts = [];
    for (const entry of value) {
      const text = extractModelText(entry, level + 1);
      if (text) parts.push(text);
    }
    return parts.join('\n').trim();
  }
  if (typeof value === 'object') {
    const keys = [
      'output_text',
      'text',
      'content',
      'message',
      'output',
      'response',
      'data',
      'choices',
    ];
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        const text = extractModelText(value[key], level + 1);
        if (text) return text;
      }
    }
  }
  return '';
}

/**
 * @param {unknown} value
 * @param {number} fallback
 * @returns {number}
 */
function clampConfidence(value, fallback) {
  const num = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(num)) return fallback;
  if (num < 0) return 0;
  if (num > 1) return num > 1 && num <= 100 ? num / 100 : 1;
  return Math.round(num * 100) / 100;
}

const PRICING_ASK =
  /\b(fee|fees|price|pricing|quote|quotation|cost|costs|estimate|budget|rate card|rates|retainer|proposal amount|how much|ballpark|order of magnitude|not to exceed|nte|scope of work|statement of work|sow|level of effort|loe|time and materials|fixed[- ]fee|day rate|bid)\b/i;
const PRICING_QUESTION = /\b(what would|what will|what does|how much|can you (provide|give|share)|could you (provide|give|share)|send (us|me) (a|your)|provide an? (estimate|indication))\b/i;
const CURRENCY_IN_TEXT = /(?:[$€£]\s?\d|\b\d[\d,]*\s?(?:k|m)\b|\busd\b|\bdollars?\b|\bper hour\b|\bhourly rate\b)/i;

/**
 * Deterministic backstop for the one failure mode that costs real money:
 * the agent originating or estimating a fee. Spec §11 Gate 3.
 *
 * Pass the WHOLE thread here, not the quote-stripped newest message. Found by
 * Gate 1 on Yancey's own mail: NIPSCO asked "can you provide an estimate of
 * timeline and cost?", and by the time Chuck replied "let's write a proposal"
 * the fee ask was one level down the quoted chain. Scanning only the newest
 * message let it through and the agent drafted a reply about a proposal.
 *
 * @param {string} text
 * @param {string} subject
 * @returns {{ triggered: boolean, reason: string }}
 */
function pricingGuard(text, subject) {
  const haystack = String(subject || '') + '\n' + String(text || '');
  const asksPrice = PRICING_ASK.test(haystack) && PRICING_QUESTION.test(haystack);
  if (!asksPrice) return { triggered: false, reason: '' };
  if (CURRENCY_IN_TEXT.test(haystack)) {
    return {
      triggered: false,
      reason: 'pricing discussed but a figure already exists in-thread',
    };
  }
  return {
    triggered: true,
    reason: 'asks for a fee/estimate and no figure exists in-thread',
  };
}

const SENSITIVE_TOPICS =
  /\b(lawsuit|litigation|subpoena|attorney|legal counsel|breach of contract|termination|severance|grievance|harassment|discrimination|press inquiry|journalist|reporter|defamation|non-?compete|indemnif|liability claim|settlement|arbitration|osha|whistleblow)\b/i;

/**
 * @param {string} senderEmail
 * @param {string} text
 * @param {string} subject
 * @param {string[]} alwaysFlagSenders
 * @returns {{ triggered: boolean, reason: string }}
 */
function sensitivityGuard(senderEmail, text, subject, alwaysFlagSenders) {
  const sender = String(senderEmail || '').toLowerCase();
  const flagged = Array.isArray(alwaysFlagSenders) ? alwaysFlagSenders : [];
  for (const entry of flagged) {
    const needle = String(entry || '').toLowerCase().trim();
    if (!needle) continue;
    if (sender === needle || (needle.startsWith('@') && sender.endsWith(needle))) {
      return { triggered: true, reason: 'sender is on the always-flag list' };
    }
  }
  const haystack = String(subject || '') + '\n' + String(text || '');
  if (SENSITIVE_TOPICS.test(haystack)) {
    return { triggered: true, reason: 'legal/HR/press subject matter detected' };
  }
  return { triggered: false, reason: '' };
}

/**
 * Apply every hard invariant the spec puts on a classification, whatever the
 * model actually returned.
 *
 * @param {any} raw
 * @param {{ cleanText: string, subject: string, senderEmail: string,
 *           confidenceMin: number, alwaysFlagSenders?: string[],
 *           threadText?: string }} ctx
 */
function normalizeClassification(raw, ctx) {
  const parsed = parseModelJson(raw) || {};
  /** @type {string[]} */
  const guards = [];

  let label = String(parsed.label || '').toLowerCase().trim();
  if (VALID_LABELS.indexOf(label) === -1) {
    label = 'flag_for_yancey';
    guards.push('unparseable_label');
  }

  let confidence = clampConfidence(parsed.confidence, 0);
  const urgency =
    VALID_URGENCY.indexOf(String(parsed.urgency || '').toLowerCase()) === -1
      ? 'fyi'
      : String(parsed.urgency).toLowerCase();

  const asks = Array.isArray(parsed.asks_of_yancey)
    ? parsed.asks_of_yancey.map((/** @type {unknown} */ a) => String(a)).filter(Boolean).slice(0, 6)
    : [];

  const pricing = pricingGuard(ctx.threadText || ctx.cleanText, ctx.subject);
  if (pricing.triggered) {
    label = 'flag_for_yancey';
    guards.push('pricing_guard');
  }

  const sensitivity = sensitivityGuard(
    ctx.senderEmail,
    ctx.cleanText,
    ctx.subject,
    ctx.alwaysFlagSenders || [],
  );
  if (sensitivity.triggered) {
    label = 'flag_for_yancey';
    guards.push('sensitivity_guard');
  }

  if (guards.length > 0) confidence = Math.max(confidence, 0.9);

  const draftable = label === 'opportunity' || label === 'admin';
  const modelWantsDraft = parsed.requires_draft !== false;
  const requiresDraft =
    draftable && modelWantsDraft && confidence >= ctx.confidenceMin;

  if (draftable && modelWantsDraft && !requiresDraft) {
    guards.push('below_confidence_min');
  }

  return {
    label,
    confidence,
    urgency,
    reasoning: String(parsed.reasoning || '').slice(0, 500),
    asks,
    requiresDraft,
    guards,
    guardFlags: guards.join(','),
  };
}

/** @type {Record<string, string>} */
const LABEL_CATEGORY = {
  opportunity: 'AI-Action',
  admin: 'AI-Action',
  flag_for_yancey: 'AI-Review',
  noise: '',
};

/**
 * Graph PATCH replaces the whole categories array, so keep what is there.
 * @param {string[]} existing
 * @param {string} label
 * @returns {string[]}
 */
function decideCategories(existing, label) {
  const base = Array.isArray(existing) ? existing.slice() : [];
  const managed = ['AI-Processed', 'AI-Action', 'AI-Review', 'AI-Drafted', 'AI-Error'];
  const kept = base.filter((c) => managed.indexOf(String(c)) === -1);
  kept.push('AI-Processed');
  const mapped = LABEL_CATEGORY[label];
  const extra = mapped === undefined ? 'AI-Review' : mapped;
  if (extra) kept.push(extra);
  return kept;
}

/**
 * @param {string[]} existing
 * @param {string} category
 * @returns {string[]}
 */
function addCategory(existing, category) {
  const base = Array.isArray(existing) ? existing.slice() : [];
  if (base.indexOf(category) === -1) base.push(category);
  return base;
}

/* ------------------------------------------------------------------ *
 * Drafting
 * ------------------------------------------------------------------ */

/**
 * @param {string} value
 * @returns {string}
 */
function escapeHtml(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * @param {string} text
 * @returns {string}
 */
function textToHtmlParagraphs(text) {
  const blocks = String(text || '')
    .replace(/\r\n?/g, '\n')
    .split(/\n{2,}/)
    .map((b) => b.trim())
    .filter(Boolean);
  if (blocks.length === 0) return '<p></p>';
  return blocks
    .map((block) => '<p>' + escapeHtml(block).replace(/\n/g, '<br>') + '</p>')
    .join('\n');
}

/**
 * Put the generated reply above the history Graph already quoted for us,
 * so the draft still reads as a proper reply.
 *
 * @param {string} draftHtml
 * @param {string} replyHtml
 * @returns {string}
 */
function injectReplyIntoDraftHtml(draftHtml, replyHtml) {
  const source = String(draftHtml || '');
  const block = String(replyHtml || '');
  if (!source.trim()) return block;
  const bodyOpen = /<body[^>]*>/i.exec(source);
  if (bodyOpen && bodyOpen.index !== undefined) {
    const at = bodyOpen.index + bodyOpen[0].length;
    return source.slice(0, at) + '\n' + block + '\n' + source.slice(at);
  }
  return block + '\n<br>\n' + source;
}

const FORBIDDEN_OPENERS = [
  /i hope this (e-?mail|message) finds you well[.,!]?/gi,
  /i hope you(?:'re| are) doing well[.,!]?/gi,
  /i hope this note finds you well[.,!]?/gi,
  /i trust this (e-?mail|message) finds you well[.,!]?/gi,
];

const CURRENCY_TOKEN = /[$€£]\s?\d[\d,]*(?:\.\d+)?\s*(?:k|m|mm|million|thousand)?|\b\d[\d,]*(?:\.\d+)?\s?(?:usd|dollars?)\b/gi;
const PERCENT_TOKEN = /\b\d[\d,]*(?:\.\d+)?\s?%/g;
const DATE_TOKEN =
  /\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?\b|\b\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/gi;
const BIG_NUMBER_TOKEN = /\b\d{1,3}(?:,\d{3})+\b|\b\d{4,}\b/g;

/**
 * @param {string} token
 * @returns {string}
 */
function digitsOf(token) {
  return String(token).replace(/\D/g, '');
}

/**
 * @param {string} text
 * @param {RegExp} pattern
 * @returns {string[]}
 */
function matchAll(text, pattern) {
  const re = new RegExp(pattern.source, pattern.flags.indexOf('g') === -1 ? pattern.flags + 'g' : pattern.flags);
  /** @type {string[]} */
  const out = [];
  let m = re.exec(String(text || ''));
  while (m !== null) {
    out.push(m[0]);
    if (m.index === re.lastIndex) re.lastIndex += 1;
    m = re.exec(String(text || ''));
  }
  return out;
}

/**
 * @param {string} draft
 * @param {string} source
 * @param {RegExp} pattern
 * @returns {string[]}
 */
function inventedTokens(draft, source, pattern) {
  const sourceDigits = matchAll(source, pattern).map(digitsOf);
  const sourceRaw = String(source || '').toLowerCase();
  /** @type {string[]} */
  const invented = [];
  for (const token of matchAll(draft, pattern)) {
    const digits = digitsOf(token);
    if (!digits) continue;
    if (sourceDigits.indexOf(digits) !== -1) continue;
    if (sourceRaw.indexOf(String(token).toLowerCase()) !== -1) continue;
    if (sourceRaw.replace(/\D/g, '').indexOf(digits) !== -1) continue;
    if (invented.indexOf(token) === -1) invented.push(token);
  }
  return invented;
}

/**
 * @param {string} text
 * @returns {number}
 */
function wordCount(text) {
  const words = String(text || '').trim().split(/\s+/).filter(Boolean);
  return words.length;
}

/**
 * Gate 2 in code: repair what is safe to repair, block what is not.
 *
 * hard  -> do not create a draft at all (the message is tagged AI-Review)
 * medium-> create the draft but also tag AI-Review so Yancey looks closer
 * soft  -> silently repaired here
 *
 * @param {string} draftText
 * @param {string} sourceText
 * @param {{ signOff?: string, signature?: string, maxWords?: number }} [options]
 */
function voiceGuard(draftText, sourceText, options) {
  const opts = options || {};
  const signOff = opts.signOff === undefined ? 'Thanks,' : opts.signOff;
  const signature = opts.signature === undefined ? 'Yancey' : opts.signature;
  const maxWords = opts.maxWords === undefined ? 220 : opts.maxWords;

  /** @type {string[]} */
  const hard = [];
  /** @type {string[]} */
  const medium = [];
  /** @type {string[]} */
  const soft = [];

  let body = String(draftText || '').trim();
  body = body.replace(/^```(?:\w+)?\s*/i, '').replace(/```\s*$/i, '').trim();
  body = body.replace(/^(?:subject|re)\s*:.*\n+/i, '').trim();

  if (!body) {
    hard.push('empty_draft');
    return { ok: false, hard, medium, soft, text: '', words: 0 };
  }

  const invented = inventedTokens(body, sourceText, CURRENCY_TOKEN);
  if (invented.length > 0) hard.push('invented_currency:' + invented.join('|'));

  const inventedPct = inventedTokens(body, sourceText, PERCENT_TOKEN);
  if (inventedPct.length > 0) hard.push('invented_percent:' + inventedPct.join('|'));

  const inventedDates = inventedTokens(body, sourceText, DATE_TOKEN);
  if (inventedDates.length > 0) medium.push('invented_date:' + inventedDates.join('|'));

  const inventedNumbers = inventedTokens(body, sourceText, BIG_NUMBER_TOKEN);
  if (inventedNumbers.length > 0) medium.push('invented_number:' + inventedNumbers.join('|'));

  for (const opener of FORBIDDEN_OPENERS) {
    if (opener.test(body)) {
      body = body.replace(opener, '').replace(/^\s*\n/, '').trim();
      soft.push('removed_forbidden_opener');
    }
  }

  body = body.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();

  const hasSignOff = new RegExp('(^|\\n)\\s*' + signOff.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(body);
  const hasSignature = new RegExp('(^|\\n)\\s*' + signature.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$', 'i').test(body);

  if (!hasSignOff || !hasSignature) {
    body = body.replace(/\n*\s*(thanks|thank you|warm regards|best regards|regards|best)\s*,?\s*\n*\s*(yancey(\s+dawson)?)?\s*$/i, '').trim();
    body = body + '\n\n' + signOff + '\n\n' + signature;
    soft.push('appended_signoff');
  }

  const words = wordCount(body);
  if (words > maxWords) medium.push('over_length:' + words);

  return { ok: hard.length === 0, hard, medium, soft, text: body, words };
}

/* ------------------------------------------------------------------ *
 * Digest
 * ------------------------------------------------------------------ */

/**
 * @param {string | Date | null | undefined} value
 * @returns {Date | null}
 */
function toDate(value) {
  if (!value) return null;
  const d = value instanceof Date ? value : new Date(String(value));
  return Number.isNaN(d.getTime()) ? null : d;
}

/**
 * @param {Date} from
 * @param {number} days
 * @returns {Date}
 */
function addBusinessDays(from, days) {
  const out = new Date(from.getTime());
  let remaining = days;
  while (remaining > 0) {
    out.setUTCDate(out.getUTCDate() + 1);
    const day = out.getUTCDay();
    if (day !== 0 && day !== 6) remaining -= 1;
  }
  return out;
}

/**
 * @param {Date} a
 * @param {Date} b
 * @returns {number}
 */
function businessDaysBetween(a, b) {
  if (b.getTime() <= a.getTime()) return 0;
  let count = 0;
  const cursor = new Date(a.getTime());
  let guard = 0;
  while (guard < 4000) {
    guard += 1;
    cursor.setUTCDate(cursor.getUTCDate() + 1);
    if (cursor.getTime() > b.getTime()) break;
    const day = cursor.getUTCDay();
    if (day !== 0 && day !== 6) count += 1;
  }
  return count;
}

/**
 * Split the log into the three digest sections and work out which rows the
 * mailbox says are already dealt with.
 *
 * @param {any[]} rows
 * @param {{ now: string | Date, liveDraftIds: string[], followupDays: number,
 *           lookbackHours?: number, autoCloseDays?: number }} options
 */
function bucketDigestRows(rows, options) {
  const now = toDate(options.now) || new Date();
  const lookbackHours = options.lookbackHours === undefined ? 24 : options.lookbackHours;
  const autoCloseDays = options.autoCloseDays === undefined ? 21 : options.autoCloseDays;
  const autoCloseBefore = new Date(now.getTime() - autoCloseDays * 86400 * 1000);
  const since = new Date(now.getTime() - lookbackHours * 3600 * 1000);
  /** @type {Record<string, boolean>} */
  const live = {};
  for (const id of options.liveDraftIds || []) live[String(id)] = true;

  /** @type {any[]} */
  const needsYou = [];
  /** @type {any[]} */
  const draftsWaiting = [];
  /** @type {any[]} */
  const followupsDue = [];
  /** @type {string[]} */
  const resolvedIds = [];
  /** @type {string[]} */
  const resolvedMessageIds = [];
  let autoClosed = 0;

  for (const row of rows || []) {
    if (!row || !row.message_id) continue;
    if (row.resolved_at) continue;

    const received = toDate(row.received_at) || now;
    const draftId = row.draft_id ? String(row.draft_id) : '';
    const status = String(row.status || '');

    if (!draftId && received.getTime() < autoCloseBefore.getTime()) {
      resolvedIds.push(String(row.id));
      resolvedMessageIds.push(String(row.message_id));
      autoClosed += 1;
      continue;
    }

    if (draftId && !live[draftId]) {
      resolvedIds.push(String(row.id));
      resolvedMessageIds.push(String(row.message_id));
      continue;
    }

    if (draftId) {
      draftsWaiting.push({
        subject: row.subject,
        sender: row.sender_name || row.sender,
        urgency: row.urgency,
        waitingDays: businessDaysBetween(toDate(row.draft_created_at) || received, now),
        webLink: row.draft_web_link || row.web_link,
        guardFlags: row.guard_flags || '',
      });
      continue;
    }

    const isOpenAction =
      row.label === 'flag_for_yancey' ||
      (row.label === 'opportunity' && status !== 'drafted') ||
      (row.label === 'admin' && status !== 'drafted');

    if (isOpenAction && received.getTime() >= since.getTime()) {
      needsYou.push({
        subject: row.subject,
        sender: row.sender_name || row.sender,
        label: row.label,
        urgency: row.urgency,
        reasoning: row.reasoning,
        asks: row.asks ? String(row.asks) : '',
        guardFlags: row.guard_flags || '',
        webLink: row.web_link,
      });
      continue;
    }

    if (row.label === 'opportunity' || row.label === 'flag_for_yancey') {
      const due = toDate(row.followup_due_at) || addBusinessDays(received, options.followupDays);
      if (due.getTime() <= now.getTime()) {
        followupsDue.push({
          subject: row.subject,
          sender: row.sender_name || row.sender,
          quietDays: businessDaysBetween(received, now),
          webLink: row.web_link,
        });
      }
    }
  }

  /** @type {Record<string, number>} */
  const urgencyRank = { today: 0, this_week: 1, fyi: 2 };
  needsYou.sort((a, b) => {
    const ra = urgencyRank[a.urgency];
    const rb = urgencyRank[b.urgency];
    return (ra === undefined ? 3 : ra) - (rb === undefined ? 3 : rb);
  });
  followupsDue.sort((a, b) => b.quietDays - a.quietDays);

  return {
    needsYou,
    draftsWaiting,
    followupsDue,
    resolvedIds,
    resolvedMessageIds,
    stats: {
      needsYou: needsYou.length,
      draftsWaiting: draftsWaiting.length,
      followupsDue: followupsDue.length,
      resolved: resolvedIds.length,
      autoClosed,
      scanned: (rows || []).length,
    },
  };
}

/**
 * @param {string} label
 * @returns {string}
 */
function labelBadge(label) {
  if (label === 'flag_for_yancey') return 'FLAGGED — no draft written';
  if (label === 'opportunity') return 'Opportunity';
  if (label === 'admin') return 'Admin';
  return String(label || '');
}

/**
 * @param {{ title: string, emptyText: string, rows: string[] }} section
 * @returns {string}
 */
function renderSection(section) {
  const inner =
    section.rows.length > 0
      ? section.rows.join('\n')
      : '<p style="margin:0;color:#6b7280;font-size:14px;">' +
        escapeHtml(section.emptyText) +
        '</p>';
  return [
    '<h2 style="margin:28px 0 10px;font-size:15px;font-weight:600;letter-spacing:.04em;text-transform:uppercase;color:#111827;">' +
      escapeHtml(section.title) +
      '</h2>',
    inner,
  ].join('\n');
}

/**
 * @param {string} subject
 * @param {string} sender
 * @param {string} meta
 * @param {string} detail
 * @param {string} webLink
 * @returns {string}
 */
function renderCard(subject, sender, meta, detail, webLink) {
  const link = webLink
    ? '<a href="' +
      escapeHtml(webLink) +
      '" style="color:#2563eb;text-decoration:none;font-size:13px;">Open in Outlook &rarr;</a>'
    : '';
  return [
    '<div style="border:1px solid #e5e7eb;border-radius:8px;padding:12px 14px;margin:0 0 8px;">',
    '<div style="font-weight:600;font-size:15px;color:#111827;">' + escapeHtml(subject || '(no subject)') + '</div>',
    '<div style="font-size:13px;color:#6b7280;margin-top:2px;">' + escapeHtml(sender || '') + (meta ? ' &middot; ' + escapeHtml(meta) : '') + '</div>',
    detail ? '<div style="font-size:13px;color:#374151;margin-top:6px;">' + escapeHtml(detail) + '</div>' : '',
    link ? '<div style="margin-top:8px;">' + link + '</div>' : '',
    '</div>',
  ]
    .filter(Boolean)
    .join('\n');
}

/**
 * @param {ReturnType<typeof bucketDigestRows>} buckets
 * @param {{ dateLabel: string, mailboxLabel?: string, mode?: string, errors?: any[] }} options
 * @returns {string}
 */
function renderDigestHtml(buckets, options) {
  const errors = options.errors || [];

  const needsYouRows = buckets.needsYou.map((r) => {
    const asks = r.asks ? String(r.asks).replace(/^\[|\]$/g, '').replace(/","/g, '; ').replace(/"/g, '') : '';
    const detail = [r.reasoning, asks ? 'Asks: ' + asks : ''].filter(Boolean).join(' — ');
    const meta = [labelBadge(r.label), r.urgency, r.guardFlags].filter(Boolean).join(' · ');
    return renderCard(r.subject, r.sender, meta, detail, r.webLink);
  });

  const draftRows = buckets.draftsWaiting.map((r) => {
    const meta = [
      r.waitingDays > 0 ? 'waiting ' + r.waitingDays + ' business day' + (r.waitingDays === 1 ? '' : 's') : 'new',
      r.urgency,
      r.guardFlags ? 'check: ' + r.guardFlags : '',
    ]
      .filter(Boolean)
      .join(' · ');
    return renderCard(r.subject, r.sender, meta, '', r.webLink);
  });

  const followupRows = buckets.followupsDue.map((r) =>
    renderCard(r.subject, r.sender, 'quiet ' + r.quietDays + ' business days', '', r.webLink),
  );

  const errorRows = errors.map((e) =>
    renderCard(
      String(e.workflow_name || 'workflow'),
      String(e.node_name || ''),
      String(e.occurred_at || ''),
      String(e.error_message || '').slice(0, 240),
      '',
    ),
  );

  const modeBanner =
    options.mode && options.mode !== 'live'
      ? '<div style="background:#fef3c7;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;margin:0 0 18px;font-size:13px;color:#92400e;">Shadow mode: the assistant is labelling only. No replies are being drafted.</div>'
      : '';

  return [
    '<div style="font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;max-width:640px;margin:0 auto;padding:24px;color:#111827;">',
    '<div style="font-size:13px;color:#6b7280;letter-spacing:.04em;text-transform:uppercase;">Upperform inbox assistant</div>',
    '<h1 style="margin:4px 0 2px;font-size:22px;font-weight:600;">Daily digest</h1>',
    '<div style="font-size:14px;color:#6b7280;margin-bottom:18px;">' +
      escapeHtml(options.dateLabel) +
      (options.mailboxLabel ? ' &middot; ' + escapeHtml(options.mailboxLabel) : '') +
      '</div>',
    modeBanner,
    '<div style="display:block;font-size:13px;color:#374151;background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px 14px;">' +
      buckets.stats.needsYou +
      ' need you &middot; ' +
      buckets.stats.draftsWaiting +
      ' drafts waiting &middot; ' +
      buckets.stats.followupsDue +
      ' follow-ups due &middot; ' +
      buckets.stats.resolved +
      ' closed since yesterday</div>',
    renderSection({
      title: 'Needs you',
      emptyText: 'Nothing new needs your attention.',
      rows: needsYouRows,
    }),
    renderSection({
      title: 'Drafts waiting in your Drafts folder',
      emptyText: 'No drafts waiting.',
      rows: draftRows,
    }),
    renderSection({
      title: 'Follow-ups due',
      emptyText: 'No threads have gone quiet past the threshold.',
      rows: followupRows,
    }),
    errorRows.length > 0
      ? renderSection({ title: 'Errors in the last 24h', emptyText: '', rows: errorRows })
      : '',
    '<p style="margin:28px 0 0;font-size:12px;color:#9ca3af;border-top:1px solid #e5e7eb;padding-top:12px;">The assistant can read your mail and write drafts. It cannot send email — every reply leaves your mailbox only when you press Send.</p>',
    '</div>',
  ]
    .filter(Boolean)
    .join('\n');
}

/**
 * @param {ReturnType<typeof bucketDigestRows>} buckets
 * @param {{ dateLabel: string }} options
 * @returns {string}
 */
function renderDigestText(buckets, options) {
  /** @type {string[]} */
  const lines = ['Upperform inbox assistant — daily digest', options.dateLabel, ''];
  lines.push('NEEDS YOU (' + buckets.stats.needsYou + ')');
  for (const r of buckets.needsYou) {
    lines.push('  - ' + (r.subject || '(no subject)') + ' — ' + (r.sender || '') + ' [' + labelBadge(r.label) + ']');
  }
  lines.push('', 'DRAFTS WAITING (' + buckets.stats.draftsWaiting + ')');
  for (const r of buckets.draftsWaiting) {
    lines.push('  - ' + (r.subject || '(no subject)') + ' — ' + (r.sender || ''));
  }
  lines.push('', 'FOLLOW-UPS DUE (' + buckets.stats.followupsDue + ')');
  for (const r of buckets.followupsDue) {
    lines.push('  - ' + (r.subject || '(no subject)') + ' — quiet ' + r.quietDays + ' business days');
  }
  lines.push('', 'The assistant cannot send email. Only you press Send.');
  return lines.join('\n');
}

export {
  addBusinessDays,
  isPassThroughHead,
  addCategory,
  bucketDigestRows,
  buildClassifierInput,
  businessDaysBetween,
  clampConfidence,
  decideCategories,
  decodeEntities,
  escapeHtml,
  extractModelText,
  findQuoteBoundaries,
  htmlToText,
  injectReplyIntoDraftHtml,
  inventedTokens,
  normalizeClassification,
  normalizeGraphMessage,
  parseModelJson,
  pricingGuard,
  renderDigestHtml,
  renderDigestText,
  sensitivityGuard,
  stripLegalFooter,
  stripQuotedHistory,
  textToHtmlParagraphs,
  truncate,
  voiceGuard,
  wordCount,
};
