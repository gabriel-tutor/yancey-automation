import { describe, expect, it } from 'vitest';
import {
  CLIENT_POLICY_TOPICS,
  GUARD_PATTERNS as BASE_PATTERNS,
  policyTopicGuard,
  pricingGuard,
  sensitivityGuard,
} from './agent-lib.js';

/**
 * Base guards plus Yancey's named categories, composed here rather than in the
 * library so the library's bundles stay tree-shakeable. Built from the list, so
 * adding a category without adding probes fails this suite.
 * @type {Record<string, RegExp>}
 */
const GUARD_PATTERNS = Object.assign({}, BASE_PATTERNS);
for (const topic of CLIENT_POLICY_TOPICS) {
  GUARD_PATTERNS['TOPIC_' + topic.key.toUpperCase()] = topic.pattern;
}

/**
 * Every term in every guard vocabulary, checked against a sentence somebody
 * might actually write.
 *
 * Three defects shipped in a single day because nothing did this:
 *   - `indemnif` inside a trailing \b could never match "indemnification",
 *     the only form anyone writes. It had never fired once.
 *   - `whistleblow` had the same problem against "whistleblower".
 *   - "how much" sat in both PRICING_ASK and PRICING_QUESTION, so one phrase
 *     satisfied a rule designed to need two independent signals.
 *
 * A term with no probe fails the suite, so a new one cannot be added to a
 * guard without demonstrating that it works.
 */

/**
 * Split the top-level alternation out of a pattern, ignoring `|` inside
 * nested groups.
 * @param {RegExp} re
 * @returns {string[]}
 */
function alternativesOf(re) {
  const src = re.source;
  const open = src.indexOf('(');
  const close = src.lastIndexOf(')');
  const inner = src.slice(open + 1, close).replace(/^\?:/, '');
  /** @type {string[]} */
  const out = [];
  let depth = 0;
  let current = '';
  for (let i = 0; i < inner.length; i += 1) {
    const ch = inner[i];
    if (ch === '\\') { current += ch + (inner[i + 1] || ''); i += 1; continue; }
    if (ch === '(') depth += 1;
    if (ch === ')') depth -= 1;
    if (ch === '|' && depth === 0) { out.push(current); current = ''; continue; }
    current += ch;
  }
  if (current) out.push(current);
  return out;
}

/**
 * A realistic sentence for each term. The point is the *word form* people use.
 * @type {Record<string, Record<string, string>>}
 */
const PROBES = {
  PRICING_ASK: {
    'fee': 'What fee are we assuming for this?',
    'fees': 'Are the fees fixed for the year?',
    'price': 'What price did we put on the last one?',
    'pricing': 'Can you share the pricing for the extension?',
    'quote': 'Please send a quote for the assessment.',
    'quotation': 'We will need a formal quotation.',
    'cost': 'Can you provide an estimate of cost?',
    'costs': 'What are the costs for a second district?',
    'estimate': 'Can you provide an estimate for the work?',
    'budget': 'What budget should we assume?',
    'rate card': 'Do you have a rate card we can circulate?',
    'rates': 'Can you share your rates for next year?',
    'retainer': 'Would this run on a retainer?',
    'proposal amount': 'What proposal amount should I put to the board?',
    'ballpark': 'Can you give a ballpark for the pigging work?',
    'order of magnitude': 'An order of magnitude figure would help.',
    'not to exceed': 'We would need a not to exceed number.',
    'nte': 'Procurement wants an NTE on the purchase order.',
    'scope of work': 'Can you provide a scope of work for the review?',
    'statement of work': 'We will need a statement of work before signing.',
    'sow': 'Can you share the SOW for the T-Week engagement?',
    'level of effort': 'What level of effort are we talking about?',
    'loe': 'Can you provide the LOE for the assessment?',
    'time and materials': 'Would this be time and materials or fixed?',
    'fixed[- ]fee': 'Is this a fixed-fee engagement?',
    'day rate': 'What day rate applies to a Principal Consultant?',
    'bid': 'Are we submitting a bid for the framework?',
  },
  PRICING_QUESTION: {
    'what would': 'What would this take to stand up?',
    'what will': 'What will the northern district need?',
    'what does': 'What does a full assessment involve?',
    'how much': 'Roughly how much would the assessment cost?',
    'can you (provide|give|share)': 'Can you provide the figure by Friday?',
    'could you (provide|give|share)': 'Could you share the numbers before the SLT?',
    'send (us|me) (a|your)': 'Please send us a summary before Thursday.',
    'provide an? (estimate|indication)': 'Please provide an estimate of timeline and cost.',
  },
  FEE_STATEMENT: {
    'fixed[- ]fee': 'The range is $175K to $200K on a fixed-fee basis.',
    'indicative range': 'The indicative range is $175K to $200K.',
    'retainer': 'We hold a $12K monthly retainer for this.',
    'not to exceed': 'Billed not to exceed $200K in total.',
    'nte': 'The NTE on the order is $200K.',
    'per hour': 'Our blended rate is $265 per hour.',
    'hourly rate': 'The hourly rate on this engagement is $265.',
    'day rate': 'The day rate is $2,400 for a Principal.',
    'our fee': 'Our fee for the assessment is $175K.',
    'the fee': 'The fee covering all five deliverables is $175K.',
    'fee is': 'The fee is $175K on a fixed basis.',
    'fee of': 'A fee of $175K covers the scoping phase.',
    'fee covers': 'That fee covers $175K of embedded support.',
    'fee would be': 'The final fee would be $200K at scoping.',
    'quoted at': 'The engagement was quoted at $175K.',
    'budget of': 'They have a budget of $200K approved.',
  },
  SENSITIVE_TOPICS: {
    'lawsuit': 'The contractor has filed a lawsuit against the utility.',
    'litigation': 'This may end up in litigation.',
    'subpoena': 'We received a subpoena for the inspection records.',
    'attorney': 'Our attorney flagged the clause.',
    'legal counsel': 'Legal counsel wants to review this first.',
    'breach of contract': 'They are alleging breach of contract.',
    'termination': 'The termination clause needs review.',
    'severance': 'His severance package is still being negotiated.',
    'grievance': 'A grievance was filed with the union.',
    'harassment': 'A harassment complaint has been raised.',
    'discrimination': 'The claim alleges discrimination in promotion.',
    'press inquiry': 'We have had a press inquiry about the outage.',
    'journalist': 'A journalist is asking about the near-miss.',
    'reporter': 'A reporter called about the pigging incident.',
    'defamation': 'They are threatening a defamation claim.',
    'non-?compete': 'His non-compete runs another eighteen months.',
    'indemni\\w*': 'They added an indemnification obligation covering third-party claims.',
    'liability claim': 'A liability claim has been lodged against us.',
    'settlement': 'The settlement figure is confidential.',
    'arbitration': 'The dispute goes to arbitration next month.',
    'osha': 'OSHA is opening an inquiry into the incident.',
    'whistleblow\\w*': 'A whistleblower complaint was raised internally.',
  },

  COMMERCIAL_TERMS: {
    'fee': 'What is the fee for the second district?',
    'fees': 'Are the fees fixed for the year?',
    'pricing': 'Their pricing team wants a call.',
    'quotation': 'We will need a formal quotation.',
    'rate card': 'Do you have a rate card we can circulate?',
    'retainer': 'Would this run on a retainer?',
    'scope change': 'That is a scope change and needs sign-off.',
    'scope changes': 'Two scope changes came out of the workshop.',
    'change order': 'Please raise a change order for the extra week.',
    'statement of work': 'The statement of work is attached.',
    'scope of work': 'Can you review the scope of work?',
    'sow': 'Is the SOW signed yet?',
    'msa': 'Is this work under the MSA?',
    'master service agreement': 'The master service agreement covers this.',
    'commercial terms': 'Legal is reviewing the commercial terms.',
    'not to exceed': 'Billed not to exceed the agreed figure.',
    'nte': 'Procurement wants an NTE on the order.',
    'time and materials': 'Would this be time and materials or fixed?',
    'day rate': 'What day rate applies to a Principal?',
    'hourly rate': 'Our blended hourly rate is on the last page.',
    'billing rate': 'The billing rate changes in January.',
    'fee schedule': 'Attached is the fee schedule for next year.',
    'purchase order': 'We need a purchase order before starting.',
  },
  TOPIC_CONTRACT: {
    'contract': 'The contract needs signing before we mobilise.',
    'contracts': 'Both contracts are up for renewal in March.',
    'subcontract': 'We would subcontract the field survey portion.',
    'nda': 'Can you sign the NDA before we share the model?',
    'non-?disclosure': 'Their non-disclosure agreement is attached.',
    'master service agreement': 'The master service agreement covers this scope.',
    'msa': 'Is this work under the MSA or a separate order?',
    'terms and conditions': 'Legal flagged the terms and conditions.',
    'addendum': 'Please review the addendum before Friday.',
    'countersign': 'I need you to countersign the amendment.',
    'executed agreement': 'The executed agreement is in the folder.',
    'signed agreement': 'We do not have a signed agreement yet.',
  },
  TOPIC_HR: {
    'human resources': 'Human resources asked about the start date.',
    'hr': 'HR needs the paperwork by Monday.',
    'new hire': 'The new hire starts on the second.',
    'hiring': 'Are we hiring for the analyst role this quarter?',
    'onboarding': 'Her onboarding is scheduled for next week.',
    'offboarding': 'We should start his offboarding checklist.',
    'performance review': 'Your performance review is due.',
    'compensation': 'The compensation question came up again.',
    'salary': 'He asked about the salary band.',
    'payroll': 'Payroll needs the timesheet corrections.',
    'employee benefits': 'The employee benefits renewal is due.',
    'open enrollment': 'Open enrollment closes on the fifteenth.',
    'resignation': 'I received his resignation this morning.',
    'disciplinary': 'This may become a disciplinary matter.',
    'headcount': 'Do we have headcount approved for this?',
  },
  TOPIC_FINANCIAL: {
    'invoice': 'The invoice for July has not arrived.',
    'invoices': 'Two invoices are still outstanding.',
    'invoicing': 'Who handles invoicing for this account?',
    'past due': 'The account is 30 days past due.',
    'overdue': 'That payment is overdue.',
    'accounts payable': 'Accounts payable is asking for the reference.',
    'accounts receivable': 'Accounts receivable flagged the balance.',
    'wire transfer': 'They want to pay by wire transfer.',
    'bank details': 'Please confirm the bank details before we pay.',
    'routing number': 'The routing number on the form looks wrong.',
    'purchase order': 'We need a purchase order before starting.',
    'expense report': 'My expense report is still unapproved.',
    'reimbursement': 'Submitting the travel reimbursement today.',
    'remittance': 'The remittance advice is attached.',
  },
  TOPIC_COMPLAINT: {
    'complaint': 'We received a complaint from the district team.',
    'complaints': 'There have been complaints about response times.',
    'dissatisfied': 'They are dissatisfied with the last deliverable.',
    'unacceptable': 'The delay is unacceptable to their leadership.',
    'escalation': 'This is heading for an escalation.',
    'escalating': 'They are escalating to the steering committee.',
    'disappointed': 'The client is disappointed with the outcome.',
    'unhappy with': 'They are unhappy with how it was handled.',
    'not happy with': 'He is not happy with the revised plan.',
  },
  TOPIC_MEDIA: {
    'media inquiry': 'We have a media inquiry about the outage.',
    'media request': 'A media request came in this morning.',
    'media contact': 'Who is the media contact for this?',
    'press release': 'Do you want to review the press release?',
    'press statement': 'They are drafting a press statement.',
    'interview request': 'An interview request landed in my inbox.',
    'on the record': 'He wants a comment on the record.',
    'for attribution': 'Is that quote for attribution?',
  },
  TOPIC_CONFIDENTIAL: {
    'confidential': 'The attached numbers are confidential.',
    'confidentiality': 'Confidentiality applies to the whole appendix.',
    'proprietary': 'That method is proprietary to them.',
    'privileged': 'This thread is privileged and should stay closed.',
    'do not share': 'Please do not share this outside the team.',
    'not for distribution': 'Draft only, not for distribution.',
    'internal use only': 'Marked internal use only.',
  },
  TOPIC_APPROVAL: {
    'approve': 'Can you approve this before Thursday?',
    'approves': 'Once he approves it we can proceed.',
    'approval': 'This needs your approval to move.',
    'approvals': 'Two approvals are still outstanding.',
    'sign-?off': 'We need your sign-off on the scope.',
    'authorization': 'Who gives authorization for the spend?',
    'authorisation': 'Their authorisation has not come through.',
    'authorize': 'Please authorize the change order.',
    'authorise': 'Can you authorise the extra day?',
    'go-?ahead': 'Give me the go-ahead and I will book it.',
  },
};

/**
 * Extra word forms that must also match — where a stem is the whole point.
 * @type {Record<string, string[]>}
 */
const STEM_FORMS = {
  'indemni\\w*': ['indemnification', 'indemnify', 'indemnities', 'indemnified', 'indemnitor'],
  'whistleblow\\w*': ['whistleblower', 'whistleblowing', 'whistleblown'],
  'non-?compete': ['non-compete', 'noncompete'],
  'fixed[- ]fee': ['fixed-fee', 'fixed fee'],
};

for (const [name, pattern] of Object.entries(GUARD_PATTERNS)) {
  describe(name + ' — every term matches a sentence someone would write', () => {
    const terms = alternativesOf(pattern);
    const probes = PROBES[name] || {};

    it('has a probe for every term in the pattern', () => {
      const missing = terms.filter((t) => !(t in probes));
      expect(missing).toEqual([]);
    });

    it('has no probe for a term the pattern no longer contains', () => {
      const stale = Object.keys(probes).filter((t) => terms.indexOf(t) === -1);
      expect(stale).toEqual([]);
    });

    for (const term of terms) {
      const probe = probes[term];
      if (!probe) continue;
      it('matches "' + term + '"', () => {
        const re = new RegExp(pattern.source, pattern.flags.replace('g', ''));
        expect(re.test(probe)).toBe(true);
      });
    }
  });
}

describe('guard vocabularies do not overlap where the rule needs two signals', () => {
  // pricingGuard requires PRICING_ASK && PRICING_QUESTION. A term in both
  // collapses that into a single-token trigger — how "given how much Upperform
  // has grown" became a pricing request.
  it('PRICING_ASK and PRICING_QUESTION share no term', () => {
    const ask = alternativesOf(BASE_PATTERNS.PRICING_ASK);
    const question = alternativesOf(BASE_PATTERNS.PRICING_QUESTION);
    expect(ask.filter((t) => question.indexOf(t) !== -1)).toEqual([]);
  });
});

describe('stems match every form of the word, not just the stem', () => {
  for (const [term, forms] of Object.entries(STEM_FORMS)) {
    for (const form of forms) {
      it(term + ' matches "' + form + '"', () => {
        const owner = Object.entries(GUARD_PATTERNS).find(
          ([, p]) => alternativesOf(p).indexOf(term) !== -1,
        );
        if (!owner) throw new Error('no pattern owns ' + term);
        const re = new RegExp(owner[1].source, owner[1].flags.replace('g', ''));
        expect(re.test('Regarding the ' + form + ' point raised.')).toBe(true);
      });
    }
  }
});

describe('the guards stay quiet on ordinary consulting mail', () => {
  const ORDINARY = [
    'Can you pull the crossarm workbook together before the Momentum Room?',
    'The Overlap Analysis found two groups running the same locate checks.',
    'Given how much Upperform has grown, we may need more space.',
    'Slide 2 might be susceptible to poking holes and getting stuck.',
    'I will send the revised deck with the additions you mention.',
    'Please do a peer review of the T+1 project plan for me.',
    'SCE plans to deploy approximately $40.6B from 2026 to 2030.',
    // Upperform's own subject matter, one word away from the new categories.
    // All four are lifted from opp-04, which Yancey labelled OPPORTUNITY and
    // which must still be draftable: "benefits" there means programme benefits,
    // and "expense" is the client's cost base, not a claim form.
    'Sustain benefits through performance monitoring that feeds the pipeline.',
    'Reducing operating expense improves cash flow and credit metrics.',
    'The contractor crews are onsite Monday morning.',
    'Directors are accountable for realizing benefits across the portfolio.',
  ];

  for (const line of ORDINARY) {
    it('pricing guard silent on: ' + line.slice(0, 46), () => {
      expect(pricingGuard(line, 'RE: Vector').triggered).toBe(false);
    });
    it('sensitivity guard silent on: ' + line.slice(0, 42), () => {
      expect(sensitivityGuard('ccoleman@partnersperform.com', line, 'RE: Vector', []).triggered)
        .toBe(false);
    });
    // The topic vocabulary moved out of sensitivityGuard, so it needs its own
    // silence check or the four regression lines above would stop guarding it.
    it('topic guard silent on: ' + line.slice(0, 46), () => {
      expect(policyTopicGuard(line, 'RE: Vector').triggered).toBe(false);
    });
  }
});

/**
 * The topic vocabulary is allowed to be broad precisely because the caller only
 * applies it to mail that would otherwise get a reply. These assert the split
 * rather than the words: the guard still fires on the text, and the classifier's
 * verdict is what decides whether that matters.
 */
describe('topic guards fire on the words but do not judge the mail', () => {
  const INVOICE = 'Invoice from JEREMIAH CHRISTIAN PLACENTE ROSALES (0028)';

  it('sees "invoice" as a financial matter', () => {
    const v = policyTopicGuard('You sent a payment request.', INVOICE);
    expect(v.triggered).toBe(true);
    expect(v.category).toBe('financial');
  });

  // Yancey's own junk set is three PayPal invoices. Folding this vocabulary
  // back into sensitivityGuard would put all three in his 07:30 digest.
  it('is not consulted by sensitivityGuard, which overrules the classifier', () => {
    expect(sensitivityGuard('service@paypal.com', 'You sent a payment request.', INVOICE, [])
      .triggered).toBe(false);
  });
});
