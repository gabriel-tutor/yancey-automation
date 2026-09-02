/**
 * Which of Yancey's emails Gate 2 rates the voice against.
 *
 * Kept in its own module so src/gate2-samples.test.js can hold it to the one
 * rule that matters: every sample must be a message the assistant would
 * actually draft a reply to under the policy that is deployed. Rating a reply
 * that will never be written teaches nothing, and the list drifted into exactly
 * that — all five earlier samples were mail his own rules hold for review, four
 * from Chuck Coleman and one from outside the domain.
 *
 * There are four rather than five because four is how many of the twenty-one
 * golden fixtures survive his rules. The narrowness is the point: after 28
 * August the assistant only ever drafts to an internal colleague who is not
 * Chuck, so that is the only register the voice check needs to cover. The old
 * spread across client, business development and copied-in was measuring
 * registers the assistant is now forbidden from writing in.
 */
/**
 * `flag` is optional: it carries a warning to Yancey about something an
 * automated check cannot see, for a sample that needs one.
 *
 * @type {Array<{ id: string, register: string, why: string, flag?: string }>}
 */
export const SAMPLES = [
  {
    id: 'opp-02',
    register: 'Colleague, with numbers',
    why:
      'Shawn reports crossarm capitalization figures and leaves an open ask. The reply has to engage with the ' +
      'substance without restating his numbers back at him or inventing one he did not give.',
  },
  {
    id: 'opp-04',
    register: 'Colleague, disagreeing',
    why:
      'Darryl pushes back on what CI actually means. Three lines of his on top of a very long quoted deck outline. ' +
      'Watch whether the reply engages with his argument or just agrees pleasantly with everything.',
  },
  {
    id: 'opp-07',
    register: 'Colleague, mid-debate',
    why:
      'Four levels into an operational argument about fuse outages and CAIDI, and you are already a participant. ' +
      'The reply has to sound like the next turn in a conversation, not an opening statement.',
  },
  {
    id: 'synthetic-06-admin',
    register: 'Scheduling',
    why:
      'Kathleen asks one question that needs one answer. The failure here is length: anything more than a couple ' +
      'of lines is the assistant padding.',
  },
];
