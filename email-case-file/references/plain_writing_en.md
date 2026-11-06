# Plain English: how to write the report for English users

Use this file instead of `plain_writing.md` when the user talks to you in English, or asks for the report in English. Then every field you write yourself is in English and follows these rules.

The reader is a manager. They use the report to run a meeting. They need to see in a few seconds: who, when, what was done, and what is stuck.

Rules adapted from [isas1/skills](https://github.com/isas1/skills) (`eli5-succinct`, MIT licence). The ASD-STE100 (Simplified Technical English) register is used here: simple language, full accuracy.

## What you may rewrite, and what you must not touch

**Write in plain English (the text you write yourself):**

- `overview_summary`, `changes_summary`, each topic's `status_summary`
- `timeline[].headline`, `timeline[].points`
- `decisions[].decision`, `pending[].issue`, `pending[].options`, `pending[].note`
- `todos[].item`, `completed[].item`, `completed[].evidence`, every `note`
- Your chat messages to the user (topic check, people check, the final handover)

**Never change (protected list):**

- `quote`: keep the original language and wording, even typos. Do not translate it.
- Names of people, companies, products, part numbers, order numbers
- Amounts, currencies, quantities, unit prices, percentages, dates, deadlines
- Email and file IDs (E001, D001), subjects, file names, attachment names, links
- Every field inside `source`
- In incremental mode, **old entries that did not change**: keep last time's text exactly. The script matches old and new entries by their text. A reworded entry looks like "removed + added".

## How to write

Assume a smart reader who has not seen this case before. Not a slow reader. A new one.

1. **Lead with the answer.** The first sentence says where the case stands. No preamble.
2. **One idea per sentence.** Short, declarative sentences.
3. **Active voice, clear subject.** "John (Acme, buyer) rejected option B." Not "Option B was not accepted."
4. **Concrete over abstract.** "Acme asked for 2 samples by 30 Sep." Not "The parties discussed sample logistics."
5. **Same word for the same thing.** If you call it "the sample", do not switch to "the prototype" or "the unit" later. Readers think a new word means a new thing.
6. **No jargon.** If a term is unavoidable, define it in the same sentence.
7. **Say what you do not know.** Do not hedge inside the sentence ("it seems that perhaps"). Write the sentence plainly and set `confidence` to `低` with the reason in `note`.
8. **No drama.** Do not add feelings, meaning or predictions that are not in the emails.

## Phrases to cut

| Pattern | Example | Fix |
| :-- | :-- | :-- |
| Inflated meaning | "marks a pivotal milestone", "lays the groundwork" | Delete. Say what happened. |
| Empty verbs | "engaged in discussions", "aligned on", "continued to follow up" | Use the real action: proposed, rejected, agreed, chased, sent |
| Vague attribution | "the client side feels", "stakeholders indicated" | Name the person and role |
| Signposting | "It is worth noting that", "Importantly,", "This means that" | Delete the lead-in, keep the point |
| Stock closers | "In summary", "Overall", "All in all" | Delete. End on the most important fact. |
| AI vocabulary | "delve", "navigate", "landscape", "seamless", "robust", "underscore", "leverage" | Use the plain word or cut it |
| Rule of three | "improve efficiency, reduce cost, and strengthen ties" | Keep only the item an email supports |
| Not X but Y | "The issue is not price but lead time" | "Lead time is the blocker." |
| Vague positivity | "Good progress", "Communication is smooth" | Give the evidence: "Spec confirmed (E014)" |
| Announcing | "Below is a summary of the case" | Delete. Write the summary. |

`headline` needs the most care: "who did what", 8–15 words. Example: "Acme rejected our offer in writing and countered at 36,500." Not "Discussion on tooling cost."

## Spelling, numbers and punctuation

- Match the user's spelling (British or American) from how they write to you. If you cannot tell, use the spelling in the emails.
- Keep amounts as written in the source, with currency: `US$ 24,000`, `€ 36,500`. Do not convert.
- Dates in the text fields: `30 Sep 2026` or `2026-09-30`. Pick one and keep it for the whole report. Date fields in the JSON stay `YYYY-MM-DD`.
- No Markdown bold or dashes inside text fields. The script handles layout.

## Chat messages to an English user

- Lead with the answer. Under 100 words unless you are listing topics or people.
- For a yes/no question, answer `Yes.` or `No.` on its own line first, then one or two sentences.
- Match how the user writes: their register, their sentence length, their words for things. Do not announce that you are doing this.
- No offers or next-step lists they did not ask for.

## What stays in Chinese

The section headings and table labels in the Word, PDF, HTML, Excel and brief files come from `build_report.py`, and they are in Traditional Chinese. Your text fields will be in English. Tell the user this once, in one line, when you hand over the first report.

## Check before you build the report

After you write `report.json` and before you run `build_report.py`, scan your text fields:

1. Any phrase from the table above? Fix it.
2. Does every `headline` show who did what?
3. Is every sentence one idea, active voice, with a clear subject?
4. Is anything on the protected list changed? Change it back.
5. Incremental mode: are unchanged old entries word for word the same as last time?

Do this yourself. Do not show the user a list of edits.
