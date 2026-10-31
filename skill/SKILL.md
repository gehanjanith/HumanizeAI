# SKILL — De-AI Curriculum Docx Humanizer (UK RQF Level 5)

> Living knowledge file. Everything decided about this project lives here.
> If the code and this file disagree, this file wins — update the code.

---

## 0. Status

- **Layout: inspected and confirmed** (see §3 and §11). All structural guessing is done.
- **Code: not written yet.** `main.py` is still an empty stub in the repo (`main.py.txt`).
- **`.env`: not created yet** — needs `DEEPSEEK_API_KEY=...`.
- Next step: write `main.py`, `skill/prompt.txt`, `requirements.txt`, `.env`, then run
  `--dry-run` → `--limit 10` → full run.

---

## 1. One-paragraph summary

A Python script, `main.py`, sitting in `e:\wife task\Catelist\HumanizeAI`. It reads the
`.docx` curriculum specification in `Input/`, finds every **unit specification table**,
and lightly rewrites four things inside each table — the **unit name**, the **unit aim**
paragraph, the **learning outcomes**, and the **assessment criteria** — so the text reads
like a human wrote it and so every **action verb sits at UK RQF / FHEQ Level 5**.
Meaning, facts, numbers and credits must not change. Result is saved to `output/` as
`.docx` with the same filename. Rewriting is done by the **DeepSeek API**
(`deepseek-chat`), key stored in `.env`.

---

## 2. Hard requirements (locked)

| # | Requirement | Status |
|---|---|---|
| R1 | Input: `.docx` from `Input/` | locked |
| R2 | Output: `.docx` to `output/`, same filename | locked |
| R3 | Engine: DeepSeek API, model `deepseek-chat` | locked |
| R4 | API key in `.env` as `DEEPSEEK_API_KEY`, loaded via `python-dotenv` | locked |
| R5 | Runs as a plain Python script (`python main.py`) | locked |
| R6 | `requirements.txt` provided | locked |
| R7 | Only these four fields are rewritten: **Unit Name, Unit Aim, Learning Outcomes, Assessment Criteria** | locked |
| R8 | Content is **only slightly paraphrased** — light human polish, not a rewrite | locked |
| R9 | Meaning must be preserved exactly: no new facts, no dropped facts | locked |
| R10 | All action verbs aligned to **UK RQF Level 5** (FHEQ Level 5 / second year of bachelor's) | locked |
| R11 | Everything else untouched: other table rows, headings, TOC, page numbers, headers/footers, images, styles, cell merging | locked |
| R12 | The file in `Input/` must never be modified | locked |

---

## 3. What the document is

- **Type**: UK higher-education curriculum specification, unit-by-unit.
  240 credits, Level 5 Extended Diploma in Accounting and Finance.
- **Shape**: the content is **table-based**. Each unit occupies its own table
  (a "unit specification table") containing labelled rows for the unit title,
  unit aim, learning outcomes, assessment criteria, indicative content, credits,
  guided learning hours, etc.
- **Implication**: a naive `doc.paragraphs` sweep finds almost nothing, because the
  prose lives inside `table.rows[..].cells[..].paragraphs`. **The parser must walk
  cells, not top-level paragraphs.**
### Confirmed layout (inspected `GA Level 5 ... Specification.docx`, Oct 2026)

- **12 tables total, every one a unit spec, each 16 rows × 5 columns.** No other tables,
  no nested tables.
- Layout uses `w:gridSpan` merges (3 + 2 columns) and `w:vMerge` on the LO column only.

| Row | Cell 0 (span 3) | Cell 3 (span 2) |
|---|---|---|
| 0 | `Mandatory Unit` header band | — |
| 1 | col 0 = unit **number** (`1.`), col 1 = unit **name** (`Business Law`) | — |
| 2 | **Unit Aim** prose (one paragraph, 1–3 runs) | — |
| 3 | label `Learning Outcomes\nThe learner will` | label `Assessment Criteria\nThe learner can` |
| 4, 5, 6 | `1. <LO text>` — vMerge restart on 4, continue on 5–6 | `1.1`, `1.2`, `1.3` |
| 7, 8, 9 | `2. <LO text>` | `2.1`, `2.2`, `2.3` |
| 10, 11, 12 | `3. <LO text>` | `3.1`, `3.2`, `3.3` |
| 13, 14, 15 | `4. <LO text>` | `4.1`, `4.2`, `4.3` |

- **4 LOs and 12 ACs per unit.** LO group size = 3 rows; derive it as
  `(len(rows) - 4) // 4` and warn if it isn't 3.
- python-docx returns the **same cell object** for every row of a vMerge group, so
  de-duplicate with `id(cell._tc)` — otherwise each LO is rewritten three times.
- Every item is a **single paragraph** with the number already in the text and **1 run**
  (sometimes a trailing empty run). The aim paragraph has 1–3 runs, run 0 holding most text.
- The labels in row 3 are **never** rewritten.
- Source text uses `’` (U+2019); a cp437 console may display it as `Æ`. Display artefact
  only — do not "repair" the character.
- `Input/` contains a stale Word lock file `~$...docx`. `Document()` will fail on it —
  **skip any filename starting with `~$`**.
- No `Unit Title` / `Unit Name` label exists anywhere; the name is simply `rows[1].cells[1]`.

### Parser strategy
Despite the confirmed layout, field lookup stays **label-anchored with row-index fallback**:
find the row containing `Learning Outcomes` and `Assessment Criteria`, then take the rows
below it. That survives a rewritten/reformatted table, while the known indices give a
fast path and a sanity check.

---

## 4. Scope: exactly what gets rewritten

### Inside each unit table — REWRITE (exact cells)
1. **Unit name** — `rows[1].cells[1]` (e.g. `Business Law`). Keep it a title; keep any unit
   code. Do **not** touch `rows[1].cells[0]` (the unit number `1.`).
2. **Unit aim** — `rows[2].cells[0]` (spans cols 0–2). Two or three sentences of prose;
   rewrite sentence by sentence.
3. **Learning outcomes** — `rows[4:16].cells[0]`, **one unique cell per LO**
   (de-dupe by `id(cell._tc)`; should yield 4). Each LO starts with an action verb →
   **level-checked** (§6).
4. **Assessment criteria** — `rows[4:16].cells[3]` (spans cols 3–4), 12 items, form
   `2.1 Analyse ...`. Each begins with an action verb → **level-checked** (§6). Numbering
   (`2.1`, `2.2`) is preserved verbatim; only the wording after the number changes.

### Everywhere — DO NOT TOUCH
- Other unit-table rows: Indicative Content, Guided Learning Hours, Credits, Unit Purpose
  reference numbers, Sector/Sub-sector, Assessment guidance tables, grading descriptors.
- Headings, cover page, TOC, list of tables, headers, footers, footnotes, page numbering.
- Tables that are not unit specs (index tables, credit frameworks, assessment grids).
- Nested-table structure, merged cells, borders, shading, column widths.
- Runs/bold/italic/fonts/size — formatting is preserved by writing new text into the
  **first run** of the target paragraph and blanking the remaining runs.
- Empty cells, labels in column 0 (`Unit Aim:` label itself is never rewritten).

---

## 5. House style for the rewrite

### Aim: "reads human, still academic"
- Slight variation, not transformation. Target **≤ 25 % of words changed** per paragraph.
- Vary sentence length — split a long 40-word sentence into two, join two choppy ones.
- Vary sentence openers so consecutive paragraphs don't all begin "This unit …".
- Register: **British/Indian academic English** ("programme", "learner", "assessment",
  "unit"; no Americanism, no slang, no first person plural).
- Keep every technical term, accounting term, standard name, course code, credit value,
  and number identical.

### Banned (AI tells)
Words: delve, leverage, robust, holistic, seamless, tapestry, landscape, realm, foster,
underscore, pivotal, multifaceted, myriad, nuanced, crucial, comprehensive, meticulous,
intricate, ever-evolving, fast-paced, game-changer.
Phrases: "it is important to note", "it is worth noting", "plays a crucial/vital role",
"in today's rapidly evolving/changing", "not only … but also", "a wide range of",
"by understanding … learners will", "this unit aims to provide learners with".
Punctuation: **no em dashes (—)** — use a comma, a full stop, or a colon.
No markdown, no bullet symbol changes, no quotes added.

### Negative constraint that beats all others
If a rewrite would change a fact, a number, a code, or a credit value → **keep the
original sentence**. Accuracy > naturalness.

---

## 6. Action verbs by RQF level

The level is read from the input **file name** (`detect_level()`), defaulting to 5 when
no level is found. `--level 3|4|5|7` overrides it.

Detection patterns tried in order:
1. `level[\s\-_]*([3457])` — e.g. `GA Level 5 Extended Diploma...`
2. `([3457])\s*(extended )?(diploma|certificate|award)` — e.g. `L4 Diploma...`
3. fallback: **5**

Verb banks live in `VERB_BANKS` in `main.py`. A curriculum at level N must open its LOs
and ACs with a verb from `VERB_BANKS[N]`; `guard()` enforces this per file.

| Level | Character | Example verbs | Count |
|---|---|---|---|
| 3 | apply knowledge to routine tasks | apply, calculate, explain, describe, prepare, record, identify, use, outline, summarise | ~54 |
| 4 | apply and analyse | analyse, examine, evaluate, apply, prepare, compare, investigate, justify, recommend | ~80 |
| 5 | analyse and evaluate information and ideas | analyse, evaluate, critically examine, appraise, assess, justify, synthesise, recommend, review | ~103 |
| 7 | critical, original, conceptual work | critically evaluate, critique, synthesise, conceptualise, deconstruct, hypothesise, theorise, create | ~65 |

Rule: never move an item **down** a level. Raise a verb or swap it for another at the
same level. The prompt states this explicitly and the guard enforces the bank.

### Level 5 reference bank (the default)

Level 5 descriptor in one line: *demonstrates understanding of a subject and the ability
to apply it; can analyse and evaluate information and ideas; can present and communicate
arguments in a structured way.*

### The Level 5 verb bank (use these)

**Analysis / evaluation (core Level 5)**
analyse · appraise · assess · compare and contrast · critically analyse ·
critically examine · critically evaluate · debate · determine · differentiate ·
discriminate · distinguish · evaluate · examine · explore · investigate · judge ·
justify · prioritise · question · review · scrutinise · synthesise · weigh

**Application / production**
apply · construct · demonstrate · develop · devise · employ · implement ·
operate · perform · produce · solve · use

**Creation / design**
compose · design · formulate · generate · initiate · plan · propose · validate

**Communication**
articulate · communicate · present · report · summarise · translate

### Verbs that are BELOW Level 5 — must be rewritten when found as an LO/AC opener

| Found verb | Replace with |
|---|---|
| know, be aware of | explain, describe *(→ then push to L5: analyse, examine)* |
| understand, grasp, appreciate | analyse, examine, evaluate, assess |
| list, name, recall, state, define | identify, classify, examine, determine |
| identify (as the main task) | examine, analyse, determine |
| describe, outline, explain | analyse, examine, investigate, appraise |
| illustrate, give examples of | demonstrate how …, apply …, illustrate by reference to … |
| demonstrate (basic skill) | apply, implement, produce |
| calculate, compute | apply, evaluate, determine |
| discuss, consider | critically examine, evaluate, debate, explore |
| recognise, remember | identify, differentiate, assess |
| show how | analyse how …, demonstrate how … |
| interpret | interpret *(keep — borderline, acceptable at L5 when followed by evaluation)* |
| carry out, prepare | undertake, produce, produce and evaluate |
| comply with, follow | apply, implement, evaluate compliance with |

### Rules for verb correction
1. Only touch the **first verb** of an LO/AC item (and any verb in the same clause that is
   clearly below level). Do **not** sprinkle "critically" everywhere.
2. The new verb must keep the sentence grammatical and the meaning identical:
   "Understand the principles of double entry" → "Examine the principles of double entry".
3. **LO ↔ AC parity must be preserved.** In UK specifications assessment criteria restate
   the LO. If LO 2 becomes "Analyse X", AC 2.1/2.2/2.3 must still read as assessment of
   that same LO. Rewrite LO first, then rewrite the matching ACs with reference to the
   new LO verb.
4. Prefer **one** Level 5 verb per item, plus one secondary. Avoid stringing
   "critically analyse and evaluate and synthesise".
5. Assessment criteria of the form `2.1 Analyse …` are usually **also** at pass level —
   do not upgrade pass criteria into a higher standard than the LO demands; align to the LO.
6. Never place a Level 5 verb where the unit content cannot support it (e.g. don't
   "critically evaluate" when the LO is a mechanical skill).

---

## 7. Prompt design (lives in `skill/prompt.txt`)

Two system prompts, chosen per call type.

### A. `AIM_SYSTEM` — for the unit aim paragraph
```
You are a senior UK higher-education curriculum writer. You are lightly editing an
AI-generated unit aim so it reads as if written by an experienced programme leader.
Keep the meaning identical. This is a light touch edit, not a rewrite: change at most
about a quarter of the words.
Preserve every fact, number, qualification, and technical term exactly.
Vary sentence length and openers. British academic English. No em dashes.
Do not use: delve, leverage, robust, holistic, seamless, foster, underscore, pivotal,
comprehensive, crucial, "it is important to note", "plays a vital role",
"in today's rapidly evolving".
Return only the rewritten paragraph as plain text, no quotes, no preamble.
```

### B. `LO_SYSTEM` — for learning outcomes and assessment criteria
```
You are editing learning outcomes and assessment criteria in a UK RQF Level 5
curriculum. Rewrite each item so it is phrased naturally by a human, and so its
opening action verb sits at Level 5 (analyse, evaluate, critically examine, appraise,
assess, justify, synthesise, design, review, compare, interpret, apply, demonstrate,
produce, plan, develop, formulate, investigate, determine).
If the existing verb is below Level 5 (know, understand, list, name, describe,
explain, identify, outline, calculate, follow), replace it with exactly one appropriate
Level 5 verb.
Never change: the item number (e.g. 2.1), the subject matter, the scope, or the meaning.
Never add a new requirement. Never use "critically" more than once in an item.
Do not change the number of items and do not reorder them.
Return JSON: {"results":[{"id":<int>,"verb":"<the new opening verb>","text":"<full rewritten item>"}]}
```

### C. User message format (both cases)
```json
{"items":[{"id":0,"kind":"lo","text":"1. Understand the principles..."}]}
```
`kind` ∈ `aim` | `lo` | `ac` | `title`.

### Deterministic post-checks (never trust the model alone)
For every returned item:
- **Number guard** — every digit sequence from the original must still appear
  (item numbers, credits, unit codes). Missing → restore original.
- **Length guard** — new/original word ratio must land in `0.75–1.35`. Outside → restore.
- **Banned-phrase guard** — any banned word/phrase or `—` present → one retry, then restore.
- **Verb guard** — result must start with (optional number +) an allowed Level 5 verb.
  Not allowed → one retry, then restore.
- **Item-count guard** — the number of items in and out must be equal. Mismatch → restore
  the whole batch and retry once at batch size 1.
- **LO/AC parity check** — LO *n* and AC *n.x* must still refer to the same subject.

---

## 8. Python module layout

```
HumanizeAI/
├── Input/                     ← drop the .docx here
├── output/                    ← humanized .docx + per-file .log.txt
├── skill/
│   ├── SKILL.md               ← this file
│   └── prompt.txt             ← editable system prompts (A and B above)
├── .env                       ← DEEPSEEK_API_KEY=...
├── .gitignore                 ← .env, __pycache__
├── requirements.txt           ← python-docx, requests, python-dotenv
└── main.py
```

### `main.py` structure
```
CONFIG            INPUT_DIR, OUTPUT_DIR, PROMPT_FILE, MODEL, BASE_URL,
                  BATCH_SIZE=8, TEMPERATURE=1.2, TIMEOUT=90
deepseek.chat()   POST https://api.deepseek.com/chat/completions,
                  messages=[system, user], response_format json_object,
                  retries: 3 with 2s/5s/10s backoff on 429/5xx/timeout
find_unit_tables()  iterate doc.tables, keep tables whose text contains BOTH
                    "Learning Outcomes" and "Assessment Criteria"
unit_targets()      per table → [("title", rows[1].cells[1]),
                                   ("aim",   rows[2].cells[0])]
                    plus LO cells = dedupe(rows[4:].cells[0], key=id(_tc))
                    plus AC cells = rows[4:].cells[3]
collect_items()     split a cell into (item_number, text) items, keeping
                    the numbering prefix; one paragraph per item, so this is
                    a pass-through that validates the `^\d+(\.\d+)*\s` prefix
humanize_batch()    JSON in / JSON out, BATCH_SIZE items per request
guard()             the deterministic checks in §7
apply_text()        write new text into paragraph.runs[0], blank the rest;
                    if 0 runs → paragraph.add_run(); never touch other runs' fonts
save()              doc.save(OUTPUT_DIR / filename)
report()            console summary + .log.txt diff table
```

### CLI
```
python main.py                 # every .docx in Input/
python main.py --file x.docx   # one file
python main.py --dry-run       # print what WOULD be rewritten, no API calls
python main.py --limit 10      # first 10 items only (cheap smoke test)
python main.py --overwrite     # allow replacing an existing file in output/
```

---

## 9. Failure behaviour

- Missing key → exit with a message naming `.env` and the variable.
- One batch fails twice → those items keep their **original** text and are listed in the log.
  The script never aborts the whole document on one API error.
- No unit-spec tables found → say so loudly and write nothing.
- Never partially save a file that failed mid-way: save once, at the end.
- Log every change: `ID | FIELD | BEFORE | AFTER` plus totals
  (`scanned / rewritten / skipped / restored`).

---

## 10. Verification checklist

1. `--dry-run` → confirm it lists only unit name, aim, LO and AC items, and that
   indicative-content/credits/GLH rows are absent.
2. `--limit 10` → check the log: numbers preserved, no banned words, all verbs ≥ Level 5.
3. Full run → open output `.docx`; confirm tables still aligned, merged cells intact,
   bold labels intact, page breaks unchanged, no empty cells.
4. Compare with `Input/` copy: headings, TOC, cover page byte-identical in text.
5. Confirm the input file is unchanged (compare size / modified time).

---

## 11. Layout facts — RESOLVED (inspected, no longer open)

Confirmed against the real file on 21 Oct 2026:

- 12 unit tables, each 16 rows × 5 cols. No nested tables. No top-level prose to process.
- Unit name = `rows[1].cells[1]`; unit number = `rows[1].cells[0]` (leave alone).
- Unit aim = `rows[2].cells[0]`.
- LOs = `rows[4:16].cells[0]`, deduplicated by `id(cell._tc)` → 4 unique cells.
- ACs = `rows[4:16].cells[3]` → 12 cells, one AC each.
- Each item is a single paragraph, 1 run, number already in the text.
- Labels live in `rows[3].cells[0]` and `rows[3].cells[3]`; never rewritten.
- `Input/` holds a stale Word lock file `~$...docx` — skip names starting with `~$`.

### Remaining minor unknowns
- A unit with other than 4 LOs → derive group size from the vMerge spans and warn.
- An empty unit-name cell → skip and log.

---

## 12. Out of scope

PDF input · rewriting table-cell prose outside the four fields · rewriting indicative
content · grammar/spelling checking · plagiarism scoring · GUI · batch folders ·
translation · preserving exact mid-paragraph bold inside a rewritten sentence
(it collapses to the first run's style, which is acceptable here).
