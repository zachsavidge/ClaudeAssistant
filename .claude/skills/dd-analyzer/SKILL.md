---
name: dd-analyzer
description: >
  Analyze due diligence materials (data room files, management meeting notes) and answer a standard
  set of business diligence questions, outputting findings to a Notion page. Use this skill whenever
  the user mentions due diligence, diligence questions, data room analysis, management meeting notes
  for a deal, investment analysis from source documents, or anything related to evaluating a business
  for acquisition. Also trigger when the user asks to "run diligence," "answer the DD questions,"
  "analyze the data room," or references the diligence question list.
---

# Due Diligence Analyzer

Analyze data room materials and management meeting notes to answer a standard set of business
diligence questions. Output goes to a single Notion page per deal.

## Core Principles

**Accuracy over completeness.** It is far better to leave a question blank or say "insufficient data"
than to give a wrong or speculative answer. Every answer should be grounded in the source materials.
If the data only partially addresses a question, answer the part you can and explicitly note what's
missing.

**Be skeptical of management claims.** Management presentations and meeting notes often contain
optimistic framing. When management says something, treat it as a claim to be verified against the
data, not as a fact. If the data contradicts or doesn't support what management said, call that out.
Prefer conclusions drawn from financial data, customer lists, and operational metrics over narrative
assertions.

**Confidence tagging.** For each answer, indicate how confident you are:
- **High** — directly supported by data in the materials
- **Medium** — reasonable inference from available data, but some assumptions involved
- **Low** — limited data; answer is a best guess based on fragments

If there genuinely isn't enough information, say so: "Not enough data to answer." That's a perfectly
good response.

**Write concisely.** No filler, no jargon where plain language works. Short sentences. If a number
speaks for itself, don't wrap it in qualifiers.

**Answer only the question asked.** Each diligence question has a specific scope. Stay within it.
Don't volunteer adjacent analysis, extra metrics, or commentary that wasn't requested. If the
question asks about revenue split by product, give the revenue split — don't add margin analysis,
growth commentary, or strategic implications. The reader knows the context; they just need the
answer. If something important comes up that doesn't fit the current question, put it in the
Red Flags section instead of shoehorning it into an unrelated answer.

## Workflow

### 1. Gather Materials

There are three possible sources of input — the user can provide any combination:

1. **A Google Drive folder URL** (preferred for cloud Claude Code sessions; works from any machine)
2. **A local folder path** on the user's drive (works for desktop / local Claude Code sessions only)
3. **A Notion page** containing additional relevant documents or notes

Ask the user for:
- The deal/company name
- The data room source — either the Drive folder URL OR the local folder path
- The Notion page URL with additional materials (optional)
- The **master Notion page** URL where the output should be created as a sub-page
- Any specific areas of focus or concern

#### Reading from a Google Drive folder (cloud workflow)

If the user provides a Drive folder URL, the data room files live in Drive (typically inside the
Renga Shared Drive or a folder shared with the `claude-skills` service account). The cloud sandbox
can pull them into ephemeral local storage for parsing — files never touch the user's local
machine and are gone when the session ends.

Use the `drive_sync.py` helper from the `translate-dataroom` skill (it's bundled in this same repo
under `.claude/skills/translate-dataroom/scripts/drive_sync.py`):

```bash
python .claude/skills/translate-dataroom/scripts/drive_sync.py down "<Drive folder URL>" /tmp/dataroom
```

This walks the entire folder structure recursively and downloads every file (Excel, PDF, DOCX,
PPTX, etc.) into `/tmp/dataroom/`. Then read and analyze from `/tmp/dataroom/` exactly as if it
were a local folder. **No translation is required** — Claude reads Japanese natively and the
underlying file-parsing libraries (`openpyxl`, `PyPDF2`, `python-docx`, `python-pptx`) all handle
Unicode correctly. Translation would only be needed if you wanted to *produce* English copies of
the source files, which is `translate-dataroom`'s job, not this skill's.

When using the Drive flow, the date-labeled folder convention from "2. Understand the Folder
Structure" still applies — the structure is preserved end-to-end.

#### Reading from a local folder (desktop workflow)

If the user provides a local folder path (e.g. `C:\Users\.../K Link Data Room\`), read directly
from disk. This path only works in desktop / local Claude Code sessions; cloud sandboxes don't
have access to the user's filesystem.

#### File types and tooling

File types may include PDFs, Excel files (.xlsx/.csv), Word docs, PowerPoints, and plain text
notes. Use whichever combination is fastest and most reliable in the current environment:

- **Excel (.xlsx / .xlsm)**: `openpyxl` directly, or invoke the `xlsx` skill for guidance on
  patterns (the `xlsx` skill is bundled in `.claude/skills/xlsx/` for cloud sessions).
- **PDF**: the Read tool (supports up to 20 pages per call via the `pages` param), or
  `pdfplumber` / `PyPDF2`, or invoke the `pdf` skill (bundled in `.claude/skills/pdf/`).
- **Word (.docx)**: `python-docx`, or invoke the `docx` skill (bundled in `.claude/skills/docx/`).
- **PowerPoint (.pptx)**: `python-pptx`, or invoke the `pptx` skill (bundled in
  `.claude/skills/pptx/`).
- **Plain text / CSV**: read directly.
- **Images**: the Read tool directly.

### 2. Understand the Folder Structure

Data room files follow a specific structure. Materials are organized into date-labeled folders
based on when data was downloaded. Each date folder contains two sub-folders:

```
data-room/
├── 2026-03-15/
│   ├── English/
│   │   ├── financial-statements.pdf
│   │   ├── customer-list.xlsx
│   │   └── ...
│   └── Japanese/
│       ├── financial-statements.pdf
│       ├── customer-list.xlsx
│       └── ...
├── 2026-03-22/
│   ├── English/
│   └── Japanese/
└── ...
```

The English and Japanese folders are mirror images — the English versions are translations of the
Japanese originals. **Only index one language set.** Use whichever is faster to process (typically
English, since that's the output language). Don't waste time reading both.

When multiple date folders exist, later dates may contain updated versions of earlier documents.
Note the download date when cataloging — if the same document appears in multiple batches, the
most recent version takes precedence.

All output is in English only.

### 3. Read and Catalog Sources

Before answering any questions, read through the materials and build a source index.
For each document, note:
- What it contains (financial statements, customer list, org chart, etc.)
- Time period covered
- Download batch date
- Reliability level (audited financials > management deck > verbal claims in notes)

This catalog serves two purposes: it helps you find information efficiently, and it becomes the
basis for source attribution in your answers.

### 4. Answer the Diligence Questions

Work through the standard question list in `references/diligence-questions.md`. The list covers
12 categories:

1. Product Suite
2. Pricing Model
3. Customer Base
4. Competition
5. Switching Costs
6. End-Market
7. Sales Motion
8. Cost Structure
9. Current Operations
10. Culture & Employee Function
11. Outsourcing
12. Financial Stats

Number every question sequentially across all categories (Q1, Q2, Q3... continuing through
all 12 sections). These numbers are how the user will reference specific answers later.

For each question:
- Pull relevant data from the source materials
- Note your confidence level
- Flag any contradictions between sources (especially data vs. management narrative)
- If the answer depends on management claims that can't be verified from the data, say so explicitly
- Record the source(s) in the internal source ledger (see step 5), keyed by question number

When management meeting notes and data room documents tell different stories, lead with what the
data shows and note the discrepancy. For example: "Data room financials show 85% gross retention
over 3 years. Management stated 'very low churn' in the Q2 meeting — the data suggests retention
is decent but not exceptional."

### 5. Build the Source Ledger

As you work through the materials, maintain an internal source ledger keyed by question number.
This ledger is NOT included in the Notion output — it stays internal so the user can query you
after delivery. For each numbered question, record:
- The question number (Q1, Q2, etc.)
- The source document name and path (local file or Notion page title)
- The specific location within the document (page number, sheet name, section heading)
- The download batch date (for local files)

When the user asks "where did Q14 come from?" or "what's the source for #7?", you should be
able to point to the exact document and location immediately.

### 6. Output to Notion

Create the findings as a **sub-page of the master Notion page** the user provides. Use the
Notion MCP tools to create the sub-page under the master page.

**Page title:** `[Company Name] — Diligence Findings`

**Sections** (one per category from the question list):
- Section heading (e.g., "Product Suite")
- Each question numbered (Q1, Q2, etc.) with the answer beneath it
- Confidence tag on each answer
- No source citations in the output — keep it clean. Sources live in the internal ledger only.

**Bottom section: "Gaps & Open Items"**
- Every question (by number) that couldn't be answered or was answered at Low confidence
- For each, what additional data would be needed

**Bottom section: "Red Flags & Contradictions"**
- Places where management claims don't match the data
- Unusual patterns in the financials or operations
- Anything that warrants follow-up

Use headings, toggle blocks for each category, and callout blocks for red flags.

## Handling Common Situations

**Partial data rooms:** Many data rooms are incomplete. Don't try to fill gaps with assumptions.
Answer what you can, mark the rest as gaps, and suggest what documents to request.

**Conflicting numbers:** If two documents show different figures for the same metric (e.g.,
revenue), flag both numbers, cite both sources, and note the discrepancy. Don't pick one
without explaining why.

**Management spin:** Watch for vague language in meeting notes ("strong pipeline," "significant
growth," "industry-leading"). When you encounter this, either find the data that backs it up
or note that the claim is unsubstantiated.

**Time series gaps:** For questions about trends (retention, pricing over 3-5 years), note
exactly which years are covered and which are missing. Don't extrapolate to fill gaps.

**Value proposition questions are about the customer, not the financials.** When the question
asks "what is the value prop," the answer should describe why customers buy the product — what
problem it solves for them, what they'd have to do without it, why they chose this vendor.
Draw from product descriptions, customer references, sales pitch materials, and meeting notes
about how the product is positioned. Don't reference revenue figures, retention rates, or
financial metrics — those belong in their own sections.

**Q4 specifically: focus on what the product does for the end user and customer, not on
comparisons to alternatives.** Describe each product/SKU on its own terms: what problem it
solves, how it works from the customer's perspective, and what capability the buyer gains.
Do not frame the value prop relative to competitors or alternative approaches (e.g., don't
say "cheaper than VDI" or "eliminates need for multiple PCs vs. physical separation"). Save
competitive comparisons for Q16 (pros/cons vs. competitors). The Q4 answer should stand on
its own — a reader should understand the product's value without needing to know what else
exists in the market.

**Q11: estimate end-customer revenue concentration even when it isn't directly available.**
Data rooms rarely present clean revenue-by-end-customer tables. When they don't, look for
proxies that can be used to estimate customer-level concentration:
- User/license counts by customer (best proxy for per-seat software businesses)
- Contract-level billing or advance payment schedules broken out by customer
- Customer cohort analyses with user counts
- Maintenance/support fee schedules by customer
- Any customer list with a quantitative metric (seats, locations, transaction volume)
Use whatever is available to build a top-1 / top-5 / top-10 / top-20 concentration view at
the end-customer level, not just the reseller/SI/channel level. If the business sells through
a channel, the SI-level concentration and end-customer concentration can tell very different
stories — present both.
Always caveat the methodology: state what proxy you used, why it's a reasonable stand-in for
revenue, and where it might be imprecise (e.g., "user license counts don't capture pricing
differences across contract vintages, but given uniform per-seat pricing they are directionally
reliable"). The reader should know exactly how you arrived at the estimate and what assumptions
are embedded in it.

**Q10: use active customers and users, not cumulative totals.** Customer and user counts
in data room files often include the full history of every contract ever signed — including
expired ones. Always filter to active contracts only (i.e., contracts where the end/expiry
date is on or after the reference date). State the reference date explicitly (e.g., "as of
October 2025") and explain the filtering logic used.

Key principles:
- **Identify the file date, not just the data.** A customer list dated October 2025 can only
  reliably tell you who was active as of that date. Do not use it to project forward — contracts
  signed or renewed after the file date won't be captured, so any forward view from that file
  only shows runoff of the existing book, not a true forecast.
- **Use active figures, not lifetime cumulative.** A company that has signed 171 customers over
  its history but only has 122 with live contracts today has 122 customers. The 171 figure is
  misleading as a description of the current base.
- **Cross-reference across multiple sources.** Check whether the active count from the primary
  source (e.g., a maintenance/user file) matches counts from other documents: monthly monitoring
  reports, cohort analyses, churn dashboards, or management's own stated figures. Report any
  discrepancies and explain them.
- **Segment classification from customer names.** When the data doesn't include an explicit
  segment field, classify customers by parsing their names (e.g., "Board of Education" →
  Education, "City Hall" → Local Government). Always caveat that this is name-based
  classification and may contain minor errors.
- **Distinguish customer count share from user count share.** The segment that has the most
  customers may not have the most users, and vice versa. Present both views — they tell
  different stories about the base composition.

**Q18: do not conflate churn with rising competitive intensity.** High customer churn is
an observable fact; the *cause* of that churn is an inference that requires careful analysis.
When assessing competitive intensity, resist the temptation to see churn data and conclude
"competition is increasing." There are at least three distinct explanations for elevated churn,
and the answer should acknowledge all of them rather than asserting one:
- **Demand shift:** Customers may be abandoning the product category entirely (e.g., moving
  to a fundamentally different technology approach), which reflects a shrinking addressable
  market — not necessarily more competition within the category.
- **Stable but high competition:** Competition may have always been this intense. If the
  business is experiencing its first major renewal cycle (e.g., 5-year contracts coming due
  for the first time), high churn at renewal could simply reveal the baseline competitive
  reality that was hidden during the initial sales period.
- **Genuinely increasing intensity:** New entrants, new partnerships, or new substitute
  products may be making the market more competitive than it was when the original contracts
  were signed.

Separate what management actually said from what synthesized/analyst documents infer. If
management named competitors but did not characterize competition as intensifying, do not
upgrade that into a stronger claim. State what the data shows, lay out the competing
explanations, and flag what additional evidence would be needed to distinguish between them
(e.g., win/loss data on specific bids, customer interviews on why they switched, historical
bid participation rates).

**Every claim must trace to a specific source.** Do not fill gaps with generic industry
knowledge or plausible-sounding inferences. If a data room document or management interview
does not explicitly state something, do not include it in the answer — even if it "sounds
right" or is likely true based on general knowledge of the industry. For example, do not
assert that a security breach drove customer adoption unless a source document specifically
says so. Do not attribute regulatory mandates to specific government agencies unless a source
names them. Plausible but unsourced claims erode the credibility of the entire analysis.
When you cannot find a source for something that seems like it should be true, either omit
it or explicitly flag it as an unverified assumption. The user can always add context they
know to be true — but they cannot easily identify which of your claims lack sourcing unless
you tell them.
