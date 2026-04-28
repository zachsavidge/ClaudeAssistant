---
name: model-builder
description: "Build a financial model spreadsheet from raw trial P&L and Balance Sheet CSV files. Takes monthly income statement and balance sheet exports, organizes them into clean source tabs, then builds a structured Output tab with P&L, CFS, B/S, and Financial Summary sections -- all driven by SUMIF formulas back to the source data. Use this skill whenever the user mentions building a model, creating a financial model from trial balance data, organizing raw financials into a model, or provides P&L and balance sheet CSV files that need to be turned into an operating model. Also trigger when the user says 'model builder', 'build the model', 'operating model', or mentions SUMIF-based financial models."
---

# Model Builder

Build an operating model spreadsheet from raw monthly P&L and Balance Sheet CSV/text files.

## Before you start: read the references

1. Read the **xlsx SKILL.md** (in the skills directory) for openpyxl patterns, formula best practices, the mandatory recalc step, and formatting standards.
2. Read **`references/output-layout.md`** (in this skill's directory). It contains the exact formatting spec — every font, fill, border, number format, and formula pattern. You must follow it precisely. The formatting section describes theme colors, bold/italic rules, number formats, and indentation for every row type. The formula wiring section shows exactly how the model P&L section connects to the raw data sections below.
3. If the input files are in Japanese, read **`references/japanese-account-mapping.md`** for the translation table.

## What this skill produces

A single `.xlsx` workbook with these tabs:

1. **P&L (original language)** — if input is Japanese, keep the original-language tab
2. **P&L** (English) — the income statement data organized by month, with helper rows for Year End, Quarter End, and Month End dates
3. **Balance Sheet (original language)** — if input is Japanese
4. **Balance Sheet** (English) — balance sheet data organized by month, with date helper rows
5. **Output tab** — the financial model

## Architecture: three layers

The Output tab has a specific three-layer architecture. Understanding this is essential.

**Layer 1: Raw Data (bottom of the sheet, ~row 200+)**
- Raw P&L section: every P&L line item pulled via SUMIF from the P&L source tab
- OpEx Grouped section: consolidates raw SGA items into categories using SUM formulas
- OpEx Adjusted section: adjusts grouped OpEx for 1x items
- Adjustments section: exec comp, retirement, other 1x line items
- Raw B/S section: every B/S line item pulled via SUMIF from the Balance Sheet tab

**Layer 2: The Model (top of the sheet, rows ~46-160)**
- P&L section: references Layer 1 rows via simple formulas like `=M{raw_row}/1000`
- CFS section: derived from P&L and B/S data
- B/S working capital section: references Raw B/S rows
- Financial Summary: deal template with dummy values

**Layer 3: Metrics (within the model area)**
- Y/Y growth percentages
- Margins as % of sales
- Effective tax rate

**CRITICAL RULES — read these carefully, they address the most common failure modes:**

### MANDATORY ARCHITECTURE — DO NOT SKIP OR SHORTCUT

The Output tab MUST contain ALL of the following sections in this order. If any section is missing, the model is broken:

```
Row ~7:    KPIs section header (blue fill bar)
Row ~46:   P&L section header (blue fill bar)
Row ~50:   Sales, COGS, Gross Profit, OpEx items, EBITDA, PBT, Tax, NI, Check
Row ~76:   Y/Y Growth rows (annual columns only)
Row ~83:   "as a % of sales" margin rows (all columns)
Row ~96:   CFS section header (blue fill bar)
Row ~118:  B/S section header (blue fill bar)
Row ~119:  A/R, Inventory, A/P, Accrued, NWC, Deferred Revenue, Debt, Cash, Net Debt
Row ~137:  Financial Summary section header (orange fill bar)
Row ~170:  Memo sections (DTA, Debt schedule, IRR placeholders)
Row ~201:  Raw P&L section header — SUMIF formulas pulling from P&L source tab
Row ~274:  OpEx Grouped section — SUM formulas grouping Raw P&L SGA rows
Row ~284:  OpEx Adjusted section — adjusts grouped OpEx for 1x items
Row ~292:  Adjustments section — exec comp, retirement, other 1x
Row ~299:  Raw B/S section header — SUMIF formulas pulling from Balance Sheet source tab
```

### THE TWO-TIER ARCHITECTURE IS MANDATORY — NO SHORTCUTS

**DO NOT put SUMIF formulas directly in the model P&L section (rows 46-95).** This is the #1 failure mode.

The correct architecture is:
1. **Raw P&L section (row ~201)**: SUMIF formulas pull data from the P&L source tab
2. **Model P&L section (row ~50)**: Simple formulas like `=M201/1000` reference the Raw P&L rows

Example of CORRECT wiring:
- Raw P&L row 206 has: `=SUMIF('P&L'!$B$6:$ZZ$6,M$5,'P&L'!$B$12:$ZZ$12)` → pulls Total Revenue (P&L source row 12, matching Output-to-source offset)
- Model Sales row 50 has: `=M206/1000` → references the raw Output row 206, divides by 1000

Example of WRONG wiring (DO NOT DO THIS):
- Model Sales row has: `=SUMIF('P&L'!$B$6:$ZZ$6,M$5,'P&L'!$B$12:$ZZ$12)/1000` → WRONG! No intermediate layer!

### THE OPEX SECTIONS ARE MANDATORY

The OpEx Grouped, OpEx Adjusted, and Adjustments sections MUST exist. Without them:
- EBITDA = just Gross Profit (WRONG — should include operating expenses)
- No visibility into expense categories
- No ability to adjust for one-time items

The model P&L's OpEx rows (Employee Comp, Outsourced Labor, Rent, Travel, Other) reference the OpEx Adjusted section, NOT the Raw P&L directly. The formula chain is:
```
P&L source tab → [SUMIF] → Raw P&L rows → [SUM] → OpEx Grouped → [± adjustments] → OpEx Adjusted → [/1000] → Model P&L OpEx rows
```

### THE MODEL P&L MUST HAVE OPEX DETAIL BETWEEN GROSS PROFIT AND EBITDA

The model P&L section must include these rows (in order):
- Sales (positive)
- COGS (negative)
- **Gross Profit** = SUM(Sales:COGS)
- Adj Employee Comp (negative) = `-OpExAdj_EmpComp_row/1000`
- Adj Outsourced Labor (negative)
- Adj Rent (negative)
- Adj Travel & Entertainment (negative)
- Other (negative, PLUG = `-RawTotalSGA/1000 - Adjustments - SUM(named OpEx) - D&A`)
- **EBITDA - adj** = `+GrossProfit + SUM(EmpComp:Other)` — this is GP minus operating expenses
- blank row
- Adjustments = `-SUM(ExecComp:Retirement)/1000`
- **EBITDA - realized** = `+Adjustments + EBITDA_adj`
- blank row
- D&A (negative) = `-OpExGrouped_DA_row/1000`
- Other below the line (placeholder)
- blank row
- **PBT** = `RawPBT_row/1000`
- Effective Tax Rate (%)
- Tax = `-RawTax_row/1000`
- **NI** = `RawNI_row/1000`
- **Check** = `ROUND(PBT,1)=ROUND(SUM(EBITDA_realized:Other_below),1)`

### NUMBERED RULES

1. **EVERY model row must have formulas in EVERY data column.** This includes the P&L section, CFS section, AND the B/S section (A/R, Inventory, A/P, Accrued, NWC, Debt, Cash, Net Debt). Do not leave data columns blank anywhere.

2. **Quarterly columns = ONE column per quarter, NOT one per month.** The Output tab has exactly one column per fiscal quarter. For a company with data from Apr 2021 through Dec 2025 with a March FY end, that's ~19 quarterly columns. The SUMIF formulas aggregate 3 months of P&L data into each quarterly column by matching on the Quarter End date. Do NOT create monthly columns.

3. **Annual columns must exist.** After the quarterly columns, add a narrow gap column (~3.7 width), then annual columns (one per fiscal year). Annual SUMIFs match on Year End dates (row 3 of the P&L tab). Include 3-5 projection years beyond the last actual. Then another gap, then LTM and NTM columns.

4. **Sign convention.** In the model P&L: Revenue is positive. COGS is negative (formula: `=-RawCOGS/1000`). OpEx items are negative (formula: `=-RawOpEx/1000`). EBITDA = Gross Profit + SUM(OpEx lines), where OpEx lines are already negative. This means EBITDA = GP - |OpEx|.

5. **The B/S section MUST reference Raw B/S rows.** Formulas like `=M{raw_bs_ar_row}/1000` for A/R, `=-M{raw_bs_ap_row}/1000` for A/P, `=-SUM(M{cash_start}:M{cash_end})/1000` for Cash, etc. The B/S section is the second most common failure — it gets labels but no formulas.

6. **Y/Y growth and % margin sections are required.** Below the main P&L, add Y/Y growth rows (annual columns only, format `0%;(0%)`) and "as a % of sales" rows (all columns, format `0%;(0%)`). These are in the reference layout and must not be skipped.

7. **Row spacing must match the example layout.** The KPIs section should have ~35 rows of placeholders (so P&L starts around row 46). Don't compress the layout — preserve the spacing from `references/output-layout.md`.

8. **ALL account names must be English EVERYWHERE.** This includes: source tabs (P&L, Balance Sheet), Raw P&L section, Raw B/S section, OpEx Grouped, OpEx Adjusted. Translate EVERY account name using the japanese-account-mapping.md reference. Do NOT leave any Japanese text (hiragana, katakana, kanji) anywhere in the workbook. For bank names like "みずほ（法人）", transliterate to "Mizuho Bank".

9. **B/S model section: Debt and Cash MUST have formulas, not zeros.** Find the relevant Raw B/S rows for short-term borrowings, long-term borrowings, and cash/bank accounts. Debt = `=M{short_term_debt_row}/1000 + M{long_term_debt_row}/1000`. Cash = `=-SUM(M{first_cash_row}:M{last_cash_row})/1000` (negative because cash is an asset, shown as negative in the model convention). If you can't identify specific rows, still write a formula referencing the closest match — never leave them as hardcoded 0.

10. **No empty formula cells in model rows.** Every model row that has a label must have a value or formula in every data column. If a line item is zero (like adjustments the user will fill in), write `0` (the number), not leave it blank/empty.

11. **B/S source tab dates are offset by 1 month from P&L.** When building the Balance Sheet source tab from freee CSVs, the first data column in the CSV corresponds to the closing balance of the PREVIOUS month, not the labeled month. To fix this: after loading the CSV data into columns, shift all date helper rows (5, 6, 7) BACK by one month. For example, if the CSV labels start at "2021-04", the first data column's Month End date should be 2021-03-31 (not 2021-04-30). This ensures the SUMIF on the Output tab matches the correct point-in-time balance.

12. **The "Other" OpEx plug formula must use the SAME column letter for the deduction.** The plug formula is: `=-SUMIF(TotalSGA)/1000 - {col}54 - {col}55 - {col}56 - {col}57 + SUMIF(D&A)/1000`. The `{col}` must match the formula's own column (e.g., AG54 for the AG column, NOT M54). This is easy to get wrong when copying formulas across columns.

13. **COGS must reference "Cost of Goods Sold" (total), not "COGS" (component).** In Japanese accounting CSVs translated to English, "COGS" (売上原価) is just one component, while "Cost of Goods Sold" (売上原価合計) is the total line. The model's COGS row must reference the total. Look for the row that includes all cost components.

14. **B/S Accrued Expenses must include ALL current liability sub-accounts.** This includes: Accrued Expenses, Accrued Liabilities, credit card accounts (AmEx cards, etc.), Income Taxes Payable, Consumption Tax Payable, and Deposits Received. Missing any of these will cause Accrued to understate.

15. **B/S Cash must include ALL bank/cash accounts.** Sum every bank account row (Mizuho, SMBC, MUFG, SBI Sumishin, Shoko Chukin, East Japan, etc.) plus the Cash row itself. Missing bank accounts will understate cash.

16. **B/S NWC should include a Deferred Tax adjustment.** NWC = A/R + Inventory + A/P + Accrued - (Income Taxes Payable + Consumption Tax Payable + Accrued Consumption Tax + Consumption Tax Received)/1000. The last term is the "Deferred Tax" effect. Note: Income Taxes Payable and Consumption Tax Payable may already be in Accrued — the double-count is intentional to match standard modeling convention.

17. **Other plug MUST reference the Adjustments row, NOT EBITDA-realized.** The Other plug formula is: `=-M{raw_total_sga_row}/1000 - M{adjustments_row} - SUM(M{first_opex}:M{last_opex}) - M{da_row}`. If you accidentally reference the EBITDA-realized row instead of Adjustments, you create a circular dependency: Other → EBITDA → EBITDA-realized → Other. EBITDA-realized depends on EBITDA, and EBITDA depends on Other. Adjustments is independent (comes from exec comp/retirement raw values) and breaks the chain.

18. **Verify ALL critical BS accounts exist in Raw B/S before wiring the model.** After building the Raw B/S section, explicitly check that these accounts have rows: Accounts Receivable, Accounts Payable, Inventory, Cash, Short-term Borrowings, Long-term Borrowings. If any are missing (e.g., due to alphabetical sorting skipping them), add them with proper SUMIF formulas before proceeding to wire the B/S model section. A missing account will cause the model B/S to silently show 0 or use a stale formula.

19. **Annual column number formatting must match quarterly formatting.** Apply the same number format to annual columns (AG+) as quarterly columns. This is the second most common formatting failure — agents format quarterly columns correctly but forget the annual columns entirely. After applying formats, explicitly verify by reading back the number_format property of at least one annual cell per row type.

20. **Other 1x (below-the-line) must capture ALL non-operating items.** The formula should include Miscellaneous Income, Miscellaneous Losses, Extraordinary Gains, and Extraordinary Losses: `=(M{misc_income}-M{misc_losses}+M{ext_gains}-M{ext_losses})/1000`. Missing any of these causes the PBT check to fail because the bridge from EBITDA-realized to PBT won't close.

21. **CFS formulas must be populated.** Every CFS row must have formulas: NI=model NI row, D&A=negate model D&A, NWC change=current-prior NWC, CFO=SUM(NI:NWC), CapEx=placeholder 0, L-FCF=CFO+CapEx. U-FCF (normalized)=EBITDA+Tax+CapEx+NWC change. Don't leave CFS rows as empty placeholders.

22. **Y/Y growth and margin formulas must be populated.** Growth rows: `=IF(prior=0,0,(current-prior)/ABS(prior))` in annual columns only. Margin rows: `=IF(Sales=0,0,item/Sales)` in all columns. Format as `0%;(0%)`. An empty growth/margin section defeats the purpose of the model.

23. **CFS rows should be EMPTY placeholders (no formulas, General format) in the template.** The example model has ALL CFS rows (NI, D&A, Chng in NWC, 1x retirement benefits, Other, CFO, CapEx, L-FCF, Debt change, Deployable FCF, Dividends, Change in Cash, U-FCF, % margin) as LABELS ONLY with no data — General format, no formulas. This is because CFS is populated later during projection work, not from historical SUMIFs. Do NOT wire CFS formulas unless the user specifically asks.

24. **B/S Deferred Tax row is required.** Between "Accrued Expenses" and "NWC ex deferred", add a "Deferred Tax" row (ITALIC). Formula: `=-SUM(M{income_taxes_payable}:M{consumption_tax_payable},M{accrued_consumption_tax_1}:M{consumption_tax_received})/1000`. The NWC formula then becomes a simple SUM: `=SUM(M{ar}:M{deferred_tax})`. This Deferred Tax row is critical because without it, NWC will be overstated by the tax liability amounts.

25. **The Accrued Expenses formula must NOT double-count items included in Deferred Tax.** If you have a Deferred Tax row that includes Income Taxes Payable and Consumption Tax Payable, then the Accrued Expenses row should EXCLUDE those items. In the example: Accrued Expenses = `-SUM(M{accrued_expenses}:M{consumption_tax_payable},M{deposits_received})/1000` (includes AccruedExp, AccruedLiab, AmEx cards, IncomeTaxPayable, ConsumpTaxPayable, DepositsReceived). The Deferred Tax row picks up IncomeTaxPayable, ConsumpTaxPayable AGAIN plus AccruedConsumpTax — this intentional double-count is a modeling convention (Deferred Tax is treated as a separate NWC item).

26. **B/S SUMIF lookup MUST use row 7 (Month End), NOT row 5 (Year End) or row 6 (Quarter End).** Balance Sheet data is point-in-time, so we match against the exact month-end date. The formula must be: `=SUMIF('Balance Sheet'!$B$7:$DX$7,{col}$5,'Balance Sheet'!$B{row}:$DX{row})`. Using row 5 (Year End) will match annual periods instead of quarterly snapshots, causing incorrect or zero values.

27. **SUMIF source row references must match YOUR P&L/BS sheet structure.** The P&L source tab may have accounts in ALPHABETICAL order (e.g., "Advertising & Promotion" at row 9, "Total Revenue" at row 63) rather than presentation order. When writing SUMIF formulas, you MUST look up each account's ACTUAL row number in your source tab. Do NOT copy row references from an example model — its source tabs have different row ordering. Build a `{account_name: row_number}` mapping and use it consistently.

28. **OpEx Grouped Employee Compensation must include ALL compensation accounts.** The correct formula is: `=SUM(M{officers_comp}:M{employee_benefits})` where the range spans Officers' Compensation, Salaries & Wages, Bonuses, Retirement Benefits, Statutory Welfare Expenses, and Employee Benefits (6 accounts). Missing any of these causes Adj Employee Comp to understate.

29. **Exec Comp adjustment formula uses $W$ (absolute column) for the base period.** The formula `=M{officers_comp_raw}-$W${officers_comp_raw}` normalizes executive compensation by subtracting the base-period value (col W = a specific reference quarter). The `$W$` is absolute — it stays fixed across all columns. This means EBITDA-adj excludes the "excess" exec comp above the normalized level.

30. **Number format escaping: avoid `\\(` in format strings.** openpyxl may escape parentheses in number formats (e.g., `#,##0_)\\(#,##0\\)` instead of `#,##0_);(#,##0)`). LibreOffice recalculation can also re-escape formats. After any recalculation, re-apply these specific formats by replacing `\\(` with `(` and `\\)` with `)`. Also fix date formats: `mmm\\-yy` should be `mmm-yy`. Apply format fixes AFTER the final LibreOffice recalculation.

31. **Source tab row ordering MUST match the raw CSV input order.** The P&L and Balance Sheet source tabs must preserve the native hierarchical/logical ordering from the CSV files (Revenue → COGS → Gross Profit → SGA → Operating Income → Non-Operating → Taxes → Net Income for P&L; Assets → Liabilities → Equity for BS). Do NOT sort accounts alphabetically. The CSV files from freee already have the correct accounting hierarchy — preserve it exactly.

32. **Raw P&L/BS sections on Output must match source tab ordering.** The Raw P&L section (rows ~201-271) and Raw B/S section (rows ~299-396) on the Output tab must list accounts in the SAME order as the source tabs. Since the source tabs match CSV order (rule 31), and the Output Raw sections match the source tabs, there is a simple 1:1 mapping: Output Raw row N → Source tab row (N - offset). For example, if Raw P&L starts at Output row 203 and source data starts at row 9, then Output row 203 → Source row 9, Output row 204 → Source row 10, etc. (offset = 194 for P&L, 292 for BS).

33. **SUMIF formulas must use wide ranges ($ZZ$ not $DX$).** To prevent brittleness when source data has more or fewer columns, all SUMIF formulas should use `$B$6:$ZZ$6` (not `$B$6:$DX$6`) for the lookup range and `$B${row}:$ZZ${row}` for the data range. This ensures formulas work regardless of source tab width.

34. **SUMIF criterion always references Output row 5.** All SUMIF formulas use the quarter-end date from the Output tab's row 5 as the criterion. The formula pattern is `{col_letter}$5` (e.g., `M$5`). Do NOT use `M$6` or `M$7` as the criterion — row 6 and row 7 on the Output tab are empty. The Output tab's row 5 ("Fiscal Period ending") contains the quarterly dates that all SUMIFs match against.

35. **P&L SUMIF lookup row = source tab row 6 (Quarter End).** For ALL quarterly columns, P&L SUMIF formulas look up against the P&L source tab's row 6 (Quarter End dates). Pattern: `=SUMIF('P&L'!$B$6:$ZZ$6, M$5, 'P&L'!$B${src_row}:$ZZ${src_row})`. For annual columns, use row 5 (Year End): `=SUMIF('P&L'!$B$5:$ZZ$5, AG$5, 'P&L'!$B${src_row}:$ZZ${src_row})`.

36. **BS SUMIF lookup row = source tab row 7 (Month End) for ALL columns.** BS is point-in-time, so we always match against Month End dates, for both quarterly and annual columns. Pattern: `=SUMIF('Balance Sheet'!$B$7:$ZZ$7, M$5, 'Balance Sheet'!$B${src_row}:$ZZ${src_row})`. The criterion (`M$5`) is the quarter-end date from the Output tab, which equals a specific month-end in the BS source tab.

37. **Row heights must match the example template.** After building the Output tab, copy all row heights from the example "Output tab" sheet to our "Output" sheet. This ensures visual consistency.

38. **Y/Y growth projections for annual columns.** For annual columns beyond the last year of actual data, projection formulas should use Y/Y growth rates. Pattern: `=prior_year_value * (1 + growth_rate)`. Growth rates are stored in the Y/Y growth section (rows ~73-84) and can be filled in by the user.

## Step-by-step workflow

### Step 1: Ingest and understand the raw data

Read every CSV file the user provides.

**CSV quirks to handle:**
- **Encoding**: Usually Shift-JIS (cp932), not UTF-8. Always try `shift_jis` or `cp932` first.
- **Multiple files**: Data often comes split across multiple CSVs covering different date ranges. Stitch them together by matching on month columns.
- **Header row**: Row 0 is a title/description. Row 1 has month labels (e.g., "2021-04"). Account names are in column A starting from row 2.
- **P&L files** have a "期間累計" (period cumulative total) column at the end — exclude this.
- **Balance Sheet files** have a "期首" (beginning of period balance) column at the start — include it but the monthly model uses the month-end columns.
- **Empty values**: Treat empty strings as zero.

**Determine:**
- Which files are P&L vs. Balance Sheet (look at account names)
- The language (English or Japanese — translate if Japanese)
- The fiscal year-end month (look at the date range; the FY end is the LAST month of the company's fiscal year)
- The time span and how many quarterly/annual columns you'll need
- Every distinct SGA line item for OpEx grouping

### Step 2: Build the source tabs

Create English-language P&L and Balance Sheet tabs. **The layout must match this EXACTLY:**

- **Row 1**: Title (e.g., "Logicaltech Inc. - Monthly P&L")
- **Row 2**: Unit (e.g., "Unit: JPY")
- **Row 3**: (empty)
- **Row 4**: (empty)
- **Row 5**: Year End — label "Year End" in A5, date formulas in data columns
- **Row 6**: Quarter End — label "Quarter End" in A6, date formulas in data columns
- **Row 7**: Month End — label "Month End" in A7, date values/formulas in data columns
- **Row 8**: Account label — "Account" in A8, month labels ("2021-04", "2021-05", etc.) in data columns
- **Row 9 onward**: Account data. Column A = English account name. Columns B onward = monthly values.

**Data starts at column B.** Not column C or M. Column A has labels/names, column B has the first month's data.

**Date helper rows — Python code (use exactly this):**
```python
import calendar
from datetime import date

first_data_col = 2  # Column B
fy_end_month = 3  # March

# Write Month End (row 7) — seed first cell, then EOMONTH formulas
first_month = month_columns[0]  # e.g. "2021-04"
y, m = int(first_month[:4]), int(first_month[5:7])
last_day = calendar.monthrange(y, m)[1]
pl_sheet.cell(7, first_data_col, date(y, m, last_day))  # B7 = 2021-04-30

for i in range(1, len(month_columns)):
    col = first_data_col + i
    prev_col_letter = get_column_letter(first_data_col + i - 1)
    pl_sheet.cell(7, col, f'=EOMONTH({prev_col_letter}7,1)')

# Write Quarter End (row 6) — formulas pointing to the last month of each quarter in row 7
# For March FY: Q1 months = 4,5,6 → quarter-end col = col of month 6 (3rd col)
#               Q2 months = 7,8,9 → quarter-end col = col of month 9 (6th col)
for i, month_label in enumerate(month_columns):
    col = first_data_col + i
    y, m = int(month_label[:4]), int(month_label[5:7])
    # Find the quarter-end month
    if m in [4,5,6]: qe_m = 6
    elif m in [7,8,9]: qe_m = 9
    elif m in [10,11,12]: qe_m = 12
    else: qe_m = 3
    # Find the column of that quarter-end month
    qe_label = f"{y}-{qe_m:02d}" if qe_m >= m else f"{y}-{qe_m:02d}"
    if qe_m < m:  # Q4 wraps to next calendar year
        qe_label = f"{y+1}-{qe_m:02d}"
    qe_col_idx = month_columns.index(qe_label) if qe_label in month_columns else i
    qe_col = first_data_col + qe_col_idx
    qe_col_letter = get_column_letter(qe_col)
    pl_sheet.cell(6, col, f'={qe_col_letter}7')

# Write Year End (row 5) — formulas pointing to the last month of the FY in row 7
for i, month_label in enumerate(month_columns):
    col = first_data_col + i
    y, m = int(month_label[:4]), int(month_label[5:7])
    ye_year = y if m <= fy_end_month else y + 1
    ye_label = f"{ye_year}-{fy_end_month:02d}"
    if ye_label in month_columns:
        ye_col = first_data_col + month_columns.index(ye_label)
        ye_col_letter = get_column_letter(ye_col)
        pl_sheet.cell(5, col, f'={ye_col_letter}7')
    else:
        # FY extends beyond data range — hardcode the date
        pl_sheet.cell(5, col, date(ye_year, fy_end_month, 31))

# Write month labels (row 8)
pl_sheet.cell(8, 1, 'Account')
for i, label in enumerate(month_columns):
    pl_sheet.cell(8, first_data_col + i, label)
```

**CRITICAL: After writing dates, verify by reading back B6 and B7.** B7 should be a date (2021-04-30). B6 should be a formula (=D7 for a March FY). If they're None, something is wrong.

**Account data (row 9+):** Write one row per account. Column A = English name (translated from Japanese using japanese-account-mapping.md). Column B onward = monthly values (numeric, from the CSV). **CRITICAL: Preserve the original CSV row ordering** — do NOT sort alphabetically. The CSV files already have a logical hierarchical order (Revenue → COGS → Gross Profit → SGA items → Operating Income → Non-Operating → Taxes → Net Income for P&L). This same ordering should appear in the source tab AND in the Raw P&L/BS sections of the Output tab (see Rule 31-32).

If input is Japanese, also create original-language tabs.

### Step 3: Build the Output tab — bottom-up

**BUILD FROM THE BOTTOM UP.** You MUST build the sections in this exact order:
1. First: Set up column structure and date rows (Step 3a)
2. Second: Raw P&L section at row ~201 (Step 3b)
3. Third: OpEx Grouped section at row ~274 (Step 3c)
4. Fourth: OpEx Adjusted section at row ~284 (Step 3d)
5. Fifth: Adjustments section at row ~292 (Step 3e)
6. Sixth: Raw B/S section at row ~299 (Step 3f)
7. ONLY THEN: Build the model sections at the top (Step 4) that reference the rows you just created

You must know the exact row numbers of your raw data sections before you can wire the model sections. This is why bottom-up is mandatory.

#### Step 3a: Set up the Output tab header and date columns

**THE OUTPUT TAB HAS QUARTERLY COLUMNS, NOT MONTHLY.** Each column represents one fiscal quarter. The SUMIF formulas aggregate 3 months into each quarterly column.

**Column structure — Python code example (follow this exactly):**

```python
from datetime import date
from openpyxl.utils import get_column_letter

# For a March FY-end company with data from Apr 2021 - Dec 2025:
fy_end_month = 3  # March

# Generate quarterly end dates: Jun 30, Sep 30, Dec 31, Mar 31, Jun 30, ...
# First quarter end = first quarter-end date after the start of data
# For Mar FY: Q1 ends Jun 30, Q2 ends Sep 30, Q3 ends Dec 31, Q4 ends Mar 31
quarterly_dates = []
# Start from Jun 30, 2021 (Q1 FY2022) through Dec 31, 2025 (Q3 FY2026)
d = date(2021, 6, 30)  # First quarter-end
while d <= date(2025, 12, 31):
    quarterly_dates.append(d)
    # Advance 3 months to next quarter-end
    m = d.month + 3
    y = d.year + (m - 1) // 12
    m = (m - 1) % 12 + 1
    # End of month
    import calendar
    d = date(y, m, calendar.monthrange(y, m)[1])

# This gives 19 quarterly dates
first_data_col = 13  # Column M

# Write dates to row 5 ("Fiscal Period ending")
for i, qdate in enumerate(quarterly_dates):
    col = first_data_col + i
    ws.cell(5, col, qdate)  # Write as date, NOT as text string

# Write year-end dates to row 4 ("Year-ending")
for i, qdate in enumerate(quarterly_dates):
    col = first_data_col + i
    # FY end for this quarter
    if qdate.month <= fy_end_month:
        fy_end = date(qdate.year, fy_end_month, 31)
    else:
        fy_end = date(qdate.year + 1, fy_end_month, 31)
    ws.cell(4, col, fy_end)

# Gap column after last quarterly col
gap_col = first_data_col + len(quarterly_dates)
ws.column_dimensions[get_column_letter(gap_col)].width = 3.7

# Annual columns start after gap
annual_start_col = gap_col + 1
# FY2022 through FY2025 actual + 3 projection years
annual_dates = [date(y, 3, 31) for y in range(2022, 2029)]
for i, adate in enumerate(annual_dates):
    col = annual_start_col + i
    ws.cell(5, col, adate)
    ws.cell(4, col, adate)
```

**KEY POINTS:**
- Row 5 must contain ACTUAL DATE VALUES (datetime objects), NOT text strings like "2021-04"
- There should be ~19 quarterly columns, NOT 57 monthly columns
- Format row 5 dates as `mmm-yy` so they display as "Jun-21", "Sep-21", etc.
- The SUMIF formulas in the Raw P&L section match on these quarter-end dates, which automatically aggregates 3 months of P&L data per quarter

**Row spacing:** Leave ~35 placeholder rows for the KPIs section so the P&L section starts around row 46. This spacing matters for readability and matches the example layout.

#### Step 3b: Build the Raw P&L section (~row 200+)

For EVERY line item in the P&L source tab, create a row in the Output tab with:
- Column B: English account name
- Quarterly columns: `=SUMIF('P&L'!$B$4:${last_col}$4, Output!M$5, 'P&L'!$B${source_row}:${last_col}${source_row})`
- Annual columns: `=SUMIF('P&L'!$B$3:${last_col}$3, Output!AG$5, 'P&L'!$B${source_row}:${last_col}${source_row})`

**CRITICAL: `${source_row}` is the ROW NUMBER IN THE P&L SOURCE TAB, NOT the row in the Output tab.**

Example: If "Total Revenue" is at row 12 in the P&L source tab, and you're writing it to row 206 in the Output tab, the SUMIF formula in Output!M206 should be:
`=SUMIF('P&L'!$B$6:$BF$6, Output!M$5, 'P&L'!$B$12:$BF$12)` ← references P&L row 12, matches against row 6 (Quarter End)
NOT: `=SUMIF('P&L'!$B$6:$BF$6, Output!M$5, 'P&L'!$B$206:$BF$206)` ← WRONG! Row 206 doesn't exist in the P&L tab

The lookup row in the SUMIF is:
- **Row 6** (Quarter End) for quarterly columns
- **Row 5** (Year End) for annual columns

In Python, build a mapping from account name to source row number:
```python
# Build lookup: {account_name: source_row_number}
pl_account_rows = {}
for row in range(9, pl_sheet.max_row + 1):  # Data starts row 9
    name = pl_sheet.cell(row, 1).value
    if name:
        pl_account_rows[name] = row

last_data_col_letter = get_column_letter(pl_sheet.max_column)  # e.g. "BF"

# Then when writing Raw P&L SUMIF formulas:
for i, (account_name, source_row) in enumerate(pl_account_rows.items()):
    output_row = raw_pl_start_row + i
    ws.cell(output_row, 2, account_name)  # Label in Output col B
    
    # Quarterly columns — match against P&L row 6 (Quarter End), criterion from Output row 5
    for data_col in quarterly_cols:
        col_letter = get_column_letter(data_col)
        formula = f"=SUMIF('P&L'!$B$6:$ZZ$6,{col_letter}$5,'P&L'!$B${source_row}:$ZZ${source_row})"
        ws.cell(output_row, data_col, formula)
    
    # Annual columns — match against P&L row 5 (Year End), criterion from Output row 5
    for data_col in annual_cols:
        col_letter = get_column_letter(data_col)
        formula = f"=SUMIF('P&L'!$B$5:$ZZ$5,{col_letter}$5,'P&L'!$B${source_row}:$ZZ${source_row})"
        ws.cell(output_row, data_col, formula)
```

#### Step 3c: Build the OpEx Grouped section

Below the Raw P&L section, create a SEPARATE section labeled "OpEx Line Item - grouped". This is NOT the same as Raw P&L — it is a GROUPING LAYER that aggregates many Raw P&L rows into ~6-8 categories. Each row is ONE category that SUMs multiple individual Raw P&L line items. You should have ~5-8 category rows total, NOT one row per individual account. Standard categories:

- **Employee Compensation** (one row): `=SUM(M{officers_row}, M{salaries_row}, M{bonuses_row}, M{welfare_row}, M{benefits_row})`
- **Outsourced Labor** (one row): `=SUM(M{outsourcing_row}, M{commission_row})`
- **Rent** (one row): `=M{rent_row}`
- **Travel, Transport & Entertainment** (one row): `=SUM(M{travel_row}, M{entertainment_row}, M{meetings_row})`
- **Depreciation & Amortization** (one row): `=SUM(M{depreciation_row}, M{goodwill_amort_row})`
- **Recruiting & Education** (one row): `=SUM(M{recruiting_row}, M{training_row})`
- **Retirement** (one row): `=M{retirement_row}`
- **Other** (one row — PLUG): `=M{raw_total_sga_row} - SUM(M{emp_comp}:M{retirement})`

The purpose of grouping is to consolidate 30+ individual SGA line items into ~8 meaningful categories. Do NOT list individual accounts here — that's what the Raw P&L section is for.

#### Step 3d: Build the OpEx Adjusted section

Below OpEx Grouped, create adjusted versions:
- Emp Comp adjusted = Emp Comp grouped - Exec Comp adjustment - Retirement adjustment
- Other categories pass through: `=+M{grouped_row}`

#### Step 3e: Build the Adjustments section

Below OpEx Adjusted:
- Exec Comp: `=M{officers_comp_raw} - $W${officers_comp_raw}` (normalize to a reference period — or default to 0)
- Retirement payments: `=M{retirement_raw}` (or 0 if user wants to fill manually)
- Other 1x: placeholder (0)

#### Step 3f: Build the Raw B/S section

Same approach as Raw P&L but matching on Month End dates (**row 7** of Balance Sheet tab, since BS is point-in-time — we want the quarter-end snapshot):
`=SUMIF('Balance Sheet'!$B$7:$ZZ$7, M$5, 'Balance Sheet'!$B${source_row}:$ZZ${source_row})`

For annual columns, also match row 7 (Month End) since B/S is point-in-time and we want the FY-end date:
`=SUMIF('Balance Sheet'!$B$7:$ZZ$7, AG$5, 'Balance Sheet'!$B${source_row}:$ZZ${source_row})`

**CRITICAL: B/S ALWAYS uses row 7 (Month End) for lookup, for BOTH quarterly AND annual columns.** This is different from P&L, which uses row 6 for quarterly and row 5 for annual. The reason: B/S is point-in-time (snapshot), so we always want to match an exact month-end date, even for the annual column. The Output tab's annual column date (e.g., 2022-03-31) will match the March month-end date in the BS source tab.

**DO NOT use row 5 (Year End) for BS lookups — it contains annual dates that won't match the monthly dates in row 7 of the BS tab, causing all values to return 0.**

**Same CRITICAL rule: `${source_row}` is the row in the Balance Sheet SOURCE TAB, not the Output tab row.**

### Step 4: Build the model sections — top of sheet

NOW build the visible model sections, referencing the raw data engine you just created. **Every row must have formulas in every data column.**

#### Step 4a: KPIs section (placeholder)

Section header with blue fill bar. Add "Key KPIs" sub-header (bold, underline). Leave a few blank rows for the user to fill. See `references/output-layout.md` for exact formatting.

#### Step 4b: P&L section

Section header with blue fill bar. Then build each row with formulas referencing the raw data rows you created in Steps 3b-3f. **EVERY formula in this section must reference a row in the Raw P&L, OpEx Adjusted, or Adjustments sections — NEVER a SUMIF to the source tab directly.**

```python
# Example formula patterns (M = first quarterly data column)
# Sales row references Raw Total Revenue row
sales_formula = f'=M{raw_total_revenue_row}/1000'

# COGS references Raw COGS (negative because costs are positive in raw data)
cogs_formula = f'=-M{raw_cogs_row}/1000'

# Gross Profit = Sales + COGS
gp_formula = f'=SUM(M{sales_row}:M{cogs_row})'

# OpEx items reference OpEx Adjusted rows (negative to show as expense)
emp_comp_formula = f'=-M{opex_adj_emp_comp_row}/1000'

# Other is a PLUG: total SGA minus known items
other_formula = f'=-M{raw_total_sga_row}/1000-M{adjustments_row}-SUM(M{first_opex_row}:M{last_named_opex_row})-M{da_row}'

# EBITDA = Gross Profit + all OpEx lines (OpEx is negative, so this subtracts)
ebitda_formula = f'=+M{gp_row}+SUM(M{first_opex_row}:M{other_row})'

# PBT comes directly from Raw PBT
pbt_formula = f'=M{raw_pbt_row}/1000'

# Tax
tax_formula = f'=-M{raw_tax_row}/1000'

# NI
ni_formula = f'=M{raw_ni_row}/1000'

# Check
check_formula = f'=ROUND(M{pbt_row},1)=ROUND(SUM(M{ebitda_realized_row}:M{other_1x_row}),1)'
```

Apply the correct formatting for each row type per the reference: bold for subtotals, gray fill for EBITDA/L-FCF/U-FCF rows, `#,##0` for revenue, `#,##0;(#,##0);\-` for costs, etc.

**Then add the Y/Y growth and margin sections** below the main P&L. Growth uses `=(current-prior)/ABS(prior)` for annual columns. Margins use `=item/sales`. Format as `0%;(0%)`.

#### Step 4c: CFS section

Section header with blue fill bar. The CFS section is a PLACEHOLDER — all rows should have labels but NO formulas (General format, empty data cells). The CFS labels are: NI, D&A, Chng in NWC, 1x retirement benefits, Other, CFO, blank, (-) CapEx, L-FCF, blank, (-) Chng in Debt, Deployable FCF, blank, (-) Dividends / Other, Change in Cash, blank, memo:, U-FCF (normalized), % margin.

These rows are populated later during projection work, not from historical data.

#### Step 4c-2: B/S section

B/S items (A/R, Inventory, A/P, Accrued, Deferred Tax) reference Raw B/S rows divided by 1000. NWC is the sum. Cash, Debt, Net Debt reference specific Raw B/S row ranges.

**B/S model section is the #2 most common failure.** Agents frequently write hardcoded 0 instead of formulas. Every cell in the B/S model section MUST have a formula. Use this Python pattern:

```python
# After building Raw B/S section, find the row numbers for key accounts
# These are the ROW NUMBERS IN THE OUTPUT TAB for the Raw B/S section
raw_bs_rows = {}  # {account_name: output_row_number}
for row in range(raw_bs_start_row, raw_bs_start_row + num_bs_accounts):
    name = ws.cell(row, 2).value
    if name:
        raw_bs_rows[name.strip()] = row

# Now wire the B/S model section (example for column M = 13)
for data_col in all_data_cols:
    cl = get_column_letter(data_col)
    
    # A/R = Raw BS Accounts Receivable / 1000
    ws.cell(ar_model_row, data_col, f'={cl}{raw_bs_rows["Accounts Receivable"]}/1000')
    
    # Inventory
    ws.cell(inv_model_row, data_col, f'={cl}{raw_bs_rows["Inventory"]}/1000')
    
    # A/P = NEGATIVE (it's a liability)
    ws.cell(ap_model_row, data_col, f'=-{cl}{raw_bs_rows["Accounts Payable"]}/1000')
    
    # Cash = negative sum of ALL bank account Raw B/S rows
    bank_rows = [raw_bs_rows[name] for name in raw_bs_rows if "bank" in name.lower() or "cash" in name.lower()]
    cash_refs = "+".join([f'{cl}{r}/1000' for r in bank_rows])
    ws.cell(cash_model_row, data_col, f'=-({cash_refs})')
    
    # NWC = SUM of working capital items
    ws.cell(nwc_model_row, data_col, f'=SUM({cl}{ar_model_row}:{cl}{accrued_model_row})')
```

**NEVER write `0` for any B/S model cell.** If you can't determine the correct formula, at least write `=0` as a formula, not the number 0.

#### Step 4d: Financial Summary section

Section header with orange fill bar. Build the deal template with:
- Dummy hardcoded values in **blue font (FF0000FF)** for Equity Price, Deal Costs, Min Cash, Close Date
- Formulas for EV, Total Uses, Multiples
- The ENTIRE section (all rows, all columns) gets the light orange fill

See `references/output-layout.md` for the exact row sequence.

#### Step 4e: Memo sections (DTA, debt schedule, IRR)

Build placeholder templates with labels in column D. Leave data columns empty or with structural formulas (Loan BoP → Amortization → EoP, Interest = BoP × Rate).

### Step 5: Format everything

Go through the entire Output tab and apply formatting per `references/output-layout.md`:
- Section headers: blue fill bar with "x" in column A
- Financial Summary: orange fill on entire section
- Bold for subtotals (Sales, Gross Profit, EBITDA, PBT, NI, CFO, L-FCF, etc.)
- Gray fill for EBITDA adj, L-FCF, U-FCF rows
- Number formats per row type
- Thick bottom border on row 5
- Thin borders on specific subtotal rows (EV, Total Uses, EBITDA % margin)

### Step 6: Recalculate and verify

Run recalc:
```bash
libreoffice --headless --calc --convert-to xlsx --outdir OUTPUT_DIR OUTPUT_FILE.xlsx
```

Fix any errors. Common issues:
- `#REF!` from wrong sheet references
- `#DIV/0!` in margins where revenue is zero
- Off-by-one row references

### Step 6b: MANDATORY SELF-VERIFICATION CHECKLIST

Before considering the model complete, you MUST verify ALL of the following by reading back cells from the workbook. If ANY check fails, fix it before proceeding.

```python
# Run this verification code and fix ANY failures
import openpyxl
wb = openpyxl.load_workbook('output.xlsx')
ws = wb['Output']  # or whatever the Output tab is named

checks = []

# 1. Model P&L Sales row has formula like =M{row}/1000 (NOT SUMIF)
sales_formula = ws.cell(MODEL_SALES_ROW, 13).value  # col M
checks.append(("Sales references raw row", "/1000" in str(sales_formula) and "SUMIF" not in str(sales_formula)))

# 2. EBITDA includes OpEx (should NOT equal Gross Profit)
ebitda_formula = ws.cell(EBITDA_ROW, 13).value
checks.append(("EBITDA includes OpEx", "SUM" in str(ebitda_formula)))

# 3. OpEx Grouped section exists with grouped categories
checks.append(("OpEx Grouped exists", ws.cell(OPEX_GROUPED_HEADER_ROW, 2).value == "OpEx Grouped"))

# 4. CFS section exists
found_cfs = any("CFS" in str(ws.cell(r, 2).value or "") or "Cash Flow" in str(ws.cell(r, 2).value or "") for r in range(90, 120))
checks.append(("CFS section exists", found_cfs))

# 5. B/S MODEL section exists (NOT Raw B/S) with A/R, Cash, Debt formulas
found_bs_model = any("B/S" in str(ws.cell(r, 2).value or "") for r in range(110, 140))
checks.append(("B/S model section exists", found_bs_model))
# Also check A/R has a formula
ar_formula = ws.cell(AR_ROW, 13).value
checks.append(("A/R has formula", ar_formula is not None and "EMPTY" not in str(ar_formula)))

# 6. Financial Summary section exists
found_fs = any("Financial Summary" in str(ws.cell(r, 2).value or "") for r in range(130, 170))
checks.append(("Financial Summary exists", found_fs))

# 7. Y/Y growth rows exist
found_yy = any("Y/Y" in str(ws.cell(r, 4).value or "") or "y/y" in str(ws.cell(r, 4).value or "") for r in range(70, 100))
checks.append(("Y/Y growth rows exist", found_yy))

# 8. % margin rows exist
found_margin = any("% of" in str(ws.cell(r, 4).value or "").lower() or "margin" in str(ws.cell(r, 4).value or "").lower() for r in range(70, 100))
checks.append(("Margin rows exist", found_margin))

# 9. English-only account names in source tabs
pl_sheet = wb['P&L']
import re
jp_chars = 0
for row in range(7, pl_sheet.max_row+1):
    val = str(pl_sheet.cell(row, 1).value or "")
    if re.search(r'[\u3000-\u9fff]', val):
        jp_chars += 1
checks.append(("English-only P&L source tab", jp_chars == 0))

for name, passed in checks:
    status = "PASS" if passed else "**FAIL**"
    print(f"{status}: {name}")
```

```python
# 10. CRITICAL: Verify columns are QUARTERLY not monthly
row5_dates = []
for col in range(13, 40):
    val = ws.cell(5, col).value
    if val is not None:
        row5_dates.append(val)
# Should have ~19 quarterly columns. If you have 50+ columns, they're monthly (WRONG).
checks.append(("Quarterly columns (not monthly)", len(row5_dates) < 30))
# Also check that row 5 has date objects, not text strings
from datetime import date, datetime
first_date = ws.cell(5, 13).value
checks.append(("Row 5 has date objects not text", isinstance(first_date, (date, datetime))))
```

```python
# 11. CRITICAL: Sales value should NOT be 0 after recalc
# For Logicaltech, Q1 FY2022 Sales should be ~90,199 thousands JPY
recalc_sales = ws.cell(MODEL_SALES_ROW, 13).value  # After loading with data_only=True
checks.append(("Sales is non-zero after recalc", recalc_sales is not None and recalc_sales != 0))
```

**If ANY check shows FAIL, you MUST go back and fix it before finishing.**

**If Sales = 0 after recalc**, the SUMIF is broken. Common causes:
- Source tab date helper rows (5, 6, 7) are empty or have wrong dates
- SUMIF data row reference points to Output tab row instead of source tab row
- SUMIF lookup row is wrong (quarterly=row 6, annual=row 5, B/S=row 7)
- Data in source tab starts at wrong column (should start at column B)

The most commonly skipped sections (in order of frequency) are:
1. **CFS section** (~row 96) — Cash flow rows referencing P&L and B/S
2. **B/S model section** (~row 118) — A/R, Inventory, A/P, Accrued, NWC, Deferred Rev, Debt, Cash, Net Debt with formulas referencing Raw B/S
3. **Financial Summary section** (~row 137) — Deal template with orange fill
4. **Y/Y growth rows** (~row 76) — Annual growth percentages
5. **% margin rows** (~row 83) — Each item as % of sales
6. **English account names** — Source tabs must be English-only

### Step 7: Explain OpEx grouping

After delivering the file, explain:
1. What categories you created
2. Which raw line items went into each
3. Why you grouped them that way
4. What's in the "Other" plug

## Key principles

1. **SUMIF is the backbone.** Every number from the source tabs flows through SUMIF. This makes the model trivially auditable.
2. **The model section must be fully wired.** Every label row in the P&L/CFS/B/S sections must have working formulas in every data column. No empty data columns.
3. **Build bottom-up.** Raw data first, then model sections that reference it. This avoids circular references and makes the row numbers concrete before you wire the model.
4. **Formatting matters.** Follow `references/output-layout.md` precisely — the bold/italic/fill/border/number-format choices are deliberate conventions that make the model readable.
5. **Values in thousands.** All model section values divide raw values by 1000.
6. **Costs are negative in the model.** Revenue is positive, COGS and OpEx are negative. Revenue + COGS + OpEx = EBITDA.
7. **Fiscal year alignment.** A March FY end means Q1 = Apr-Jun, Q2 = Jul-Sep, etc. Get the Year End and Quarter End helper rows right and everything follows.
8. **Projection years.** For annual columns beyond the last actual data, use Y/Y growth rate formulas: `=prior_year * (1 + growth_rate)`. Growth rates are in the Y/Y growth section and can be filled by the user. For quarterly projections beyond actuals, reference the same quarter from the prior year times (1 + growth rate).
9. **Source tab ordering = CSV ordering.** Never sort source tab accounts alphabetically. Preserve the hierarchical accounting order from the raw CSV input (Revenue → COGS → SGA → Operating Income → etc.). This makes the Raw P&L/BS sections on the Output tab naturally align with the source tabs via a simple row offset.
10. **Non-brittle SUMIF ranges.** Always use `$ZZ$` as the column bound in SUMIF ranges (not `$DX$` or any specific column). This ensures formulas keep working if more or fewer months of data are loaded into the source tabs.
11. **P&L annual SUMIF uses Row 5 (Year End), not Row 6 (Quarter End).** Quarterly columns use P&L source Row 6 (Quarter End) for lookup because we want to aggregate months within a quarter. But annual columns (AO-AR) must use P&L source Row 5 (Year End) so that all 12 months with the same fiscal year end date are summed together. Using Row 6 for annual columns only picks up Q4 data.
12. **BS SUMIF always uses Row 7 (Month End) for both quarterly and annual.** For quarterly columns, the quarter-end date matches the last month of the quarter. For annual columns, the fiscal year-end date (e.g., 2022-03-31) matches the specific month-end balance, which is correct for balance sheet point-in-time figures.
13. **Japanese account name splits.** Some accounts exist in the CSV under TWO variant names (e.g., 仮受消費税 and 仮受消税, or 仮払消費税 and 仮払消税). Earlier fiscal years may only have data under one variant, while later years have data under the other. Both must be loaded into separate BS source tab rows, and both must be referenced in Output formulas (e.g., Deferred Tax = -SUM(row364:row365, row369:row370)/1000).
14. **December fiscal year data.** When BS CSV data includes the last month (e.g., Dec 2025 for a file covering Apr-Dec 2025), that month must be loaded as a new column in the BS source tab with correct Year End, Quarter End, and Month End header dates. Missing this causes all quarter-end SUMIF lookups for that period to return 0.
15. **Column width XML precision.** When fixing column widths via XML, always split multi-column `<col>` ranges (min≠max) into individual single-column definitions. openpyxl reads multi-column ranges incorrectly, applying the width only to the first column. Use precise decimal widths like `13.28515625` (not rounded 13.29).
16. **XML-level edits preserve cached values.** After LibreOffice recalculation, any openpyxl modification clears formula cache. Use XML-level edits (unzip xlsx, edit XML, re-zip) for post-recalc cosmetic fixes (column widths, format strings) to preserve computed values.
17. **KPI section is always blank.** Rows 38-49 (the KPI section between the header area and the P&L model) must be left empty — no formulas, no values in data columns. Keep only row labels in column B if they exist in the template. The KPI section is company-specific and populated manually by the user, never auto-generated.
18. **Mirror column widths exactly from example.** Copy the complete `<cols>` XML element from the example template's Output tab into the generated model. This includes hidden columns (M-X grouped with outlineLevel=1), separator columns (AK=3.71), and the full width scheme: A-B=1.71, C-K=2.71, L=13, M-AF=13.29, AK=3.71, AL-AZ=13.29, BA-BD=21, BE=12. Use multi-column ranges in XML (min/max) to keep the definition compact.
