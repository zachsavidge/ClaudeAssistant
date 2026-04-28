# Output Tab Layout Specification

This is the authoritative reference for the Output tab. Follow it precisely — every font style, fill color, number format, and formula pattern described here must be reproduced exactly.

## Global defaults

- **Font**: Calibri, size 11, black
- **Alignment**: Left-aligned, no indent (unless stated otherwise)
- **Number format**: General (unless stated otherwise)

## Column layout

| Columns | Purpose | Width |
|---------|---------|-------|
| A | Section marker ("x") | 1.7 |
| B | Primary labels | default (~8.4) |
| C | Sub-labels (level 1 indent) | 2.7 |
| D | Sub-labels (level 2 indent) | default |
| E-K | Spacers (unused) | default |
| L | Spacer | 13.0 |
| M onward | First quarterly data column | 13.3 |
| (gap col, e.g. AK) | Spacer between quarterly and annual | 3.7 |
| Annual cols | One per fiscal year | 13.3 |
| (gap col) | Spacer | default |
| BA | LTM | 21.0 |
| BB | NTM | default |

**Quarterly columns**: 4 per fiscal year. For a March FY end: Q1 ends Jun 30, Q2 ends Sep 30, Q3 ends Dec 31, Q4 ends Mar 31.

**Annual columns**: One per fiscal year. Dates generated with `=EDATE(prior_col, 12)`. Include projection years.

## Date header rows

### Row 4: "Year-ending"
- B4: "Year-ending" — **bold, italic**
- Data columns (M onward): formula pointing to the Q4 date for that fiscal year (all 4 quarters in a FY show the same year-ending). **Bold, black (FF000000), center-aligned, format=`mmm-yy`**

### Row 5: "Fiscal Period ending"
- B5: "Fiscal Period ending" — **bold, italic**
- Quarterly data columns: actual quarter-end dates. First quarterly cell is a hardcoded date; subsequent cells use `=EOMONTH(prior, 3)`. **Bold, blue font (FF0000FF), center-aligned, format=`mmm-yy`, thick bottom border**
- Annual data columns: fiscal year-end dates via `=EDATE(prior, 12)`. **Bold, black (FF000000), center-aligned, format=`mmm-yy`, thick bottom border**
- LTM/NTM columns: text labels "LTM" / "NTM". Center-aligned.

## Formatting rules by element type

### Section headers (KPIs, P&L, CFS)
- **Col A**: "x" in **blue font (FF0000FF), center-aligned**
- **Col B**: Section name, **bold**
- **Entire row** (cols B through all data columns): **solid fill, theme color 4 (blue), tint 0.80** (this produces light blue ~#D6E4F0)

### Financial Summary header
- **Col A**: "x" in **blue font (FF0000FF), center-aligned**
- **Col B**: "Financial Summary", **bold, underline=single**
- **Entire row**: **solid fill, theme color 5 (orange), tint 0.80** (produces light orange ~#FDEADA)
- The entire Financial Summary section (all rows from header through multiples) gets this orange fill on all cells

### Sub-section headers (Key KPIs, User Licenses, Customers, B/S, etc.)
- **Bold, underline=single**

### Metric label headers (Y/Y growth, as a % sales)
- **Underline=single** (not bold)

### Row-level formatting in the P&L section

| Row type | Label column | Label style | Data column style | Number format |
|----------|-------------|-------------|-------------------|---------------|
| Recurring, 1x revenue | C (indent=1) | regular | regular | `#,##0` |
| **Sales (Total Revenue)** | C | **bold** | **bold** | `#,##0` |
| COGS | C (indent=1) | regular | regular | `#,##0;(#,##0);\-` |
| **Gross Profit** | C | **bold** | **bold** | `#,##0` |
| OpEx line items (Emp Comp, Rent, etc.) | D | regular | regular | `#,##0;(#,##0);\-` |
| Other (plug) | D | regular | regular | `#,##0;(#,##0);\-` |
| **EBITDA - adj** | C | **bold** | **bold** | `#,##0_);(#,##0)` |
| ↳ EBITDA row also gets | — | — | **light gray fill** (theme=0, tint=-0.05) | — |
| (-) Adjustments | D | regular | regular | `#,##0;(#,##0);\-` |
| **EBITDA - realized** | C | **bold** | **bold** | `#,##0_);(#,##0)` |
| D&A, Interest | D | regular | regular | `#,##0;(#,##0);\-` |
| Other 1x | D | regular | regular | `#,##0;(#,##0);\-` |
| **PBT** | C | **bold** | **bold** | `#,##0;(#,##0);\-` |
| Tax | C (indent=1) | regular | regular | `#,##0;(#,##0);\-` |
| **NI** | C | **bold** | **bold** | `#,##0` |
| check | C | **italic** | **italic, right-aligned** | `#,##0` |
| Y/Y growth values | D | regular | regular/italic | `0%;(0%)` |
| % of sales values | D | regular | regular/italic | `0%;(0%)` |
| Effective tax rate | D | regular | **italic** | `0%;(0%)` |

### Row-level formatting in the CFS section

| Row type | Label column | Label style | Data style | Number format |
|----------|-------------|-------------|------------|---------------|
| NI | B | regular | regular | `#,##0` |
| D&A, NWC, etc. | B (indent=1) | regular | regular | `#,##0;(#,##0);\-` |
| **CFO** | B | **bold** | **bold** | `#,##0_);(#,##0)` |
| (-) CapEx | B (indent=1) | regular | regular | `#,##0;(#,##0);\-` |
| **L-FCF** | B | **bold** | **bold** | `#,##0_);(#,##0)` |
| ↳ L-FCF also gets | — | — | **light gray fill** (theme=0, tint=-0.05) | — |
| (-) Chng in Debt | B (indent=1) | regular | regular | — |
| **Deployable FCF** | B | **bold** | **bold** | — |
| **Change in Cash** | B | **bold** | **bold** | — |
| memo: | B | **italic** | — | — |
| **U-FCF (normalized)** | B | **bold** | **bold** | — |
| ↳ U-FCF also gets | — | — | **light gray fill** (theme=0, tint=-0.05) | — |
| % margin | B | **italic** | — | — |

### B/S section (within CFS area)

| Row type | Label column | Label style | Data style | Number format |
|----------|-------------|-------------|------------|---------------|
| B/S header | B | **bold, underline** | — | — |
| A/R, Inventory, A/P, etc. | C | regular | **italic** | `#,##0;(#,##0);\-` |
| **NWC ex deferred** | B | **bold** | **bold** | `#,##0;(#,##0);\-` |
| Deferred Revenue | B | regular | regular | `#,##0;(#,##0);\-` |
| Debt, Cash | B (indent=1) | regular | regular | `#,##0;(#,##0);\-` |
| **Net Debt** | B | **bold** | **bold** | `_(* #,##0_);_(* (##0);_(* \-??_);_(@_)` |

### Bottom-of-model sections

| Section | Header style |
|---------|-------------|
| Raw P&L | B: **bold, black font (FF000000)** |
| OpEx grouped | B: **underline** |
| OpEx adj | B: **underline** |
| Adj | B: regular |
| Raw B/S | B: **bold** |

Data in raw sections: green font for SUMIF cross-sheet references would be ideal, but black is acceptable. Number format: `#,##0` or General.

---

## Formula wiring: how the model section connects to raw data

This is the most critical part. The model P&L (top of the sheet) references the raw data sections (bottom of the sheet). Every row in the model section must have formulas in every data column (quarterly AND annual).

### Notation
- `{raw_pl_XXX}` = row number where "XXX" appears in the Raw P&L section
- `{raw_bs_XXX}` = row number where "XXX" appears in the Raw B/S section
- `{opex_adj_XXX}` = row number in the OpEx Adjusted section
- `{opex_grp_D&A}` = row number in OpEx Grouped for D&A
- `{adj_exec}`, `{adj_retire}` = row numbers in the Adjustments sub-section

All references below are relative — they refer to rows within the SAME Output tab. The Raw sections use SUMIF to pull from source tabs; the model section references those Raw rows.

### P&L section formulas

```
Sales (Total Revenue):     =M{raw_pl_total_revenue}/1000
COGS:                      =-M{raw_pl_cogs}/1000
Gross Profit:              =SUM(M{sales_row}:M{cogs_row})

Adj Employee Comp:         =-M{opex_adj_emp_comp}/1000
Outsourced Labor:          =-M{opex_adj_contract}/1000
Rent:                      =-M{opex_adj_rent}/1000
Travel, Transport & Ent:   =-M{opex_adj_travel}/1000
Other (plug):              =-M{raw_pl_total_sga}/1000 - M{adjustments_row} - SUM(M{first_opex}:M{last_named_opex}) - M{da_row}

EBITDA - adj:              =+M{gross_profit_row} + SUM(M{first_opex}:M{other_row})

(-) Adjustments:           =-SUM(M{adj_exec}:M{adj_retire})/1000
EBITDA - realized:         =+M{adjustments_row} + M{ebitda_adj_row}

D&A:                       =-M{opex_grp_da}/1000
Interest:                  =-M{raw_pl_interest_expense}/1000
Other 1x:                  =SUM(M{raw_extraordinary_gains}, -M{raw_extraordinary_losses}, ...)/1000

PBT:                       =M{raw_pl_pbt}/1000

Tax:                       =-M{raw_pl_tax}/1000
NI:                        =M{raw_pl_ni}/1000

check:                     =ROUND(M{pbt_row},1)=ROUND(SUM(M{ebitda_realized}:M{other_1x}),1)
```

### Y/Y Growth formulas (annual columns only)
```
Revenue growth:            =(AO{sales} - AN{sales}) / ABS(AN{sales})
EBITDA growth:             =(AO{ebitda} - AN{ebitda}) / ABS(AN{ebitda})
```

### Margin formulas (as % of sales)
```
Gross Profit %:            =M{gp}/M{sales}
EBITDA %:                  =M{ebitda}/M{sales}
OpEx item %:               =M{opex_item}/M{sales}
Effective tax rate:        =M{tax}/M{pbt}
```

### CFS formulas (annual columns)
```
NI:                        =AO{ni_row}  (from P&L section)
D&A:                       =AO{da_row}  (from P&L section, positive)
Chng in NWC:               =AO{nwc_row} - AN{nwc_row}
CFO:                       =SUM(AO{ni_cfs}:AO{other_cfs})
(-) CapEx:                 placeholder (user fills, blue font)
L-FCF:                     =AO{cfo} + AO{capex}
(-) Chng in Debt:          =AO{debt} - AN{debt}
Deployable FCF:            =AO{lfcf} + AO{chng_debt}
Change in Cash:            =AO{deployable} + AO{dividends}
U-FCF (normalized):        =AO{ebitda_adj} + AO{tax} + AO{capex} + AO{chng_nwc}
% margin:                  =AO{ufcf} / AO{sales}
```

### B/S formulas (quarterly columns — point-in-time)
```
A/R:                       =M{raw_bs_ar}/1000
Inventory:                 =M{raw_bs_inventory}/1000
A/P:                       =-M{raw_bs_ap}/1000
Accrued Expenses:          =-SUM(M{raw_bs_accrued_start}:M{raw_bs_accrued_end})/1000
Deferred Tax:              =-SUM(M{raw_bs_def_tax_start}:M{raw_bs_def_tax_end})/1000
NWC ex deferred:           =SUM(M{ar_row}:M{def_tax_row})
Deferred Revenue:          =(M{raw_bs_unearned} + M{raw_bs_deposits})/1000
Debt:                      =M{raw_bs_lt_borrowings}/1000 + M{raw_bs_st_borrowings}/1000
Cash:                      =-SUM(M{raw_bs_cash_start}:M{raw_bs_cash_end})/1000
Net Debt:                  =SUM(M{debt_row}:M{cash_row})
```

### Raw P&L SUMIF formulas
For each P&L line item, in quarterly columns:
```
=SUMIF('P&L'!$B$4:${last_col}$4, 'Output tab'!M$5, 'P&L'!$B${source_row}:${last_col}${source_row})
```
In annual columns:
```
=SUMIF('P&L'!$B$3:${last_col}$3, 'Output tab'!AO$5, 'P&L'!$B${source_row}:${last_col}${source_row})
```
Where row 4 of the P&L tab is the Quarter End row, row 3 is the Year End row.

### Raw B/S SUMIF formulas
For each BS line item (balance sheet is point-in-time, so match on month-end):
```
=SUMIF('Balance Sheet'!$B$5:${last_col}$5, 'Output tab'!M$5, 'Balance Sheet'!$B${source_row}:${last_col}${source_row})
```
Where row 5 of the BS tab is the Month End row.

### OpEx Grouped formulas
Each grouped category is a SUM of the relevant raw P&L rows:
```
Employee Compensation:  =SUM(M{officers}, M{salaries}, M{bonuses}, M{welfare}, M{benefits})
Outsourced Labor:       =SUM(M{outsourcing}, M{commissions})
Rent:                   =M{rent_raw}
Travel & Ent:           =SUM(M{travel}, M{entertainment}, M{meetings})
D&A:                    =SUM(M{depreciation}, M{goodwill_amort})
```

### OpEx Adjusted formulas
```
Emp Comp - adj:         =M{grouped_emp} - M{adj_exec} - M{adj_retire}
Outsourced Labor:       =+M{grouped_contract}
Travel & Ent:           =+M{grouped_travel}
Rent:                   =+M{grouped_rent}
```

### LTM / NTM formulas
```
LTM Revenue:  =AR{sales} * ($BC$138) + (1-$BC$138) * AS{sales}
NTM Revenue:  =AS{sales} * $BC$138 + (1-$BC$138) * AT{sales}
```
Where $BC$138 holds a weight fraction based on close date position in the fiscal year.

---

## Full row-by-row template

This is the exact sequence of rows. Row numbers will vary depending on how many KPI placeholder rows you include, but the ORDER and STRUCTURE must match.

```
Row 2:   B = "[Company] Model"                  (bold, size 14)
Row 4:   B = "Year-ending"                      (bold, italic)
Row 5:   B = "Fiscal Period ending"             (bold, italic)
         [blank row]
Row 7:   SECTION: KPIs                          (blue fill bar)
Row 9:   B = "Key KPIs"                         (bold, underline)
         [KPI placeholder rows in col C/D]
         [blank rows for user licenses, customers, revenue by product — all placeholders]
         [blank row]
Row ~46: SECTION: P&L                           (blue fill bar)
         [blank row]
         C = "Recurring"                         (indent 1)
         C = "1x"                                (indent 1)
         C = "Sales"                             (bold — total revenue)
         [blank row]
         C = "COGS"                              (indent 1)
         C = "Gross Profit"                      (bold)
         [blank row]
         D = "Adj Employee Compensation"
         D = "Outsourced labor"
         D = "Rent"
         D = "Travel, transport & entertainment"
         D = "Other"
         C = "EBITDA - adj"                      (bold, gray fill)
         [blank row]
         D = "(-) Adjustments"
         C = "EBITDA - realized"                 (bold)
         [blank row]
         D = "D&A"
         D = "Interest"
         D = "Other 1x"
         C = "PBT"                               (bold)
         [blank row]
         C = "Tax"                               (indent 1)
         C = "NI"                                (bold)
         C = "check"                             (italic)
         [blank row]
         D = "Y/Y growth (%)"                   (underline)
         D = "Recurring Revenue"
         D = "1x revenue"
         D = "Total Revenue"
         D = "EBITDA - adj"
         [blank row]
         D = "Adj Employee Compensation"
         D = "Outsourced labor"
         D = "Rent"
         D = "Travel, transport & entertainment"
         D = "Other"
         [blank row]
         D = "as a % sales"                      (underline)
         D = "Gross Profit"
         D = "EBITDA - adj"                      (thin borders top/bottom/left)
         [blank row]
         D = each OpEx item as % of sales
         D = "Adj as a % of sales"
         D = "Effected tax rate (as % of PBT)"
         [blank row]
SECTION: CFS                                    (blue fill bar)
         [blank row]
         B = "NI"
         B = "D&A"                               (indent 1)
         B = "Chng in NWC"                       (indent 1)
         B = "1x retirement benefits"            (indent 1)
         B = "Other"                             (indent 1)
         B = "CFO"                               (bold)
         [blank row]
         B = "(-) CapEx"                         (indent 1)
         B = "L-FCF"                             (bold, gray fill)
         [blank row]
         B = "(-) Chng in Debt"                  (indent 1)
         B = "Deployable FCF"                    (bold)
         [blank row]
         B = "(-) Dividends / Other"             (indent 1)
         B = "Change in Cash"                    (bold)
         [blank row]
         B = "memo:"                             (italic)
         B = "U-FCF (normalized)"                (bold, gray fill)
         B = "% margin"                          (italic)
         [blank row]
         B = "B/S"                               (bold, underline)
         C = "A/R"
         C = "Inventory"
         C = "A/P"
         C = "Accrued Expenses"
         C = "Deferred Tax"
         B = "NWC ex deferred"                   (bold)
         C = "as a % of sales"
         [blank row]
         B = "Deferred Revenue"
         C = "as a % of sales (annualized)"
         [blank row]
         B = "Debt"                              (indent 1)
         B = "Cash"                              (indent 1)
         B = "Net Debt"                          (bold)
         [blank row]
         B = "Interest"
         C = "% rate"
         [blank row]
SECTION: Financial Summary                      (orange fill bar — bold, underline)
         [entire section has orange fill on all cells]
         [blank row]
         D = "Equity Price"                      (italic)
         D = "Debt Outstanding"                  (italic)
         D = "Excess Cash"                       (italic)
         C = "Enterprise Value"                  (thin borders top/bottom/left)
         D = "Deal Costs"
         C = "Total Uses"                        (bold, thin borders top/bottom/left)
         C = "memo:"                             (italic)
         C = "EV adj for tax shield"
         C = "Est. Close Date"
         [blank row]
         C = "Multiples (headline)"              (bold, underline)
         C = "EV/Total revenue"
         C = "EV/EBITDA - adj"
         C = "EV/U-FCF"
         C = "P/L-FCF (pre amort)"
         [blank row]
         C = "Multiples (inclusive of deal cost)" (bold, underline)
         C = "EV/Total revenue"
         C = "EV/EBITDA"
         C = "EV/U-FCF"
         [blank rows]
         [memo sections: DTA, debt schedule, IRR — col D, regular formatting]
         [blank rows]
--- BELOW THE FOLD (raw data engine) ---
HEADER:  Raw P&L                                (bold)
         [every P&L line item, one row each, SUMIF formulas]
         [blank row]
HEADER:  OpEx Line Item - grouped                (underline)
         [grouped categories with SUM formulas]
         [blank row]
HEADER:  OpEx Line Item - adj                    (underline)
         [adjusted categories]
         [blank row]
HEADER:  Adj                                     (regular)
         [adjustment line items]
         [blank rows]
HEADER:  Raw B/S                                 (bold)
         [every B/S line item, one row each, SUMIF formulas]
```
