# Data Cleaning Log — Retail Orders Dataset
**Name:** Aditya Yogesh Choudhari , bitsom_ba_2511025  
**Dataset:** raw_orders.xlsx  
**Total Raw Records:** 932  
**Final Clean Records:** 912   
**Tool Used:** Microsoft Excel (formulas, Find & Replace, Remove Duplicates, PivotTables)

---

## 1. Issues Found

### 1.1 Text Field Issues

| Field | Issue | Count |
|---|---|---|
| customer_name | Leading/trailing whitespace | 5 |
| customer_name | Double embedded spaces (e.g. "Vikram  Iyer") | 5 |
| customer_name | Wrong case (ALL CAPS or all lowercase) | 4 |
| segment | Whitespace padding | 11 |
| segment | Embedded double space ("Small  Business") | 2 |
| segment | Case issues ("corporate", "SMALL BUSINESS") | 4 |
| region | Whitespace padding | 7 |
| region | Wrong case ("NORTH", "west", "EAST") | 9 |
| category | Whitespace padding | 9 |
| category | Case issues ("OFFICE SUPPLIES", "TECHNOLOGY", "FURNITURE") | 8 |
| sub_category | Case issues ("storage", "LABELS", "COPIERS", "MACHINES") | 7 |
| ship_mode | Case issues ("STANDARD CLASS", "FIRST CLASS") | 4 |
| ship_mode | Double space ("Standard  Class") | 1 |
| payment_status | Trailing whitespace | 4 |
| order_status | Trailing whitespace | 8 |
| order_status | Case issues ("completed", "COMPLETED", "cancelled", "CANCELLED") | 5 |

### 1.2 Date Issues

- **Mixed date formats used across the raw dataset:**
  - `YYYY-MM-DD` (ISO format)
  - `DD-MM-YYYY` (Indian format)
  - Ambiguous slash format (treated as `DD/MM/YYYY`, day-first)
  - `DD Mon YYYY` (e.g. "21 Jul 2024")
- **Ship date before order date:** 93 rows (flagged using a day-first reading of ambiguous dates — see assumption note below)
- **Same-day shipping:** valid, documented separately
- No truly unparseable dates found

### 1.3 Missing Values

| Field | Missing Count | Fill Strategy |
|---|---|---|
| region | 25 | Filled "Unknown" |
| ship_mode | 21 | Filled "Unknown" |
| discount | 18 | Filled 0.0 (all had valid quantity/unit_price) |

*Note: raw counts were 26 (region) and 22 (ship_mode); one row in each case was part of an exact-duplicate pair removed during deduplication, leaving 25 and 21 unique rows respectively in the final cleaned dataset.*

### 1.4 Discount Issues

| Issue | Count |
|---|---|
| Negative discount (e.g. -0.14, -0.19) | 15 |
| High discount, decimal form (e.g. 0.55, 0.65) | 7 |
| High discount, percentage-string form (e.g. "70%", "85%") | 8 |
| Missing discount | 18 |

*Note: high decimal and percentage-string discounts (7 + 8 = 15) were both flagged under a single `high_discount` category in the cleaned file's `discount_flag` column, rather than split into separate categories. Negative discount count is 15 after deduplication (one of the 16 raw negative-discount rows belonged to an exact-duplicate pair that was removed).*

### 1.5 Duplicates

| Type | Count |
|---|---|
| Exact duplicate rows | 40 (20 pairs) |
| Conflicting duplicate order_IDs (same ID, different data) | 12 unique IDs, 24 rows flagged |

**Full list of conflicting order IDs** (same order_id, different data — flagged for manual review, not deleted):

```
ORD-2024-10124    ORD-2024-10143    ORD-2024-10273    ORD-2024-10332
ORD-2024-10424    ORD-2025-10091    ORD-2025-10171    ORD-2025-10225
ORD-2025-10315    ORD-2025-10336    ORD-2025-10374    ORD-2025-10572
```

**Examples of conflicting duplicates:**
- `ORD-2025-10091` appears twice with different sales (9445.10 vs 9549.44) and different order_status (Completed vs Returned)
- `ORD-2024-10143` appears twice with different order_status and different ship dates, including one row flagged as an invalid ship-before-order-date record
- `ORD-2024-10124` appears twice with different sales and shipping_delay_days values

### 1.6 Sales Calculation Mismatches

- Formula used: `calculated_sales = quantity × unit_price × (1 − cleaned_discount)`
- **49 rows** had a difference of more than ₹1.00 between reported `sales` and `calculated_sales`
- Likely causes: rounding done at source system, or incorrect discount value captured
- Action: calculated_sales column uses the correct formula; original `sales` column retained for reference

---

## 2. Cleaning Actions Performed

### 2.1 Text Standardization
- Applied `Title Case` to: customer_name, segment, region, state, city, category, sub_category, ship_mode, payment_status, order_status
- Specifically normalized:
  - "Small  Business" → "Small Business"
  - "Home  Office" → "Home Office"  
  - "Standard  Class" → "Standard Class"
  - "Office  Supplies" → "Office Supplies"
  - "PRIYA MENON" / "mira das" → "Priya Menon" / "Mira Das"

### 2.2 Date Standardization
- Standardized all dates to a consistent format
- Ambiguous slash-formatted dates were interpreted as day-first (DD/MM/YYYY), consistent with the Indian business context
- Created `shipping_delay_days` as `ship_date - order_date`
- Set `shipping_delay_days` blank for rows where ship date was before order date
- Created `order_month` and `order_year` extracted from cleaned order_date

### 2.3 Missing Value Handling
- **region:** Filled 25 missing values with "Unknown" (post-deduplication count); flagged in data_quality_flag
- **ship_mode:** Filled 21 missing values with "Unknown" (post-deduplication count); flagged in data_quality_flag
- **discount:** Filled 18 missing values with 0.0 after verifying quantity and unit_price were valid

### 2.4 Discount Cleaning
- **Percentage strings (70%, 85%):** Stripped the `%` symbol and divided by 100 (e.g. "70%" → 0.70)
- **Negative discounts (15 rows):** Left original value in `cleaned_discount`; flagged as `negative_discount`
- **High discounts, both decimal and percentage-string forms (15 rows total: 7 decimal + 8 percentage-string):** Left as-is in `cleaned_discount`; flagged as `high_discount` — both sub-types were merged into one flag category rather than tracked separately

### 2.5 Duplicate Handling
- **Exact duplicates:** Found 40 rows that were complete copies (20 pairs). Removed all 20, keeping the first occurrence of each pair. Final row count: 912.
- **Conflicting duplicates:** Found 12 order_ids appearing more than once with different data (different sales, status, or dates). Both records retained in cleaned_orders.xlsx with `duplicate_conflict_flag = True` (24 rows total) — not deleted.
- **Process note:** Duplicate detection was done in two passes. First, a row-fingerprint formula (concatenating order_id, order_date, customer_id, customer_name, quantity, unit_price, sales, checked with COUNTIF) caught the obvious exact duplicates. A second full-row Remove Duplicates pass caught a small number of additional exact duplicates that the fingerprint missed, likely due to minor formatting differences in non-fingerprint columns at the time. The conflicting-duplicate flag was rebuilt only after all exact duplicates were fully removed — checking for conflicts before removing exact duplicates gives an inflated, inaccurate count, since exact-duplicate pairs also trip the "same order_id appears more than once" check.
- Business decision: Cannot silently delete conflicting records — flagged for manual review by data owner.
- **Downstream reporting fix:** The `duplicate_conflict_flag` column exists specifically so conflicting rows can be excluded from aggregate calculations without deleting them. This was applied as a Report Filter (`duplicate_conflict_flag = FALSE`) on all 6 pivot tables in `pivot_summary.xlsx` after it was discovered that the pivots had initially been built from the unfiltered range, double-counting all 24 conflicting rows and inflating total sales by ₹2,59,685.13. See README.md for the full before/after numbers.

### 2.6 Calculated Columns Added

| Column | Formula / Logic |
|---|---|
| cleaned_discount | Standardized discount (% strings converted, blanks = 0) |
| calculated_sales | quantity × unit_price × (1 − cleaned_discount) |
| calculated_profit | calculated_sales − cost |
| profit_margin | calculated_profit / calculated_sales |
| shipping_delay_days | ship_date_parsed − order_date_parsed (days) |
| order_month | Month extracted from order_date |
| order_year | Year extracted from order_date |
| date_issue_flag | Flags rows where ship_date < order_date |
| data_quality_flag | CLEAN or WARNING/INVALID with specific reason codes |
| duplicate_conflict_flag | TRUE if row is a conflicting duplicate |

---

## 3. Business Rules Applied

| Rule Area | Rule | Applied Action |
|---|---|---|
| Raw file | Do not overwrite raw_orders.xlsx | Created cleaned_orders.xlsx separately ✓ |
| Missing region | Fill as Unknown and flag | Done — 25 rows ✓ |
| Missing ship_mode | Fill as Unknown and flag | Done — 21 rows ✓ |
| Missing discount | Treat as 0 if all other fields valid | Done — 18 rows ✓ |
| Negative discount | Flag as invalid | Done — 15 rows flagged ✓ |
| High discount | Flag unusually high (decimal or percentage-string) | Done — 15 rows flagged ✓ |
| Cancelled orders | Not in completed-sales pivot | Only Completed + Paid in pivot summaries ✓ |
| Failed payments | Not in completed-sales pivot | Excluded from pivot summaries ✓ |
| Refunded orders | Summarized separately | Captured in "Non-Completed by Region" pivot ✓ |
| Ship date before order date | Flag as invalid | Done — 93 rows flagged ✓ |
| Conflicting duplicates | Flag, do not delete silently | Done — 24 rows flagged ✓ |

---

## 4. Assumptions Made

1. **Date format ambiguity:** For dates like `11/05/2024`, I treated them as `MM/DD/YYYY` (US format) based on consistency with surrounding records in that batch. Some misclassifications may exist where dates are genuinely `DD/MM/YYYY`.

2. **Missing discount = 0:** Where discount was blank and quantity/unit_price were both valid numbers, I assumed the transaction had no discount applied. If a discount was intentionally omitted for a different reason, this assumption is incorrect.

3. **Missing region:** Filled as "Unknown" rather than inferring from state/city, since state-to-region mapping was not provided in the business_rules sheet.

4. **Conflicting duplicates:** Kept both records for manual review. In a production scenario, a unique primary key system or merge rule from the source systems would be needed.

5. **High discounts (>50% decimal):** Flagged as suspicious but not removed — some business scenarios (clearance sales, bulk orders) may legitimately have high discounts.

6. **Sales mismatch:** Where calculated_sales differs from reported sales, I kept `calculated_sales` as the authoritative value since it follows the documented formula. The reported `sales` column retained for audit trail.

---

## 5. Records Removed

| Removal Reason | Count |
|---|---|
| Exact duplicate rows (kept first occurrence) | 20 |

No other records were permanently removed. All flagged records are retained with quality flags for human review.

---

## 6. Limitations of This Cleaning Process

1. **Cannot verify state-to-region mapping** without a reference lookup table — 25 rows left as "Unknown".

2. **Date ambiguity (day-first vs month-first) is the single biggest source of uncertainty in this analysis.** Ambiguous slash-formatted dates were read as day-first (DD/MM/YYYY), which produced 93 "ship before order date" flags which is  a meaningfully large number. Reading the same ambiguous dates as month-first instead would produce a much smaller flagged set (roughly 22 records based on a cross-check). This is a genuine judgment call, not a confirmed fact, and should be verified against the source system before any of these 93 records are treated as confirmed data errors.

3. **Sales calculation mismatches** (49 rows) could stem from rounding rules at the source system, partial shipments, or mid-order amendments — root cause cannot be confirmed without source system access.

4. **Conflicting duplicate order_IDs** (12 IDs, 24 rows) require input from the source system owners to determine which copy is canonical.

5. **Customer names with common issues** (double spaces, case) were normalized, but similar names could represent different people (e.g. "Mira Das" might be two different customers with the same name).
