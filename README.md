# Part 1: Business Data Cleaning, Validation & Excel Reporting

## Assignment Part Title
Business Data Cleaning, Validation & Excel Reporting — Retail Orders Dataset

---

## Business Problem Summary

A retail company operating across multiple Indian states exports order-level sales data from several internal systems. The data lands in a raw, messy state every reporting cycle ,dates in inconsistent formats, text fields with random casing and whitespace, discount values stored as percentages instead of decimals, negative discounts, missing geography fields, and duplicate records that sometimes carry conflicting data.

The business needs a clean, validated, analysis-ready dataset and a set of summary reports that can go directly into management review decks. Without this cleaning, any analysis or decision made on the raw data would be unreliable.

---

## Dataset Used

- **File:** `raw_orders.xlsx`
- **Raw records:** 932
- **Columns:** 21 (order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status)
- **Date range:** January 2024 – October 2025
- **Source:** Retail company order management system (multi-system export)
- **Currency:** Indian Rupees (₹)

---

## Tools Used

- **Microsoft Excel** — For all cleaning, formula-building, pivot tables, and charting done directly in Excel
- Built-in Excel functions: TRIM, PROPER, SUBSTITUTE, DATEVALUE, COUNTIF, COUNTIFS, IF, IFERROR
- PivotTables and PivotCharts for all summary reporting

---

## Steps Performed

### Step 1: Data Profiling
Before touching the data, went through each column to identify:
- Inconsistent text formatting (whitespace, casing, double spaces)
- Mixed date formats
- Discount anomalies (negative values, percentage strings, unusually high values)
- Duplicate rows (both exact and conflicting)
- Calculation mismatches between reported and computed sales

### Step 2: Text Field Cleaning
- Trimmed leading/trailing whitespace and collapsed double spaces across all text columns (customer_name, segment, region, state, city, category, sub_category, ship_mode, payment_status, order_status)
- Applied consistent Proper Case using `=PROPER(TRIM())`
- Fixed specific compound-word spacing issues (e.g. "Small  Business" → "Small Business", "Standard  Class" → "Standard Class") using Find & Replace

### Step 3: Date Standardization
- Identified multiple mixed date formats in the raw export (ISO, DD-MM-YYYY, MM/DD/YYYY, DD Mon YYYY)
- Converted `order_date` and `ship_date` into a consistent date format
- Built `shipping_delay_days` as ship_date − order_date
- Flagged rows where ship_date fell before order_date using `=IF(ship_date<order_date, "INVALID: Ship date before order date", "")`
- Extracted `order_month` and `order_year` from the cleaned order_date

### Step 4: Missing Values
- **region:** Missing values filled with "Unknown" using Find & Replace (entire-cell match), per business rule
- **ship_mode:** Same approach — missing values filled with "Unknown"
- **discount:** Missing values filled with 0 after confirming quantity and unit_price were valid for those rows

### Step 5: Discount Cleaning
- Rows with percentage-string discounts (e.g. "70%") converted to proper decimal format
- Negative discount values flagged as invalid, not silently corrected
- Discounts above 50% flagged as a warning for review
- Built a `cleaned_discount` column holding the standardized values

### Step 6: Duplicate Handling
This was done in two passes, refined along the way:

1. **First pass:** Built a row-fingerprint formula (concatenating order_id, order_date, customer_id, customer_name, quantity, unit_price, sales) checked with COUNTIF to catch obvious exact duplicates.
2. **Second pass:** Ran Excel's built-in Remove Duplicates (full row, all columns except the formula-driven helper/flag columns) to catch a small number of additional exact duplicates the fingerprint missed due to minor formatting differences.
3. Final exact-duplicate removal: **20 rows removed**, bringing the dataset from 932 to **912 rows**.
4. Conflicting duplicates (same order_id, different underlying data) were identified using `=IF(COUNTIF(order_id_range, order_id)>1, "TRUE", "FALSE")`, applied only after exact duplicates were fully removed — checking before that point gives an inflated, inaccurate count. **12 unique conflicting order_ids found (24 rows flagged)**, not deleted, marked with `duplicate_conflict_flag = TRUE` for manual review.

### Step 7: Calculated Columns
Added these columns to `cleaned_orders.xlsx`:
- `cleaned_discount` — standardized discount value
- `calculated_sales` — quantity × unit_price × (1 − cleaned_discount)
- `calculated_profit` — calculated_sales − cost
- `profit_margin` — calculated_profit / calculated_sales
- `shipping_delay_days` — ship_date − order_date
- `order_month`, `order_year` — extracted from order_date
- `date_issue_flag` — flags invalid ship-before-order-date rows
- `data_quality_flag` — CLEAN or WARNING with reason
- `exact_duplicate` — UNIQUE/DUPLICATE flag used during cleanup (retained for audit trail)
- `duplicate_conflict_flag` — TRUE for conflicting duplicate records

### Step 8: Data Quality Report
Built `data_quality_report.xlsx` covering: executive summary, missing values, duplicate analysis (with full list of conflicting order IDs), discount issues, date issues, calculation mismatches, order status consistency, and final clean-vs-flagged record counts.

### Step 9: Pivot Summary Report
Built `pivot_summary.xlsx` using native Excel PivotTables, with one sheet per breakdown:
1. **By Region** — sales and profit by region, sorted descending, with a clustered column PivotChart
2. **By Category** — category and sub-category breakdown, sorted descending
3. **By Ship Mode** — order counts by shipping method
4. **By Segment** — average profit margin by customer segment, sorted descending
5. **Non-Completed by Region** — cancelled/returned orders by region, filtered view
6. **Monthly Sales Trend** — chronological monthly sales (2024–2025), with a line PivotChart showing the trend visually

All currency figures formatted in ₹ (Indian Rupees) to match the business context.

### Step 9a: Correcting a Double-Counting Bug in the Pivot Tables

During review, it was identified that the 6 pivot tables in `pivot_summary.xlsx` were built from the full `raw_data` range without excluding rows flagged `duplicate_conflict_flag = TRUE`. Since the 12 conflicting duplicate order IDs each contribute **two** rows with different sales figures, this meant every pivot was silently double-counting these disputed orders — both versions of each conflicting order were being summed as if they were two separate, valid transactions.

**Quantified impact:** Across the 24 affected rows (12 order IDs × 2 copies), this inflated total `calculated_sales` by **₹2,59,685.13**. The By Region pivot's Grand Total dropped from ₹90,51,689.76 to **₹87,92,004.63** once corrected — a drop that matches the inflated amount exactly, confirming the fix was applied correctly rather than just cosmetically.

**Fix applied:** `duplicate_conflict_flag` was added as a Report Filter on all 6 pivot tables (By Region, By Category, By Ship Mode, By Segment, Non-Completed by Region, Monthly Sales Trend), filtered to show only `FALSE` (i.e. non-conflicting) rows. This was verified at the pivot-cache level, not just visually, to confirm the filter was genuinely excluding the flagged rows rather than just being present and unapplied.

This is the correct way to handle conflicting duplicates in downstream reporting: the rows stay visible and flagged in `cleaned_orders.xlsx` for manual review (per the business rule of never silently deleting conflicting data), but they are excluded from aggregate summary calculations until the data owner resolves which version is correct.

---

## Business Rules Applied

| Rule | Action Taken |
|---|---|
| Raw file preservation | raw_orders.xlsx untouched; cleaned_orders.xlsx created separately |
| Missing region | Filled "Unknown", flagged |
| Missing ship_mode | Filled "Unknown", flagged |
| Missing discount | Treated as 0 when all other fields valid |
| Negative discount | Flagged as invalid, not auto-corrected |
| High/percent-format discount | Flagged as warning |
| Cancelled orders | Excluded from completed-sales pivot summaries |
| Failed payments | Excluded from completed-sales pivot summaries |
| Refunded/returned orders | Captured separately in Non-Completed by Region sheet |
| Ship date before order date | Flagged invalid |
| Conflicting duplicates | Flagged for manual review, not deleted |

---

## Summary of Data Quality Issues Found

| Category | Issue | Count |
|---|---|---|
| Duplicates | Exact duplicate rows removed | 20 |
| Duplicates | Conflicting duplicate order IDs (24 rows flagged) | 12 |
| Missing | Missing region (filled Unknown) | 25 |
| Missing | Missing ship_mode (filled Unknown) | 21 |
| Missing | Missing discount (filled 0) | 18 |
| Dates | Ship date before order date | 93 |
| Discount | Negative discount values | 15 |
| Discount | High discount, including percentage-format (e.g. 70%, 85%) | 15 |
| Calculation | Sales mismatch | 49 |

**Final dataset: 912 deduplicated records — 751 fully clean, 161 carrying at least one quality flag for review.**

---

## Summary of Final Pivot Reports

| Pivot | Key Finding |
|---|---|
| By Region | South leads in total sales, but East shows a higher profit despite slightly lower sales — South converts revenue to profit less efficiently than East |
| By Category | Technology generates strong profit relative to volume; Furniture has higher volume but thinner margins |
| By Ship Mode | Standard and Second Class dominate order volume; Same Day is lower-volume |
| By Segment | Small Business segment shows the highest average profit margin (~30%), Consumer the lowest (~25%) |
| Non-Completed by Region | Cancelled/Returned orders are spread fairly evenly across regions, no single region stands out as a major outlier |
| Monthly Sales Trend | Visible month-to-month volatility across both years rather than a single seasonal pattern; no single quarter dominates |

---

## Key Business Insights

1. **South and East have similar sales, but East converts more efficiently to profit** — worth investigating what's driving the cost/margin difference between these two regions.
2. **49 sales calculation mismatches** suggest inconsistent rounding or discount application at the source system.
3. **93 ship-before-order-date records** were flagged using a day-first (DD/MM) reading of ambiguous slash-formatted dates. This is a meaningfully large share of the dataset (~10%), and the count is sensitive to that interpretation choice — reading the same ambiguous dates month-first instead would flag a much smaller set (~22 records). This ambiguity is a genuine limitation of the raw export, not a cleaning error, and is worth flagging to the data owner directly.
4. **12 conflicting duplicate order IDs** point to a data integrity issue upstream — the same order ID is being reused across source systems being merged.
5. **"Unknown" region records (25)** could likely be resolved with a state-to-region reference table, which wasn't provided for this assignment.

---

## Assumptions and Limitations

**Assumptions:**
- Ambiguous slash-formatted dates (e.g. "03/01/2024") were treated as day-first (DD/MM), consistent with the Indian business context of this dataset
- Missing discount = no discount applied (0), not a data loss issue
- The first occurrence of any exact duplicate pair is treated as the canonical record

**Limitations:**
- The day-first date assumption is a genuine judgment call, not a verified fact — the raw export does not specify its date convention, and reading the same ambiguous dates month-first instead would produce a materially different (much smaller) ship-before-order-date count. This is flagged as the single largest source of uncertainty in this analysis and should be confirmed with the source system owner before the 93 flagged records are acted on.
- Sales mismatches could not be fully root-caused without access to the source system's rounding/calculation logic
- Conflicting duplicate order IDs require input from the data owner to resolve which record is correct
- No state-to-region mapping table was available, so "Unknown" region values could not be inferred

---

## Screenshots Included

| File | Contents |
|---|---|
| `screenshots/raw_data_preview.png` | Raw dataset before cleaning, showing format inconsistencies |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with calculated columns and quality flags |
| `screenshots/pivot_summary_1.png` | By Region pivot table with clustered column chart (Sales vs Profit by Region) |
| `screenshots/pivot_summary_2.png` | Monthly Sales Trend pivot table with line chart (2024–2025) |

---

## Repository Structure

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx           ← Original untouched dataset
│   └── cleaned_orders.xlsx       ← Cleaned, validated, analysis-ready dataset (912 rows)
├── outputs/
│   ├── data_quality_report.xlsx  ← Full data quality report
│   ├── pivot_summary.xlsx        ← 6 pivot tables + charts, filtered to exclude conflicting duplicates
│   └── cleaning_log.md           ← Full documentation of cleaning decisions
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md
```
