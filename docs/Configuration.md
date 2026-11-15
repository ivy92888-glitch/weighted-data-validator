# Configuration

This document explains every adjustable constant in `src/WeightedDataValidator.bas`.
Source code is the single source of truth — values listed here reference *names* only.

## Constants

### Scoring Weights

| Constant | Type | Default | What it does |
|---|---|---|---|
| `WEIGHT_FREQUENCY` | Double | 0.4 | Weight for "how many times this value appeared" in the reference DB |
| `WEIGHT_RECENCY`   | Double | 0.6 | Weight for "how recently this value was last seen" |

**How to tune:** If your reference DB has many historical stale records
(e.g. old deprecated codes that still appear often), bump up `WEIGHT_RECENCY`.
If your data rarely changes and frequency is the real signal, raise
`WEIGHT_FREQUENCY`. The two should sum to roughly 1.0.

### Text Processing

| Constant | Type | Default | What it does |
|---|---|---|---|
| `MIN_NAME_LEN`      | Integer | 3  | Ignore extracted suggestions shorter than this |
| `MAX_TRUNCATE_LEN`  | Integer | 18 | Truncation threshold for non-numeric description clauses |
| `NAME_SEPARATOR`    | String  | `/` | Delimiter when picking last alternative (`A / B / C` → `C`) |
| `RULE_MARK`         | String  | `-` | Prefix prepended to rule annotations |

### Color Palette

All use Excel's `RGB(r, g, b)`. Adjust freely to match your company template.

| Constant | Default | Purpose |
|---|---|---|
| `CLR_SUGGESTED` | RGB(226, 239, 218) | Background of cells the validator wants you to review |
| `CLR_PART_ERR`  | RGB(127, 150, 152) | Color of the *wrong* part inside a modified description |
| `CLR_PART_OK`   | RGB(198,  89,  17) | Color of the *correct* part inside a modified description |
| `CLR_HEADER`    | RGB(255, 255,   0) | Background of the Feedback sheet header row |

### Priority Keyword Arrays

| Variable | Type | What it does |
|---|---|---|
| `primaryKeywords`   | Array | High-priority suggestion keywords, ordered from highest to lowest |
| `secondaryKeywords` | Array | Low-priority keywords — only checked if no primary keyword matched |

**To customize:** edit the `Initialize` subroutine to change the arrays.
The algorithm walks them in array-index order — smaller index = higher
priority.

## Workbook Layout Assumptions

| Sheet Name | Required Columns (letter + header) |
|---|---|
| **Data** | A: ID, B: Raw Link, C: Category Name (CN), D: Category Name (EN), E: Suggested Name, F: Suggested Code, G: Suggested Rate, H: Notes |
| **ReferenceCodes** | A: Lookup Code (category), B: Code (candidate value), C: Notes / Source |
| **Rules** | A: Keyword to match, B: Annotation text to append |

Row 1 is always the header row. Data starts at row 2.

## Second Domain Example — E-commerce SKU Normalization

To prove generality, here's how you'd re-point this tool at a completely
different use case:

### Original scenario → Second domain mapping

| Original scenario (customs declaration) | Second scenario (e-commerce SKU) |
|---|---|
| Code consistency check | **SKU consistency check across listings** |
| Anti-dumping rule engine | **Restricted-material rule engine** (flammable, food-contact, etc.) |
| Euro-code table as reference DB | **Historical SKU usage log** as reference DB |
| `amazon.*` → `amazon.de` domain normalize | **Any marketplace URL → canonical product-page format** |
| Color-coded Feedback sheet | Same output — mechanism is fully generic |

### What you actually change

1. **No code changes needed.** Just populate the three sheets with new data:
   - `ReferenceCodes` = your historical SKU assignments
   - `Rules` = your restricted-material annotations
2. **If column layout differs**, read the source constants at the top of the
   code and adjust the column-letter references in each subroutine. Or copy
   the same layout convention — the tool was designed around it for a reason.

That's it. Same code, same algorithm, different domain.

---

## Reference Database Schema Templates

> ⚠️ **Reference database tables are NOT included in this repository.**
> They contain your organization's historical records and are for **your use
> only.** This section explains what columns your `ReferenceCodes` sheet
> *can* have — pick and choose based on your domain.
>
> The tool was originally used alongside a larger workbook that maintained
> three lookup tables internally (each with ~10,000+ rows of historical
> records). These tables drove the weighted-scoring algorithm's data source.
> Below are the genericized schemas for reference. Adjust freely — the code
> only reads `Lookup Code` + `Code`, the rest are optional metadata.

### Template A — Full Rich Lookup (6 columns)

Used when you have rich historical context per record.

| Column | Header | Purpose | Example (neutral) |
|---|---|---|---|
| A | Lookup Term | What you group records by (the "category" the validator looks up) | Garden Cart |
| B | Candidate Value | The value that gets scored by frequency + recency | 7326909890 |
| C | External Link | Product / record URL (only used for context, not by the algorithm) | https://amazon.example-domain/dp/B0X123 |
| D | Material / Spec | Optional metadata — what material or spec was it? | plastic |
| E | Batch / Unit ID | Optional — which batch or container was this from? | BATCH-2024-003 |
| F | Display Name (EN) | Human-readable English name for reference | Garden Utility Cart |

### Template B — Compact Lookup (10 columns)

When your reference table comes from an operational system that records
everything — prices, shipping weights, sources, etc.

| Column | Header | Purpose | Example (neutral) |
|---|---|---|---|
| A | Lookup Term | Grouping term | Lantern |
| B | Display Name (EN) | Human-readable English name | Paper Lantern |
| C | Candidate Value | The value to score | 4823908580 |
| D | Percent Parameter | Percent value historically seen with this value | 0 |
| E | Unit Price | Historical selling price (for reference only) | 10.90 |
| F | Unit Weight | Historical weight per unit | 0.23 kg |
| G | External Link | Product URL | https://amazon.example-domain/dp/B0X123 |
| H | Material / Spec | Material note | paper |
| I | Source | Who / what system entered this record | Team-A |
| J | Price per Unit | Derived metric (not used by validator) | 47.39 |

### Template C — Code + Description (4 columns)

A simpler reference for lookup + description pairs. Useful when you don't
need full history tracking.

| Column | Header | Purpose | Example (neutral) |
|---|---|---|---|
| A | Lookup Term | Grouping term | Pebble |
| B | Candidate Value | Code / identifier to score | 2517101000 |
| C | Percent Parameter | Percent value | 0 |
| D | Description | Human-readable description of the code | Pebbles, gravel, broken stone |

### Which template to pick?

Start with **Template A** if you're not sure — it has the minimum viable
columns for the weighted-scoring algorithm to work well.

All three templates can coexist in the same workbook as different sheets.
Point the `CalculateBestCode` function's `refSheetName` argument at
whichever one contains your historical records.

> 💡 **Row ordering matters.** The weighted algorithm walks the sheet
> **bottom-up**. New records should go at the bottom (higher row numbers).
> Older records stay above. This is how "recency" is calculated.
