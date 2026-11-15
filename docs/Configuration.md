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
