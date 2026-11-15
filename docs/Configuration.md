# Configuration

This document explains every adjustable constant in `src/WeightedDataValidator.bas`.
Source code is the single source of truth — values listed here reference *names* only.

## Two-Phase Pipeline

The code has **two independent pipelines**. Run them together via `RunFullPipeline()`,
or individually via `RunPipelinePhase1()` / `RunPipelinePhase2()`.

| Pipeline | What it does | Safe to rerun? |
|---|---|---|
| **Phase 1 — Consistency Check** | Validates + marks cells in-place on the Data sheet | ✅ Color marks are overwritten |
| **Phase 2 — Process + Report** | Cleans legacy marks, sanitizes links, generates Feedback sheet | ✅ Feedback sheet is replaced each run |

---

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
| `MAX_TRUNCATE_LEN` | Integer | 18 | Truncation threshold for non-numeric description clauses |
| `NAME_SEPARATOR`   | String  | `/` | Delimiter when picking last alternative (`A / B / C` → `C`) |
| `RULE_MARK`        | String  | `-` | Prefix prepended to rule annotations |
| `TARGET_DOMAIN`    | String  | `"amazon.example-domain"` | Domain suffix used when reconstructing sanitized entity links |

### Column Letter Constants

Adjust these to match your Data sheet layout. Each is a single Excel column letter.

| Constant | Default | What it is |
|---|---|---|
| `COL_LOOKUP`           | A | Grouping key — what you search for in ReferenceCodes |
| `COL_LINK`             | B | External / entity link — gets sanitized in Phase 2 |
| `COL_NAME_RAW`         | C | Original name (human display) |
| `COL_NAME_EN_RAW`      | D | Original English / display name — parsed for suggestions in Phase 1 |
| `COL_CODE_RAW`         | E | Original identifier / code — parsed for suggestions in Phase 1 |
| `COL_SUGGESTION`       | F | Match result written by Phase 1 ("code-rate" or "N/A") |
| `COL_SUGGESTION_NAME`  | G | Structured suggested name — written by Phase 1 |
| `COL_SUGGESTION_CODE`  | H | Structured suggested code / value — written by Phase 1 |
| `COL_RATE`             | I | Structured percent parameter — written by Phase 1 |
| `COL_METADATA`         | J | Optional notes — rule annotations appended here |

### Worksheet Name Constants

| Constant | Default | What it is |
|---|---|---|
| `SHEET_DATA`        | `"Data"` | Your main working sheet (what you open and run the macro on) |
| `SHEET_REF_CODES`   | `"ReferenceCodes"` | Reference table for weighted scoring lookup |
| `SHEET_RULES`       | `"Rules"` | Optional keyword → annotation rules |
| `SHEET_FEEDBACK`    | `"Feedback"` | Generated report sheet (recreated each Phase 2 run) |
| `SHEET_REF_LOOKUP`  | `"ReferenceLookup"` | Optional table for code → rate lookups |

### Color Palette

All use Excel's `RGB(r, g, b)`. Adjust freely.

| Constant | Default | Purpose |
|---|---|---|
| `CLR_SUGGESTED_BG`   | (226, 239, 218) | Light green — cell background when Phase 1 detects a free-text suggestion |
| `CLR_WRONG_PART`     | (127, 150, 152) | Slate gray-green — wrong portion inside a description |
| `CLR_RIGHT_PART`     | (198,  89,  17) | Orange — correct portion inside a description |
| `CLR_HEADER_BG`       | (255, 255,   0) | Yellow — Feedback sheet header row background |
| `CLR_MATCH_OK_FONT`   | (  0, 128,   0) | Dark green font — Phase 1 writes this on the match column when original code matches history |
| `CLR_MISS_FONT`       | (192,   0,   0) | Dark red font — Phase 1 writes this when no match or N/A |
| `CLR_PREVIOUS_MARK`   | (146, 208,  80) | Legacy green — cleared in Phase 2 so old marks don't carry over |

### Priority Keyword Arrays

These are **Populated at runtime** in `InitializePipeline()`, not constants.

| Variable | Type | What it does |
|---|---|---|
| `primaryKeywords`   | Array | High-priority suggestion keywords, ordered highest → lowest |
| `secondaryKeywords` | Array | Low-priority keywords — only checked if no primary keyword matched |

**Defaults:**
```
primaryKeywords   = Array("top", "try", "suggest", "recommend")
secondaryKeywords = Array("may", "let", "allow", "consider")
```

The `ExtractSuggestion()` function walks both arrays. Lower index = higher
priority. Keywords from the primary array always beat secondary.

---

## Workbook Layout Assumptions

| Sheet | Required Columns |
|---|---|
| **Data** | Row 1–3 = headers; data starts row 4. Columns A through J as defined above. |
| **ReferenceCodes** | A: Lookup Term (grouping key), B: Candidate Value (the value to score), C+: Optional metadata. |
| **Rules** | A: Keyword, B: Annotation (free-text note to append on match). |
| **ReferenceLookup** | A: Lookup Term, B: Code, C: Percent parameter (optional). |

Row ordering for **ReferenceCodes** matters for weighted scoring.
Newest records always go at the **bottom** (highest row numbers).
The algorithm walks bottom-up — higher row number = "more recent".

---

## Second Domain Example — E-commerce SKU Normalization

To prove generality, here's how you'd re-point this tool at a completely
different use case:

| Original scenario | Second scenario |
|---|---|
| Code consistency check | **SKU consistency check across listings** |
| Rule engine on product names | **Restricted-material rule engine** (flammable, food-contact, etc.) |
| Euro-code table as reference DB | **Historical SKU usage log** as reference DB |
| Any marketplace URL → normalize domain | **Any marketplace URL → canonical product-page format** |
| Color-coded Feedback sheet | Same mechanism — fully generic |

### What you actually change

1. **No code changes needed.** Just populate the three sheets with new data.
2. **If column layout differs**, read the `COL_*` constants at the top of the
   code and adjust the column letters. Or copy the same layout convention.
3. Adjust `primaryKeywords` / `secondaryKeywords` in `InitializePipeline()`
   to match your domain's free-text suggestion style.

That's it. Same code, same algorithm, different domain.
