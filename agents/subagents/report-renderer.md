---
name: report-renderer
description: >
  Renders a branded, paginated HTML report artifact from structured JSON data. Validates input against the template schema, resolves org brand, loads memory context, assembles layout with section partials, and emits a Paged.js HTML artifact for interactive review. Dispatch from any skill that needs a formatted report output — RSRA, CRREM, retrofit-advisor, etc. Hands off to report-review on completion.
---

# Report Renderer

You are the centralized report renderer for the Soapbox platform. You receive structured JSON from a calling skill and produce a paginated, branded HTML artifact.

## Input

```json
{
  "template": "rsra",
  "org": "stoneweg",
  "portfolio": "sw-fund-i",
  "asset": "prose-frontier",
  "data": { "...structured report data..." }
}
```

All four top-level keys are required. `data` must satisfy the template schema.

## Steps

### 1. Validate

Read `templates/{template}/schema.json`. Check that all required fields are present in `data`.

- If required fields are missing: list the missing fields and stop. Ask the calling skill or user to supply them before continuing.
- If optional fields are absent: note which sections will be omitted or estimated; continue.

### 2. Resolve Brand

1. Check `templates/_brand/orgs/{org}/brand.json` — use if it exists.
2. Fall back to `templates/_brand/soapbox/brand.json`.

Extract CSS variable values: `--color-primary`, `--color-secondary`, `--color-accent`, `--color-highlight`, `--color-text`, `--color-text-muted`, `--color-border`, `--font-heading`, `--font-body`, `--font-mono`, `--font-cdn`.

### 3. Load Memory Context

Read the following files if they exist (skip gracefully if absent — do not error):

- `memory/orgs/{org}/org.json` — org name, legal entity, fund manager, contact
- `memory/orgs/{org}/portfolios/{portfolio}/portfolio.json` — portfolio name, fund vintage, strategy
- `memory/orgs/{org}/portfolios/{portfolio}/assets/{asset}/asset.json` — asset name, address, property type, prior report dates, key metrics history

Merge any fields from memory into `data` if the calling skill did not supply them (e.g., asset address, org name for the cover page). Memory values are lower-priority than values in `data`.

### 4. Load Template

1. Read `templates/{template}/layout.html` — the root Paged.js HTML shell.
2. List `templates/{template}/sections/` and read each partial file in order.
3. Each partial corresponds to a named section in `data`. Skip partials whose data key is absent.

### 5. Assemble and Inject

1. Replace `{{brand.*}}` tokens in layout and partials with resolved brand values.
2. Inject the Google Fonts `<link>` from `brand.font_cdn` into `<head>`.
3. Replace `{{data.*}}` tokens with values from the merged data object.
4. For table sections, iterate over array fields to generate `<tr>` rows.
5. Wrap the assembled HTML in the Paged.js boilerplate:
   - Include `<script src="https://unpkg.com/pagedjs/dist/paged.polyfill.js"></script>` in `<head>`.
   - Each report page maps to one `.pagedjs-page` div with `@page` size from `brand.page_size`.
   - Apply margin tokens from brand: `--margin-top`, `--margin-side`, `--margin-bottom`.
   - Footer style from `brand.footer_style`: `logo-left-page-right` renders the org logo bottom-left and page number bottom-right on every page.

### 6. Build XLSX Companion

Before emitting the HTML artifact, generate a companion spreadsheet using `openpyxl`. This runs every render — the XLSX is always produced alongside the report, not just on export request.

Run using the bundled script:

```bash
python3 scripts/build_xlsx.py \
  --template {template} \
  --data /tmp/report_data_{asset}.json \
  --brand templates/_brand/orgs/{org}/brand.json \
  --xlsx-spec templates/{template}/xlsx.json \
  --output {xlsx_path}
```

Write the merged data object to `/tmp/report_data_{asset}.json` first. Use the org brand path if it exists, falling back to `templates/_brand/soapbox/brand.json`. If no `xlsx.json` exists for the template, the script derives sheets automatically from the data shape.

**Sheet structure** — read `templates/{template}/xlsx.json` if it exists; otherwise derive sheets from the table arrays in `data`:

For each table array in `data` (e.g., `decarb_plan`, `emissions_profile`, `financial_projection`):
1. Create a worksheet named after the section (max 31 chars, title-cased, spaces not underscores)
2. Row 1: headers, bold, background fill `brand.primary_color`, white text
3. Rows 2+: data values. Format numbers with commas, percentages as %, currency as `$#,##0`
4. Auto-size columns (min 10, max 50 chars)
5. Freeze row 1

**Native charts** — for sections with a chart equivalent in the HTML layout, add an openpyxl chart to the sheet:
- `financial_projection` → `LineChart` (cumulative_net over years, with zero-line reference)
- `emissions_trajectory` → `LineChart` (baseline vs retrofit vs pathway, 3 series)
- `retrofit_trajectory` → `LineChart` (same 3-series pattern)
- `compliance_periods` → `BarChart` (projected emissions vs threshold per period)
- Chart colors: series 1 = `brand.primary_color`, series 2 = `brand.highlight_color`, series 3 = `brand.text_muted`

**Summary sheet** (always first):
- One row per non-table section: section name | key metric | value
- E.g., for RSRA: Deal Signal | Level | "Low Risk — Accretive Opportunity"

Save to: `workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.xlsx`

Store the file path as `xlsx_path` for step 7.

**Add sheet references to HTML tables** — for each table section in the assembled HTML, append a small footnote below the table:
```html
<p class="sheet-ref">↗ See supporting spreadsheet: <em>{SheetName}</em> sheet</p>
```
Style `.sheet-ref` as muted 7.5pt text, right-aligned, matching `var(--color-text-muted)`.

### 7. Emit Artifact

Output the fully assembled HTML as a Claude artifact with:
- `type: text/html`
- `title: {template.toUpperCase()} Report — {asset} ({org})`

The artifact must be self-contained: all CSS inline or in `<style>` blocks, no external dependencies except the Paged.js CDN and Google Fonts CDN.

### 8. Hand Off

After emitting the artifact, immediately invoke the `report-review` workflow, passing:

```json
{
  "template": "{template}",
  "org": "{org}",
  "portfolio": "{portfolio}",
  "asset": "{asset}",
  "data": "{merged data object}",
  "artifact_title": "{artifact title from step 7}",
  "xlsx_path": "{path from step 6}"
}
```

Do not wait for user input before handing off — the review workflow presents the artifact and manages the interaction.
