---
name: report-review
description: >
  Interactive review and export workflow for rendered reports. Presents the Paged.js artifact to the user, accepts revision requests, re-renders on change, and on approval exports to PDF (Playwright), PPTX (python-pptx), and/or XLSX (openpyxl). Writes export paths and report metadata back to asset memory. Called by report-renderer after every render.
---

# Report Review

You manage the interactive review cycle and export pipeline for all Soapbox reports.

## Inputs (from report-renderer)

```json
{
  "template": "rsra",
  "org": "stoneweg",
  "portfolio": "sw-fund-i",
  "asset": "prose-frontier",
  "data": { "...merged data object..." },
  "artifact_title": "RSRA Report — prose-frontier (stoneweg)"
}
```

## Steps

### 1. Present

Tell the user:
> "Your **{artifact_title}** is ready. Review it in the artifact pane. Let me know if you'd like any changes, or say **approve** to export."

### 2. Revision Loop

If the user requests a change:

1. Apply the change to `data` (or to the layout/brand token if it's a structural or style change).
2. Re-dispatch `report-renderer` with the updated inputs.
3. After the new artifact is emitted, return to step 1.

Handle common revision types:
- **Content change** — update the relevant field in `data` and re-render.
- **Section reorder** — update `data.section_order` (array of section names) and re-render.
- **Brand/style change** — update the brand token in `data.brand_overrides` and re-render; do not modify the stored `brand.json`.
- **Add/remove section** — add or remove the section key from `data` and re-render.

Continue the loop until the user says "approve", "looks good", "export", or an equivalent.

### 3. Export Prompt

On approval, ask:
> "Export as **PDF**, **PowerPoint**, **Excel**, or **all three**?"

If the user has already specified a format earlier in the conversation, use that and skip the prompt.

### 4. PDF Export

1. Use the Playwright MCP to navigate to the artifact.
   - If a local artifact URL is available, use it. Otherwise ask the user to open the artifact and share the URL, or export via `print` from the browser.
2. Wait for Paged.js to finish paginating (wait for `window.PagedPolyfill.done` or a 3-second settle delay).
3. Call `page.pdf()` with:
   - `format`: match `brand.page_size` (default: `Letter`)
   - `printBackground`: `true`
   - `margin`: `{ top: brand.margin_top, right: brand.margin_side, bottom: brand.margin_bottom, left: brand.margin_side }`
4. Save to: `workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.pdf`
5. Confirm file was written and report the path.

### 5. PPTX Export

Run `python3` with `python-pptx` installed (install via `pip install python-pptx` if absent):

1. Create a new `Presentation()` with slide size matching `brand.page_size`.
2. For each section in `data.sections` (in order):
   - Add a slide using a blank layout.
   - Add a title text box from `section.title`.
   - For table sections: render an `add_table()` with headers and rows from `section.rows`.
   - For narrative sections: render a text box from `section.body`.
   - Apply brand colors to title and table headers (`brand.primary_color`, `brand.highlight_color`).
3. Save to: `workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.pptx`
4. Confirm file was written and report the path.

### 6. XLSX Export

Run `python3` with `openpyxl` installed (install via `pip install openpyxl` if absent):

1. Create a new `Workbook()`.
2. For each table section in `data.sections`:
   - Add a worksheet named from `section.title` (truncated to 31 chars).
   - Write headers in row 1, bold, with background fill from `brand.primary_color`.
   - Write data rows starting at row 2.
   - Auto-size columns.
3. For non-table sections (narrative, deal signal): add a single "Summary" sheet with a text cell per section.
4. Save to: `workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.xlsx`
5. Confirm file was written and report the path.

### 7. Update Memory

Write or update `memory/orgs/{org}/portfolios/{portfolio}/assets/{asset}/asset.json`:

- `last_report_date`: today's date (ISO 8601)
- `last_report_template`: `{template}`
- `reports`: append `{ "date": "...", "template": "...", "pdf": "...", "pptx": "...", "xlsx": "..." }` to the array (create array if absent)
- Preserve all other existing fields

If the file does not exist, create it with only the fields above.

### 8. Return

Return to the calling skill:

```json
{
  "status": "exported",
  "pdf": "workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.pdf",
  "pptx": "workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.pptx",
  "xlsx": "workspace/orgs/{org}/portfolios/{portfolio}/assets/{asset}/reports/{template}-{YYYY-MM-DD}.xlsx"
}
```

Omit keys for formats the user did not request.
