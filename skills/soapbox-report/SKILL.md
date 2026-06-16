---
name: soapbox-report
description: >
  Meta-skill for creating new report templates in the soapbox-report plugin. Guides the user through naming the template, defining the schema (required and optional data fields), designing the page layout (cover, sections, footer), writing section partials, and generating an example render. Use when the user wants to create a report template, add a new report type, design a report layout, build a new Soapbox report, or add a template for a new skill. Triggers on: create a report template, add a new report type, design a report layout, new template, report template, build a report, add a template, template for RSRA/CRREM/retrofit/any skill.
---

# Soapbox Report: Template Creator

You help users design and build new report templates for the `soapbox-report` plugin. A template is a reusable layout that `report-renderer` assembles with live data from any calling skill.

## What You Build

A complete template lives at `templates/{template-name}/` and contains:

```
templates/{template-name}/
  schema.json          # required + optional data fields
  layout.html          # Paged.js root shell
  sections/            # one .html partial per report section
    cover.html
    {section-name}.html
    ...
  example/
    data.json          # sample data matching schema
    render-notes.md    # brief description of the example render
```

---

## Step 1: Name the Template

Ask the user:
> "What should this template be called? Use a short, lowercase, hyphenated identifier (e.g., `rsra`, `crrem`, `retrofit-plan`, `annual-esg`)."

Validate: lowercase letters, digits, hyphens only. No spaces. Must be unique — check `templates/` for existing names.

---

## Step 2: Define the Schema

Ask the user to describe the data this report will display. For each field, determine:

| Field | Type | Required? | Description |
|---|---|---|---|
| `asset_name` | string | yes | Displayed on cover page and header |
| `...` | ... | ... | ... |

Common field types: `string`, `number`, `date` (ISO 8601), `array<object>` (for tables), `array<string>` (for lists), `object` (for nested groupings).

Write `templates/{template-name}/schema.json`:

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "{Template Name} Report Data",
  "type": "object",
  "required": ["field1", "field2"],
  "properties": {
    "field1": { "type": "string", "description": "..." },
    "field2": { "type": "number", "description": "..." }
  }
}
```

Show the schema to the user and confirm before proceeding.

---

## Step 3: Design the Layout

Ask the user:
> "Describe the pages and sections for this report. A typical layout is: **Cover → Executive Summary → [Data Sections] → Appendix**. What sections do you need?"

For each section, collect:
- **Section name** (used as the partial filename and the `data` key)
- **Section type**: `cover`, `narrative` (prose), `table`, `chart-placeholder`, `kpi-grid`, `appendix`
- **Content**: what data fields go in it

Write `templates/{template-name}/layout.html` — the Paged.js root shell:

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>{{data.asset_name}} — {Template Display Name}</title>
  <link rel="stylesheet" href="{{brand.font_cdn}}">
  <style>
    :root {
      --color-primary: {{brand.primary_color}};
      --color-secondary: {{brand.secondary_color}};
      --color-accent: {{brand.accent_color}};
      --color-highlight: {{brand.highlight_color}};
      --color-text: {{brand.text_color}};
      --color-text-muted: {{brand.text_muted}};
      --color-border: {{brand.border_color}};
      --font-heading: {{brand.font_heading}}, sans-serif;
      --font-body: {{brand.font_body}}, sans-serif;
      --font-mono: {{brand.font_mono}}, monospace;
    }
    @page {
      size: {{brand.page_size}};
      margin: {{brand.margin_top}} {{brand.margin_side}} {{brand.margin_bottom}};
    }
    /* base reset, typography, table styles */
  </style>
  <script src="https://unpkg.com/pagedjs/dist/paged.polyfill.js"></script>
</head>
<body>
  {{> cover}}
  {{> section-name}}
  {{> ...}}
</body>
</html>
```

Explain the `{{> partial-name}}` slot convention to the user: each slot is replaced by the corresponding section partial during rendering.

---

## Step 4: Write Section Partials

For each section, write `templates/{template-name}/sections/{section-name}.html`.

**Cover partial pattern:**
```html
<div class="page cover-page">
  <div class="cover-logo"><img src="{{brand.logo_path}}" alt="{{brand.name}}"></div>
  <h1 class="cover-title">{{data.asset_name}}</h1>
  <p class="cover-subtitle">{Template Display Name}</p>
  <p class="cover-date">{{data.report_date}}</p>
  <div class="cover-footer">{{data.org_name}} · Confidential</div>
</div>
```

**Table section partial pattern:**
```html
<div class="page section-page">
  <h2 class="section-title">{Section Display Name}</h2>
  <table class="data-table">
    <thead>
      <tr>
        {{#each data.{section_key}.headers}}<th>{{this}}</th>{{/each}}
      </tr>
    </thead>
    <tbody>
      {{#each data.{section_key}.rows}}
      <tr>
        {{#each this}}<td>{{this}}</td>{{/each}}
      </tr>
      {{/each}}
    </tbody>
  </table>
  {{#if data.{section_key}.footnote}}
  <p class="footnote">{{data.{section_key}.footnote}}</p>
  {{/if}}
</div>
```

**Narrative section partial pattern:**
```html
<div class="page section-page">
  <h2 class="section-title">{Section Display Name}</h2>
  <div class="narrative">{{data.{section_key}.body}}</div>
</div>
```

Write each partial. After writing, show the user the full list of partials and confirm the set is complete.

---

## Step 5: Create an Example

Ask the user:
> "Provide sample values for each required field, or I can generate placeholder values from the schema."

Write `templates/{template-name}/example/data.json` with valid sample data satisfying the schema.

Write `templates/{template-name}/example/render-notes.md` — a brief description (3–5 lines) of what the example render should look like: which sections appear, what the cover says, any notable layout choices.

---

## Step 6: Confirm and Test Render

Show the user a summary:

```
Template: {template-name}
Schema:   {N} required fields, {M} optional fields
Sections: cover → {section-1} → {section-2} → ...
Files:    templates/{template-name}/schema.json
          templates/{template-name}/layout.html
          templates/{template-name}/sections/{N files}
          templates/{template-name}/example/data.json
          templates/{template-name}/example/render-notes.md
```

Ask: "Ready to test render with the example data?"

If yes, dispatch `report-renderer` with:
```json
{
  "template": "{template-name}",
  "org": "soapbox",
  "portfolio": "example",
  "asset": "example-asset",
  "data": { "...contents of example/data.json..." }
}
```

The rendered artifact confirms the template is wired correctly. Iterate on layout or partial issues before declaring the template done.
