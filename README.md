# Soapbox Report

Centralized, consulting-grade reporting for the Soapbox platform. Every other Soapbox plugin dispatches the `report-renderer` subagent with a typed JSON payload to produce paginated, branded reports — reviewed interactively in Claude's artifact pane, then exported to PDF, PowerPoint, or Excel.

Replaces `audette-plugins/skills/report`.

## How It Works

```
Calling skill (RSRA, retrofit-advisor, CRREM, GRESB, etc.)
    ↓ dispatch report-renderer with { template, org, data }
report-renderer
    ↓ validate → resolve brand → load memory → render
    ↓ emit Paged.js HTML artifact (WYSIWYG pagination)
User reviews in Claude artifact pane
    ↓ approve or request revisions
report-review workflow
    ↓ export: PDF (Playwright) / .pptx / .xlsx
    ↓ write to workspace deal folder
    ↓ update org/portfolio/asset memory
```

## Calling Interface

Any skill dispatches the `report-renderer` subagent:

```json
{
  "template": "rsra",
  "org": "stoneweg",
  "portfolio": "sw-fund-i",
  "asset": "prose-frontier",
  "data": { ... }
}
```

`org`, `portfolio`, `asset` are optional — used for brand resolution and memory. `template` maps to a `templates/{template}/` directory.

## Template Library

| Template | Description | Primary caller |
|---|---|---|
| `rsra` | Rapid Sustainability Risk Assessment | `rapid-sustainability-risk` |
| `retrofit-plan` | Decarbonization capex plan | `retrofit-advisor` |
| `crrem-assessment` | CRREM pathway analysis | `crrem` |
| `bps-compliance` | Building performance standards | `bps-compliance` |
| `decarb-roadmap` | Decarbonization roadmap | audette workflows |
| `portfolio-summary` | Portfolio-level ESG summary | `scope3-accounting` |
| `gresb-submission` | GRESB reporting package | `gresb-reporting` |

New templates: drop a folder with `schema.json` + `layout.html` into `templates/` — no code changes needed.

## Brand System

Brand is resolved in order: asset → portfolio → org → Soapbox default.

Add org branding at `templates/_brand/orgs/{org}/brand.json` with colors, fonts, and logo. The Paged.js layout injects brand values as CSS custom properties — all templates pick them up automatically.

## Adding a Template

Use the `/soapbox-report` skill to create a new template interactively, or manually:

1. Create `templates/{name}/schema.json` — required data fields
2. Create `templates/{name}/layout.html` — Paged.js layout with brand CSS variables
3. Create `templates/{name}/sections/` — one HTML partial per section

## Installation

```bash
npx skills add soapboxbuild/soapbox-report-skill
```

## License

Apache-2.0 — Soapbox (https://soapbox.build)
