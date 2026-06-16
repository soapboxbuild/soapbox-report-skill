#!/usr/bin/env node
/**
 * export_pdf.js — Node.js fallback for Paged.js HTML → PDF export
 *
 * Used by export_pdf.py when the Python playwright package is unavailable.
 * Requires: npm install playwright  (or npx playwright)
 *
 * Usage:
 *   node export_pdf.js --html /path/to/report.html --output /path/to/report.pdf [options]
 *
 * Options:
 *   --html         Path to the source HTML file (required)
 *   --output       Destination PDF path (required)
 *   --page-size    Letter or A4 (default: Letter)
 *   --margin-top   Top margin, e.g. 18mm (default: 18mm)
 *   --margin-side  Side margin, e.g. 16mm (default: 16mm)
 *   --margin-bottom Bottom margin, e.g. 20mm (default: 20mm)
 *   --timeout      Paged.js wait timeout in ms (default: 30000)
 */

"use strict";

const path = require("path");
const fs = require("fs");

// ---------------------------------------------------------------------------
// Arg parsing
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const args = {
    html: null,
    output: null,
    pageSize: "Letter",
    marginTop: "18mm",
    marginSide: "16mm",
    marginBottom: "20mm",
    timeout: 30000,
  };

  for (let i = 2; i < argv.length; i++) {
    const key = argv[i];
    const val = argv[i + 1];
    switch (key) {
      case "--html":           args.html        = val; i++; break;
      case "--output":         args.output      = val; i++; break;
      case "--page-size":      args.pageSize    = val; i++; break;
      case "--margin-top":     args.marginTop   = val; i++; break;
      case "--margin-side":    args.marginSide  = val; i++; break;
      case "--margin-bottom":  args.marginBottom = val; i++; break;
      case "--timeout":        args.timeout     = parseInt(val, 10); i++; break;
      default:
        if (key.startsWith("--")) {
          console.error(`Unknown option: ${key}`);
          process.exit(1);
        }
    }
  }

  if (!args.html) { console.error("--html is required"); process.exit(1); }
  if (!args.output) { console.error("--output is required"); process.exit(1); }

  args.html = path.resolve(args.html);
  args.output = path.resolve(args.output);

  if (!fs.existsSync(args.html)) {
    console.error(`HTML file not found: ${args.html}`);
    process.exit(1);
  }

  return args;
}

// ---------------------------------------------------------------------------
// Serve the HTML file on a local HTTP server to avoid file:// CDN restrictions
// ---------------------------------------------------------------------------

function serveFile(htmlPath) {
  const http = require("http");
  const dir = path.dirname(htmlPath);
  const basename = path.basename(htmlPath);

  const server = http.createServer((req, res) => {
    // Only serve files within the HTML's directory (security boundary)
    const safePath = path.join(dir, path.normalize(req.url).replace(/^(\.\.[/\\])+/, ""));
    fs.readFile(safePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }
      const ext = path.extname(safePath).toLowerCase();
      const mime = {
        ".html": "text/html",
        ".css":  "text/css",
        ".js":   "application/javascript",
        ".svg":  "image/svg+xml",
        ".png":  "image/png",
        ".jpg":  "image/jpeg",
        ".woff": "font/woff",
        ".woff2":"font/woff2",
      }[ext] || "application/octet-stream";
      res.writeHead(200, { "Content-Type": mime });
      res.end(data);
    });
  });

  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      resolve({ server, url: `http://127.0.0.1:${port}/${basename}` });
    });
  });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const args = parseArgs(process.argv);

  let playwright;
  try {
    playwright = require("playwright");
  } catch {
    console.error(
      "playwright is not installed. Run:\n  npm install playwright\n  npx playwright install chromium"
    );
    process.exit(2);
  }

  // Ensure output directory exists
  const outDir = path.dirname(args.output);
  fs.mkdirSync(outDir, { recursive: true });

  // Serve locally to avoid file:// CDN load failures
  const { server, url } = await serveFile(args.html);

  let browser;
  try {
    browser = await playwright.chromium.launch({ headless: true });
    const context = await browser.newContext();
    const page = await context.newPage();

    // Navigate and wait for full network idle (CDN fonts, scripts)
    await page.goto(url, { waitUntil: "networkidle", timeout: 15000 });

    // Wait for Paged.js to complete pagination
    // Strategy 1: window.PagedPolyfill.done promise (Paged.js >= 0.4)
    // Strategy 2: presence of .pagedjs-page elements + 2s settle
    try {
      await page.waitForFunction(
        `window.PagedPolyfill && typeof window.PagedPolyfill.then === 'function'
          ? window.PagedPolyfill.then(() => true)
          : document.querySelectorAll('.pagedjs-page').length > 0`,
        { timeout: args.timeout }
      );
    } catch {
      // Fallback: wait for at least one pagedjs-page div
      await page.waitForFunction(
        `document.querySelectorAll('.pagedjs-page').length > 0`,
        { timeout: args.timeout }
      );
    }

    // Extra settle — Paged.js can still be adjusting layout
    await page.waitForTimeout(2000);

    // Generate PDF
    const format = args.pageSize.toLowerCase() === "a4" ? "A4" : "Letter";
    await page.pdf({
      path: args.output,
      format,
      printBackground: true,
      margin: {
        top: args.marginTop,
        right: args.marginSide,
        bottom: args.marginBottom,
        left: args.marginSide,
      },
    });

    console.log(args.output);
  } finally {
    if (browser) await browser.close();
    server.close();
  }
}

main().catch((err) => {
  console.error("export_pdf.js error:", err.message || err);
  process.exit(1);
});
