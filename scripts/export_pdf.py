#!/usr/bin/env python3
"""
export_pdf.py — Paged.js HTML → PDF export

Primary path: Python playwright (sync API) with Chromium.
Fallback path: Node.js export_pdf.js (shells out via subprocess).

Usage:
    python export_pdf.py \\
        --html /path/to/report.html \\
        --output /path/to/report.pdf \\
        [--page-size Letter|A4] \\
        [--margin-top 18mm] \\
        [--margin-side 16mm] \\
        [--margin-bottom 20mm] \\
        [--timeout 30000] \\
        [--force-node]

Exits with code 0 on success (prints output path to stdout).
Exits with code 1 on error (prints to stderr).
Exits with code 2 if neither Python nor Node playwright is available.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path


# ---------------------------------------------------------------------------
# Local HTTP server (avoids file:// CDN restrictions)
# ---------------------------------------------------------------------------

class _SilentHandler(SimpleHTTPRequestHandler):
    """SimpleHTTPRequestHandler that suppresses access logs."""

    def log_message(self, format, *args):  # noqa: A002
        pass

    def log_error(self, format, *args):  # noqa: A002
        pass


def _serve_directory(directory: str) -> tuple[HTTPServer, int]:
    """Start a silent HTTP server rooted at *directory* on a random port."""
    os.chdir(directory)
    server = HTTPServer(("127.0.0.1", 0), _SilentHandler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, port


# ---------------------------------------------------------------------------
# Python playwright path
# ---------------------------------------------------------------------------

def _export_via_python(
    html_path: Path,
    output_path: Path,
    page_size: str,
    margin_top: str,
    margin_side: str,
    margin_bottom: str,
    timeout_ms: int,
) -> None:
    """Export using playwright's Python sync API."""
    from playwright.sync_api import sync_playwright  # type: ignore[import]

    html_dir = str(html_path.parent.resolve())
    html_name = html_path.name

    server, port = _serve_directory(html_dir)
    url = f"http://127.0.0.1:{port}/{html_name}"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            try:
                context = browser.new_context()
                page = context.new_page()

                # Navigate — wait for network idle so CDN fonts/scripts load
                page.goto(url, wait_until="networkidle", timeout=15_000)

                # Wait for Paged.js to finish pagination
                _wait_for_pagedjs(page, timeout_ms)

                # Extra settle — Paged.js may still be adjusting layout
                page.wait_for_timeout(2_000)

                fmt = "A4" if page_size.lower() == "a4" else "Letter"
                output_path.parent.mkdir(parents=True, exist_ok=True)

                page.pdf(
                    path=str(output_path),
                    format=fmt,
                    print_background=True,
                    margin={
                        "top": margin_top,
                        "right": margin_side,
                        "bottom": margin_bottom,
                        "left": margin_side,
                    },
                )
            finally:
                browser.close()
    finally:
        server.shutdown()


def _wait_for_pagedjs(page, timeout_ms: int) -> None:
    """Wait for Paged.js to complete pagination using two strategies."""
    # Strategy 1: window.PagedPolyfill promise (Paged.js >= 0.4)
    try:
        page.wait_for_function(
            """
            () => {
                if (window.PagedPolyfill && typeof window.PagedPolyfill.then === 'function') {
                    return window.PagedPolyfill.then(() => true);
                }
                return document.querySelectorAll('.pagedjs-page').length > 0;
            }
            """,
            timeout=timeout_ms,
        )
        return
    except Exception:
        pass

    # Strategy 2: wait for at least one .pagedjs-page element
    page.wait_for_function(
        "() => document.querySelectorAll('.pagedjs-page').length > 0",
        timeout=timeout_ms,
    )


# ---------------------------------------------------------------------------
# Node.js fallback path
# ---------------------------------------------------------------------------

def _find_node() -> str | None:
    """Return the path to node, or None if not found."""
    for candidate in ("node", "nodejs"):
        try:
            result = subprocess.run(
                [candidate, "--version"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            if result.returncode == 0:
                return candidate
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None


def _export_via_node(
    html_path: Path,
    output_path: Path,
    page_size: str,
    margin_top: str,
    margin_side: str,
    margin_bottom: str,
    timeout_ms: int,
) -> None:
    """Export by shelling out to export_pdf.js."""
    node = _find_node()
    if not node:
        print(
            "Node.js is not available. Install it to use the Node.js fallback:\n"
            "  https://nodejs.org/",
            file=sys.stderr,
        )
        sys.exit(2)

    js_script = Path(__file__).parent / "export_pdf.js"
    if not js_script.exists():
        print(
            f"Node.js fallback script not found: {js_script}\n"
            "Ensure export_pdf.js is in the same directory as export_pdf.py.",
            file=sys.stderr,
        )
        sys.exit(2)

    # Check that the playwright npm package is available
    check = subprocess.run(
        [node, "-e", "require('playwright')"],
        capture_output=True,
        text=True,
        cwd=str(js_script.parent),
    )
    if check.returncode != 0:
        print(
            "playwright npm package is not installed. Install it:\n"
            "  cd scripts && npm install playwright\n"
            "  npx playwright install chromium",
            file=sys.stderr,
        )
        sys.exit(2)

    cmd = [
        node, str(js_script),
        "--html", str(html_path.resolve()),
        "--output", str(output_path.resolve()),
        "--page-size", page_size,
        "--margin-top", margin_top,
        "--margin-side", margin_side,
        "--margin-bottom", margin_bottom,
        "--timeout", str(timeout_ms),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"export_pdf.js error:\n{result.stderr}", file=sys.stderr)
        sys.exit(result.returncode)

    # Node script prints the output path to stdout
    print(result.stdout.strip())


# ---------------------------------------------------------------------------
# Python playwright availability check
# ---------------------------------------------------------------------------

def _python_playwright_available() -> bool:
    """Return True if the playwright Python package and chromium are ready."""
    try:
        import playwright  # noqa: F401
    except ImportError:
        return False

    # Check that chromium is installed (playwright install chromium)
    try:
        from playwright.sync_api import sync_playwright  # type: ignore[import]
        with sync_playwright() as p:
            exec_path = p.chromium.executable_path
            return os.path.exists(exec_path)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Export a Paged.js HTML report to PDF.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--html", required=True, help="Path to the source HTML file")
    p.add_argument("--output", required=True, help="Destination PDF path")
    p.add_argument(
        "--page-size",
        default="Letter",
        choices=["Letter", "A4", "letter", "a4"],
        help="PDF page size (default: Letter)",
    )
    p.add_argument("--margin-top", default="18mm", help="Top margin (default: 18mm)")
    p.add_argument("--margin-side", default="16mm", help="Side margins (default: 16mm)")
    p.add_argument("--margin-bottom", default="20mm", help="Bottom margin (default: 20mm)")
    p.add_argument(
        "--timeout",
        type=int,
        default=30_000,
        help="Paged.js wait timeout in milliseconds (default: 30000)",
    )
    p.add_argument(
        "--force-node",
        action="store_true",
        help="Skip Python playwright and use the Node.js fallback directly",
    )
    return p


def main() -> None:
    args = _build_parser().parse_args()

    html_path = Path(args.html).resolve()
    output_path = Path(args.output).resolve()

    if not html_path.exists():
        print(f"HTML file not found: {html_path}", file=sys.stderr)
        sys.exit(1)

    # Normalise page size to title case for downstream
    page_size = args.page_size.capitalize()  # "letter" → "Letter", "a4" → "A4" via special case
    if page_size.lower() == "a4":
        page_size = "A4"

    common = dict(
        html_path=html_path,
        output_path=output_path,
        page_size=page_size,
        margin_top=args.margin_top,
        margin_side=args.margin_side,
        margin_bottom=args.margin_bottom,
        timeout_ms=args.timeout,
    )

    if not args.force_node and _python_playwright_available():
        try:
            _export_via_python(**common)
            print(str(output_path))
            return
        except Exception as exc:
            print(
                f"Python playwright export failed ({exc}); falling back to Node.js.",
                file=sys.stderr,
            )
            # Fall through to Node fallback

    # Node.js fallback
    _export_via_node(**common)


if __name__ == "__main__":
    main()
