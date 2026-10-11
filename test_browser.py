#!/usr/bin/env python3
"""
test_browser.py — Headless browser test using Playwright.

Opens attendance_app.html, pastes the 8-line example, triggers
"Download all", intercepts the three blob downloads, then verifies
each workbook with openpyxl.

Run: python3 test_browser.py
"""

import asyncio
import base64
import hashlib
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path

import openpyxl
from playwright.async_api import async_playwright

ROOT = Path(__file__).parent
APP_HTML = ROOT / "attendance_app.html"
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT / "_test_output"
OUT_DIR.mkdir(exist_ok=True)

EIGHT_LINES = """Samra Syed - R24EH133 - 5th sem
Shreyans Mehta - R24EH148 - 5th sem AIDS C
ARJUN R24EA016- 5 TH SEM AI ML A
Karempudi Devendran Chowdary — R25EA105 — 3rd sem — AIML - B
Jeevan-R24EH051 5th sem -AI:DS A
Mahika S-R24EJ061-5th Sem CSIT B
Gagan
NoNameXYZ - R99ZZ999"""

EVENT_NAME = "Test Event"
EVENT_DATE = "2026-10-15"

errors = []

def ok(msg): print(f"  ✓ PASS: {msg}")
def fail(msg): errors.append(msg); print(f"  ✗ FAIL: {msg}")
def check(cond, p, f): (ok if cond else fail)(p if cond else f)


async def run():
    downloads = {}  # filename -> bytes

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        ctx = await browser.new_context(accept_downloads=True)
        page = await ctx.new_page()

        print(f"\nOpening {APP_HTML.name} …")
        await page.goto(f"file://{APP_HTML.absolute()}")
        await page.wait_for_load_state("networkidle")

        # Fill event name
        await page.fill("#event-name", EVENT_NAME)
        # Fill date
        await page.fill("#event-date", EVENT_DATE)
        # Pick Post Lunch preset
        await page.select_option("#session-preset", "pl")
        await page.wait_for_timeout(200)

        print("Pasting 8-line example …")
        await page.fill("#student-input", EIGHT_LINES)
        await page.click("#parse-btn")
        await page.wait_for_timeout(500)

        # Check warnings rendered for Gagan (ambiguous) and R99ZZ999 (not found)
        warn_text = await page.inner_text("#warnings-zone")
        check("Gagan" in warn_text, "Warnings zone mentions Gagan", "Gagan not mentioned in warnings")
        check("NoNameXYZ" in warn_text or "R99ZZ999" in warn_text,
              "Warnings zone mentions NoNameXYZ / R99ZZ999",
              "NoNameXYZ / R99ZZ999 not in warnings")

        # Check Group 2 card appears
        g2_text = await page.inner_text("#groups-zone")
        check("Group 2" in g2_text, "Group 2 card rendered", "Group 2 card missing")

        # Download each group
        for g in [1, 2, 3]:
            dl_btn = page.locator(f"#dl-g{g}")
            if not await dl_btn.count():
                print(f"  (Group {g} download button not found — skipping)")
                continue

            print(f"  Downloading Group {g} …")
            async with page.expect_download() as dl_info:
                await dl_btn.click()
            dl = await dl_info.value
            fname = dl.suggested_filename
            dest = OUT_DIR / fname
            await dl.save_as(dest)
            downloads[g] = (fname, dest)
            print(f"    → {fname}")

        await browser.close()

    # ── Verify downloads with openpyxl ────────────────────────────────────────
    print("\n=== Verifying downloaded workbooks ===")

    for g, (fname, fpath) in downloads.items():
        print(f"\n  Group {g}: {fname}")
        wb = openpyxl.load_workbook(fpath)

        # Link Sheet must be present
        check("Link Sheet" in wb.sheetnames,
              f"G{g} has Link Sheet", f"G{g} missing Link Sheet")

        # Verify Link Sheet content matches original
        orig_path = DATA_DIR / f"Attendance_21-9_Group{g}_Badging.xlsx"
        with zipfile.ZipFile(orig_path) as zorig, zipfile.ZipFile(fpath) as znew:
            def link_bytes(z):
                wb_xml = z.read('xl/workbook.xml').decode()
                rels_xml = z.read('xl/_rels/workbook.xml.rels').decode()
                m = re.search(r'name="Link Sheet"[^>]*r:id="([^"]*)"', wb_xml)
                if not m: return b""
                rid = m.group(1)
                rm = re.search(rf'Id="{rid}"[^>]*Target="([^"]*)"', rels_xml)
                if not rm: return b""
                target = rm.group(1)
                path = target if target.startswith('xl/') else 'xl/' + target
                return z.read(path)

            def norm(b):
                s = b.decode('utf-8', errors='replace')
                s = re.sub(r'<\?xml[^?]*\?>\s*', '', s)
                return re.sub(r'\s+', ' ', s).strip()

            match = norm(link_bytes(zorig)) == norm(link_bytes(znew))
            check(match, f"G{g} Link Sheet content identical to original",
                  f"G{g} Link Sheet content DIFFERS from original!")

        # For each data sheet: check row structure
        for shname in wb.sheetnames:
            if "Link" in shname:
                continue
            ws = wb[shname]
            # Row 1: title (merged A1:L1)
            check(ws.cell(1, 1).value is not None,
                  f"G{g}/{shname}: row 1 has title",
                  f"G{g}/{shname}: row 1 title missing")
            # Row 2: headers
            check(ws.cell(2, 1).value == "SL.No",
                  f"G{g}/{shname}: row 2 col A = 'SL.No'",
                  f"G{g}/{shname}: row 2 col A = '{ws.cell(2,1).value}'")
            check(ws.cell(2, 2).value == "SRN",
                  f"G{g}/{shname}: row 2 col B = 'SRN'",
                  f"G{g}/{shname}: row 2 col B = '{ws.cell(2,2).value}'")

            # Find last row with data
            max_r = ws.max_row
            footer_row = None
            for r in range(max_r, 2, -1):
                v = ws.cell(r, 1).value
                if v is not None:
                    footer_row = r
                    break

            if footer_row:
                footer_val = ws.cell(footer_row, 1).value
                check("Extracurricular" in str(footer_val),
                      f"G{g}/{shname}: footer row {footer_row} has signature",
                      f"G{g}/{shname}: footer row {footer_row} has '{footer_val}'")

                # Check merged cell on footer
                merged = [str(m) for m in ws.merged_cells.ranges]
                footer_merge = f"A{footer_row}:D{footer_row}"
                check(footer_merge in merged,
                      f"G{g}/{shname}: footer merge cell {footer_merge} present",
                      f"G{g}/{shname}: footer merge cell {footer_merge} missing. Merges: {merged}")

                # Check SL.No is numbered from 1 (only when data rows are present)
                if footer_row > 7:  # 3 (start) + 0 students + 4 (GAP) = 7 means no data
                    check(ws.cell(3, 1).value in [1, 1.0],
                          f"G{g}/{shname}: first data row SL.No=1",
                          f"G{g}/{shname}: first data row SL.No={ws.cell(3,1).value}")

            # Group 2 should have cloned 3rd sem sheet if Karempudi is there
            if g == 2 and "3rd" in shname:
                # Should have at least Karempudi (R25EA105)
                usns = [ws.cell(r, 2).value for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value not in [None, "Extracurricular Vertical Head"]]
                check("R25EA105" in usns,
                      f"G2/3rd Sem: Karempudi (R25EA105) present",
                      f"G2/3rd Sem: Karempudi not found. USNs: {usns}")

    # ── Verify XML declarations ────────────────────────────────────────────────
    print("\n=== Verifying XML declarations ===")
    for g, (fname, fpath) in downloads.items():
        with zipfile.ZipFile(fpath) as z:
            for name in z.namelist():
                if not name.endswith('.xml'): continue
                data = z.read(name)
                try:
                    text = data.decode('utf-8', errors='replace')
                except:
                    continue
                decl_count = text.count('<?xml')
                check(decl_count <= 1,
                      f"G{g}/{name}: exactly 1 XML decl (found {decl_count})",
                      f"G{g}/{name}: found {decl_count} XML declarations!")

    # ── Summary ────────────────────────────────────────────────────────────────
    print("\n" + "="*60)
    if errors:
        print(f"FAILED — {len(errors)} check(s) failed:")
        for e in errors:
            print(f"  • {e}")
        sys.exit(1)
    else:
        print("ALL BROWSER TESTS PASSED")


if __name__ == "__main__":
    asyncio.run(run())
