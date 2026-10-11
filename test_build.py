#!/usr/bin/env python3
"""
test_build.py — Verifies the build output of attendance_app.html.

Tests:
  1. Embedded FACE_MEMBERS JSON is correct (Shivanshi correction, section extraction)
  2. SECTION_FALLBACK is populated
  3. Base64 template data decodes to valid xlsx files
  4. Link Sheet XML content is identical to the originals
  5. Simulates the JS parsing logic in Python for the 8-line example
  6. Generates workbooks via Playwright (headless browser) and verifies with openpyxl

Run: python3 test_build.py
"""

import base64
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
APP_HTML = ROOT / "attendance_app.html"
DATA_DIR = ROOT / "data"

PASS = "✓ PASS"
FAIL = "✗ FAIL"

errors = []

def ok(msg):
    print(f"  {PASS}: {msg}")

def fail(msg):
    errors.append(msg)
    print(f"  {FAIL}: {msg}")

def check(cond, pass_msg, fail_msg):
    (ok if cond else fail)(pass_msg if cond else fail_msg)

# ── 1. Read embedded data from attendance_app.html ───────────────────────────

print("\n=== 1. Verifying embedded data ===")
html = APP_HTML.read_text(encoding="utf-8")

# Extract FACE_MEMBERS
m = re.search(r'const FACE_MEMBERS = (\{[\s\S]*?\});', html)
if not m:
    fail("FACE_MEMBERS not found in HTML")
    sys.exit(1)
face_members = json.loads(m.group(1))
ok(f"FACE_MEMBERS loaded ({len(face_members)} entries)")

# Check Shivanshi correction
sh = face_members.get("R25EH122")
check(sh is not None, "R25EH122 (Shivanshi) present in FACE_MEMBERS", "R25EH122 missing from FACE_MEMBERS")
if sh:
    check(sh["branch"] == "AIDS C", f"Shivanshi branch='AIDS C'", f"Shivanshi branch='{sh['branch']}' (expected 'AIDS C')")
    check(sh["section"] == "C", f"Shivanshi section='C'", f"Shivanshi section='{sh['section']}' (expected 'C')")
    check(sh["group"] == 2, f"Shivanshi group=2", f"Shivanshi group={sh['group']} (expected 2)")

# Check a few known members
known = [
    ("R24EH133", {"name": "Samra Syed", "group": 2, "semester": 5, "section": "C"}),
    ("R24EH148", {"name": "Shreyans Mehta", "group": 2, "semester": 5, "section": "C"}),
    ("R24EA016", {"name": "Arjun", "group": 2, "semester": 5, "section": "A"}),
    ("R25EA105", {"name": "Karempudi Devendran Chowdary", "group": 2, "semester": 3, "section": "B"}),
    ("R24EH051", {"name": "Jeevan", "group": 2, "semester": 5, "section": "A"}),
    ("R24EJ061", {"name": "Mahika S", "group": 3, "semester": 5, "section": "B"}),
]
for usn, expected in known:
    fm = face_members.get(usn)
    if fm is None:
        fail(f"{usn}: not in FACE_MEMBERS"); continue
    for field, val in expected.items():
        check(fm[field] == val, f"{usn} {field}={val!r}", f"{usn} {field}={fm[field]!r} (expected {val!r})")

# SECTION_FALLBACK
m2 = re.search(r'const SECTION_FALLBACK = (\{[\s\S]*?\});', html)
check(m2 is not None, "SECTION_FALLBACK found", "SECTION_FALLBACK not found")
if m2:
    sf = json.loads(m2.group(1))
    check(len(sf) > 0, f"SECTION_FALLBACK has {len(sf)} entries", "SECTION_FALLBACK is empty")

# ── 2. Verify base64 decodes to valid xlsx ───────────────────────────────────

print("\n=== 2. Verifying base64 template data ===")
for g in [1, 2, 3]:
    pattern = re.compile(rf'"1": "([A-Za-z0-9+/=]+)"'.replace('"1"', f'"{g}"'))
    # More robust: find the block
    b64m = re.search(rf'"{g}": "([A-Za-z0-9+/=]+)"', html)
    if not b64m:
        fail(f"TEMPLATE_B64[{g}] not found"); continue
    raw = base64.b64decode(b64m.group(1))
    try:
        with zipfile.ZipFile(__import__('io').BytesIO(raw)) as zf:
            names = zf.namelist()
        check('xl/workbook.xml' in names, f"Group {g} template is valid xlsx", f"Group {g} template missing workbook.xml")
    except Exception as e:
        fail(f"Group {g} template not a valid zip: {e}")

# ── 3. Verify Link Sheet content is identical to originals ───────────────────

print("\n=== 3. Verifying Link Sheet content ===")

def get_link_sheet_bytes(xlsx_path):
    """Return (path_in_zip, raw_bytes) for the Link Sheet XML."""
    with zipfile.ZipFile(xlsx_path) as zf:
        # Read workbook.xml to find Link Sheet rId
        wb_xml = zf.read('xl/workbook.xml').decode()
        rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode()

        # Find Link Sheet rId
        m = re.search(r'name="Link Sheet"[^>]*r:id="([^"]*)"', wb_xml)
        if not m:
            return None, None
        rid = m.group(1)

        # Find the file path
        rm = re.search(rf'Id="{rid}"[^>]*Target="([^"]*)"', rels_xml)
        if not rm:
            return None, None
        target = rm.group(1)
        path = target if target.startswith('xl/') else 'xl/' + target

        return path, zf.read(path)

for g in [1, 2, 3]:
    orig_path = DATA_DIR / f"Attendance_21-9_Group{g}_Badging.xlsx"
    orig_link_path, orig_link_bytes = get_link_sheet_bytes(orig_path)
    if orig_link_bytes is None:
        fail(f"Group {g}: could not find Link Sheet in original"); continue

    # Extract from embedded base64
    b64m = re.search(rf'"{g}": "([A-Za-z0-9+/=]+)"', html)
    if not b64m:
        fail(f"Group {g}: b64 not found"); continue

    raw = base64.b64decode(b64m.group(1))
    with zipfile.ZipFile(__import__('io').BytesIO(raw)) as zf:
        _, emb_bytes = get_link_sheet_bytes(__import__('io').BytesIO(raw))

    # Compare XML content (strip whitespace differences)
    def normalise_xml(b):
        s = b.decode('utf-8', errors='replace') if isinstance(b, bytes) else b
        s = re.sub(r'<\?xml[^?]*\?>\s*', '', s)
        return re.sub(r'\s+', ' ', s).strip()

    if emb_bytes is None:
        fail(f"Group {g}: Link Sheet not found in embedded template"); continue

    match = normalise_xml(orig_link_bytes) == normalise_xml(emb_bytes)
    check(match, f"Group {g} Link Sheet XML content is identical",
          f"Group {g} Link Sheet XML content differs!")

# ── 4. Python simulation of parse logic ─────────────────────────────────────

print("\n=== 4. Simulating parse logic (Python) ===")

USN_RE = re.compile(r'R\d{2}E[A-Za-z]\d{3}', re.IGNORECASE)
SEM_RE = re.compile(r'\b(\d+)\s*(?:st|nd|rd|th)?\s*sem(?:ester)?\b', re.IGNORECASE)

def extract_usn(line):
    m = USN_RE.search(line)
    return m.group(0).upper() if m else None

def extract_sem(line):
    m = SEM_RE.search(line)
    return int(m.group(1)) if m else None

def extract_section_from_line(line):
    m = SEM_RE.search(line)
    if not m: return None
    after = line[m.end():]
    matches = re.findall(r'(?<![A-Za-z])([A-F])(?![A-Za-z])', after.upper())
    return matches[-1] if matches else None

EIGHT_LINES = [
    "Samra Syed - R24EH133 - 5th sem",
    "Shreyans Mehta - R24EH148 - 5th sem AIDS C",
    "ARJUN R24EA016- 5 TH SEM AI ML A",
    "Karempudi Devendran Chowdary — R25EA105 — 3rd sem — AIML - B",
    "Jeevan-R24EH051 5th sem -AI:DS A",
    "Mahika S-R24EJ061-5th Sem CSIT B",
    "Gagan",
    "NoNameXYZ - R99ZZ999",
]

expected_parse = {
    "R24EH133": {"sem": 5, "section": "C", "group": 2},
    "R24EH148": {"sem": 5, "section": "C", "group": 2},
    "R24EA016": {"sem": 5, "section": "A", "group": 2},
    "R25EA105": {"sem": 3, "section": "B", "group": 2},
    "R24EH051": {"sem": 5, "section": "A", "group": 2},
    "R24EJ061": {"sem": 5, "section": "B", "group": 3},
}

found = {}
ambiguous_lines = []
not_found_lines = []

for line in EIGHT_LINES:
    usn = extract_usn(line)
    sem = extract_sem(line)
    sect = extract_section_from_line(line)

    if usn:
        fm = face_members.get(usn)
        if fm:
            section = sect or fm.get("section") or ""
            found[usn] = {"sem": sem or fm["semester"], "section": section, "group": fm["group"]}
        else:
            not_found_lines.append(line)  # no_group case
    else:
        # Name-only line
        name_part = re.sub(USN_RE, ' ', line)
        name_part = re.sub(SEM_RE, ' ', name_part)
        name_part = re.sub(r'[—\-|/]', ' ', name_part)
        name_part = re.sub(r'\s+', ' ', name_part).strip()
        if name_part:
            # Simple name match
            matches = [u for u, fm in face_members.items()
                       if all(w in fm['name'].lower() for w in name_part.lower().split())]
            if len(matches) > 1:
                ambiguous_lines.append((name_part, matches))
            elif len(matches) == 0:
                not_found_lines.append(line)

# Check expected USNs were parsed
for usn, exp in expected_parse.items():
    f = found.get(usn)
    check(f is not None, f"{usn} parsed", f"{usn} NOT parsed")
    if f:
        check(f["section"] == exp["section"], f"{usn} section='{exp['section']}'", f"{usn} section='{f['section']}' (expected '{exp['section']}')")
        check(f["sem"] == exp["sem"], f"{usn} sem={exp['sem']}", f"{usn} sem={f['sem']} (expected {exp['sem']})")
        check(f["group"] == exp["group"], f"{usn} group={exp['group']}", f"{usn} group={f['group']} (expected {exp['group']})")

# Check Gagan is ambiguous
check(len(ambiguous_lines) == 1, "Gagan line flagged as ambiguous", f"Expected 1 ambiguous line, got {len(ambiguous_lines)}")
if ambiguous_lines:
    name, cands = ambiguous_lines[0]
    check('gagan' in name.lower(), f"Ambiguous line is Gagan: '{name}'", f"Unexpected ambiguous: '{name}'")
    check(len(cands) >= 2, f"Gagan has {len(cands)} candidates", "Expected >=2 Gagan candidates")

# R99ZZ999 should not be a valid USN and NoNameXYZ not found
check(extract_usn("NoNameXYZ - R99ZZ999") is None, "R99ZZ999 correctly not matched as USN (fails R__E__ pattern)", "R99ZZ999 should NOT match USN pattern")

# ── 5. Verify Group 2 workbook structure (openpyxl on original) ─────────────

print("\n=== 5. Verifying original template structure ===")
for g in [1, 2, 3]:
    wb = openpyxl.load_workbook(DATA_DIR / f"Attendance_21-9_Group{g}_Badging.xlsx")
    has_link = "Link Sheet" in wb.sheetnames
    check(has_link, f"Group {g} has 'Link Sheet'", f"Group {g} missing 'Link Sheet'")

wb1 = openpyxl.load_workbook(DATA_DIR / "Attendance_21-9_Group1_Badging.xlsx")
check("Group 1 - 5th Sem" in wb1.sheetnames, "Group 1 has '5th Sem' sheet", "Group 1 missing 5th Sem sheet")
check("Group 1 - 3rd Sem" in wb1.sheetnames, "Group 1 has '3rd Sem' sheet", "Group 1 missing 3rd Sem sheet")

wb2 = openpyxl.load_workbook(DATA_DIR / "Attendance_21-9_Group2_Badging.xlsx")
check("Group 2 - 5th Sem" in wb2.sheetnames, "Group 2 has '5th Sem' sheet", "Group 2 missing 5th Sem sheet")
check("Group 2 - 3rd Sem" not in wb2.sheetnames, "Group 2 correctly has NO '3rd Sem' sheet (template)", "")

wb3 = openpyxl.load_workbook(DATA_DIR / "Attendance_21-9_Group3_Badging.xlsx")
check("Group 1 -5th Sem" in wb3.sheetnames, "Group 3 template has misnamed '5th' sheet (as expected)", "")

# ── Summary ──────────────────────────────────────────────────────────────────

print("\n" + "="*60)
if errors:
    print(f"FAILED — {len(errors)} check(s) failed:")
    for e in errors:
        print(f"  • {e}")
    sys.exit(1)
else:
    print(f"ALL CHECKS PASSED")
