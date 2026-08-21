#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ตรวจว่ากฎทุกข้อ **ไม่ฟ้องเทมเพลตทางการของหอสมุด**

รัน: python3 scripts/test_official_templates.py

ทำไมต้องมีไฟล์นี้แยกจาก test_false_positives.py
--------------------------------------------------
`test_false_positives.py` ทดสอบกฎทีละข้อด้วยข้อความสังเคราะห์ ซึ่งจับ regex ที่จับ
เกินได้ แต่**จับไม่ได้**เมื่อความผิดพลาดอยู่ที่การอ่านค่าจาก OOXML (เช่นอ่านช่อง
ตัวหนาผิดช่อง หรือเอา character style ไปเทียบกับเกณฑ์ของ paragraph style)

ไฟล์นี้จึงรันตัวตรวจ **ทั้งชุด** กับเทมเพลตทางการจริง แล้วบังคับว่าต้องได้ 0 finding
ยกเว้นรายการที่ขึ้นทะเบียนไว้ใน KNOWN_TEMPLATE_DEFECTS ว่าเป็นข้อบกพร่องของ
ตัวเทมเพลตเอง (ไม่ใช่ของกฎ)

ตอนเพิ่มไฟล์นี้ครั้งแรกจับ false positive ได้ 5 ชนิดที่หลุดมานาน:
  1. TU_Chapter ถูกฟ้องว่า "ไม่หนา" — เพราะอ่าน w:b (ละติน = ปิด) แทน w:bCs (ไทย = เปิด)
  2. สไตล์ "… Char" 5 ตัวถูกเอาไปเทียบกับเกณฑ์ของ paragraph style
  3. "ภาคผนวก" 20 pt บนหน้าคั่นถูกฟ้องว่าเพี้ยนจาก TU_Chapter (18 pt)
  4. "THESIS OR DISSERTATION" (placeholder ให้เลือก) ถูกอ่านว่าเล่มนี้เป็นดุษฎีนิพนธ์
     แล้วฟ้องต่อว่า "ยังมีคำว่าวิทยานิพนธ์ปนอยู่"
  5. รายการอ้างอิงแบบเรียงเลขถูกตรวจด้วยกฎการเรียงตัวอักษร
"""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "fixtures"
SCRIPTS = ROOT / "scripts"

# ข้อบกพร่องที่อยู่ใน **ตัวเทมเพลตเอง** ยืนยันแล้วด้วยการอ่าน OOXML โดยตรง
# กฎที่จับสิ่งเหล่านี้ทำงานถูกต้อง จึงไม่นับเป็น false positive
#
# บทที่ 7–8 ของเทมเพลตไทยมีย่อหน้าที่ใช้สไตล์ Normal + จัดรูปแบบด้วยมือ แทนที่จะใช้
# TU_Sub-heading — เป็นความไม่เรียบร้อยของไฟล์ต้นแบบ ไม่ใช่ของกฎ
KNOWN_TEMPLATE_DEFECTS = [
    (r"TULIBS_Thesis-template-Thai",
     r"หัวข้อย่อยระดับ [12] \(X\.X\.X(\.X)?\) “[78]\.[\d.]+ หัวข้อย่อยระดับที่ \d”"
     r" ใช้สไตล์ Normal"),
]

# ระดับที่ยอมให้ผ่าน — info เป็นข้อมูลประกอบ ไม่ใช่การฟ้อง
IGNORED_SEVERITIES = {"info", "ok"}

FAIL = []


def is_known_defect(fixture_name, issue):
    for file_pat, issue_pat in KNOWN_TEMPLATE_DEFECTS:
        if re.search(file_pat, fixture_name) and re.search(issue_pat, issue):
            return True
    return False


def run_checks(docx_path, tmp_prefix):
    auto = f"/tmp/{tmp_prefix}_auto.json"
    deep = f"/tmp/{tmp_prefix}_deep.json"
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "check_all.py"), str(docx_path),
         "--auto-json", auto, "--deep-json", deep],
        capture_output=True, text=True, timeout=300)
    if proc.returncode != 0:
        FAIL.append(f"{docx_path.name}: check_all.py ล้มเหลว\n{proc.stderr[-1500:]}")
        return []
    findings = []
    for path in (auto, deep):
        try:
            with open(path, encoding="utf-8") as fh:
                findings += json.load(fh).get("findings", [])
        except (OSError, json.JSONDecodeError) as exc:
            FAIL.append(f"{docx_path.name}: อ่าน {path} ไม่ได้ ({exc})")
    return findings


def main():
    fixtures = sorted(FIXTURES.glob("*.docx")) if FIXTURES.is_dir() else []
    if not fixtures:
        print(f"ข้าม: ไม่พบไฟล์เทมเพลตใน {FIXTURES}")
        print("      วางไฟล์ TULIBS_Thesis-template-*.docx ไว้ในโฟลเดอร์นี้แล้วรันใหม่")
        return 0

    total_checked = 0
    for fixture in fixtures:
        findings = run_checks(fixture, fixture.stem[:20].replace(" ", "_"))
        total_checked += 1
        unexpected = []
        for f in findings:
            if f.get("severity") in IGNORED_SEVERITIES:
                continue
            issue = f.get("issue", "")
            if is_known_defect(fixture.name, issue):
                continue
            unexpected.append(f"[{f.get('severity')}] {f.get('category')}: {issue[:130]}")
        if unexpected:
            FAIL.append(
                f"เทมเพลตทางการ {fixture.name!r} ถูกฟ้อง {len(unexpected)} จุด — "
                f"เทมเพลตของหอสมุดคือเกณฑ์ ถ้ากฎฟ้องเทมเพลต แปลว่ากฎผิด:\n    "
                + "\n    ".join(unexpected[:10]))
        else:
            print(f"  ok  {fixture.name} — 0 finding (ไม่นับ info/ข้อบกพร่องที่ขึ้นทะเบียน)")

    if FAIL:
        print(f"\nFAILED ({len(FAIL)}):")
        for f in FAIL:
            print("  -", f)
        return 1
    print(f"\nผ่านทั้งหมด — ตรวจเทมเพลตทางการ {total_checked} ไฟล์ ไม่มี false positive")
    return 0


if __name__ == "__main__":
    sys.exit(main())
