#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""รันชุดทดสอบทั้งหมดของสกิลในคำสั่งเดียว — รันหลังแก้กฎใด ๆ ก่อนใช้จริง

    python3 scripts/run_tests.py

ประกอบด้วย 3 ด่านที่จับคนละอย่าง:

  1. test_false_positives.py  — ยิงกฎทีละข้อด้วยข้อความสังเคราะห์
                                จับ regex ที่จับเกิน (URL, ชื่อเมือง, ปีในวงเล็บ)
  2. test_official_templates.py — รันตัวตรวจทั้งชุดกับเทมเพลตทางการจริง
                                จับความผิดพลาดตอนอ่านค่าจาก OOXML ซึ่งด่านแรกมองไม่เห็น
  3. test_read_only.py        — รันขั้นตอนจริงครบวงจรแล้วยืนยันว่าไฟล์ต้นฉบับ
                                sha256 เท่าเดิมและไม่มีไฟล์งอกในโฟลเดอร์

ด่านที่ 2 สำคัญกว่าที่คิด: ตอนเพิ่มเข้ามาครั้งแรกจับ false positive ได้ 5 ชนิด
ที่ค้างอยู่ในสกิลมานาน ทั้งที่ด่านแรกผ่านหมดทุกครั้ง
"""
import subprocess
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
SUITES = [
    ("กฎรายข้อ (false positives)", "test_false_positives.py"),
    ("เทมเพลตทางการ (ต้องได้ 0)", "test_official_templates.py"),
    ("อ่านอย่างเดียว (ห้ามแตะต้นฉบับ)", "test_read_only.py"),
]


def main():
    failed = []
    for label, script in SUITES:
        path = SCRIPTS / script
        if not path.exists():
            print(f"⏭  ข้าม {label}: ไม่พบ {script}")
            continue
        print(f"\n{'='*68}\n▶  {label}\n{'='*68}")
        proc = subprocess.run([sys.executable, str(path)], text=True)
        if proc.returncode != 0:
            failed.append(label)

    print(f"\n{'='*68}")
    if failed:
        print(f"❌ ไม่ผ่าน {len(failed)} ชุด: " + " · ".join(failed))
        print("   อย่าเพิ่งใช้กฎที่แก้ไป — แก้ให้ผ่านก่อน")
        return 1
    print("✅ ผ่านทุกชุด")
    return 0


if __name__ == "__main__":
    sys.exit(main())
