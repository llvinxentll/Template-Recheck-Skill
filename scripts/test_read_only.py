#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ด่าน 3 — พิสูจน์ว่าสกิลไม่แตะไฟล์ต้นฉบับ

กฎ "🔒 อ่านอย่างเดียว" ใน SKILL.md จะเชื่อถือได้ก็ต่อเมื่อมีอะไรบังคับ
ด่านนี้รันขั้นตอนจริงทั้งหมดกับสำเนาเทมเพลตทางการ แล้วยืนยันสองอย่าง:

  1. ไฟล์ต้นฉบับมี sha256 เท่าเดิม  → ไม่มีใครเขียนทับ
  2. โฟลเดอร์ต้นฉบับมีไฟล์เท่าเดิม → ไม่มีไฟล์งอก
     (ตัวที่งอกง่ายสุดคือ `.~lock.<ชื่อ>#` ของ LibreOffice ตอนเรนเดอร์
      ถ้าเผลอชี้ soffice ไปที่ต้นฉบับแทนสำเนา)

ด่านนี้จับกรณีที่โค้ดยัง "ทำงานถูก" แต่ทิ้งร่องรอยไว้ในโฟลเดอร์ของนักศึกษา
ซึ่งด่าน 1 และ 2 มองไม่เห็นเลย
"""
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
FIXTURES = ROOT / "fixtures"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run(args, cwd=None):
    return subprocess.run([sys.executable] + args, cwd=cwd,
                          stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                          text=True, timeout=1800)


def check_one(fixture, workroot):
    src_dir = Path(tempfile.mkdtemp(prefix="tulibs-ro-src-", dir=workroot))
    out_dir = Path(tempfile.mkdtemp(prefix="tulibs-ro-out-", dir=workroot))
    src = src_dir / fixture.name
    shutil.copyfile(fixture, src)

    before_hash = sha256(src)
    before_ls = sorted(os.listdir(src_dir))

    auto = out_dir / "auto.json"
    deep = out_dir / "deep.json"
    report = out_dir / "report.docx"

    steps = [
        ("check_all", [str(SCRIPTS / "check_all.py"), str(src),
                       "--auto-json", str(auto), "--deep-json", str(deep)]),
        ("page_truth", [str(SCRIPTS / "page_truth.py"), str(src),
                        "-o", str(out_dir / "pagemap.json")]),
        ("make_report", [str(SCRIPTS / "make_report.py"), str(auto), str(deep),
                         "--docx", str(src), "-o", str(report),
                         "--title", "read-only test"]),
    ]

    problems = []
    for label, args in steps:
        proc = run(args)
        # page_truth คืน 3 เมื่อเรนเดอร์ไม่ได้ = พฤติกรรมที่ออกแบบไว้ ไม่ใช่ error
        ok = proc.returncode == 0 or (label == "page_truth" and proc.returncode == 3)
        if not ok:
            problems.append("%s ล้มเหลว (exit %s): %s"
                            % (label, proc.returncode, proc.stdout.strip()[-300:]))

    after_hash = sha256(src)
    after_ls = sorted(os.listdir(src_dir))

    if before_hash != after_hash:
        problems.append("ไฟล์ต้นฉบับถูกเขียนทับ (sha256 เปลี่ยน)")
    if before_ls != after_ls:
        extra = sorted(set(after_ls) - set(before_ls))
        missing = sorted(set(before_ls) - set(after_ls))
        problems.append("โฟลเดอร์ต้นฉบับเปลี่ยน — งอก %s หาย %s" % (extra, missing))

    return problems


def main():
    fixtures = sorted(FIXTURES.glob("*.docx")) if FIXTURES.is_dir() else []
    if not fixtures:
        print("⏭  ข้าม: ไม่พบไฟล์ทดสอบใน fixtures/")
        return 0

    failed = 0
    with tempfile.TemporaryDirectory(prefix="tulibs-readonly-") as workroot:
        for fx in fixtures:
            problems = check_one(fx, workroot)
            if problems:
                failed += 1
                print("  ไม่ผ่าน  %s" % fx.name)
                for p in problems:
                    print("      - %s" % p)
            else:
                print("  ok  %s — ต้นฉบับไม่ถูกแตะ ไม่มีไฟล์งอกในโฟลเดอร์" % fx.name)

    print()
    if failed:
        print("❌ %d ไฟล์ละเมิดกฎอ่านอย่างเดียว" % failed)
        print("   ทุกการเรนเดอร์ต้องทำจากสำเนาใน temp dir และห้ามเขียนข้างต้นฉบับ")
        return 1
    print("ผ่านทั้งหมด — สกิลไม่แตะไฟล์ต้นฉบับและไม่สร้างไฟล์ในโฟลเดอร์ต้นฉบับ")
    return 0


if __name__ == "__main__":
    sys.exit(main())
