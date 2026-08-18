#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""เสนอ "บรรทัดส่วนเกินบนหน้าปก" ที่ควรเพิ่มเข้ารายการตรวจ — สำหรับคนอนุมัติ ไม่ใช่กฎอัตโนมัติ

ทำไมต้องมีขั้นให้คนอนุมัติ
--------------------------
เราลองทำตัวตรวจแบบ "เทียบหน้าปกกับเทมเพลตแล้วฟ้องทุกบรรทัดที่ไม่ตรง" มาแล้วสองวิธี
และวัดกับวิทยานิพนธ์จริง 9 เล่ม — ทั้งสองวิธีให้ false positive ล้วน:

1. ฟ้อง "บรรทัดที่ไม่มีในเทมเพลต" → 5–6 แถวต่อเล่ม และทุกแถวถูกต้องอยู่แล้ว
   ("คณะพาณิชยศาสตร์และการบัญชี มหาวิทยาลัยธรรมศาสตร์", "ปีการศึกษา 2568") เพราะ
   **เนื้อหาบนหน้าปกส่วนใหญ่คือข้อมูลของนักศึกษาเอง** ที่มาแทนที่ช่องว่างในเทมเพลต
   แยกด้วยข้อความอย่างเดียวไม่ได้ว่าอันไหนคือข้อมูลที่กรอก อันไหนคือข้อความแปลกปลอม

2. ฟ้อง "บรรทัดคงที่ของเทมเพลตที่ขาดไป" → 10–12 แถวต่อเล่ม และทุกแถวถูกต้องอยู่แล้ว
   เพราะเทมเพลตมีทางเลือกที่ใช้พร้อมกันไม่ได้อยู่ในไฟล์เดียว (วิทยานิพนธ์ / สารนิพนธ์ /
   การค้นคว้าอิสระ · ตารางกรรมการแบบ 3 คน กับแบบ 5 คน · บรรทัด "(ถ้ามี)")
   "มีในเทมเพลต" จึงไม่เท่ากับ "ทุกเล่มต้องมี"

สคริปต์นี้จึงทำหน้าที่ต่างออกไป: **หา "ผู้ต้องสงสัย" มาให้คนตัดสิน** แล้วรายการที่อนุมัติ
ค่อยเข้าไปอยู่ใน `COVER_DEPRECATED` ของ `check_deep.py` ซึ่งเป็นกฎที่แม่นเพราะเป็นรายการปิด

เกณฑ์คัดผู้ต้องสงสัย: เป็นบรรทัดในส่วนหน้า · สั้นแบบข้อความสำเร็จรูป · โผล่ในหลายเล่ม
(แปลว่าไม่ใช่ข้อมูลเฉพาะคน) · และไม่ปรากฏในเทมเพลตทั้ง 3 ไฟล์เลย

ใช้:
    python3 scripts/find_cover_extras.py "โฟลเดอร์ที่มีไฟล์ .docx" [--min-files 2]
"""
import argparse
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

try:
    from docx import Document
except ImportError:
    sys.exit("ต้องติดตั้ง python-docx ก่อน:  pip install python-docx")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_deep import COVER_DEPRECATED, front_matter_lines  # noqa: E402

TEMPLATES = Path(__file__).resolve().parent.parent / "references" / "templates"


def normalise(text):
    return re.sub(r"\s+", " ", text or "").strip()


def template_corpus():
    """ข้อความเทมเพลตทั้งหมดรวมกัน — ใช้เช็คว่าวลีนี้ 'ไม่มีในเทมเพลตเลย' จริงไหม"""
    blob = []
    for path in sorted(TEMPLATES.glob("TULIBS-*.md")):
        blob.append(normalise(path.read_text(encoding="utf-8")).lower())
    if not blob:
        sys.exit(f"ไม่พบไฟล์เทมเพลตใน {TEMPLATES} — ต้องมีก่อนถึงจะบอกได้ว่าอะไร 'ไม่มีในเทมเพลต'")
    return "\n".join(blob)


def looks_like_boilerplate(line):
    """ข้อความสำเร็จรูปสั้น ๆ ไม่ใช่ชื่อเรื่อง/ชื่อคน/ประโยคยาว"""
    if not (6 <= len(line) <= 60):
        return False
    if re.search(r"\d{3,}", line):          # เลขยาว = รหัสนักศึกษา/ปี/เลขหน้า
        return False
    if line.count(" ") > 8:
        return False
    return True


def main():
    ap = argparse.ArgumentParser(description="เสนอบรรทัดส่วนเกินบนหน้าปกให้คนอนุมัติ")
    ap.add_argument("folder")
    ap.add_argument("--min-files", type=int, default=2,
                    help="ต้องโผล่กี่เล่มขึ้นไปถึงจะถือว่าเป็นข้อความสำเร็จรูป (ค่าเริ่มต้น 2)")
    args = ap.parse_args()

    corpus = template_corpus()
    known = {c.lower() for _, c in COVER_DEPRECATED}
    seen = defaultdict(set)

    files = [p for p in sorted(Path(args.folder).glob("*.docx"))
             if not p.name.startswith("~$")]
    if not files:
        sys.exit(f"ไม่พบไฟล์ .docx ใน {args.folder}")

    for path in files:
        try:
            doc = Document(path)
        except Exception as exc:
            print(f"ข้าม {path.name}: {exc}", file=sys.stderr)
            continue
        for _, line in front_matter_lines(doc):
            line = normalise(line)
            if not looks_like_boilerplate(line):
                continue
            if line.lower() in corpus:       # มีในเทมเพลต = ไม่ใช่ส่วนเกิน
                continue
            seen[line].add(path.name)

    rows = [(line, files_) for line, files_ in seen.items()
            if len(files_) >= args.min_files and line.lower() not in known]
    rows.sort(key=lambda r: (-len(r[1]), r[0]))

    print(f"ตรวจ {len(files)} ไฟล์ · เทียบกับเทมเพลต {len(list(TEMPLATES.glob('TULIBS-*.md')))} ไฟล์\n")
    if not rows:
        print("ไม่พบผู้ต้องสงสัยใหม่ — รายการใน COVER_DEPRECATED ครอบคลุมเท่าที่คลังนี้มี")
        return
    print("ผู้ต้องสงสัย (ยังไม่ใช่กฎ — ต้องตรวจเองก่อนว่าเป็นข้อความส่วนเกินจริง):\n")
    for line, files_ in rows:
        print(f"  พบใน {len(files_)} เล่ม : “{line}”")
        for name in sorted(files_)[:3]:
            print(f"        - {name[:60]}")
    print("\nอนุมัติแล้วให้เพิ่มลง COVER_DEPRECATED ใน scripts/check_deep.py")
    print("รูปแบบ:  (r\"<regex>\", \"<ข้อความที่จะแสดงในรายงาน>\"),")
    print("แล้วรัน scripts/test_false_positives.py ก่อนใช้จริงเสมอ")


if __name__ == "__main__":
    main()
