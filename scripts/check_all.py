#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""รัน format + deep checks โดย unzip/parse DOCX เพียงครั้งเดียว.

ใช้แทนการเรียก check_docx.py และ check_deep.py แยก process เมื่อต้องการผลทั้งสองชุด:

    python3 scripts/check_all.py thesis.docx \
        --auto-json work/auto.json --deep-json work/deep.json
"""
import argparse
import json
import os
import sys
import time

from docx import Document

import check_deep
import check_docx


def _write_json(path, payload):
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def run_all(path, auto_json="auto.json", deep_json="deep.json", profile=None,
            report_out=None, title=None, time_note=None, page_engine="markers"):
    started = time.time()
    doc = Document(path)
    profile = profile or check_docx.detect_profile(doc)
    page_data = check_docx.resolve_page_data(path, doc, page_engine)

    auto_payload, auto_findings = check_docx.inspect_document(
        doc, path, profile, started_at=started, page_data=page_data)
    deep_started = time.time()
    deep_payload, deep_findings = check_deep.inspect_document(
        doc, path, profile, started_at=deep_started, page_data=page_data)

    _write_json(auto_json, auto_payload)
    _write_json(deep_json, deep_payload)

    total = time.time() - started
    print("=" * 70)
    print(f"TULIBS combined check — {os.path.basename(path)}   profile: {profile}")
    print("format: 🔴 {critical}  🟠 {major}  🟡 {minor}".format(**auto_payload["counts"]))
    print("deep:   🔴 {critical}  🟠 {major}  🟡 {minor}".format(**deep_payload["counts"]))
    print(f"รวม {total:.2f} วินาที · โหลด DOCX ครั้งเดียว")
    print(f"แผ่นงาน: {page_data.get('total_pages', 0)} · แหล่งเลขแผ่น: "
          f"{page_data.get('source', 'unavailable')} · "
          f"จับคู่ข้อความ {page_data.get('match_ratio', 0):.1%}")
    section_pages = auto_payload.get("section_pages", [])
    if section_pages:
        print("Section pages: " + "; ".join(
            f"{heading}=แผ่นที่ {page}" for heading, page in section_pages[:16]))
    elif not auto_payload.get("has_page_numbers"):
        print("⚠ ระบุแผ่นงานไม่ได้ — รายงานจะอ้างตำแหน่งด้วยข้อความที่ยกมาแทนเลขแผ่น\n"
              "  (ไม่กระทบผลตรวจ: รายงานยึดข้อความสำหรับ Ctrl+F อยู่แล้ว)")
    if page_data.get("warning"):
        print(f"⚠ {page_data['warning']}")
    print(f"JSON written: {auto_json}, {deep_json}")

    if report_out:
        from make_report import build_report, merge
        merged = merge([auto_payload, deep_payload])
        build_report(merged, report_out, title, time_note or f"{total:.1f} วินาที")
        print(f"DOCX report written: {report_out}")

    return auto_payload, deep_payload


def main():
    ap = argparse.ArgumentParser(
        description="ตรวจ TULIBS format + deep/APA checks โดยโหลด DOCX ครั้งเดียว")
    ap.add_argument("docx")
    ap.add_argument("--profile", choices=list(check_docx.PROFILES), default=None)
    ap.add_argument("--auto-json", default="auto.json")
    ap.add_argument("--deep-json", default="deep.json")
    ap.add_argument("--report", default=None, help="สร้างรายงาน DOCX จากผลอัตโนมัติทั้งสองชุด")
    ap.add_argument("--title", default=None)
    ap.add_argument("--time", dest="time_note", default=None)
    ap.add_argument("--page-engine", choices=("auto", "rendered", "markers"),
                    default="markers",
                    help="markers (ค่าเริ่มต้น) = ใช้ตำแหน่งแบ่งหน้าที่ Word บันทึกไว้ในไฟล์ "
                         "ไม่ต้องใช้ LibreOffice · rendered = ตัวเลือกเสริม จัดหน้าใหม่เป็น PDF "
                         "เฉพาะเครื่องที่เรียก LibreOffice ได้")
    args = ap.parse_args()
    if not os.path.exists(args.docx):
        sys.exit(f"File not found: {args.docx}")
    try:
        run_all(args.docx, args.auto_json, args.deep_json, args.profile,
                args.report, args.title, args.time_note, args.page_engine)
    except RuntimeError as exc:
        sys.exit(f"คำนวณแผ่นงานไม่สำเร็จ: {exc}")


if __name__ == "__main__":
    main()
