#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
page_truth.py — บอกว่า "เลขแผ่นที่สคริปต์เดาไว้ เชื่อได้แค่ไหน"

ทำไมต้องมี
----------
`--page-engine markers` เดาเลขแผ่นจาก `w:lastRenderedPageBreak` ซึ่งเป็น**แคช**
จากตอนที่ Word เรนเดอร์ครั้งสุดท้าย ไม่ใช่ผลการจัดหน้าปัจจุบัน สองอย่างจึงพลาดได้:

  1. ไฟล์ที่แก้แล้วเซฟด้วยโปรแกรมอื่น → marker ค้างอยู่ที่ตำแหน่งเก่า
  2. บางหน้า Word ไม่ได้บันทึก marker ไว้ → ตัวนับข้ามไป

ผลคือเลขแผ่น**คลาดสะสม** ยิ่งท้ายเล่มยิ่งเพี้ยน วัดกับเล่มจริง 120 แผ่น:
marker บอก 115 แผ่น (ขาด 5) ทั้งที่แก้บั๊กนับซ้ำแล้ว

สคริปต์นี้เรนเดอร์ PDF จริงแล้วเทียบให้ ได้คำตอบ 3 อย่าง:
  - เลขแผ่นจริงของข้อความที่สนใจ (pagemap.json)
  - ค่า drift ระหว่าง marker กับของจริง
  - ธง `page_claims_allowed` — finding ประเภท "ข้ามหน้า/เกิน 1 หน้า" อ้างได้ไหม

วิธีใช้
------
    python3 scripts/page_truth.py thesis.docx                 # รายงานบนจอ
    python3 scripts/page_truth.py thesis.docx -o pagemap.json # เขียนไฟล์ให้ agent ใช้
    python3 scripts/page_truth.py thesis.docx --probe "บทคัดย่อ" --probe "คำสำคัญ:"

ไม่มี LibreOffice / ถูก sandbox บล็อก → exit 3 พร้อมบอกว่า page_claims_allowed=false
(ตรวจรูปแบบอย่างอื่นยังทำได้ครบ ดู references/troubleshooting.md)
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# drift ที่ยอมให้ยังถือว่า marker พอเชื่อได้สำหรับ "จัดลำดับ" (ไม่ใช่สำหรับอ้างในรายงาน)
ORDER_TOLERANCE = 2


def _find_soffice():
    return (os.environ.get("TULIBS_SOFFICE")
            or shutil.which("soffice")
            or shutil.which("libreoffice"))


def render_pages(docx_path):
    """คืน list ข้อความรายแผ่นจาก PDF จริง; None ถ้าเรนเดอร์ไม่ได้"""
    soffice = _find_soffice()
    if not soffice:
        return None, "ไม่พบ soffice/libreoffice ในเครื่อง"
    tmp = tempfile.mkdtemp(prefix="tulibs-pagetruth-")
    try:
        # เรนเดอร์จาก **สำเนา** เสมอ — soffice วางไฟล์ล็อก .~lock.<ชื่อ># ไว้ข้าง ๆ
        # ไฟล์ที่เปิด ถ้าชี้ไปที่ต้นฉบับก็เท่ากับไปเขียนในโฟลเดอร์ของนักศึกษา
        # (สกิลนี้ห้ามแตะไฟล์ต้นฉบับและโฟลเดอร์ต้นฉบับ — ดู "อ่านอย่างเดียว" ใน SKILL.md)
        work = os.path.join(tmp, "src.docx")
        shutil.copyfile(docx_path, work)
        profile = os.path.join(tmp, "lo-profile")
        os.makedirs(profile, exist_ok=True)
        r = subprocess.run(
            [soffice, "-env:UserInstallation=file://" + profile,
             "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
             "--convert-to", "pdf", "--outdir", tmp, work],
            capture_output=True, text=True, timeout=900)
        pdf = os.path.join(tmp, "src.pdf")
        if not os.path.exists(pdf):
            return None, "soffice ไม่ได้สร้าง PDF: " + (r.stderr or r.stdout or "")[:200]
        if not shutil.which("pdftotext"):
            return None, "ไม่พบ pdftotext (ติดตั้ง poppler-utils)"
        t = subprocess.run(["pdftotext", "-layout", pdf, "-"],
                           capture_output=True, text=True, timeout=300).stdout
        pages = t.split("\f")
        if pages and not pages[-1].strip():
            pages.pop()
        return pages, None
    except subprocess.TimeoutExpired:
        return None, "เรนเดอร์ไม่ทันเวลา"
    except Exception as e:                                    # noqa: BLE001
        return None, "%s: %s" % (type(e).__name__, e)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def marker_index(docx_path):
    """คืน (page_by_para_index, total, paragraphs) จากตัวนับ marker ของสกิล"""
    from docx import Document
    import check_docx
    doc = Document(docx_path)
    idx, has_markers = check_docx.build_page_index(doc)
    total = max(idx) if idx else 1
    return idx, total, doc.paragraphs, has_markers


def real_page_of(text, pages):
    n = _norm(text)
    if not n:
        return None
    for i, pg in enumerate(pages, 1):
        if n in _norm(pg):
            return i
    return None


def build(docx_path, probes=None, min_len=12):
    idx, marker_total, paras, has_markers = marker_index(docx_path)
    pages, err = render_pages(docx_path)

    out = {
        "file": os.path.basename(docx_path),
        "marker_total": marker_total,
        "has_markers": has_markers,
        "rendered_total": None,
        "drift": None,
        "engine": "markers",
        "page_claims_allowed": False,
        "render_error": err,
        "probes": [],
        "pagemap": {},
    }
    if pages is None:
        return out

    out["rendered_total"] = len(pages)
    out["drift"] = len(pages) - marker_total
    out["engine"] = "rendered"
    out["page_claims_allowed"] = True

    # ข้อความยาวพอที่ค้นได้ไม่กำกวม → เก็บเลขแผ่นจริงไว้ให้ agent
    for i, p in enumerate(paras):
        t = (p.text or "").strip()
        if len(t) < min_len:
            continue
        key = t[:80]
        if key in out["pagemap"]:
            continue
        rp = real_page_of(t[:60], pages)
        if rp:
            out["pagemap"][key] = {"sheet": rp, "marker": idx[i] if i < len(idx) else None}

    for probe in (probes or []):
        m = None
        for i, p in enumerate(paras):
            if probe in (p.text or ""):
                m = idx[i] if i < len(idx) else None
                break
        out["probes"].append({"text": probe, "marker": m,
                              "real": real_page_of(probe, pages)})
    return out


def main():
    ap = argparse.ArgumentParser(description="วัดความน่าเชื่อถือของเลขแผ่นและสร้าง page map จริง")
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", help="เขียน pagemap JSON ลงไฟล์")
    ap.add_argument("--probe", action="append", default=[],
                    help="ข้อความที่อยากรู้เลขแผ่นจริง (ใส่ซ้ำได้)")
    a = ap.parse_args()

    if not os.path.exists(a.docx):
        print("ไม่พบไฟล์: " + a.docx, file=sys.stderr)
        return 2

    d = build(a.docx, a.probe)

    print("ไฟล์            : %s" % d["file"])
    print("marker เดาได้   : %s แผ่น%s" % (d["marker_total"],
                                          "" if d["has_markers"] else "  (ไฟล์ไม่มี marker เลย)"))
    if d["rendered_total"] is None:
        print("เรนเดอร์จริง    : ทำไม่ได้ — %s" % d["render_error"])
        print()
        print("⚠️  page_claims_allowed = false")
        print("    ห้ามออก finding ที่อ้างการข้ามหน้า/เกินหน้า (B1 องค์ประกอบปกอยู่หน้าเดียว,")
        print("    B3 บทคัดย่อ ≤ 1 หน้า, B4 กิตติกรรมประกาศ ≤ 1 หน้า, แต่ละบทขึ้นหน้าใหม่)")
        print("    เกณฑ์อื่นที่อ่านจาก OOXML ตรวจได้ตามปกติ")
        if a.out:
            with open(a.out, "w", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False, indent=2)
            print("เขียน %s แล้ว" % a.out)
        return 3

    print("เรนเดอร์จริง    : %s แผ่น" % d["rendered_total"])
    print("drift           : %+d แผ่น" % d["drift"])
    print("page map        : %d ย่อหน้า" % len(d["pagemap"]))
    print()
    if d["probes"]:
        print("%-46s %8s %8s" % ("ข้อความ", "marker", "จริง"))
        print("-" * 66)
        for p in d["probes"]:
            flag = "" if p["marker"] == p["real"] else "   ← ไม่ตรง"
            print("%-46s %8s %8s%s" % (p["text"][:44], p["marker"], p["real"], flag))
        print()
    if abs(d["drift"]) > ORDER_TOLERANCE:
        print("⚠️  marker คลาด %+d แผ่น — ใช้เลขแผ่นจาก pagemap นี้เท่านั้น" % d["drift"])
    print("✅ page_claims_allowed = true (มีเลขแผ่นจริงให้อ้างแล้ว)")

    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        print("เขียน %s แล้ว" % a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
