#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pagemap_service.py — [ตัวเลือกเสริมสำหรับผู้ดูแล ไม่ใช่ขั้นตอนปกติของสกิล]

*** การตรวจตามปกติไม่ต้องใช้ไฟล์นี้และไม่ต้องใช้ LibreOffice เลย ***
ค่าเริ่มต้นของสกิลคือ --page-engine markers ซึ่งอ่านตำแหน่งแบ่งหน้าที่ Word บันทึกไว้ใน
ไฟล์ (w:lastRenderedPageBreak) ทำงานได้ทุกที่ ไม่ต้องติดตั้งอะไร ไม่ต้องเปิด Terminal ค้าง
ไฟล์นี้มีไว้เฉพาะกรณีที่ต้องการเลขแผ่นแบบ render (--page-engine rendered) บนเครื่องที่
LibreOffice ถูก sandbox บล็อก เช่นรันผ่าน Codex CLI บน macOS

ทำไมต้องมีไฟล์นี้
-----------------
บน macOS ตัว soffice จะสร้าง NSApplication เสมอ แม้สั่ง --headless (mac build มีแต่
vclplug_osx ไม่มี svp แบบ Linux) การสร้าง NSApplication ต้องคุยกับ LaunchServices/
WindowServer ซึ่ง sandbox ของ Codex บล็อกไว้ → โปรเซสถูก abort() ทันที
(crash log ชี้ที่ +[NSApplication sharedApplication] → _RegisterApplication → abort)

การตรวจ "รูปแบบ" ทั้งหมด (ฟอนต์/ขนาด/ระยะขอบ/สไตล์/ระยะบรรทัด) อ่านจาก OOXML ตรง ๆ
อยู่แล้ว ไม่ต้องใช้ LibreOffice. สิ่งเดียวที่ต้อง render จริงคือ **เลขแผ่นงานจริง**
(อะไรตกไปอยู่หน้าไหน) ไฟล์นี้จึงแยกงานส่วนนั้นออกมาเป็นบริการต่างหาก:

    [โปรเซสใน sandbox]  --คิวไฟล์ใน /tmp-->  [worker ที่ผู้ใช้เปิดไว้ใน Terminal ปกติ]
                        <--ข้อความรายแผ่น--

คิวเป็น "ไฟล์" ล้วน ไม่ใช้ network จึงทำงานได้แม้ sandbox ปิดเน็ตทั้งหมด

วิธีใช้
------
1) เปิด worker ใน Terminal.app ปกติ (นอก Codex) ค้างไว้ 1 ครั้ง:

       python3 scripts/pagemap_service.py worker

2) จากนั้นรันการตรวจตามปกติได้เลย — check_docx.py จะตรวจเจอ worker แล้วส่งงานให้เอง:

       python3 scripts/check_all.py thesis.docx --page-engine rendered \
           --auto-json work/auto.json --deep-json work/deep.json

3) วินิจฉัยสภาพแวดล้อมเมื่อมีปัญหา:

       python3 scripts/pagemap_service.py doctor

ตัวแปรสภาพแวดล้อม
-----------------
    TULIBS_PAGEMAP_DIR       โฟลเดอร์คิว (ค่าเริ่มต้น /tmp/tulibs-pagemap)
    TULIBS_PAGEMAP_TIMEOUT   วินาทีที่รอผลจาก worker (ค่าเริ่มต้น 240)
    TULIBS_PAGEMAP_DISABLE   ตั้งเป็น 1 เพื่อบังคับไม่ใช้ worker (เรียก soffice ตรง ๆ)
    TULIBS_SOFFICE           path ของ soffice ถ้าไม่ได้อยู่ใน PATH
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

DEFAULT_QUEUE = "/tmp/tulibs-pagemap"
HEARTBEAT_NAME = "worker.alive"
HEARTBEAT_MAX_AGE = 20.0        # วินาที — เกินนี้ถือว่า worker ตายแล้ว
HEARTBEAT_INTERVAL = 5.0
POLL_INTERVAL = 0.25
JOB_TTL = 3600.0                # เก็บกวาดไฟล์ค้างที่เก่ากว่านี้

MAC_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

START_HINT = (
    "วิธีแก้: เปิด worker ค้างไว้ใน Terminal ปกติ (นอก sandbox/Codex) หนึ่งครั้ง แล้วรันคำสั่งเดิมซ้ำ\n"
    "    python3 scripts/pagemap_service.py worker\n"
    "หรือถ้าไม่ต้องการเลขแผ่นงานจริง ให้ใช้ --page-engine markers "
    "(ตรวจรูปแบบได้ครบเหมือนเดิม แต่เลขแผ่นจะอ้างจากมาร์กเกอร์ของ Word)"
)


class SandboxBlockedError(RuntimeError):
    """soffice ถูก sandbox ฆ่า (SIGABRT/SIGKILL) — ไม่ใช่ไฟล์เสีย"""


# --------------------------------------------------------------------------
# จำว่า "เรียก soffice ที่นี่ไม่ได้" — เพื่อไม่ให้เรียกซ้ำแล้วเกิด crash report รัว ๆ
# --------------------------------------------------------------------------
BLOCKED_NAME = "soffice-blocked.json"
# ตัวแปรแวดล้อมที่บอกว่าเรากำลังอยู่ใน sandbox ของ Codex (มีบ้างไม่มีบ้างแล้วแต่รุ่น)
SANDBOX_ENV_HINTS = ("CODEX_SANDBOX", "CODEX_SANDBOX_NETWORK_DISABLED",
                     "TULIBS_NO_RENDER")


def blocked_marker(qdir=None):
    return os.path.join(queue_dir() if qdir is None else qdir, BLOCKED_NAME)


def remember_blocked(reason, qdir=None):
    """บันทึกว่าเครื่อง/สภาพแวดล้อมนี้เรียก soffice ไม่ได้"""
    try:
        _ensure_dirs(qdir)
        with open(blocked_marker(qdir), "w", encoding="utf-8") as fh:
            json.dump({"reason": str(reason), "soffice": find_soffice() or "",
                       "platform": sys.platform, "at": time.time()}, fh,
                      ensure_ascii=False)
    except OSError:
        pass


def blocked_reason(qdir=None):
    """คืนเหตุผลถ้ารู้อยู่แล้วว่าเรียก soffice ไม่ได้ — ไม่งั้นคืน None

    ตรวจสองทาง: ตัวแปรแวดล้อมที่ประกาศตรง ๆ และไฟล์จำจากครั้งก่อน
    เจตนาคือ **ห้ามลองเรียกซ้ำ** เพราะบน macOS การเรียกที่ถูกบล็อกแต่ละครั้ง
    ทำให้ระบบเขียน crash report ของ soffice ใหม่ทุกครั้ง ผู้ใช้เห็นแล้วตกใจ
    ทั้งที่งานตรวจยังเดินต่อได้ด้วย marker
    """
    for name in SANDBOX_ENV_HINTS:
        if os.environ.get(name):
            return (f"ปิดการจัดหน้าใหม่ไว้ผ่านตัวแปรแวดล้อม {name} "
                    "(สภาพแวดล้อมนี้เรียก LibreOffice ไม่ได้)")
    try:
        with open(blocked_marker(qdir), encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return None
    if data.get("soffice") and data["soffice"] != (find_soffice() or ""):
        return None                      # เปลี่ยน LibreOffice แล้ว ให้ลองใหม่ได้
    return data.get("reason") or "เคยลองเรียก soffice แล้วถูก sandbox ปิดกั้น"


def clear_blocked(qdir=None):
    try:
        os.remove(blocked_marker(qdir))
        return True
    except OSError:
        return False


# --------------------------------------------------------------------------
# ที่อยู่คิว
# --------------------------------------------------------------------------
def queue_dir():
    return os.environ.get("TULIBS_PAGEMAP_DIR") or DEFAULT_QUEUE


def _paths(qdir=None):
    qdir = qdir or queue_dir()
    return {
        "queue": qdir,
        "jobs": os.path.join(qdir, "jobs"),
        "out": os.path.join(qdir, "out"),
        "beat": os.path.join(qdir, HEARTBEAT_NAME),
    }


def _ensure_dirs(qdir=None):
    p = _paths(qdir)
    for key in ("queue", "jobs", "out"):
        os.makedirs(p[key], exist_ok=True)
    return p


def worker_alive(qdir=None):
    """True เมื่อมี worker เขียน heartbeat ล่าสุดภายใน HEARTBEAT_MAX_AGE วินาที"""
    if os.environ.get("TULIBS_PAGEMAP_DISABLE") == "1":
        return False
    beat = _paths(qdir)["beat"]
    try:
        return (time.time() - os.path.getmtime(beat)) <= HEARTBEAT_MAX_AGE
    except OSError:
        return False


def worker_info(qdir=None):
    try:
        with open(_paths(qdir)["beat"], encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


# --------------------------------------------------------------------------
# การแปลงจริง (ต้องรันในโปรเซสที่ไม่ถูก sandbox)
# --------------------------------------------------------------------------
def find_soffice():
    return (os.environ.get("TULIBS_SOFFICE")
            or shutil.which("soffice")
            or shutil.which("libreoffice")
            or (MAC_SOFFICE if os.path.exists(MAC_SOFFICE) else None))


def _pdf_to_pages(pdf_path):
    """ดึงข้อความรายแผ่นจาก PDF — ใช้ pdftotext ก่อน ไม่มีค่อย fallback pypdf"""
    pdftotext = shutil.which("pdftotext")
    if pdftotext:
        proc = subprocess.run(
            [pdftotext, "-layout", "-enc", "UTF-8", pdf_path, "-"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
        if proc.returncode == 0:
            pages = proc.stdout.decode("utf-8", "replace").split("\f")
            if pages and pages[-1] == "":
                pages.pop()
            if pages:
                return pages
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader  # type: ignore
        except ImportError:
            raise RuntimeError(
                "ไม่พบ pdftotext (poppler) และ pypdf สำหรับอ่านข้อความรายแผ่น — "
                "ติดตั้งอย่างใดอย่างหนึ่ง: brew install poppler | "
                "pip install pypdf --break-system-packages")
    reader = PdfReader(pdf_path)
    pages = [(page.extract_text() or "") for page in reader.pages]
    if not pages:
        raise RuntimeError("PDF ที่จัดหน้าแล้วไม่มีแผ่นงาน")
    return pages


def convert_docx_to_pages(path, timeout=240):
    """สั่ง LibreOffice จัดหน้า .docx เป็น PDF แล้วคืนข้อความรายแผ่น (list[str])

    ต้องรันในโปรเซสที่ไม่ถูก sandbox — ถ้าถูกบล็อกจะยก SandboxBlockedError
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("ไม่พบ LibreOffice/soffice สำหรับคำนวณแผ่นงานจริง")
    known = blocked_reason()
    if known:
        # เคยรู้แล้วว่าเรียกไม่ได้ → ไม่เรียกซ้ำ ไม่ให้เกิด crash report อีก
        raise SandboxBlockedError(known)
    with tempfile.TemporaryDirectory(prefix="tulibs-page-map-") as tmp:
        profile_dir = os.path.join(tmp, "lo-profile")
        os.makedirs(profile_dir, exist_ok=True)
        command = [
            soffice, f"-env:UserInstallation=file://{profile_dir}",
            "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
            "--convert-to", "pdf", "--outdir", tmp, os.path.abspath(path),
        ]
        proc = subprocess.run(command, stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE, timeout=timeout)
        pdf_path = os.path.join(
            tmp, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        if proc.returncode or not os.path.exists(pdf_path):
            detail = (proc.stderr or proc.stdout).decode("utf-8", "replace").strip()
            rc = proc.returncode
            # -6/134 = SIGABRT (NSApplication ถูก sandbox บล็อก), -9/137 = SIGKILL
            if rc in (-6, 134, -9, 137, -4, 132, -11, 139):
                message = (f"LibreOffice ถูกปิดกั้นโดย sandbox (signal {abs(rc)}): "
                           "soffice บน macOS ต้องเข้าถึง LaunchServices/WindowServer "
                           "ซึ่ง sandbox ของ Codex ไม่อนุญาต")
                remember_blocked(message)   # จำไว้ ครั้งหน้าจะไม่เรียกซ้ำ
                raise SandboxBlockedError(message)
            raise RuntimeError(
                f"จัดหน้า DOCX เป็น PDF ไม่สำเร็จ (exit {rc}): {detail or 'unknown error'}")
        return _pdf_to_pages(pdf_path)


# --------------------------------------------------------------------------
# ฝั่งลูกค้า (รันใน sandbox ได้) — ส่งงานเข้าคิวแล้วรอผล
# --------------------------------------------------------------------------
def submit_job(path, timeout=None, qdir=None):
    """ส่ง .docx ให้ worker แปลง แล้วคืนข้อความรายแผ่น (list[str])"""
    timeout = timeout or float(os.environ.get("TULIBS_PAGEMAP_TIMEOUT", 240))
    p = _ensure_dirs(qdir)
    job_id = f"{int(time.time())}-{uuid.uuid4().hex[:8]}"
    docx_tmp = os.path.join(p["jobs"], job_id + ".docx.part")
    docx_final = os.path.join(p["jobs"], job_id + ".docx")
    request = os.path.join(p["jobs"], job_id + ".request")
    result = os.path.join(p["out"], job_id + ".json")

    shutil.copyfile(path, docx_tmp)
    os.replace(docx_tmp, docx_final)           # atomic: worker เห็นไฟล์ครบเสมอ
    with open(request, "w", encoding="utf-8") as fh:
        json.dump({"id": job_id, "name": os.path.basename(path),
                   "submitted": time.time()}, fh)

    deadline = time.time() + timeout
    try:
        while time.time() < deadline:
            if os.path.exists(result):
                with open(result, encoding="utf-8") as fh:
                    payload = json.load(fh)
                if not payload.get("ok"):
                    raise RuntimeError(payload.get("error", "worker แปลงไฟล์ไม่สำเร็จ"))
                return payload["pages"]
            if not worker_alive(qdir):
                raise RuntimeError(
                    "pagemap worker หยุดทำงานระหว่างรอผล\n" + START_HINT)
            time.sleep(POLL_INTERVAL)
        raise RuntimeError(
            f"รอผลจาก pagemap worker เกิน {timeout:.0f} วินาที "
            "(ไฟล์ใหญ่มากให้เพิ่ม TULIBS_PAGEMAP_TIMEOUT)")
    finally:
        for f in (docx_tmp, docx_final, request, result):
            try:
                os.remove(f)
            except OSError:
                pass


# --------------------------------------------------------------------------
# worker
# --------------------------------------------------------------------------
def _write_heartbeat(p, soffice):
    tmp = p["beat"] + ".part"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump({"pid": os.getpid(), "ts": time.time(), "soffice": soffice,
                   "extractor": "pdftotext" if shutil.which("pdftotext") else "pypdf"},
                  fh)
    os.replace(tmp, p["beat"])


def _sweep(p):
    now = time.time()
    for folder in (p["jobs"], p["out"]):
        for name in os.listdir(folder):
            f = os.path.join(folder, name)
            try:
                if now - os.path.getmtime(f) > JOB_TTL:
                    os.remove(f)
            except OSError:
                pass


def run_worker(qdir=None):
    try:
        sys.stdout.reconfigure(line_buffering=True)   # ให้เห็น log ทันทีแม้ redirect
    except Exception:                                 # noqa: BLE001
        pass
    p = _ensure_dirs(qdir)
    soffice = find_soffice()
    if not soffice:
        sys.exit("ไม่พบ LibreOffice/soffice — ติดตั้งก่อน หรือกำหนด TULIBS_SOFFICE")
    print("=" * 68)
    print("TULIBS pagemap worker กำลังทำงาน — ปล่อยหน้าต่างนี้ค้างไว้")
    print(f"  คิว     : {p['queue']}")
    print(f"  soffice : {soffice}")
    print(f"  ดึงข้อความ: {'pdftotext' if shutil.which('pdftotext') else 'pypdf (fallback)'}")
    print("  หยุดด้วย Ctrl+C")
    print("=" * 68)

    last_beat = 0.0
    last_sweep = 0.0
    try:
        while True:
            now = time.time()
            if now - last_beat >= HEARTBEAT_INTERVAL:
                _write_heartbeat(p, soffice)
                last_beat = now
            if now - last_sweep >= 300:
                _sweep(p)
                last_sweep = now

            requests = sorted(f for f in os.listdir(p["jobs"]) if f.endswith(".request"))
            if not requests:
                time.sleep(POLL_INTERVAL)
                continue

            for req in requests:
                job_id = req[:-len(".request")]
                req_path = os.path.join(p["jobs"], req)
                docx_path = os.path.join(p["jobs"], job_id + ".docx")
                out_path = os.path.join(p["out"], job_id + ".json")
                try:
                    os.remove(req_path)          # กันประมวลผลซ้ำ
                except OSError:
                    continue
                if not os.path.exists(docx_path):
                    continue
                started = time.time()
                try:
                    pages = convert_docx_to_pages(docx_path)
                    payload = {"ok": True, "pages": pages,
                               "elapsed": round(time.time() - started, 2)}
                    print(f"[{time.strftime('%H:%M:%S')}] ✓ {job_id} "
                          f"→ {len(pages)} แผ่น ({payload['elapsed']}s)")
                except Exception as exc:                     # noqa: BLE001
                    payload = {"ok": False, "error": str(exc)}
                    print(f"[{time.strftime('%H:%M:%S')}] ✗ {job_id} → {exc}")
                tmp_out = out_path + ".part"
                with open(tmp_out, "w", encoding="utf-8") as fh:
                    json.dump(payload, fh, ensure_ascii=False)
                os.replace(tmp_out, out_path)
    except KeyboardInterrupt:
        print("\nปิด worker แล้ว")
    finally:
        try:
            os.remove(p["beat"])
        except OSError:
            pass


# --------------------------------------------------------------------------
# doctor
# --------------------------------------------------------------------------
def run_doctor(qdir=None):
    p = _paths(qdir)
    soffice = find_soffice()
    print("=" * 68)
    print("TULIBS pagemap doctor")
    print("=" * 68)
    print(f"queue dir        : {p['queue']}")
    print(f"soffice          : {soffice or '— ไม่พบ —'}")
    print(f"pdftotext        : {shutil.which('pdftotext') or '— ไม่พบ —'}")
    try:
        import pypdf  # noqa: F401
        print("pypdf            : มี (fallback ได้)")
    except ImportError:
        print("pypdf            : — ไม่พบ —")

    alive = worker_alive(qdir)
    info = worker_info(qdir)
    print(f"worker           : {'ทำงานอยู่ (pid %s)' % info.get('pid') if alive else 'ไม่ได้เปิด'}")
    known = blocked_reason(qdir)
    print(f"สถานะ soffice     : {'ถูกจำว่าเรียกไม่ได้ — ' + known if known else 'ยังไม่เคยถูกบล็อก'}")
    if known:
        print(f"  (ล้างสถานะนี้ด้วย: python3 {os.path.basename(__file__)} reset)")

    if not soffice:
        print("\n→ ไม่มี LibreOffice: ใช้ --page-engine markers ได้ "
              "(ตรวจรูปแบบครบ แต่เลขแผ่นอ้างจากมาร์กเกอร์ Word)")
        return 1

    if known and not alive:
        print("\n→ ข้ามการทดสอบเรียก soffice เพราะรู้อยู่แล้วว่าถูกบล็อก "
              "(เรียกซ้ำมีแต่จะสร้าง crash report เพิ่ม)")
        print("\n" + START_HINT)
        return 2

    print("\nทดสอบเรียก soffice ตรง ๆ จากโปรเซสนี้ ...")
    with tempfile.TemporaryDirectory(prefix="tulibs-doctor-") as tmp:
        probe = os.path.join(tmp, "probe.docx")
        try:
            from docx import Document
            d = Document()
            d.add_paragraph("probe")
            d.save(probe)
        except Exception as exc:                              # noqa: BLE001
            print(f"  สร้างไฟล์ทดสอบไม่ได้: {exc}")
            return 1
        try:
            pages = convert_docx_to_pages(probe, timeout=180)
            print(f"  ✓ สำเร็จ ({len(pages)} แผ่น) — โปรเซสนี้เรียก LibreOffice ได้โดยตรง "
                  "ไม่จำเป็นต้องเปิด worker")
            return 0
        except SandboxBlockedError as exc:
            print(f"  ✗ {exc}")
            print("\n" + START_HINT)
            return 2
        except Exception as exc:                              # noqa: BLE001
            print(f"  ✗ {exc}")
            return 1


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "doctor"
    if mode == "worker":
        run_worker()
        return 0
    if mode == "doctor":
        return run_doctor()
    if mode == "reset":
        print("ล้างสถานะ 'soffice ถูกบล็อก' แล้ว — ครั้งหน้าจะลองจัดหน้าใหม่อีกครั้ง"
              if clear_blocked() else "ไม่มีสถานะที่ต้องล้าง")
        return 0
    if mode in ("-h", "--help", "help"):
        print(__doc__)
        return 0
    sys.exit(f"โหมดไม่ถูกต้อง: {mode} (ใช้ worker | doctor | reset)")


if __name__ == "__main__":
    sys.exit(main() or 0)
