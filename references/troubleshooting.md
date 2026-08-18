# แก้ปัญหาสภาพแวดล้อม (เปิดเมื่อเจอ error เท่านั้น)

ไฟล์นี้แยกออกมาจาก SKILL.md เพราะเป็นเรื่องที่เจอเฉพาะบางเครื่อง — ไม่ต้องอ่านตอนตรวจปกติ

## Codex บน macOS — soffice จะถูก sandbox ฆ่า (SIGABRT) เสมอ

อาการ: crash report ของ `soffice` เด้งขึ้นมา (`NSApplication → _RegisterApplication → abort`). **ไม่ใช่ไฟล์นักศึกษาเสีย และลง LibreOffice ใหม่ก็ไม่หาย** — seatbelt ของ Codex ไม่ให้เข้า LaunchServices/WindowServer

สกิลนี้ **จำไว้หลังเจอครั้งแรก** (`/tmp/tulibs-pagemap/soffice-blocked.json`) แล้วจะไม่เรียกซ้ำอีก จึงไม่มี crash report ซ้ำ ๆ และงานตรวจเดินต่อด้วย marker ทันที

อยากได้เลขแผ่นจริงใน Codex → **เปิด worker ค้างไว้ใน Terminal ปกติ (นอก Codex) หน้าต่างเดียว ครั้งเดียว**:

```bash
python3 ~/Downloads/tulibs-thesis-docx-checker/scripts/pagemap_service.py worker
```

แล้วรันตรวจใน Codex ตามปกติ — สคริปต์จะส่งงานให้ worker ผ่านคิวไฟล์ใน `/tmp` (ไม่ใช้เน็ต) ได้ `source: rendered_pdf` เหมือนรันบน Linux ทุกประการ (ทดสอบแล้วกับเล่ม 135 หน้า: จับคู่ข้อความ 85.4% เลขแผ่นตรง)

| สั่ง | ผล |
|---|---|
| `pagemap_service.py doctor` | บอกว่าเรียก soffice ได้ไหม / ถูกจำว่าบล็อกอยู่ไหม / worker เปิดอยู่ไหม |
| `pagemap_service.py reset` | ล้างสถานะ "ถูกบล็อก" เพื่อให้ลองจัดหน้าใหม่อีกครั้ง |
| `TULIBS_NO_RENDER=1` | สั่งข้ามการจัดหน้าใหม่ไปเลย (ใช้ marker ล้วน ไม่แตะ soffice) |

**ห้ามเรียก `soffice` / `libreoffice` เองด้วยมือ** — ให้สคริปต์เป็นคนเรียกเท่านั้น เพราะบน macOS ที่ถูก sandbox จะถูกฆ่าด้วย SIGABRT (`NSApplication → _RegisterApplication → abort`) แล้วทำให้เข้าใจผิดว่าไฟล์นักศึกษาเสีย. `resolve_page_data()` ดักกรณีนี้ให้แล้วและถอยไปใช้ marker เอง; `scripts/pagemap_service.py` มีไว้ให้ผู้ดูแลเปิด worker นอก sandbox ถ้าอยากได้เลขแผ่นจริงบนเครื่องที่ถูกบล็อก

**อย่าเขียนสคริปต์ใหม่** — ต่อยอดสองตัวนี้. ไฟล์เปิดไม่ได้ ให้แจ้งผู้ใช้แล้วทำขั้น 3 ด้วยตาอย่างเดียว.


## ไฟล์เปิดไม่ได้ / .doc เก่า

- `.doc` เก่าหรือไฟล์ที่ export จาก Google Docs: แปลงเป็น .docx ก่อน (`libreoffice --headless --convert-to docx`) แล้วค่อยตรวจ
- ไฟล์เปิดด้วย python-docx ไม่ได้: แจ้งผู้ใช้ตรง ๆ ว่าไฟล์เสียหรือถูกป้องกันไว้ อย่าเดาผลตรวจจากชื่อไฟล์
