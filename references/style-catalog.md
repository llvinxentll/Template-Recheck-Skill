# TULIBS Named-Style Catalog (ค่าจริงจาก styles.xml)

ตารางนี้คือ "ค่าที่ถูกต้อง" ของแต่ละ named style ในเทมเพลต ใช้เทียบว่าไฟล์นักศึกษา
แก้สไตล์เพี้ยนไปหรือไม่ (เช่น เปลี่ยนขนาด/ฟอนต์/การเยื้องของ TU_Paragraph_Normal)
sz = ascii size (ละติน), szCs = complex-script size (ไทย), หน่วยเป็น pt

## โปรไฟล์ thai / english (rev.2024) — ฟอนต์ TH Sarabun New ทั้งหมด

| Style name | sz | szCs | bold | align | firstLine | left |
|---|---|---|---|---|---|---|
| Normal | (docDefault 16) | 16 | ✗ | — | — | — |
| TU_Chapter | — | 18 | ✓ | — (จัดกึ่งกลางที่ย่อหน้า) | — | — |
| TU_Paragraph_Normal | 16 | 16 | ✗ | — | **0.80"** | — |
| TU_Main Heading _Chapter1 … _Chapter8 | 16 | 16 | ✓ | — | — | 0.0" (thai) / **0.25"** (english) |
| TU_Sub-heading 1 | 16 | 16 | ✓ | — | **0.80"** | — |
| TU_Sub-heading 2 | (16) | 16 | ✓ | — | **1.10"** | — |
| TU_Sub-heading 3 | (16) | 16 | ✓ | — | **1.40"** | — |
| TU_Para_Sub-heading 1 | 16 | 16 | ✗ | — | **1.19"** | — |
| TU_Para_Sub-heading 2 | 16 | 16 | ✗ | — | **1.63"** | — |
| TU_Para_Sub-heading 3 | 16 | 16 | ✗ | — | **1.63"** | — |
| heading 1 (built-in, = ใช้กับ TU_Chapter) | 18 | 20 | ✓ | center | — | 0.0" |

## โปรไฟล์ english-times (rev.2023) — ละติน Times New Roman / ไทย TH Sarabun New

| Style name | sz (Times) | szCs (Sarabun) | bold | firstLine | left |
|---|---|---|---|---|---|
| Normal | (docDefault 12) | 12 | ✗ | — | — |
| TU_Chapter | 14 | 18 | ✓ | — | — |
| TU_Paragraph_Normal | 12 | 16 | ✗ | **0.80"** | — |
| TU_Main Heading _ChapterN | **12** | 16 | ✓ | — | **0.25"** |
| TU_Sub-heading 1 | 12 | 16 | ✓ | **0.80"** | — |
| TU_Sub-heading 2 | 12 | 16 | ✓ | **1.10"** | — |
| TU_Sub-heading 3 | 12 | 16 | ✓ | **1.40"** | — |
| TU_Para_Sub-heading 1/2/3 | 12 | 16 | ✗ | 1.19 / 1.63 / 1.63" | — |
| heading 1 | 14 | 20 | ✓ (center, line 1.5) | — | — |

## วิธีใช้ตารางนี้ตอนตรวจ

1. เปิด `student.docx` ด้วย python-docx อ่าน `doc.styles`
2. เช็คว่าสไตล์ TU_* ข้างต้นยังมีครบ (ขาด = ไฟล์อาจไม่ได้มาจากเทมเพลต → major)
3. สุ่มเทียบค่าของสไตล์หลัก (TU_Paragraph_Normal firstLine ต้อง 0.80", TU_Main Heading left ตรง profile, TU_Sub-heading 1–3 firstLine ตรงชั้น)
4. ถ้านักศึกษาพิมพ์เนื้อหาโดย **ไม่ใช้สไตล์** (ใช้ Normal + จัดมือ) ให้ตรวจที่ระดับ run/paragraph แทน (ฟอนต์/ขนาด/เยื้อง/หนา) แล้วแนะนำให้กลับไปใช้สไตล์ของเทมเพลตเพื่อความสม่ำเสมอ

## สัญญาณว่าเทมเพลตถูกดัดแปลง (ตรวจเน้น)

- TU_Paragraph_Normal firstLine ≠ 0.80"  → นักศึกษาแก้ค่าเยื้อง
- TU_Main Heading left เปลี่ยน (เช่น thai ควร 0" แต่เป็น 0.25")
- ฟอนต์ในสไตล์ถูกเปลี่ยนจาก TH Sarabun New เป็นฟอนต์อื่น
- ขนาด body ในสไตล์ถูกดันเป็น 14/15/17 pt เพื่อยัดหน้า
