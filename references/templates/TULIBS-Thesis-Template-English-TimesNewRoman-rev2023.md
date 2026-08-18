# TULIBS Thesis Template — English (Times New Roman) rev.2023 — ข้อมูลอ้างอิงสำหรับตรวจรูปแบบ

> เอกสารนี้เป็นการถอด (extract) เนื้อหาและรูปแบบทั้งหมดจากไฟล์ต้นฉบับ
> `TULIBS-Thesis-template-English-Times_new_rev_2023.docx`
> เพื่อใช้เป็นข้อมูลอ้างอิงให้ AI ตรวจสอบรูปแบบวิทยานิพนธ์ของนักศึกษา
> ทุกย่อหน้าจะระบุ **ชื่อสไตล์ Word** ไว้ในวงเล็บเหลี่ยม `[...]` นำหน้าข้อความ
> พร้อมรายละเอียดฟอนต์ ขนาด การจัดวาง ระยะย่อหน้า และคอมเมนต์คำแนะนำ

## 0. สรุปข้อกำหนดสำคัญ (ตารางอ้างอิงเร็วสำหรับการตรวจ)

- **ขนาดกระดาษ:** 21.0 × 29.7 ซม. (A4)
- **ระยะขอบมาตรฐาน:** บน 3.81 ซม. / ล่าง 2.54 ซม. / ซ้าย 3.81 ซม. / ขวา 2.54 ซม.
- **ฟอนต์เริ่มต้นทั้งเอกสาร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

| สไตล์ (styleId) | ฟอนต์ | ขนาด | น้ำหนัก | การจัดวาง | ย่อหน้าบรรทัดแรก | ระยะก่อน | ระยะหลัง | ระยะบรรทัด |
|---|---|---|---|---|---|---|---|---|
| `Normal` (`Normal`) | Times New Roman | 12 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 0 | 0 | single |
| `heading 1` (`Heading1`) | Times New Roman | 20 pt | หนา | กึ่งกลาง (center) | 0 | 0 | 0 | 1.5 เท่า |
| `heading 2` (`Heading2`) | Times New Roman | 16.5 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 3` (`Heading3`) | Times New Roman | 15 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 4` (`Heading4`) | Times New Roman | 20 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 5` (`Heading5`) | Times New Roman | 20 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 6` (`Heading6`) | Times New Roman | 15 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 7` (`Heading7`) | Times New Roman | 15 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 8` (`Heading8`) | Times New Roman | 13 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `heading 9` (`Heading9`) | Times New Roman | 13 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 2pt | 0 | single |
| `TU_Chapter` (`TUChapter`) | TH Sarabun New | 18 pt | หนา | กึ่งกลาง (center) | 0 | 0 | 0 | 1.5 เท่า |
| `TU_Main Heading _Chapter1` (`TUMainHeadingChapter1`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter2` (`TUMainHeadingChapter2`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter3` (`TUMainHeadingChapter3`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter4` (`TUMainHeadingChapter4`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter5` (`TUMainHeadingChapter5`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter6` (`TUMainHeadingChapter6`) | Times New Roman | 12 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter7` (`TUMainHeadingChapter7`) | Times New Roman | 12 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Main Heading_Chapter8` (`TUMainHeadingChapter8`) | Times New Roman | 12 pt | หนา | ตามค่าเริ่มต้น | 0 | 0pt | 0 | single |
| `TU_Para_Sub-heading 1` (`TUParaSub-heading1`) | Times New Roman | 12 pt | ปกติ | ตามค่าเริ่มต้น | 1.19 นิ้ว | 0 | 0 | single |
| `TU_Para_Sub-heading 2` (`TUParaSub-heading2`) | Times New Roman | 12 pt | ปกติ | ตามค่าเริ่มต้น | 1.63 นิ้ว | 0 | 0 | single |
| `TU_Para_Sub-heading 3` (`TUParaSub-heading3`) | Times New Roman | 12 pt | ปกติ | ตามค่าเริ่มต้น | 1.63 นิ้ว | 0 | 0 | single |
| `TU_Paragraph_Normal` (`TUParagraphNormal`) | Times New Roman | 12 pt | ปกติ | ตามค่าเริ่มต้น | 0.8 นิ้ว | 0 | 0 | single |
| `TU_Sub-heading 1` (`TUSub-heading1`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 0.8 นิ้ว | 0pt | 0 | single |
| `TU_Sub-heading 2` (`TUSub-heading2`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 1.1 นิ้ว | 0pt | 0 | single |
| `TU_Sub-heading 3` (`TUSub-heading3`) | TH Sarabun New | 16 pt | หนา | ตามค่าเริ่มต้น | 1.4 นิ้ว | 0pt | 0 | single |
| `caption` (`Caption`) | Angsana New | 11 pt | ปกติ | ตามค่าเริ่มต้น | 0 | 0 | 10pt | single |

## 1. การตั้งค่าหน้ากระดาษ (Page Setup) ของทุกส่วนในเอกสาร

### ส่วนที่ 1

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- เลขหน้า: เริ่มนับที่=1
- Header (default): word/header1.xml

### ส่วนที่ 2

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- เลขหน้า: เริ่มนับที่=1
- Header (default): word/header2.xml

### ส่วนที่ 3

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- เลขหน้า: เริ่มนับที่=1
- Header (default): word/header3.xml

### ส่วนที่ 4

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- เลขหน้า: เริ่มนับที่=1
- Header (default): word/header4.xml

### ส่วนที่ 5

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- เลขหน้า: เริ่มนับที่=1
- Header (default): word/header5.xml

### ส่วนที่ 6

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 7

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 8

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 9

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 10

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 11

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 12

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 13

- ขนาดกระดาษ: 16838 × 11906 twips = 11.693 × 8.268 นิ้ว (29.7 × 21.0 ซม.) แนว landscape
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.0 นิ้ว (2.54 ซม.) / ขวา 1.5 นิ้ว (3.81 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Header (default): word/header6.xml
- Footer (default): word/footer1.xml

### ส่วนที่ 14

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Footer (default): word/footer2.xml

### ส่วนที่ 15

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 16

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 0.49 นิ้ว (1.25 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Header (default): word/header7.xml

### ส่วนที่ 17

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Header (default): word/header8.xml

### ส่วนที่ 18

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Header (default): word/header9.xml

### ส่วนที่ 19

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว

### ส่วนที่ 20

- ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
- ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
- ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
- Header (default): word/header10.xml

## 2. รายละเอียดสไตล์ (Word Styles) ที่กำหนดไว้ในเทมเพลต

ค่าที่แสดงคือ **ค่าที่มีผลจริง** หลังรวม docDefaults + สไตล์แม่ (basedOn) แล้ว

### ค่าเริ่มต้นของทั้งเอกสาร (docDefaults)

- ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
- ย่อหน้า: —

### สไตล์: `TU_Chapter`  (styleId = `TUChapter`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 1` (`Heading1`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา
- **ย่อหน้า:** จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1

### สไตล์: `TU_Chapter Char`  (styleId = `TUChapterChar`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 1 Char` (`Heading1Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading _Chapter1`  (styleId = `TUMainHeadingChapter1`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=1 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="1.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading _Chapter1 Char`  (styleId = `TUMainHeadingChapter1Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading_Chapter2`  (styleId = `TUMainHeadingChapter2`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=2 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="2.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter2 Char`  (styleId = `TUMainHeadingChapter2Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading_Chapter3`  (styleId = `TUMainHeadingChapter3`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=3 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="3.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter3 Char`  (styleId = `TUMainHeadingChapter3Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading_Chapter4`  (styleId = `TUMainHeadingChapter4`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=4 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="4.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter4 Char`  (styleId = `TUMainHeadingChapter4Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading_Chapter5`  (styleId = `TUMainHeadingChapter5`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=5 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="5.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter5 Char`  (styleId = `TUMainHeadingChapter5Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Main Heading_Chapter6`  (styleId = `TUMainHeadingChapter6`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
- **ย่อหน้า:** ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=6 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="6.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter6 Char`  (styleId = `TUMainHeadingChapter6Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

### สไตล์: `TU_Main Heading_Chapter7`  (styleId = `TUMainHeadingChapter7`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
- **ย่อหน้า:** ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=7 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="7.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter7 Char`  (styleId = `TUMainHeadingChapter7Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

### สไตล์: `TU_Main Heading_Chapter8`  (styleId = `TUMainHeadingChapter8`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 2` (`Heading2`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
- **ย่อหน้า:** ระยะก่อน=0pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=8 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="8.%1" เริ่มที่=1

### สไตล์: `TU_Main Heading_Chapter8 Char`  (styleId = `TUMainHeadingChapter8Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 2 Char` (`Heading2Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16.5pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Para_Sub-heading 1`  (styleId = `TUParaSub-heading1`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips)

### สไตล์: `TU_Para_Sub-heading 2`  (styleId = `TUParaSub-heading2`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips)

### สไตล์: `TU_Para_Sub-heading 3`  (styleId = `TUParaSub-heading3`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips)

### สไตล์: `TU_Paragraph_Normal`  (styleId = `TUParagraphNormal`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips)

### สไตล์: `TU_Sub-heading 1`  (styleId = `TUSub-heading1`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 3` (`Heading3`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, outlineLvl=2, keepNext

### สไตล์: `TU_Sub-heading 1 Char`  (styleId = `TUSub-heading1Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 3 Char` (`Heading3Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#1F4D78

### สไตล์: `TU_Sub-heading 2`  (styleId = `TUSub-heading2`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 4` (`Heading4`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, outlineLvl=3, keepNext

### สไตล์: `TU_Sub-heading 2 Char`  (styleId = `TUSub-heading2Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 4 Char` (`Heading4Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `TU_Sub-heading 3`  (styleId = `TUSub-heading3`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `heading 5` (`Heading5`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
- **ย่อหน้า:** ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, outlineLvl=4, keepNext

### สไตล์: `TU_Sub-heading 3 Char`  (styleId = `TUSub-heading3Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Heading 5 Char` (`Heading5Char`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=TH Sarabun New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา, สี=#2E74B5

### สไตล์: `Default Paragraph Font`  (styleId = `DefaultParagraphFont`, ประเภท = character)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

### สไตล์: `Footer Char`  (styleId = `FooterChar`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

### สไตล์: `Header Char`  (styleId = `HeaderChar`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

### สไตล์: `Heading 1 Char`  (styleId = `Heading1Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

### สไตล์: `Heading 2 Char`  (styleId = `Heading2Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16.5pt, สี=#2E74B5

### สไตล์: `Heading 3 Char`  (styleId = `Heading3Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, สี=#1F4D78

### สไตล์: `Heading 4 Char`  (styleId = `Heading4Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, ตัวเอียง, สี=#2E74B5

### สไตล์: `Heading 5 Char`  (styleId = `Heading5Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, สี=#2E74B5

### สไตล์: `Heading 6 Char`  (styleId = `Heading6Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, สี=#1F4D78

### สไตล์: `Heading 7 Char`  (styleId = `Heading7Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, ตัวเอียง, สี=#1F4D78

### สไตล์: `Heading 8 Char`  (styleId = `Heading8Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=10.5pt, ขนาด(ไทย/CS)=13pt, สี=#272727

### สไตล์: `Heading 9 Char`  (styleId = `Heading9Char`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=10.5pt, ขนาด(ไทย/CS)=13pt, ตัวเอียง, สี=#272727

### สไตล์: `Normal`  (styleId = `Normal`, ประเภท = paragraph)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

### สไตล์: `annotation reference`  (styleId = `CommentReference`, ประเภท = character)
- อ้างอิงจากสไตล์: `Default Paragraph Font` (`DefaultParagraphFont`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=8pt, ขนาด(ไทย/CS)=9pt

### สไตล์: `caption`  (styleId = `Caption`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=11pt, ตัวเอียง
- **ย่อหน้า:** ระยะหลัง=10pt

### สไตล์: `footer`  (styleId = `Footer`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

### สไตล์: `header`  (styleId = `Header`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

### สไตล์: `heading 1`  (styleId = `Heading1`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา
- **ย่อหน้า:** จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1

### สไตล์: `heading 2`  (styleId = `Heading2`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=13pt, ขนาด(ไทย/CS)=16.5pt, สี=#2E74B5
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=1 รูปแบบ=decimal แม่แบบข้อความ="%1.%2" เริ่มที่=1

### สไตล์: `heading 3`  (styleId = `Heading3`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, สี=#1F4D78
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=2, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=2 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3" เริ่มที่=1

### สไตล์: `heading 4`  (styleId = `Heading4`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, ตัวเอียง, สี=#2E74B5
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=3, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=3 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4" เริ่มที่=1

### สไตล์: `heading 5`  (styleId = `Heading5`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, สี=#2E74B5
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=4, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=4 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4.%5" เริ่มที่=1

### สไตล์: `heading 6`  (styleId = `Heading6`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, สี=#1F4D78
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=5, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=5 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4.%5.%6" เริ่มที่=1

### สไตล์: `heading 7`  (styleId = `Heading7`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=15pt, ตัวเอียง, สี=#1F4D78
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=6, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=6 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4.%5.%6.%7" เริ่มที่=1

### สไตล์: `heading 8`  (styleId = `Heading8`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=10.5pt, ขนาด(ไทย/CS)=13pt, สี=#272727
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=7, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=7 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4.%5.%6.%7.%8" เริ่มที่=1

### สไตล์: `heading 9`  (styleId = `Heading9`, ประเภท = paragraph)
- อ้างอิงจากสไตล์: `Normal` (`Normal`)
- ย่อหน้าถัดไปใช้สไตล์: `Normal`
- **ตัวอักษร:** ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=10.5pt, ขนาด(ไทย/CS)=13pt, ตัวเอียง, สี=#272727
- **ย่อหน้า:** ระยะก่อน=2pt, outlineLvl=8, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=8 รูปแบบ=decimal แม่แบบข้อความ="%1.%2.%3.%4.%5.%6.%7.%8.%9" เริ่มที่=1

## 3. รูปแบบเลขลำดับอัตโนมัติ (Numbering / Multilevel list)

### numId `0` → abstractNumId `None`

| ระดับ (ilvl) | รูปแบบตัวเลข | แม่แบบข้อความ | เริ่มที่ | ผูกกับสไตล์ |
|---|---|---|---|---|

## 4. เนื้อหาทั้งหมดของเทมเพลต เรียงตามลำดับหน้าจริง

**[Normal]** 		[IMAGE]
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS OR DISSERTATION TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=16pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using Times New Roman for all pages Using 16 point bold font with an uppercase in each alphabet except scientific name."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** BY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]**  FIRSTNAME SURNAME
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะหลัง=5pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** A THESIS OR DISSERTATION SUBMITTED IN PARTIAL FULFILLMENT OF THE REQUIREMENTS FOR THE DEGREE OF DEGREE TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

**[Normal]** DEPARTMENT  
FACULTY  
THAMMASAT UNIVERSITY  
ACADEMIC YEAR 20XX
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 3 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS OR DISSERTATION TITLE  

  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=16pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 16 point bold font with an uppercase in each alphabet except scientific name.."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** BY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]**  FIRSTNAME SURNAME
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=ชิดขอบทั้งสองข้าง (justified), ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะก่อน=12pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** A THESIS OR DISSERTATION SUBMITTED IN PARTIAL FULFILLMENT OF THE REQUIREMENTS FOR THE DEGREE OF DEGREE TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา

**[Normal]** DEPARTMENT  
FACULTY  
THAMMASAT UNIVERSITY  
ACADEMIC YEAR 20XX
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 3 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=14pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THAMMASAT UNIVERSITY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** FACULTY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font with an uppercase in each alphabet."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS OR DISSERTATION
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** BY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** FIRSTNAME SURNAME
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** ENTITLED
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** was approved as partial fulfillment of the requirements for
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** the degree of Degree Title
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

**[Normal]**   
on Approval date e.g. January 1, 2011 
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "M/D/Y of submitting the complete thesis/dissertation."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| Chairman `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
|---|---|
| Member and Advisor `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
| Member `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
| Dean `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THAMMASAT UNIVERSITY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** FACULTY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font with an uppercase in each alphabet."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS OR DISSERTATION
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** BY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** FIRSTNAME SURNAME
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** ENTITLED
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** THESIS TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** was approved as partial fulfillment of the requirements for
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** the degree of Degree Title
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

**[Normal]**   
on Approval date e.g. January 1, 2011 
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "M/D/Y of submitting the complete thesis/dissertation."

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| Chairman `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
|---|---|
| Member and Advisor `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
| Member and Co-adviser `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
| Member `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |
| Dean `[Normal]` | [IMAGE] `[Normal]`<br>(Academic Title Firstname Surname, Ph.D./M.D.) `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| Thesis Title `[Normal]` | THESIS TITLE `[Normal]` |
|---|---|
| Author `[Normal]` | Firstname Surname `[Normal]` |
| Degree `[Normal]` | Degree Title `[Normal]` |
| Major Field/Faculty/University `[Normal]` | Major Field `[Normal]`<br>Faculty `[Normal]`<br>Thammasat University `[Normal]` |
| Thesis Advisor `[Normal]`<br>Thesis Co-Advisor (If any) `[Normal]` | Academic Title Firstname Surname `[Normal]`<br>Academic Title Firstname Surname `[Normal]` |
| Academic Year `[Normal]` | 20xx `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** ABSTRACT
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Keywords: Insert keyword here, Insert keyword here, Insert keyword here  

  - ย่อหน้า: ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 3 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** ACKNOWLEDGEMENTS  

  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert new text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** 		Firstname Surname
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 2 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** TABLE OF CONTENTS
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** 		       Page
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| ABSTRACT `[Normal]` | (1) `[Normal]` |
|---|---|
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| ACKNOWLEDGEMENTS `[Normal]` | (2) `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| LIST OF TABLES (If any) `[Normal]` | (8) `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| LIST OF FIGURES (If any) `[Normal]` | (9) `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| LIST OF ABBREVIATIONS (If any) `[Normal]` | (10) `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| CHAPTER 1 INTRODUCTION `[Normal]` | 1 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| 1.1 Main heading `[Normal]` | 1 `[Normal]` |
| 1.1.1 Subheading 1 `[Normal]` | 3 `[Normal]` |
| 1.1.1.1 Subheading 2 `[Normal]` | 4 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 5 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 6 `[Normal]` |
| 1.2 Main heading `[Normal]` | 8 `[Normal]` |
| 1.2.1 Subheading 1 `[Normal]` | 9 `[Normal]` |
| 1.2.1.1 Subheading 2 `[Normal]` | 9 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 10 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 11 `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| CHAPTER 2 REVIEW OF LITERATURE `[Normal]` | 14 `[Normal]` |
|---|---|
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| 2.1 Main heading `[Normal]` | 14 `[Normal]` |
| 2.1.1 Subheading 1 `[Normal]` | 15 `[Normal]` |
| 2.1.1.1 Subheading 2 `[Normal]` | 16 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 17 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 18 `[Normal]` |
| 2.2 Main heading `[Normal]` | 19 `[Normal]` |
| 2.2.1 Subheading 1 `[Normal]` | 19 `[Normal]` |
| 2.2.1.1 Subheading 2 `[Normal]` | 20 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 21 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 22 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| CHAPTER 3 RESEARCH METHODOLOGY `[Normal]` | 23 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| 3.1 Main heading `[Normal]` | 23 `[Normal]` |
| 3.1.1 Subheading 1 `[Normal]` | 24 `[Normal]` |
| 3.1.1.1 Subheading 2 `[Normal]` | 25 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 26 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 27 `[Normal]` |
| 3.2 Main heading `[Normal]` | 28 `[Normal]` |
| 3.2.1 Subheading 1 `[Normal]` | 29 `[Normal]` |
| 3.2.1.1 Subheading 2 `[Normal]` | 30 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 32 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 33 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| CHAPTER 4 RESULTS AND DISCUSSION `[Normal]` | 35 `[Normal]` |
|---|---|
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| 4.1 Main heading `[Normal]` | 36 `[Normal]` |
| 4.1.1 Subheading 1 `[Normal]` | 37 `[Normal]` |
| 4.1.1.1 Subheading 2 `[Normal]` | 38 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 39 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 40 `[Normal]` |
| 4.2 Main heading `[Normal]` | 41 `[Normal]` |
| 4.2.1 Subheading 1 `[Normal]` | 42 `[Normal]` |
| 4.2.1.1 Subheading 2 `[Normal]` | 43 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 44 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 45 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| CHAPTER 5 CONCLUSIONS AND RECOMMENDATIONS `[Normal]` | 50 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| 5.1 Main heading `[Normal]` | 50 `[Normal]` |
| 5.1.1 Subheading 1 `[Normal]` | 51 `[Normal]` |
| 5.1.1.1 Subheading 2 `[Normal]` | 52 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 53 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 54 `[Normal]` |
| 5.2 Main heading `[Normal]` | 55 `[Normal]` |
| 5.2.1 Subheading 1 `[Normal]` | 56 `[Normal]` |
| 5.2.1.1 Subheading 2 `[Normal]` | 57 `[Normal]` |
| (1) Subheading 3 `[Normal]` | 58 `[Normal]` |
| (2) Subheading 3 `[Normal]` | 59 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| REFERENCES `[Normal]` | 60 `[Normal]` |
|---|---|
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| APPENDICES `[Normal]` | `[Normal: ว่าง]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| APPENDIX A `[Normal]` | 63 `[Normal]` |
| APPENDIX B `[Normal]` | 64 `[Normal]` |
| APPENDIX C `[Normal]` | 65 `[Normal]` |
| `[Normal: ว่าง]` | `[Normal: ว่าง]` |
| BIOGRAPHY `[Normal]` | 67 `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** LIST OF TABLES
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** 	Tables	 Page
  - ย่อหน้า: จัดวาง=ชิดขวา (right), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt


**[TABLE — ตาราง]**

*ความกว้าง=8010 (dxa) | เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| 1.1 Insert table title `[Normal]` | 10 `[Normal]` |
|---|---|
| 1.2 Insert table title `[Normal]` | 11 `[Normal]` |
| 2.1 Insert table title `[Normal]` | 12 `[Normal]` |
| 2.2 Insert table title `[Normal]` | 13 `[Normal]` |
| 3.1 Insert table title `[Normal]` | 14 `[Normal]` |
| 3.2 Insert table title `[Normal]` | 15 `[Normal]` |
| 4.1 Insert table title `[Normal]` | 16 `[Normal]` |
| 4.2 Insert table title `[Normal]` | 17 `[Normal]` |
| 5.1 Insert table title `[Normal]` | 18 `[Normal]` |
| 5.2 Insert table title `[Normal]` | 19 `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** LIST OF FIGURES
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** 	Figures	 Page
  - ย่อหน้า: ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt


**[TABLE — ตาราง]**

*ความกว้าง=8105 (dxa) | เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| 1.1 Insert figure title `[Normal]` | 20 `[Normal]` |
|---|---|
| 1.2 Insert figure title `[Normal]` | 21 `[Normal]` |
| 2.1 Insert figure title `[Normal]` | 22 `[Normal]` |
| 2.2 Insert figure title `[Normal]` | 23 `[Normal]` |
| 3.1 Insert figure title `[Normal]` | 24 `[Normal]` |
| 3.2 Insert figure title `[Normal]` | 25 `[Normal]` |
| 4.1 Insert figure title `[Normal]` | 26 `[Normal]` |
| 4.2 Insert figure title `[Normal]` | 27 `[Normal]` |
| 5.1 Insert figure title `[Normal]` | 28 `[Normal]` |
| 5.2 Insert figure title `[Normal]` | 29 `[Normal]` |

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะหลัง=8pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

`— PAGE BREAK: ขึ้นหน้าใหม่ —`

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** LIST OF ABBREVIATIONS
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| Symbols/Abbreviations `[Normal]` | Terms `[Normal]` |
|---|---|
| `[Normal: ว่าง]`<br>Insert text here `[Normal]`<br>Insert text here `[Normal]`<br>Insert text here `[Normal]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]`<br>`[Normal: ว่าง]` | `[Normal: ว่าง]`<br>Insert text here `[Normal]`<br>Insert text here `[Normal]`<br>Insert text here `[Normal]` |

**[TU_Chapter]** CHAPTER 1 
  - ย่อหน้า: จัดวาง=ชิดซ้าย (left), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1

**[heading 1]** CHAPTER 2   
INTRODUCTION
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading _Chapter1]** 1.1 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=1 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="1.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point bold font"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		1.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point bold font"

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 1.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Para_Sub-heading 3]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading _Chapter1]** 1.2 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=1 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="1.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		1.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 1.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 3]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 3   
REVIEW OF LITERATURE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter2]** 2.1 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=2 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="2.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		2.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 2.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter2]** 2.2 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=2 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="2.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		2.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 2.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 3]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 4   
RESEARCH METHODOLOGY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter3]** 3.1 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=3 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="3.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		3.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 3.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Para_Sub-heading 2]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter3]** 3.2 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=3 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="3.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		3.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 3.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 3]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 5   
RESULTS AND DISCUSSION
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter4]** 4.1 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=4 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="4.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		4.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 4.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter4]** 4.2 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=4 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="4.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		4.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 4.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 6   
CONCLUSIONS AND RECOMMENDATIONS
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter5]** 5.1 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=5 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="5.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		5.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 5.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Para_Sub-heading 2]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter5]** 5.2 Main heading
  - ย่อหน้า: เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=5 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="5.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		5.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 5.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 7   
CHAPTER TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter6]** 6.1 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=6 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="6.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		6.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 6.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter6]** 6.2 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=6 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="6.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		6.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 6.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[heading 1]** CHAPTER 8   
CHAPTER TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter7]** 7.1 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=7 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="7.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		7.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 7.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter7]** 7.2 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=7 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="7.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		7.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 7.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter6]** 6.3 
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=6 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="6.%1" เริ่มที่=1

**[heading 1]** CHAPTER 9   
CHAPTER TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext, เลขลำดับอัตโนมัติ: numId=16 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="CHAPTER %1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=20pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter8]** 8.1 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=8 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="8.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		8.1.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 8.1.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Main Heading_Chapter8]** 8.2 Main heading
  - ย่อหน้า: ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext, เลขลำดับอัตโนมัติ: numId=8 ilvl=0 รูปแบบ=decimal แม่แบบข้อความ="8.%1" เริ่มที่=1
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Sub-heading 1]** 		8.2.1 Subheading 1
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=2, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 1]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.19 นิ้ว (1714 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 2]** 8.2.1.1 Subheading 2
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.1 นิ้ว (1584 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=3, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[TU_Para_Sub-heading 2]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Sub-heading 3]** (1) Subheading 3
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.4 นิ้ว (2016 twips), ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=4, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=16pt, ตัวหนา

**[Normal]** Insert text here
  - ย่อหน้า: ย่อหน้าบรรทัดแรก=1.63 นิ้ว (2347 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Main Heading_Chapter8]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.25 นิ้ว, เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext

`[TU_Main Heading_Chapter8]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.25 นิ้ว, เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext

**[TU_Main Heading_Chapter8]** Table xx
  - ย่อหน้า: แขวน=0.25 นิ้ว, เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[TU_Main Heading_Chapter8]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.25 นิ้ว, เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext


**[TABLE — ตาราง]**

| `[Normal: ว่าง]` | `[Normal: ว่าง]` | `[Normal: ว่าง]` |
|---|---|---|
| `[Normal: ว่าง]` | `[Normal: ว่าง]` | `[Normal: ว่าง]` |

`[TU_Main Heading_Chapter8]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.25 นิ้ว, เยื้องซ้าย=0.25 นิ้ว, ระยะก่อน=0pt, ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=1, keepNext

**[TU_Chapter]** REFERENCES
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Books and Book Articles
  - ย่อหน้า: ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 12 point normal font"

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Articles
  - ย่อหน้า: ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Electronic Media
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Other Materials
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[TU_Chapter]** REFERENCES
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[Normal]** 1.	Insert text here
  - ย่อหน้า: ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here (continued)
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** 2. 	Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here (continued)
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** 				3. 	Insert text here	
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here (continued)
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** 4. 	Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here (continued)
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** 5. 	Insert text here
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

**[Normal]** Insert text here (continued)
  - ย่อหน้า: แขวน=0.5 นิ้ว, เยื้องซ้าย=0.5 นิ้ว, ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** APPENDICES
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=16pt, ขนาด(ไทย/CS)=16pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 16 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** APPENDIX A
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

**[TU_Chapter]** INSERT APPENDIX TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point bold font with an uppercase in each alphabet"

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here  

  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt
  - 💬 **คอมเมนต์คำแนะนำในเทมเพลต:** "Using 14 point normal font"

**[TU_Chapter]** APPENDIX B
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

**[TU_Chapter]** INSERT APPENDIX TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here  

  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Paragraph_Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** APPENDIX C
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

**[TU_Chapter]** INSERT APPENDIX TITLE
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะก่อน=6pt, ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Paragraph_Normal]** Insert text here  

  - ย่อหน้า: ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)
  - มีช่องกรอกข้อมูลแบบฟอร์ม (form field) 1 ช่องในย่อหน้านี้
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=12pt

`[TU_Paragraph_Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.8 นิ้ว (1152 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)

**[TU_Chapter]** BIOGRAPHY
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center), ระยะบรรทัด=1.5 เท่า (line=360 auto), outlineLvl=0, keepNext
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Times New Roman, ขนาด(ละติน)=14pt, ขนาด(ไทย/CS)=18pt, ตัวหนา

`[Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.985 นิ้ว (1418 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)


**[TABLE — ตาราง]**

*เส้นขอบ: top=none, left=none, bottom=none, right=none, insideH=none, insideV=none*

| Name `[Normal]` | Firstname Surname `[Normal]` |
|---|---|
| Educational Attainment `[Normal]` | Academic Year: Insert only bachelor and graduate degrees `[Normal]` |
| Scholarship (If any) `[Normal]` | Year XXXX: Scholarship title `[Normal]` |
| `[Normal: ว่าง]`<br>Publications `[Normal]`<br>`[Normal: ว่าง]` | `[Normal: ว่าง]` |
| Insert based on reference style `[Normal]`<br>`[Normal: ว่าง]` |

`[TU_Paragraph_Normal]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — ย่อหน้าบรรทัดแรก=0.0 นิ้ว (0 twips), ระยะบรรทัด=1.5 เท่า (line=360 auto)


> **[SECTION BREAK — ตั้งค่าหน้ากระดาษของส่วนที่จบตรงนี้]**

> - ขนาดกระดาษ: 11906 × 16838 twips = 8.268 × 11.693 นิ้ว (21.0 × 29.7 ซม.) แนว portrait
> - ระยะขอบ: บน 1.5 นิ้ว (3.81 ซม.) / ล่าง 1.0 นิ้ว (2.54 ซม.) / ซ้าย 1.5 นิ้ว (3.81 ซม.) / ขวา 1.0 นิ้ว (2.54 ซม.)
> - ระยะหัวกระดาษ (header) จากขอบบน 1.0 นิ้ว (2.54 ซม.), ท้ายกระดาษ (footer) จากขอบล่าง 0.49 นิ้ว (1.25 ซม.), gutter 0.0 นิ้ว
> - Header (default): word/header10.xml

## 5. เนื้อหาหัวกระดาษและท้ายกระดาษ (Headers / Footers)

### `word/footer1.xml`

**[footer]** [IMAGE]5959
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

### `word/footer2.xml`

`[footer]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

### `word/header1.xml`

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center)

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center)

### `word/header2.xml`

**[header]** (English title page in case of using different language)
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, สี=#FF0000

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center)

### `word/header3.xml`

**[header]** (Choose one of the following:)
  - ย่อหน้า: จัดวาง=กึ่งกลาง (center)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt, สี=#FF0000

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center)

### `word/header4.xml`

**[header]** ({FIELD: PAGE   \* MERGEFORMAT}2)
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

### `word/header5.xml`

> **[CONTENT CONTROL / ฟิลด์อัตโนมัติ —  [gallery=Page Numbers (Top of Page)]]**

**[header]** {FIELD: PAGE   \* MERGEFORMAT}2
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

### `word/header6.xml`

**[header]** 60
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=กึ่งกลาง (center)

### `word/header7.xml`

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=ชิดขวา (right)

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=ชิดขวา (right)

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*

### `word/header8.xml`

> **[CONTENT CONTROL / ฟิลด์อัตโนมัติ —  [gallery=Page Numbers (Top of Page)]]**

**[header]** 61
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=ชิดขวา (right)

### `word/header9.xml`

> **[CONTENT CONTROL / ฟิลด์อัตโนมัติ —  [gallery=Page Numbers (Top of Page)]]**

**[header]** 63
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

`[header]` *(ย่อหน้าว่างสำหรับเว้นระยะ)*  — จัดวาง=ชิดขวา (right)

### `word/header10.xml`

**[header]** [IMAGE](For optional)(For optional)67
  - ย่อหน้า: จัดวาง=ชิดขวา (right)
  - ตัวอักษร: ฟอนต์: ascii=Times New Roman, hAnsi=Times New Roman, cs (ไทย/complex)=Angsana New, ขนาด(ละติน)=12pt, ขนาด(ไทย/CS)=20pt

## 6. คอมเมนต์ทั้งหมดในเทมเพลต (คำแนะนำการจัดรูปแบบจากหอสมุด)

| # | ผู้เขียน | ข้อความคำแนะนำ |
|---|---|---|
| 0 | admin | Using Times New Roman for all pages Using 16 point bold font with an uppercase in each alphabet except scientific name. |
| 3 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 5 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 6 | admin | Using 16 point bold font with an uppercase in each alphabet except scientific name.. |
| 7 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 9 | admin | Using 12 point normal font with an uppercase in each alphabet. |
| 10 | admin | Using 12 point normal font |
| 11 | admin | M/D/Y of submitting the complete thesis/dissertation. |
| 13 | admin | Using 12 point normal font with an uppercase in each alphabet. |
| 14 | admin | Using 12 point normal font |
| 15 | admin | M/D/Y of submitting the complete thesis/dissertation. |
| 17 | admin | Using 12 point normal font with an uppercase in each alphabet except scientific name. |
| 18 | admin | Using 12 point normal font |
| 21 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 22 | admin | Using 12 point normal font |
| 23 | admin | Using 12 point normal font |
| 24 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 25 | admin | Using 12 point normal font |
| 26 | admin | Using 12 point normal font with an uppercase in each alphabet |
| 28 | admin | Using 12 point normal font |
| 35 | admin | Single digit number will be aligned at the left |
| 43 | admin | Numbering of tables of each chapter |
| 45 | admin | Using 12 point normal font |
| 48 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 49 | admin | Numbering of figures of each chapter |
| 50 | admin | Using 12 point normal font |
| 52 | admin | Using 12 point normal font |
| 53 | admin | Using 12 point bold font |
| 56 | admin | Using 12 point bold font |
| 63 | admin | Using 12 point normal font |
| 64 | admin | Using 16 point bold font with an uppercase in each alphabet |
| 66 | admin | Using 14 point bold font with an uppercase in each alphabet |
| 67 | admin | Using 14 point normal font |
| 69 | admin | Using 12 point normal font |
