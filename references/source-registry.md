# Source registry และ version guard

ตรวจสถานะแหล่งทางการล่าสุดก่อนอ้างว่าไฟล์ “ผ่านเกณฑ์ปัจจุบัน”:

| แหล่ง | URL ทางการ | สถานะที่ยืนยัน 2026-08-04 |
|---|---|---|
| คู่มือนักศึกษา TU e-Thesis | https://tuethesis.library.tu.ac.th/manual/TU%20e-Thesis-Manual_Student_TH.pdf | ฉบับปรับปรุงครั้งที่ 12 (มีนาคม 2569/2026) และชี้ไปหน้า Thesis Manual/Templates ด้านล่าง |
| Thesis Manual and Templates | https://library.tu.ac.th/bbs/content/94_628 | เป็นหน้าที่คู่มือปัจจุบันอ้างถึง แต่ automated fetch อาจได้ HTTP 403; ต้องเปิดผ่าน browser/ให้ผู้ใช้แนบ template เพื่อยืนยัน revision ของไฟล์จริง |
| APA Style — DOIs and URLs | https://apastyle.apa.org/style-grammar-guidelines/references/dois-urls | หน้าอ้างอิงทางการสำหรับรูปแบบ DOI/URL |

## กฎการใช้งาน

1. ถือ profile ใน `docx-spec.md` เป็น **pinned ruleset ที่สกัดจาก template rev.2024/rev.2023** ไม่ใช่คำยืนยันว่าเป็น revision ล่าสุดในปีปัจจุบัน. การ recheck ปกติด้วย profile ที่ระบุ **ไม่ต้อง browse**.
2. Browse/เปิดหน้า official templates เฉพาะเมื่อผู้ใช้ขอ “ล่าสุด/current”, แนบ template revision ใหม่ หรือไม่ได้ยืนยัน profile. จากนั้นเทียบ `styles.xml`/`sectPr`. ถ้าเข้าถึงไม่ได้ ให้ระบุในรายงานว่า “ตรวจด้วย pinned ruleset; revision ปัจจุบันยังไม่ยืนยัน”.
3. ถ้าคณะ/สถาบันมี template เฉพาะ (เช่น SIIT) ให้ใช้ template ของหน่วยงานนั้นเป็นหลักและอย่าเอาเกณฑ์ TULIBS ทั่วไปไปตัดสินโดยอัตโนมัติ.
4. บันทึกชื่อ profile/revision และวันที่ตรวจไว้ในรายงานทุกครั้ง.
