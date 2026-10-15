# Check EDI Able

เครื่องมือตรวจสอบเอกสารขาเข้าเรือก่อนจัดทำใบขนสินค้าขาเข้า (Thailand ocean import) —
เทียบ **MANIFEST** (รายงานจากสายเรือ, .xls/.xlsx **หรือ PDF "CARGO MANIFEST"**) กับ **ENTER**
(ฟอร์มของ Able, .pdf) ทีละ B/L ย่อย แล้วสร้างรายงาน `EDI.xlsx` แบบไฮไลต์สีบอกจุดที่ต้องแก้ก่อนยื่นใบขน

> **ข้อจำกัด:** หน้าเว็บรองรับเฉพาะ **ENTER ฟอร์มของ Able**
> ถ้าเจอฟอร์มของเจ้าอื่นหน้าเว็บจะแจ้งว่าไม่ใช่ฟอร์มของ Able · ไฟล์ CNTRS ไม่นำมาเทียบ

## Check EDI Able บนเว็บ (ให้เพื่อนใช้ได้ ไม่ต้องติดตั้ง)

https://rattanao.github.io/check-edi-Able/

1. ลากไฟล์ **MANIFEST** (.xls/.xlsx หรือ PDF) และ **ENTER ของ Able** (.pdf ได้หลายไฟล์) มาวางพร้อมกัน
   (ตั้งชื่ออะไรก็ได้ — PDF ที่เป็น CARGO MANIFEST จะถูกจัดเป็น MANIFEST ให้เอง, PDF อื่นเป็น ENTER;
   ไฟล์ที่ชื่อมี `CNTRS` จะถูกข้าม)
2. ใส่ SHED NO. ที่แจ้ง (ถ้ามี) แล้วกด **ตรวจสอบ**
3. ดูผลบนหน้าเว็บ (ตารางสีเหมือน EDI.xlsx, ค้นหา/กรอง "แสดงเฉพาะที่ไม่ตรง") หรือกด **ดาวน์โหลด EDI.xlsx**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่มีการอัปโหลดไปที่ใด

โค้ดหน้าเว็บอยู่ที่ `docs/` (`index.html` + `core.js` — พอร์ตจาก `build_edi.py` เฉพาะฟอร์มของ Able,
ฐานเดียวกับ [Check EDI MAC-NELS](https://rattanao.github.io/MAC-NELS/) และ
[Check EDI Penanshin](https://rattanao.github.io/Penanshin/))

## ฟอร์ม ENTER ที่รองรับ

หัวตาราง `MARKS & NOS` / `DESCRIPTIONS OF GOODS` / `GROSS WT.& MEAS` แล้วต่อบล็อก:

```
                      10 PALLETS                 6,405.000 KGS.
                                                    13.450 CBM
<marks (ซ้าย)>        <description (กลาง)>
                      " CARGO MOVEMENT = 7 "      ← (ถ้ามี) รหัส CARGO MOVEMENT
CONSIGNEE :           <ชื่อผู้รับ หลายบรรทัดได้>
Notify Party :        <ชื่อ notify>              ← (ถ้ามี)
CONT NO.  :           <เลขตู้>/<ขนาด>
B/L  NO.  :           <เลข B/L>          A       ← ตัวอักษรต่อท้าย B/L แยกอยู่ทางขวา
```

ยอดรวมท้ายเอกสาร `TOTAL  n PACKAGES  n KGS.  n CBM` ใช้เทียบกับ MANIFEST

## MANIFEST แบบ PDF

รับไฟล์ PDF "CARGO MANIFEST" (พิมพ์แนวนอน) ได้เหมือน .xls — แต่ PDF นี้ **ไม่มีช่อง
SHED NO. / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID** ให้ตรวจ — B/L ที่เป็น TRANSIT จะขึ้น
**สีส้ม** "ไม่มีใน MANIFEST PDF" พร้อมบอกค่าที่ต้องเป็น (เช่น `(MM)`, `7-TRANSIT`) ให้ตรวจในระบบเอง

## ตรวจอะไรบ้าง

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS), TOTAL PACKAGE, PACKAGING
- GROSS WEIGHT, MEASUREMENT
- MARKS, DESCRIPTION — เทียบทีละคู่ ไม่นับช่องว่าง/วรรคตอน แต่คำที่เกิน/ขาดต้องเตือน (สีส้ม)
- REEFER TEMP, DG (CLASS/UN), TRANSIT/TRANSHIPMENT — พบฝั่งใดก็เตือนเสมอ
- SHED NO., ประเทศปลายทาง, CARGO MOVEMENT, TAX ID เทียบ NOTIFY — **ถูกต้องตามกฎ = แสดง `-`**
  (แสดงค่าจริงเฉพาะตอนไม่ตรง/ต้องตรวจ)
- ข้อความสีแดงที่ลูกค้าพิมพ์ทับ PDF (FreeText) — ใช้กับ B/L ที่อยู่ใต้ข้อความ, `ทุก BL` = ใช้กับทุก B/L
- ยอดรวมทั้งฉบับเทียบกับ MANIFEST GRAND TOTAL และ B/L ที่มีใน ENTER แต่ตกหล่นจาก MANIFEST (วิกฤต)

สี: แดง ⚠ = ไม่ตรง · ส้ม = ต้องให้คนยืนยัน · เหลือง = ไม่มี ENTER เทียบ · เขียว = ผ่าน (คอลัมน์สรุป)

## ใช้งานแบบโปรแกรม Python (ในเครื่อง)

`build_edi.py` อ่าน ENTER ได้หลายฟอร์ม (Able, MAC-NELS, Penanshin, CMS / THE M CORP, AGN, Quality, Cargoport, DSV ฯลฯ)
และ MANIFEST ได้ทั้ง .xls/.xlsx และ PDF

```bash
pip install -r requirements.txt
python build_edi.py --manifest MANIFEST.pdf --enter ENTER.pdf --outdir output
```

เพิ่ม `--shed 0141` ถ้ามี SHED NO. ที่แจ้ง · ถ้า MANIFEST เป็นของทั้งลำเรือ แต่ ENTER เป็นของลูกค้ารายเดียว
ให้เพิ่ม `--only-enter`

## วิธีเปิดเว็บให้เพื่อนใช้ (GitHub Pages)

1. ขึ้น GitHub: https://github.com/Rattanao/check-edi-Able → **Settings → Pages**
2. **Build and deployment → Source:** *Deploy from a branch*
3. **Branch:** `main` และโฟลเดอร์ `/ (root)` → กด **Save**
4. รอ ~1 นาที ลิงก์จะขึ้นที่หัวหน้า Pages: `https://rattanao.github.io/check-edi-Able/`
   (หน้าแรกของ repo จะเด้งไป `docs/` เอง)
5. ส่งลิงก์ให้เพื่อนได้เลย — แก้โค้ดแล้ว `git push` หน้าเว็บอัปเดตเองภายในไม่กี่นาที

## หมายเหตุเรื่องข้อมูลลูกค้า

`input/`, `output/` และไฟล์ `.xls/.xlsx/.pdf` ถูก git-ignore ไว้ — repo นี้เก็บเฉพาะโปรแกรม
ไม่เก็บเอกสารจริงของลูกค้า

---
Project by Ms. Rattana Mudlae — Position Import DOC.
