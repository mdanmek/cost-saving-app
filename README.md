# cost-saving-app — IC processor (Python, รันในเบราว์เซอร์)

หน้าเว็บ: **https://mdanmek.github.io/cost-saving-app/**

แพ็กเกจนี้ทำงานแทน Power Query ของไฟล์ IC (Price Change Template) รับไฟล์ input แล้วสร้างไฟล์ output สี่ไฟล์ หน้าเว็บใช้ [stlite](https://github.com/whitphx/stlite) (Streamlit บน Pyodide) โค้ด Python ทั้งหมดรันในแท็บของเบราว์เซอร์ ไม่มี server และไม่เก็บข้อมูล

เอกสารออกแบบ การตัดสินใจ (D54, D55 …) และสถานะของโครงการอยู่ใน repository ส่วนตัวของโครงการ (`PROJECT_CONTEXT.md`) repo นี้เก็บเฉพาะโค้ดและ config ที่เปิดเผยได้ ชื่อ decision ในโค้ดและ test (เช่น D53) อ้างถึงเอกสารนั้น

## ใช้หน้าเว็บ

หน้าเว็บมีสามโหมด เลือกที่แถบด้านซ้าย (decision D56)

```
ทีมข้อมูล   master + CSV PO ─▶ [โหมดทีมข้อมูล] ─▶ template IC + แพ็กเกจ (1 ไฟล์ต่อ commodity) ─▶ SharePoint
Sourcing   ดาวน์โหลด template ─▶ คีย์ ─▶ [โหมด Sourcing] ─▶ IC Processed, for Contract
           IC Processed ─▶ ยื่นใน SharePoint ─▶ manager approve
PD         ดาวน์โหลด IC Processed ─▶ [โหมด PD, รอบถัดไป] ─▶ ไฟล์ upload Source List, Info Record, FOC
```

**Sourcing**

1. ดาวน์โหลด template รุ่นล่าสุดจาก SharePoint คีย์ Header และ Lines (คอลัมน์เรียงแบบ noCT_Input ของไฟล์เดิม ไม่ต้องกรอกคำอธิบายสินค้า ชื่อ Vendor และ MFG) Quit / Inactive ไม่ต้องกรอกราคา ถ้าต้องการกลุ่ม Plant ของตัวเอง ให้เพิ่มแถวใน dimPlantGroup (Plant List ใส่รหัสหรือชื่อ Plant คั่นด้วย `,`)
2. อัปโหลด template ทั้งไฟล์ แล้วกด **ประมวลผล** ครั้งแรกรอประมาณ 20–60 วินาทีให้โหลด Python และ template ที่มี spending ทั้งปีอาจใช้เวลาอีกประมาณหนึ่งนาที
3. ดู **IC_Check** ก่อน ถ้ามีแถว Level = Error ต้องแก้ใน template แล้วอัปโหลดใหม่ ระหว่างนั้นปุ่มดาวน์โหลดจะถูกปิด
4. ดาวน์โหลด 2 ไฟล์ (D57): **IC Processed** ไฟล์เดียวที่ยื่นใน SharePoint มี Summary, IC_Check, ตาราง `Input_forSavingCollect` สำหรับไฟล์ Project และ Out_PD สำหรับ PD · **for Contract** นำเข้า Ariba

ถ้าอยากลองโดยไม่มีไฟล์จริง เลือก **ข้อมูลตัวอย่าง (สมมติ)** ที่แถบด้านข้าง มีชุดที่ไม่มี Error และชุดที่มี Error และดาวน์โหลด template ตัวอย่าง (ว่าง) ไปดูรูปแบบได้

ค่า **ERP System ID** ของ Ariba เป็นค่าของระบบองค์กร จึงไม่อยู่ใน repo ให้พิมพ์ที่แถบด้านข้างทุกครั้ง

ส่วน **ขั้นสูง** ใช้ไฟล์ IC แบบเดิม (Price_Change_Template: noCT_Input หรือ CT_Input ตามที่ Sourcing แก้หลังรีเฟรช อ่านตาม `costsaving/legacy.py` ต้องกรอก Key Date ที่หน้าเว็บ) คู่กับ CSV PO หรือใช้ไฟล์ master / CSV แทนตารางในแพ็กเกจเพื่อลองแพ็กเกจรุ่นใหม่ (Package Version จะต่อท้ายด้วย `+ CSV`)

**ทีมข้อมูล** อัปโหลดไฟล์ master (dimMasterInfo_Raw หรือ MasterInfo_ActiveTH, dimUOM_Raw, dimPlant_Raw, dimPlantGroup) และ CSV PO line เลือก Commodity แล้วกด **สร้าง template** ได้ไฟล์ `IC_Template_{Package Version}.xlsx` นำไปวางใน SharePoint แทนรุ่นเดิม ช่วง spending ค่าเริ่มต้นคือถึง PO Date ล่าสุดในไฟล์ ย้อนหลัง 13 เดือน

### template และแพ็กเกจ

template เป็นไฟล์เดียวเหมือนไฟล์ IC ปัจจุบัน ส่วนที่ Sourcing คีย์และส่วนที่ทีมข้อมูลควบคุมอยู่คนละชีต

| ชีต | ใครกรอก | เนื้อหา |
|---|---|---|
| Header, Lines | Sourcing | การต่อรองหนึ่งครั้ง (มี dropdown จาก config) |
| dimPlantGroup | ทีมข้อมูลตั้งต้น, Sourcing เพิ่มได้ | กลุ่ม Plant |
| Package_Info | ทีมข้อมูล | Package Version, Commodity, ช่วง spending |
| dimMasterInfo_Raw, dimUOM_Raw, dimPlant_Raw | ทีมข้อมูล | master |
| fact_po_lines | ทีมข้อมูล | PO line ของ commodity ในช่วง spending (ใช้หา LPP และ Summary 12 เดือน) |
| dimMaterial | ทีมข้อมูล | Material → Commodity จาก spending ทุก commodity (ใช้ตรวจ E10) |

ทุกตารางเริ่มที่ A1 เพื่อให้แอปอ่านแบบ streaming ได้ แอปเก็บไว้เฉพาะ PO line ของ Material ที่อยู่ใน Lines จึงใช้หน่วยความจำน้อยแม้ fact_po_lines มีเป็นแสนแถว แต่ยังต้อง parse ทุกแถว (ประมาณ 25 วินาทีต่อ 160,000 แถวบน CPython)

IC_Check ของแพ็กเกจ: **P01** spending ในแพ็กเกจเก่ากว่า Key Date เกิน 45 วัน, **P02** spending ไม่ครบ 12 เดือนก่อน Key Date, **P03** Commodity ใน Header ไม่ตรงกับแพ็กเกจ (Error)

ไฟล์ **IC Processed** เก็บสิ่งที่คีย์ (Header, Lines, dimPlantGroup), Out_PD, dimPlant_Raw, IC_Check, ตาราง Collect (หนึ่งแถวต่อขั้น FOC และ Project Name `[Strategy]Method_ชื่อ` แบบไฟล์เดิม) และชีต Processed ที่มี Package Version และ **Input Checksum (SHA-256)** ของสิ่งที่คีย์ โหมด PD จะใช้ checksum นี้ตรวจว่าไฟล์ถูกแก้หลังประมวลผลหรือไม่ ไฟล์นี้ไม่คัดลอกแพ็กเกจไปด้วย จึงเล็ก

### ข้อมูลไปที่ไหน

ไฟล์ที่อัปโหลดอยู่ในหน่วยความจำของแท็บนั้นเท่านั้น ปิดแท็บแล้วหายทั้งหมด หน้าเว็บโหลดแค่ตัวโปรแกรม (stlite, Pyodide, pandas, openpyxl) จาก `cdn.jsdelivr.net` และ PyPI และไม่ส่งข้อมูลกลับไปที่ใด ส่วนนี้มีสามชั้น:

- `site/index.html` ตั้ง Content-Security-Policy ให้หน้าเว็บเชื่อมต่อได้เฉพาะ host ข้างบน ถ้ามีโค้ดพยายามส่งข้อมูลไป host อื่น เบราว์เซอร์จะบล็อก
- ปิด usage stats ของ Streamlit (`browser.gatherUsageStats = false`)
- test `test_app_code_has_no_network_calls` กันไม่ให้โค้ดใน `costsaving/` และ `streamlit_app.py` import ไลบรารีเครือข่าย

## ไหลอย่างไร

```
template IC + แพ็กเกจ      Header, Lines, dimPlantGroup + master + fact_po_lines + dimMaterial
(หรือแบบเดิม: ไฟล์ IC + ไฟล์ master + CSV PO line prc_spending_sap_{year}.csv)
                           อ่านเฉพาะ PO line ของ Material ที่อยู่ใน Lines
          │
          ▼
   agreement_line  (grain: Line × Plant; Plant Group แตกเป็นรายโรงพยาบาล, ว่าง = All Plant)
          │  Baseline = min(LPP Hospital, SDC, Dr. Pharma) หยุดค่าที่ Key Date
          ▼
   *_IC_Processed.xlsx   Summary, IC_Check, Header, Lines, Input_forSavingCollect, Out_PD, dimPlant_Raw,
                         Out_Contract, checksum → SharePoint (approve) → IC Collected (Project) และ PD
   *_for_Contract.xlsx   หนึ่งชีตต่อ Vendor (Ariba)       → Sourcing import
```

## โครงสร้าง

| ไฟล์ | หน้าที่ |
|---|---|
| `costsaving/config/*.csv` | กติกาธุรกิจ: Commodity, Method → Calc Type, dimCombination, dimPD, ค่าคงที่ เพิ่ม Method หรือ Commodity โดยแก้ CSV ไม่ต้องแก้โค้ด |
| `costsaving/config.py` | อ่าน config และฟังก์ชันแปลงค่า (`round_half_up` ปัดแบบ Excel) |
| `costsaving/prices.py` | ราคาต่อ Base Unit หัก FOC ขั้นลึกสุด ปัด 2 ตำแหน่ง (D53) |
| `costsaving/readers.py` | อ่านตาราง Excel ตามชื่อตาราง และ CSV PO line ทีละ chunk |
| `costsaving/ic.py` | กติกาทั้งหมดของไฟล์ IC |
| `costsaving/writers.py` | เขียนไฟล์ output เป็น bytes ให้เบราว์เซอร์ดาวน์โหลด |
| `costsaving/template.py` | สร้าง template + แพ็กเกจ (โหมดทีมข้อมูล) และอ่าน template แบบ streaming (โหมด Sourcing) |
| `costsaving/pipeline.py` | `process_ic(...)` เรียกครั้งเดียวจากไฟล์ input ถึงไฟล์ output, checksum ของไฟล์ IC Processed |
| `costsaving/webui.py` | สิ่งที่หน้าเว็บทำ (รันจากไฟล์ที่อัปโหลด, เรียง IC_Check, กันดาวน์โหลดเมื่อมี Error, zip) แยกจาก Streamlit เพื่อให้ pytest ทดสอบได้ |
| `costsaving/sample.py` | ข้อมูลสมมติ ใช้ทั้งใน test และปุ่มข้อมูลตัวอย่างของหน้าเว็บ |
| `streamlit_app.py` | หน้าเว็บ (วาด widget รอบ `webui`) |
| `site/index.html`, `site/build.py` | หน้าเว็บ stlite และสคริปต์สร้าง `_site/` สำหรับ GitHub Pages |
| `tests/` | test ที่ต้องได้ตัวเลขตรงชีต Expected, test ของหน้าเว็บ และ test ของการ build |
| `.github/workflows/tests.yml` | pytest กับ pandas 2.2 (แบบที่ Pyodide ใช้) และ pandas 3 |
| `.github/workflows/pages.yml` | push เข้า `main` → รัน test → build `_site/` → deploy GitHub Pages |

## รันบนเครื่อง

```bash
pip install -r requirements.txt
python -m pytest -q
```

ลองหน้าเว็บ stlite บนเครื่อง (ต้องต่อเน็ตเพื่อโหลด stlite จาก CDN):

```bash
python site/build.py
python -m http.server 8000 -d _site    # เปิด http://localhost:8000
```

หรือรันเป็น Streamlit ปกติ (ไม่ผ่าน Pyodide): `pip install streamlit` แล้ว `streamlit run streamlit_app.py`

เรียกจาก Python ตรง ๆ:

```python
from costsaving.pipeline import process_ic

result, files = process_ic('IC.xlsx', po_files=['prc_spending_sap_2026.csv'], master_file='master.xlsx')
print(result['IC_Check'])
for name, data in files.items():
    open(name, 'wb').write(data)
```

## เปลี่ยนเวอร์ชัน stlite

เวอร์ชันถูก pin ไว้ใน `site/index.html` (ตอนนี้ `@stlite/browser@1.10.0` ซึ่งใช้ Pyodide 0.29 / Python 3.13) stlite รุ่นใหม่อาจมาพร้อม Pyodide และ pandas รุ่นใหม่ ให้เปลี่ยนทั้งสองบรรทัด (css และ js) พร้อมกัน แล้วลองด้วยข้อมูลตัวอย่างทั้งสองชุดก่อน merge

## กติกาของ repo public

- ห้าม commit ข้อมูลจริงทุกชนิด: ราคา ชื่อผู้ขาย รหัส Vendor เลข PO เลขสัญญา ปริมาณ ชื่อโรงพยาบาลจริง path ของ SharePoint ไฟล์ Excel / CSV จริง (`.gitignore` กัน `*.xlsx` และ `*.csv` นอก `costsaving/config/`)
- test และข้อมูลตัวอย่างใช้ข้อมูลสมมติใน `costsaving/sample.py` เท่านั้น (รหัสขึ้นต้นด้วย `TST-`, Vendor `T9…`, Plant `T1…`)
- ค่าของระบบองค์กร (เช่น ERP System ID) ไม่อยู่ใน config ให้ผู้ใช้กรอกตอนรัน
