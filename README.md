# cost-saving-app — IC processor (Python, รันในเบราว์เซอร์)

หน้าเว็บ: **https://mdanmek.github.io/cost-saving-app/**

แพ็กเกจนี้ทำงานแทน Power Query ของไฟล์ IC (Price Change Template) รับไฟล์ input แล้วสร้างไฟล์ output สี่ไฟล์ หน้าเว็บใช้ [stlite](https://github.com/whitphx/stlite) (Streamlit บน Pyodide) โค้ด Python ทั้งหมดรันในแท็บของเบราว์เซอร์ ไม่มี server และไม่เก็บข้อมูล

เอกสารออกแบบ การตัดสินใจ (D54, D55 …) และสถานะของโครงการอยู่ใน repository ส่วนตัวของโครงการ (`PROJECT_CONTEXT.md`) repo นี้เก็บเฉพาะโค้ดและ config ที่เปิดเผยได้ ชื่อ decision ในโค้ดและ test (เช่น D53) อ้างถึงเอกสารนั้น

## ใช้หน้าเว็บ

1. เปิดหน้าเว็บ ครั้งแรกรอประมาณ 20–60 วินาทีให้โหลด Python, pandas และ openpyxl
2. อัปโหลด **ไฟล์ IC** (.xlsx มีตาราง Header, Lines, CT_Leadtime), **ไฟล์ master** (ไม่บังคับ ถ้าตาราง master อยู่ในไฟล์ IC แล้ว) และ **CSV PO line** `prc_spending_sap_{year}.csv` (เลือกได้หลายไฟล์)
3. กด **ประมวลผล** แล้วดู **IC_Check** ก่อน ถ้ามีแถว Level = Error ต้องแก้ในไฟล์ IC แล้วอัปโหลดใหม่ ระหว่างนั้นปุ่มดาวน์โหลดจะถูกปิด
4. ดู **Summary** และ Summary ราย Plant แล้วดาวน์โหลดไฟล์ทีละไฟล์ หรือทั้ง 4 ไฟล์เป็น .zip

ถ้าอยากลองโดยไม่มีไฟล์จริง เลือก **ข้อมูลตัวอย่าง (สมมติ)** ที่แถบด้านข้าง มีสองชุด คือชุดที่ไม่มี Error (ลองดาวน์โหลดได้) และชุดที่มี Error (ดูว่าระบบกันการดาวน์โหลดอย่างไร)

ค่า **ERP System ID** ของ Ariba เป็นค่าของระบบองค์กร จึงไม่อยู่ใน repo ให้พิมพ์ที่แถบด้านข้างทุกครั้ง ถ้าเว้นว่าง ไฟล์ for_Contract จะมีคอลัมน์นี้ว่างและ IC_Check จะเตือน

### ข้อมูลไปที่ไหน

ไฟล์ที่อัปโหลดอยู่ในหน่วยความจำของแท็บนั้นเท่านั้น ปิดแท็บแล้วหายทั้งหมด หน้าเว็บโหลดแค่ตัวโปรแกรม (stlite, Pyodide, pandas, openpyxl) จาก `cdn.jsdelivr.net` และ PyPI และไม่ส่งข้อมูลกลับไปที่ใด ส่วนนี้มีสามชั้น:

- `site/index.html` ตั้ง Content-Security-Policy ให้หน้าเว็บเชื่อมต่อได้เฉพาะ host ข้างบน ถ้ามีโค้ดพยายามส่งข้อมูลไป host อื่น เบราว์เซอร์จะบล็อก
- ปิด usage stats ของ Streamlit (`browser.gatherUsageStats = false`)
- test `test_app_code_has_no_network_calls` กันไม่ให้โค้ดใน `costsaving/` และ `streamlit_app.py` import ไลบรารีเครือข่าย

## ไหลอย่างไร

```
ไฟล์ IC (.xlsx)            ตาราง Header, Lines, CT_Leadtime
ไฟล์ master (.xlsx)        dimUOM_Raw, dimPlant_Raw, dimPlantGroup, dimMasterInfo_Raw
                           (ไม่ต้องใช้ถ้าตารางเหล่านี้อยู่ในไฟล์ IC แล้ว เช่นไฟล์ IC เดิม)
CSV PO line                prc_spending_sap_{year}.csv (อ่านเฉพาะ Material ที่อยู่ใน Lines)
          │
          ▼
   agreement_line  (grain: Line × Plant; Plant Group แตกเป็นรายโรงพยาบาล, ว่าง = All Plant)
          │  Baseline = min(LPP Hospital, SDC, Dr. Pharma) หยุดค่าที่ Key Date
          ▼
   *_for_PD.xlsx         Out_PD, Out_PD_Leadtime        → ไฟล์ PD Setup
   *_for_Contract.xlsx   หนึ่งชีตต่อ Vendor (Ariba)       → Sourcing import
   *_for_Collect.xlsx    ตาราง Input_forSavingCollect    → ไฟล์ Cost Saving
   *_for_Summary.xlsx    Summary, Summary_ByPlant, Summary_Lines, IC_Check
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
| `costsaving/pipeline.py` | `process_ic(...)` เรียกครั้งเดียวจากไฟล์ input ถึงไฟล์ output |
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
