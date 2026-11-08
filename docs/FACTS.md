# FACTS — decision ที่โค้ดต้องทำตาม และสถานะในโค้ด

สรุป ณ 2026-11-08 (ขั้นที่ 0 ของรอบ "Sourcing: gap fixes" + "PD mode") · ผู้สรุป: ไห่อัน
แหล่ง: `PROJECT_CONTEXT.md` ส่วน 5 (D1–D57), `design/*.md`, `current-state/powerquery/*.m` และ `drafts/` ใน repository ส่วนตัวของโครงการ เทียบกับโค้ดใน repo นี้ที่ `main` (commit `1e672c4`)
ไฟล์นี้ไม่มีข้อมูลจริง ตัวอย่างทุกตัวเป็นค่าสมมติ

สถานะ: **ทำแล้ว** = มีโค้ดและ test · **บางส่วน** = มีโค้ดแต่ขาดบางข้อ · **ยังไม่ทำ** · **ขัดกัน** = โค้ดหรือเอกสารขัดกับ decision ล่าสุด

---

## 1. โหมด Sourcing

| Decision | สิ่งที่โค้ดต้องทำ | ไฟล์ / ฟังก์ชัน | สถานะ |
|---|---|---|---|
| D11, D14, P4 | LPP = PO ล่าสุด **ก่อน** Key Date รวมทั้งใบ (ของแถมรวมในใบ) ย้อนหลัง 12 เดือน | `ic.po_window`, `ic.lpp_tables` | ทำแล้ว (`test_lpp_latest_po_whole_po_strictly_before_key_date`) |
| D12, D13 | Min LPP ทุก Method · All Plant = ต่ำสุดของ PO ล่าสุดแต่ละ Plant · By Plant = ของ Plant นั้น | `ic.build_agreement_lines` | ทำแล้ว |
| D15 | Baseline กรอกเองได้ (LPP Hospital Manual) และ Reference Material Code | `ic.build_agreement_lines` | ทำแล้ว (`test_reference_material_code_and_manual`) |
| D6 | Growth ว่าง = 5% (`constants.csv`) | `ic.build_agreement_lines` | ทำแล้ว |
| D3, D53 | Pool Purchase = All Vendor ใช้ Transfer Price · Parallel = By Vendor (`dim_method.csv`) | `ic.summary_lines` | ทำแล้ว (`test_pool_purchase_uses_transfer_price`) |
| D53 | ราคา Base Unit ปัด 2 ตำแหน่ง แบบ half away from zero | `config.round_half_up`, `prices.base_price` | ทำแล้ว (`test_round_half_up_two_decimals`) |
| หลักการ (Base Unit) | เทียบราคาและคูณปริมาณที่ Base Unit · LPP หาจาก Material + Base Unit + Plant · spend 12 เดือนเป็นปริมาณ Base Unit | `ic.lpp_tables`, `ic.summary_lines` | ทำแล้ว |
| D47, D52 | ไม่มี Fail · PD Action เลือกรายบรรทัดตาม Strategy → Method → PD Action (dim_combination) | `config/dim_combination.csv`, `ic.build_check` E13 | ทำแล้ว (dropdown ใน template แสดง PD Action ทุกตัว ไม่กรองตาม Method — E13 จับแทน) |
| D48 | ไม่มี Item Category ในไฟล์ IC | `ic.build_agreement_lines` ใส่ Standard ภายใน ไม่ส่งใน Out_PD | ทำแล้ว |
| D49 | หนึ่งไฟล์หนึ่ง Commodity | E10 (`dimMaterial` ของแพ็กเกจ), P03 | ทำแล้ว |
| D44, D45 (S4), P14 | Out_Contract ตามไฟล์ Ariba: Commodity A/B, Plant ว่างเมื่อ All Plant, Item Text = Trade Name ตัวใหญ่, Price = Final ระดับ Order Unit ก่อนหัก FOC, Quantity 999999, Tax Code VY · ไม่มีแถวเมื่อ PD Action Needs Contract = FALSE | `ic.contract_base`, `ic.out_contract`, `writers.write_for_contract` | ทำแล้ว |
| D51 | master รูปแบบเดิม (dimMasterInfo_Raw, dimUOM_Raw, dimPlant_Raw, dimPlantGroup) Base Unit มาจาก Order Unit | `readers`, `ic.prepare_masters` | ทำแล้ว |
| D56 | template ไฟล์เดียวมีแพ็กเกจในตัว อ่านแบบ streaming | `template.build_template`, `template.read_template` | ทำแล้ว |
| D56 | CT_Leadtime ใช้เฉพาะ Method ที่ `is CT` ใน config | `ic.is_ct`, `ic.out_pd_leadtime`, W09 / W10 | ทำแล้ว |
| D56, D57 Q4 | Sourcing เพิ่มกลุ่ม Plant ได้ แต่แก้กลุ่มของทีมข้อมูลไม่ได้ (P04 Error) | `ic.changed_official_groups`, `build_check` | ทำแล้ว |
| D57 Q1 | **IC Processed ไฟล์เดียว** มี Summary, IC_Check, สิ่งที่คีย์, ตาราง `Input_forSavingCollect`, Out_PD, Out_Contract · ดาวน์โหลดเหลือ IC Processed + for Contract | `pipeline.output_files`, `writers.write_processed` | ทำแล้วในโค้ด · **README และ docstring ของ `streamlit_app.py` ยังเขียน "4 ไฟล์" (ล้าสมัย)** |
| D57 Q1, AGENTS | ตาราง Collect: ชื่อและลำดับ 29 คอลัมน์เดิม ต่อท้ายได้เท่านั้น | `ic.OUT_COLLECT_COLUMNS` | ทำแล้ว: 29 คอลัมน์แรกตรงกับ `Input_forSavingCollect` ของไฟล์ IC เดิม และครอบ 28 คอลัมน์ที่ไฟล์ Project expand (`func_project`) · **ค่าในบางคอลัมน์ต่างจากไฟล์เดิม ดูข้อ 3.1** |
| D57 Q3 | Package Version แสดงใน Summary · ประมวลผลซ้ำใช้แพ็กเกจล่าสุด | `ic.summary`, `pipeline.processed_stamp` | ทำแล้ว |
| D57 Q5 | เตือนเมื่อ spending ในแพ็กเกจเก่ากว่า Key Date เกิน 45 วัน | `build_check` P01 (`constants.csv` Package / Max Age Days) | ทำแล้ว |
| D57 Q6 | ประมวลผลไม่เกิน 2 นาที | `webui.benchmark` | ทำแล้ว (วัดบน Pyodide 17.3 วินาที / 160,009 แถว สมมติ) |
| D57 Q7 | ทางเดิม (ไฟล์ IC + master + CSV) เก็บไว้ช่วงเปลี่ยนผ่าน | `pipeline.load_ic_inputs`, ส่วน "ขั้นสูง" ของหน้าเว็บ | ทำแล้ว |
| P16 (D57) | Input Checksum (SHA-256) ของสิ่งที่คีย์ ใน IC Processed | `pipeline.input_checksum` | ทำแล้ว · **ยังไม่มี test ว่าอ่านกลับจากไฟล์ IC Processed แล้วได้ค่าเดิม** (ลองมือกับข้อมูลตัวอย่างแล้วตรง แต่โหมด PD ต้องพึ่ง จึงต้องมี test) |
| template-spec ข้อ 7 | กฎ E01–E13, H01–H07, W01–W08 | `ic.build_check` | บางส่วน: **W04 (Base Unit ของ Reference Material Code ไม่ตรงกับรายการใหม่) ยังไม่ทำ** · E09 (consignment มี FOC) อยู่ฝั่ง PD เพราะไฟล์ IC ไม่รู้ consignment (D48) |
| template-spec ข้อ 8 | Summary มี 10 รายการ Saving สูงสุด / ต่ำสุด | — | ยังไม่ทำ (P13 เปลี่ยน Summary เป็นราย Plant แบบเดิม จึงไม่ถือเป็นช่องว่าง เว้นแต่ 瑞安 ต้องการ) |

### ข้อที่ต้องตัดสินก่อนแก้ (เปลี่ยนตัวเลขปลายทาง)

**3.1 ค่าในตาราง Collect ต่างจากไฟล์ IC เดิม** (`current-state/powerquery/ic_price_change_template.m` query `Input_forSavingCollect`)

| เรื่อง | ไฟล์ IC เดิม | แอปตอนนี้ | ผลต่อไฟล์ปลายทางเดิม |
|---|---|---|---|
| grain ของ FOC | แตกหนึ่งแถว**ต่อขั้น FOC** และเพิ่มแถวไม่มีของแถม (`1F0`) เช่น `10F1` → 2 แถว (ราคาไม่หัก FOC และราคาหัก 10F1) | หนึ่งแถวต่อ Line × Plant ใช้ขั้นลึกสุด และ Final FOC เป็นข้อความเต็ม เช่น `10F1,20F3` | ไฟล์ dispatch อ่าน `Final Buy` จากข้อความก่อน `F` แล้วตัดแถว PO ที่ Order Quantity < Final Buy ทิ้ง (Check Quantity) ถ้า Collect มีแต่ขั้นลึกสุด PO ที่ซื้อน้อยกว่าขั้นนั้นจะหายจาก Actual |
| Project Name | `[Saving Strategy]Method_Project Name` | Project Name ตามที่คีย์ | ชื่อ project ใน Power BI จะไม่เป็นรูปแบบเดียวกับปีก่อน |

## 2. โหมด PD

แหล่ง logic: `drafts/powerquery/setup_shared.m`, `pd_inputs.m`, `pd_upload_foc.m`, `pd_upload_source_list.m`, `pd_upload_info_record.m` (ร่าง M ยังไม่เคยรัน, D54 ให้เก็บไว้เป็นเอกสารอ้างอิงกติกา) และชีต Expected ของ `pd_test_workbook.xlsx` ที่ `drafts/tools/build_upload_test_workbook.py` สร้าง

| Decision | สิ่งที่โค้ดต้องทำ | ที่มาใน M | สถานะ |
|---|---|---|---|
| D39, D56 | โหมด PD ในหน้าเว็บ: อัปโหลด IC Processed (หนึ่งหรือหลายไฟล์) + รายงาน SAP → Upload_Check → ดาวน์โหลด | `pd_inputs.m` | ยังไม่ทำ (หน้าเว็บมีแค่ข้อความ placeholder) |
| P16 | Input Checksum ไม่ตรง = Error | — (ใหม่) | ยังไม่ทำ |
| D52 | upload เฉพาะ PD Type ใน `Upload PD Types` (Update, Add) ที่เหลือเป็น `Upload_PD_Manual` · No Setup ไม่มีงาน | `listUploadPDTypes`, `Upload_PD_Manual` | ยังไม่ทำ |
| D34, D45 (S1) | Purchase Org 1000 · Source List / Info Record ราย Plant · All Plant แตกเป็นทุก Plant ของ commodity | `fnExpandAllPlant`, `dimPlant_SetupScope` | ยังไม่ทำ · **รายชื่อ Plant มาจากไหนในเว็บ: ดูคำถาม** |
| D26, D35, D40, D43 | Contract No. + Contract Item ของแต่ละบรรทัด (Contract Match Type: Contract item / Contract only / Ambiguous / Not found / No contract / Not needed) | `Upload_Lines`, `Upload_Contract_Match` | ยังไม่ทำ |
| Q16, Q21 (ข้อเสนอ) | เลข Contract: ถ้าไม่ได้กรอก ให้จับจาก ME1L ด้วย Vendor + Material + ช่วงวันที่ ถ้าไม่เจอให้ PD พิมพ์เอง · เป็น config | — (ใหม่) | ยังไม่ทำ |
| D35, D50 (P24) | อ่าน ZMMP1004, ZMMR1019 (หาแถวหัวตารางเอง) และ ME1L | `fnReadSAPReport`, `sap_current_*` | ยังไม่ทำ · **หัวคอลัมน์ของ ME1L ยังไม่รู้ (P24)** |
| D37, D45 (S2, S3), D50 | Source List: ทั้งชุด Vendor ต่อ Material × Plant · Replace / Keep ตาม `Source List Mode` · Fixed vendor YES / NO · Blocked NO · MRP 1 · ME01 manual เมื่อ Source List เดิมมี Vendor เดียวที่ต้องตัด · ชุดไหนมี Issue ห้าม upload ทั้งชุด | `pd_upload_source_list.m` | ยังไม่ทำ (`dim_method.csv` มี `Source List Mode` แล้ว) |
| D37, D50 | Info Record ราย Plant: Category 0 / 2, Vendor Material No. = Setup Material No., Telephone = Agreement ID, Conversion เป็นเศษส่วนจำนวนเต็ม, Pl. Deliv. Time (ยา 4 / เวชภัณฑ์ CT_Leadtime ไม่มีใช้ 6), Purchasing Group จาก Material Type, Standard / Minimum Qty (ยา 1 / 1), Tax code ตาม category, วันที่ `dd.MM.yyyy` | `pd_upload_info_record.m` | ยังไม่ทำ · `dim_commodity.csv` ยังไม่มีคอลัมน์ของ PD |
| D35, D37, D50 | FOC สี่ชีต: Add ต่อขั้น (Scale no. 1, 2, 3), FreeGdsQty = Minimum qty, Base Unit = Order Unit, CalculType / FreeGoodsCat. = 2 · Delete เมื่อยกเลิกของแถม (ค่าจาก ZMMR1019) · เปลี่ยนขั้นไม่ลบ (มากับ Contract ใหม่) · consignment มี FOC = Issue | `pd_upload_foc.m` | ยังไม่ทำ |
| D46, D28 | consignment ไม่มี Contract / Source List / FOC มีแค่ Info Record category 2 | `cfgItemCategory` | ยังไม่ทำ · **ไฟล์ IC ไม่รู้ consignment (D48, Q13 ค) ต้องมีแหล่งในโหมด PD** |
| หลักการ (grain) | ทุก join กับรายงาน SAP ต้องไม่ทำให้แถวเพิ่ม · golden test จากชีต Expected | Checks ท้ายแต่ละไฟล์ .m | ยังไม่ทำ |

grain ที่จะใช้ (จาก M และ data-contract 6.4): Out_PD = Line × Plant scope · Source List upload = Material × Plant × Vendor · Info Record = Supplier × Material × Plant × Category · FOC = Contract × Material × (Plant หรือว่าง) × Order Unit × Valid from × Scale no. · ZMMP1004 = Material × Plant × Vendor × Valid from · ZMMR1019 = Contract × Material × Plant × Order Unit × Valid from × Scale no. · ME1L = Contract × Contract Item

## 3. เอกสารที่ล้าสมัยหรือขัดกัน (ยึด decision ล่าสุด)

| ไฟล์ | จุด | ยึด |
|---|---|---|
| repo นี้ `README.md`, `streamlit_app.py` (docstring), หัวข้อ "ใช้หน้าเว็บ" | ดาวน์โหลด 4 ไฟล์ (for Collect, for Summary แยก) | D57 Q1: 2 ไฟล์ |
| `design/template-spec.md` ข้อ 3, 7 | โครงไฟล์ Power Query (`[Config]` จาก IC_Config, `Out_Collect` / `Summary` เป็นชีตแยก, Prefill) และตารางกฎไม่มี P01–P04, W09, W10 | D54, D56, D57 |
| `design/setup-outputs.md` บรรทัดเป้าหมาย และส่วน 7 | `Out_Upload_*` อยู่ในไฟล์คีย์ · ร่าง M เป็นแผนหลัก | D39 (ไฟล์ PD Setup แยก), D54 (Python) |
| `design/data-contract.md` ส่วน 3 | Template Version `v2.0` | `constants.csv` v2.1 |
| `drafts/powerquery/setup_shared.m` (`cfgItemCategory`) | consignment Needs Contract "รอ Q13 ก" | D46 ปิดแล้ว |
| `ROADMAP.md` ส่วนเครื่องมือ | ยังไม่ระบุ web app Python | D54 |
| `PROJECT_CONTEXT.md` ส่วน 10 | ลำดับรายการไม่เรียงตามเวลา และ "งานถัดไป" หลายชุดซ้อนกัน | ทำใหม่ตอนจบรอบ |
| D47 กับ D52 | D47 "PD เหลือแค่ Update" | D52 (ตีความใหม่ รอ 瑞安 ยืนยัน) โค้ดทำตาม D52 |
