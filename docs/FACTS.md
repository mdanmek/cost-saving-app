# FACTS — decision ที่โค้ดต้องทำตาม และสถานะในโค้ด

สรุป ณ 2026-11-08 · อัปเดตหลังเฟส 1 (ขั้นที่ 0 ของรอบ "Sourcing: gap fixes" + "PD mode") · ผู้สรุป: ไห่อัน
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
| D56 | CT_Leadtime ใช้เฉพาะ Method ที่ `is CT` ใน config | — | แทนด้วย D61: เลิกใช้ CT_Leadtime |
| D56, D57 Q4 | Sourcing เพิ่มกลุ่ม Plant ได้ แต่แก้กลุ่มของทีมข้อมูลไม่ได้ (P04 Error) | `ic.changed_official_groups`, `build_check` | ทำแล้ว |
| D57 Q1 | **IC Processed ไฟล์เดียว** มี Summary, IC_Check, สิ่งที่คีย์, ตาราง `Input_forSavingCollect`, Out_PD, Out_Contract · ดาวน์โหลดเหลือ IC Processed + for Contract | `pipeline.output_files`, `writers.write_processed` | ทำแล้วในโค้ด · **README และ docstring ของ `streamlit_app.py` ยังเขียน "4 ไฟล์" (ล้าสมัย)** |
| D57 Q1, AGENTS | ตาราง Collect: ชื่อและลำดับ 29 คอลัมน์เดิม ต่อท้ายได้เท่านั้น | `ic.OUT_COLLECT_COLUMNS` | ทำแล้ว: 29 คอลัมน์แรกตรงกับ `Input_forSavingCollect` ของไฟล์ IC เดิม และครอบ 28 คอลัมน์ที่ไฟล์ Project expand (`func_project`) · **ค่าในบางคอลัมน์ต่างจากไฟล์เดิม ดูข้อ 3.1** |
| D57 Q3 | Package Version แสดงใน Summary · ประมวลผลซ้ำใช้แพ็กเกจล่าสุด | `ic.summary`, `pipeline.processed_stamp` | ทำแล้ว |
| D57 Q5 | เตือนเมื่อ spending ในแพ็กเกจเก่ากว่า Key Date เกิน 45 วัน | `build_check` P01 (`constants.csv` Package / Max Age Days) | ทำแล้ว |
| D57 Q6 | ประมวลผลไม่เกิน 2 นาที | `webui.benchmark` | ทำแล้ว (วัดบน Pyodide 17.3 วินาที / 160,009 แถว สมมติ) |
| D57 Q7 | ทางเดิม (ไฟล์ IC + master + CSV) เก็บไว้ช่วงเปลี่ยนผ่าน | `pipeline.load_ic_inputs`, `legacy.read_legacy`, ส่วน "ขั้นสูง" ของหน้าเว็บ | ขัดกัน → แก้แล้ว: เดิมอ่านไฟล์ IC จริงไม่ได้ (ดู 1.1) |
| P16 (D57) | Input Checksum (SHA-256) ของสิ่งที่คีย์ ใน IC Processed | `pipeline.input_checksum`, `pipeline.canonical_text` | บางส่วน → แก้แล้ว: อ่านกลับจากไฟล์ IC เดิมได้ค่าไม่ตรง (`test_processed_file_checksum_round_trips`) |
| template-spec ข้อ 7 | กฎ E01–E13, H01–H07, W01–W08 | `ic.build_check` | ทำแล้ว (W04 เพิ่มในเฟส 1) · E09 (consignment มี FOC) อยู่ฝั่ง PD เพราะไฟล์ IC ไม่รู้ consignment (D48) |
| template-spec ข้อ 8 | Summary มี 10 รายการ Saving สูงสุด / ต่ำสุด | — | ยังไม่ทำ (P13 เปลี่ยน Summary เป็นราย Plant แบบเดิม จึงไม่ถือเป็นช่องว่าง เว้นแต่ 瑞安 ต้องการ) |

### 1.1 สิ่งที่ Sourcing คีย์จริงในไฟล์ IC เดิม (ตรวจโครงไฟล์ Price_Change_Template v.20261102 ที่ 瑞安 แนบ 2026-11-08 · บันทึกเฉพาะชื่อชีต ตาราง และคอลัมน์)

`varCT` (ชีต Summary) เลือกทาง ทั้งสองทางลงเอยที่ตารางหัวคอลัมน์ 48 คอลัมน์ชุดเดียวกัน (`Input_noCT` / `Input_CT`) แล้ว query `Input` ประมวลผลเหมือนกัน

| ทาง | Sourcing กรอก | ไม่ได้กรอก (สูตรหรือ query) |
|---|---|---|
| noCT | เซลล์เหนือตาราง: Commodity, Saving Strategy · ตาราง `Input_noCT` 28 คอลัมน์: Method, PD, Plant Group, Material Code, Material Description, Vendor ID, Vendor Name, MFG, Order Unit, Alternative Material Code, Alternative Order Unit, Per for Setup, First Price, First FOC, Nego Price, Nego FOC, Final Price, Final FOC, Effective Date, Valid Date, Project Name, Local or Import, Supplier Product Code, Supplier Product Name, Trade Name, TMT Code or UCEP Code, Quotation Date, Remark | Commodity, Saving Strategy (อ้างเซลล์), Order Unit Text, ISO Code, Conversion, Base Unit, ราคา Base Unit สามตัว, Avoid, PPV, %PPV, Expected, Combination ID, Plant List |
| CT | ชีต `CT_Vendor`: Vendor (`varVendor` = `รหัส_ชื่อ`), Quotation Date เหนือตาราง, ตาราง `CT` (ฟอร์มใบเสนอราคา 16 คอลัมน์ รวมประเทศผู้ผลิตและเลขทะเบียน อย.) · ชีต `CT_Leadtime` · รีเฟรช `CT_Input` แล้วกรอก Commodity, Saving Strategy (เหนือตาราง) และ Method, PD, Plant Group รายแถว | ที่เหลือของ `Input_CT` มาจาก query: Vendor จาก `varVendor`, Effective / Valid Date จาก parameter ของ Power Query, Project Name = `CT{ปี}-{ชื่อ Vendor}`, Nego = Final, First Price จาก MasterInfo, Per for Setup = 1 |

ข้อเท็จจริงเพิ่ม: ประเทศผู้ผลิตและเลขทะเบียน อย. อยู่ในฟอร์ม CT แต่ไม่ถูกส่งเข้า `Input_CT` · แถวของ CT ที่ไม่ได้กรอก Method และ PD ไม่มี Combination ID จึงไม่ถูกนับ · ไฟล์เดิมไม่มี Key Date (Baseline คือ First Price) · `factSpending_Raw` ในไฟล์ไม่มีเลข PO และวันที่ PO จึงหา LPP ไม่ได้ ต้องใช้ CSV PO line · สูตร LET ใช้ขั้น FOC ที่**เขียนสุดท้าย** แอปใช้ขั้นที่ซื้อมากสุด (ตรงกันเมื่อเขียนจากน้อยไปมาก)

**ช่องว่างที่พบ:** ทางเดิม (D57 Q7) ของแอปอ่านเฉพาะตาราง `Header` / `Lines` ซึ่งไฟล์จริงไม่มี จึง**อ่านไฟล์ IC จริงไม่ได้** · แก้แล้วใน `costsaving/legacy.py`: อ่านตามกติกาข้างบน ทั้งสองทางเป็น source รูปเดียวกับคอลัมน์ที่กรอกของ `Input_noCT` แล้วผ่าน `ic.run_ic` ชุดเดียวกับ template (First Price ไม่ใช้, D12) · Key Date ผู้ใช้กรอกที่หน้าเว็บ · กฎใหม่ L00 (varCT ไม่ใช่ CT / noCT), L01 (Method / Effective / Valid / Project Name มีหลายค่าในไฟล์เดียว), L03 (Warning: ข้ามแถว CT ที่ไม่ได้กรอก Method และ PD) · L02 ที่เคยใส่ถูกเอาออก (D61) · test ใน `tests/test_legacy.py`
**ช่องว่างที่สอง:** Input Checksum เปลี่ยนเมื่ออ่านไฟล์ IC Processed กลับ (วันที่กลับมาเป็น datetime, 1 กลับมาเป็น 1.0) แก้แล้วด้วย `pipeline.canonical_text`

### 1.2 ตัดสินแล้วในรอบ 2026-11-08 (grilling Q1–Q18 · บันทึกเป็น D58–D63 ใน PROJECT_CONTEXT) และแก้ในเฟส 1

| เรื่อง | ตัดสิน | ไฟล์ / ฟังก์ชัน | test |
|---|---|---|---|
| D58 (Q1, Q13) template กรอกแบบ noCT_Input | Header ครั้งเดียว + Lines เรียงตามคอลัมน์ที่กรอกของ noCT_Input ไม่มี First Price / FOC (D12) ระบบเติม Material Description, Vendor Name, MFG · Template Version v2.2 | `ic.LINE_COLUMNS`, `template.build_template` | `test_blank_template_layout` |
| D59 (Q2) Quit / Inactive* ไม่ต้องมีราคา | `dim_pd.csv` คอลัมน์ `Needs Price` = FALSE → ไม่ขึ้น E03 ไม่นับ Cost Saving (ถ้าย้ายรหัส / Vendor ให้กรอกราคาของรายการใหม่) | `ic.build_agreement_lines`, `ic.build_check` | `test_quit_and_inactive_need_no_price_and_are_not_collected` |
| D60 (Q4) Collect รูปแบบเดิม | หนึ่งแถวต่อขั้น Final FOC + แถว `1F0` · Nego FOC = ขั้นลึกสุดที่ไม่เกินขั้น Final · Project Name `[Strategy]Method_ชื่อ` · ราคา Base Unit ของแต่ละขั้นปัดตาม D53 | `ic.collect_tiers`, `ic.out_collect` | `test_collect_one_row_per_foc_tier_like_the_current_file` · เทียบชีต Collect ของไฟล์ตัวอย่างจริงสองไฟล์ใน session (ไม่ลง repo): จำนวนแถวและค่าที่เทียบได้ตรงทุกแถว |
| D61 (Q12) เลิก CT_Leadtime | ไม่มีชีต CT_Leadtime, Out_PD_Leadtime, W09 / W10, `is CT` · ไฟล์เดิมทาง CT: CT_Input ตามที่ Sourcing แก้หลังรีเฟรชคือจุดตั้งต้น (เพิ่ม / ลบแถวได้ จึงไม่มี L02) | `legacy.read_legacy` | `test_ct_input_is_the_starting_point_even_when_edited`, `test_no_leadtime_output_and_no_ct_rule` |
| D62 (Q14) Alternative Material Code / Order Unit | = Reference Material Code / Reference Order Unit ใช้หา LPP และ spending ของรหัสเดิม · W04 เตือนเมื่อ Base Unit ต่างจากรหัสเดิม | `ic.build_agreement_lines`, `ic.summary_lines`, `ic.build_check` | `test_w04_reference_material_with_another_base_unit` |
| D57 Q1 / Q3 | ดาวน์โหลด 2 ไฟล์ · IC Processed มี `dimPlant_Raw` ต่อท้าย (Q15 สำหรับโหมด PD) | `writers.write_processed` | `test_processed_file_has_inputs_out_pd_and_checksum` |

## 2. โหมด PD (setup) — ทำในเฟส 2 (`costsaving/setup.py`)

แหล่ง logic: `drafts/powerquery/setup_shared.m`, `pd_inputs.m`, `pd_upload_*.m` (ร่าง M ไม่เคยรัน, D54) และชีต Expected ของ `pd_test_workbook.xlsx` · ข้อมูลสมมติชุดเดียวกันอยู่ใน `costsaving/sample_pd.py`

| Decision | สิ่งที่โค้ดทำ | ไฟล์ / ฟังก์ชัน | test |
|---|---|---|---|
| D39, D56, D63 | หน้าเว็บโหมด PD: IC Processed หลายไฟล์ + ZMMP1004, ZMMR1019, ME1L, ME3L (+ ตาราง Tax ของ Vendor ไม่บังคับ) → Upload_Check → ดาวน์โหลด 4 ไฟล์ ประมวลผลในเบราว์เซอร์ ไม่แก้ CSP | `streamlit_app.pd_mode`, `webui.run_setup_uploads` | `test_page_pd_mode_*` |
| P16, Q5 | Input Checksum ไม่ตรง = Error ห้ามดาวน์โหลด · ไฟล์ที่ไม่ใช่ IC Processed / อัปโหลดซ้ำ / Agreement ID ซ้ำ = Error | `setup.read_processed`, `run_setup` | `test_edited_file_is_an_error_and_blocks_download`, `test_file_that_is_not_processed_is_an_error` |
| D52, Q6 | upload เฉพาะ PD Type ใน `PD / Upload PD Types` (Update, Add) ที่เหลือ (No Sale / No UOM) เป็น `Upload_PD_Manual` · No Setup ไม่มีงาน | `run_setup` | `test_inactive_and_quit_go_to_the_manual_list` |
| D26, D35, Q7, Q10 | Contract No. + Item: จับจาก ME3L (Doc Type ใน `PD / Contract Doc Types`) ด้วย Vendor + Material + Order Unit + Plant (ว่าง = ทุก Plant) + ช่วงวันที่ทับกัน · เลขที่ PD พิมพ์ใช้จำกัดผลลัพธ์ · ไม่พบแต่พิมพ์มาเลขเดียว = Contract only (FOC ไปได้ Source List รอ) | `run_setup` | `test_typed_contract_number_restricts_me3l`, `test_contract_found_in_me3l_without_typing` |
| D34, D45 S1, Q15 | All Plant แตกเป็นทุก Plant ใน `dimPlant_Raw` ของไฟล์ IC Processed (Source List, Info Record) · FOC คงแถวเดียว Plant ว่าง | `run_setup.expand` | `test_ic_processed_from_the_sourcing_mode_is_read_by_the_pd_mode` |
| D37, D45 S2/S3, D50 | Source List ทั้งชุด Vendor ต่อ Material × Plant · Replace / Keep · Fixed vendor YES / NO · Blocked NO · MRP 1 · ME01 manual · ชุดมี Issue รอทั้งชุด | `setup._source_list` | `test_source_list`, `test_kept_vendor_waits_with_its_set` |
| D37, D50, Q11, Q16, Q17 | Info Record ราย Plant: category จาก ME1L (consignment = 2), Purchasing Group / Pl. Deliv. Time / Minimum Qty จาก ME1L ถ้าไม่มีใช้ค่าตั้งต้นของ commodity, Tax code ตาม category + ตาราง Vendor Tax + Plant ที่ PD พิมพ์, คีย์ชนกันจากหลายไฟล์ใช้ Effective Date ล่าสุด (D23) | `setup._info_record` | `test_info_record`, `test_same_info_record_in_two_files_keeps_the_latest_effective_date` |
| D35, D37, D50 | FOC สี่ชีต: Add ต่อขั้น, Delete เมื่อยกเลิก (ค่าจาก ZMMR1019), เปลี่ยนขั้นไม่ลบ, consignment มี FOC = E09 | `setup._foc_add`, `setup._foc_delete` | `test_foc_add`, `test_foc_delete_only_when_cancelled` |
| grain | ทุก join ใช้ key ที่ไม่ซ้ำ: ME3L ตัดซ้ำที่ Contract × Item, ZMMP1004 ที่ Material × Plant × Vendor × Valid from, ME1L ใช้แถวแรก (Note) · รายงานซ้ำสองเท่าได้ผลเท่าเดิม | `run_setup` | `test_duplicate_report_rows_do_not_add_upload_rows` · golden: `test_row_counts_match_expected_sheet` (FOC 4/1/1/0, Source List 9, ME01 1, Info Record 15) |
| P24 | ชื่อหัวคอลัมน์ของรายงาน SAP อยู่ใน `config/sap_reports.csv` แก้ได้ไม่ต้องแก้โค้ด หาแถวหัวตารางเอง | `setup.read_sap_report` | `test_me3l_ignores_purchase_orders_and_reads_after_title_rows` |

รีวิว #7 (2026-11-08): บรรทัด By Plant ที่ไม่มีรหัส Plant ไม่ออกเป็น Info Record / FOC แบบไม่มี Plant · Source List หนึ่ง Vendor fix ต่อ Material × Plant และหนึ่งแถวต่อ Vendor · ไฟล์ที่ขาด Header / Lines เป็น Error ของไฟล์นั้น · ความเร็ว: ME3L ทำ index ตาม Vendor × Material × Order Unit รอบสมมติ 40 ไฟล์ × 100 บรรทัด กับรายงานละ 40,000 แถว จาก 64 เหลือ 10 วินาทีบน CPython (ผลเท่าเดิม)

Upload_Check: **Error** = ทั้งรอบห้ามดาวน์โหลด · **Warning** แยกชนิด `Issue` (แถวนี้ไม่อยู่ในไฟล์ upload), `Manual` (PD ทำเอง), `Note` (แถวไปได้ แต่มีช่องที่ต้องตรวจ)

ยังไม่ได้ทำ / รอ: template ไฟล์ upload สำหรับงาน inactive (Q6 ข), P25 Standard / Minimum Qty ของเวชภัณฑ์, C11 Registration Type / Special Feature, ชื่อหัวคอลัมน์จริงของ Info Record template คอลัมน์ O / P (setup-outputs เขียนแบบย่อ)

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
