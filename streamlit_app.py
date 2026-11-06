'''Cost Saving — IC processor page. Runs in the browser with stlite (Streamlit on Pyodide).

Every uploaded file stays in the memory of this browser tab: the Python code runs inside the page
and nothing is sent to a server. The logic is in costsaving/webui.py (tested by pytest); this file
only draws the widgets.
'''
import streamlit as st

from costsaving import webui

st.set_page_config(page_title='Cost Saving — IC processor', layout='wide')

st.title('Cost Saving — ประมวลผลไฟล์ IC')
st.caption(
    'ไฟล์ทุกไฟล์ประมวลผลในเบราว์เซอร์นี้เท่านั้น ไม่ถูกส่งออกไปที่ใด '
    'ปิดแท็บแล้วข้อมูลหายทั้งหมด (โค้ด Python รันในหน้าเว็บด้วย Pyodide)'
)

with st.sidebar:
    st.header('ตั้งค่า')
    erp_system_id = st.text_input(
        'ERP System ID (สำหรับไฟล์ Contract ของ Ariba)',
        help='ค่านี้เป็นของระบบองค์กร จึงไม่เก็บไว้ในโค้ด ถ้าว่าง ไฟล์ for_Contract จะมีคอลัมน์นี้ว่างและ IC_Check จะเตือน',
    )
    sample_choice = st.radio(
        'ข้อมูลตัวอย่าง (สมมติ)',
        ['ไม่ใช้ (อัปโหลดไฟล์เอง)', 'ตัวอย่างไม่มี Error', 'ตัวอย่างที่มี Error'],
        help='ไฟล์ IC, master และ CSV PO ที่สร้างจากข้อมูลสมมติ ใช้ลองหน้าเว็บโดยไม่ต้องมีไฟล์จริง',
    )
    use_sample = sample_choice != 'ไม่ใช้ (อัปโหลดไฟล์เอง)'
    sample_with_errors = sample_choice == 'ตัวอย่างที่มี Error'

st.subheader('1. อัปโหลดไฟล์')
c1, c2, c3 = st.columns(3)
with c1:
    ic_up = st.file_uploader('ไฟล์ IC (.xlsx) — ตาราง Header, Lines, CT_Leadtime', type=['xlsx', 'xlsm'],
                             disabled=use_sample)
with c2:
    master_up = st.file_uploader(
        'ไฟล์ master (.xlsx, ไม่บังคับ) — dimUOM_Raw, dimPlant_Raw, dimPlantGroup, dimMasterInfo_Raw',
        type=['xlsx', 'xlsm'], disabled=use_sample,
        help='ไม่ต้องใช้ถ้าตารางเหล่านี้อยู่ในไฟล์ IC แล้ว (เช่นไฟล์ IC เดิม)',
    )
with c3:
    po_up = st.file_uploader('CSV PO line (prc_spending_sap_{year}.csv) เลือกได้หลายไฟล์', type=['csv'],
                             accept_multiple_files=True, disabled=use_sample,
                             help='อ่านเฉพาะ Material ที่อยู่ใน Lines ไฟล์ทั้งปีอาจใช้เวลาหลายสิบวินาที')

if use_sample:
    st.info('กำลังใช้ข้อมูลตัวอย่าง (สมมติ) — ช่องอัปโหลดถูกปิดไว้')
    signature = ('sample', sample_choice, erp_system_id)
else:
    signature = (
        None if ic_up is None else (ic_up.name, ic_up.size),
        None if master_up is None else (master_up.name, master_up.size),
        tuple((f.name, f.size) for f in po_up or []),
        erp_system_id,
    )

ready = use_sample or ic_up is not None
if st.button('ประมวลผล', type='primary', disabled=not ready):
    with st.spinner('กำลังประมวลผล...'):
        if use_sample:
            ic_bytes, po_list, master_bytes = webui.sample_inputs(sample_with_errors)
        else:
            ic_bytes = ic_up.getvalue()
            master_bytes = master_up.getvalue() if master_up is not None else None
            po_list = [f.getvalue() for f in po_up or []]
        st.session_state['run'] = webui.run_uploads(ic_bytes, po_list, master_bytes, erp_system_id)
        st.session_state['signature'] = signature

run = st.session_state.get('run')
if run is None:
    st.stop()
stale = st.session_state.get('signature') != signature
if stale:
    st.warning('ไฟล์หรือการตั้งค่าเปลี่ยนไปหลังประมวลผลครั้งล่าสุด กด "ประมวลผล" อีกครั้ง ผลด้านล่างเป็นของรอบก่อน '
               'และปุ่มดาวน์โหลดถูกปิดจนกว่าจะประมวลผลใหม่')
if run.problem:
    st.error(run.problem)
    st.stop()

st.subheader('2. IC_Check')
m1, m2 = st.columns(2)
m1.metric('Error (ต้องแก้ก่อนยื่น)', run.errors)
m2.metric('Warning', run.warnings)
check = webui.sorted_check(run.result['IC_Check'])
if check.empty:
    st.success('ไม่มี Error และ Warning')
else:
    if run.errors:
        st.error(f'มี Error {run.errors} รายการ แก้ในไฟล์ IC แล้วอัปโหลดใหม่ ระหว่างนี้ดาวน์โหลดไฟล์ output ไม่ได้')
    st.dataframe(check, hide_index=True)

st.subheader('3. Summary')
st.dataframe(webui.summary_for_display(run.result['Summary']), hide_index=True, height=560)
st.caption('Summary ราย Plant')
st.dataframe(run.result['Summary_ByPlant'], hide_index=True)
with st.expander('Summary_Lines (รายบรรทัด × Plant)'):
    st.dataframe(run.result['Summary_Lines'], hide_index=True)

st.subheader('4. ดาวน์โหลด')
can_download = run.can_download and not stale
if not run.can_download:
    st.caption('ปุ่มดาวน์โหลดจะใช้ได้เมื่อ IC_Check ไม่มี Error')
elif stale:
    st.caption('ปุ่มดาวน์โหลดจะใช้ได้เมื่อกด "ประมวลผล" กับไฟล์ชุดปัจจุบัน')
cols = st.columns(len(run.files) + 1)
for col, (name, data) in zip(cols, run.files.items()):
    col.download_button('for ' + name.rsplit('_for_', 1)[-1].removesuffix('.xlsx'), data=data, file_name=name,
                        mime=webui.XLSX_MIME, disabled=not can_download, key=f'dl_{name}',
                        help=name)
cols[-1].download_button('ทั้ง 4 ไฟล์ (.zip)', data=webui.zip_files(run.files) if can_download else b'',
                         file_name='IC_outputs.zip', mime='application/zip',
                         disabled=not can_download, key='dl_zip')
