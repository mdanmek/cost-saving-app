'''Cost Saving — IC processor page. Runs in the browser with stlite (Streamlit on Pyodide).

Every uploaded file stays in the memory of this browser tab: the Python code runs inside the page
and nothing is sent to a server. The logic is in costsaving/webui.py (tested by pytest); this file
only draws the widgets.

Modes (D56):
    Sourcing   template with data package (from SharePoint) → IC_Check, Summary → 4 files
    ทีมข้อมูล    master workbook + PO CSVs → template with data package, to publish on SharePoint
    PD         (next round) processed IC files + SAP reports → upload files for Source List, Info Record, FOC
'''
import streamlit as st

from costsaving import webui
from costsaving.config import load_config

st.set_page_config(page_title='Cost Saving — IC processor', layout='wide')

MODES = ['Sourcing: ประมวลผล template IC', 'ทีมข้อมูล: สร้าง template + แพ็กเกจ', 'PD: setup (ยังไม่พร้อม)']

with st.sidebar:
    st.header('โหมด')
    mode = st.radio('เลือกโหมด', MODES, label_visibility='collapsed')

st.title('Cost Saving — ประมวลผลไฟล์ IC')
st.caption(
    'ไฟล์ทุกไฟล์ประมวลผลในเบราว์เซอร์นี้เท่านั้น ไม่ถูกส่งออกไปที่ใด '
    'ปิดแท็บแล้วข้อมูลหายทั้งหมด (โค้ด Python รันในหน้าเว็บด้วย Pyodide)'
)


# ---------------------------------------------------------------- Sourcing
def sourcing():
    with st.sidebar:
        st.header('ตั้งค่า')
        erp_system_id = st.text_input(
            'ERP System ID (สำหรับไฟล์ Contract ของ Ariba)',
            help='ค่านี้เป็นของระบบองค์กร จึงไม่เก็บไว้ในโค้ด ถ้าว่าง ไฟล์ for_Contract จะมีคอลัมน์นี้ว่างและ IC_Check จะเตือน',
        )
        sample_choice = st.radio(
            'ข้อมูลตัวอย่าง (สมมติ)',
            ['ไม่ใช้ (อัปโหลดไฟล์เอง)', 'ตัวอย่างไม่มี Error', 'ตัวอย่างที่มี Error'],
            help='template + แพ็กเกจที่สร้างจากข้อมูลสมมติ ใช้ลองหน้าเว็บโดยไม่ต้องมีไฟล์จริง',
        )
        st.download_button('ดาวน์โหลด template ตัวอย่าง (ว่าง)', data=webui.sample_blank_template(),
                           file_name='IC_Template_sample.xlsx', mime=webui.XLSX_MIME, key='dl_sample_template',
                           help='รูปแบบเดียวกับ template ที่ทีมข้อมูลวางใน SharePoint แต่ข้อมูลเป็นของสมมติ')
    use_sample = sample_choice != 'ไม่ใช้ (อัปโหลดไฟล์เอง)'
    sample_with_errors = sample_choice == 'ตัวอย่างที่มี Error'

    st.subheader('1. อัปโหลด template ที่คีย์แล้ว')
    ic_up = st.file_uploader(
        'template IC (.xlsx) — ดาวน์โหลดรุ่นล่าสุดจาก SharePoint คีย์ Header, Lines (และ CT_Leadtime ถ้าเป็น CT) แล้วอัปโหลดทั้งไฟล์',
        type=['xlsx', 'xlsm'], disabled=use_sample)
    with st.expander('ขั้นสูง: ไฟล์ IC แบบเดิม หรือแทนข้อมูลในแพ็กเกจ'):
        st.caption('ใช้เมื่อเทียบกับไฟล์ IC เดิม (ไม่มีแพ็กเกจ) หรือทีมข้อมูลลองแพ็กเกจรุ่นใหม่ก่อนประกาศ '
                   'ถ้าอัปโหลดที่นี่ จะใช้แทนตารางในแพ็กเกจของ template')
        c1, c2 = st.columns(2)
        with c1:
            master_up = st.file_uploader('ไฟล์ master (.xlsx) — dimUOM_Raw, dimPlant_Raw, dimPlantGroup, dimMasterInfo_Raw',
                                         type=['xlsx', 'xlsm'], disabled=use_sample)
        with c2:
            po_up = st.file_uploader('CSV PO line (prc_spending_sap_{year}.csv) เลือกได้หลายไฟล์', type=['csv'],
                                     accept_multiple_files=True, disabled=use_sample)

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
        with st.spinner('กำลังประมวลผล... (template ที่มี spending ทั้งปีอาจใช้เวลาประมาณหนึ่งนาที)'):
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
            st.error(f'มี Error {run.errors} รายการ แก้ใน template แล้วอัปโหลดใหม่ ระหว่างนี้ดาวน์โหลดไฟล์ output ไม่ได้')
        st.dataframe(check, hide_index=True)

    st.subheader('3. Summary')
    st.dataframe(webui.summary_for_display(run.result['Summary']), hide_index=True, height=560)
    st.caption('Summary ราย Plant')
    st.dataframe(run.result['Summary_ByPlant'], hide_index=True)
    with st.expander('Summary_Lines (รายบรรทัด × Plant)'):
        st.dataframe(run.result['Summary_Lines'], hide_index=True)

    st.subheader('4. ดาวน์โหลด')
    st.caption('**IC Processed** คือไฟล์เดียวที่ยื่นใน SharePoint (IC Request) มี Summary, ตาราง Collect สำหรับไฟล์ '
               'Cost Saving และ Out_PD สำหรับ PD อยู่ในตัว · **for Contract** ใช้ import เข้า Ariba หลัง approve')
    can_download = run.can_download and not stale
    if not run.can_download:
        st.caption('ปุ่มดาวน์โหลดจะใช้ได้เมื่อ IC_Check ไม่มี Error')
    elif stale:
        st.caption('ปุ่มดาวน์โหลดจะใช้ได้เมื่อกด "ประมวลผล" กับไฟล์ชุดปัจจุบัน')
    cols = st.columns(len(run.files))
    for col, (name, data) in zip(cols, run.files.items()):
        col.download_button(webui.file_label(name), data=data, file_name=name, mime=webui.XLSX_MIME,
                            disabled=not can_download, key=f'dl_{name}', help=name)


# ---------------------------------------------------------------- data team
def data_team():
    cfg = load_config()
    st.subheader('สร้าง template + แพ็กเกจข้อมูล (ทีมข้อมูล)')
    st.caption('ได้ไฟล์ template ว่างหนึ่งไฟล์ต่อ commodity ที่มี master และ spending อยู่ในตัว '
               'นำไปวางใน SharePoint ให้ Sourcing ดาวน์โหลด ทำใหม่ทุกรอบอัปเดตแพ็กเกจ')
    use_sample = st.checkbox('ใช้ข้อมูลตัวอย่าง (สมมติ)')
    commodity = st.selectbox('Commodity', list(cfg.dim_commodity['Commodity']))
    c1, c2 = st.columns(2)
    with c1:
        master_up = st.file_uploader('ไฟล์ master (.xlsx) — dimMasterInfo_Raw หรือ MasterInfo_ActiveTH, dimUOM_Raw, '
                                     'dimPlant_Raw, dimPlantGroup', type=['xlsx', 'xlsm'], disabled=use_sample)
    with c2:
        po_up = st.file_uploader('CSV PO line (prc_spending_sap_{year}.csv) เลือกได้หลายไฟล์ เช่นปีนี้และปีก่อน',
                                 type=['csv'], accept_multiple_files=True, disabled=use_sample)
    custom = st.checkbox('กำหนดช่วง spending เอง',
                         help=f"ค่าเริ่มต้น: ถึง PO Date ล่าสุดในไฟล์ ย้อนหลัง {int(cfg.const_number('Package', 'Spending Months'))} เดือน")
    spend_from = spend_to = None
    if custom:
        d1, d2 = st.columns(2)
        spend_from = d1.date_input('Spending From')
        spend_to = d2.date_input('Spending To')
    ready = use_sample or (master_up is not None and po_up)
    if st.button('สร้าง template', type='primary', disabled=not ready):
        with st.spinner('กำลังสร้าง template... (spending ทั้งปีอาจใช้เวลาหลายสิบวินาที)'):
            st.session_state['build'] = webui.run_build(
                None if use_sample else master_up.getvalue(),
                [] if use_sample else [f.getvalue() for f in po_up],
                commodity, spend_from, spend_to, sample_data=use_sample)
    with st.expander('ทดสอบความเร็วด้วยข้อมูลสมมติขนาดจริง'):
        st.caption('สร้าง template ที่มี PO line สมมติตามจำนวนที่เลือก แล้วประมวลผลเหมือน Sourcing เพื่อดูว่าเครื่องนี้ใช้เวลาเท่าไร '
                   'เกณฑ์ที่ตั้งไว้: ประมวลผลไม่เกิน 2 นาที')
        n_rows = st.select_slider('จำนวน PO line', options=[10_000, 40_000, 80_000, 160_000, 240_000], value=160_000)
        if st.button('เริ่มทดสอบ'):
            bar = st.progress(0.0)
            with st.spinner('กำลังทดสอบ... อาจใช้เวลาหลายนาที'):
                timings, size_mb, n, problem = webui.benchmark(n_rows, progress=bar.progress)
            st.session_state['bench'] = (timings, size_mb, n, problem)
        if 'bench' in st.session_state:
            timings, size_mb, n, problem = st.session_state['bench']
            if problem:
                st.error(problem)
            st.dataframe([{'ขั้น': k, 'วินาที': round(v, 1)} for k, v in timings] +
                         [{'ขั้น': f'ขนาด template ({n:,} PO lines)', 'วินาที': None}], hide_index=True)
            st.caption(f'ขนาดไฟล์ {size_mb:.1f} MB')

    build = st.session_state.get('build')
    if build is None:
        st.stop()
    if build.problem:
        st.error(build.problem)
        st.stop()
    st.success(f"สร้าง {build.file_name} แล้ว")
    st.dataframe([{'Field': k, 'Value': str(v)} for k, v in build.info.items()], hide_index=True)
    st.download_button(f'ดาวน์โหลด {build.file_name}', data=build.data, file_name=build.file_name,
                       mime=webui.XLSX_MIME, key='dl_template')


# ---------------------------------------------------------------- PD
def pd_mode():
    st.subheader('PD: setup')
    st.info('โหมดนี้อยู่ในรอบถัดไป: อัปโหลดไฟล์ IC Processed ที่ approve แล้ว + รายงาน SAP (ZMMP1004, ZMMR1019, ME1L) '
            'แล้วได้ไฟล์ upload Source List, Info Record และ FOC')


if mode == MODES[0]:
    sourcing()
elif mode == MODES[1]:
    data_team()
else:
    pd_mode()
