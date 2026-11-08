'''PD setup mode: golden test against the Expected sheet of pd_test_workbook.xlsx (fictitious data), join grain,
file checks, and the hand-over from the Sourcing mode.'''
import io

import pandas as pd
import pytest
from openpyxl import load_workbook

from costsaving import sample, sample_pd, webui
from costsaving.config import load_config
from costsaving.setup import (SetupFile, conversion_fraction, read_processed, read_sap_report, run_setup, sap_date,
                              vendor_key)


def vendor_tax_bytes():
    buf = io.BytesIO()
    sample_pd.vendor_tax_df().to_csv(buf, index=False)
    return buf.getvalue()


def run_round(files=None, reports=None, tax_plants=sample_pd.TAX_PLANT):
    f, r = sample_pd.round_inputs()
    return webui.run_setup_uploads(files if files is not None else f, reports if reports is not None else r,
                                   vendor_tax_bytes(), tax_plants, stamp='TEST')


@pytest.fixture(scope='module')
def golden():
    r = run_round()
    assert r.problem is None
    return r


def rows(df, **where):
    out = df
    for k, v in where.items():
        out = out[out[k] == v]
    return out


# ---------------------------------------------------------------- Expected sheet of pd_test_workbook.xlsx
def test_row_counts_match_expected_sheet(golden):
    res = golden.result
    counts = {k: len(res[k]) for k in ('Cont. Add No Plant', 'Cont. Add with Plant', 'Cont. Delete with plant',
                                       'Cont. Delete No plant', 'Load-SL', 'ME01 (manual)', 'Template Info')}
    assert counts == {'Cont. Add No Plant': 4, 'Cont. Add with Plant': 1, 'Cont. Delete with plant': 1,
                      'Cont. Delete No plant': 0, 'Load-SL': 9, 'ME01 (manual)': 1, 'Template Info': 15}
    assert golden.errors == 0 and golden.can_download


def test_foc_add(golden):
    add = golden.result['Cont. Add No Plant']
    got = sorted(zip(add['Material'], add['Scale no.'], add['Minimum qty'], add['Add. quantity'], add['Contract No.']))
    assert got == [('TST-MED-0001', 1, 10, 1, '8199000001'), ('TST-MED-0002', 1, 10, 1, '8199000001'),
                   ('TST-MED-0002', 2, 20, 3, '8199000001'), ('TST-MSP-0001', 1, 5, 1, '8199000002')]
    assert (add['FreeGdsQty'] == add['Minimum qty']).all() and (add['Base Unit'] == add['Order Unit']).all()
    assert set(add['Valid from']) == {'01.12.2026'} and set(add['CalculType']) == {'2'} and set(add['Option']) == {'A'}
    plant = golden.result['Cont. Add with Plant']
    assert list(zip(plant['Material'], plant['Plant'], plant['Minimum qty'], plant['Add. quantity'])) == [('TST-MED-0004', 'T102', 10, 2)]


def test_foc_delete_only_when_cancelled(golden):
    d = golden.result['Cont. Delete with plant']
    assert list(zip(d['Material'], d['Plant'], d['Option'], d['Valid from'])) == [('TST-MED-0003', 'T101', 'D', '01.04.2026')]
    base = golden.result['Upload_FOC_Base']
    changed = base[base.get('Delete Reason') == 'Changed']
    assert list(changed['Material']) == ['TST-MED-0004']  # tier changed on the same contract: not deleted (D50)


def test_source_list(golden):
    sl = golden.result['Load-SL']
    got = sorted(zip(sl['Material'], sl['Plant'], sl['Vendor'], sl['Fixed vendor']))
    assert got == [
        ('TST-MED-0001', 'T101', 'T900000001', 'YES'), ('TST-MED-0001', 'T102', 'T900000001', 'YES'),
        ('TST-MED-0002', 'T101', 'T900000001', 'NO'), ('TST-MED-0002', 'T102', 'T900000001', 'NO'),
        ('TST-MED-0003', 'T101', 'T900000001', 'NO'), ('TST-MED-0004', 'T102', 'T900000001', 'NO'),
        ('TST-MSP-0001', 'T101', 'T900000002', 'YES'), ('TST-MSP-0001', 'T102', 'T900000002', 'YES'),
        ('TST-MSP-0001', 'T701', 'T900000002', 'YES'),
    ]
    assert set(sl['Blocked']) == {'NO'} and set(sl['Purchase Org.']) == {'1000'} and set(sl['Materials Planning(MRP)']) == {'1'}
    assert dict(zip(zip(sl['Material'], sl['Plant']), sl['Agreement Item']))[('TST-MED-0003', 'T101')] == '3'
    # Vendor C (Replace: dropped), Vendor G (blocked), MSP-0002 (no contract item), MED-0005 (no contract) are not loaded
    assert not {'T900000003', 'T900000007', 'T900000006'} & set(sl['Vendor'])
    me01 = golden.result['ME01 (manual)']
    assert list(zip(me01['Material Code'], me01['Plant'], me01['Vendor ID'])) == [('TST-MSP-0001', 'T102', 'T900000003')]


def test_kept_vendor_waits_with_its_set(golden):
    base = golden.result['Upload_SourceList_Base']
    kept = base[(base['Row Source'] == 'Existing (kept)') & (base['Material'] == 'TST-MED-0005')]
    assert list(kept['Vendor']) == ['T900000006'] and kept['Issue'].str.startswith('รอแถวอื่น').all()


def test_info_record(golden):
    ir = golden.result['Template Info']
    key = dict(zip(zip(ir['Material Code'], ir['Plant']), ir.to_dict('records')))
    assert set(ir['Telephone'].dropna()) <= set(sample_pd.IC_FILES) and ir['Telephone'].notna().all()
    assert key[('TST-MED-0003', 'T101')]['Numerator for Conversion'] == 25
    assert key[('TST-MED-0003', 'T101')]['Denominator for Conversion of Order Unit to Base Unit'] == 10
    assert key[('TST-MSP-0003', 'T701')]['Info Record Category'] == '2' and key[('TST-MSP-0003', 'T701')]['Tax code'] == 'V7'
    assert key[('TST-MSP-0004', 'T102')]['Info Record Category'] == '2' and key[('TST-MSP-0004', 'T102')]['Tax code'] == 'VY'
    assert key[('TST-MED-0001', 'T101')]['Info Record Category'] == '0' and pd.isna(key[('TST-MED-0001', 'T101')]['Tax code'])
    assert {key[k]['Purchasing Group'] for k in key if k[0].startswith('TST-MED')} == {'1A0'}
    assert {key[k]['Purchasing Group'] for k in key if k[0].startswith('TST-MSP')} == {'1B0'}
    assert key[('TST-MSP-0001', 'T101')]['Pl. Deliv. Time'] == 7  # current value in ME1L
    assert key[('TST-MSP-0001', 'T102')]['Pl. Deliv. Time'] == 6 and key[('TST-MED-0001', 'T101')]['Pl. Deliv. Time'] == 4
    assert key[('TST-MED-0001', 'T101')]['Standard Qty (per order unit)'] == 1
    assert pd.isna(key[('TST-MSP-0001', 'T101')]['Standard Qty (per order unit)'])  # P25: Note
    assert key[('TST-MED-0001', 'T101')]['Net Price'] == 100 and key[('TST-MED-0001', 'T101')]['Valid From'] == '01.12.2026'
    assert not any(k[0] == 'TST-MED-0006' for k in key)  # No Setup (Fail) gets nothing


def test_check_messages(golden):
    c = golden.result['Upload_Check']
    issues = set(zip(c.loc[c['Type'] == 'Issue', 'Material Code'], c.loc[c['Type'] == 'Issue', 'Message'].str[:12]))
    assert ('TST-MED-0005', 'ยังไม่มีเลข Contract'[:12]) in issues
    msgs = ' | '.join(c['Message'])
    for text in ('E04 FOC ผิดรูปแบบ', 'E09 Item Category', 'ไม่ลบ: FOC เปลี่ยน', 'ไม่พบ Contract Item ใน ME3L',
                 'ได้เลข Contract แต่ไม่พบ Contract Item', 'ต้องลบ Vendor เดิมเองที่ ME01'):
        assert text in msgs, text
    assert 'TST-MED-0006' not in set(c['Material Code'])


# ---------------------------------------------------------------- grain: SAP reports never add rows
def test_duplicate_report_rows_do_not_add_upload_rows():
    files, reports = sample_pd.round_inputs()
    cfg = load_config()
    loaded = {k: read_sap_report(v, k, cfg)[0] for k, v in reports.items()}
    base = run_setup([read_processed(n, d) for n, d, *_ in files], loaded, cfg, tax_plants=[sample_pd.TAX_PLANT])
    doubled = {k: pd.concat([v, v], ignore_index=True) for k, v in loaded.items()}
    f2 = [read_processed(n, d) for n, d, *_ in files]
    for f, (_, _, agr, contract) in zip(f2, files):
        f.agreement_id = agr
    again = run_setup(f2, doubled, cfg, tax_plants=[sample_pd.TAX_PLANT])
    assert len(again['setup_lines']) == len(base['setup_lines']) == sum(len(f.out_pd) for f in f2)
    for sheet in ('Template Info', 'Cont. Add No Plant', 'Cont. Add with Plant', 'Cont. Delete with plant'):
        assert len(again[sheet]) == len(base[sheet]), sheet
    assert len(again['Load-SL']) == len(base['Load-SL']) and len(again['ME01 (manual)']) == len(base['ME01 (manual)'])


def test_me3l_ignores_purchase_orders_and_reads_after_title_rows():
    cfg = load_config()
    df, problem = read_sap_report(sample_pd.zmmp1004_bytes(), 'ZMMP1004', cfg)  # header in row 6
    assert problem is None and len(df) == 5
    me3l, _ = read_sap_report(sample_pd.me3l_bytes(), 'ME3L', cfg)
    assert set(me3l['Doc Type']) == {'ZARB', 'ZFRE'}
    res = run_round().result
    lines = res['setup_lines']
    assert set(lines.loc[lines['Material Code'] == 'TST-MSP-0002', 'Contract Match Type']) == {'Contract only'}


def test_report_with_other_headers_is_a_readable_problem():
    r = run_round(reports={'ZMMP1004': sample_pd.me3l_bytes(), 'ZMMR1019': None, 'ME1L': None, 'ME3L': None})
    assert r.problem and 'ZMMP1004' in r.problem and 'sap_reports.csv' in r.problem


# ---------------------------------------------------------------- file checks (Error = no download)
def test_edited_file_is_an_error_and_blocks_download():
    files, reports, vt, tp = webui.sample_setup_inputs(with_error=True)
    r = webui.run_setup_uploads(files, reports, vendor_tax_bytes(), tp, stamp='TEST')
    assert r.errors >= 1 and not r.can_download
    assert r.result['Upload_Check']['Message'].str.startswith('Input Checksum ไม่ตรง').any()


def test_file_that_is_not_processed_is_an_error():
    files, reports = sample_pd.round_inputs()
    files.append(('old_IC.xlsx', sample.ic_workbook_bytes(), 'X-1', ''))
    r = run_round(files=files)
    assert r.errors >= 1 and r.result['Upload_Check']['Message'].str.startswith('ไม่ใช่ไฟล์ IC Processed').any()


def test_missing_source_list_report_is_an_error_others_are_warnings():
    files, reports = sample_pd.round_inputs()
    r = run_round(reports=dict(reports, ZMMP1004=None, ME1L=None))
    c = r.result['Upload_Check']
    assert set(c.loc[c['Upload'].str.startswith('ZMMP1004'), 'Level']) == {'Error'}
    assert set(c.loc[c['Upload'].str.startswith('ME1L'), 'Level']) == {'Warning'}
    assert not r.can_download


def test_typed_contract_number_restricts_me3l():
    files, reports = sample_pd.round_inputs()
    files = [(n, d, a, '8199999999' if a == 'MED-20261201-001' else c) for n, d, a, c in files]
    lines = run_round(files=files).result['setup_lines']
    med = lines[lines['Agreement ID'] == 'MED-20261201-001']
    assert set(med['Contract Match Type']) == {'Contract only'} and set(med['SAP Contract No.']) == {'8199999999'}


def test_contract_found_in_me3l_without_typing():
    files, reports = sample_pd.round_inputs()
    files = [(n, d, a, '') for n, d, a, c in files]  # PD types nothing (Q7: ME3L first)
    lines = run_round(files=files).result['setup_lines']
    med = lines[(lines['Agreement ID'] == 'MED-20261201-001')]
    assert set(med['Contract Match Type']) == {'Contract item'} and set(med['SAP Contract No.']) == {'8199000001'}


def test_same_info_record_in_two_files_keeps_the_latest_effective_date():
    cfg = load_config()
    files, reports = sample_pd.round_inputs()
    loaded = {k: read_sap_report(v, k, cfg)[0] for k, v in reports.items()}
    a = read_processed('a.xlsx', sample_pd.processed_bytes('MED-20261201-002'))
    b = read_processed('b.xlsx', sample_pd.processed_bytes('MED-20261201-002'))
    a.agreement_id, b.agreement_id = 'A-1', 'B-1'
    b.out_pd['Effective Date'] = pd.Timestamp('2027-01-01')
    res = run_setup([a, b], loaded, cfg)
    ir = res['Template Info']
    assert list(ir['Telephone']) == ['B-1']
    assert res['Upload_Check']['Message'].str.contains('Effective Date ล่าสุด').any()


def test_inactive_and_quit_go_to_the_manual_list():
    cfg = load_config()
    f = read_processed('x.xlsx', sample_pd.processed_bytes('MED-20261201-002'))
    f.out_pd['PD Type'], f.out_pd['PD Action'] = 'No UOM', 'Inactive UOM'
    res = run_setup([f], {'ZMMP1004': pd.DataFrame(columns=['Material Code', 'Plant', 'Vendor ID', 'Valid from', 'Valid to',
                                                             'Purchase Org.', 'Order Unit', 'Agreement', 'Agreement Item',
                                                             'MRP', 'Blocked', 'Fixed vendor'])}, cfg)
    assert len(res['Upload_PD_Manual']) == 1 and res['Upload_PD_Manual']['Task'].iloc[0].startswith('ทำเอง: Inactive UOM')
    assert res['Template Info'].empty and res['Load-SL'].empty


# ---------------------------------------------------------------- Sourcing mode → PD mode
def test_ic_processed_from_the_sourcing_mode_is_read_by_the_pd_mode():
    from costsaving.pipeline import process_ic
    clean = [x for x in sample.LINES if x['line'] not in webui.SAMPLE_ERROR_LINES]
    result, files = process_ic(sample.template_bytes(lines=clean))
    name, data = next((k, v) for k, v in files.items() if k.endswith('_IC_Processed.xlsx'))
    f = read_processed(name, data)
    assert f.checksum_ok and f.problems == []
    assert len(f.out_pd) == len(result['Out_PD']) and set(f.plants['Plant']) == {'T101', 'T102', 'T103'}
    r = webui.run_setup_uploads([(name, data, 'MED-20261106-001', '')], {'ZMMP1004': sample_pd.zmmp1004_bytes()}, stamp='T')
    assert r.problem is None and r.errors == 0
    lines = r.result['setup_lines']
    assert len(lines) == len(result['Out_PD'])  # setup lines keep the Out_PD grain
    # All Plant lines of Update / Add are split into the 3 plants of dimPlant_Raw for the Info Record
    ir = r.result['Upload_InfoRecord_Base']
    one = ir[ir['Material Code'] == 'TST-MED-0001']
    assert sorted(one['Plant']) == ['T101', 'T102', 'T103']


# ---------------------------------------------------------------- helpers
def test_helpers():
    assert vendor_key('0900000123 TEST NAME') == vendor_key('900000123') == vendor_key('900000123_TEST') == '900000123'
    assert conversion_fraction(10) == (10, 1, 0) and conversion_fraction(2.5) == (25, 10, 1)
    assert sap_date('2026-12-01') == '01.12.2026' and sap_date('01.12.2026') == '01.12.2026'


def test_upload_files_layout(golden):
    wb = load_workbook(io.BytesIO(golden.files['Upload_FOC_TEST.xlsx']))
    assert wb.sheetnames == ['Cont. Add with Plant', 'Cont. Add No Plant', 'Cont. Delete with plant', 'Cont. Delete No plant']
    ws = wb['Cont. Add No Plant']
    assert [c.value for c in ws[1]][:3] == ['Option', 'Contract No.', 'Material'] and ws.max_row == 5
    sl = load_workbook(io.BytesIO(golden.files['Upload_SourceList_TEST.xlsx']))
    assert sl.sheetnames == ['Load-SL', 'ME01 (manual)'] and sl['Load-SL']['B2'].number_format == '@'
    ir = load_workbook(io.BytesIO(golden.files['Upload_InfoRecord_TEST.xlsx']))['Template Info']
    assert [c.value for c in ir[1]][-1] == 'Client' and ir.max_row == 16


# ---------------------------------------------------------------- review of #7 (2026-11-08)
def loaded_reports(extra_me3l=()):
    cfg = load_config()
    _, reports = sample_pd.round_inputs()
    loaded = {k: read_sap_report(v, k, cfg)[0] for k, v in reports.items()}
    if extra_me3l:
        loaded['ME3L'] = pd.concat([loaded['ME3L'], pd.DataFrame(list(extra_me3l), columns=loaded['ME3L'].columns)],
                                   ignore_index=True)
    return cfg, loaded


def me3l_item(no, item, vendor, mat, uom, plant=None):
    from datetime import datetime
    return {'Doc Type': 'ZARB', 'Contract No.': no, 'Contract Item': item, 'Valid from': datetime(2026, 12, 1),
            'Valid to': datetime(2027, 11, 30), 'Plant': plant, 'Material Code': mat, 'Vendor ID': vendor,
            'Order Unit': uom, 'Deletion Indicator': None}


def test_by_plant_line_without_plant_never_goes_to_all_plants():
    cfg, loaded = loaded_reports()
    f = read_processed('c.xlsx', sample_pd.processed_bytes('MED-20261201-001'))
    f.agreement_id, f.contract_numbers = 'C-1', ['8199000001']
    f.out_pd.loc[f.out_pd['Line No'] == 4, 'Plant'] = None  # By Plant line, FOC 10F2
    res = run_setup([f], loaded, cfg)
    assert res['Template Info']['Plant'].notna().all()
    assert 'TST-MED-0004' not in set(res['Cont. Add No Plant']['Material'])
    assert res['Upload_Check']['Message'].str.startswith('By Plant แต่ไม่มีรหัส Plant').any()


def test_only_one_fixed_vendor_per_material_and_plant():
    cfg, loaded = loaded_reports([me3l_item('8199000002', '2', 'T900000009', 'TST-MSP-0001', 'BOX')])
    f = read_processed('b.xlsx', sample_pd.processed_bytes('MSP-20261201-001'))
    f.agreement_id, f.contract_numbers = 'B-1', ['8199000002']
    second = f.out_pd.iloc[[0]].copy()
    second['Vendor ID'], second['Line No'] = 'T900000009', 3  # second Bidding winner: fixed by Fix Vendor Methods
    f.out_pd = pd.concat([f.out_pd, second], ignore_index=True)
    res = run_setup([f], loaded, cfg)
    sl = res['Load-SL']
    assert sl[sl['Material'] == 'TST-MSP-0001'].empty  # whole set waits
    assert res['Upload_Check']['Message'].str.contains('หลาย Vendor ที่ต้อง fix').any()


def test_one_source_list_row_per_vendor():
    cfg, loaded = loaded_reports([me3l_item('8199000001', '5', 'T900000001', 'TST-MED-0003', 'BOX', 'T101')])
    f = read_processed('d.xlsx', sample_pd.processed_bytes('MED-20261201-001'))
    f.agreement_id, f.contract_numbers = 'D-1', ['8199000001']
    other_unit = f.out_pd.iloc[[2]].copy()  # line 3 (TST-MED-0003, T101) again with another Order Unit
    other_unit['Order Unit'], other_unit['Line No'] = 'BOX', 9
    f.out_pd = pd.concat([f.out_pd, other_unit], ignore_index=True)
    res = run_setup([f], loaded, cfg)
    sl = res['Load-SL']
    assert sl[(sl['Material'] == 'TST-MED-0003') & (sl['Plant'] == 'T101')].empty
    assert res['Upload_Check']['Message'].str.contains('Vendor เดียวกันมีหลายแถว').any()


def test_processed_file_without_lines_is_an_error_of_that_file_only():
    from openpyxl import load_workbook as lw
    data = sample_pd.processed_bytes('MED-20261201-002')
    wb = lw(io.BytesIO(data))
    del wb['Lines']
    buf = io.BytesIO()
    wb.save(buf)
    files, reports = sample_pd.round_inputs()
    files.append(('no_lines.xlsx', buf.getvalue(), 'X-9', ''))
    r = run_round(files=files)
    assert r.problem is None and r.errors == 1 and not r.can_download
    c = r.result['Upload_Check']
    assert c.loc[c['Level'] == 'Error', 'IC File'].tolist() == ['no_lines.xlsx']
    assert len(r.result['Template Info']) == 15  # the other files are still processed
