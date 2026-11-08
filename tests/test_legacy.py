'''The current IC file (Price_Change_Template) keyed as Sourcing does it today: noCT_Input, or CT_Vendor +
CT_Leadtime + refreshed CT_Input. Both paths must give the same numbers as the template (fictitious data).'''
from datetime import date
import io

import pandas as pd
import pytest

from costsaving import sample, webui
from costsaving.pipeline import process_ic

KEY_DATE = date(2026, 11, 6)  # same Key Date as sample.HEADER
# line 10 uses LPP Hospital Manual, which the current file does not have
LINES = [x for x in sample.LINES if x['line'] not in webui.SAMPLE_ERROR_LINES | {10}]
CT_LINES = [dict(x, vendor=sample.V1) for x in LINES]  # CT = one vendor per file
COLS = ['Line No', 'Spend Plant', 'Baseline Price Base Unit', 'Baseline Source', 'Final Price Base Unit',
        'Spending Amount', 'Hospital Saving', 'Avoid']


def legacy(mode='noCT', lines=LINES, key_date=KEY_DATE, **kw):
    return process_ic(sample.legacy_workbook_bytes(mode, lines, **kw), po_files=[io.BytesIO(sample.po_csv_bytes())],
                      key_date=key_date)


def summary_lines(result):
    key = ['Line No', 'Spend Plant']
    return result['Summary_Lines'][COLS].sort_values(key, na_position='first').reset_index(drop=True)


def errors(result):
    c = result['IC_Check']
    return list(c.loc[c['Level'] == 'Error', 'Message'])


def test_noct_input_gives_the_same_numbers_as_the_template():
    old, files = legacy('noCT')
    new, _ = process_ic(sample.template_bytes(lines=LINES))
    assert errors(old) == []
    pd.testing.assert_frame_equal(summary_lines(old), summary_lines(new), check_dtype=False)
    assert len(old['Out_Collect']) == len(new['Out_Collect'])
    assert any(k.endswith('_IC_Processed.xlsx') for k in files)


def test_first_price_is_not_used():
    '''The current file keys First Price; the Baseline is Min LPP (D12), so First Price changes nothing.'''
    a = summary_lines(legacy('noCT')[0])
    rows = sample.legacy_workbook_bytes('noCT', LINES)
    assert rows  # First Price is 999 in every fixture row
    assert (a['Baseline Price Base Unit'].dropna() < 999).all()


def test_ct_path_is_processed_like_noct():
    '''CT: vendor form + refreshed CT_Input + typed Method / PD / Plant Group → same source as noCT_Input.'''
    ct, _ = legacy('CT', CT_LINES, method='Yearly Contract')
    noct, _ = legacy('noCT', CT_LINES, method='Yearly Contract')
    assert errors(ct) == []
    a, b = summary_lines(ct), summary_lines(noct)
    # CT has no Nego Price (Nego = Final in Input_CT), so only Avoid may differ
    pd.testing.assert_frame_equal(a.drop(columns='Avoid'), b.drop(columns='Avoid'), check_dtype=False)
    assert (a['Avoid'] == 0).all()


def test_ct_carries_registration_country_and_leadtime():
    result, _ = legacy('CT', CT_LINES, method='Yearly Contract')
    out = result['Out_PD']
    assert set(out['Registration No.']) == {'REG-TEST'} and set(out['Country of Manufacture']) == {'Thailand'}
    lt = result['Out_PD_Leadtime']
    assert len(lt) == len(sample.LEADTIME) and set(lt['Vendor ID']) == {sample.V1}


def test_ct_input_not_refreshed_is_an_error():
    result, _ = legacy('CT', CT_LINES, method='Yearly Contract', stale_ct=True)
    assert any(m.startswith('L02') for m in errors(result))


def test_ct_rows_without_method_and_pd_are_skipped_with_a_warning():
    lines = [dict(x, pd=None) if x['line'] == 2 else x for x in CT_LINES]
    wb = sample.legacy_workbook_bytes('CT', lines, method='Yearly Contract')
    # blank the Method of that row too, as in a CT_Input row Sourcing did not classify
    from openpyxl import load_workbook
    book = load_workbook(io.BytesIO(wb))
    ws = book['CT_Input']
    header = [c.value for c in ws[5]]
    for row in ws.iter_rows(min_row=6):
        if row[header.index('Material Code')].value == 'TST-MED-0002':
            row[header.index('Method')].value = None
    buf = io.BytesIO()
    book.save(buf)
    result, _ = process_ic(buf.getvalue(), po_files=[io.BytesIO(sample.po_csv_bytes())], key_date=KEY_DATE)
    c = result['IC_Check']
    assert c['Message'].str.startswith('L03').any() and errors(result) == []
    assert 'TST-MED-0002' not in set(result['Out_PD']['Material Code'])


def test_several_methods_in_one_file_is_an_error():
    lines = [dict(x) for x in LINES]
    data = sample.legacy_workbook_bytes('noCT', lines)
    from openpyxl import load_workbook
    book = load_workbook(io.BytesIO(data))
    ws = book['noCT_Input']
    header = [c.value for c in ws[5]]
    ws.cell(6, header.index('Effective Date') + 1).value = date(2027, 1, 1)
    buf = io.BytesIO()
    book.save(buf)
    result, _ = process_ic(buf.getvalue(), po_files=[io.BytesIO(sample.po_csv_bytes())], key_date=KEY_DATE)
    assert any(m.startswith('L01') for m in errors(result))


def test_current_file_needs_a_key_date():
    result, _ = legacy('noCT', key_date=None)
    assert any(m.startswith('H06') for m in errors(result))


def test_web_page_passes_the_key_date():
    r = webui.run_uploads(sample.legacy_workbook_bytes('noCT', LINES), [sample.po_csv_bytes()], key_date=KEY_DATE)
    assert r.problem is None and r.errors == 0 and r.can_download


@pytest.mark.parametrize('mode', ['noCT', 'CT'])
def test_processed_file_checksum_round_trips(mode):
    '''The PD mode re-computes the checksum from IC Processed: reading the file back must give the same value.'''
    from costsaving.pipeline import input_checksum
    from costsaving.ic import ICInputs
    from costsaving.readers import open_workbook, read_header, read_table
    result, files = legacy(mode, CT_LINES if mode == 'CT' else LINES, method='Yearly Contract' if mode == 'CT' else 'Renegotiation')
    data = next(v for k, v in files.items() if k.endswith('_IC_Processed.xlsx'))
    wb = open_workbook(data)
    stamp = dict(zip(read_table(wb, 'Processed')['Field'], read_table(wb, 'Processed')['Value']))
    back = ICInputs(header=read_header(wb), lines=read_table(wb, 'Lines'), masterinfo=None, uom=None, plant=None,
                    plant_group=read_table(wb, 'dimPlantGroup', required=False), po_lines=None,
                    leadtime=read_table(wb, 'CT_Leadtime', required=False),
                    package={'Package Version': stamp['Package Version']})
    assert input_checksum(back) == stamp['Input Checksum (SHA-256)'] == input_checksum(result['inputs'])
