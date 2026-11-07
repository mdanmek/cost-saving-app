'''Template with data package (D56): build (data team) → key → process (Sourcing).'''
from datetime import date
import io

import pandas as pd
import pytest
from openpyxl import load_workbook

from costsaving import sample, webui
from costsaving.config import load_config
from costsaving.pipeline import input_checksum, process_ic
from costsaving.readers import normalize_po_lines
from costsaving.template import PACKAGE_PO_COLUMNS, build_template, read_template

CLEAN = [x for x in sample.LINES if x['line'] not in webui.SAMPLE_ERROR_LINES]


def measures(result):
    s = result['Summary']
    return dict(zip(s['Measure'], s['Value']))


def stamp(files):
    data = next(v for k, v in files.items() if k.endswith('_IC_Processed.xlsx'))
    ws = load_workbook(io.BytesIO(data))['Processed']
    return {r[0]: r[1] for r in ws.iter_rows(min_row=2, values_only=True)}


# ---------------------------------------------------------------- same numbers as the earlier format
def test_template_gives_the_same_numbers_as_ic_file_plus_csv():
    '''Once the Error lines are fixed, the package path and the IC file + CSV path must agree line by line.'''
    new, _ = process_ic(sample.template_bytes(lines=CLEAN))
    old, _ = process_ic(io.BytesIO(sample.ic_workbook_bytes(lines=CLEAN)), po_files=[io.BytesIO(sample.po_csv_bytes())])
    cols = ['Line No', 'Spend Plant', 'Baseline Price Base Unit', 'Baseline Source', 'Spending Amount', 'Hospital Saving', 'Avoid']
    key = ['Line No', 'Spend Plant']
    a = new['Summary_Lines'][cols].sort_values(key, na_position='first').reset_index(drop=True)
    b = old['Summary_Lines'][cols].sort_values(key, na_position='first').reset_index(drop=True)
    pd.testing.assert_frame_equal(a, b, check_dtype=False)
    for m in ('Hospital Saving (Annualize)', 'Spending 12 months (matched)', 'Avoid (positive only, not in Cost Saving, Q3)'):
        assert measures(new)[m] == pytest.approx(measures(old)[m])


def test_error_lines_are_still_found_with_a_one_commodity_package():
    '''The package holds Medicine spending only; dimMaterial still tells that line 7 is Medical Supply (E10).'''
    result, _ = process_ic(sample.template_bytes())
    codes = {(m[:3], ln) for m, ln in zip(result['IC_Check']['Message'], result['IC_Check']['Line No'])}
    assert {('E10', 7), ('E02', 8), ('E04', 9)} <= codes


# ---------------------------------------------------------------- the template file
def test_blank_template_layout():
    wb = load_workbook(io.BytesIO(sample.template_bytes(blank=True)))
    assert wb.sheetnames[:5] == ['README', 'Header', 'Lines', 'CT_Leadtime', 'dimPlantGroup']
    assert {'Package_Info', 'dimMasterInfo_Raw', 'dimUOM_Raw', 'dimPlant_Raw', 'fact_po_lines', 'dimMaterial',
            'dimPlantGroup_Official', 'lists'} <= set(wb.sheetnames)
    assert wb['lists'].sheet_state == 'hidden'
    for name in ('Header', 'Lines', 'CT_Leadtime', 'dimPlantGroup', 'fact_po_lines'):
        assert name in wb[name].tables, name
        assert wb[name]['A1'].value is not None  # every table starts at A1 (streaming read)
    header = {r[0]: r[1] for r in wb['Header'].iter_rows(min_row=2, values_only=True)}
    assert header['Commodity'] == 'Medicine' and header['Method'] is None
    dropdowns = {str(dv.sqref): dv.formula1 for dv in wb['Header'].data_validations.dataValidation}
    assert dropdowns['B2'].startswith('lists!') and dropdowns['B4'].startswith('lists!')
    assert [c.value for c in wb['fact_po_lines'][1]] == PACKAGE_PO_COLUMNS


def test_package_holds_one_commodity_inside_the_window():
    wb = load_workbook(io.BytesIO(sample.template_bytes(blank=True)))
    info = {r[0]: r[1] for r in wb['Package_Info'].iter_rows(min_row=2, values_only=True)}
    assert info['Package Version'] == 'MED-20261106' and info['Commodity'] == 'Medicine'
    po = list(wb['fact_po_lines'].iter_rows(min_row=2, values_only=True))
    # 11 sample PO lines − 1 Medical Supply − 1 PO of 2025-10-01 (window From = To − 13 months = 2025-10-06)
    assert len(po) == info['Spending Rows'] == 9
    assert all(not str(r[3]).startswith('TST-MSP') for r in po)


def test_streaming_read_keeps_only_materials_of_lines():
    t = read_template(sample.template_bytes(lines=[dict(line=1, mat='TST-MED-0001', uom='BX100T', final=100, pd='Update')]))
    assert set(t['po_lines']['Material Code']) == {'TST-MED-0001'}
    assert t['po_lines']['Plant Name'].notna().all()  # filled from dimPlant_Raw


# ---------------------------------------------------------------- package checks
def test_old_package_warns_p01():
    header = [(f, date(2027, 3, 1) if f == 'Key Date' else v) for f, v in sample.HEADER]
    result, _ = process_ic(sample.template_bytes(header=header, lines=CLEAN))
    assert result['IC_Check']['Message'].str.startswith('P01').any()


def test_short_package_warns_p02():
    result, _ = process_ic(sample.template_bytes(lines=CLEAN, spend_from=date(2026, 6, 1), spend_to=date(2026, 11, 6)))
    assert result['IC_Check']['Message'].str.startswith('P02').any()


def test_commodity_must_match_package_p03():
    header = [(f, 'Medical Supply' if f == 'Commodity' else v) for f, v in sample.HEADER]
    result, _ = process_ic(sample.template_bytes(header=header, lines=CLEAN))
    assert result['IC_Check'].loc[result['IC_Check']['Message'].str.startswith('P03'), 'Level'].eq('Error').all()


# ---------------------------------------------------------------- Plant Group keyed by Sourcing (D56)
def test_official_group_changed_is_an_error_p04():
    '''Sourcing may add groups but not change the data team's groups (D57).'''
    keyed = pd.DataFrame(sample.PLANT_GROUP, columns=['Plant Group', 'Plant List', 'Count'])
    name = keyed.loc[0, 'Plant Group']
    changed = keyed.copy()
    changed.loc[0, 'Plant List'] = 'TH1'
    added = pd.concat([keyed, pd.DataFrame([('MY GROUP', 'T101', 1)], columns=keyed.columns)], ignore_index=True)
    reordered = keyed.copy()
    reordered.loc[0, 'Plant List'] = ', '.join(reversed([p.strip() for p in str(keyed.loc[0, 'Plant List']).split(',')]))

    def p04(plant_group):
        inputs = process_ic(sample.template_bytes(lines=CLEAN))[0]['inputs']
        inputs.plant_group = plant_group
        from costsaving.ic import changed_official_groups
        return changed_official_groups(inputs)
    assert p04(changed) == [name]
    assert p04(added) == [] and p04(reordered) == []


def test_official_group_check_runs_in_the_pipeline(monkeypatch):
    m = sample.master_frames()
    official = m['plant_group']
    keyed = official.copy()
    keyed.loc[0, 'Plant List'] = 'TH1'
    data, _ = build_template(load_config(), m['masterinfo'], m['uom'], m['plant'], official,
                             normalize_po_lines(sample.po_df()), 'Medicine', header=sample.HEADER, lines=sample.lines_df(CLEAN))
    from costsaving import template as tpl
    real = tpl.read_template

    def edited(source):
        t = real(source)
        t['plant_group'] = keyed
        return t
    monkeypatch.setattr('costsaving.pipeline.read_template', edited)
    result, _ = process_ic(data)
    p04 = result['IC_Check'][result['IC_Check']['Message'].str.startswith('P04')]
    assert len(p04) == 1 and p04['Level'].iloc[0] == 'Error'


def test_plant_list_accepts_codes_or_names():
    def run(plant_list):
        m = sample.master_frames()
        m['plant_group'] = pd.DataFrame([('MINE', plant_list, 2)], columns=['Plant Group', 'Plant List', 'Count'])
        data, _ = build_template(load_config(), m['masterinfo'], m['uom'], m['plant'], m['plant_group'],
                                 normalize_po_lines(sample.po_df()), 'Medicine', header=sample.HEADER,
                                 lines=sample.lines_df([dict(line=1, mat='TST-MED-0001', uom='BX100T', final=100, pg='MINE', pd='Update')]))
        result, _ = process_ic(data)
        return sorted(result['agreement_line']['Plant']), result
    by_code, r = run('T101, T102')
    by_name, _ = run('TH1,TH2')
    assert by_code == by_name == ['T101', 'T102']
    assert not r['IC_Check']['Message'].str.startswith('E08').any()


# ---------------------------------------------------------------- processed file (approve → PD)
def test_processed_file_has_inputs_out_pd_and_checksum():
    result, files = process_ic(sample.template_bytes(lines=CLEAN))
    data = next(v for k, v in files.items() if k.endswith('_IC_Processed.xlsx'))
    wb = load_workbook(io.BytesIO(data))
    assert {'Processed', 'Summary', 'Header', 'Lines', 'forSavingCollect', 'Out_PD', 'Out_PD_Leadtime', 'Out_Contract',
            'IC_Check', 'dimPlantGroup'} <= set(wb.sheetnames)
    assert 'fact_po_lines' not in wb.sheetnames  # the package is not copied
    s = stamp(files)
    assert s['Package Version'] == 'MED-20261106' and len(s['Input Checksum (SHA-256)']) == 64
    assert set(result['Out_Collect']['Package Version']) == {'MED-20261106'}


def test_checksum_changes_only_when_keyed_data_changes():
    a = process_ic(sample.template_bytes(lines=CLEAN))[0]['inputs']
    b = process_ic(sample.template_bytes(lines=CLEAN))[0]['inputs']
    changed = [dict(x, final=x['final'] + 1) if x['line'] == 1 else x for x in CLEAN]
    c = process_ic(sample.template_bytes(lines=changed))[0]['inputs']
    assert input_checksum(a) == input_checksum(b) != input_checksum(c)


# ---------------------------------------------------------------- data team mode
def test_run_build_sample_and_missing_inputs():
    r = webui.run_build(None, [], 'Medicine', sample_data=True)
    assert r.problem is None and r.file_name == 'IC_Template_MED-20261106.xlsx' and r.info['Spending Rows'] == 9
    assert webui.run_build(None, [], 'Medicine').problem == 'ยังไม่ได้อัปโหลดไฟล์ master'


def test_run_build_from_uploaded_files_matches_sample():
    r = webui.run_build(sample.master_workbook_bytes(), [sample.po_csv_bytes()], 'Medicine')
    assert r.problem is None and r.info['Spending Rows'] == 9
    result, _ = process_ic(r.data)  # blank template: processes without crashing, Header errors only
    assert result['IC_Check']['Message'].str.startswith('H02').any()


def test_template_with_csv_override_marks_the_package():
    result, _ = process_ic(sample.template_bytes(lines=CLEAN), po_files=[io.BytesIO(sample.po_csv_bytes())])
    assert measures(result)['Package Version'] == 'MED-20261106 + CSV'
