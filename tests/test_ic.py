'''The IC process must give the numbers of the Expected sheet (IC_Template_v2_test.xlsx).
Each test names the rule (decision number in PROJECT_CONTEXT.md) it protects.'''
import pandas as pd
import pytest
from openpyxl import load_workbook

from costsaving.config import load_config
from costsaving.ic import OUT_COLLECT_COLUMNS, OUT_PD_COLUMNS
from costsaving.pipeline import process_ic
from costsaving.prices import base_price, parse_foc

from . import fixture_ic as fx


@pytest.fixture(scope='module')
def run():
    result, files = process_ic(fx.as_file(fx.ic_workbook_bytes()), po_files=[fx.as_file(fx.po_csv_bytes())])
    return result, files


def line(df, no, plant=None):
    rows = df[(df['Line No'] == no) & ((df['Plant'].isna()) if plant is None else (df['Plant'] == plant))]
    assert len(rows) == 1, f'line {no} plant {plant}: {len(rows)} rows'
    return rows.iloc[0]


def measure(result, name):
    s = result['Summary']
    return s.loc[s['Measure'] == name, 'Value'].iloc[0]


# ---------------------------------------------------------------- prices (D53)
def test_round_half_up_two_decimals():
    assert base_price(0.125, None, 1, 2) == 0.13   # Python round() would give 0.12
    assert base_price(100, '10F1', 100, 2) == 0.91  # 100 × 10/11 ÷ 100 = 0.909…


def test_foc_deepest_tier_and_invalid():
    assert base_price(45, '10F1,20F3', 1, 2) == 39.13  # 45 × 20/23
    assert not all(t['valid'] for t in parse_foc('5F'))
    assert not any(t['valid'] for t in parse_foc('10F1,10F2'))  # repeated buy → all invalid


# ---------------------------------------------------------------- row counts (grain)
def test_row_counts(run):
    result, _ = run
    counts = {k: len(result[k]) for k in
              ('agreement_line', 'Out_PD', 'Out_Contract', 'Out_Collect', 'Summary_Lines')}
    # Collect: 10 collected lines + one row per FOC tier as the current IC file writes it (Q4):
    # line 1 '10F1' → 1F0 + 10F1, line 2 '10F1,20F3' → 1F0 + 10F1 + 20F3
    assert counts == {'agreement_line': 12, 'Out_PD': 12, 'Out_Contract': 11, 'Out_Collect': 13,
                      'Summary_Lines': 11}


# ---------------------------------------------------------------- Baseline (D11, D13, D14, P4)
def test_lpp_latest_po_whole_po_strictly_before_key_date(run):
    sl = run[0]['Summary_Lines']
    r = line(sl, 1)
    # T101 latest PO 4500000002 = (1050 + 0) / (1000 + 100) = 0.95; T102 = 1.20; PO on the Key Date ignored
    assert r['Baseline Price Base Unit'] == 0.95
    assert r['Baseline Source'] == 'PO 4500000002'


def test_plant_group_splits_rows(run):
    sl = run[0]['Summary_Lines']
    assert line(sl, 3, 'T101')['Baseline Price Base Unit'] == 13.0
    assert line(sl, 3, 'T102')['Baseline Source'] == 'No Baseline'


def test_reference_material_code_and_manual(run):
    sl = run[0]['Summary_Lines']
    r6 = line(sl, 6)
    assert (r6['Baseline Price Base Unit'], r6['Spend Plant'], r6['Hospital Saving']) == (10.0, 'T102', 300.0)
    r10 = line(sl, 10)
    assert (r10['Baseline Price Base Unit'], r10['Baseline Source']) == (5.0, 'Manual')


# ---------------------------------------------------------------- Summary (D9, Q3)
def test_summary_totals(run):
    result, _ = run
    assert measure(result, 'Hospital Saving (Annualize)') == pytest.approx(1351)
    assert measure(result, '  Gain') == pytest.approx(1401)
    assert measure(result, '  Loss') == pytest.approx(-50)
    assert measure(result, 'Hospital Saving (Annualize + Growth)') == pytest.approx(1351 * 1.05)
    assert measure(result, 'Avoid (positive only, not in Cost Saving, Q3)') == pytest.approx(714)
    assert measure(result, 'Spending 12 months (matched)') == pytest.approx(13250)


def test_summary_by_plant(run):
    bp = run[0]['Summary_ByPlant']
    assert sorted(bp['Plant']) == ['T101', 'T102']
    assert bp['Hospital Saving'].sum() == pytest.approx(1351)


def test_line_values_match_expected(run):
    sl = run[0]['Summary_Lines']
    expected = {  # line: (Final Price Base Unit, Baseline, Hospital Saving, Avoid)
        1: (0.91, 0.95, 64.0, 304.0), 2: (39.13, 48.0, 887.0, 360.0), 4: (8.0, 7.5, -50.0, 50.0),
        7: (2.5, 2.6, 100.0, 0.0),
    }
    for no, (final, base, hs, avoid) in expected.items():
        r = line(sl, no)
        assert (r['Final Price Base Unit'], r['Baseline Price Base Unit']) == (final, base), no
        assert r['Hospital Saving'] == pytest.approx(hs), no
        assert r['Avoid'] == pytest.approx(avoid), no


# ---------------------------------------------------------------- PD Action (D52, P14)
def test_pd_action_not_collected_has_no_contract(run):
    result, _ = run
    assert 11 not in set(result['Out_Collect']['Line No'])
    assert 11 not in set(result['Summary_Lines']['Line No'])
    assert len(result['Out_Contract']) == 11
    out_pd = result['Out_PD']
    assert line(out_pd, 11)['PD Type'] == 'No Sale'
    assert line(out_pd, 4)['PD Type'] == 'Add'


# ---------------------------------------------------------------- Check codes (template-spec section 7)
def test_check_codes(run):
    check = run[0]['IC_Check']
    got = {(r['Level'], r['Message'][:3], None if pd.isna(r['Line No']) else r['Line No'])
           for _, r in check.iterrows() if r['Message'][:1] in 'EW' and r['Message'][1:3].isdigit()}
    expected = {
        ('Error', 'E10', 7), ('Error', 'E02', 8), ('Error', 'E04', 9),
        ('Warning', 'W01', 3), ('Warning', 'W01', 5), ('Warning', 'W01', 8),
        ('Warning', 'W02', 4), ('Warning', 'W02', 9), ('Warning', 'W07', 4), ('Warning', 'W08', 8),
    }
    assert got == expected
    # file-level reminder, shown once (not once per line)
    assert (check['Message'].str.startswith('ERP System ID')).sum() == 1


def test_header_errors():
    header = [(f, None if f == 'Key Date' else v) for f, v in fx.HEADER]
    result, _ = process_ic(fx.as_file(fx.ic_workbook_bytes(header=header)), po_files=[fx.as_file(fx.po_csv_bytes())])
    assert any(m.startswith('H06') for m in result['IC_Check']['Message'])


# ---------------------------------------------------------------- data contract (P5)
def test_collect_columns_keep_v1_order(run):
    collect = run[0]['Out_Collect']
    assert list(collect.columns) == OUT_COLLECT_COLUMNS
    assert OUT_COLLECT_COLUMNS[:3] == ['Commodity', 'Effective Date', 'Valid Date']
    assert OUT_COLLECT_COLUMNS[28] == 'PD'
    # First Price Base Unit carries the Baseline (P12)
    assert (collect['First Price Base Unit'].fillna(-1) == collect['Baseline Price Base Unit'].fillna(-1)).all()


def test_out_pd_columns(run):
    assert list(run[0]['Out_PD'].columns) == OUT_PD_COLUMNS


# ---------------------------------------------------------------- files
def test_output_files(run):
    _, files = run
    assert sorted(n.rsplit('_', 1)[1] for n in files) == ['Contract.xlsx', 'Processed.xlsx']
    processed = next(v for k, v in files.items() if k.endswith('_IC_Processed.xlsx'))
    wb = load_workbook(fx.as_file(processed))
    table_name = load_config().const('IC', 'Collect Table Name')
    assert table_name in wb['forSavingCollect'].tables  # the Project file finds the table in the submitted file (D57)
    rows = list(wb['forSavingCollect'].iter_rows(values_only=True))
    assert list(rows[0]) == list(run[0]['Out_Collect'].columns) and len(rows) - 1 == len(run[0]['Out_Collect'])
    assert {'Processed', 'Summary', 'IC_Check', 'Out_PD', 'Out_Contract', 'Summary_Lines'} <= set(wb.sheetnames)
    contract = next(v for k, v in files.items() if k.endswith('_for_Contract.xlsx'))
    sheets = load_workbook(fx.as_file(contract)).sheetnames
    assert sheets == ['Contract_T900000001', 'Contract_T900000002']  # one contract per vendor


def test_masters_in_separate_workbook():
    result, _ = process_ic(fx.as_file(fx.ic_workbook_bytes(with_masters=False)),
                           po_files=[fx.as_file(fx.po_csv_bytes())],
                           master_file=fx.as_file(fx.master_workbook_bytes()))
    assert measure(result, 'Hospital Saving (Annualize)') == pytest.approx(1351)


def test_pool_purchase_uses_transfer_price():
    '''Pool Purchase = All Vendor (spend of any vendor) and hospital pays the Transfer Price.'''
    header = [(f, {'Method': 'Pool Purchase', 'Saving Strategy': 'Auction'}.get(f, v)) for f, v in fx.HEADER]
    lines = [dict(line=1, mat='TST-MED-0001', uom='BX100T', final=80, transfer=90, trade='Testamol', pd='Fix Winner')]
    result, _ = process_ic(fx.as_file(fx.ic_workbook_bytes(header=header, lines=lines)),
                           po_files=[fx.as_file(fx.po_csv_bytes())])
    sl = result['Summary_Lines']
    # All Vendor: T101 (V1, 1600 TAB) + T102 (V2, 2000 TAB); paid 0.90, final 0.80, baseline 0.95
    assert sorted(sl['Spend Plant']) == ['T101', 'T102']
    assert sl['Hospital Saving'].sum() == pytest.approx((0.95 - 0.90) * 3600)
    assert sl['Group Margin'].sum() == pytest.approx((0.90 - 0.80) * 3600)
    assert not any(result['IC_Check']['Level'] == 'Error')


# ---------------------------------------------------------------- CT_Leadtime is not passed on (Q12, 2026-11-08)
def test_no_leadtime_output_and_no_ct_rule():
    header = [(f, {'Method': 'Yearly Contract'}.get(f, v)) for f, v in fx.HEADER]
    result, files = process_ic(fx.as_file(fx.ic_workbook_bytes(header=header)), po_files=[fx.as_file(fx.po_csv_bytes())])
    assert 'Out_PD_Leadtime' not in result
    assert not result['IC_Check']['Message'].str[:3].isin(['W09', 'W10']).any()


# ---------------------------------------------------------------- Collect in the current shape (Q4, 2026-11-08)
def test_collect_one_row_per_foc_tier_like_the_current_file(run):
    result, _ = run
    c = result['Out_Collect']
    two = c[c['Material Code'] == 'TST-MED-0002'].sort_values('Final Price Base Unit', ascending=False)
    assert list(two['Final FOC']) == ['1F0', '10F1', '20F3']
    assert list(two['Nego FOC']) == ['1F0', '10F1', '10F1']  # deepest Nego tier not above the Final tier
    assert list(two['Final Price Base Unit']) == [45.0, 40.91, 39.13]  # 45 × buy ÷ (buy + free), D53
    assert list(two['Nego Price Base Unit']) == [47.0, 42.73, 42.73]
    assert list(two['Avoid Base Unit']) == [2.0, 1.82, 3.6]
    assert (two['PPV Base Unit'] == (two['First Price Base Unit'] - two['Final Price Base Unit']).round(2)).all()
    no_foc = c[c['Material Code'] == 'TST-MED-0005']
    assert len(no_foc) == 1 and no_foc['Final FOC'].isna().all()  # no FOC → one row, FOC blank
    assert set(c['Project Name']) == {'[Negotiation]Renegotiation_TEST Renegotiation Vendor A'}


def test_quit_and_inactive_need_no_price_and_are_not_collected():
    lines = [dict(x, final=None) if x['line'] == 11 else x for x in fx.LINES]  # line 11 = Inactive Vendor
    result, _ = process_ic(fx.as_file(fx.ic_workbook_bytes(lines=lines)), po_files=[fx.as_file(fx.po_csv_bytes())])
    check = result['IC_Check']
    assert not ((check['Line No'] == 11) & check['Message'].str.startswith('E03')).any()
    assert 11 not in set(result['Out_Collect']['Line No']) and 11 in set(result['Out_PD']['Line No'])


def test_w04_reference_material_with_another_base_unit():
    lines = [dict(x, ref='TST-MED-0003') if x['line'] == 6 else x for x in fx.LINES]  # TUB vs TAB
    result, _ = process_ic(fx.as_file(fx.ic_workbook_bytes(lines=lines)), po_files=[fx.as_file(fx.po_csv_bytes())])
    check = result['IC_Check']
    assert ((check['Line No'] == 6) & check['Message'].str.startswith('W04')).any()
    assert not check['Message'].str.startswith('W04').sum() > 1
