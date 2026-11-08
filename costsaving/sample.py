'''Fictitious IC test case (no real company data). Used by the tests and by the "sample data" button
of the web page, so anyone can try the app without real files. Same data as drafts/tools/build_upload_test_workbook.py,
so the expected numbers are the ones in the Expected sheet of IC_Template_v2_test.xlsx.

Builds the files a user would upload: an IC workbook (Header, Lines, CT_Leadtime + master tables)
and a PO-line CSV in the format of prc_spending_sap_{year}.csv.
'''
import io
from datetime import date

import pandas as pd

from costsaving.writers import add_table_sheet, new_workbook, to_bytes

V1, V2 = 'T900000001', 'T900000002'

HEADER = [
    ('Commodity', 'Medicine'), ('Saving Strategy', 'Negotiation'), ('Method', 'Renegotiation'),
    ('Vendor ID', V1), ('Project Name', 'TEST Renegotiation Vendor A'),
    ('Effective Date', date(2026, 12, 1)), ('Valid Date', date(2027, 11, 30)), ('Key Date', date(2026, 11, 6)),
    ('Quotation Date', date(2026, 10, 30)), ('Plant Group', None), ('Growth %', None), ('Remark', None),
]
MASTERINFO = [  # Material, Description, Vendor, Vendor Name, MFG, Order Unit, FOC, Order Unit Price
    ('TST-MED-0001', 'TESTAMOL 500 MG TAB', V1, 'Test Vendor A', 'MFG-A', 'BX100T', None, 105),
    ('TST-MED-0002', 'TESTACILLIN 250 MG SYR', V1, 'Test Vendor A', 'MFG-A', 'BOT', None, 50),
    ('TST-MED-0003', 'TESTAZOLE CREAM 5 G', V1, 'Test Vendor A', 'MFG-A', 'PK2T', None, 32),
    ('TST-MED-0004', 'TESTOVIT TAB', V2, 'Test Vendor B', 'MFG-B', 'BX10T', None, 78),
    ('TST-MED-0004', 'TESTOVIT TAB', V1, 'Test Vendor A', 'MFG-A', 'BX10T', None, 80),
    ('TST-MED-0005', 'TESTOPRIL 5 MG TAB', V1, 'Test Vendor A', 'MFG-A', 'BX10T', None, 62),
    ('TST-MED-0006', 'TESTOLOL 50 MG TAB (new code)', V1, 'Test Vendor A', 'MFG-A', 'BX30T', None, 280),
    ('TST-MED-0096', 'TESTOLOL 50 MG TAB (old code)', V1, 'Test Vendor A', 'MFG-A', 'BX30T', None, 300),
    ('TST-MED-0008', 'TESTOCET 10 MG TAB', V1, 'Test Vendor A', 'MFG-A', 'BX10T', None, 50),
    ('TST-MSP-0001', 'TEST GLOVE SIZE M', V1, 'Test Vendor A', 'MFG-A', 'BX100P', None, 260),
]
UOM = [
    ('BX100T', 'Box 100 Tablet', 100, 'TAB', 'X90'), ('BX10T', 'Box 10 Tablet', 10, 'TAB', 'X91'),
    ('BX30T', 'Box 30 Tablet', 30, 'TAB', 'X92'), ('BOT', 'Bottle', 1, 'BOT', 'BO'),
    ('PK2T', 'Pack 2.5 Tube', 2.5, 'TUB', 'X93'), ('BX100P', 'Box 100 Piece', 100, 'PCS', 'X94'),
]
PLANT = [
    ('Group T', 'T101', 'TH1', 'Premium', 'BKK/Boundaries'),
    ('Group T', 'T102', 'TH2', 'Standard', 'BKK/Boundaries'),
    ('Group T', 'T103', 'TH3', 'Standard', 'Upcountry'),
]
PLANT_GROUP = [('BKK', 'TH1,TH2', 2), ('UPC', 'TH3', 1), ('TH1', 'TH1', 1)]
LEADTIME = [
    (1, None, 'Hospital Area BKK (test)', 'General', None, 1000, 5, '5 working days'),
    (2, None, 'Hospital Area Upcountry (test)', 'General', None, 2500, 15, '15 working days'),
]
PO = [  # plant, PO, date, material, vendor, order unit, base unit, base qty, amount, commodity
    ('T101', '4500000001', date(2026, 5, 1), 'TST-MED-0001', V1, 'BX100T', 'TAB', 500, 400, 'Medicine'),
    ('T101', '4500000002', date(2026, 10, 20), 'TST-MED-0001', V1, 'BX100T', 'TAB', 1000, 1050, 'Medicine'),
    ('T101', '4500000002', date(2026, 10, 20), 'TST-MED-0001', V1, 'BX100T', 'TAB', 100, 0, 'Medicine'),
    ('T102', '4500000003', date(2026, 8, 15), 'TST-MED-0001', V2, 'BX100T', 'TAB', 2000, 2400, 'Medicine'),
    ('T101', '4500000004', date(2026, 11, 6), 'TST-MED-0001', V1, 'BX100T', 'TAB', 100, 50, 'Medicine'),
    ('T103', '4500000005', date(2025, 10, 1), 'TST-MED-0001', V1, 'BX100T', 'TAB', 100, 60, 'Medicine'),
    ('T101', '4500000010', date(2026, 10, 1), 'TST-MED-0002', V1, 'BOT', 'BOT', 100, 4800, 'Medicine'),
    ('T101', '4500000020', date(2026, 7, 1), 'TST-MED-0003', V1, 'PK2T', 'TUB', 50, 650, 'Medicine'),
    ('T101', '4500000030', date(2026, 10, 5), 'TST-MED-0004', V2, 'BX10T', 'TAB', 100, 750, 'Medicine'),
    ('T102', '4500000040', date(2026, 9, 1), 'TST-MED-0096', V1, 'BX30T', 'TAB', 300, 3000, 'Medicine'),
    ('T101', '4500000050', date(2026, 9, 15), 'TST-MSP-0001', V1, 'BX100P', 'PCS', 1000, 2600, 'Medical Supply'),
]
LINES = [
    dict(line=1, mat='TST-MED-0001', uom='BX100T', nego=110, final=100, ffoc='10F1', trade='Testamol', pd='Update'),
    dict(line=2, mat='TST-MED-0002', uom='BOT', nego=47, nfoc='10F1', final=45, ffoc='10F1,20F3', trade='Testacillin', pd='Update'),
    dict(line=3, mat='TST-MED-0003', uom='PK2T', final=30, pg='BKK', trade='Testazole', pd='Update'),
    dict(line=4, mat='TST-MED-0004', uom='BX10T', vendor=V2, nego=85, final=80, trade='Testovit', pd='Add Vendor'),
    dict(line=5, mat='TST-MED-0005', uom='BX10T', final=60, trade='Testopril', pd='Add UOM'),
    dict(line=6, mat='TST-MED-0006', uom='BX30T', final=270, ref='TST-MED-0096', trade='Testolol', pd='Update'),
    dict(line=7, mat='TST-MSP-0001', uom='BX100P', final=250, trade='Testglove', pd='Update'),
    dict(line=8, mat='TST-MED-0002', uom='CTN', final=400, trade='Testacillin', pd='Update'),
    dict(line=9, mat='TST-MED-0004', uom='BX10T', final=78, ffoc='5F', trade='Testovit', pd='Update'),
    dict(line=10, mat='TST-MED-0008', uom='BX10T', final=45, manual=50, trade='Testocet', pd='Update'),
    dict(line=11, mat='TST-MED-0096', uom='BX30T', final=300, trade='Testolol', pd='Inactive Vendor'),
]
PO_HEADERS = ['Commodity Full', 'Plant', 'Plant Name', 'PO Number', 'PO Date', 'Material Code', 'Vendor ID',
              'Order Unit', 'Base Unit', 'Spending Quantity (Base Unit)', 'Spending Amount THB']


def lines_df(lines=None):
    from .ic import LINE_COLUMNS
    rows = []
    for x in lines or LINES:
        rows.append({
            'Line No': x['line'], 'PD Action': x.get('pd'), 'Plant Group': x.get('pg'), 'Material Code': x['mat'],
            'Vendor ID': x.get('vendor'), 'Order Unit': x['uom'], 'Reference Material Code': x.get('ref'),
            'Nego Price': x.get('nego'), 'Nego FOC': x.get('nfoc'), 'Final Price': x.get('final'), 'Final FOC': x.get('ffoc'),
            'Local or Import': 'Local', 'Trade Name': x.get('trade'), 'Quotation Date': date(2026, 10, 30),
            'Transfer Price': x.get('transfer'), 'LPP Hospital Manual': x.get('manual'),
            'Registration No.': 'REG-TEST', 'Country of Manufacture': 'Thailand',
        })
    return pd.DataFrame(rows, columns=LINE_COLUMNS)


def po_df(po=None):
    plant_name = {p[1]: p[2] for p in PLANT}
    return pd.DataFrame(
        [[('01_' if p[9] == 'Medicine' else '02_') + p[9], p[0], plant_name[p[0]], p[1], p[2].isoformat(),
          p[3], p[4], p[5], p[6], p[7], p[8]] for p in po or PO],
        columns=PO_HEADERS)


def ic_workbook_bytes(header=None, lines=None, with_masters=True):
    wb = new_workbook()
    add_table_sheet(wb, 'Header', pd.DataFrame(header or HEADER, columns=['Field', 'Value']), 'Header')
    add_table_sheet(wb, 'Lines', lines_df(lines), 'Lines',
                    {'Material Code', 'Vendor ID', 'Reference Material Code', 'Final FOC', 'Nego FOC', 'Order Unit'})
    if with_masters:
        add_master_sheets(wb)
    return to_bytes(wb)


def add_master_sheets(wb):
    add_table_sheet(wb, 'MasterInfo', pd.DataFrame(MASTERINFO, columns=[
        'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG', 'Order Unit', 'FOC',
        'Order Unit Price']), 'dimMasterInfo_Raw', {'Material Code', 'Vendor ID'})
    add_table_sheet(wb, 'dimUOM_Raw', pd.DataFrame(UOM, columns=[
        'Order Unit', 'Order Unit Text', 'Conversion Rate', 'Base Unit', 'ISO Code']), 'dimUOM_Raw')
    add_table_sheet(wb, 'dimPlant_Raw', pd.DataFrame(PLANT, columns=[
        'Group', 'Plant', 'Plant Name', 'Hospital Tier', 'Location']), 'dimPlant_Raw')
    add_table_sheet(wb, 'dimPlantGroup', pd.DataFrame(PLANT_GROUP, columns=['Plant Group', 'Plant List', 'Count']),
                    'dimPlantGroup')


def master_workbook_bytes():
    wb = new_workbook()
    add_master_sheets(wb)
    return to_bytes(wb)


def po_csv_bytes(po=None):
    return po_df(po).to_csv(index=False).encode('utf-8-sig')


def as_file(data):
    return io.BytesIO(data)


# ---------------------------------------------------------------- template with data package (D56)
def master_frames():
    return {
        'masterinfo': pd.DataFrame(MASTERINFO, columns=['Material Code', 'Material Description', 'Vendor ID', 'Vendor Name',
                                                        'MFG', 'Order Unit', 'FOC', 'Order Unit Price']),
        'uom': pd.DataFrame(UOM, columns=['Order Unit', 'Order Unit Text', 'Conversion Rate', 'Base Unit', 'ISO Code']),
        'plant': pd.DataFrame(PLANT, columns=['Group', 'Plant', 'Plant Name', 'Hospital Tier', 'Location']),
        'plant_group': pd.DataFrame(PLANT_GROUP, columns=['Plant Group', 'Plant List', 'Count']),
    }


def template_bytes(header=None, lines=None, blank=False, spend_from=None, spend_to=None):
    '''The sample as the data team would publish it: template + package, built by template.build_template.
    blank=True gives the empty template (only Commodity set), as Sourcing downloads it.'''
    from datetime import datetime, timezone
    from .config import load_config
    from .readers import normalize_po_lines
    from .template import build_template
    m = master_frames()
    data, _ = build_template(
        load_config(), m['masterinfo'], m['uom'], m['plant'], m['plant_group'], normalize_po_lines(po_df()), 'Medicine',
        spend_from=spend_from, spend_to=spend_to,
        header=None if blank else (header or HEADER),
        lines=None if blank else lines_df(lines),
        created=datetime(2026, 11, 6, tzinfo=timezone.utc),
    )
    return data


# ---------------------------------------------------------------- the current IC file (Price_Change_Template)
# Same layout as the current file (sheet / table / cell names and the 48 input columns), fictitious values.
INPUT_COLUMNS = [
    'Commodity', 'Saving Strategy', 'Method', 'PD', 'Plant Group', 'Material Code', 'Material Description',
    'Vendor ID', 'Vendor Name', 'MFG', 'Order Unit', 'Order Unit Text', 'ISO Code', 'Conversion', 'Base Unit',
    'Alternative Material Code', 'Alternative Order Unit', 'a1', 'Per for Setup', 'First Price', 'First FOC',
    'Nego Price', 'Nego FOC', 'Final Price', 'Final FOC', 'a2', 'First Price Base Unit', 'Nego Price Base Unit',
    'Final Price Base Unit', 'a3', 'Avoid Base Unit', 'PPV Base Unit', '%PPV Base Unit', 'Expected', 'a4',
    'Effective Date', 'Valid Date', 'Project Name', 'Local or Import', 'Supplier Product Code',
    'Supplier Product Name', 'Trade Name', 'TMT Code or UCEP Code', 'Quotation Date', 'Remark', 'a5',
    'Combination ID', 'Plant List',
]
# formula columns of the current file: written as formulas (no cached value), never read by the app
INPUT_FORMULAS = {
    'Order Unit Text': '=XLOOKUP(1,1,1)', 'ISO Code': '=XLOOKUP(1,1,1)', 'Conversion': '=XLOOKUP(1,1,1)',
    'Base Unit': '=XLOOKUP(1,1,1)', 'First Price Base Unit': '=1', 'Nego Price Base Unit': '=1',
    'Final Price Base Unit': '=1', 'Avoid Base Unit': '=0', 'PPV Base Unit': '=0', '%PPV Base Unit': '=0',
    'Expected': '="Constant"', 'Combination ID': '="NA"', 'Plant List': '=""',
}
CT_FORM_COLUMNS = [
    'No.', 'Material Code', 'Material Description', 'Order Unit', 'UOM Description', 'a0', 'Final Price', 'Final FOC',
    'Supplier Product Code', 'Supplier Product Name', 'Trade Name', 'TMT Code or UCEP Code', 'ประเทศผู้ผลิตสินค้า',
    'Regulatory no. (เลขทะเบียน อย.)', 'Remark', 'Local or Import', 'Quotation Date',
]


def legacy_workbook_bytes(mode='noCT', lines=None, method='Renegotiation', strategy='Negotiation', stale_ct=False):
    '''The current IC file, keyed as Sourcing does it today (fictitious values).
    noCT: typed columns of Input_noCT. CT: table CT (vendor form) + CT_Leadtime + Input_CT as refreshed from
    CT with Method / PD / Plant Group typed in it. stale_ct=True leaves one CT row out of Input_CT.'''
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.table import Table
    from openpyxl.utils import get_column_letter
    lines = lines or LINES
    eff, val, quote = date(2026, 12, 1), date(2027, 11, 30), date(2026, 10, 30)
    vendor_name = {V1: 'Test Vendor A', V2: 'Test Vendor B'}
    wb = Workbook()
    summary = wb.active
    summary.title = 'Summary'
    summary['B5'] = mode
    names = {'varCT': 'Summary!$B$5'}

    def table(ws, name, top, columns, rows):
        for c, col in enumerate(columns, 2):
            ws.cell(top, c, col)
        for r, row in enumerate(rows, top + 1):
            for c, col in enumerate(columns, 2):
                ws.cell(r, c, row.get(col))
        last = top + max(len(rows), 1)
        ws.add_table(Table(displayName=name, ref=f'B{top}:{get_column_letter(len(columns) + 1)}{last}'))

    def input_row(x, ct):
        vendor = V1 if ct else x.get('vendor', V1)
        row = {
            'Method': method, 'PD': x.get('pd'), 'Plant Group': x.get('pg'), 'Material Code': x['mat'],
            'Material Description': None, 'Vendor ID': vendor, 'Vendor Name': vendor_name[vendor], 'Order Unit': x['uom'],
            'Alternative Material Code': x.get('ref'), 'Per for Setup': 1 if ct else None,
            'First Price': 999, 'Nego Price': x['final'] if ct else x.get('nego'),
            'Nego FOC': x.get('ffoc') if ct else x.get('nfoc'), 'Final Price': x['final'], 'Final FOC': x.get('ffoc'),
            'Effective Date': eff, 'Valid Date': val,
            'Project Name': f'CT2026-{vendor_name[V1]}' if ct else 'TEST Renegotiation Vendor A',
            'Local or Import': 'Local', 'Trade Name': x.get('trade'), 'Quotation Date': quote,
        }
        if ct:
            row.update({'Commodity': 'Medicine', 'Saving Strategy': strategy})
        for col, formula in INPUT_FORMULAS.items():
            row[col] = formula
        if not ct:
            row.update({'Commodity': '=Commodity', 'Saving Strategy': '=SavingStrategy'})
        return row

    if mode == 'CT':
        form = wb.create_sheet('CT_Vendor')
        form['B2'], form['B4'] = f'{V1}_{vendor_name[V1]}', quote
        names['varVendor'] = 'CT_Vendor!$B$2'
        table(form, 'CT', 6, CT_FORM_COLUMNS, [{
            'No.': i, 'Material Code': x['mat'], 'Order Unit': x['uom'], 'Final Price': x['final'],
            'Final FOC': x.get('ffoc'), 'Trade Name': x.get('trade'), 'ประเทศผู้ผลิตสินค้า': 'Thailand',
            'Regulatory no. (เลขทะเบียน อย.)': 'REG-TEST', 'Local or Import': 'Local', 'Quotation Date': '=$B$4',
        } for i, x in enumerate(lines, 1)])
        lt = wb.create_sheet('CT_Leadtime')
        lt['B3'] = '=varVendor'
        table(lt, 'CT_Leadtime', 6, ['No.', 'Hospital Area', 'Product Type', 'a0', 'Minimum Order (Baht)',
                                     'Leadtime (Days)', 'Remark'],
              [{'No.': r[0], 'Hospital Area': r[2], 'Product Type': r[3], 'Minimum Order (Baht)': r[5],
                'Leadtime (Days)': r[6], 'Remark': r[7]} for r in LEADTIME])
        ws = wb.create_sheet('CT_Input')
        ws['D2'], ws['H2'] = strategy, 'Medicine'
        names.update({'SavingStrategy_CT': 'CT_Input!$D$2', 'Commodity_CT': 'CT_Input!$H$2'})
        table(ws, 'Input_CT', 5, INPUT_COLUMNS, [input_row(x, True) for x in (lines[1:] if stale_ct else lines)])
    else:
        ws = wb.create_sheet('noCT_Input')
        ws['D2'], ws['G2'] = strategy, 'Medicine'
        names.update({'SavingStrategy': 'noCT_Input!$D$2', 'Commodity': 'noCT_Input!$G$2'})
        table(ws, 'Input_noCT', 5, INPUT_COLUMNS, [input_row(x, False) for x in lines])
    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    master = new_workbook()
    add_master_sheets(master)
    for src in master.worksheets:  # masters sit in the same file, as in the current IC file
        dst = wb.create_sheet(src.title)
        for row in src.iter_rows(values_only=True):
            dst.append(row)
        for t in src.tables.values():
            dst.add_table(Table(displayName=t.displayName, ref=t.ref))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
