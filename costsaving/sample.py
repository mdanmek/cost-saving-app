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
LINE_HEADERS = [
    'Line No', 'Material Code', 'Material Description', 'Order Unit', 'UOM Description', 'a0',
    'Final Price', 'Final FOC', 'Supplier Product Code', 'Supplier Product Name', 'Trade Name',
    'TMT Code or UCEP Code', 'Country of Manufacture', 'Registration No.', 'Remark', 'Local or Import', 'Quotation Date',
    'Vendor ID', 'Plant Group', 'PD Action', 'Nego Price', 'Nego FOC', 'Per for Setup',
    'Transfer Price', 'LPP SaveDrug', 'LPP DrPharma', 'Reference Material Code', 'LPP Hospital Manual',
    'Growth %', 'Fix Vendor Requested', 'UCEP / Note for Info',
    'Registration Type', 'Special Feature', 'Standard Qty', 'Minimum Qty',
]
PO_HEADERS = ['Commodity Full', 'Plant', 'Plant Name', 'PO Number', 'PO Date', 'Material Code', 'Vendor ID',
              'Order Unit', 'Base Unit', 'Spending Quantity (Base Unit)', 'Spending Amount THB']


def lines_df(lines=None):
    rows = []
    for x in lines or LINES:
        rows.append([
            x['line'], x['mat'], None, x['uom'], None, None,
            x['final'], x.get('ffoc'), None, None, x.get('trade'),
            None, 'Thailand', 'REG-TEST', None, 'Local', date(2026, 10, 30),
            x.get('vendor'), x.get('pg'), x.get('pd'), x.get('nego'), x.get('nfoc'), None,
            x.get('transfer'), None, None, x.get('ref'), x.get('manual'),
            None, None, None, None, None, None, None,
        ])
    return pd.DataFrame(rows, columns=LINE_HEADERS)


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
    add_table_sheet(wb, 'CT_Leadtime', pd.DataFrame(LEADTIME, columns=[
        'No.', 'Vendor ID', 'Hospital Area', 'Product Type', 'a0', 'Minimum Order (Baht)', 'Leadtime (Days)', 'Remark']),
        'CT_Leadtime')
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


def template_bytes(header=None, lines=None, leadtime=None, blank=False, spend_from=None, spend_to=None):
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
        leadtime=None if blank else (LEADTIME if leadtime is None else leadtime),
        created=datetime(2026, 11, 6, tzinfo=timezone.utc),
    )
    return data
