'''Fictitious PD setup round (no real company data): the same ten cases as pd_test_workbook.xlsx built by
drafts/tools/build_upload_test_workbook.py in the project repository, so the expected rows are the ones of its
Expected sheet. Used by the tests and the "sample data" button of the PD mode.

Changes against that workbook, from the decisions of 2026-11-08:
  - consignment comes from ME1L (Text for Item Cat.), no longer from an Item Category column (Q11)
  - contracts come from ME3L (Doc Type ZARB) instead of a hand-made table (Q10)
  - the plant with its own tax rule is a fictitious code typed on the page (TAX_PLANT), not a real one (Q16)
  - Pl. Deliv. Time of a Medical Supply line comes from ME1L instead of CT_Leadtime (Q11, Q12)
'''
from datetime import date

import pandas as pd

from .ic import OUT_PD_COLUMNS, ICInputs
from .writers import add_table_sheet, new_workbook, to_bytes

EFF, VAL = date(2026, 12, 1), date(2027, 11, 30)
SAP_FROM, SAP_TO = date(2026, 4, 1), date(2027, 3, 31)
TAX_PLANT = 'T701'

IC_FILES = {  # Agreement ID: (file name, contract number typed by PD)
    'MED-20261201-001': ('IC_TEST_MED_001_IC_Processed.xlsx', '8199000001'),
    'MSP-20261201-001': ('IC_TEST_MSP_001_IC_Processed.xlsx', '8199000002'),
    'MSP-20261201-002': ('IC_TEST_MSP_002_IC_Processed.xlsx', None),   # consignment: no contract
    'MED-20261201-002': ('IC_TEST_MED_002_IC_Processed.xlsx', None),   # contract not created yet
    'MED-20261201-003': ('IC_TEST_MED_003_IC_Processed.xlsx', None),   # Fail / No Setup
}
PLANTS = {
    'Medicine': [('T101', 'Test Hospital 1'), ('T102', 'Test Hospital 2')],
    'Medical Supply': [('T101', 'Test Hospital 1'), ('T102', 'Test Hospital 2'), (TAX_PLANT, 'Test Hospital 701')],
}
LINES = [
    dict(agr='MED-20261201-001', line=1, com='Medicine', method='Renegotiation', mat='TST-MED-0001', vendor='T900000001',
         vname='Test Vendor A', uom='BOX', conv=10, scope='All Plant', price=100, foc='10F1', pd='Update', trade='Testamol', fixreq='✓'),
    dict(agr='MED-20261201-001', line=2, com='Medicine', method='Renegotiation', mat='TST-MED-0002', vendor='T900000001',
         vname='Test Vendor A', uom='BOT', conv=1, scope='All Plant', price=45, foc='10F1,20F3', pd='Update', trade='Testacillin'),
    dict(agr='MED-20261201-001', line=3, com='Medicine', method='Renegotiation', mat='TST-MED-0003', vendor='T900000001',
         vname='Test Vendor A', uom='PAC', conv=2.5, scope='By Plant', plant='T101', price=30, pd='Update', trade='Testazole'),
    dict(agr='MED-20261201-001', line=4, com='Medicine', method='Renegotiation', mat='TST-MED-0004', vendor='T900000001',
         vname='Test Vendor A', uom='BOX', conv=10, scope='By Plant', plant='T102', price=80, foc='10F2', pd='Update', trade='Testovit'),
    dict(agr='MSP-20261201-001', line=1, com='Medical Supply', method='Bidding', mat='TST-MSP-0001', vendor='T900000002',
         vname='Test Vendor B', uom='BOX', conv=100, scope='All Plant', price=250, foc='5F1', pd='Update', spc='GL-M-01', trade='Testglove'),
    dict(agr='MSP-20261201-001', line=2, com='Medical Supply', method='Bidding', mat='TST-MSP-0002', vendor='T900000002',
         vname='Test Vendor B', uom='EA', conv=1, scope='All Plant', price=3.5, foc='5F', pd='Update', spc='SY-05'),
    dict(agr='MSP-20261201-002', line=1, com='Medical Supply', method='Renegotiation', mat='TST-MSP-0003', vendor='T900000004',
         vname='Test Vendor D', uom='EA', conv=1, scope='By Plant', plant=TAX_PLANT, price=120, foc='10F1', pd='Update', spc='CT-14', trade='Testcath'),
    dict(agr='MSP-20261201-002', line=2, com='Medical Supply', method='Renegotiation', mat='TST-MSP-0004', vendor='T900000004',
         vname='Test Vendor D', uom='EA', conv=1, scope='By Plant', plant='T102', price=15, pd='Update', spc='DR-10', trade='Testdress'),
    dict(agr='MED-20261201-002', line=1, com='Medicine', method='Renegotiation', mat='TST-MED-0005', vendor='T900000005',
         vname='Test Vendor E', uom='BOX', conv=10, scope='By Plant', plant='T101', price=60, foc='12F2', pd='Update', trade='Testopril'),
    dict(agr='MED-20261201-003', line=1, com='Medicine', method='Renegotiation', mat='TST-MED-0006', vendor='T900000008',
         vname='Test Vendor H', uom='BOX', conv=10, scope='All Plant', price=90, pd='PD No Setup', pd_type='No Setup', trade='Testolol'),
]
CODE_FIELD = {'Medicine': 'trade', 'Medical Supply': 'spc'}


def out_pd_row(x):
    code, foc = x.get(CODE_FIELD[x['com']]), x.get('foc')
    smn = None if foc is None and code is None else ('_' + code if foc is None else '_'.join(v for v in (foc, code) if v))
    row = {c: None for c in OUT_PD_COLUMNS}
    row.update({
        'Line No': x['line'], 'PD Type': x.get('pd_type', 'Update'), 'Commodity': x['com'], 'Method': x['method'],
        'Plant Scope': x['scope'], 'Plant': x.get('plant'), 'Plant Name': None,
        'Material Code': x['mat'], 'Material Description': f"TEST ITEM {x['mat']}", 'Vendor ID': x['vendor'],
        'Vendor Name': x['vname'], 'MFG': 'MFG-TEST', 'Order Unit': x['uom'], 'Order Unit Text': x['uom'].title(),
        'Conversion': x['conv'], 'Base Unit': 'EA', 'Per for Setup': 1, 'Final Price': x['price'],
        'Price for Setup': x['price'], 'Final FOC': foc, 'Setup Material No.': smn,
        'Supplier Product Code': x.get('spc'), 'Trade Name': x.get('trade'), 'Local or Import': 'Local',
        'Registration No.': 'REG-TEST', 'Fix Vendor Requested': x.get('fixreq'),
        'Material Group': 'MG-TEST', 'Material Type': 'ZMED' if x['com'] == 'Medicine' else 'ZMSP',
        'Effective Date': EFF, 'Valid Date': VAL, 'Project Name': 'TEST PROJECT', 'Template Version': 'v2.2',
        'PD Action': x['pd'], 'Combination ID': 'TEST',
    })
    return row


def processed_bytes(agreement, edited=False, with_plants=True):
    '''A fictitious IC Processed file as the Sourcing mode writes it: Processed (checksum), Header, Lines, Out_PD,
    dimPlant_Raw. edited=True changes a keyed value after processing (the checksum no longer matches).'''
    from .pipeline import input_checksum
    lines = [x for x in LINES if x['agr'] == agreement]
    com = lines[0]['com']
    header = {'Commodity': com, 'Method': lines[0]['method'], 'Project Name': 'TEST PROJECT',
              'Effective Date': EFF, 'Valid Date': VAL, 'Key Date': date(2026, 11, 6)}
    keyed = pd.DataFrame([{'Line No': x['line'], 'PD Action': x['pd'], 'Material Code': x['mat'], 'Vendor ID': x['vendor'],
                           'Order Unit': x['uom'], 'Final Price': x['price'], 'Final FOC': x.get('foc')} for x in lines])
    checksum = input_checksum(ICInputs(header=header, lines=keyed, masterinfo=None, uom=None, plant=None,
                                       plant_group=None, po_lines=None, package={'Package Version': 'TST-20261106'}))
    if edited:
        keyed.loc[0, 'Final Price'] = keyed.loc[0, 'Final Price'] - 1
    wb = new_workbook()
    add_table_sheet(wb, 'Processed', pd.DataFrame([('Template Version', 'v2.2'), ('Package Version', 'TST-20261106'),
                                                   ('Input Checksum (SHA-256)', checksum)], columns=['Field', 'Value']), 'Processed')
    add_table_sheet(wb, 'Header', pd.DataFrame(list(header.items()), columns=['Field', 'Value']), 'Header')
    add_table_sheet(wb, 'Lines', keyed, 'Lines', ('Material Code', 'Vendor ID', 'Order Unit', 'Final FOC'))
    add_table_sheet(wb, 'Out_PD', pd.DataFrame([out_pd_row(x) for x in lines], columns=OUT_PD_COLUMNS), 'Out_PD',
                    ('Material Code', 'Vendor ID', 'Plant', 'Order Unit', 'Final FOC'))
    if with_plants:
        add_table_sheet(wb, 'dimPlant_Raw', pd.DataFrame([('Group T', p, n) for p, n in PLANTS[com]],
                                                         columns=['Group', 'Plant', 'Plant Name']), 'dimPlant_Raw', ('Plant',))
    return to_bytes(wb)


def _report_bytes(title_rows, header, rows):
    '''An SAP export: title rows, then the header row, then data (like ZMMP1004 / ZMMR1019 / ME1L / ME3L).'''
    from openpyxl import Workbook
    import io
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data'
    for t in title_rows:
        ws.append(list(t))
    ws.append(header)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


ZMMR1019_HEADER = ['Cond.record no.', 'Contract No.', 'Vendor', 'Name', 'Material', 'Material Description', 'Purchase Org.',
                   'Plant', 'Infotype', 'Order Unit', 'Unit text', 'Valid from', 'Valid to', 'Scale Status', 'Scale no.',
                   'Minimum qty', 'FreeGdsQty', 'Unit', 'Add. quantity', 'Add. qty unit', 'CalculType', 'FreeGoodsCat.']
ZMMP1004_HEADER = ['Number', 'Material', 'Description', 'Plant', 'Valid from', 'Valid to', 'Vendor', 'Name', 'Purchase Org.',
                   'Order Unit', 'Measurement unit text', 'Agreement', 'Agreement Item', 'Materials Planning(MRP)', 'Blocked',
                   'Fixed vendor', 'Fixed Agreement Item', 'Fixed issuing plant', 'MPN Material']
ME1L_HEADER = ['Purchasing Info Rec.', 'Text for Item Cat.', 'Plant', 'Purchasing Group', 'Planned Deliv. Time',
               'Minimum Order Qty', 'Order Unit', 'Variable Purchase Order Unit', 'Purch. org. data', 'Price origin',
               'Net Price', 'Currency', 'Material', 'Material Description', 'Supplier/Supplying Plant', 'Valid to']
ME3L_HEADER = ['Purchasing Doc. Type', 'Purchasing Document', 'Item', 'Document Date', 'Validity Per. Start',
               'Validity Period End', 'Plant', 'Material Group', 'Material', 'Short Text', 'Supplier/Supplying Plant',
               'Purchasing Group', 'Deletion Indicator', 'Order Unit', 'Net Price', 'Currency', 'Price Unit']


def zmmr1019_bytes():
    def row(no, mat, plant, uom):  # current FOC 10F1 on contract 8199000001
        return [no, '8199000001', 'T900000001', 'Test Vendor A', mat, 'TEST', '1000', plant, '', uom, uom.title(),
                SAP_FROM.strftime('%d.%m.%Y'), SAP_TO.strftime('%d.%m.%Y'), '', '1', 10, 10, uom, 1, uom, '2', '2']
    return _report_bytes([('Report Free Goods for Purchase',), ('', 'Run by: TEST'), ('', 'Run Date: 06.11.2026'), ()],
                         ZMMR1019_HEADER,
                         [row('0000000001', 'TST-MED-0003', 'T101', 'PAC'), row('0000000002', 'TST-MED-0004', 'T102', 'BOX'),
                          row('0000000003', 'TST-MED-0001', '', 'BOX')])


def zmmp1004_bytes():
    def sl(no, mat, plant, vendor, uom, agr, blocked, fixed):
        return [str(no), mat, 'TEST', plant, SAP_FROM, SAP_TO, vendor, 'Test Vendor', '1000', uom, uom.title(), agr, '10',
                '1', blocked, fixed, '', '', '']
    return _report_bytes([('Report Source List',), ('', 'Date: 06.11.2026'), ('', 'Time: 00:00:00'), ('', 'Records: 5'), ()],
                         ZMMP1004_HEADER, [
        sl(1, 'TST-MSP-0001', 'T101', 'T900000002', 'BOX', '8199000090', 'NO', 'NO'),
        sl(2, 'TST-MSP-0001', 'T101', 'T900000003', 'BOX', '8199000091', 'NO', 'YES'),
        sl(3, 'TST-MSP-0001', 'T102', 'T900000003', 'BOX', '8199000091', 'NO', 'YES'),
        sl(4, 'TST-MED-0005', 'T101', 'T900000006', 'BOX', '8199000092', 'NO', 'NO'),
        sl(5, 'TST-MED-0001', 'T101', 'T900000007', 'BOX', '8199000093', 'YES', 'NO'),
    ])


def me1l_bytes():
    def ir(cat, plant, group, deliv, minimum, uom, mat, vendor):
        return ['5300000001', cat, plant, group, deliv, minimum, uom, '2', 'No', 'Condition', 0, '', mat, 'TEST',
                f'{vendor} TEST VENDOR', date(9999, 12, 31)]
    return _report_bytes([], ME1L_HEADER, [
        ir('Consignment', TAX_PLANT, '1B0', 0, 0, 'EA', 'TST-MSP-0003', 'T900000004'),
        ir('Consignment', 'T102', '1B0', 0, 0, 'EA', 'TST-MSP-0004', 'T900000004'),
        ir('Standard', 'T101', '1B0', 7, 0, 'BOX', 'TST-MSP-0001', 'T900000002'),
    ])


def me3l_bytes():
    def item(no, it, plant, mat, uom, vendor):
        return ['ZARB', no, it, EFF, EFF, VAL, plant or '', 'MG-TEST', mat, 'TEST', f'{vendor}  TEST VENDOR', '1A0', '', uom,
                1.0, 'THB', 1]
    return _report_bytes([], ME3L_HEADER, [
        item('8199000001', '1', None, 'TST-MED-0001', 'BOX', 'T900000001'),
        item('8199000001', '2', None, 'TST-MED-0002', 'BOT', 'T900000001'),
        item('8199000001', '3', 'T101', 'TST-MED-0003', 'PAC', 'T900000001'),
        item('8199000001', '4', 'T102', 'TST-MED-0004', 'BOX', 'T900000001'),
        item('8199000002', '1', None, 'TST-MSP-0001', 'BOX', 'T900000002'),
        # a purchase order in the same export (not a contract): must be ignored
        ['ZFRE', '4500000001', '1', EFF, None, None, '', 'MG-TEST', 'TST-MSP-0002', 'TEST', 'T900000002  TEST VENDOR',
         '1B0', '', 'EA', 1.0, 'THB', 1],
    ])


def vendor_tax_df():
    return pd.DataFrame([('T900000004', 'V7')], columns=['Vendor ID', 'Vendor Tax Code'])


def round_inputs():
    '''(files [(name, bytes, agreement id, contract numbers text)], reports {name: bytes}) of the sample round.'''
    files = [(name, processed_bytes(agr), agr, contract or '') for agr, (name, contract) in IC_FILES.items()]
    reports = {'ZMMP1004': zmmp1004_bytes(), 'ZMMR1019': zmmr1019_bytes(), 'ME1L': me1l_bytes(), 'ME3L': me3l_bytes()}
    return files, reports
