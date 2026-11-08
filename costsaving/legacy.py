'''Read the current IC file (Price_Change_Template v.20261102) as Sourcing keys it today.

What Sourcing keys in that file (checked on the structure of the real file, no values kept here):

    varCT = noCT   sheet noCT_Input, table Input_noCT: the typed columns only (NOCT_TYPED_COLUMNS);
                   Commodity and Saving Strategy are cells above the table (names Commodity, SavingStrategy)
    varCT = CT     CT_Vendor and CT_Leadtime are what the vendor sends back. CT_Input (query table Input_CT)
                   maps it to the current Info Record (First Price) so Sourcing does not map by hand, then
                   Sourcing types Method / PD / Plant Group and may change values or add / delete rows there.
                   So CT_Input, as Sourcing left it, is the starting point (2026-11-08), like noCT_Input.

Input_CT and Input_noCT have the same 48 columns: both are read the same way, as ONE source shaped like
the typed columns of Input_noCT, then go through the same process as the template (ic.run_ic).
Formula columns (Order Unit Text, Conversion, Base Unit, base prices, PPV, Expected, Combination ID,
Plant List) are not read: the app computes them from the masters. CT_Leadtime is not used (Q12).
CT_Vendor only adds what Input_CT does not carry (country of manufacture, registration no.), when found.

First Price / First FOC are read by the old file but not used: the Baseline is Min LPP (D12).
The old file has no Key Date (the Baseline was First Price), so the caller passes it.
'''
import pandas as pd

from .config import clean, to_text
from .ic import LINE_COLUMNS, records, to_date
from .readers import read_table

INPUT_TABLE = {'CT': 'Input_CT', 'noCT': 'Input_noCT'}
HEADER_NAMES = {'CT': ('Commodity_CT', 'SavingStrategy_CT'), 'noCT': ('Commodity', 'SavingStrategy')}

# columns Sourcing types in Input_noCT (every other column of the table is a formula)
NOCT_TYPED_COLUMNS = [
    'Method', 'PD', 'Plant Group', 'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG',
    'Order Unit', 'Alternative Material Code', 'Alternative Order Unit', 'Per for Setup',
    'First Price', 'First FOC', 'Nego Price', 'Nego FOC', 'Final Price', 'Final FOC',
    'Effective Date', 'Valid Date', 'Project Name', 'Local or Import', 'Supplier Product Code',
    'Supplier Product Name', 'Trade Name', 'TMT Code or UCEP Code', 'Quotation Date', 'Remark',
]
# vendor quote form (table CT) → names used by the app; only what Input_CT does not carry
CT_FORM_EXTRA = {'ประเทศผู้ผลิตสินค้า': 'Country of Manufacture', 'Regulatory no. (เลขทะเบียน อย.)': 'Registration No.'}
# per-file values: the app keys them once in Header, the old file repeats them on every row
HEADER_FROM_ROWS = ['Method', 'Effective Date', 'Valid Date', 'Project Name']


def name_value(workbook, name):
    '''Value of the cell a defined name points to (e.g. varCT → Summary!$B$5); None when missing.'''
    if name not in workbook.defined_names:
        return None
    for sheet, coord in workbook.defined_names[name].destinations:
        if sheet in workbook.sheetnames:
            return clean(workbook[sheet][coord.replace('$', '')].value)
    return None


def is_legacy(workbook):
    '''The current IC file: has Input_noCT or Input_CT and no Header table of the app's format.'''
    tables = {n for ws in workbook.worksheets for n in ws.tables}
    return 'Header' not in tables and bool(tables & set(INPUT_TABLE.values()))


def _distinct(values):
    out = []
    for v in values:
        v = clean(v)
        key = v.date() if hasattr(v, 'date') and callable(v.date) else v
        if v is not None and key not in [x[0] for x in out]:
            out.append((key, v))
    return [x[1] for x in out]


def read_legacy(workbook, key_date=None):
    '''Current IC file → (header dict, lines DataFrame in LINE_COLUMNS, messages).
    messages = [(level, upload, text)] for IC_Check.'''
    messages = []
    var_ct = to_text(name_value(workbook, 'varCT'))
    mode = 'CT' if var_ct == 'CT' else 'noCT'
    if var_ct not in INPUT_TABLE:
        messages.append(('Error', 'Header', f"L00 varCT (Summary) ต้องเป็น CT หรือ noCT แต่ได้ '{var_ct}' ใช้ noCT ไปก่อน"))
    source = read_table(workbook, INPUT_TABLE[mode])
    for col in NOCT_TYPED_COLUMNS:
        if col not in source:
            source[col] = None
    source = source[NOCT_TYPED_COLUMNS].copy()
    source = source[source['Material Code'].map(to_text).notna()].reset_index(drop=True)

    extra = {}
    if mode == 'CT':
        # an item of the vendor form that Sourcing left without Method and PD is not part of this
        # negotiation: the current file drops it (no Combination ID), so it is left out here too
        blank = source['Method'].map(to_text).isna() & source['PD'].map(to_text).isna()
        if blank.any():
            messages.append(('Warning', 'CT_Input', f'L03 ข้าม {int(blank.sum())} รายการใน CT_Input ที่ไม่ได้กรอก Method และ PD '
                             '(ไฟล์เดิมก็ไม่นับรายการเหล่านี้) ถ้าต้องการให้นับ กรอก Method และ PD'))
        source = source[~blank].reset_index(drop=True)
        form = read_table(workbook, 'CT', required=False)
        for r in records(form):
            key = (to_text(r.get('Material Code')), to_text(r.get('Order Unit')))
            if key[0]:  # one row per Material × Order Unit: a duplicate in the form must not add rows
                extra.setdefault(key, {v: clean(r.get(k)) for k, v in CT_FORM_EXTRA.items()})

    commodity_name, strategy_name = HEADER_NAMES[mode]
    header = {'Commodity': to_text(name_value(workbook, commodity_name)),
              'Saving Strategy': to_text(name_value(workbook, strategy_name))}
    for col in HEADER_FROM_ROWS:
        values = _distinct(source[col])
        header[col] = values[0] if values else None
        if len(values) > 1:
            messages.append(('Error', INPUT_TABLE[mode], f'L01 ไฟล์นี้มี {col} {len(values)} ค่า '
                             f'ระบบใหม่ใช้ค่าเดียวต่อไฟล์ (Header) แยกไฟล์ตาม {col}'))
    quotation = _distinct(source['Quotation Date'])
    header['Quotation Date'] = quotation[0] if len(quotation) == 1 else None
    header['Key Date'] = key_date
    header.update({'Vendor ID': None, 'Plant Group': None, 'Growth %': None, 'Remark': None})

    rows = []
    for i, r in enumerate(source.to_dict('records'), 1):
        x = extra.get((to_text(r['Material Code']), to_text(r['Order Unit'])), {})
        rows.append({
            'Line No': i, 'PD Action': r['PD'], 'Plant Group': r['Plant Group'], 'Material Code': r['Material Code'],
            'Vendor ID': r['Vendor ID'], 'Order Unit': r['Order Unit'],
            # Alternative Material Code / Order Unit = the old code / unit of a moved item (Q14)
            'Reference Material Code': r['Alternative Material Code'], 'Reference Order Unit': r['Alternative Order Unit'],
            'Per for Setup': r['Per for Setup'], 'Nego Price': r['Nego Price'], 'Nego FOC': r['Nego FOC'],
            'Final Price': r['Final Price'], 'Final FOC': r['Final FOC'], 'Local or Import': r['Local or Import'],
            'Supplier Product Code': r['Supplier Product Code'], 'Supplier Product Name': r['Supplier Product Name'],
            'Trade Name': r['Trade Name'], 'TMT Code or UCEP Code': r['TMT Code or UCEP Code'],
            'Quotation Date': to_date(r['Quotation Date']), 'Remark': r['Remark'],
            'Country of Manufacture': x.get('Country of Manufacture'), 'Registration No.': x.get('Registration No.'),
        })
    lines = pd.DataFrame(rows, columns=LINE_COLUMNS, dtype=object)
    return header, lines, messages
