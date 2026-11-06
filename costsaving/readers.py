'''Read the input files into DataFrames.

Everything is read by Excel TABLE name (not sheet position), so the same reader works for the new
IC input workbook and for the current IC file (Price_Change_Template), which already holds the
master tables dimUOM_Raw, dimPlant_Raw, dimPlantGroup and dimMasterInfo_Raw.
'''
import io

import pandas as pd
from openpyxl import load_workbook

from .config import clean, to_number, to_text

# columns of the PO-line CSV (prc_spending_sap_{year}.csv), as used by the dispatch file
PO_COLUMNS = [
    'Commodity Full', 'Plant', 'Plant Name', 'PO Number', 'PO Date', 'Material Code', 'Vendor ID',
    'Order Unit', 'Base Unit', 'Spending Quantity (Base Unit)', 'Spending Amount THB',
]
PO_TEXT_COLUMNS = ['Commodity Full', 'Plant', 'Plant Name', 'PO Number', 'Material Code', 'Vendor ID', 'Order Unit', 'Base Unit']


def _as_file(source):
    if isinstance(source, (bytes, bytearray)):
        return io.BytesIO(source)
    return source


def open_workbook(source):
    '''source = path, file object or bytes (what a browser upload gives).'''
    return load_workbook(_as_file(source), data_only=True, read_only=False)


def list_tables(workbook):
    return {name: ws.title for ws in workbook.worksheets for name in ws.tables.keys()}


def read_table(workbook, table_name, required=True):
    '''Excel table → DataFrame (header row = table header). Missing table → None or KeyError.'''
    for ws in workbook.worksheets:
        if table_name in ws.tables:
            ref = ws.tables[table_name].ref
            rows = [[cell.value for cell in row] for row in ws[ref]]
            header = [str(h).strip() if h is not None else f'Column{i + 1}' for i, h in enumerate(rows[0])]
            df = pd.DataFrame(rows[1:], columns=header, dtype=object)
            df = df.map(clean)
            return df.dropna(how='all').reset_index(drop=True)
    if required:
        raise KeyError(f'ไม่พบตาราง {table_name} ในไฟล์ (ตารางที่มี: {", ".join(sorted(list_tables(workbook)))})')
    return None


def read_header(workbook, table_name='Header'):
    '''Header table Field | Value → dict.'''
    df = read_table(workbook, table_name)
    return {to_text(r['Field']): clean(r['Value']) for _, r in df.iterrows() if to_text(r['Field'])}


def read_masterinfo(workbook):
    '''Current Info Records. Accepts the IC-file table dimMasterInfo_Raw, or the monthly export
    table MasterInfo_ActiveTH (Net Price / Per, Vendor Material Group = MFG).
    grain after: Material × Vendor × Order Unit.'''
    df = read_table(workbook, 'dimMasterInfo_Raw', required=False)
    if df is None:
        raw = read_table(workbook, 'MasterInfo_ActiveTH')
        raw['Order Unit Price'] = [
            None if to_number(n) is None or not to_number(p) else round(to_number(n) / to_number(p), 2)
            for n, p in zip(raw['Net Price'], raw['Per'])
        ]
        raw = raw.rename(columns={'Vendor Material Group': 'MFG'})
        raw = raw.sort_values('FOC', ascending=False, na_position='last', key=lambda s: s.astype(str))
        df = (raw.groupby(['Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG', 'Order Unit'],
                          dropna=False, sort=False)
                 .agg(**{'FOC': ('FOC', 'first'), 'Order Unit Price': ('Order Unit Price', 'min')})
                 .reset_index())
    return normalize_masterinfo(df)


def normalize_masterinfo(df):
    '''Codes as text, price as number, one row per Material × Vendor × Order Unit.'''
    df = df.copy()
    for col in ['Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG', 'Order Unit', 'FOC', 'Order Unit Price']:
        if col not in df:
            df[col] = None
    for col in ['Material Code', 'Vendor ID', 'Order Unit', 'FOC']:
        df[col] = df[col].map(to_text)
    df['Order Unit Price'] = df['Order Unit Price'].map(to_number)
    return df.drop_duplicates(['Material Code', 'Vendor ID', 'Order Unit']).reset_index(drop=True)


def read_po_lines(sources, materials=None, chunksize=200_000):
    '''PO lines from one or more CSV files. Only the needed columns are kept and, when materials is
    given, only those materials — the yearly file is large and the browser has limited memory.'''
    keep = set(materials) if materials else None
    parts = []
    for source in sources if isinstance(sources, (list, tuple)) else [sources]:
        reader = pd.read_csv(_as_file(source), usecols=lambda c: c in PO_COLUMNS, dtype=str,
                             keep_default_na=False, chunksize=chunksize, encoding='utf-8-sig')
        for chunk in reader:
            if keep is not None:
                chunk = chunk[chunk['Material Code'].str.strip().isin(keep)]
            parts.append(chunk)
    if not parts:
        return pd.DataFrame(columns=PO_COLUMNS)
    return normalize_po_lines(pd.concat(parts, ignore_index=True))


def normalize_po_lines(df):
    df = df.copy()
    for col in PO_TEXT_COLUMNS:
        df[col] = df[col].map(to_text)
    df['PO Date'] = pd.to_datetime(df['PO Date'], errors='coerce').dt.date
    for col in ['Spending Quantity (Base Unit)', 'Spending Amount THB']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df
