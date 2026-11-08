'''Write the IC outputs as Excel files (bytes), so the browser can offer them as downloads.

    IC_Processed.xlsx  the one file submitted in SharePoint (D57): IC Request → IC Approve → IC Collected
                       Processed (checksum), Summary, IC_Check, what was keyed, table Input_forSavingCollect
                       (read by the Project file), Out_PD + dimPlant_Raw (PD mode), Out_Contract, Summary_ByPlant / _Lines
    for_Contract.xlsx  one sheet per Vendor (Ariba)     → Sourcing imports into Ariba after approval
'''
import io
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF')
PRICE_HINTS = ('Price', 'Saving', 'Gain', 'Loss', 'Margin', 'Value', 'Avoid', 'Amount', 'PPV', 'LPP', 'Baseline')


def _cell_value(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, 'item') and not isinstance(v, (str, bytes)):
        return v.item()
    return v


def _safe_table_name(name):
    return ''.join(ch if ch.isalnum() or ch == '_' else '_' for ch in name)[:250]


def add_table_sheet(wb, sheet_name, df, table_name=None, text_columns=()):
    ws = wb.create_sheet(sheet_name[:31])
    columns = list(df.columns)
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for row in df.itertuples(index=False):
        ws.append([_cell_value(v) for v in row])
    for i, col in enumerate(columns, 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = min(max(len(str(col)) + 2, 10), 40)
        for cell in ws[letter][1:]:
            if isinstance(cell.value, (date, datetime)):
                cell.number_format = 'yyyy-mm-dd'
            elif col in text_columns and cell.value is not None:
                cell.value, cell.number_format = str(cell.value), '@'
            elif isinstance(cell.value, float) and any(h in col for h in PRICE_HINTS):
                cell.number_format = '#,##0.00'
    ws.freeze_panes = 'A2'
    if columns:
        last_row = max(ws.max_row, 2)  # an Excel table needs at least one data row
        ref = f'A1:{get_column_letter(len(columns))}{last_row}'
        table = Table(displayName=_safe_table_name(table_name or sheet_name), ref=ref)
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(table)
    return ws


def to_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def new_workbook():
    wb = Workbook()
    wb.remove(wb.active)
    return wb


TEXT_CODES = ('Material Code', 'Vendor ID', 'Plant', 'Spend Plant', 'Item ID', 'Material Number', 'PO Number',
              'LPP PO Number', 'Reference Material Code')


def write_processed(result, inputs, stamp, collect_table_name):
    '''The one file Sourcing submits in SharePoint (D57). It travels IC Request → IC Approve →
    IC Collected like the current IC file, so it carries the table Input_forSavingCollect that the
    Project file reads, and what PD needs (Out_PD). The data package is not copied: the file stays small.'''
    wb = new_workbook()
    add_table_sheet(wb, 'Processed', pd.DataFrame(stamp, columns=['Field', 'Value']), 'Processed')
    add_table_sheet(wb, 'Summary', result['Summary'], 'Summary')
    add_table_sheet(wb, 'IC_Check', result['IC_Check'], 'IC_Check', TEXT_CODES)
    add_table_sheet(wb, 'Header', pd.DataFrame(list(inputs.header.items()), columns=['Field', 'Value']), 'Header')
    add_table_sheet(wb, 'Lines', inputs.lines, 'Lines', TEXT_CODES + ('Order Unit', 'Reference Order Unit', 'Final FOC', 'Nego FOC'))
    if inputs.plant_group is not None:
        add_table_sheet(wb, 'dimPlantGroup', inputs.plant_group, 'dimPlantGroup')
    add_table_sheet(wb, 'forSavingCollect', result['Out_Collect'], collect_table_name, TEXT_CODES)
    add_table_sheet(wb, 'Out_PD', result['Out_PD'], 'Out_PD', TEXT_CODES)
    # plants of the commodity: the PD mode splits an All Plant line into these (S1, Q15)
    if inputs.plant is not None:
        add_table_sheet(wb, 'dimPlant_Raw', inputs.plant, 'dimPlant_Raw', TEXT_CODES)
    add_table_sheet(wb, 'Out_Contract', result['Out_Contract'], 'Out_Contract', TEXT_CODES)
    add_table_sheet(wb, 'Summary_ByPlant', result['Summary_ByPlant'], 'Summary_ByPlant', TEXT_CODES)
    add_table_sheet(wb, 'Summary_Lines', result['Summary_Lines'], 'Summary_Lines', TEXT_CODES)
    return to_bytes(wb)


def write_for_contract(result):
    '''One sheet per Vendor: one Ariba contract per Vendor (C1). The Vendor ID column is not part of the import.'''
    wb = new_workbook()
    df = result['Out_Contract']
    if df.empty:
        add_table_sheet(wb, 'Contract', df.drop(columns=['Vendor ID']), 'Contract')
    for i, (vendor, part) in enumerate(df.groupby('Vendor ID', sort=True), 1):
        part = part.drop(columns=['Vendor ID']).copy()
        part['Item Number'] = range(1, len(part) + 1)
        add_table_sheet(wb, f'Contract_{vendor}', part, f'Contract_{i}', TEXT_CODES)
    return to_bytes(wb)


# ---------------------------------------------------------------- PD setup mode
def plain_sheet(wb, sheet_name, df):
    '''Header in row 1, values from row 2, no Excel table: the layout of the PD upload templates.
    Every value is written as text except numbers, so codes keep their leading zeros.'''
    ws = wb.create_sheet(sheet_name[:31])
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in df.itertuples(index=False):
        values = [_cell_value(v) for v in row]
        ws.append(values)
        for cell in ws[ws.max_row]:
            if isinstance(cell.value, str):
                cell.number_format = '@'
    for i, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(col)) + 2, 10), 40)
    return ws


def write_setup_files(result, stamp):
    '''Upload files of the PD setup mode (setup-outputs sections 1, 2, 2A) + Upload_Check.'''
    from .setup import FOC_SHEETS
    out = {}
    wb = new_workbook()
    plain_sheet(wb, 'Load-SL', result['Load-SL'])
    plain_sheet(wb, 'ME01 (manual)', result['ME01 (manual)'])
    out[f'Upload_SourceList_{stamp}.xlsx'] = to_bytes(wb)
    wb = new_workbook()
    plain_sheet(wb, 'Template Info', result['Template Info'])
    out[f'Upload_InfoRecord_{stamp}.xlsx'] = to_bytes(wb)
    wb = new_workbook()
    for sheet in FOC_SHEETS.values():
        plain_sheet(wb, sheet, result[sheet])
    out[f'Upload_FOC_{stamp}.xlsx'] = to_bytes(wb)
    wb = new_workbook()
    add_table_sheet(wb, 'Upload_Check', result['Upload_Check'], 'Upload_Check', TEXT_CODES)
    add_table_sheet(wb, 'Upload_PD_Manual', result['Upload_PD_Manual'], 'Upload_PD_Manual', TEXT_CODES)
    lines = result['setup_lines']
    keep = [c for c in ['IC File', 'Agreement ID', 'Line No', 'PD Type', 'PD Action', 'Commodity', 'Method', 'Plant Scope',
                        'Plant', 'Material Code', 'Vendor ID', 'Order Unit', 'Final Price', 'Final FOC', 'Category',
                        'Contract List', 'Contract Match Type', 'SAP Contract No.', 'SAP Contract Item',
                        'Effective Date', 'Valid Date'] if c in lines]
    add_table_sheet(wb, 'setup_lines', lines[keep] if keep else lines, 'setup_lines', TEXT_CODES)
    out[f'Upload_Check_{stamp}.xlsx'] = to_bytes(wb)
    return out
