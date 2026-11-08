'''Template IC with its data package — one Excel file, as today (D56).

The data team builds the template for one commodity and puts it on SharePoint. Sourcing downloads
it, keys Header / Lines (+ their own Plant Groups), and the app processes it.

    sheet               who fills it        what it is
    ─────────────────   ─────────────────   ─────────────────────────────────────────────────────
    README              data team (app)     how to use, package version
    Header              Sourcing            one agreement (Field | Value, dropdowns)
    Lines               Sourcing            one row per item (typed columns of noCT_Input, Q13)
    dimPlantGroup       Sourcing may add    Plant Group → Plant List (codes or names, comma separated)
    ─── package (controlled by the data team until the next package round) ───────────────────
    Package_Info        data team           version, commodity, spending window
    dimMasterInfo_Raw   data team           current Info Records: Material × Vendor × Order Unit
    dimUOM_Raw          data team           Order Unit → Base Unit, Conversion
    dimPlant_Raw        data team           Plant → Plant Name, Group
    fact_po_lines       data team           PO lines of the commodity in the spending window
    dimPlantGroup_Official data team        copy of the data team's groups: these names cannot be changed (D57)
    dimMaterial         data team           Material → Commodity, from the PO lines of every commodity
                                            (E10: a Medical Supply item keyed in a Medicine file)
    lists               (hidden)            dropdown values from config

Every sheet holds one table that starts at A1 (header row 1), so the file can be read in streaming
mode: the PO lines can be ~160,000 rows, and loading them cell by cell would not fit the memory of
the browser. Only the PO lines of the materials in Lines are kept.
'''
from datetime import date, datetime, timezone
import io
import warnings

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .config import Config, clean, to_number, to_text
from .ic import LINE_COLUMNS, add_months, to_date
from .readers import PO_COLUMNS, normalize_masterinfo, normalize_po_lines

HEADER_FIELDS = [
    'Commodity', 'Saving Strategy', 'Method', 'Vendor ID', 'Project Name', 'Effective Date', 'Valid Date',
    'Key Date', 'Quotation Date', 'Plant Group', 'Growth %', 'Remark',
]
PLANT_GROUP_COLUMNS = ['Plant Group', 'Plant List', 'Count']
MASTERINFO_COLUMNS = ['Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG', 'Order Unit',
                      'FOC', 'Order Unit Price']
UOM_COLUMNS = ['Order Unit', 'Order Unit Text', 'Conversion Rate', 'Base Unit', 'ISO Code']
PLANT_COLUMNS = ['Group', 'Plant', 'Plant Name', 'Hospital Tier', 'Location']
# Commodity Full and Plant Name are left out of the package (one commodity per file; name from dimPlant_Raw):
# fewer cells to read in the browser
PACKAGE_PO_COLUMNS = [c for c in PO_COLUMNS if c not in ('Commodity Full', 'Plant Name')]
MATERIAL_COLUMNS = ['Material Code', 'Commodity']
PACKAGE_SHEET = 'Package_Info'
PO_SHEET = 'fact_po_lines'
INPUT_ROWS_WITH_DROPDOWN = 2000

HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
INPUT_FILL = PatternFill('solid', fgColor='375623')
HEADER_FONT = Font(bold=True, color='FFFFFF')


# ---------------------------------------------------------------- build (data team)
def package_po_lines(po, commodity, spend_from, spend_to):
    '''PO lines of one commodity inside [spend_from, spend_to]. po = normalize_po_lines output.'''
    df = po.copy()
    suffix = df['Commodity Full'].map(lambda v: None if v is None else (v.split('_', 1)[1] if '_' in v else v))
    df = df[(suffix == commodity) & df['PO Date'].map(lambda d: d is not None and not pd.isna(d) and spend_from <= d <= spend_to)]
    return df.sort_values(['PO Date', 'PO Number', 'Material Code'], kind='stable')[PACKAGE_PO_COLUMNS].reset_index(drop=True)


def material_commodity(po):
    '''Material → Commodity from all PO lines (any commodity). A material bought under two
    commodities keeps the one with the most lines.'''
    df = po[['Material Code', 'Commodity Full']].dropna().copy()
    if df.empty:
        return pd.DataFrame(columns=MATERIAL_COLUMNS)
    df['Commodity'] = df['Commodity Full'].map(lambda v: v.split('_', 1)[1] if '_' in v else v)
    counts = df.groupby(['Material Code', 'Commodity']).size().reset_index(name='n')
    counts = counts.sort_values(['Material Code', 'n', 'Commodity'], ascending=[True, False, True])
    return counts.drop_duplicates('Material Code')[MATERIAL_COLUMNS].reset_index(drop=True)


def default_window(po, cfg: Config):
    '''Spending To = latest PO Date in the data; From = To − Package Spending Months.'''
    dates = [d for d in po['PO Date'] if d is not None and not pd.isna(d)]
    if not dates:
        return None, None
    to = max(dates)
    return add_months(to, -int(cfg.const_number('Package', 'Spending Months'))), to


def package_version(cfg: Config, commodity, spend_to):
    code = {r['Commodity']: r['Commodity Code'] for _, r in cfg.dim_commodity.iterrows()}.get(commodity, commodity)
    return f'{code}-{spend_to:%Y%m%d}'


def _style_header(cells, fill=HEADER_FILL):
    for c in cells:
        c.fill, c.font = fill, HEADER_FONT
    return cells


def _value(v):
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


def _table_sheet(wb, name, columns, rows, fill=HEADER_FILL, widths=None, hidden=False, min_rows=1, validations=()):
    '''Write-only sheet with one Excel table at A1 named like the sheet.'''
    ws = wb.create_sheet(name)
    if hidden:
        ws.sheet_state = 'hidden'
    for i, col in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(col, min(max(len(col) + 2, 10), 40))
    ws.freeze_panes = 'A2'
    for dv in validations:
        ws.data_validations.dataValidation.append(dv)
    header = []
    for col in columns:
        cell = WriteOnlyCell(ws, value=col)
        cell.fill, cell.font = fill, HEADER_FONT
        header.append(cell)
    ws.append(header)
    n = 0
    for row in rows:
        ws.append([_value(v) for v in row])
        n += 1
    for _ in range(max(0, min_rows - n)):
        ws.append([None] * len(columns))
        n += 1
    table = Table(displayName=name, ref=f'A1:{get_column_letter(len(columns))}{n + 1}')
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    table._initialise_columns()
    for tc, col in zip(table.tableColumns, columns):
        tc.name = col
    with warnings.catch_warnings():  # columns are set above; openpyxl warns in write-only mode anyway
        warnings.simplefilter('ignore', UserWarning)
        ws.add_table(table)
    return ws


def _list_validation(sheet_range, cells):
    dv = DataValidation(type='list', formula1=sheet_range, allow_blank=True, showErrorMessage=False)
    dv.add(cells)
    return dv


def build_template(cfg: Config, masterinfo, uom, plant, plant_group, po, commodity,
                   spend_from=None, spend_to=None, header=None, lines=None, created=None):
    '''Returns (xlsx bytes, Package_Info dict). masterinfo / uom / plant / plant_group = DataFrames
    as read from the master workbook; po = normalized PO lines (all commodities is fine).
    header / lines prefill the input sheets (used for the sample file).'''
    if spend_from is None or spend_to is None:
        f, t = default_window(po, cfg)
        spend_from, spend_to = spend_from or f, spend_to or t
    if spend_to is None:
        raise ValueError('ไม่มี PO Date ในไฟล์ PO line จึงกำหนดช่วง spending ไม่ได้')
    po_pkg = package_po_lines(po, commodity, spend_from, spend_to)
    info = {
        'Package Version': package_version(cfg, commodity, spend_to),
        'Commodity': commodity,
        'Spending From': spend_from, 'Spending To': spend_to,
        'Spending Rows': len(po_pkg),
        'Template Version': cfg.const('IC', 'Template Version'),
        'Created (UTC)': (created or datetime.now(timezone.utc)).strftime('%Y-%m-%d %H:%M'),
    }

    # dropdown values (hidden sheet 'lists'): Commodity, Saving Strategy, Method, PD Action
    lists = {
        'Commodity': list(cfg.dim_commodity['Commodity']),
        'Saving Strategy': list(dict.fromkeys(cfg.dim_strategy_method['Strategy'])),
        'Method': list(cfg.dim_method['Method']),
        'PD Action': list(cfg.dim_pd['PD Action']),
    }
    depth = max(len(v) for v in lists.values())
    list_rows = [[(lists[k][i] if i < len(lists[k]) else None) for k in lists] for i in range(depth)]

    def list_ref(field):
        col = get_column_letter(list(lists).index(field) + 1)
        return f'lists!${col}$2:${col}${len(lists[field]) + 1}'

    wb = Workbook(write_only=True)
    readme = wb.create_sheet('README')
    readme.column_dimensions['A'].width = 120
    for text in [
        'Template IC — คีย์ Header และ Lines แล้วอัปโหลดไฟล์นี้ทั้งไฟล์ที่ https://mdanmek.github.io/cost-saving-app/',
        'ห้ามแก้ชื่อชีต ชื่อคอลัมน์ หรือลบแถวหัวตาราง ทุกตารางต้องเริ่มที่เซลล์ A1',
        'dimPlantGroup: เพิ่มกลุ่มใหม่ได้ (Plant List = รหัสหรือชื่อ Plant คั่นด้วย ,) แต่ห้ามแก้กลุ่มที่ทีมข้อมูลตั้งไว้ ถ้าต้องการชุด Plant อื่นให้ตั้งชื่อกลุ่มใหม่',
        'Package_Info, dimMasterInfo_Raw, dimUOM_Raw, dimPlant_Raw, fact_po_lines, dimMaterial, dimPlantGroup_Official เป็นข้อมูลที่ทีมข้อมูลควบคุม ห้ามแก้',
        f"Package {info['Package Version']} · spending {spend_from:%Y-%m-%d} ถึง {spend_to:%Y-%m-%d} · {info['Spending Rows']:,} PO lines",
    ]:
        readme.append([text])

    header_values = dict(header or [])
    header_values.setdefault('Commodity', commodity)
    header_rows = [(f, header_values.get(f)) for f in HEADER_FIELDS]
    pos = {f: i + 2 for i, f in enumerate(HEADER_FIELDS)}
    _table_sheet(wb, 'Header', ['Field', 'Value'], header_rows, fill=INPUT_FILL, widths={'Field': 22, 'Value': 40},
                 validations=[_list_validation(list_ref(f), f'B{pos[f]}') for f in ('Commodity', 'Saving Strategy', 'Method')])

    line_rows = [] if lines is None else [[r.get(c) for c in LINE_COLUMNS] for r in lines.to_dict('records')]
    pd_col = get_column_letter(LINE_COLUMNS.index('PD Action') + 1)
    _table_sheet(wb, 'Lines', LINE_COLUMNS, line_rows, fill=INPUT_FILL,
                 validations=[_list_validation(list_ref('PD Action'), f'{pd_col}2:{pd_col}{INPUT_ROWS_WITH_DROPDOWN}')])
    _table_sheet(wb, 'dimPlantGroup', PLANT_GROUP_COLUMNS,
                 [[r.get(c) for c in PLANT_GROUP_COLUMNS] for r in plant_group.to_dict('records')], fill=INPUT_FILL,
                 widths={'Plant List': 80})

    _table_sheet(wb, PACKAGE_SHEET, ['Field', 'Value'], list(info.items()), widths={'Field': 22, 'Value': 30})
    mi = normalize_masterinfo(masterinfo)
    _table_sheet(wb, 'dimMasterInfo_Raw', MASTERINFO_COLUMNS, mi[MASTERINFO_COLUMNS].itertuples(index=False))
    _table_sheet(wb, 'dimUOM_Raw', UOM_COLUMNS, uom.reindex(columns=UOM_COLUMNS).itertuples(index=False))
    _table_sheet(wb, 'dimPlant_Raw', PLANT_COLUMNS, plant.reindex(columns=PLANT_COLUMNS).itertuples(index=False))
    _table_sheet(wb, PO_SHEET, PACKAGE_PO_COLUMNS, po_pkg.itertuples(index=False))
    _table_sheet(wb, 'dimMaterial', MATERIAL_COLUMNS, material_commodity(po).itertuples(index=False))
    _table_sheet(wb, 'dimPlantGroup_Official', PLANT_GROUP_COLUMNS,
                 [[r.get(c) for c in PLANT_GROUP_COLUMNS] for r in plant_group.to_dict('records')], widths={'Plant List': 80})
    _table_sheet(wb, 'lists', list(lists), list_rows, hidden=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), info


# ---------------------------------------------------------------- read (Sourcing)
def _open_stream(source):
    '''source = bytes or a path. Streaming mode: cells are read one row at a time, never all at once.'''
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    return load_workbook(source, read_only=True, data_only=True)


def has_package(source_bytes):
    '''True when the workbook is a template with a data package (has the Package_Info sheet).'''
    try:
        wb = _open_stream(source_bytes)
    except Exception:  # noqa: BLE001 — not an xlsx: the caller reports it
        return False
    try:
        return PACKAGE_SHEET in wb.sheetnames
    finally:
        wb.close()


def _rows(ws):
    it = ws.iter_rows(values_only=True)
    header = None
    for row in it:
        if row and any(v is not None for v in row):
            header = [str(h).strip() if h is not None else f'Column{i + 1}' for i, h in enumerate(row)]
            break
    return header, it


def read_sheet(wb, name, required=True, keep_column=None, keep_values=None):
    '''Sheet with a header in its first non-empty row → DataFrame. With keep_column / keep_values only
    the rows whose value in that column is in keep_values are kept, while streaming (one set lookup
    per row: this runs ~160,000 times on the PO lines).'''
    if name not in wb.sheetnames:
        if required:
            raise KeyError(f'ไม่พบชีต {name} ใน template (ห้ามเปลี่ยนชื่อชีต)')
        return None
    header, it = _rows(wb[name])
    if header is None:
        return pd.DataFrame()
    width = len(header)
    out = []
    if keep_column is not None:
        idx = header.index(keep_column)
        wanted = {str(v) for v in keep_values}
        for row in it:
            v = row[idx] if idx < len(row) else None
            if v is not None and (v if isinstance(v, str) else to_text(v)).strip() in wanted:
                out.append(list(row[:width]) + [None] * (width - len(row)))
    else:
        for row in it:
            row = list(row[:width]) + [None] * (width - len(row))
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
                continue
            out.append(row)
    df = pd.DataFrame(out, columns=header, dtype=object)
    return df.map(clean)


def read_template(source):
    '''Template with package → dict of DataFrames, ready for ICInputs. Only the PO lines of the
    materials in Lines (and Reference Material Code) are kept.'''
    wb = _open_stream(source)
    try:
        info_df = read_sheet(wb, PACKAGE_SHEET)
        info = {to_text(r['Field']): clean(r['Value']) for _, r in info_df.iterrows() if to_text(r['Field'])}
        for k in ('Spending From', 'Spending To'):
            info[k] = to_date(info.get(k))
        header_df = read_sheet(wb, 'Header')
        header = {to_text(r['Field']): clean(r['Value']) for _, r in header_df.iterrows() if to_text(r['Field'])}
        lines = read_sheet(wb, 'Lines')
        for col in LINE_COLUMNS:
            if col not in lines:
                lines[col] = None
        materials = set(lines['Material Code'].map(to_text).dropna())
        materials |= set(lines['Reference Material Code'].map(to_text).dropna())
        plant = read_sheet(wb, 'dimPlant_Raw')
        po = read_sheet(wb, PO_SHEET, keep_column='Material Code', keep_values=materials)
        if po.empty:
            po = pd.DataFrame(columns=PACKAGE_PO_COLUMNS)
        plant_name = {to_text(r['Plant']): to_text(r['Plant Name']) for _, r in plant.iterrows()}
        commodity = to_text(info.get('Commodity'))
        po['Commodity Full'] = commodity
        po['Plant Name'] = po['Plant'].map(lambda p: plant_name.get(to_text(p)))
        return {
            'package': info,
            'header': header,
            'lines': lines,
            'plant_group': read_sheet(wb, 'dimPlantGroup'),
            'official_plant_group': read_sheet(wb, 'dimPlantGroup_Official', required=False),
            'masterinfo': normalize_masterinfo(read_sheet(wb, 'dimMasterInfo_Raw')),
            'uom': read_sheet(wb, 'dimUOM_Raw'),
            'plant': plant,
            'po_lines': normalize_po_lines(po[PO_COLUMNS]),
            'material_commodity': {to_text(r['Material Code']): to_text(r['Commodity'])
                                   for _, r in read_sheet(wb, 'dimMaterial', required=False, keep_column='Material Code', keep_values=materials).iterrows()}
            if 'dimMaterial' in wb.sheetnames else {},
        }
    finally:
        wb.close()
