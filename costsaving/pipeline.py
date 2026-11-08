'''One call from uploaded files to output files. The browser UI (stlite) and the tests both use this.

    IC workbook (Header, Lines)             ─┐
    master workbook (dimUOM_Raw, dimPlant_Raw, ├─▶ run_ic ─▶ {file name: xlsx bytes}
       dimPlantGroup, dimMasterInfo_Raw)       │
    PO-line CSV(s) prc_spending_sap_{year}  ───┘

The master tables can sit in the IC workbook itself (the current Price_Change_Template already holds
them); then the master workbook is not needed. Nothing is written to disk: input and output are bytes.
'''
from datetime import datetime, timezone
import hashlib
import io
import json

from .config import clean, load_config, to_text
from .ic import ICInputs, run_ic
from .readers import normalize_po_lines, open_workbook, read_header, read_masterinfo, read_po_lines, read_table
from .legacy import is_legacy, read_legacy
from .template import has_package, read_template
from . import writers

MASTER_TABLES = ['dimUOM_Raw', 'dimPlant_Raw', 'dimPlantGroup']


def _find_table(workbooks, name, required=True):
    for wb in workbooks:
        df = read_table(wb, name, required=False)
        if df is not None:
            return df
    if required:
        raise KeyError(f'ไม่พบตาราง {name} ทั้งในไฟล์ IC และไฟล์ master')
    return None


def _find_masterinfo(workbooks):
    errors = []
    for wb in workbooks:
        try:
            return read_masterinfo(wb)
        except KeyError as e:
            errors.append(str(e))
    raise KeyError('ไม่พบตาราง dimMasterInfo_Raw หรือ MasterInfo_ActiveTH')


def load_ic_inputs(ic_file, po_files=(), master_file=None, po_table=None, key_date=None):
    '''ic_file / master_file = path, file object or bytes. po_files = list of CSV sources.
    po_table = name of a PO-line table inside the IC workbook (used by the test workbook).
    ic_file may be the current IC file (Price_Change_Template, noCT_Input / CT_Input): it is read as
    Sourcing keys it today (legacy.py), and key_date is needed because that file has none.'''
    ic_wb = open_workbook(ic_file)
    workbooks = [ic_wb] + ([open_workbook(master_file)] if master_file is not None else [])
    messages = []
    if is_legacy(ic_wb):
        header, lines, messages = read_legacy(ic_wb, key_date)
    else:
        header = read_header(ic_wb)
        lines = read_table(ic_wb, 'Lines')
    materials = set(lines['Material Code'].map(to_text).dropna())
    if 'Reference Material Code' in lines:
        materials |= set(lines['Reference Material Code'].map(to_text).dropna())
    if po_table:
        po = _find_table(workbooks, po_table)
        po = normalize_po_lines(po)
    else:
        po = read_po_lines(list(po_files), materials=materials)
    return ICInputs(
        header=header, lines=lines,
        masterinfo=_find_masterinfo(workbooks),
        uom=_find_table(workbooks, 'dimUOM_Raw'),
        plant=_find_table(workbooks, 'dimPlant_Raw'),
        plant_group=_find_table(workbooks, 'dimPlantGroup'),
        po_lines=po,
        messages=messages,
    )


def _bytes(source):
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, 'read'):
        data = source.read()
        if hasattr(source, 'seek'):
            source.seek(0)
        return data
    with open(source, 'rb') as f:
        return f.read()


def load_template_inputs(ic_bytes, po_files=(), master_file=None):
    '''Template with data package (D56). A master workbook or PO CSVs, when given, replace the
    package tables (advanced use: test a new package before it is published).'''
    t = read_template(ic_bytes)
    inputs = ICInputs(header=t['header'], lines=t['lines'], masterinfo=t['masterinfo'], uom=t['uom'],
                      plant=t['plant'], plant_group=t['plant_group'], po_lines=t['po_lines'],
                      package=t['package'], material_commodity=t['material_commodity'],
                      official_plant_group=t['official_plant_group'])
    if master_file is not None:
        wb = open_workbook(master_file)
        inputs.masterinfo = read_masterinfo(wb)
        inputs.uom = read_table(wb, 'dimUOM_Raw')
        inputs.plant = read_table(wb, 'dimPlant_Raw')
    if po_files:
        materials = set(inputs.lines['Material Code'].map(to_text).dropna())
        materials |= set(inputs.lines['Reference Material Code'].map(to_text).dropna())
        inputs.po_lines = read_po_lines(list(po_files), materials=materials)
        inputs.package = dict(inputs.package or {}, **{'Package Version': (inputs.package or {}).get('Package Version', '') + ' + CSV'})
    return inputs


def canonical_text(value):
    '''One text per value, whatever type it comes back as from Excel: date and datetime at midnight → 2026-12-01,
    100.0 → 100, codes and text stripped. Written then read back, the processed file must give the same checksum.'''
    from datetime import date, datetime
    import math
    value = clean(value)
    if value is None:
        return None
    if hasattr(value, 'to_pydatetime'):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        return value.date().isoformat() if value.time() == datetime.min.time() else value.isoformat(sep=' ')
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)) or hasattr(value, 'item'):
        number = float(value)
        if math.isnan(number):
            return None
        return str(int(number)) if number.is_integer() else repr(number)
    return str(value).strip()


def input_checksum(inputs):
    '''SHA-256 of what Sourcing keyed (Header, Lines, dimPlantGroup) + package version.
    Written in the processed file, so the PD mode can tell whether the file was edited after approval.
    Values go through canonical_text, and empty trailing rows / columns do not count.'''
    def table(df):
        if df is None:
            return []
        rows = [[canonical_text(v) for v in row] for row in df.itertuples(index=False)]
        return [[str(c)] + [r[i] for r in rows] for i, c in enumerate(df.columns)
                if any(r[i] is not None for r in rows)]
    payload = {
        'header': {str(k): canonical_text(v) for k, v in sorted(inputs.header.items(), key=lambda kv: str(kv[0]))},
        'lines': table(inputs.lines), 'plant_group': table(inputs.plant_group),
        'package': to_text((inputs.package or {}).get('Package Version')),
    }
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()


def processed_stamp(inputs, cfg, processed_at=None):
    p = inputs.package or {}
    return [
        ('Template Version', cfg.const('IC', 'Template Version')),
        ('Package Version', to_text(p.get('Package Version'))),
        ('Spending From', p.get('Spending From')), ('Spending To', p.get('Spending To')),
        ('Processed (UTC)', (processed_at or datetime.now(timezone.utc)).strftime('%Y-%m-%d %H:%M')),
        ('Input Checksum (SHA-256)', input_checksum(inputs)),
    ]


def output_files(result, cfg, inputs, processed_at=None):
    '''IC_Processed = the one file submitted in SharePoint (D57); for_Contract = Ariba import.'''
    project_name = to_text(inputs.header.get('Project Name'))
    stem = ''.join(ch if ch.isalnum() or ch in '-_ ' else '_' for ch in (project_name or 'IC')).strip().replace(' ', '_')
    return {
        f'{stem}_IC_Processed.xlsx': writers.write_processed(result, inputs, processed_stamp(inputs, cfg, processed_at),
                                                             cfg.const('IC', 'Collect Table Name')),
        f'{stem}_for_Contract.xlsx': writers.write_for_contract(result),
    }


def process_ic(ic_file, po_files=(), master_file=None, po_table=None, cfg=None, processed_at=None, key_date=None):
    '''Returns (result tables, output files). The caller decides whether to block on errors:
    result['IC_Check'] with Level = Error must be empty before the file is submitted.

    ic_file = template with data package (D56), or an IC workbook without package plus a master
    workbook / PO CSVs (the earlier format, kept for comparing with the current IC file, D57 Q7).
    key_date = Key Date for the current IC file, which has none (ignored for the template).'''
    cfg = cfg or load_config()
    data = _bytes(ic_file)
    if po_table is None and has_package(data):
        inputs = load_template_inputs(data, po_files, master_file)
    else:
        inputs = load_ic_inputs(io.BytesIO(data), po_files, master_file, po_table, key_date)
    result = run_ic(inputs, cfg)
    result['inputs'] = inputs
    return result, output_files(result, cfg, inputs, processed_at)
