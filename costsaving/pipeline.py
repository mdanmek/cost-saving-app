'''One call from uploaded files to output files. The browser UI (stlite) and the tests both use this.

    IC workbook (Header, Lines, CT_Leadtime) ─┐
    master workbook (dimUOM_Raw, dimPlant_Raw, ├─▶ run_ic ─▶ {file name: xlsx bytes}
       dimPlantGroup, dimMasterInfo_Raw)       │
    PO-line CSV(s) prc_spending_sap_{year}  ───┘

The master tables can sit in the IC workbook itself (the current Price_Change_Template already holds
them); then the master workbook is not needed. Nothing is written to disk: input and output are bytes.
'''
from .config import load_config, to_text
from .ic import ICInputs, run_ic
from .readers import normalize_po_lines, open_workbook, read_header, read_masterinfo, read_po_lines, read_table
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


def load_ic_inputs(ic_file, po_files=(), master_file=None, po_table=None):
    '''ic_file / master_file = path, file object or bytes. po_files = list of CSV sources.
    po_table = name of a PO-line table inside the IC workbook (used by the test workbook).'''
    ic_wb = open_workbook(ic_file)
    workbooks = [ic_wb] + ([open_workbook(master_file)] if master_file is not None else [])
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
        leadtime=read_table(ic_wb, 'CT_Leadtime', required=False),
    )


def output_files(result, cfg, project_name=None):
    stem = ''.join(ch if ch.isalnum() or ch in '-_ ' else '_' for ch in (project_name or 'IC')).strip().replace(' ', '_')
    return {
        f'{stem}_for_PD.xlsx': writers.write_for_pd(result),
        f'{stem}_for_Contract.xlsx': writers.write_for_contract(result),
        f'{stem}_for_Collect.xlsx': writers.write_for_collect(result, cfg.const('IC', 'Collect Table Name')),
        f'{stem}_for_Summary.xlsx': writers.write_for_summary(result),
    }


def process_ic(ic_file, po_files=(), master_file=None, po_table=None, cfg=None):
    '''Returns (result tables, output files). The caller decides whether to block on errors:
    result['IC_Check'] with Level = Error must be empty before the file is submitted.'''
    cfg = cfg or load_config()
    inputs = load_ic_inputs(ic_file, po_files, master_file, po_table)
    result = run_ic(inputs, cfg)
    return result, output_files(result, cfg, to_text(inputs.header.get('Project Name')))
