'''What the web page does, without Streamlit, so pytest can test it.

    uploads (bytes) ──▶ run_uploads ──▶ RunResult
                                         ├─ result   tables (IC_Check, Summary, ...)
                                         ├─ files    {file name: xlsx bytes}  (IC Processed, for Contract)
                                         ├─ errors / warnings  counts from IC_Check
                                         └─ problem  text when the input could not be read

streamlit_app.py only draws widgets around this. Nothing here touches the network or the disk:
inputs and outputs are bytes in memory.
'''
from dataclasses import dataclass, field
import io

import pandas as pd

from .config import load_config
from .pipeline import process_ic
from .readers import normalize_po_lines, open_workbook, read_masterinfo, read_po_lines, read_table
from .template import build_template
from . import sample

LEVEL_ORDER = {'Error': 0, 'Warning': 1}


@dataclass
class RunResult:
    result: dict = field(default_factory=dict)
    files: dict = field(default_factory=dict)
    errors: int = 0
    warnings: int = 0
    problem: str | None = None

    @property
    def can_download(self):
        '''Outputs are offered only when IC_Check has no Error (Error ต้องแก้ก่อน).'''
        return self.problem is None and self.errors == 0 and bool(self.files)


def make_config(erp_system_id=None):
    '''ERP System ID belongs to the company system, so it is not in the public config; the user types it.'''
    overrides = {}
    if erp_system_id and str(erp_system_id).strip():
        overrides[('Contract', 'ERP System ID')] = str(erp_system_id).strip()
    return load_config(overrides=overrides or None)


def run_uploads(ic_bytes, po_bytes_list=(), master_bytes=None, erp_system_id=None, key_date=None):
    '''One run from the uploaded files. Reading problems (missing table, wrong file type) come back
    as RunResult.problem instead of an exception, so the page can show them in Thai.'''
    if not ic_bytes:
        return RunResult(problem='ยังไม่ได้อัปโหลดไฟล์ IC')
    try:
        cfg = make_config(erp_system_id)
        result, files = process_ic(
            ic_bytes,
            po_files=[io.BytesIO(b) for b in po_bytes_list if b],
            master_file=io.BytesIO(master_bytes) if master_bytes else None,
            cfg=cfg, key_date=key_date,
        )
    except KeyError as e:
        return RunResult(problem=str(e).strip("'\""))
    except Exception as e:  # noqa: BLE001 — any reading error must reach the user, not crash the page
        return RunResult(problem=f'อ่านไฟล์ไม่สำเร็จ: {type(e).__name__}: {e}')
    check = result['IC_Check']
    return RunResult(
        result=result, files=files,
        errors=int((check['Level'] == 'Error').sum()),
        warnings=int((check['Level'] == 'Warning').sum()),
    )


def sorted_check(check):
    '''Errors first, then warnings, then by line.'''
    if check.empty:
        return check
    order = check['Level'].map(LEVEL_ORDER).fillna(9)
    return (check.assign(_o=order)
                 .sort_values(['_o', 'Line No'], na_position='first', kind='stable')
                 .drop(columns='_o').reset_index(drop=True))


def summary_for_display(summary):
    '''Summary mixes text, dates and numbers in one column; show it as text so the table renders
    the same in every pandas / Arrow version (numbers with thousand separators).'''
    def fmt(measure, v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ''
        if isinstance(v, bool):
            return str(v)
        if str(measure).lstrip().startswith('%') and isinstance(v, (int, float)):
            return f'{v:.1%}'
        if isinstance(v, int):
            return f'{v:,}'
        if isinstance(v, float):
            return f'{v:,.2f}'
        if hasattr(v, 'isoformat'):
            return v.isoformat()
        return str(v)
    out = summary.copy()
    out['Value'] = [fmt(m, v) for m, v in zip(out['Measure'], out['Value'])]
    return out


SAMPLE_ERROR_LINES = {7, 8, 9}  # E10 commodity, E02 Order Unit, E04 FOC in costsaving/sample.py


def sample_inputs(with_errors=False):
    '''Fictitious template with data package (costsaving/sample.py) to try the page: (template, [], None).
    with_errors=False drops the lines that are wrong on purpose, so the downloads can be tried.'''
    lines = sample.LINES if with_errors else [x for x in sample.LINES if x['line'] not in SAMPLE_ERROR_LINES]
    return sample.template_bytes(lines=lines), [], None


def file_label(name):
    '''Download button label from the output file name.'''
    if name.endswith('_IC_Processed.xlsx'):
        return 'IC Processed (ไฟล์ยื่น SharePoint)'
    return 'for ' + name.rsplit('_for_', 1)[-1].removesuffix('.xlsx')


# ---------------------------------------------------------------- data team: build the template + package
@dataclass
class BuildResult:
    data: bytes = b''
    file_name: str = ''
    info: dict = field(default_factory=dict)
    problem: str | None = None


def run_build(master_bytes, po_bytes_list, commodity, spend_from=None, spend_to=None, sample_data=False):
    '''Master workbook (dimMasterInfo_Raw or MasterInfo_ActiveTH, dimUOM_Raw, dimPlant_Raw, dimPlantGroup)
    + PO-line CSVs → blank template with the data package of one commodity.'''
    try:
        cfg = load_config()
        if sample_data:
            m = sample.master_frames()
            masterinfo, uom, plant, plant_group = m['masterinfo'], m['uom'], m['plant'], m['plant_group']
            po = normalize_po_lines(sample.po_df())
        else:
            if not master_bytes:
                return BuildResult(problem='ยังไม่ได้อัปโหลดไฟล์ master')
            if not any(po_bytes_list):
                return BuildResult(problem='ยังไม่ได้อัปโหลด CSV PO line')
            wb = open_workbook(master_bytes)
            masterinfo = read_masterinfo(wb)
            uom, plant, plant_group = (read_table(wb, n) for n in ('dimUOM_Raw', 'dimPlant_Raw', 'dimPlantGroup'))
            po = read_po_lines([io.BytesIO(b) for b in po_bytes_list if b])
        data, info = build_template(cfg, masterinfo, uom, plant, plant_group, po, commodity, spend_from, spend_to)
    except KeyError as e:
        return BuildResult(problem=str(e).strip("'\""))
    except Exception as e:  # noqa: BLE001
        return BuildResult(problem=f'สร้าง template ไม่สำเร็จ: {type(e).__name__}: {e}')
    code = info['Package Version']
    return BuildResult(data=data, file_name=f'IC_Template_{code}.xlsx', info=info)


def benchmark(n_rows=160_000, progress=None):
    '''Speed test on fictitious data: a package with n_rows PO lines (≈ one commodity-year) is built,
    then processed like a keyed template. Returns [(step, seconds)] and the template size in MB.'''
    import random
    import time
    from datetime import date, timedelta
    cfg = load_config()
    rng = random.Random(1)
    base = sample.po_df()
    d0 = date(2025, 11, 7)
    rows = [['01_Medicine', f'T1{i % 40:02d}', f'TH{i % 40}', f'46{i:08d}', (d0 + timedelta(days=i % 364)).isoformat(),
             f'TST-X{i % 8000:05d}', f'T9{i % 300:08d}', 'BX10T', 'TAB', rng.randint(1, 500), round(rng.random() * 10000, 2)]
            for i in range(n_rows)]
    po = normalize_po_lines(pd.concat([base, pd.DataFrame(rows, columns=base.columns)], ignore_index=True))
    m = sample.master_frames()
    lines = sample.lines_df([x for x in sample.LINES if x['line'] not in SAMPLE_ERROR_LINES])
    timings = []
    t = time.perf_counter()
    data, info = build_template(cfg, m['masterinfo'], m['uom'], m['plant'], m['plant_group'], po, 'Medicine',
                                header=sample.HEADER, lines=lines)
    timings.append(('สร้าง template (ทีมข้อมูล)', time.perf_counter() - t))
    if progress:
        progress(0.5)
    t = time.perf_counter()
    r = run_uploads(data)
    timings.append(('ประมวลผล template (Sourcing)', time.perf_counter() - t))
    if progress:
        progress(1.0)
    return timings, len(data) / 2 ** 20, info['Spending Rows'], r.problem


def sample_blank_template():
    '''Blank template of the sample package, as Sourcing would download it from SharePoint.'''
    return sample.template_bytes(blank=True)


XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


# ---------------------------------------------------------------- PD setup mode
SETUP_REPORTS = ['ZMMP1004', 'ZMMR1019', 'ME1L', 'ME3L']


@dataclass
class SetupRunResult:
    result: dict = field(default_factory=dict)
    files: dict = field(default_factory=dict)
    errors: int = 0
    warnings: int = 0
    problem: str | None = None

    @property
    def can_download(self):
        '''Upload files are offered only when Upload_Check has no Error (ห้ามดาวน์โหลดเมื่อมี Error).'''
        return self.problem is None and self.errors == 0 and bool(self.files)


def read_vendor_tax(data):
    '''Optional table Vendor ID | Vendor Tax Code (xlsx first sheet or csv), header in the first row.'''
    from .setup import _rows_of
    rows = _rows_of(data)
    header = [str(v).strip() if v is not None else '' for v in rows[0]] if rows else []
    if 'Vendor ID' not in header or 'Vendor Tax Code' not in header:
        raise KeyError('ตาราง Tax ของ Vendor ต้องมีคอลัมน์ Vendor ID และ Vendor Tax Code ในแถวแรก')
    i, j = header.index('Vendor ID'), header.index('Vendor Tax Code')
    return pd.DataFrame([(r[i], r[j]) for r in rows[1:] if r and r[i] is not None], columns=['Vendor ID', 'Vendor Tax Code'])


def run_setup_uploads(ic_files, report_bytes, vendor_tax_bytes=None, tax_plants_text='', stamp=None):
    '''ic_files = [(file name, bytes, Agreement ID typed, contract numbers typed)];
    report_bytes = {'ZMMP1004' | 'ZMMR1019' | 'ME1L' | 'ME3L': bytes or None}.'''
    from datetime import datetime, timezone
    from .setup import SetupFile, parse_contract_numbers, read_processed, read_sap_report, run_setup
    from . import writers
    if not ic_files:
        return SetupRunResult(problem='ยังไม่ได้อัปโหลดไฟล์ IC Processed')
    try:
        cfg = load_config()
        files = []
        for name, data, agreement, contracts in ic_files:
            f = read_processed(name, data)
            f.agreement_id = (str(agreement).strip() or None) if agreement is not None else None
            f.contract_numbers = parse_contract_numbers(contracts)
            files.append(f)
        reports, problems = {}, []
        for name in SETUP_REPORTS:
            data = report_bytes.get(name)
            if not data:
                reports[name] = None
                continue
            df, problem = read_sap_report(data, name, cfg)
            if problem:
                problems.append(problem)
            reports[name] = df
        if problems:
            return SetupRunResult(problem='; '.join(problems))
        vendor_tax = read_vendor_tax(vendor_tax_bytes) if vendor_tax_bytes else None
        tax_plants = [p.strip() for p in str(tax_plants_text or '').replace(';', ',').split(',') if p.strip()]
        result = run_setup(files, reports, cfg, tax_plants=tax_plants, vendor_tax=vendor_tax)
    except KeyError as e:
        return SetupRunResult(problem=str(e).strip("'\""))
    except Exception as e:  # noqa: BLE001 — any reading error must reach the user, not crash the page
        return SetupRunResult(problem=f'อ่านไฟล์ไม่สำเร็จ: {type(e).__name__}: {e}')
    check = result['Upload_Check']
    stamp = stamp or datetime.now(timezone.utc).strftime('%Y%m%d-%H%M')
    return SetupRunResult(result=result, files=writers.write_setup_files(result, stamp),
                          errors=int((check['Level'] == 'Error').sum()),
                          warnings=int((check['Level'] == 'Warning').sum()))


def sample_setup_inputs(with_error=False):
    '''Fictitious PD round (costsaving/sample_pd.py). with_error=True adds an IC Processed file edited after
    processing (checksum Error), to try the download block.'''
    from . import sample_pd
    files, reports = sample_pd.round_inputs()
    if with_error:
        files.append(('IC_TEST_EDITED_IC_Processed.xlsx', sample_pd.processed_bytes('MED-20261201-002', edited=True),
                      'MED-20261201-099', ''))
    return files, reports, sample_pd.vendor_tax_df(), sample_pd.TAX_PLANT
