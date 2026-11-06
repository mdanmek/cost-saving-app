'''What the web page does, without Streamlit, so pytest can test it.

    uploads (bytes) ──▶ run_uploads ──▶ RunResult
                                         ├─ result   tables (IC_Check, Summary, ...)
                                         ├─ files    {file name: xlsx bytes}  (the 4 downloads)
                                         ├─ errors / warnings  counts from IC_Check
                                         └─ problem  text when the input could not be read

streamlit_app.py only draws widgets around this. Nothing here touches the network or the disk:
inputs and outputs are bytes in memory.
'''
from dataclasses import dataclass, field
import io
import zipfile

import pandas as pd

from .config import load_config
from .pipeline import process_ic
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


def run_uploads(ic_bytes, po_bytes_list=(), master_bytes=None, erp_system_id=None):
    '''One run from the uploaded files. Reading problems (missing table, wrong file type) come back
    as RunResult.problem instead of an exception, so the page can show them in Thai.'''
    if not ic_bytes:
        return RunResult(problem='ยังไม่ได้อัปโหลดไฟล์ IC')
    try:
        cfg = make_config(erp_system_id)
        result, files = process_ic(
            io.BytesIO(ic_bytes),
            po_files=[io.BytesIO(b) for b in po_bytes_list if b],
            master_file=io.BytesIO(master_bytes) if master_bytes else None,
            cfg=cfg,
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


def zip_files(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return buf.getvalue()


SAMPLE_ERROR_LINES = {7, 8, 9}  # E10 commodity, E02 Order Unit, E04 FOC in costsaving/sample.py


def sample_inputs(with_errors=False):
    '''Fictitious files (costsaving/sample.py) to try the page: (ic, [po csv], master).
    with_errors=False drops the lines that are wrong on purpose, so the downloads can be tried.'''
    lines = sample.LINES if with_errors else [x for x in sample.LINES if x['line'] not in SAMPLE_ERROR_LINES]
    return (sample.ic_workbook_bytes(lines=lines, with_masters=False), [sample.po_csv_bytes()],
            sample.master_workbook_bytes())


XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
