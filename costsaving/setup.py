'''PD setup mode (D39, D56, D63): many IC Processed files + current SAP reports → upload files.

Same rules as drafts/powerquery/pd_inputs.m, pd_upload_source_list.m, pd_upload_info_record.m and
pd_upload_foc.m, with the decisions of 2026-11-08 (ME1L = current Info Records, ME3L = contracts,
Agreement ID and contract numbers typed by PD on the page, plants of the commodity from IC Processed).

    IC Processed (1..n) ── Out_PD, dimPlant_Raw, checksum ─┐
    typed per file: Agreement ID, SAP Contract No.         ├─▶ setup_lines (grain: Out_PD row = Line × Plant scope)
    ME3L  contracts (Contract No. × Item)                 ─┤        + Contract No. / Item, category
    ME1L  current Info Records (Vendor × Material × Plant) ─┘
             │
             ├─▶ Source List  (Material × Plant × Vendor; whole vendor set of a Material × Plant, + ZMMP1004)
             ├─▶ Info Record  (Supplier × Material × Plant × Info Record Category)
             ├─▶ FOC          (Contract × Material × Plant-or-blank × Order Unit × Valid from × Scale no., + ZMMR1019)
             ├─▶ PD manual    (PD Types not uploaded, e.g. Inactive / Quit)
             └─▶ Upload_Check (Error = nothing may be downloaded · Warning: Issue = row left out, Note = row goes)

Every join below is on a key made unique first (dict with setdefault, or an explicit duplicate check),
so a SAP report never adds rows to the lines.
'''
from dataclasses import dataclass, field
from datetime import date, datetime
import io

import pandas as pd
from openpyxl import load_workbook

from .config import Config, clean, is_blank, to_bool, to_number, to_text
from .ic import records, to_date
from .prices import parse_foc

SOURCE_LIST_COLUMNS = [
    'Material', 'Plant', 'Valid from', 'Valid to', 'Vendor', 'Purchase Org.',
    'Order Unit', 'Agreement', 'Agreement Item', 'Fixed vendor', 'Blocked', 'Materials Planning(MRP)',
]
INFO_RECORD_COLUMNS = [
    'Supplier', 'Old Vendor Code', 'Supplier Name', 'Material Code', 'Old Item Code',
    'Material Description', 'Purchasing Org', 'Plant', 'Info Record Category', 'Vendor Material No.',
    'Supplier Subrange', 'Sale Person', 'Telephone', 'Order Unit',
    'Denominator for Conversion of Order Unit to Base Unit', 'Numerator for Conversion',
    'Var. Order Unit', 'Pl. Deliv. Time', 'Purchasing Group',
    'Standard Qty (per order unit)', 'Minimum Qty (per order unit)',
    'Tol. Overdl', 'Conf. Ctr', 'Net Price', 'Currency', 'Price Unit', 'Order Price Unit',
    'Valid From', 'Valid to', 'MFG_ID', 'Tax code', 'Local / Import',
    'Registration No.', 'Registration Type', 'Trade Product Use (TPU)', 'Special Feature',
    'Tradename', 'Client',
]
FOC_COLUMNS = [
    'Option', 'Contract No.', 'Material', 'Plant', 'Order Unit', 'Valid from', 'Valid to', 'Scale no.',
    'Minimum qty', 'FreeGdsQty', 'Base Unit', 'Add. quantity', 'Add. qty unit', 'CalculType', 'FreeGoodsCat.',
]
FOC_SHEETS = {('A', True): 'Cont. Add with Plant', ('A', False): 'Cont. Add No Plant',
              ('D', True): 'Cont. Delete with plant', ('D', False): 'Cont. Delete No plant'}
ME01_COLUMNS = ['Material Code', 'Plant', 'Vendor ID', 'SAP Valid from', 'SAP Valid to', 'SAP Agreement', 'Note']
MANUAL_COLUMNS = ['IC File', 'Agreement ID', 'Line No', 'PD Type', 'PD Action', 'PD Action Description',
                  'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'Order Unit', 'Plant Scope',
                  'Plant', 'Effective Date', 'Valid Date', 'Task']
CHECK_COLUMNS = ['Level', 'Type', 'Upload', 'Upload Sheet', 'IC File', 'Agreement ID', 'Line No', 'Material Code',
                 'Plant', 'Vendor ID', 'Message']
CONTRACT_MATCH_ISSUE = {
    'No contract': 'ยังไม่มีเลข Contract: ไม่พบใน ME3L และ PD ไม่ได้กรอก',
    'Ambiguous': 'พบหลาย Contract Item ที่ตรง Vendor × Material × Order Unit × Plant × ช่วงวันที่: กรอกเลข Contract ของไฟล์นี้',
    'Not found': 'เลข Contract ที่กรอกมีหลายเลข แต่ไม่พบ item ของรายการนี้ใน ME3L',
}
REPORT_LABEL = {'ZMMP1004': 'ZMMP1004 (Source List)', 'ZMMR1019': 'ZMMR1019 (FOC)',
                'ME1L': 'ME1L (Info Record)', 'ME3L': 'ME3L (Contract)'}


# ---------------------------------------------------------------- small helpers
def vendor_key(value):
    '''Vendor codes differ between reports: '0900000123 NAME' (ME1L), '900000123' (ZMMP1004), '900000123'
    (IC). Key = first token, leading zeros removed. Used for joins only; outputs keep the code as given.'''
    text = to_text(value)
    if text is None:
        return None
    token = text.replace('_', ' ').split()[0]
    return token.lstrip('0') or '0'


def code_key(value):
    '''Material / contract numbers: text, leading zeros kept (SAP shows them the same in every report).'''
    return to_text(value)


def any_date(value):
    '''datetime, date, '01.12.2026' (SAP export as text) or '2026-12-01' → date.'''
    value = clean(value)
    if isinstance(value, str) and len(value) == 10 and value[2] == '.' and value[5] == '.':
        try:
            return datetime.strptime(value, '%d.%m.%Y').date()
        except ValueError:
            return None
    return to_date(value)


def sap_date(value):
    '''dd.MM.yyyy as text (D50, P22); always the Gregorian year whatever the PC locale.'''
    d = any_date(value)
    return None if d is None else f'{d.day:02d}.{d.month:02d}.{d.year:04d}'


def is_true(value):
    return to_bool(value) is True or to_text(value) in ('/', '✓')


def join_messages(messages):
    text = '; '.join(m for m in messages if m)
    return text or None


def overlaps(a_from, a_to, b_from, b_to):
    '''Periods overlap; a blank bound is open.'''
    return (a_from is None or b_to is None or a_from <= b_to) and (a_to is None or b_from is None or a_to >= b_from)


# ---------------------------------------------------------------- reading the inputs
def _rows_of(source):
    '''xlsx (first sheet) or csv → list of row tuples. source = bytes or path.'''
    data = source if isinstance(source, (bytes, bytearray)) else open(source, 'rb').read()
    if data[:2] == b'PK':
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        try:
            ws = wb['Data'] if 'Data' in wb.sheetnames else wb.worksheets[0]
            return [tuple(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    df = pd.read_csv(io.BytesIO(data), header=None, dtype=str, keep_default_na=False, encoding='utf-8-sig')
    return [tuple(r) for r in df.itertuples(index=False)]


def read_sap_report(source, report, cfg: Config):
    '''SAP export → DataFrame with the logical columns of config/sap_reports.csv. The header row is found
    (title rows above it are skipped, e.g. ZMMP1004 row 6, ZMMR1019 row 5); a repeated header name takes
    its first column. Rows with a blank first mapped column (footer, empty lines) are dropped.
    Returns (DataFrame or None, problem text or None).'''
    mapping = [(r['Field'], r['Header']) for r in records(cfg.sap_reports) if r['Report'] == report]
    wanted = [h for _, h in mapping]
    rows = _rows_of(source)
    for i, row in enumerate(rows[:60]):
        names = [str(v).strip() if v is not None else '' for v in row]
        if all(h in names for h in wanted):
            pos = {h: names.index(h) for h in wanted}
            out = []
            for r in rows[i + 1:]:
                values = {f: clean(r[pos[h]]) if pos[h] < len(r) else None for f, h in mapping}
                if all(v is None for v in values.values()):
                    continue
                out.append(values)
            return pd.DataFrame(out, columns=[f for f, _ in mapping], dtype=object), None
    missing = sorted(set(wanted) - {str(v).strip() for row in rows[:60] for v in row if v is not None})
    return None, (f"{REPORT_LABEL.get(report, report)}: ไม่พบหัวคอลัมน์ {', '.join(missing)} "
                  '(แก้ชื่อหัวคอลัมน์ได้ที่ config/sap_reports.csv)')


@dataclass
class SetupFile:
    '''One IC Processed file of the round, with what PD typed for it on the page.'''
    name: str
    out_pd: pd.DataFrame
    plants: pd.DataFrame = None          # dimPlant_Raw: plants an All Plant line is split into (S1, Q15)
    agreement_id: str = None              # SharePoint property, typed by PD (D22, Q7)
    contract_numbers: list = field(default_factory=list)  # typed by PD; blank → found in ME3L (Q7, Q10)
    checksum_ok: bool = True
    problems: list = field(default_factory=list)          # file-level errors found while reading


def read_processed(name, data):
    '''IC Processed bytes → SetupFile. The checksum of what Sourcing keyed is computed again: a different
    value means the file was edited after it was processed (Error, Q5).'''
    from .ic import ICInputs
    from .pipeline import input_checksum
    from .readers import open_workbook, read_header, read_table
    try:
        wb = open_workbook(data)
    except Exception as e:  # noqa: BLE001 — not an xlsx
        return SetupFile(name=name, out_pd=pd.DataFrame(), checksum_ok=False,
                         problems=[f'เปิดไฟล์ไม่ได้ ({type(e).__name__}) ไม่ใช่ไฟล์ IC Processed'])
    out_pd = read_table(wb, 'Out_PD', required=False)
    stamp_df = read_table(wb, 'Processed', required=False)
    if out_pd is None or stamp_df is None:
        return SetupFile(name=name, out_pd=pd.DataFrame(), checksum_ok=False,
                         problems=['ไม่ใช่ไฟล์ IC Processed (ไม่มีตาราง Out_PD / Processed) ประมวลผลในโหมด Sourcing ก่อน'])
    stamp = {to_text(r['Field']): clean(r['Value']) for r in records(stamp_df)}
    missing = [t for t in ('Header', 'Lines') if read_table(wb, t, required=False) is None]
    if missing:  # the checksum cannot be computed: this file only is refused, the others of the round go on
        return SetupFile(name=name, out_pd=pd.DataFrame(), checksum_ok=False,
                         problems=[f"ไฟล์ IC Processed ไม่มีตาราง {', '.join(missing)} ตรวจ checksum ไม่ได้ ให้ Sourcing ประมวลผลใหม่"])
    keyed = ICInputs(header=read_header(wb), lines=read_table(wb, 'Lines'), masterinfo=None, uom=None, plant=None,
                     plant_group=read_table(wb, 'dimPlantGroup', required=False), po_lines=None,
                     package={'Package Version': stamp.get('Package Version')})
    ok = input_checksum(keyed) == stamp.get('Input Checksum (SHA-256)')
    problems = [] if ok else ['Input Checksum ไม่ตรง: ไฟล์ถูกแก้หลังประมวลผล ให้ Sourcing ประมวลผลใหม่และยื่น approve อีกครั้ง']
    return SetupFile(name=name, out_pd=out_pd, plants=read_table(wb, 'dimPlant_Raw', required=False),
                     checksum_ok=ok, problems=problems)


def parse_contract_numbers(text):
    '''"8100000001, 8100000002" (or ; space new line) → list without blanks or repeats.'''
    out = []
    for part in str(text or '').replace(';', ',').replace('\n', ',').replace(' ', ',').split(','):
        part = part.strip()
        if part and part not in out:
            out.append(part)
    return out


# ---------------------------------------------------------------- the run
class _Check:
    def __init__(self):
        self.rows = []

    def add(self, level, kind, upload, sheet, message, row=None, plant=None):
        row = row or {}
        self.rows.append({
            'Level': level, 'Type': kind, 'Upload': upload, 'Upload Sheet': sheet,
            'IC File': row.get('IC File'), 'Agreement ID': row.get('Agreement ID'), 'Line No': row.get('Line No'),
            'Material Code': row.get('Material Code'), 'Plant': plant if plant is not None else row.get('Plant'),
            'Vendor ID': row.get('Vendor ID'), 'Message': message,
        })

    def frame(self):
        df = pd.DataFrame(self.rows, columns=CHECK_COLUMNS)
        if df.empty:
            return df
        df = df.drop_duplicates().copy()
        df['_l'] = df['Level'].map({'Error': 0, 'Warning': 1}).fillna(9)
        df['_t'] = df['Type'].map({'Error': 0, 'Issue': 1, 'Manual': 2, 'Note': 3}).fillna(9)
        return (df.sort_values(['_l', '_t', 'Upload', 'IC File', 'Line No'], kind='stable', na_position='first')
                  .drop(columns=['_l', '_t']).reset_index(drop=True))


def run_setup(files, reports, cfg: Config, tax_plants=(), vendor_tax=None):
    '''files = [SetupFile]; reports = {'ZMMP1004' | 'ZMMR1019' | 'ME1L' | 'ME3L': DataFrame from read_sap_report
    or None}; tax_plants = plant codes whose consignment Info Record takes the plant tax code (typed on the page,
    not in the public config, Q16); vendor_tax = DataFrame Vendor ID | Vendor Tax Code (optional).
    Returns a dict of tables (uploads, bases, Upload_Check, Upload_PD_Manual, setup_lines).'''
    chk = _Check()
    upload_types = set(cfg.const_list('PD', 'Upload PD Types'))
    no_setup = cfg.const('PD', 'No Setup PD Type')
    commodity = {r['Commodity']: r for r in records(cfg.dim_commodity)}
    method_mode = {r['Method']: to_text(r.get('Source List Mode')) for r in records(cfg.dim_method)}
    pd_info = {r['PD Action']: r for r in records(cfg.dim_pd)}
    cat_by_name = {r['Item Category']: r for r in records(cfg.item_category)}
    cat_by_me1l = {to_text(r['ME1L Text']): r['Item Category'] for r in records(cfg.item_category)}
    purch_group = {to_text(r['Material Type']): to_text(r['Purchasing Group']) for r in records(cfg.purchasing_group)}
    tax_map = {(to_text(r['Info Record Category']), to_text(r['Vendor Tax Code'])): to_text(r['Tax code'])
               for r in records(cfg.tax_code_map)}
    tax_plants = {to_text(p) for p in tax_plants if to_text(p)}
    vendor_tax_code = {}
    for r in records(vendor_tax):
        vendor_tax_code.setdefault(vendor_key(r.get('Vendor ID')), to_text(r.get('Vendor Tax Code')))

    # ---- files: file-level checks (Error blocks every download)
    seen_agreements, seen_names = {}, set()
    for f in files:
        for p in f.problems:
            chk.add('Error', 'Error', 'IC File', f.name, p, {'IC File': f.name, 'Agreement ID': f.agreement_id})
        if f.name in seen_names:
            chk.add('Error', 'Error', 'IC File', f.name, 'อัปโหลดไฟล์นี้ซ้ำ', {'IC File': f.name})
        seen_names.add(f.name)
        if f.agreement_id:
            if f.agreement_id in seen_agreements:
                chk.add('Error', 'Error', 'IC File', f.name,
                        f'Agreement ID {f.agreement_id} ซ้ำกับไฟล์ {seen_agreements[f.agreement_id]}', {'IC File': f.name})
            seen_agreements.setdefault(f.agreement_id, f.name)
        else:
            chk.add('Warning', 'Note', 'IC File', f.name, 'ไม่ได้กรอก Agreement ID: Telephone ของ Info Record จะว่าง (D37)',
                    {'IC File': f.name})

    # ---- reports
    for name, df in reports.items():
        if df is None:
            level = 'Error' if name == 'ZMMP1004' else 'Warning'
            effect = {
                'ZMMP1004': 'สร้าง Source List ไม่ได้ เพราะต้องรู้ Vendor เดิมทั้งชุด (S2) — ถ้าไม่มี Source List เดิม ให้อัปโหลดรายงานที่ว่าง',
                'ZMMR1019': 'ไม่ได้ตรวจ FOC เดิม จึงไม่มีแถวลบ FOC',
                'ME1L': 'ไม่รู้ว่ารายการไหนเป็น consignment (ถือเป็น Standard) และใช้ค่าเริ่มต้นแทน Info Record ปัจจุบัน',
                'ME3L': 'ใช้เฉพาะเลข Contract ที่กรอก และหา Contract Item ไม่ได้ (Source List จะรอ)',
            }[name]
            chk.add(level, 'Error' if level == 'Error' else 'Note', REPORT_LABEL[name], name, f'ไม่มีรายงาน {REPORT_LABEL[name]}: {effect}')

    me1l = {}
    me1l_vm = {}
    dup_me1l = set()
    for r in records(reports.get('ME1L')):
        key = (vendor_key(r['Vendor ID']), code_key(r['Material Code']), to_text(r['Plant']))
        if key in me1l:
            dup_me1l.add(key)
        me1l.setdefault(key, r)
        me1l_vm.setdefault(key[:2], []).append(r)
    contract_types = set(cfg.const_list('PD', 'Contract Doc Types'))
    me3l, seen_items = [], set()
    for r in records(reports.get('ME3L')):
        key = (code_key(r.get('Contract No.')), to_text(r.get('Contract Item')))  # grain: Contract No. × Item
        if key in seen_items or (contract_types and to_text(r.get('Doc Type')) not in contract_types) \
                or not is_blank(r.get('Deletion Indicator')):
            continue
        seen_items.add(key)
        me3l.append(r)
    # index by Vendor × Material × Order Unit, values normalized once: a round of many files × a large ME3L
    # would otherwise compare every line with every contract item (lines × items)
    me3l_index = {}
    for c in me3l:
        c = dict(c, **{'_plant': to_text(c['Plant']), '_from': any_date(c['Valid from']), '_to': any_date(c['Valid to']),
                       '_no': code_key(c['Contract No.'])})
        me3l_index.setdefault((vendor_key(c['Vendor ID']), code_key(c['Material Code']), to_text(c['Order Unit'])), []).append(c)

    def category_of(row, plant):
        '''Item Category of a line at a plant from ME1L (Q11); no row → Standard.'''
        hit = me1l.get((row['Vendor Key'], row['Material Code'], plant))
        text = to_text(hit.get('Item Category Text')) if hit else None
        return cat_by_me1l.get(text, 'Standard')

    def line_category(row):
        '''By Plant → the category at that plant. All Plant → Consignment only when ME1L knows the
        vendor × material and every plant of it is consignment (no contract / FOC for those).'''
        if row['Plant']:
            return category_of(row, row['Plant'])
        hits = me1l_vm.get((row['Vendor Key'], row['Material Code']), [])
        cats = {cat_by_me1l.get(to_text(h.get('Item Category Text')), 'Standard') for h in hits}
        return cats.pop() if len(cats) == 1 else 'Standard'

    # ---- setup lines: grain = Out_PD row (Line × Plant scope) of every file
    lines = []
    for f in files:
        if f.problems and not f.checksum_ok and f.out_pd.empty:
            continue
        plants = [{'Plant': to_text(r.get('Plant')), 'Plant Name': to_text(r.get('Plant Name'))}
                  for r in records(f.plants) if to_text(r.get('Plant'))]
        for r in records(f.out_pd):
            row = {k: clean(v) for k, v in r.items()}
            for k in ('Material Code', 'Vendor ID', 'Order Unit', 'Plant', 'Final FOC', 'PD Type', 'Commodity', 'Method'):
                row[k] = to_text(row.get(k))
            row['Line No'] = int(to_number(row['Line No'])) if to_number(row.get('Line No')) is not None else None
            row['Effective Date'], row['Valid Date'] = any_date(row.get('Effective Date')), any_date(row.get('Valid Date'))
            row['Per for Setup'] = to_number(row.get('Per for Setup')) or 1
            row['Final Price'] = to_number(row.get('Final Price'))
            row['Conversion'] = to_number(row.get('Conversion'))
            row.update({'IC File': f.name, 'Agreement ID': f.agreement_id, 'Contract List': list(f.contract_numbers),
                        'Vendor Key': vendor_key(row['Vendor ID']), 'File Plants': plants,
                        'Plant Key': row['Plant'] or ''})
            com = commodity.get(row['Commodity'], {})
            row['Commodity Config'] = com
            p = pd_info.get(row.get('PD Action'), {})
            row['Category'] = line_category(row)
            cat = cat_by_name.get(row['Category'], {})
            row['Needs Contract'] = p.get('Needs Contract') is not False and cat.get('Needs Contract') is not False
            row['Is Upload'] = row['PD Type'] in upload_types
            lines.append(row)

    # ---- contract number and item per line (ME3L, typed numbers; D26, D35, D43, Q7, Q10)
    for row in lines:
        typed = row['Contract List']
        cands = [c for c in me3l_index.get((row['Vendor Key'], row['Material Code'], row['Order Unit']), [])
                 if (c['_plant'] is None or c['_plant'] == row['Plant'])
                 and overlaps(c['_from'], c['_to'], row['Effective Date'], row['Valid Date'])
                 and (not typed or c['_no'] in typed)]
        if len(cands) > 1:  # several contracts over time: the one that starts on the Effective Date wins
            same_start = [c for c in cands if c['_from'] == row['Effective Date']]
            cands = same_start if len(same_start) == 1 else cands
        row['Contract Candidates'] = len(cands)
        if not row['Needs Contract']:
            kind, no, item = 'Not needed', None, None
        elif len(cands) == 1:
            kind, no, item = 'Contract item', code_key(cands[0]['Contract No.']), to_text(cands[0]['Contract Item'])
        elif len(cands) > 1:
            kind, no, item = 'Ambiguous', None, None
        elif len(typed) == 1:
            kind, no, item = 'Contract only', typed[0], None
        elif typed:
            kind, no, item = 'Not found', None, None
        else:
            kind, no, item = 'No contract', None, None
        row.update({'Contract Match Type': kind, 'SAP Contract No.': no, 'SAP Contract Item': item})
        if row['Is Upload'] and kind in CONTRACT_MATCH_ISSUE:
            chk.add('Warning', 'Issue', 'Contract No.', kind, CONTRACT_MATCH_ISSUE[kind], row)
        if row['Is Upload'] and kind == 'Contract only':
            chk.add('Warning', 'Note', 'Contract No.', kind,
                    'ได้เลข Contract แต่ไม่พบ Contract Item ใน ME3L: FOC ไปได้ Source List รอ', row)

    # ---- PD manual tasks (D52, Q6)
    manual = []
    for row in lines:
        if row['Is Upload'] or row['PD Type'] == no_setup:
            continue
        if row['PD Type'] is None:
            chk.add('Warning', 'Issue', 'PD manual', None, 'ไม่มี PD Type: ไฟล์ IC ไม่ได้เลือก PD Action', row)
            continue
        task = f"ทำเอง: {row.get('PD Action') or row['PD Type']}" + (
            f" — {row['PD Action Description']}" if row.get('PD Action Description') else '')
        manual.append({c: row.get(c) for c in MANUAL_COLUMNS} | {'Task': task})
        chk.add('Warning', 'Manual', 'PD manual', row['PD Type'], task, row)

    def expand(rows):
        '''All Plant → one row per plant of the file's dimPlant_Raw (S1, D45, Q15). grain: line × Plant.'''
        out = []
        for row in rows:
            if row['Plant'] or row.get('Plant Scope') == 'By Plant':
                out.append(dict(row))
            elif row['File Plants']:
                out.extend(dict(row, Plant=p['Plant'], **{'Plant Name': p['Plant Name'], 'Plant Key': p['Plant']})
                           for p in row['File Plants'])
            else:
                out.append(dict(row, **{'Expand Issue': 'All Plant แต่ไฟล์ IC Processed ไม่มี dimPlant_Raw (ประมวลผลใหม่ด้วยแอปรุ่นปัจจุบัน)'}))
        return out

    upload_lines = [r for r in lines if r['Is Upload']]
    source_list = _source_list(expand(upload_lines), reports.get('ZMMP1004'), category_of, cat_by_name, method_mode, cfg, chk)
    info_record = _info_record(expand(upload_lines), category_of, cat_by_name, me1l, dup_me1l, purch_group, tax_map,
                               tax_plants, vendor_tax_code, reports.get('ME1L') is not None, cfg, chk)
    foc_add = _foc_add(upload_lines, cat_by_name, cfg, chk)
    foc_delete = _foc_delete(upload_lines, foc_add, reports.get('ZMMR1019'), cat_by_name, cfg, chk)

    check = chk.frame()
    lines_df = pd.DataFrame([{k: v for k, v in r.items() if k not in ('File Plants', 'Commodity Config')} | {
        'Contract List': ', '.join(r['Contract List'])} for r in lines])
    out = {
        'Upload_Check': check,
        'Upload_PD_Manual': pd.DataFrame(manual, columns=MANUAL_COLUMNS),
        'setup_lines': lines_df,
        'Upload_SourceList_Base': source_list['base'],
        'Load-SL': source_list['upload'],
        'ME01 (manual)': source_list['me01'],
        'Upload_InfoRecord_Base': info_record['base'],
        'Template Info': info_record['upload'],
        'Upload_FOC_Base': pd.concat([foc_add['base'], foc_delete['base']], ignore_index=True),
    }
    for (option, has_plant), sheet in FOC_SHEETS.items():
        base = foc_add['ready'] if option == 'A' else foc_delete['ready']
        rows = [r for r in base if (r['Plant'] is not None) == has_plant]
        rows.sort(key=lambda r: (r['Contract No.'] or '', r['Material'] or '', r['Plant'] or '', r['Scale no.'] or 0))
        out[sheet] = pd.DataFrame(rows, columns=FOC_COLUMNS)
    return out


# ---------------------------------------------------------------- Source List (setup-outputs section 1)
def _source_list(rows, zmmp1004, category_of, cat_by_name, method_mode, cfg, chk):
    '''The load file for one Material × Plant must hold the WHOLE set of vendors still in use (S2):
    PD removes a vendor by leaving it out. Replace (Bidding, Pool) keeps only the new vendors, Keep keeps
    the current ones (S3). One row not ready → the whole Material × Plant waits.'''
    c = lambda name: cfg.const('Source List', name)  # noqa: E731
    new = []
    for row in rows:
        cat = cat_by_name.get(category_of(row, row['Plant']) if row['Plant'] else row['Category'], {})
        if cat.get('Needs Source List') is False:
            continue  # consignment: no Source List (D28)
        com = row['Commodity Config']
        fix_methods = [m.strip() for m in (to_text(com.get('Fix Vendor Methods')) or '').split(',') if m.strip()]
        fixed = is_true(row.get('Fix Vendor Requested')) or row.get('Method') in fix_methods
        mode = method_mode.get(row.get('Method'))
        issue = join_messages([
            row.get('Expand Issue'),
            'ไม่มี Plant' if not row['Plant'] and not row.get('Expand Issue') else None,
            (f"ไม่มีเลข Contract ({row['Contract Match Type']})" if row['SAP Contract No.'] is None else
             'ไม่พบ Contract Item ใน ME3L (Source List ต้องใช้ Agreement Item)' if row['SAP Contract Item'] is None else None),
            f"ยังไม่กำหนด Source List Mode ของ Method '{row.get('Method')}' ใน dim_method.csv (S3)" if not mode else None,
            'ไม่มี Effective Date / Valid Date' if not row['Effective Date'] or not row['Valid Date'] else None,
        ])
        new.append({'row': row, 'mode': mode, 'fixed': fixed, 'issue': issue, 'source': 'New', 'out': {
            'Material': row['Material Code'], 'Plant': row['Plant'], 'Valid from': sap_date(row['Effective Date']),
            'Valid to': sap_date(row['Valid Date']), 'Vendor': row['Vendor ID'], 'Purchase Org.': c('Purchase Org.'),
            'Order Unit': row['Order Unit'], 'Agreement': row['SAP Contract No.'], 'Agreement Item': row['SAP Contract Item'],
            'Fixed vendor': c('Fixed vendor') if fixed else c('Not Fixed vendor'), 'Blocked': c('Blocked'),
            'Materials Planning(MRP)': c('Materials Planning(MRP)')}})
    sets = {}
    for x in new:
        if x['row']['Plant']:
            s = sets.setdefault((x['row']['Material Code'], x['row']['Plant']),
                                {'vendors': set(), 'eff': None, 'modes': set(), 'blank_mode': False, 'fixed': False})
            s['vendors'].add(x['row']['Vendor Key'])
            eff = x['row']['Effective Date']
            s['eff'] = eff if s['eff'] is None or (eff and eff < s['eff']) else s['eff']
            s['blank_mode'] |= x['mode'] is None
            if x['mode']:
                s['modes'].add(x['mode'])
            s['fixed'] |= x['fixed']
    kept, me01 = [], []
    current = [] if zmmp1004 is None else records(zmmp1004)
    active, seen_rows = {}, set()
    for r in current:
        row_key = (code_key(r['Material Code']), to_text(r['Plant']), vendor_key(r['Vendor ID']), any_date(r['Valid from']))
        if row_key in seen_rows:  # grain of ZMMP1004: Material × Plant × Vendor × Valid from
            continue
        seen_rows.add(row_key)
        if is_true(r.get('Blocked')):
            continue  # a blocked vendor is not in use (PD does not block, D37)
        key = (code_key(r['Material Code']), to_text(r['Plant']))
        s = sets.get(key)
        if s is None or (s['eff'] and any_date(r['Valid to']) and any_date(r['Valid to']) < s['eff']):
            continue
        active.setdefault(key, []).append(r)
    for key, rs in active.items():
        s = sets[key]
        mode = None if s['blank_mode'] or len(s['modes']) != 1 else next(iter(s['modes']))
        vendor_count = len({vendor_key(r['Vendor ID']) for r in rs})
        for r in rs:
            vk = vendor_key(r['Vendor ID'])
            if vk in s['vendors']:
                continue  # replaced by the new row
            if mode == 'Replace':
                if vendor_count == 1:
                    me01.append({'Material Code': key[0], 'Plant': key[1], 'Vendor ID': to_text(r['Vendor ID']),
                                 'SAP Valid from': sap_date(r['Valid from']), 'SAP Valid to': sap_date(r['Valid to']),
                                 'SAP Agreement': to_text(r['Agreement']),
                                 'Note': 'ต้องลบ Vendor เดิมเองที่ ME01 (Source List เดิมมี Vendor เดียว)'})
                continue  # dropped: left out of the load
            fixed = False if s['fixed'] else is_true(r.get('Fixed vendor'))  # only one fixed vendor
            kept.append({'row': {'Material Code': key[0], 'Plant': key[1], 'Vendor ID': to_text(r['Vendor ID'])},
                         'source': 'Existing (kept)',
                         'issue': None if mode == 'Keep' else 'Source List Mode ของ Material × Plant นี้ว่างหรือขัดกันระหว่าง agreement (S3)',
                         'out': {'Material': key[0], 'Plant': key[1], 'Valid from': sap_date(r['Valid from']),
                                 'Valid to': sap_date(r['Valid to']), 'Vendor': to_text(r['Vendor ID']),
                                 'Purchase Org.': to_text(r['Purchase Org.']), 'Order Unit': to_text(r['Order Unit']),
                                 'Agreement': to_text(r['Agreement']), 'Agreement Item': to_text(r['Agreement Item']),
                                 'Fixed vendor': c('Fixed vendor') if fixed else c('Not Fixed vendor'),
                                 'Blocked': c('Blocked'), 'Materials Planning(MRP)': to_text(r['MRP'])}})
    # SAP: one fixed source per Material × Plant, and one record per vendor in the same period
    per_vendor, fixed_vendors = {}, {}
    for x in new:
        key = (x['row']['Material Code'], x['row']['Plant'])
        per_vendor.setdefault(key + (x['row']['Vendor Key'],), []).append(x)
        if x['fixed']:
            fixed_vendors.setdefault(key, set()).add(x['row']['Vendor Key'])
    for xs in per_vendor.values():
        if len(xs) > 1:
            for x in xs:
                x['issue'] = join_messages([x['issue'], 'Vendor เดียวกันมีหลายแถวใน Material × Plant นี้ (เช่น หลาย Order Unit): Source List ชนกัน'])
    for x in new:
        if len(fixed_vendors.get((x['row']['Material Code'], x['row']['Plant']), ())) > 1:
            x['issue'] = join_messages([x['issue'], 'มีหลาย Vendor ที่ต้อง fix ใน Material × Plant เดียวกัน (SAP fix ได้รายเดียว) เลือก Vendor ที่ fix ใน IC'])
    all_rows = new + kept
    bad_sets = {(x['row']['Material Code'], x['row'].get('Plant')) for x in all_rows if x['issue']}
    base, ready = [], []
    for x in all_rows:
        key = (x['row']['Material Code'], x['row'].get('Plant'))
        issue = x['issue'] or ('รอแถวอื่นใน Material × Plant เดียวกัน (ต้อง upload ทั้งชุด)' if key in bad_sets else None)
        base.append(x['out'] | {'Row Source': x['source'], 'Issue': issue, 'IC File': x['row'].get('IC File'),
                                'Line No': x['row'].get('Line No')})
        if issue:
            chk.add('Warning', 'Issue', 'Source List', 'Load-SL', issue, x['row'])
        else:
            ready.append(x['out'])
    for m in me01:
        chk.add('Warning', 'Manual', 'Source List', 'ME01 (manual)', m['Note'], m)
    ready.sort(key=lambda o: (o['Material'] or '', o['Plant'] or '', o['Vendor'] or ''))
    return {'base': pd.DataFrame(base), 'upload': pd.DataFrame(ready, columns=SOURCE_LIST_COLUMNS),
            'me01': pd.DataFrame(me01, columns=ME01_COLUMNS)}


# ---------------------------------------------------------------- Info Record (setup-outputs section 2)
def conversion_fraction(conversion):
    '''10 → (10, 1), 2.5 → (25, 10): whole numbers (D37). Returns (numerator, denominator, decimals) or None.'''
    if conversion is None or conversion <= 0:
        return None
    text = f'{conversion:.10f}'.rstrip('0').rstrip('.')
    decimals = len(text.split('.')[1]) if '.' in text else 0
    denominator = 10 ** decimals
    return int(round(conversion * denominator)), denominator, decimals


def _info_record(rows, category_of, cat_by_name, me1l, dup_me1l, purch_group, tax_map, tax_plants, vendor_tax_code,
                 has_me1l, cfg, chk):
    c = lambda name: cfg.const('Info Record', name)  # noqa: E731
    existing_types = set(cfg.const_list('PD', 'Existing Info Record PD Types'))
    new_types = set(cfg.const_list('PD', 'New Info Record PD Types'))
    items = []
    for row in rows:
        plant = row['Plant']
        category = category_of(row, plant) if plant else row['Category']
        cat = cat_by_name.get(category, {})
        ir_cat = to_text(cat.get('Info Record Category'))
        com = row['Commodity Config']
        cur = me1l.get((row['Vendor Key'], row['Material Code'], plant), {}) if plant else {}
        material_type = to_text(row.get('Material Type')) or to_text(com.get('Default Material Type'))
        group = to_text(cur.get('Purchasing Group')) or purch_group.get(material_type)
        deliv = to_number(cur.get('Planned Deliv. Time'))
        deliv = deliv if deliv else to_number(com.get('Default Pl. Deliv. Time'))
        minimum = to_number(row.get('Minimum Qty'))
        if minimum is None:
            current_min = to_number(cur.get('Minimum Order Qty'))
            minimum = current_min if current_min else to_number(com.get('Default Minimum Qty'))
        standard = to_number(row.get('Standard Qty'))
        standard = standard if standard is not None else to_number(com.get('Default Standard Qty'))
        tax = None
        if cat.get('Needs Tax Code') is True:
            tax = (c('Plant Rule Tax Code') if plant in tax_plants
                   else tax_map.get((ir_cat, vendor_tax_code.get(row['Vendor Key']))))
        frac = conversion_fraction(row['Conversion'])
        out = {
            'Supplier': row['Vendor ID'], 'Old Vendor Code': None, 'Supplier Name': row.get('Vendor Name'),
            'Material Code': row['Material Code'], 'Old Item Code': None,
            'Material Description': row.get('Material Description'), 'Purchasing Org': c('Purchasing Org'),
            'Plant': plant, 'Info Record Category': ir_cat, 'Vendor Material No.': to_text(row.get('Setup Material No.')),
            'Supplier Subrange': None, 'Sale Person': to_text(row.get('UCEP / Note for Info')),
            'Telephone': row.get('Agreement ID'), 'Order Unit': row['Order Unit'],
            'Denominator for Conversion of Order Unit to Base Unit': frac[1] if frac else None,
            'Numerator for Conversion': frac[0] if frac else None,
            'Var. Order Unit': c('Var. Order Unit'), 'Pl. Deliv. Time': deliv, 'Purchasing Group': group,
            'Standard Qty (per order unit)': standard, 'Minimum Qty (per order unit)': minimum,
            'Tol. Overdl': c('Tol. Overdl'), 'Conf. Ctr': c('Conf. Ctr'),
            'Net Price': None if row['Final Price'] is None else row['Final Price'] * row['Per for Setup'],
            'Currency': c('Currency'), 'Price Unit': row['Per for Setup'], 'Order Price Unit': row['Order Unit'],
            'Valid From': sap_date(row['Effective Date']), 'Valid to': sap_date(row['Valid Date']),
            'MFG_ID': to_text(row.get('MFG')), 'Tax code': tax, 'Local / Import': to_text(row.get('Local or Import')),
            'Registration No.': to_text(row.get('Registration No.')), 'Registration Type': to_text(row.get('Registration Type')),
            'Trade Product Use (TPU)': to_text(row.get('Trade Product Use (TPU)')) or to_text(row.get('TMT Code or UCEP Code')),
            'Special Feature': to_text(row.get('Special Feature')), 'Tradename': to_text(row.get('Trade Name')), 'Client': None,
        }
        issue = [
            row.get('Expand Issue'),
            'By Plant แต่ไม่มีรหัส Plant (Info Record ทำราย Plant, D34)' if not plant and not row.get('Expand Issue') else None,
            f"ไม่พบ Item Category '{category}' ใน item_category.csv" if not cat else None,
            'ไม่มี Conversion ของ Order Unit' if frac is None else None,
            'E03 Final Price ว่างหรือ ≤ 0' if not row['Final Price'] or row['Final Price'] <= 0 else None,
            'ไม่มี Effective Date / Valid Date' if not row['Effective Date'] or not row['Valid Date'] else None,
            f"ไม่พบ Purchasing Group (ME1L ไม่มี และ Material Type '{material_type}' ไม่มีใน purchasing_group.csv)" if not group else None,
            'หา Tax code ไม่ได้: อัปโหลดตาราง Tax ของ Vendor หรือกรอก Plant ที่ใช้ Tax code เฉพาะ (D37)'
            if cat.get('Needs Tax Code') is True and not tax else None,
        ]
        note = [
            'Pl. Deliv. Time ว่าง (ตรวจ dim_commodity.csv)' if deliv is None else None,
            'Standard / Minimum Qty ว่าง (เวชภัณฑ์รอ P25)' if standard is None or minimum is None else None,
            'Conversion มีทศนิยมมากกว่า 3 ตำแหน่ง ตรวจ Numerator / Denominator' if frac and frac[2] > 3 else None,
            'Vendor Material No. ว่าง (ไม่มี FOC และไม่มี Trade Name / Supplier Product Code)' if not out['Vendor Material No.'] else None,
            'ME1L มีหลายแถวของ Vendor × Material × Plant นี้: ใช้แถวแรก' if plant and (row['Vendor Key'], row['Material Code'], plant) in dup_me1l else None,
        ]
        if has_me1l and plant:  # Q11 (4): the PD Action against what SAP has now
            if row['PD Type'] in existing_types and not cur:
                note.append(f"PD Action {row.get('PD Action')} แต่ไม่พบ Info Record ปัจจุบันใน ME1L (อาจต้อง Add)")
            if row['PD Type'] in new_types and cur and to_text(cur.get('Order Unit')) == row['Order Unit']:
                note.append(f"PD Action {row.get('PD Action')} แต่ ME1L มี Info Record ของ Order Unit นี้อยู่แล้ว (อาจเป็น Update)")
        items.append({'row': row, 'out': out, 'issue': issue, 'note': note})

    # one Info Record = Supplier × Material × Plant × Category with one Order Unit (Q17)
    groups = {}
    for x in items:
        o = x['out']
        groups.setdefault((x['row']['Vendor Key'], o['Material Code'], o['Plant'], o['Info Record Category']), []).append(x)
    for key, xs in groups.items():
        if len(xs) < 2:
            continue
        files = {x['row']['IC File'] for x in xs}
        latest = max(x['row']['Effective Date'] or date.min for x in xs)
        winners = [x for x in xs if (x['row']['Effective Date'] or date.min) == latest]
        if len(files) > 1 and len(winners) == 1:  # two IC files on the same Info Record: latest Effective Date (D23)
            for x in xs:
                if x is winners[0]:
                    x['note'].append('Info Record เดียวกันอยู่ในหลายไฟล์ IC: ใช้ไฟล์ที่ Effective Date ล่าสุด (D23)')
                else:
                    x['issue'].append(f"ไม่ upload: ไฟล์ {winners[0]['row']['IC File']} มี Effective Date ล่าสุดของ Info Record นี้ (D23)")
        else:
            units = {x['out']['Order Unit'] for x in xs}
            for x in xs:
                x['issue'].append('Supplier × Material × Plant × Category เดียวกันมีหลาย Order Unit (Info Record ชนกัน)'
                                  if len(units) > 1 else 'แถวซ้ำในไฟล์ upload (Supplier × Material × Plant × Category)')
    base, ready = [], []
    for x in items:
        issue, note = join_messages(x['issue']), join_messages(x['note'])
        base.append(x['out'] | {'Issue': issue, 'Note': note, 'IC File': x['row']['IC File'], 'Line No': x['row']['Line No']})
        for kind, text in (('Issue', issue), ('Note', note)):
            for t in (text or '').split('; '):
                if t:
                    chk.add('Warning', kind, 'Info Record', 'Template Info', t, x['row'])
        if not issue:
            ready.append(x['out'])
    ready.sort(key=lambda o: (o['Supplier'] or '', o['Material Code'] or '', o['Plant'] or ''))
    return {'base': pd.DataFrame(base), 'upload': pd.DataFrame(ready, columns=INFO_RECORD_COLUMNS)}


# ---------------------------------------------------------------- FOC (setup-outputs section 2A)
def _foc_row(cfg, option, contract, material, plant, order_unit, valid_from, valid_to, scale, minimum, free):
    return {'Option': option, 'Contract No.': contract, 'Material': material, 'Plant': plant, 'Order Unit': order_unit,
            'Valid from': sap_date(valid_from), 'Valid to': sap_date(valid_to), 'Scale no.': scale,
            'Minimum qty': minimum, 'FreeGdsQty': minimum,      # = Minimum qty always (D35)
            'Base Unit': order_unit,                            # Order Unit always, in this file only (D35)
            'Add. quantity': free, 'Add. qty unit': order_unit,
            'CalculType': cfg.const('FOC', 'CalculType'), 'FreeGoodsCat.': cfg.const('FOC', 'FreeGoodsCat.')}


def _foc_add(rows, cat_by_name, cfg, chk):
    '''One row per FOC tier (D35); All Plant stays one row with Plant blank (FOC is tied to the Contract, D34).'''
    items = []
    for row in rows:
        if is_blank(row.get('Final FOC')):
            continue
        cat = cat_by_name.get(row['Category'], {})
        tiers = parse_foc(row['Final FOC'])
        valid = bool(tiers) and all(t['valid'] for t in tiers)
        for t in tiers if valid else [{'buy': None, 'free': None, 'scale_no': None}]:
            out = _foc_row(cfg, 'A', row['SAP Contract No.'], row['Material Code'], row['Plant'], row['Order Unit'],
                           row['Effective Date'], row['Valid Date'], t['scale_no'],
                           None if t['buy'] is None else int(t['buy']), None if t['free'] is None else int(t['free']))
            issue = [
                'E09 Item Category นี้ไม่มีของแถม (consignment ไม่มี FOC, D35)' if cat.get('Allows FOC') is False else None,
                f"E04 FOC ผิดรูปแบบหรือมีขั้นซ้ำ: {row['Final FOC']}" if not valid else None,
                # a blank Plant means "all plants" in the FOC file: a By Plant line must never land there
                'By Plant แต่ไม่มีรหัส Plant' if row.get('Plant Scope') == 'By Plant' and not row['Plant'] else None,
                (f"ไม่มีเลข Contract ({row['Contract Match Type']}) — FOC ผูกกับ Contract (D34)"
                 if row['SAP Contract No.'] is None and row['Contract Match Type'] != 'Not needed' else None),
                'ไม่มี Effective Date / Valid Date' if not row['Effective Date'] or not row['Valid Date'] else None,
            ]
            items.append({'row': row, 'out': out, 'issue': issue, 'tier': (t['buy'], t['free'])})
    keys = {}
    for x in items:
        o = x['out']
        k = (o['Contract No.'], o['Material'], o['Plant'] or '', o['Order Unit'], o['Valid from'], o['Scale no.'])
        keys[k] = keys.get(k, 0) + 1
        x['key'] = k
    base, ready = [], []
    for x in items:
        if keys[x['key']] > 1 and x['out']['Contract No.']:
            x['issue'].append('แถวซ้ำในไฟล์ upload (Contract × Material × Plant × Order Unit × Valid from × Scale no.)')
        issue = join_messages(x['issue'])
        sheet = FOC_SHEETS[('A', x['out']['Plant'] is not None)]
        base.append(x['out'] | {'Issue': issue, 'Note': None, 'Upload Sheet': sheet, 'IC File': x['row']['IC File'],
                                'Line No': x['row']['Line No']})
        for t in (issue or '').split('; '):
            if t:
                chk.add('Warning', 'Issue', 'FOC', sheet, t, x['row'])
        if not issue:
            ready.append(x['out'])
    return {'base': pd.DataFrame(base), 'ready': ready, 'items': items}


def _foc_delete(rows, foc_add, zmmr1019, cat_by_name, cfg, chk):
    '''Delete when the free goods are cancelled (D37). A current SAP record is in scope when a line has the same
    Contract × Material × Plant (blank = blank) × Order Unit and the periods overlap.
      Cancelled  the line has no Final FOC                → Delete row (values of the SAP record)
      Unchanged  the same tier is in the new FOC          → nothing
      Changed    the tier is not in the new FOC           → no delete: a new FOC comes with a new Contract (D50)'''
    scope = {}
    for row in rows:
        if row['SAP Contract No.'] and cat_by_name.get(row['Category'], {}).get('Allows FOC') is not False:
            scope.setdefault((row['SAP Contract No.'], row['Material Code'], row['Plant Key'], row['Order Unit']), []).append(row)
    new_tiers = {}
    for x in foc_add['items']:
        new_tiers.setdefault((x['row']['IC File'], x['row']['Line No']), set()).add(x['tier'])
    picked = {}
    for r in records(zmmr1019):
        plant = to_text(r['Plant'])
        key = (code_key(r['Contract No.']), code_key(r['Material Code']), plant or '', to_text(r['Order Unit']))
        tier = (to_number(r['Minimum qty']), to_number(r['Add. quantity']))
        for row in scope.get(key, []):
            if not overlaps(any_date(r['Valid from']), any_date(r['Valid to']), row['Effective Date'], row['Valid Date']):
                continue
            if is_blank(row.get('Final FOC')):
                reason = 'Cancelled'
            elif tier in new_tiers.get((row['IC File'], row['Line No']), set()):
                reason = 'Unchanged'
            else:
                reason = 'Changed'
            if reason == 'Unchanged':
                continue
            sap_key = key + (any_date(r['Valid from']), to_text(r['Scale no.']))
            prev = picked.get(sap_key)
            # one SAP record overlapping several agreements: the latest Effective Date (D23)
            if prev is None or (row['Effective Date'] or date.min) > (prev[1]['Effective Date'] or date.min):
                picked[sap_key] = (r, row, reason)
    base, ready = [], []
    for r, row, reason in picked.values():
        scale = to_number(r['Scale no.'])
        out = _foc_row(cfg, 'D', code_key(r['Contract No.']), code_key(r['Material Code']), to_text(r['Plant']),
                       to_text(r['Order Unit']), r['Valid from'], r['Valid to'], None if scale is None else int(scale),
                       to_number(r['Minimum qty']), to_number(r['Add. quantity']))
        issue = 'ไม่ลบ: FOC เปลี่ยนต้องมาพร้อม Contract ใหม่ (D50)' if reason == 'Changed' else None
        note = ('Contract นี้มี FOC เดิมคนละขั้นกับ FOC ใหม่ ตรวจว่า Sourcing สร้าง Contract ใหม่แล้วหรือยัง (D50)'
                if reason == 'Changed' else None)
        sheet = FOC_SHEETS[('D', out['Plant'] is not None)]
        base.append(out | {'Issue': issue, 'Note': note, 'Delete Reason': reason, 'Upload Sheet': sheet,
                           'IC File': row['IC File'], 'Line No': row['Line No']})
        if issue:
            chk.add('Warning', 'Issue', 'FOC', sheet, issue, row)
            chk.add('Warning', 'Note', 'FOC', sheet, note, row)
        else:
            ready.append(out)
    return {'base': pd.DataFrame(base), 'ready': ready}
