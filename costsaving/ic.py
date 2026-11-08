'''IC file process: Header + Lines + masters + PO lines → the four outputs.

Same rules as drafts/powerquery/ic_inputs.m and ic_outputs.m (decisions D11–D15, D39, D47–D53):

    Header ─┐
    Lines ──┼─▶ agreement_lines ─▶ for PD (Out_PD, Out_PD_Leadtime)
    masters ┘          ▲            for Contract (Out_Contract)
    PO lines ─▶ LPP ───┘            for Collect (Out_Collect)
             └▶ 12-month spend ───▶ for Summary (Summary, Summary_ByPlant, Summary_Lines)
                                    + Check (Error / Warning)

grain of agreement_lines: Line × Plant scope. A blank Plant Group = All Plant (1 row, Plant blank);
a Plant Group = 1 row per plant of the group.
'''
from dataclasses import dataclass
from datetime import date, datetime
import calendar

import pandas as pd

from .config import Config, clean, is_blank, to_bool, to_number, to_text, round_half_up
from .prices import base_price, is_foc_valid

LINE_COLUMNS = [
    'Line No', 'Material Code', 'Material Description', 'Order Unit', 'UOM Description', 'a0',
    'Final Price', 'Final FOC', 'Supplier Product Code', 'Supplier Product Name', 'Trade Name',
    'TMT Code or UCEP Code', 'Country of Manufacture', 'Registration No.', 'Remark', 'Local or Import', 'Quotation Date',
    'Vendor ID', 'Plant Group', 'PD Action', 'Nego Price', 'Nego FOC', 'Per for Setup',
    'Transfer Price', 'LPP SaveDrug', 'LPP DrPharma', 'Reference Material Code', 'LPP Hospital Manual',
    'Growth %', 'Fix Vendor Requested', 'UCEP / Note for Info',
    'Registration Type', 'Special Feature', 'Standard Qty', 'Minimum Qty',
]
TEXT_FIELDS = [
    'Material Code', 'Order Unit', 'Final FOC', 'Supplier Product Code', 'Supplier Product Name', 'Trade Name',
    'TMT Code or UCEP Code', 'Country of Manufacture', 'Registration No.', 'Remark', 'Local or Import',
    'Vendor ID', 'Plant Group', 'PD Action', 'Nego FOC', 'Reference Material Code', 'Fix Vendor Requested',
    'UCEP / Note for Info', 'Registration Type', 'Special Feature',
]
NUMBER_FIELDS = [
    'Final Price', 'Nego Price', 'Per for Setup', 'Transfer Price', 'LPP SaveDrug', 'LPP DrPharma',
    'LPP Hospital Manual', 'Growth %', 'Standard Qty', 'Minimum Qty',
]

# Out_PD = data contract between the IC file and the PD Setup file (setup_shared.m colsOut_PD)
OUT_PD_COLUMNS = [
    'Line No', 'PD Type', 'Commodity', 'Method', 'Plant Scope', 'Plant', 'Plant Name',
    'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'MFG',
    'Order Unit', 'Order Unit Text', 'UOM ISO Code', 'Conversion', 'Base Unit',
    'Per for Setup', 'Final Price', 'Price for Setup', 'Final FOC', 'Setup Material No.',
    'Supplier Product Code', 'Supplier Product Name', 'Trade Name', 'TMT Code or UCEP Code', 'Local or Import',
    'Registration No.', 'Registration Type', 'Trade Product Use (TPU)', 'Special Feature',
    'Standard Qty', 'Minimum Qty', 'Fix Vendor Requested', 'UCEP / Note for Info',
    'Material Group', 'Material Type', 'Effective Date', 'Valid Date', 'Project Name', 'Expected', 'Remark',
    'Template Version', 'Info Record Price', 'Info Record FOC', 'Country of Manufacture',
    'PD Action', 'PD Action Description', 'Combination ID', 'Package Version',
]
OUT_PD_LEADTIME_COLUMNS = [
    'Commodity', 'Vendor ID', 'Hospital Area', 'Product Type', 'Minimum Order (Baht)', 'Leadtime (Days)', 'Remark',
    'Effective Date', 'Valid Date',
]
# Collect: the first 29 columns keep the names and order of Input_forSavingCollect (v.20261102); v2 appended (P5)
OUT_COLLECT_COLUMNS = [
    'Commodity', 'Effective Date', 'Valid Date', 'Saving Strategy', 'Method', 'Project Name',
    'Plant', 'Plant Name', 'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name',
    'Order Unit', 'Order Unit Text', 'Conversion', 'Base Unit',
    'First Price', 'First FOC', 'Nego Price', 'Nego FOC', 'Final Price', 'Final FOC',
    'First Price Base Unit', 'Nego Price Base Unit', 'Final Price Base Unit', 'Avoid Base Unit', 'PPV Base Unit',
    'Expected', 'PD',
    'Baseline Price Base Unit', 'Baseline Source', 'IC Key Date', 'Transfer Price Base Unit', 'Growth %',
    'LPP SaveDrug Base Unit', 'LPP DrPharma Base Unit', 'IC Template Version',
    'Line No', 'Plant Scope', 'LPP Hospital', 'LPP Plant Name', 'LPP PO Number', 'LPP PO Date',
    'Reference Material Code', 'LPP Hospital Manual Base Unit', 'Quotation Date', 'Combination ID',
    'Package Version',
]
OUT_CONTRACT_COLUMNS = [
    'Item Number', 'Type', 'Item Name', 'Description', 'Commodity', 'Region', 'Item ID', 'ERP System ID', 'Plant',
    'Item Text', 'Price', 'Quantity', 'UOM.Quantity', 'Item Category', 'Material Number', 'Material Group',
    'Supplier Part Id', 'Tax Code', 'Remark', 'Advanced Pricing Conditions', 'Material Type', 'Material PO Text',
    'Quotation Number', 'Discount Amount', 'Discount Percentage', 'Surcharge Amount', 'Lead Time',
    'Down Payment%', 'Retention%', 'Incoterms', 'Incoterms Location', 'External System Line Number',
]
SUMMARY_LINE_COLUMNS = [
    'Line No', 'Plant Scope', 'Plant', 'Plant Name', 'Group', 'Spend Plant', 'Spend Plant Name',
    'Material Code', 'Material Description', 'Vendor ID', 'Vendor Name', 'Order Unit', 'Base Unit',
    'Calc Type', 'Winner Count', 'Duration Fraction', 'Spending Quantity', 'Spending Amount',
    'Baseline Price Base Unit', 'Baseline Source', 'Final Price Base Unit', 'Paid Price Base Unit',
    'Transfer Price Base Unit', 'Hospital Saving', 'Gain', 'Loss', 'Group Margin', 'Sourcing Value', 'Avoid',
    'Growth %', 'Hospital Saving + Growth', 'Sourcing Value + Growth', 'Expected', 'Rule Applied',
]
CHECK_COLUMNS = ['Level', 'Upload', 'Line No', 'Material Code', 'Plant', 'Vendor ID', 'Message']


@dataclass
class ICInputs:
    header: dict
    lines: pd.DataFrame
    masterinfo: pd.DataFrame
    uom: pd.DataFrame
    plant: pd.DataFrame
    plant_group: pd.DataFrame
    po_lines: pd.DataFrame
    leadtime: pd.DataFrame = None
    package: dict = None   # Package_Info of the template (D56); None for files without a package
    material_commodity: dict = None  # dimMaterial of the package: Material → Commodity (all commodities)
    official_plant_group: pd.DataFrame = None  # data team's groups in the package (D57: names locked)
    messages: list = None  # [(level, upload, text)] found while reading the input (e.g. the current IC file)


# ---------------------------------------------------------------- small helpers
def to_date(value):
    value = clean(value)
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors='coerce')
    return None if pd.isna(parsed) else parsed.date()


def add_months(d, months):
    month = d.month - 1 + months
    year = d.year + month // 12
    month = month % 12 + 1
    return date(year, month, min(d.day, calendar.monthrange(year, month)[1]))


def join_messages(messages):
    text = '; '.join(m for m in messages if m)
    return text or None


def records(df):
    return [] if df is None else df.to_dict('records')


# ---------------------------------------------------------------- masters
def prepare_masters(inputs):
    '''Lookups built once. Each key is unique, so a lookup never adds rows (no join fan-out).'''
    masterinfo = {}
    material_name, vendor_name = {}, {}
    for r in records(inputs.masterinfo):
        key = (to_text(r['Material Code']), to_text(r['Vendor ID']), to_text(r['Order Unit']))
        masterinfo.setdefault(key, r)
        material_name.setdefault(key[0], clean(r.get('Material Description')))
        vendor_name.setdefault(key[1], clean(r.get('Vendor Name')))
    uom = {}
    for r in records(inputs.uom):
        uom.setdefault(to_text(r['Order Unit']), {
            'Order Unit Text': clean(r.get('Order Unit Text')),
            'Conversion': to_number(r.get('Conversion Rate', r.get('Conversion'))),
            'Base Unit': to_text(r.get('Base Unit')),
            'UOM ISO Code': to_text(r.get('ISO Code')),
        })
    plant_by_name = {}
    for r in records(inputs.plant):
        plant_by_name.setdefault(to_text(r['Plant Name']), {
            'Plant': to_text(r['Plant']), 'Plant Name': to_text(r['Plant Name']),
            'Group': to_text(r.get('Group')), 'Location': to_text(r.get('Location')),
        })
    plant_by_code = {v['Plant']: v for v in plant_by_name.values()}
    # Plant List: Plant codes or Plant Names, comma separated (Sourcing builds their own groups, D56)
    groups = {}
    for r in records(inputs.plant_group):
        items = [n.strip() for n in (to_text(r.get('Plant List')) or '').split(',') if n.strip()]
        plants = []
        for n in items:
            p = plant_by_code.get(n) or plant_by_name.get(n)
            plants.append({'Plant': p['Plant'], 'Plant Name': p['Plant Name']} if p else {'Plant': None, 'Plant Name': n})
        groups.setdefault(to_text(r['Plant Group']), plants)
    return {
        'masterinfo': masterinfo, 'material_name': material_name, 'vendor_name': vendor_name,
        'uom': uom, 'plant_by_code': plant_by_code, 'groups': groups,
    }


# ---------------------------------------------------------------- PO lines → LPP and spend
def po_window(po_lines, key_date, lookback_months, materials):
    '''PO lines of the materials of this file inside [Key Date − lookback, Key Date).
    Strictly before the Key Date: a PO on the Key Date may already carry the new price.'''
    if po_lines is None or po_lines.empty or key_date is None:
        return pd.DataFrame(columns=list(po_lines.columns if po_lines is not None else []) + ['Commodity'])
    start = add_months(key_date, -int(lookback_months))
    df = po_lines[po_lines['Material Code'].isin(set(materials))]
    df = df[df['PO Date'].map(lambda d: d is not None and not pd.isna(d) and start <= d < key_date)].copy()
    df['Commodity'] = df['Commodity Full'].map(lambda v: None if v is None else (v.split('_', 1)[1] if '_' in v else v))
    return df


def lpp_tables(window, decimals):
    '''Whole PO (D11): free-goods lines (amount 0) are summed into their PO, so the price includes FOC.
    LPP by plant = latest PO of Plant × Material × Base Unit (ties → lower price), not tied to vendor.
    LPP all plant = lowest of the plants' latest POs (D13).'''
    if window.empty:
        return {}, {}
    po = (window.groupby(['Plant', 'Plant Name', 'PO Number', 'PO Date', 'Material Code', 'Base Unit'], dropna=False)
                .agg(qty=('Spending Quantity (Base Unit)', 'sum'), amt=('Spending Amount THB', 'sum'))
                .reset_index())
    po = po[(po['qty'] > 0) & (po['amt'] > 0)].copy()
    po['price'] = [round_half_up(a / q, decimals) for a, q in zip(po['amt'], po['qty'])]
    po['ord'] = po['PO Date'].map(lambda d: d.toordinal())
    po = po.sort_values(['ord', 'price'], ascending=[False, True])
    by_plant = {}
    for r in po.to_dict('records'):
        by_plant.setdefault((r['Plant'], r['Material Code'], r['Base Unit']), {
            'LPP Hospital': r['price'], 'LPP Plant Name': r['Plant Name'],
            'LPP PO Number': r['PO Number'], 'LPP PO Date': r['PO Date'],
        })
    all_plant = {}
    for (plant, mat, base), v in sorted(by_plant.items(), key=lambda kv: (kv[1]['LPP Hospital'], -kv[1]['LPP PO Date'].toordinal())):
        all_plant.setdefault((mat, base), v)
    return by_plant, all_plant


def spend_table(window):
    '''12-month spend before the Key Date. grain: Plant × Material × Vendor × Order Unit × Base Unit.'''
    if window.empty:
        return []
    sp = (window.groupby(['Plant', 'Plant Name', 'Material Code', 'Vendor ID', 'Order Unit', 'Base Unit'], dropna=False)
                .agg(qty=('Spending Quantity (Base Unit)', 'sum'), amt=('Spending Amount THB', 'sum'))
                .reset_index())
    return sp.to_dict('records')


# ---------------------------------------------------------------- agreement lines
def read_lines(lines_df):
    rows = []
    for i, r in enumerate(records(lines_df), 1):
        row = {c: clean(r.get(c)) for c in LINE_COLUMNS}
        for c in TEXT_FIELDS:
            row[c] = to_text(row[c])
        for c in NUMBER_FIELDS:
            row[c] = to_number(row[c])
        if row['Material Code'] is None:
            continue
        # same cleaning as the current Input query: Setup Material No. must not contain ' or _
        if row['Supplier Product Code']:
            row['Supplier Product Code'] = row['Supplier Product Code'].replace("'", '').replace('_', '')
        row['Line No'] = int(to_number(row['Line No'])) if to_number(row['Line No']) is not None else i
        row['Quotation Date'] = to_date(row['Quotation Date'])
        rows.append(row)
    return rows


def build_agreement_lines(inputs, cfg: Config):
    h = inputs.header
    decimals = int(cfg.const_number('IC', 'Base Price Decimals'))
    key_date = to_date(h.get('Key Date'))
    growth_header = to_number(h.get('Growth %'))
    if growth_header is None:
        growth_header = cfg.const_number('IC', 'Default Growth %')
    m = prepare_masters(inputs)
    lines = read_lines(inputs.lines)
    materials = {x['Material Code'] for x in lines} | {x['Reference Material Code'] for x in lines if x['Reference Material Code']}
    window = po_window(inputs.po_lines, key_date, cfg.const_number('IC', 'LPP Lookback Months'), materials)
    by_plant, all_plant = lpp_tables(window, decimals)
    material_commodity = {}
    for r in window.to_dict('records'):
        if r.get('Commodity'):
            material_commodity.setdefault(r['Material Code'], r['Commodity'])
    # the package holds the spending of one commodity only; dimMaterial knows every commodity (E10)
    material_commodity.update(inputs.material_commodity or {})
    comb = {(r['Strategy'], r['Method'], r['PD Action']): r for r in records(cfg.dim_combination)}
    pd_info = {r['PD Action']: r for r in records(cfg.dim_pd)}
    commodity = {r['Commodity']: r for r in records(cfg.dim_commodity)}.get(to_text(h.get('Commodity')), {})

    out = []
    for x in lines:
        row = dict(x)
        row.update({
            'Commodity': to_text(h.get('Commodity')), 'Saving Strategy': to_text(h.get('Saving Strategy')),
            'Method': to_text(h.get('Method')), 'Project Name': to_text(h.get('Project Name')),
            'Effective Date': to_date(h.get('Effective Date')), 'Valid Date': to_date(h.get('Valid Date')),
            'IC Key Date': key_date, 'Quotation Date': x['Quotation Date'] or to_date(h.get('Quotation Date')),
            'Template Version': cfg.const('IC', 'Template Version'),
            'Package Version': to_text((inputs.package or {}).get('Package Version')),
        })
        # line value wins over the header default
        row['Vendor ID'] = x['Vendor ID'] or to_text(h.get('Vendor ID'))
        row['Plant Group'] = x['Plant Group'] or to_text(h.get('Plant Group'))
        row['Growth %'] = x['Growth %'] if x['Growth %'] is not None else growth_header
        row['Per for Setup'] = x['Per for Setup'] if x['Per for Setup'] is not None else 1
        # PD Action → Combination ID, PD Type (= PD Sheet), is Collect Cost Saving (D52)
        c = comb.get((row['Saving Strategy'], row['Method'], row['PD Action']), {})
        p = pd_info.get(row['PD Action'], {})
        row['Combination ID'] = c.get('Combination ID')
        row['PD Type'] = c.get('PD Sheet')
        row['is Collect Cost Saving'] = c.get('is Collect Cost Saving')
        row['PD Action Description'] = p.get('PD Action Description')
        row['PD Needs Contract'] = p.get('Needs Contract')
        if row['Fix Vendor Requested'] is None and p.get('Fix Vendor') is True:
            row['Fix Vendor Requested'] = '✓'
        # masters
        mi = m['masterinfo'].get((row['Material Code'], row['Vendor ID'], row['Order Unit']), {})
        row['MFG'] = clean(mi.get('MFG'))
        row['Info Record FOC'] = to_text(mi.get('FOC'))
        row['Info Record Price'] = to_number(mi.get('Order Unit Price'))
        row['Material Description'] = m['material_name'].get(row['Material Code'])
        row['Vendor Name'] = m['vendor_name'].get(row['Vendor ID'])
        row['Material Commodity'] = material_commodity.get(row['Material Code'])
        # Base Unit and Conversion come from the Order Unit code (dimUOM_Raw), as in the current file
        u = m['uom'].get(row['Order Unit'], {})
        row.update({k: u.get(k) for k in ('Order Unit Text', 'UOM ISO Code', 'Conversion', 'Base Unit')})
        row['Material Group'] = row['Material Type'] = None
        row['Item Category'] = 'Standard'
        # prices at base unit, FOC deducted
        nego_price, nego_foc = (x['Nego Price'], x['Nego FOC']) if x['Nego Price'] is not None else (x['Final Price'], x['Final FOC'])
        conv = row['Conversion']
        row['Nego Price Base Unit'] = base_price(nego_price, nego_foc, conv, decimals)
        row['Final Price Base Unit'] = base_price(x['Final Price'], x['Final FOC'], conv, decimals)
        row['Transfer Price Base Unit'] = base_price(x['Transfer Price'], None, conv, decimals)
        row['LPP SaveDrug Base Unit'] = base_price(x['LPP SaveDrug'], None, conv, decimals)
        row['LPP DrPharma Base Unit'] = base_price(x['LPP DrPharma'], None, conv, decimals)
        row['LPP Hospital Manual Base Unit'] = base_price(x['LPP Hospital Manual'], None, conv, decimals)
        row['LPP Lookup Material'] = x['Reference Material Code'] or x['Material Code']
        # Setup Material No. = FOC_code (code column from dimCommodity), as funcPD of the current file
        code = row.get(commodity.get('Setup Code Field')) if commodity.get('Setup Code Field') else None
        foc = x['Final FOC']
        row['Setup Material No.'] = None if foc is None and code is None else ('_' + code if foc is None else '_'.join(v for v in (foc, code) if v))
        row['Trade Product Use (TPU)'] = None
        row['Price for Setup'] = None if x['Final Price'] is None else x['Final Price'] * row['Per for Setup']
        # Plant Group → one row per plant; blank = All Plant
        if row['Plant Group'] is None:
            scope, plants, found = 'All Plant', [{'Plant': None, 'Plant Name': None}], True
        else:
            group = m['groups'].get(row['Plant Group'])
            found = bool(group) and all(g['Plant'] for g in group)
            scope, plants = 'By Plant', group or [{'Plant': None, 'Plant Name': None}]
        for pl in plants:
            r = dict(row)
            r.update({'Plant Scope': scope, 'Plant': pl['Plant'], 'Plant Name': pl['Plant Name'],
                      'Plant Key': pl['Plant'] or '', 'is Plant Group Found': found})
            lpp_key = (r['LPP Lookup Material'], r['Base Unit'])
            lpp = all_plant.get(lpp_key) if scope == 'All Plant' else by_plant.get((r['Plant'],) + lpp_key)
            lpp = lpp or {}
            r.update({k: lpp.get(k) for k in ('LPP Hospital', 'LPP Plant Name', 'LPP PO Number', 'LPP PO Date')})
            # manual value replaces only the hospital side; the min with SDC / Dr. Pharma still applies (P4)
            hosp = r['LPP Hospital Manual Base Unit'] if r['LPP Hospital Manual Base Unit'] is not None else r['LPP Hospital']
            r['LPP Hospital Final'] = hosp
            cands = [v for v in (hosp, r['LPP SaveDrug Base Unit'], r['LPP DrPharma Base Unit']) if v is not None]
            baseline = min(cands) if cands else None
            r['Baseline Price Base Unit'] = baseline
            if baseline is None:
                r['Baseline Source'] = 'No Baseline'
            elif baseline == hosp:
                r['Baseline Source'] = 'Manual' if r['LPP Hospital Manual Base Unit'] is not None else f"PO {r['LPP PO Number']}"
            elif baseline == r['LPP SaveDrug Base Unit']:
                r['Baseline Source'] = 'Save Drug'
            else:
                r['Baseline Source'] = 'Dr. Pharma'
            final_b, nego_b = r['Final Price Base Unit'], r['Nego Price Base Unit']
            r['PPV Base Unit'] = None if baseline is None or final_b is None else round_half_up(baseline - final_b, decimals)
            r['%PPV Base Unit'] = None if r['PPV Base Unit'] is None or not baseline else r['PPV Base Unit'] / baseline
            r['Avoid Base Unit'] = None if nego_b is None or final_b is None else round_half_up(nego_b - final_b, decimals)
            ppv = r['PPV Base Unit']
            r['Expected'] = None if ppv is None else ('Gain' if ppv > 0 else 'Loss' if ppv < 0 else 'Constant')
            out.append(r)
    out.sort(key=lambda r: (r['Line No'], r['Plant'] or ''))
    return out, window


# ---------------------------------------------------------------- outputs
def out_pd(lines):
    return pd.DataFrame([{c: r.get(c) for c in OUT_PD_COLUMNS} for r in lines], columns=OUT_PD_COLUMNS)


def plant_set(plant_list, plant):
    '''Plant List text → set of Plant codes (codes or names accepted; unknown names kept as text).'''
    by_code = {to_text(r['Plant']): to_text(r['Plant']) for r in records(plant)}
    by_name = {to_text(r['Plant Name']): to_text(r['Plant']) for r in records(plant)}
    items = [n.strip() for n in (to_text(plant_list) or '').split(',') if n.strip()]
    return frozenset(by_code.get(n) or by_name.get(n) or n for n in items)


def changed_official_groups(inputs):
    '''Groups whose name is in the data team's copy but whose plants differ (D57).'''
    official = {to_text(r['Plant Group']): plant_set(r.get('Plant List'), inputs.plant)
                for r in records(inputs.official_plant_group) if to_text(r.get('Plant Group'))}
    keyed = {}
    for r in records(inputs.plant_group):
        name = to_text(r.get('Plant Group'))
        if name:
            keyed.setdefault(name, []).append(plant_set(r.get('Plant List'), inputs.plant))
    return sorted(n for n, sets in keyed.items() if n in official and any(s != official[n] for s in sets))


def is_ct(header, cfg: Config):
    '''CT_Leadtime is used only for a CT Method, as in the current file (D56): dim_method.is CT.'''
    flags = {r['Method']: r['is CT'] for r in records(cfg.dim_method)}
    return flags.get(to_text(header.get('Method'))) is True


def leadtime_rows(inputs):
    return [r for r in records(inputs.leadtime)
            if to_text(r.get('Hospital Area')) and (to_number(r.get('Leadtime (Days)')) is not None
                                                     or to_number(r.get('Minimum Order (Baht)')) is not None)]


def out_pd_leadtime(inputs, cfg: Config):
    h = inputs.header
    rows = []
    if not is_ct(h, cfg):
        return pd.DataFrame(rows, columns=OUT_PD_LEADTIME_COLUMNS)
    for r in records(inputs.leadtime):
        area = to_text(r.get('Hospital Area'))
        days, minimum = to_number(r.get('Leadtime (Days)')), to_number(r.get('Minimum Order (Baht)'))
        if area is None or (days is None and minimum is None):
            continue
        rows.append({
            'Commodity': to_text(h.get('Commodity')), 'Vendor ID': to_text(r.get('Vendor ID')) or to_text(h.get('Vendor ID')),
            'Hospital Area': area, 'Product Type': to_text(r.get('Product Type')),
            'Minimum Order (Baht)': minimum, 'Leadtime (Days)': days, 'Remark': clean(r.get('Remark')),
            'Effective Date': to_date(h.get('Effective Date')), 'Valid Date': to_date(h.get('Valid Date')),
        })
    return pd.DataFrame(rows, columns=OUT_PD_LEADTIME_COLUMNS)


def out_collect(lines):
    rows = []
    for r in lines:
        if r.get('is Collect Cost Saving') is not True or not r.get('Final Price Base Unit'):
            continue
        c = dict(r)
        # First Price / First FOC no longer exist (D12); First Price Base Unit carries the Baseline (P12)
        c.update({'First Price': None, 'First FOC': None, 'First Price Base Unit': r['Baseline Price Base Unit'],
                  'PD': r.get('PD Action'), 'IC Template Version': r.get('Template Version')})
        rows.append({col: c.get(col) for col in OUT_COLLECT_COLUMNS})
    return pd.DataFrame(rows, columns=OUT_COLLECT_COLUMNS)


def contract_base(lines, cfg: Config):
    '''Ariba line items (section 2B). One row per line except PD Actions with Needs Contract = FALSE.'''
    na = cfg.const('Contract', 'N/A')
    commodity = {r['Commodity']: r for r in records(cfg.dim_commodity)}
    erp = cfg.const('Contract', 'ERP System ID')
    rows = []
    for r in lines:
        if r.get('PD Needs Contract') is False:
            continue
        com = commodity.get(r.get('Commodity'), {})
        text_field = com.get('Ariba Item Text Field')
        item_text = r.get(text_field) if text_field else None
        out = {
            'Item Number': None, 'Type': cfg.const('Contract', 'Type'),
            'Item Name': r.get('Material Description'), 'Description': r.get('Material Description'),
            'Commodity': com.get('Ariba Commodity'), 'Region': cfg.const('Contract', 'Region'),
            'Item ID': r.get('Material Code'), 'ERP System ID': erp,
            # All Plant → blank (C5); By Plant → plant code as text
            'Plant': None if r.get('Plant Scope') == 'All Plant' else r.get('Plant'),
            'Item Text': None if item_text is None else str(item_text).upper(),
            'Price': r.get('Final Price'), 'Quantity': cfg.const_number('Contract', 'Quantity'),
            'UOM.Quantity': r.get('Order Unit'), 'Item Category': cfg.const('Contract', 'Item Category'),
            'Material Number': r.get('Material Code'), 'Material Group': r.get('Material Group'),
            'Supplier Part Id': r.get('Setup Material No.'), 'Tax Code': cfg.const('Contract', 'Tax Code'),
            'Remark': r.get('Remark'), 'Advanced Pricing Conditions': None, 'Material Type': r.get('Material Type'),
            'Material PO Text': na, 'Quotation Number': na, 'Discount Amount': na, 'Discount Percentage': na,
            'Surcharge Amount': na, 'Lead Time': na, 'Down Payment%': na, 'Retention%': na,
            'Incoterms': None, 'Incoterms Location': na, 'External System Line Number': None,
        }
        rows.append({'line': r, 'out': out})
    keys = {}
    for x in rows:
        k = (x['line'].get('Vendor ID'), x['out']['Material Number'], x['out']['UOM.Quantity'], x['out']['Plant'])
        keys[k] = keys.get(k, 0) + 1
    for x in rows:
        r, o = x['line'], x['out']
        k = (r.get('Vendor ID'), o['Material Number'], o['UOM.Quantity'], o['Plant'])
        x['issue'] = join_messages([
            'ไม่มี Vendor ID (หนึ่ง Contract ต่อหนึ่ง Vendor)' if not r.get('Vendor ID') else None,
            'ไม่มี Material Description (Item Name บังคับ)' if not r.get('Material Description') else None,
            'E03 Final Price ว่างหรือ ≤ 0' if not r.get('Final Price') or r['Final Price'] <= 0 else None,
            'แถวซ้ำใน Contract เดียวกัน (Vendor × Material × Order Unit × Plant)' if keys[k] > 1 else None,
            'Item Text บังคับ: กรอก Trade Name (C9)' if o['Item Text'] is None else None,
            'ไม่พบ Commodity ของ Ariba ใน dim_commodity (C4)' if o['Commodity'] is None else None,
        ])
    return rows


def out_contract(contract_rows):
    rows = [x['out'] | {'_vendor': x['line'].get('Vendor ID')} for x in contract_rows if x['issue'] is None]
    rows.sort(key=lambda o: (o['_vendor'] or '', o['Material Number'] or '', o['Plant'] or ''))
    return pd.DataFrame(rows, columns=OUT_CONTRACT_COLUMNS + ['_vendor']).rename(columns={'_vendor': 'Vendor ID'})


def summary_lines(lines, window, inputs, cfg: Config):
    '''Annualize estimate (D9): 12-month spend before the Key Date × (Baseline − Paid Price).
    grain: Line × Plant scope × spend Plant; only lines with is Collect Cost Saving = TRUE.'''
    method = {r['Method']: r for r in records(cfg.dim_method)}
    calc = {r['Calc Type']: r for r in records(cfg.dim_calc_type)}
    plant_group = {to_text(r['Plant']): to_text(r.get('Group')) for r in records(inputs.plant)}
    spend = spend_table(window)
    collect = [r for r in lines if r.get('is Collect Cost Saving') is True]
    winners = {}
    for r in collect:
        winners.setdefault((r['Material Code'], r['Plant Key']), set()).add(r['Vendor ID'])
    out = []
    for r in collect:
        mt = method.get(r['Method'], {})
        ct = calc.get(mt.get('Calc Type'), {})
        uses_transfer = mt.get('Uses Transfer Price') is True
        winner_count = len(winners[(r['Material Code'], r['Plant Key'])])
        split = 1 / winner_count if ct.get('Split by Winner') is True and winner_count > 1 else 1
        duration = 1
        if ct.get('Prorate by Duration') is True and r['Effective Date'] and r['Valid Date']:
            duration = min(1, ((r['Valid Date'] - r['Effective Date']).days + 1) / 365)
        # spend of this line: Material + Base Unit (+ Plant when By Plant) (+ Vendor + Order Unit when By Vendor);
        # a moved code uses the spend of the old code
        matches = [s for s in spend
                   if s['Material Code'] == r['LPP Lookup Material'] and s['Base Unit'] == r['Base Unit']
                   and (r['Plant'] is None or s['Plant'] == r['Plant'])
                   and (ct.get('Match Vendor') is not True or s['Vendor ID'] == r['Vendor ID'])
                   and (ct.get('Match Order Unit') is not True or s['Order Unit'] == r['Order Unit'])]
        for s in matches or [None]:
            qty = None if s is None else s['qty'] * split * duration
            amt = None if s is None else s['amt'] * split * duration
            paid = r['Transfer Price Base Unit'] if uses_transfer and r['Transfer Price Base Unit'] is not None else r['Final Price Base Unit']
            base, final = r['Baseline Price Base Unit'], r['Final Price Base Unit']
            hs = None if base is None or paid is None or qty is None else (base - paid) * qty
            gm = None
            if uses_transfer and r['Transfer Price Base Unit'] is not None and final is not None and qty is not None:
                gm = (r['Transfer Price Base Unit'] - final) * qty
            sv = None if hs is None and gm is None else (hs or 0) + (gm or 0)
            avoid = 0 if r['Avoid Base Unit'] is None or qty is None or r['Avoid Base Unit'] <= 0 else r['Avoid Base Unit'] * qty
            g = r['Growth %'] or 0
            rule = [mt.get('Calc Type'),
                    f'÷ {winner_count} winners' if split < 1 else None,
                    f'× {duration:.2f} of a year' if duration < 1 else None,
                    'paid = Transfer' if paid != final else None,
                    'no spend in 12 months' if qty is None else None,
                    'no Baseline → Managed Spend only' if base is None else None]
            spend_plant = None if s is None else s['Plant']
            out.append({
                'Line No': r['Line No'], 'Plant Scope': r['Plant Scope'], 'Plant': r['Plant'], 'Plant Name': r['Plant Name'],
                'Group': plant_group.get(spend_plant), 'Spend Plant': spend_plant,
                'Spend Plant Name': None if s is None else s['Plant Name'],
                'Material Code': r['Material Code'], 'Material Description': r['Material Description'],
                'Vendor ID': r['Vendor ID'], 'Vendor Name': r['Vendor Name'], 'Order Unit': r['Order Unit'],
                'Base Unit': r['Base Unit'], 'Calc Type': mt.get('Calc Type'), 'Winner Count': winner_count,
                'Duration Fraction': duration, 'Spending Quantity': qty, 'Spending Amount': amt,
                'Baseline Price Base Unit': base, 'Baseline Source': r['Baseline Source'],
                'Final Price Base Unit': final, 'Paid Price Base Unit': paid,
                'Transfer Price Base Unit': r['Transfer Price Base Unit'],
                'Hospital Saving': hs, 'Gain': hs if hs is not None and hs > 0 else 0,
                'Loss': hs if hs is not None and hs < 0 else 0, 'Group Margin': gm, 'Sourcing Value': sv,
                'Avoid': avoid, 'Growth %': g,
                'Hospital Saving + Growth': None if hs is None else hs * (1 + g),
                'Sourcing Value + Growth': None if sv is None else sv * (1 + g),
                'Expected': r['Expected'], 'Rule Applied': ' · '.join(x for x in rule if x),
            })
    df = pd.DataFrame(out, columns=SUMMARY_LINE_COLUMNS)
    return df.sort_values(['Line No', 'Spend Plant'], na_position='last', kind='stable').reset_index(drop=True)


def summary_by_plant(sl):
    df = sl[sl['Spend Plant'].notna()]
    if df.empty:
        return pd.DataFrame(columns=['Group', 'Plant', 'Plant Name', 'Spending Amount', 'Hospital Saving', 'Gain', 'Loss',
                                     'Group Margin', 'Avoid', '% Saving to Managed Spend'])
    g = (df.groupby(['Group', 'Spend Plant', 'Spend Plant Name'], dropna=False)
           .agg(**{'Spending Amount': ('Spending Amount', 'sum'), 'Hospital Saving': ('Hospital Saving', 'sum'),
                   'Gain': ('Gain', 'sum'), 'Loss': ('Loss', 'sum'), 'Group Margin': ('Group Margin', 'sum'),
                   'Avoid': ('Avoid', 'sum')})
           .reset_index().rename(columns={'Spend Plant': 'Plant', 'Spend Plant Name': 'Plant Name'}))
    g['% Saving to Managed Spend'] = [None if not a else h / a for h, a in zip(g['Hospital Saving'], g['Spending Amount'])]
    return g


def summary(sl, lines, check, header, package=None):
    def total(col):
        return float(pd.to_numeric(sl[col], errors='coerce').fillna(0).sum())

    def count_lines(df):
        return int(df['Line No'].nunique())

    spend = total('Spending Amount')
    hs = total('Hospital Saving')
    rows = [
        ('Commodity', header.get('Commodity')), ('Saving Strategy', header.get('Saving Strategy')),
        ('Method', header.get('Method')), ('Project Name', header.get('Project Name')),
        ('Effective Date', to_date(header.get('Effective Date'))), ('Valid Date', to_date(header.get('Valid Date'))),
        ('Key Date (Baseline frozen)', to_date(header.get('Key Date'))),
        ('Package Version', to_text((package or {}).get('Package Version'))),
        ('Spending in package', None if not package else
         f"{package.get('Spending From'):%Y-%m-%d} to {package.get('Spending To'):%Y-%m-%d}"
         if package.get('Spending From') and package.get('Spending To') else None),
        ('Lines', len({r['Line No'] for r in lines})),
        ('Lines counted for Cost Saving', count_lines(sl)),
        ('Lines without Baseline', count_lines(sl[sl['Baseline Source'] == 'No Baseline'])),
        ('Lines with Loss', count_lines(sl[sl['Expected'] == 'Loss'])),
        ('Lines without spend in 12 months', count_lines(sl) - count_lines(sl[sl['Spending Quantity'].notna()])),
        ('Spending 12 months (matched)', spend),
        ('Hospital Saving (Annualize)', hs), ('  Gain', total('Gain')), ('  Loss', total('Loss')),
        ('% Hospital Saving to Managed Spend', hs / spend if spend else None),
        ('Hospital Saving (Annualize + Growth)', total('Hospital Saving + Growth')),
        ('Group Margin (Pool)', total('Group Margin')), ('Sourcing Value (Annualize)', total('Sourcing Value')),
        ('Sourcing Value (Annualize + Growth)', total('Sourcing Value + Growth')),
        ('Avoid (positive only, not in Cost Saving, Q3)', total('Avoid')),
        ('Errors (must fix before submit)', int((check['Level'] == 'Error').sum())),
        ('Warnings', int((check['Level'] == 'Warning').sum())),
    ]
    return pd.DataFrame(rows, columns=['Measure', 'Value'])


# ---------------------------------------------------------------- Check (template-spec section 7)
def build_check(lines, contract_rows, header, cfg: Config, inputs=None):
    msgs = []

    def add(level, upload, message, r=None):
        msgs.append({'Level': level, 'Upload': upload, 'Line No': None if r is None else r.get('Line No'),
                     'Material Code': None if r is None else r.get('Material Code'),
                     'Plant': None if r is None else r.get('Plant'),
                     'Vendor ID': None if r is None else r.get('Vendor ID'), 'Message': message})

    commodity, method, strategy = to_text(header.get('Commodity')), to_text(header.get('Method')), to_text(header.get('Saving Strategy'))
    effective, valid = to_date(header.get('Effective Date')), to_date(header.get('Valid Date'))
    commodities = set(cfg.dim_commodity['Commodity'])
    methods = set(cfg.dim_method['Method'])
    allowed = {(r['Strategy'], r['Method']) for r in records(cfg.dim_strategy_method)}
    for text in [
        'H01 ยังไม่ได้เลือก Commodity' if not commodity else (None if commodity in commodities else f"H01 Commodity '{commodity}' ไม่มีใน dim_commodity"),
        'H02 ยังไม่ได้เลือก Method' if not method else (None if method in methods else f"H02 Method '{method}' ไม่มีใน dim_method"),
        'H03 ยังไม่ได้เลือก Saving Strategy' if not strategy else None,
        f"H07 Method '{method}' ใช้กับ Saving Strategy นี้ไม่ได้ (dim_strategy_method)" if strategy and method and (strategy, method) not in allowed else None,
        'H04 ยังไม่ได้กรอก Project Name' if not to_text(header.get('Project Name')) else None,
        'H05 ยังไม่ได้กรอก Effective Date / Valid Date' if not effective or not valid else ('E07 Valid Date ก่อน Effective Date' if valid < effective else None),
        'H06 ยังไม่ได้กรอก Key Date (วันที่หยุดค่า Baseline, D14)' if not to_date(header.get('Key Date')) else None,
    ]:
        if text:
            add('Error', 'Header', text)

    # data package (D56): the Baseline is only as recent as the package
    package = (inputs.package if inputs is not None else None) or {}
    key_date = to_date(header.get('Key Date'))
    if package:
        p_from, p_to = package.get('Spending From'), package.get('Spending To')
        p_com = to_text(package.get('Commodity'))
        if commodity and p_com and commodity != p_com:
            add('Error', 'Package', f"P03 Commodity ใน Header ({commodity}) ไม่ตรงกับแพ็กเกจ ({p_com}) ใช้ template ของ commodity นี้")
        if key_date and p_to and (key_date - p_to).days > cfg.const_number('Package', 'Max Age Days'):
            add('Warning', 'Package', f"P01 แพ็กเกจ {package.get('Package Version')} มี spending ถึง {p_to:%Y-%m-%d} "
                f"เก่ากว่า Key Date เกิน {int(cfg.const_number('Package', 'Max Age Days'))} วัน ขอ template รุ่นใหม่จากทีมข้อมูล")
        if key_date and p_from and add_months(key_date, -int(cfg.const_number('IC', 'LPP Lookback Months'))) < p_from:
            add('Warning', 'Package', f"P02 spending ในแพ็กเกจเริ่ม {p_from:%Y-%m-%d} ไม่ครบ 12 เดือนก่อน Key Date: LPP และ Summary อาจขาดบางส่วน")
    if inputs is not None and inputs.official_plant_group is not None and inputs.plant_group is not None:
        for name in changed_official_groups(inputs):
            add('Error', 'dimPlantGroup', f'P04 กลุ่ม {name} ถูกแก้ Plant List ต่างจากกลุ่มของทีมข้อมูล '
                'ถ้าต้องการชุด Plant อื่นให้เพิ่มกลุ่มใหม่ชื่ออื่น แล้วคืนค่ากลุ่มนี้ตาม template')
    if inputs is not None:
        ct, lt = is_ct(header, cfg), leadtime_rows(inputs)
        if ct and not lt:
            add('Warning', 'CT_Leadtime', 'W09 Method นี้เป็น CT แต่ยังไม่กรอก CT_Leadtime (PD จะใช้ค่าเริ่มต้น)')
        if not ct and lt:
            add('Warning', 'CT_Leadtime', 'W10 กรอก CT_Leadtime แต่ Method นี้ไม่ใช่ CT จึงไม่ส่งต่อให้ PD (ใช้เฉพาะ CT)')

    for level, upload, text in (inputs.messages if inputs is not None else None) or []:
        add(level, upload, text)

    uses_transfer = {r['Method']: r['Uses Transfer Price'] for r in records(cfg.dim_method)}
    ppv_warning = cfg.const_number('IC', 'PPV Warning %')
    keys = {}
    for r in lines:
        k = (r['Material Code'], r['Vendor ID'], r['Order Unit'], r['Plant Key'])
        keys[k] = keys.get(k, 0) + 1
    for r in lines:
        k = (r['Material Code'], r['Vendor ID'], r['Order Unit'], r['Plant Key'])
        transfer = uses_transfer.get(r['Method']) is True
        errors = [
            'E01 ไม่พบ Material Code ใน MasterInfo' if r['Material Description'] is None else None,
            ('E13 ยังไม่ได้เลือก PD Action' if r['PD Action'] is None else
             (f"E13 PD Action '{r['PD Action']}' ใช้กับ Strategy / Method นี้ไม่ได้ (dim_combination)" if r['Combination ID'] is None else None)),
            (f"E10 สินค้านี้เป็น {r['Material Commodity']} ไม่ตรงกับ Commodity ของไฟล์ (แยกไฟล์ตามทีม, D49)"
             if r['Material Description'] and r['Material Commodity'] and r['Material Commodity'] != r['Commodity'] else None),
            ('E02 ยังไม่ได้กรอก Order Unit' if r['Order Unit'] is None else
             (f"E02 ไม่พบ Order Unit {r['Order Unit']} ใน dimUOM_Raw" if r['Conversion'] is None else None)),
            ('E11 ไม่มี Vendor ID (กรอกที่ Header หรือ Lines)' if r['Vendor ID'] is None else
             (f"E11 ไม่พบ Vendor ID {r['Vendor ID']} ใน MasterInfo" if r['Vendor Name'] is None else None)),
            'E03 Final Price ว่างหรือ ≤ 0' if not r['Final Price'] or r['Final Price'] <= 0 else None,
            'E04 FOC ผิดรูปแบบ (ซื้อFแถม เช่น 10F1 หรือ 10F1,20F3)' if not (is_foc_valid(r['Final FOC']) and is_foc_valid(r['Nego FOC'])) else None,
            'E05 รายการซ้ำ (Material + Vendor + Order Unit + Plant)' if keys[k] > 1 else None,
            'E06 Method นี้ต้องกรอก Transfer Price' if transfer and r['Transfer Price'] is None else None,
            (f"E08 ไม่พบ Plant Group '{r['Plant Group']}' หรือชื่อ Plant ใน Plant List ไม่มีใน dimPlant_Raw"
             if r['is Plant Group Found'] is False else None),
        ]
        pct = r['%PPV Base Unit']
        warnings = [
            ('W01 หา Baseline ไม่ได้ ถ้าเป็นสินค้าใหม่ปล่อยได้ ถ้าไม่ใช่ให้กรอก LPP Hospital Manual หรือ Reference Material Code'
             if r['Baseline Source'] == 'No Baseline' else None),
            'W02 Final Price Base Unit สูงกว่า Baseline (จะเป็น Loss)' if r['Expected'] == 'Loss' else None,
            'W03 Final Price สูงกว่า Nego Price' if r['Nego Price'] is not None and r['Final Price'] is not None and r['Final Price'] > r['Nego Price'] else None,
            f'W05 PPV {pct:.0%} ของ Baseline (มักเป็นการคีย์หน่วยผิด)' if pct is not None and abs(pct) > ppv_warning else None,
            ('W06 Transfer Price ต่ำกว่า Final Price (SDC ขายต่ำกว่าทุน)'
             if transfer and r['Transfer Price Base Unit'] is not None and r['Final Price Base Unit'] is not None
             and r['Transfer Price Base Unit'] < r['Final Price Base Unit'] else None),
            (f"W07 Final Price สูงกว่าราคา Info Record ปัจจุบัน ({r['Info Record Price']:,.2f})"
             if r['Info Record Price'] is not None and r['Final Price'] is not None and r['Final Price'] > r['Info Record Price'] else None),
            ('W08 Vendor นี้ยังไม่มี Info Record ของ Material × Order Unit นี้ (PD ต้อง Add)'
             if r['Info Record Price'] is None and r['Material Description'] and r['Vendor Name'] else None),
        ]
        for text in errors:
            if text:
                add('Error', 'Lines', text, r)
        for text in warnings:
            if text:
                add('Warning', 'Lines', text, r)
    for x in contract_rows:
        for text in (x['issue'] or '').split('; '):
            if text:
                add('Error', 'Contract (Ariba)', text, x['line'])
    if contract_rows and not cfg.const('Contract', 'ERP System ID'):
        add('Warning', 'Contract (Ariba)', 'ERP System ID ว่าง (constants.csv): ใส่ค่าก่อน import เข้า Ariba')
    df = pd.DataFrame(msgs, columns=CHECK_COLUMNS)
    # a line-level message repeats on each plant of a Plant Group: keep it once per line
    return df.drop_duplicates(['Level', 'Upload', 'Line No', 'Material Code', 'Message']).reset_index(drop=True)


# ---------------------------------------------------------------- whole run
def run_ic(inputs: ICInputs, cfg: Config):
    '''Returns every table of the four outputs plus Check and agreement_line.'''
    lines, window = build_agreement_lines(inputs, cfg)
    contract_rows = contract_base(lines, cfg)
    check = build_check(lines, contract_rows, inputs.header, cfg, inputs)
    sl = summary_lines(lines, window, inputs, cfg)
    return {
        'agreement_line': pd.DataFrame(lines),
        'Out_PD': out_pd(lines),
        'Out_PD_Leadtime': out_pd_leadtime(inputs, cfg),
        'Out_Contract': out_contract(contract_rows),
        'Out_Collect': out_collect(lines),
        'Summary_Lines': sl,
        'Summary_ByPlant': summary_by_plant(sl),
        'Summary': summary(sl, lines, check, inputs.header, inputs.package),
        'IC_Check': check,
    }
