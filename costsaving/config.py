'''Config tables of the IC process (dimCombination, dimMethod, ...) and small value helpers.

The CSV files in costsaving/config/ hold the business rules, so a Method or a PD Action is added by
editing a table, not the code (PROJECT_INSTRUCTIONS: rules live in config). They are public-safe:
no plant, vendor or material codes and no system ids. Values that belong to the company system
(e.g. the Ariba ERP System ID) are left blank here and passed in by the user at run time.
'''
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import math

import pandas as pd

CONFIG_DIR = Path(__file__).parent / 'config'

BOOL_COLUMNS = {
    'dim_method': ['Uses Transfer Price', 'is CT'],
    'dim_calc_type': ['Match Vendor', 'Match Order Unit', 'Prorate by Duration', 'Split by Winner'],
    'dim_pd': ['is Collect Cost Saving', 'Fix Vendor', 'Needs Contract'],
    'dim_combination': ['is Collect Cost Saving'],
}


def is_blank(value):
    '''None, NaN, pd.NA and empty / space-only text are all "no value".'''
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if value is pd.NA or value is pd.NaT:
        return True
    if isinstance(value, str) and value.strip() == '':
        return True
    return False


def clean(value):
    '''Blank → None, text → stripped text, everything else unchanged.'''
    if is_blank(value):
        return None
    if isinstance(value, str):
        return value.strip()
    return value


def to_text(value):
    '''Codes must stay text: 100.0 read from Excel becomes '100', not '100.0'.'''
    value = clean(value)
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def to_number(value):
    value = clean(value)
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_bool(value):
    value = clean(value)
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    return str(value).strip().upper() in ('TRUE', 'Y', 'YES', '1', 'X', '✓')


def round_half_up(value, decimals):
    '''Excel ROUND and SAP round half away from zero (0.125 → 0.13); Python round() would give 0.12.'''
    if value is None:
        return None
    quant = Decimal(1).scaleb(-int(decimals))
    rounded = Decimal(repr(float(value))).quantize(quant, rounding=ROUND_HALF_UP)
    return float(rounded)


@dataclass
class Config:
    dim_commodity: pd.DataFrame
    dim_method: pd.DataFrame
    dim_calc_type: pd.DataFrame
    dim_strategy_method: pd.DataFrame
    dim_pd: pd.DataFrame
    dim_combination: pd.DataFrame
    constants: dict = field(default_factory=dict)

    def const(self, group, name):
        key = (group, name)
        if key not in self.constants:
            raise KeyError(f'constants.csv has no row {group} / {name}')
        return self.constants[key]

    def const_number(self, group, name):
        return float(self.const(group, name))


def load_config(config_dir=CONFIG_DIR, overrides=None):
    '''Read the config CSVs. overrides = {(group, field): value} for values the user passes in.'''
    tables = {}
    for name in BOOL_COLUMNS.keys() | {'dim_commodity', 'dim_strategy_method'}:
        df = pd.read_csv(Path(config_dir) / f'{name}.csv', dtype=str, keep_default_na=False)
        df = df.map(clean)
        for col in BOOL_COLUMNS.get(name, []):
            df[col] = df[col].map(to_bool)
        tables[name] = df
    const_df = pd.read_csv(Path(config_dir) / 'constants.csv', dtype=str, keep_default_na=False)
    constants = {(r['Group'], r['Field']): clean(r['Value']) for _, r in const_df.iterrows()}
    if overrides:
        constants.update(overrides)
    return Config(constants=constants, **tables)
