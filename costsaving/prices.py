'''FOC text and base-unit prices.

FOC is written "buyFfree" per tier, several tiers separated by commas: "10F1" = buy 10 get 1,
"10F1,20F3" = two tiers. The base-unit price takes FOC off the order-unit price:

    base price = price × buy ÷ (buy + free) ÷ conversion

With several tiers the deepest tier (largest buy) is used — the same tier as the current IC file
when tiers are written small to large.
'''
from .config import is_blank, round_half_up


def parse_foc(foc):
    '''"10F1,20F3" → [{'buy': 10, 'free': 1, 'valid': True}, ...] sorted by buy, scale_no from 1.

    A tier is valid when it is <whole number > 0>F<whole number > 0>. A repeated buy makes every
    tier invalid. Blank FOC → [].
    '''
    if is_blank(foc):
        return []
    tiers = []
    for raw in str(foc).upper().split(','):
        text = raw.strip()
        if not text:
            continue
        parts = text.split('F')
        buy = free = None
        if len(parts) == 2:
            try:
                buy, free = float(parts[0]), float(parts[1])
            except ValueError:
                buy = free = None
        valid = (buy is not None and free is not None and buy > 0 and free > 0
                 and buy.is_integer() and free.is_integer())
        tiers.append({'tier': text, 'buy': buy, 'free': free, 'valid': valid})
    if len({t['buy'] for t in tiers}) < len(tiers):
        for t in tiers:
            t['valid'] = False
    tiers.sort(key=lambda t: (t['buy'] is None, t['buy'] or 0))
    for i, t in enumerate(tiers, 1):
        t['scale_no'] = i
    return tiers


def is_foc_valid(foc):
    '''Blank FOC is valid (no free goods). Otherwise every tier must be valid.'''
    tiers = parse_foc(foc)
    if is_blank(foc):
        return True
    return bool(tiers) and all(t['valid'] for t in tiers)


def base_price(price, foc, conversion, decimals):
    '''Order-unit price (+ FOC) → base-unit price, rounded like Excel ROUND. Invalid FOC → no FOC.'''
    if price is None or conversion in (None, 0):
        return None
    valid = [t for t in parse_foc(foc) if t['valid']]
    factor = 1.0
    if valid:
        deepest = max(valid, key=lambda t: t['buy'])
        factor = deepest['buy'] / (deepest['buy'] + deepest['free'])
    return round_half_up(price * factor / conversion, decimals)
