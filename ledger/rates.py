"""Exchange-rate helpers and month-pair logic."""
import json


def as_rates(rate, n):
    """Per-sheet exchange rates. `rate` may be one number or a list of n numbers."""
    if isinstance(rate, (list, tuple)):
        out = [float(x) if x else 0.0 for x in rate]
        if len(out) != n:
            raise ValueError("One exchange rate is needed for every sheet.")
    else:
        out = [float(rate or 0)] * n
    if any(x <= 0 for x in out):
        raise ValueError("Every sheet needs a USD rate greater than 0.")
    return out


def parse_rates(rates_str, n, fallback=None):
    """Form field -> list of n rates (JSON list, same order as the files)."""
    vals = None
    if rates_str:
        try:
            vals = json.loads(rates_str)
        except json.JSONDecodeError:
            raise ValueError("Invalid rates.")
        if not isinstance(vals, list) or len(vals) != n:
            raise ValueError("One USD rate per sheet is required.")
    elif fallback:
        vals = [fallback] * n
    else:
        raise ValueError("USD rate missing.")
    return as_rates(vals, n)


def same_rate(rates):
    return rates[0] if len(set(rates)) == 1 else None


def rates_text(month_names, rates):
    if len(set(rates)) == 1:
        return f"PKR {rates[0]:,.2f} per USD"
    return "; ".join(f"{m}: PKR {r:,.2f}" for m, r in zip(month_names, rates)) + " per USD"


def month_pairs(names):
    """Every pair of selected months (later vs earlier).
    Returns a list of (a_index, b_index, title, is_overall)."""
    n = len(names)
    out = []
    for a in range(n - 1):
        for b in range(a + 1, n):
            ov = (a == 0 and b == n - 1 and n >= 3)
            title = (f"OVERALL: {names[b]} vs {names[a]}" if ov
                     else f"{names[b]} vs {names[a]}")
            out.append((a, b, title, ov))
    out.sort(key=lambda p: (p[1] - p[0], p[0]))
    return out
