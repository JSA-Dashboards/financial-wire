"""Nearby futures: last settle plus day and week change in nominal terms (not %).

Run by push_wire.py on your PC (needs MASSIVE_API_KEY in financial_wire/.env);
the hosted app only reads the result from Snowflake, so the key never leaves your PC.

Day  = last settle minus the prior session's settle.
Week = last settle minus the settle five sessions back, on the SAME contract
       (so a roll never shows up as a move).
Units are the contract's own quote: grains in cents/bu, meal in $/short ton,
soybean oil and cattle in cents/lb.
"""
from datetime import date

import massive_api as mx

# name, Massive product code, unit, decimals shown
PRODUCTS = [
    ("Corn", "ZC", "\u00a2/bu", 2),
    ("Soybeans", "ZS", "\u00a2/bu", 2),
    ("Soybean meal", "ZM", "$/st", 2),
    ("Soybean oil", "ZL", "\u00a2/lb", 2),
    ("SRW wheat", "ZW", "\u00a2/bu", 2),
    ("HRW wheat", "KE", "\u00a2/bu", 2),
    ("Live cattle", "LE", "\u00a2/lb", 3),
]


def build(api_key: str, as_of: date | None = None) -> list[dict]:
    as_of = as_of or date.today()
    nearby = {}
    for name, code, _, _ in PRODUCTS:
        contracts = mx.get_active_contract_tickers(code, api_key, as_of)
        if contracts:
            nearby[name] = contracts[0]["ticker"]
    bars = mx.get_daily_bars_many(list(nearby.values()), api_key)

    rows = []
    for name, code, unit, dec in PRODUCTS:
        t = nearby.get(name)
        df = bars.get(t) if t else None
        if df is None or len(df) < 2:
            continue
        s = df["settle"]
        last = float(s.iloc[-1])
        rows.append({
            "name": name, "ticker": t, "unit": unit, "decimals": dec,
            "asof": str(df.index[-1]),
            "last": round(last, 4),
            "day": round(last - float(s.iloc[-2]), 4),
            "week": round(last - float(s.iloc[-6]), 4) if len(s) >= 6 else None,
        })
    return rows


def fmt_change(v, dec: int) -> str:
    return "n/a" if v is None else f"{v:+.{dec}f}"
