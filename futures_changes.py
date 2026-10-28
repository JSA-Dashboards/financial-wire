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
    cr = crush_row(bars, nearby)
    if cr:
        rows.append(cr)
    try:
        eth = ethanol_row(api_key, as_of)
        if eth:
            rows.append(eth)
    except Exception as e:   # ethanol is thin and flaky; never drop the grain rows for it
        print("ETHANOL ERROR:", e)
    return rows


def _crush(beans_c: float, meal_st: float, oil_c: float) -> float:
    """Board crush in $/bu: 44 lb meal + 11 lb oil per bushel, less the beans (cents -> $)."""
    return 0.022 * meal_st + 0.11 * oil_c - beans_c / 100.0


def crush_row(bars: dict, nearby: dict) -> dict | None:
    """Nearby board crush from the three nearby contracts, day/week change on the crush itself."""
    try:
        tb, tm, to = nearby["Soybeans"], nearby["Soybean meal"], nearby["Soybean oil"]
        df = bars[tb]["settle"].to_frame("b").join(bars[tm]["settle"].rename("m"), how="inner")             .join(bars[to]["settle"].rename("o"), how="inner")
    except KeyError:
        return None
    if len(df) < 2:
        return None
    c = (0.022 * df["m"] + 0.11 * df["o"] - df["b"] / 100.0)
    last = float(c.iloc[-1])
    return {"name": "Board crush", "ticker": f"{tb}/{tm}/{to}", "unit": "$/bu", "decimals": 3,
            "asof": str(df.index[-1]), "last": round(last, 4),
            "day": round(last - float(c.iloc[-2]), 4),
            "week": round(last - float(c.iloc[-6]), 4) if len(c) >= 6 else None}


def ethanol_row(api_key: str, as_of: date) -> dict | None:
    """CME Chicago Ethanol (Platts) futures, nearest contract with a settlement.
    A futures PROXY for the Platts assessment (it settles against the month's average).
    Day = settle - previous settlement (from the snapshot, which exists even on no-trade
    days). Week = vs the last bar 5-9 days earlier, or n/a: the contract rarely trades."""
    contracts = mx.get_active_contract_tickers("CU", api_key, as_of, limit=1000)[:6]
    snaps = mx.get_snapshots([c["ticker"] for c in contracts], api_key)
    for c in contracts:
        sess = (snaps.get(c["ticker"]) or {}).get("session") or {}
        settle, prev = sess.get("settlement_price"), sess.get("previous_settlement")
        if settle:
            week = None
            bars = mx.get_daily_bars(c["ticker"], api_key)
            if len(bars):
                old = bars[[5 <= (as_of - d).days <= 9 for d in bars.index]]
                if len(old):
                    week = round(float(settle) - float(old["settle"].iloc[-1]), 4)
            return {"name": "Chicago ethanol (Platts fut)", "ticker": c["ticker"], "unit": "$/gal", "decimals": 4,
                    "asof": str(as_of), "last": round(float(settle), 4),
                    "day": round(float(settle) - float(prev), 4) if prev else None, "week": week}
    return None


def fmt_change(v, dec: int) -> str:
    return "n/a" if v is None else f"{v:+.{dec}f}"
