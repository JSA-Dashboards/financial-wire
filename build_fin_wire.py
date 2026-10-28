"""Render fin_input.json (from fetch_fin.ps1) into an Outlook-safe HTML email.

No summarizing step: the snapshot, reading list, history and PJ's ag-read
paragraphs are all the senders' own text, laid into a fixed template.

    python build_fin_wire.py            # -> fin_output.html
    python build_fin_wire.py --open     # also open it in the browser
"""
import html
import json
import re
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent
SECTION_ORDER = ["snapshot", "reading", "macro", "history"]  # reorder freely

INK, ACCENT, MUTED, RULE = "#32373c", "#0693e3", "#6b7280", "#d9dde1"
GREEN, RED, AMBER = "#2e7d32", "#c62828", "#b26a00"
FONT = "font-family:Arial,Helvetica,sans-serif;"


def esc(s: str) -> str:
    return html.escape(s or "", quote=True)


def heading(text: str) -> str:
    return (f'<tr><td style="padding:22px 0 6px 0;{FONT}font-size:18px;font-weight:bold;'
            f'color:{INK};border-bottom:2px solid {ACCENT};">{esc(text)}</td></tr>')


def row(inner: str, pad: str = "10px 0 0 0") -> str:
    return f'<tr><td style="padding:{pad};{FONT}font-size:14px;line-height:1.5;color:{INK};">{inner}</td></tr>'


def change_color(name: str, change: str) -> str:
    up = change.strip().startswith("+")
    # Yields and the dollar going up is a headwind for grains - his own snapshot paints yield +0.01 red.
    if re.search(r"yield", name, re.I):
        return RED if up else GREEN
    return GREEN if up else RED


def snapshot(d: dict) -> str:
    rows = d.get("snapshot") or []
    if not rows:
        return ""
    body = ""
    for r in rows:
        col = change_color(r["name"], r["change"])
        body += (f'<tr><td style="padding:9px 4px;border-bottom:1px solid {RULE};{FONT}font-size:14px;'
                 f'font-weight:bold;color:{INK};">{esc(r["name"])}</td>'
                 f'<td align="right" style="padding:9px 4px;border-bottom:1px solid {RULE};{FONT}font-size:14px;'
                 f'color:{INK};">{esc(r["level"])}</td>'
                 f'<td align="right" style="padding:9px 4px;border-bottom:1px solid {RULE};{FONT}font-size:14px;'
                 f'font-weight:bold;color:{col};width:70px;">{esc(r["change"])}</td></tr>')
    asof = d.get("snapshotAsOf")
    note = f'<div style="{FONT}font-size:11px;color:{MUTED};padding-top:4px;">Bloomberg, as of {esc(asof)}</div>' if asof else ""
    return heading("Market Snapshot") + row(f'<table width="100%" cellpadding="0" cellspacing="0">{body}</table>{note}', "6px 0 0 0")


def reading(d: dict) -> str:
    items = d.get("reading") or []
    if not items:
        return ""
    lis = "".join(f'<li style="padding:3px 0;">{esc(i)}</li>' for i in items)
    return heading("What We're Reading") + row(f'<ul style="margin:0;padding-left:20px;">{lis}</ul>')


def macro(d: dict) -> str:
    paras = d.get("macro") or []
    if not paras:
        return ""
    ps = "".join(f'<p style="margin:0 0 12px 0;">{esc(p)}</p>' for p in paras)
    return heading("Macro & Ag Read") + row(ps) + \
        f'<tr><td style="{FONT}font-size:11px;color:{MUTED};">PJ Quaid, StoneX</td></tr>'


def history(d: dict) -> str:
    h = d.get("history")
    if not h:
        return ""
    return heading("Today in Markets History") + row(esc(h))


SECTIONS = {"snapshot": snapshot, "reading": reading, "macro": macro, "history": history}


def build(d: dict) -> str:
    today = datetime.now().strftime("%A, %B %d, %Y").replace(" 0", " ")
    parts = "".join(SECTIONS[k](d) for k in SECTION_ORDER)

    gaps = d.get("gaps") or []
    if gaps:
        li = "".join(f"<li>{esc(g)}</li>" for g in gaps)
        parts += (f'<tr><td style="padding:22px 0 0 0;"><table width="100%" cellpadding="10" cellspacing="0" '
                  f'style="background:#fff4e0;border-left:4px solid {AMBER};"><tr><td style="{FONT}font-size:13px;'
                  f'color:{INK};"><b>Did not make this wire</b><ul style="margin:6px 0 0 0;padding-left:18px;">{li}</ul>'
                  f'</td></tr></table></td></tr>')

    srcs = esc("; ".join(d.get("sources") or []))
    foot = (f'<tr><td style="padding:24px 0 0 0;{FONT}font-size:11px;color:{MUTED};border-top:1px solid {RULE};">'
            f'Sources: {srcs}. For internal reading only; do not forward.</td></tr>')

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{esc(d.get('subject',''))}</title></head>
<body style="margin:0;padding:0;background:#ffffff;">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0" style="width:640px;max-width:100%;">
<tr><td style="padding:20px 0 4px 0;{FONT}font-size:24px;font-weight:bold;color:{INK};">Morning Financial Wire</td></tr>
<tr><td style="{FONT}font-size:13px;color:{MUTED};">{esc(today)}</td></tr>
{parts}
{foot}
</table></td></tr></table></body></html>"""


if __name__ == "__main__":
    data = json.loads((HERE / "fin_input.json").read_text(encoding="utf-8-sig"))
    out = HERE / "fin_output.html"
    out.write_text(build(data), encoding="utf-8")
    print(f"Wrote {out}")
    if "--open" in sys.argv:
        webbrowser.open(out.as_uri())
