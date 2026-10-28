"""Run on your PC each morning (Task Scheduler, ~6:00 AM CT weekdays):
read Outlook, then push the results to Snowflake for the hosted app.

    python push_wire.py            # fetch + push
    python push_wire.py --dry      # fetch only, print what would be pushed
"""
import json
import subprocess
import sys
from pathlib import Path

import headlines
import wire_store

HERE = Path(__file__).parent
DRY = "--dry" in sys.argv


def main() -> int:
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                    str(HERE / "fetch_fin.ps1")], check=True, timeout=300)
    wire = json.loads((HERE / "fin_input.json").read_text(encoding="utf-8-sig"))

    errors: list[str] = []
    pj = headlines.pj_items(errors)
    bbg = headlines.bbg_items(errors)
    for e in errors:
        print("MAILBOX ERROR:", e)

    as_of = wire_store.today_ct()
    payloads = {"wire": wire, "pj": pj, "bbg": bbg}
    for kind, data in payloads.items():
        n = len(data) if isinstance(data, list) else len(wire.get("macro", [])) + len(wire.get("snapshot", []))
        print(f"{kind:5} {n} item(s)")
    if DRY:
        return 0

    # Never overwrite a good day with an empty one (Outlook closed, COM failure).
    if not wire.get("snapshot") and not wire.get("macro") and not pj and not bbg:
        print("Nothing fetched - leaving Snowflake untouched.")
        return 2
    for kind, data in payloads.items():
        wire_store.push(kind, data, as_of)
    print(f"Pushed {as_of} to Snowflake.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
