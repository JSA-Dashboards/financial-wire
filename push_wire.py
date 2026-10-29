"""Run on your PC each morning (Task Scheduler, ~6:00 AM CT weekdays):
read Outlook, then push the results to Snowflake for the hosted app.

    python push_wire.py            # fetch + push
    python push_wire.py --dry      # fetch only, print what would be pushed
    python push_wire.py --clean    # 4:00 AM: wipe yesterday (nothing is pulled)
    python push_wire.py --scheduled  # 6:30-8:30 AM task: full pull to 7:10, then only on a Refresh request
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import futures_changes
import headlines
import wire_store

HERE = Path(__file__).parent
DRY = "--dry" in sys.argv
SCHEDULED = "--scheduled" in sys.argv
CLEAN = "--clean" in sys.argv
TRIED = HERE / ".last_scheduled_try"   # stops a failing pull from retrying every 10 minutes


def should_run() -> bool:
    """--scheduled fires every 10 minutes from 6:30 AM Central. Through 7:10 every run is a full pull
    (the emails land one by one); after that only when the app's Refresh button has asked for one."""
    import time
    now = wire_store.ct_now()
    if 6 * 60 + 30 <= now.hour * 60 + now.minute <= 7 * 60 + 14:
        return True
    if not wire_store.refresh_pending():
        return False
    if TRIED.exists() and time.time() - TRIED.stat().st_mtime < 1800:
        return False   # tried within 30 min and the request is still open: the pull is failing, don't hammer Outlook
    TRIED.write_text(now.isoformat())
    return True


def clean_slate() -> int:
    """4:00 AM: drop everything saved for yesterday, here and in Snowflake. Pulls nothing."""
    n = wire_store.clear_all()
    for name in ("fin_input.json", "pj_candidates.json", "bbg_candidates.json", "mkts.json",
                 "headline_candidates.json", "headline_candidates.html", "fin_output.html"):
        (HERE / name).unlink(missing_ok=True)
    TRIED.unlink(missing_ok=True)
    print(f"Clean slate: removed {n} saved row(s) from Snowflake and the local copies.")
    return 0


def main() -> int:
    wire_store.config()   # loads financial_wire/.env into the environment (Snowflake + MASSIVE_API_KEY)
    if CLEAN:
        return clean_slate()
    if SCHEDULED and not should_run():
        return 0
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
    key = os.environ.get("MASSIVE_API_KEY")
    if key:
        try:
            payloads["mkts"] = futures_changes.build(key, as_of)
            (HERE / "mkts.json").write_text(json.dumps(payloads["mkts"]), encoding="utf-8")
        except Exception as e:   # a Massive blip must not block the rest of the wire
            print("FUTURES ERROR:", e)
    else:
        print("MASSIVE_API_KEY not set in financial_wire/.env - skipping the futures table.")
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
