#!/usr/bin/env python3
"""Day-by-day change log for the dashboard (the data behind the Changes view).

Each run saves the day's latest schedule as history/snapshots/YYYY-MM-DD.json (overwritten by every
run that day, so it is the end-of-day state). Today's snapshot is compared with the last snapshot
of an earlier day, and the differences are written to history/changes.json:

    {"days": {"2026-10-26": {"vs": "2026-10-25", "events": [ {...}, ... ]}, ...}}

An event is {"type": "changed" | "removed", pol, service, vessel, vessel_code, vyg_bound,
chg_h (changed only; + delayed, - advanced), "old": {etb, etd, dry, rf}, "new": {...} (changed only)}.
  changed  = ETB or ETD moved by 30 minutes or more
  removed  = the call was in the earlier snapshot, is gone now, and its ETA lies inside the date
             range the new snapshot covers (calls that merely left the 2-month window do not count)
"""
import json
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent
SNAP_DIR = HERE / "history" / "snapshots"
LOG = HERE / "history" / "changes.json"
KEEP = ("vessel_code", "vyg_bound", "pol", "service", "vessel", "eta", "etb", "etd", "cutoff_dry", "cutoff_reefer")
FMT = "%Y-%m-%d %H:%M"


def compact(records):
    return [{k: r[k] for k in KEEP} for r in records]


def _key(r):
    return (r["vessel_code"], r["vyg_bound"], r["pol"])


def _hours(h):
    return int(h) if h == int(h) else round(h * 2) / 2


def _times(r):
    return {"etb": r["etb"], "etd": r["etd"], "dry": r["cutoff_dry"], "rf": r["cutoff_reefer"]}


def diff(prev, cur):
    """Events that turn the earlier snapshot `prev` into the later one `cur`."""
    P = lambda t: datetime.strptime(t, FMT)
    pm = {_key(o): o for o in prev}
    cm = {_key(r): r for r in cur}
    events = []
    base = lambda r: {k: r[k] for k in ("pol", "service", "vessel", "vessel_code", "vyg_bound")}
    for k, r in cm.items():
        o = pm.get(k)
        if not o:
            continue
        h_b = (P(r["etb"]) - P(o["etb"])).total_seconds() / 3600
        h_d = (P(r["etd"]) - P(o["etd"])).total_seconds() / 3600
        if abs(h_b) >= 0.5 or abs(h_d) >= 0.5:
            events.append({"type": "changed", **base(r), "chg_h": _hours(h_b if abs(h_b) >= 0.5 else h_d),
                           "old": _times(o), "new": _times(r)})
    if cur:
        lo, hi = min(r["eta"] for r in cur), max(r["eta"] for r in cur)
        for k, o in pm.items():
            if k not in cm and lo <= o["eta"] <= hi:
                events.append({"type": "removed", **base(o), "old": _times(o)})
    events.sort(key=lambda e: (e["pol"] != "THBKK", (e.get("new") or e["old"])["etb"]))
    return events


def _load_log():
    try:
        return json.loads(LOG.read_text(encoding="utf-8"))
    except Exception:
        return {"days": {}}


def _earlier_snapshot(day):
    """(date, snapshot) of the latest saved snapshot before `day`, or (None, None)."""
    days = sorted(p.stem for p in SNAP_DIR.glob("*.json") if p.stem < day)
    if not days:
        return None, None
    return days[-1], json.loads((SNAP_DIR / f"{days[-1]}.json").read_text(encoding="utf-8"))


def record_day(records, day):
    """Save `records` as the snapshot of `day` (YYYY-MM-DD) and refresh that day's change events.

    Returns the whole change log (dict).
    """
    SNAP_DIR.mkdir(parents=True, exist_ok=True)
    cur = compact(records)
    log = _load_log()
    vs, prev = _earlier_snapshot(day)
    if prev is not None:
        log["days"][day] = {"vs": vs, "events": diff(prev, cur)}
    (SNAP_DIR / f"{day}.json").write_text(json.dumps(cur, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    LOG.parent.mkdir(parents=True, exist_ok=True)
    LOG.write_text(json.dumps(log, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    return log


def load_log():
    return _load_log()
