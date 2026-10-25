#!/usr/bin/env python3
"""One-off: rebuild history/ from every published version of index.html in git.

For each Bangkok calendar day the last committed index.html is that day's snapshot; consecutive days
are then compared with history.record_day(). Safe to re-run (snapshots/log are rewritten).
"""
import json
import re
import subprocess
from datetime import datetime, timedelta, timezone

import history

BKK = timezone(timedelta(hours=7))


def git(*a):
    return subprocess.run(["git", *a], capture_output=True, encoding="utf-8", check=True).stdout


def main():
    last_of_day = {}
    for line in git("log", "--reverse", "--format=%H %cI", "--", "index.html").splitlines():
        sha, iso = line.split()
        day = datetime.fromisoformat(iso).astimezone(BKK).strftime("%Y-%m-%d")
        last_of_day[day] = sha  # later commits of the same day overwrite earlier ones
    n = 0
    for day in sorted(last_of_day):
        html = git("show", f"{last_of_day[day]}:index.html")
        m = re.search(r"const RAW = (\[.*?\]);\nconst WHARF_MAP", html, re.S)
        if not m:
            continue
        raw = json.loads(m.group(1))
        if not raw or "etb" not in raw[0]:
            print(f"{day}: older page layout without ETB - skipped")
            continue
        log = history.record_day(raw, day)
        ev = log["days"].get(day)
        print(f"{day}: {len(raw)} calls" + (f", {len(ev['events'])} events vs {ev['vs']}" if ev else " (first snapshot)"))
        n += 1
    print(f"{n} days rebuilt")


if __name__ == "__main__":
    main()
