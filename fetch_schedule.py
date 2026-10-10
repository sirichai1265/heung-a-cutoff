#!/usr/bin/env python3
"""
Pull the vessel schedule for Bangkok (THBKK) and Laem Chabang (THLCH) straight
from the Heung-A e-Service site, for the dashboard's automatic refresh.

Source: https://ebiz.heungaline.com/Schedule/GetScheduleJsonData  (public JSON,
the same call the Schedule page makes). Uses only the Python standard library.

What the site gives us, per port P:
    Outbound (bnd=O, pol=P)  -> every departure from P:  ETD, wharf, next ports
    Inbound  (bnd=I, pod=P)  -> every arrival at P:      ETA at P (the berthing time)

How that maps onto the dashboard's fields (checked against the internal CUT
file: wharf and next port match 100%, ETD ~97%, and the site's arrival time
equals the file's ETB, with the file's ETA being ETB - 1h in 66 of 69 cases):
    ETB = site arrival time        ETA = ETB - 1h        ETD = site departure

Limitation: the site has no arrival row for voyages that start their rotation
at Bangkok / Laem Chabang (about 45% of calls). For those the ETA/ETB are taken
from a matching row of the latest internal CUT .xls if one exists, otherwise
ESTIMATED as ETD minus the median port stay seen in the rest of the data. Such
rows carry eta_src="est" so the dashboard can label them.
"""
import json
import statistics
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta

BASE = "https://ebiz.heungaline.com/Schedule/GetScheduleJsonData"
PORTS = ("THBKK", "THLCH")
# Typical port stay (hours, ETB->ETD) in the internal CUT files; used only when the site
# publishes no arrival times for a port. Bangkok ~30h, Laem Chabang ~11h.
DEFAULT_DWELL_H = {"THBKK": 30, "THLCH": 11}
FMT = "%Y-%m-%d %H:%M"
# The server answers HTTP 500 unless an Accept-Language header is present.
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://ebiz.heungaline.com/Schedule",
}


def _get(params, tries=4):
    url = BASE + "?" + urllib.parse.urlencode(params)
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=120) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:  # network blip / 5xx: back off and retry
            last = e
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"schedule request failed after {tries} tries: {last}")


def fetch_feed(today=None, months_ahead_days=62, back_out_days=3, back_in_days=14):
    """Return {port: {'O': [...rows], 'I': [...rows]}} for a 2-month window."""
    today = today or datetime.now()
    to = (today + timedelta(days=months_ahead_days)).strftime("%Y%m%d") + "2400"
    feed = {}
    for port in PORTS:
        feed[port] = {
            "O": _get({"bnd": "O", "pol": port, "pod": "", "todt": to,
                       "fmdt": (today - timedelta(days=back_out_days)).strftime("%Y%m%d")}),
            "I": _get({"bnd": "I", "pol": "", "pod": port, "todt": to,
                       "fmdt": (today - timedelta(days=back_in_days)).strftime("%Y%m%d")}),
        }
    return feed


def build_records(feed, derive_times, known=None):
    """Turn the raw feed into dashboard records.

    derive_times(pol, eta_dt, etd_dt) -> (cutoff_dry, cutoff_reefer, opengate)
    known: records from an internal CUT .xls, used to fill missing arrivals.
    Returns (records, wharf_names, stats).
    """
    P = lambda s: datetime.strptime(s, FMT)
    known_by_key = {}
    op_liner = {}
    for k in known or []:
        known_by_key[(k["vessel_code"], k["vyg_bound"], k["pol"], k["etd"])] = k
        op_liner.setdefault(k["vessel_code"], k["op_liner"])

    deps = {}
    arrs = defaultdict(list)
    wharf_names = {}
    for port, d in feed.items():
        for r in d["O"]:
            wharf_names[r["POLW"]] = r["POLWNM"].strip()
            key = (r["VSL"], r["VYG"], port, r["ETD"])
            if key not in deps or r["ETA"] < deps[key]["ETA"]:  # first (nearest) next port
                deps[key] = r
        for r in d["I"]:
            wharf_names[r["PODW"]] = r["PODWNM"].strip()
            arrs[(r["VSL"], r["VYG"], port)].append(r)

    # first pass: real arrivals
    rows, dwell = [], defaultdict(list)
    for (vsl, vyg, port, etd), r in deps.items():
        etd_dt = P(etd)
        cands = [a for a in arrs.get((vsl, vyg, port), [])
                 if P(a["ETA"]) <= etd_dt and etd_dt - P(a["ETA"]) <= timedelta(days=7)]
        a = max(cands, key=lambda a: a["ETA"]) if cands else None
        if a:
            dwell[port].append((etd_dt - P(a["ETA"])).total_seconds() / 3600)
        rows.append((vsl, vyg, port, r, a))

    median_dwell = {p: statistics.median(v) for p, v in dwell.items() if v}
    stats = {"web": 0, "xls": 0, "est": 0, "median_dwell_h": median_dwell}

    records = []
    for vsl, vyg, port, r, a in rows:
        etd_dt = P(r["ETD"])
        key = (vsl, vyg, port, r["ETD"])
        if a:
            etb_dt, src = P(a["ETA"]), "web"
        elif key in known_by_key:
            etb_dt, src = P(known_by_key[key]["etb"]), "xls"
        else:
            etb_dt = etd_dt - timedelta(hours=median_dwell.get(port, DEFAULT_DWELL_H[port]))
            src = "est"
        stats[src] += 1
        eta_dt = etb_dt - timedelta(hours=1)
        dry, reefer, opengate = derive_times(port, eta_dt, etd_dt)
        if src == "xls":
            eta_dt = P(known_by_key[key]["eta"])
            dry, reefer, opengate = derive_times(port, eta_dt, etd_dt)
        records.append({
            "service": r["SVC"], "vessel_code": vsl, "vessel": r["VSLNM"].strip(),
            "op_liner": op_liner.get(vsl, ""), "seq": "", "vyg": vyg[:-1], "bound": vyg[-1:],
            "vyg_bound": vyg, "wharf": r["POLW"], "pol": port, "pod": r["POD"],
            "no_d": "N", "no_l": "N", "used": "",
            "eta": eta_dt.strftime(FMT), "etb": etb_dt.strftime(FMT), "etd": etd_dt.strftime(FMT),
            "cutoff_dry": dry.strftime(FMT), "cutoff_reefer": reefer.strftime(FMT),
            "opengate": opengate.strftime(FMT), "eta_src": src,
        })
    records.sort(key=lambda x: x["eta"])
    return records, wharf_names, stats
