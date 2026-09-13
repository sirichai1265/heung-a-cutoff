#!/usr/bin/env python3
"""
Rebuild "Cut off / Open gate dashboard.html" from a daily CUT_OFF-style .xls file
and a Wharf.xls code-to-name lookup.

Usage:
    python build_dashboard.py 9-9-CUT.xls
        -> rebuilds both "Cut off - Open Gate Dashboard.html" and index.html
           using the built-in wharf map.
    python build_dashboard.py 9-9-CUT.xls Wharf.xls   # optional explicit wharf lookup

Or just run update.ps1 (see that file) to rebuild + git push in one step.

Requirements:
    pip install xlrd

Rules applied (matches the standing spec):
    - Cut off (Dry)    = ETA - 24 hours
    - Cut off (Reefer) = ETA - 1 hour
    - 1st Return       = ETD - 5 days (counting ETD as day 1, time ignored)
    - Rows where Skip == "Y" are excluded
    - Wharf codes are resolved to full names using Wharf.xls (code shown underneath)
"""
import argparse
import json
import sys
from datetime import datetime, timedelta

try:
    import xlrd
except ImportError:
    sys.exit("Missing dependency. Install it first:  pip install xlrd")


# Wharf code -> full name. Used when no Wharf.xls is supplied. Edit here if a
# new wharf code appears (the build prints a warning listing any it doesn't know).
DEFAULT_WHARF_MAP = {
    "BKK00": "BKK LESSOR'S DEPOT",
    "BKK01": "PAT TERMINAL 2 (PORT AUTHORITY OF THAILAND)",
    "BKK02": "UNITHAI CONTAINER TERMINAL",
    "BKK03": "BDS TERMINAL",
    "BKK04": "PAT TERMINAL 1 (PORT AUTHORITY OF THAILAND)",
    "BKK05": "THAI HANJIN LAT KRABANG",
    "BKK06": "THAI HANJIN LATKRABANG K.PUKPIK",
    "BKK07": "STAR PACIFIC",
    "BKK08": "THAI INTER DEPOT AND TRANSPORT",
    "BKK09": "GREATING FORTUNE CONTAINER SERVICE (THAILAND) CO.,LTD",
    "BKK10": "Thai Sugar Container Terminal",
    "BKK11": "Siam River Port Co., Ltd (SRP)",
    "BKK21": "TIGER DEPOT",
    "BKK22": "555",
    "BKK23": "CDS(CONTAINER DEPOT SERVICE CO LTD)",
    "BKK24": "YJC DEPOT SERVICE CO LTD",
    "BKK25": "Smart Logistics Service (Thailand ) Co.,Ltd.",
    "BKK26": "B.C. DEPOT CO.,LTD. (KLONGTOEY)",
    "BKK27": "B.C. DEPOT CO.,LTD. ( Bang-Na KM.18)",
    "BKK28": "YJC DEPOT SERVICES CO LTD ( YJC BKK2 )",
    "BKK99": "SHPR OR CNEE'S PREMISE",
    "BKKF1": "FMC SERVICES KM.21",
    "BKKM1": "PORT AUTHORITY OF THAILAND (P.A.T.)",
    "BKKM3": "SAHATHAI COASTAL SEAPORT CO LTD / Code 0513",
    "BKKM4": "Thai Prosperity Terminal (TPT)",
    "BKKY4": "BMT PACIFIC LTD.",
    "BKKY5": "Sintanachote Co.,Ltd.",
    "BKKY6": "B.C. DEPOT CO,LTD. (BANG NA-TRAD  KM. 13)",
    "BKKZZ": "BKK LOADING/DISCHARGING PIER",
    "LCH00": "LCH LESSOR'S DEPOT",
    "LCH01": "ESCO (EASTERN SEA LCH CNTR TML/B3)",
    "LCH02": "A2 ( Thai Laemchabang Terminal, TLT / 허치슨 )",
    "LCH03": "B4 TIPS CONTAINERR TERMINAL",
    "LCH04": "LCMT Company LTD, ( under LCB1 Group)  A0",
    "LCH05": "B5 LCIT (LAEM CHABANG INTERNATIONAL TERMINAL CO., LTD)",
    "LCH06": "A3 (Hutchison Laemchabang Terminal Limited, HLT)",
    "LCH07": "B1 ( LCB container terminal 1 , LCB1 )",
    "LCH08": "Hutchison laemchabang terminal( C1,C2 )",
    "LCH09": "Hutchison laemchabang terminal( D1 )",
    "LCH10": "LCIT(C3) LAEM CHABANG INTERNATIONAL TERMINAL CO., LTD",
    "LCH20": "YJC DEPOT (YJC THAILAND)",
    "LCH21": "Reefer Express CO.,LTD.",
    "LCH22": "CONTPOOL LCH",
    "LCH23": "SINGSAMUT 28 DEPOT CO LTD",
    "LCH24": "CDS (CONTAINER DEPOT SERVICE)",
    "LCH25": "SCS Yard Co Ltd.",
    "LCH26": "PH Depot Co Ltd ( Laemchabang )",
    "LCH27": "HAST Logistics Co.,Ltd.",
    "LCH28": "CELLO",
    "LCH51": "THAI HANJIN LATKRABANG",
    "LCH52": "SCT (Siam Container Terminal)",
    "LCH53": "JWD (FOR DG CARGO)",
    "LCH54": "TIFFA ICD LKB",
    "LCH55": "ICD LKT - Esco Ladkrabang (GATE 2)",
    "LCH99": "SHPR OR CNEE'S PREMISE",
    "LCHA3": "Hutchison Laemchabang Terminal Ltd. A3 / Code 2829",
    "LCHB1": "LCB CONTAINER TERMINAL 1 LTD. (B1) / Code 2811",
    "LCHB2": "EVERGREEN CONTAINER TERMINAL (THAILAND) LTD. / Code 2812",
    "LCHB3": "EASTERN SEA LAEM CHABANG TERMINAL CO. LTD. / Code 2813",
    "LCHB4": "TIPS CO. LTD. B4 / Code 2814",
    "LCHC1": "HUTCHISON LAEMCHABANG TERMINAL LTD (C1&C2) / Code 2836",
    "LCHD1": "Hutchison laemchabang Terminal LTD ( D1 ) (customs code 2840)",
    "LCHE1": "THAI ENGKONG LCH",
    "LCHF1": "FORTRESS LCH",
    "LCHG1": "GREATING FORTUNE LCH",
    "LCHKC": "Kittichai  Container  Depot  Co.,Ltd.",
    "LCHL1": "LAEMCHABANG INTER DEPOT",
    "LCHM1": "THAI LAEMCHABANG TERMINAL CO. LTD.",
    "LCHM2": "HUTCHISON LAEMCHABANG TERMINAL LIMITED",
    "LCHM3": "LCB CONTAINER TERMINAL 1 LTD.",
    "LCHM4": "LAEM CHABANG INTERNATIONAL TERMINAL CO. LTD.",
    "LCHM5": "TIPS CO. LTD.",
    "LCHM6": "EASTERN SEA LAEM CHABANG TERMINAL CO. LTD.",
    "LCHM7": "KERRY SIAM SEAPORT LIMITED.",
    "LCHS1": "SRITHAI FREIGHT FORWARDER LCH",
    "LCHY2": "JWD INFO LOGISTICS CO., LTD",
    "LCHY3": "EVERGREEN CONTAINER TERMINAL (THAILAND) LTD.",
    "LCHY4": "SIAM CONTAINER TRANSPORT AND TERMINAL CO., LTD",
    "LCHY5": "PW DEPOT CO., LTD.",
    "LCHY6": "Smart Logistics Service (Thailand ) Co.,Ltd. / Laem Chabang",
    "LCHZZ": "LCH LOADING/DISCHARGING PIER",
}


def parse_dt(v):
    if isinstance(v, str):
        return datetime.strptime(v.strip(), "%Y-%m-%d %H:%M")
    return v


def load_cutoff_records(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    records = []
    skipped = 0
    for r in range(1, sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        (service, vessel_code, vessel_name, op_liner, seq, vyg, bound,
         vyg_bound, wharf, pol, pod, skip, no_d, no_l, used,
         eta, etb, etd) = row

        if skip == "Y":
            skipped += 1
            continue

        eta_dt = parse_dt(eta)
        etb_dt = parse_dt(etb)
        etd_dt = parse_dt(etd)

        cutoff_dry = eta_dt - timedelta(hours=24)
        cutoff_reefer = eta_dt - timedelta(hours=1)
        # 1st Return: 5 calendar days counting ETD as day 1, time ignored
        etd_date = datetime(etd_dt.year, etd_dt.month, etd_dt.day)
        opengate = etd_date - timedelta(days=4)

        records.append({
            "service": service, "vessel_code": vessel_code, "vessel": vessel_name,
            "op_liner": op_liner, "seq": seq, "vyg": vyg, "bound": bound,
            "vyg_bound": vyg_bound, "wharf": wharf, "pol": pol, "pod": pod,
            "no_d": no_d, "no_l": no_l, "used": used,
            "eta": eta_dt.strftime("%Y-%m-%d %H:%M"),
            "etb": etb_dt.strftime("%Y-%m-%d %H:%M"),
            "etd": etd_dt.strftime("%Y-%m-%d %H:%M"),
            "cutoff_dry": cutoff_dry.strftime("%Y-%m-%d %H:%M"),
            "cutoff_reefer": cutoff_reefer.strftime("%Y-%m-%d %H:%M"),
            "opengate": opengate.strftime("%Y-%m-%d %H:%M"),
        })

    records.sort(key=lambda x: x["eta"])
    print(f"Loaded {len(records)} voyages ({skipped} Skip=Y rows excluded) from {xls_path}")
    return records


def load_wharf_map(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    m = {}
    for r in range(1, sh.nrows):
        code = str(sh.cell_value(r, 0)).strip()
        name = str(sh.cell_value(r, 1)).strip()
        if code:
            m[code] = name
    print(f"Loaded {len(m)} wharf codes from {xls_path}")
    return m


def check_wharf_coverage(records, wharf_map):
    used = {r["wharf"] for r in records}
    missing = used - set(wharf_map.keys())
    if missing:
        print(f"Warning: {len(missing)} wharf code(s) have no name mapping "
              f"(add them to DEFAULT_WHARF_MAP in build_dashboard.py): {sorted(missing)}")


TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>HEUNG A LINE — Cut Off / Open Gate Dashboard - THBKK & THLCH</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --navy:#0b2540; --navy2:#123a5e; --navy3:#1a4a72;
    --canvas:#eef1f2; --card:#ffffff; --line:#dde3e7; --line-strong:#c3ccd2;
    --text:#132230; --text2:#586975; --muted:#93a2ac;
    --bkk:#2563a6; --bkk-bg:#e7f0fa;
    --lch:#1f8a72; --lch-bg:#e4f4ef;
    --coral:#c1483f; --coral-bg:#fbebe9;
    --amber:#c1791f; --amber-bg:#faf0dd; --amber-strong:#dd9a35;
    --teal:#1a7862; --teal-bg:#e4f4ef;
  }
  *{box-sizing:border-box;}
  body{margin:0;font-family:'Space Grotesk',-apple-system,'Segoe UI',sans-serif;background:var(--canvas);color:var(--text);}
  .mono{font-family:'JetBrains Mono',ui-monospace,monospace;}
  .wrap{max-width:none;margin:0;padding:0 16px 60px;}

  /* Hero */
  .hero{background:var(--navy);background-image:linear-gradient(160deg,var(--navy) 0%,#0e2f52 60%,var(--navy2) 100%);color:#fff;padding:28px 20px 24px;margin-bottom:22px;}
  .hero-inner{max-width:none;margin:0;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;}
  .hero h1{font-size:24px;font-weight:700;margin:0 0 6px;letter-spacing:-.01em;}
  .hero p{font-size:13px;color:#a9bccd;margin:0;max-width:520px;line-height:1.5;}
  .hero-mark{display:flex;align-items:center;gap:8px;color:#7fa8cb;font-size:11px;text-transform:uppercase;letter-spacing:.08em;font-weight:600;}
  .hero-clock{display:flex;align-items:center;gap:7px;font-family:'JetBrains Mono',ui-monospace,monospace;font-size:12.5px;color:#c9dbea;margin-top:12px;letter-spacing:.02em;}
  .hero-clock .dot{width:6px;height:6px;border-radius:50%;background:#3ecf8e;box-shadow:0 0 0 3px rgba(62,207,142,.2);flex:none;}
  .hero-logo{background:#fff;border-radius:12px;padding:12px 20px;flex:none;box-shadow:0 2px 12px rgba(0,0,0,.2);}
  .hero-logo img{height:64px;width:auto;display:block;}
  @media (max-width:760px){ .hero-logo{padding:9px 14px;} .hero-logo img{height:44px;} }


  /* Board summary strip */
  .board{background:var(--navy);border-radius:14px;padding:4px 4px;margin-bottom:20px;display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));overflow:hidden;}
  .board .seg-stat{padding:14px 18px;border-right:1px dashed rgba(255,255,255,.16);}
  .board .seg-stat:last-child{border-right:none;}
  .board .seg-stat .lbl{font-size:10.5px;color:#8ba6bf;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;}
  .board .seg-stat .val{font-family:'JetBrains Mono';font-size:19px;font-weight:600;color:#fff;line-height:1.25;}
  .board .seg-stat .val.small{font-size:12.5px;font-weight:500;}
  .board .seg-stat .val .amber{color:var(--amber-strong);}

  .controls{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px;}
  .controls input[type=text]{flex:1;min-width:180px;padding:10px 13px;border:1px solid var(--line-strong);border-radius:9px;font-size:13.5px;background:var(--card);font-family:inherit;}
  .controls select{padding:10px 11px;border:1px solid var(--line-strong);border-radius:9px;font-size:13px;background:var(--card);color:var(--text);font-family:inherit;}
  .seg{display:flex;border:1px solid var(--line-strong);border-radius:9px;overflow:hidden;background:var(--card);}
  .seg button{border:none;background:transparent;padding:9px 14px;font-size:13px;cursor:pointer;color:var(--text2);font-family:inherit;}
  .seg button.active{background:var(--navy);color:#fff;}
  .seg button:not(:last-child){border-right:1px solid var(--line);}
  .clearbtn{padding:9px 13px;border:1px solid var(--line-strong);border-radius:9px;background:var(--card);font-size:13px;cursor:pointer;color:var(--text2);font-family:inherit;}
  .clearbtn:hover{background:#f3f5f6;}

  #tableHolder{overflow:auto;max-height:calc(100vh - 16px);border:1px solid var(--line);border-radius:12px;}
  table{width:100%;max-width:none;table-layout:fixed;border-collapse:separate;border-spacing:0;background:var(--card);font-size:12.8px;}
  thead th{position:sticky;top:0;z-index:6;background:var(--navy);text-align:left;padding:11px 12px;font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:#9fb7cc;font-weight:600;cursor:pointer;white-space:nowrap;box-shadow:0 2px 6px rgba(11,37,64,.35);overflow:hidden;text-overflow:ellipsis;}
  thead th:hover{color:#fff;}
  thead th.sorted{color:var(--amber-strong);}
  tbody td{padding:10px 12px;border-bottom:1px solid var(--line);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  th.wharfcell, td.wharfcell{white-space:normal;}
  th.polcell, td.polcell{padding-left:8px;padding-right:8px;text-align:center;overflow:visible;}
  /* column widths: pol, service, vessel, vessel-name, wharf (same as vessel-name), next-port, then the 5 date columns share the rest */
  thead th:nth-child(1), tbody td:nth-child(1){width:8%;}
  thead th:nth-child(2), tbody td:nth-child(2){width:5%;}
  thead th:nth-child(3), tbody td:nth-child(3){width:6%;}
  thead th:nth-child(4), tbody td:nth-child(4){width:12.25%;}
  thead th:nth-child(5), tbody td:nth-child(5){width:12.25%;}
  thead th:nth-child(6), tbody td:nth-child(6){width:5.5%;}
  thead th:nth-child(7), tbody td:nth-child(7),
  thead th:nth-child(8), tbody td:nth-child(8),
  thead th:nth-child(9), tbody td:nth-child(9),
  thead th:nth-child(10), tbody td:nth-child(10),
  thead th:nth-child(11), tbody td:nth-child(11){width:10.2%;}
  .dtcell.dtcell-lg .d{font-size:15px;font-weight:600;}
  .dtcell.dtcell-lg .t{font-size:13px;}
  tbody tr{position:relative;}
  tbody tr td:first-child{box-shadow:inset 4px 0 0 0 var(--rowaccent, transparent);}
  tbody tr:last-child td{border-bottom:none;}
  tbody tr:hover{background:#f6f8f9;}
  .pol-badge{font-size:11px;font-weight:600;padding:2px 9px;border-radius:5px;font-family:'JetBrains Mono';}
  .pol-bkk{background:var(--bkk-bg);color:var(--bkk);}
  .pol-lch{background:var(--lch-bg);color:var(--lch);}
  .vessel{font-weight:600;}
  .vcode{font-family:'JetBrains Mono';font-size:10.5px;font-weight:600;color:var(--text2);background:var(--canvas);border:1px solid var(--line-strong);padding:0 5px;border-radius:4px;margin-left:6px;vertical-align:1px;}
  .sub{color:var(--muted);font-size:11.5px;margin-top:1px;}
  .dtcell{display:flex;flex-direction:column;line-height:1.4;font-family:'JetBrains Mono';}
  .dtcell .d{font-weight:500;}
  .dtcell .t{color:var(--muted);font-size:11.3px;}
  .flag{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:5px;margin-left:5px;vertical-align:1px;text-transform:uppercase;letter-spacing:.03em;font-family:'Space Grotesk';}
  .flag-soon{background:var(--amber-bg);color:var(--amber);}
  .flag-past{background:var(--coral-bg);color:var(--coral);}
  .flag-open{background:var(--teal-bg);color:var(--teal);}
  .noload{background:#eef0e9;color:#65695a;font-size:10px;padding:1px 6px;border-radius:5px;font-family:'Space Grotesk';}
  .empty{padding:50px 20px;text-align:center;color:var(--muted);font-size:14px;}
  .count{font-size:12px;color:var(--muted);margin:10px 2px 0;font-family:'JetBrains Mono';}

  /* Mobile ticket-stub card view */
  .mcards{display:flex;flex-direction:column;gap:14px;}
  .mcard{display:flex;background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;position:relative;}
  .mcard-main{flex:1;padding:16px 18px;min-width:0;}
  .mcard-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;}
  .mcard-vessel{font-weight:700;font-size:17px;}
  .mcard-sub{color:var(--muted);font-size:13px;margin-top:3px;}
  .mcard-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px 18px;margin-top:14px;font-family:'JetBrains Mono';}
  .mcard-item .lbl{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:5px;font-family:'Space Grotesk';}
  .mcard-item .val{font-size:16px;font-weight:600;line-height:1.35;}
  .mcard-item .val .t{color:var(--muted);font-size:13.5px;font-weight:400;margin-left:5px;}
  .mcard-item .val .flag{font-size:11px;padding:2px 8px;}
  .mcard-item .val .noload{font-size:11px;}
  .mcard-vessel .vcode{font-size:11.5px;}
  .mcard-item.full{grid-column:1 / -1;}
  .mcard-stub{width:46px;flex:none;background:var(--navy);position:relative;display:flex;align-items:center;justify-content:center;background-image:linear-gradient(180deg,var(--navy),#0e2f52);}
  .mcard-stub::before{content:'';position:absolute;left:0;top:0;bottom:0;width:0;border-left:2px dashed rgba(255,255,255,.3);}
  .mcard-stub::after{content:'';position:absolute;top:-7px;left:-7px;width:14px;height:14px;border-radius:50%;background:var(--canvas);box-shadow:0 calc(100% - 14px) 0 0 var(--canvas);}
  .mcard-stub-pol{writing-mode:vertical-rl;transform:rotate(180deg);color:#fff;font-family:'JetBrains Mono';font-weight:600;font-size:12.5px;letter-spacing:.08em;}

  @media (max-width:760px){
    #tableHolder{max-height:none;overflow:visible;border:none;}
    .board{grid-template-columns:repeat(2,1fr);}
    .hero h1{font-size:20px;}
  }
</style>
</head>
<body>
<div class="hero">
  <div class="hero-inner">
    <div>
      <h1>Cut off and 1st return</h1>
      <div class="hero-clock"><span class="dot"></span><span id="liveClock"></span></div>
    </div>
    <div class="hero-logo">
      <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAiQAAAB3CAYAAAAzdlsAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAASdEVYdFNvZnR3YXJlAEdyZWVuc2hvdF5VCAUAAD9ESURBVHhe7Z0HlBVF+vZ3/9+awxpW4gQkI4iAKCBJkohEyRkJEgQkiAgCguQsWXLOSE4DEhQBSRIkSFBJirKKWddc33lqpnqq3+6+3X1v901Tv3OeA3Nvdbid6umqt976B1ME5MCBA6xq5cqsQN68LGdyckRUuFAhNnL4cLprCoVCoVDEDf+gHygYmz1rFiuYP7/BGFDlS05mxZKSWf2EJDY6ayLb8UACe8+F9jyQwBZlTmCDsyayBglJrEiScRtmKlK4MFu/bh3dbYVCoVAoYhZlSNLo2rkzy5Ujh6Hyh3IlJ7MCycmsefYktjmT0Vj4pbFZElnpxNTt032Ccj/4IBs2ZAj9KQqFQqFQxBwZ2pAcOniQlStd2lDRQ/mTk1m3bEnsXROjEAm980ACey1rEiubaNxXKG+uXGzJ4sX0JyoUCoVCERNkSEPSp3dv09aQx5OS2bzMiWy3iSGINs3NnMiKm3Tx4HdNnjiR/mSFQqFQKKKaDGVIJk2YYDAi6A5plj2J7TSp9GNFazInsIYJSQZzkidnTrZq5Up6GBQKhUKhiDoyhCFp3rQpj7eQK+vHkpLZskyJPLCUVvCxrIWZE1lR0nICE9a7Vy96WBQKhUKhiBri2pDUql7d0GrQNnv0xIX4JXQ5PZ89iRUmxqR927b0ECkUCoVCERXErSGpW7u2vvsiOZmPWqGVd7xrWNYklpuYshc6dqSHS6FQKBSKiBJ3hmTpkiWsQL58WuWLXCHTMqARoRqfJZGbMtmYDB40iB4+hUKhUCgiQtwYklkzZ+oCVlH5TlBGRKeUTAmsaqK+xaRyxYr0UCoUCoVCEXbiwpB0bN9e9+ZfITG8CcxiTeszJbAnpPiSwgULstOnTtHDqlAoFApF2Ih5Q1K5QgWtYn0kKZl1zp4UdyNn/BASrTUlQ4VbtWhBD69CoZD49fe/2Naj19izo95ltzdZzv5Rb7Gp/tVwCSvbfzt7eeFRduziN3Q1CoXChJg1JAsXLGCPFSumVaYVE5PYFtUq4lr9sulNyTNPP00PddTRbe4RdkdT68pAKNcL7uf76TrnsGE9ZsrRcS27eP0nujh7dclxQ1mvlNPk99zaZJmhHNU9LVeyebs+oYsGBMeOrofqruYr2MAVJ+iinFFrTxvKm+n2pstZtaG76OKOyd15vWGdZrqtyTK6qGN2nfqSZWm72rBON8L1OmP7BbrqsJKni7NjJStX5/XchIUDp9cMhPs0FNq9ecCwTjOhnBWP99lqKG+mf9ZbzHovOkoXd0ynmQcN6zRTo/Hv0UU17m250lDeK3lJTBqSGW++qYsXqZeQFBPZVaNVU7Mk6AJenyxblh7yqEIZknSi1ZCMWHPKUN5K/6y/mPVdfIyuwhF+G5KmE/by/aPrC1aonJ4euottP/EF3ZTvPDv6XcP+2OnmRksDVnReEmuGpHhvZ4YEQtlgUYYkipk/d67OjLRQXTSeCIaukdSF81TlyvTQRw3KkKQTrYZkzPozhvJ2QheHW/wyJJ1mHOStN3Q9XqremD10s74xeNWHhu27Ueup79NVek6sGZLHXnFuSKDbmy5jQ986SVdjizIkUcyjRYpolSZm4F2lumk804LM6S0lSDvfs3t3evijAmVI0olWQzJ8tfMWEqG209xXen4YkurDdhuW90Nl+m2jm/aNemPct47IKtprC12l58S7IYFydjLew3YoQxKlYIiqMCPtsifxwExaqSqFpoWZE3TDgtu2bk1PQ8RRhiSdaDUko9c5r1xk5e26ga4qIF4bktL9thmW9VNFe22mu+ALoV6XTd7YS1fpORnBkEDd5rrbd2VIohB5aC+6aeI9/XskNSNLAp90UBzvIYMH09MRUZQhSSdaDYmbykWW2wBALw1Jw3HvGZYLh9B64TdJHdYativr/zVYYvhMFs7LyLWn6Wo9xc01Ew2G5NGXtxjKO1GRl9yZUGVIooiJEyawfLlz84oRleTkOEl2dqxkFsYO3sR1Y8qdbF/27IYykRRGLFVMTDcl6CqLFpQhSSdaDQn6yml5p8Kbp1O8MiRTU87ZVsp+Cr/j51//oLvlCbVGvmPYnizcS93nHmH/Vz/w73+o20a6ak+JNUPyeJAtJBCO9eCVH9JVmqIMSRTxTNWqWqVYNjGZJzzbmyWBfT/vDq1Cd6NLL9yrq3g/aXOfoYwTna3/H0Ml7kbnm9yvrevv929iHz0b/PpOPf0A+37u7YZ9dKozta233T1bonb8O3XoQE9PRIh1Q4LKccz602z3qS9d6cD5r+jmotaQDHcxysZMCAAc5iAA0AtD8sEnN1imNm8ZlgkktBhUGrSDnf3se9Z51iF+HPosPsb2nLnOHnlps6F8ID3QehVbc+AKW33gCt01T+i98Khhm7ISO6xljcfvNXxOhRFHfhJrhsTNKBszVRy4g67SlHAakhlvXzA8d+zkJVFtSEaPGqVVhvmTk9nqtABWGJJvpt3B/th1M/t7n7GCNejATeyv927i5c83vl9X4V5+8R7+OXvfZDkT/fXezezXTbew42UyGypvpzqQJxv7c/fNuvXCGNFyTvVJu/vYr5tvYX/svpmbG7rPZhLH47s5t7P9ydatMyulmJJCBQrQUxQRotmQvL7S2WgGVKT0xrZTLBmSUEd1QE8N3klXayBUQ9JxhrOHvdDtTZaxDtMPsrm7PtYMiYwwJIgT+OGX3wO2Ttz/3Cq2+YPP2MYjn3FD4hfZnl9j2LasdtPeZ/e2sq+wMPx50pazdPWekdEMyf81WMLGbThDV2vAE0Pi4PxCQ946aXju2MlLotaQ9OvbVzMjdROS2KYAo2nefzA7+2XtrYZKF7rc/R5DeSsdK5WZmxe6jr/338SOPRG8AdEpUwK7PuJuwzb+2Hkz259kbQzc6GSVTNw40W38+c7N7IqL4yG0LlMCK5iWaj5vrlxs8sSJ9HSFlXgxJF4QtYbE4XGwEyr3QIRqSEr2TTGUtVKbqe/zuIFTV76jqzEFhqT+2D2s+vDUUTulXk1hRz7+mv8f+42EazAkflK45yZ2U8Olht8iq+aId9iw1fZdbGgVQrCyX8SaIQk2qJUKuV76LT1OV68RbkMSSaLSkNSuUUMzI/MzGytIKlTAf+8zr4BPVHBuJM42+I9hHXw9u2/mpoeWD0YH82Vjv6XcYtgGdLFj8K0ksmBsft9uvo3P+vybHcyfzbCMnTAxX2UppqRHt270tIUNZUjSiVZD4uQ4FHhxg+2DEm/lgR7WoRiSLrMPGcpZCfElk7eco6twBAwJKjakkIchCRdIwEZ/BxVaafovPc4NCVq17m6xwlBGFrK9+kWsGZJHe9sHtaI7DKKfUwXKwaMMSYR5pFAhXukVTE6NGaGVI9W+hOzsx6W3GSpf6MpLzlsEjhTNyltD6DrQ2oDWE1o+GH3c2jpm5dvpd7J92UI3PiefMm8h+WPHLex88/sN5Z0Iyeeey56eOK1qBBOnKUOSTrQaEidBrXc1W87GbfyIZbfpUoAqDnybboITiiHJ2s5ZOvjUESangjYkkaD/shOG32GmLrPTK3YYJicVl19DgGPOkDgYZQMD99KCDwyfUyHIdeQa8wlOlSGJEL/88gufeRYVXqmkZLbRgRmBTpTLbIjJEEbiVLUHDOWtdKbGA4Z1QGhpQcsGLe9W+xKzs29nBA7GPVkpk2E5t3o/h3ULydXe94T0W16Q5r4pWrgwPYVhAYbklsaBm6EhzD+CJn83erjnJnZ388BviRDKmBmSVxYdM5SlQgZQZL7E73CqaSnmlWG0GhKneUhenHPY0TwxmKxuylZj/EIohuSFmc5bSPwO6PQaJ8cUwlu+THsHFfWgFc5Gh7gl1gxJiT7OuvtaTd5v+MxKGHpO8cKQ4BlAy1Ohaw8GlT57rIQWNa+JKkMy5PXXtbTwtROS2E6TCtFMezMnsB8Wmo8y+bz/vw3lrXSkcFb257vmxgamh5Z3I+zj8bKZedDpb1tuYT8uNt/fq684b9GxEkbdmLWQwKScb3a/obwbjc+ayCqkdd1gOHYkwM3gxJD4qVAMSTCq/Lp5RH6sG5IGY/ewzG2djXBBLo03t53XbScUQ1J+wNuGcma6s9lyNjBAJbx87yVWe+Q7ngnHJFSeHrrT8DvMRHOLIEcGLUOVq7P7lkcnxKshuafFSt6qRD83E+JJ5pJ7NVyGxK0eDCLrrB1RY0h69+qlvXk3TnCX+OzDiuZdFDyGxIWR+PR58+6UXzfeaijrVuca/oevB+s7XjozuzHpTsN2oG+m3WlY1q0utLjfNDj3lzW3so/qWg/xdaqpWdKHAk984w16Kn1HGZJ0Yt2QYKQKUsYnd7TvZ4duabSULXjnU207oRgS5HWh5cyEWV0D4bSbzo1CAYnl6PrMVHPEboMhmbD5rKGcmRqOta78giVeDQmEId1OuwifJ9tUhiTMvNSjh1bBNUlwP1neZ33/bah8of9tvJUPEablrXR9xF2GdUDfzbzDUNatRFcNWnIQIHuyaibLIbrHQxzRgxFBMGN0vdCnHe7zJAnb7MwJ2jkbP24cPaW+ogxJOrFuSPBQB70WHHWcth1N4IJQDMkDrZ21zKDVIhDRZkiQ14Suz0z3tVpFF+WGBL+XlqXCcb/8lfH6D4VYMyROh/0iLgOGBF0cdsnnIFyrFQel3+/KkISZrp07a5Vb6+xJhsrPTifKZzbtasFnH1ZwHpPxSVuLFpLNobWQHHkkq2Y+ENSKz97PmY39sto8EPdyt3t5HAhdj1N9/Jz57/j5rdvYmTrOY2oCaalkSMaOGUNPqa8oQ5JOvBgSMGDZCT6qhpYx08M9NvFlMKyVfmcmM0PiNBlaXZvU7l4bErN9dQOuFbpOM3WaeYguynFiaO5otpwVe9nbCffi1ZCIUWIwJE6T5iGQ+tUlx/hyiOug35tJGRIP6P/qq7xSQ/KtVkGYEeiLQca8HtBPy28zlLUSRup8O9084PTaAOdxKFR7s2Rn381KXe+PS25P/y5TAvt2pvn2fk+5hX3oYrgy1dWX7zGsE0LrjJsg30DCxIZN0kbd5H7wQXpafUUZknScGJJ/t1jBpqXoYy/syPa8ffOyH4YE6clpGTNlbvMWX6aUwzwiZpW8E9MFVRu2iy6qw2tDAqMULE5GdCAmBq0gk7eaB0oj0JguYyYkffOSeDckACO18PyiZayEpHWvOOyCU4bEA0aOGMErNcxTs9bhqBoqBIuajbJBC8nJys5bSKxaFhCESss6FbqMftuaGjtyfdjdPLhVfHesdObULLEm2/xhgWReXMqqpeenlbexM7W8MSTQIqmV5OGHHqKn1jeUIUnHyXA+5JZ4Y+NHdNGAOOnvxjEYs94802QwhgRM2HSWtZ7yvqEcFVpSJmz+SEs6ZiczQ+K0JaFQWmuMFV4bEgwnDRYn1wO6W6598wtdVAOBw06GtEIYXuwVGcGQgKXvXWS3NrZ/kYCaTdjHh5zTz82kDIkH5MmZk1doLUOYwffoY1nY79vMh7l6IRibQw9lNWzXiT54NAv7best7H/rbmWnqxvNwKUX72V/7TU3JRfJvDtOdbhwVp75la6Pr7PzvZ5lhIUmSQGuxcI0AR8MidM8JKiI3SjUPCRODAlaLFCRoxvFqbYdv0Y3xXESB4FjhbdeNyB4lK6HCr9j0mbjUFwQrCEBv/7+l6MKEQ/Zs59/76gSRksSxUlrgtBYC+MFvDYkDcYFP8qmTP/thvVRVR2yM6AhAbg+6XJmgnH1ioxkSAYsP8Eaj3c2u7TTrsXmE/fptiPjxJDc0ngZm/n2BcOzx0pv7b9MNxMyETUk7Vq35hVZkaTkgKnh7YShvbTi9VJBd9lkSmDfz7mDx4+ceuYBXeuI0JdD7jZNxoYWn2Dny7n2mnkX1ndz7nDVYuREiPkRhmTk8OH0FPuCG0PillANidPKCW+p9AYPJCtDktjBPqkYWgdaTkoPBHUCXYeZ8JCbv9s8NiVUQ4L+c7t4EpgMzAODvC70OyozQ+LEdAkFqgAxUgXmzE7oKqHrNROu72BBNxpdnyzkc+m/7LjBiFNhjhW7dUE95gW/r5SMYkgEiMGhZc3ktGsxkCFxYtohBODSZ4+V4s6QDBs6lFdk+ZKdJ0EzE0aV/LHTvxYSxHTQbToR4lLEHDuHH7ZoYcmUwL4aZzG6Z1Zwo3s+bW/RZbPsNna6prGVJhRhwkNhSDDPTTiIF0PiBQjupOumQtpzt82rTh7YMDq7TppPrhWKIQEwJIW62/+2lpP3899HP6cyMySYw4WWsxK6vaxwakicjJZAiwNmgg6GJ151NkrJS6GFzisymiGB8cNEjf+0GXmT0H4NKz/AvuWr/hjrljU3hiSSRNSQ5M+Th1dkCGZ1O9RX1gfFUrtFaAWMlgnMT0PLW+nDSpkM64DQZXO4kIWhCCCMrhHdNYcKWi+P2XbNYmAgzORLy9vpSNEslrEpFzvdy40SXSYUTZa6bUqXLElPs+coQ5IOmt/puq3k5s3byTwdOAdffGve9B+qIQGpD2z7igP98XamBGUoMD1ZHMTKQEjKFipOms0xgV8w7PjwC0fr90PIJeMFGc2QAARy2+XgwbWNmYHp51RmWV4FypDY8Hzbtnx0BiqxN7IkGio5NzpWMgv7/W2jIeGp459x3iJwodn9hnVAyHBKy9oJwaz/HX0XT1B2zsYUHciVjbde0O1CV3q6z9wK00HXA2GUj1kcS6jqky3dkHRs356eas9RhiSdjjPs37qF/tPamHfCDLy10WXNdG/Llb4aErBg9yc8eyVdzq3MWkgA0qDTsmZC/ohQZrpFDIqTHBTIlBoMGKpr18Xll0I1B4KMaEjAa8tPcDNBl3MrzCxthTIkARBDfaGu2YIb6isLlTatfCGeGM0kbsNMe7Nmt8yeijgPWt5OZ2r/h8eGwARgfhpkkw0klKPbhTAUma7bTlbDoNEFhFYgWj5Urc+UwPKnnc9wDAFWhiQdBNE6jYVA0zAefoHAMF6nsQ7lX7OeodRLQ4LRA3Q5t7IyJIinoGUD6XGb/TVjzcEr7F8N7c8RKg2YwWBw0qLll+5rtZLuTlBkVEMCJm05a9vKZ6dGqoUkODBEVBgSxCDQCs6trIa5orvkQB7nE8l93s88OPbHRe6H4GKEDF1PsBLJ1JzKarZfdD25XZfQKWSW3Wc9N9DSzImsTkJqgGulChXoKfeUaDYkry6xr+Dw4CncYxPvF3Yjq1lWnxzobE4WCEGNdD4YcPzSNyyXw6ynQvitVnhlSAAMyfiNH4X0wDbrsgEwJMg4atdsLguTkDmdYK7lpH2OWi5wPdM07k5pO9V+mLTfKjfA2pw6ZYBLc+hUBbptpJuKSkPy1GDn3a9mCjTsF62ZtLyZMLqNPnfs5CURMSStW7XSDMmQrKF110BWw34xnPZ0DeddFOebmHfZ/PG2uy4bxIQgCRliWJAy/uuJdzqSWRwM9NXYuxy39ECXupiboR8WBZcYbV+27HzfEJeCuBj6PbQlUwIrnJR6TkeNGEFPuafEuiEJVph7xYx7WtqPhvBDVQbvpLui4aUhAchW2XfJMXaXw9YbKqsWEgBDMn3bedeGJ1/XDTzh3Cdf/si++/l39u1Pv7GDF75mvRce5a0GtHwg5emygf32x1901xyBY+3E9PgpdB2GSkY2JACGBC2YwZ7LQJMyOjUkwchLImJI8qUFsyI7ayija4SiyZBgnhjEaQgDQL8PpLP1/2PYNgRjgzgTWt5KXhuSE0+mpuaHKbEKzkUOGeSSwXkt9fjj9JR7ijIkejATLS3rt5ACGwF5VnhtSMDvf/7FBgb5wA5kSMDLC4+yemPeNSwXDj352tus86xDQRsSJzlbIKQuP3bxG1c6+uk3LG9X+5YzJMkLlYxuSACG7jodDkzVIMCEh8qQWNCuTRtPW0egCy3NjcRvKbewgwWcV+Sf9bHospFTvtvoWInM7M9dt3ATcb7Z/YbvAwkJy35eZR7c+sUg864SM8F0mCVbw0ieC63uN5S305EiWbV1nGtkvTzMJXLK4NyO83F+G2VI9MCQjF4f3rfksv0DN9X6YUhAs4l7eeAnXY+d7AwJQPNzsJVBsEKrjFVyOSfAyNB1Wqn6sN10cUdgNBJdl5nKD3ibLuoKZUhSQSubk0zJVI3Hm3fpAmVILCiYPz+vsEJJFU8VTYbk81dT14GROQfyOt+20KWu5q0bfOiww2yxnhqSTPrZlAPlMcH8No3TWkkqli9PT71nKENiBMNFyzrIVeCFei34gE3ffoHugg6/DAnoOd95hlUhs9TxZsCQwGzR5f0QzAhG34RiSLC8EyN6e9NlPAA6GGBInLTCtJ0W3JBlgTIkqcCQPDNsF7/f6boCKeCwX2VIzBk8aJDWQtI3mzctJKioUWHTChj6+Ln7DeWtdKxUZsPy0F97buYp4Gl5M301PnWkzi9rbmMHgzAk+xOzW84C/PmAfzvKh4IuLMu8Js/fx0cU0WWshC4o/H6xvN2wYYya8nu0jTIk5hTttZlXCghcpct6IcxAiliB9m/axwv4bUiQ6I2uK5CcGhIAQ3L4469db8ONKgx8m83YfiFgSnonOG0tChTw6AQnZgH3ZCg42UYwijVDInhuyn7DugKpqUXQO1CGxIIqlSpphmS5i0DNQOLzxaQYY0j+3nczO1M7cAUq61wj8xgOZIGlZc10pk5q7Ah07TXnXSyy0G2DFO90H8TvwTw1dBmqy93Mh0Hz2X6fdj7s92C+rLq09hiCHCip2tuZEliVRP/zkShDYg0MyeQtZ/lIELp8sMIbeJ1R77Ddp75kI9acops0xU9DArAfLSbtM6zPSm4MCYAh2X78Gpu+/bwneVAg5CFBjAAyb1rNkuwGp5UqhCkGQsHJRIRI4R9sYjegDIkeGJKpKeccB1sHHParDIk5+XLn5hXWwyHOXyPLb0OC1gZalupA7mzsp+XpLRs/LXPezUOF7Ky/bjRv8bnayz5RmpUh+XHpbTw/Ci1vUKbU9POY5Vhe/sbEO21bV9qlddmUfOwxeuo9A5Vdyb4pPLAykGoMd99njpk16XrMVG3oLvb5DWNSMAzfo2W9Ema2dQIMydC3TrLVB67wfCf0AeJUMCKl+21jH3x6g607dJUbEqdgjhu6/2YKNP+GHSv2XeJdLHSdZirRZytd3BEwJLN2fMwOnP+alei71XHlICt3l/Vs5f5LrGD3jTxI1CuKv7LV8DutFGw6ekG3uYfZo72dbW/Y6uByWUzdes6wLi/07Oh36aZ4viBazkwoZ0XTCXsN5c00ecs5uqhjei86yk0RXaeZXll0lC6ugVY/Wt4reUlYDcnoUaN4ZVUiKZm95lFAK/TFEPNEYAyjUxzmIcHQVqtsqT8uvZ2drGLdsoCuGYxCkZfhaevrO6j8iTC8FwnazPKIQDBe+7IZlxMSQ47pchDiWmA06DJ0+W+mmieIw3GGWaHLyBJdNn7Oa6MMiTOEIUFa8SGrTrKHe25i56/9wD+fuDnVtOBNffjqU3z46tZj19jn3/zCcnRayzrNPMROX/1OMyRuiRdDIhCG5Mf//cE2HL7KR+RguO+eM9f5BH9rD15hGw5/xt49c53PQHxnsxV8GCfejg9d+JobEq8JpyEBypDoCYchAcqQ+ESTRo14ZYUEWksyB29IzjU0b8mw0w8Lbmfv50h/w0dODas5XwLp+si7udGxasUwEwJmrYJSz9R6wHTGXzshSdnx0pl5DhP6nR+68pJ960z/rKlp5AvkzcsufvopvQQUEcaJIVE4I5AhUSgU7gmrISldqhSvrNqGOJmeMiSpCrchwRw5dN+ppqVNtIcWkolvvEEvAYVCoVAoTAmrISlcsGBqwKMH89coRacwFQDOcZ6cOVn7du3oJaBQKBQKhSlhNSSopFBZdVOGJG6FifaQYyZXjhzsqUqV6CWgUCgUCoUpYTUkyE0BQ/KSMiRxqw2ZElieNENSqkQJegkoFAqFQmFKWA0JKikYkr4ejrBRii7BkORLMyTFixall4BCoVAoFKaE1ZCIhGgDlCGJW2E+mwJphqTIww/TS8A3/vjjD/b29u1syuTJ9CuFQiHRrEkT9kSJEuyRQoW4KpQrR4sYqFenjvb8hso88QQtElYWL1rEypcpw2rXqMGea9mSvdSjB9u0MT0B2ueff64rH27q162rHSs8C2dMn06LKEyIiCEZn7+QoSJTig9tzpLEiuTJy8/zQ/ny0UvAF3Czi/ikEsWLs4cfekj38AwktOKsXLGC9erZ06Atm4MfY3/16tWw6LPPPqObDonLly+z06dPa7p08SIt4imoOHbv2sXeWrVK084dO2gxx3zy8cfavp85c4b99/p1WsSSa9eu6X47dOXyZVrMEV9//bVhXW7k9XmVQeX9ePHi2j1QrWpVWiT1t1+5ov3do1s33X3zbO3auvLhZv26dYZ7+bX+/fk927B+fW4CRo0YQRcLCwvmzdMGcEAIVRg6eDAt5gtIt0CPi5latWhBF+XUeOYZXTk/80mZERFDMipfQUNFphQf2pQ5iRXKlSssLSS1qlfn15PoCgxWmJm4UIEChs8nBDlsWSQADIfwAPISPKjk9Tdq0IAW8ZTt27axBtLbJFS7Zk1azBHlSpc2HB83b/Ko0OjybZ57jhazBa10dD1u9fJLL9HVegaG48vHyuw31iQV07y5c3nLivi73rPP0kXCCgxJ9aef1t23Pbp3Zw3q1dP+xnNh/NixdFFfkY8RFSr31W+9RRfxlFANSYumTXXlSpcsSYv4SlgNiag4xnfqzL5ctlIpDnVyxmyWP82QoLXCT0SiPQjbw8NpyeLFPG09PsNb3Nw5cwwa0K8fq/Tkk7xM4wYNWOtWrQw3LN5qlCGJLUOCbgh6fNByhuZ9J3hlSCZPmmRYj1v5aUgWLVzIKqZd/9CLXbrovl+7Zo1hf2a8+abu7w7PP69bJhK81LOnrqUH56r/q6/qupfy58nDFi5YQBf1nDmzZ2sz2QthmpTx48ZprbdCjxUrxlK2hpY52IoF8+cbnndmSkkxn0MK14K8r5UrVKBFfCWshkScmDGjRtGvFHHCF9eucXMQjmG/eGiLGweVGJrXYUjg6vEZuozojQjBaMCsiGXfGDeO/4sYFGFmihYuTDfnmPlz5/K3NyvJ+y00b84cVvaJJ7jkz/GQo8vLerZWLbr5kIhlQyLOOxWuxTenTqXFDfhlSFAx0fNmJ1yTfoFusSoVK2r71+eVV3Tf10xreYRwH+A+jkZD8ua0aax82bLaPj3z9NP8cxiSd3bvZsWKFOExMqF0vTqhT+/evPVDPj54UVi6eDE3JGtWrzbtRi5TqhTbsmkTXV1EwbUg7yNaysJJWA1J4UKF+I98feBA+pUiTrh86RI/x+FIjDbwtde0G6dq5cra56Uef9xw8wcSHihg3dq12hsXsgr7hZUhEcifw5CEk1g2JIHOe7++fWlxA14ZElT4dD12xpIqkoYEZkN817FDB97i17NHD+0zmPaxY8ZwYy8LXVUH3g9utl+0MNAYLjvhWoVhwj7BdDasV0/3feeOHdnzbdsalsN9Hirjxo417ebFfnTt0oVtT0nhhgTAkOzbu5fHtIjJZanQIos4nXPnQpv3JlQylCERqePRrPb333/TrxVxAAIUcY7DkTr+1T59+LbwsF+7di0PxEMgI4LKRM4bJxLN42hdQXMqPnv6qafo5jxDGZJ0vDQkVi0kQp06dKCL6PDbkEQLgQzJIMnkuxVaAYLt5mzXpo1hfX4p2BfiD44c4bEzocasQVgH7cqRlT9vXtdxOgjmpesJJJSnZChD0qRxY/4jWzRrxo4fO0a/VsQBGPGCc4wmS79HaODBQm8ydHl07dxZ2wfRDSILD00RBY/WEfQz0/V43RUiowxJOl4aEtpC0v3FF3VvsXVszmk8GxI52NNKGLU1Yfx47e9HixRhTRo25BIB5Py3FCjAz5n4Tqhls2b8twdDNBuS1wcNMqxDCC9egwcN0p45brRv3z5WvVo105aWRx5+mC1bupTuSkCUIXGJCPYrV6YM72dXxB8jhw/n5xg3qt+GBM3G9CajhsQMJ4YEpjlU0PRO12ul4UOHasvR78zUuVMn3ba8Ip4MCUZPYYir+BvN5XUDDFeNZ0PipLJCfBWCGMXf/fv108yGWQyJl4YEo2ZoF5ATvdCxo7ZPo0eONHxvJsSXuAFB0dQ0IF7l0MGDbNrUqXzoOuIiB7/+uishZg2GZOaMGezUqVOsfdu2vOWkVfPm7OeffnJtSLA/uJ+EFs6fr+1v0Uce4edG/h7lKRnOkIg+NPT9/fe//6VFFDEORrDg/KKf2W/kGwcGCG94yCUhB7UicEwuB6EVBw8E8TfemHCDoolUDJuza953gjIk9nhpSGhAMETfbnGdLF+2jC7KiWdDcvjQIa0iwrUt58kYNmQI27pliy6YFerZvbu2vJkhiQYmT5yo7VPzJk3o154AQwINGjiQmwSnQ2vdConezp09y/MiwZC4hRoSBPPK66/21FO67xE6gTg/ocYNG+rKo/ta/l7Iz674sBoSUKVSJe0H31pzAvtHvcVKcaL/V3c+y1G0Cj+3XlTodsgP1VUrVmifU0OCBwjelJEXBZ+bGRKACk1Ey7/cq5e0peBQhsQevw0JzjWCLeV+f/x/2pQpdPEMY0jkShxCgOaePXt03Q7ohogFQ1JJatF5uGBB+rWnhMOQhArynCBo99ChQ+z8+fNsunTepkyapLX0lC1dmnXr2tWwD07UsnlzulnPCLshwYUuflhStZcNlZpSbCtzufb83CKo1G/kCkhOzUwNCYzLxg0b2BNpn1sZElQaouJy289sBk23DX3//feuY0iwDCR3SSAi3w/i0ZAsXbLEMCwTLScUrwwJHfaLN035zdSJ8KbsF3hZkPcP9ypSncOQYAqGmdOn89FyMtFqSERguxBiYPxGzjeC3DforkL3IJ43SMOPZwuuG3m/RBeN+BtDqTdt2KAzN05S+NuB2DexPuRakg0JUutjGDT+j6BmtJDI++hUcWVIxMnM9eCD7KaaUwwVmlLs6p91F7JMRWvy81uxfHl66j0H26A3C+a1EE2PuNlRiSMJFJJB4TM8SNGPLFpX+vbpo2tpEfIi7bSZIQF4G6W5UU6fOqUtJ6dRR5IqgTIkgQlkSMy+Q2ItGb8MSTDyK9X47FmzDKNEYNbwVn39yy8N+xGsgmHzpk2G+8JO8+fN07VIheO5Qw3Jjh07tJYHVPj79+0zGBIgX4MIEl65fLnnhkTOsItAZtmQoEtHTg0/a+ZM3bNm9syZpjEkchkIzy+/CLshkaOps9cYYKjUlGJXN9WZyfIUfJSfW7wx+E2tGjV0Nw/klSGZatKk7xYrQxIsypAExsx0iJYzdInRLjQa4JoRDIn8Bi0LI2q2paQYPg9WweAk8JYK97h8n40YNoyu1nOoIUELiXjmYCjvqJEjDYYE3WJoFcH/YQrmzJrF1+W1IUGciFgfMlDLhgRpEXBvib/xTKSYGZJwEnZDgpER+dJGNeTOmYv9q9abhopNKRa1iGV6NLViQcUZDkMiHgLFHnmEV2wIav3qq6+0VPC42Z0GtdLJunLnzEk35xplSOwJlyEBY0eP1sWwQWjWFvhlSPD2TJNz2WnD+vV0tZ5AK0qq5k2bGvbFbHg9hC4KdIXS8lAwBGtIEBsh/kZXhN9EcwyJfA8gDbwyJA6QExjdXWeiSeWmFGv6V915LGexyvyc4o0hHCCwU1xHeJAKQ0JjSOQbDDIzJCuWLdOV8SIIURkSe8JpSH7++WfWNC0XkpDI0gu8MiRWQa200g4kPwzJkkWLAibjgswmpHtSSs9O9ZzFJG3B0PullwwZazEflrw9GA75e7T4yPUJngP/+9//6Ko9JZoNiXy8evfqpQyJE9Lz+udgN9dQhiQe9O/6b7Kk4qnDBsM1IZNcgVStUkX73M6QoMuGzmXzhpQQCvKi+ZR2EQSSGmWTvk0/DcnAAQN4k7pcpka1anweJL8NSaQxmw0ZwvF+pXdv/vuRk0fkiMJxoWXN5FWrhJkhkbsghKghwTlCsjZ8B8NVvGhRumpPiVZDQlvmdu3cqQugV4bEAjm6N7FCe0PlphRb+letaSxv3tQHAiL2kSHTTzBfBoQgOJHqHTlExOev9u3LP0POm507dmifC+EzeVbLZUuW8M/lYXCN6tfXLRMMypDYE05DIkBgJ51aYMjrr8e1IUGri9gXBLGK0RYQRhxdOH+e/36z8lToIi2SNoeMkF/Xo5khMWPY0KG6MsFeQ04JdHyCkZO5luxAPJxYH559NA+JMiQWwJBg0iPxUMhaZ7ihklOKHSVWTb85O7ZvT0+358g3TLgUDGZdNniQIaMk/dxq2C/e9kQmTNVlExg3hmTF8uWGocD4nCaHQhePW6LNkOAFQexHoYceYhMnTNBlZKWBoG1bt9btO1q05TgS5P7AJHB03hsvKlWKU0OCzMpyGb+HJUejIXmyXDltfThHypC44NTJk1rK7hZtO7NdJ6+x3ae+VIpBvTF3tXYBo7vEb+QbJlwKBgxLpOmigVtDIsD4f9FUjVTVfpARDAnASxEdNo4ZoyH5s6ZBZP+MRkMipu0Q89OIJIHQpIkTtbJTJ0/mXZ10/83ykODYyJ8hfsHr69KpIcEQVnR5yOX8bKmlgb4YYkszmgYS7UrGMyFUkKpArA+TgypD4hI0U+NHo3np+vXr9GtFDHDp0iXt4sV5RHIlv6EBgFSrVq7kY+7p52b65JNP2MgRIwyfUwWD14YkHGQUQwKaNGpkKE8VTBKoaDMkAIbkwoULXGaGpGe3bqYBrzgXaAkxMyRIN7906VJtOhAhBMZ6lamZmkbMQ4P7lRpHtAgsXrhQ1xXXWLp2kTujWJEiruexsQLPBLEdmDzEayCo3qkwOkk2fl60kMgjyJC1VhkSl7RLaxrEQ1fNaxObYJSCuHi9SCYWKmLIL4aWL1q0iH6tA2mW0UqHpns6ssBPlCFJJ5KGZMnixbqRWmbCqA23RKMhAYEMCdIx0Lga5PQR58LMkIA5s2ez4cOG6ZKtoXJF3g03HDl8mI8IwbLIL0T3xU6NGzXiBokGnKKlAMiBzDAzH544QXfBFbIh8UKhGhI8v+T14blMDQkmAURAv3jJwmzDlAxvSBChjR+O5rYbN27QIoooB6MWxMUbjvgRKzBXDW1qRmWDlOtmvLdnDw+ElcujQkOkut9EsyFBiwDdN7dycwwjaUiAaKW1Eiplt2CeHLqeYPTKyy/TVXuCmSHBfYy5WiqUL88DXtHKKP92K0MCZs+cySZNmMCvVbSMBFO5YlQb/f1OBOOCrKcYJg1DglE3CLoV3+OFAwYRyRHl5dCaEAqyIcHxQvwRTf8fSMidIs8yHswxkxHB/RDML0ZKUUOC7q+UlBRDtluhMaNHSyNgk/moLFpGlh9E1JAAEaiH6Zz9GHuv8A80neLcoYUBD/ZIGBKMmKFTg6PpedaMGbbZLhFYjb5g+jaGt6x5aUMfg6HLCy/o1hcOIVDRCzKiIUGrQW2TrL8QgiXdEouGBPz22288hwdSn8OQyAQyJGDve+/xiv+HH37Qfe4UjHKiv58KdQVGAuHNvlnjxmz50qXs/LlzmiEBH330kS6IF8KLCu5p+T7HHDihYBZDQk1HIHkdQ3LlyhU+ygYtVTiWwMyQeHF/C/lBxA0JgsvkmwNTLyuin6+//lpLwoNmUXnYajhAYBidkwOCOYIRwQPSCejLxVh9akqgnj160OKOiGVDIrq8QlEsGRIAQ7J82TLTawAjTtwiJ6MKReE2JMAqqZidIQkVxDOgFQP3b+lSpXilun//ft6KjjoBRkM2JIGAIVm/dq2h+0aoerVqIb/hR1uXjTAkMJPHjx/nnylDEiSFpTHxZoE2iujjiy++YMXTmgmfqlyZfu076Asd0K+fdt2geRFNtpjR0q5lhAJDgpk4Dxw4wNqmzbWEygkP4WDYvWuXoXnTbx07epTuRlBMmTyZlStThjdDBysnJkCAYMBWzZvrlg9mqC2oVrWqYV9Eki870O1Qp2ZNbRQT+s6bNWkSVKstDAnSoKPSp/vjRri+/SCQIbHCb0MCENfx5Rdf8BmHhSEJljWrV/Pngfw7hYLJLUOJNkNihpkhWbxokSHQPlj5QVQYEvT/4wYUBy6xeh9Drgul6NFNdWawfA+lJkbCRT5+3Dh6SsMCWgUQmLrn3XfZB0eO8AdQKMCQoEKGqYEyoiFRxD+XLl5kZ8+e5aMuICcDCnCNyWbJzyG1XnHm9GmeLwWZmBHkiW4WPCu8MCTffvutbtTMf69fN5iMQEJmafl4TnzjDbqJkPns6lXdLL146Yp2osKQAOTdlwNzEit3NlSEStGhm2pPY3lypwaQ1qlVi57KsCIMiZcIQ6JQxCPBGJJYxcyQeIGZIYk2lCEJEcQFiLHw2Z5oZqgIlSKrf9ZbxLIWr81y5Ujta0cQa6QNiUKhUCjig6gyJJj8CuP+eStJjhzs9prjDZWiUuR0U715LEeR9OQ78twXCoVCoVCEQlQZEgBDMnjQIK3Sy1JjoKFiVIqMEiqnTz6HxGORMiSnT51iS5cs4f3aSG5lhhxfEezkeAp/wIgJcW5wHhFz4zeXL1/WuijOnz9Pv3bFR2fO6PY/I4LROJgPCMOiEQ+B0UjrgwgAjndE1xi6jtDNE02sXbNG95wMNVmcF0SdIQHICSCSXOXJ+xD7V+3phspRKXxK7apJTyJWtnRp3yZ3cwJuHiRhwr4ULVyYfs0R+wqh71gRPVQiSap6dO9Oi3iOPJwZw0dDASMBxbrCnbgu0iABmlmKeVmYoFCRCp5POCaPP/ooj+MIlcOHDvGcXRgJFip0nqBgg/i9JCoNCZCzaOYoWoXdVG++oaJUCo9urTuL5XmkND8XGNuPoZqRRBmS2AXDRWkFxh+GLoYKB4MyJKEzeeJEXe4fGBNkRsX0AtSkYOSkwjtDQoddw0yEijIkLoAhEScTyla8Nn9Tp5Wlkv+6t2Z6qnO4c2VIFMFSVErrLQuGwU+8NCRffvkl7waEjhw5Qr+OW9A1I58zpGjHhHYwJMj3IqdCx3cKZUjcErWGBGAUhzhY+R5+nN367AxDZankr7KXTs/sN2zoUG5IIo2fhgQTUE2bOpW90LEj69a1K48ROOoix8epkye19NBHP/hA9x1iD8R3aHq1A/PtiPK7d+/mM2KjJQHzhXTt0oWnt3drDk+dOsUnGsOINgyFnD5tGjv54Yc8sy22s+Ptt+kinoI03+K84HfIb9ahpvMOhJeGBFkx5fNiB64DUR7XF0DlhLmWkBALMVFuufjpp3zuEbQ4Ids1plDwG9z74hhi2oWff/qJGxLw/v79vOVIZEdFS0q7Nm3oKiw5efIk/z24LhGPMqB//5DzCkUDXhgSTDw4ZtQo7djjnunfrx+/ni5evEiLm4IYlmFDhvBzgqRuGAKsDIlLYEjkiZJyFHtKdd2ESbfUnc0eerqjduyRbtmrKcVDxS9DMm7MGMt0004TQSHrq1iGJmBCELD4Dtky7XiiRAleFg8grJdOHgjhwY+mdCcgZwudKh5CyvW6tWvz/+P3+0XP7t21bWK6gdHSQxZ6c9o0uohneGlI3HbZIPMsymKa+l4W57FUiRI8s6YTcB3JrRFCuBb6vvIKLe4JSNwlbwup9mFIZGBwcZznzp7NTZtTQwLzYTYNBITsubGMF4YEgcP0uAg56epECxZdju9T8eK63F8YTBJpotqQAGTjPHTwYHqQa86c7ERarn6Ff8B9yxevm/lJ/AaGxKxitRIelHZMnTJFeyjiYf/G+PG8taNunTraejB3jx1eGpLSJUvqfgcq0vXr1vH5KipXqKB9jv22G+0hT16GqRo2b9rEW2xQuWL+EPEd7jM/QItQjWee4dtAWn6M0JCPFd92/vzsfZ9GREXSkNAKAd3ReIaNHTNGdx0Peu01uqgB2cTBPKI1D+uqKbU8PVm2LF0sZLR0DGnTNMCQeAFaWEQrWYF8+fg9smDePNa6ZUttbiGkfw+2Mo80XhgSPIueTXthgHC/o7UKrUmYvTcQOJbiOOJfTIiKexH3nzyzL+T0xc1Pot6QABgS+abGW8bff/9Niyk8Yt/evbqJxlBxxLshkd8U0DWFrhAYEgxRFTNz4qFgh5eGpHyZMrrfgXsAhmTq5MmGlhy7N6XCBQtqZbEfsiHBXDLiO0xd7geowOSWARxjzBOD3yRm/Ia86Bs3I5KGBNeBKI9WEoxUgYmYP2+ezlg6mR8G50eUnzB+PDckeHmAmRGfwzx4DUZ1iPXDWHqFfO5xjoQh2bRhg84oQ8i4Gmt4ZUjMYkicGBJ54sr6detqhmTs6NG69UHzQpxw0AtiwpAAPLjE7LIQbmTkFlB4Bypg+SGANzn056KSjSZkQ/Jily5afglZbVq31n6HnSHBZHyiLFpH6CRSz0lva4jbCISfhkTumqEBhoH6f/EAQkyMKAvzJYOJyMR3fhmSMtJMvJg6HoYE/Pjjj3wSO/EdjrUfRIsh6dS+ve47+bfbBfZieg1R9omSJXXXaIf27bVuHD/OoVyxoeXOCxDTINZZ4rHH2NrVq7XvYEiGSPclhGd+rOGFIQFmhsQJ8ovbgvnztc/f3b2bdxPK6wz0DAkXMWNIAAwJKgS5vxFTdKvWktD5/fffWZdOnbTjihl80ecfjdAYEmpGIPlGs2uKRL+7XD6QMMtwIGRDgorX6rtgDIn8BuMmIA2tPQhCFGURmyIjGxI/umxEDIVT4Z72mmgxJNTcy/tlZ0i6vPCC4VhZadbMmXTxkBgxfLhu/cuWLqVFeMtPjxdf5LFmToKj5dm6MVMzBS2B8jZhWmKNSBoSnAN5GZg8GTzj5e8DPUPCRUwZEoCmaRw42ZQgWh1TViuC49133tG1jJQvW5ZfrNEKNSRmyDeanSFZsWyZrjwCqWE8zGQX4/CK9BaLt0qZ59u00b4LxpAsXLBA+86NIblx4wZrUK+eVhYVh0y/Pn207/x4u0YSPXlf7YTKz2viwZDAZMjHCbEi9PoUOnz4MF08JNB6IV+P6C7AVB8ycosHhOdyIBD3JMrieY5uRBn52EAwtrGGX4YELaR2wJCUfOwxbRm0OAnwzJe75SHVZRMkc2bN4g/n4lK/P7oXvvvuO1pUYQOOpXxRPlO1Kjck0YwfhkQOXpXfPvCQRVAdHpj16tTRLWfGpIkTtfXA5ImRE1MmTdLtUwsH/fxeGpJuL76oKy8y7aISwO8Tn3vdQjJzxgzddvu9+qqhSwySfw8q8N9++42uKiTiwZCgVULuqsM1K8BoFHyG67Rl8+a65bwCwazyucRvF5Xs6JEjdUO4sT+LFy2iq9ABQ4KsrmK/5akoUGGKexwqVqSIwbBQpkyerF1PKVu30q8NbNywQSs/PcD9Ewp+GRLaymkFhviKZeSu67dWrjTEoWEkVSCQal4crzmzZ9OvPSEmDQnAwxldOHJQGHKVPFi9l2EIq5KJ6i5kSU/om9JbtmjBDUm047UhAQ3r19ctg77X/OSGhUGwAy14GE4nLyeEh65o2XMynNErQwJmz5ypiyORhQeT6Gv2etgvKhJ5W1YBl+g2Eq10MEhOWpDcQN+2A6lCuXJ0cR2RMiRAjFQS4tcpGQLspDk/GEaNGOG4e/PlXr3o4qY4aT3DCCKYFztDgvMmlnFyz8u/xa9jJgyJEwUaso3fbzY0GvlJAoHcPnQZWfI6h6fFdVkhnytk6PWDmDUkAlQkcgI1rkersJvrzzNWwkpct9WdxXIXTn/bwQgMBLTCkMQCfhgSJGFCHE3njh0NNz4e+uvWrNEZgkCg0qHrqFyxIvv4wgVttIv8dmuFl4YEPNeiBT8WouLHPuKN+7OrV7X1Yf+8RCTTwtszWiZGDBtGi3BgSBC0K/8mLwNc48WQ4C0V+T8QbE73GwZ6W0pKSG/idqDSQlcA7jt6jUN1atViq1ascGVIMNpIHgUmhOu098sv8yReToh3Q4IEjcgHJS+Dui8QMCQ9e/QwdM+gtQvJEOX12XWVKkPiECTWQf+qfIPkyZOP3Vt9uKEyzsi6ud48lrN4VZYrZ3q8CCpKJKSCIYkV/JrtF4bkgyNHeBI0xNDANGAEzrZt27ghcQoqnXd272aNGzRgtWvW5MHBGGJ74fx5bZ+2OMhCuWrlSt3vOHPmjPad25k6Dxw4wM8x+v3xIIK5wb/Xrl3j/xfXg11l7BZUNtg/VOJohg8Emvjl34RcCV6BcyivO5DsKnS3s/1u375dK4/suzLyfuH/ToAhQSxT186dWZVKlXisErocPv30U25I/AaG5JsbN9isGTP4PVKpQgU+ZBm/TRgSN+AawW9/vl07HtyKSrJvnz7s3Nmz3JA4BedNHEsn9zyu/507d/LrHpW2H+D5RK8vK2F/AgFDgoy4MK5VK1fm598upg3gtyF1xgsdOrAqFSvy6wX3GtI7yNu3e4bg/hVl0d3lB3FhSAAMydNVqvAHBBLsiAds3nwF2d21xhgq54ymB2oMYXly6Zt2kZkPhkQR/8jDoNFSIaeux9ukeINympFWoYh1YPDxMobWO6+GMitCI24MiQCGBEm85Mx2osXkvpojDRV1vCt7zYEs14P6mTgRS4Am8mhIFawIDwg6RCCtfB2YyY8htwpFNII3feTiSNmyxTYAVxEe4s6QCAYNHKhLp8yVIwdLrvwC+0eczxr8f/UWsrxVO7G8D+kDCpGLAtHeCHBUZDzatGrF084jTwRiH0SQLYJwEY1/9uxZZUgUGYbvv/+eC4ZEER3ErSEBMCSYQRKR/TSoJ88jZdgddacbKvNY1111JrK8BfWjPFDpIMvjUot4C0XGwc6QKBQKRaSIa0Mi+Omnn3iwFcZZ08jwXA8+yKo27cYGLD3KBq44EZPqMXU7K1erFW8B0pmunDnZyhUr+Fh/GBKFQqFQKKKVDGFIBDAkSL5DTQmEyOVvv/2WLhL1oPuFJriBEKSFnBgwJAqFQqFQRDsZypAAGBJkrMNwMjNjgs8wK+LVK1foolEBUuSPGzvWkLQLQrcUZoZEPgdhSBQKhUKhiAUynCERwJAcPHiQjxO3msoeiXkwoVWkU9IjTS/mV6H7JwwUku8geRB+CwyJQqFQKBSxRoY1JAJhSJBkBv9/JG3eElrxyyalSaNGbM+773o+y/Cff/7JtmzaxGrXqKGbX8RMwoQgcZf4v0KhUCgUsUqGNyQywpAgAyFmPUXLBP6mZsBKMDKI58Ckf8heiAn/kK4b6X0Ro9K8SRPeHfRkuXI8VXIg40OFAFWkoN6wbh0rUrgwW7t2rTIhCoVCoYgblCEJAAwJUm4/VakSn8AKQyZr1aihm9XSF+XIwfLlycNbYpAmvXChQmzk8OGaIVEoFAqFIt5QhsQFwpCkpKSwpo0bc6MAc3L82DGe9RQtIggmRetH3ty5TVtAEHgKs/FYsWJ84jFMC4/gU0y8hm4aTJqF+QpkQ6JQKBQKRbyjDIkHeGVIFAqFQqHIqChDolAoFAqFIuIoQ6JQKBQKhSLiKEOiUCgUCoUi4ihDolAoFAqFIuIoQ6JQKBQKhSLiKEOiUCgUCoUi4ihDolAoFAqFIuIoQ6JQKBQKhSLi/H//b/etu8Y8nAAAAABJRU5ErkJggg==" alt="HEUNG A LINE — Heung A Line Co., Ltd.">
    </div>
  </div>
</div>
<div class="wrap">

<div class="board" id="summary"></div>

<div class="controls">
  <div class="seg" id="viewSeg">
    <button data-v="desktop" class="active">🖥️ Computer</button>
    <button data-v="mobile">📱 Mobile</button>
  </div>
  <div class="seg" id="polSeg">
    <button data-v="ALL" class="active">All</button>
    <button data-v="THBKK">THBKK</button>
    <button data-v="THLCH">THLCH</button>
  </div>
  <input type="text" id="search" placeholder="Search vessel, voyage, service, wharf...">
  <select id="serviceFilter"><option value="ALL">All services</option></select>
  <select id="timeFilter">
    <option value="ALL">All time</option>
    <option value="7">Next 7 days</option>
    <option value="14">Next 14 days</option>
    <option value="30">Next 30 days</option>
  </select>
  <button class="clearbtn" id="clearBtn">Clear filters</button>
</div>

<div id="tableHolder"></div>
<div class="count" id="rowCount"></div>
</div>

<script>
const RAW = __RAW_JSON__;
const WHARF_MAP = __WHARF_JSON__;

function parseDT(s){ return new Date(s.replace(' ','T')+':00'); }
const NOW = new Date();

(function liveClock(){
  const el = document.getElementById('liveClock');
  if(!el) return;
  function tick(){
    const now = new Date();
    el.textContent = now.toLocaleDateString('en-US',{weekday:'long', year:'numeric', month:'long', day:'numeric'});
  }
  tick();
  setInterval(tick, 60000);
})();

const data = RAW.map(r=>({
  ...r,
  etaDT: parseDT(r.eta),
  etdDT: parseDT(r.etd),
  cutoffDryDT: parseDT(r.cutoff_dry),
  cutoffReeferDT: parseDT(r.cutoff_reefer),
  opengateDT: parseDT(r.opengate),
}));

const state = { pol:'ALL', q:'', service:'ALL', time:'ALL', sortKey:'eta', sortDir:1, view: (window.innerWidth <= 760 ? 'mobile' : 'desktop') };

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtDT(dt){
  const d = String(dt.getDate()).padStart(2,'0') + ' ' + MONTHS[dt.getMonth()];
  const t = dt.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',hour12:false});
  return {d,t};
}

function hoursUntil(dt){ return (dt - NOW)/36e5; }

function flagFor(dt, urgentHours, pastLabel, pastClass){
  const h = hoursUntil(dt);
  if(h < 0) return `<span class="flag ${pastClass || 'flag-past'}">${pastLabel || 'Closed'}</span>`;
  if(h <= urgentHours) return '<span class="flag flag-soon">Due soon</span>';
  return '';
}

// populate service filter
const services = [...new Set(data.map(r=>r.service))].sort();
const svcSel = document.getElementById('serviceFilter');
services.forEach(s=>{
  const o = document.createElement('option');
  o.value = s; o.textContent = s;
  svcSel.appendChild(o);
});

document.querySelectorAll('#viewSeg button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#viewSeg button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    state.view = btn.dataset.v;
    render();
  });
});
document.querySelectorAll('#polSeg button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#polSeg button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    state.pol = btn.dataset.v;
    render();
  });
});
document.getElementById('search').addEventListener('input', e=>{ state.q = e.target.value.trim().toLowerCase(); render(); });
svcSel.addEventListener('change', e=>{ state.service = e.target.value; render(); });
document.getElementById('timeFilter').addEventListener('change', e=>{ state.time = e.target.value; render(); });
document.getElementById('clearBtn').addEventListener('click', ()=>{
  state.pol='ALL'; state.q=''; state.service='ALL'; state.time='ALL';
  document.getElementById('search').value='';
  svcSel.value='ALL';
  document.getElementById('timeFilter').value='ALL';
  document.querySelectorAll('#polSeg button').forEach(b=>b.classList.remove('active'));
  document.querySelector('#polSeg button[data-v="ALL"]').classList.add('active');
  render();
});

const COLS = [
  {key:'pol', label:'POL'},
  {key:'service', label:'Service'},
  {key:'vessel_code', label:'Vessel'},
  {key:'vessel', label:'Vessel Name / Voyage'},
  {key:'wharf', label:'Wharf'},
  {key:'pod', label:'Next Port'},
  {key:'opengate', label:'1st Return'},
  {key:'cutoff_dry', label:'Cut off (Dry)'},
  {key:'cutoff_reefer', label:'Cut off (Reefer)'},
  {key:'eta', label:'ETA'},
  {key:'etd', label:'ETD'},
];

function sortData(rows){
  const {sortKey, sortDir} = state;
  const dtKeyMap = {eta:'etaDT', etd:'etdDT', cutoff_dry:'cutoffDryDT', cutoff_reefer:'cutoffReeferDT', opengate:'opengateDT'};
  return rows.slice().sort((a,b)=>{
    let av, bv;
    if(dtKeyMap[sortKey]){ av=a[dtKeyMap[sortKey]]; bv=b[dtKeyMap[sortKey]]; return sortDir*(av-bv); }
    av = (a[sortKey]||'').toString().toLowerCase();
    bv = (b[sortKey]||'').toString().toLowerCase();
    if(av<bv) return -1*sortDir;
    if(av>bv) return 1*sortDir;
    return 0;
  });
}

function filterData(){
  let rows = data;
  if(state.pol !== 'ALL') rows = rows.filter(r=>r.pol === state.pol);
  if(state.service !== 'ALL') rows = rows.filter(r=>r.service === state.service);
  if(state.time !== 'ALL'){
    const days = parseInt(state.time,10);
    const limit = new Date(NOW.getTime() + days*24*3600*1000);
    rows = rows.filter(r=> r.etaDT <= limit && r.etaDT >= new Date(NOW.getTime()-24*3600*1000));
  }
  if(state.q){
    const q = state.q;
    rows = rows.filter(r=>
      r.vessel.toLowerCase().includes(q) ||
      (r.vessel_code||'').toLowerCase().includes(q) ||
      r.vyg_bound.toLowerCase().includes(q) ||
      r.service.toLowerCase().includes(q) ||
      r.wharf.toLowerCase().includes(q) ||
      (WHARF_MAP[r.wharf]||'').toLowerCase().includes(q) ||
      r.pod.toLowerCase().includes(q) ||
      r.pol.toLowerCase().includes(q)
    );
  }
  return rows;
}

function render(){
  document.querySelectorAll('#viewSeg button').forEach(b=>{
    b.classList.toggle('active', b.dataset.v === state.view);
  });
  const filtered = sortData(filterData());
  renderSummary(filtered);
  if(state.view === 'mobile'){
    renderCards(filtered);
  } else {
    renderTable(filtered);
  }
  document.getElementById('rowCount').textContent = `Showing ${filtered.length} of ${data.length} voyages`;
}

function rowAccent(r){
  const dryPast = hoursUntil(r.cutoffDryDT) < 0;
  const reeferPast = hoursUntil(r.cutoffReeferDT) < 0;
  const dryDue = hoursUntil(r.cutoffDryDT) >= 0 && hoursUntil(r.cutoffDryDT) <= 24;
  const reeferDue = hoursUntil(r.cutoffReeferDT) >= 0 && hoursUntil(r.cutoffReeferDT) <= 6;
  if(dryPast || reeferPast) return 'var(--coral)';
  if(dryDue || reeferDue) return 'var(--amber-strong)';
  const gateDue = hoursUntil(r.opengateDT) >= 0 && hoursUntil(r.opengateDT) <= 24;
  if(gateDue) return 'var(--amber-strong)';
  return 'transparent';
}

function renderSummary(rows){
  const upcoming7 = rows.filter(r=> hoursUntil(r.etaDT) >=0 && hoursUntil(r.etaDT) <= 24*7).length;
  const nextCutoff = rows.filter(r=>hoursUntil(r.cutoffDryDT)>=0).sort((a,b)=>a.cutoffDryDT-b.cutoffDryDT)[0];
  const nextGate = rows.filter(r=>hoursUntil(r.opengateDT)>=0).sort((a,b)=>a.opengateDT-b.opengateDT)[0];
  const bkkCount = rows.filter(r=>r.pol==='THBKK').length;
  const lchCount = rows.filter(r=>r.pol==='THLCH').length;

  const cards = [
    {lbl:'Voyages shown', val: rows.length, small:false},
    {lbl:'THBKK / THLCH', val: `${bkkCount} / ${lchCount}`, small:false},
    {lbl:'Arriving next 7 days', val: upcoming7, small:false},
    {lbl:'Nearest cut off (dry)', val: nextCutoff ? `${nextCutoff.vessel} — ${fmtDT(nextCutoff.cutoffDryDT).d} ${fmtDT(nextCutoff.cutoffDryDT).t}` : '—', small:true},
    {lbl:'Nearest 1st Return', val: nextGate ? `${nextGate.vessel} — ${fmtDT(nextGate.opengateDT).d}` : '—', small:true},
  ];
  document.getElementById('summary').innerHTML = cards.map(c=>`
    <div class="seg-stat"><div class="lbl">${c.lbl}</div><div class="val ${c.small?'small':''}">${c.val}</div></div>
  `).join('');
}

function renderTable(rows){
  const holder = document.getElementById('tableHolder');
  if(rows.length===0){
    holder.innerHTML = '<div class="empty">No results match the current filters</div>';
    return;
  }
  let thead = '<thead><tr>' + COLS.map(c=>{
    const sorted = state.sortKey===c.key;
    const arrow = sorted ? (state.sortDir===1?' ▲':' ▼') : '';
    const wharfCls = c.key==='wharf' ? ' wharfcell' : '';
    const polCls = c.key==='pol' ? ' polcell' : '';
    return `<th data-key="${c.key}" class="${sorted?'sorted':''}${wharfCls}${polCls}">${c.label}${arrow}</th>`;
  }).join('') + '</tr></thead>';

  let tbody = '<tbody>' + rows.map(r=>{
    const eta = fmtDT(r.etaDT), etd = fmtDT(r.etdDT);
    const cd = fmtDT(r.cutoffDryDT), cr = fmtDT(r.cutoffReeferDT), og = fmtDT(r.opengateDT);
    const polClass = r.pol==='THBKK' ? 'pol-bkk' : 'pol-lch';
    const noLoad = r.no_l === 'Y' ? '<span class="noload">No load</span>' : '';
    return `<tr style="--rowaccent:${rowAccent(r)}">
      <td class="polcell"><span class="pol-badge ${polClass}">${r.pol}</span></td>
      <td>${r.service}</td>
      <td class="mono" style="font-weight:600">${r.vessel_code}</td>
      <td><div class="vessel">${r.vessel}</div><div class="sub">${r.vyg_bound} · ${r.op_liner}</div></td>
      <td class="wharfcell"><div class="vessel" style="font-weight:500">${WHARF_MAP[r.wharf]||r.wharf}</div><div class="sub">${r.wharf}</div></td>
      <td>${r.pod}</td>
      <td><div class="dtcell"><span class="d">${og.d}${flagFor(r.opengateDT,24,'Open','flag-open')}</span></div></td>
      <td><div class="dtcell"><span class="d">${cd.d}${flagFor(r.cutoffDryDT,24,'Closed')}</span><span class="t">${cd.t}</span></div></td>
      <td><div class="dtcell"><span class="d">${cr.d}${flagFor(r.cutoffReeferDT,6,'Closed')}${noLoad}</span><span class="t">${cr.t}</span></div></td>
      <td><div class="dtcell dtcell-lg"><span class="d">${eta.d}</span><span class="t">${eta.t}</span></div></td>
      <td><div class="dtcell dtcell-lg"><span class="d">${etd.d}</span><span class="t">${etd.t}</span></div></td>
    </tr>`;
  }).join('') + '</tbody>';

  holder.innerHTML = `<table>${thead}${tbody}</table>`;

  holder.querySelectorAll('th').forEach(th=>{
    th.addEventListener('click', ()=>{
      const key = th.dataset.key;
      if(state.sortKey === key){ state.sortDir *= -1; }
      else { state.sortKey = key; state.sortDir = 1; }
      render();
    });
  });
}

function renderCards(rows){
  const holder = document.getElementById('tableHolder');
  if(rows.length===0){
    holder.innerHTML = '<div class="empty">No results match the current filters</div>';
    return;
  }
  holder.innerHTML = '<div class="mcards">' + rows.map(r=>{
    const eta = fmtDT(r.etaDT), etd = fmtDT(r.etdDT);
    const cd = fmtDT(r.cutoffDryDT), cr = fmtDT(r.cutoffReeferDT), og = fmtDT(r.opengateDT);
    const noLoad = r.no_l === 'Y' ? '<span class="noload">No load</span>' : '';
    return `<div class="mcard">
      <div class="mcard-main">
        <div class="mcard-top">
          <div>
            <div class="mcard-vessel">${r.vessel}<span class="vcode">${r.vessel_code}</span></div>
            <div class="mcard-sub">${r.vyg_bound} · ${r.op_liner} · ${r.service}</div>
          </div>
        </div>
        <div class="mcard-grid">
          <div class="mcard-item"><div class="lbl">Wharf</div><div class="val">${WHARF_MAP[r.wharf]||r.wharf}</div></div>
          <div class="mcard-item"><div class="lbl">Next Port</div><div class="val">${r.pod}</div></div>
          <div class="mcard-item"><div class="lbl">ETA</div><div class="val">${eta.d}<span class="t">${eta.t}</span></div></div>
          <div class="mcard-item"><div class="lbl">ETD</div><div class="val">${etd.d}<span class="t">${etd.t}</span></div></div>
          <div class="mcard-item full"><div class="lbl">1st Return</div><div class="val">${og.d}${flagFor(r.opengateDT,24,'Open','flag-open')}</div></div>
          <div class="mcard-item"><div class="lbl">Cut off (Dry)</div><div class="val">${cd.d}<span class="t">${cd.t}</span>${flagFor(r.cutoffDryDT,24,'Closed')}</div></div>
          <div class="mcard-item"><div class="lbl">Cut off (Reefer)</div><div class="val">${cr.d}<span class="t">${cr.t}</span>${flagFor(r.cutoffReeferDT,6,'Closed')}${noLoad}</div></div>
        </div>
      </div>
      <div class="mcard-stub"><div class="mcard-stub-pol">${r.pol}</div></div>
    </div>`;
  }).join('') + '</div>';
}

render();
</script>
</body>
</html>
"""


def build_html(records, wharf_map, source_filename):
    html = TEMPLATE.replace("__RAW_JSON__", json.dumps(records, ensure_ascii=False))
    html = html.replace("__WHARF_JSON__", json.dumps(wharf_map, ensure_ascii=False))
    html = html.replace("__SOURCE_FILENAME__", source_filename)
    return html


DEFAULT_OUTPUTS = ["Cut off - Open Gate Dashboard.html", "index.html"]


def main():
    ap = argparse.ArgumentParser(description="Rebuild the Cut off / Open gate dashboard.")
    ap.add_argument("cutoff_xls", help="Path to the daily CUT_OFF-style .xls file")
    ap.add_argument("wharf_xls", nargs="?",
                     help="Optional Wharf.xls (Wharf code -> Wharf Name). "
                          "If omitted, the built-in DEFAULT_WHARF_MAP is used.")
    ap.add_argument("-o", "--output", action="append",
                     help="Output HTML path (repeatable). "
                          "Default: writes both %s" % " and ".join(DEFAULT_OUTPUTS))
    args = ap.parse_args()

    records = load_cutoff_records(args.cutoff_xls)
    if args.wharf_xls:
        wharf_map = load_wharf_map(args.wharf_xls)
    else:
        wharf_map = dict(DEFAULT_WHARF_MAP)
        print(f"Using built-in wharf map ({len(wharf_map)} codes)")
    check_wharf_coverage(records, wharf_map)

    src_name = args.cutoff_xls.replace("\\", "/").split("/")[-1]
    html = build_html(records, wharf_map, src_name)

    for out in (args.output or DEFAULT_OUTPUTS):
        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Done -> {out}")


if __name__ == "__main__":
    main()
