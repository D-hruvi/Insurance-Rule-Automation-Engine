"""
Digit 2W Insurance Commission Grid Processor  v3
Key fixes:
  - POSP = PayIn * 0.8 always (confirmed from reference)
  - Fuel order: Petrol, LPG, Diesel, CNG  (matches reference)
  - Power values stored as formatted strings: '3.00', '7.00', etc.
  - SAOD Royal Enfield VehicleType = Bike (not ALL)
  - SAOD cluster case-insensitive matching fixed
  - Make field: 'Others' -> EXCLUDE list with reference brand order
  - Rule name: STATE_new_com_N_TW format using cluster abbreviation
  - 1+1 ordering: all RR rows, then all TP rows per cluster block
  - EV bonus rows at end of each 1+1 cluster block
"""

import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

OUTPUT_HEADERS = [
    "Rule Code", "Rule Name *", "IC Code *", "Product Type *", "Group",
    "Rule Type *", "Cover Type", "Business Type", "Vehicle Age", "State",
    "RTO", "Vehicle Category", "Vehicle Type", "Fuel Type", "Make", "Model",
    "Owner Type", "Usage Type", "Booking Mode", "Cover Selection Type",
    "Covers", "Addon Selection Type", "Addons", "CC From", "CC To",
    "Power From", "Power To", "GVW From", "GVW To", "Carrying From",
    "Carrying To", "NCB Type", "NCB From", "NCB To", "IDV From", "IDV To",
    "OD Discount From", "OD Discount To", "Effect Start Date *",
    "Effect End Date *", "PayIn (Commision Type)", "PayIn (Reward Type)",
    "PayIn (Amount Percentage)", "PayIn (OD Amount)", "PayIn (TP Amount)",
    "POSP (Commision Type)", "POSP (Reward Type)", "POSP (Amount Percentage)",
    "POSP (OD Amount)", "POSP (TP Amount)"
]

MISP_RATE = 22.5
FUEL_PETROL = "Petrol, LPG, Diesel, CNG"
FUEL_ALL    = "ALL"
FUEL_EV     = "Electric"
EXCLUDE_OTHERS = "EXCLUDE: HERO MOTOCORP, BAJAJ, HONDA, ROYAL ENFIELD, TVS, SUZUKI, YAMAHA"

RTO_STATE_NAMES = {
    "AN": "ANDAMAN ISLANDS",   "AP": "ANDHRA PRADESH",
    "AR": "ARUNACHAL PRADESH", "AS": "ASSAM",
    "BR": "BIHAR",             "CG": "CHHATTISGARH",
    "CH": "CHANDIGARH",        "DD": "DAMAN AND DIU",
    "DL": "DELHI",             "DN": "DADRA AND NAGAR HAVELI",
    "GA": "GOA",               "GJ": "GUJARAT",
    "HP": "HIMACHAL PRADESH",  "HR": "HARYANA",
    "JH": "JHARKHAND",         "JK": "JAMMU KASHMIR",
    "KA": "KARNATAKA",         "KL": "KERALA",
    "LA": "LADAKH",            "LD": "LAKSHADWEEP",
    "MH": "MAHARASHTRA",       "ML": "MEGHALAYA",
    "MN": "MANIPUR",           "MP": "MADHYA PRADESH",
    "MZ": "MIZORAM",           "NL": "NAGALAND",
    "OD": "ODISHA",            "OR": "ODISHA",
    "PB": "PUNJAB",            "PY": "PUDUCHERRY",
    "RJ": "RAJASTHAN",         "SK": "SIKKIM",
    "TG": "TELANGANA",         "TN": "TAMIL NADU",
    "TR": "TRIPURA",           "TS": "TELANGANA",
    "UA": "UTTARAKHAND",       "UK": "UTTARAKHAND",
    "UP": "UTTAR PRADESH",     "WB": "WEST BENGAL",
}

CLUSTER_ABBR = {
    "Andaman": "ANDAMAN", "APTS_Bad": "TS_AP_Bad", "APTS_Good": "TS_AP_Good",
    "APTS_Good1": "TS_AP_Good1", "APTS_Good2": "TS_AP_Good2",
    "APTS_Ref": "TS_AP_ref", "BR_Bad": "BR", "BR_Good": "BR",
    "CG+MP Bad": "CG_MP", "CG,MP,HR,HP ,WB REF": "REF", "CG_Good": "CG",
    "Chandigarh": "CH", "DL_Delhi": "DL", "GJ_ABS": "GJ", "GJ_Bad": "GJ",
    "GJ_Good": "GJ", "Goa": "GA", "HP_Bad": "HP", "HP_Good": "HP",
    "HR_Bad": "HR", "HR_Good": "HR", "J&K_Bad": "JK", "J&K_Jammu": "JK",
    "J&K_Srinagar": "JK", "Jharkhand": "JH", "KA_Bad": "KA",
    "KA_Bangalore": "KA", "MH_Bad": "MH", "MH_Good": "MH",
    "MH_Mumbai": "MH", "MH_Pune": "MH", "MP_Good": "MP", "NCR": "NCR",
    "NE_Bad": "NE", "NE_OR_Good": "NE_OR", "NE_OR_GOOD": "NE_OR",
    "NE_Ref": "NE", "OR_Bad": "OR", "PB_Bad": "PB", "PB_Good": "PB",
    "RJ_Bad": "RJ", "RJ_Good": "RJ", "RJ_Jaipur": "RJ",
    "ROM+RJ+GJ+UP+UK+OR REF": "ROM_REF", "TN+KL Good": "TN_KL",
    "TN+KL+KA REF": "TN_KL_KA", "TN+KL_Bad": "TN_KL", "TN_Chennai": "TN",
    "UP+UK_Bad": "UP_UK", "UP+UK_Bad2": "UP_UK", "UP+UK_Good": "UP_UK",
    "UP+UK_Good2": "UP_UK", "WB_Bad": "WB", "WB_Good": "WB",
    "WB_Kolkata": "WB",
    # SAOD clusters
    "AHMEDABAD,SURAT,BARODA": "GJ", "AP": "AP", "Assam": "AS", "BR": "BR",
    "CG": "CG", "Chennai_Coimbatore": "TN", "DL": "DL", "GA": "GA",
    "GJ": "GJ", "HR": "HR", "HP": "HP", "JK": "JK", "JH": "JH",
    "KA": "KA", "KL": "KL", "MP": "MP", "MH_Bad": "MH", "MH": "MH",
    "MH_Mumbai,PUNE": "MH", "NE": "NE", "OR": "OR", "PB": "PB",
    "PY": "PY", "RJ": "RJ", "TN": "TN", "TG": "TG", "UP": "UP",
    "UK": "UK", "WB": "WB",
}

def _abbr(cluster):
    return CLUSTER_ABBR.get(cluster,
        cluster.replace(" ","").replace(",","_").replace("+","_")[:10].upper())

# ── PayIn helpers ─────────────────────────────────────────────

def _pct(val):
    if val is None or val == 0:
        return "0"
    if isinstance(val, str):
        return "0"
    v = round(float(val) * 100, 4)
    s = f"{v:.4f}".rstrip("0").rstrip(".")
    return s

def resolve_payin(cd2):
    if cd2 is None or cd2 == 0:
        return ("net", "0")
    if isinstance(cd2, str):
        u = cd2.strip().upper()
        if u == "MISP":
            return ("od", str(MISP_RATE))
        if u == "D":
            return ("net", "0")
        try:
            return ("net", _pct(float(cd2)))
        except Exception:
            return ("net", "0")
    return ("net", _pct(cd2))

def posp_amt(payin_pct):
    """POSP = PayIn * 0.8 always"""
    try:
        p = float(payin_pct)
    except Exception:
        return "0"
    if p == 0:
        return "0"
    r = round(p * 0.8, 4)
    return f"{r:.4f}".rstrip("0").rstrip(".")

def fmt_power(val):
    """Format power as '3.00' style string"""
    if val is None:
        return None
    return f"{float(val):.2f}"

# ── Row builder ───────────────────────────────────────────────

def build_row(rule_name, state, rto,
              cover_type, biz_type, veh_age,
              vtype, fuel, make, model,
              ccf, cct, pwf, pwt,
              payin_type, payin_pct,
              eff_start, eff_end):
    p = posp_amt(payin_pct)
    return [
        None, rule_name, "DIGIT", "TW", None,
        "PAYINPAYOUT", cover_type, biz_type, veh_age, state, rto,
        None, vtype, fuel, make, model,
        "ALL", None, "any", "na", None, "na", None,
        ccf, cct, pwf, pwt,
        None, None, None, None,
        "na", None, None, None, None, None, None,
        eff_start, eff_end,
        payin_type, "percentage", payin_pct, "0", "0",
        "net", "percentage", p, "0", "0",
    ]

# ── Data loading ──────────────────────────────────────────────

def load_input_data(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    d = {}

    ws_rto = wb["2W RTO's"]
    rto_rows = list(ws_rto.iter_rows(values_only=True))
    d["rto"] = {}
    for r in rto_rows[2:]:
        if r[1]:
            d["rto"][r[1]] = {"1p1": r[2], "1p5": r[3], "saod": r[4]}

    # Build cluster->rto lookups; saod uses uppercase keys from RTO col
    for prod in ("1p1", "1p5", "saod"):
        key = f"c_{prod}"
        d[key] = {}
        for rto, c in d["rto"].items():
            cl = c.get(prod)
            if cl:
                cl_key = cl.strip().upper() if prod == "saod" else cl.strip()
                d[key].setdefault(cl_key, []).append(rto)

    ws = wb["TW 1+5"]
    d["tw_1p5"] = [
        {"cluster": r[1], "make": r[2], "seg": r[3], "cd2": r[5]}
        for r in ws.iter_rows(values_only=True) if r[1] and r[2] and r[3]
    ]

    ws = wb["TW 1+1 & SATP"]
    d["tw_1p1"] = [
        {"cluster": r[1], "seg": r[2], "cd2_1p1": r[4], "cd2_satp": r[5]}
        for r in ws.iter_rows(values_only=True) if r[1] and r[2] and r[3] is not None
    ]

    ws = wb["TW SAOD with Flexi Options"]
    d["tw_saod"] = [
        {"cluster": r[1], "seg": r[2],
         "yr1": r[7], "yr2": r[8], "yr3": r[9], "yr4": r[10]}
        for r in ws.iter_rows(values_only=True) if r[1] and r[2] and r[6] is not None
    ]

    wb.close()
    return d

def get_all_states(d):
    states = set()
    for rto in d["rto"]:
        s = RTO_STATE_NAMES.get(rto[:2].upper())
        if s:
            states.add(s)
    return sorted(states)

def _state_from_rtos(rtos):
    states = set()
    for rto in rtos:
        s = RTO_STATE_NAMES.get(rto[:2].upper())
        if s:
            states.add(s)
    return ", ".join(sorted(states))

def _rto_field(rtos):
    return ", ".join(sorted(rtos)) if rtos else "ALL"

# ── TW 1+5 segment → rows ─────────────────────────────────────

def seg_rows_1p5(make, seg, cd2, rule_pfx, state, rto, counter, eff_s, eff_e):
    rows = []
    s = str(seg).strip()
    pt, pp = resolve_payin(cd2)

    def add(vt, fuel, ccf=None, cct=None, pwf=None, pwt=None, use_pp=True):
        nonlocal counter
        name = f"{rule_pfx}_new_com_{counter}_TW"
        rows.append(build_row(name, state, rto,
                              "Comprehensive", "new", "ALL",
                              vt, fuel, make, "ALL",
                              ccf, cct,
                              fmt_power(pwf), fmt_power(pwt),
                              pt if use_pp else "net",
                              pp if use_pp else "0",
                              eff_s, eff_e))
        counter += 1

    # BAJAJ EV 3-7KW → 3 power-band rows
    if make == "BAJAJ" and s == "3-7 KW":
        add("ALL", FUEL_EV, pwf=3.00, pwt=7.00)
        add("ALL", FUEL_EV, pwf=0.10, pwt=2.90, use_pp=False)
        add("ALL", FUEL_EV, pwf=7.10, pwt=100.00, use_pp=False)
        return rows, counter

    # EV power bands
    if s in ("< 3 KW", "< 7 KW"):
        add("ALL", FUEL_EV, pwf=0.10, pwt=2.90)
    elif s in (">3 KW", "> 3 KW", ">7 KW", "> 7 KW"):
        add("ALL", FUEL_EV, pwf=3.10, pwt=100.00)
    elif s == "3-7 KW":
        add("ALL", FUEL_EV, pwf=3.00, pwt=7.00)
        add("ALL", FUEL_EV, pwf=0.10, pwt=2.90, use_pp=False)
        add("ALL", FUEL_EV, pwf=7.10, pwt=100.00, use_pp=False)
    elif s == "EV":
        add("ALL", FUEL_EV)
    elif s == "SCOOTER/EV":
        add("Scooter", FUEL_EV)
    # Bike CC bands
    elif s == "MC <=155":
        add("Bike", FUEL_PETROL, ccf=1, cct=155)
    elif s == "MC >155":
        add("Bike", FUEL_PETROL, ccf=156, cct=9999)
    elif s == "MC":
        add("Bike", FUEL_PETROL)
    elif s == "SCOOTER":
        add("Scooter", FUEL_ALL)
    elif s == "SCOOTER/MC":
        add("ALL", FUEL_PETROL)
    elif s in ("< =350", "<= 350"):
        add("ALL", FUEL_PETROL, ccf=1, cct=350)
    elif s in ("> 350", ">350"):
        add("ALL", FUEL_PETROL, ccf=351, cct=9999)
    elif s in ("All", "ALL"):
        add("ALL", FUEL_ALL)
    else:
        add("ALL", FUEL_ALL)

    return rows, counter

# ── 1+1 segment info ──────────────────────────────────────────

def veh_info_1p1(seg):
    s = str(seg).strip()
    if s == "SC/EV":
        return ("Scooter", FUEL_PETROL, None, None)
    if s in ("MC <= 180 Hero/Honda", "MC <= 180 Hero/Honda/TVS", "MC <= 180 Others"):
        return ("Bike", FUEL_PETROL, 1, 180)
    if s in ("MC_180-350_RE", "MC_180-350_HONDA/JAWA/Avenger",
             "MC_180-350_Other than RE", "MC_180-350_Others"):
        return ("Bike", FUEL_PETROL, 181, 350)
    if s == "MC>350":
        return ("Bike", FUEL_PETROL, 351, 9999)
    return ("ALL", FUEL_ALL, None, None)

def make_field_1p1(seg):
    s = str(seg).strip()
    if s == "SC/EV":
        return "ALL"
    if s == "MC <= 180 Hero/Honda":
        return "HERO MOTOCORP, HONDA"
    if s == "MC <= 180 Hero/Honda/TVS":
        return "HERO MOTOCORP, HONDA, TVS"
    if s == "MC <= 180 Others":
        return "EXCLUDE: HERO MOTOCORP, HONDA"
    if s == "MC_180-350_RE":
        return "ROYAL ENFIELD"
    if s == "MC_180-350_HONDA/JAWA/Avenger":
        return "HONDA, JAWA MOTORCYCLE, BAJAJ"
    if s in ("MC_180-350_Other than RE", "MC_180-350_Others"):
        return "EXCLUDE: BAJAJ, HONDA, ROYAL ENFIELD, JAWA MOTORCYCLE"
    if s == "MC>350":
        return "ALL"
    return "ALL"

# ── SAOD segment ──────────────────────────────────────────────

def veh_info_saod(seg):
    s = str(seg).strip()
    if s == "MC <155":
        return ("Bike",   FUEL_PETROL, 1,    155,  None, None, "EXCLUDE: ROYAL ENFIELD")
    if s == "MC>155":
        return ("Bike",   FUEL_PETROL, 156,  9999, None, None, "EXCLUDE: ROYAL ENFIELD")
    if s == "RE":
        return ("Bike",   FUEL_ALL,    None, None, None, None, "ROYAL ENFIELD")
    if s == "SC":
        return ("Scooter",FUEL_PETROL, None, None, None, None, "EXCLUDE: ROYAL ENFIELD")
    if s == "SC_EV":
        return ("Scooter",FUEL_EV,     None, None, None, None, "EXCLUDE: ROYAL ENFIELD")
    return ("ALL", FUEL_ALL, None, None, None, None, "ALL")

# ── Generators ────────────────────────────────────────────────

def gen_1p5(d, eff_s, eff_e, counter, state_filter=None):
    rows = []
    clusters = {}
    for e in d["tw_1p5"]:
        clusters.setdefault(e["cluster"], []).append(e)

    for cluster in sorted(clusters):
        entries = clusters[cluster]
        rtos = sorted(d["c_1p5"].get(cluster, []))
        if not rtos:
            continue
        state = _state_from_rtos(rtos)
        if state_filter and state_filter not in state.split(", "):
            continue
        rto_f = _rto_field(rtos)
        pfx = _abbr(cluster)

        for e in entries:
            make = e["make"]
            mf = EXCLUDE_OTHERS if make == "Others" else make
            new_rows, counter = seg_rows_1p5(mf, e["seg"], e["cd2"],
                                             pfx, state, rto_f,
                                             counter, eff_s, eff_e)
            rows.extend(new_rows)

    return rows, counter


def gen_1p1(d, eff_s, eff_e, counter, state_filter=None):
    rows = []
    clusters = {}
    for e in d["tw_1p1"]:
        clusters.setdefault(e["cluster"], []).append(e)

    for cluster in sorted(clusters):
        entries = clusters[cluster]
        rtos = sorted(d["c_1p1"].get(cluster, []))
        if not rtos:
            continue
        state = _state_from_rtos(rtos)
        if state_filter and state_filter not in state.split(", "):
            continue
        rto_f = _rto_field(rtos)
        pfx = _abbr(cluster)

        rr_rows = []
        tp_rows = []
        has_ev = False
        ev_cd2_rr = None
        ev_cd2_tp = None

        for e in entries:
            seg = e["seg"]
            vt, fuel, ccf, cct = veh_info_1p1(seg)
            mf = make_field_1p1(seg)

            if seg == "SC/EV":
                has_ev = True
                ev_cd2_rr = e["cd2_1p1"]
                ev_cd2_tp = e["cd2_satp"]
                fuel = FUEL_PETROL

            pt1, pp1 = resolve_payin(e["cd2_1p1"])
            pts, pps = resolve_payin(e["cd2_satp"])

            name_rr = f"{pfx}_all_RR_{counter}_TW"
            rr_rows.append(build_row(name_rr, state, rto_f,
                                     "Comprehensive", "renew, rollover", "ALL",
                                     vt, fuel, mf, "ALL",
                                     ccf, cct, None, None,
                                     pt1, pp1, eff_s, eff_e))
            counter += 1

            name_tp = f"{pfx}_all_TP_{counter}_TW"
            tp_rows.append(build_row(name_tp, state, rto_f,
                                     "TP", "ALL", "ALL",
                                     vt, fuel, mf, "ALL",
                                     ccf, cct, None, None,
                                     pts, pps, eff_s, eff_e))
            counter += 1

        rows.extend(rr_rows)
        rows.extend(tp_rows)

        if has_ev:
            pt1, pp1 = resolve_payin(ev_cd2_rr)
            pts, pps = resolve_payin(ev_cd2_tp)
            rows.append(build_row(f"{pfx}_ALL_EV_RR", state, rto_f,
                                  "Comprehensive", "renew, rollover", "ALL",
                                  "ALL", FUEL_EV, "ALL", "ALL",
                                  None, None, None, None,
                                  pt1, pp1, eff_s, eff_e))
            rows.append(build_row(f"{pfx}_TP_EV_1", state, rto_f,
                                  "TP", "ALL", "ALL",
                                  "ALL", FUEL_EV, "ALL", "ALL",
                                  None, None, None, None,
                                  pts, pps, eff_s, eff_e))

    return rows, counter


def gen_saod(d, eff_s, eff_e, counter, state_filter=None):
    rows = []
    clusters = {}
    for e in d["tw_saod"]:
        clusters.setdefault(e["cluster"], []).append(e)

    yr_keys = ["yr1", "yr2", "yr3", "yr4"]
    yr_ages = ["1", "2", "3", "4"]

    for cluster in sorted(clusters):
        entries = clusters[cluster]
        rtos = sorted(d["c_saod"].get(cluster.strip().upper(), []))
        if not rtos:
            continue
        state = _state_from_rtos(rtos)
        if state_filter and state_filter not in state.split(", "):
            continue
        rto_f = _rto_field(rtos)
        pfx = _abbr(cluster)

        for yr_key, yr_age in zip(yr_keys, yr_ages):
            for e in entries:
                vt, fuel, ccf, cct, pwf, pwt, mf = veh_info_saod(e["seg"])
                pt, pp = resolve_payin(e[yr_key])
                name = f"{pfx}_all_Od_{counter}_TW"
                rows.append(build_row(name, state, rto_f,
                                      "SAOD", "renew, rollover", yr_age,
                                      vt, fuel, mf, "ALL",
                                      ccf, cct, pwf, pwt,
                                      pt, pp, eff_s, eff_e))
                counter += 1

    return rows, counter

# ── Output writer ─────────────────────────────────────────────

def write_output_excel(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(OUTPUT_HEADERS)
    hfill = PatternFill("solid", fgColor="366092")
    hfont = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for row in rows:
        ws.append(row)
    for col in ws.columns:
        ml = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(ml + 2, 45)
    wb.save(path)

# ── Public API ────────────────────────────────────────────────

def generate_for_state(d, state_name, eff_s, eff_e, counter_start=1):
    rows, counter = [], counter_start
    r, counter = gen_1p5(d, eff_s, eff_e, counter, state_filter=state_name)
    rows.extend(r)
    r, counter = gen_saod(d, eff_s, eff_e, counter, state_filter=state_name)
    rows.extend(r)
    r, counter = gen_1p1(d, eff_s, eff_e, counter, state_filter=state_name)
    rows.extend(r)
    return rows, counter

def process_all(input_path, output_dir, eff_s, eff_e,
                states=None, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)
    if progress_callback:
        progress_callback("Loading source data...", 0, 1)
    d = load_input_data(input_path)
    all_states = get_all_states(d)
    targets = [s for s in all_states if (states is None or s in states)]
    generated = []
    counter = 1
    for idx, state in enumerate(targets):
        if progress_callback:
            progress_callback(f"Processing {state}...", idx, len(targets))
        rows, counter = generate_for_state(d, state, eff_s, eff_e, counter)
        if not rows:
            continue
        fname = state.replace(", ", "_").replace(" ", "").replace(",", "_") + "_2W.xlsx"
        out = os.path.join(output_dir, fname)
        write_output_excel(rows, out)
        generated.append(out)
    if progress_callback:
        progress_callback("Complete!", len(targets), len(targets))
    return generated

if __name__ == "__main__":
    import time
    INPUT = "/mnt/user-data/uploads/Large_Insurance_Brokers_Feb_26_h_M.xlsx"
    os.makedirs("/tmp/out_v3", exist_ok=True)
    d = load_input_data(INPUT)

    # Quick check Andaman
    rows, _ = generate_for_state(d, "ANDAMAN ISLANDS", "2026-02-01", "2026-02-28")
    print(f"Andaman: {len(rows)} rows (ref=61)")

    # Quick check AP
    rows, _ = generate_for_state(d, "ANDHRA PRADESH", "2026-02-01", "2026-02-28")
    print(f"AP: {len(rows)} rows (ref=139)")

    t0 = time.time()
    files = process_all(INPUT, "/tmp/out_v3", "2026-02-01", "2026-02-28")
    print(f"All states done in {time.time()-t0:.1f}s — {len(files)} files")
