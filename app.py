import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys
import io
import re

st.set_page_config(layout="wide", page_title="KIM Schedule 2026")

# =========================
# CUSTOM CSS
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

[data-testid="stSidebar"] {
    background: #0f1923 !important;
    border-right: 1px solid #1e2d3d;
}
[data-testid="stSidebar"] * { color: #c8d8e8 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #7a9bb5 !important; font-size: 11px; font-weight: 600;
    letter-spacing: 0.08em; text-transform: uppercase;
}
[data-testid="stSidebar"] [data-baseweb="select"] {
    background: #1a2a3a !important; border-color: #2a3f55 !important;
}

.main .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1600px; }

.dashboard-header {
    background: linear-gradient(135deg, #0f1923 0%, #1a2f47 50%, #0d2137 100%);
    border-radius: 12px; padding: 24px 32px; margin-bottom: 24px;
    display: flex; align-items: center; justify-content: space-between;
    border: 1px solid #1e3a55;
}
.header-title { font-size: 24px; font-weight: 700; color: #f0f6ff; margin: 0; }
.header-subtitle { font-size: 13px; color: #6b92b5; margin-top: 4px; }
.header-badge {
    background: rgba(59,130,246,0.15); border: 1px solid rgba(59,130,246,0.3);
    color: #7eb8f7; padding: 6px 14px; border-radius: 20px;
    font-size: 12px; font-weight: 600;
}

.metric-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px; }
.metric-card {
    background: #ffffff; border: 1px solid #e8ecf0; border-radius: 10px;
    padding: 14px 18px; position: relative; overflow: hidden;
}
.metric-card::before {
    content: ''; position: absolute; left: 0; top: 0; bottom: 0;
    width: 3px; border-radius: 3px 0 0 3px;
}
.metric-card.blue::before  { background: #3b82f6; }
.metric-card.teal::before  { background: #14b8a6; }
.metric-card.amber::before { background: #f59e0b; }
.metric-card.rose::before  { background: #f43f5e; }
.metric-label { font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #94a3b8; margin-bottom: 4px; }
.metric-value { font-size: 26px; font-weight: 700; color: #0f1923; font-family: 'DM Mono', monospace; }
.metric-sub { font-size: 11px; color: #94a3b8; margin-top: 2px; }

.section-header {
    font-size: 12px; font-weight: 700; color: #334155; letter-spacing: 0.08em;
    text-transform: uppercase; margin-bottom: 10px; padding-bottom: 6px;
    border-bottom: 2px solid #e2e8f0;
}

/* === PLANNING TABLE === */
.plan-table-wrap { overflow: auto; max-height: 75vh; border: 1px solid #e2e8f0; border-radius: 8px; }
table.plan { border-collapse: collapse; font-size: 11px; white-space: nowrap; min-width: 100%; }
table.plan th, table.plan td { border: 0.5px solid #e2e8f0; padding: 2px 6px; }
table.plan thead th {
    background: #f8fafc; font-weight: 600; text-align: center;
    position: sticky; top: 0; z-index: 10; font-size: 10px;
}
table.plan thead tr:nth-child(1) th { top: 0; z-index: 12; }
table.plan thead tr:nth-child(2) th { top: 22px; z-index: 11; }
table.plan thead tr:nth-child(3) th { top: 44px; z-index: 10; }
table.plan tbody tr:hover { background: #f0f6ff !important; }
.th-month { background: #1e3a55 !important; color: #7eb8f7 !important; font-weight: 700 !important; }
.th-meta { background: #f1f5f9 !important; min-width: 36px; }
.td-name {
    min-width: 280px; max-width: 360px; white-space: normal; font-size: 11px;
    background: #ffffff; position: sticky; left: 120px; z-index: 5;
    border-right: 2px solid #cbd5e1 !important;
}
.th-name { min-width: 280px; text-align: left !important; position: sticky; left: 120px; z-index: 13 !important; }
.td-col0 { width: 36px; text-align: center; position: sticky; left: 0; z-index: 5; background: #f8fafc; }
.td-col1 { width: 42px; text-align: center; position: sticky; left: 36px; z-index: 5; background: #f8fafc; }
.td-col2 { width: 42px; text-align: center; position: sticky; left: 78px; z-index: 5; background: #f8fafc; }
.th-col0 { position: sticky !important; left: 0; z-index: 14 !important; }
.th-col1 { position: sticky !important; left: 36px; z-index: 14 !important; }
.th-col2 { position: sticky !important; left: 78px; z-index: 14 !important; }
.cell-has {
    text-align: center; font-weight: 600; cursor: default;
    font-family: 'DM Mono', monospace; font-size: 11px;
}
.cell-empty { background: transparent; }
.section-row td { font-weight: 700; font-size: 11px; padding: 3px 8px; }

/* Conflicts */
.conflict-card {
    border-radius: 8px; padding: 10px 14px; margin-bottom: 8px;
    font-size: 13px; border-left: 4px solid;
}
.conflict-card.room { background: #fff5f5; border-color: #ef4444; border: 1px solid #fecaca; border-left: 4px solid #ef4444; }
.conflict-card.person { background: #faf5ff; border-color: #8b5cf6; border: 1px solid #ddd6fe; border-left: 4px solid #8b5cf6; }
.conflict-title { font-weight: 700; font-size: 13px; margin-bottom: 3px; }
.conflict-detail { color: #64748b; font-size: 11px; font-family: 'DM Mono', monospace; }

/* Calendar */
.cal-wrap { font-family: 'DM Sans', sans-serif; margin-bottom: 28px; }
.cal-title { font-size: 16px; font-weight: 700; margin-bottom: 8px; color: #0f1923; }
.cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; }
.cal-hdr { text-align: center; font-weight: 700; font-size: 10px; padding: 4px 0; color: #94a3b8; background: #f8fafc; border-radius: 4px 4px 0 0; letter-spacing: 0.05em; }
.cal-cell { border: 1px solid #e8ecf2; border-radius: 6px; padding: 4px; min-height: 80px; background: #fff; box-sizing: border-box; position: relative; }
.cal-cell-wknd { background: #fafbfc; }
.cal-cell-today { border: 2px solid #3b82f6; background: #eff6ff; }
.cal-cell-conflict { border: 2px solid #f97316 !important; }
.cal-cell-empty { border: none; background: transparent; min-height: 80px; }
.cal-daynum { font-weight: 700; font-size: 11px; color: #475569; margin-bottom: 2px; font-family: 'DM Mono', monospace; }
.cal-conflict-dot { position: absolute; top: 3px; right: 4px; width: 6px; height: 6px; background: #f97316; border-radius: 50%; }
.cal-ev { padding: 2px 4px; margin-top: 2px; border-radius: 3px; font-size: 9px; line-height: 1.3; word-break: break-word; }
.cal-ev-time { font-weight: 700; display: block; font-family: 'DM Mono', monospace; }
.cal-ev-name { display: block; font-size: 8.5px; opacity: 0.9; }
.cal-more { font-size: 9px; color: #94a3b8; font-style: italic; margin-top: 2px; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #f1f5f9; border-radius: 8px; padding: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 6px; font-weight: 600; font-size: 13px; padding: 6px 14px; }
.stTabs [aria-selected="true"] { background: #fff !important; color: #0f1923 !important; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }

.stDownloadButton button {
    background: #0f1923 !important; color: #7eb8f7 !important;
    border: 1px solid #1e3a55 !important; border-radius: 6px !important;
    font-weight: 600 !important; font-size: 12px !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("""
<div class="dashboard-header">
  <div>
    <div class="header-title">KIM — Sitzungsplanung 2026</div>
    <div class="header-subtitle">Klinik für Intensivmedizin · Inselspital Bern</div>
  </div>
  <div class="header-badge">2026 PLANUNG</div>
</div>
""", unsafe_allow_html=True)

# =========================
# DATA LOADING — two files
# =========================
WEEKS = [
    "2025-12-29","2026-01-05","2026-01-12","2026-01-19","2026-01-26",
    "2026-02-02","2026-02-09","2026-02-16","2026-02-23",
    "2026-03-02","2026-03-09","2026-03-16","2026-03-23","2026-03-30",
    "2026-04-06","2026-04-13","2026-04-20","2026-04-27",
    "2026-05-04","2026-05-11","2026-05-18","2026-05-25",
    "2026-06-01","2026-06-08","2026-06-15","2026-06-22","2026-06-29",
    "2026-07-06","2026-07-13","2026-07-20","2026-07-27",
    "2026-08-03","2026-08-10","2026-08-17","2026-08-24","2026-08-31",
    "2026-09-07","2026-09-14","2026-09-21","2026-09-28",
    "2026-10-05","2026-10-12","2026-10-19","2026-10-26",
    "2026-11-02","2026-11-09","2026-11-16","2026-11-23","2026-11-30",
    "2026-12-07","2026-12-14","2026-12-21",
]
WEEK_NUMS = list(range(1, 53))  # 52 weeks

MONTHS = {
    0: "Januar", 5: "Februar", 9: "März", 13: "April",
    18: "Mai", 22: "Juni", 26: "Juli", 31: "August",
    35: "September", 39: "Oktober", 44: "November", 48: "Dezember"
}

MONTH_SPANS = []
m_keys = sorted(MONTHS.keys())
for i, k in enumerate(m_keys):
    end = m_keys[i+1] - 1 if i+1 < len(m_keys) else 51
    MONTH_SPANS.append({"label": MONTHS[k], "start": k, "end": end})

SECTION_COLORS = {
    "IB Pflege":       ("#dbeafe", "#1e40af"),
    "Wund & Rehab":    ("#fef3c7", "#92400e"),
    "Geräte & ECMO":   ("#fee2e2", "#991b1b"),
    "EPIC":            ("#e0f2fe", "#0c4a6e"),
    "Kommunikation":   ("#fce7f3", "#831843"),
    "Einführung":      ("#dcfce7", "#14532d"),
    "Ausbildung":      ("#ede9fe", "#4c1d95"),
    "Events":          ("#f3f4f6", "#374151"),
}

CELL_COLORS = {
    "IB Pflege":       "#bfdbfe",
    "Wund & Rehab":    "#fde68a",
    "Geräte & ECMO":   "#fca5a5",
    "EPIC":            "#7dd3fc",
    "Kommunikation":   "#f9a8d4",
    "Einführung":      "#86efac",
    "Ausbildung":      "#c4b5fd",
    "Events":          "#d1d5db",
}

# Sidebar — upload both files
st.sidebar.markdown("""
<div style="padding:12px 0 8px 0;">
  <div style="font-size:9px;font-weight:700;letter-spacing:0.15em;color:#4a6f8a;text-transform:uppercase;margin-bottom:2px;">KIM ICU</div>
  <div style="font-size:15px;font-weight:700;color:#e8f4ff;margin-bottom:12px;">Sitzungsplanung</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown('<div style="font-size:10px;color:#7a9bb5;font-weight:600;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;">Datei hochladen</div>', unsafe_allow_html=True)

planning_file = st.sidebar.file_uploader(
    "Jahresplanung (Fachgruppentage…xlsx)",
    type=["xlsx"],
    key="planning",
    help="Fachgruppentage_und_Sitzungen_KIM_2026.xlsx"
)
events_file = st.sidebar.file_uploader(
    "Sitzungsdaten (Sitzungsdaten_sort.xlsx)",
    type=["xlsx"],
    key="events",
    help="Sitzungsdaten_sort.xlsx"
)


@st.cache_data(show_spinner=False)
def load_planning(file_bytes):
    """Parse the Fachgruppentage planning matrix."""
    import io as _io
    df_raw = pd.read_excel(_io.BytesIO(file_bytes), sheet_name="Sitzungsdaten 2026", header=None)

    # Detect sections by background color (fallback: row structure)
    # We'll use positional heuristics matching the known layout
    section_map = {
        range(3, 22):  "IB Pflege",
        range(23, 34): "Wund & Rehab",
        range(36, 52): "Geräte & ECMO",
        range(54, 58): "EPIC",
        range(60, 64): "Kommunikation",
        range(65, 70): "Einführung",
        range(69, 83): "Ausbildung",
        range(85, 88): "Events",
    }

    def get_section(row_idx):
        for rng, sec in section_map.items():
            if row_idx in rng:
                return sec
        return "Sonstiges"

    rows = []
    for r in range(3, min(88, len(df_raw))):
        tag   = df_raw.iloc[r, 0]
        fa_ib = df_raw.iloc[r, 1]
        wer   = df_raw.iloc[r, 2]
        name  = df_raw.iloc[r, 3]

        if pd.isna(name) or str(name).strip() in ["nan", ""]:
            continue

        values = []
        for c in range(4, 56):
            if c >= df_raw.shape[1]:
                values.append(None)
                continue
            val = df_raw.iloc[r, c]
            if pd.isna(val) or isinstance(val, bool):
                values.append(None)
            else:
                try:
                    v = float(val)
                    values.append(v if v != 0 else None)
                except Exception:
                    values.append(None)

        rows.append({
            "tag":     str(tag).strip()   if pd.notna(tag)   else "",
            "fa_ib":   str(fa_ib).strip() if pd.notna(fa_ib) else "",
            "wer":     str(wer).strip()   if pd.notna(wer)   else "",
            "name":    str(name).strip(),
            "values":  values,
            "section": get_section(r),
        })
    return rows


@st.cache_data(show_spinner=False)
def load_events(file_bytes):
    """Parse Sitzungsdaten_sort.xlsx — the detailed event list."""
    import io as _io
    df = pd.read_excel(_io.BytesIO(file_bytes), header=None)
    # First row is header
    df.columns = [str(c) for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    # Normalise column names
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if "datum" in cl:    col_map[c] = "Datum"
        elif "zeit" in cl:   col_map[c] = "Zeit"
        elif "ort" in cl:    col_map[c] = "Ort"
        elif "event" in cl or "veranst" in cl: col_map[c] = "Event"
        elif "person" in cl: col_map[c] = "Personen"
        elif "bemerk" in cl: col_map[c] = "Bemerkungen"
        elif "tag" == cl:    col_map[c] = "Tag"
    df = df.rename(columns=col_map)
    for col in ["Datum","Zeit","Ort","Event","Personen","Bemerkungen","Tag"]:
        if col not in df.columns:
            df[col] = ""
    df["Datum"]    = pd.to_datetime(df["Datum"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Datum"])
    df["Zeit"]     = df["Zeit"].fillna("").astype(str)
    df["Personen"] = df["Personen"].fillna("").astype(str)
    df["Ort"]      = df["Ort"].fillna("").astype(str)
    df["Event"]    = df["Event"].fillna("").astype(str)
    df["Month"]    = df["Datum"].dt.month
    df["Day"]      = df["Datum"].dt.day
    return df.sort_values("Datum")


# =========================
# CATEGORY + COLOR for events list
# =========================
BASE_COLORS = {
    "EPIC": "#4FC3F7", "Geräte & Beatmung": "#E57373",
    "Workshop": "#FFD54F", "Schulung/Kurs": "#81C784",
    "Sitzung": "#FF8A65", "Einführung": "#4DB6AC",
    "Lernwerkstatt": "#F48FB1", "Führung/Austausch": "#CE93D8",
    "ICU": "#90CAF9", "Kommunikation": "#FFCC80",
    "Planung": "#A5D6A7", "Sonstiges": "#B0BEC5",
}

def get_category(event):
    e = str(event).lower()
    if "epic" in e:                                              return "EPIC"
    if any(x in e for x in ["ecmo","lvad","impella","assist","prismax","beatmung"]): return "Geräte & Beatmung"
    if "workshop" in e:                                          return "Workshop"
    if any(x in e for x in ["schulung","kurs","basiskurs","aufbaukurs","refresher"]): return "Schulung/Kurs"
    if any(x in e for x in ["sitzung","fachgruppe","fg ","superuser","fachforum"]): return "Sitzung"
    if "einführung" in e or "einblick" in e:                     return "Einführung"
    if any(x in e for x in ["lernwerkstatt","repe","simulation","kimsim"]): return "Lernwerkstatt"
    if any(x in e for x in ["führungsdialog","austausch"]):      return "Führung/Austausch"
    if "icu" in e:                                               return "ICU"
    if any(x in e for x in ["kommunikation","aggressions"]):     return "Kommunikation"
    if any(x in e for x in ["planung","bürotag"]):               return "Planung"
    return "Sonstiges"

def make_color_map(events):
    def gen_shades(base_hex, n):
        base_hex = base_hex.lstrip("#")
        r,g,b = int(base_hex[0:2],16)/255, int(base_hex[2:4],16)/255, int(base_hex[4:6],16)/255
        h,l,s = colorsys.rgb_to_hls(r,g,b)
        shades = []
        for i in range(n):
            nl = max(0.35, min(0.75, l+(i-n/2)*0.08))
            r2,g2,b2 = colorsys.hls_to_rgb(h, nl, s)
            shades.append("#{:02x}{:02x}{:02x}".format(int(r2*255),int(g2*255),int(b2*255)))
        return shades

    groups = {}
    for ev in events:
        cat = get_category(ev)
        groups.setdefault(cat, []).append(ev)
    cmap = {}
    for cat, evs in groups.items():
        base = BASE_COLORS.get(cat, "#B0BEC5")
        shades = gen_shades(base, len(evs))
        for ev, color in zip(sorted(evs), shades):
            cmap[ev] = color
    return cmap


# =========================
# TIME HELPERS
# =========================
def parse_time_range(zeit_str):
    s = str(zeit_str).strip()
    m = re.match(r"(\d{1,2})[:\.](\d{2})\s*[-–]\s*(\d{1,2})[:\.](\d{2})", s)
    if m:
        try:
            return (datetime.time(int(m.group(1)), int(m.group(2))),
                    datetime.time(int(m.group(3)), int(m.group(4))))
        except Exception:
            return None
    return None

def times_overlap(t1, t2):
    if not t1 or not t2: return False
    s1, e1 = t1; s2, e2 = t2
    return s1 < e2 and s2 < e1

def fmt_week_date(d):
    if not d: return ""
    parts = d.split("-")
    return f"{parts[2]}.{parts[1]}.{parts[0][2:]}"


# =========================
# CONFLICT DETECTION (event list)
# =========================
def find_conflicts(df_ev):
    room_conflicts, person_conflicts = [], []
    for date, day_df in df_ev.groupby("Datum"):
        rows = day_df.reset_index(drop=True).to_dict("records")
        for i in range(len(rows)):
            for j in range(i+1, len(rows)):
                r1, r2 = rows[i], rows[j]
                t1 = parse_time_range(r1["Zeit"])
                t2 = parse_time_range(r2["Zeit"])
                ort1, ort2 = str(r1["Ort"]).strip(), str(r2["Ort"]).strip()
                if ort1 and ort2 and ort1.lower() == ort2.lower() and times_overlap(t1, t2):
                    room_conflicts.append({
                        "Datum": date.strftime("%d.%m.%Y"), "Raum": ort1,
                        "Event 1": r1["Event"], "Zeit 1": r1["Zeit"],
                        "Event 2": r2["Event"], "Zeit 2": r2["Zeit"],
                    })
                p1s = {p.strip() for p in str(r1["Personen"]).split("/") if p.strip()}
                p2s = {p.strip() for p in str(r2["Personen"]).split("/") if p.strip()}
                shared = p1s & p2s
                if shared and times_overlap(t1, t2):
                    person_conflicts.append({
                        "Datum": date.strftime("%d.%m.%Y"), "Person": ", ".join(shared),
                        "Event 1": r1["Event"], "Zeit 1": r1["Zeit"],
                        "Event 2": r2["Event"], "Zeit 2": r2["Zeit"],
                    })
    return room_conflicts, person_conflicts


# =========================
# iCAL EXPORT
# =========================
def make_ical(df_ev):
    import hashlib
    lines = [
        "BEGIN:VCALENDAR","VERSION:2.0",
        "PRODID:-//KIM ICU Schedule//DE","CALSCALE:GREGORIAN",
        "METHOD:PUBLISH","X-WR-CALNAME:KIM ICU Schedule 2026",
        "X-WR-TIMEZONE:Europe/Zurich",
    ]
    for _, row in df_ev.iterrows():
        uid = hashlib.md5(f"{row['Datum']}{row['Event']}{row['Zeit']}".encode()).hexdigest()
        date_str = row["Datum"].strftime("%Y%m%d")
        t = parse_time_range(row["Zeit"])
        if t:
            dtstart = f"DTSTART;TZID=Europe/Zurich:{date_str}T{t[0].strftime('%H%M%S')}"
            dtend   = f"DTEND;TZID=Europe/Zurich:{date_str}T{t[1].strftime('%H%M%S')}"
        else:
            dtstart = f"DTSTART;VALUE=DATE:{date_str}"
            dtend   = f"DTEND;VALUE=DATE:{date_str}"
        summary  = str(row["Event"]).replace(",","\\,").replace(";","\\;")
        location = str(row["Ort"]).replace(",","\\,")
        desc     = f"Person: {row['Personen']}\\nOrt: {row['Ort']}"
        if pd.notna(row.get("Bemerkungen","")) and str(row.get("Bemerkungen","")).strip():
            desc += f"\\nBemerkung: {row['Bemerkungen']}"
        lines += [
            "BEGIN:VEVENT", f"UID:{uid}@kim.insel.ch",
            f"DTSTAMP:{datetime.datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            dtstart, dtend, f"SUMMARY:{summary}",
            f"LOCATION:{location}", f"DESCRIPTION:{desc}", "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


# =========================
# PLANNING TABLE RENDERER
# =========================
def render_planning_table(plan_rows, filter_month=0, filter_section="", search_q=""):
    # Determine visible week indices
    if filter_month > 0:
        month_names = ["","Januar","Februar","März","April","Mai","Juni",
                       "Juli","August","September","Oktober","November","Dezember"]
        ms = next((s for s in MONTH_SPANS if s["label"] == month_names[filter_month]), None)
        vis_weeks = list(range(ms["start"], ms["end"]+1)) if ms else list(range(52))
    else:
        vis_weeks = list(range(52))

    # Filter rows
    vis_rows = [r for r in plan_rows
                if (not filter_section or r["section"] == filter_section)
                and (not search_q or search_q.lower() in r["name"].lower())]

    # Build visible month spans for header
    vis_month_spans = []
    for ms in MONTH_SPANS:
        cols = [i for i in range(ms["start"], ms["end"]+1) if i in vis_weeks]
        if cols:
            vis_month_spans.append({"label": ms["label"], "cols": cols})

    html = '<div class="plan-table-wrap"><table class="plan">'

    # === HEADER ROW 1: month names ===
    html += "<thead><tr>"
    html += '<th class="th-col0 th-meta" rowspan="3">Tag</th>'
    html += '<th class="th-col1 th-meta" rowspan="3">FA/<br>IB</th>'
    html += '<th class="th-col2 th-meta" rowspan="3">Wer?</th>'
    html += '<th class="th-name th-meta" rowspan="3">Bezeichnung</th>'
    for ms in vis_month_spans:
        html += f'<th class="th-month" colspan="{len(ms["cols"])}">{ms["label"]}</th>'
    html += "</tr>"

    # === HEADER ROW 2: week numbers ===
    html += "<tr>"
    for wi in vis_weeks:
        html += f'<th class="th-meta" style="font-size:9px;color:#64748b">{WEEK_NUMS[wi]}</th>'
    html += "</tr>"

    # === HEADER ROW 3: dates ===
    html += "<tr>"
    for wi in vis_weeks:
        html += f'<th class="th-meta" style="font-size:9px;color:#94a3b8;font-family:DM Mono,monospace">{fmt_week_date(WEEKS[wi])}</th>'
    html += "</tr></thead><tbody>"

    # === DATA ROWS ===
    last_section = None
    for row in vis_rows:
        # Section separator
        if row["section"] != last_section:
            last_section = row["section"]
            bg, fg = SECTION_COLORS.get(row["section"], ("#f8fafc","#334155"))
            total_cols = 4 + len(vis_weeks)
            html += (f'<tr class="section-row">'
                     f'<td colspan="{total_cols}" style="background:{bg};color:{fg}">'
                     f'{row["section"]}</td></tr>')

        cell_color = CELL_COLORS.get(row["section"], "#e2e8f0")

        html += "<tr>"
        html += f'<td class="td-col0" style="font-size:10px;color:#64748b">{row["tag"]}</td>'
        html += f'<td class="td-col1" style="font-size:10px;color:#64748b">{row["fa_ib"]}</td>'
        html += f'<td class="td-col2" style="font-size:10px;color:#64748b">{row["wer"]}</td>'
        html += f'<td class="td-name">{row["name"]}</td>'

        for wi in vis_weeks:
            v = row["values"][wi] if wi < len(row["values"]) else None
            if v is not None:
                disp = int(v) if v == int(v) else round(v, 1)
                # All scheduled dates for tooltip
                all_dates = [f"KW{WEEK_NUMS[i]} {fmt_week_date(WEEKS[i])}" 
                             for i,val in enumerate(row["values"]) if val is not None]
                tip = f"{row['name']} | KW{WEEK_NUMS[wi]} ({fmt_week_date(WEEKS[wi])}): {disp} | Alle: {', '.join(all_dates)}"
                html += (f'<td class="cell-has" '
                         f'style="background:{cell_color};color:#1e3a5f" '
                         f'title="{tip}">{disp}</td>')
            else:
                html += '<td class="cell-empty"></td>'
        html += "</tr>"

    html += "</tbody></table></div>"
    return html, len(vis_rows), len(vis_weeks)


# =========================
# SIDEBAR FILTERS
# =========================
st.sidebar.markdown("---")
st.sidebar.markdown('<div style="font-size:9px;font-weight:700;letter-spacing:0.12em;color:#4a6f8a;text-transform:uppercase;margin-bottom:6px;">Filter</div>', unsafe_allow_html=True)

plan_rows = []
df_ev = pd.DataFrame()

if planning_file:
    plan_rows = load_planning(planning_file.read())
    planning_file.seek(0)

if events_file:
    df_ev = load_events(events_file.read())
    events_file.seek(0)

# Sidebar filters (for events list)
month_options = ["Ganzes Jahr"] + ["Januar","Februar","März","April","Mai","Juni",
                                    "Juli","August","September","Oktober","November","Dezember"]
selected_month_name = st.sidebar.selectbox("Monat", month_options)
selected_month_num  = month_options.index(selected_month_name)  # 0 = all

if not df_ev.empty:
    event_filter  = st.sidebar.multiselect("Event-Typ",  sorted(df_ev["Event"].unique()))
    person_filter = st.sidebar.multiselect("Person",     sorted(df_ev["Personen"].unique()))
    room_filter   = st.sidebar.multiselect("Raum",       sorted(df_ev["Ort"].unique()))
else:
    event_filter = person_filter = room_filter = []

section_filter = st.sidebar.selectbox(
    "Bereich (Planungstabelle)",
    ["Alle"] + list(SECTION_COLORS.keys())
)
section_filter_val = "" if section_filter == "Alle" else section_filter

# iCal export (events list)
if not df_ev.empty:
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div style="font-size:9px;font-weight:700;letter-spacing:0.12em;color:#4a6f8a;text-transform:uppercase;margin-bottom:6px;">iCal Export</div>', unsafe_allow_html=True)
    filtered_ev = df_ev.copy()
    if selected_month_num > 0: filtered_ev = filtered_ev[filtered_ev["Month"]==selected_month_num]
    if event_filter:   filtered_ev = filtered_ev[filtered_ev["Event"].isin(event_filter)]
    if person_filter:  filtered_ev = filtered_ev[filtered_ev["Personen"].isin(person_filter)]
    if room_filter:    filtered_ev = filtered_ev[filtered_ev["Ort"].isin(room_filter)]

    ical_data = make_ical(filtered_ev)
    st.sidebar.download_button(
        label=f"Gefilterte Events ({len(filtered_ev)}) → .ics",
        data=ical_data.encode("utf-8"),
        file_name="kim_schedule_filtered.ics", mime="text/calendar",
    )
    ical_all = make_ical(df_ev)
    st.sidebar.download_button(
        label=f"Alle Events ({len(df_ev)}) → .ics",
        data=ical_all.encode("utf-8"),
        file_name="kim_schedule_2026_all.ics", mime="text/calendar",
    )
else:
    filtered_ev = df_ev.copy()

# =========================
# UPLOAD PROMPT
# =========================
if not plan_rows and df_ev.empty:
    st.info("Bitte links die Excel-Dateien hochladen: **Fachgruppentage_und_Sitzungen_KIM_2026.xlsx** (Jahresplanung) und optional **Sitzungsdaten_sort.xlsx** (Detaillierte Sitzungsliste).")
    st.stop()

# =========================
# METRICS
# =========================
rc, pc = (find_conflicts(filtered_ev) if not filtered_ev.empty else ([], []))
total_plan_entries = sum(1 for r in plan_rows for v in r["values"] if v is not None)
busy_weeks = max(
    (sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)),
    default=0
)

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card blue">
    <div class="metric-label">Planungseinträge</div>
    <div class="metric-value">{total_plan_entries}</div>
    <div class="metric-sub">in der Jahresplanung</div>
  </div>
  <div class="metric-card teal">
    <div class="metric-label">Ereignisarten</div>
    <div class="metric-value">{len(plan_rows)}</div>
    <div class="metric-sub">verschiedene Ereignisse</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">Raum-Konflikte</div>
    <div class="metric-value">{len(rc)}</div>
    <div class="metric-sub">Doppelbuchungen erkannt</div>
  </div>
  <div class="metric-card rose">
    <div class="metric-label">Personen-Konflikte</div>
    <div class="metric-value">{len(pc)}</div>
    <div class="metric-sub">Überlappungen erkannt</div>
  </div>
</div>
""", unsafe_allow_html=True)

# =========================
# TABS
# =========================
tab_plan, tab_list, tab_calendar, tab_conflicts, tab_ical, tab_heatmap, tab_bar = st.tabs([
    "📋 Jahresplanung", "📄 Sitzungsliste", "📅 Kalender",
    "⚠️ Konflikte", "📤 iCal Export", "🔥 Heatmap", "📊 Statistik"
])

# ======================================
# TAB 1 — JAHRESPLANUNG (PRIMARY)
# ======================================
with tab_plan:
    st.markdown('<div class="section-header">Jahresplanung 2026 — Wochenübersicht</div>', unsafe_allow_html=True)

    if not plan_rows:
        st.warning("Bitte Fachgruppentage-Datei hochladen.")
    else:
        c1, c2 = st.columns([3, 1])
        with c1:
            search_q = st.text_input("Suchen", placeholder="Ereignis suchen...", label_visibility="collapsed")
        with c2:
            st.caption(f"{len(plan_rows)} Ereignisarten · 52 Wochen")

        table_html, n_rows, n_weeks = render_planning_table(
            plan_rows,
            filter_month   = selected_month_num,
            filter_section = section_filter_val,
            search_q       = search_q,
        )
        st.markdown(table_html, unsafe_allow_html=True)

        # Legend
        st.markdown("<br>", unsafe_allow_html=True)
        leg_cols = st.columns(len(SECTION_COLORS))
        for col, (sec, (bg, fg)) in zip(leg_cols, SECTION_COLORS.items()):
            cell_bg = CELL_COLORS.get(sec, "#e2e8f0")
            col.markdown(
                f'<div style="background:{cell_bg};color:{fg};padding:3px 8px;border-radius:4px;font-size:10px;font-weight:600;text-align:center">{sec}</div>',
                unsafe_allow_html=True
            )


# ======================================
# TAB 2 — SITZUNGSLISTE
# ======================================
with tab_list:
    st.markdown('<div class="section-header">Detaillierte Sitzungsliste</div>', unsafe_allow_html=True)
    if filtered_ev.empty:
        st.info("Sitzungsdaten-Datei hochladen für die Detailansicht.")
    else:
        display_df = filtered_ev.copy()
        display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")
        cols = [c for c in ["Tag","Datum","Zeit","Event","Personen","Ort","Bemerkungen"] if c in display_df.columns]
        st.dataframe(display_df[cols], use_container_width=True, hide_index=True)


# ======================================
# TAB 3 — KALENDER
# ======================================
with tab_calendar:
    if filtered_ev.empty:
        st.info("Sitzungsdaten-Datei hochladen für die Kalenderansicht.")
    else:
        cmap = make_color_map(list(filtered_ev["Event"].unique()))
        conflict_dates = set(c["Datum"] for c in rc + pc)
        months_to_show = range(1,13) if selected_month_num == 0 else [selected_month_num]
        today_date = datetime.date.today()
        WDAYS = ["Mo","Di","Mi","Do","Fr","Sa","So"]

        for m in months_to_show:
            month_df = filtered_ev[filtered_ev["Month"] == m]
            if month_df.empty: continue
            mname = datetime.date(1900,m,1).strftime("%B")
            html  = f'<div class="cal-wrap"><div class="cal-title">{mname} 2026</div><div class="cal-grid">'
            for d in WDAYS:
                html += f'<div class="cal-hdr">{d}</div>'
            for week in calendar.monthcalendar(2026, m):
                for di, day in enumerate(week):
                    if day == 0:
                        html += '<div class="cal-cell-empty"></div>'; continue
                    is_today   = (datetime.date(2026,m,day) == today_date)
                    is_wknd    = (di >= 5)
                    ds         = datetime.date(2026,m,day).strftime("%d.%m.%Y")
                    has_conf   = ds in conflict_dates
                    cls = "cal-cell"
                    if is_today:  cls += " cal-cell-today"
                    elif is_wknd: cls += " cal-cell-wknd"
                    if has_conf:  cls += " cal-cell-conflict"
                    day_df = month_df[month_df["Day"] == day].copy()
                    day_df["_s"] = day_df["Zeit"].str[:5]
                    day_df = day_df.sort_values("_s")
                    dot = '<div class="cal-conflict-dot"></div>' if has_conf else ''
                    html += f'<div class="{cls}">{dot}<div class="cal-daynum">{day:02d}</div>'
                    for idx, (_,row) in enumerate(day_df.iterrows()):
                        if idx == 3:
                            html += f'<div class="cal-more">+{len(day_df)-3} weitere</div>'; break
                        color = cmap.get(row["Event"], "#CCCCCC")
                        hx = color.lstrip("#")
                        r2,g2,b2 = int(hx[0:2],16),int(hx[2:4],16),int(hx[4:6],16)
                        txt = "#000" if (r2*299+g2*587+b2*114)/1000 > 140 else "#fff"
                        html += (f'<div class="cal-ev" style="background:{color};color:{txt}">'
                                 f'<span class="cal-ev-time">{str(row["Zeit"])[:11]}</span>'
                                 f'<span class="cal-ev-name">{str(row["Event"])[:40]}</span></div>')
                    html += '</div>'
            html += '</div></div>'
            st.markdown(html, unsafe_allow_html=True)

            with st.expander(f"Tagesdetails {mname}", expanded=False):
                for day in range(1,32):
                    try: datetime.date(2026,m,day)
                    except ValueError: break
                    day_df = month_df[month_df["Day"]==day]
                    if day_df.empty: continue
                    day_df = day_df.copy()
                    day_df["_s"] = day_df["Zeit"].str[:5]
                    day_df = day_df.sort_values("_s")
                    ds = datetime.date(2026,m,day).strftime("%d.%m.%Y")
                    wname = datetime.date(2026,m,day).strftime("%a")
                    conf_ind = " ⚠️" if ds in conflict_dates else ""
                    st.markdown(f"**{wname}, {ds}** — {len(day_df)} Event(s){conf_ind}")
                    for _,row in day_df.iterrows():
                        color = cmap.get(row["Event"], "#CCCCCC")
                        bem = str(row.get("Bemerkungen","")) if pd.notna(row.get("Bemerkungen","")) else ""
                        st.markdown(
                            f'<div style="border-left:4px solid {color};padding:6px 12px;margin-bottom:6px;'
                            f'background:#fafbfc;border-radius:0 6px 6px 0;font-size:12px">'
                            f'<b style="font-family:DM Mono,monospace">{row["Zeit"]}</b> — <b>{row["Event"]}</b><br>'
                            f'<span style="color:#64748b">👤 {row["Personen"]}</span><br>'
                            f'<span style="color:#64748b">📍 {row["Ort"]}</span>'
                            + (f'<br><span style="color:#94a3b8;font-size:11px">{bem}</span>' if bem else '')
                            + '</div>',
                            unsafe_allow_html=True
                        )
            st.markdown("---")


# ======================================
# TAB 4 — KONFLIKTE
# ======================================
with tab_conflicts:
    st.markdown('<div class="section-header">Konflikt-Checker</div>', unsafe_allow_html=True)
    if filtered_ev.empty:
        st.info("Sitzungsdaten-Datei hochladen für Konfliktprüfung.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"#### 🏢 Raum-Konflikte ({len(rc)})")
            if not rc:
                st.success("Keine Raum-Doppelbuchungen.")
            else:
                for c in rc:
                    st.markdown(
                        f'<div class="conflict-card room"><div class="conflict-title">'
                        f'{c["Raum"]} — {c["Datum"]}</div>'
                        f'<div class="conflict-detail">{c["Zeit 1"]} → {c["Event 1"]}<br>'
                        f'{c["Zeit 2"]} → {c["Event 2"]}</div></div>',
                        unsafe_allow_html=True
                    )
        with col2:
            st.markdown(f"#### 👤 Personen-Konflikte ({len(pc)})")
            if not pc:
                st.success("Keine Personen-Überlappungen.")
            else:
                for c in pc:
                    st.markdown(
                        f'<div class="conflict-card person"><div class="conflict-title">'
                        f'{c["Person"]} — {c["Datum"]}</div>'
                        f'<div class="conflict-detail">{c["Zeit 1"]} → {c["Event 1"]}<br>'
                        f'{c["Zeit 2"]} → {c["Event 2"]}</div></div>',
                        unsafe_allow_html=True
                    )
        if rc or pc:
            all_c = []
            for c in rc:
                all_c.append({"Typ":"Raum","Datum":c["Datum"],"Raum/Person":c["Raum"],
                               "Event 1":c["Event 1"],"Zeit 1":c["Zeit 1"],
                               "Event 2":c["Event 2"],"Zeit 2":c["Zeit 2"]})
            for c in pc:
                all_c.append({"Typ":"Person","Datum":c["Datum"],"Raum/Person":c["Person"],
                               "Event 1":c["Event 1"],"Zeit 1":c["Zeit 1"],
                               "Event 2":c["Event 2"],"Zeit 2":c["Zeit 2"]})
            buf = io.StringIO()
            pd.DataFrame(all_c).to_csv(buf, index=False)
            st.download_button("Konflikte als CSV", buf.getvalue(), "konflikte_2026.csv", "text/csv")


# ======================================
# TAB 5 — iCAL EXPORT
# ======================================
with tab_ical:
    st.markdown('<div class="section-header">iCal / Outlook Export</div>', unsafe_allow_html=True)
    if filtered_ev.empty:
        st.info("Sitzungsdaten-Datei hochladen für iCal-Export.")
    else:
        st.info("Tipp: Verwende die Filter in der Sidebar, um gezielt Events einer Person oder eines Raumes zu exportieren.")
        ca, cb = st.columns(2)
        with ca:
            st.markdown(f"#### Gefilterte Events ({len(filtered_ev)})")
            st.download_button(
                "Gefilterte Events als .ics",
                make_ical(filtered_ev).encode("utf-8"),
                "kim_filtered.ics", "text/calendar"
            )
            st.caption("Outlook: Datei → Öffnen & Exportieren → Importieren → iCalendar")
        with cb:
            st.markdown(f"#### Alle Events ({len(df_ev)})")
            st.download_button(
                "Alle Events als .ics",
                make_ical(df_ev).encode("utf-8"),
                "kim_alle_2026.ics", "text/calendar"
            )
        st.markdown("---")
        st.markdown("#### Einzelner Event-Typ")
        ev_type = st.selectbox("Typ wählen", sorted(df_ev["Event"].unique()))
        single_df = df_ev[df_ev["Event"]==ev_type]
        st.markdown(f"**{len(single_df)} Termine** für *{ev_type}*")
        st.download_button(
            f"'{ev_type[:30]}' als .ics",
            make_ical(single_df).encode("utf-8"),
            f"kim_{ev_type[:20].replace(' ','_')}.ics", "text/calendar"
        )


# ======================================
# TAB 6 — HEATMAP
# ======================================
with tab_heatmap:
    st.markdown('<div class="section-header">Event-Dichte 2026 — GitHub-Style Heatmap</div>', unsafe_allow_html=True)
    if filtered_ev.empty:
        # Use planning data for heatmap if no event list
        if plan_rows:
            day_counts = {}
            for row in plan_rows:
                for i, v in enumerate(row["values"]):
                    if v is not None and i < len(WEEKS):
                        d = datetime.datetime.strptime(WEEKS[i], "%Y-%m-%d").date()
                        day_counts[d] = day_counts.get(d, 0) + 1
            st.caption("Heatmap basiert auf der Jahresplanungstabelle (Wochen-Granularität).")
        else:
            st.info("Dateien hochladen."); st.stop()
    else:
        day_counts = filtered_ev.groupby(filtered_ev["Datum"].dt.date).size().to_dict()

    max_count = max(day_counts.values(), default=1)

    def hmap_color(n, mx):
        if n == 0: return "#f1f5f9"
        t = n/mx
        if t < 0.25:  return "#bfdbfe"
        elif t < 0.5: return "#60a5fa"
        elif t < 0.75:return "#2563eb"
        return "#1e3a8a"

    jan1  = datetime.date(2026, 1, 1)
    start = jan1 - datetime.timedelta(days=jan1.weekday())
    cells = '<div style="font-family:DM Mono,monospace">'
    cells += '<div style="display:flex;margin-left:36px;margin-bottom:4px">'
    shown_m = {}
    for w in range(53):
        d = start + datetime.timedelta(weeks=w)
        m = d.strftime("%b")
        if d.month not in shown_m:
            shown_m[d.month] = True
            cells += f'<div style="flex:1;font-size:9px;color:#94a3b8">{m}</div>'
        else:
            cells += '<div style="flex:1"></div>'
    cells += '</div><div style="display:flex;gap:2px">'
    cells += '<div style="display:flex;flex-direction:column;gap:2px;width:32px;flex-shrink:0">'
    for wd in ["Mo","Di","Mi","Do","Fr","Sa","So"]:
        cells += f'<div style="height:14px;line-height:14px;font-size:9px;color:#94a3b8;text-align:right;padding-right:4px">{wd}</div>'
    cells += '</div><div style="display:flex;gap:2px;flex:1">'
    for w in range(53):
        cells += '<div style="display:flex;flex-direction:column;gap:2px;flex:1">'
        for wd in range(7):
            d = start + datetime.timedelta(weeks=w, days=wd)
            if d.year != 2026:
                cells += '<div style="height:14px;border-radius:2px;background:#f8fafc"></div>'
            else:
                n = day_counts.get(d, 0)
                col = hmap_color(n, max_count)
                cells += f'<div title="{d.strftime("%d.%m.%Y")}: {n}" style="height:14px;border-radius:2px;background:{col};cursor:default"></div>'
        cells += '</div>'
    cells += '</div></div>'
    cells += '<div style="display:flex;align-items:center;gap:6px;margin-top:10px;font-size:10px;color:#94a3b8">'
    cells += '<span>Weniger</span>'
    for c in ["#f1f5f9","#bfdbfe","#60a5fa","#2563eb","#1e3a8a"]:
        cells += f'<div style="width:14px;height:14px;border-radius:2px;background:{c}"></div>'
    cells += '<span>Mehr</span></div></div>'
    st.markdown(cells, unsafe_allow_html=True)

    # Bubble chart
    if not filtered_ev.empty:
        import plotly.graph_objects as go
        st.markdown('<div class="section-header" style="margin-top:24px">Events nach Kategorie & Datum</div>', unsafe_allow_html=True)
        filtered_ev["Kategorie"] = filtered_ev["Event"].apply(get_category)
        bubble_df = filtered_ev.groupby(["Datum","Kategorie"]).size().reset_index(name="Anzahl")
        fig = go.Figure()
        for cat, color in BASE_COLORS.items():
            d = bubble_df[bubble_df["Kategorie"]==cat]
            if d.empty: continue
            fig.add_trace(go.Scatter(
                x=d["Datum"], y=d["Kategorie"], mode="markers",
                marker=dict(size=d["Anzahl"]*14, color=color, opacity=0.8,
                            line=dict(width=1, color="white"), sizemode="diameter"),
                name=cat,
                text=[f"{r['Datum'].strftime('%d.%m.%Y')}: {r['Anzahl']} Events" for _,r in d.iterrows()],
                hovertemplate="%{text}<extra></extra>",
            ))
        fig.update_layout(
            height=420, plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(range=["2026-01-01","2026-12-31"], tickformat="%b", dtick="M1", gridcolor="#f1f5f9"),
            yaxis=dict(gridcolor="#f1f5f9"),
            legend=dict(orientation="h", yanchor="bottom", y=-0.4, xanchor="center", x=0.5),
            margin=dict(l=10,r=10,t=10,b=10), paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig, use_container_width=True)


# ======================================
# TAB 7 — STATISTIK
# ======================================
with tab_bar:
    import plotly.graph_objects as go
    st.markdown('<div class="section-header">Statistik & Jahresübersicht</div>', unsafe_allow_html=True)

    # Planning matrix: events per month
    if plan_rows:
        st.markdown("##### Events pro Monat (Jahresplanung)")
        month_counts = {}
        for ms in MONTH_SPANS:
            cnt = sum(1 for r in plan_rows for i in range(ms["start"], ms["end"]+1)
                      if i < len(r["values"]) and r["values"][i] is not None)
            month_counts[ms["label"]] = cnt

        fig_m = go.Figure(go.Bar(
            x=list(month_counts.keys()), y=list(month_counts.values()),
            marker_color="#378ADD", opacity=0.85,
        ))
        fig_m.update_layout(
            height=280, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10,r=10,t=10,b=10),
            xaxis=dict(gridcolor="#f1f5f9"), yaxis=dict(title="Anzahl", gridcolor="#f1f5f9"),
        )
        st.plotly_chart(fig_m, use_container_width=True)

        st.markdown("##### Events pro Bereich")
        sec_counts = {}
        for r in plan_rows:
            cnt = sum(1 for v in r["values"] if v is not None)
            sec_counts[r["section"]] = sec_counts.get(r["section"], 0) + cnt

        fig_s = go.Figure(go.Bar(
            x=list(sec_counts.values()), y=list(sec_counts.keys()),
            orientation="h",
            marker_color=[CELL_COLORS.get(s,"#ccc") for s in sec_counts],
        ))
        fig_s.update_layout(
            height=320, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10,r=10,t=10,b=10),
            xaxis=dict(title="Anzahl Einträge", gridcolor="#f1f5f9"),
            yaxis=dict(gridcolor="#f1f5f9"),
        )
        st.plotly_chart(fig_s, use_container_width=True)

    # Event list: stacked bar by day
    if not filtered_ev.empty:
        st.markdown("##### Events aus Sitzungsdaten — Stapeldiagramm pro Tag")
        filtered_ev["Kategorie"] = filtered_ev["Event"].apply(get_category)
        full_range = pd.date_range("2026-01-01","2026-12-31")
        pivot = filtered_ev.groupby(["Datum","Kategorie"]).size().unstack(fill_value=0)
        pivot = pivot.reindex(full_range, fill_value=0)
        fig_bar_stacked = go.Figure()
        for cat in pivot.columns:
            fig_bar_stacked.add_trace(go.Bar(
                x=pivot.index, y=pivot[cat], name=cat,
                marker_color=BASE_COLORS.get(cat,"#B0BEC5")
            ))
        fig_bar_stacked.update_layout(
            barmode="stack", height=380, plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(tickformat="%b", dtick="M1", gridcolor="#f1f5f9"),
            yaxis=dict(title="Anzahl", gridcolor="#f1f5f9"),
            legend=dict(orientation="h", yanchor="bottom", y=-0.45, xanchor="center", x=0.5),
            margin=dict(l=10,r=10,t=10,b=20), paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_bar_stacked, use_container_width=True)
