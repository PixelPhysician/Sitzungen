import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import re
from google.oauth2 import service_account
import gspread

st.set_page_config(layout="wide", page_title="KIM Sitzungsplanung 2026")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
[data-testid="stSidebar"]{background:#0f1923 !important;border-right:1px solid #1e2d3d;}
[data-testid="stSidebar"] *{color:#c8d8e8 !important;}
[data-testid="stSidebar"] .stSelectbox label,[data-testid="stSidebar"] .stMultiSelect label{
  color:#7a9bb5 !important;font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;}
[data-testid="stSidebar"] [data-baseweb="select"]{background:#1a2a3a !important;border-color:#2a3f55 !important;}
.main .block-container{padding-top:1.5rem;padding-bottom:3rem;max-width:1700px;}
.dashboard-header{background:linear-gradient(135deg,#0f1923 0%,#1a2f47 50%,#0d2137 100%);
  border-radius:10px;padding:20px 28px;margin-bottom:18px;
  display:flex;align-items:center;justify-content:space-between;border:1px solid #1e3a55;}
.header-title{font-size:22px;font-weight:700;color:#f0f6ff;margin:0;}
.header-subtitle{font-size:12px;color:#6b92b5;margin-top:3px;}
.header-badge{background:rgba(59,130,246,.15);border:1px solid rgba(59,130,246,.3);
  color:#7eb8f7;padding:5px 12px;border-radius:16px;font-size:11px;font-weight:600;}
.metric-row{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px;}
.metric-card{background:#fff;border:1px solid #e8ecf0;border-radius:8px;padding:12px 16px;position:relative;overflow:hidden;}
.metric-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:3px 0 0 3px;}
.metric-card.blue::before{background:#3b82f6;}.metric-card.teal::before{background:#14b8a6;}.metric-card.amber::before{background:#f59e0b;}
.metric-label{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;margin-bottom:3px;}
.metric-value{font-size:24px;font-weight:700;color:#0f1923;font-family:'DM Mono',monospace;}
.metric-sub{font-size:10px;color:#94a3b8;margin-top:1px;}
/* Planning table */
.plan-wrap{overflow:auto;max-height:80vh;border:1px solid #e2e8f0;border-radius:8px;}
table.plan{border-collapse:collapse;font-size:11px;white-space:nowrap;}
table.plan th,table.plan td{border:.5px solid #e2e8f0;padding:2px 5px;}
table.plan thead th{background:#f8fafc;font-weight:600;text-align:center;position:sticky;z-index:10;}
table.plan thead tr:nth-child(1) th{top:0;z-index:12;}
table.plan thead tr:nth-child(2) th{top:22px;z-index:11;}
table.plan thead tr:nth-child(3) th{top:44px;z-index:10;}
table.plan tbody tr:hover{background:#f0f6ff !important;}
.th-month{background:#1e3a55 !important;color:#7eb8f7 !important;font-weight:700 !important;font-size:11px !important;}
.th-meta{background:#f1f5f9 !important;}
.th-c0{position:sticky !important;left:0px;  z-index:14 !important;min-width:34px;max-width:34px;}
.th-c1{position:sticky !important;left:34px; z-index:14 !important;min-width:40px;max-width:40px;}
.th-c2{position:sticky !important;left:74px; z-index:14 !important;min-width:38px;max-width:38px;}
.th-c3{position:sticky !important;left:112px;z-index:14 !important;min-width:58px;max-width:58px;}
.th-c4{position:sticky !important;left:170px;z-index:14 !important;min-width:50px;max-width:50px;}
.th-c5{position:sticky !important;left:220px;z-index:14 !important;min-width:50px;max-width:50px;}
.th-name{position:sticky !important;left:270px;z-index:14 !important;min-width:270px;text-align:left !important;background:#f1f5f9 !important;}
.td-c0{position:sticky;left:0px;  z-index:5;background:#f8fafc;min-width:34px; max-width:34px; text-align:center;font-size:10px;color:#64748b;}
.td-c1{position:sticky;left:34px; z-index:5;background:#f8fafc;min-width:40px; max-width:40px; text-align:center;font-size:10px;color:#64748b;}
.td-c2{position:sticky;left:74px; z-index:5;background:#f8fafc;min-width:38px; max-width:38px; text-align:center;font-size:10px;color:#64748b;}
.td-c3{position:sticky;left:112px;z-index:5;background:#f8fafc;min-width:58px; max-width:58px; text-align:center;font-size:10px;color:#64748b;}
.td-c4{position:sticky;left:170px;z-index:5;background:#f8fafc;min-width:50px; max-width:50px; text-align:center;font-size:10px;color:#475569;font-family:'DM Mono',monospace;}
.td-c5{position:sticky;left:220px;z-index:5;background:#f8fafc;min-width:50px; max-width:50px; text-align:center;font-size:10px;color:#475569;font-family:'DM Mono',monospace;}
.td-name{position:sticky;left:270px;z-index:5;background:#fff;
  min-width:270px;max-width:380px;white-space:normal;font-size:11px;
  border-right:2px solid #cbd5e1 !important;}
.cell-has{text-align:center;font-weight:600;font-family:'DM Mono',monospace;font-size:11px;}
.cell-empty{background:transparent;}
.section-row td{font-weight:700;font-size:11px;padding:3px 8px;}
.stTabs [data-baseweb="tab-list"]{gap:4px;background:#f1f5f9;border-radius:8px;padding:4px;}
.stTabs [data-baseweb="tab"]{border-radius:6px;font-weight:600;font-size:12px;padding:5px 14px;}
.stTabs [aria-selected="true"]{background:#fff !important;color:#0f1923 !important;box-shadow:0 1px 3px rgba(0,0,0,.1);}
.stDownloadButton button{background:#0f1923 !important;color:#7eb8f7 !important;
  border:1px solid #1e3a55 !important;border-radius:6px !important;font-weight:600 !important;font-size:11px !important;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="dashboard-header">
  <div>
    <div class="header-title">KIM — Sitzungsplanung 2026</div>
    <div class="header-subtitle">Klinik fur Intensivmedizin · Inselspital Bern</div>
  </div>
  <div class="header-badge">2026 PLANUNG</div>
</div>
""", unsafe_allow_html=True)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
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
WEEK_NUMS = list(range(1, 53))

MONTHS = {0:"Januar",5:"Februar",9:"März",13:"April",18:"Mai",22:"Juni",
          26:"Juli",31:"August",35:"September",39:"Oktober",44:"November",48:"Dezember"}
MONTH_SPANS = []
_mk = sorted(MONTHS.keys())
for _i, _k in enumerate(_mk):
    _end = _mk[_i+1]-1 if _i+1 < len(_mk) else 51
    MONTH_SPANS.append({"label": MONTHS[_k], "start": _k, "end": _end})
MONTH_LIST = ["Januar","Februar","März","April","Mai","Juni",
              "Juli","August","September","Oktober","November","Dezember"]

# Section colors keyed by whatever is in col D of the sheet.
# Currently letters A-H, will be replaced with real names later.
# Format: (header_bg, header_fg, cell_bg)
SECTION_COLORS = {
    "A": ("#dbeafe","#1e40af","#bfdbfe"),
    "B": ("#fef3c7","#92400e","#fde68a"),
    "C": ("#fee2e2","#991b1b","#fca5a5"),
    "D": ("#e0f2fe","#0c4a6e","#7dd3fc"),
    "E": ("#fce7f3","#831843","#f9a8d4"),
    "F": ("#dcfce7","#14532d","#86efac"),
    "G": ("#ede9fe","#4c1d95","#c4b5fd"),
    "H": ("#f3f4f6","#374151","#d1d5db"),
    "":  ("#f8fafc","#475569","#e2e8f0"),
}

# ── GOOGLE SHEETS CONNECTION ──────────────────────────────────────────────────
@st.cache_resource
def get_gspread_client():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=60, show_spinner=False)
def load_sheet(spreadsheet_url: str, sheet_name: str) -> list[list]:
    """Return all values from a sheet as a list of rows (list of lists)."""
    client = get_gspread_client()
    sh     = client.open_by_url(spreadsheet_url)
    ws     = sh.worksheet(sheet_name)
    return ws.get_all_values()


# ── PLANNING PARSER ───────────────────────────────────────────────────────────
def parse_planning(raw_rows: list[list]) -> list[dict]:
    """
    Parse the Jahresplanung sheet.

    Sheet layout (rows are 0-indexed here):
      Row 0: empty / "2026" / month names
      Row 1: week numbers 1-52
      Row 2: week start dates
      Row 3: column headers:
               A=Tag  B=FA/IB  C=Wer?  D=Bereich
               E=Zeit Start  F=Zeit Ende  G=Notizen  H=event name
               I..BH = 52 week value columns
      Row 4+: data rows
    """
    # Find header row: the one containing "Tag" in col A
    hdr_idx = None
    for i, row in enumerate(raw_rows):
        if row and str(row[0]).strip().lower() == "tag":
            hdr_idx = i
            break
    if hdr_idx is None:
        st.error("Headerzeile mit 'Tag' in Spalte A nicht gefunden.")
        return []

    data_start = hdr_idx + 1

    # Detect name column from header row
    # Expected: col 0=Tag,1=FA/IB,2=Wer?,3=Bereich,4=ZeitStart,5=ZeitEnd,6=Notizen,7=Name
    hdr = raw_rows[hdr_idx]
    name_col = 7  # default
    for ci, val in enumerate(hdr):
        v = str(val).strip().lower()
        # The name column is the last meta col before the week numbers start.
        # We detect it as: not a known meta field, appears after col 3.
        if ci >= 4 and v not in ("zeit start","zeit ende","notizen","von","bis","notiz","notes"):
            # Check if the NEXT columns look like numbers (week cols)
            if ci + 1 < len(hdr):
                nxt = str(hdr[ci+1]).strip()
                if nxt.isdigit() or nxt in ("1","2","3"):
                    name_col = ci
                    break

    week_start_col = name_col + 1

    rows = []
    for row in raw_rows[data_start:]:
        if not row or len(row) <= name_col:
            continue
        name = str(row[name_col]).strip()
        if not name:
            continue

        def g(ci):
            if ci < len(row): return str(row[ci]).strip()
            return ""

        bereich = g(3)

        # Parse 52 week values
        values = []
        for j in range(52):
            ci = week_start_col + j
            raw = g(ci)
            if raw in ("","0","0.0","-"):
                values.append(None)
            else:
                raw2 = raw.replace(",",".")
                try:
                    v = float(raw2)
                    values.append(v if v != 0 else None)
                except ValueError:
                    values.append(raw if raw else None)  # keep "X" etc.

        rows.append({
            "tag":        g(0),
            "fa_ib":      g(1),
            "wer":        g(2),
            "bereich":    bereich,
            "zeit_start": g(4),
            "zeit_end":   g(5),
            "notizen":    g(6),
            "name":       name,
            "values":     values,
        })
    return rows


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_date(d):
    if not d: return ""
    p = d.split("-"); return f"{p[2]}.{p[1]}.{p[0][2:]}"

def parse_time_range(s):
    m = re.match(r"(\d{1,2})[:\.](\d{2})\s*[-–]\s*(\d{1,2})[:\.](\d{2})", str(s).strip())
    if m:
        try: return (datetime.time(int(m.group(1)),int(m.group(2))),
                     datetime.time(int(m.group(3)),int(m.group(4))))
        except: pass
    return None

def times_overlap(t1, t2):
    if not t1 or not t2: return False
    return t1[0] < t2[1] and t2[0] < t1[1]


# ── PLANNING TABLE RENDERER ───────────────────────────────────────────────────
def render_table(plan_rows, vis_month=0, bereich_filter="", search=""):
    if vis_month > 0:
        ms    = next((s for s in MONTH_SPANS if s["label"]==MONTH_LIST[vis_month-1]), None)
        vis_w = list(range(ms["start"], ms["end"]+1)) if ms else list(range(52))
    else:
        vis_w = list(range(52))

    vis_rows = [r for r in plan_rows
                if (not bereich_filter or r["bereich"] == bereich_filter)
                and (not search or search.lower() in r["name"].lower())]

    vis_ms = []
    for ms in MONTH_SPANS:
        cols = [i for i in range(ms["start"], ms["end"]+1) if i in vis_w]
        if cols: vis_ms.append({"label": ms["label"], "cols": cols})

    html = '<div class="plan-wrap"><table class="plan"><thead><tr>'
    html += '<th class="th-c0 th-meta" rowspan="3">Tag</th>'
    html += '<th class="th-c1 th-meta" rowspan="3">FA/<br>IB</th>'
    html += '<th class="th-c2 th-meta" rowspan="3">Wer?</th>'
    html += '<th class="th-c3 th-meta" rowspan="3">Bereich</th>'
    html += '<th class="th-c4 th-meta" rowspan="3">von</th>'
    html += '<th class="th-c5 th-meta" rowspan="3">bis</th>'
    html += '<th class="th-name th-meta" rowspan="3">Bezeichnung</th>'
    for ms in vis_ms:
        html += f'<th class="th-month" colspan="{len(ms["cols"])}">{ms["label"]}</th>'
    html += "</tr><tr>"
    for wi in vis_w:
        html += f'<th class="th-meta" style="font-size:9px;color:#64748b">{WEEK_NUMS[wi]}</th>'
    html += "</tr><tr>"
    for wi in vis_w:
        html += f'<th class="th-meta" style="font-size:9px;color:#94a3b8;font-family:DM Mono,monospace">{fmt_date(WEEKS[wi])}</th>'
    html += "</tr></thead><tbody>"

    last_bereich = None
    for row in vis_rows:
        b = row["bereich"]
        # Section separator row when bereich changes
        if b != last_bereich:
            last_bereich = b
            bg, fg, _ = SECTION_COLORS.get(b, SECTION_COLORS[""])
            html += (f'<tr class="section-row"><td colspan="{7+len(vis_w)}" '
                     f'style="background:{bg};color:{fg};font-weight:700;padding:4px 8px">'
                     f'Bereich {b}</td></tr>') if b else ""

        _, _, cell_bg = SECTION_COLORS.get(b, SECTION_COLORS[""])

        html += ("<tr>"
                 f'<td class="td-c0">{row["tag"]}</td>'
                 f'<td class="td-c1">{row["fa_ib"]}</td>'
                 f'<td class="td-c2">{row["wer"]}</td>'
                 f'<td class="td-c3">{b}</td>'
                 f'<td class="td-c4">{row["zeit_start"]}</td>'
                 f'<td class="td-c5">{row["zeit_end"]}</td>'
                 f'<td class="td-name">{row["name"]}</td>')
        for wi in vis_w:
            v = row["values"][wi] if wi < len(row["values"]) else None
            if v is not None:
                disp = int(v) if isinstance(v,float) and v==int(v) else (round(v,1) if isinstance(v,float) else v)
                all_w = [f"KW{WEEK_NUMS[i]} {fmt_date(WEEKS[i])}"
                         for i,val in enumerate(row["values"]) if val is not None and i < 52]
                tip = f"{row['name']} | KW{WEEK_NUMS[wi]} ({fmt_date(WEEKS[wi])}): {disp} | Alle Termine: {', '.join(all_w)}"
                html += (f'<td class="cell-has" style="background:{cell_bg};color:#1e3a5f" '
                         f'title="{tip}">{disp}</td>')
            else:
                html += '<td class="cell-empty"></td>'
        html += "</tr>"

    html += "</tbody></table></div>"
    return html


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("""
<div style="padding:12px 0 8px 0;">
  <div style="font-size:9px;font-weight:700;letter-spacing:.15em;color:#4a6f8a;text-transform:uppercase;margin-bottom:2px;">KIM ICU BERN</div>
  <div style="font-size:15px;font-weight:700;color:#e8f4ff;margin-bottom:4px;">Sitzungsplanung</div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown('<div style="font-size:10px;color:#7a9bb5;font-weight:600;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;">Google Sheet</div>', unsafe_allow_html=True)

sheet_url = st.sidebar.text_input(
    "Spreadsheet URL",
    placeholder="https://docs.google.com/spreadsheets/d/...",
    help="Muss fuer das Service Account freigegeben sein",
)
planning_sheet = st.sidebar.text_input("Blattname Jahresplanung", value="Sitzungsdaten 2026")

if st.sidebar.button("Neu laden"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
month_sel = st.sidebar.selectbox("Monat", ["Ganzes Jahr"] + MONTH_LIST)
month_idx = MONTH_LIST.index(month_sel) + 1 if month_sel != "Ganzes Jahr" else 0

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
if not sheet_url.strip():
    st.info("Bitte die Google-Sheets-URL links eingeben.")
    st.stop()

with st.spinner("Lade Daten..."):
    try:
        raw = load_sheet(sheet_url.strip(), planning_sheet.strip())
    except Exception as e:
        st.error(f"Fehler beim Laden: {e}")
        st.stop()

plan_rows = parse_planning(raw)

if not plan_rows:
    st.warning("Keine Daten geladen. Blattname korrekt?")
    st.stop()

# Bereich filter (uses whatever letters/values are in col D)
all_bereiche = sorted({r["bereich"] for r in plan_rows if r["bereich"]})
bereich_sel  = st.sidebar.selectbox("Bereich", ["Alle"] + all_bereiche)
bereich_filter = "" if bereich_sel == "Alle" else bereich_sel

# ── METRICS ───────────────────────────────────────────────────────────────────
total      = sum(1 for r in plan_rows for v in r["values"] if v is not None)
n_events   = len(plan_rows)
busy_weeks = max((sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)), default=0)

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card blue">
    <div class="metric-label">Planungseintraege</div>
    <div class="metric-value">{total}</div>
    <div class="metric-sub">geplante Eintraege</div>
  </div>
  <div class="metric-card teal">
    <div class="metric-label">Ereignisarten</div>
    <div class="metric-value">{n_events}</div>
    <div class="metric-sub">verschiedene Ereignisse</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">Max. Ereignisse / Woche</div>
    <div class="metric-value">{busy_weeks}</div>
    <div class="metric-sub">intensivste Woche</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── MAIN TABLE ────────────────────────────────────────────────────────────────
c1, c2 = st.columns([3,1])
with c1:
    search_q = st.text_input("Suchen", placeholder="Ereignis suchen...", label_visibility="collapsed")
with c2:
    st.caption(f"{len(plan_rows)} Ereignisse · letzte Aktualisierung: {datetime.datetime.now().strftime('%H:%M')}")

st.markdown(
    render_table(plan_rows, vis_month=month_idx, bereich_filter=bereich_filter, search=search_q),
    unsafe_allow_html=True,
)

# Legend
if all_bereiche:
    st.markdown("<br>", unsafe_allow_html=True)
    leg_cols = st.columns(min(len(all_bereiche), 8))
    for col, b in zip(leg_cols, all_bereiche):
        _, fg, cell_bg = SECTION_COLORS.get(b, SECTION_COLORS[""])
        col.markdown(
            f'<div style="background:{cell_bg};color:{fg};padding:3px 8px;border-radius:4px;'
            f'font-size:10px;font-weight:600;text-align:center">Bereich {b}</div>',
            unsafe_allow_html=True,
        )
