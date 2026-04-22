import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys
import io
import re
from google.oauth2 import service_account
import gspread

st.set_page_config(layout="wide", page_title="KIM Sitzungsplanung 2026")

# ── CSS ───────────────────────────────────────────────────────────────────────
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
.metric-row{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:16px;}
.metric-card{background:#fff;border:1px solid #e8ecf0;border-radius:8px;padding:12px 16px;position:relative;overflow:hidden;}
.metric-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:3px 0 0 3px;}
.metric-card.blue::before{background:#3b82f6;}.metric-card.teal::before{background:#14b8a6;}
.metric-card.amber::before{background:#f59e0b;}.metric-card.rose::before{background:#f43f5e;}
.metric-label{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;margin-bottom:3px;}
.metric-value{font-size:24px;font-weight:700;color:#0f1923;font-family:'DM Mono',monospace;}
.metric-sub{font-size:10px;color:#94a3b8;margin-top:1px;}
.plan-wrap{overflow:auto;max-height:78vh;border:1px solid #e2e8f0;border-radius:8px;}
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
.th-c3{position:sticky !important;left:112px;z-index:14 !important;min-width:56px;max-width:56px;}
.th-c4{position:sticky !important;left:168px;z-index:14 !important;min-width:50px;max-width:50px;}
.th-c5{position:sticky !important;left:218px;z-index:14 !important;min-width:50px;max-width:50px;}
.th-name{position:sticky !important;left:268px;z-index:14 !important;min-width:270px;text-align:left !important;background:#f1f5f9 !important;}
.td-c0{position:sticky;left:0px;  z-index:5;background:#f8fafc;min-width:34px; max-width:34px; text-align:center;font-size:10px;color:#64748b;}
.td-c1{position:sticky;left:34px; z-index:5;background:#f8fafc;min-width:40px; max-width:40px; text-align:center;font-size:10px;color:#64748b;}
.td-c2{position:sticky;left:74px; z-index:5;background:#f8fafc;min-width:38px; max-width:38px; text-align:center;font-size:10px;color:#64748b;}
.td-c3{position:sticky;left:112px;z-index:5;background:#f8fafc;min-width:56px; max-width:56px; text-align:center;font-size:10px;color:#64748b;}
.td-c4{position:sticky;left:168px;z-index:5;background:#f8fafc;min-width:50px; max-width:50px; text-align:center;font-size:10px;color:#475569;font-family:'DM Mono',monospace;}
.td-c5{position:sticky;left:218px;z-index:5;background:#f8fafc;min-width:50px; max-width:50px; text-align:center;font-size:10px;color:#475569;font-family:'DM Mono',monospace;}
.td-name{position:sticky;left:268px;z-index:5;background:#fff;
  min-width:270px;max-width:380px;white-space:normal;font-size:11px;
  border-right:2px solid #cbd5e1 !important;}
.cell-has{text-align:center;font-weight:600;font-family:'DM Mono',monospace;font-size:11px;}
.cell-empty{background:transparent;}
.section-row td{font-weight:700;font-size:11px;padding:3px 8px;}
.cf-card{border-radius:6px;padding:8px 12px;margin-bottom:6px;font-size:12px;}
.cf-room{background:#fff5f5;border:1px solid #fecaca;border-left:4px solid #ef4444;}
.cf-title{font-weight:700;font-size:12px;margin-bottom:2px;}
.cf-detail{color:#64748b;font-size:11px;font-family:'DM Mono',monospace;}
.cal-wrap{margin-bottom:24px;}
.cal-title{font-size:15px;font-weight:700;margin-bottom:7px;color:#0f1923;}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;}
.cal-hdr{text-align:center;font-weight:700;font-size:10px;padding:4px 0;color:#94a3b8;background:#f8fafc;border-radius:4px 4px 0 0;}
.cal-cell{border:1px solid #e8ecf2;border-radius:5px;padding:4px;min-height:78px;background:#fff;box-sizing:border-box;position:relative;}
.cal-cell-wknd{background:#fafbfc;}.cal-cell-today{border:2px solid #3b82f6;background:#eff6ff;}
.cal-cell-empty{border:none;background:transparent;min-height:78px;}
.cal-daynum{font-weight:700;font-size:10px;color:#475569;margin-bottom:2px;font-family:'DM Mono',monospace;}
.cal-ev{padding:2px 4px;margin-top:2px;border-radius:2px;font-size:9px;line-height:1.3;word-break:break-word;}
.cal-ev-time{font-weight:700;display:block;font-family:'DM Mono',monospace;}
.cal-ev-name{display:block;font-size:8.5px;opacity:.9;}
.cal-more{font-size:9px;color:#94a3b8;font-style:italic;margin-top:2px;}
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

WEEKDAY_OFFSETS = {
    "mo": 0, "di": 1, "mi": 2, "do": 3, "fr": 4, "sa": 5, "so": 6,
    "montag": 0, "dienstag": 1, "mittwoch": 2, "donnerstag": 3,
    "freitag": 4, "samstag": 5, "sonntag": 6,
}

MONTHS = {0:"Januar",5:"Februar",9:"März",13:"April",18:"Mai",22:"Juni",
          26:"Juli",31:"August",35:"September",39:"Oktober",44:"November",48:"Dezember"}
MONTH_SPANS = []
_mk = sorted(MONTHS.keys())
for _i,_k in enumerate(_mk):
    _end = _mk[_i+1]-1 if _i+1<len(_mk) else 51
    MONTH_SPANS.append({"label":MONTHS[_k],"start":_k,"end":_end})
MONTH_LIST = ["Januar","Februar","März","April","Mai","Juni",
              "Juli","August","September","Oktober","November","Dezember"]

BEREICH_COLORS = {
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

BASE_COLORS = {
    "EPIC":"#4FC3F7","Geräte & Beatmung":"#E57373","Workshop":"#FFD54F",
    "Schulung/Kurs":"#81C784","Sitzung":"#FF8A65","Einführung":"#4DB6AC",
    "Lernwerkstatt":"#F48FB1","Führung/Austausch":"#CE93D8","ICU":"#90CAF9",
    "Kommunikation":"#FFCC80","Planung":"#A5D6A7","Sonstiges":"#B0BEC5",
}

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
def sanitise_url(url: str) -> str:
    url = url.strip()
    url = re.sub(r"/u/\d+/", "/", url)
    m = re.match(r"(https://docs\.google\.com/spreadsheets/d/[^/?#]+)", url)
    return m.group(1) if m else url

@st.cache_resource
def get_client():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=120, show_spinner=False)
def fetch_first_sheet(spreadsheet_url: str) -> tuple[str, list]:
    """Open the spreadsheet and return (sheet_title, all_values) for the first worksheet."""
    gc = get_client()
    wb = gc.open_by_url(spreadsheet_url)
    ws = wb.worksheets()[0]          # always use first sheet
    return ws.title, ws.get_all_values()

# ── PLANNING PARSER ───────────────────────────────────────────────────────────
def parse_planning(raw: list) -> list:
    """
    Actual sheet layout (from screenshot):
      Row 4  = header:
        Col A=Tag  B=FA IB  C=Wer?  D=Bereich  E=Personen  F=Raum
        G=Notizen  H=WB relevant?  I=Zeit Start  J=Zeit Ende
        K=ganze Tage Pflegende IB  (= EVENT NAME)
        L onward = week columns (week 1, 2, 3 …)

    We find the header row by scanning for a row whose col A is "Tag" (case-insensitive).
    Then we scan rightward from col K to find where the week numbers start,
    identifying the name column as the last non-numeric header before the weeks.
    """
    hdr_idx = None
    for i, row in enumerate(raw):
        if row and str(row[0]).strip().lower() == "tag":
            hdr_idx = i
            break

    if hdr_idx is None:
        # Fallback: scan every row for "tag" anywhere in first 3 cols
        for i, row in enumerate(raw):
            for ci in range(min(3, len(row))):
                if str(row[ci]).strip().lower() == "tag":
                    hdr_idx = i
                    break
            if hdr_idx is not None:
                break

    if hdr_idx is None:
        return []

    hdr = raw[hdr_idx]

    # Fixed columns from screenshot:
    #  0=Tag 1=FA_IB 2=Wer? 3=Bereich 4=Personen 5=Raum 6=Notizen
    #  7=WB_relevant 8=ZeitStart 9=ZeitEnd 10=EventName  11+=weeks
    # But we auto-detect name_col and week_start_col to be robust.

    # Find week_start_col: first col (after col 4) whose header is a small integer (1-52)
    week_start_col = None
    for ci in range(4, len(hdr)):
        v = str(hdr[ci]).strip()
        if v.isdigit() and 1 <= int(v) <= 52:
            week_start_col = ci
            break

    if week_start_col is None:
        return []

    # name_col = the last non-empty, non-integer column before week_start_col
    name_col = week_start_col - 1
    for ci in range(week_start_col - 1, 3, -1):
        v = str(hdr[ci]).strip()
        if v and not v.isdigit():
            name_col = ci
            break

    # Derive meta column indices relative to fixed sheet layout
    # We use whatever is actually in cols 0-9; if the sheet has a different order
    # we still show what's there. Map by position:
    COL_TAG        = 0
    COL_FA_IB      = 1
    COL_WER        = 2
    COL_BEREICH    = 3
    COL_ZEIT_START = 8   # col I
    COL_ZEIT_END   = 9   # col J
    # name_col is col K (index 10) in the real sheet

    def g(row, ci):
        if ci < len(row):
            v = str(row[ci]).strip()
            return "" if v.upper() in ("NAN","NONE","#REF!","#N/A","#VALUE!","#NAME?") else v
        return ""

    rows = []
    for row in raw[hdr_idx + 1:]:
        if len(row) <= name_col:
            continue
        name = str(row[name_col]).strip()
        if not name or name.upper() in ("NAN", "NONE", ""):
            continue

        # Parse up to 52 week values
        values = []
        for j in range(52):
            ci = week_start_col + j
            raw_val = g(row, ci)
            if not raw_val or raw_val in ("0", "0.0", "-"):
                values.append(None)
            else:
                try:
                    v = float(raw_val.replace(",", "."))
                    values.append(v if v != 0 else None)
                except ValueError:
                    values.append(raw_val)   # keep "X" markers

        rows.append({
            "tag":        g(row, COL_TAG),
            "fa_ib":      g(row, COL_FA_IB),
            "wer":        g(row, COL_WER),
            "bereich":    g(row, COL_BEREICH).upper(),
            "zeit_start": g(row, COL_ZEIT_START),
            "zeit_end":   g(row, COL_ZEIT_END),
            "notizen":    "",   # col G not critical for display
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

def times_overlap(t1,t2):
    if not t1 or not t2: return False
    return t1[0]<t2[1] and t2[0]<t1[1]

def get_category(event):
    e = str(event).lower()
    if "epic" in e:                                                           return "EPIC"
    if any(x in e for x in ["ecmo","lvad","impella","assist","prismax","beatmung"]): return "Geräte & Beatmung"
    if "workshop" in e:                                                       return "Workshop"
    if any(x in e for x in ["schulung","kurs","basiskurs","aufbaukurs","refresher"]): return "Schulung/Kurs"
    if any(x in e for x in ["sitzung","fachgruppe","fg ","superuser","fachforum"]):   return "Sitzung"
    if "einführung" in e or "einblick" in e:                                  return "Einführung"
    if any(x in e for x in ["lernwerkstatt","repe","simulation","kimsim"]):   return "Lernwerkstatt"
    if any(x in e for x in ["führungsdialog","austausch"]):                   return "Führung/Austausch"
    if "icu" in e:                                                            return "ICU"
    if any(x in e for x in ["kommunikation","aggressions"]):                  return "Kommunikation"
    if any(x in e for x in ["planung","bürotag"]):                            return "Planung"
    return "Sonstiges"

def make_ical(df_ev):
    import hashlib
    lines=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//KIM ICU//DE","CALSCALE:GREGORIAN",
           "METHOD:PUBLISH","X-WR-CALNAME:KIM ICU 2026","X-WR-TIMEZONE:Europe/Zurich"]
    for _,row in df_ev.iterrows():
        uid=hashlib.md5(f"{row['Datum']}{row['Event']}{row['Zeit']}".encode()).hexdigest()
        ds=row["Datum"].strftime("%Y%m%d"); t=parse_time_range(row["Zeit"])
        if t:
            dtstart=f"DTSTART;TZID=Europe/Zurich:{ds}T{t[0].strftime('%H%M%S')}"
            dtend  =f"DTEND;TZID=Europe/Zurich:{ds}T{t[1].strftime('%H%M%S')}"
        else:
            dtstart=f"DTSTART;VALUE=DATE:{ds}"; dtend=f"DTEND;VALUE=DATE:{ds}"
        summary=str(row["Event"]).replace(",","\\,").replace(";","\\;")
        desc=f"Person: {row['Personen']}\\nOrt: {row['Ort']}"
        lines+=["BEGIN:VEVENT",f"UID:{uid}@kim.insel.ch",
                f"DTSTAMP:{datetime.datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                dtstart,dtend,f"SUMMARY:{summary}",
                f"LOCATION:{str(row['Ort']).replace(',','\\,')}",
                f"DESCRIPTION:{desc}","END:VEVENT"]
    lines.append("END:VCALENDAR"); return "\r\n".join(lines)

# ── DATE RESOLVER ─────────────────────────────────────────────────────────────
def resolve_event_date(row: dict, week_idx: int) -> datetime.date:
    tag         = str(row.get("tag", "")).strip()
    week_monday = datetime.datetime.strptime(WEEKS[week_idx], "%Y-%m-%d").date()

    # 1. Explicit date DD.MM. / DD.MM.YY / DD.MM.YYYY
    m = re.search(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?", tag)
    if m:
        try:
            day=int(m.group(1)); month=int(m.group(2))
            yr=m.group(3); year=2026
            if yr: year=int(yr) if len(yr)==4 else 2000+int(yr)
            return datetime.date(year, month, day)
        except ValueError:
            pass

    # 2. German weekday abbreviation (whole-string match)
    if tag.lower().strip() in WEEKDAY_OFFSETS:
        return week_monday + datetime.timedelta(days=WEEKDAY_OFFSETS[tag.lower().strip()])

    return week_monday

# ── PLANNING TABLE ────────────────────────────────────────────────────────────
def render_table(plan_rows, vis_month=0, bereich_filter="", search=""):
    if vis_month > 0:
        ms = next((s for s in MONTH_SPANS if s["label"]==MONTH_LIST[vis_month-1]), None)
        vis_w = list(range(ms["start"],ms["end"]+1)) if ms else list(range(52))
    else:
        vis_w = list(range(52))

    vis_rows = [r for r in plan_rows
                if (not bereich_filter or r["bereich"]==bereich_filter)
                and (not search or search.lower() in r["name"].lower())]

    vis_ms = []
    for ms in MONTH_SPANS:
        cols=[i for i in range(ms["start"],ms["end"]+1) if i in vis_w]
        if cols: vis_ms.append({"label":ms["label"],"cols":cols})

    html='<div class="plan-wrap"><table class="plan"><thead><tr>'
    html+='<th class="th-c0 th-meta" rowspan="3">Tag</th>'
    html+='<th class="th-c1 th-meta" rowspan="3">FA/<br>IB</th>'
    html+='<th class="th-c2 th-meta" rowspan="3">Wer?</th>'
    html+='<th class="th-c3 th-meta" rowspan="3">Bereich</th>'
    html+='<th class="th-c4 th-meta" rowspan="3">von</th>'
    html+='<th class="th-c5 th-meta" rowspan="3">bis</th>'
    html+='<th class="th-name th-meta" rowspan="3">Bezeichnung</th>'
    for ms in vis_ms: html+=f'<th class="th-month" colspan="{len(ms["cols"])}">{ms["label"]}</th>'
    html+="</tr><tr>"
    for wi in vis_w: html+=f'<th class="th-meta" style="font-size:9px;color:#64748b">{WEEK_NUMS[wi]}</th>'
    html+="</tr><tr>"
    for wi in vis_w: html+=f'<th class="th-meta" style="font-size:9px;color:#94a3b8;font-family:DM Mono,monospace">{fmt_date(WEEKS[wi])}</th>'
    html+="</tr></thead><tbody>"

    last_b = None
    for row in vis_rows:
        b = row["bereich"]
        if b != last_b:
            last_b = b
            bg,fg,_ = BEREICH_COLORS.get(b, BEREICH_COLORS[""])
            if b:
                html+=(f'<tr class="section-row"><td colspan="{7+len(vis_w)}" '
                       f'style="background:{bg};color:{fg}">Bereich {b}</td></tr>')
        _,_,cell_bg = BEREICH_COLORS.get(b, BEREICH_COLORS[""])
        html+=("<tr>"
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
                         for i,val in enumerate(row["values"]) if val is not None and i<52]
                tip = f"{row['name']} | KW{WEEK_NUMS[wi]} ({fmt_date(WEEKS[wi])}): {disp}"
                html+=f'<td class="cell-has" style="background:{cell_bg};color:#1e3a5f" title="{tip}">{disp}</td>'
            else:
                html+='<td class="cell-empty"></td>'
        html+="</tr>"
    html+="</tbody></table></div>"
    return html

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("""
<div style="padding:12px 0 8px 0;">
  <div style="font-size:9px;font-weight:700;letter-spacing:.15em;color:#4a6f8a;text-transform:uppercase;margin-bottom:2px;">KIM ICU BERN</div>
  <div style="font-size:15px;font-weight:700;color:#e8f4ff;margin-bottom:4px;">Sitzungsplanung</div>
</div>
""", unsafe_allow_html=True)
st.sidebar.markdown("---")

if st.sidebar.button("🔄 Neu laden"):
    st.cache_data.clear(); st.rerun()

st.sidebar.markdown("---")
month_sel  = st.sidebar.selectbox("Monat", ["Ganzes Jahr"] + MONTH_LIST)
month_idx  = MONTH_LIST.index(month_sel) + 1 if month_sel != "Ganzes Jahr" else 0

# ── LOAD DATA (first sheet, URL from secrets) ─────────────────────────────────
raw_url = st.secrets.get("sheet_url", "")
if not raw_url:
    st.error("Keine `sheet_url` in secrets.toml gefunden.")
    st.stop()

sheet_url = sanitise_url(raw_url)

with st.spinner("Lade Sitzungsplanung..."):
    try:
        sheet_title, raw = fetch_first_sheet(sheet_url)
    except Exception as e:
        st.error(f"Verbindung fehlgeschlagen: {e}")
        st.stop()

plan_rows = parse_planning(raw)

# ── DEBUG expander: show if parse failed ──────────────────────────────────────
if not plan_rows:
    with st.expander("🔍 Debug: Rohdaten (erste 8 Zeilen)", expanded=True):
        st.caption(f"Blatt geladen: **{sheet_title}** · {len(raw)} Zeilen")
        for i, r in enumerate(raw[:8]):
            st.text(f"Zeile {i+1}: {r[:15]}")
    st.warning("Parser hat keine Ereignisse gefunden. Bitte Debug-Info oben prüfen.")
    st.stop()

all_bereiche   = sorted({r["bereich"] for r in plan_rows if r["bereich"]})
bereich_sel    = st.sidebar.selectbox("Bereich", ["Alle"] + all_bereiche)
bereich_filter = "" if bereich_sel == "Alle" else bereich_sel

# ── METRICS ───────────────────────────────────────────────────────────────────
total = sum(1 for r in plan_rows for v in r["values"] if v is not None)
busy  = max((sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)), default=0)

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card blue">
    <div class="metric-label">Planungseintraege</div>
    <div class="metric-value">{total}</div>
    <div class="metric-sub">geplante Eintraege 2026</div>
  </div>
  <div class="metric-card teal">
    <div class="metric-label">Ereignisarten</div>
    <div class="metric-value">{len(plan_rows)}</div>
    <div class="metric-sub">verschiedene Ereignisse</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">Max. pro Woche</div>
    <div class="metric-value">{busy}</div>
    <div class="metric-sub">intensivste Woche</div>
  </div>
  <div class="metric-card rose">
    <div class="metric-label">Letzte Aktualisierung</div>
    <div class="metric-value" style="font-size:16px">{datetime.datetime.now().strftime('%H:%M')}</div>
    <div class="metric-sub">{sheet_title}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_plan, tab_cal, tab_conf, tab_ical, tab_heat, tab_stat = st.tabs([
    "Jahresplanung","Kalender","Konflikte","iCal Export","Heatmap","Statistik"
])

# ── TAB 1: JAHRESPLANUNG ──────────────────────────────────────────────────────
with tab_plan:
    c1, c2 = st.columns([3,1])
    with c1:
        search_q = st.text_input("Suchen", placeholder="Ereignis suchen...", label_visibility="collapsed")
    with c2:
        st.caption(f"{len(plan_rows)} Ereignisse · 52 Wochen · {sheet_title}")
    st.markdown(render_table(plan_rows, vis_month=month_idx, bereich_filter=bereich_filter, search=search_q),
                unsafe_allow_html=True)
    if all_bereiche:
        st.markdown("<br>", unsafe_allow_html=True)
        leg_cols = st.columns(min(len(all_bereiche), 8))
        for col, b in zip(leg_cols, all_bereiche):
            _,fg,cell_bg = BEREICH_COLORS.get(b, BEREICH_COLORS[""])
            col.markdown(f'<div style="background:{cell_bg};color:{fg};padding:3px 8px;border-radius:4px;'
                         f'font-size:10px;font-weight:600;text-align:center">Bereich {b}</div>',
                         unsafe_allow_html=True)

# ── TAB 2: KALENDER ───────────────────────────────────────────────────────────
with tab_cal:
    cal_events = []
    for row in plan_rows:
        for i, v in enumerate(row["values"]):
            if v is not None and i < len(WEEKS):
                actual_date = resolve_event_date(row, i)
                cal_events.append({
                    "date": actual_date, "name": row["name"],
                    "zeit_start": row["zeit_start"], "zeit_end": row["zeit_end"],
                    "bereich": row["bereich"],
                    "month": actual_date.month, "day": actual_date.day,
                })
    if not cal_events:
        st.info("Keine Daten fuer Kalenderansicht.")
    else:
        cal_df      = pd.DataFrame(cal_events)
        cmap2       = {b: BEREICH_COLORS.get(b, BEREICH_COLORS[""])[2] for b in all_bereiche}
        months_show = range(1,13) if month_idx == 0 else [month_idx]
        today_d     = datetime.date.today()
        WDAYS       = ["Mo","Di","Mi","Do","Fr","Sa","So"]
        for m in months_show:
            mdf = cal_df[cal_df["month"] == m]
            if mdf.empty: continue
            mname = datetime.date(1900,m,1).strftime("%B")
            html  = f'<div class="cal-wrap"><div class="cal-title">{mname} 2026</div><div class="cal-grid">'
            for d in WDAYS: html += f'<div class="cal-hdr">{d}</div>'
            for week in calendar.monthcalendar(2026, m):
                for di, day in enumerate(week):
                    if day == 0: html += '<div class="cal-cell-empty"></div>'; continue
                    is_today = (datetime.date(2026,m,day) == today_d)
                    is_wknd  = (di >= 5)
                    cls = "cal-cell" + (" cal-cell-today" if is_today else " cal-cell-wknd" if is_wknd else "")
                    ddf = mdf[mdf["day"] == day]
                    html += f'<div class="{cls}"><div class="cal-daynum">{day:02d}</div>'
                    for idx, (_, ev_row) in enumerate(ddf.iterrows()):
                        if idx == 3: html += f'<div class="cal-more">+{len(ddf)-3} weitere</div>'; break
                        color    = cmap2.get(ev_row["bereich"], "#e2e8f0")
                        time_str = f'{ev_row["zeit_start"]}-{ev_row["zeit_end"]}' if ev_row["zeit_start"] else ""
                        html += (f'<div class="cal-ev" style="background:{color};color:#1e3a5f">'
                                 f'<span class="cal-ev-time">{time_str}</span>'
                                 f'<span class="cal-ev-name">{str(ev_row["name"])[:38]}</span></div>')
                    html += '</div>'
            html += '</div></div>'
            st.markdown(html, unsafe_allow_html=True)

# ── TAB 3: KONFLIKTE ──────────────────────────────────────────────────────────
with tab_conf:
    st.markdown("#### Zeitkonflikte")
    st.caption("Tage, an denen mehrere Ereignisse zur gleichen Zeit geplant sind.")
    conflicts = []
    for wi in range(52):
        week_rows = [r for r in plan_rows if wi < len(r["values"]) and r["values"][wi] is not None]
        if len(week_rows) < 2: continue
        for i in range(len(week_rows)):
            for j in range(i+1, len(week_rows)):
                r1, r2  = week_rows[i], week_rows[j]
                t1_str  = f"{r1['zeit_start']}-{r1['zeit_end']}" if r1['zeit_start'] and r1['zeit_end'] else ""
                t2_str  = f"{r2['zeit_start']}-{r2['zeit_end']}" if r2['zeit_start'] and r2['zeit_end'] else ""
                t1 = parse_time_range(t1_str); t2 = parse_time_range(t2_str)
                if times_overlap(t1, t2):
                    d1 = resolve_event_date(r1, wi); d2 = resolve_event_date(r2, wi)
                    if d1 == d2:
                        conflicts.append({"KW": WEEK_NUMS[wi], "Datum": d1.strftime("%d.%m.%y"),
                                          "Ereignis 1": r1["name"], "Zeit 1": t1_str,
                                          "Ereignis 2": r2["name"], "Zeit 2": t2_str})
    if not conflicts:
        st.success("Keine Zeitkonflikte gefunden.")
    else:
        st.warning(f"{len(conflicts)} Konflikt(e) gefunden")
        for c in conflicts:
            st.markdown(
                f'<div class="cf-card cf-room"><div class="cf-title">KW{c["KW"]} — {c["Datum"]}</div>'
                f'<div class="cf-detail">{c["Zeit 1"]} → {c["Ereignis 1"]}<br>'
                f'{c["Zeit 2"]} → {c["Ereignis 2"]}</div></div>', unsafe_allow_html=True)
        buf = io.StringIO(); pd.DataFrame(conflicts).to_csv(buf, index=False)
        st.download_button("Konflikte als CSV", buf.getvalue(), "konflikte_2026.csv", "text/csv")

# ── TAB 4: iCAL ───────────────────────────────────────────────────────────────
with tab_ical:
    st.markdown("#### iCal Export")
    ical_rows = []
    for row in plan_rows:
        for i, v in enumerate(row["values"]):
            if v is not None and i < len(WEEKS):
                actual_date = resolve_event_date(row, i)
                zeit = f"{row['zeit_start']}-{row['zeit_end']}" if row["zeit_start"] and row["zeit_end"] else ""
                ical_rows.append({"Datum": pd.Timestamp(actual_date), "Event": row["name"],
                                  "Zeit": zeit, "Ort": "", "Personen": row["wer"] or "",
                                  "Bemerkungen": row["notizen"] or ""})
    if not ical_rows:
        st.info("Keine Daten verfugbar.")
    else:
        ical_df = pd.DataFrame(ical_rows)
        if month_idx > 0: ical_df = ical_df[ical_df["Datum"].dt.month == month_idx]
        if bereich_filter:
            n_b = {r["name"] for r in plan_rows if r["bereich"] == bereich_filter}
            ical_df = ical_df[ical_df["Event"].isin(n_b)]
        ca, cb = st.columns(2)
        with ca:
            st.markdown(f"**{len(ical_df)} Termine** (gefiltert)")
            st.download_button("Gefiltert als .ics", make_ical(ical_df).encode("utf-8"), "kim_filtered.ics","text/calendar")
        with cb:
            all_ical = pd.DataFrame(ical_rows)
            st.markdown(f"**{len(all_ical)} Termine** (alles)")
            st.download_button("Alle als .ics", make_ical(all_ical).encode("utf-8"), "kim_alle_2026.ics","text/calendar")
        st.caption("Outlook: Datei > Oeffnen & Exportieren > Importieren > iCalendar (.ics)")

# ── TAB 5: HEATMAP ────────────────────────────────────────────────────────────
with tab_heat:
    st.markdown("#### Event-Dichte 2026")
    day_counts = {}
    for row in plan_rows:
        for i, v in enumerate(row["values"]):
            if v is not None and i < len(WEEKS):
                d = resolve_event_date(row, i)
                day_counts[d] = day_counts.get(d, 0) + 1
    mx = max(day_counts.values(), default=1)
    def hc(n,mx):
        if n==0: return "#f1f5f9"
        t=n/mx
        if t<0.25: return "#bfdbfe"
        elif t<0.5: return "#60a5fa"
        elif t<0.75: return "#2563eb"
        return "#1e3a8a"
    jan1=datetime.date(2026,1,1); start=jan1-datetime.timedelta(days=jan1.weekday())
    cells='<div style="font-family:DM Mono,monospace"><div style="display:flex;margin-left:36px;margin-bottom:4px">'
    shown_m={}
    for w in range(53):
        d=start+datetime.timedelta(weeks=w); m=d.strftime("%b")
        if d.month not in shown_m: shown_m[d.month]=True; cells+=f'<div style="flex:1;font-size:9px;color:#94a3b8">{m}</div>'
        else: cells+='<div style="flex:1"></div>'
    cells+='</div><div style="display:flex;gap:2px"><div style="display:flex;flex-direction:column;gap:2px;width:32px;flex-shrink:0">'
    for wd in ["Mo","Di","Mi","Do","Fr","Sa","So"]:
        cells+=f'<div style="height:14px;line-height:14px;font-size:9px;color:#94a3b8;text-align:right;padding-right:4px">{wd}</div>'
    cells+='</div><div style="display:flex;gap:2px;flex:1">'
    for w in range(53):
        cells+='<div style="display:flex;flex-direction:column;gap:2px;flex:1">'
        for wd in range(7):
            d=start+datetime.timedelta(weeks=w,days=wd)
            if d.year!=2026: cells+='<div style="height:14px;border-radius:2px;background:#f8fafc"></div>'
            else:
                n=day_counts.get(d,0); col=hc(n,mx)
                cells+=f'<div title="{d.strftime("%d.%m.%Y")}: {n}" style="height:14px;border-radius:2px;background:{col}"></div>'
        cells+='</div>'
    cells+='</div></div><div style="display:flex;align-items:center;gap:6px;margin-top:10px;font-size:10px;color:#94a3b8"><span>Weniger</span>'
    for c in ["#f1f5f9","#bfdbfe","#60a5fa","#2563eb","#1e3a8a"]:
        cells+=f'<div style="width:14px;height:14px;border-radius:2px;background:{c}"></div>'
    cells+='<span>Mehr</span></div></div>'
    st.markdown(cells, unsafe_allow_html=True)

# ── TAB 6: STATISTIK ──────────────────────────────────────────────────────────
with tab_stat:
    import plotly.graph_objects as go
    st.markdown("#### Statistik")
    month_counts={ms["label"]:sum(1 for r in plan_rows
                  for i in range(ms["start"],ms["end"]+1)
                  if i<len(r["values"]) and r["values"][i] is not None)
                  for ms in MONTH_SPANS}
    fig_m=go.Figure(go.Bar(x=list(month_counts.keys()),y=list(month_counts.values()),marker_color="#378ADD",opacity=0.85))
    fig_m.update_layout(height=260,plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10,r=10,t=10,b=10),xaxis=dict(gridcolor="#f1f5f9"),yaxis=dict(title="Anzahl",gridcolor="#f1f5f9"))
    st.plotly_chart(fig_m,use_container_width=True)
    b_counts={}
    for r in plan_rows:
        b_counts[r["bereich"]]=b_counts.get(r["bereich"],0)+sum(1 for v in r["values"] if v is not None)
    b_colors=[BEREICH_COLORS.get(b,BEREICH_COLORS[""])[2] for b in b_counts]
    fig_b=go.Figure(go.Bar(x=list(b_counts.values()),y=[f"Bereich {b}" for b in b_counts.keys()],
        orientation="h",marker_color=b_colors))
    fig_b.update_layout(height=300,plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10,r=10,t=10,b=10),xaxis=dict(title="Anzahl",gridcolor="#f1f5f9"),yaxis=dict(gridcolor="#f1f5f9"))
    st.plotly_chart(fig_b,use_container_width=True)
