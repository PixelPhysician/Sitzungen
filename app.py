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
.cf-person{background:#faf5ff;border:1px solid #ddd6fe;border-left:4px solid #8b5cf6;}
.cf-title{font-weight:700;font-size:12px;margin-bottom:2px;}
.cf-detail{color:#64748b;font-size:11px;font-family:'DM Mono',monospace;}
.cal-wrap{margin-bottom:24px;}
.cal-title{font-size:15px;font-weight:700;margin-bottom:7px;color:#0f1923;}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;}
.cal-hdr{text-align:center;font-weight:700;font-size:10px;padding:4px 0;color:#94a3b8;background:#f8fafc;border-radius:4px 4px 0 0;}
.cal-cell{border:1px solid #e8ecf2;border-radius:5px;padding:4px;min-height:78px;background:#fff;box-sizing:border-box;position:relative;}
.cal-cell-wknd{background:#fafbfc;}.cal-cell-today{border:2px solid #3b82f6;background:#eff6ff;}
.cal-cell-conf{border:2px solid #f97316 !important;}.cal-cell-empty{border:none;background:transparent;min-height:78px;}
.cal-daynum{font-weight:700;font-size:10px;color:#475569;margin-bottom:2px;font-family:'DM Mono',monospace;}
.cal-conf-dot{position:absolute;top:3px;right:4px;width:5px;height:5px;background:#f97316;border-radius:50%;}
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

MONTHS = {0:"Januar",5:"Februar",9:"März",13:"April",18:"Mai",22:"Juni",
          26:"Juli",31:"August",35:"September",39:"Oktober",44:"November",48:"Dezember"}
MONTH_SPANS = []
_mk = sorted(MONTHS.keys())
for _i,_k in enumerate(_mk):
    _end = _mk[_i+1]-1 if _i+1<len(_mk) else 51
    MONTH_SPANS.append({"label":MONTHS[_k],"start":_k,"end":_end})
MONTH_LIST = ["Januar","Februar","März","April","Mai","Juni",
              "Juli","August","September","Oktober","November","Dezember"]

# Colors keyed by Bereich letter (A-H) as they appear in the sheet.
# Tuple: (section_header_bg, section_header_fg, cell_bg)
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
@st.cache_resource
def get_client():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=120, show_spinner=False)
def fetch_sheet(spreadsheet_url: str, sheet_name: str) -> list:
    return get_client().open_by_url(spreadsheet_url).worksheet(sheet_name).get_all_values()

# ── PLANNING PARSER ───────────────────────────────────────────────────────────
def parse_planning(raw: list) -> list:
    """
    Sheet layout (row 4 = header, row 5+ = data):
      Col A=Tag  B=FA/IB  C=Wer?  D=Bereich  E=ZeitStart  F=ZeitEnd
      G=Notizen  H=EventName  I..BH=52 week value columns
    Finds the header row by looking for 'Tag' in col A.
    """
    hdr_idx = None
    for i, row in enumerate(raw):
        if row and str(row[0]).strip().lower() == "tag":
            hdr_idx = i
            break
    if hdr_idx is None:
        return []

    # Auto-detect which column the event name is in and where weeks start.
    # We know: col 0=Tag,1=FA/IB,2=Wer?,3=Bereich,4=ZeitStart,5=ZeitEnd,6=Notizen,7=Name → weeks at 8
    # But we detect it robustly from the header row.
    hdr = raw[hdr_idx]
    name_col = 7  # default
    week_start_col = 8

    # Find the first column after col 3 whose header is NOT a time/meta field
    # and whose next column looks like a week number (1,2,3...)
    for ci in range(4, min(len(hdr), 15)):
        v = str(hdr[ci]).strip().lower()
        nxt = str(hdr[ci+1]).strip() if ci+1 < len(hdr) else ""
        if v not in ("zeit start","start","von","zeit ende","ende","bis","notizen","notiz","notes","") and nxt in [str(x) for x in range(1,6)]:
            name_col = ci
            week_start_col = ci + 1
            break

    rows = []
    for row in raw[hdr_idx+1:]:
        if len(row) <= name_col:
            continue
        name = str(row[name_col]).strip()
        if not name or name.upper() == "NAN":
            continue

        def g(ci):
            if ci < len(row):
                v = str(row[ci]).strip()
                return "" if v.upper() in ("NAN","NONE","#REF!","#N/A","#VALUE!") else v
            return ""

        # Parse 52 week values
        values = []
        for j in range(52):
            ci = week_start_col + j
            raw_val = g(ci)
            if not raw_val or raw_val in ("0","0.0","-"):
                values.append(None)
            else:
                try:
                    v = float(raw_val.replace(",","."))
                    values.append(v if v != 0 else None)
                except ValueError:
                    values.append(raw_val)  # keep "X" or other markers

        rows.append({
            "tag":        g(0),
            "fa_ib":      g(1),
            "wer":        g(2),
            "bereich":    g(3).upper(),   # normalise to uppercase letter
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

def times_overlap(t1,t2):
    if not t1 or not t2: return False
    return t1[0]<t2[1] and t2[0]<t1[1]

def get_category(event):
    e = str(event).lower()
    if "epic" in e:                                                          return "EPIC"
    if any(x in e for x in ["ecmo","lvad","impella","assist","prismax","beatmung"]): return "Geräte & Beatmung"
    if "workshop" in e:                                                      return "Workshop"
    if any(x in e for x in ["schulung","kurs","basiskurs","aufbaukurs","refresher"]): return "Schulung/Kurs"
    if any(x in e for x in ["sitzung","fachgruppe","fg ","superuser","fachforum"]):   return "Sitzung"
    if "einführung" in e or "einblick" in e:                                 return "Einführung"
    if any(x in e for x in ["lernwerkstatt","repe","simulation","kimsim"]):  return "Lernwerkstatt"
    if any(x in e for x in ["führungsdialog","austausch"]):                  return "Führung/Austausch"
    if "icu" in e:                                                           return "ICU"
    if any(x in e for x in ["kommunikation","aggressions"]):                 return "Kommunikation"
    if any(x in e for x in ["planung","bürotag"]):                           return "Planung"
    return "Sonstiges"

def make_color_map(events):
    def shades(base,n):
        base=base.lstrip("#"); r,g,b=int(base[0:2],16)/255,int(base[2:4],16)/255,int(base[4:6],16)/255
        h,l,s=colorsys.rgb_to_hls(r,g,b); out=[]
        for i in range(n):
            nl=max(0.35,min(0.75,l+(i-n/2)*0.08)); r2,g2,b2=colorsys.hls_to_rgb(h,nl,s)
            out.append("#{:02x}{:02x}{:02x}".format(int(r2*255),int(g2*255),int(b2*255)))
        return out
    groups={}
    for ev in events: groups.setdefault(get_category(ev),[]).append(ev)
    cmap={}
    for cat,evs in groups.items():
        base=BASE_COLORS.get(cat,"#B0BEC5")
        for ev,col in zip(sorted(evs),shades(base,len(evs))): cmap[ev]=col
    return cmap

def find_conflicts(df_ev):
    rc,pc=[],[]
    for date,ddf in df_ev.groupby("Datum"):
        rows=ddf.reset_index(drop=True).to_dict("records")
        for i in range(len(rows)):
            for j in range(i+1,len(rows)):
                r1,r2=rows[i],rows[j]
                t1,t2=parse_time_range(r1["Zeit"]),parse_time_range(r2["Zeit"])
                o1,o2=str(r1["Ort"]).strip(),str(r2["Ort"]).strip()
                if o1 and o2 and o1.lower()==o2.lower() and times_overlap(t1,t2):
                    rc.append({"Datum":date.strftime("%d.%m.%Y"),"Raum":o1,
                        "Event 1":r1["Event"],"Zeit 1":r1["Zeit"],
                        "Event 2":r2["Event"],"Zeit 2":r2["Zeit"]})
                p1={p.strip() for p in str(r1["Personen"]).split("/") if p.strip()}
                p2={p.strip() for p in str(r2["Personen"]).split("/") if p.strip()}
                shared=p1&p2
                if shared and times_overlap(t1,t2):
                    pc.append({"Datum":date.strftime("%d.%m.%Y"),"Person":", ".join(shared),
                        "Event 1":r1["Event"],"Zeit 1":r1["Zeit"],
                        "Event 2":r2["Event"],"Zeit 2":r2["Zeit"]})
    return rc,pc

def make_ical(df_ev):
    import hashlib
    lines=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//KIM ICU//DE","CALSCALE:GREGORIAN",
           "METHOD:PUBLISH","X-WR-CALNAME:KIM ICU 2026","X-WR-TIMEZONE:Europe/Zurich"]
    for _,row in df_ev.iterrows():
        uid=hashlib.md5(f"{row['Datum']}{row['Event']}{row['Zeit']}".encode()).hexdigest()
        ds=row["Datum"].strftime("%Y%m%d"); t=parse_time_range(row["Zeit"])
        if t:
            dtstart=f"DTSTART;TZID=Europe/Zurich:{ds}T{t[0].strftime('%H%M%S')}"
            dtend=f"DTEND;TZID=Europe/Zurich:{ds}T{t[1].strftime('%H%M%S')}"
        else:
            dtstart=f"DTSTART;VALUE=DATE:{ds}"; dtend=f"DTEND;VALUE=DATE:{ds}"
        summary=str(row["Event"]).replace(",","\\,").replace(";","\\;")
        desc=f"Person: {row['Personen']}\\nOrt: {row['Ort']}"
        if pd.notna(row.get("Bemerkungen","")) and str(row.get("Bemerkungen","")).strip():
            desc+=f"\\nBemerkung: {row['Bemerkungen']}"
        lines+=["BEGIN:VEVENT",f"UID:{uid}@kim.insel.ch",
                f"DTSTAMP:{datetime.datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                dtstart,dtend,f"SUMMARY:{summary}",
                f"LOCATION:{str(row['Ort']).replace(',','\\,')}",
                f"DESCRIPTION:{desc}","END:VEVENT"]
    lines.append("END:VCALENDAR"); return "\r\n".join(lines)

# ── PLANNING TABLE ────────────────────────────────────────────────────────────
def render_table(plan_rows, vis_month=0, bereich_filter="", search=""):
    if vis_month>0:
        ms=next((s for s in MONTH_SPANS if s["label"]==MONTH_LIST[vis_month-1]),None)
        vis_w=list(range(ms["start"],ms["end"]+1)) if ms else list(range(52))
    else:
        vis_w=list(range(52))

    vis_rows=[r for r in plan_rows
              if (not bereich_filter or r["bereich"]==bereich_filter)
              and (not search or search.lower() in r["name"].lower())]

    vis_ms=[]
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

    last_b=None
    for row in vis_rows:
        b=row["bereich"]
        if b!=last_b:
            last_b=b
            bg,fg,_=BEREICH_COLORS.get(b,BEREICH_COLORS[""])
            if b:
                html+=(f'<tr class="section-row"><td colspan="{7+len(vis_w)}" '
                       f'style="background:{bg};color:{fg}">Bereich {b}</td></tr>')

        _,_,cell_bg=BEREICH_COLORS.get(b,BEREICH_COLORS[""])
        html+=("<tr>"
               f'<td class="td-c0">{row["tag"]}</td>'
               f'<td class="td-c1">{row["fa_ib"]}</td>'
               f'<td class="td-c2">{row["wer"]}</td>'
               f'<td class="td-c3">{b}</td>'
               f'<td class="td-c4">{row["zeit_start"]}</td>'
               f'<td class="td-c5">{row["zeit_end"]}</td>'
               f'<td class="td-name">{row["name"]}</td>')
        for wi in vis_w:
            v=row["values"][wi] if wi<len(row["values"]) else None
            if v is not None:
                disp=int(v) if isinstance(v,float) and v==int(v) else (round(v,1) if isinstance(v,float) else v)
                all_w=[f"KW{WEEK_NUMS[i]} {fmt_date(WEEKS[i])}"
                       for i,val in enumerate(row["values"]) if val is not None and i<52]
                tip=f"{row['name']} | KW{WEEK_NUMS[wi]} ({fmt_date(WEEKS[wi])}): {disp} | Alle: {', '.join(all_w)}"
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
st.sidebar.markdown('<div style="font-size:10px;color:#7a9bb5;font-weight:600;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;">Google Sheet</div>', unsafe_allow_html=True)

sheet_url = st.sidebar.text_input(
    "Spreadsheet URL",
    value=st.secrets.get("sheet_url",""),
    placeholder="https://docs.google.com/spreadsheets/d/...",
)
planning_tab = st.sidebar.text_input("Blattname", value="Sitzungsdaten 2026")

if st.sidebar.button("Neu laden"):
    st.cache_data.clear(); st.rerun()

st.sidebar.markdown("---")
month_sel=st.sidebar.selectbox("Monat",["Ganzes Jahr"]+MONTH_LIST)
month_idx=MONTH_LIST.index(month_sel)+1 if month_sel!="Ganzes Jahr" else 0

# ── LOAD PLANNING DATA ────────────────────────────────────────────────────────
if not sheet_url.strip():
    st.info("Bitte die Google-Sheets-URL links eingeben.")
    st.stop()

with st.spinner("Lade Sitzungsplanung..."):
    try:
        raw=fetch_sheet(sheet_url.strip(), planning_tab.strip())
    except Exception as e:
        st.error(f"Fehler: {e}")
        st.stop()

plan_rows=parse_planning(raw)
if not plan_rows:
    st.warning("Keine Daten geladen — Blattname korrekt?")
    st.stop()

all_bereiche=sorted({r["bereich"] for r in plan_rows if r["bereich"]})
bereich_sel=st.sidebar.selectbox("Bereich",["Alle"]+all_bereiche)
bereich_filter="" if bereich_sel=="Alle" else bereich_sel

# ── METRICS ───────────────────────────────────────────────────────────────────
total=sum(1 for r in plan_rows for v in r["values"] if v is not None)
busy=max((sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)),default=0)

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
    <div class="metric-sub">Daten aus Google Sheets</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_plan,tab_cal,tab_conf,tab_ical,tab_heat,tab_stat = st.tabs([
    "Jahresplanung","Kalender","Konflikte","iCal Export","Heatmap","Statistik"
])

# ── TAB 1: JAHRESPLANUNG ──────────────────────────────────────────────────────
with tab_plan:
    c1,c2=st.columns([3,1])
    with c1:
        search_q=st.text_input("Suchen",placeholder="Ereignis suchen...",label_visibility="collapsed")
    with c2:
        st.caption(f"{len(plan_rows)} Ereignisse · 52 Wochen")

    st.markdown(render_table(plan_rows,vis_month=month_idx,bereich_filter=bereich_filter,search=search_q),
                unsafe_allow_html=True)

    # Color legend
    if all_bereiche:
        st.markdown("<br>",unsafe_allow_html=True)
        leg_cols=st.columns(min(len(all_bereiche),8))
        for col,b in zip(leg_cols,all_bereiche):
            _,fg,cell_bg=BEREICH_COLORS.get(b,BEREICH_COLORS[""])
            col.markdown(f'<div style="background:{cell_bg};color:{fg};padding:3px 8px;border-radius:4px;'
                         f'font-size:10px;font-weight:600;text-align:center">Bereich {b}</div>',
                         unsafe_allow_html=True)

# ── TAB 2: KALENDER ───────────────────────────────────────────────────────────
with tab_cal:
    # Build a simple event list from the planning grid for calendar display
    # Each scheduled week cell becomes a calendar entry on the Monday of that week
    cal_events=[]
    for row in plan_rows:
        for i,v in enumerate(row["values"]):
            if v is not None and i<len(WEEKS):
                d=datetime.datetime.strptime(WEEKS[i],"%Y-%m-%d").date()
                # find the actual day: use Monday of that week
                cal_events.append({
                    "date": d,
                    "name": row["name"],
                    "zeit_start": row["zeit_start"],
                    "zeit_end": row["zeit_end"],
                    "bereich": row["bereich"],
                    "val": v,
                    "month": d.month,
                    "day": d.day,
                })

    if not cal_events:
        st.info("Keine Daten fuer Kalenderansicht.")
    else:
        cal_df=pd.DataFrame(cal_events)
        cmap2={b:BEREICH_COLORS.get(b,BEREICH_COLORS[""])[2] for b in all_bereiche}
        months_show=range(1,13) if month_idx==0 else [month_idx]
        today_d=datetime.date.today()
        WDAYS=["Mo","Di","Mi","Do","Fr","Sa","So"]

        for m in months_show:
            mdf=cal_df[cal_df["month"]==m]
            if mdf.empty: continue
            mname=datetime.date(1900,m,1).strftime("%B")
            html=f'<div class="cal-wrap"><div class="cal-title">{mname} 2026</div><div class="cal-grid">'
            for d in WDAYS: html+=f'<div class="cal-hdr">{d}</div>'
            for week in calendar.monthcalendar(2026,m):
                for di,day in enumerate(week):
                    if day==0: html+='<div class="cal-cell-empty"></div>'; continue
                    is_today=(datetime.date(2026,m,day)==today_d)
                    is_wknd=(di>=5)
                    cls="cal-cell"
                    if is_today: cls+=" cal-cell-today"
                    elif is_wknd: cls+=" cal-cell-wknd"
                    ddf=mdf[mdf["day"]==day]
                    html+=f'<div class="{cls}"><div class="cal-daynum">{day:02d}</div>'
                    for idx,(_,row) in enumerate(ddf.iterrows()):
                        if idx==3: html+=f'<div class="cal-more">+{len(ddf)-3} weitere</div>'; break
                        color=cmap2.get(row["bereich"],"#e2e8f0")
                        time_str=f'{row["zeit_start"]}-{row["zeit_end"]}' if row["zeit_start"] else ""
                        html+=(f'<div class="cal-ev" style="background:{color};color:#1e3a5f">'
                               f'<span class="cal-ev-time">{time_str}</span>'
                               f'<span class="cal-ev-name">{str(row["name"])[:38]}</span></div>')
                    html+='</div>'
            html+='</div></div>'
            st.markdown(html,unsafe_allow_html=True)

# ── TAB 3: KONFLIKTE ──────────────────────────────────────────────────────────
with tab_conf:
    st.markdown("#### Zeitkonflikte in der Wochenplanung")
    st.caption("Zeigt Wochen, in denen mehrere Ereignisse mit gleicher Zeit Start/Ende geplant sind.")

    # Find weeks with multiple events at the same time
    conflicts=[]
    for wi in range(52):
        week_rows=[r for r in plan_rows if wi<len(r["values"]) and r["values"][wi] is not None]
        if len(week_rows)<2: continue
        # Check for time overlaps within same week
        for i in range(len(week_rows)):
            for j in range(i+1,len(week_rows)):
                r1,r2=week_rows[i],week_rows[j]
                # Build time strings from zeit_start/zeit_end
                t1_str=f"{r1['zeit_start']}-{r1['zeit_end']}" if r1['zeit_start'] and r1['zeit_end'] else ""
                t2_str=f"{r2['zeit_start']}-{r2['zeit_end']}" if r2['zeit_start'] and r2['zeit_end'] else ""
                t1=parse_time_range(t1_str)
                t2=parse_time_range(t2_str)
                if times_overlap(t1,t2):
                    conflicts.append({
                        "KW": WEEK_NUMS[wi],
                        "Datum": fmt_date(WEEKS[wi]),
                        "Ereignis 1": r1["name"],
                        "Zeit 1": t1_str,
                        "Ereignis 2": r2["name"],
                        "Zeit 2": t2_str,
                    })

    if not conflicts:
        st.success("Keine Zeitkonflikte gefunden.")
    else:
        st.warning(f"{len(conflicts)} Konflikt(e) gefunden")
        for c in conflicts:
            st.markdown(
                f'<div class="cf-card cf-room"><div class="cf-title">KW{c["KW"]} — {c["Datum"]}</div>'
                f'<div class="cf-detail">{c["Zeit 1"]} → {c["Ereignis 1"]}<br>'
                f'{c["Zeit 2"]} → {c["Ereignis 2"]}</div></div>',
                unsafe_allow_html=True)
        buf=io.StringIO(); pd.DataFrame(conflicts).to_csv(buf,index=False)
        st.download_button("Konflikte als CSV",buf.getvalue(),"konflikte_2026.csv","text/csv")

# ── TAB 4: iCAL ───────────────────────────────────────────────────────────────
with tab_ical:
    st.markdown("#### iCal Export")
    st.caption("Exportiert alle geplanten Ereignisse als .ics Datei fuer Outlook / Apple Calendar / Google Calendar.")

    # Build a proper dataframe for iCal from the planning grid
    ical_rows=[]
    for row in plan_rows:
        for i,v in enumerate(row["values"]):
            if v is not None and i<len(WEEKS):
                d=datetime.datetime.strptime(WEEKS[i],"%Y-%m-%d").date()
                zeit=""
                if row["zeit_start"] and row["zeit_end"]:
                    zeit=f"{row['zeit_start']}-{row['zeit_end']}"
                ical_rows.append({
                    "Datum":pd.Timestamp(d),
                    "Event":row["name"],
                    "Zeit":zeit,
                    "Ort":"",
                    "Personen":row["wer"] or "",
                    "Bemerkungen":row["notizen"] or "",
                })

    if not ical_rows:
        st.info("Keine Daten verfugbar.")
    else:
        ical_df=pd.DataFrame(ical_rows)
        # Filter by month if set
        if month_idx>0:
            ical_df=ical_df[ical_df["Datum"].dt.month==month_idx]
        if bereich_filter:
            names_in_bereich={r["name"] for r in plan_rows if r["bereich"]==bereich_filter}
            ical_df=ical_df[ical_df["Event"].isin(names_in_bereich)]

        ca,cb=st.columns(2)
        with ca:
            st.markdown(f"**{len(ical_df)} Termine** (gefiltert)")
            st.download_button("Gefiltert als .ics",make_ical(ical_df).encode("utf-8"),
                "kim_filtered.ics","text/calendar")
        with cb:
            all_ical_df=pd.DataFrame(ical_rows)
            st.markdown(f"**{len(all_ical_df)} Termine** (alles)")
            st.download_button("Alle als .ics",make_ical(all_ical_df).encode("utf-8"),
                "kim_alle_2026.ics","text/calendar")

        st.caption("Outlook: Datei > Oeffnen & Exportieren > Importieren > iCalendar (.ics)")

# ── TAB 5: HEATMAP ────────────────────────────────────────────────────────────
with tab_heat:
    st.markdown("#### Event-Dichte 2026")
    # Count events per week-start date
    day_counts={}
    for row in plan_rows:
        for i,v in enumerate(row["values"]):
            if v is not None and i<len(WEEKS):
                d=datetime.datetime.strptime(WEEKS[i],"%Y-%m-%d").date()
                day_counts[d]=day_counts.get(d,0)+1

    mx=max(day_counts.values(),default=1)
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
    st.markdown(cells,unsafe_allow_html=True)

# ── TAB 6: STATISTIK ──────────────────────────────────────────────────────────
with tab_stat:
    import plotly.graph_objects as go
    st.markdown("#### Statistik")

    # Events per month
    month_counts={ms["label"]:sum(1 for r in plan_rows
                  for i in range(ms["start"],ms["end"]+1)
                  if i<len(r["values"]) and r["values"][i] is not None)
                  for ms in MONTH_SPANS}
    fig_m=go.Figure(go.Bar(x=list(month_counts.keys()),y=list(month_counts.values()),
        marker_color="#378ADD",opacity=0.85))
    fig_m.update_layout(height=260,plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10,r=10,t=10,b=10),
        xaxis=dict(gridcolor="#f1f5f9"),yaxis=dict(title="Anzahl Eintraege",gridcolor="#f1f5f9"))
    st.plotly_chart(fig_m,use_container_width=True)

    # Events per Bereich
    b_counts={}
    for r in plan_rows:
        b_counts[r["bereich"]]=b_counts.get(r["bereich"],0)+sum(1 for v in r["values"] if v is not None)
    b_colors=[BEREICH_COLORS.get(b,BEREICH_COLORS[""])[2] for b in b_counts]
    fig_b=go.Figure(go.Bar(x=list(b_counts.values()),y=[f"Bereich {b}" for b in b_counts.keys()],
        orientation="h",marker_color=b_colors))
    fig_b.update_layout(height=300,plot_bgcolor="rgba(0,0,0,0)",paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10,r=10,t=10,b=10),
        xaxis=dict(title="Anzahl Eintraege",gridcolor="#f1f5f9"),yaxis=dict(gridcolor="#f1f5f9"))
    st.plotly_chart(fig_b,use_container_width=True)
