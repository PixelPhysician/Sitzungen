import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys
import io
import re
from google.oauth2 import service_account
import gspread
import plotly.graph_objects as go

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="KIM Sitzungsplanung 2026")

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1i3UXnhEVvuhpB-MwSeOyoX4xfkMCViSfILyKFUB9VeQ"
PLANNING_TAB = "Sitzungsdaten 2026"

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

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
@st.cache_resource
def get_client():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=120)
def fetch_sheet(url, sheet):
    return get_client().open_by_url(url).worksheet(sheet).get_all_values()

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
with st.spinner("Lade Sitzungsplanung..."):
    raw = fetch_sheet(DEFAULT_SHEET_URL, PLANNING_TAB)

# ── PARSER ────────────────────────────────────────────────────────────────────
def parse_planning(raw):
    hdr_idx = next(i for i,r in enumerate(raw) if r and r[0].lower()=="tag")
    rows=[]
    for r in raw[hdr_idx+1:]:
        if len(r)<8 or not r[7]: continue
        vals=[]
        for i in range(52):
            try:
                v=float(r[8+i].replace(",",".")) if r[8+i] else None
                vals.append(v if v else None)
            except:
                vals.append(None)
        rows.append({
            "tag":r[0],"fa_ib":r[1],"wer":r[2],
            "bereich":r[3].upper(),"zeit_start":r[4],
            "zeit_end":r[5],"name":r[7],"values":vals
        })
    return rows

plan_rows = parse_planning(raw)

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("### KIM Planung")
if st.sidebar.button("↺ Neu laden"):
    st.cache_data.clear()
    st.rerun()

month_sel = st.sidebar.selectbox("Monat", ["Ganzes Jahr"] + MONTH_LIST)
month_idx = MONTH_LIST.index(month_sel)+1 if month_sel!="Ganzes Jahr" else 0

bereiche = sorted({r["bereich"] for r in plan_rows if r["bereich"]})
bereich_sel = st.sidebar.selectbox("Bereich", ["Alle"] + bereiche)
bereich_filter = "" if bereich_sel=="Alle" else bereich_sel

# ── METRICS ───────────────────────────────────────────────────────────────────
total = sum(1 for r in plan_rows for v in r["values"] if v is not None)
busy = max((sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)), default=0)

c1,c2,c3,c4 = st.columns(4)
c1.metric("Einträge", total)
c2.metric("Events", len(plan_rows))
c3.metric("Max/Woche", busy)
c4.metric("Zeit", datetime.datetime.now().strftime('%H:%M'))

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_plan,tab_cal,tab_conf,tab_heat,tab_stat = st.tabs([
    "Plan","Kalender","Konflikte","Heatmap","Statistik"
])

# ── PLAN ──────────────────────────────────────────────────────────────────────
with tab_plan:
    for r in plan_rows:
        if bereich_filter and r["bereich"]!=bereich_filter:
            continue
        active=[WEEK_NUMS[i] for i,v in enumerate(r["values"]) if v is not None]
        if active:
            st.write(f"**{r['name']}** ({r['bereich']}) → KW {active[:6]}")

# ── KALENDER ──────────────────────────────────────────────────────────────────
with tab_cal:
    events=[]
    for r in plan_rows:
        for i,v in enumerate(r["values"]):
            if v is not None:
                d=datetime.datetime.strptime(WEEKS[i],"%Y-%m-%d").date()
                events.append((d,r["name"]))
    df=pd.DataFrame(events,columns=["date","name"])
    st.dataframe(df)

# ── KONFLIKTE ─────────────────────────────────────────────────────────────────
with tab_conf:
    st.info("Simple overlap check")
    conflicts=[]
    for i in range(52):
        ev=[r["name"] for r in plan_rows if r["values"][i] is not None]
        if len(ev)>1:
            conflicts.append((i+1,ev))
    for kw,ev in conflicts[:10]:
        st.write(f"KW{kw}: {ev}")

# ── HEATMAP ───────────────────────────────────────────────────────────────────
with tab_heat:
    counts=[sum(1 for r in plan_rows if r["values"][i] is not None) for i in range(52)]
    st.bar_chart(counts)

# ── STAT ──────────────────────────────────────────────────────────────────────
with tab_stat:
    b_counts={}
    for r in plan_rows:
        b_counts[r["bereich"]] = b_counts.get(r["bereich"],0)+sum(1 for v in r["values"] if v)
    fig = go.Figure(go.Bar(x=list(b_counts.keys()), y=list(b_counts.values())))
    st.plotly_chart(fig, use_container_width=True)
