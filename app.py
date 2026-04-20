import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys
import io

st.set_page_config(layout="wide", page_title="KIM Schedule 2026")

# =========================
# CUSTOM CSS – clean clinical look
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Dark sidebar */
[data-testid="stSidebar"] {
    background: #0f1923 !important;
    border-right: 1px solid #1e2d3d;
}
[data-testid="stSidebar"] * {
    color: #c8d8e8 !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #7a9bb5 !important;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
[data-testid="stSidebar"] [data-baseweb="select"] {
    background: #1a2a3a !important;
    border-color: #2a3f55 !important;
}

/* Main area */
.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1400px;
}

/* Header banner */
.dashboard-header {
    background: linear-gradient(135deg, #0f1923 0%, #1a2f47 50%, #0d2137 100%);
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border: 1px solid #1e3a55;
    position: relative;
    overflow: hidden;
}
.dashboard-header::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(59,130,246,0.12) 0%, transparent 70%);
    pointer-events: none;
}
.header-title {
    font-size: 26px;
    font-weight: 700;
    color: #f0f6ff;
    letter-spacing: -0.02em;
    margin: 0;
}
.header-subtitle {
    font-size: 13px;
    color: #6b92b5;
    margin-top: 4px;
    font-weight: 400;
    letter-spacing: 0.03em;
}
.header-badge {
    background: rgba(59,130,246,0.15);
    border: 1px solid rgba(59,130,246,0.3);
    color: #7eb8f7;
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.05em;
}
.header-cross {
    font-size: 42px;
    margin-right: 16px;
    opacity: 0.8;
}

/* Metric cards */
.metric-row {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 24px;
}
.metric-card {
    background: #ffffff;
    border: 1px solid #e8ecf0;
    border-radius: 10px;
    padding: 16px 20px;
    position: relative;
    overflow: hidden;
}
.metric-card::before {
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    border-radius: 3px 0 0 3px;
}
.metric-card.blue::before  { background: #3b82f6; }
.metric-card.teal::before  { background: #14b8a6; }
.metric-card.amber::before { background: #f59e0b; }
.metric-card.rose::before  { background: #f43f5e; }
.metric-label {
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #94a3b8;
    margin-bottom: 6px;
}
.metric-value {
    font-size: 28px;
    font-weight: 700;
    color: #0f1923;
    letter-spacing: -0.03em;
    font-family: 'DM Mono', monospace;
}
.metric-sub {
    font-size: 11px;
    color: #94a3b8;
    margin-top: 2px;
}

/* Section headers */
.section-header {
    font-size: 13px;
    font-weight: 700;
    color: #334155;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 2px solid #e2e8f0;
}

/* Conflict alert */
.conflict-card {
    background: #fff7ed;
    border: 1px solid #fed7aa;
    border-left: 4px solid #f97316;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 10px;
    font-size: 13px;
}
.conflict-card.room { border-left-color: #ef4444; background: #fff5f5; border-color: #fecaca; }
.conflict-card.person { border-left-color: #8b5cf6; background: #faf5ff; border-color: #ddd6fe; }
.conflict-title {
    font-weight: 700;
    font-size: 13px;
    margin-bottom: 4px;
}
.conflict-detail {
    color: #64748b;
    font-size: 12px;
    font-family: 'DM Mono', monospace;
}

/* Heatmap */
.heatmap-wrap {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
}
.heatmap-grid {
    display: grid;
    grid-template-columns: 32px repeat(53, 1fr);
    gap: 2px;
    align-items: center;
}
.hm-cell {
    width: 100%;
    aspect-ratio: 1;
    border-radius: 2px;
    cursor: default;
    transition: transform 0.1s;
}
.hm-cell:hover { transform: scale(1.4); z-index: 10; position: relative; }
.hm-label {
    font-size: 9px;
    color: #94a3b8;
    text-align: right;
    padding-right: 4px;
    line-height: 1;
}

/* iCal export buttons */
.ical-btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #0f1923;
    color: #7eb8f7;
    border: 1px solid #1e3a55;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    margin-right: 8px;
    margin-bottom: 8px;
    text-decoration: none;
    font-family: 'DM Sans', sans-serif;
}

/* Expander tweaks */
[data-testid="stExpander"] {
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    margin-bottom: 8px !important;
}
[data-testid="stExpander"] summary {
    font-weight: 600 !important;
    color: #1e293b !important;
    padding: 12px 16px !important;
}

/* Calendar */
.cal-wrap { font-family: 'DM Sans', sans-serif; margin-bottom: 32px; }
.cal-title { font-size: 17px; font-weight: 700; margin-bottom: 10px; color: #0f1923; letter-spacing: -0.01em; }
.cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; }
.cal-hdr { text-align: center; font-weight: 700; font-size: 10px; padding: 5px 0; color: #94a3b8;
           background: #f8fafc; border-radius: 4px 4px 0 0; letter-spacing: 0.06em; }
.cal-cell { border: 1px solid #e8ecf2; border-radius: 6px; padding: 5px; min-height: 90px;
            background: #fff; box-sizing: border-box; }
.cal-cell-wknd { background: #fafbfc; }
.cal-cell-today { border: 2px solid #3b82f6; background: #eff6ff; }
.cal-cell-empty { border: none; background: transparent; min-height: 90px; }
.cal-daynum { font-weight: 700; font-size: 11px; color: #475569; margin-bottom: 3px;
              font-family: 'DM Mono', monospace; }
.cal-ev { padding: 2px 5px; margin-top: 2px; border-radius: 3px; font-size: 9px; line-height: 1.4; word-break: break-word; }
.cal-ev-time { font-weight: 700; display: block; font-family: 'DM Mono', monospace; }
.cal-ev-name { display: block; font-size: 8.5px; opacity: 0.9; }
.cal-more { font-size: 9px; color: #94a3b8; font-style: italic; margin-top: 3px; }
.cal-conflict { position: relative; }
.cal-conflict::after { content: ''; font-size: 8px; position: absolute; top: 2px; right: 4px; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #f1f5f9; border-radius: 8px; padding: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 6px; font-weight: 600; font-size: 13px; padding: 6px 16px; }
.stTabs [aria-selected="true"] { background: #fff !important; color: #0f1923 !important; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }

/* Download button */
.stDownloadButton button {
    background: #0f1923 !important;
    color: #7eb8f7 !important;
    border: 1px solid #1e3a55 !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    font-size: 12px !important;
}
.stDownloadButton button:hover {
    background: #1a2f47 !important;
    border-color: #3b82f6 !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
st.markdown("""
<div class="dashboard-header">
  <div style="display:flex;align-items:center;">
    <div>
      <div class="header-title">KIM — ICU Schedule Dashboard</div>
      <div class="header-subtitle">Klinik für Intensivmedizin · Inselspital Bern · 2026</div>
    </div>
  </div>
  <div class="header-badge">2026 PLANUNG</div>
</div>
""", unsafe_allow_html=True)

GITHUB_URL = "https://raw.githubusercontent.com/PixelPhysician/Sitzungen/main/Sitzungsdaten_sort.xlsx"

@st.cache_data(ttl=300)
def load_data():
    return pd.read_excel(GITHUB_URL)

df = load_data()

# =========================
# DATEN BEREINIGEN
# =========================
df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Datum"])
df["Zeit"] = df["Zeit"].fillna("Zeit unbekannt")
df["Personen"] = df["Personen"].fillna("Person unbekannt")
df["Ort"] = df["Ort"].fillna("Ort unbekannt")
df["Event"] = df["Event"].fillna("Thema unbekannt")
df["Month"] = df["Datum"].dt.month
df["Day"] = df["Datum"].dt.day
df["Weekday"] = df["Datum"].dt.strftime("%a")

# =========================
# KATEGORIEN & FARBEN
# =========================
def get_category(event):
    e = str(event).lower()
    if "epic" in e: return "EPIC"
    elif any(x in e for x in ["ecmo","lvad","impella","assist","prismax","beatmung"]): return "Geräte & Beatmung"
    elif "workshop" in e: return "Workshop"
    elif any(x in e for x in ["schulung","kurs","basiskurs","aufbaukurs","refresher"]): return "Schulung/Kurs"
    elif any(x in e for x in ["sitzung","fachgruppe","fg ","superuser","fachforum"]): return "Sitzung"
    elif "einführung" in e or "einblick" in e: return "Einführung"
    elif any(x in e for x in ["lernwerkstatt","repe","simulation","kimsim"]): return "Lernwerkstatt"
    elif any(x in e for x in ["führungsdialog","austausch"]): return "Führung/Austausch"
    elif "icu" in e: return "ICU"
    elif any(x in e for x in ["kommunikation","aggressions"]): return "Kommunikation"
    elif any(x in e for x in ["planung","bürotag"]): return "Planung"
    else: return "Sonstiges"

BASE_COLORS = {
    "EPIC": "#4FC3F7", "Geräte & Beatmung": "#E57373", "Workshop": "#FFD54F",
    "Schulung/Kurs": "#81C784", "Sitzung": "#FF8A65", "Einführung": "#4DB6AC",
    "Lernwerkstatt": "#F48FB1", "Führung/Austausch": "#CE93D8", "ICU": "#90CAF9",
    "Kommunikation": "#FFCC80", "Planung": "#A5D6A7", "Sonstiges": "#B0BEC5",
}

def generate_shades(base_hex, n):
    base_hex = base_hex.lstrip('#')
    r,g,b = int(base_hex[0:2],16)/255, int(base_hex[2:4],16)/255, int(base_hex[4:6],16)/255
    h,l,s = colorsys.rgb_to_hls(r,g,b)
    shades = []
    for i in range(n):
        new_l = max(0.35, min(0.75, l+(i-n/2)*0.08))
        r2,g2,b2 = colorsys.hls_to_rgb(h, new_l, s)
        shades.append('#{:02x}{:02x}{:02x}'.format(int(r2*255),int(g2*255),int(b2*255)))
    return shades

df["Kategorie"] = df["Event"].apply(get_category)
event_groups = {}
for event in df["Event"].unique():
    cat = get_category(event)
    event_groups.setdefault(cat, []).append(event)

event_color_map = {}
for cat, events in event_groups.items():
    base = BASE_COLORS.get(cat, "#B0BEC5")
    shades = generate_shades(base, len(events))
    for event, color in zip(sorted(events), shades):
        event_color_map[event] = color

def get_color(event):
    return event_color_map.get(event, "#CCCCCC")

# =========================
# CONFLICT DETECTION
# =========================
def parse_time_range(zeit_str):
    """Parse '08:00-16:00' -> (datetime.time, datetime.time) or None"""
    s = str(zeit_str).strip()
    if "unbekannt" in s.lower() or "ganzer" in s.lower():
        return None
    # try HH:MM-HH:MM
    import re
    m = re.match(r'(\d{1,2})[:\.](\d{2})\s*[-–]\s*(\d{1,2})[:\.](\d{2})', s)
    if m:
        try:
            t_start = datetime.time(int(m.group(1)), int(m.group(2)))
            t_end   = datetime.time(int(m.group(3)), int(m.group(4)))
            return (t_start, t_end)
        except:
            return None
    return None

def times_overlap(t1, t2):
    """Check if two (start, end) time tuples overlap"""
    if t1 is None or t2 is None:
        return False
    s1, e1 = t1
    s2, e2 = t2
    return s1 < e2 and s2 < e1

def find_conflicts(df_input):
    room_conflicts = []
    person_conflicts = []

    for date, day_df in df_input.groupby("Datum"):
        day_df = day_df.reset_index(drop=True)
        rows = day_df.to_dict("records")

        for i in range(len(rows)):
            for j in range(i+1, len(rows)):
                r1, r2 = rows[i], rows[j]
                t1 = parse_time_range(r1["Zeit"])
                t2 = parse_time_range(r2["Zeit"])

                # Room conflict
                ort1, ort2 = str(r1["Ort"]).strip(), str(r2["Ort"]).strip()
                if (ort1 not in ["Ort unbekannt",""] and
                    ort2 not in ["Ort unbekannt",""] and
                    ort1.lower() == ort2.lower() and
                    times_overlap(t1, t2)):
                    room_conflicts.append({
                        "Datum": date.strftime("%d.%m.%Y"),
                        "Raum": ort1,
                        "Event 1": r1["Event"], "Zeit 1": r1["Zeit"],
                        "Event 2": r2["Event"], "Zeit 2": r2["Zeit"],
                    })

                # Person conflict
                p1s = [p.strip() for p in str(r1["Personen"]).split("/") if p.strip() and p.strip() != "Person unbekannt"]
                p2s = [p.strip() for p in str(r2["Personen"]).split("/") if p.strip() and p.strip() != "Person unbekannt"]
                shared = set(p1s) & set(p2s)
                if shared and times_overlap(t1, t2):
                    person_conflicts.append({
                        "Datum": date.strftime("%d.%m.%Y"),
                        "Person": ", ".join(shared),
                        "Event 1": r1["Event"], "Zeit 1": r1["Zeit"],
                        "Event 2": r2["Event"], "Zeit 2": r2["Zeit"],
                    })

    return room_conflicts, person_conflicts

# =========================
# iCAL GENERATION
# =========================
def make_ical(df_input):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//KIM ICU Schedule//DE",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:KIM ICU Schedule 2026",
        "X-WR-TIMEZONE:Europe/Zurich",
    ]
    import re, hashlib
    for _, row in df_input.iterrows():
        uid = hashlib.md5(f"{row['Datum']}{row['Event']}{row['Zeit']}".encode()).hexdigest()
        date_str = row["Datum"].strftime("%Y%m%d")

        t = parse_time_range(row["Zeit"])
        if t:
            dtstart = f"DTSTART;TZID=Europe/Zurich:{date_str}T{t[0].strftime('%H%M%S')}"
            dtend   = f"DTEND;TZID=Europe/Zurich:{date_str}T{t[1].strftime('%H%M%S')}"
        else:
            dtstart = f"DTSTART;VALUE=DATE:{date_str}"
            dtend   = f"DTEND;VALUE=DATE:{date_str}"

        summary = str(row["Event"]).replace(",","\\,").replace(";","\\;")
        location = str(row["Ort"]).replace(",","\\,")
        description = f"Person: {row['Personen']}\\nOrt: {row['Ort']}"
        if "Bemerkungen" in row and pd.notna(row.get("Bemerkungen","")):
            description += f"\\nBemerkung: {row['Bemerkungen']}"

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}@kim.insel.ch",
            f"DTSTAMP:{datetime.datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            dtstart, dtend,
            f"SUMMARY:{summary}",
            f"LOCATION:{location}",
            f"DESCRIPTION:{description}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)

# =========================
# FILTER SIDEBAR
# =========================
st.sidebar.markdown("""
<div style="padding: 16px 0 8px 0;">
  <div style="font-size:9px;font-weight:700;letter-spacing:0.15em;color:#4a6f8a;text-transform:uppercase;margin-bottom:2px;">ICU SCHEDULE</div>
  <div style="font-size:16px;font-weight:700;color:#e8f4ff;margin-bottom:16px;">Filter & Export</div>
</div>
""", unsafe_allow_html=True)

month_options = ["Ganzes Jahr"] + sorted(df["Month"].unique())
selected_month = st.sidebar.selectbox(
    "Monat",
    month_options,
    index=0,
    format_func=lambda x: x if x=="Ganzes Jahr" else datetime.date(1900,x,1).strftime('%B')
)
event_filter  = st.sidebar.multiselect("Event-Typ", sorted(df["Event"].unique()))
person_filter = st.sidebar.multiselect("Person",    sorted(df["Personen"].unique()))
room_filter   = st.sidebar.multiselect("Raum",      sorted(df["Ort"].unique()))

filtered_df = df.copy()
if selected_month != "Ganzes Jahr": filtered_df = filtered_df[filtered_df["Month"]==selected_month]
if event_filter:  filtered_df = filtered_df[filtered_df["Event"].isin(event_filter)]
if person_filter: filtered_df = filtered_df[filtered_df["Personen"].isin(person_filter)]
if room_filter:   filtered_df = filtered_df[filtered_df["Ort"].isin(room_filter)]
filtered_df = filtered_df.sort_values("Datum")

# Sidebar legend
st.sidebar.markdown("---")
st.sidebar.markdown('<div style="font-size:9px;font-weight:700;letter-spacing:0.12em;color:#4a6f8a;text-transform:uppercase;margin-bottom:8px;">Kategorien</div>', unsafe_allow_html=True)
for cat, color in BASE_COLORS.items():
    count = len(filtered_df[filtered_df["Kategorie"]==cat])
    if count > 0:
        st.sidebar.markdown(f"""
        <div style="display:flex;align-items:center;font-size:11px;margin-bottom:4px;color:#c8d8e8;">
            <div style="width:8px;height:8px;background:{color};margin-right:8px;border-radius:2px;flex-shrink:0;"></div>
            {cat} <span style="color:#4a6f8a;margin-left:auto;font-family:'DM Mono',monospace;font-size:10px;">{count}</span>
        </div>""", unsafe_allow_html=True)

# Sidebar iCal export
st.sidebar.markdown("---")
st.sidebar.markdown('<div style="font-size:9px;font-weight:700;letter-spacing:0.12em;color:#4a6f8a;text-transform:uppercase;margin-bottom:8px;">iCal Export</div>', unsafe_allow_html=True)
ical_data = make_ical(filtered_df)
st.sidebar.download_button(
    label=f" Gefilterte Events ({len(filtered_df)}) als .ics",
    data=ical_data.encode("utf-8"),
    file_name="kim_schedule_filtered.ics",
    mime="text/calendar",
)
ical_all = make_ical(df)
st.sidebar.download_button(
    label=f" Alle Events ({len(df)}) als .ics",
    data=ical_all.encode("utf-8"),
    file_name="kim_schedule_2026_all.ics",
    mime="text/calendar",
)

# =========================
# METRICS
# =========================
conflict_room_all, conflict_person_all = find_conflicts(filtered_df)
n_conflict_days = len(set([c["Datum"] for c in conflict_room_all + conflict_person_all]))

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card blue">
    <div class="metric-label">Total Events</div>
    <div class="metric-value">{len(filtered_df)}</div>
    <div class="metric-sub">im gewählten Zeitraum</div>
  </div>
  <div class="metric-card teal">
    <div class="metric-label">Veranstaltungstage</div>
    <div class="metric-value">{filtered_df['Datum'].nunique()}</div>
    <div class="metric-sub">Tage mit Events</div>
  </div>
  <div class="metric-card amber">
    <div class="metric-label">Raum-Konflikte</div>
    <div class="metric-value">{len(conflict_room_all)}</div>
    <div class="metric-sub">Doppelbuchungen erkannt</div>
  </div>
  <div class="metric-card rose">
    <div class="metric-label">Personen-Konflikte</div>
    <div class="metric-value">{len(conflict_person_all)}</div>
    <div class="metric-sub">Überlappungen erkannt</div>
  </div>
</div>
""", unsafe_allow_html=True)

# =========================
# TABS
# =========================
tab_heatmap, tab_conflicts, tab_calendar, tab_ical, tab_list, tab_bar = st.tabs([
    " Jahres-Heatmap", "️ Konflikt-Checker", " Kalender", " iCal Export", " Liste", " Balken"
])

# ---- HEATMAP TAB ----
with tab_heatmap:
    st.markdown('<div class="section-header">Event-Dichte 2026</div>', unsafe_allow_html=True)
    st.caption("Jede Zelle = ein Tag. Farbe = Anzahl Events (ähnlich GitHub Contribution Graph).")

    # Build day counts
    day_counts = filtered_df.groupby(filtered_df["Datum"].dt.date).size().to_dict()

    max_count = max(day_counts.values()) if day_counts else 1

    def heatmap_color(n, max_n):
        if n == 0: return "#eef2f7"
        t = n / max_n
        # Deep blue palette
        if t < 0.25:  return "#bfdbfe"
        elif t < 0.5: return "#60a5fa"
        elif t < 0.75:return "#2563eb"
        else:         return "#1e3a8a"

    WEEKDAY_LABELS = ["Mo","Di","Mi","Do","Fr","Sa","So"]

    # Build 53-week grid
    jan1 = datetime.date(2026, 1, 1)
    start = jan1 - datetime.timedelta(days=jan1.weekday())  # Monday before Jan 1

    cells_html = '<div class="heatmap-wrap">'
    # Month labels row
    cells_html += '<div style="display:flex;margin-left:36px;margin-bottom:4px;">'
    months_shown = {}
    for w in range(53):
        d = start + datetime.timedelta(weeks=w)
        m = d.strftime("%b")
        if d.month not in months_shown:
            months_shown[d.month] = True
            cells_html += f'<div style="flex:1;font-size:9px;color:#94a3b8;font-family:DM Mono,monospace;">{m}</div>'
        else:
            cells_html += '<div style="flex:1;"></div>'
    cells_html += '</div>'

    # Grid rows (7 weekdays)
    cells_html += '<div style="display:flex;gap:2px;">'
    # Weekday labels column
    cells_html += '<div style="display:flex;flex-direction:column;gap:2px;width:32px;flex-shrink:0;">'
    for wd in WEEKDAY_LABELS:
        cells_html += f'<div style="height:14px;line-height:14px;font-size:9px;color:#94a3b8;text-align:right;padding-right:4px;font-family:DM Mono,monospace;">{wd}</div>'
    cells_html += '</div>'

    # Week columns
    cells_html += '<div style="display:flex;gap:2px;flex:1;">'
    for w in range(53):
        cells_html += '<div style="display:flex;flex-direction:column;gap:2px;flex:1;">'
        for wd in range(7):
            d = start + datetime.timedelta(weeks=w, days=wd)
            if d.year != 2026:
                cells_html += '<div style="height:14px;border-radius:2px;background:#f8fafc;flex:1;"></div>'
            else:
                n = day_counts.get(d, 0)
                color = heatmap_color(n, max_count)
                tooltip = f"{d.strftime('%d.%m.%Y')}: {n} Event{'s' if n!=1 else ''}"
                cells_html += f'<div title="{tooltip}" style="height:14px;border-radius:2px;background:{color};flex:1;cursor:default;" onmouseover="this.style.outline=\'2px solid #3b82f6\'" onmouseout="this.style.outline=\'none\'"></div>'
        cells_html += '</div>'
    cells_html += '</div></div>'

    # Legend
    cells_html += '<div style="display:flex;align-items:center;gap:6px;margin-top:12px;font-size:10px;color:#94a3b8;">'
    cells_html += '<span>Weniger</span>'
    for color in ["#eef2f7","#bfdbfe","#60a5fa","#2563eb","#1e3a8a"]:
        cells_html += f'<div style="width:14px;height:14px;border-radius:2px;background:{color};"></div>'
    cells_html += '<span>Mehr</span></div>'
    cells_html += '</div>'

    st.markdown(cells_html, unsafe_allow_html=True)

    # Also keep the original bubble chart
    st.markdown('<div class="section-header" style="margin-top:28px;">Events pro Tag nach Kategorie (Bubbles)</div>', unsafe_allow_html=True)
    import plotly.graph_objects as go

    bubble_df = filtered_df.groupby(["Datum","Kategorie"]).size().reset_index(name="Anzahl")
    fig_bubble = go.Figure()
    for cat, color in BASE_COLORS.items():
        cat_df = bubble_df[bubble_df["Kategorie"]==cat]
        if cat_df.empty: continue
        fig_bubble.add_trace(go.Scatter(
            x=cat_df["Datum"], y=cat_df["Kategorie"], mode="markers",
            marker=dict(size=cat_df["Anzahl"]*14, color=color, opacity=0.8,
                        line=dict(width=1, color="white"), sizemode="diameter"),
            name=cat,
            text=[f"{r['Datum'].strftime('%d.%m.%Y')}: {r['Anzahl']} Events" for _,r in cat_df.iterrows()],
            hovertemplate="%{text}<extra></extra>",
        ))
    fig_bubble.update_layout(
        height=450, plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(range=["2026-01-01","2026-12-31"], tickformat="%b", dtick="M1",
                   gridcolor="#f1f5f9"),
        yaxis=dict(gridcolor="#f1f5f9"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
        margin=dict(l=10,r=10,t=10,b=10), paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_bubble, use_container_width=True)

# ---- CONFLICT TAB ----
with tab_conflicts:
    st.markdown('<div class="section-header">Konflikt-Checker</div>', unsafe_allow_html=True)
    st.caption("Erkennt Doppelbuchungen von Räumen und Personen bei überlappenden Zeiten.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"####  Raum-Konflikte ({len(conflict_room_all)})")
        if not conflict_room_all:
            st.success(" Keine Raum-Doppelbuchungen gefunden!")
        else:
            for c in conflict_room_all:
                st.markdown(f"""
                <div class="conflict-card room">
                  <div class="conflict-title"> {c['Raum']} — {c['Datum']}</div>
                  <div class="conflict-detail">
                    {c['Zeit 1']}  →  {c['Event 1']}<br>
                    {c['Zeit 2']}  →  {c['Event 2']}
                  </div>
                </div>""", unsafe_allow_html=True)

    with col2:
        st.markdown(f"####  Personen-Konflikte ({len(conflict_person_all)})")
        if not conflict_person_all:
            st.success(" Keine Personen-Überlappungen gefunden!")
        else:
            for c in conflict_person_all:
                st.markdown(f"""
                <div class="conflict-card person">
                  <div class="conflict-title"> {c['Person']} — {c['Datum']}</div>
                  <div class="conflict-detail">
                    {c['Zeit 1']}  →  {c['Event 1']}<br>
                    {c['Zeit 2']}  →  {c['Event 2']}
                  </div>
                </div>""", unsafe_allow_html=True)

    if conflict_room_all or conflict_person_all:
        st.markdown("---")
        st.markdown("**Export Konflikte als CSV:**")
        all_conflicts = []
        for c in conflict_room_all:
            all_conflicts.append({"Typ":"Raum-Konflikt","Datum":c["Datum"],"Raum/Person":c["Raum"],
                                   "Event 1":c["Event 1"],"Zeit 1":c["Zeit 1"],
                                   "Event 2":c["Event 2"],"Zeit 2":c["Zeit 2"]})
        for c in conflict_person_all:
            all_conflicts.append({"Typ":"Personen-Konflikt","Datum":c["Datum"],"Raum/Person":c["Person"],
                                   "Event 1":c["Event 1"],"Zeit 1":c["Zeit 1"],
                                   "Event 2":c["Event 2"],"Zeit 2":c["Zeit 2"]})
        conflict_df = pd.DataFrame(all_conflicts)
        csv_buf = io.StringIO()
        conflict_df.to_csv(csv_buf, index=False)
        st.download_button(" Konflikte als CSV herunterladen", csv_buf.getvalue(),
                           "konflikte_2026.csv", "text/csv")

# ---- CALENDAR TAB ----
with tab_calendar:
    months_to_show = range(1,13) if selected_month=="Ganzes Jahr" else [selected_month]
    WEEKDAYS_DE = ["Mo","Di","Mi","Do","Fr","Sa","So"]
    today = datetime.date.today()

    # Build set of conflict dates
    conflict_dates = set([c["Datum"] for c in conflict_room_all + conflict_person_all])

    cal_css = """
    <style>
    .cal-wrap{font-family:'DM Sans',sans-serif;margin-bottom:32px;}
    .cal-title{font-size:17px;font-weight:700;margin-bottom:10px;color:#0f1923;letter-spacing:-0.01em;}
    .cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;}
    .cal-hdr{text-align:center;font-weight:700;font-size:10px;padding:5px 0;color:#94a3b8;background:#f8fafc;border-radius:4px 4px 0 0;letter-spacing:0.06em;}
    .cal-cell{border:1px solid #e8ecf2;border-radius:6px;padding:5px;min-height:90px;background:#fff;box-sizing:border-box;position:relative;}
    .cal-cell-wknd{background:#fafbfc;}
    .cal-cell-today{border:2px solid #3b82f6;background:#eff6ff;}
    .cal-cell-conflict{border:2px solid #f97316 !important;}
    .cal-cell-empty{border:none;background:transparent;min-height:90px;}
    .cal-daynum{font-weight:700;font-size:11px;color:#475569;margin-bottom:3px;font-family:'DM Mono',monospace;}
    .cal-conflict-dot{position:absolute;top:4px;right:5px;width:6px;height:6px;background:#f97316;border-radius:50%;}
    .cal-ev{padding:2px 5px;margin-top:2px;border-radius:3px;font-size:9px;line-height:1.4;word-break:break-word;}
    .cal-ev-time{font-weight:700;display:block;font-family:'DM Mono',monospace;}
    .cal-ev-name{display:block;font-size:8.5px;opacity:0.9;}
    .cal-more{font-size:9px;color:#94a3b8;font-style:italic;margin-top:3px;}
    </style>
    """
    st.markdown(cal_css, unsafe_allow_html=True)

    for m in months_to_show:
        month_df = filtered_df[filtered_df["Month"]==m]
        if month_df.empty: continue
        month_name = datetime.date(1900,m,1).strftime('%B')
        cal_weeks = calendar.monthcalendar(2026,m)

        html = f'<div class="cal-wrap"><div class="cal-title">{month_name} 2026</div><div class="cal-grid">'
        for d in WEEKDAYS_DE:
            html += f'<div class="cal-hdr">{d}</div>'

        for week in cal_weeks:
            for i, day in enumerate(week):
                if day == 0:
                    html += '<div class="cal-cell-empty"></div>'; continue
                is_today  = (datetime.date(2026,m,day)==today)
                is_wknd   = (i>=5)
                date_str  = datetime.date(2026,m,day).strftime("%d.%m.%Y")
                has_conflict = date_str in conflict_dates

                cls = "cal-cell"
                if is_today:    cls += " cal-cell-today"
                elif is_wknd:   cls += " cal-cell-wknd"
                if has_conflict: cls += " cal-cell-conflict"

                day_df = month_df[month_df["Day"]==day].copy()
                day_df["_s"] = day_df["Zeit"].str.replace("Zeit unbekannt","00:00",regex=False)
                day_df = day_df.sort_values("_s")

                conflict_dot = '<div class="cal-conflict-dot" title="Konflikt erkannt"></div>' if has_conflict else ''
                html += f'<div class="{cls}">{conflict_dot}<div class="cal-daynum">{day:02d}</div>'

                for idx, (_,row) in enumerate(day_df.iterrows()):
                    if idx==3:
                        html += f'<div class="cal-more">+{len(day_df)-3} weitere</div>'; break
                    color = get_color(row["Event"])
                    hx = color.lstrip("#")
                    r2,g2,b2 = int(hx[0:2],16),int(hx[2:4],16),int(hx[4:6],16)
                    txt = "#000" if (r2*299+g2*587+b2*114)/1000>140 else "#fff"
                    zeit = str(row["Zeit"])[:11]
                    name = str(row["Event"])[:40]
                    html += (f'<div class="cal-ev" style="background:{color};color:{txt};">'
                             f'<span class="cal-ev-time">{zeit}</span>'
                             f'<span class="cal-ev-name">{name}</span></div>')
                html += '</div>'
        html += '</div></div>'
        st.markdown(html, unsafe_allow_html=True)

        st.markdown(f"**Tagesdetails {month_name}:**")
        for day in range(1,32):
            try: datetime.date(2026,m,day)
            except ValueError: break
            day_df = month_df[month_df["Day"]==day].copy()
            if day_df.empty: continue
            day_df["_s"] = day_df["Zeit"].str.replace("Zeit unbekannt","00:00",regex=False)
            day_df = day_df.sort_values("_s")
            weekday_name = datetime.date(2026,m,day).strftime("%a")
            date_str = datetime.date(2026,m,day).strftime("%d.%m.%Y")
            conflict_indicator = " ️" if date_str in conflict_dates else ""
            with st.expander(f"{weekday_name}, {date_str} —  {len(day_df)} Event(s){conflict_indicator}"):
                detail_html = ""
                for _,row in day_df.iterrows():
                    color = get_color(row["Event"])
                    bem_val = row.get("Bemerkungen","")
                    bem = str(bem_val) if pd.notna(bem_val) and str(bem_val).strip() else ""
                    detail_html += (
                        f'<div style="border-left:4px solid {color};padding:8px 14px;margin-bottom:8px;'
                        f'background:#fafbfc;border-radius:0 6px 6px 0;">'
                        f'<b style="font-size:13px;font-family:DM Mono,monospace;">{row["Zeit"]}</b>'
                        f'&nbsp;—&nbsp;<b style="color:#0f1923;">{row["Event"]}</b><br>'
                        f'<span style="font-size:11px;color:#64748b;"> {row["Personen"]}</span><br>'
                        f'<span style="font-size:11px;color:#64748b;"> {row["Ort"]}</span>'
                        + (f'<br><span style="font-size:10px;color:#94a3b8;font-style:italic;"> {bem}</span>' if bem else "")
                        + '</div>'
                    )
                    # Individual iCal
                    single_ical = make_ical(pd.DataFrame([row]))
                st.markdown(detail_html, unsafe_allow_html=True)
        st.markdown("---")

# ---- iCal TAB ----
with tab_ical:
    st.markdown('<div class="section-header">iCal / Outlook Export</div>', unsafe_allow_html=True)
    st.markdown("Exportiere Events direkt in **Outlook, Apple Calendar oder Google Calendar**.")
    st.info(" Tipp: Verwende die Filter in der Sidebar, um nur bestimmte Events, Personen oder Räume zu exportieren — dann kannst du gezielt z.B. alle ECMO-Schulungen oder Events einer Person exportieren.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("####  Gefilterte Events")
        st.markdown(f"**{len(filtered_df)} Events** entsprechen den aktuellen Filtern.")
        ical_filtered = make_ical(filtered_df)
        st.download_button(" Gefilterte Events als .ics herunterladen",
                           ical_filtered.encode("utf-8"),
                           "kim_schedule_filtered.ics", "text/calendar")
        st.markdown("*Importiere die .ics Datei in Outlook: Datei → Öffnen & Exportieren → Importieren/Exportieren → iCalendar importieren*")

    with col_b:
        st.markdown("####  Alle Events (ganzes Jahr)")
        st.markdown(f"**{len(df)} Events** total in 2026.")
        ical_all = make_ical(df)
        st.download_button(" Alle Events als .ics herunterladen",
                           ical_all.encode("utf-8"),
                           "kim_schedule_2026_all.ics", "text/calendar")

    st.markdown("---")
    st.markdown("####  Einzelne Events exportieren")
    st.caption("Wähle einen spezifischen Event-Typ für den Export:")
    event_types = sorted(df["Event"].unique())
    selected_event_export = st.selectbox("Event für Export auswählen", event_types)
    single_event_df = df[df["Event"]==selected_event_export]
    st.markdown(f"**{len(single_event_df)} Termine** für *{selected_event_export}*")
    ical_single = make_ical(single_event_df)
    st.download_button(f" '{selected_event_export[:30]}' als .ics",
                       ical_single.encode("utf-8"),
                       f"kim_{selected_event_export[:20].replace(' ','_')}.ics", "text/calendar")

# ---- LIST TAB ----
with tab_list:
    st.markdown('<div class="section-header">Listenansicht</div>', unsafe_allow_html=True)
    display_df = filtered_df.copy()
    display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")
    cols_show = [c for c in ["Tag","Datum","Zeit","Event","Personen","Ort","Bemerkungen"] if c in display_df.columns]
    st.dataframe(display_df[cols_show], use_container_width=True, hide_index=True)

# ---- BAR TAB ----
with tab_bar:
    import plotly.graph_objects as go
    st.markdown('<div class="section-header">Jahresübersicht: Events pro Tag (gestapelt)</div>', unsafe_allow_html=True)
    full_range = pd.date_range(start="2026-01-01", end="2026-12-31")
    pivot = filtered_df.groupby(["Datum","Kategorie"]).size().unstack(fill_value=0)
    pivot = pivot.reindex(full_range, fill_value=0)
    fig_bar = go.Figure()
    for cat in pivot.columns:
        fig_bar.add_trace(go.Bar(x=pivot.index, y=pivot[cat], name=cat,
                                  marker_color=BASE_COLORS.get(cat,"#B0BEC5")))
    fig_bar.update_layout(
        barmode="stack", height=420,
        xaxis=dict(title="Datum", tickformat="%b", dtick="M1", gridcolor="#f1f5f9"),
        yaxis=dict(title="Anzahl Events", gridcolor="#f1f5f9"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.4, xanchor="center", x=0.5),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10,r=10,t=10,b=20),
    )
    st.plotly_chart(fig_bar, use_container_width=True)
