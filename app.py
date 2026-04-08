import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys

st.set_page_config(layout="wide")
st.title("ICU Schedule Dashboard 2026")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

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
        if "epic" in e:
            return "EPIC"
        elif "ecmo" in e or "lvad" in e or "impella" in e or "assist" in e or "prismax" in e or "beatmung" in e:
            return "Geräte & Beatmung"
        elif "workshop" in e:
            return "Workshop"
        elif "schulung" in e or "kurs" in e or "basiskurs" in e or "aufbaukurs" in e or "refresher" in e:
            return "Schulung/Kurs"
        elif "sitzung" in e or "fachgruppe" in e or "fg " in e or "superuser" in e or "fachforum" in e:
            return "Sitzung"
        elif "einführung" in e or "einblick" in e:
            return "Einführung"
        elif "lernwerkstatt" in e or "repe" in e or "simulation" in e or "kimsim" in e:
            return "Lernwerkstatt"
        elif "führungsdialog" in e or "austausch" in e:
            return "Führung/Austausch"
        elif "icu" in e:
            return "ICU"
        elif "kommunikation" in e or "aggressions" in e:
            return "Kommunikation"
        elif "planung" in e or "bürotag" in e:
            return "Planung"
        else:
            return "Sonstiges"

    BASE_COLORS = {
        "EPIC":                "#4FC3F7",
        "Geräte & Beatmung":   "#E57373",
        "Workshop":            "#FFD54F",
        "Schulung/Kurs":       "#81C784",
        "Sitzung":             "#FF8A65",
        "Einführung":          "#4DB6AC",
        "Lernwerkstatt":       "#F48FB1",
        "Führung/Austausch":   "#CE93D8",
        "ICU":                 "#90CAF9",
        "Kommunikation":       "#FFCC80",
        "Planung":             "#A5D6A7",
        "Sonstiges":           "#B0BEC5",
    }

    def generate_shades(base_hex, n):
        base_hex = base_hex.lstrip('#')
        r = int(base_hex[0:2], 16) / 255
        g = int(base_hex[2:4], 16) / 255
        b = int(base_hex[4:6], 16) / 255
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        shades = []
        for i in range(n):
            new_l = max(0.35, min(0.75, l + (i - n / 2) * 0.08))
            r2, g2, b2 = colorsys.hls_to_rgb(h, new_l, s)
            shades.append('#{:02x}{:02x}{:02x}'.format(int(r2 * 255), int(g2 * 255), int(b2 * 255)))
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
    # FILTER SIDEBAR
    # =========================

    st.sidebar.header("Filter")

    month_options = ["Ganzes Jahr"] + sorted(df["Month"].unique())

    selected_month = st.sidebar.selectbox(
        "Monat auswählen",
        month_options,
        index=0,
        format_func=lambda x: x if x == "Ganzes Jahr" else datetime.date(1900, x, 1).strftime('%B')
    )

    event_filter = st.sidebar.multiselect("Event-Typ", sorted(df["Event"].unique()))
    person_filter = st.sidebar.multiselect("Person", sorted(df["Personen"].unique()))
    room_filter = st.sidebar.multiselect("Raum", sorted(df["Ort"].unique()))

    filtered_df = df.copy()

    if selected_month != "Ganzes Jahr":
        filtered_df = filtered_df[filtered_df["Month"] == selected_month]
    if event_filter:
        filtered_df = filtered_df[filtered_df["Event"].isin(event_filter)]
    if person_filter:
        filtered_df = filtered_df[filtered_df["Personen"].isin(person_filter)]
    if room_filter:
        filtered_df = filtered_df[filtered_df["Ort"].isin(room_filter)]

    filtered_df = filtered_df.sort_values("Datum")

    # =========================
    # BUBBLES ÜBERSICHT
    # =========================

    with st.expander("Übersicht: Events pro Tag (Bubbles)", expanded=True):
        import plotly.graph_objects as go
        import numpy as np

        bubble_df = filtered_df.groupby(["Datum", "Kategorie"]).size().reset_index(name="Anzahl")

        fig_bubble = go.Figure()

        for cat, color in BASE_COLORS.items():
            cat_df = bubble_df[bubble_df["Kategorie"] == cat]
            if cat_df.empty:
                continue
            fig_bubble.add_trace(go.Scatter(
                x=cat_df["Datum"],
                y=cat_df["Kategorie"],
                mode="markers",
                marker=dict(
                    size=cat_df["Anzahl"] * 14,
                    color=color,
                    opacity=0.85,
                    line=dict(width=1, color="white"),
                    sizemode="diameter",
                ),
                name=cat,
                text=[f"{row['Datum'].strftime('%d.%m.%Y')}: {row['Anzahl']} Events" for _, row in cat_df.iterrows()],
                hovertemplate="%{text}<extra></extra>",
            ))

        fig_bubble.update_layout(
            height=500,
            xaxis=dict(
                range=["2026-01-01", "2026-12-31"],
                title="Datum",
                tickformat="%b",
                dtick="M1",
            ),
            yaxis=dict(title=""),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=20, b=20),
        )

        st.plotly_chart(fig_bubble, use_container_width=True)

    # =========================
    # KALENDER ANSICHT
    # =========================

    with st.expander("Kalenderansicht", expanded=False):

        months_to_show = range(1, 13) if selected_month == "Ganzes Jahr" else [selected_month]
        WEEKDAYS_DE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
        today = datetime.date.today()

        cal_css = """
        <style>
        .cal-wrap { font-family: sans-serif; margin-bottom: 32px; }
        .cal-title { font-size: 18px; font-weight: 700; margin-bottom: 8px; color: #222; }
        .cal-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 3px;
        }
        .cal-hdr {
            text-align: center;
            font-weight: 700;
            font-size: 11px;
            padding: 5px 0;
            border-bottom: 2px solid #ccc;
            color: #555;
            background: #f5f5f5;
            border-radius: 4px 4px 0 0;
        }
        .cal-cell {
            border: 1px solid #e2e2e2;
            border-radius: 5px;
            padding: 4px;
            min-height: 95px;
            background: #fff;
            box-sizing: border-box;
        }
        .cal-cell-wknd { background: #f7f7f7; }
        .cal-cell-today { border: 2px solid #f0a500; background: #fffbee; }
        .cal-cell-empty { border: none; background: transparent; min-height: 95px; }
        .cal-daynum { font-weight: 700; font-size: 11px; color: #333; margin-bottom: 2px; }
        .cal-ev {
            padding: 2px 4px;
            margin-top: 2px;
            border-radius: 4px;
            font-size: 9.5px;
            line-height: 1.35;
            word-break: break-word;
        }
        .cal-ev-time { font-weight: 700; display: block; }
        .cal-ev-name { display: block; font-size: 8.5px; opacity: 0.92; }
        .cal-more { font-size: 9px; color: #888; font-style: italic; margin-top: 2px; }
        </style>
        """
        st.markdown(cal_css, unsafe_allow_html=True)

        for m in months_to_show:
            month_df = filtered_df[filtered_df["Month"] == m]
            if month_df.empty:
                continue

            month_name = datetime.date(1900, m, 1).strftime('%B')
            cal_weeks = calendar.monthcalendar(2026, m)

            # Build full month HTML in one shot
            html = f'<div class="cal-wrap"><div class="cal-title">{month_name} 2026</div><div class="cal-grid">'

            for d in WEEKDAYS_DE:
                html += f'<div class="cal-hdr">{d}</div>'

            for week in cal_weeks:
                for i, day in enumerate(week):
                    if day == 0:
                        html += '<div class="cal-cell-empty"></div>'
                        continue

                    is_today = (datetime.date(2026, m, day) == today)
                    is_wknd = (i >= 5)

                    if is_today:
                        cls = "cal-cell cal-cell-today"
                    elif is_wknd:
                        cls = "cal-cell cal-cell-wknd"
                    else:
                        cls = "cal-cell"

                    day_df = month_df[month_df["Day"] == day].copy()
                    day_df["_s"] = day_df["Zeit"].str.replace("Zeit unbekannt", "00:00", regex=False)
                    day_df = day_df.sort_values("_s")

                    html += f'<div class="{cls}"><div class="cal-daynum">{day:02d}</div>'

                    for idx, (_, row) in enumerate(day_df.iterrows()):
                        if idx == 3:
                            rem = len(day_df) - 3
                            html += f'<div class="cal-more">+{rem} weitere</div>'
                            break
                        color = get_color(row["Event"])
                        hx = color.lstrip("#")
                        r2,g2,b2 = int(hx[0:2],16), int(hx[2:4],16), int(hx[4:6],16)
                        txt = "#000" if (r2*299+g2*587+b2*114)/1000 > 140 else "#fff"
                        zeit = str(row["Zeit"])[:11]
                        name = str(row["Event"])[:40]
                        html += (
                            f'<div class="cal-ev" style="background:{color};color:{txt};">'
                            f'<span class="cal-ev-time">{zeit}</span>'
                            f'<span class="cal-ev-name">{name}</span>'
                            f'</div>'
                        )

                    html += '</div>'  # close cal-cell

            html += '</div></div>'  # close grid + wrap
            st.markdown(html, unsafe_allow_html=True)

            # Detail expanders per day, listed below the calendar grid
            st.markdown(f"**Tagesdetails {month_name}** — alle Events aufklappbar:")
            for day in range(1, 32):
                try:
                    datetime.date(2026, m, day)
                except ValueError:
                    break
                day_df = month_df[month_df["Day"] == day].copy()
                if day_df.empty:
                    continue
                day_df["_s"] = day_df["Zeit"].str.replace("Zeit unbekannt", "00:00", regex=False)
                day_df = day_df.sort_values("_s")
                weekday_name = datetime.date(2026, m, day).strftime("%a")
                with st.expander(f"{weekday_name}, {day:02d}.{m:02d}.2026  —  {len(day_df)} Event(s)"):
                    detail_html = ""
                    for _, row in day_df.iterrows():
                        color = get_color(row["Event"])
                        bem_val = row.get("Bemerkungen", "")
                        bem = str(bem_val) if pd.notna(bem_val) and str(bem_val).strip() else ""
                        detail_html += (
                            f'<div style="border-left:4px solid {color};padding:6px 12px;'
                            f'margin-bottom:8px;background:#fafafa;border-radius:4px;">'
                            f'<b style="font-size:13px;">{row["Zeit"]}</b> &nbsp;—&nbsp; <b>{row["Event"]}</b><br>'
                            f'<span style="font-size:12px;color:#555;">Person: {row["Personen"]}</span><br>'
                            f'<span style="font-size:12px;color:#555;">Raum: {row["Ort"]}</span>'
                            + (f'<br><span style="font-size:11px;color:#999;">Bemerkung: {bem}</span>' if bem else "")
                            + '</div>'
                        )
                    st.markdown(detail_html, unsafe_allow_html=True)

            st.markdown("---")

    # =========================
    # LISTENANSICHT
    # =========================

    with st.expander("Listenansicht", expanded=False):
        display_df = filtered_df.copy()
        display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")
        cols_show = [c for c in ["Tag", "Datum", "Zeit", "Event", "Personen", "Ort", "Bemerkungen"] if c in display_df.columns]
        st.dataframe(
            display_df[cols_show],
            use_container_width=True,
            hide_index=True,
        )

    # =========================
    # JAHRESÜBERSICHT STACKED BAR
    # =========================

    with st.expander("Jahresübersicht: Events pro Tag (gestapelt)", expanded=False):
        import plotly.graph_objects as go

        full_range = pd.date_range(start="2026-01-01", end="2026-12-31")

        pivot = filtered_df.groupby(["Datum", "Kategorie"]).size().unstack(fill_value=0)
        pivot = pivot.reindex(full_range, fill_value=0)

        fig_bar = go.Figure()

        for cat in pivot.columns:
            color = BASE_COLORS.get(cat, "#B0BEC5")
            fig_bar.add_trace(go.Bar(
                x=pivot.index,
                y=pivot[cat],
                name=cat,
                marker_color=color,
            ))

        fig_bar.update_layout(
            barmode="stack",
            height=400,
            xaxis=dict(
                title="Datum",
                tickformat="%b",
                dtick="M1",
            ),
            yaxis=dict(title="Anzahl Events"),
            legend=dict(orientation="h", yanchor="bottom", y=-0.4, xanchor="center", x=0.5),
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=10, t=20, b=20),
        )

        st.plotly_chart(fig_bar, use_container_width=True)

    # =========================
    # LEGENDE SIDEBAR
    # =========================

    st.sidebar.markdown("---")
    st.sidebar.markdown("##### Kategorien")

    for cat, color in BASE_COLORS.items():
        count = len(filtered_df[filtered_df["Kategorie"] == cat])
        if count > 0:
            st.sidebar.markdown(f"""
            <div style="display:flex;align-items:center;font-size:11px;margin-bottom:3px;">
                <div style="width:10px;height:10px;background:{color};margin-right:6px;border-radius:2px;flex-shrink:0;"></div>
                {cat} <span style="color:#999;margin-left:4px;">({count})</span>
            </div>
            """, unsafe_allow_html=True)
