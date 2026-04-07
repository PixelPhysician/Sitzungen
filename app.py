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

    with st.expander("Übersicht: Events pro Tag", expanded=True):
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

        for m in months_to_show:
            month_df = filtered_df[filtered_df["Month"] == m]
            if month_df.empty:
                continue

            month_name = datetime.date(1900, m, 1).strftime('%B')
            st.markdown(f"### {month_name} 2026")

            header_cols = st.columns(7)
            for i, d in enumerate(WEEKDAYS_DE):
                header_cols[i].markdown(
                    f"<div style='text-align:center;font-weight:bold;font-size:13px;padding:4px;border-bottom:2px solid #ddd;'>{d}</div>",
                    unsafe_allow_html=True
                )

            cal = calendar.monthcalendar(2026, m)

            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day == 0:
                        cols[i].markdown(
                            "<div style='min-height:80px;'></div>",
                            unsafe_allow_html=True
                        )
                        continue

                    day_df = month_df[month_df["Day"] == day].copy()
                    day_df["Zeit_sort"] = day_df["Zeit"].str.replace("Zeit unbekannt", "00:00")
                    day_df = day_df.sort_values("Zeit_sort")

                    today = datetime.date.today()
                    is_today = (datetime.date(2026, m, day) == today)
                    is_weekend = (i >= 5)

                    bg_day = "#fff8e1" if is_today else ("#f9f9f9" if is_weekend else "white")
                    border = "2px solid #F9A825" if is_today else "1px solid #e0e0e0"

                    visible = day_df.head(3)
                    hidden = day_df.iloc[3:]

                    events_html = ""
                    for _, row in visible.iterrows():
                        color = get_color(row["Event"])
                        zeit = str(row['Zeit'])[:11]
                        events_html += f"""
                        <div style="
                            background:{color};
                            padding:3px 5px;
                            margin-top:3px;
                            border-radius:5px;
                            font-size:10px;
                            line-height:1.3;
                            overflow:hidden;
                        ">
                            <span style="font-weight:600;">{zeit}</span><br>
                            <span style="font-size:9px;">{str(row['Event'])[:35]}</span>
                        </div>"""

                    if len(hidden) > 0:
                        events_html += f"""
                        <div style="font-size:10px;margin-top:3px;color:#777;font-style:italic;">
                            +{len(hidden)} weitere
                        </div>"""

                    html = f"""
                    <div style="
                        border:{border};
                        border-radius:6px;
                        padding:5px;
                        min-height:90px;
                        background:{bg_day};
                        font-size:12px;
                    ">
                        <div style="font-weight:bold;font-size:12px;color:#333;">{day:02d}</div>
                        {events_html}
                    </div>"""

                    cols[i].markdown(html, unsafe_allow_html=True)

                    if len(day_df) > 0:
                        with cols[i].expander(f"Details {day:02d}.{m:02d}"):
                            for _, row in day_df.iterrows():
                                color = get_color(row["Event"])
                                st.markdown(f"""
                                <div style="border-left:4px solid {color};padding:6px 10px;margin-bottom:6px;background:#fafafa;border-radius:4px;">
                                    <b>{row['Zeit']}</b> — {row['Event']}<br>
                                    <span style="font-size:12px;color:#555;"> {row['Personen']}</span><br>
                                    <span style="font-size:12px;color:#555;"> {row['Ort']}</span>
                                </div>
                                """, unsafe_allow_html=True)

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
