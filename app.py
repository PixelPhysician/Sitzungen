import streamlit as st
import pandas as pd
import datetime
import calendar
import colorsys

st.set_page_config(layout="wide")

st.title("TEST ONLY! ICU Schedule Dashboard 2026")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    # =========================
    # CLEAN DATA
    # =========================

    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["Datum"])

    df["Zeit"] = df["Zeit"].fillna("time unknown")
    df["Personen"] = df["Personen"].fillna("person unknown")
    df["Ort"] = df["Ort"].fillna("room unknown")
    df["Event"] = df["Event"].fillna("topic unknown")

    df["Month"] = df["Datum"].dt.month
    df["Day"] = df["Datum"].dt.day
    df["Weekday"] = df["Datum"].dt.strftime("%a")

    # =========================
    # CATEGORY COLORS
    # =========================

    def get_category(event):
        e = str(event).lower()

        if "epic" in e:
            return "epic"
        elif "icu" in e:
            return "icu"
        elif "workshop" in e:
            return "workshop"
        elif "schulung" in e or "kurs" in e:
            return "training"
        elif "sitzung" in e:
            return "meeting"
        elif "team" in e:
            return "team"
        elif "einführung" in e:
            return "intro"
        elif "ecmo" in e or "lvad" in e or "impella" in e:
            return "device"
        else:
            return "other"

    BASE_COLORS = {
        "epic": "#4FC3F7",
        "icu": "#E57373",
        "workshop": "#FFD54F",
        "training": "#81C784",
        "meeting": "#FF8A65",
        "team": "#BA68C8",
        "intro": "#4DB6AC",
        "device": "#7986CB",
        "other": "#B0BEC5"
    }

    def generate_shades(base_hex, n):
        base_hex = base_hex.lstrip('#')
        r = int(base_hex[0:2], 16) / 255
        g = int(base_hex[2:4], 16) / 255
        b = int(base_hex[4:6], 16) / 255

        h, l, s = colorsys.rgb_to_hls(r, g, b)

        shades = []
        for i in range(n):
            new_l = max(0.35, min(0.75, l + (i - n/2) * 0.08))
            r2, g2, b2 = colorsys.hls_to_rgb(h, new_l, s)

            shades.append('#{:02x}{:02x}{:02x}'.format(
                int(r2*255), int(g2*255), int(b2*255)
            ))

        return shades

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
    # FILTERS
    # =========================

    st.sidebar.header("Filters")

    month_options = ["All Year"] + sorted(df["Month"].unique())

    selected_month = st.sidebar.selectbox(
        "Select Month",
        month_options,
        index=0,
        format_func=lambda x: x if x == "All Year" else datetime.date(1900, x, 1).strftime('%B')
    )

    event_filter = st.sidebar.multiselect("Event type", sorted(df["Event"].unique()))
    person_filter = st.sidebar.multiselect("Person", sorted(df["Personen"].unique()))
    room_filter = st.sidebar.multiselect("Room", sorted(df["Ort"].unique()))

    filtered_df = df.copy()

    if selected_month != "All Year":
        filtered_df = filtered_df[filtered_df["Month"] == selected_month]

    if event_filter:
        filtered_df = filtered_df[filtered_df["Event"].isin(event_filter)]

    if person_filter:
        filtered_df = filtered_df[filtered_df["Personen"].isin(person_filter)]

    if room_filter:
        filtered_df = filtered_df[filtered_df["Ort"].isin(room_filter)]

    filtered_df = filtered_df.sort_values("Datum")

    # =========================
    # YEAR OVERVIEW (IMPROVED)
    # =========================

    if selected_month == "All Year":
        with st.expander("Year Overview", expanded=True):

            mode = st.radio(
                "Display Mode",
                ["Colorful", "Grayscale"],
                horizontal=True
            )

            timeline = filtered_df.groupby("Datum").size().reset_index(name="Count")

            # FULL DATE RANGE (important fix)
            full_range = pd.date_range(start="2026-01-01", end="2026-12-31")
            timeline = timeline.set_index("Datum").reindex(full_range, fill_value=0).reset_index()
            timeline.columns = ["Datum", "Count"]

            timeline["Date_str"] = timeline["Datum"].dt.strftime("%d.%m")

            if mode == "Grayscale":
                st.bar_chart(timeline.set_index("Date_str")["Count"])
            else:
                # pseudo-colorful (still clean)
                st.bar_chart(timeline.set_index("Date_str")["Count"])

    # =========================
    # CALENDAR VIEW (ALL MONTHS)
    # =========================

    with st.expander("Calendar View", expanded=False):

        months_to_show = range(1, 13) if selected_month == "All Year" else [selected_month]

        for m in months_to_show:

            month_df = filtered_df[filtered_df["Month"] == m]

            if month_df.empty:
                continue

            st.markdown(f"## {datetime.date(1900, m, 1).strftime('%B')}")

            cal = calendar.monthcalendar(2026, m)

            cols = st.columns(7)
            for i, d in enumerate(["Mo","Di","Mi","Do","Fr","Sa","So"]):
                cols[i].markdown(f"**{d}**")

            for week in cal:
                cols = st.columns(7)

                for i, day in enumerate(week):
                    if day == 0:
                        cols[i].markdown("")
                    else:
                        day_df = month_df[month_df["Day"] == day].copy()

                        day_df["Zeit_sort"] = day_df["Zeit"].str.replace("ganzer Tag","00:00")
                        day_df = day_df.sort_values("Zeit_sort")
 
                        visible_events = day_df.head(3)
                        hidden_events = day_df.iloc[3:]
                        
                        html = f"<b>{day:02d}</b><br>"
                        
                        # --- visible events ---
                        for _, row in visible_events.iterrows():
                            color = get_color(row["Event"])
                        
                            html += f"""
                            <div style="
                                background:{color};
                                padding:4px;
                                margin-top:4px;
                                border-radius:6px;
                                font-size:11px;
                            ">
                                <b>{row['Zeit']}</b><br>
                                {row['Event']}
                            </div>
                            """
                        
                        # --- "more" indicator ---
                        if len(hidden_events) > 0:
                            html += f"""
                            <div style="
                                font-size:11px;
                                margin-top:4px;
                                color:#555;
                            ">
                                +{len(hidden_events)} more
                            </div>
                            """
                        
                        cols[i].markdown(html, unsafe_allow_html=True)
                        
                        # --- expandable full list BELOW calendar ---
                        if len(hidden_events) > 0:
                            with cols[i].expander("Show all"):
                                for _, row in day_df.iterrows():
                                    st.markdown(f"""
                                    **{row['Zeit']}** — {row['Event']}  
                                    {row['Personen']} | {row['Ort']}
                                    """)

    # =========================
    # LIST VIEW
    # =========================

    with st.expander("List View", expanded=False):

        display_df = filtered_df.copy()
        display_df["Datum"] = display_df["Datum"].dt.strftime("%d.%m.%Y")

        st.dataframe(
            display_df[[
                "Weekday","Datum","Zeit","Event","Personen","Ort","Bemerkungen"
            ]],
            use_container_width=True
        )

    # =========================
    # LEGEND
    # =========================

    st.sidebar.markdown("---")
    st.sidebar.markdown("###### Categories")

    for cat, color in BASE_COLORS.items():
        st.sidebar.markdown(f"""
        <div style="display:flex;align-items:center;font-size:11px;margin-bottom:2px;">
            <div style="width:8px;height:8px;background:{color};margin-right:5px;border-radius:2px;"></div>
            {cat}
        </div>
        """, unsafe_allow_html=True)
