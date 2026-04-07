import streamlit as st
import pandas as pd
import datetime

st.set_page_config(layout="wide")


COLOR_PALETTE = [  
    "#4FC3F7", "#81C784", "#FFD54F", "#FF8A65",
    "#BA68C8", "#90A4AE", "#F06292", "#AED581",
    "#7986CB", "#4DB6AC", "#FFB74D", "#A1887F",
    "#64B5F6", "#DCE775", "#9575CD", "#E57373"
]

 
st.title("TEST ONLY! ICU Schedule Dashboard 2026")

# =========================
# LOAD DATA
# =========================

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # --- CLEAN DATA ---
    df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors="coerce")

    df = df.dropna(subset=["Datum"])

    # Extract useful columns
    df["Month"] = df["Datum"].dt.month
    df["Day"] = df["Datum"].dt.day
    df["Date_str"] = df["Datum"].dt.strftime("%d.%m.%Y")

    # =========================
    # SIDEBAR FILTERS
    # =========================

    st.sidebar.header("Filters")

    month_options = ["All Year"] + sorted(df["Month"].unique())

    selected_month = st.sidebar.selectbox(
        "Select Month",
        month_options,
        index=0,
        format_func=lambda x: x if x == "All Year" else datetime.date(1900, x, 1).strftime('%B')
    )

    event_filter = st.sidebar.multiselect(
        "Event type",
        df["Event"].dropna().unique()
    )

    person_filter = st.sidebar.multiselect(
        "Person",
        df["Personen"].dropna().unique()
    )

    # Apply filters
    if selected_month == "All Year":
        filtered_df = df.copy()
    else:
        filtered_df = df[df["Month"] == selected_month]

    if event_filter:
        filtered_df = filtered_df[filtered_df["Event"].isin(event_filter)]

    if person_filter:
        filtered_df = filtered_df[filtered_df["Personen"].isin(person_filter)]

    # =========================
    # COLOR MAPPING
    # =========================
    
    unique_events = sorted(filtered_df["Event"].dropna().unique())   
    
    event_color_map = {
        event: COLOR_PALETTE[i % len(COLOR_PALETTE)]
        for i, event in enumerate(unique_events)
    }

    
    # =========================
    # CALENDAR VIEW
    # =========================
    with st.expander("Calendar View", expanded=False):   
    # dont need? st.header("Calendar View")

    # dont need? days = sorted(filtered_df["Day"].unique())

    for day in range(1, 32):
        day_df = filtered_df[filtered_df["Day"] == day]

        if not day_df.empty:
            st.markdown(f"### {day:02d}.{selected_month:02d}.2026")

            for _, row in day_df.iterrows():
                color = get_color(row["Event"])

                st.markdown(
                    f"""
                    <div style="
                        background-color:{color};
                        padding:8px;
                        border-radius:8px;
                        margin-bottom:5px;
                    ">
                        <b>{row['Zeit']}</b> — {row['Event']}<br>
                         {row['Personen'] if pd.notna(row['Personen']) else '-'}<br>
                         {row['Ort'] if pd.notna(row['Ort']) else '-'}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

    # =========================
    # TABLE VIEW
    # =========================

    st.header("List View")

    st.dataframe(
        filtered_df.sort_values("Datum")[[
            "Datum", "Zeit", "Event", "Personen", "Ort", "Bemerkungen"
        ]],
        use_container_width=True
    )

    # =========================
    # TIMELINE OVERVIEW
    # =========================

    st.header("Year Overview")

    df["HasEvent"] = 1

    timeline = filtered_df.groupby("Datum").size().reset_index(name="Count")

    timeline["Date_str"] = timeline["Datum"].dt.strftime("%d.%m")
    
    st.bar_chart(
        timeline.set_index("Date_str")["Count"]
    )
