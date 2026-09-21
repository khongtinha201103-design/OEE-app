import streamlit as st
import pyodbc
import pandas as pd
import plotly.express as px
from streamlit_autorefresh import st_autorefresh   # pip install streamlit-autorefresh
import io
from datetime import datetime
import openpyxl   # ✅ ensure openpyxl is available for Excel export

# -----------------------------
# 1. Auto-refresh every 60s
# -----------------------------
st_autorefresh(interval=60000, key="refresh")

# -----------------------------
# 2. Load data from SQL Server
# -----------------------------
import pyodbc

# Option A: Using FreeTDS (Most reliable on Linux Streamlit Cloud)
conn = pyodbc.connect(
    'DRIVER={FreeTDS};'
    'SERVER=your_public_db_host_or_ip;'
    'PORT=1433;'
    'DATABASE=SiplaceOIS;'
    'UID=sa;'
    'PWD=Siplace%Sa.1.Pwd;'
    'TDS_Version=8.0;'
)
    )
    query = "SELECT * FROM dbo.v_MachineTimelineReport"
    df = pd.read_sql(query, conn)
    conn.close()
    return df

st.title("📊 Machine Timeline Dashboard")

@st.cache_data
def get_data():
    return load_data()

df = get_data()

# -----------------------------
# 3. Filters
# -----------------------------
machines = df['Machine_ID'].unique()
default_selection = ["1080"] if "1080" in machines else []
if len(machines) > 1:
    default_selection.append(machines[0])

selected_machines = st.multiselect("Select Machines (choose 2)", machines, default=default_selection)

min_date = pd.to_datetime(df['Start_Time']).min().date()
max_date = pd.to_datetime(df['End_Time']).max().date()
date_range = st.date_input("Select Date Range", [min_date, max_date])

df_timeframe = df[
    (pd.to_datetime(df['Start_Time']).dt.date >= date_range[0]) &
    (pd.to_datetime(df['End_Time']).dt.date <= date_range[1])
]

recipes_sorted = (
    df_timeframe[['Recipe','End_Time']]
    .dropna()
    .sort_values("End_Time", ascending=False)
    .drop_duplicates(subset=["Recipe"])
)
recipes = recipes_sorted['Recipe'].tolist()
selected_recipe = st.selectbox("Select Recipe", recipes, index=0) if recipes else None

# -----------------------------
# 4. Custom Colors
# -----------------------------
color_map = {
    "Running": "green",
    "Starved": "lightblue",
    "Interruption": "orange",
    "Fault": "red",
    "Maintenance": "darkblue",
    "Not Scheduled": "gray",
    "Setup": "pink",
    "Blocked": "beige"
}

# -----------------------------
# 5. Timelines stacked vertically with timestamp annotation
# -----------------------------
remarks_data = []

for machine in selected_machines[:2]:
    filtered_df = df_timeframe[
        (df_timeframe['Machine_ID'] == machine) &
        ((df_timeframe['Recipe'] == selected_recipe) if selected_recipe else True)
    ]
    fig = px.timeline(
        filtered_df,
        x_start="Start_Time",
        x_end="End_Time",
        y="Conveyor",
        color="MappedState",
        title=f"Timeline - {machine}",
        color_discrete_map=color_map,
        hover_data=["MappedState","Start_Time","End_Time"]
    )
    st.plotly_chart(fig, use_container_width=True)

    st.write(f"📝 Annotate Timeline for {machine}")
    start_time = st.time_input(f"Start time for {machine}", value=datetime.now().time(), key=f"start_{machine}")
    end_time = st.time_input(f"End time for {machine}", value=datetime.now().time(), key=f"end_{machine}")
    remark = st.text_area(f"Remark for {machine}", key=f"remark_{machine}")

    if remark:
        remarks_data.append({
            "Machine": machine,
            "Start": start_time.strftime("%H:%M:%S"),
            "End": end_time.strftime("%H:%M:%S"),
            "Remark": remark
        })

# -----------------------------
# Helper function: format duration
# -----------------------------
def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.0f} sec"
    elif seconds < 3600:
        return f"{seconds/60:.1f} min"
    else:
        return f"{seconds/3600:.1f} hr"

# -----------------------------
# 6. Pareto charts stacked vertically
# -----------------------------
for machine in selected_machines[:2]:
    filtered_df = df_timeframe[
        (df_timeframe['Machine_ID'] == machine) &
        ((df_timeframe['Recipe'] == selected_recipe) if selected_recipe else True)
    ]

    # --- Main State Pareto ---
    downtime_main = (
        filtered_df[filtered_df['MappedState'] != 'Running']
        .groupby("MappedState")['DurationSeconds'].sum()
        .reset_index()
        .sort_values("DurationSeconds", ascending=False)
    )
    downtime_main["FormattedDuration"] = downtime_main["DurationSeconds"].apply(format_duration)

    fig_main = px.bar(
        downtime_main,
        x="MappedState",
        y="DurationSeconds",
        title=f"Main State Pareto - {machine}",
        labels={"DurationSeconds": "Duration"},
        color="MappedState",
        color_discrete_map=color_map,
        text="FormattedDuration"
    )
    fig_main.update_traces(textposition="outside")
    fig_main.update_yaxes(
        tickvals=downtime_main["DurationSeconds"],
        ticktext=downtime_main["FormattedDuration"]
    )
    st.plotly_chart(fig_main, use_container_width=True)

    # --- Technical State Pareto ---
    main_states = downtime_main['MappedState'].unique().tolist()
    selected_main_state = st.selectbox(f"Filter Technical Pareto ({machine})", main_states, key=f"tech_{machine}")

    downtime_tech_filtered = (
        filtered_df[(filtered_df['MappedState'] == selected_main_state)]
        .groupby("TechnicalState")['DurationSeconds'].sum()
        .reset_index()
        .sort_values("DurationSeconds", ascending=False)
    )
    downtime_tech_filtered["FormattedDuration"] = downtime_tech_filtered["DurationSeconds"].apply(format_duration)

    fig_tech = px.bar(
        downtime_tech_filtered,
        x="TechnicalState",
        y="DurationSeconds",
        title=f"Technical State Pareto - {machine} ({selected_main_state})",
        labels={"DurationSeconds": "Duration"},
        color="TechnicalState",
        text="FormattedDuration"
    )
    fig_tech.update_traces(textposition="outside")
    fig_tech.update_yaxes(
        tickvals=downtime_tech_filtered["DurationSeconds"],
        ticktext=downtime_tech_filtered["FormattedDuration"]
    )
    st.plotly_chart(fig_tech, use_container_width=True)

# -----------------------------
# 7. Export Remarks to Excel (using openpyxl)
# -----------------------------
if remarks_data:
    remarks_df = pd.DataFrame(remarks_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        remarks_df.to_excel(writer, index=False, sheet_name="Remarks")
    st.download_button(
        label="Download Remarks as Excel",
        data=buffer.getvalue(),
        file_name="timeline_remarks.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
