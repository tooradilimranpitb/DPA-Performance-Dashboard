import json
import os
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

# --- Page Configuration ---
st.set_page_config(
    page_title="DPA Performance & Target Management Dashboard",
    page_icon="📊",
    layout="wide",
)

# --- Force-scrollable dropdown fix (JS-based, DOM-structure independent) ---
components.html(
    """
    <script>
    (function () {
      function fixDropdowns() {
        try {
          const doc = window.parent.document;
          const boxes = doc.querySelectorAll('ul[role="listbox"], div[role="listbox"]');
          boxes.forEach(function (el) {
            el.style.setProperty('max-height', '260px', 'important');
            el.style.setProperty('overflow-y', 'auto', 'important');
            el.style.setProperty('overflow-x', 'hidden', 'important');

            let node = el.parentElement;
            let hops = 0;
            while (node && hops < 6) {
              const style = window.parent.getComputedStyle(node);
              if (style.overflow === 'hidden' || style.overflowY === 'hidden') {
                node.style.setProperty('overflow', 'visible', 'important');
              }
              if (node.getAttribute && node.getAttribute('data-baseweb') === 'popover') {
                break;
              }
              node = node.parentElement;
              hops += 1;
            }
          });
        } catch (e) {}
      }
      try {
        const doc = window.parent.document;
        const observer = new MutationObserver(fixDropdowns);
        observer.observe(doc.body, { childList: true, subtree: true });
        fixDropdowns();
      } catch (e) {}
    })();
    </script>
    """,
    height=0,
)

# --- Styling & Theme ---
st.markdown("""
    <style>
    .main { background-color: var(--background-color, #f8f9fa); }
    
    [data-testid="stVirtualDropdown"],
    ul[role="listbox"],
    div[role="listbox"] {
        max-height: 260px !important;
        overflow-y: auto !important;
    }

    [data-testid="stVirtualDropdown"]::-webkit-scrollbar,
    ul[role="listbox"]::-webkit-scrollbar,
    div[role="listbox"]::-webkit-scrollbar {
        width: 8px;
    }
    [data-testid="stVirtualDropdown"]::-webkit-scrollbar-track,
    ul[role="listbox"]::-webkit-scrollbar-track,
    div[role="listbox"]::-webkit-scrollbar-track {
        background: transparent;
    }
    [data-testid="stVirtualDropdown"]::-webkit-scrollbar-thumb,
    ul[role="listbox"]::-webkit-scrollbar-thumb,
    div[role="listbox"]::-webkit-scrollbar-thumb {
        background-color: rgba(128, 128, 128, 0.5);
        border-radius: 4px;
    }
    
    div[data-testid="stMetric"] {
        background-color: var(--secondary-background-color, #ffffff);
        padding: 18px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    div[data-testid="stMetric"] label {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: var(--text-color, #374151) !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: var(--primary-color, #1e3a8a) !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- App Header ---
header_col1, header_col2 = st.columns([1, 6])
with header_col1:
  st.image("appLogoIcon.png", width=110)
with header_col2:
  st.title("DPA Performance Analytics & Target Management Dashboard")
st.markdown("---")

DEFAULT_FILE_PATH = "dpa_data.xlsx"
SUPERVISOR_SHEET_NAME = "Supervisor_Mapping"

DEFAULT_SUPERVISOR_MAPPING = {
    "ATTOCK": "SALAMT KHAN",
    "CHAKWAL": "SALAMT KHAN",
    "GUJRAT": "SALAMT KHAN",
    "JEHLUM": "SALAMT KHAN",
    "RAWALPINDI": "SALAMT KHAN",
    "BHAKHAR": "ZULQANAIN HADIER",
    "KHUSHAB": "ZULQANAIN HADIER",
    "MANDI BAHAUDDIN": "ZULQANAIN HADIER",
    "MIANWALI": "ZULQANAIN HADIER",
    "SARGODHA": "ZULQANAIN HADIER",
    "BAHAWALNAGAR": "ZESHAN",
    "MULTAN": "ZESHAN",
    "OKARA": "ZESHAN",
    "PAKPATTAN": "ZESHAN",
    "SAHIWAL": "ZESHAN",
    "BAHAWALPUR": "ZULFQAR ALI",
    "LODHRAN": "ZULFQAR ALI",
    "RAHIM YAR KHAN": "ZULFQAR ALI",
    "RAJANPUR": "ZULFQAR ALI",
    "D G KHAN": "SHAZAD KHOSSA",
    "KHANEWAL": "SHAZAD KHOSSA",
    "LAYYAH": "SHAZAD KHOSSA",
    "MUZAFFARGARH": "SHAZAD KHOSSA",
    "VEHARI": "SHAZAD KHOSSA",
    "CHINIOT": "MUHAMMAD ASIF",
    "FAISALABAD": "MUHAMMAD ASIF",
    "HAFIZABAD": "MUHAMMAD ASIF",
    "JHANG": "MUHAMMAD ASIF",
    "TOBA TEK SINGH": "MUHAMMAD ASIF",
    "FREEDKOT LAHORE": "MUSHAID",
    "SHADRA LAHORE": "MUSHAID",
    "SHEKUPURA": "MUSHAID",
    "NANKANA SAHIB": "MUSHAID",
    "GUJRANWALA": "MUNEEB",
    "KASUR": "MUNEEB",
    "LAHORE DHA": "MUNEEB",
    "NAROWAL": "MUNEEB",
    "OPF LAHORE": "MUNEEB",
    "ALICOMPLEX": "MUNEEB",
    "SIALKOT": "MUNEEB",
}


def load_excel_data(file_path):
  xls = pd.ExcelFile(file_path)
  sheets_data = {}
  for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet)
    sheets_data[sheet] = df
  return sheets_data


def save_hierarchical_to_excel(
    hierarchical_data, supervisor_mapping, file_path=DEFAULT_FILE_PATH
):
  try:
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
      target_year = (
          "2026" if "2026" in hierarchical_data else list(hierarchical_data.keys())[0]
      )
      for sheet_name, df in hierarchical_data[target_year].items():
        if sheet_name != SUPERVISOR_SHEET_NAME:
          df.to_excel(writer, sheet_name=str(sheet_name), index=False)

      sup_df = pd.DataFrame(
          list(supervisor_mapping.items()), columns=["District", "Supervisor"]
      )
      sup_df.to_excel(writer, sheet_name=SUPERVISOR_SHEET_NAME, index=False)
  except Exception as e:
    st.error(f"Error saving data to file: {e}")


# --- Persistent Session State Data Initialization ---
if "hierarchical_data" not in st.session_state:
  if os.path.exists(DEFAULT_FILE_PATH):
    try:
      sheets_data = load_excel_data(DEFAULT_FILE_PATH)
      if SUPERVISOR_SHEET_NAME in sheets_data:
        sup_df = sheets_data.pop(SUPERVISOR_SHEET_NAME)
        st.session_state["supervisor_mapping"] = {
            str(row["District"]).strip().upper(): str(row["Supervisor"]).strip()
            for _, row in sup_df.dropna(subset=["District", "Supervisor"]).iterrows()
        }
      else:
        st.session_state["supervisor_mapping"] = dict(DEFAULT_SUPERVISOR_MAPPING)

      st.session_state["hierarchical_data"] = {"2026": sheets_data}
    except Exception:
      st.session_state["hierarchical_data"] = {
          "2026": {
              "jan-2026": pd.DataFrame(
                  columns=["DISTRICT", "DPA NAME", "1", "2", "3", "TOTAL PAGES"]
              )
          }
      }
      st.session_state["supervisor_mapping"] = dict(DEFAULT_SUPERVISOR_MAPPING)
  else:
    st.session_state["hierarchical_data"] = {
        "2026": {
            "jan-2026": pd.DataFrame(
                columns=["DISTRICT", "DPA NAME", "1", "2", "3", "TOTAL PAGES"]
            )
        }
    }
    st.session_state["supervisor_mapping"] = dict(DEFAULT_SUPERVISOR_MAPPING)
    save_hierarchical_to_excel(
        st.session_state["hierarchical_data"],
        st.session_state["supervisor_mapping"],
    )

if "supervisor_mapping" not in st.session_state:
  st.session_state["supervisor_mapping"] = dict(DEFAULT_SUPERVISOR_MAPPING)

if "is_authorized" not in st.session_state:
  st.session_state["is_authorized"] = False

# --- Sidebar Controls & Authorization / Authentication ---
st.sidebar.header("⚙️ Dashboard Controls")

st.sidebar.markdown("### 🔐 Admin Authentication")
if not st.session_state["is_authorized"]:
  admin_pass = st.sidebar.text_input(
      "Enter Admin Password to Modify/Upload",
      type="password",
      key="admin_password_input",
  )
  if st.sidebar.button("🔓 Authenticate"):
    if admin_pass == "adminpass123":
      st.session_state["is_authorized"] = True
      st.sidebar.success("Successfully authenticated!")
      st.rerun()
    else:
      st.sidebar.error("Incorrect password. Access denied.")
else:
  st.sidebar.success("Status: Authorized (Admin Mode)")
  if st.sidebar.button("🔒 Logout"):
    st.session_state["is_authorized"] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 Year & Month Management")

available_years = list(st.session_state["hierarchical_data"].keys())
selected_year = st.sidebar.selectbox("Select Year", available_years)

available_months = list(
    st.session_state["hierarchical_data"][selected_year].keys()
)

selected_sheet = st.sidebar.selectbox(
    "Select Month / Report Sheet", available_months
)

if st.session_state["is_authorized"]:
  uploaded_file = st.sidebar.file_uploader(
      "Upload New Workbook (.xlsx)", type=["xlsx"]
  )

  if uploaded_file is not None:
    try:
      with open(DEFAULT_FILE_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())

      sheets_data = load_excel_data(DEFAULT_FILE_PATH)
      if SUPERVISOR_SHEET_NAME in sheets_data:
        sup_df = sheets_data.pop(SUPERVISOR_SHEET_NAME)
        st.session_state["supervisor_mapping"] = {
            str(row["District"]).strip().upper(): str(row["Supervisor"]).strip()
            for _, row in sup_df.dropna(subset=["District", "Supervisor"]).iterrows()
        }
      st.session_state["hierarchical_data"]["2026"] = sheets_data
      st.success("Workbook successfully loaded and saved to environment!")
      st.rerun()
    except PermissionError:
      st.sidebar.error(
          "Permission Denied: Please close 'dpa_data.xlsx' in Excel and try"
          " uploading again."
      )

  with st.sidebar.expander("🛠️ Manage Years & Months", expanded=False):
    new_year = st.text_input("New Year (e.g., 2027)", key="input_new_year")
    if st.button("➕ Add Year"):
      if new_year and new_year not in st.session_state["hierarchical_data"]:
        st.session_state["hierarchical_data"][new_year] = {
            "jan-2027": pd.DataFrame(
                columns=["DISTRICT", "DPA NAME", "1", "2", "3", "TOTAL PAGES"]
            )
        }
        save_hierarchical_to_excel(
            st.session_state["hierarchical_data"],
            st.session_state["supervisor_mapping"],
        )
        st.success(f"Added year: {new_year}")
        st.rerun()
      elif new_year in st.session_state["hierarchical_data"]:
        st.warning("Year already exists.")

    st.markdown("---")

    new_month_name = st.text_input(
        "New Month Name (e.g., feb-2026)", key="input_new_month"
    )
    if st.button("➕ Add Month Sheet"):
      if (
          new_month_name
          and new_month_name
          not in st.session_state["hierarchical_data"][selected_year]
      ):
        reference_df = st.session_state["hierarchical_data"][selected_year][
            selected_sheet
        ]
        st.session_state["hierarchical_data"][selected_year][new_month_name] = (
            pd.DataFrame(columns=reference_df.columns)
        )
        save_hierarchical_to_excel(
            st.session_state["hierarchical_data"],
            st.session_state["supervisor_mapping"],
        )
        st.success(f"Added {new_month_name} under {selected_year}")
        st.rerun()
      elif (
          new_month_name
          in st.session_state["hierarchical_data"][selected_year]
      ):
        st.warning("Month sheet already exists in this year.")

    st.markdown("---")
    if len(st.session_state["hierarchical_data"][selected_year]) > 1:
      month_to_delete = st.selectbox(
          "Select Month to Delete", available_months, key="del_month_select"
      )
      if st.button("🗑️ Delete Month Sheet", type="primary", key="btn_del_month"):
        del st.session_state["hierarchical_data"][selected_year][month_to_delete]
        save_hierarchical_to_excel(
            st.session_state["hierarchical_data"],
            st.session_state["supervisor_mapping"],
        )
        st.success(f"Deleted month: {month_to_delete}")
        st.rerun()
    else:
      st.info("At least one month sheet must remain per year.")

  st.sidebar.markdown("---")
  with st.sidebar.expander("🧑‍💼 Manage Supervisor Mapping", expanded=False):
    st.caption(
        "Assign each district to a supervisor. Saved directly into the Excel file"
        f" under sheet `{SUPERVISOR_SHEET_NAME}`."
    )
    mapping_df = pd.DataFrame(
        list(st.session_state["supervisor_mapping"].items()),
        columns=["District", "Supervisor"],
    ).sort_values(by=["Supervisor", "District"]).reset_index(drop=True)

    edited_mapping_df = st.data_editor(
        mapping_df,
        num_rows="dynamic",
        use_container_width=True,
        key="supervisor_mapping_editor",
    )

    if st.button("💾 Save Supervisor Mapping"):
      cleaned = edited_mapping_df.dropna(subset=["District", "Supervisor"])
      new_mapping = {
          str(row["District"]).strip().upper(): str(row["Supervisor"]).strip()
          for _, row in cleaned.iterrows()
          if str(row["District"]).strip() and str(row["Supervisor"]).strip()
      }
      st.session_state["supervisor_mapping"] = new_mapping
      save_hierarchical_to_excel(
          st.session_state["hierarchical_data"], new_mapping
      )
      st.success("Supervisor mapping saved to Excel workbook.")
      st.rerun()
else:
  st.sidebar.info(
      "🔒 File uploads, adding/deleting months, and editing supervisor/data"
      " mappings are locked. Please authenticate above."
  )

raw_df = st.session_state["hierarchical_data"][selected_year][
    selected_sheet
].copy()


# --- Data Cleaning & Preprocessing Pipeline ---
def clean_dpa_dataframe(df):
  if df is None or df.empty:
    return df, pd.DataFrame(), "", "", None, []

  df.columns = [str(c).strip().upper() for c in df.columns]

  name_col = next(
      (c for c in df.columns if "NAME" in c or "DPA" in c),
      df.columns[1] if len(df.columns) > 1 else df.columns[0],
  )
  district_col = next((c for c in df.columns if "DISTRICT" in c), df.columns[0])
  total_col = next(
      (c for c in df.columns if "TOTAL" in c and "PAGE" in c), None
  )

  df_filtered = df.dropna(subset=[name_col])
  df_filtered = df_filtered[
      ~df_filtered[name_col].astype(str).str.contains("NAME|SR|TOTAL", na=False)
  ]

  meta_cols = {name_col, district_col, total_col, "SR NO."}
  date_cols = [
      c for c in df_filtered.columns if c not in meta_cols and c in df_filtered.columns
  ]

  id_vars = [col for col in [district_col, name_col] if col in df_filtered.columns]
  melted = df_filtered.melt(
      id_vars=id_vars,
      value_vars=date_cols,
      var_name="DATE_STR",
      value_name="PAGES",
  )
  melted.rename(
      columns={district_col: "DISTRICT", name_col: "NAME"}, inplace=True
  )

  melted["PAGES"] = pd.to_numeric(melted["PAGES"], errors="coerce").fillna(0)
  melted["PARSED_DATE"] = pd.to_datetime(melted["DATE_STR"], errors="coerce")

  melted = melted[melted["PARSED_DATE"].dt.day_name() != "Sunday"]

  return df_filtered, melted, district_col, name_col, total_col, date_cols


clean_df, melted_df, district_col, name_col, total_col, date_cols = (
    clean_dpa_dataframe(raw_df)
)

# --- Summary Statistics & Metrics ---
st.subheader(
    f"📈 Executive Performance Report — {selected_sheet} {selected_year}"
)

total_scanned_all = (
    melted_df["PAGES"].sum()
    if total_col is None and not melted_df.empty
    else (
        clean_df[total_col].dropna().sum()
        if total_col and not clean_df.empty
        else 0
    )
)
active_dpas = melted_df["NAME"].nunique() if not melted_df.empty else 0
avg_per_dpa = total_scanned_all / active_dpas if active_dpas > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total Scanned Pages", f"{int(total_scanned_all):,}")
col2.metric("Active DPAs", f"{active_dpas}")
col3.metric("Avg Pages / DPA", f"{int(avg_per_dpa):,}")
col4.metric(
    "Working Days Tracked",
    f"{melted_df['PARSED_DATE'].nunique() if not melted_df.empty else 0}",
)
col5.metric(
    "Districts Covered",
    f"{melted_df['DISTRICT'].nunique() if not melted_df.empty else 0}",
)

st.markdown("---")

# --- Tabs ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Performance Dashboard",
    "📅 Daily & Weekly Breakdown",
    "🔮 Current Month Projections",
    "🎯 Target Setting for Next Month",
    "📝 Live Data Sheet Editor",
    "🧑‍💼 Supervisor Performance",
])

with tab1:
  st.subheader("🏆 Enhanced DPA Leaderboard & Performance Evaluation")
  if not melted_df.empty:
    dpa_summary = (
        melted_df.groupby(["DISTRICT", "NAME"])["PAGES"]
        .agg(["sum", "mean", "max", "count"])
        .reset_index()
    )
    dpa_summary.columns = [
        "District",
        "DPA Name",
        "Total Pages",
        "Daily Average",
        "Max Daily Peak",
        "Active Days",
    ]
    dpa_summary = dpa_summary.sort_values(by="Total Pages", ascending=False)
    mean_output = dpa_summary["Total Pages"].mean()
    dpa_summary["Performance Status"] = dpa_summary["Total Pages"].apply(
        lambda x: (
            "🌟 Exceptional"
            if x >= mean_output * 1.25
            else ("✅ On Track" if x >= mean_output * 0.75 else "⚠️ Needs Support")
        )
    )
    st.dataframe(
        dpa_summary.iloc[:, 1:].style.format({
            "Total Pages": "{:,.0f}",
            "Daily Average": "{:,.1f}",
            "Max Daily Peak": "{:,.0f}",
        }),
        use_container_width=True,
    )
    exec_csv = dpa_summary.iloc[:, 1:].to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Executive Summary Report (CSV)",
        data=exec_csv,
        file_name=(
            f"executive_summary_report_{selected_sheet}_{selected_year}.csv"
        ),
        mime="text/csv",
    )
    st.markdown("---")
    fig_bar = px.bar(
        dpa_summary,
        x="DPA Name",
        y="Total Pages",
        color="District",
        title=(
            f"Comparative Output per DPA for {selected_sheet}"
            f" {selected_year} (Sundays Excluded)"
        ),
        text_auto=".2s",
        template="plotly_white",
    )
    fig_bar.update_layout(xaxis={"categoryorder": "total descending"})
    st.plotly_chart(fig_bar, use_container_width=True)
  else:
    st.info("No data available to display leaderboard.")

with tab2:
  st.subheader("📅 Periodic Performance Breakdown (Daily & Fortnightly)")
  if not melted_df.empty:
    melted_df["WEEK"] = melted_df["PARSED_DATE"].dt.isocalendar().week
    melted_df["FORTNIGHT"] = melted_df["PARSED_DATE"].dt.day.apply(
        lambda x: "First Half (1-15)" if x <= 15 else "Second Half (16-31)"
    )
    daily_trend = (
        melted_df.groupby("PARSED_DATE")["PAGES"].sum().reset_index()
    )
    fig_daily = px.line(
        daily_trend,
        x="PARSED_DATE",
        y="PAGES",
        markers=True,
        title=(
            f"Department-Wide Daily Scanning Trend — {selected_sheet}"
            f" {selected_year} (Sundays Excluded)"
        ),
        template="plotly_white",
    )
    st.plotly_chart(fig_daily, use_container_width=True)
    fortnight_summary = (
        melted_df.groupby(["NAME", "FORTNIGHT"])["PAGES"]
        .sum()
        .unstack()
        .fillna(0)
    )
    st.markdown("#### Fortnightly Output Distribution")
    st.dataframe(fortnight_summary, use_container_width=True)
  else:
    st.info("No trend data available.")

with tab3:
  st.subheader("🔮 Current Month Projection & Pace Analysis")
  if not melted_df.empty:
    valid_days_passed = melted_df[melted_df["PAGES"] > 0][
        "PARSED_DATE"
    ].nunique()
    projection_df = (
        melted_df.groupby(["DISTRICT", "NAME"])["PAGES"]
        .agg(Actual_So_Far="sum", Active_Days=lambda x: (x > 0).sum())
        .reset_index()
    )
    if active_dpas > 0 and valid_days_passed > 0:
      projection_df["Daily_Run_Rate"] = (
          projection_df["Actual_So_Far"] / projection_df["Active_Days"]
      )
      remaining_days_est = np.maximum(0, 22 - projection_df["Active_Days"])
      projection_df["Projected_Month_Total"] = projection_df[
          "Actual_So_Far"
      ] + (projection_df["Daily_Run_Rate"] * remaining_days_est)
      st.dataframe(
          projection_df.iloc[:, 1:].style.format({
              "Daily_Run_Rate": "{:.1f}",
              "Projected_Month_Total": "{:.0f}",
          }),
          use_container_width=True,
      )
      fig_proj = px.bar(
          projection_df,
          x="NAME",
          y=["Actual_So_Far", "Projected_Month_Total"],
          barmode="group",
          title="Actual Output vs Projected Month-End Output",
          template="plotly_white",
      )
      st.plotly_chart(fig_proj, use_container_width=True)
    else:
      st.info("Insufficient daily activity data to calculate projections.")
  else:
    st.info("No data available for projections.")

with tab4:
  st.subheader("🎯 Automated Target Setting for Upcoming Month")
  if not melted_df.empty:
    growth_factor = (
        st.slider("Target Growth / Adjustment Multiplier (%)", -20, 50, 10, 5)
        / 100.0
    )
    target_df = (
        melted_df.groupby(["DISTRICT", "NAME"])["PAGES"].sum().reset_index()
    )
    target_df.rename(columns={"PAGES": "Current_Month_Actual"}, inplace=True)
    target_df["Recommended_Next_Month_Target"] = (
        target_df["Current_Month_Actual"] * (1 + growth_factor)
    ).round(-2)
    st.dataframe(
        target_df.iloc[:, 1:].style.format({
            "Current_Month_Actual": "{:,.0f}",
            "Recommended_Next_Month_Target": "{:,.0f}",
        }),
        use_container_width=True,
    )
    csv_data = target_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Next Month Targets CSV",
        data=csv_data,
        file_name=f"next_month_targets_{selected_sheet}_{selected_year}.csv",
        mime="text/csv",
    )
  else:
    st.info("No data available to generate targets.")

with tab5:
  st.subheader(
      f"📝 Live Spreadsheet Data Editor — {selected_sheet} {selected_year}"
  )
  current_sheet_data = st.session_state["hierarchical_data"][selected_year][
      selected_sheet
  ]

  if st.session_state["is_authorized"]:
    edited_raw_df = st.data_editor(
        current_sheet_data,
        num_rows="dynamic",
        use_container_width=True,
        key=f"data_editor_{selected_year}_{selected_sheet}",
    )
    if st.button("💾 Save & Refresh Dashboard Metrics"):
      st.session_state["hierarchical_data"][selected_year][
          selected_sheet
      ] = edited_raw_df
      save_hierarchical_to_excel(
          st.session_state["hierarchical_data"],
          st.session_state["supervisor_mapping"],
      )
      st.success(
          f"Changes successfully saved and stored permanently for {selected_sheet}"
          f" ({selected_year})!"
      )
      st.rerun()
  else:
    st.warning(
        "🔒 Data editing is locked. Please authenticate in the sidebar to"
        " modify sheet data."
    )
    st.dataframe(current_sheet_data, use_container_width=True)

  @st.cache_data
  def convert_df_to_csv(df):
    return df.to_csv(index=False).encode("utf-8")

  csv_export = convert_df_to_csv(current_sheet_data)
  st.download_button(
      label=f"📥 Download {selected_sheet} {selected_year} Sheet as CSV",
      data=csv_export,
      file_name=f"updated_{selected_sheet}_{selected_year}_data.csv",
      mime="text/csv",
  )

with tab6:
  st.subheader(
      f"🧑‍💼 Supervisor Performance Roll-Up — {selected_sheet} {selected_year}"
  )
  st.markdown(
      "Each district's total pages are rolled up under its assigned supervisor"
      f" (stored in the `{SUPERVISOR_SHEET_NAME}` sheet inside the Excel"
      " workbook). Sundays are automatically excluded."
  )

  if not melted_df.empty:
    district_totals = (
        melted_df.groupby("DISTRICT")["PAGES"].sum().reset_index()
    )
    district_totals["SUPERVISOR"] = (
        district_totals["DISTRICT"]
        .str.strip()
        .str.upper()
        .map(st.session_state["supervisor_mapping"])
        .fillna("Unassigned")
    )

    unassigned_districts = sorted(
        district_totals.loc[
            district_totals["SUPERVISOR"] == "Unassigned", "DISTRICT"
        ].unique()
    )
    if unassigned_districts:
      st.warning(
          "These districts aren't mapped to a supervisor yet, so their pages"
          " are grouped under 'Unassigned': "
          + ", ".join(unassigned_districts)
          + ". Add them in the sidebar's 'Manage Supervisor Mapping' panel."
      )

    supervisor_summary = (
        district_totals.groupby("SUPERVISOR")["PAGES"]
        .sum()
        .reset_index()
        .rename(columns={"PAGES": "Total Pages"})
        .sort_values(by="Total Pages", ascending=False)
        .reset_index(drop=True)
    )

    mean_supervisor_output = supervisor_summary["Total Pages"].mean()

    def _supervisor_perf_color(value):
      if value >= mean_supervisor_output * 1.1:
        return "background-color: #2ecc71; color: white; font-weight: 700;"
      elif value >= mean_supervisor_output * 0.9:
        return "background-color: #f1c40f; color: #1f2937; font-weight: 700;"
      else:
        return "background-color: #e74c3c; color: white; font-weight: 700;"

    st.markdown("#### Supervisor Totals")
    st.dataframe(
        supervisor_summary.style.format({"Total Pages": "{:,.0f}"}).map(
            _supervisor_perf_color, subset=["Total Pages"]
        ),
        use_container_width=True,
    )

    supervisor_csv = supervisor_summary.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Supervisor Summary (CSV)",
        data=supervisor_csv,
        file_name=(
            f"supervisor_performance_{selected_sheet}_{selected_year}.csv"
        ),
        mime="text/csv",
    )

    st.markdown("---")

    fig_supervisor = px.bar(
        supervisor_summary,
        x="SUPERVISOR",
        y="Total Pages",
        title=(
            f"Total Pages Scanned per Supervisor — {selected_sheet}"
            f" {selected_year} (Sundays Excluded)"
        ),
        text_auto=".2s",
        template="plotly_white",
    )
    fig_supervisor.update_layout(xaxis={"categoryorder": "total descending"})
    fig_supervisor.add_hline(
        y=mean_supervisor_output,
        line_dash="dash",
        line_color="gray",
        annotation_text="Average",
    )
    st.plotly_chart(fig_supervisor, use_container_width=True)

    st.markdown("---")
    st.markdown("#### District Breakdown by Supervisor")
    for supervisor_name in supervisor_summary["SUPERVISOR"]:
      supervisor_districts = district_totals[
          district_totals["SUPERVISOR"] == supervisor_name
      ][["DISTRICT", "PAGES"]].sort_values(by="PAGES", ascending=False)
      supervisor_total = supervisor_districts["PAGES"].sum()
      with st.expander(
          f"{supervisor_name} — Total: {supervisor_total:,.0f} pages"
      ):
        st.dataframe(
            supervisor_districts.rename(
                columns={"DISTRICT": "District", "PAGES": "Total Pages"}
            ).style.format({"Total Pages": "{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )
  else:
    st.info("No data available to calculate supervisor performance.")

st.markdown("---")