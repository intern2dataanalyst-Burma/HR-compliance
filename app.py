import io
import zipfile
from io import BytesIO
from pathlib import Path
import os

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from openpyxl import load_workbook


custom_css = """
<style>
/* Global Background & Font */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: radial-gradient(circle at top, #ffffff 0, #f3f6fb 45%, #e5edf7 100%) !important;
    color: #1e293b;
}

.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 2rem !important;
}

[data-testid="stDataFrame"] {
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
}

.stDownloadButton > button {
    background: linear-gradient(to right, #1d9bf0, #34d399) !important;
    color: white !important;
    border: none !important;
    border-radius: 999px !important;
    font-weight: bold !important;
    padding: 0.5rem 1.5rem !important;
}
</style>
"""


BASE_DIR = Path(__file__).resolve().parent
DATA_CANDIDATES = [
    BASE_DIR / "data" / "real_master_compliance.xlsx",
    BASE_DIR / "sample_data" / "dummy_master_compliance.xlsx",
]
TEMPLATES_DIR = BASE_DIR / "templates"


@st.cache_data(show_spinner=False, ttl=300)
def load_master_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    try:
        excel_url = st.secrets["EXCEL_FILE_URL"]
        response = requests.get(excel_url, timeout=15)
        response.raise_for_status()
        file_bytes = io.BytesIO(response.content)
        dfs = pd.read_excel(file_bytes, sheet_name=[0, 1], skiprows=1, engine="openpyxl")
        df_conso = dfs[0].copy()
        df_units = dfs[1].copy()
        df_conso.columns = df_conso.columns.astype(str).str.strip()
        df_units.columns = df_units.columns.astype(str).str.strip()

        if 'Unit' in df_conso.columns and 'Unit' in df_units.columns:
            df_units_subset = df_units[['Unit', 'State']].drop_duplicates()
            df_conso = df_conso.merge(df_units_subset, on='Unit', how='left')

        return df_conso, df_units
    except Exception as exc:
        st.error(f"Failed to download master data from Google Drive: {exc}")
        st.stop()


@st.cache_data(show_spinner=False)
def build_merged_view() -> pd.DataFrame:
    df_conso, df_units = load_master_data()
    df = df_conso.copy()

    canonical_columns = {
        "Code": ["Empl_Code", "Code", "Employee Code", "Emp Code"],
        "Employee Name": ["Name_of_the_employee", "Employee Name", "Employee_Name", "Name"],
        "Father Name": ["Father_Name", "Father Name"],
        "Spouse Name": ["Spouce_Name", "Spouse Name", "Spouse_Name"],
        "UAN No": ["UAN No", "UAN", "UAN Number"],
        "PF No": ["PF No", "PF", "Provident Fund No"],
        "ESIC Old No": ["ESIC Old No", "ESIC No", "ESIC"],
        "PAN": ["PAN", "Pan"],
        "Joining_Date": ["Joining_Date", "Joining Date", "Join Date", "Joining_date"],
        "Exit_Date": ["Exit_Date", "Exit Date", "Exit_Date"],
        "Designation": ["Designation", "Designation Name"],
        "Department": ["Department", "Department Name"],
        "Days Paid": ["Days Paid", "Days_Paid"],
        "Days Present": ["Days Present", "Days_Present"],
        "Earned Basic": ["Earned Basic", "Earned_Basic", "Basic"],
        "Earned HRA": ["Earned HRA", "Earned_HRA", "HRA"],
        "Earned Gross Salary": ["Earned Gross Salary", "Earned_Gross_Salary", "Gross Salary", "Gross"],
        "Prov Fund": ["Prov Fund", "Prov_Fund", "PFund"],
        "ESIC": ["ESIC", "ESIC Amount"],
        "PTax": ["PTax", "Professional Tax"],
        "TDS": ["TDS", "Tax Deducted"],
        "Total Deductions": ["Total Deductions", "Total_Deductions", "Deduction", "Total Deduction"],
        "Net Paid": ["Net Paid", "Net_Paid", "Net Salary"],
    }

    for canonical_name, aliases in canonical_columns.items():
        matched_column = next((alias for alias in aliases if alias in df.columns), None)
        if matched_column is not None:
            df[canonical_name] = df[matched_column]
        else:
            df[canonical_name] = pd.Series(["" for _ in range(len(df))])

    state_candidate = next((col for col in ["State", "Region", "Location", "State/Region", "State_Region"] if col in df.columns), None)
    if state_candidate is not None:
        df["State"] = df[state_candidate]
    else:
        df["State"] = pd.Series(["" for _ in range(len(df))])

    if "Unit" in df.columns:
        df["Unit"] = df["Unit"]
    else:
        df["Unit"] = pd.Series(["" for _ in range(len(df))])

    month_candidate = next((col for col in ["Month Year", "Month", "Pay Month", "Month_Year", "Period", "For the Month", "Date"] if col in df.columns), None)
    if month_candidate is not None:
        df["Month Year"] = df[month_candidate]
    else:
        df["Month Year"] = pd.Series(["" for _ in range(len(df))])

    for column in ["Employee Name", "Department", "Designation", "Unit", "State", "Month Year"]:
        df[column] = df[column].fillna("")

    df["Net Paid"] = pd.to_numeric(df.get("Net Paid", pd.Series([0 for _ in range(len(df))])), errors="coerce").fillna(0)

    for unit_column in ["Unit", "Address", "Address1", "Address_1"]:
        if unit_column in df_units.columns:
            df["Unit_Address"] = df_units[unit_column]
            break

    if "Unit" in df_units.columns and "Unit" in df.columns:
        unit_lookup = df_units[[col for col in df_units.columns if col in ["Unit", "Address", "Address1", "Address_2", "Registration No", "RegistrationNo"]]]
        if not unit_lookup.empty:
            df = df.merge(unit_lookup, on="Unit", how="left")

    return df


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return "".join(ch.lower() for ch in str(value) if ch.isalnum())


def find_header_row(sheet) -> int:
    for row_idx in range(1, 26):
        row_values = [sheet.cell(row=row_idx, column=col_idx).value for col_idx in range(1, 26)]
        normalized_values = [normalize_header(value) for value in row_values if value is not None]
        if any(keyword in " ".join(normalized_values) for keyword in ["employee", "code", "name", "father", "spouse", "designation", "department", "days", "basic", "hra", "gross", "deduct", "net", "month", "period"]):
            return row_idx
    return 10


def get_value_for_header(header_value: object, row: dict) -> object:
    normalized = normalize_header(header_value)

    if any(keyword in normalized for keyword in ["code", "empl", "employeeid"]):
        return row.get("Code", "")
    if any(keyword in normalized for keyword in ["name", "employee"]) and "father" not in normalized and "spouse" not in normalized:
        return row.get("Employee Name", "")
    if "father" in normalized:
        return row.get("Father Name", "")
    if "spouse" in normalized:
        return row.get("Spouse Name", "")
    if "uan" in normalized:
        return row.get("UAN No", "")
    if "pf" in normalized and "no" in normalized:
        return row.get("PF No", "")
    if "esic" in normalized and "old" in normalized:
        return row.get("ESIC Old No", "")
    if "pan" in normalized:
        return row.get("PAN", "")
    if "join" in normalized or "appoint" in normalized:
        return row.get("Joining_Date", "")
    if "exit" in normalized:
        return row.get("Exit_Date", "")
    if "design" in normalized:
        return row.get("Designation", "")
    if "depart" in normalized:
        return row.get("Department", "")
    if "day" in normalized and "paid" in normalized:
        return row.get("Days Paid", "")
    if "day" in normalized and "present" in normalized:
        return row.get("Days Present", "")
    if "basic" in normalized:
        return row.get("Earned Basic", "")
    if "hra" in normalized:
        return row.get("Earned HRA", "")
    if "gross" in normalized:
        return row.get("Earned Gross Salary", "")
    if "prov" in normalized or "provident" in normalized:
        return row.get("Prov Fund", "")
    if "esic" in normalized:
        return row.get("ESIC", "")
    if "ptax" in normalized or "tax" in normalized and "professional" in normalized:
        return row.get("PTax", "")
    if "tds" in normalized:
        return row.get("TDS", "")
    if "deduct" in normalized or "totaldeduct" in normalized:
        return row.get("Total Deductions", "")
    if "net" in normalized or "salary" in normalized:
        return row.get("Net Paid", "")
    if "month" in normalized or "period" in normalized:
        return row.get("Month Year", "")
    return ""


def write_header_month(sheet, selected_month: str) -> None:
    target_keywords = ["month", "for the period ending", "wage month", "period"]
    for row_idx in range(1, 6):
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if cell.value is None:
                continue
            normalized_value = normalize_header(cell.value)
            if any(keyword in normalized_value for keyword in target_keywords):
                next_cell = sheet.cell(row=row_idx, column=col_idx + 1)
                if next_cell.value is None or normalize_header(str(next_cell.value)) != normalize_header(selected_month):
                    next_cell.value = selected_month
                else:
                    cell.value = selected_month
                return


def generate_dynamic_form(filtered_df: pd.DataFrame, template_source: io.BytesIO, selected_month: str) -> BytesIO:
    template_source.seek(0)
    workbook = load_workbook(template_source)
    sheet = workbook.active
    write_header_month(sheet, selected_month)
    header_row = find_header_row(sheet)

    header_cells = [sheet.cell(row=header_row, column=col_idx).value for col_idx in range(1, 51)]

    for row_idx, row in enumerate(filtered_df.to_dict(orient="records"), start=header_row + 1):
        for col_idx, header_value in enumerate(header_cells, start=1):
            if header_value is None:
                continue
            sheet.cell(row=row_idx, column=col_idx, value=get_value_for_header(header_value, row))

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


st.set_page_config(page_title="HR Compliance Engine", layout="wide")
st.markdown(custom_css, unsafe_allow_html=True)

if "forms_generated" not in st.session_state:
    st.session_state.forms_generated = False

if "generated_forms" not in st.session_state:
    st.session_state.generated_forms = {}


def reset_forms():
    st.session_state.forms_generated = False
    st.session_state.generated_forms = {}

header_html = """
<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-bottom: 1.5rem;">
    <img src="https://drive.google.com/thumbnail?id=1OpUw3MCFGLRs7GQ4xezk5ouqYTvy6yv9&sz=w1000"
         style="width: 120px; margin-bottom: 8px; filter: drop-shadow(0 4px 6px rgba(0,0,0,0.1));">
    <h1 style="color: #1d4ed8; font-family: 'Inter', sans-serif; font-weight: 800; font-size: 2.1rem; margin: 0; padding: 0; text-align: center;">
        HR Statutory Compliance Engine
    </h1>
</div>
"""
st.markdown(header_html, unsafe_allow_html=True)

merged_df = build_merged_view()
filtered_conso_df = merged_df.copy()

st.markdown("### Filters")
col1, col2, col3, col4 = st.columns(4)

with col1:
    if "State" in merged_df.columns and not merged_df["State"].dropna().eq("").all():
        state_list = ["All"] + sorted(merged_df['State'].dropna().astype(str).unique().tolist())
        selected_state = st.selectbox("State / Region", state_list, key="selected_state", on_change=reset_forms)
    else:
        selected_state = "All"
        st.info("State / Region column not found in the workbook; showing all records.")

with col2:
    if "Unit" in merged_df.columns:
        if selected_state == "All":
            unit_list = ["All"] + sorted(merged_df['Unit'].dropna().astype(str).unique().tolist())
        else:
            filtered_by_state = merged_df[merged_df['State'].astype(str) == str(selected_state)]
            unit_list = ["All"] + sorted(filtered_by_state['Unit'].dropna().astype(str).unique().tolist())
        selected_unit = st.selectbox("Unit", unit_list, key="selected_unit", on_change=reset_forms)
    else:
        selected_unit = "All"
        st.info("Unit column not found; using all records.")

with col3:
    if "Department" in merged_df.columns:
        department_list = ["All"] + sorted(merged_df['Department'].dropna().astype(str).unique().tolist())
        selected_department = st.selectbox("Department", department_list, key="selected_department", on_change=reset_forms)
    else:
        selected_department = "All"
        st.info("Department column not found.")

with col4:
    if 'Month Year' in merged_df.columns:
        month_list = ["All"] + sorted(merged_df['Month Year'].dropna().astype(str).unique().tolist())
    else:
        month_list = ["All"]
    selected_month = st.selectbox("Month-Year", month_list, key="selected_month", on_change=reset_forms)

filtered_df = merged_df.copy()

if selected_state != "All" and "State" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['State'].astype(str) == str(selected_state)]

if selected_unit != "All" and "Unit" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['Unit'].astype(str) == str(selected_unit)]

if selected_department != "All" and "Department" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['Department'].astype(str) == str(selected_department)]

if selected_month != "All" and 'Month Year' in filtered_df.columns:
    filtered_df = filtered_df[filtered_df['Month Year'].astype(str) == str(selected_month)]

total_emps = len(filtered_df)
total_paid = filtered_df['Net Paid'].sum() if 'Net Paid' in filtered_df.columns else 0

kpi_html = f"""
<div style="display: flex; gap: 20px; margin-bottom: 20px;">
    <div style="flex: 1; background: #ffffff; padding: 18px 24px; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <p style="margin: 0; color: #64748b; font-size: 0.85rem; font-weight: 600;">Total Employees</p>
        <h2 style="margin: 4px 0 0 0; color: #1d4ed8; font-size: 1.8rem; font-weight: 700;">{total_emps}</h2>
    </div>
    <div style="flex: 1; background: #ffffff; padding: 18px 24px; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <p style="margin: 0; color: #64748b; font-size: 0.85rem; font-weight: 600;">Total Net Paid</p>
        <h2 style="margin: 4px 0 0 0; color: #1d4ed8; font-size: 1.8rem; font-weight: 700;">₹ {total_paid:,.2f}</h2>
    </div>
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# Visual analytics section
st.markdown('<div class="custom-card" style="padding: 20px; margin-bottom: 24px;">', unsafe_allow_html=True)
_, radio_col = st.columns([3, 1])
with radio_col:
    chart_view = st.radio(
        "View By",
        ["Unit", "Department"],
        horizontal=True,
        key="chart_view_toggle",
    )

if selected_unit != "All" and "Department" in filtered_df.columns:
    chart_group_col = "Department"
    chart_title = "Department-Wise Employee Distribution"
else:
    chart_group_col = "Unit" if chart_view == "Unit" else "Department"
    chart_title = f"{chart_group_col}-Wise Employee Distribution"

if chart_group_col == "Unit" and "State" in filtered_df.columns:
    unit_counts = (
        filtered_df.groupby(["Unit", "State"])
        .size()
        .reset_index(name="Headcount")
        .sort_values("Headcount", ascending=False)
    )
else:
    unit_counts = (
        filtered_df.groupby(chart_group_col)
        .size()
        .reset_index(name="Headcount")
        .sort_values("Headcount", ascending=False)
    )

bar_kwargs = {
    "x": chart_group_col,
    "y": "Headcount",
    "title": chart_title,
    "text": "Headcount",
}
if chart_group_col == "Unit" and "State" in filtered_df.columns:
    bar_kwargs["color"] = "State"
else:
    bar_kwargs["color_discrete_sequence"] = ["#1d4ed8"]

region_chart = px.bar(unit_counts, **bar_kwargs)
region_chart.update_traces(textposition="outside")
region_chart.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=20, r=20, t=40, b=50),
)
region_chart.update_xaxes(tickangle=-30, showgrid=False, zeroline=False)
region_chart.update_yaxes(showgrid=False, zeroline=False)

status_column = None
for candidate in ["Status", "Compliance Status"]:
    if candidate in filtered_df.columns:
        status_column = candidate
        break

if status_column and not filtered_df[status_column].dropna().empty:
    status_summary = (
        filtered_df.groupby(status_column)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    status_colors = {
        "Compliant": "#10b981",
        "Pending": "#f59e0b",
        "Expired": "#ef4444",
    }
    status_chart = px.pie(
        status_summary,
        names=status_column,
        values="Count",
        hole=0.4,
        title="Compliance Status Breakdown",
        color_discrete_sequence=[status_colors.get(x, "#636efa") for x in status_summary[status_column]],
    )
    status_chart.update_traces(textinfo="percent+label")
else:
    if "Department" in filtered_df.columns and "Net Paid" in filtered_df.columns:
        status_summary = (
            filtered_df.groupby("Department", dropna=False)["Net Paid"]
            .sum()
            .reset_index()
            .sort_values("Net Paid", ascending=False)
            .head(8)
        )
        status_chart = px.bar(
            status_summary,
            x="Department",
            y="Net Paid",
            title="Total Net Paid by Department",
            color_discrete_sequence=["#1d4ed8"],
        )
        status_chart.update_xaxes(showgrid=False, zeroline=False)
        status_chart.update_yaxes(showgrid=False, zeroline=False)
    else:
        placeholder_df = pd.DataFrame({"Metric": ["No status data available"], "Value": [1]})
        status_chart = px.pie(
            placeholder_df,
            names="Metric",
            values="Value",
            hole=0.4,
            title="Status data unavailable",
        )
        status_chart.update_traces(textinfo="none")

status_chart.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=0, r=0, t=40, b=0),
)

chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.plotly_chart(region_chart, use_container_width=True, key="region_chart")
with chart_col2:
    st.plotly_chart(status_chart, use_container_width=True, key="status_chart")
st.markdown('</div>', unsafe_allow_html=True)

st.subheader("Filtered Employee Preview")
display_columns = ['Empl_Code', 'Name_of_the_employee', 'Department', 'Designation', 'Unit', 'Net Paid']
preview_df = filtered_df[[col for col in display_columns if col in filtered_df.columns]]
st.dataframe(preview_df, use_container_width=True, hide_index=True)

st.markdown("### Statutory Form Automation")
run_automation = st.button("Run Automation Engine")

if run_automation:
    if selected_state == "All":
        st.warning("Please choose a specific state or region to run the automation.")
    else:
        st.session_state.generated_forms = {}
        missing_forms = []
        form_names = ["Form A", "Form C", "Form D", "Form E", "Form IV", "Form V"]
        state_urls = st.secrets.get("templates", {}).get(selected_state, {})

        if not state_urls:
            st.error(f"No template URLs configured for {selected_state}")
            st.stop()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            with st.spinner("Compiling statutory forms from Google Drive templates..."):
                for form_name in form_names:
                    if form_name not in state_urls:
                        missing_forms.append(form_name)
                        continue

                    try:
                        response = requests.get(state_urls[form_name], timeout=15)
                        response.raise_for_status()
                        template_bytes = io.BytesIO(response.content)
                        form_bytes = generate_dynamic_form(filtered_df, template_bytes, str(selected_month))
                        zip_file.writestr(
                            f"{selected_state}_{selected_month}_{form_name}.xlsx",
                            form_bytes.getvalue(),
                        )
                        st.session_state.generated_forms[form_name] = form_bytes.getvalue()
                    except Exception as exc:
                        missing_forms.append(f"{form_name} ({exc})")

        zip_buffer.seek(0)
        st.session_state.forms_generated = len(st.session_state.generated_forms) > 0

        if missing_forms:
            st.warning(f"Templates skipped or failed for {selected_state}: {', '.join(missing_forms)}")
        if not st.session_state.forms_generated:
            st.error("No forms were generated for the selected state. Please verify the configured template URLs.")

if st.session_state.forms_generated:
    st.success("Forms generated successfully! Ready for download.")
    safe_state = "All" if selected_state == "All" else "".join(ch if ch.isalnum() else "_" for ch in str(selected_state))
    safe_month = "All" if selected_month == "All" else "".join(ch if ch.isalnum() else "_" for ch in str(selected_month))
    st.download_button(
        label="📦 Download All Generated Forms (ZIP)",
        data=zip_buffer.getvalue() if "zip_buffer" in locals() else b"",
        file_name=f"HR_Compliance_{safe_state}_{safe_month}.zip",
        mime="application/zip",
    )

