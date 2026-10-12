import calendar
import io
import os
import re
import shutil
import smtplib
import tempfile
import time
import zipfile
from copy import copy
from datetime import datetime
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter


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

# Constants for master file sync
MASTER_FILE_PATH = "data/Master_CONSO_DATA_ALL_UNITS.xlsx"
ONEDRIVE_MASTER_URL = "https://hungerpangs-my.sharepoint.com/:x:/p/ghiri_s/IQA8NPZyhwnyTo-xjBGG3LdmAUinyrrhOflimsUMPRcYQPs?download=1"

config_dir = BASE_DIR / ".streamlit"
config_dir.mkdir(parents=True, exist_ok=True)
config_file = config_dir / "config.toml"
if not config_file.exists() or "enableStaticServing" not in config_file.read_text(encoding="utf-8"):
    with open(config_file, "a", encoding="utf-8") as f:
        f.write("\n[server]\nenableStaticServing = true\n")

STATIC_ARCHIVE_DIR = BASE_DIR / "static" / "HR_Compliance_Archive"
STATIC_ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
STATIC_WEB_PREFIX = "app/static/HR_Compliance_Archive"

user_home = Path.home()
onedrive_path_secret = st.secrets.get("archive", {}).get("onedrive_path", None)
if onedrive_path_secret and Path(onedrive_path_secret).parent.exists():
    ONEDRIVE_ARCHIVE_DIR = Path(onedrive_path_secret)
else:
    onedrive_candidates = [
        user_home / "OneDrive - Hunger Pangs Pvt Ltd" / "HR_Compliance_Archive",
        user_home / "OneDrive" / "HR_Compliance_Archive",
    ]
    ONEDRIVE_ARCHIVE_DIR = next((path for path in onedrive_candidates if path.parent.exists()), None)

if ONEDRIVE_ARCHIVE_DIR:
    ONEDRIVE_ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)


def fetch_onedrive_file_with_retry(url: str, max_retries: int = 4, delay: int = 2) -> bytes:
    """
    Downloads a file from SharePoint/OneDrive with browser headers 
    and automatic exponential backoff retries against 503/429 errors.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
    }
    
    last_exception = None
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
            if response.status_code == 200:
                head = response.content[:200].lower()
                if b"<html" in head or b"<!doctype html" in head:
                    raise RuntimeError("OneDrive returned a login webpage instead of the direct Excel file.")
                return response.content
            
            response.raise_for_status()
        except Exception as err:
            last_exception = err
            if attempt < max_retries:
                time.sleep(delay * attempt)
            else:
                break

    raise last_exception


@st.cache_data(ttl=3600)
def get_active_master_data():
    if not os.path.exists(MASTER_FILE_PATH):
        os.makedirs(os.path.dirname(MASTER_FILE_PATH), exist_ok=True)
        content = fetch_onedrive_file_with_retry(ONEDRIVE_MASTER_URL)
        with open(MASTER_FILE_PATH, "wb") as f:
            f.write(content)
            
    df = pd.read_excel(MASTER_FILE_PATH, sheet_name="Conso_Data", skiprows=1, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()
    if "Month-Year" in df.columns:
        df["Month-Year"] = df["Month-Year"].astype(str).str.strip()
    return df


def email_file_to_outlook(file_bytes, filename):
    try:
        if hasattr(file_bytes, "getvalue"):
            file_bytes = file_bytes.getvalue()

        email_secrets = st.secrets.get("email_config", {})
        sender_email = email_secrets.get("sender_email")
        sender_password = email_secrets.get("sender_password")
        recipient_email = email_secrets.get("recipient_email", sender_email)

        if not sender_email or not sender_password:
            return False, "Email sender credentials missing in secrets.toml."

        msg = EmailMessage()
        msg["Subject"] = f"[HR Compliance Sync] {filename}"
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg.set_content(
            "Automated HR Compliance form generated from Streamlit portal. "
            "Attached is the file for Power Automate OneDrive archiving."
        )

        if filename.lower().endswith(".zip"):
            maintype, subtype = "application", "zip"
        else:
            maintype, subtype = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        msg.add_attachment(file_bytes, maintype=maintype, subtype=subtype, filename=filename)

        with smtplib.SMTP("smtp.gmail.com", 587, timeout=15) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)

        return True, "Emailed to Outlook for automatic OneDrive ingestion!"
    except Exception as e:
        return False, f"Email dispatch skipped/failed: {str(e)}"


def write_conso_data_sheet(master_path: str, new_df: pd.DataFrame) -> None:
    workbook = load_workbook(master_path)
    if "Conso_Data" in workbook.sheetnames:
        sheet_index = workbook.sheetnames.index("Conso_Data")
        del workbook["Conso_Data"]
        sheet = workbook.create_sheet("Conso_Data", sheet_index)
    else:
        sheet = workbook.create_sheet("Conso_Data")

    columns = list(new_df.columns)
    for col_idx, column_name in enumerate(columns, start=1):
        sheet.cell(row=1, column=col_idx, value=get_column_letter(col_idx))
        sheet.cell(row=2, column=col_idx, value=str(column_name))

    for row_offset, row in enumerate(new_df.itertuples(index=False, name=None), start=3):
        for col_idx, value in enumerate(row, start=1):
            if pd.isna(value):
                value = None
            elif hasattr(value, "isoformat"):
                pass 
            sheet.cell(row=row_offset, column=col_idx, value=value)

    workbook.save(master_path)


def normalize_unit_key(value: object) -> str:
    return re.sub(r"[\s\-_]+", "", str(value)).strip().lower()


@st.cache_data(ttl=3600, show_spinner=False)
def load_master_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        if Path(MASTER_FILE_PATH).exists():
            with open(MASTER_FILE_PATH, "rb") as f:
                file_bytes = io.BytesIO(f.read())
        else:
            excel_url = st.secrets.get("EXCEL_FILE_URL") or ONEDRIVE_MASTER_URL
            content = fetch_onedrive_file_with_retry(excel_url)
            file_bytes = io.BytesIO(content)
            os.makedirs(os.path.dirname(MASTER_FILE_PATH), exist_ok=True)
            with open(MASTER_FILE_PATH, "wb") as f:
                f.write(content)

        df_conso = pd.read_excel(file_bytes, sheet_name="Conso_Data", skiprows=1, engine="openpyxl")
        df_units = pd.read_excel(file_bytes, sheet_name="Units_Master", skiprows=1, engine="openpyxl")
        df_mapping_rules = pd.read_excel(file_bytes, sheet_name="Sheet1", header=None, engine="openpyxl")
        df_col_ref = pd.read_excel(file_bytes, sheet_name="Coloums", skiprows=3, engine="openpyxl")

        def _try_load_optional_sheet(candidate_names: list[str]) -> pd.DataFrame:
            for candidate in candidate_names:
                try:
                    file_bytes.seek(0)
                    sheet_df = pd.read_excel(file_bytes, sheet_name=candidate, engine="openpyxl")
                    sheet_df.columns = sheet_df.columns.astype(str).str.strip()
                    return sheet_df
                except (ValueError, KeyError):
                    continue
            return pd.DataFrame()

        df_emp_master = _try_load_optional_sheet(["Emp_Master", "Employee_Master", "EmpMaster", "Emp Mast"])
        df_leave_register = _try_load_optional_sheet(["Leave_Register", "Leave_Reg", "LeaveReg", "Leave Reg"])
        df_attendance_register = _try_load_optional_sheet(["Attendance_Register", "Attd_Reg", "AttdReg", "Attd Reg"])
        file_bytes.seek(0)

        df_conso.columns = df_conso.columns.astype(str).str.strip()
        df_units.columns = df_units.columns.astype(str).str.strip()

        if "Unit" in df_conso.columns:
            df_conso["Unit"] = df_conso["Unit"].astype(str).str.strip()
            df_conso["Unit_Display"] = df_conso["Unit"].astype(str).str.strip()
            df_conso["Unit_Clean"] = df_conso["Unit"].apply(normalize_unit_key)
        else:
            df_conso["Unit"] = pd.Series(["" for _ in range(len(df_conso))])
            df_conso["Unit_Display"] = pd.Series(["" for _ in range(len(df_conso))])
            df_conso["Unit_Clean"] = ""

        if "Unit" in df_units.columns:
            df_units["Unit"] = df_units["Unit"].astype(str).str.strip()
            df_units["Unit_Clean"] = df_units["Unit"].apply(normalize_unit_key)
        else:
            df_units["Unit"] = pd.Series(["" for _ in range(len(df_units))])
            df_units["Unit_Clean"] = ""

        if "Unit_Clean" in df_conso.columns and "Unit_Clean" in df_units.columns:
            units_state_col = next(
                (c for c in ["State", "STATE", "State/Region", "Region"] if c in df_units.columns),
                None,
            )
            if units_state_col is not None:
                state_lookup = df_units[["Unit_Clean", units_state_col]].drop_duplicates(subset=["Unit_Clean"])
                if units_state_col != "State":
                    state_lookup = state_lookup.rename(columns={units_state_col: "State"})
                df_conso = df_conso.merge(state_lookup, on="Unit_Clean", how="left", suffixes=("", "_master"))
                if "State_master" in df_conso.columns:
                    df_conso["State"] = df_conso["State"].fillna(df_conso["State_master"])
                    df_conso.drop(columns=["State_master"], inplace=True)
            else:
                st.warning(
                    "Units_Master has no 'State' column, so outlets could not be matched to a "
                    "State this run. State-based filtering and form generation may be incomplete "
                    "until the 'State' column is restored."
                )

        return df_conso, df_units, df_mapping_rules, df_col_ref, df_emp_master, df_leave_register, df_attendance_register
    except Exception as exc:
        st.error(f"Failed to download master data from OneDrive: {exc}")
        st.stop()


@st.cache_data(show_spinner=False)
def build_merged_view() -> pd.DataFrame:
    df_conso, df_units, df_mapping_rules, df_col_ref, df_emp_master, df_leave_register, df_attendance_register = load_master_data()
    df = df_conso.copy()

    conso_code_col = next((c for c in df.columns if "code" in c.lower()), None)

    if not df_emp_master.empty and "Unit" in df.columns:
        code_col = next((c for c in df_emp_master.columns if "code" in c.lower()), None)
        if code_col and conso_code_col:
            df = df.merge(
                df_emp_master.add_suffix("_empmast").rename(columns={f"{code_col}_empmast": code_col}),
                left_on=conso_code_col,
                right_on=code_col,
                how="left",
                suffixes=("", "_empmast_dup"),
            )

    if df_attendance_register is not None and not df_attendance_register.empty:
        attendance_code_col = next((c for c in df_attendance_register.columns if "code" in c.lower()), None)
        if attendance_code_col and conso_code_col:
            df = df.merge(
                df_attendance_register.add_suffix("_attendance").rename(
                    columns={f"{attendance_code_col}_attendance": attendance_code_col}
                ),
                left_on=conso_code_col,
                right_on=attendance_code_col,
                how="left",
                suffixes=("", "_attendance_dup"),
            )

    if df_leave_register is not None and not df_leave_register.empty:
        leave_code_col = next((c for c in df_leave_register.columns if "code" in c.lower()), None)
        if leave_code_col and conso_code_col:
            df = df.merge(
                df_leave_register.add_suffix("_leave").rename(
                    columns={f"{leave_code_col}_leave": leave_code_col}
                ),
                left_on=conso_code_col,
                right_on=leave_code_col,
                how="left",
                suffixes=("", "_leave_dup"),
            )

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
        df["State"] = df[state_candidate].fillna("").astype(str).str.strip()
    else:
        df["State"] = pd.Series(["" for _ in range(len(df))])

    if "Unit" in df.columns:
        df["Unit_Raw"] = df["Unit"].astype(str).str.strip()
        df["Unit_Display"] = df["Unit_Raw"]
        df["Unit_Clean"] = df["Unit_Raw"].apply(normalize_unit_key)
    else:
        df["Unit_Raw"] = pd.Series(["" for _ in range(len(df))])
        df["Unit_Display"] = pd.Series(["" for _ in range(len(df))])
        df["Unit_Clean"] = pd.Series(["" for _ in range(len(df))])

    month_candidate = next((col for col in ["Month-Year", "Month Year", "Month", "Pay Month", "Month_Year", "Period", "For the Month", "Date"] if col in df.columns), None)
    if month_candidate is not None:
        df["Month-Year"] = df[month_candidate]
    else:
        df["Month Year"] = pd.Series(["" for _ in range(len(df))])

    for column in ["Employee Name", "Department", "Designation", "Unit_Raw", "Unit_Display", "State", "Month-Year"]:
        if column in df.columns:
            df[column] = df[column].fillna("")

    df["Net Paid"] = pd.to_numeric(df.get("Net Paid", pd.Series([0 for _ in range(len(df))])), errors="coerce").fillna(0)

    address_column = next(
        (c for c in ["Address", "Address1", "Address_1"] if c in df_units.columns),
        None,
    )
    if address_column and "Unit_Clean" in df_units.columns and "Unit_Clean" in df.columns:
        address_lookup = df_units[["Unit_Clean", address_column]].drop_duplicates(subset=["Unit_Clean"])
        address_dict = dict(
            zip(address_lookup["Unit_Clean"].astype(str).str.strip(), address_lookup[address_column])
        )
        df["Unit_Address"] = df["Unit_Clean"].map(address_dict)

    if "Unit_Clean" in df.columns and "Unit_Clean" in df_units.columns:
        unit_name_col = df_units.columns[2] if len(df_units.columns) > 2 else "Unit"
        if unit_name_col in df_units.columns and unit_name_col != "Unit_Clean":
            try:
                unit_lookup = df_units[["Unit_Clean", unit_name_col]].drop_duplicates(subset=["Unit_Clean"])
                unit_dict = dict(zip(unit_lookup["Unit_Clean"].astype(str).str.strip(), 
                                    unit_lookup[unit_name_col].astype(str).str.strip()))
                df["Unit_Display"] = df["Unit_Clean"].map(unit_dict).fillna(df["Unit_Raw"].astype(str).str.strip())
            except Exception as e:
                st.warning(f"Unit name lookup issue: {e}")

    return df


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return "".join(ch.lower() for ch in str(value) if ch.isalnum())


HEADER_KEYWORDS = ["employee", "code", "name", "father", "spouse", "designation", "department",
                    "days", "basic", "hra", "gross", "deduct", "net", "month", "period"]


def find_header_row(sheet) -> int:
    best_row = 10
    best_score = 0
    for row_idx in range(1, 26):
        row_values = [sheet.cell(row=row_idx, column=col_idx).value for col_idx in range(1, 26)]
        normalized_values = [normalize_header(value) for value in row_values if value is not None]
        distinct_matching_cells = sum(
            1 for value in normalized_values if any(keyword in value for keyword in HEADER_KEYWORDS)
        )
        if distinct_matching_cells >= 2 and distinct_matching_cells > best_score:
            best_score = distinct_matching_cells
            best_row = row_idx
    return best_row


def get_unit_master_details(df_units: pd.DataFrame | None, selected_unit: str | None) -> dict[str, str]:
    if df_units is None or selected_unit is None:
        return {"unit_name": "", "unit_address": "", "ward_circle": ""}

    df_units = df_units.copy()
    df_units.columns = df_units.columns.astype(str).str.strip()

    if "Unit_Clean" not in df_units.columns:
        if "Unit" in df_units.columns:
            df_units["Unit_Clean"] = df_units["Unit"].apply(normalize_unit_key)
        else:
            return {"unit_name": "", "unit_address": "", "ward_circle": ""}

    clean_target = normalize_unit_key(selected_unit)
    match = df_units[df_units["Unit_Clean"].apply(normalize_unit_key) == clean_target]
    if match.empty:
        return {"unit_name": str(selected_unit).strip(), "unit_address": "", "ward_circle": ""}

    row = match.iloc[0]

    addr_col = next((col for col in df_units.columns if "address" in str(col).lower()), None)
    unit_address = ""
    if addr_col and pd.notna(row[addr_col]):
        unit_address = str(row[addr_col]).strip()
    elif len(row) > 4 and pd.notna(row.iloc[4]):
        unit_address = str(row.iloc[4]).strip()

    ward_col = next((col for col in df_units.columns if any(w in str(col).lower() for w in ["ward", "circle"])), None)
    ward_circle = ""
    if ward_col and pd.notna(row[ward_col]):
        ward_circle = str(row[ward_col]).strip()
    elif len(row) > 12 and pd.notna(row.iloc[12]):
        ward_circle = str(row.iloc[12]).strip()

    return {
        "unit_name": str(row["Unit"]).strip() if "Unit" in row.index else str(selected_unit).strip(),
        "unit_address": unit_address,
        "ward_circle": ward_circle,
    }


def get_unit_master_column_value(df_units: pd.DataFrame | None, selected_unit: str | None, column_letter: str) -> object:
    if df_units is None or df_units.empty or not selected_unit or not column_letter:
        return ""
    df_units = df_units.copy()
    df_units.columns = df_units.columns.astype(str).str.strip()
    unit_col = next((c for c in df_units.columns if str(c).strip().lower().rstrip() == "unit"), None)
    if unit_col is None:
        return ""
    match = df_units[df_units[unit_col].apply(normalize_unit_key) == normalize_unit_key(selected_unit)]
    if match.empty:
        return ""
    col_idx0 = column_letter_to_index(column_letter) - 1
    if col_idx0 >= len(match.columns):
        return ""
    value = match.iloc[0, col_idx0]
    return "" if pd.isna(value) else value


def _lookup_col_ref_value(letter: str, df_col_ref: pd.DataFrame | None, row_data: dict) -> object:
    """Look up a Coloums registry entry by its short code (e.g. 'B', 'AT1', 'EN') and
    return the matching field's value from the current employee row. Matching is
    tolerant of stray whitespace, hyphens, and underscores in the registry's own code
    column ('AT-1' vs 'AT_1' vs 'AT1')."""
    if not letter:
        return None
    col_name = get_col_registry(df_col_ref).get(_clean_tag(letter))
    if col_name is not None and col_name in row_data:
        val = row_data.get(col_name, "")
        return val if not pd.isna(val) else ""
    return None


COL_PREFIX_RE = re.compile(r"^col[\s\-_]+", re.IGNORECASE)

# Physical first-data-row for each (State, Form Name) pair, verified directly against
# the real uploaded template for each. This MUST be keyed by state as well as form name:
# multiple states reuse the same form name/number (e.g. "FORM A" exists for both
# Chandigarh and Gujarat, "Form X"/"Form XI"/"Form XII" exist for both Telangana and
# West Bengal) with genuinely different physical row layouts. A flat form-name-only
# dict silently collides across states -- this was confirmed to have caused Gujarat's
# FORM A, Form M, and Form IV A to use Chandigarh/placeholder row numbers instead of
# their own verified ones.
FORM_START_ROWS_BY_STATE = {
    ("CHANDIGARH", "FORM A"): 19,
    ("CHANDIGARH", "FORM C"): 9,
    ("CHANDIGARH", "FORM D"): 8,
    ("CHANDIGARH", "FORM E"): 9,
    ("CHANDIGARH", "FORM IV"): 10,
    ("CHANDIGARH", "FORM V"): 11,
    ("GUJARAT", "FORM A"): 10,
    ("GUJARAT", "FORM P"): 13,
    ("GUJARAT", "FORM M"): 14,
    ("GUJARAT", "FORM IV A"): 10,
    ("GUJARAT", "FORM C"): 12,
    ("TELANGANA", "FORM X"): 7,
    ("TELANGANA", "FORM XI"): 8,
    ("TELANGANA", "FORM XII"): 6,
    ("TELANGANA", "FORM XXII"): 10,
    ("TELANGANA", "FORM XXIII"): 9,
    ("TELANGANA", "FORM XXV"): 10,
    ("TELANGANA", "FORM III"): 11,
    ("WEST BENGAL", "FORM J"): 9,
    ("WEST BENGAL", "FORM M"): 10,
    ("WEST BENGAL", "FORM U"): 10,
    ("WEST BENGAL", "FORM X"): 8,
    ("WEST BENGAL", "FORM XI"): 8,
    ("WEST BENGAL", "FORM XII"): 7,
}


def get_form_start_row(selected_state: str | None, form_name: str | None, default: int = 9) -> int:
    state_key = str(selected_state).strip().upper() if selected_state else ""
    form_key = str(form_name).strip().upper() if form_name else ""
    return FORM_START_ROWS_BY_STATE.get((state_key, form_key), default)


def _looks_like_numbering_row(sheet, row_idx: int, max_col_scan: int = 5) -> bool:
    """Detect a legend/numbering row (e.g. '1','2','3'... or 'Sr. No.') sitting at
    start_row, so employee data never gets pasted directly on top of it. A genuine
    legend row is purely small sequential numeric markers with no real text -- if
    any of the first few columns holds actual text (an employee name, a real
    multi-digit code, etc.), this bails out immediately and treats it as real data,
    not a legend, to avoid ever skipping a genuine first employee row."""
    numeric_hits = 0
    label_hits = 0
    for col_idx in range(1, max_col_scan + 1):
        val = sheet.cell(row=row_idx, column=col_idx).value
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue
        normalized = normalize_header(text)
        if normalized in ("srno", "slno", "sino", "serialno"):
            label_hits += 1
            continue
        if re.fullmatch(r"\d{1,2}(\.0)?", text):
            numeric_hits += 1
            continue
        return False
    return (numeric_hits + label_hits) >= 2


INTERNAL_ANNOTATION_PATTERNS = ["salreg", "empmast", "emplmast", "empmaster", "salaryreg"]


def _clear_internal_annotation_row(sheet, row_idx: int, max_col_scan: int = 60) -> None:
    """Clear cells containing known internal build-note annotations (like 'Sal Reg' or
    'Emp Mast' -- leftover data-source notes from template construction) without
    touching genuine statutory sub-headers or numbering legends that might share the
    same row. Deliberately does NOT blank the whole row -- only cells whose text
    matches one of these specific known junk patterns, unlike the old blanket
    reference-row-clearing logic that was removed earlier in this project for
    destroying real sub-headers."""
    if row_idx < 1:
        return
    for col_idx in range(1, max_col_scan + 1):
        cell = sheet.cell(row=row_idx, column=col_idx)
        if cell.value is None:
            continue
        normalized = normalize_header(cell.value)
        if any(pattern in normalized for pattern in INTERNAL_ANNOTATION_PATTERNS):
            safe_write(sheet, cell.coordinate, None)


def _is_col_prefixed(rule_text: str) -> bool:
    return bool(COL_PREFIX_RE.match(rule_text.strip()))

def _extract_col_letter(token: str) -> str:
    token = token.strip()
    token = COL_PREFIX_RE.sub("", token)
    match = re.match(r"[A-Za-z0-9]+", token)
    return match.group(0).upper() if match else ""


EMBEDDED_COL_REF_RE = re.compile(r"\{\s*col[\s\-_]*([a-zA-Z0-9]+)\s*\}", re.IGNORECASE)


def _substitute_embedded_col_refs(text: str, df_col_ref: pd.DataFrame | None, row_data: dict) -> str:
    def _replace(match: re.Match) -> str:
        letter = match.group(1).upper()
        value = _lookup_col_ref_value(letter, df_col_ref, row_data)
        return "" if value in (None, "") else str(value)
    return EMBEDDED_COL_REF_RE.sub(_replace, text)


# ---------------------------------------------------------------------------
# Tag registry + arithmetic formula engine
# ---------------------------------------------------------------------------
_COL_REF_CACHE: dict = {}


def _clean_tag(value: object) -> str:
    return re.sub(r"[\s\-_]+", "", str(value)).strip().upper()


def get_col_registry(df_col_ref: pd.DataFrame | None) -> dict[str, str]:
    """{CLEAN_TAG: conso column name} built once per Coloums DataFrame (first match wins)."""
    if df_col_ref is None or df_col_ref.empty or df_col_ref.shape[1] < 2:
        return {}
    cached = _COL_REF_CACHE.get(id(df_col_ref))
    if cached is not None and cached[0] is df_col_ref:
        return cached[1]
    registry: dict[str, str] = {}
    for code, name in zip(df_col_ref.iloc[:, 0], df_col_ref.iloc[:, 1]):
        if pd.isna(code) or pd.isna(name):
            continue
        key = _clean_tag(code)
        if key and key not in registry:
            registry[key] = str(name).strip()
    if len(_COL_REF_CACHE) > 8:
        _COL_REF_CACHE.clear()
    _COL_REF_CACHE[id(df_col_ref)] = (df_col_ref, registry)
    return registry


# "Col_BU", "Col _BC", "COL-ES", "Col - CQ", "Col BU" ...
_FORMULA_COL_TAG_RE = re.compile(r"col[\s_\-]*([A-Za-z]{1,3}\d{0,2})(?![A-Za-z0-9])", re.IGNORECASE)
_FORMULA_BARE_TAG_RE = re.compile(r"([A-Za-z]{1,3}\d{0,2})(?![A-Za-z0-9])")
_FORMULA_NUM_RE = re.compile(r"\d+(?:\.\d+)?")
_FORMULA_LONE_COL_RE = re.compile(r"col(?![A-Za-z0-9])", re.IGNORECASE)


def tokenize_col_formula(rule_text: str, registry: dict[str, str] | None = None) -> tuple[list, list[str]]:
    """Turn a rule like 'Col_BU -( Col _BC + Col _BD)' into tokens.
    Returns (tokens, notes). Tokens: ('tag','BU') ('num',1.0) ('op','+') ('(',) (')',).
    Parsing stops at the first piece of descriptive text after a complete operand
    (e.g. 'Col_B Code' or 'Col_N Date of entry into service:')."""
    registry = registry or {}
    text = str(rule_text).replace("\xa0", " ").replace("{", "(").replace("}", ")").replace("[", "(").replace("]", ")")
    text = text.replace("–", "-").replace("—", "-").replace("×", "*")
    tokens: list = []
    notes: list[str] = []
    pos = 0
    expect_operand = True
    while pos < len(text):
        ch = text[pos]
        if ch.isspace():
            pos += 1
            continue
        if expect_operand:
            if ch == "(":
                tokens.append(("(",)); pos += 1; continue
            if ch in "+-":  # unary sign
                tokens.append(("num", 0.0)); tokens.append(("op", ch)); pos += 1; continue
            m = _FORMULA_COL_TAG_RE.match(text, pos)
            if m:
                tokens.append(("tag", _clean_tag(m.group(1)))); pos = m.end(); expect_operand = False; continue
            m = _FORMULA_LONE_COL_RE.match(text, pos)
            if m:  # 'COL (BU - ...)' -- the word Col wrapping a group
                pos = m.end(); continue
            m = _FORMULA_NUM_RE.match(text, pos)
            if m:
                tokens.append(("num", float(m.group(0)))); pos = m.end(); expect_operand = False; continue
            m = _FORMULA_BARE_TAG_RE.match(text, pos)
            if m and _clean_tag(m.group(1)) in registry and tokens:
                tokens.append(("tag", _clean_tag(m.group(1)))); pos = m.end(); expect_operand = False; continue
            break
        else:
            if ch in "+-*/":
                tokens.append(("op", ch)); pos += 1; expect_operand = True; continue
            if ch == ")":
                tokens.append((")",)); pos += 1; continue
            if ch == "(":  # 'Col_BU ( Col BO+Col BI )' -- missing operator, read as subtraction
                tokens.append(("op", "-")); tokens.append(("(",)); pos += 1; expect_operand = True
                notes.append("missing operator before '(' read as '-'")
                continue
            break  # trailing descriptive text
    # drop a dangling operator / balance parentheses
    while tokens and tokens[-1][0] == "op":
        tokens.pop()
    depth = 0
    cleaned = []
    for tok in tokens:
        if tok[0] == "(":
            depth += 1
        elif tok[0] == ")":
            if depth == 0:
                continue
            depth -= 1
        cleaned.append(tok)
    cleaned.extend([(")",)] * depth)
    return cleaned, notes


def _to_number(value: object):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if pd.isna(value) else float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def evaluate_col_formula(tokens: list, get_value) -> tuple[object, list[str]]:
    """Evaluate tokens with + - * / and parentheses (no eval()). get_value(tag) -> raw value.
    Missing / non-numeric operands count as 0. Returns (result, problems); result is ''
    when every operand was blank."""
    problems: list[str] = []
    seen_any = [False]
    pos = [0]

    def peek():
        return tokens[pos[0]] if pos[0] < len(tokens) else None

    def parse_expr():
        val = parse_term()
        while peek() and peek()[0] == "op" and peek()[1] in "+-":
            op = tokens[pos[0]][1]; pos[0] += 1
            rhs = parse_term()
            val = val + rhs if op == "+" else val - rhs
        return val

    def parse_term():
        val = parse_atom()
        while peek() and peek()[0] == "op" and peek()[1] in "*/":
            op = tokens[pos[0]][1]; pos[0] += 1
            rhs = parse_atom()
            if op == "*":
                val *= rhs
            elif rhs == 0:
                problems.append("division by zero -> 0"); val = 0.0
            else:
                val /= rhs
        return val

    def parse_atom():
        tok = peek()
        if tok is None:
            return 0.0
        pos[0] += 1
        if tok[0] == "num":
            return tok[1]
        if tok[0] == "tag":
            raw = get_value(tok[1])
            number = _to_number(raw)
            if raw is None:
                problems.append(f"tag {tok[1]} not found")
            if number is None:
                return 0.0
            seen_any[0] = True
            return number
        if tok[0] == "(":
            val = parse_expr()
            if peek() and peek()[0] == ")":
                pos[0] += 1
            return val
        return 0.0

    result = parse_expr()
    if not seen_any[0]:
        return "", problems
    result = round(result, 2)
    if result == int(result):
        result = int(result)
    return result, problems


def resolve_formula_rule(rule_text: str, df_col_ref: pd.DataFrame | None, row_data: dict) -> object:
    """Resolve a Col-prefixed rule: a single tag (raw value returned) or an arithmetic expression."""
    registry = get_col_registry(df_col_ref)
    tokens, _ = tokenize_col_formula(rule_text, registry)
    tags = [t for t in tokens if t[0] == "tag"]
    if not tags:
        return ""
    if len(tokens) == 1:  # plain single tag
        value = _lookup_col_ref_value(tags[0][1], df_col_ref, row_data)
        return value if value is not None else ""
    return evaluate_col_formula(tokens, lambda tag: _lookup_col_ref_value(tag, df_col_ref, row_data))[0]



def _refresh_month_in_text(text: str, active_month: str | None, selected_year: int | None) -> str:
    """Fixed-text rules such as 'During wage month June 2026' carry a stale month.
    Re-point 'wage month / period / month of <Month YYYY>' to the selected month."""
    if not text or active_month in {"All", "", None}:
        return text
    year_value = selected_year if selected_year is not None else datetime.now().year
    try:
        month_name = datetime(int(year_value), datetime.strptime(str(active_month)[:3], "%b").month, 1).strftime("%B")
    except ValueError:
        return text
    pattern = re.compile(r"((?:wage\s+)?(?:month|period)(?:\s+of|\s+ending)?\s*[:\-]*\s*)" + MONTH_NAME_ALTERNATION + r"[\s,\-]*(\d{4}|\d{2})\b", re.IGNORECASE)
    return pattern.sub(lambda m: f"{m.group(1)}{month_name} {year_value}", text)


def resolve_col_rule(
    rule_str: object,
    df_col_ref: pd.DataFrame | None,
    row_data: dict,
    idx: int = 0,
    active_month: str | None = None,
    selected_year: int | None = None,
    df_units: pd.DataFrame | None = None,
    selected_unit: str | None = None,
) -> object:
    if rule_str is None:
        return ""

    rule_text = str(rule_str).strip()
    if not rule_text:
        return ""

    if EMBEDDED_COL_REF_RE.search(rule_text):
        rule_text = _substitute_embedded_col_refs(rule_text, df_col_ref, row_data)

    lowered_rule = rule_text.lower()

    if "fixed text" in lowered_rule:
        cleaned = re.sub(r"[\(\-]?\s*fixed\s*text\s*\)?", "", rule_text, flags=re.IGNORECASE)
        return _refresh_month_in_text(cleaned.strip(" -"), active_month, selected_year)

    if lowered_rule == "auto_generated" or lowered_rule.startswith("auto"):
        return idx + 1

    if lowered_rule == "wage month":
        month_value = active_month if active_month not in {"All", "", None} else "July"
        year_value = selected_year if selected_year is not None else datetime.now().year
        return f"{month_value[:3]}-{str(year_value)[-2:]}"

    active_month_name = active_month if active_month not in {"All", "", None} else "July"
    year_value = selected_year if selected_year is not None else datetime.now().year
    try:
        month_num = datetime.strptime(active_month_name[:3], "%b").month
    except ValueError:
        month_num = datetime.now().month
    days_in_month = calendar.monthrange(int(year_value), month_num)[1]

    if "1st day" in lowered_rule and "month" in lowered_rule:
        return datetime(int(year_value), month_num, 1).strftime("%d-%b-%Y")
    if "last day" in lowered_rule and "month" in lowered_rule:
        return datetime(int(year_value), month_num, days_in_month).strftime("%d-%b-%Y")
    if "days in cal" in lowered_rule:
        return days_in_month

    if "unit master" in lowered_rule and "col" in lowered_rule:
        parts = re.split(r"\s*\+\s*", rule_text)
        resolved_parts = []
        for part in parts:
            match = re.search(r"col_([a-z]+)", part.strip(), flags=re.IGNORECASE)
            if match:
                value = get_unit_master_column_value(df_units, selected_unit, match.group(1))
                if value not in (None, ""):
                    resolved_parts.append(str(value).strip())
        return ", ".join(resolved_parts)

    if _is_col_prefixed(rule_text):
        return resolve_formula_rule(rule_text, df_col_ref, row_data)

    if rule_text.isalpha() and rule_text.isupper() and 1 <= len(rule_text) <= 3:
        value = _lookup_col_ref_value(rule_text, df_col_ref, row_data)
        if value is not None:
            return value

    return rule_text


def get_value_for_header(
    header_value: object,
    row: dict,
    unit_master_details: dict[str, str] | None = None,
    df_mapping_rules: pd.DataFrame | None = None,
    df_col_ref: pd.DataFrame | None = None,
    idx: int = 0,
    df_units: pd.DataFrame | None = None,
    selected_unit: str | None = None,
) -> object:
    normalized = normalize_header(header_value)

    if df_mapping_rules is not None and not df_mapping_rules.empty:
        for _, rule_row in df_mapping_rules.iterrows():
            values = [str(value).strip() for value in rule_row.dropna().tolist() if str(value).strip()]
            if len(values) < 2:
                continue
            if normalize_header(values[0]) == normalized:
                return resolve_col_rule(
                    values[1], df_col_ref, row, idx, df_units=df_units, selected_unit=selected_unit
                )

    # ---------------------------------------------------------------------
    # ULTIMATE FAILSAFE: If Excel Sheet1 rules are empty or fail to read, 
    # Python forcefully catches the Form E and Form IV headers directly here.
    # ---------------------------------------------------------------------
    if "wagespayable" in normalized:
        for key in ["Fixed Monthly Gross", "Earned Gross Salary", "Gross Salary", "Net Paid"]:
            if key in row and pd.notna(row[key]) and str(row[key]).strip() != "":
                return row[key]
                
    if "normalearnings" in normalized or "totalearning" in normalized:
        for key in ["Earned Gross Salary", "Fixed Monthly Gross", "Gross Salary", "Net Paid"]:
            if key in row and pd.notna(row[key]) and str(row[key]).strip() != "":
                return row[key]
                
    if "amountdeducted" in normalized:
        for key in ["Total Deductions", "Deduction", "Total Deduction"]:
            if key in row and pd.notna(row[key]) and str(row[key]).strip() != "":
                return row[key]

    if "lwf" in normalized:
        for key in ["Employer LWF", "Employee LWF", "LWF Employer", "LWF Employee", "LWF"]:
            if key in row and pd.notna(row[key]) and str(row[key]).strip() != "":
                return row[key]

    if "overtime" in normalized:
        for key in ["Overtime", "Over Time", "O_TIME_HRS", "OT_WAGERATE_PERHR"]:
            if key in row and pd.notna(row[key]) and str(row[key]).strip() != "":
                return row[key]
    # ---------------------------------------------------------------------

    if "code" in normalized:
        return row.get("Code", "")
    if any(keyword in normalized for keyword in ["srno", "slno", "sino", "serialno"]):
        return idx + 1
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
        return row.get("Month-Year", "")
    return ""


def column_letter_to_index(column_letter: str) -> int:
    column_letter = column_letter.upper()
    total = 0
    for char in column_letter:
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total


def _coerce_numeric_for_write(value):
    """Turn numeric-looking values into real int/float so Excel treats them as
    numbers instead of text, and whole numbers display cleanly (no trailing
    '.0'). Leaves genuinely non-numeric text and leading-zero codes untouched,
    since those must stay text (e.g. an employee code '0102')."""
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else value
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return value
        unsigned = stripped[1:] if stripped.startswith("-") else stripped
        if re.fullmatch(r"\d+", unsigned) and not unsigned.startswith("0"):
            try:
                return int(stripped)
            except ValueError:
                return value
        if re.fullmatch(r"\d+\.\d+", unsigned):
            try:
                as_float = float(stripped)
                return int(as_float) if as_float.is_integer() else as_float
            except ValueError:
                return value
        return value
    # numpy int64/float64 and similar numeric-but-not-Python-native types
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return value
    return int(as_float) if as_float.is_integer() else as_float


def safe_write(sheet, cell_coordinate: str, value) -> None:
    cell = sheet[cell_coordinate]
    value = _coerce_numeric_for_write(value)

    # Prevent scientific notation (e.g., 1.7991E+15) for large IDs like UAN/Accounts/Signatures,
    # and give whole numbers a clean '0' display format (no decimals, no thousands separator,
    # so a narrow column shows the real digits instead of "##").
    is_whole_number = isinstance(value, int) and not isinstance(value, bool)
    is_long_number = value is not None and str(value).strip().lstrip('-').isdigit() and len(str(value).strip().lstrip('-')) >= 10
    if is_whole_number or is_long_number:
        cell.number_format = '0'
    elif isinstance(value, float):
        cell.number_format = '0.##'

    if isinstance(cell, MergedCell):
        for merged_range in sheet.merged_cells.ranges:
            if cell_coordinate in merged_range:
                top_left_coord = merged_range.coord.split(":")[0]
                sheet[top_left_coord].value = value
                if is_whole_number or is_long_number:
                    sheet[top_left_coord].number_format = '0'
                elif isinstance(value, float):
                    sheet[top_left_coord].number_format = '0.##'
                return
    cell.value = value


TEMPLATE_MARKER_FILLS = {"FFBDD6EE", "FFFFFF00", "FF9CC2E5"}

def is_reference_cell(cell) -> bool:
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.type == "rgb":
        return fill.fgColor.rgb == "FFFFFF00"
    return False

def detect_table_end_row(sheet, start_row: int, max_scan: int = 500) -> int:
    for row_idx in range(start_row, start_row + max_scan):
        if row_idx > sheet.max_row:
            break
        for col_idx in range(1, 11):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is not None and "TOTAL" in str(value).strip().upper():
                return row_idx

    last_marked_row = start_row
    blank_streak = 0
    for row_idx in range(start_row, start_row + max_scan):
        if row_idx > sheet.max_row:
            break
        row_has_marker = False
        for col_idx in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            fill = cell.fill
            if fill and fill.fgColor and fill.fgColor.type == "rgb" and fill.fgColor.rgb in TEMPLATE_MARKER_FILLS:
                row_has_marker = True
                break
        if row_has_marker:
            last_marked_row = row_idx
            blank_streak = 0
        else:
            blank_streak += 1
            if blank_streak >= 5:
                break
    return last_marked_row


def copy_cell_style(source_cell, target_cell) -> None:
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.alignment = copy(source_cell.alignment)


MONTH_NAME_ALTERNATION = r"(January|February|March|April|May|June|July|August|September|October|November|December)"
DATE_IN_TEXT_RE = re.compile(MONTH_NAME_ALTERNATION + r"[\s\-]*\d{2,4}", re.IGNORECASE)

MONTH_NAMES_LOWER = {
    "january", "february", "march", "april", "may", "june", "july",
    "august", "september", "october", "november", "december",
}


# ---------------------------------------------------------------------------
# Header injection (establishment name / address / month)
# ---------------------------------------------------------------------------
COMPANY_LEGAL_NAME = "Hunger Pangs Private Limited"
# 'company'      -> "Hunger Pangs Private Limited"
# 'unit'         -> "Ahmedabad Palladium"
# 'company_unit' -> "Hunger Pangs Private Limited - Ahmedabad Palladium"
ESTABLISHMENT_NAME_STYLE = "company"

_LABEL_SPLIT_RE = re.compile(r"\s*(?:[:：]\s*[-–—]*|\s[-–—]+\s*|[-–—]+\s)\s*")


def _split_label_value(text: str) -> tuple[str, str | None, str]:
    """'Name of establishment :- Old' -> ('Name of establishment', 'Old', ' :- ').
    No separator -> (text, None, '')."""
    m = _LABEL_SPLIT_RE.search(text)
    if not m or m.start() == 0:
        return text.strip(), None, ""
    return text[: m.start()].strip(), text[m.end():].strip(), text[m.start(): m.end()]


def _classify_header_label(label: str) -> str | None:
    """Which injectable field is this template label asking for?
    Returns 'est_name', 'est_name_address', 'address', 'employer_address' or None.
    Table column headers are never passed in (callers only scan rows above the table)."""
    n = normalize_header(label)
    if not n or len(n) > 80:
        return None
    has_addr = "address" in n
    is_est = any(k in n for k in ("establishment", "nameoftheshop", "nameofshop", "shopname", "nameoffactory", "factoryname"))
    is_emp = any(k in n for k in ("employer", "manager", "occupier", "owner"))
    if is_emp:
        return "employer_address" if has_addr else None
    if is_est and has_addr and n.startswith(("nameof", "nameand")) :
        return "est_name_address"
    if is_est and not has_addr and (n.startswith(("nameof", "name")) or n in ("establishmentname", "theestablishmentname")
                                    or "establishmentname" in n or "nameofthe" in n):
        return "est_name"
    if has_addr and (n.startswith("address") or n.startswith("addressofthe")):
        return "address"
    return None


def _merged_end_col(sheet, row: int, col: int) -> int:
    for rng in sheet.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.max_col
    return col


def _find_value_cell(sheet, row: int, col: int, max_reach: int = 12):
    """The cell that holds (or should hold) the value for a label-only cell: the first
    non-empty plain value to the right of the label's merged range, otherwise the first
    empty writable cell. Stops at the next label (text ending in ':')."""
    start = _merged_end_col(sheet, row, col) + 1
    first_empty = None
    for c in range(start, min(sheet.max_column, start + max_reach) + 1):
        cell = sheet.cell(row=row, column=c)
        val = cell.value
        if val is None or (isinstance(val, str) and not val.strip()):
            if first_empty is None and not isinstance(cell, MergedCell):
                first_empty = cell
            continue
        if isinstance(val, str):
            lbl, existing, _sep = _split_label_value(val)
            if val.strip().endswith(":") or (existing is not None and _classify_header_label(lbl)):
                break
        return cell
    return first_empty


def get_header_values(unit_master_details: dict[str, str], df_units=None, selected_unit=None) -> dict[str, str]:
    unit = (unit_master_details.get("unit_name") or "").strip()
    address = (unit_master_details.get("unit_address") or "").strip()
    if ESTABLISHMENT_NAME_STYLE == "unit":
        est = unit
    elif ESTABLISHMENT_NAME_STYLE == "company_unit" and unit:
        est = f"{COMPANY_LEGAL_NAME} - {unit}"
    else:
        est = COMPANY_LEGAL_NAME
    # optional Units_Master override for the legal establishment name
    if df_units is not None and selected_unit:
        try:
            cols = {str(c).strip().lower(): c for c in df_units.columns}
            key = next((cols[k] for k in ("establishment name", "legal name", "legal entity", "company name") if k in cols), None)
            if key is not None:
                unit_col = next((c for c in df_units.columns if str(c).strip().lower() == "unit"), None)
                match = df_units[df_units[unit_col].apply(normalize_unit_key) == normalize_unit_key(selected_unit)]
                if not match.empty and pd.notna(match.iloc[0][key]) and str(match.iloc[0][key]).strip():
                    est = str(match.iloc[0][key]).strip()
        except Exception:
            pass
    employer = ""
    if df_units is not None and selected_unit:
        try:
            employer = str(get_unit_master_column_value(df_units, selected_unit, "N") or "").strip()
        except Exception:
            employer = ""
    return {"est_name": est, "address": address, "employer_address": employer or address, "unit": unit}


def _header_write(sheet, coordinate: str, value) -> None:
    """safe_write for title-block values, always in black text (some templates mark the
    month placeholder in red)."""
    safe_write(sheet, coordinate, value)
    cell = sheet[coordinate]
    if isinstance(cell, MergedCell):
        for rng in sheet.merged_cells.ranges:
            if coordinate in rng:
                cell = sheet.cell(rng.min_row, rng.min_col)
                break
    font = copy(cell.font)
    font.color = "FF000000"
    cell.font = font


def inject_dynamic_headers(sheet, unit_name, unit_address, selected_month, selected_year,
                           header_row: int | None = None, values: dict[str, str] | None = None) -> list[str]:
    """Write establishment name / address (and employer name+address) into the sheet's title
    block. Works on the rows ABOVE the data-table header only, so table column headings such as
    'Name & Address of nominee' are never touched. Handles both layouts:
      * label and value in one cell  ('Address: old value')  -> replaced in place
      * label cell + separate value cell ('Address' | ... | 'old value') -> value cell written,
        label left exactly as the template has it.
    Returns a log of what was written (used by the diagnostics)."""
    if values is None:
        values = {"est_name": unit_name or "", "address": unit_address or "",
                  "employer_address": unit_address or "", "unit": unit_name or ""}
    limit = (header_row if header_row else 25)
    limit = max(1, min(limit, 25))
    log: list[str] = []

    # first pass: find the labels
    found = []
    for row_idx in range(1, limit):
        for col_idx in range(1, min(sheet.max_column, 30) + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str) or not cell.value.strip():
                continue
            label, existing, sep = _split_label_value(cell.value)
            kind = _classify_header_label(label)
            if kind:
                found.append((cell, label, existing, sep, kind))
    kinds_present = {k for *_, k in found}

    for cell, label, existing, sep, kind in found:
        if kind == "est_name_address":
            # a separate 'Address' label elsewhere gets the address; otherwise write both here
            value = values["est_name"] if "address" in kinds_present else f"{values['est_name']}, {values['address']}".strip(", ")
        else:
            value = values.get(kind, "")
        if not value:
            continue
        if kind == "employer_address":
            # templates already carry the employer's name/address -- only fill an empty slot
            if existing:
                continue
            target = _find_value_cell(sheet, cell.row, cell.column)
            if target is not None and target.value in (None, ""):
                _header_write(sheet, target.coordinate, value)
                log.append(f"{cell.coordinate}: {kind} -> {target.coordinate}")
            continue
        if existing:  # label + old value in one cell
            _header_write(sheet, cell.coordinate, f"{label}{sep if sep.strip() else ': '}{value}")
            log.append(f"{cell.coordinate}: {kind} (in place)")
            continue
        target = _find_value_cell(sheet, cell.row, cell.column)
        if target is not None:
            _header_write(sheet, target.coordinate, value)
            log.append(f"{cell.coordinate}: {kind} -> {target.coordinate}")
        else:
            _header_write(sheet, cell.coordinate, f"{label.rstrip(' :')}: {value}")
            log.append(f"{cell.coordinate}: {kind} (appended to label)")
    return log


def _first_of_month(selected_month: str, selected_year: int):
    try:
        month_num = datetime.strptime(str(selected_month)[:3], "%b").month
    except ValueError:
        month_num = datetime.now().month
    return datetime(int(selected_year), month_num, 1)


_DAY_MONTH_YEAR_RE = re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)?\s+" + MONTH_NAME_ALTERNATION + r"[\s,]+(\d{4})\b", re.IGNORECASE)


_NUMERIC_DATE_RE = re.compile(r"\b(\d{1,2})([/\-.])(\d{1,2})([/\-.])(?:19|20)\d{2}\b")


def _ordinal(day: int) -> str:
    suffix = "th" if 10 <= day % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def _refresh_period_range_text(text: str, selected_month: str, selected_year: int) -> str:
    """Re-point 'From 1st June 2026 To 30th June 2026' style ranges to the selected month
    (days <= 15 become the 1st, later days become the last day of the month)."""
    first = _first_of_month(selected_month, selected_year)
    last_day = calendar.monthrange(first.year, first.month)[1]
    month_name = first.strftime("%B")

    def _sub(m: re.Match) -> str:
        day = _ordinal(1 if int(m.group(1)) <= 15 else last_day)
        return f"{day} {month_name} {first.year}"

    text = _DAY_MONTH_YEAR_RE.sub(_sub, text)

    def _sub_numeric(m: re.Match) -> str:  # 01/06/2026, 30-06-2026
        day = 1 if int(m.group(1)) <= 15 else last_day
        return f"{day:0{len(m.group(1))}d}{m.group(2)}{first.month:0{len(m.group(3))}d}{m.group(4)}{first.year}"

    return _NUMERIC_DATE_RE.sub(_sub_numeric, text)


def _month_in_style(original: str, selected_month: str, selected_year: int) -> str:
    """Write the selected month the way the template wrote its own: 'June 2026', 'June-26', 'Jun-2026'..."""
    first = _first_of_month(selected_month, selected_year)
    word = re.match(r"[A-Za-z]+", original.strip()).group(0)
    name = first.strftime("%B") if len(word) > 3 else first.strftime("%b")
    if word.isupper():
        name = name.upper()
    sep_match = re.search(r"[A-Za-z]+([\s\-]*)(\d+)", original)
    sep = sep_match.group(1) if sep_match else "-"
    digits = sep_match.group(2) if sep_match else "00"
    year = str(first.year) if len(digits) == 4 else str(first.year)[-2:]
    return f"{name}{sep or ' '}{year}"


def write_header_month(sheet, selected_month: str, selected_year: int | None = None, header_row: int | None = None) -> None:
    """Make every month / wage-period marker in the title block show the selected month:
      * text cells with an embedded 'Month-YY' / 'Month YYYY' -> that part is replaced
      * 'Month:' style labels with a separate value cell -> value cell written (never the label)
      * stale date cells (a 1st-of-month date left in the template) -> set to the 1st of the month
    Rows at/after the data-table header are never touched (a 'Month' data column stays intact)."""
    if selected_month in {"All", "", None}:
        selected_month = "July"
    if selected_year is None:
        selected_year = datetime.now().year
    keywords = ("month", "forperiod", "periodending", "wageperiod", "period")
    first_day = _first_of_month(selected_month, selected_year)
    month_label = first_day.strftime("%B %Y")
    limit = max(1, min(header_row if header_row else 12, 25))
    done_rows: set[int] = set()

    for row_idx in range(1, limit + 1):
        for col_idx in range(1, min(sheet.max_column, 60) + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            value = cell.value
            if row_idx == limit:
                # the table-header row itself: only refresh explicit dates such as
                # 'For the Wage Period Ending : 30th June 2026' (a plain 'Month' column is left alone)
                if isinstance(value, str) and (_DAY_MONTH_YEAR_RE.search(value) or _NUMERIC_DATE_RE.search(value)):
                    _header_write(sheet, cell.coordinate, _refresh_period_range_text(value, selected_month, selected_year))
                elif isinstance(value, str) and DATE_IN_TEXT_RE.search(value) and any(k in normalize_header(value) for k in keywords):
                    m = DATE_IN_TEXT_RE.search(value)
                    _header_write(sheet, cell.coordinate, value[: m.start()] + _month_in_style(m.group(0), selected_month, selected_year) + value[m.end():])
                continue
            if isinstance(value, datetime):  # stale date left in the template
                if value.day == 1:
                    _header_write(sheet, cell.coordinate, first_day)
                continue
            if not isinstance(value, str):
                continue
            if _DAY_MONTH_YEAR_RE.search(value) or _NUMERIC_DATE_RE.search(value):  # 'From: 1st June 2026  To: 30th June 2026'
                refreshed = _refresh_period_range_text(value, selected_month, selected_year)
                if refreshed != value:
                    _header_write(sheet, cell.coordinate, refreshed)
                continue
            norm = normalize_header(value)
            if not any(k in norm for k in keywords) or len(norm) > 120 or "auto" in norm:
                # no month label, but an explicit 'June 2026' in the title block
                # (e.g. 'Name of the employer : ...      JUNE 2026', 'Year: June 2026')
                full = re.search(MONTH_NAME_ALTERNATION + r"[\s,\-]*(?:19|20)\d{2}\b", value, re.IGNORECASE)
                if full and "auto" not in norm:
                    _header_write(sheet, cell.coordinate, value[: full.start()] + _month_in_style(full.group(0), selected_month, selected_year) + value[full.end():])
                continue
            m = DATE_IN_TEXT_RE.search(value)
            if m:
                _header_write(sheet, cell.coordinate, value[: m.start()] + _month_in_style(m.group(0), selected_month, selected_year) + value[m.end():])
                continue
            mn = re.search(r"\b" + MONTH_NAME_ALTERNATION + r"\b", value, re.IGNORECASE)
            if mn:  # 'Month: June Year: 2026' -> month name and the year are separate
                new_text = value[: mn.start()] + _first_of_month(selected_month, selected_year).strftime("%B") + value[mn.end():]
                new_text = re.sub(r"\b(19|20)\d{2}\b", str(selected_year), new_text)
                _header_write(sheet, cell.coordinate, new_text)
                continue
            if row_idx in done_rows:
                continue
            # label-only cell: find the month value cell to its right (outside any merge)
            label_text, existing, sep = _split_label_value(value)
            if len(norm) > 30 and not existing:
                continue
            if existing:  # 'Month: <something not a month>' -> overwrite value
                _header_write(sheet, cell.coordinate, f"{label_text}{sep if sep.strip() else ': '}{month_label}")
                done_rows.add(row_idx)
                continue
            start = _merged_end_col(sheet, row_idx, col_idx) + 1
            for c in range(start, min(sheet.max_column, start + 8) + 1):
                probe = sheet.cell(row=row_idx, column=c)
                pv = probe.value
                if pv is None or (isinstance(pv, str) and not pv.strip()):
                    if isinstance(probe, MergedCell):
                        continue
                    _header_write(sheet, probe.coordinate, month_label)
                    done_rows.add(row_idx)
                    break
                if isinstance(pv, datetime):
                    _header_write(sheet, probe.coordinate, first_day)
                    done_rows.add(row_idx)
                    break
                ptxt = str(pv).strip()
                if ptxt.lower().rstrip(".") in MONTH_NAMES_LOWER or DATE_IN_TEXT_RE.search(ptxt):
                    _header_write(sheet, probe.coordinate, selected_month if ptxt.lower().rstrip(".") in MONTH_NAMES_LOWER else month_label)
                    # separate year cell after a month-name cell
                    for yc in range(c + 1, min(sheet.max_column, c + 4) + 1):
                        ycell = sheet.cell(row=row_idx, column=yc)
                        if re.fullmatch(r"\d{4}(\.0)?", str(ycell.value or "").strip()):
                            _header_write(sheet, ycell.coordinate, selected_year)
                            break
                    done_rows.add(row_idx)
                    break
                if isinstance(pv, str) and pv.strip().endswith(":"):
                    break


def inject_form_dates(sheet, form_name: str, selected_month: str, selected_year: int) -> None:
    if selected_month in {"All", "", None}:
        selected_month = "July"
    year_suffix = str(selected_year)[-2:]
    month_label = f"{selected_month[:3]}-{year_suffix}"

    if form_name == "Form A":
        cell_a8 = sheet["A8"]
        current_text = str(cell_a8.value or "")
        marker = "shall take effect from (date)"
        marker_idx = current_text.find(marker)
        if marker_idx != -1:
            prefix = current_text[: marker_idx + len(marker)]
            new_text = f"{prefix} {selected_month}-{year_suffix}"
        else:
            new_text = current_text
        safe_write(sheet, "A8", new_text)
    elif form_name == "Form E":
        safe_write(sheet, "A5", f"Wage Period- {selected_month}-{selected_year}")
    elif form_name == "Form IV":
        safe_write(sheet, "A4", f"Month Ending -{selected_month.upper()} {year_suffix}")
    elif form_name == "Form V":
        safe_write(sheet, "A5", f"For the period ending - {selected_month}-{selected_year}")


def verify_generated_form(
    file_path: str | os.PathLike[str],
    form_name: str | None = None,
    known_codes: set[str] | None = None,
    selected_state: str | None = None,
) -> list[str]:
    errors: list[str] = []
    try:
        workbook = load_workbook(file_path, data_only=True)
        sheet = workbook.active
    except Exception as exc:
        return [f"Unable to open generated workbook for QA: {exc}"]

    start_row = get_form_start_row(selected_state, form_name, default=15)
    known_codes = known_codes or set()

    for row_idx in range(1, start_row):
        for col_idx in range(1, 15):
            value = str(sheet.cell(row=row_idx, column=col_idx).value or "")
            if not value.strip():
                continue
            if "ALL-" in value.upper() or "ALL 20" in value.upper():
                errors.append(f"Invalid date string '{value}' found at row {row_idx}, col {col_idx}")
            if "empl_code" in value.lower():
                errors.append(f"Leaked column-header text '{value}' found in header block at row {row_idx}, col {col_idx}")
            elif value.strip() in known_codes:
                errors.append(f"Leaked employee-code value '{value}' found in header block at row {row_idx}, col {col_idx}")

    name_found = False
    for col_idx in range(1, 6):
        candidate_value = str(sheet.cell(row=start_row, column=col_idx).value or "").strip()
        if candidate_value:
            name_found = True
            break
    if not name_found:
        errors.append(f"Employee Name missing at starting row {start_row}")

    for row_idx in range(start_row, min(start_row + 3, sheet.max_row + 1)):
        row_has_data = any(str(sheet.cell(row=row_idx, column=col_idx).value or "").strip() for col_idx in range(1, 8))
        if row_has_data:
            first_value = str(sheet.cell(row=row_idx, column=1).value or "").strip()
            second_value = str(sheet.cell(row=row_idx, column=2).value or "").strip()
            third_value = str(sheet.cell(row=row_idx, column=3).value or "").strip()
            if not any([first_value, second_value, third_value]):
                errors.append(f"Empty employee row detected at row {row_idx}")

    return errors


def find_form_rule_row(df_mapping_rules: pd.DataFrame | None, form_name: str | None, selected_state: str | None) -> int | None:
    if df_mapping_rules is None or df_mapping_rules.empty or form_name is None:
        return None
    if df_mapping_rules.shape[1] < 2:
        return None
    target = str(form_name).strip().upper()
    target_state = str(selected_state).strip().upper() if selected_state is not None else ""
    for row_idx in range(df_mapping_rules.shape[0]):
        candidate = df_mapping_rules.iloc[row_idx, 1]
        state_candidate = df_mapping_rules.iloc[row_idx, 0]
        state_value = str(state_candidate).strip().upper() if pd.notna(state_candidate) else ""
        if pd.notna(candidate) and str(candidate).strip().upper() == target and state_value == target_state:
            return row_idx
    return None

def _is_numeric_or_formula_value_for_totals(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        if value.startswith("="):
            return True
        candidate = value.strip().replace(",", "")
        if candidate:
            try:
                float(candidate)
                return True
            except ValueError:
                return False
    return False

NON_SUMMABLE_HEADER_TOKENS = {
    "signature", "account", "aadhar", "aadhaar", "pan", "uan", "regn",
    "registration", "licence", "license", "mobile", "phone", "ward", "circle",
    "number", "code", "serial",
}

def _is_identifier_header(header_value: object) -> bool:
    header_raw = str(header_value or "").lower()
    tokens = set(re.findall(r"[a-z]+", header_raw))
    return bool(tokens & NON_SUMMABLE_HEADER_TOKENS) or "a/c" in header_raw

def write_totals_row(sheet, header_row: int, start_row: int, last_data_row: int, data_col_count: int) -> int:
    totals_row_idx = last_data_row + 1
    label_written = False

    for col_idx in range(1, data_col_count + 1):
        col_letter = get_column_letter(col_idx)
        target_cell = sheet.cell(row=totals_row_idx, column=col_idx)
        source_cell = sheet.cell(row=last_data_row, column=col_idx)

        copy_cell_style(source_cell, target_cell)
        target_cell.fill = PatternFill(fill_type=None)
        target_cell.value = None

        if col_idx <= 2:
            continue

        header_text = sheet.cell(row=header_row, column=col_idx).value
        is_identifier_column = _is_identifier_header(header_text)

        column_values = [
            sheet.cell(row=r, column=col_idx).value for r in range(start_row, last_data_row + 1)
        ]
        populated_values = [v for v in column_values if v is not None and str(v).strip() != ""]

        if (
            not is_identifier_column
            and populated_values
            and all(_is_numeric_or_formula_value_for_totals(v) for v in populated_values)
        ):
            target_cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last_data_row})"
            target_cell.font = Font(name=source_cell.font.name, size=source_cell.font.size, bold=True)
        elif not label_written and populated_values:
            target_cell.value = "Total"
            target_cell.font = Font(name=source_cell.font.name, size=source_cell.font.size, bold=True)
            label_written = True

    return totals_row_idx


def apply_smart_formatting(sheet, start_row: int, last_content_row: int, data_col_count: int) -> None:
    """
    Dynamically adjusts column widths and row heights based ONLY on the data rows.
    This prevents ###### errors without destroying the rotated header layouts of Gov forms.
    Also normalizes font weight and computes an explicit fallback row height, since not
    every spreadsheet application (Google Sheets, LibreOffice) auto-calculates wrapped-text
    row height the way Excel does when height is left as None.
    """
    if last_content_row < start_row:
        return

    column_widths: dict[int, float] = {}

    for col_idx in range(1, data_col_count + 1):
        col_letter = get_column_letter(col_idx)
        max_data_len = 0

        # Scan only the DATA rows to determine required width
        for row_idx in range(start_row, last_content_row + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)

            # Force wrap text on data cells to allow vertical expansion
            if cell.alignment:
                cell.alignment = Alignment(
                    horizontal=cell.alignment.horizontal,
                    vertical=cell.alignment.vertical,
                    wrap_text=True,
                    text_rotation=cell.alignment.text_rotation
                )
            else:
                cell.alignment = Alignment(wrap_text=True)

            # Normalize font weight so a data row can't inherit stray bold/non-bold
            # inconsistency from whichever column happened to be used as the style
            # reference. Data rows are always regular weight; the totals row is
            # bolded separately by write_totals_row, which runs after this.
            if cell.font:
                cell.font = Font(
                    name=cell.font.name,
                    size=cell.font.size,
                    bold=False,
                    italic=cell.font.italic,
                    color=cell.font.color,
                )

            val = cell.value
            if val is not None and str(val).strip() != "":
                # Estimate visual width of the data string
                lines = str(val).split('\n')
                longest_line = max(len(line) for line in lines)
                max_data_len = max(max_data_len, longest_line)

        # Gently adjust column width if the data is wide, capping at 35 so it doesn't stretch infinitely
        current_width = sheet.column_dimensions[col_letter].width
        if current_width is None:
            current_width = 8.5  # Excel default
        if max_data_len > 0:
            required_width = max_data_len + 1.2
            if required_width > current_width:
                current_width = min(required_width, 35)
                sheet.column_dimensions[col_letter].width = current_width
        column_widths[col_idx] = current_width

    # Row height: let Excel auto-fit where it can (height=None), but also compute an
    # explicit fallback so wrapped multi-line text still displays correctly in viewers
    # that don't recalculate wrapped-row height on open.
    for row_idx in range(start_row, last_content_row + 1):
        max_lines_needed = 1
        for col_idx in range(1, data_col_count + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is None or str(val).strip() == "":
                continue
            width = max(column_widths.get(col_idx, 8.5), 1)
            for line in str(val).split('\n'):
                lines_needed = max(1, -(-len(line) // int(width)))  # ceiling division
                max_lines_needed = max(max_lines_needed, lines_needed)
        sheet.row_dimensions[row_idx].height = max(15, max_lines_needed * 15)


def run_diagnostic_mapping_test(form_name, filtered_df, df_mapping_rules, df_col_ref, template_bytes, selected_state=None):
    if filtered_df.empty:
        st.error("No employees found to test mapping.")
        return None

    row = filtered_df.iloc[0].to_dict()
    template_bytes.seek(0)
    workbook = load_workbook(template_bytes)
    sheet = workbook.active
    
    header_row = find_header_row(sheet)
    header_cells = [sheet.cell(row=header_row, column=col_idx).value for col_idx in range(1, 52)]
    
    start_row = get_form_start_row(selected_state, form_name, default=max(header_row + 2, 15))
    if _looks_like_numbering_row(sheet, start_row):
        start_row += 1
    form_rule_row = find_form_rule_row(df_mapping_rules, form_name, selected_state)
    
    logs = []
    
    for col_idx in range(1, 25):
        col_letter = get_column_letter(col_idx)
        header_val = str(header_cells[col_idx-1]).replace("\n", " ") if col_idx <= len(header_cells) else ""
        source_cell = sheet.cell(row=start_row, column=col_idx)
        
        is_ref = is_reference_cell(source_cell)
        is_formula = isinstance(source_cell.value, str) and source_cell.value.startswith("=")
        
        rule = ""
        if form_rule_row is not None:
            rule_col_index = col_idx + 6
            if rule_col_index < df_mapping_rules.shape[1]:
                rule_candidate = df_mapping_rules.iloc[form_rule_row, rule_col_index]
                if pd.notna(rule_candidate):
                    rule = str(rule_candidate).strip()
                    if rule == 'nan': rule = ""
                
        mapped_target = ""
        extracted_value = ""
        note = ""

        if is_ref:
            extracted_value = "[SKIPPED: YELLOW TEMPLATE CELL]"
            note = f"Protected template constant: {source_cell.value}"
        elif is_formula:
            extracted_value = "[SKIPPED: TEMPLATE FORMULA]"
            note = f"Protected template constant: {source_cell.value}"
        elif _is_col_prefixed(rule):
            registry = get_col_registry(df_col_ref)
            tokens, parse_notes = tokenize_col_formula(rule, registry)
            sub_letters = [t[1] for t in tokens if t[0] == "tag"]
            mapped_target = " ".join(
                (t[1] if t[0] in ("tag", "op") else str(t[1]) if t[0] == "num" else t[0]) for t in tokens
            )
            missing_letters = [l for l in sub_letters if l not in registry]

            resolved_value = resolve_col_rule(rule, df_col_ref, row, 0)
            if resolved_value not in (None, ""):
                extracted_value = str(resolved_value)
                note = "✅ Exact match successful" if len(tokens) == 1 else "✅ Formula evaluated"
            elif missing_letters:
                note = f"❌ Rule letter(s) {', '.join(missing_letters)} not found in Coloums mapping sheet!"
            else:
                extracted_value = "BLANK"
                note = f"❌ TARGET '{mapped_target}' resolved but this employee's data is blank for it."
            if parse_notes:
                note += " (" + "; ".join(parse_notes) + ")"
        elif rule:
            resolved_value = resolve_col_rule(rule, df_col_ref, row, 0)
            extracted_value = str(resolved_value) if resolved_value not in (None, "") else "BLANK"
            note = f"Evaluated via manual rule: {rule}"
        else:
            extracted_value = "..."
            note = "Fallback to direct header matching"

        logs.append({
            "Excel Column": f"{col_letter} ({col_idx})",
            "Template Header": header_val[:35] + "..." if len(header_val) > 35 else header_val,
            "Sheet1 Rule": rule,
            "Mapped To (Coloums)": mapped_target,
            "Value Found in Data": extracted_value,
            "Diagnostic Note": note
        })
        
    logs_df = pd.DataFrame(logs)
    st.markdown("This table shows exactly how the code is matching data to the template columns for the first employee. Use this to spot misspelled headers or missing data.")
    st.dataframe(logs_df, use_container_width=True)
    return logs_df


def generate_dynamic_form(
    filtered_df: pd.DataFrame,
    template_source: io.BytesIO,
    selected_month: str,
    selected_year: int,
    form_name: str | None = None,
    selected_state: str | None = None,
    df_mapping_rules: pd.DataFrame | None = None,
    df_col_ref: pd.DataFrame | None = None,
    df_units: pd.DataFrame | None = None,
    selected_unit: str | None = None,
) -> BytesIO:
    template_source.seek(0)
    workbook = load_workbook(template_source)
    sheet = workbook.active

    selected_month = selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July"

    unit_master_details = get_unit_master_details(df_units, selected_unit)
    header_row = find_header_row(sheet)

    write_header_month(sheet, selected_month, selected_year, header_row=header_row)
    inject_dynamic_headers(
        sheet,
        unit_master_details.get("unit_name", ""),
        unit_master_details.get("unit_address", ""),
        selected_month,
        selected_year,
        header_row=header_row,
        values=get_header_values(unit_master_details, df_units, selected_unit),
    )
    if form_name:
        inject_form_dates(sheet, form_name, selected_month, selected_year)

    header_cells = [sheet.cell(row=header_row, column=col_idx).value for col_idx in range(1, 51)]

    form_key = str(form_name).strip().upper() if form_name else ""
    start_row = get_form_start_row(selected_state, form_name, default=max(header_row + 2, 15))
    if _looks_like_numbering_row(sheet, start_row):
        start_row += 1
    _clear_internal_annotation_row(sheet, start_row - 1)
    table_end_row = detect_table_end_row(sheet, start_row)
    row_count = len(filtered_df)

    active_month = (selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July")
    month_label = f"{active_month[:3]}-{str(selected_year)[-2:]}"
    wage_month_columns = {"Form E": "E", "Form D": "F"}
    form_rule_row = find_form_rule_row(df_mapping_rules, form_name, selected_state)

    data_col_count = min(sheet.max_column, 51)

    for offset in range(row_count):
        row = filtered_df.iloc[offset].to_dict()
        row_idx = start_row + offset
        for col_idx in range(1, data_col_count + 1):
            target_cell = sheet.cell(row=row_idx, column=col_idx)
            source_cell = sheet.cell(row=start_row, column=col_idx)
            copy_cell_style(source_cell, target_cell)

            if is_reference_cell(source_cell):
                continue

            if isinstance(target_cell.value, str) and target_cell.value.startswith("="):
                continue

            rule_value = ""
            if df_mapping_rules is not None and not df_mapping_rules.empty and form_rule_row is not None:
                rule_col_index = col_idx + 6
                if rule_col_index < df_mapping_rules.shape[1]:
                    candidate = df_mapping_rules.iloc[form_rule_row, rule_col_index]
                    if pd.notna(candidate) and str(candidate).strip():
                        rule_value = str(candidate).strip()

            if rule_value:
                cell_value = resolve_col_rule(
                    rule_value,
                    df_col_ref,
                    row,
                    offset,
                    active_month=active_month,
                    selected_year=selected_year,
                    df_units=df_units,
                    selected_unit=selected_unit,
                )
            else:
                header_value = header_cells[col_idx - 1] if col_idx <= len(header_cells) else None
                cell_value = get_value_for_header(
                    header_value,
                    row,
                    unit_master_details=unit_master_details,
                    df_mapping_rules=df_mapping_rules,
                    df_col_ref=df_col_ref,
                    idx=offset,
                    df_units=df_units,
                    selected_unit=selected_unit,
                )
            
            safe_write(sheet, f"{get_column_letter(col_idx)}{row_idx}", cell_value)

        if form_name in wage_month_columns:
            month_col_letter = wage_month_columns[form_name]
            month_col_idx = column_letter_to_index(month_col_letter)
            month_cell = sheet.cell(row=row_idx, column=month_col_idx)
            copy_cell_style(sheet.cell(row=start_row, column=month_col_idx), month_cell)
            safe_write(sheet, month_cell.coordinate, month_label)

    # ADD TOTALS ROW SAFELY
    totals_row_idx = start_row - 1
    if row_count > 0:
        last_data_row = start_row + row_count - 1
        totals_row_idx = write_totals_row(sheet, header_row, start_row, last_data_row, data_col_count)

    # Clear remaining unused template rows inside the blue/yellow boundary (Leaves footers untouched)
    clear_start_row = totals_row_idx + 1
    if table_end_row >= clear_start_row:
        amount_to_delete = table_end_row - clear_start_row + 1
        sheet.delete_rows(clear_start_row, amount_to_delete)

    # VISUAL FIX 2: Smart Auto-Fit (Prevents ###### without destroying headers)
    last_content_row = totals_row_idx if row_count > 0 else start_row - 1
    apply_smart_formatting(sheet, start_row, last_content_row, data_col_count)

    # VISUAL FIX 3: Strip Background Colors Only (Preserves Text Rotation!)
    for row_idx in range(start_row, last_content_row + 1):
        for col_idx in range(1, data_col_count + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            # Remove the yellow/blue highlights
            cell.fill = PatternFill(fill_type=None)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# Mapping Editor helpers (template header <-> Conso Data header -> Sheet1 rule)
# ---------------------------------------------------------------------------
RULE_KIND_KEEP = "Keep current rule"
RULE_KIND_CONSO = "Conso Data column"
RULE_KIND_FORMULA = "Formula (Conso columns with + - × ÷)"
RULE_KIND_FIXED = "Fixed text"
RULE_KIND_AUTO = "Serial number (auto)"
RULE_KIND_WAGE_MONTH = "Wage month"
RULE_KIND_UNIT = "Unit Master column(s)"
RULE_KIND_BLANK = "Blank (clear rule)"
RULE_KINDS = [RULE_KIND_KEEP, RULE_KIND_CONSO, RULE_KIND_FORMULA, RULE_KIND_FIXED,
              RULE_KIND_AUTO, RULE_KIND_WAGE_MONTH, RULE_KIND_UNIT, RULE_KIND_BLANK]
FORMULA_OPERATORS = {"+": "+", "−": "-", "×": "*", "÷": "/"}


def conso_header_options(df_col_ref: pd.DataFrame | None) -> dict[str, str]:
    """{tag: 'BU · Earned Gross Salary'} in Coloums-sheet order, skipping separator rows."""
    options: dict[str, str] = {}
    if df_col_ref is None or df_col_ref.empty or df_col_ref.shape[1] < 2:
        return options
    for code, name in zip(df_col_ref.iloc[:, 0], df_col_ref.iloc[:, 1]):
        if pd.isna(code) or pd.isna(name):
            continue
        tag = _clean_tag(code)
        label = str(name).strip()
        if not tag or label in {":", ""} or tag in options:
            continue
        options[tag] = f"{tag} · {label}"
    return options


def build_rule_string(kind: str, *, tags: list[str] | None = None, operators: list[str] | None = None,
                      text: str = "", unit_letters: list[str] | None = None) -> str:
    """The exact text stored in Sheet1 for a UI selection (round-trips through resolve_col_rule)."""
    if kind == RULE_KIND_CONSO and tags:
        return f"Col_{tags[0]}"
    if kind == RULE_KIND_FORMULA and tags:
        parts = [f"Col_{tags[0]}"]
        for op, tag in zip(operators or [], tags[1:]):
            parts.append(f"{FORMULA_OPERATORS.get(op, op)} Col_{tag}")
        return " ".join(parts)
    if kind == RULE_KIND_FIXED:
        return f"{text.strip()} (Fixed Text)" if text.strip() else ""
    if kind == RULE_KIND_AUTO:
        return "Auto_generated"
    if kind == RULE_KIND_WAGE_MONTH:
        return "Wage Month"
    if kind == RULE_KIND_UNIT and unit_letters:
        return " + ".join(f"Unit Master_Col_{l}" for l in unit_letters)
    return ""


def describe_rule(rule: object, df_col_ref: pd.DataFrame | None) -> str:
    """Human readable version of a stored rule, e.g. 'Col_BU - Col_BC' -> 'Earned Gross Salary − Earned Basic'."""
    text = "" if rule is None or (not isinstance(rule, str) and pd.isna(rule)) else str(rule).strip()
    if not text:
        return ""
    if not _is_col_prefixed(text):
        return text
    registry = get_col_registry(df_col_ref)
    tokens, _ = tokenize_col_formula(text, registry)
    pretty = {"+": "+", "-": "−", "*": "×", "/": "÷"}
    out = []
    for tok in tokens:
        if tok[0] == "tag":
            out.append(registry.get(tok[1], f"?{tok[1]}"))
        elif tok[0] == "op":
            out.append(pretty[tok[1]])
        elif tok[0] == "num":
            out.append(f"{tok[1]:g}")
        else:
            out.append(tok[0])
    return " ".join(out)


def suggest_conso_tag(template_label: str, conso_options: dict[str, str]) -> str | None:
    """Best-guess Conso Data tag for a template header (shown as a hint, never applied automatically)."""
    import difflib
    target = normalize_header(template_label.split("›")[-1])
    if not target or target in {"noheader"}:
        return None
    best_tag, best_score = None, 0.0
    for tag, label in conso_options.items():
        name = normalize_header(label.split("·", 1)[-1])
        if not name:
            continue
        score = difflib.SequenceMatcher(None, target, name).ratio()
        if name in target or target in name:
            score = max(score, 0.75)
        if score > best_score:
            best_tag, best_score = tag, score
    return best_tag if best_score >= 0.6 else None


def get_template_columns(template_bytes, selected_state: str | None, form_name: str | None, max_cols: int = 51) -> list[dict]:
    """Template column list [{'col': 1, 'letter': 'A', 'label': 'Code'}, ...] built from the
    table's header rows (merged parent headings are joined with their sub-headings)."""
    template_bytes.seek(0)
    sheet = load_workbook(template_bytes).active
    header_row = find_header_row(sheet)
    start_row = get_form_start_row(selected_state, form_name, default=max(header_row + 2, 15))
    last_header_row = max(header_row, start_row - 1)
    anchors = {}
    for rng in sheet.merged_cells.ranges:
        anchors[(rng.min_row, rng.min_col, rng.max_row, rng.max_col)] = sheet.cell(rng.min_row, rng.min_col).value

    def text_at(r: int, c: int):
        cell = sheet.cell(r, c)
        if cell.value is not None:
            return cell.value
        for (r1, c1, r2, c2), val in anchors.items():
            if r1 <= r <= r2 and c1 <= c <= c2:
                return val
        return None

    columns = []
    last_col = min(sheet.max_column, max_cols)
    for c in range(1, last_col + 1):
        parts: list[str] = []
        for r in range(header_row, last_header_row + 1):
            v = text_at(r, c)
            if v is None:
                continue
            t = re.sub(r"\s+", " ", str(v)).strip()
            if not t or re.fullmatch(r"\d{1,2}(\.0)?", t) or t in parts:
                continue
            parts.append(t)
        columns.append({"col": c, "letter": get_column_letter(c), "label": " › ".join(parts) or "(no header)"})
    while columns and columns[-1]["label"] == "(no header)":
        columns.pop()
    return columns


def get_rule_for_column(df_mapping_rules: pd.DataFrame, rule_row: int, col_idx: int) -> str:
    j = col_idx + 6
    if j >= df_mapping_rules.shape[1]:
        return ""
    v = df_mapping_rules.iloc[rule_row, j]
    return "" if pd.isna(v) else str(v).strip()


def save_rule_edits(master_path: str, rule_row: int, edits: dict[int, str]) -> bytes:
    """Write {template_col_idx: rule_text} into Sheet1 of the master workbook (all other
    sheets are left as they are) and return the saved file's bytes."""
    workbook = load_workbook(master_path)
    sheet = workbook["Sheet1"]
    for col_idx, rule_text in edits.items():
        sheet.cell(row=rule_row + 1, column=col_idx + 7).value = rule_text if rule_text != "" else None
    workbook.save(master_path)
    with open(master_path, "rb") as f:
        return f.read()



st.set_page_config(page_title="HR Compliance Engine", layout="wide")


def check_password() -> bool:
    if st.session_state.get("authenticated", False):
        return True

    st.title("🔒 HR Compliance Portal - Secure Login")
    st.caption("Please enter your organization access code to open the dashboard.")
    password_input = st.text_input("Access Code", type="password", key="login_code")
    if st.button("Log In", type="primary"):
        correct_code = st.secrets.get("app_security", {}).get("access_code", "HungerPangs2026")
        if password_input == correct_code:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Incorrect Access Code. Please contact HR IT.")

    return False


if not check_password():
    st.stop()

st.markdown(custom_css, unsafe_allow_html=True)

if "forms_generated" not in st.session_state:
    st.session_state.forms_generated = False

if "generated_forms" not in st.session_state:
    st.session_state.generated_forms = {}

header_html = """
<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-bottom: 1.5rem;">
    <img src="https://drive.google.com/thumbnail?id=1OpUw3MCFGLRs7GQ4xezk5ouqYTvy6yv9&sz=w1000"
         style="max-height: 120px; width: auto; object-fit: contain; margin-bottom: 12px; filter: drop-shadow(0 4px 6px rgba(0,0,0,0.1));">
    <h1 style="color: #1d4ed8; font-family: 'Inter', sans-serif; font-weight: 800; font-size: 2.1rem; margin: 0; padding: 0; text-align: center;">
        HR Statutory Compliance Engine
    </h1>
</div>
"""
st.markdown(header_html, unsafe_allow_html=True)

# Load a cached/simple master view for uploader validation
try:
    df_master = get_active_master_data()
except Exception as e:
    st.error(f"❌ Failed to fetch or read Master Data from OneDrive: {e}")
    if st.button("🔄 Retry Connection to OneDrive"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# Sidebar: Update Master Data
st.sidebar.subheader("📥 Update Master Data")

if st.sidebar.button("🔄 Sync Live Mapping Rules", type="secondary"):
    with st.spinner("Fetching latest rules from OneDrive..."):
        try:
            content = fetch_onedrive_file_with_retry(ONEDRIVE_MASTER_URL)
            fresh_bytes = io.BytesIO(content)
            
            df_sheet1 = pd.read_excel(fresh_bytes, sheet_name="Sheet1", header=None, engine="openpyxl")
            df_coloums = pd.read_excel(fresh_bytes, sheet_name="Coloums", header=None, engine="openpyxl")
            
            local_wb = load_workbook(MASTER_FILE_PATH)
            
            if "Sheet1" in local_wb.sheetnames:
                del local_wb["Sheet1"]
            ws1 = local_wb.create_sheet("Sheet1")
            for r_idx, row in enumerate(df_sheet1.itertuples(index=False, name=None), 1):
                for c_idx, val in enumerate(row, 1):
                    ws1.cell(row=r_idx, column=c_idx, value=val if pd.notna(val) else None)
                    
            if "Coloums" in local_wb.sheetnames:
                del local_wb["Coloums"]
            wsc = local_wb.create_sheet("Coloums")
            for r_idx, row in enumerate(df_coloums.itertuples(index=False, name=None), 1):
                for c_idx, val in enumerate(row, 1):
                    wsc.cell(row=r_idx, column=c_idx, value=val if pd.notna(val) else None)
                    
            local_wb.save(MASTER_FILE_PATH)
            st.cache_data.clear()
            st.sidebar.success("✅ Rules synced perfectly! Run the diagnostic again.")
        except Exception as e:
            st.sidebar.error(f"❌ Sync failed: {e}")

months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
years = [2025, 2026, 2027, 2028, 2029, 2030]

upload_month = st.sidebar.selectbox("Select Data Month", months)
upload_year = st.sidebar.selectbox("Select Data Year", years, index=1)

_legacy_row_count = (
    int(df_master["Month-Year"].isna().sum()) if "Month-Year" in df_master.columns else len(df_master)
)
clear_legacy_rows = False
if _legacy_row_count:
    st.sidebar.warning(
        f"⚠️ {_legacy_row_count} existing rows have no Month-Year (from before monthly tracking was added). "
        "They'll be kept as-is unless cleared below."
    )
    clear_legacy_rows = st.sidebar.checkbox(
        f"Clear these {_legacy_row_count} untracked legacy rows on next upload",
        value=False,
        help="Leave unchecked to keep them untouched alongside every future monthly upload. "
             "Check this once your monthly uploads have caught up to replace them.",
    )

uploaded_file = st.sidebar.file_uploader("Upload Monthly Conso Data (.xlsx)", type=["xlsx"])
if uploaded_file and st.sidebar.button("Validate & Replace Master Data", type="primary"):
    try:
        uploaded_bytes = uploaded_file.getbuffer().tobytes()

        # --- Smart Header Detector ---
        new_conso_df = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=0, engine="openpyxl")
        new_conso_df.columns = new_conso_df.columns.astype(str).str.strip()
        missing_cols_normal = set(df_master.columns) - set(new_conso_df.columns)

        if len(missing_cols_normal) > 5:
            new_conso_df_skip1 = pd.read_excel(
                io.BytesIO(uploaded_bytes), sheet_name=0, engine="openpyxl", skiprows=1
            )
            new_conso_df_skip1.columns = new_conso_df_skip1.columns.astype(str).str.strip()
            missing_cols_skip1 = set(df_master.columns) - set(new_conso_df_skip1.columns)

            if len(missing_cols_skip1) < len(missing_cols_normal):
                new_conso_df = new_conso_df_skip1
        # --- End Smart Header Detector ---

        new_conso_df["Year"] = int(upload_year)
        new_conso_df["Month-Year"] = f"{upload_month}-{upload_year}"
        new_conso_df["Month-Year"] = new_conso_df["Month-Year"].astype(str).str.strip()
    except Exception as e:
        st.sidebar.error(f"❌ Unable to read uploaded file: {e}")
        st.stop()

    missing_cols = set(df_master.columns) - set(new_conso_df.columns)
    if missing_cols:
        st.sidebar.error(f"❌ Missing columns: {missing_cols}")
    else:
        try:
            existing_df = pd.read_excel(MASTER_FILE_PATH, sheet_name="Conso_Data", skiprows=1, engine="openpyxl")
            existing_df.columns = existing_df.columns.astype(str).str.strip()
        except Exception as e:
            st.sidebar.error(f"❌ Failed reading existing master data before append: {e}")
            st.stop()

        code_col = next((c for c in existing_df.columns if "code" in c.lower()), None)
        if code_col and code_col in new_conso_df.columns:
            existing_df[code_col] = existing_df[code_col].astype(str).str.strip()
            new_conso_df[code_col] = new_conso_df[code_col].astype(str).str.strip()

        target_month_year = new_conso_df["Month-Year"].iloc[0] if not new_conso_df.empty else f"{upload_month}-{upload_year}"

        if "Month-Year" in existing_df.columns:
            existing_my = existing_df["Month-Year"].astype(str).str.strip()
            is_legacy = existing_df["Month-Year"].isna() | existing_my.isin({"", "nan", "None"})
            is_target_month = (~is_legacy) & existing_my.eq(str(target_month_year).strip())
        else:
            is_legacy = pd.Series(True, index=existing_df.index)
            is_target_month = pd.Series(False, index=existing_df.index)

        keep_mask = ~is_target_month
        if clear_legacy_rows:
            keep_mask &= ~is_legacy
        existing_minus_month = existing_df[keep_mask]
        months_replaced = int(is_target_month.sum())
        legacy_cleared = int(is_legacy.sum()) if clear_legacy_rows else 0

        combined_df = pd.concat([existing_minus_month, new_conso_df], ignore_index=True, sort=False)

        try:
            write_conso_data_sheet(MASTER_FILE_PATH, combined_df)
        except Exception as e:
            st.sidebar.error(f"❌ Failed updating Conso_Data sheet in master workbook: {e}")
            st.stop()

        with open(MASTER_FILE_PATH, "rb") as f:
            updated_master_bytes = f.read()

        email_success_1, msg_1 = email_file_to_outlook(updated_master_bytes, "Master_CONSO_DATA_ALL_UNITS.XLSX")

        archive_buffer = BytesIO()
        with pd.ExcelWriter(archive_buffer, engine="openpyxl") as writer:
            new_conso_df.to_excel(writer, sheet_name="Conso_Data", index=False)
        archive_buffer.seek(0)
        backup_filename = f"CONSO_DATA_{upload_month}_{upload_year}.xlsx"
        email_success_2, msg_2 = email_file_to_outlook(archive_buffer.getvalue(), backup_filename)

        st.cache_data.clear()

        if email_success_1 and email_success_2:
            months_now = combined_df["Month-Year"].nunique() if "Month-Year" in combined_df.columns else 1
            replace_note = f" (replaced {int(months_replaced)} existing rows for this month)" if months_replaced else ""
            legacy_note = f" Cleared {legacy_cleared} untracked legacy rows." if legacy_cleared else ""
            st.sidebar.success(
                f"✅ Appended {target_month_year} ({len(new_conso_df)} employees){replace_note}.{legacy_note} "
                f"Master now holds {months_now} month(s), {len(combined_df)} total rows. Archived as {backup_filename}!"
            )
        else:
            st.sidebar.warning(f"⚠️ Updated locally, but OneDrive sync issue: {msg_1} | {msg_2}")

        st.rerun()

[df_conso, df_units, df_mapping_rules, df_col_ref, df_emp_master, df_leave_register, df_attendance_register] = load_master_data()
merged_df = build_merged_view()
filtered_df = merged_df.copy()

if "generated_zip" not in st.session_state:
    st.session_state.generated_zip = None

if "last_generated_files" not in st.session_state:
    st.session_state.last_generated_files = []


def reset_forms():
    st.session_state.forms_generated = False
    st.session_state.generated_forms = {}
    st.session_state.generated_zip = None
    st.session_state.last_generated_files = []


def save_form_to_archive(
    form_bytes: BytesIO,
    selected_state: str,
    selected_unit: str,
    selected_month: str,
    selected_year: int,
    form_name: str,
) -> tuple[Path, Path]:
    active_month = (selected_month.split("-")[0] if selected_month != "All" else "July")
    active_month_year = f"{active_month}-{selected_year}"

    target_static_dir = STATIC_ARCHIVE_DIR / str(selected_state) / str(selected_unit) / active_month_year
    target_static_dir.mkdir(parents=True, exist_ok=True)
    static_file_path = target_static_dir / f"{selected_state}_{selected_unit}_{form_name}_{active_month_year}.xlsx"
    with open(static_file_path, "wb") as f:
        f.write(form_bytes.getvalue())

    if ONEDRIVE_ARCHIVE_DIR:
        target_od_dir = ONEDRIVE_ARCHIVE_DIR / str(selected_state) / str(selected_unit) / active_month_year
        target_od_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(static_file_path, target_od_dir / static_file_path.name)
        return static_file_path, target_od_dir

    return static_file_path, target_static_dir


def compile_statutory_forms(
    form_names: list[str],
    filtered_df: pd.DataFrame,
    selected_state: str,
    selected_unit: str,
    selected_month: str,
    selected_year: int,
) -> tuple[dict[str, bytes], list[str], Path | None]:
    generated: dict[str, bytes] = {}
    errors: list[str] = []
    state_urls = st.secrets.get("templates", {}).get(selected_state, {})
    archive_target_dir: Path | None = None
    if not state_urls:
        errors.append(f"No template URLs configured for {selected_state}")
        return generated, errors, archive_target_dir

    active_month = (selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July")

    for form_name in form_names:
        if form_name not in state_urls:
            errors.append(form_name)
            continue
        try:
            response = requests.get(state_urls[form_name], timeout=15)
            response.raise_for_status()
            template_bytes = io.BytesIO(response.content)
            form_bytes = generate_dynamic_form(
                filtered_df,
                template_bytes,
                active_month,
                int(selected_year),
                form_name=form_name,
                selected_state=selected_state,
                df_mapping_rules=df_mapping_rules,
                df_col_ref=df_col_ref,
                df_units=df_units,
                selected_unit=selected_unit,
            )

            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
                temp_file.write(form_bytes.getvalue())
                temp_path = temp_file.name

            known_codes = {
                str(c).strip()
                for c in filtered_df.get("Code", pd.Series(dtype=str)).tolist()
                if str(c).strip()
            }
            qa_errors = verify_generated_form(temp_path, form_name, known_codes=known_codes, selected_state=selected_state)
            os.remove(temp_path)
            if qa_errors:
                errors.append(f"{form_name} QA Warning: {', '.join(qa_errors)}")

            file_path, target_dir = save_form_to_archive(
                form_bytes,
                selected_state,
                selected_unit,
                selected_month,
                int(selected_year),
                form_name,
            )
            archive_target_dir = target_dir
            generated[form_name] = form_bytes.getvalue()
        except Exception as exc:
            errors.append(f"{form_name} ({exc})")

    return generated, errors, archive_target_dir


def build_archive_zip(file_paths: list[Path]) -> BytesIO:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for path in file_paths:
            zip_file.write(path, arcname=path.name)
    zip_buffer.seek(0)
    return zip_buffer


tab1, tab2, tab3, tab4 = st.tabs(["⚡ Statutory Form Generator", "Generated Forms", "Notice Dashboard", "🧩 Mapping Editor"])

with tab1:
    st.markdown("### Selection Filters")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if "State" in merged_df.columns and not merged_df["State"].dropna().eq("").all():
            state_list = ["All"] + sorted(merged_df["State"].dropna().astype(str).unique().tolist())
            selected_state = st.selectbox("State / Region", state_list, key="selected_state", on_change=reset_forms)
        else:
            selected_state = "All"
            st.info("State / Region column not found in the workbook; showing all records.")

    with col2:
        if "Unit_Display" in merged_df.columns:
            if selected_state == "All":
                unit_options = sorted(merged_df["Unit_Display"].dropna().astype(str).unique().tolist())
            else:
                unit_options = sorted(
                    merged_df.loc[merged_df["State"].astype(str) == str(selected_state), "Unit_Display"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
            selected_unit = st.selectbox("Outlet / Brand Unit", ["All"] + unit_options, key="selected_unit", on_change=reset_forms)
        else:
            selected_unit = "All"
            st.info("Unit column not found; using all records.")

    with col3:
        if "Month-Year" in merged_df.columns:
            month_list = ["All"] + sorted(merged_df["Month-Year"].dropna().astype(str).unique().tolist())
        else:
            month_list = ["All"]
        selected_month = st.selectbox("Month-Year", month_list, key="selected_month", on_change=reset_forms)

    with col4:
        current_year = datetime.now().year
        year_options = [str(year) for year in range(current_year - 2, current_year + 3)]
        selected_year = st.selectbox("Year", year_options, index=year_options.index(str(current_year)), key="selected_year", on_change=reset_forms)

    if selected_state != "All" and "State" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["State"].astype(str) == str(selected_state)]

    if selected_unit != "All" and "Unit_Display" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Unit_Display"].astype(str) == str(selected_unit)]

    if selected_month not in {"All", "", None} and "Month-Year" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Month-Year"].astype(str) == str(selected_month)]

    total_emps = len(filtered_df)
    kpi_html = f"""
    <div style="display: flex; gap: 20px; margin-bottom: 20px;">
        <div style="flex: 1; background: #ffffff; padding: 18px 24px; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
            <p style="margin: 0; color: #64748b; font-size: 0.85rem; font-weight: 600;">Total Employees</p>
            <h2 style="margin: 4px 0 0 0; color: #1d4ed8; font-size: 1.8rem; font-weight: 700;">{total_emps}</h2>
        </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    if selected_state == "All" or selected_unit == "All":
        st.warning("⚠️ Please select a specific State and Outlet/Brand Unit to enable form generation.")
    else:
        state_urls = st.secrets.get("templates", {}).get(selected_state, {})
        if state_urls:
            available_forms = sorted(state_urls.keys())
            st.info(f"Available statutory form templates for {selected_state}: {', '.join(available_forms)}")
        else:
            st.warning(f"No statutory form templates configured for {selected_state}.")

        gen_col1, gen_col2, debug_col = st.columns(3)
        with gen_col1:
            selected_form = st.selectbox(
                "Select Statutory Form",
                available_forms if available_forms else ["None"],
                key="selected_form",
            )
            if st.button("Generate Selected Form", key="generate_selected"):
                if selected_month in {"All", "", None}:
                    st.error("❌ Select a specific Month-Year before generating a statutory form - 'All' would mix multiple months' payroll for the same employee, which isn't a valid compliance form.")
                    st.stop()
                with st.status("⚡ Generating forms and pushing to cloud OneDrive...", expanded=True) as status:
                    status.write("📄 Compiling statutory form templates...")
                    generated, errors, archive_dir = compile_statutory_forms(
                        [selected_form],
                        filtered_df,
                        selected_state,
                        selected_unit,
                        selected_month,
                        int(selected_year),
                    )
                    status.write(f"☁️ Auto-pushed to OneDrive folder: `{archive_dir}`")
                    active_month = (selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July")
                    active_month_year = f"{active_month}-{selected_year}"
                    filename = f"{selected_state}_{selected_unit}_{selected_form}_{active_month_year}.xlsx"
                    if selected_form in generated:
                        status.write("📧 Dispatching file to Outlook for cloud OneDrive sync...")
                        email_success, email_msg = email_file_to_outlook(
                            generated[selected_form], filename
                        )
                        if email_success:
                            status.write(f"☁️ {email_msg}")
                        else:
                            status.caption(f"⚠️ {email_msg}")
                    status.update(label="✅ Form generated and synced to OneDrive successfully!", state="complete")
                
                if errors:
                    for error in errors:
                        st.warning(error)
                if generated:
                    st.success(f"✅ Saved to OneDrive: `{archive_dir}`")
                    form_content = generated[selected_form]
                    st.download_button(
                        label=f"📄 Download {selected_form} (.xlsx)",
                        data=form_content,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                        key=f"dl_instant_{selected_form}",
                    )

        with gen_col2:
            if st.button("Generate All Forms (ZIP)", key="generate_bulk"):
                if selected_month in {"All", "", None}:
                    st.error("❌ Select a specific Month-Year before generating statutory forms - 'All' would mix multiple months' payroll for the same employee, which isn't a valid compliance form.")
                    st.stop()
                with st.status("⚡ Generating forms and pushing to cloud OneDrive...", expanded=True) as status:
                    status.write("📄 Compiling statutory form templates...")
                    state_urls = st.secrets.get("templates", {}).get(selected_state, {})
                    all_forms = list(state_urls.keys())
                    generated, errors, archive_dir = compile_statutory_forms(
                        all_forms,
                        filtered_df,
                        selected_state,
                        selected_unit,
                        selected_month,
                        int(selected_year),
                    )
                    status.write(f"☁️ Auto-pushed to OneDrive folder: `{archive_dir}`")
                    active_month = (selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July")
                    active_month_year = f"{active_month}-{selected_year}"
                    if generated:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for form_name, content in generated.items():
                                zip_file.writestr(
                                    f"{selected_state}_{selected_unit}_{form_name}_{active_month_year}.xlsx",
                                    content,
                                )
                        zip_buffer.seek(0)
                        email_filename = f"HR_Compliance_{selected_state}_{selected_unit}_{active_month_year}.zip"
                        status.write("📧 Dispatching archive to Outlook for cloud OneDrive sync...")
                        email_success, email_msg = email_file_to_outlook(zip_buffer, email_filename)
                        if email_success:
                            status.write(f"☁️ {email_msg}")
                        else:
                            status.caption(f"⚠️ {email_msg}")
                    status.update(label="✅ All forms generated and synced to OneDrive successfully!", state="complete")
                
                if errors:
                    for error in errors:
                        st.warning(error)
                if generated:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for form_name, content in generated.items():
                            zip_file.writestr(
                                f"{selected_state}_{selected_unit}_{form_name}_{active_month_year}.xlsx",
                                content,
                            )
                    zip_buffer.seek(0)
                    st.session_state.forms_generated = True
                    st.session_state.generated_forms = generated
                    st.session_state.generated_zip = zip_buffer.getvalue()
                    st.success(f"✅ All forms generated and archived in OneDrive: `{archive_dir}`")
                    st.download_button(
                        label="📦 Download All Generated Forms (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"HR_Compliance_{selected_state}_{selected_unit}_{active_month_year}.zip",
                        mime="application/zip",
                        key="dl_instant_zip",
                    )
                    with st.expander("Individual Form Downloads"):
                        for form_name, content in generated.items():
                            st.download_button(
                                label=f"Download {form_name}",
                                data=content,
                                file_name=f"{selected_state}_{selected_unit}_{form_name}_{active_month_year}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                                key=f"dl_individual_{form_name}",
                            )
        
        with debug_col:
            if st.button("🛠️ Run Diagnostic Data Test (All Forms)", key="run_diagnostic"):
                if selected_month in {"All", "", None}:
                    st.error("❌ Select a specific Month-Year to run the diagnostic test.")
                    st.stop()

                state_urls = st.secrets.get("templates", {}).get(selected_state, {})
                if not state_urls:
                    st.error("No template URLs configured for the selected state.")
                    st.stop()

                diagnostic_frames: dict[str, pd.DataFrame] = {}

                for form_name, template_url in state_urls.items():
                    try:
                        response = requests.get(template_url, timeout=15)
                        response.raise_for_status()
                        template_bytes = io.BytesIO(response.content)
                    except Exception as exc:
                        st.warning(f"⚠️ Could not download template for {form_name}: {exc}")
                        continue

                    st.markdown(f"### Diagnostic Trace for {form_name}")
                    logs_df = run_diagnostic_mapping_test(
                        form_name,
                        filtered_df,
                        df_mapping_rules,
                        df_col_ref,
                        template_bytes,
                        selected_state,
                    )
                    if logs_df is not None and not logs_df.empty:
                        diagnostic_frames[form_name] = logs_df

                if diagnostic_frames:
                    diagnostic_excel = io.BytesIO()
                    used_sheet_names: set[str] = set()
                    with pd.ExcelWriter(diagnostic_excel, engine="openpyxl") as writer:
                        for form_name, logs_df in diagnostic_frames.items():
                            sheet_name = re.sub(r'[\[\]:*?/\\]', "_", form_name).strip()[:31] or "Diagnostic"
                            base_sheet_name = sheet_name
                            suffix = 1
                            while sheet_name in used_sheet_names:
                                suffix_str = f"_{suffix}"
                                sheet_name = f"{base_sheet_name[: 31 - len(suffix_str)]}{suffix_str}"
                                suffix += 1
                            used_sheet_names.add(sheet_name)
                            logs_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    diagnostic_excel.seek(0)

                    active_month_diag = (
                        selected_month.split("-")[0] if selected_month not in {"All", "", None} else "July"
                    )
                    diagnostic_filename = (
                        f"Diagnostic_{selected_state}_{selected_unit}_"
                        f"{active_month_diag}-{selected_year}.xlsx"
                    )
                    st.download_button(
                        label="📥 Download Diagnostic Report (All Forms, .xlsx)",
                        data=diagnostic_excel.getvalue(),
                        file_name=diagnostic_filename,
                        mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                        key="dl_diagnostic_all_forms",
                    )

with tab2:
    st.markdown("### OneDrive Form Archive Browser")
    if st.button("🗑️ Clear All OneDrive Archived Forms", key="clear_all_archive_forms"):
        clear_archive_directory(STATIC_ARCHIVE_DIR)
        if ONEDRIVE_ARCHIVE_DIR:
            clear_archive_directory(ONEDRIVE_ARCHIVE_DIR)
        st.success("All OneDrive archived forms were removed. Refresh the app to confirm the archive is empty.")

    archived_states = sorted([d.name for d in STATIC_ARCHIVE_DIR.iterdir() if d.is_dir()])
    if not archived_states:
        st.info("No archived forms found in OneDrive.")
    else:
        selected_archive_state = st.selectbox("Select Archived State", archived_states, key="archive_state")
        archive_state_dir = STATIC_ARCHIVE_DIR / selected_archive_state
        archived_units = sorted([d.name for d in archive_state_dir.iterdir() if d.is_dir()])
        if not archived_units:
            st.info("No outlets found for the selected archived state.")
        else:
            selected_archive_unit = st.selectbox("Select Archived Outlet", archived_units, key="archive_unit")
            archive_unit_dir = archive_state_dir / selected_archive_unit
            archived_months = sorted([d.name for d in archive_unit_dir.iterdir() if d.is_dir()])
            if not archived_months:
                st.info("No wage-month folders found for the selected outlet.")
            else:
                selected_archive_month = st.selectbox("Select Month-Year", archived_months, key="archive_month")
                archive_month_dir = archive_unit_dir / selected_archive_month
                archived_files = sorted([f for f in archive_month_dir.glob("*.xlsx") if f.is_file()])
                if not archived_files:
                    st.info("No archived forms found for this selection.")
                else:
                    archive_df = pd.DataFrame(
                        [
                            {
                                "Form Name": archive_file.stem.replace(
                                    f"{selected_archive_state}_{selected_archive_unit}_", ""
                                ),
                                "State / Region": selected_archive_state,
                                "Outlet / Brand Unit": selected_archive_unit,
                                "Wage Month": selected_archive_month,
                                "File Size (KB)": round(archive_file.stat().st_size / 1024, 1),
                            }
                            for archive_file in archived_files
                        ]
                    )

                    st.dataframe(archive_df, use_container_width=True, hide_index=True)

                    st.markdown("#### Download Archived Forms")
                    for archive_file in archived_files:
                        form_label = archive_file.stem.replace(f"{selected_archive_state}_{selected_archive_unit}_", "")
                        dl_col1, dl_col2 = st.columns([3, 1])
                        with dl_col1:
                            st.write(f"📄 {form_label}")
                        with dl_col2:
                            st.download_button(
                                label="📥 Download",
                                data=archive_file.read_bytes(),
                                file_name=archive_file.name,
                                mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                                key=f"archive_dl_{selected_archive_state}_{selected_archive_unit}_{selected_archive_month}_{archive_file.name}",
                            )

with tab3:
    st.markdown("###  Notice Closure Portal")
    st.caption("Securely log compliance notices and view the live dashboard below.")
    
    # Live Google Workspace URL embedded securely into Streamlit
    GAS_WEB_APP_URL = "https://script.google.com/a/macros/burmaburma.in/s/AKfycbyiDwUqhB5WgqcAhot3NBnWUiPIJ-0wuKwqhtEF6szXSZVBcr8C6GKu_Yi4XDmyvw8_EQ/exec"
    
    # Embed the HTML web app seamlessly into Streamlit with scrolling enabled for the data table
    components.iframe(GAS_WEB_APP_URL, height=900, scrolling=True)


with tab4:
    st.markdown("### Mapping Editor")
    st.caption(
        "Pair each column of a statutory template with a Conso Data header (or a formula built from "
        "Conso Data headers). Saving writes the matching tag into the Sheet1 rules for you - no more "
        "hand-typing Col_XX tags."
    )

    rule_states = sorted({str(v).strip() for v in df_mapping_rules.iloc[:, 0].dropna() if str(v).strip()}) if not df_mapping_rules.empty else []
    if not rule_states:
        st.warning("Sheet1 has no state/form rows to edit.")
    else:
        me_col1, me_col2 = st.columns(2)
        with me_col1:
            me_state = st.selectbox("State", rule_states, key="me_state")
        me_forms = [
            str(df_mapping_rules.iloc[i, 1]).strip()
            for i in range(df_mapping_rules.shape[0])
            if pd.notna(df_mapping_rules.iloc[i, 0]) and str(df_mapping_rules.iloc[i, 0]).strip().upper() == me_state.upper()
            and pd.notna(df_mapping_rules.iloc[i, 1])
        ]
        with me_col2:
            me_form = st.selectbox("Form", me_forms, key="me_form")
        me_rule_row = find_form_rule_row(df_mapping_rules, me_form, me_state)

        # Template: configured URL first, uploaded file as fallback / override
        me_template = None
        me_upload = st.file_uploader(
            "Template (optional - a blank template or a previously generated form of this type)",
            type=["xlsx"], key=f"me_upload_{me_state}_{me_form}",
        )
        if me_upload is not None:
            me_template = io.BytesIO(me_upload.getvalue())
        else:
            me_url = st.secrets.get("templates", {}).get(me_state, {}).get(me_form)
            if me_url:
                try:
                    resp = requests.get(me_url, timeout=15)
                    resp.raise_for_status()
                    me_template = io.BytesIO(resp.content)
                except Exception as exc:
                    st.info(f"Could not fetch the configured template ({exc}). Upload one above to see its headers.")

        conso_options = conso_header_options(df_col_ref)
        conso_tags = list(conso_options.keys())
        unit_letters_all = [get_column_letter(i + 1) for i in range(len(df_units.columns))]
        unit_letter_labels = {get_column_letter(i + 1): f"{get_column_letter(i + 1)} · {name}" for i, name in enumerate(df_units.columns)}

        if me_template is not None:
            me_columns = get_template_columns(me_template, me_state, me_form)
        else:
            width = max(df_mapping_rules.shape[1] - 7, 0)
            me_columns = [{"col": c, "letter": get_column_letter(c), "label": "(upload template to see header)"}
                          for c in range(1, width + 1) if get_rule_for_column(df_mapping_rules, me_rule_row or 0, c)]

        sample_row = filtered_df.iloc[0].to_dict() if not filtered_df.empty else (df_conso.iloc[0].to_dict() if not df_conso.empty else {})

        if me_rule_row is None:
            st.error(f"No row for {me_state} / {me_form} in Sheet1.")
        elif not me_columns:
            st.info("No template columns found.")
        else:
            st.markdown("Pick a source for each column. Leave **Keep current rule** to change nothing.")
            head = st.columns([0.7, 2.2, 2.2, 2.2, 3.2, 2.0])
            for h, t in zip(head, ["Col", "Template header", "Current rule", "Source", "Conso Data header(s)", "Preview (1st employee)"]):
                h.markdown(f"**{t}**")

            pending: dict[int, str] = {}
            for entry in me_columns:
                c = entry["col"]
                key = f"me_{me_state}_{me_form}_{c}"
                current = get_rule_for_column(df_mapping_rules, me_rule_row, c)
                cols = st.columns([0.7, 2.2, 2.2, 2.2, 3.2, 2.0])
                cols[0].write(entry["letter"])
                cols[1].write(entry["label"])
                cols[2].caption(f"{current or '-'}" + (f"  \n= {describe_rule(current, df_col_ref)}" if _is_col_prefixed(current) else ""))
                kind = cols[3].selectbox("Source", RULE_KINDS, key=f"{key}_kind", label_visibility="collapsed")
                new_rule = None
                with cols[4]:
                    if kind == RULE_KIND_CONSO:
                        tag = st.selectbox("Conso header", conso_tags, format_func=lambda t: conso_options[t],
                                           key=f"{key}_tag", label_visibility="collapsed")
                        new_rule = build_rule_string(kind, tags=[tag])
                    elif kind == RULE_KIND_FORMULA:
                        n_terms = st.number_input("Terms", min_value=2, max_value=8, value=2, step=1, key=f"{key}_n")
                        f_tags, f_ops = [], []
                        for t_i in range(int(n_terms)):
                            tcols = st.columns([1, 4])
                            if t_i > 0:
                                f_ops.append(tcols[0].selectbox("op", list(FORMULA_OPERATORS), key=f"{key}_op{t_i}", label_visibility="collapsed"))
                            else:
                                tcols[0].write("")
                            f_tags.append(tcols[1].selectbox("Conso header", conso_tags, format_func=lambda t: conso_options[t],
                                                             key=f"{key}_t{t_i}", label_visibility="collapsed"))
                        new_rule = build_rule_string(kind, tags=f_tags, operators=f_ops)
                    elif kind == RULE_KIND_FIXED:
                        new_rule = build_rule_string(kind, text=st.text_input("Text", key=f"{key}_txt", label_visibility="collapsed"))
                    elif kind == RULE_KIND_UNIT:
                        picked = st.multiselect("Units_Master column(s)", unit_letters_all, format_func=lambda l: unit_letter_labels[l],
                                                key=f"{key}_unit", label_visibility="collapsed")
                        new_rule = build_rule_string(kind, unit_letters=picked)
                    elif kind in (RULE_KIND_AUTO, RULE_KIND_WAGE_MONTH):
                        new_rule = build_rule_string(kind)
                    elif kind == RULE_KIND_BLANK:
                        new_rule = ""
                if kind != RULE_KIND_KEEP and new_rule is not None and (new_rule != "" or kind == RULE_KIND_BLANK):
                    pending[c] = new_rule
                    try:
                        preview = resolve_col_rule(new_rule, df_col_ref, sample_row, 0, active_month="April", selected_year=int(selected_year),
                                                   df_units=df_units, selected_unit=sample_row.get("Unit"))
                    except Exception as exc:
                        preview = f"error: {exc}"
                    cols[5].write(str(preview) if preview not in (None, "") else "(blank)")
                    cols[4].caption(f"Will write: `{new_rule or '(empty)'}`")
                elif kind == RULE_KIND_KEEP and not current:
                    hint = suggest_conso_tag(entry["label"], conso_options)
                    if hint:
                        cols[5].caption(f"💡 {conso_options[hint]}")
                elif current:
                    try:
                        cur_val = resolve_col_rule(current, df_col_ref, sample_row, 0, active_month="April", selected_year=int(selected_year),
                                                   df_units=df_units, selected_unit=sample_row.get("Unit"))
                    except Exception as exc:
                        cur_val = f"error: {exc}"
                    cols[5].caption(str(cur_val) if cur_val not in (None, "") else "(blank)")

            st.divider()
            st.write(f"**{len(pending)}** column(s) changed for {me_state} / {me_form}.")
            save_col, dl_col = st.columns(2)
            with save_col:
                if st.button("💾 Save mapping to master workbook", disabled=not pending, key="me_save"):
                    try:
                        updated_bytes = save_rule_edits(MASTER_FILE_PATH, me_rule_row, pending)
                        st.session_state["me_last_saved_master"] = updated_bytes
                        mail_ok, mail_msg = email_file_to_outlook(updated_bytes, "Master_CONSO_DATA_ALL_UNITS.XLSX")
                        st.cache_data.clear()
                        st.success(f"Saved {len(pending)} rule(s). " + (mail_msg if mail_ok else f"Local copy updated ({mail_msg}) - use the download to refresh SharePoint."))
                    except Exception as exc:
                        st.error(f"Could not save mapping: {exc}")
            with dl_col:
                if st.session_state.get("me_last_saved_master"):
                    st.download_button(
                        "📥 Download updated master workbook",
                        data=st.session_state["me_last_saved_master"],
                        file_name="Master_CONSO_DATA_ALL_UNITS.XLSX",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="me_download_master",
                    )
            st.caption(
                "The new rules are used the next time a form is generated. The master workbook is stored on the "
                "app server, so upload the downloaded copy back to SharePoint if the app is redeployed."
            )
