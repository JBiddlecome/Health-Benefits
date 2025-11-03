
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Employee Hours Analyzer", page_icon="📊", layout="wide")

st.title("📊 Employee Hours Analyzer")
st.markdown(
    """
    Upload **Employee List** and **Payroll** Excel files, pick a preset date range (2nd → 1st),
    and get an output table of employees with **Total Hours ≥ 360**.
    """
)

# -----------------------------
# Helper functions
# -----------------------------
def normalize_cols(df):
    # strip whitespace and unify case
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_emp_id(value):
    """Return a consistent string representation for employee identifiers."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, np.integer)):
        return str(int(value))

    if isinstance(value, (float, np.floating)):
        if np.isnan(value):
            return None
        if float(value).is_integer():
            return str(int(value))
        return str(value).strip()

    value_str = str(value).strip()
    if not value_str:
        return None

    # Common Excel artefact: identifiers like "123.0"
    if value_str.replace(".", "", 1).isdigit():
        try:
            as_float = float(value_str)
            if as_float.is_integer():
                return str(int(as_float))
        except ValueError:
            pass

    return value_str


def align_employee_ids(emp_series: pd.Series, payroll_series: pd.Series):
    """Return cleaned copies of employee and payroll ID columns with identical formatting."""

    emp_clean = emp_series.apply(clean_emp_id)
    payroll_clean = payroll_series.apply(clean_emp_id)

    # Find the widest purely numeric identifier across both sources so we can zero-fill consistently
    combined = pd.concat([emp_clean.dropna(), payroll_clean.dropna()], ignore_index=True)
    numeric_vals = [val for val in combined if isinstance(val, str) and val.isdigit()]
    pad_width = max((len(val) for val in numeric_vals), default=None)

    def format_id(value: str):
        if not isinstance(value, str):
            return value
        formatted = value.zfill(pad_width) if pad_width and value.isdigit() else value
        return formatted.upper()

    emp_aligned = emp_clean.apply(format_id).astype("string")
    payroll_aligned = payroll_clean.apply(format_id).astype("string")

    return emp_aligned, payroll_aligned

EXCEL_EPOCH = datetime(1899, 12, 30)


def parse_date(value):
    """Coerce assorted Excel date representations into pandas Timestamps."""
    if pd.isna(value):
        return pd.NaT

    # Already a datetime-like object
    if isinstance(value, (pd.Timestamp, datetime, np.datetime64)):
        return pd.to_datetime(value, errors="coerce")

    # Excel serial numbers come through as floats/ints. Ignore booleans which also
    # inherit from int.
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        if np.isnan(value) or value <= 0:
            return pd.NaT
        excel_like = float(value) < 60000  # covers Excel serials through ~2064
        if excel_like:
            try:
                return pd.Timestamp(EXCEL_EPOCH) + pd.to_timedelta(float(value), unit="D")
            except (OverflowError, ValueError):
                pass
        value = str(int(value)) if float(value).is_integer() else str(value)

    # Strings and other objects fall back to pandas' parser
    value_str = str(value).strip()
    if not value_str:
        return pd.NaT
    return pd.to_datetime(value_str, errors="coerce")

def month_2nd_to_1st(year:int, month:int):
    """
    For a given calendar month (1-12), returns (start, end) where:
      start = YYYY-MM-02 00:00:00 of that month
      end   = YYYY-(month+1)-01 23:59:59 of next month (inclusive)
    We'll treat end as inclusive logic by adding a day and using < next_day in filters.
    """
    start = datetime(year, month, 2)
    # end is the 1st of next month at 23:59:59 inclusive -> we compute exclusive upper bound as next_day after end
    end_inclusive = (start + relativedelta(months=1)).replace(day=1)
    # exclusive bound is end_inclusive + 1 day
    exclusive_upper = end_inclusive + relativedelta(days=1)
    return start, end_inclusive, exclusive_upper

def build_month_options(min_year=2020, max_year=None):
    # Build a list of (label, (y,m)) tuples for recent years to choose from.
    # Default to covering a reasonable recent range.
    today = datetime.today()
    if max_year is None:
        max_year = today.year + 1
    options = []
    for y in range(max_year, min_year-1, -1):
        for m in range(12, 0, -1):
            start, end_inclusive, _ = month_2nd_to_1st(y, m)
            label = f"{start.strftime('%m/%d/%Y')} - {end_inclusive.strftime('%m/%d/%Y')}"
            options.append((label, (y, m)))
    return options

def safe_number(x):
    try:
        if pd.isna(x): 
            return 0.0
        return float(str(x).replace(',', '').strip())
    except Exception:
        return 0.0

def to_excel_download(df, filename="results.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
    output.seek(0)
    st.download_button(
        label="⬇️ Download results as Excel",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# File uploaders
# -----------------------------
col1, col2 = st.columns(2)
with col1:
    emp_file = st.file_uploader("Employee List Example (Excel .xlsx)", type=["xlsx"])
with col2:
    payroll_file = st.file_uploader("Payroll Example (Excel .xlsx)", type=["xlsx"])

with st.expander("📎 Expected columns & notes"):
    st.markdown("""
    **Employee List Example** must include at least:
    - `Employee ID`
    - `First Name`
    - `Last Name`
    - `Start Date`
    - `Rehire Date` *(can be blank)*
    - `Status` *(exclude **Terminated** or **Resigned**)*

    **Payroll Example** must include at least:
    - `#Emp`
    - `First Name`
    - `Last Name`
    - `Reg H (e)`
    - `OT H (e)`
    - `DT H (e)`
    - `Non-Worked Hours (e)`
    """)

# -----------------------------
# Date range preset picker
# -----------------------------
st.subheader("Step 1 — Pick a preset date range (2nd → 1st)")
month_options = build_month_options(min_year=2018)  # adjust as needed
labels = [opt[0] for opt in month_options]
default_idx = 0  # most recent is first

selected_label = st.selectbox("Choose a period", labels, index=default_idx, help="Preset periods run from the 2nd of a month through the 1st of the following month.")
selected_year, selected_month = month_options[labels.index(selected_label)][1]
start_dt, end_inclusive_dt, exclusive_upper_dt = month_2nd_to_1st(selected_year, selected_month)

st.info(f"Using **{start_dt.strftime('%m/%d/%Y')}** through **{end_inclusive_dt.strftime('%m/%d/%Y')}** (inclusive).")

debug_mode = st.checkbox(
    "Show troubleshooting diagnostics",
    value=False,
    help="Enable extra tables to compare employee IDs and hours while debugging empty results.",
)

# -----------------------------
# Process files when both are provided
# -----------------------------
if emp_file is not None and payroll_file is not None:
    try:
        emp_df_raw = pd.read_excel(emp_file, engine="openpyxl")
        payroll_df_raw = pd.read_excel(payroll_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read one of the files: {e}")
        st.stop()

    emp_df = normalize_cols(emp_df_raw.copy())
    payroll_df = normalize_cols(payroll_df_raw.copy())

    # Validate required columns
    emp_required = ["Employee ID", "First Name", "Last Name", "Start Date", "Rehire Date", "Status"]
    missing_emp = [c for c in emp_required if c not in emp_df.columns]
    if missing_emp:
        st.error(f"Employee List is missing required columns: {missing_emp}")
        st.stop()

    payroll_required = ["#Emp", "First Name", "Last Name", "Reg H (e)", "OT H (e)", "DT H (e)", "Non-Worked Hours (e)"]
    missing_pay = [c for c in payroll_required if c not in payroll_df.columns]
    if missing_pay:
        st.error(f"Payroll file is missing required columns: {missing_pay}")
        st.stop()

    # Normalize Employee IDs in both datasets before any filtering so formats match exactly
    emp_df["Employee ID"], payroll_df["#Emp"] = align_employee_ids(
        emp_df["Employee ID"], payroll_df["#Emp"]
    )

    if debug_mode:
        st.markdown("#### 🔍 ID alignment preview")
        st.caption("First few IDs after cleaning to verify formatting and leading zeros.")
        st.dataframe(emp_df[["Employee ID"]].head(10))
        st.dataframe(payroll_df[["#Emp"]].head(10))

    # Parse dates in Employee List
    emp_df["Start Date"] = emp_df["Start Date"].apply(parse_date)
    emp_df["Rehire Date"] = emp_df["Rehire Date"].apply(parse_date)

    # Filter status
    mask_active = ~emp_df["Status"].str.strip().str.lower().isin(["terminated", "resigned"])
    emp_active = emp_df.loc[mask_active].copy()

    # Keep employees whose Start Date OR Rehire Date falls within the selected period (inclusive end)
    # We'll treat end as inclusive by using < exclusive_upper_dt
    in_range = (
        ((emp_active["Start Date"] >= start_dt) & (emp_active["Start Date"] < exclusive_upper_dt)) |
        ((emp_active["Rehire Date"] >= start_dt) & (emp_active["Rehire Date"] < exclusive_upper_dt))
    )
    emp_window = emp_active.loc[in_range, ["Employee ID", "First Name", "Last Name", "Start Date", "Rehire Date"]].copy()

    st.subheader("Step 2 — Employees in window (excluding Terminated/Resigned)")
    st.caption("Employees with Start Date or Rehire Date in the selected window.")
    st.dataframe(emp_window, use_container_width=True)

    # Use Employee IDs to filter Payroll rows
    # Convert to Python set for membership checks, ignoring missing IDs
    emp_ids = set(emp_window["Employee ID"].dropna())

    if debug_mode:
        st.markdown("#### 🔍 Employee ID comparison")
        payroll_ids = set(payroll_df["#Emp"].dropna())
        st.write("Employee IDs in window", sorted(emp_ids)[:15])
        st.write("Sample payroll IDs", sorted(payroll_ids)[:15])
        missing_from_payroll = sorted(emp_ids - payroll_ids)
        st.write(
            "IDs missing from payroll",
            missing_from_payroll[:15],
            len(missing_from_payroll),
        )

    payroll_filtered = payroll_df[payroll_df["#Emp"].isin(emp_ids)].copy()

    if debug_mode:
        st.markdown("#### 🔍 Matched payroll rows before hour filter")
        st.write("Matched payroll rows", payroll_filtered.shape[0])
        hour_cols = ["Reg H (e)", "OT H (e)", "DT H (e)", "Non-Worked Hours (e)"]
        preview = payroll_filtered[["#Emp", *hour_cols]].head(10)
        preview_numeric = preview.assign(
            **{col: preview[col].apply(safe_number) for col in hour_cols}
        )
        st.dataframe(preview_numeric)

    # Compute Total Hours = Reg H (e) + OT H (e) + DT H (e) + Non-Worked Hours (e)
    for col in ["Reg H (e)", "OT H (e)", "DT H (e)", "Non-Worked Hours (e)"]:
        payroll_filtered[col] = payroll_filtered[col].apply(safe_number)

    payroll_filtered["Total Hours"] = (
        payroll_filtered["Reg H (e)"]
        + payroll_filtered["OT H (e)"]
        + payroll_filtered["DT H (e)"]
        + payroll_filtered["Non-Worked Hours (e)"]
    )

    # Remove rows where Total Hours < 360
    payroll_final = payroll_filtered[payroll_filtered["Total Hours"] >= 360].copy()

    if debug_mode:
        st.markdown("#### 🔍 Final results overview")
        st.write("Rows meeting 360-hour threshold", payroll_final.shape[0])

    # Display only the requested columns
    out_cols = ["#Emp", "First Name", "Last Name", "Total Hours"]
    missing_out = [c for c in out_cols if c not in payroll_final.columns]
    if missing_out:
        st.error(f"Missing expected columns in payroll after processing: {missing_out}")
        st.stop()

    st.subheader("Step 3 — Results (Total Hours ≥ 360)")
    st.dataframe(payroll_final[out_cols].sort_values(["Last Name", "First Name"]), use_container_width=True)

    # Download
    to_excel_download(payroll_final[out_cols], filename="employee_total_hours_360_plus.xlsx")

else:
    st.warning("Please upload both files to proceed.")
