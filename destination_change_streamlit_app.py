from __future__ import annotations

import os
import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from destination_change_unified_flow import (
    PriorityRule,
    detect_timeline_vendors,
    fmt_date,
    load_fwk3_from_production,
    normalize_pct,
    normalize_whse,
    parse_user_date,
    process_files,
    saturday_of_current_week,
)

st.set_page_config(page_title="Destination Change", layout="wide")

st.title("Destination Change App")
st.caption("Multi-vendor PSW support with Firm PO Reconciliation Gap logic")

with st.expander("Backend logic summary", expanded=False):
    st.markdown(
        """
- **F Wk3 for optimizer** = main vendor PSW `F` at Target Week only.
- **Other Vendor Supply** = explicit other/sub vendor PSW `F` supply when uploaded and matched as other vendor.
- **Firm PO Reconciliation Gap** = Timeline Firm PO at mapped ETA week minus PSW F used for reconciliation.
- **Total Supply Added to SI** = Main Vendor F Wk3 + Other Vendor Supply + Firm PO Reconciliation Gap.
- **New SI** = Current SI + Total Supply Added to SI.
- **New SI-SS** = Current SI-SS + Total Supply Added to SI.
- **SI After** = New SI + Net Destination Change.
- **Sub vendor DC columns** are suggestion-only columns in Optimized Data, calculated after the main-vendor result.
- **Optional OSQP second-pass** can add separate sheets for main-vendor and sub-vendor what-if optimization.
        """
    )

st.subheader("1. Upload input files")
col1, col2 = st.columns(2)

with col1:
    plan_file = st.file_uploader("PlanDetailTimeline.csv", type=["csv"], key="plan")
    due_files = st.file_uploader(
        "DueDateCalc.xlsx files",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        key="due",
        help=(
            "Upload order should match the vendor order detected from PlanDetailTimeline. "
            "If only one file is uploaded, all vendors use the same transit time. "
            "If fewer files than vendors are uploaded, the last uploaded DueDateCalc is used as fallback."
        ),
    )

with col2:
    psw_files = st.file_uploader(
        "PSW / Production Schedule.csv files",
        type=["csv"],
        accept_multiple_files=True,
        key="psw",
        help="Upload one or more PSW/Production Schedule files. If both vendors are in the same file, upload it once. Vendor matching to Timeline decides main vs other. Second/subsequent files are treated as other vendor sources unless vendor matching indicates otherwise in backend rules.",
    )


# Preview vendor order from PlanDetailTimeline so users can upload matching DueDateCalc files.
if plan_file is not None:
    try:
        with tempfile.TemporaryDirectory() as vendor_tmp:
            preview_plan_path = Path(vendor_tmp) / Path(plan_file.name).name
            preview_plan_path.write_bytes(plan_file.getvalue())
            vendor_preview_df = detect_timeline_vendors(str(preview_plan_path))
        if not vendor_preview_df.empty:
            st.subheader("Detected vendor order for DueDateCalc mapping")
            st.caption(
                "Upload DueDateCalc files in this order. "
                "Example: vendor order #1 uses DueDateCalc file #1; vendor order #2 uses DueDateCalc file #2. "
                "If one DueDateCalc is shared by all vendors, upload only one file."
            )
            st.dataframe(vendor_preview_df, use_container_width=True, hide_index=True)
            if due_files:
                due_count = len(due_files)
                vendor_count = len(vendor_preview_df)
                if due_count == 1 and vendor_count > 1:
                    st.info(f"{vendor_count} vendors detected. One DueDateCalc uploaded, so all vendors will use the same transit time.")
                elif due_count < vendor_count:
                    st.warning(
                        f"{vendor_count} vendors detected but only {due_count} DueDateCalc files uploaded. "
                        "The last uploaded DueDateCalc will be used as fallback for remaining vendors."
                    )
                else:
                    st.success("DueDateCalc file count is enough for the detected vendor order.")
        else:
            st.warning("No vendor column/value was detected from PlanDetailTimeline. The app will use the first DueDateCalc as default transit time.")
    except Exception as exc:
        st.warning(f"Could not detect vendor order from PlanDetailTimeline yet: {exc}")

st.subheader("2. Week setup")
default_current = saturday_of_current_week()
default_target = default_current + timedelta(days=14)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    target_week_input = st.text_input("Target Week", value=fmt_date(default_target), help="Example: 5/23/2026")
with c2:
    current_week_input = st.text_input("Current Week", value=fmt_date(default_current), help="Usually keep default unless needed")
with c3:
    offset_mode = st.selectbox("Offset mode", ["legacy_compatible", "due_date"], index=0)

use_osqp_second_pass = st.checkbox(
    "Add OSQP second-pass sheets optional",
    value=False,
    help="Creates separate what-if sheets for OSQP main-vendor and sub-vendor optimization. The current Optimized Data sheet is kept unchanged.",
)

try:
    target_week_preview = parse_user_date(target_week_input)
except Exception:
    target_week_preview = default_target

st.subheader("3. Optional priority rules")
st.markdown(
    "Leave blank if there are no priority warehouses. "
    "**Value examples:** `50` = 50%, `0.5` = 50%, `1` = 100%, `100` = 100%."
)

priority_rules = {}

if psw_files:
    try:
        with tempfile.TemporaryDirectory() as preview_tmp:
            first_psw_path = Path(preview_tmp) / Path(psw_files[0].name).name
            first_psw_path.write_bytes(psw_files[0].getvalue())
            f_preview, _ = load_fwk3_from_production(str(first_psw_path), target_week_preview)
            whse_options = sorted(
                f_preview["Whse"].dropna().astype(str).unique().tolist(),
                key=lambda x: (len(x), x),
            )
    except Exception as exc:
        whse_options = []
        st.warning(f"Could not preview warehouse list from the first PSW file yet: {exc}")

    default_rows = pd.DataFrame([
        {"Whse": "", "Mode": "SI", "Value": None},
    ])

    priority_table = st.data_editor(
        default_rows,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Whse": st.column_config.SelectboxColumn(
                "Whse",
                options=[""] + whse_options,
                help="Select a priority warehouse.",
            ),
            "Mode": st.column_config.SelectboxColumn(
                "Mode",
                options=["SI", "SS"],
                help="SI = cover toward SI = 0. SS = target SI / SS percentage.",
            ),
            "Value": st.column_config.NumberColumn(
                "Value",
                help="Examples: 50 = 50%, 0.5 = 50%, 1 = 100%, 100 = 100%.",
            ),
        },
        key="priority_rules_editor",
    )

    for _, row in priority_table.iterrows():
        whse = normalize_whse(row.get("Whse", ""))
        mode = str(row.get("Mode", "")).strip().upper()
        value = row.get("Value")
        if not whse or whse.lower() == "nan":
            continue
        if mode not in {"SI", "SS"}:
            continue
        if pd.isna(value):
            continue
        priority_rules[whse] = PriorityRule(whse=whse, mode=mode, value=normalize_pct(value))
else:
    st.info("Upload the first PSW file to preview available warehouses for priority rules.")


def save_uploaded_file(uploaded, folder: str) -> str:
    safe_name = Path(uploaded.name).name
    path = os.path.join(folder, safe_name)
    base, ext = os.path.splitext(path)
    idx = 1
    while os.path.exists(path):
        path = f"{base}_{idx}{ext}"
        idx += 1
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return path

st.divider()

if st.button("Run Full Flow", type="primary"):
    if plan_file is None:
        st.error("Please upload PlanDetailTimeline.csv.")
        st.stop()
    if not psw_files:
        st.error("Please upload at least one PSW / Production Schedule.csv file.")
        st.stop()
    if not due_files:
        st.error("Please upload at least one DueDateCalc.xlsx file.")
        st.stop()
    # Multiple DueDateCalc files are supported; upload order follows the detected vendor order.

    try:
        target_week = parse_user_date(target_week_input)
        current_week = parse_user_date(current_week_input)
    except Exception as e:
        st.error(f"Invalid week input: {e}")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            plan_path = save_uploaded_file(plan_file, tmpdir)
            psw_paths = [save_uploaded_file(f, tmpdir) for f in psw_files]
            due_paths = [save_uploaded_file(f, tmpdir) for f in due_files]

            output_path = os.path.join(tmpdir, f"destination_change_multivendor_{target_week.strftime('%Y%m%d')}.xlsx")

            with st.spinner("Running Destination Change full flow..."):
                final_path = process_files(
                    plan_detail_csv=plan_path,
                    production_schedule_csv=psw_paths[0],
                    due_date_calc_xlsx=due_paths[0],
                    output_path=output_path,
                    target_week=target_week,
                    current_week=current_week,
                    priority_rules=priority_rules,
                    offset_mode=offset_mode,
                    psw_csv_paths=psw_paths,
                    other_due_date_calc_xlsx=due_paths[1] if len(due_paths) > 1 else None,
                    due_date_calc_xlsx_list=due_paths,
                    use_osqp_second_pass=use_osqp_second_pass,
                )

            with open(final_path, "rb") as f:
                data = f.read()

            st.success("Done. Download the output Excel below.")
            st.download_button(
                label="Download Output Excel",
                data=data,
                file_name=os.path.basename(final_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.info(
                "Tip: Review the audit columns/sheets for Timeline Firm PO, PSW F Used for Reconciliation, "
                "Firm PO Reconciliation Gap, Total Supply Added to SI, and optional OSQP sheets if enabled."
            )
        except Exception as e:
            st.exception(e)
