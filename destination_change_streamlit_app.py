from __future__ import annotations

import os
import tempfile
from datetime import timedelta
from pathlib import Path

import streamlit as st

from destination_change_unified_flow import (
    PriorityRule,
    fmt_date,
    normalize_pct,
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
        help="Upload order: 1st file = main/default vendor transit. 2nd file = sub/other vendor transit. If only 1 file is uploaded, sub vendor uses the main/default transit.",
    )

with col2:
    psw_files = st.file_uploader(
        "PSW / Production Schedule.csv files",
        type=["csv"],
        accept_multiple_files=True,
        key="psw",
        help="Upload one or more PSW/Production Schedule files. If both vendors are in the same file, upload it once. Vendor matching to Timeline decides main vs other. Second/subsequent files are treated as other vendor sources unless vendor matching indicates otherwise in backend rules.",
    )

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

st.subheader("3. Optional priority rules")
st.markdown("Enter one rule per line using `Whse, Mode, Value`. Mode can be `SI` or `SS`. Example: `17, SI, 50` or `28, SS, 0`.")
priority_text = st.text_area("Priority rules", value="", height=100, placeholder="17, SI, 50\n28, SS, 0")


def parse_priority_rules(text: str):
    rules = {}
    errors = []
    for line_no, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) != 3:
            errors.append(f"Line {line_no}: expected format Whse, Mode, Value")
            continue
        whse, mode, value = parts
        mode = mode.upper()
        if mode not in {"SI", "SS"}:
            errors.append(f"Line {line_no}: Mode must be SI or SS")
            continue
        try:
            val = normalize_pct(float(value))
        except Exception:
            errors.append(f"Line {line_no}: Value is not numeric")
            continue
        rules[str(whse).strip()] = PriorityRule(whse=str(whse).strip(), mode=mode, value=val)
    return rules, errors


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
    if len(due_files) > 2:
        st.warning("More than 2 DueDateCalc files were uploaded. The app will use only the first two files.")

    priority_rules, priority_errors = parse_priority_rules(priority_text)
    if priority_errors:
        st.error("Priority rule errors:\n" + "\n".join(priority_errors))
        st.stop()

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
            due_paths = [save_uploaded_file(f, tmpdir) for f in due_files[:2]]

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
                "Firm PO Reconciliation Gap, and Total Supply Added to SI."
            )
        except Exception as e:
            st.exception(e)
