from __future__ import annotations

import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from destination_change_unified_flow import (
    PriorityRule,
    fmt_date,
    load_fwk3_from_production,
    normalize_pct,
    normalize_whse,
    parse_user_date,
    process_files,
    saturday_of_current_week,
)

st.set_page_config(
    page_title="Destination Change Multi-Vendor Flow",
    page_icon="📦",
    layout="wide",
)

st.title("Destination Change Multi-Vendor Flow")
st.caption("Streamlit app with PSW upload-order logic: first PSW = main vendor, second/subsequent PSW = other vendor supply.")

with st.expander("Logic summary", expanded=False):
    st.markdown(
        """
        **Input files:**
        1. `PlanDetailTimeline.csv` raw export file. Timeline weeks are treated as ETA.
        2. One or more `PSW / Production Schedule.csv` raw export files. PSW weeks are treated as ETD.
        3. `DueDateCalc.xlsx` transit/offset file.

        **PSW upload order rule:**
        - 1st PSW file = **main vendor source**.
        - 2nd and later PSW files = **other/sub vendor source**.

        **Inventory logic:**
        - Main vendor Target Week quantity becomes `F Wk3` and is used for optimizer allocation.
        - Other vendor quantity is added only to `New SI` and `New SI-SS`; it is not reallocated by optimizer.
        - If only one `DueDateCalc.xlsx` is uploaded, main and other vendors use the same warehouse transit time.
        - If a PSW row has its own transit columns such as Transit Days, Delivery Days, Lead Time, or Transit Weeks, that row-level transit is used.
        """
    )

st.header("1) Upload input files")
col1, col2 = st.columns(2)
with col1:
    plan_file = st.file_uploader("PlanDetailTimeline raw CSV", type=["csv"])
with col2:
    due_file = st.file_uploader("DueDateCalc Excel", type=["xlsx", "xlsm", "xls"])

psw_files = st.file_uploader(
    "PSW / Production Schedule raw CSV files",
    type=["csv"],
    accept_multiple_files=True,
    help="Upload order matters: first file = main vendor; second/subsequent files = other vendor(s).",
)

if psw_files:
    st.info("PSW order: " + " | ".join([f"{i+1}. {f.name}" for i, f in enumerate(psw_files)]))
    if len(psw_files) == 1:
        st.caption("Only one PSW file uploaded. The app will run with main vendor supply only; no other vendor supply file is included.")
else:
    st.info("Upload at least one PSW / Production Schedule CSV. The first uploaded file is treated as the main vendor file.")

st.header("2) Week setup")
def_current = saturday_of_current_week()
def_target = def_current + timedelta(days=14)

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    target_week_input = st.date_input("Target Week / Wk3", value=def_target)
with col2:
    current_week_input = st.date_input("Current Week", value=def_current)
with col3:
    offset_mode = st.selectbox(
        "ETA → ETD offset mode",
        options=["legacy_compatible", "due_date"],
        index=0,
        help="legacy_compatible keeps the current SI-SS_WANEK logic when available. due_date uses ceil(Delivery Days / 7).",
    )

target_week = target_week_input if isinstance(target_week_input, date) else parse_user_date(str(target_week_input))
current_week = current_week_input if isinstance(current_week_input, date) else parse_user_date(str(current_week_input))

st.header("3) Priority rules optional")
st.markdown(
    "Leave blank if there are no priority warehouses. "
    "**Value examples:** `50` = 50%, `0.5` = 50%, `1` = 100%, `100` = 100%."
)

priority_rules = {}

if psw_files:
    try:
        with tempfile.TemporaryDirectory() as preview_tmp:
            first_psw_path = Path(preview_tmp) / psw_files[0].name
            first_psw_path.write_bytes(psw_files[0].getvalue())
            f_preview, _ = load_fwk3_from_production(str(first_psw_path), target_week)
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

st.header("4) Run and download")
output_name = st.text_input(
    "Output file name",
    value=f"destination_change_multivendor_{target_week.strftime('%Y%m%d')}.xlsx",
)

ready = plan_file is not None and due_file is not None and bool(psw_files)
run_clicked = st.button("Run Full Flow", type="primary", disabled=not ready)

if not ready:
    st.info("Please upload PlanDetailTimeline, DueDateCalc, and at least one PSW / Production Schedule CSV before running.")

if run_clicked:
    try:
        with st.spinner("Processing full multi-vendor Destination Change flow..."):
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                plan_path = tmp_path / plan_file.name
                due_path = tmp_path / due_file.name
                output_path = tmp_path / output_name

                plan_path.write_bytes(plan_file.getvalue())
                due_path.write_bytes(due_file.getvalue())

                psw_paths = []
                for i, uploaded in enumerate(psw_files, start=1):
                    safe_name = f"psw_{i}_{uploaded.name}"
                    path = tmp_path / safe_name
                    path.write_bytes(uploaded.getvalue())
                    psw_paths.append(str(path))

                # The backend still requires production_schedule_csv. Use the first PSW file as the main production source.
                final_path = process_files(
                    plan_detail_csv=str(plan_path),
                    production_schedule_csv=psw_paths[0],
                    due_date_calc_xlsx=str(due_path),
                    output_path=str(output_path),
                    target_week=target_week,
                    current_week=current_week,
                    priority_rules=priority_rules,
                    offset_mode=offset_mode,
                    psw_csv_paths=psw_paths,
                )
                final_bytes = Path(final_path).read_bytes()

        st.success("Done. Output Excel is ready.")
        st.download_button(
            label="Download Output Excel",
            data=final_bytes,
            file_name=Path(final_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            f"Target Week: {fmt_date(target_week)} | Current Week: {fmt_date(current_week)} | "
            f"Offset mode: {offset_mode} | PSW files: {len(psw_files)} | Priority rules: {len(priority_rules)}"
        )
    except Exception as exc:
        st.error("The app encountered an error while processing the files.")
        st.exception(exc)
