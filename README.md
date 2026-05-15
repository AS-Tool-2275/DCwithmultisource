# Destination Change Multi-Vendor Streamlit App

Run locally:

```bash
pip install -r requirements.txt
streamlit run destination_change_streamlit_app.py
```

Upload:
1. PlanDetailTimeline raw CSV
2. DueDateCalc.xlsx
3. One or more PSW / Production Schedule raw CSV files

PSW upload order matters:
- First PSW file = main vendor source
- Second/subsequent PSW files = other/sub vendor source

If only one DueDateCalc file is provided, all vendors use the same warehouse transit time unless the PSW row contains its own transit columns.
