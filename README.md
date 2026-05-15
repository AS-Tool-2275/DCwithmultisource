# Destination Change Multi-Vendor Streamlit App

Run locally:

```bash
pip install -r requirements.txt
streamlit run destination_change_streamlit_app.py
```

Upload:
1. PlanDetailTimeline raw CSV
2. DueDateCalc Excel files
   - First file = main/default vendor transit time
   - Second file = sub/other vendor transit time
   - If only one DueDateCalc is uploaded, both vendor groups use the same warehouse transit time
3. One or more PSW / Production Schedule raw CSV files
   - First PSW file = main vendor source
   - Second/subsequent PSW files = other/sub vendor source

PSW weeks are treated as ETD. PlanDetailTimeline weeks are treated as ETA.

Other/sub vendor supply is used to update New SI and New SI-SS only. It is not used for destination allocation.
