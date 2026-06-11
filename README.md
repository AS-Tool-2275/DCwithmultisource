# Destination Change App

Run locally:

```bash
pip install -r requirements.txt
streamlit run destination_change_streamlit_app.py
```

Files:
- `destination_change_streamlit_app.py`
- `destination_change_unified_flow.py`
- `requirements.txt`

The app supports:
- PSW-based vendor order detection for DueDateCalc mapping
- Priority rules including SI = 0 hard lock
- Zero-SS fallback equalization pass
- Optional OSQP second-pass sheets
