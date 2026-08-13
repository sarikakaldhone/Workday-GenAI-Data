"""
PeopleSoft -> Workday HR Data Migration  |  Streamlit Web App
=============================================================
Run via bat file:  run_workday_mapping_agent.bat
Run directly:      streamlit run workday_mapping_agent.py

This file is a launcher shim — it adds the ps_to_workday_migration
project to sys.path and then delegates entirely to app.py there.
"""

import sys
from pathlib import Path

# ── Point to the project that contains app.py and all modules ─────────────
PROJECT_DIR = Path(
    r"C:\Users\tejas.raj.srivastava\Desktop\ps_to_workday_migration\ps_to_workday"
)

if not PROJECT_DIR.exists():
    import streamlit as st
    st.error(
        f"Project directory not found:\n\n`{PROJECT_DIR}`\n\n"
        "Please update the `PROJECT_DIR` path in this file to match your machine."
    )
    st.stop()

if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

# ── Run app.py in-place ────────────────────────────────────────────────────
app_py = PROJECT_DIR / "app.py"
exec(compile(app_py.read_text(encoding="utf-8"), str(app_py), "exec"), {"__file__": str(app_py)})
