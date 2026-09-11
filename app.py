"""
app.py - Main Entrypoint for Streamlit Community Cloud / Hugging Face Spaces / Docker
WCL GEMOY - 5G NR26 Fast & Accurate Analytics
"""
import os
import runpy

app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "wcl_web_app.py")
runpy.run_path(app_path, run_name="__main__")
