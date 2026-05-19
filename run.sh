#!/usr/bin/env bash
echo " CO-PO Attainment System v3.0 — NBA Tool"
command -v python3 &>/dev/null || { echo "Python3 not found"; exit 1; }
pip3 install streamlit pandas numpy plotly openpyxl matplotlib reportlab --quiet --upgrade
streamlit run app.py --server.headless false --browser.gatherUsageStats false
