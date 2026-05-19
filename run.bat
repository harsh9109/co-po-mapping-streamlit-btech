@echo off
title CO-PO Attainment System v3.0
color 1F
echo.
echo  ╔═══════════════════════════════════════════════════════╗
echo  ║   CO-PO Attainment ^& Academic Analytics System v3.0  ║
echo  ║   NBA Accreditation Tool                              ║
echo  ╚═══════════════════════════════════════════════════════╝
echo.
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python not found! Install from https://www.python.org
    pause & exit /b 1
)
echo  [1/3] Python found.
echo  [2/3] Installing packages (first run may take a minute)...
pip install streamlit pandas numpy plotly openpyxl matplotlib reportlab --quiet --upgrade
echo  [3/3] Launching...  http://localhost:8501
streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
