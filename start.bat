@echo off
cd /d "D:\vibecoding\excelhelper\development"
call venv\Scripts\activate
echo.
echo ============================================
echo   Excel Helper - Starting...
echo   Address: http://localhost:8501
echo   Press Ctrl+C to stop the server.
echo ============================================
echo.
timeout /t 3 /nobreak >nul
start "" http://localhost:8501
streamlit run app.py
