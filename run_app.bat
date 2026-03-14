@echo off
cd /d "%~dp0"
call .venv\Scripts\activate
echo.
echo App corriendo. Abrir en el navegador:
echo http://127.0.0.1:8501
echo.
streamlit run app.py --server.port 8501 --server.headless true
pause
