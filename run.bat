@echo off
cd /d "%~dp0"
echo Starting ERP vs Tally Reconciliation...
echo Open http://127.0.0.1:5000 in your browser.
start "" http://127.0.0.1:5000
python app.py
pause
