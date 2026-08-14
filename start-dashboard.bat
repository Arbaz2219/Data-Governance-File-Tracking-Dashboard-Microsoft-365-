@echo off
REM LDP Security Dashboard - starts backend (port 3001) and frontend (port 5173)
cd /d "%~dp0"
start "Dashboard Backend (3001)" cmd /k node dynamicFetch.js
start "Dashboard Frontend (5173)" cmd /k "cd client && npm run dev"
echo Backend: http://localhost:3001  ^|  Frontend: http://localhost:5173
