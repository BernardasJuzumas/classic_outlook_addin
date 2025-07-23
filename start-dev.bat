@echo off
echo Starting Outlook Add-in Development Server...
echo.

REM Check if node_modules exists
if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    echo.
)

echo Starting server at https://localhost:3000
echo Press Ctrl+C to stop the server
echo.

call npm start
