Write-Host "Starting Outlook Add-in Development Server..." -ForegroundColor Green
Write-Host ""

# Check if node_modules exists
if (!(Test-Path "node_modules")) {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    npm install
    Write-Host ""
}

Write-Host "Starting server at https://localhost:3000" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Yellow
Write-Host ""

# Start the development server
npm start
