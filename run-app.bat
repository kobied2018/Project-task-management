@echo off
cd /d "%~dp0"
if not exist "node_modules" (
  echo Dependencies are not installed. Run npm install first.
  pause
  exit /b 1
)
start "Fieldnotes development server" /D "%~dp0" cmd /k "npm run dev"
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "$deadline = (Get-Date).AddSeconds(60); do { try { $response = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:5173' -TimeoutSec 5 -ErrorAction Stop; $session = Invoke-RestMethod -Uri 'http://127.0.0.1:5173/api/session' -TimeoutSec 5 -ErrorAction Stop; if ($response.StatusCode -eq 200 -and $session.demo -eq $true) { exit 0 }; $lastError = 'The server did not return a local development session.' } catch { $lastError = $_.Exception.Message }; Start-Sleep -Milliseconds 500 } while ((Get-Date) -lt $deadline); Write-Error ('Fieldnotes startup check failed: ' + $lastError); exit 1"
if errorlevel 1 (
  echo The web app or its API did not become ready at http://localhost:5173.
  echo Check the development server window for errors, including ports already in use.
  pause
  exit /b 1
)
start "" "http://localhost:5173"
