# Активация .venv и запуск entrypoint main.py
# Запускай в PowerShell в корне проекта

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (!(Test-Path ".venv\\Scripts\\Activate.ps1")) {
  Write-Host "⚠️ .venv not found. Run: .\\venv_setup.ps1" -ForegroundColor Yellow
  exit 1
}

Write-Host "Activating venv..."
.\\./.venv\\Scripts\\Activate.ps1

Write-Host "Running: python main.py"
python main.py