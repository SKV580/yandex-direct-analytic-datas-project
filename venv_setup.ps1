# Создание окружения (.venv) и установка зависимостей
# Запускай в PowerShell в корне проекта (где requirements.txt)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Creating venv: .venv"
python -m venv .venv

Write-Host "Activating venv..."
.\\./.venv\\Scripts\\Activate.ps1

Write-Host "Upgrading pip..."
python -m pip install --upgrade pip

Write-Host "Installing requirements..."
pip install -r requirements.txt

Write-Host "✅ Done. You can run: .\\start.ps1"