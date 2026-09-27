$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path '.env')) {
    Copy-Item '.env.example' '.env'
    Write-Host 'Created .env. Add your local credentials there, save it, then run this script again.'
    notepad '.env'
    exit 0
}
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    python -m venv .venv
}
& '.venv\Scripts\python.exe' -m pip install -r requirements.txt
& '.venv\Scripts\python.exe' -m agent.app
