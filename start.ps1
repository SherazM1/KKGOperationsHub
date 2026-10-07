$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Virtual environment missing. Create .venv and install requirements.txt first."
}
Push-Location -LiteralPath $projectRoot
try {
    & $pythonPath -m streamlit run (Join-Path $projectRoot "app\main.py") @args
    if ($LASTEXITCODE -ne 0) { throw "Streamlit exited with code $LASTEXITCODE." }
} finally {
    Pop-Location
}
