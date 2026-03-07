<#
Run helper for PowerShell
Usage (PowerShell):
  Open PowerShell in the project root and run:
    .\run_project.ps1

This script will create a virtual environment at <project_root>\venv if missing,
activate it, install requirements, run migrations, and start the dev server.
#>

$projectRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Write-Host "Project root: $projectRoot"

$venvPath = Join-Path $projectRoot 'venv'
$activateScript = Join-Path $venvPath 'Scripts\Activate.ps1'

function Ensure-Python311 {
  Write-Host "Checking for Python 3.11..."

  # Check py launcher first
  if (Get-Command py -ErrorAction SilentlyContinue) {
    try {
      $out = & py -3.11 --version 2>&1
      if ($LASTEXITCODE -eq 0) { Write-Host "Found py -3.11"; return 'py -3.11' }
    } catch {}
  }

  # Check current python executable
  try {
    $ver = & python --version 2>&1
    if ($ver -match 'Python 3\.11') { Write-Host "Found python (3.11)"; return 'python' }
  } catch {}

  Write-Host "Python 3.11 not found. Attempting per-user install of Python 3.11..."

  $installerUrl = 'https://www.python.org/ftp/python/3.11.6/python-3.11.6-amd64.exe'
  $installerPath = Join-Path $env:TEMP 'python-3.11.6-amd64.exe'

  if (-not (Test-Path $installerPath)) {
    Write-Host "Downloading Python installer to $installerPath"
    try {
      Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
    } catch {
      Write-Error "Failed to download Python installer: $_"
      throw
    }
  } else {
    Write-Host "Installer already downloaded: $installerPath"
  }

  Write-Host "Running Python installer (per-user, silent)..."
  $args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1'
  $proc = Start-Process -FilePath $installerPath -ArgumentList $args -Wait -PassThru
  if ($proc.ExitCode -ne 0) {
    Write-Error "Python installer failed with exit code $($proc.ExitCode)"
    throw "Installer failed"
  }

  # Re-check
  if (Get-Command py -ErrorAction SilentlyContinue) {
    try { & py -3.11 --version > $null; return 'py -3.11' } catch {}
  }
  try { $ver = & python --version 2>&1; if ($ver -match 'Python 3\.11') { return 'python' } } catch {}

  throw "Python 3.11 installation succeeded but could not locate python executable. Please restart your shell or add Python to PATH."
}

$pyCmd = Ensure-Python311

if (-not (Test-Path $activateScript)) {
  Write-Host "Creating virtual environment at $venvPath using $pyCmd..."
  & $pyCmd -m venv $venvPath
}

Write-Host "Activating virtual environment..."
. $activateScript

Write-Host "Upgrading pip and installing requirements (preferring binary wheels)..."
python -m pip install --upgrade pip setuptools wheel
python -m pip install --prefer-binary --no-cache-dir -r (Join-Path $projectRoot 'requirements.txt')

Set-Location $projectRoot
Write-Host "Applying migrations..."
python manage.py migrate

Write-Host "Starting Django development server on 127.0.0.1:8000"
python manage.py runserver 127.0.0.1:8000
