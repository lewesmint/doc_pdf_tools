$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$targetPy = @(
    'extract_structure.py'
)

$lintTargets = @(
    'extract_structure.py'
    'scripts/generate_sample_documents.py'
    'scripts/regenerate_from_json.py'
    'scripts/run_iterative_roundtrip.py'
)

$typeTargets = @(
    'extract_structure.py'
    'scripts/generate_sample_documents.py'
    'scripts/regenerate_from_json.py'
    'scripts/run_iterative_roundtrip.py'
)

foreach ($f in $targetPy) {
    if (-not (Test-Path $f)) {
        Write-Host "Missing target file: $f"
        exit 1
    }
}

# Python resolution order:
# 1) explicit override via PYTHON_BIN
# 2) active virtualenv interpreter (if VIRTUAL_ENV is set)
# 3) workspace-local .venv
# 4) current shell python
$pythonBin = $env:PYTHON_BIN
if (-not $pythonBin -and $env:VIRTUAL_ENV) {
    $venvPython = Join-Path $env:VIRTUAL_ENV 'Scripts/python.exe'
    if (Test-Path $venvPython) {
        $pythonBin = $venvPython
    }
}
if (-not $pythonBin) {
    $localVenvPython = Join-Path $scriptDir '.venv/Scripts/python.exe'
    if (Test-Path $localVenvPython) {
        $pythonBin = $localVenvPython
    }
}
if (-not $pythonBin) {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $pythonBin = $pythonCmd.Source
    }
}
if (-not $pythonBin) {
    Write-Host 'Could not locate Python interpreter.'
    exit 1
}

Write-Host "Using Python: $pythonBin"
& $pythonBin --version
Write-Host ''

Write-Host '===== INSTALL_LINT_TOOLS ====='
& $pythonBin -m pip install --upgrade pip
& $pythonBin -m pip install -r requirements.txt
& $pythonBin -m pip install pytest ruff mypy bandit pylint pyright
Write-Host '[exit:0]'
Write-Host ''

$invokeCheck = {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Exe,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    Write-Host "===== $Name ====="
    try {
        & $Exe @Arguments
        if ($LASTEXITCODE -ne $null) {
            $exitCode = $LASTEXITCODE
        }
        else {
            $exitCode = 0
        }
    }
    catch {
        $exitCode = 1
    }

    Write-Host "[exit:$exitCode]"
    Write-Host ''
}

$pyCompileArgs = @('-m', 'py_compile') + $targetPy
$ruffFixArgs = @('-m', 'ruff', 'check', '--fix') + $lintTargets
$ruffArgs = @('-m', 'ruff', 'check') + $lintTargets
$mypyArgs = @('-m', 'mypy')
$pytestArgs = @('-m', 'pytest', '-q')
$banditArgs = @('-m', 'bandit') + $lintTargets
$pylintArgs = @('-m', 'pylint') + $lintTargets
$pyrightArgs = @('-m', 'pyright', '--pythonpath', $pythonBin) + $typeTargets

& $invokeCheck -Name 'PY_COMPILE' -Exe $pythonBin -Arguments $pyCompileArgs
& $invokeCheck -Name 'RUFF_FIX' -Exe $pythonBin -Arguments $ruffFixArgs
& $invokeCheck -Name 'RUFF' -Exe $pythonBin -Arguments $ruffArgs
& $invokeCheck -Name 'MYPY' -Exe $pythonBin -Arguments $mypyArgs
& $invokeCheck -Name 'PYTEST' -Exe $pythonBin -Arguments $pytestArgs
& $invokeCheck -Name 'BANDIT' -Exe $pythonBin -Arguments $banditArgs
& $invokeCheck -Name 'PYLINT' -Exe $pythonBin -Arguments $pylintArgs
& $invokeCheck -Name 'PYRIGHT' -Exe $pythonBin -Arguments $pyrightArgs
