$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $scriptDir "wb_photo_report.py"

$pythonCandidates = @(
    "C:\Users\m.cventarnih.KARI\AppData\Local\Programs\Python\Python313\python.exe"
)

try {
    $pythonCommand = Get-Command python -ErrorAction Stop
    if ($pythonCommand.Source) {
        $pythonCandidates += $pythonCommand.Source
    }
} catch {
}

$pythonPath = $pythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $pythonPath) {
    Write-Error "Python не найден. Установите Python или укажите путь к интерпретатору."
    exit 1
}

& $pythonPath $scriptPath @args
exit $LASTEXITCODE
