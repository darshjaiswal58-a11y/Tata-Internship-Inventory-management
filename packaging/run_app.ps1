$ErrorActionPreference = "Stop"

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $appDir "python\python.exe"
$app = Join-Path $appDir "app.py"
$url = "http://127.0.0.1:8000"
$logDir = Join-Path $appDir "logs"
$stdoutLog = Join-Path $logDir "server.out.log"
$stderrLog = Join-Path $logDir "server.err.log"

if (!(Test-Path $python)) {
    throw "Bundled Python not found: $python"
}
if (!(Test-Path $app)) {
    throw "App file not found: $app"
}

$running = $false
try {
    Invoke-WebRequest -Uri "$url/api/me" -UseBasicParsing -TimeoutSec 2 | Out-Null
    $running = $true
} catch {
    $running = $false
}

if (!$running) {
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    Start-Process -FilePath $python -ArgumentList "`"$app`"" -WorkingDirectory $appDir -WindowStyle Hidden -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog
    Start-Sleep -Seconds 6
}

Start-Process $url
