# 大家资产持仓信用主体舆情日报 - Windows定时任务入口
# 放在与 daily_runner.py 同一目录

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonPath = "C:\Users\mengi\AppData\Local\Programs\Python\Python314\python.exe"
$logFile = Join-Path $scriptDir "daily_log.txt"

Set-Location $scriptDir

$sep = "=================================================="
$ts1 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Output $sep | Out-File $logFile -Append -Encoding utf8
Write-Output "[$ts1] 开始执行..." | Out-File $logFile -Append -Encoding utf8
Write-Output $sep | Out-File $logFile -Append -Encoding utf8

# 执行 daily_runner.py
& $pythonPath (Join-Path $scriptDir "daily_runner.py") 2>&1 | Out-File $logFile -Append -Encoding utf8

$exitCode = $LASTEXITCODE
$ts2 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Output "" | Out-File $logFile -Append -Encoding utf8
Write-Output "[$ts2] 完成，退出码: $exitCode" | Out-File $logFile -Append -Encoding utf8
Write-Output $sep | Out-File $logFile -Append -Encoding utf8

exit $exitCode
