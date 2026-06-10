@echo off
REM 大家资产持仓信用主体舆情日报 — Windows定时任务入口
REM 每个工作日9:00由任务计划程序自动执行

set LOGFILE=E:\claude coding\舆情跟踪\daily_log.txt

echo ================================================== >> "%LOGFILE%"
echo [%date% %time%] 开始执行每日舆情报告生成... >> "%LOGFILE%"
echo ================================================== >> "%LOGFILE%"

cd /d "E:\claude coding\舆情跟踪"

if errorlevel 1 (
    echo [错误] 无法切换到工作目录 >> "%LOGFILE%"
    exit /b 1
)

C:\Users\mengi\AppData\Local\Programs\Python\Python314\python.exe daily_runner.py >> "%LOGFILE%" 2>&1

set EXITCODE=%ERRORLEVEL%
echo. >> "%LOGFILE%"
echo [%date% %time%] 执行完成，退出码: %EXITCODE% >> "%LOGFILE%"
echo ================================================== >> "%LOGFILE%"

exit /b %EXITCODE%
