@echo off
chcp 65001 >nul
echo ============================================
echo  王者荣耀内战数据分析系统 - 局域网分享
echo ============================================
echo.

:: 获取本机IP地址
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    for /f "tokens=1" %%b in ("%%a") do (
        set LOCAL_IP=%%b
    )
)

echo 本机局域网IP地址: %LOCAL_IP%
echo.
echo 启动后，同一WiFi下的朋友可以访问:
echo   http://%LOCAL_IP%:5000
echo.
echo 正在启动服务...
echo ============================================
echo.

cd /d "%~dp0"
C:\ProgramData\anaconda3\python.exe app.py
pause
