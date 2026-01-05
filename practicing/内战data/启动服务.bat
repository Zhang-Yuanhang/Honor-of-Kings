@echo off
echo ============================================
echo  王者荣耀内战数据分析系统 - 启动器
echo ============================================
echo.
echo 正在启动Web服务...
echo 启动后请访问: http://127.0.0.1:5000
echo.
cd /d "%~dp0"
C:\ProgramData\anaconda3\python.exe app.py
pause
