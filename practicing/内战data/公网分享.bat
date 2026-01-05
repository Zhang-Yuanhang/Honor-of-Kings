@echo off
chcp 65001 >nul
echo ============================================
echo  王者荣耀内战数据分析系统 - 公网分享版
echo ============================================
echo.
echo 此脚本将帮助你把本地服务分享给朋友
echo.
echo 请选择内网穿透工具:
echo   1. cpolar (国内，速度快)
echo   2. ngrok  (国外，需要注册)
echo   3. 仅启动本地服务
echo.
set /p choice="请输入选项 (1/2/3): "

if "%choice%"=="1" (
    echo.
    echo [cpolar 使用说明]
    echo 1. 访问 https://www.cpolar.com/ 注册账号
    echo 2. 下载并安装 cpolar
    echo 3. 运行: cpolar http 5000
    echo 4. 复制生成的公网地址分享给朋友
    echo.
    echo 正在启动本地服务...
    start cmd /k "cd /d %~dp0 && C:\ProgramData\anaconda3\python.exe app.py"
    echo.
    echo 服务已启动，请在另一个终端运行: cpolar http 5000
    pause
) else if "%choice%"=="2" (
    echo.
    echo [ngrok 使用说明]
    echo 1. 访问 https://ngrok.com/ 注册账号
    echo 2. 下载 ngrok 并配置 authtoken
    echo 3. 运行: ngrok http 5000
    echo 4. 复制生成的公网地址分享给朋友
    echo.
    echo 正在启动本地服务...
    start cmd /k "cd /d %~dp0 && C:\ProgramData\anaconda3\python.exe app.py"
    echo.
    echo 服务已启动，请在另一个终端运行: ngrok http 5000
    pause
) else (
    echo.
    echo 正在启动本地服务...
    cd /d %~dp0
    C:\ProgramData\anaconda3\python.exe app.py
    pause
)
