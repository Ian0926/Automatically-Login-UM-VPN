@echo off
title Pulse VPN 自动登录脚本
echo 正在启动Pulse VPN自动登录...
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到Python环境，请先安装Python 3.7+
    pause
    exit /b 1
)

:: 检查依赖是否安装
python -c "import pyautogui" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖...
    pip install -r requirements.txt
)

:: 运行主脚本
python pulse_vpn_auto_login.py

echo.
echo 操作完成！
pause