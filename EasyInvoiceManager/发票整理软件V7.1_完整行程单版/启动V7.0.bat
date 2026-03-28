@echo off
chcp 65001 >nul
title 发票智能整理工具 V7.0 - 纯API版
echo ==========================================
echo   发票智能整理工具 V7.0 - 纯API版
echo ==========================================
echo.
echo 说明：
echo   本版本只使用百度API进行识别
echo   需要提前配置API密钥才能使用
echo.
echo 首次使用请先申请百度API：
echo   https://ai.baidu.com/tech/ocr
echo   每天500次免费额度
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

echo [启动] 正在启动程序...
python main.py

if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出
    pause
)
