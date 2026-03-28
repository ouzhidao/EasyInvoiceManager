@echo off
chcp 65001 >nul
title 发票智能整理工具 V7.0 - 纯API模式
echo ==========================================
echo   发票智能整理工具 V7.0 - 纯API模式
echo ==========================================
echo.
echo 此模式将：
echo - 只使用百度API进行识别（不加载本地PaddleOCR）
echo - 需要提前配置百度OCR API密钥
echo.
echo 如果尚未配置API，请先在软件中点击"配置API"按钮
echo.
pause

set FORCE_API_ONLY=1
python main.py

if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出
    pause
)
