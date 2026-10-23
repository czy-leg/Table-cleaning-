@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
title 表格清洗助手

set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"

echo.
echo ========================================
echo        表格清洗助手 - 本地启动
echo ========================================
echo.

if exist "%VENV_PYTHON%" goto check_python

echo [1/3] 首次运行，正在创建 Python 虚拟环境...
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m venv ".venv"
) else (
    where python >nul 2>nul
    if errorlevel 1 goto python_missing
    python -m venv ".venv"
)

if errorlevel 1 goto venv_failed
if not exist "%VENV_PYTHON%" goto venv_failed

:check_python
"%VENV_PYTHON%" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 goto python_too_old

echo [2/3] 检查运行依赖...
"%VENV_PYTHON%" -c "import streamlit, pandas, openpyxl, pytest" >nul 2>nul
if not errorlevel 1 goto launch

echo 正在安装依赖，首次安装需要联网，请稍候...
set "PIP_DISABLE_PIP_VERSION_CHECK=1"
"%VENV_PYTHON%" -m pip install -r "requirements.txt"
if errorlevel 1 goto install_failed

:launch
echo [3/3] 正在启动应用，浏览器通常会自动打开...
echo 关闭黑色命令窗口或按 Ctrl+C 即可停止应用。
echo.
start "" "http://127.0.0.1:8501"
"%VENV_PYTHON%" -m streamlit run "app.py"
if errorlevel 1 goto app_failed
exit /b 0

:python_missing
echo.
echo [错误] 未找到 Python。请先从 https://www.python.org/downloads/ 安装 Python 3.10 或更高版本。
echo 安装时请勾选“Add Python to PATH”，完成后重新双击 run.bat。
goto failed

:python_too_old
echo.
echo [错误] 当前 Python 版本低于 3.10。请安装 Python 3.10 或更高版本后重试。
goto failed

:venv_failed
echo.
echo [错误] 虚拟环境创建失败。请确认当前目录可写，并确认 Python 安装完整。
goto failed

:install_failed
echo.
echo [错误] 依赖安装失败。请检查网络连接，或运行以下命令查看详细错误：
echo "%VENV_PYTHON%" -m pip install -r requirements.txt
goto failed

:app_failed
echo.
echo [错误] 应用启动失败。请查看上方错误信息，或按 README 的“常见问题”排查。
goto failed

:failed
echo.
pause
exit /b 1
