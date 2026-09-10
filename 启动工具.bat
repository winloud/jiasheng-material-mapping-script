@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 嘉盛物料映射工具

set "PY_CMD=py"

:menu
cls
echo ============================================================
echo                 嘉盛物料映射工具
echo ============================================================
echo.
echo   [1] 导入客户物料清单
echo       输入目录：data\客户清单\待处理
echo.
echo   [2] 合并外部映射汇总表
echo       输入目录：data\映射合并\待合并
echo.
echo   [Q] 退出
echo.
echo ============================================================

choice /C 12Q /N /M "请选择 [1/2/Q]："

if errorlevel 3 goto :eof
if errorlevel 2 goto merge
if errorlevel 1 goto customer

:customer
cls
echo ============================================================
echo  1 - 导入客户物料清单
echo ============================================================
echo.
call :check_python
if errorlevel 1 goto after_task
"%PY_CMD%" "app\customer_import.py"
goto after_task

:merge
cls
echo ============================================================
echo  2 - 合并外部映射汇总表
echo ============================================================
echo.
call :check_python
if errorlevel 1 goto after_task
"%PY_CMD%" "app\mapping_merge.py"
goto after_task

:check_python
"%PY_CMD%" --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python Launcher：py
    echo 请确认 Python 已正确安装。
    exit /b 1
)

echo 检查依赖...
"%PY_CMD%" -c "import openpyxl, xlrd" >nul 2>&1
if errorlevel 1 (
    echo 首次运行，正在安装 openpyxl 和 xlrd...
    "%PY_CMD%" -m pip install openpyxl xlrd
    if errorlevel 1 (
        echo [错误] Python 依赖安装失败。
        exit /b 1
    )
)
echo.
exit /b 0

:after_task
echo.
echo ============================================================
echo  本次任务已结束
echo ============================================================
echo.
echo   M - 回主菜单
echo   Q - 退出工具
echo.
choice /C MQ /N /M "请选择 [M/Q]："

if errorlevel 2 goto :eof
if errorlevel 1 goto menu