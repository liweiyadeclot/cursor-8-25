@echo off
chcp 65001 >nul
echo ============================================================
echo 工作流执行器
echo ============================================================
echo.

if "%~1"=="" (
    echo 用法: run_workflow.bat ^<工作流文件^>
    echo.
    echo 示例:
    echo   run_workflow.bat workflows\simple_mcp_test.json
    echo   run_workflow.bat workflows\reimburse_example.json
    echo.
    echo 可用工作流:
    if exist workflows (
        dir /b workflows\*.json 2>nul
    )
    exit /b 1
)

python run_workflow.py %*

