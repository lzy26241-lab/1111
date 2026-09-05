@echo off
chcp 65001 >nul
title Doubao Workspace 一键同步

echo ========================================
echo   Doubao Workspace 一键同步到 GitHub
echo ========================================
echo.

cd /d "D:\Doubao_Workspace"

set PATH=D:\Program Files\Git\cmd;%PATH%

echo [1/3] 正在检测文件变化...
git add .
git status --short

echo.
echo [2/3] 正在提交...
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%"
set "Min=%dt:~10,2%"
set "SS=%dt:~12,2%"
set "timestamp=%YYYY%-%MM%-%DD% %HH%:%Min%:%SS%"

git commit -m "自动同步：%timestamp%" 2>nul

if %errorlevel% neq 0 (
    echo 没有新的变化需要提交。
    echo.
    pause
    exit /b 0
)

echo.
echo [3/3] 正在推送到 GitHub...
git push

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo   同步成功！
    echo ========================================
) else (
    echo.
    echo ========================================
    echo   推送失败，请检查网络或认证状态
    echo ========================================
)

echo.
pause
