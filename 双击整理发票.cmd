@echo off
setlocal EnableExtensions
chcp 65001 >nul
title 发票整理工具（便携版）

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "PYTHON=%ROOT%\runtime\python\python.exe"
set "PROGRAM=%ROOT%\runtime\invoice_organizer.py"
set "PYTHONHOME=%ROOT%\runtime\python"
set "PYTHONPATH=%ROOT%\runtime\python\Lib\site-packages"
set "PYTHONNOUSERSITE=1"
set "PYTHONIOENCODING=utf-8"

if not exist "%PYTHON%" goto :runtime_missing
if not exist "%PROGRAM%" goto :runtime_missing

echo 正在启动发票整理，处理进度会持续显示在下方...
echo.
"%PYTHON%" -u "%PROGRAM%" "%ROOT%"
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" echo 整理完成，请打开新生成的“整理结果”目录。
if not "%RC%"=="0" echo 部分材料无法唯一配对，请查看新生成结果目录中的“待确认.txt”。
echo.
pause
exit /b %RC%

:runtime_missing
echo 便携运行库不完整，请确认 runtime 文件夹与本 CMD 文件一起复制和解压。
echo.
pause
exit /b 10
