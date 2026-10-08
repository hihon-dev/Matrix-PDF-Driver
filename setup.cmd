@echo off
setlocal
rem Ghostscript and the driver need the Visual C++ 2015-2019 runtime (x86).
rem Windows 7 usually lacks it, so install the bundled one when missing.
if not exist "%SystemRoot%\System32\vcruntime140.dll" goto runtime
if not exist "%SystemRoot%\System32\msvcp140.dll" goto runtime
if not exist "%SystemRoot%\System32\ucrtbase.dll" goto runtime
goto driver

:runtime
"%~dp0vcredist_x86.exe" /install /quiet /norestart

:driver
if "%~1"=="" (
    start "" /wait "%~dp0MatrixPdfDriver.exe" --setup
) else (
    start "" /wait "%~dp0MatrixPdfDriver.exe" --install %*
)
exit /b %errorlevel%
