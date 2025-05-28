@echo off
setlocal enabledelayedexpansion

set SCRIPT="C:\PDFCrop\pdfcrop.exe"

set ARGS=

:loop
if "%~1"=="" goto run
set ARGS=!ARGS! "%~1"
shift
goto loop

:run
%SCRIPT% %ARGS%