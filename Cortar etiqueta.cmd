@echo off
setlocal enabledelayedexpansion

set SCRIPT="CAMINHO PARA O SCRIPT PYTHON"

set ARGS=

:loop
if "%~1"=="" goto run
set ARGS=!ARGS! "%~1"
shift
goto loop

:run
%SCRIPT% %ARGS%