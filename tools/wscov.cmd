@echo off
setlocal

set "ROOT=%~dp0.."
set "IN=%ROOT%\samples\sut\Calculator.wsc"
set "OUT_WSC=%ROOT%\out\Calculator.__cov__.wsc"
set "MAP=%ROOT%\out\coverage-map.json"
set "TESTS=%ROOT%\samples\tests\calculator"
set "OUT_DIR=%ROOT%\out"

echo [wscov] instrument...
cscript //nologo "%ROOT%\tools\wscov_instrument.vbs" "%IN%" "%OUT_WSC%" "%MAP%"
if errorlevel 1 exit /b %errorlevel%

echo [wscov] run...
cscript //nologo "%ROOT%\tools\wscov_run.vbs" "%OUT_WSC%" "Calculator" "%TESTS%" "%MAP%" "%OUT_DIR%"
exit /b %errorlevel%
