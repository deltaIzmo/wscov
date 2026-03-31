@echo off
setlocal EnableExtensions

set "ROOT=%~dp0.."

echo [wscov] regression: calculator
call "%ROOT%\tools\wscov_apply_target.cmd" ^
  "%ROOT%\samples\sut\Calculator.wsc" ^
  "%ROOT%\samples\tests\calculator" ^
  "%ROOT%\out\regression\calculator" ^
  "Calculator"
if errorlevel 1 exit /b %errorlevel%

echo [wscov] regression: branchy
call "%ROOT%\tools\wscov_apply_target.cmd" ^
  "%ROOT%\samples\sut\Branchy.wsc" ^
  "%ROOT%\samples\tests\branchy" ^
  "%ROOT%\out\regression\branchy" ^
  "Branchy"
if errorlevel 1 exit /b %errorlevel%

echo [wscov] regression passed
exit /b 0
