@echo off
setlocal EnableExtensions

if "%~3"=="" goto :usage

set "ROOT=%~dp0.."
set "INPUT=%~f1"
set "TESTS=%~f2"
set "OUT_DIR=%~f3"
set "COMPONENT=%~4"

if not exist "%INPUT%" (
  echo [wscov] ERROR: input WSC not found: %INPUT%
  exit /b 2
)

if not exist "%TESTS%" (
  echo [wscov] ERROR: tests directory not found: %TESTS%
  exit /b 2
)

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

set "BASE=%~n1"
set "OUT_WSC=%OUT_DIR%\%BASE%.__cov__.wsc"
set "MAP=%OUT_DIR%\coverage-map.json"
set "INSTR_LOG=%OUT_DIR%\instrument.log"
set "RUN_LOG=%OUT_DIR%\run.log"

echo [wscov] instrument: %INPUT%
cscript //nologo "%ROOT%\tools\wscov_instrument.vbs" "%INPUT%" "%OUT_WSC%" "%MAP%" > "%INSTR_LOG%" 2>&1
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo [wscov] instrument failed. exit=%RC%
  type "%INSTR_LOG%"
  exit /b %RC%
)

echo [wscov] run tests: %TESTS%
if "%COMPONENT%"=="" (
  cscript //nologo "%ROOT%\tools\wscov_run.vbs" "%OUT_WSC%" "%TESTS%" "%MAP%" "%OUT_DIR%" > "%RUN_LOG%" 2>&1
) else (
  cscript //nologo "%ROOT%\tools\wscov_run.vbs" "%OUT_WSC%" "%COMPONENT%" "%TESTS%" "%MAP%" "%OUT_DIR%" > "%RUN_LOG%" 2>&1
)
set "RC=%ERRORLEVEL%"

type "%RUN_LOG%"
echo [wscov] logs:
echo   %INSTR_LOG%
echo   %RUN_LOG%
exit /b %RC%

:usage
echo Usage: tools\wscov_apply_target.cmd ^<input.wsc^> ^<testsDir^> ^<outDir^> [componentId]
echo Example:
echo   tools\wscov_apply_target.cmd samples\sut\Calculator.wsc samples\tests\calculator out\m4\calculator Calculator
exit /b 2
