@echo off
cd /d "%~dp0"
if not exist "%~dp0output" mkdir "%~dp0output"
echo Running CATIA VBScript inspect...
cscript //nologo "%~dp0catia_spec_inspect.vbs"
if exist "%~dp0output\catia_spec_inspect_vbs.json" (
  echo Done. Output created: "%~dp0output\catia_spec_inspect_vbs.json"
  if exist "%~dp0output\catia_spec_inspect_vbs.csv" echo CSV output created: "%~dp0output\catia_spec_inspect_vbs.csv"
) else (
  echo Warning: output file not created. Please ensure CATIA is running and the active document is open.
)
pause
