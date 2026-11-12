@echo off
rem ============================================================================
rem  PassRe launcher
rem    Double-click .............. opens the PassRe window
rem    Drag an .xlsx onto this ... opens the window with that workbook loaded
rem    PassRe.bat <verb> <file> .. command line: unprotect | reprotect | toggle | status
rem ============================================================================
setlocal
set "APP=%~dp0PassRe.ps1"
set "PS=powershell -STA -NoProfile -ExecutionPolicy Bypass"

rem No arguments -> open the GUI (detached, no console window).
if "%~1"=="" (
    start "" %PS% -WindowStyle Hidden -File "%APP%"
    goto :eof
)

rem A verb -> run on the console so output is visible.
for %%V in (unprotect reprotect toggle status) do if /I "%~1"=="%%V" (
    %PS% -File "%APP%" %*
    goto :eof
)

rem Otherwise treat the first argument as a file to open in the GUI.
start "" %PS% -WindowStyle Hidden -File "%APP%" "%~1"
goto :eof
