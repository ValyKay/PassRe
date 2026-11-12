@echo off
rem ============================================================================
rem  PassRe launcher
rem    Double-click .............. opens the PassRe window
rem    Drag an .xlsx onto this ... opens the window with that workbook loaded
rem    PassRe.bat <verb> <file> .. command line: unprotect | reprotect | toggle | status
rem    PassRe.bat update ......... self-update from GitHub (-CheckOnly to just check, -Force: no prompt)
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
for %%V in (unprotect reprotect toggle status update) do if /I "%~1"=="%%V" goto :cli

rem Otherwise treat the first argument as a file to open in the GUI. -Gui makes a bad path (e.g. one
rem cut at '&' because Explorer didn't quote it) show up as "Not found" in the window, not vanish.
start "" %PS% -WindowStyle Hidden -File "%APP%" -Gui "%~1"
goto :eof

:cli
rem One line, parsed before PowerShell runs: "update" can replace this .bat meanwhile, and cmd would
rem otherwise resume reading the new file mid-way. "call exit /b" hands PowerShell's exit code back
rem (a bare "goto :eof" or "exit /b" can lose it under "cmd /c").
%PS% -File "%APP%" %* & call exit /b %%errorlevel%%
