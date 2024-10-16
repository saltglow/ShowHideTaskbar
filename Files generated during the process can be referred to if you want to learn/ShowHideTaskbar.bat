@echo off
REM Declare to use code page UTF-8
chcp 65001 >NUL

SET REGKEY=HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced
SET REGVAL=TaskbarAutoHideInTabletMode

SET DISPLAY_TEXT_ENABLED=Auto-hide taskbar feature enabled
SET DISPLAY_TEXT_DISABLED=Auto-hide taskbar feature disabled
SET TIMEOUT_SECOND=2

REM Query the value in registry
SET TASKBAR_AUTO_HIDE_ENABLED=FOR /F "tokens=2,*" %%A IN ('REG Query %REGKEY% /V %REGVAL% ^|findstr %REGVAL%') DO (
    SET TASKBAR_AUTO_HIDE_ENABLED=%%B
)

REM Checking if the value is set
IF not defined TASKBAR_AUTO_HIDE_ENABLED (
    SET TASKBAR_AUTO_HIDE_ENABLED=0
)

IF %TASKBAR_AUTO_HIDE_ENABLED% EQU 0 (
    REM Enable Auto-hide in tablet mode
    REG ADD %REGKEY% /V %REGVAL% /T REG_DWORD /D 1 /F >NUL
    ECHO %DISPLAY_TEXT_ENABLED%
) ELSE (
    REM Disable Auto-hide in tablet mode
    REG ADD %REGKEY% /V %REGVAL% /T REG_DWORD /D 0 /F >NUL
    ECHO %DISPLAY_TEXT_DISABLED%
)

TIMEOUT %TIMEOUT_SECOND% >NUL

REM To kill and restart explorer to take effects
TASKKILL /f /im explorer.exe
START explorer.exe
