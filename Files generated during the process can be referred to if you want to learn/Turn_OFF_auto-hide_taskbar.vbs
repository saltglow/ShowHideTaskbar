' Created by: Shawn Brink
' Created on: May 28th 2017
' Tutorial: https://www.tenforums.com/tutorials/23817-turn-off-auto-hide-taskbar-desktop-mode-windows-10-a.html

Const HKEY_CURRENT_USER = &H80000001
Set objShell = CreateObject("WScript.Shell")

' Add registry entry
objShell.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3\Settings", _
    Chr(&H30) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
    Chr(&HFE) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & _
    Chr(&H02) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
    Chr(&H03) & Chr(&H0) & Chr(&H0) & Chr(&H0) & _
    Chr(&H9B) & Chr(&H00) & Chr(&H00) & Chr(&H00) & _
    Chr(&H64) & Chr(&H00) & Chr(&H00) & Chr(&H00) & _
    Chr(&H00) & Chr(&H00) & Chr(&H00) & Chr(&H0C) & _
    Chr(&H08) & Chr(&H00) & Chr(&H00) & Chr(&H00) & _
    Chr(&H0F) & Chr(&H00) & Chr(&H00) & Chr(&H00) & _
    Chr(&H70) & Chr(&H08) & Chr(&H00) & Chr(&H00) & _
    Chr(&HF0) & Chr(&H00) & Chr(&H00) & Chr(&H00) & _
    Chr(&H00) & Chr(&H10) & Chr(&H00) & Chr(&H00) & _
    Chr(&H00)

' Restart Explorer
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
For Each objProcess In colProcesses
    objProcess.Terminate()
Next

' Wait for a while to ensure the process is terminated
WScript.Sleep 1000

' Restart Explorer
objShell.Run "explorer.exe"

' End the script
WScript.Quit
