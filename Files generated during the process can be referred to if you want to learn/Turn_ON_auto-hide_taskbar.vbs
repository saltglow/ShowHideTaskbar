Set objShell = CreateObject("WScript.Shell")

' Run REG ADD command
command = "REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3 /v Settings /t REG_BINARY /d 30000000feffffff03000000030000009b00000064000000000000000c080000000f000070080000f000000001000000 /f"
objShell.Run command, 0, True

' Terminate explorer.exe process
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
For Each objProcess In colProcesses
    objProcess.Terminate()
Next

WScript.Sleep 1000
objShell.Run "explorer.exe"
WScript.Quit
