:: Created by: Shawn Brink
:: Created on: May 28th 2017
:: Tutorial: https://www.tenforums.com/tutorials/23817-turn-off-auto-hide-taskbar-desktop-mode-windows-10-a.html


REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3" /v Settings /t REG_BINARY /d 30000000feffffff03000000030000009b00000064000000000000000c080000000f000070080000f000000001000000 /f

taskkill /f /im explorer.exe
start explorer.exe
exit