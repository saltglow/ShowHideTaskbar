' 自动获取脚本路径
Dim SCRIPT_PATH
SCRIPT_PATH = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Dim TOKEN_FILE
TOKEN_FILE = SCRIPT_PATH & "\token" ' 无扩展名的token文件

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 检查token文件是否存在
If objFSO.FileExists(TOKEN_FILE) Then
    ' 删除token文件
    On Error Resume Next ' 开始错误处理
    objFSO.DeleteFile TOKEN_FILE
    On Error GoTo 0 ' 重置错误处理
    ' 执行让任务栏自动显示的脚本
    Set objShell = CreateObject("WScript.Shell")
    
    ' 运行 REG ADD 命令
    command = "REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3 /v Settings /t REG_BINARY /d 30000000feffffff03000000030000009b00000064000000000000000c080000000f000070080000f000000001000000 /f"
    objShell.Run command, 0, True

    ' 终止 explorer.exe 进程
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
    For Each objProcess In colProcesses
        objProcess.Terminate()
    Next

    WScript.Sleep 1000
    objShell.Run "explorer.exe"
Else
    ' 如果token文件不存在，则生成token文件
    Dim tokenStream
    Set tokenStream = objFSO.CreateTextFile(TOKEN_FILE, True) ' 创建token文件
    tokenStream.Close ' 直接关闭文件，不写入任何内容

    ' 执行让任务栏锁定的脚本
    Set objShell = CreateObject("WScript.Shell")

    ' 添加注册表项
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

    ' 重新启动资源管理器
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'explorer.exe'")
    For Each objProcess In colProcesses
        objProcess.Terminate()
    Next

    ' 等待一段时间以确保进程终止
    WScript.Sleep 1000

    ' 重新启动资源管理器
    objShell.Run "explorer.exe"
End If

WScript.Quit
