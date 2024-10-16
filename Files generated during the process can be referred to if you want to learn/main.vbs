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
Else
    ' 如果token文件不存在，则生成token文件
    Dim tokenStream
    Set tokenStream = objFSO.CreateTextFile(TOKEN_FILE, True) ' 创建token文件
    tokenStream.Close ' 直接关闭文件，不写入任何内容
End If

WScript.Quit
