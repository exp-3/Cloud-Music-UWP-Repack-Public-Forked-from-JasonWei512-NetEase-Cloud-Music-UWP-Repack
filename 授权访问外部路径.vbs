Option Explicit

Dim WshShell, intChoice, strCommand

' ���� WshShell ����
Set WshShell = CreateObject("WScript.Shell")

' ����ѡ��Ի���
intChoice = MsgBox("��ѡ����Ȩ��ʽ��" & vbCrLf & "��� ���ǡ� һ����ȨĬ������Ŀ¼" & vbCrLf & "��� ���� ��Ȩ�����Զ���Ŀ¼", _
                   vbYesNoCancel + vbQuestion + vbSystemModal, "��Ȩѡ��")

' �����û�ѡ��ִ����Ӧ����
Select Case intChoice
    Case vbYes ' �û�ѡ���ǡ���Ĭ����Ȩ��
        strCommand = "data\0.exe %USERPROFILE%\Music *S-1-15-2-4148197969-2579484590-2200292714-3209550610-568264259-1141328317-1602574124"
    Case vbNo ' �û�ѡ�񡰷񡱣��Զ�����Ȩ��
        strCommand = "data\0.exe _CUSTOM *S-1-15-2-4148197969-2579484590-2200292714-3209550610-568264259-1141328317-1602574124"
    Case vbCancel ' �û�ѡ��ȡ����رնԻ���
        WScript.Quit ' �˳��ű�
End Select

' ����ѡ�������
WshShell.Run strCommand, 0, True

' ����
Set WshShell = Nothing
