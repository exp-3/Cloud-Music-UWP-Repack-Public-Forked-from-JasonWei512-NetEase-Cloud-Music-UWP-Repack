Option Explicit

Dim WshShell, strUserProfile, strMusicPath, strMusicQuickPath, strCloudMusicData, strCloudMusicDownload, strCloudMusicPath, strDesktopPath, strPublicUser, strPublicMusic, strPublicMusicQuick
Set WshShell = CreateObject("WScript.Shell")

' ��ȡ����·��
strUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
strPublicUser = WshShell.ExpandEnvironmentStrings("%Public%")
strDesktopPath = WshShell.SpecialFolders("Desktop")

' ���·���Ƿ���ȷ��ȡ
If InStr(strUserProfile, "%") > 0 Then
    MsgBox "�޷���ȡ�û��ļ���·�����˵���ϵͳ�������������ش����⣡", vbOKOnly, "���ش���"
    WScript.Quit
End If
If InStr(strPublicUser, "%") > 0 Then
    strPublicUser = strUserProfile & "\..\Public"
End If

' ����·��
strMusicPath = strUserProfile & "\Music"
strMusicQuickPath = strMusicPath & "\CloudMusic"
strPublicMusic = strPublicUser & "\Music"
strPublicMusicQuick = strPublicMusic & "\CloudMusic"
strCloudMusicData = strUserProfile & "\AppData\Local\Packages\cloudmusic.uwp_6p888gkwt396e"
strCloudMusicDownload = strCloudMusicData & "\LocalState\download"
strCloudMusicPath = strCloudMusicDownload & "\music"

' ��鸽���Ƿ����
If Not FolderExists("data\") Then
    MsgBox "ȱʧ������ĸ����ļ������Ƚ����߰��������ļ�һ���ѹ������ʹ�á�", vbOKOnly, "����"
    WScript.Quit
End If
If Not FileExists("data\1.bi") Then
    MsgBox "ȱʧ������ĸ����ļ������Ƚ����߰��������ļ�һ���ѹ������ʹ�á�", vbOKOnly, "����"
    WScript.Quit
End If
If Not FileExists("data\2.bi") Then
    MsgBox "ȱʧ������ĸ����ļ������Ƚ����߰��������ļ�һ���ѹ������ʹ�á�", vbOKOnly, "����"
    WScript.Quit
End If

' ����Ƿ��Ѱ�װ������UWP
If Not FolderExists(strCloudMusicData) Then
    MsgBox "�ƺ�δ��װ������UWP���޷����������Ȱ�װappx���������ʹ�ô˲�����", vbOKOnly, "����"
    WScript.Quit
End If

' ��ֲ���ô洢�ļ�
Sub CopySettingsFile()
    If Not FolderExists(strCloudMusicData & "\Settings") Then
        CreateFolder strCloudMusicData & "\Settings"
    End If
    If FileExists(strCloudMusicData & "\Settings\settings.dat") Then
        Dim objFSO2
        Set objFSO2 = CreateObject("Scripting.FileSystemObject")
        objFSO2.DeleteFile(strCloudMusicData & "\Settings\settings.dat")
        Set objFSO2 = Nothing
    End If
    Dim objFSO1
    Set objFSO1 = CreateObject("Scripting.FileSystemObject")
    objFSO1.CopyFile "data\1.bi", strCloudMusicData & "\Settings\settings.dat"
    Set objFSO1 = Nothing
End Sub
CopySettingsFile

' ����Music�ļ��������������
If Not FolderExists(strMusicPath) Then
    CreateFolder strMusicPath
End If
If Not FolderExists(strPublicMusic) Then
    CreateFolder strPublicMusic
End If

' ���������������ļ��������������
If Not FolderExists(strCloudMusicDownload) Then
    CreateFolder strCloudMusicDownload
End If
If Not FolderExists(strCloudMusicPath) Then
    CreateFolder strCloudMusicPath
End If

' ����������
CreateSymbolicLink strMusicQuickPath, strCloudMusicDownload
CreateSymbolicLink strPublicMusicQuick, strCloudMusicDownload

' �������ϴ�����ݷ�ʽ
CreateShortcut strDesktopPath, strCloudMusicDownload, "������UWP ����Ŀ¼"

' ���������ļ�������ͼ
Sub PutThumbnail()
    Dim objFSO3
    Set objFSO3 = CreateObject("Scripting.FileSystemObject")
    If Not FileExists(strCloudMusicDownload & "\Folder.jpg") Then
        objFSO3.CopyFile "data\2.bi", strCloudMusicDownload & "\Folder.jpg"
        ' ��������ͼ����Ϊϵͳ+����
        objFSO3.GetFile(strCloudMusicDownload & "\Folder.jpg").Attributes = 2 + 4
    End If
    ' ����imagetemp�ļ�������Ϊ����
    If FolderExists(strCloudMusicDownload & "\imagetemp") Then
        objFSO3.GetFolder(strCloudMusicDownload & "\imagetemp").Attributes = 2
    End If
    Set objFSO3 = Nothing
End Sub
PutThumbnail

' ����ļ�(��)�Ƿ���ڵĺ���
Function FolderExists(FolderPath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = objFSO.FolderExists(FolderPath)
    Set objFSO = Nothing
End Function
Function FileExists(FilePath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    FileExists = objFSO.FileExists(FilePath)
    Set objFSO = Nothing
End Function

' �����ļ��еĺ���
Sub CreateFolder(FolderPath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder(FolderPath)
    Set objFSO = Nothing
End Sub

' ���������ӵĺ���
Sub CreateSymbolicLink(LinkPath, LinkTarget)
    If FolderExists(LinkTarget) Then ' ���Ŀ���ļ����Ƿ����
        ' ȷ�� LinkPath �����ڣ����� mklink �����ʧ��
        If Not FolderExists(LinkPath) Then
            ' ʹ�ù���ԱȨ������ mklink ����
            WshShell.Run "cmd /c mklink /d """ & LinkPath & """ """ & LinkTarget & """", 0, True
        End If
    Else
        WScript.Echo "����Ŀ���ļ��в�����: " & LinkTarget
    End If
End Sub

' ������ݷ�ʽ�ĺ���
Sub CreateShortcut(DesktopPath, TargetPath, ShortcutName)
    Dim objShortcut
    Set objShortcut = WshShell.CreateShortcut(DesktopPath & "\" & ShortcutName & ".lnk")
    objShortcut.TargetPath = TargetPath
    objShortcut.Save
    Set objShortcut = Nothing
End Sub

' ������ʾ�������
MsgBox "�������", vbOKOnly, "��ʾ"

