Option Explicit

Dim WshShell, strUserProfile, strMusicPath, strMusicQuickPath, strCloudMusicData, strCloudMusicDownload, strCloudMusicPath, strDesktopPath, strPublicUser, strPublicMusic, strPublicMusicQuick
Set WshShell = CreateObject("WScript.Shell")

' 获取变量路径
strUserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
strPublicUser = WshShell.ExpandEnvironmentStrings("%Public%")
strDesktopPath = WshShell.SpecialFolders("Desktop")

' 检查路径是否被正确获取
If InStr(strUserProfile, "%") > 0 Then
    MsgBox "无法获取用户文件夹路径，此电脑系统环境变量存在重大问题！", vbOKOnly, "严重错误"
    WScript.Quit
End If
If InStr(strPublicUser, "%") > 0 Then
    strPublicUser = strUserProfile & "\..\Public"
End If

' 定义路径
strMusicPath = strUserProfile & "\Music"
strMusicQuickPath = strMusicPath & "\CloudMusic"
strPublicMusic = strPublicUser & "\Music"
strPublicMusicQuick = strPublicMusic & "\CloudMusic"
strCloudMusicData = strUserProfile & "\AppData\Local\Packages\cloudmusic.uwp_6p888gkwt396e"
strCloudMusicDownload = strCloudMusicData & "\LocalState\download"
strCloudMusicPath = strCloudMusicDownload & "\music"

' 检查附件是否存在
If Not FolderExists("data\") Then
    MsgBox "缺失了所需的附属文件，请先将工具包的所有文件一起解压缩后再使用。", vbOKOnly, "错误"
    WScript.Quit
End If
If Not FileExists("data\1.bi") Then
    MsgBox "缺失了所需的附属文件，请先将工具包的所有文件一起解压缩后再使用。", vbOKOnly, "错误"
    WScript.Quit
End If
If Not FileExists("data\2.bi") Then
    MsgBox "缺失了所需的附属文件，请先将工具包的所有文件一起解压缩后再使用。", vbOKOnly, "错误"
    WScript.Quit
End If

' 检查是否已安装云音乐UWP
If Not FolderExists(strCloudMusicData) Then
    MsgBox "似乎未安装云音乐UWP，无法继续。请先安装appx软件包后再使用此补丁。", vbOKOnly, "错误"
    WScript.Quit
End If

' 移植设置存储文件
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

' 创建Music文件夹如果它不存在
If Not FolderExists(strMusicPath) Then
    CreateFolder strMusicPath
End If
If Not FolderExists(strPublicMusic) Then
    CreateFolder strPublicMusic
End If

' 创建云音乐下载文件夹如果它不存在
If Not FolderExists(strCloudMusicDownload) Then
    CreateFolder strCloudMusicDownload
End If
If Not FolderExists(strCloudMusicPath) Then
    CreateFolder strCloudMusicPath
End If

' 创建软链接
CreateSymbolicLink strMusicQuickPath, strCloudMusicDownload
CreateSymbolicLink strPublicMusicQuick, strCloudMusicDownload

' 在桌面上创建快捷方式
CreateShortcut strDesktopPath, strCloudMusicDownload, "云音乐UWP 下载目录"

' 放置下载文件夹缩略图
Sub PutThumbnail()
    Dim objFSO3
    Set objFSO3 = CreateObject("Scripting.FileSystemObject")
    If Not FileExists(strCloudMusicDownload & "\Folder.jpg") Then
        objFSO3.CopyFile "data\2.bi", strCloudMusicDownload & "\Folder.jpg"
        ' 设置缩略图属性为系统+隐藏
        objFSO3.GetFile(strCloudMusicDownload & "\Folder.jpg").Attributes = 2 + 4
    End If
    ' 设置imagetemp文件夹属性为隐藏
    If FolderExists(strCloudMusicDownload & "\imagetemp") Then
        objFSO3.GetFolder(strCloudMusicDownload & "\imagetemp").Attributes = 2
    End If
    Set objFSO3 = Nothing
End Sub
PutThumbnail

' 检查文件(夹)是否存在的函数
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

' 创建文件夹的函数
Sub CreateFolder(FolderPath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder(FolderPath)
    Set objFSO = Nothing
End Sub

' 创建软链接的函数
Sub CreateSymbolicLink(LinkPath, LinkTarget)
    If FolderExists(LinkTarget) Then ' 检查目标文件夹是否存在
        ' 确保 LinkPath 不存在，否则 mklink 命令会失败
        If Not FolderExists(LinkPath) Then
            ' 使用管理员权限运行 mklink 命令
            WshShell.Run "cmd /c mklink /d """ & LinkPath & """ """ & LinkTarget & """", 0, True
        End If
    Else
        WScript.Echo "错误，目标文件夹不存在: " & LinkTarget
    End If
End Sub

' 创建快捷方式的函数
Sub CreateShortcut(DesktopPath, TargetPath, ShortcutName)
    Dim objShortcut
    Set objShortcut = WshShell.CreateShortcut(DesktopPath & "\" & ShortcutName & ".lnk")
    objShortcut.TargetPath = TargetPath
    objShortcut.Save
    Set objShortcut = Nothing
End Sub

' 弹窗提示操作完成
MsgBox "操作完成", vbOKOnly, "提示"

