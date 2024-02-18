Option Explicit

Dim WshShell, intChoice, strCommand

' 创建 WshShell 对象
Set WshShell = CreateObject("WScript.Shell")

' 弹出选择对话框
intChoice = MsgBox("请选择授权方式：" & vbCrLf & "点击 “是” 一键授权默认下载目录" & vbCrLf & "点击 “否” 授权其他自定义目录", _
                   vbYesNoCancel + vbQuestion + vbSystemModal, "授权选择")

' 根据用户选择执行相应操作
Select Case intChoice
    Case vbYes ' 用户选择“是”（默认授权）
        strCommand = "data\0.exe %USERPROFILE%\Music *S-1-15-2-4148197969-2579484590-2200292714-3209550610-568264259-1141328317-1602574124"
    Case vbNo ' 用户选择“否”（自定义授权）
        strCommand = "data\0.exe _CUSTOM *S-1-15-2-4148197969-2579484590-2200292714-3209550610-568264259-1141328317-1602574124"
    Case vbCancel ' 用户选择取消或关闭对话框
        WScript.Quit ' 退出脚本
End Select

' 运行选择的命令
WshShell.Run strCommand, 0, True

' 清理
Set WshShell = Nothing
