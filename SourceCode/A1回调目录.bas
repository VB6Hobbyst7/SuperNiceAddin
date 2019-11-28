Attribute VB_Name = "A1回调目录"
'开发提示
Sub inDevNow()
    Response = MsgBox("这个功能正在开发当中，将于今后解锁。" & vbNewLine & "需要现在检查更新吗？", 68, "功能未上线")
    If Response = vbYes Then    ' 用户按下“是”
        Call update
    Else    ' 用户按下“否”
    End If
End Sub

'主要回调函数
'负责处理每个功能的调用

Sub mainCallback(Control As IRibbonControl)
    Select Case Control.ID
        Case "iupdate"
            Call update
        Case "feedback"
            Call feedback
        Case "goDev"
            Call MyDebug
        Case "feedback1" '导出所有代码
            Call ExportAllVBC
        Case "functionhelp"
            Call towiki
        Case "strmd5"
            Call strmd5
        '其它
        Case Else
            MsgBox "我点了" & Control.ID
            Call inDevNow
    End Select
End Sub


