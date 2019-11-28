Attribute VB_Name = "关于模块"
Public Sub frechUpdateInfo()
    '获取更新信息
    Dim http As Object
    Dim item1name, item1url As String
    Set http = CreateObject("Microsoft.XMLHTTP")
    
    '按需更新下方地址
    http.Open "GET", "https://api.github.com/repos/mattholy/SuperNiceAddin/releases/latest", False
    
    http.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    http.SEND
    If http.Status = 200 Then
        Set X = CreateObject("ScriptControl"): X.Language = "JScript"
        Set updateInfo = X.eval("eval(" & http.responseText & ")")
        itemversion = CallByName(updateInfo, "tag_name", VbGet)
        updateLog = CallByName(updateInfo, "body", VbGet)
        updateat = CallByName(updateInfo, "published_at", VbGet)
        item1name = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "name", VbGet)
        item1url = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "browser_download_url", VbGet)
        ThisWorkbook.Sheets("配置").Range("配置!C4").value = item1url
        ThisWorkbook.Sheets("配置").Range("配置!C2").value = itemversion
        ThisWorkbook.Sheets("配置").Range("配置!C3").value = updateLog
        ThisWorkbook.Sheets("配置").Range("配置!C5").value = Replace(Replace(updateat, "T", " "), "Z", "")
        ThisWorkbook.Sheets("配置").Range("配置!C6").value = Now()
    Else
        MsgBox "由于网络问题，无法更新数据。" & vbNewLine & "错误代码：http." & http.Status, vbOKOnly + vbInformation, "更新失败"
    End If
End Sub

Public Sub update()
    '开始更新
    If ThisWorkbook.Sheets("配置").Range("配置!C1").value = ThisWorkbook.Sheets("配置").Range("配置!C2").value Then
        MsgBox "对不起，暂时没有更新可用。请过段时间再试。", vbInformation, "暂无更新"
    Else
        updateCheck.currversion.Caption = ThisWorkbook.Sheets("配置").Range("配置!C1").value
        updateCheck.newversion.Caption = ThisWorkbook.Sheets("配置").Range("配置!C2").value
        updateCheck.updatetime.Caption = ThisWorkbook.Sheets("配置").Range("配置!C5").value
        updateCheck.updateLog.text = ThisWorkbook.Sheets("配置").Range("配置!C3").value
        updateCheck.Show
    End If
End Sub

Sub feedback()
    '反馈
    ThisWorkbook.FollowHyperlink "mailto:ganyuanhao@tinman.cn?subject=超棒插件问题反馈与建议&body=当前版本：" _
        & ThisWorkbook.Sheets("配置").Range("配置!C1").value _
        & "%0d%0a%0d%0a[请问有什么建议或者问题呢？]%0d%0a"
End Sub

Sub MyDebug()
    'Debug模式
    If Environ("username") = "甘元浩" Then
        Call groupDevShow
        ThisWorkbook.IsAddin = False
        Exit Sub
    End If
    mypass = Application.InputBox("你即将切换到开发模式，请勿在任何生产数据打开的情况下这么做，除非你很清楚你在做什么。", "开发模式警告", "请在此输入解锁码")
    Select Case mypass
        Case "mattholy"
            ThisWorkbook.IsAddin = False
        Case Else
            MsgBox mypass
        End Select
End Sub
Sub exitDebug()
    ThisWorkbook.IsAddin = True
End Sub
