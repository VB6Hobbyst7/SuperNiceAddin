Attribute VB_Name = "更新模块"
Public Sub frechUpdateInfo()
    '获取更新信息
    Dim http As Object
    Dim item1name, item1url As String
    Set http = CreateObject("Microsoft.XMLHTTP")
    
    '按需更新下方地址
    http.Open "GET", "https://api.github.com/repos/mattholy/RA-AutomationTool/releases/latest", False
    
    http.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    http.SEND
    If http.Status = 200 Then
        Set x = CreateObject("ScriptControl"): x.Language = "JScript"
        Set updateInfo = x.eval("eval(" & http.responseText & ")")
        itemversion = CallByName(updateInfo, "tag_name", VbGet)
        updateLog = CallByName(updateInfo, "body", VbGet)
        updateat = CallByName(updateInfo, "published_at", VbGet)
        item1name = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "name", VbGet)
        item1url = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "browser_download_url", VbGet)
        ThisWorkbook.Sheets("更新信息").Range("更新信息!B4").Value = item1url
        ThisWorkbook.Sheets("更新信息").Range("更新信息!B2").Value = itemversion
        ThisWorkbook.Sheets("更新信息").Range("更新信息!B3").Value = updateLog
        ThisWorkbook.Sheets("更新信息").Range("更新信息!B5").Value = Replace(Replace(updateat, "T", " "), "Z", "")
        ThisWorkbook.Sheets("更新信息").Range("更新信息!B6").Value = Now()
    Else
        MsgBox "由于网络问题，无法更新数据。" & vbNewLine & "错误代码：http." & http.Status, vbOKOnly + vbInformation, "更新失败"
    End If
End Sub

Public Sub update()
    '开始更新
    If ThisWorkbook.Sheets("更新信息").Range("更新信息!B1").Value = ThisWorkbook.Sheets("更新信息").Range("更新信息!B2").Value Then
        MsgBox "对不起，暂时没有更新可用。请过段时间再试。", vbInformation, "暂无更新"
    Else
        updateCheck.currversion.Caption = ThisWorkbook.Sheets("更新信息").Range("更新信息!B1").Value
        updateCheck.newversion.Caption = ThisWorkbook.Sheets("更新信息").Range("更新信息!B2").Value
        updateCheck.updatetime.Caption = ThisWorkbook.Sheets("更新信息").Range("更新信息!B5").Value
        updateCheck.updateLog.text = ThisWorkbook.Sheets("更新信息").Range("更新信息!B3").Value
        updateCheck.Show
    End If
End Sub

Sub feedback()
    '反馈
    ThisWorkbook.FollowHyperlink "mailto:ganyuanhao@tinman.cn?subject=超棒插件问题反馈与建议&body=当前版本：" _
        & ThisWorkbook.Sheets("更新信息").Range("更新信息!B1").Value _
        & "%0d%0a%0d%0a[请问有什么建议或者问题呢？]%0d%0a"
End Sub
