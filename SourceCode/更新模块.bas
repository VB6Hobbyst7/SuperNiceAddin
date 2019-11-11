Attribute VB_Name = "����ģ��"
Public Sub frechUpdateInfo()
    '��ȡ������Ϣ
    Dim http As Object
    Dim item1name, item1url As String
    Set http = CreateObject("Microsoft.XMLHTTP")
    
    '��������·���ַ
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
        ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B4").Value = item1url
        ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B2").Value = itemversion
        ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B3").Value = updateLog
        ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B5").Value = Replace(Replace(updateat, "T", " "), "Z", "")
        ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B6").Value = Now()
    Else
        MsgBox "�����������⣬�޷��������ݡ�" & vbNewLine & "������룺http." & http.Status, vbOKOnly + vbInformation, "����ʧ��"
    End If
End Sub

Public Sub update()
    '��ʼ����
    If ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B1").Value = ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B2").Value Then
        MsgBox "�Բ�����ʱû�и��¿��á������ʱ�����ԡ�", vbInformation, "���޸���"
    Else
        updateCheck.currversion.Caption = ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B1").Value
        updateCheck.newversion.Caption = ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B2").Value
        updateCheck.updatetime.Caption = ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B5").Value
        updateCheck.updateLog.text = ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B3").Value
        updateCheck.Show
    End If
End Sub

Sub feedback()
    '����
    ThisWorkbook.FollowHyperlink "mailto:ganyuanhao@tinman.cn?subject=����������ⷴ���뽨��&body=��ǰ�汾��" _
        & ThisWorkbook.Sheets("������Ϣ").Range("������Ϣ!B1").Value _
        & "%0d%0a%0d%0a[������ʲô������������أ�]%0d%0a"
End Sub
