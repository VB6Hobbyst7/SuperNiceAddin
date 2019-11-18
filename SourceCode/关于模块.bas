Attribute VB_Name = "����ģ��"
Public Sub frechUpdateInfo()
    '��ȡ������Ϣ
    Dim http As Object
    Dim item1name, item1url As String
    Set http = CreateObject("Microsoft.XMLHTTP")
    
    '��������·���ַ
    http.Open "GET", "https://api.github.com/repos/mattholy/SuperNiceAddin/releases/latest", False
    
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
        ThisWorkbook.Sheets("����").Range("����!C4").Value = item1url
        ThisWorkbook.Sheets("����").Range("����!C2").Value = itemversion
        ThisWorkbook.Sheets("����").Range("����!C3").Value = updateLog
        ThisWorkbook.Sheets("����").Range("����!C5").Value = Replace(Replace(updateat, "T", " "), "Z", "")
        ThisWorkbook.Sheets("����").Range("����!C6").Value = Now()
    Else
        MsgBox "�����������⣬�޷��������ݡ�" & vbNewLine & "������룺http." & http.Status, vbOKOnly + vbInformation, "����ʧ��"
    End If
End Sub

Public Sub update()
    '��ʼ����
    If ThisWorkbook.Sheets("����").Range("����!C1").Value = ThisWorkbook.Sheets("����").Range("����!C2").Value Then
        MsgBox "�Բ�����ʱû�и��¿��á������ʱ�����ԡ�", vbInformation, "���޸���"
    Else
        updateCheck.currversion.Caption = ThisWorkbook.Sheets("����").Range("����!C1").Value
        updateCheck.newversion.Caption = ThisWorkbook.Sheets("����").Range("����!C2").Value
        updateCheck.updatetime.Caption = ThisWorkbook.Sheets("����").Range("����!C5").Value
        updateCheck.updateLog.text = ThisWorkbook.Sheets("����").Range("����!C3").Value
        updateCheck.Show
    End If
End Sub

Sub feedback()
    '����
    ThisWorkbook.FollowHyperlink "mailto:ganyuanhao@tinman.cn?subject=����������ⷴ���뽨��&body=��ǰ�汾��" _
        & ThisWorkbook.Sheets("����").Range("����!C1").Value _
        & "%0d%0a%0d%0a[������ʲô������������أ�]%0d%0a"
End Sub

Sub MyDebug()
    'Debugģʽ
    If Environ("username") = "��Ԫ��" Then
        Call groupDevShow
        ThisWorkbook.IsAddin = False
        Exit Sub
    End If
    mypass = Application.InputBox("�㼴���л�������ģʽ���������κ��������ݴ򿪵��������ô��������������������ʲô��", "����ģʽ����", "���ڴ����������")
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
