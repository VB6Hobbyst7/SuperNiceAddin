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
        Set X = CreateObject("ScriptControl"): X.Language = "JScript"
        Set updateInfo = X.eval("eval(" & http.responseText & ")")
        itemversion = CallByName(updateInfo, "tag_name", VbGet)
        updateLog = CallByName(updateInfo, "body", VbGet)
        updateat = CallByName(updateInfo, "published_at", VbGet)
        item1name = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "name", VbGet)
        item1url = CallByName(CallByName(CallByName(updateInfo, "assets", VbGet), "0", VbGet), "browser_download_url", VbGet)
        ThisWorkbook.Sheets("����").Range("����!C4").value = item1url
        ThisWorkbook.Sheets("����").Range("����!C2").value = itemversion
        ThisWorkbook.Sheets("����").Range("����!C3").value = updateLog
        ThisWorkbook.Sheets("����").Range("����!C5").value = Replace(Replace(updateat, "T", " "), "Z", "")
        ThisWorkbook.Sheets("����").Range("����!C6").value = Now()
    Else
        MsgBox "�����������⣬�޷��������ݡ�" & vbNewLine & "������룺http." & http.Status, vbOKOnly + vbInformation, "����ʧ��"
    End If
End Sub

Public Sub update()
    '��ʼ����
    If ThisWorkbook.Sheets("����").Range("����!C1").value = ThisWorkbook.Sheets("����").Range("����!C2").value Then
        MsgBox "�Բ�����ʱû�и��¿��á������ʱ�����ԡ�", vbInformation, "���޸���"
    Else
        updateCheck.currversion.Caption = ThisWorkbook.Sheets("����").Range("����!C1").value
        updateCheck.newversion.Caption = ThisWorkbook.Sheets("����").Range("����!C2").value
        updateCheck.updatetime.Caption = ThisWorkbook.Sheets("����").Range("����!C5").value
        updateCheck.updateLog.text = ThisWorkbook.Sheets("����").Range("����!C3").value
        updateCheck.Show
    End If
End Sub

Sub feedback()
    '����
    ThisWorkbook.FollowHyperlink "mailto:ganyuanhao@tinman.cn?subject=����������ⷴ���뽨��&body=��ǰ�汾��" _
        & ThisWorkbook.Sheets("����").Range("����!C1").value _
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
