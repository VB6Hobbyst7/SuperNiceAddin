Attribute VB_Name = "A1�ص�Ŀ¼"
'������ʾ
Sub inDevNow()
    Response = MsgBox("����������ڿ������У����ڽ�������" & vbNewLine & "��Ҫ���ڼ�������", 68, "����δ����")
    If Response = vbYes Then    ' �û����¡��ǡ�
        Call update
    Else    ' �û����¡���
    End If
End Sub

'��Ҫ�ص�����
'������ÿ�����ܵĵ���

Sub mainCallback(Control As IRibbonControl)
    Select Case Control.ID
        Case "iupdate"
            Call update
        Case "feedback"
            Call feedback
        Case "goDev"
            Call MyDebug
        Case "feedback1" '�������д���
            Call ExportAllVBC
        Case "functionhelp"
            Call towiki
        Case "strmd5"
            Call strmd5
        '����
        Case Else
            MsgBox "�ҵ���" & Control.ID
            Call inDevNow
    End Select
End Sub


