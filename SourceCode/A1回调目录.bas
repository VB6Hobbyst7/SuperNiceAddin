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

Sub mainCallback(control As IRibbonControl)
    Select Case control.ID
        Case "iupdate"
            Call update
        Case "feedback"
            Call feedback
            
        '����
        Case Else
            MsgBox "�ҵ���" & control.ID
            Call inDevNow
    End Select
End Sub


