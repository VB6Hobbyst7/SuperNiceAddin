Attribute VB_Name = "Tab�µ�Group����"
Public Rib As IRibbonUI
'GroupID��Ӧ��vba����
Public groupDev As Boolean

Sub GetVisible(Control As IRibbonControl, ByRef Visible)
    Select Case Control.ID
        Case "dev"
            Visible = groupDev
        Case Else
    End Select
End Sub

'Group��ʼ��
Sub RibbonOnLoad(Ribbon As IRibbonUI)
    groupDev = False
    Set Rib = Ribbon
End Sub

Public Sub groupDevShow()
    groupDev = True
    Rib.Invalidate
End Sub
Public Sub groupDevHide()
    groupDev = False
    Rib.Invalidate
End Sub
