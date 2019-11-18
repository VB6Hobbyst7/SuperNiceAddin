Attribute VB_Name = "Tab下的Group控制"
Public Rib As IRibbonUI
'GroupID对应的vba对象
Public groupDev As Boolean

Sub GetVisible(Control As IRibbonControl, ByRef Visible)
    Select Case Control.ID
        Case "dev"
            Visible = groupDev
        Case Else
    End Select
End Sub

'Group初始化
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
