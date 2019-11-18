Attribute VB_Name = "MyDev"
Sub ExportAllVBC()
    Dim ExportPath As String, ExtendName As String
    Dim vbc As VBComponent
    ExportPath = ThisWorkbook.Path
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbc.Type
        Case vbext_ct_ClassModule, vbext_ct_Document    '组件属性为类模块、EXCEL对象
            ExtendName = ".Cls"    '设置导出文件的扩展名
        Case vbext_ct_MSForm    '组件属性为窗体
            ExtendName = ".frm"
        Case vbext_ct_StdModule    '组件属性为模块时
            ExtendName = ".Bas"
        End Select
        If ExtendName <> "" Then
            vbc.Export ExportPath & "\SourceCode\" & vbc.Name & ExtendName
        End If
    Next
    MsgBox "导出完成"
End Sub
