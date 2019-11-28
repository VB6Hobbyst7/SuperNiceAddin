Attribute VB_Name = "数据处理"
Sub strmd5()
    For Each Rng In Selection
        If Rng <> "" Then
            Rng.value = Md5_String_Calc(Application.WorksheetFunction.Clean(Rng.value))
        End If
    Next
End Sub
