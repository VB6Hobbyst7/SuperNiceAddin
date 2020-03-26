Attribute VB_Name = "数据处理"
Sub strmd5()
    For Each Rng In Selection
        If Rng <> "" Then
            Rng.value = MD5Str(Application.WorksheetFunction.Clean(Rng.value))
        End If
    Next
End Sub

'数据清洗
Sub iclean()
    Dim i
    With Selection
        For i = .Count To 1 Step -1
            .Cells(i) = Application.Clean(.Cells(i))
            .Cells(i) = Application.Trim(.Cells(i))
        Next i
    End With
End Sub

'转成字符串
Sub totext()
    With Selection
        For i = .Count To 1 Step -1
            .Cells(i).Value2 = "'" + Trim(Str(.Cells(i).Value2))
        Next i
    End With
End Sub


