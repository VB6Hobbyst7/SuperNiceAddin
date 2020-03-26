Attribute VB_Name = "文档程序"
Sub report()
    ThisWorkbook.Sheets("报表").Copy After:=ActiveWorkbook.Sheets(1)
End Sub

