Attribute VB_Name = "模块2"
Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
'

'
    Sheets("报表").Select
    Sheets("报表").Copy After:=Workbooks("工作簿1").Sheets(1)
    Windows("SuperNice.xlam").Activate
End Sub
