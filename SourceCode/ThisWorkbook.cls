VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    '获取操作系统
    Select Case Left(Application.OperatingSystem, 7)
        Case "Windows"
            '引导用户修改名称；默认名称无法使用插件
            If Environ("username") = "JoJo" Then
                If MsgBox("请不要使用默认用户名登录，它会造成异常。" & Chr(13) & "是否前去修改用户名？", vbYesNo + vbCritical, "用户名冲突") = vbYes Then
                    Shell "cmd.exe /c  control userpasswords2"
                End If
                ThisWorkbook.Close
            End If
            '当前版本
            ThisWorkbook.Sheets("配置").Range("配置!C1").Value = "1.0r"
            '设备信息
            ThisWorkbook.Sheets("配置").Range("配置!C7").Value = Environ("username")
            ThisWorkbook.Sheets("配置").Range("配置!C8").Value = Application.OperatingSystem
            ThisWorkbook.Sheets("配置").Range("配置!C9").Value = Environ("computername")
            '网络信息
            strComputer = "."
            Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
            Set colIP = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
            For Each IP In colIP
                If Not IsNull(IP.IPAddress) Then
                    For i = LBound(IP.IPAddress) To UBound(IP.IPAddress)
                        LANstr = IP.Description(i)  '网卡名称
                        ThisWorkbook.Sheets("配置").Range("配置!C1" & i + 1).Value = IP.IPAddress(i)  '网卡IP地址
                        ThisWorkbook.Sheets("配置").Range("配置!C10").Value = IP.MacAddress      '网卡的MAC地址
                    Next
                End If
            Next
            '每天检查一次更新
            If Now() - ThisWorkbook.Sheets("配置").Range("配置!C6").Value > 1 Then
                Call frechUpdateInfo
            End If
            '如果有更新则拉起更新对话框
            If ThisWorkbook.Sheets("配置").Range("配置!C1").Value <> ThisWorkbook.Sheets("配置").Range("配置!C2").Value Then
                Call update
            End If
        Case Else
            MsgBox "当前系统并非Windows系统，本插件将会运行在限制模式下，一些功能和活动可能会被禁用，以保证Excel稳定和数据安全。稍后你可以切换到开发模式来解除这些限制。", vbInformation, "限制模式"
            ThisWorkbook.Close
        End Select
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call exitDebug
    ThisWorkbook.Close (1)
End Sub
