Attribute VB_Name = "ƒ£øÈ1"
Sub ∞¥≈•1_Click()

GetMyIP

End Sub

Public Sub GetMyIP()
Dim strComputer As String
Dim objWMI As Object
Dim colIP As Object
Dim IP As Object
Dim I As Integer
Dim LANstr As String, IPstr As String, MACstr As String
strComputer = "."
    Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colIP = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    For Each IP In colIP
        If Not IsNull(IP.IPAddress) Then
            For I = LBound(IP.IPAddress) To UBound(IP.IPAddress)
                LANstr = IP.Description(I)  'Õ¯ø®√˚≥∆
                ThisWorkbook.Sheets("≈‰÷√").Range("≈‰÷√!C1" & I + 1).value = IP.IPAddress(I)  'Õ¯ø®IPµÿ÷∑
                ThisWorkbook.Sheets("≈‰÷√").Range("≈‰÷√!C10").value = IP.MacAddress      'Õ¯ø®µƒMACµÿ÷∑
                MsgBox "Õ¯ø®√˚≥∆£∫" & LANstr & vbCrLf & "IPµÿ÷∑£∫" & IPstr & vbCrLf & "MACµÿ÷∑£∫" & MACstr, vbInformation, LANstr
            Next
        End If
    Next
End Sub

