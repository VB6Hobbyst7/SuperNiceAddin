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
    '��ȡ����ϵͳ
    Select Case Left(Application.OperatingSystem, 7)
        Case "Windows"
            '�����û��޸����ƣ�Ĭ�������޷�ʹ�ò��
            If Environ("username") = "JoJo" Then
                If MsgBox("�벻Ҫʹ��Ĭ���û�����¼����������쳣��" & Chr(13) & "�Ƿ�ǰȥ�޸��û�����", vbYesNo + vbCritical, "�û�����ͻ") = vbYes Then
                    Shell "cmd.exe /c  control userpasswords2"
                End If
                ThisWorkbook.Close
            End If
            '��ǰ�汾
            ThisWorkbook.Sheets("����").Range("����!C1").Value = "1.0r"
            '�豸��Ϣ
            ThisWorkbook.Sheets("����").Range("����!C7").Value = Environ("username")
            ThisWorkbook.Sheets("����").Range("����!C8").Value = Application.OperatingSystem
            ThisWorkbook.Sheets("����").Range("����!C9").Value = Environ("computername")
            '������Ϣ
            strComputer = "."
            Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
            Set colIP = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
            For Each IP In colIP
                If Not IsNull(IP.IPAddress) Then
                    For i = LBound(IP.IPAddress) To UBound(IP.IPAddress)
                        LANstr = IP.Description(i)  '��������
                        ThisWorkbook.Sheets("����").Range("����!C1" & i + 1).Value = IP.IPAddress(i)  '����IP��ַ
                        ThisWorkbook.Sheets("����").Range("����!C10").Value = IP.MacAddress      '������MAC��ַ
                    Next
                End If
            Next
            'ÿ����һ�θ���
            If Now() - ThisWorkbook.Sheets("����").Range("����!C6").Value > 1 Then
                Call frechUpdateInfo
            End If
            '����и�����������¶Ի���
            If ThisWorkbook.Sheets("����").Range("����!C1").Value <> ThisWorkbook.Sheets("����").Range("����!C2").Value Then
                Call update
            End If
        Case Else
            MsgBox "��ǰϵͳ����Windowsϵͳ���������������������ģʽ�£�һЩ���ܺͻ���ܻᱻ���ã��Ա�֤Excel�ȶ������ݰ�ȫ���Ժ�������л�������ģʽ�������Щ���ơ�", vbInformation, "����ģʽ"
            ThisWorkbook.Close
        End Select
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call exitDebug
    ThisWorkbook.Close (1)
End Sub