VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updateCheck 
   Caption         =   "�������������Ϣ"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "updateCheck.frx":0000
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "updateCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call updateFile
    updateCheck.Hide
    Choice = MsgBox("���²���Ѿ����ص���" & Environ("userprofile") & "\Desktop\SuperNiceAddin.xlam" & "����ر����д򿪵�Excel�����������ļ��ƶ�������ļ��н��и��ǡ�" & Chr(10) & "�Ƿ���Ҫ�򿪲���ļ���", vbYesNo + vbInformation, "�������")
    If Choice = vbYes Then
        Shell "explorer.exe " & Environ("userprofile") & "\AppData\Roaming\Microsoft\AddIns", vbNormalFocus
    Else
    End If
End Sub

Private Sub CommandButton2_Click()
    updateCheck.Hide
End Sub
