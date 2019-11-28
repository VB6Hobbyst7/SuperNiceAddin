VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updateCheck 
   Caption         =   "超棒插件更新信息"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "updateCheck.frx":0000
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "updateCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call updateFile
    updateCheck.Hide
    Choice = MsgBox("最新插件已经下载到了" & Environ("userprofile") & "\Desktop\SuperNiceAddin.xlam" & "，请关闭所有打开的Excel，并将最新文件移动到插件文件夹进行覆盖。" & Chr(10) & "是否需要打开插件文件夹", vbYesNo + vbInformation, "下载完成")
    If Choice = vbYes Then
        Shell "explorer.exe " & Environ("userprofile") & "\AppData\Roaming\Microsoft\AddIns", vbNormalFocus
    Else
    End If
End Sub

Private Sub CommandButton2_Click()
    updateCheck.Hide
End Sub
