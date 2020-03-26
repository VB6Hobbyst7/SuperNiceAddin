Attribute VB_Name = "处理函数"
Public Function SplitText(Source As String, Separator As String, Choice As Integer)
'处理切割字符串
    SplitTextPool = split(Source, Separator)
    SplitText = SplitTextPool(Choice - 1)
End Function

Public Function SplitPath(Fullpath As String, ResultFlag As Integer) As String
'处理文件路径分割
'ResultFlag=0  获取路径
'ResultFlag=1  获取文件名
'ResultFlag=2  获取扩展名
    Dim SplitPos As Integer, DotPos As Integer
    SplitPos = InStrRev(Fullpath, "\")
    DotPos = InStrRev(Fullpath, ".")
    Select Case ResultFlag
        Case 0
            SplitPath = Left(Fullpath, SplitPos - 1)
        Case 1
            If DotPos = 0 Then DotPos = Len(Fullpath) + 1
            SplitPath = Mid(Fullpath, SplitPos + 1, DotPos - SplitPos - 1)
        Case 2
            If DotPos = 0 Then DotPos = Len(Fullpath)
            SplitPath = Mid(Fullpath, DotPos + 1)
        Case Else
            Err.Raise vbObjectError + 1, "SplitPath Function", "Invalid Parameter!"
    End Select
End Function

'sha-256加密
Public Function SHA256Str(sMessage As String)

    Dim clsX As CSHA256
    Set clsX = New CSHA256

    SHA256Str = clsX.SHA256(sMessage)

    Set clsX = Nothing

End Function
