Attribute VB_Name = "������"
Public Function SplitText(Source As String, Separator As String, Choice As Integer)
'�����и��ַ���
    SplitTextPool = split(Source, Separator)
    SplitText = SplitTextPool(Choice - 1)
End Function

Public Function SplitPath(Fullpath As String, ResultFlag As Integer) As String
'�����ļ�·���ָ�
'ResultFlag=0  ��ȡ·��
'ResultFlag=1  ��ȡ�ļ���
'ResultFlag=2  ��ȡ��չ��
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



