Attribute VB_Name = "∏¸–¬ƒ£øÈ"
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub updateFile()
    Dim nUrl As String, localFilename As String, lngRetVal As Long
    nUrl = ThisWorkbook.Sheets("≈‰÷√").Range("≈‰÷√!C4").value
    localFilename = Environ("userprofile") & "\Desktop\SuperNice.xlam"
    lngRetVal = URLDownloadToFile(0, nUrl, localFilename, 0, 0)
    If lngRetVal = 0 Then
        DeleteUrlCacheEntry nUrl
    End If
End Sub
Sub A()
MsgBox Environ("HOMEPATH")
End Sub
