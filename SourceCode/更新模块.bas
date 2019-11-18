Attribute VB_Name = "¸üÐÂÄ£¿é"
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Sub updateFile()
 Dim nUrl As String, localFilename As String, lngRetVal As Long
    nUrl = "https://github.com/mattholy/SuperNiceAddin/releases/download/v0.1a/SuperNice.xlam"
    localFilename = Environ("userprofile") & "\Í¼Æ¬2.XLAM"
    lngRetVal = URLDownloadToFile(0, nUrl, localFilename, 0, 0)
    If lngRetVal = 0 Then
        DeleteUrlCacheEntry nUrl
    End If
End Sub
Sub A()
MsgBox Environ("HOMEPATH")
End Sub
