Attribute VB_Name = "modUI"
Option Explicit

Private Const IEXPLORE_PATH = "C:\Program Files\Internet Explorer\IExplore.exe"

Public Sub OpenEmail(emailAddress As String)
On Error GoTo ErrHan
    If LCase$(Left$(emailAddress, Len("mailto:"))) <> "mailto:" Then
        emailAddress = "mailto:" & emailAddress
    End If
    
    OpenProtocolHandler emailAddress
    Exit Sub
ErrHan:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly Or vbExclamation, "FiberTrack - Open Email Application"
    Err.Clear
End Sub

Public Sub OpenWebSite(webSiteAddress As String)
On Error GoTo ErrHan
    If LCase$(Left$(webSiteAddress, Len("http://"))) <> "http://" Then
        webSiteAddress = "http://" & webSiteAddress
    End If
    
    OpenProtocolHandler webSiteAddress
    Exit Sub
ErrHan:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly Or vbExclamation, "FiberTrack - Open Web Site"
    Err.Clear
End Sub

Private Sub OpenProtocolHandler(url As String)
On Error GoTo ErrHan
    Dim fileSystemObject As New fileSystemObject
    If fileSystemObject.FileExists(IEXPLORE_PATH) Then
        Shell IEXPLORE_PATH & " " & url, vbNormalFocus
    Else
        MsgBox "Unable to find path for Internet Explorer", vbOKOnly Or vbExclamation, "FiberTrack - Open Internet Protocol Handler"
    End If
    
    Exit Sub
ErrHan:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly Or vbExclamation, "FiberTrack - Open Internet Protocol Handler"
    Err.Clear
End Sub
