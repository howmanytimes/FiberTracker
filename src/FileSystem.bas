Attribute VB_Name = "modFileSystem"
Option Explicit

'If file doesn't exist, won't throw error; will return false and display msgbox if problem
Public Function SafeKill(s_FileName As String) As Boolean
    On Error Resume Next
    Kill s_FileName
    Select Case Err.Number
        Case 53     'File Not Found
            'Ignore
            Err.Clear
            SafeKill = True
        Case 0
            'Ignore
            SafeKill = True
        Case Else
            Debug.Assert False      'Display msgbox here :-)
            SafeKill = False
    End Select
End Function
