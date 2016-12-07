Attribute VB_Name = "modResource"
Option Explicit

'Function to load picture from "Custom" resource.
'Allows display of non-"standard" picture types like JPG and GIF
'which are supported by the LoadPicture function but not by the LoadResPicture function.
Public Function GetCustomPicture(l_CustomResID As Long) As IPictureDisp

    Dim arrPict()   As Byte
    Dim o_FSO       As New Scripting.FileSystemObject
    Dim o_Stream    As Scripting.TextStream
    Dim s_FileName  As String
    
    Debug.Print "Loading Picture from Custom Resource ID " & l_CustomResID
    arrPict = LoadResData(l_CustomResID, "CUSTOM")
    
    s_FileName = o_FSO.GetSpecialFolder(TemporaryFolder) & "\" & o_FSO.GetTempName
    Debug.Print "Saving Picture to temporary file " & s_FileName
    
    Set o_Stream = o_FSO.CreateTextFile(s_FileName, OverWrite:=True, Unicode:=False)
    
    Dim n As Long
    
    For n = LBound(arrPict) To UBound(arrPict)
        o_Stream.Write Chr$(arrPict(n))
    Next n
    
    o_Stream.Close
    
    Debug.Print "Loading Picture from file " & s_FileName
    Set GetCustomPicture = LoadPicture(s_FileName)
    
    Debug.Print "Deleting file " & s_FileName
    SafeKill s_FileName
    
    Set o_Stream = Nothing
    Set o_FSO = Nothing

End Function

