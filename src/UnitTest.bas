Attribute VB_Name = "modUnitTest"
Option Explicit

Public Function runUnitTests() As Boolean

    If runFTDataFileTest() Then
        Debug.Print "Passed RunFTDataFileTest."
        runUnitTests = True
    Else
        Debug.Print "Failed RunFTDataFileTest."
        Debug.Assert False
        runUnitTests = False
        Exit Function
    End If

End Function

Private Function runFTDataFileTest() As Boolean
'    runFTDataFileTest = True
'    Exit Function
On Error GoTo ErrHan
    Dim df1         As New FTDataFile
    Dim line1       As New FTDataLine
    Dim df2         As New FTDataFile
    Dim line2       As New FTDataLine
    Dim FileName    As String
    Dim n           As Long
    Dim lTime       As Long
    
'    FileName = App.Path & "\" & "test3.xls"
    
    df1.FileFormat = Text
    df1.AppendMode = False
'    df1.FileName = FileName
    df1.DenierTarget = 100
    df1.IntegrationTime = 2
    df1.LineSpeed = 50
    df1.SensorOnline(1) = True
    df1.SensorOnline(3) = True
    df1.SensorOnline(4) = True
    FileName = df1.FileName
    
    If Not df1.OpenFile Then
        Debug.Assert False
        runFTDataFileTest = False
        Exit Function
    End If
    
    df1.WriteHeader
    For n = 1 To 10000
        Set line1 = New FTDataLine
        lTime = n * df1.IntegrationTime
        line1.Seconds = lTime Mod 60
        line1.Minutes = Int(lTime / 60) Mod 60
        line1.Hours = Int(Int(lTime / 60) / 60)
        line1.Avg(1) = 98!
        df1.WriteDataLine line1
    Next n
    df1.WriteFooter
    df1.CloseFile
    'runFTDataFileTest = True
    'Exit Function
    
    df2.ReadFile FileName
    df2.CloseFile
    Debug.Assert df2.FileFormat = Text
    Debug.Assert df2.DenierTarget = 100
    Debug.Assert df2.IntegrationTime = 2
    Debug.Assert df2.LineSpeed = 50
    Debug.Assert df2.SensorOnline(1) = True
    Debug.Assert df2.SensorOnline(2) = False
    Debug.Assert df2.SensorOnline(3) = True
    Debug.Assert df2.SensorOnline(4) = True
    Debug.Assert df2.SensorOnline(5) = False
    Debug.Assert df2.SensorOnline(6) = False
    Debug.Assert df2.NumSensors = 3
    
    For n = 0 To 9999
        Debug.Assert df2.DenierValues(1, n) = 98!
    Next n
    
    df2.EraseFile
    
    runFTDataFileTest = True
    
    Exit Function
ErrHan:
    MsgBox "Error!"
    Debug.Assert False
    Exit Function
    
End Function

