VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTDataLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FTTimeUnitType
    tutError = -1000
    tutUnknown = 0
    tutString = 1
    tutHours = 2
    tutMinutes = 3
    tutSeconds = 4
End Enum

Public Hours As Long
Public Minutes As Long
Public Seconds As Long
Private m_Avg(1 To LAST_SENSOR) As Double
Private m_CvValue(1 To LAST_SENSOR) As Double

Public Property Get Avg(index As Integer) As Double
    Avg = m_Avg(index)
End Property

Public Property Let Avg(index As Integer, value As Double)
    m_Avg(index) = value
End Property

Public Property Get Cv(index As Integer) As Double
    Cv = m_CvValue(index)
End Property

Public Property Let Cv(index As Integer, value As Double)
    m_CvValue(index) = value
End Property

Private Sub Class_Initialize()
    Hours = 0
    Minutes = 0
    Seconds = 0
    Dim indexAverage As Integer
    For indexAverage = LBound(m_Avg) To UBound(m_Avg)
        m_Avg(indexAverage) = 0
    Next indexAverage
    Dim index As Integer
    For index = LBound(m_CvValue) To UBound(m_CvValue)
        m_CvValue(index) = 0
    Next index
End Sub

Public Function FormatTime(Optional Units As FTTimeUnitType = tutString) As String
    Select Case Units
        Case tutHours
            FormatTime = FormatTimeAsHours
        Case tutMinutes
            FormatTime = FormatTimeAsMinutes
        Case tutSeconds
            FormatTime = FormatTimeAsSeconds
        Case tutString
            FormatTime = FormatTimeAsString
        Case Else
            Debug.Assert False
    End Select

End Function

Public Function TimeAsHours() As Double
    TimeAsHours = Hours + (Minutes / 60!) + (Seconds / 3600!)
End Function

Private Function FormatTimeAsHours() As String
    FormatTimeAsHours = Format$(TimeAsHours, "##0.00####")
End Function

Public Function TimeAsMinutes() As Double
    TimeAsMinutes = (CDbl(Hours) * 60!) + CDbl(Minutes) + (CDbl(Seconds) / 60!)
End Function

Private Function FormatTimeAsMinutes() As String
    FormatTimeAsMinutes = Format$(TimeAsMinutes, "####0.0##")
End Function

Public Function TimeAsSeconds() As Long
    TimeAsSeconds = (CLng(Hours) * 3600) + (CLng(Minutes) * 60) + CLng(Seconds)
End Function

Private Function FormatTimeAsSeconds() As String
    FormatTimeAsSeconds = CStr(TimeAsSeconds)
End Function

Private Function FormatTimeAsString() As String
    FormatTimeAsString = StringFormat("{0}:{1}:{2}", Format$(Hours, "00"), Format$(Minutes, "00"), Format$(Seconds, "00"))
End Function

Public Function ParseTime(inputTime As String, inputTimeUnitType As FTTimeUnitType) As FTTimeUnitType
On Error GoTo ErrHan
    Dim timeUnitType As FTTimeUnitType
    inputTime = Trim$(inputTime)
    If inputTimeUnitType <> tutUnknown Then
        timeUnitType = inputTimeUnitType
    Else
        If InStr(inputTime, ":") > 0 Then
            timeUnitType = tutString
        ElseIf InStr(inputTime, ".") > 0 Then
            timeUnitType = tutMinutes
        Else
            timeUnitType = tutSeconds
        End If
    End If
    Dim singleHours As Single
    Dim singleMinutes As Single
    Dim longMinutes As Long
    Dim singleSeconds As Single
    Select Case timeUnitType
        Case tutMinutes
            singleHours = CSng(inputTime)
            singleSeconds = singleMinutes - Int(singleMinutes)
            longMinutes = Int(singleMinutes)
            If longMinutes >= 60 Then
                Hours = CInt(longMinutes \ 60)
                Minutes = CInt(longMinutes Mod 60)
            Else
                Minutes = CInt(longMinutes)
            End If
            Seconds = CInt(singleSeconds * 60)
        Case tutSeconds
            Dim longSeconds As Long
            longSeconds = CLng(inputTime)
            longMinutes = longSeconds \ 60
            Seconds = CInt(longSeconds Mod 60)
            Hours = CInt(longMinutes \ 60)
            Minutes = CInt(longMinutes Mod 60)
        Case tutHours
            singleHours = CSng(inputTime)
            Dim doubleMinutes   As Double
            Hours = CInt(Int(singleHours))
            'Use double to keep all possible precision
            doubleMinutes = 60# * (CDbl(singleHours) - CDbl(Int(singleHours)))
            Minutes = Int(doubleMinutes)
            Seconds = CSng(doubleMinutes) - Int(doubleMinutes)
            Seconds = CInt(singleSeconds)
        Case tutString
            Dim currentPosition As Integer
            Dim currentTempLine As String
            currentTempLine = inputTime
            currentPosition = InStr(currentTempLine, ":")
            Hours = CInt(Mid$(currentTempLine, 1, currentPosition - 1))
            currentTempLine = Mid$(currentTempLine, currentPosition + 1)
            currentPosition = InStr(currentTempLine, ":")
            Minutes = CInt(Mid$(currentTempLine, 1, currentPosition - 1))
            currentTempLine = Mid$(currentTempLine, currentPosition + 1)
            Seconds = CInt(currentTempLine)
        Case Else
            Debug.Assert False
    End Select
    'Make sure in and out are same
    'Debug.Assert inputTime = FormatTime(timeUnitType)
    ParseTime = timeUnitType
    Exit Function
ErrHan:
    ParseTime = tutError
    MsgBox "Error parsing time data: " & vbCrLf & inputTime & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, "FiberTrack Data File Read Error"
    Exit Function
End Function
