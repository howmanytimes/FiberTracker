VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
#If SCALE_DIAMETER Then
Private mi_DenierTarget             As Single
Private mi_DenierMax                As Single
Private mi_DenierCalibration        As Single
#ElseIf SCALE_DENIERTEMP Then
Private mi_DenierTarget             As Long
Private mi_DenierMax                As Long
Private mi_DenierCalibration        As Long
#Else
Private mi_DenierTarget             As Long
Private mi_DenierMax                As Long
Private mi_DenierCalibration        As Long
#End If

Private mi_LineSpeed                As Integer
Private mby_IntegrationTime         As Byte
Private mby_DenierTargetTolerance   As Byte
Private mi_ZeroCalInterval          As Integer
Private mby_PlotInterval            As Byte
Private mby_COMPort                 As Byte
Private mi_PlotViewIndex            As Integer
Private mi_DenierRangeIndex         As Integer
Private mi_COMPortIndex             As Integer
Private mby_PlotChannel             As Byte
Private mby_Level1SlubTolerance     As Long
Private mby_Level2SlubTolerance     As Long
Private mby_Level1Length            As Long
Private mby_Level2Length            As Long

Private msng_SensorCalFactor(1 To LAST_SENSOR)      As Single
Private mi_SensorCalValue(1 To LAST_SENSOR)         As Integer
Private mi_SensorZeroValue(1 To LAST_SENSOR)        As Integer
Private mb_SensorEnabled(1 To LAST_SENSOR)          As Boolean

Public Property Let Level1Length(ByVal vData As Long)
    mby_Level1Length = vData
End Property

Public Property Get Level1Length() As Long
    Level1Length = mby_Level1Length
End Property

Public Property Let Level1SlubTolerance(ByVal vData As Long)
    mby_Level1SlubTolerance = vData
End Property

Public Property Get Level1SlubTolerance() As Long
    Level1SlubTolerance = mby_Level1SlubTolerance
End Property

Public Property Let Level2Length(ByVal vData As Long)
    mby_Level2Length = vData
End Property

Public Property Get Level2Length() As Long
    Level2Length = mby_Level2Length
End Property

Public Property Let Level2SlubTolerance(ByVal vData As Long)
    mby_Level2SlubTolerance = vData
End Property

Public Property Get Level2SlubTolerance() As Long
    Level2SlubTolerance = mby_Level2SlubTolerance
End Property

Public Property Let PlotChannel(ByVal vData As Byte)
    mby_PlotChannel = vData
End Property

Public Property Get PlotChannel() As Byte
    PlotChannel = mby_PlotChannel
End Property

Public Property Let ComPortIndex(ByVal vData As Integer)
    mi_COMPortIndex = vData
End Property

Public Property Get ComPortIndex() As Integer
    ComPortIndex = mi_COMPortIndex
End Property

Public Property Let DenierRangeIndex(ByVal vData As Integer)
    mi_DenierRangeIndex = vData
End Property

Public Property Get DenierRangeIndex() As Integer
    DenierRangeIndex = mi_DenierRangeIndex
End Property

Public Property Let PlotViewIndex(ByVal vData As Integer)
    mi_PlotViewIndex = vData
End Property

Public Property Get PlotViewIndex() As Integer
    PlotViewIndex = mi_PlotViewIndex
End Property

Public Property Let ComPort(ByVal vData As Byte)
    mby_COMPort = vData
End Property

Public Property Get ComPort() As Byte
    ComPort = mby_COMPort
End Property

Public Property Let PlotInterval(ByVal vData As Byte)
    mby_PlotInterval = vData
End Property

Public Property Get PlotInterval() As Byte
    PlotInterval = mby_PlotInterval
End Property

Public Property Let ZeroCalInterval(ByVal vData As Integer)
    mi_ZeroCalInterval = vData
End Property

Public Property Get ZeroCalInterval() As Integer
    ZeroCalInterval = mi_ZeroCalInterval
End Property

#If SCALE_DIAMETER Then
Public Property Let DenierCalibration(ByVal vData As Single)
    mi_DenierCalibration = vData
End Property

Public Property Get DenierCalibration() As Single
    DenierCalibration = mi_DenierCalibration
End Property

Public Property Let DenierMax(ByVal vData As Single)
    mi_DenierMax = vData
End Property

Public Property Get DenierMax() As Single
    DenierMax = mi_DenierMax
End Property

Public Property Let DenierTarget(ByVal vData As Single)
    mi_DenierTarget = vData
End Property

Public Property Get DenierTarget() As Single
    DenierTarget = mi_DenierTarget
End Property
#Else
Public Property Let DenierCalibration(ByVal vData As Integer)
    mi_DenierCalibration = vData
End Property

Public Property Get DenierCalibration() As Integer
    DenierCalibration = mi_DenierCalibration
End Property

Public Property Let DenierMax(ByVal vData As Integer)
    mi_DenierMax = vData
End Property

Public Property Get DenierMax() As Integer
    DenierMax = mi_DenierMax
End Property

Public Property Let DenierTarget(ByVal vData As Integer)
    mi_DenierTarget = vData
End Property

Public Property Get DenierTarget() As Integer
    DenierTarget = mi_DenierTarget
End Property
#End If

Public Property Let DenierTargetTolerance(ByVal vData As Byte)
    mby_DenierTargetTolerance = vData
End Property

Public Property Get DenierTargetTolerance() As Byte
    DenierTargetTolerance = mby_DenierTargetTolerance
End Property

Public Property Let IntegrationTime(ByVal vData As Byte)
    mby_IntegrationTime = vData
End Property

Public Property Get IntegrationTime() As Byte
    IntegrationTime = mby_IntegrationTime
End Property

Public Property Let LineSpeed(ByVal vData As Integer)
    mi_LineSpeed = vData
End Property

Public Property Get LineSpeed() As Integer
    LineSpeed = mi_LineSpeed
End Property

Public Property Get SensorEnabled(index As Integer) As Boolean
    SensorEnabled = mb_SensorEnabled(index)
End Property

Public Property Let SensorEnabled(index As Integer, vData As Boolean)
    mb_SensorEnabled(index) = vData
End Property

Public Property Get SensorCalFactor(index As Integer) As Single
    SensorCalFactor = msng_SensorCalFactor(index)
End Property

Public Property Let SensorCalFactor(index As Integer, vData As Single)
    msng_SensorCalFactor(index) = vData
End Property

Public Property Get SensorCalValue(index As Integer) As Integer
    SensorCalValue = mi_SensorCalValue(index)
End Property

Public Property Let SensorCalValue(index As Integer, vData As Integer)
    mi_SensorCalValue(index) = vData
End Property

Public Property Get SensorZeroValue(index As Integer) As Integer
    SensorZeroValue = mi_SensorZeroValue(index)
End Property

Public Property Let SensorZeroValue(index As Integer, vData As Integer)
    mi_SensorZeroValue(index) = vData
End Property

Public Function ReadFromFile(inputFileName As String) As Boolean
On Error GoTo ErrHan
    Dim numberLine As Integer
    Dim index As Integer
    numberLine = FreeFile
    Debug.Print "FTSetup: Reading from " & inputFileName & " . . .";
    Open inputFileName For Input As #numberLine
    Input #numberLine, mi_LineSpeed, mi_DenierTarget, mi_DenierMax, mby_IntegrationTime
    Input #numberLine, mby_DenierTargetTolerance
    Input #numberLine, mi_ZeroCalInterval, mby_PlotInterval, mby_COMPort
    Input #numberLine, mi_PlotViewIndex, mi_DenierRangeIndex, mi_COMPortIndex
    Input #numberLine, mi_DenierCalibration, mby_PlotChannel
    For index = 1 To LAST_SENSOR
        Input #numberLine, mb_SensorEnabled(index), mi_SensorZeroValue(index)
        Input #numberLine, msng_SensorCalFactor(index), mi_SensorCalValue(index)
    Next index
    Input #numberLine, mby_Level1SlubTolerance, mby_Level1Length
    Input #numberLine, mby_Level2SlubTolerance, mby_Level2Length
    Close #numberLine
    Debug.Print "Done."
    ReadFromFile = True
    Exit Function
ErrHan:
    If Err.Number = 53 Then
        Debug.Print "Not Found."
        ReadFromFile = False
        Exit Function
    End If
    MsgBox "Error occurred reading setup file " & inputFileName & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbOKOnly, App.Title
End Function

Public Function WriteToFile(inputFileName As String) As Boolean
On Error GoTo ErrHan
    Dim numberLine As Integer
    Dim index As Integer
    numberLine = FreeFile
    Debug.Print "FTSetup: Writing to " & inputFileName & " . . .";
    Open inputFileName For Output As #numberLine
    Write #numberLine, mi_LineSpeed, mi_DenierTarget, mi_DenierMax, mby_IntegrationTime, mby_DenierTargetTolerance
    Write #numberLine, mi_ZeroCalInterval, mby_PlotInterval, mby_COMPort
    Write #numberLine, mi_PlotViewIndex, mi_DenierRangeIndex, mi_COMPortIndex, mi_DenierCalibration, mby_PlotChannel
    For index = 1 To LAST_SENSOR
        Write #numberLine, mb_SensorEnabled(index), mi_SensorZeroValue(index), msng_SensorCalFactor(index), mi_SensorCalValue(index)
    Next index
    Write #numberLine, mby_Level1SlubTolerance, mby_Level1Length, mby_Level2SlubTolerance, mby_Level2Length
    Close #numberLine
    Debug.Print "Done."
    WriteToFile = True
    Exit Function
ErrHan:
    If Err.Number = 53 Then
        Debug.Print "Not Found."
        WriteToFile = False
        Exit Function
    End If
    MsgBox "Error occurred saving setup file " & inputFileName & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbOKOnly, App.Title
End Function
