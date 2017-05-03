VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTDataFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'
'   FTDataFile encapsulates both read and write functionality for FiberTrack Data.
'
'   Each FiberTrack data reading is stored internally as an FTDataLine object.  See
'   FTDataLine for more information.
'
'   Supported file formats are enumerated in DataFileFormatEnum.  Text files
'   should have the exact same format as v2.3 output.  Excel files should have the
'   same format as "Excel" files in v2.3: the footer is on the bottom, and the file
'   is largely useless.  CSV files are a new format I added which work much better in
'   Excel.  Comma-separated values are a universal data interchange standard and so should
'   work well in other spreadsheet and database applications.
'
'   Examples:
'
'
'   Chris Muldoon
'   December 12, 2001

'#Const DATAFILE_SHOW_OUTPUT = -1

'FT3: FiberTrack Ver. 3 Data
'FTS: FiberTrack Setup
'CSV: Comma-Separated Values
'AID: AnalogInput Data
'AIF: AnalogInput FFT Data

Public Enum DataFileFormatEnum
    Unknown = 0
    Text = 1         '.DAT
    Excel = 2        '.XLS
    Csv = 3          '.CSV
    Fbt = 4          '.FBT format
End Enum

Private mb_AppendMode                       As Boolean
Private mb_CompareMode                      As Boolean      '.DAT, but special "Compare" file
Private mo_DataLines()                      As FTDataLine

#If SCALE_DIAMETER Then
Private mi_DenierTarget                     As Single
#ElseIf SCALE_DENIERTEMP Then
Private mi_DenierTarget                     As Integer
#Else
Private mi_DenierTarget                     As Integer
#End If

Private mb_DiameterMode                     As Boolean      '"Diameter-mode" DAT.
Private mb_FileOpen                         As Boolean
Private ms_FileName                         As String
Private mDataFileFormatEnum                 As DataFileFormatEnum
Private ms_Info                             As String
Private mbyt_IntegrationTime                As Byte
Private mi_LineSpeed                        As Integer      'm/sec
Private mb_OverWrite                        As Boolean
Private mi_NumSensors                       As Integer
Private ml_NumValues                        As Long
Private mb_SensorOnline(1 To LAST_SENSOR)   As Boolean
Private ms_Title                            As String
Private ms_ThreadLine                       As String

Private mFNum                               As Integer
Private me_TimeFormat                       As FTTimeUnitType
Private ml_CurLineNum                       As Long         'Used for error reporting, corresponds to actual line number w/in file (as opposed to data index, which l_NumLines corresponds to in ReadDataArray)


Private Const HEADER_TIME_PRE = "Integration Time = "
Private Const HEADER_TIME_POST = " Sec.   "

#If SCALE_DIAMETER Then
Private Const HEADER_DNR_PRE = "Target Diameter = "
Private Const HEADER_DNR_POST = "  "
#ElseIf SCALE_DENIERTEMP Then
Private Const HEADER_DNR_PRE = "Target Temperature = "
Private Const HEADER_DNR_POST = "  "
#Else
Private Const HEADER_DNR_PRE = "Target Denier = "
Private Const HEADER_DNR_POST = "    "
#End If

Private Const HEADER_SPEED_PRE = "Line Speed = "
Private Const HEADER_SPEED_POST = " Meters/sec."
Private Const HEADER_TIME = "  Time"
Private Const HEADER_SENSOR = "Sensor "
Private Const HEADER_CV = "CV"
Private Const HEADER_ORIG = "ORIG AVG"

Private Const ERRUI_TITLE = "FiberTrack DataFile"
Private Const FILE_EXT_TEXT = ".dat"
Private Const FILE_EXT_EXCEL = ".xls"
Private Const FILE_EXT_CSV = ".csv"
Private Const FILE_EXT_FT3 = ".ft3"
Private Const FILE_ROOT_NAME = "Fiber"

Public Property Let AppendMode(value As Boolean)
    mb_AppendMode = value
    If value Then Debug.Assert False      'Append mode doesn't really work.
End Property

'"Compare mode" is somewhat abandoned now that Excel is working.  It allows for
'division and subtraction of one set of data by another
Public Property Get CompareMode() As Boolean
    CompareMode = mb_CompareMode
End Property

Public Property Let CompareMode(value As Boolean)
    mb_CompareMode = value
End Property

Public Property Get DataLine(index As Long) As FTDataLine
    Debug.Assert index <= UBound(mo_DataLines)
    Set DataLine = mo_DataLines(index)
End Property

#If SCALE_DIAMETER Then
Public Property Get DenierTarget() As Single
    DenierTarget = mi_DenierTarget
End Property

Public Property Let DenierTarget(value As Single)
    mi_DenierTarget = value
End Property
#Else
Public Property Get DenierTarget() As Integer
    DenierTarget = mi_DenierTarget
End Property

Public Property Let DenierTarget(value As Integer)
    mi_DenierTarget = value
End Property
#End If

Public Property Get DenierValues(sensorIndex As Integer, index As Long) As Single
    Debug.Assert index <= UBound(mo_DataLines)
    Debug.Assert index >= LBound(mo_DataLines)
    DenierValues = mo_DataLines(index).avg(sensorIndex)
End Property

Public Property Get FileFormat() As DataFileFormatEnum
    FileFormat = mDataFileFormatEnum
End Property

Public Property Let FileFormat(value As DataFileFormatEnum)
    mDataFileFormatEnum = value
End Property

Public Property Get FileName() As String
    SetFileName
    FileName = ms_FileName
End Property

Public Property Let FileName(value As String)
    If Not mb_FileOpen Then
        ms_FileName = value
    Else
        Err.Raise 40000, Description:="File Already Open.  Cannot rename."
    End If
End Property
    
Public Property Get Info() As String
    Info = ms_Info
End Property

Public Property Get IntegrationTime() As Byte
    IntegrationTime = mbyt_IntegrationTime
End Property

Public Property Let IntegrationTime(value As Byte)
    mbyt_IntegrationTime = value
End Property

Public Property Get IsDiameter() As Boolean
    IsDiameter = mb_DiameterMode
End Property

'Line speed in meters per minute
Public Property Get LineSpeed() As Integer
    LineSpeed = mi_LineSpeed
End Property

Public Property Let LineSpeed(value As Integer)
    mi_LineSpeed = value
End Property

'Is overwriting existing files allowed?
Public Property Let OverWrite(value As Boolean)
Attribute OverWrite.VB_Description = "Is overwriting existing files allowed?"
    mb_OverWrite = value
End Property

Public Property Get NumSensors() As Integer
    NumSensors = mi_NumSensors
End Property

Public Property Get NumValues() As Long
    Debug.Assert ml_NumValues = UBound(mo_DataLines)
    NumValues = ml_NumValues
End Property

Public Property Get SensorOnline(index As Integer) As Boolean
    SensorOnline = mb_SensorOnline(index)
End Property

Public Property Let SensorOnline(index As Integer, value As Boolean)
    mb_SensorOnline(index) = value
End Property

Public Property Get Title() As String
    Title = ms_Title
End Property
' Package name
Public Property Get ThreadLine() As String
    ThreadLine = ms_ThreadLine
End Property

Public Function CreateCompareFile(inputDataFile As FTDataFile, sensorIndex1 As Integer, sensorIndex2 As Integer) As FTDataFile
On Error GoTo ErrHan
    Dim n As Long
    Dim index As Integer
    Dim numberOfValues As Long
    Dim outputDataFile As New FTDataFile
    Dim dataFile As FTDataFile
    Set dataFile = Me
    
    If dataFile.IntegrationTime <> inputDataFile.IntegrationTime Then
        MsgBox "Unable to compare specified files.  Files have different integration times.", vbOKOnly Or vbInformation, ERRUI_TITLE
        Exit Function
    End If

    outputDataFile.CompareMode = True
    outputDataFile.IntegrationTime = dataFile.IntegrationTime
    outputDataFile.LineSpeed = 0
    outputDataFile.DenierTarget = 100
    outputDataFile.FileName = App.Path & "\" & "compare.dat"
    
    For index = 1 To 4
        outputDataFile.SensorOnline(index) = True
    Next index
    outputDataFile.FileFormat = Text
    outputDataFile.AppendMode = False
    If Not outputDataFile.OpenFile Then
        Debug.Assert False
        Exit Function
    End If
    outputDataFile.WriteHeader
    'Choose the minimum number of values
    numberOfValues = dataFile.NumValues
    If numberOfValues > inputDataFile.NumValues Then
        numberOfValues = inputDataFile.NumValues
    End If
    Const PRETEND_INFINITY As Single = 10000!
    Dim value1 As Single
    Dim value2 As Single
    Dim substract As Single
    Dim divide As Single
    For n = 1 To numberOfValues
        value1 = dataFile.DenierValues(sensorIndex1, n)
        value2 = inputDataFile.DenierValues(sensorIndex2, n)
        substract = value1 - value2
        If value2 = 0 Then
            If value1 = 0 Then
                divide = 1
            Else
                divide = PRETEND_INFINITY
            End If
        Else
            divide = value1 / value2
        End If
        Dim currentDataLine  As New FTDataLine
        currentDataLine.avg(1) = value1
        currentDataLine.avg(2) = value2
        currentDataLine.avg(3) = substract
        currentDataLine.avg(4) = divide
        outputDataFile.WriteDataLine currentDataLine
    Next n
    outputDataFile.CloseFile
    Set CreateCompareFile = outputDataFile
    Exit Function
ErrHan:
    MsgBox "Error occurred creating comparison file." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbOKOnly, ERRUI_TITLE
    Err.Clear
    Exit Function
End Function

Public Sub CloseFile()
    Close #mFNum
    mb_FileOpen = False
End Sub

Public Sub EraseFile()
    Kill ms_FileName
End Sub

Public Sub WriteHeader()
    If Not mb_FileOpen Then
        Err.Raise 40000, Description:="File not open."
    End If
    SetTitleAndInfo
    Select Case mDataFileFormatEnum
        Case Text
            WriteHeaderText
        Case Csv
            WriteHeaderCSV
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub WriteHeaderText()
    Debug.Assert mDataFileFormatEnum = Text
    Dim index As Integer
    WriteLine ms_Title
    WriteLine "       --------------------------------------------------------------"
    WriteLine ""
    WriteLine ms_Info
    WriteLine ""
    WriteLine ""
    WriteComma ""
    Dim line As String
    Dim blank As String
    blank = "             "
    For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
        If mb_SensorOnline(index) Then
            line = CStr(GetIniSetting("Application", "NameOfThreadLine") & " " & RTrim(SensorInfos(index).Package))   ' CHANGED
            If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                line = line & blank
            End If
            If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                line = line & blank
            End If
            WriteComma Replace(line, vbNewLine, "")
        End If
    Next index
    WriteLine ""
    WriteComma HEADER_TIME
    For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
        If mb_SensorOnline(index) Then
            WriteComma HEADER_SENSOR & index
            If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                WriteComma HEADER_CV & index
            End If
            If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                WriteComma HEADER_ORIG & index
            End If
        End If
    Next index
    WriteLine ""
    WriteLine "--------------------------------------------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub WriteHeaderCSV()
    Dim index As Integer
    Dim line As String
    line = HEADER_TIME
    For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
        If mb_SensorOnline(index) Then
            line = line & ", " & HEADER_SENSOR & CStr(index) & " " & GetIniSetting("Application", "NameOfThreadLine") & " " & RTrim(SensorInfos(index).Package) ' CHANGED
            If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                line = line & ", " & HEADER_CV & CStr(index)
            End If
            If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                line = line & ", " & HEADER_ORIG & CStr(index)
            End If
        End If
    Next index
    WriteLine line
End Sub

Public Sub WriteFooter()
    If Not mb_FileOpen Then
        Err.Raise 40000, Description:="File not open."
    End If
    If mDataFileFormatEnum <> Excel Then Exit Sub
    Dim index As Integer
    WriteLine ""
    WriteComma HEADER_TIME
    For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
        If mb_SensorOnline(index) Then
            WriteComma HEADER_SENSOR & index & " " & GetIniSetting("Application", "NameOfThreadLine") & " " & SensorInfos(index).Package  ' CHANGED
            If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                WriteComma HEADER_CV & index
            End If
            If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                WriteComma HEADER_ORIG & index
            End If
        End If
    Next index
    WriteLine ""
    WriteLine ""
    SetTitleAndInfo
    WriteLine ms_Title
    WriteLine ms_Info
End Sub

Public Sub WriteDataLine(currentDataLine As FTDataLine)
    If Not mb_FileOpen Then
        Err.Raise 40000, Description:="File not open."
    End If
    Dim index As Integer
    Dim line As String
    Select Case mDataFileFormatEnum
        Case Text, Excel
            WriteTime currentDataLine
            For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
                If mb_SensorOnline(index) Then
                    WriteData currentDataLine.avg(index)
                    If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                        WriteData currentDataLine.cv(index)
                    End If
                    If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                        WriteData currentDataLine.OrigAvg(index)
                    End If
                End If
            Next index
            'Go to next line
            WriteLine ""
        Case Csv
            line = FormatTime(currentDataLine)
            For index = LBound(mb_SensorOnline) To UBound(mb_SensorOnline)
                If mb_SensorOnline(index) Then
                    line = line & ", " & FormatData(currentDataLine.avg(index))
                    If CBool(GetIniSetting("Application", "IncludeCvColumn")) Then
                        line = line & ", " & FormatData(currentDataLine.cv(index))
                    End If
                    If CBool(GetIniSetting("Application", "IncludeOriginalAvgColumn")) Then
                        line = line & ", " & FormatData(currentDataLine.OrigAvg(index))
                    End If
                End If
            Next index
            WriteLine line
        Case Else
            Debug.Assert False
    End Select
End Sub

'Write mode only - Set title and info header/footer lines.
Private Sub SetTitleAndInfo()
#If SCALE_DIAMETER Then
    If mb_CompareMode Then
        ms_Title = "          STC Diameter Comparison Data File - " & CStr(now)
    Else
        ms_Title = "          STC Diameter Monitoring Data File - " & CStr(now)
    End If
    ms_Info = HEADER_TIME_PRE & mbyt_IntegrationTime & HEADER_TIME_POST & HEADER_DNR_PRE & Format$(mi_DenierTarget, "0.000") & HEADER_DNR_POST & HEADER_SPEED_PRE & mi_LineSpeed & HEADER_SPEED_POST
#ElseIf SCALE_DENIERTEMP Then
    If mb_CompareMode Then
        ms_Title = "          STC Temperature Comparison Data File - " & CStr(now)
    Else
        ms_Title = "          STC Temperature Monitoring Data File - " & CStr(now)
    End If
    ms_Info = HEADER_TIME_PRE & mbyt_IntegrationTime & HEADER_TIME_POST & HEADER_DNR_PRE & Format$(mi_DenierTarget, "0.000") & HEADER_DNR_POST & HEADER_SPEED_PRE & mi_LineSpeed & HEADER_SPEED_POST
#Else
    If mb_CompareMode Then
        ms_Title = "            STC Denier Comparison Data File - " & CStr(now)
    Else
        ms_Title = "            STC Denier Monitoring Data File - " & CStr(now)
    End If
    ms_Info = HEADER_TIME_PRE & mbyt_IntegrationTime & HEADER_TIME_POST & HEADER_DNR_PRE & mi_DenierTarget & HEADER_DNR_POST & HEADER_SPEED_PRE & mi_LineSpeed & HEADER_SPEED_POST
#End If
End Sub

Private Sub WriteTime(currentDataLine As FTDataLine)
    WriteSemi FormatTime(currentDataLine)
End Sub

Private Function FormatTime(currentDataLine As FTDataLine) As String
    With currentDataLine
        FormatTime = Format$(.Hours, "00") & ":" & Format(.Minutes, "00") & ":" & Format(.Seconds, "00")
    End With
End Function

Private Sub WriteData(floatAverage As Single)
    WriteComma ""
    WriteSemi FormatData(floatAverage)
End Sub

Private Function FormatData(floatAverage As Single) As String
#If SCALE_DIAMETER Then
    FormatData = Format$(floatAverage, "0.000")
#Else
    FormatData = Format$(floatAverage, "####.0")
#End If
End Function

'Write appending with semi-colon
Private Sub WriteSemi(value As String)
    Debug.Assert mb_FileOpen
    Print #mFNum, value;
#If DATAFILE_SHOW_OUTPUT Then
    Debug.Print value;
#End If
End Sub

Private Sub WriteComma(value As String)
    Debug.Assert mb_FileOpen
    Print #mFNum, value,
#If DATAFILE_SHOW_OUTPUT Then
    Debug.Print value,
#End If
End Sub

Private Sub WriteLine(value As String)
    Debug.Assert mb_FileOpen
    Print #mFNum, value
#If DATAFILE_SHOW_OUTPUT Then
    Debug.Print value
#End If
End Sub

'Opens file for writing
Public Function OpenFile() As Boolean
On Error GoTo ErrHan
    If mb_FileOpen Then
        MsgBox "Unexpected condition: File Already Opened.", vbCritical Or vbOKOnly, ERRUI_TITLE
        Debug.Assert False
        OpenFile = False
        Exit Function
    End If
    SetFileName
    If Not mb_OverWrite Then
        'Don't kill an existing file
        Dim fso As New Scripting.fileSystemObject
        If fso.FileExists(ms_FileName) Then
            Do Until Not fso.FileExists(ms_FileName)
                ms_FileName = ""
                SetFileName
            Loop
        End If
    End If
    If mb_AppendMode Then
        Open ms_FileName For Append As #mFNum
    Else
        Open ms_FileName For Output As #mFNum
    End If
    mb_FileOpen = True
    OpenFile = mb_FileOpen
    Exit Function
ErrHan:
    If Err.Number = 70 Then
        'Permission Denied (e.g. file open by another process)
        If ms_FileName = GetDefaultFileName Then
            'System-specified filename
            ms_FileName = ""
            'increment the incrementor
            SetFileName
            'Try again
            Resume
        Else
            'User-specified filename
            Dim uResult As VbMsgBoxResult
            uResult = MsgBox("Error 70 (Permission Denied) occurred while attempting to write to file.  File may be in use." & vbCrLf & "File: " & ms_FileName & vbCrLf & vbCrLf & "Press Yes to retry opening file, or No to cancel.", vbCritical Or vbRetryCancel, ERRUI_TITLE)
            If uResult = vbRetry Then
                Resume
            Else
                mb_FileOpen = False
                OpenFile = False
                Exit Function
            End If
        End If
    Else
        MsgBox "Error occurred opening file." & vbCrLf & "File: " & ms_FileName & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbOKOnly, ERRUI_TITLE
        mb_FileOpen = False
        OpenFile = False
    End If
End Function

'If filename hasn't been specified, increments RunNumber and sets filename (ms_FileName variable)
Private Sub SetFileName()
    Dim tempFileName As String
    ms_FileName = Trim$(ms_FileName)
    If ms_FileName = "" Then
        'Kludge: use old logic for filename incrementor
        RunNumber = RunNumber + 1
        ms_FileName = GetDefaultFileName()
    Else
        If InStr(ms_FileName, "\") = 0 Then
            'Set path
            ms_FileName = App.Path & "\" & ms_FileName
        End If
    End If
End Sub

Private Function GetDefaultFileName() As String
    Dim result As String
    If mDataFileFormatEnum = Unknown Then
        'Not enough information to set the filename
        result = ""
    Else
        result = GetDefaultPath & "\" & FILE_ROOT_NAME & CStr(RunNumber) & "_" & Format(now, GetIniSetting("Application", "FileDateFormatAppend")) & GetFileExtension(mDataFileFormatEnum) ' CHANGED
    End If
    GetDefaultFileName = result
End Function

Private Function GetFileExtension(inputDataFileFormatEnum As DataFileFormatEnum) As String
    Dim result As String
    Select Case inputDataFileFormatEnum
        Case Unknown
            'Not enough information to set the filename
            result = ""
        Case Excel
            result = FILE_EXT_EXCEL
        Case Text
            result = FILE_EXT_TEXT
        Case Csv
            result = FILE_EXT_CSV
        Case Else
            Debug.Assert False
    End Select
    GetFileExtension = result
End Function

Private Function GetDefaultPath() As String
    GetDefaultPath = App.Path
End Function

'READ MODE FUNCTIONS
Public Function ReadFile(inputFileName As String) As Boolean
On Error GoTo ErrHan
    Dim isRead As Boolean
    ms_FileName = inputFileName
    Select Case LCase$(Right$(ms_FileName, Len(".ext")))
        Case FILE_EXT_EXCEL
            mDataFileFormatEnum = Excel
            ReadFile = ReadExcelFile(ms_FileName)
            Exit Function
        Case FILE_EXT_CSV
            mDataFileFormatEnum = Csv
            MsgBox "Comma-Separated Values files do not contain" & vbCrLf & "Date, Target Denier, or other metadata.", vbOKOnly Or vbInformation, "Note"
        Case Else       'Including ".dat"
            mDataFileFormatEnum = Text
    End Select
    Open ms_FileName For Input As #mFNum
    isRead = ReadHeader
    If isRead Then isRead = ReadDataArray
    Close #mFNum
    ReadFile = isRead
    Exit Function
ErrHan:
    MsgBox "Error occurred while processing file " & ms_FileName & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, ERRUI_TITLE
    Close #mFNum
    ReadFile = False
    Exit Function
End Function

'Read an "Excel" file.
'Return false if file is binary.  Read file twice.  First time through, read footer.
'Second time through, parse data.
Private Function ReadExcelFile(excelFileName As String) As Boolean
On Error GoTo ErrHan
    Dim fso As Scripting.fileSystemObject
    Dim inTextStream As TextStream
    Dim outTextStream As TextStream
    Dim currentLine As FTDataLine
    Dim currentInfo As String
    Dim currentTitle As String
    Dim currentThreadLine As String
    Dim columns As String
    Dim isInfo As Boolean
    Dim isTitle As Boolean
    Dim isThreadLine As Boolean
    Dim isColumns As Boolean
    Dim isReturn As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(excelFileName) Then
        MsgBox "Specified source file " & excelFileName & " does not exist.", vbCritical Or vbOKOnly, "Error Reading Excel Data File"
        ReadExcelFile = False
        Exit Function
    End If
    Set inTextStream = fso.OpenTextFile(excelFileName, ForReading, False)
    Debug.Assert Not inTextStream Is Nothing
    isReturn = True
    'First Pass Through
    Do While Not inTextStream.AtEndOfStream
        Dim lineValue  As String
        lineValue = inTextStream.ReadLine
        'Read Footer
        If Left$(Trim$(lineValue), Len("Time")) = "Time" Then
            isColumns = True
            isReturn = ParseColumnLine(lineValue)
        ElseIf Left$(Trim$(lineValue), Len("STC")) = "STC" Then
            isTitle = True
            ms_Title = lineValue
            isReturn = ParseTitleLine(lineValue)
            isReturn = ParseColumnLine(lineValue)
        ElseIf Left$(Trim$(lineValue), Len(GetIniSetting("Application", "NameOfThreadLine"))) = GetIniSetting("Application", "NameOfThreadLine") Then
            isThreadLine = True
            ms_ThreadLine = lineValue
            isReturn = True 'ParseTitleLine(lineValue)
        ElseIf Left$(Trim$(lineValue), Len("Integration")) = "Integration" Then
            isInfo = True
            ms_Info = lineValue
            isReturn = ParseInfoLine(lineValue)
        End If
        If Not isReturn Then
            ReadExcelFile = False
            Exit Function
        End If
    Loop
    inTextStream.Close
    If Not isColumns Or Not isTitle Or Not isInfo Then
        MsgBox "File format error reading " & excelFileName & vbCrLf & "Try opening this file using Microsoft Excel." & vbCrLf & "Code: " & (CInt(isColumns) * 1 + CInt(isTitle) * 2 + CInt(isInfo) * 4), vbCritical Or vbOKOnly, "Error Reading Excel Data File"
        ReadExcelFile = False
        Exit Function
    End If
    'Second Pass!!!
    Set inTextStream = fso.OpenTextFile(excelFileName, ForReading, False)
    Debug.Assert Not inTextStream Is Nothing
    Dim numberOfLines  As Long
    numberOfLines = 0
    ml_CurLineNum = 0
    Do While Not inTextStream.AtEndOfStream
        ml_CurLineNum = ml_CurLineNum + 1
        lineValue = inTextStream.ReadLine
        If Len(Trim$(lineValue)) = 0 Then Exit Do          'Start of footer
        Set currentLine = ParseDataLine(lineValue)
        If currentLine Is Nothing Then
            ReadExcelFile = False
            Exit Function
        End If
        ReDim Preserve mo_DataLines(0 To numberOfLines)    'This is *REALLY* slow
        Set mo_DataLines(numberOfLines) = currentLine
        numberOfLines = numberOfLines + 1
    Loop
    'Fix for first-data-line problems (old program wrote really large, bogus time)
    If mo_DataLines(0).TimeAsSeconds > mo_DataLines(1).TimeAsSeconds Then
        mo_DataLines(0).Hours = 0
        mo_DataLines(0).Minutes = 0
        mo_DataLines(0).Seconds = 0
    End If
    ml_NumValues = UBound(mo_DataLines)                 'Should be numberOfLines - 1, but we'll be safe
    ReadExcelFile = True
    Exit Function
ErrHan:
    ReadExcelFile = False
    Select Case Err.Number
        Case 429, 430
            MsgBox "Error instancing SCRRUN.DLL." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, "FiberTrack DataFile Error"
        Case Else
            MsgBox "Error reading Excel Data File at line " & ml_CurLineNum & "." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, "FiberTrack DataFile Error"
    End Select
    Err.Clear
End Function

Private Function ReadHeader() As Boolean
    Dim index As Integer
    Dim columns As String
    Dim headers(1 To 10) As String
    Dim isReturn As Boolean
    
    Const LINE_TITLE = 1
    Const LINE_DASHES1 = 2
    Const LINE_EMPTY1 = 3
    Const LINE_INFO = 4
    Const LINE_EMPTY2 = 5
    Const LINE_EMPTY3 = 6
    Const LINE_THREADLINE = 7
    Const LINE_EMPTY4 = 8
    Const LINE_COLUMNS = 9
    Const LINE_DASHES2 = 10
    
    ml_CurLineNum = 0
    For index = LBound(headers) To UBound(headers)
        ml_CurLineNum = ml_CurLineNum + 1
        headers(index) = ReadLine
    Next index
    ms_Title = headers(LINE_TITLE)
    ms_Info = headers(LINE_INFO)                ' Contains Denier, Speed, Time
    ms_ThreadLine = headers(LINE_THREADLINE)    ' Contains ThreadLines headers
    columns = headers(LINE_COLUMNS)             ' Contains Column headers
    
    isReturn = ParseTitleLine(ms_Title)
    If isReturn Then
        isReturn = ParseInfoLine(ms_Info)
    End If
'    If isReturn Then
'        isReturn = ParseThreadLineLine(ms_ThreadLine)
'    End If
    If isReturn Then
        isReturn = ParseColumnLine(columns)
    End If
    ReadHeader = isReturn
End Function

Private Function ParseTitleLine(inputTitle As String) As Boolean
    mb_CompareMode = (InStr(ms_Title, "Comparison") > 0)
    mb_DiameterMode = (InStr(ms_Title, "Diameter") > 0)
#If Not SCALE_DIAMETER Then
    If mb_DiameterMode Then
        MsgBox "This file was saved in Diameter mode and cannot be read " & vbCrLf & "by this software version.", vbExclamation Or vbOKOnly, StringFormat("{0} DataFile Error", GetIniSetting("Constants", "Name"))
        ParseTitleLine = False
        Exit Function
    End If
#ElseIf SCALE_DENIERTEMP Then
    If Not mb_DiameterMode Then
        MsgBox "This file was saved in Temperature mode and may not work properly.", vbInformation Or vbOKOnly, StringFormat("{0} DataFile Warning", GetIniSetting("Constants", "Name"))
    End If
#ElseIf SCALE_DIAMETER Then
    If Not mb_DiameterMode Then
        'Just warn, because earlier Diameter versions didn't tag the file
        MsgBox "This file was saved in Denier mode and may not work properly.", vbInformation Or vbOKOnly, StringFormat("{0} DataFile Warning", GetIniSetting("Constants", "Name"))
    End If
#End If
    ParseTitleLine = True
End Function

Private Function ParseInfoLine(inputInfo As String) As Boolean
    Dim currentDenier As String
    Dim currentTime As String
    Dim currentSpeed As String
    currentDenier = mid2$(inputInfo, HEADER_DNR_PRE, HEADER_DNR_POST & HEADER_SPEED_PRE)
    currentTime = mid2$(inputInfo, HEADER_TIME_PRE, HEADER_TIME_POST)
    currentSpeed = mid2$(inputInfo, HEADER_SPEED_PRE, HEADER_SPEED_POST)
    mi_DenierTarget = Int(currentDenier)
    mi_LineSpeed = Int(currentSpeed)
    mbyt_IntegrationTime = CByte(Int(currentTime))
    ParseInfoLine = True
End Function

Private Function ParseColumnLine(columns As String) As Boolean
    Dim tempLine As String
    Const TIME_STRING = "Time"
    columns = Trim$(columns)
    Dim position   As Integer
    position = InStr(columns, TIME_STRING)
    Debug.Assert position > 0
    tempLine = Trim$(Mid$(columns, position + Len(TIME_STRING)))
    Dim index As Integer
    Dim indexValue As String
    Dim oWordCounter As CWords
    Dim oWord As CWord
     
    Set oWordCounter = New CWords
    Call oWordCounter.CountWords(tempLine)
    For Each oWord In oWordCounter
        Debug.Print oWord.Word, oWord.Count
        If LCase$(oWord.Word) = LCase$(Trim$(HEADER_SENSOR)) Then
            Dim counter As Integer
            counter = 1
            Do
                If counter > oWord.Count Then Exit Do
                mb_SensorOnline(counter) = True
                mi_NumSensors = mi_NumSensors + 1
                counter = counter + 1
            Loop
        End If
    Next
    ParseColumnLine = True
End Function

'Private Function ParseThreadLineLine(inputThreadLine As String) As Boolean
'    Dim tempLine As String
'    Dim threadLineValue As String
'    threadLineValue = GetIniSetting("Application", "NameOfThreadLine")
'    inputThreadLine = Trim$(inputThreadLine)
'    Dim position As Integer
''    position = InStr(inputThreadLine, threadLineValue)
''    Debug.Assert position > 0
'    tempLine = inputThreadLine 'Trim$(Mid$(inputThreadLine, position + Len(threadLineValue)))
'    Dim index As Integer
'    Dim indexValue As String
'    Do
'        position = InStr(tempLine, threadLineValue)
'        If position <= 0 Then Exit Do
'        indexValue = (Mid$(tempLine, position + Len(threadLineValue), 1))
'        Debug.Assert IsNumeric(indexValue)
'        index = Int(indexValue)
'        If index >= LBound(mb_SensorOnline) And index <= UBound(mb_SensorOnline) Then
'            mb_SensorOnline(index) = True
'            mi_NumSensors = mi_NumSensors + 1
'        End If
'        tempLine = Trim$(Mid$(tempLine, position + Len(threadLineValue) + 1))
'    Loop
'    ParseThreadLineLine = True
'End Function

Private Function ReadDataArray() As Boolean
    Dim numberOfLines As Long
    Dim currentDataLine As FTDataLine
    numberOfLines = 0
    Do While Not AtEndOfFile
        ml_CurLineNum = ml_CurLineNum + 1                              'For error reporting
        Set currentDataLine = ParseDataLine(ReadLine)
        If currentDataLine Is Nothing Then
            ReadDataArray = False
            Exit Function
        End If
        ReDim Preserve mo_DataLines(0 To numberOfLines)            'This is *REALLY* slow
        Set mo_DataLines(numberOfLines) = currentDataLine
        numberOfLines = numberOfLines + 1
    Loop
    ml_NumValues = UBound(mo_DataLines)        'Should be numberOfLines - 1, but we'll be safe
    ReadDataArray = True
End Function

Private Function ParseDataLine(inputLine As String) As FTDataLine
    Dim position As Integer
    Dim index As Integer
    Dim currentDataLine  As New FTDataLine
    Dim tempLine As String
    Dim currentTime As String
    tempLine = Trim$(inputLine)
    If Len(tempLine) = 0 Then
        MsgBox "Error parsing data file " & ms_FileName & vbCrLf & "Empty Data Line at line number " & CStr(ml_CurLineNum), vbCritical Or vbOKOnly, "Error Parsing Data File"
        Set ParseDataLine = Nothing
        Exit Function
    Else
        position = InStr(tempLine, Space$(1))
        'This is for the case when no sensors are enabled,
        'or for first two data points in old FT.
        If position <= 0 Then position = Len(tempLine)
        currentTime = Trim$(Mid$(tempLine, 1, position))
        Dim resultTimeFormatEnum As FTTimeUnitType
        'Should maintain consistent time format throughout file.
        'Allow anything goes first time; next times must be consistent with first.
        If me_TimeFormat = tutUnknown Then
            resultTimeFormatEnum = currentDataLine.ParseTime(currentTime, tutUnknown)
        Else
            resultTimeFormatEnum = currentDataLine.ParseTime(currentTime, me_TimeFormat)
        End If
        If me_TimeFormat = tutUnknown Then
            me_TimeFormat = resultTimeFormatEnum
        End If
        If resultTimeFormatEnum = tutError Then
            Set ParseDataLine = Nothing
            Exit Function
        End If
    End If
    tempLine = Trim$(Mid$(tempLine, position))
    If Len(tempLine) > 0 Then
        Debug.Print tempLine
        Dim values As Variant
        Dim newValues As Variant
        Dim newIndex As Integer
        values = Split(tempLine, Space$(1))
        ReDim newValues(LBound(values) To UBound(values))
        For index = LBound(values) To UBound(values)
            If values(index) <> "" Then
                newIndex = newIndex + 1
                newValues(newIndex) = values(index)
            End If
        Next index
        ReDim Preserve newValues(LBound(values) To newIndex)
        Debug.Print newIndex
        Dim numberOfSensors As Integer
        Dim numberOfSensor As Integer
        numberOfSensor = 1
        If (CBool(GetIniSetting("Application", "IncludeCvColumn"))) Then
            numberOfSensors = (UBound(newValues) / 2)
            Dim isEven As Boolean
            For index = (LBound(newValues) + 1) To UBound(newValues)
                isEven = ((index Mod 2) = 0)
                If isEven Then
                    ' x is even
                    currentDataLine.cv(numberOfSensor) = CDbl(newValues(index))
                Else
                    ' x is odd
                    currentDataLine.avg(numberOfSensor) = CDbl(newValues(index))
                End If
                If index <= numberOfSensors And isEven Then
                    numberOfSensor = numberOfSensor + 1
                    Debug.Print numberOfSensor
                End If
            Next index
        Else
            ' handle single sensor lines without Cv column
            For index = (LBound(newValues) + 1) To UBound(newValues)
                currentDataLine.avg(index) = CDbl(newValues(index))
                currentDataLine.OrigAvg(index) = CDbl(newValues(index))
            Next index
        End If
    Else
        'Do nothing: "OLD" FT problem where blank data in first two rows.
    End If
    Set ParseDataLine = currentDataLine
End Function

Private Function mid2$(headerLine As String, leftValue As String, rightValue As String)
    Dim leftIndex As Integer
    Dim rightIndex As Integer
    leftIndex = InStr(headerLine, leftValue) + Len(leftValue) - 1
    rightIndex = InStr(headerLine, rightValue)
    mid2$ = Mid$(headerLine, leftIndex, rightIndex - leftIndex)
End Function

Private Function ReadLine() As String
    Dim value As String
    Line Input #1, value
    ReadLine = value
End Function

Private Function AtEndOfFile() As Boolean
    AtEndOfFile = EOF(mFNum)
End Function

Private Sub Class_Initialize()
    mFNum = FreeFile
    mDataFileFormatEnum = Unknown      'Require explicit client choice
    me_TimeFormat = tutUnknown
End Sub
