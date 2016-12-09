VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTExcelController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_GENERIC_OBJECTS = -1

'OptCompGuard
#If USE_GENERIC_OBJECTS Then
Private mo_ExcelObject        As Object
#Else
Private mo_ExcelObject        As Excel.Application
#End If

#If SCALE_DIAMETER Then
Private mi_DenierTarget As Single
#ElseIf SCALE_DENIERTEMP Then
Private mi_DenierTarget As Integer
#Else
Private mi_DenierTarget As Integer
#End If

Private ms_Header       As String
Private msi_IntTime     As Byte
Private mi_LineSpeed    As Integer
Private mi_NumSensors   As Integer
Private ml_NumValues    As Long

Public Sub ViewDataFileInExcel(datFileName As String)
On Error GoTo ErrHan
    Dim oldMousePointer As Integer
    oldMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    If Connect() Then
        Dim excelFileName As String
        'Excel File Name is DAT with a different extension.
        excelFileName = Mid$(datFileName, 1, Len(datFileName) - Len(".dat")) & ".xls"
        If LCase$(excelFileName) = LCase$(datFileName) Then
            'If source file is a fake ".xls", then set the new filename to be a random, time-based name
            excelFileName = Mid$(datFileName, 1, Len(datFileName) - Len(".dat")) & "-" & CStr(CLng(Timer)) & ".xls"
        End If
        Dim leaveWorkbookOpen As Boolean
        leaveWorkbookOpen = ConvertDATToExcel(datFileName, excelFileName)
        Disconnect leaveWorkbookOpen
    End If
ErrHan:
    'Fall Through!
    Screen.MousePointer = oldMousePointer
    If Err.Number = 0 Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly Or vbExclamation, "Viewing Data File in Excel"
    Err.Clear
    Exit Sub
End Sub

#If USE_GENERIC_OBJECTS Then
Private Function ApplyFormatting(workBook As Object) As Boolean
    Dim worksheet As Object
#Else
Private Function ApplyFormatting(workBook As Excel.workBook) As Boolean
    Dim worksheet As Excel.worksheet
#End If

    'Format Data Worksheet
    Set worksheet = workBook.Worksheets.Item(1)
    worksheet.name = "Data"
    worksheet.Rows(1).Font.Bold = True         'Format header row typeface
    AddChart worksheet, False
    ApplyFormatting = True
End Function

Private Function Connect() As Boolean
On Error Resume Next
    Dim excelProgramId As String
    excelProgramId = GetIniSetting("Application", "ExcelId")
    Dim excelPath As String
    excelPath = GetIniSetting("Application", "ExcelPath")
    Dim excelStartApp As String
    excelStartApp = StringFormat("{0} /e", excelPath)
    Dim oldTimer As Single
    oldTimer = Timer
    'Wait for any existing Excel process to become available
    Do While (Timer - oldTimer < 2!) And (mo_ExcelObject Is Nothing)
        Set mo_ExcelObject = GetObject(Class:=excelProgramId)
    Loop
    If mo_ExcelObject Is Nothing Then
        'Print errors if any
        Debug.Print StringFormat("GetObject failed -- Error {0}: {1}", Err.Number, Err.Description)
        Err.Clear
        'Create new Excel process
        oldTimer = Timer
        Do While (Timer - oldTimer < 2!) And (mo_ExcelObject Is Nothing)
            Set mo_ExcelObject = CreateObject(excelProgramId)
        Loop
        If mo_ExcelObject Is Nothing Then
            If Err.Number Then
                Debug.Print StringFormat("CreateObject failed -- Error {0}: {1}", Err.Number, Err.Description)
                Err.Clear
            End If
            oldTimer = Timer
            Do While (Timer - oldTimer < 2!) And (mo_ExcelObject Is Nothing)
                #If Not USE_GENERIC_OBJECTS Then
                    Set mo_ExcelObject = New Excel.Application
                #End If
            Loop
        End If
    
        If (mo_ExcelObject Is Nothing) Then
            If Err.Number Then
                'Print errors from both CreateObject and New attempts (e.g. print CreateObject error if New succeeds)
                Debug.Print StringFormat("ERROR {0}: {1}", Err.Number, Err.Description)
                If mo_ExcelObject Is Nothing Then
                    'Report fatal errors to user
                    If Err.Number = -2147221231 Then
                        MsgBox StringFormat("Error instancing Microsoft Excel.  If error persists, try rebooting.{0}Error {1}: {2}", vbCrLf, CStr(Err.Number), Err.Description), vbCritical Or vbOKOnly, GetIniSetting("Constants", "Name")
                    Else
                        MsgBox StringFormat("Error instancing Microsoft Excel.{0}Error {1}: {2}", vbCrLf, CStr(Err.Number), Err.Description), vbCritical Or vbOKOnly, GetIniSetting("Constants", "Name")
                    End If
                    Connect = False
                    Exit Function
                End If
                Err.Clear
            End If
        End If
    End If
    mo_ExcelObject.Visible = False
    mo_ExcelObject.Interactive = False
    Connect = True
End Function

Private Function Disconnect(leaveWorkbookOpen As Boolean) As Boolean
    If leaveWorkbookOpen Then
        Disconnect = ActivateExcelApp
    Else
        Disconnect = CloseExcelApp
    End If
    Set mo_ExcelObject = Nothing
End Function

Private Function ConvertDATToExcel(datFileName As String, excelFileName As String) As Boolean
    Debug.Assert Len(datFileName) > 0
    Debug.Assert Len(excelFileName) > 0
    Debug.Assert Not mo_ExcelObject Is Nothing
    Dim cvsFileName   As String
    cvsFileName = GetTempCSVFileName
    If Len(cvsFileName) > 0 Then
        If ConvertDATToCSV(datFileName, cvsFileName) Then
            ConvertDATToExcel = ConvertCSVToExcel(cvsFileName, excelFileName)
            'Delete Temporary File
            If Not SafeKill(cvsFileName) Then
                Debug.Assert False
            End If
        End If
    End If
End Function

Private Function ConvertCSVToExcel(cvsFileName As String, excelFileName As String) As Boolean
#If USE_GENERIC_OBJECTS Then
    Dim workBook       As Object
    Dim worksheet      As Object
#Else
    Dim workBook       As Excel.workBook
    Dim worksheet      As Excel.worksheet
#End If
    Debug.Assert Not mo_ExcelObject Is Nothing
    Debug.Assert Len(cvsFileName) > 0
    Debug.Assert Len(excelFileName) > 0
    If Not SafeKill(excelFileName) Then
        ConvertCSVToExcel = False
        Exit Function
    End If
    Set workBook = mo_ExcelObject.Workbooks.Open(cvsFileName, Editable:=True, AddToMRU:=False)
#If USE_GENERIC_OBJECTS Then
    workBook.SaveAs excelFileName, FileFormat:=43, CreateBackup:=True, AddToMRU:=True
#Else
    workBook.SaveAs excelFileName, FileFormat:=xlExcel9795, CreateBackup:=True, AddToMRU:=True
#End If
    If ApplyFormatting(workBook) Then
        workBook.Save
        ConvertCSVToExcel = True
    Else
        ConvertCSVToExcel = False
    End If
End Function

Private Function ConvertDATToCSV(datFileName As String, cvsFileName As String) As Boolean
    Dim datFile As New FTDataFile
    Dim csvFile As New FTDataFile
    Dim index As Integer
    Dim counter As Long
    If Not datFile.ReadFile(datFileName) Then
        ConvertDATToCSV = False
        Exit Function
    End If
    With datFile
        msi_IntTime = .IntegrationTime
        mi_DenierTarget = .DenierTarget
        mi_LineSpeed = .LineSpeed
        mi_NumSensors = .NumSensors
        ml_NumValues = .NumValues
        ms_Header = .Title & vbLf & .Info & vbLf & .FileName
    End With
    With csvFile
        .AppendMode = False
        .CompareMode = False
        .FileFormat = Csv
        .FileName = cvsFileName
        If mi_NumSensors = 0 Then
            .SensorOnline(1) = True     'Kludge to make display less f'd up.
        Else
            For index = 1 To LAST_SENSOR
                .SensorOnline(index) = datFile.SensorOnline(index)
            Next index
        End If
        If Not .OpenFile Then
            ConvertDATToCSV = False
            Exit Function
        End If
        .WriteHeader
        For counter = 0 To ml_NumValues
            .WriteDataLine datFile.DataLine(counter)
        Next counter
        .WriteFooter
        .CloseFile
    End With
    ConvertDATToCSV = True
End Function

Private Function ActivateExcelApp() As Boolean
    mo_ExcelObject.Visible = True
    mo_ExcelObject.Interactive = True
    Debug.Print "Activating Window: " & mo_ExcelObject.Caption
    AppActivate mo_ExcelObject.Caption
End Function

Private Function CloseExcelApp() As Boolean
    If Not (mo_ExcelObject.ActiveWorkbook Is Nothing) Then
        mo_ExcelObject.ActiveWorkbook.Close SaveChanges:=False
    End If
    If mo_ExcelObject.Workbooks.Count = 0 Then
        mo_ExcelObject.Quit
    End If
End Function

#If USE_GENERIC_OBJECTS Then
Private Sub AddChart(worksheet As Object, isPlaceInWorksheet As Boolean)
    Dim o_Dest          As Object
    Dim o_ChartObjects  As Object
    Dim o_ChartObject   As Object
    Dim o_Chart         As Object
#Else
Private Sub AddChart(worksheet As Excel.worksheet, isPlaceInWorksheet As Boolean)
    
    Dim o_Dest          As Excel.range
    Dim o_ChartObjects  As Excel.ChartObjects
    Dim o_ChartObject   As Excel.ChartObject
    Dim o_Chart         As Excel.Chart
#End If
    If isPlaceInWorksheet Then
        Set o_Dest = worksheet.range("F8")     'Place chart in cell F7.
        Set o_ChartObjects = worksheet.ChartObjects
        Set o_ChartObject = o_ChartObjects.add(o_Dest.Left, o_Dest.Top, o_Dest.Width * 12, o_Dest.Height * 20)
        Set o_Chart = o_ChartObject.Chart
    Else
        Set o_Chart = worksheet.Application.Charts.add(After:=worksheet)
        o_Chart.name = "Line Chart"
    End If
        
#If USE_GENERIC_OBJECTS Then
    o_Chart.SetSourceData worksheet.UsedRange, 2
    
    o_Chart.Type = 4
    
    Dim o_XAxis     As Object
    Dim o_YAxis     As Object
    Dim o_Legend    As Object
    
    Set o_XAxis = o_Chart.Axes(1)
    Set o_YAxis = o_Chart.Axes(2)
#Else
    o_Chart.SetSourceData worksheet.UsedRange, xlByColumns
    
    o_Chart.Type = XlChartType.xlLine
    
    Dim o_Axes      As Excel.Axes
    Dim o_XAxis     As Excel.Axis
    Dim o_YAxis     As Excel.Axis
    Dim o_Legend    As Excel.Legend
    
    Set o_Axes = o_Chart.Axes
    Set o_XAxis = o_Axes.Item(xlCategory)
    Set o_YAxis = o_Axes.Item(xlValue)
#End If
        
    o_YAxis.MaximumScaleIsAuto = True
    o_YAxis.MinimumScaleIsAuto = True
    o_YAxis.HasMajorGridlines = True
    
    o_XAxis.HasTitle = True
    o_XAxis.AxisTitle.Caption = "Time Elapsed"
    
#If USE_GENERIC_OBJECTS Then
    o_XAxis.CategoryType = -4105
#Else
    o_XAxis.CategoryType = xlAutomaticScale
#End If

    o_YAxis.HasTitle = True
#If SCALE_DIAMETER Then
    o_YAxis.AxisTitle.Caption = "Diameter"
#ElseIf SCALE_DENIERTEMP Then
    o_YAxis.AxisTitle.Caption = "Temperature"
#Else
    o_YAxis.AxisTitle.Caption = "Denier"
#End If

    o_Chart.HasTitle = True
    o_Chart.ChartTitle.Caption = ms_Header      'Line Speed, Denier Target, Int. Time data
    o_Chart.ChartTitle.Font.Size = 14
'    o_Chart.ChartTitle.AutoScaleFont = True
    
    o_Chart.Deselect
End Sub

Private Function GetTempCSVFileName() As String
    Dim fso As New Scripting.fileSystemObject
    Dim tempFile As String
    ' tempFile = fso.GetFolder(App.Path) & "\" & fso.GetTempName & ".csv"
    tempFile = fso.GetSpecialFolder(TemporaryFolder).Path & "\" & fso.GetTempName & ".csv"
    Debug.Print tempFile
    GetTempCSVFileName = tempFile
End Function

#If False Then
'Older way of converting. Just keeping for reference. Obsolete and may be deleted.
Sub Main()
    ConvertDATToExcel "C:\STI\CODE\TEST2.DAT", "C:\STI\CODE\TEST2.XLS", True
End Sub

Public Function ConvertDATToCSV2(datFileName As String, cvsFileName As String) As Boolean
    Dim fso As New Scripting.fileSystemObject
    Dim textStreamOut As TextStream
    Dim textStreamIn As TextStream
    Set textStreamIn = fso.OpenTextFile(datFileName, ForReading, False)
    Set textStreamOut = fso.CreateTextFile(cvsFileName, True)
    While Not textStreamIn.AtEndOfStream
        Dim line As String
        line = textStreamIn.ReadLine
        If Left$(line, 1) = "#" Then
            'handle header
            HandleHeader line
        Else
            textStreamOut.WriteLine line
        End If
    Wend
    textStreamOut.Close
    textStreamIn.Close
    ConvertDATToCSV2 = True
End Function

Private Sub HandleHeader(line As String)
    ms_Header = ms_Header & line & vbLf
End Sub

Private Sub AddComment(worksheet As Excel.worksheet)
    Dim commentCell   As range
    Dim commentColumn As Integer
    commentColumn = UBound(worksheet.UsedRange.Formula, 2) + 3
    Set commentCell = worksheet.Cells.Item("2", commentColumn)
    If commentCell Is Nothing Then Debug.Assert False
    
    commentCell.AddComment ms_Header
    commentCell.Comment.Shape.Width = commentCell.Comment.Shape.Width * 2
End Sub

Private Sub PrintInfo(range As Excel.range)
On Error GoTo ErrHan
    Debug.Print ""
    With range
        Debug.Print "Rows: " & LBound(.Formula, 1) & " To " & UBound(.Formula, 1)
        Debug.Print "Columns: " & LBound(.Formula, 2) & " To " & UBound(.Formula, 2)
        Debug.Print "Row: " & .Row
        Debug.Print "Column: " & .Column
        Debug.Print "Address: " & .address
        Debug.Print "Address-Local: " & .AddressLocal
        Debug.Print "Value: " & .value
        Debug.Print "Value2: " & .value2
    End With
    Exit Sub
ErrHan:
    Debug.Print StringFormat("ERROR: {0} - {1}", Err.Number, Err.Description)
    Resume Next
End Sub

#End If
