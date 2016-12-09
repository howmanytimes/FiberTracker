Attribute VB_Name = "modGlobals"
Option Explicit

'Software version
'SCALE_GENERIC, SCALE_DENIER, SCALE_DIAMETER can be defined in

'Set true to compile cv code - V1.6
Public Const COMPUTE_CV         As Boolean = True
'Display factor from legacy code - V1.6
Public Const DISFAC = 2048
'For plotting function
Public Const MIN_DENIER         As Integer = 0
'Change this for last channel #
Public Const LAST_SENSOR        As Integer = 8
'Show debug-related controls on screen
Public Const DEBUG_CONTROLS     As Boolean = False

Public RunInHours               As Long
Public RunInMinutes             As Long
Public RunInSeconds             As Long
Public RunNumber                As Long      'For data file creation
'***User supplied global information.
Public LineSpeed                As Long      'Line speed in meters/Sec.

#If SCALE_DIAMETER Then
Public Max_Denier               As Double
Public iMin_Denier              As Double
Public Calibration_Denier       As Double      'Separate denier setting for calibration - V1.7
Public Target_Denier            As Double
#ElseIf SCALE_DENIERTEMP Then
Public Max_Denier               As Long
Public iMin_Denier              As Long
Public Calibration_Denier       As Long      'Separate denier setting for calibration - V1.7
Public Target_Denier            As Long
#Else
Public Max_Denier               As Long
Public iMin_Denier              As Long
Public Calibration_Denier       As Long      'Separate denier setting for calibration - V1.7
Public Target_Denier            As Long
#End If

Public Integration_time         As Byte         'Integration time in seconds
Public Target_denier_tol        As Byte
Public Level1_slub_tol          As Byte
Public Level1_length            As Byte
Public Level2_slub_tol          As Byte
Public Level2_length            As Byte
Public ZeroCal_Interval         As Integer      'Zerocal interval in seconds
Public Plot_Interval            As Byte         'Plot window interval in minutes
'***End of user supplied information.

'Public IsSensorInit             As Boolean      'system undergoing initialization startup
Public TempComPort              As Byte

'Defines current system status
Public Type SystemInfo
    IsFormInit      As Boolean          'Completed initial load of software
    ComPort         As Byte             'Holds current Comm Port
    Enabled         As Boolean          'Completed system initialization
    Zero_cal        As Boolean          'Undergoing zero cal sequence
    Do_ZeroCal      As Boolean          'Flag to perform zero cals
    Gain_cal        As Boolean          'Undergoing gain cal sequence
    IsStartup       As Boolean          'Undergoing initial startup sequence
    IsRunning       As Boolean          'System currently running
    IsWait2Seconds  As Boolean          'Waiting 2 sec. for Comm routine
    IsWaitInitTime  As Boolean          'Waiting for integration period timer
    StartTime       As Variant          'For system runtime
    Next_Zerocal    As Integer          'Contains no. of seconds to next zerocal cycle
    PlotChannel     As Byte             'Current plotting channel
End Type

Public SystemStatus             As SystemInfo

'Defines current data packet as received from sensor.
Public Type SensorInfo
#If SCALE_DIAMETER Then
    AverageDiameter As Double
    Lowest          As Double           'Lowest average denier
    Highest         As Double           'Highest average denier
    LastAverage     As Double           'Last integration value
    SumOfAverages   As Double           'Accumulator for average denier
#ElseIf SCALE_DENIERTEMP Then
    AverageDiameter As Long
    Lowest          As Long          'Lowest average denier
    Highest         As Long          'Highest average denier
    LastAverage     As Long          'Last integration value
    SumOfAverages   As Long             'Accumulator for average denier
#Else
    AverageDiameter As Long
    Lowest          As Long          'Lowest average denier
    Highest         As Long          'Highest average denier
    LastAverage     As Long          'Last integration value
    SumOfAverages   As Long             'Accumulator for average denier
#End If
    Cv_Summary      As Long
'JW 5/25/00 Added this to store raw sensor data
    Raw_Denier      As Long          'Store the sensor raw denier
    Level1_Slub     As Long
    Level2_Slub     As Long
    PIC_Status      As Byte
    Zero_Value      As Long
    Cal_Value       As Long
    Cal_Factor      As Double           'Calibration constant
    Num_Cycles      As Long             'Count of integration values
    Error_Info      As Long
    PIC_Counter     As Byte
    Reserved        As Long
    Enabled         As Boolean          'Enabled via user interface
    Online          As Boolean          'Sensor communicates ok with host
    Out_to_Lunch    As Boolean          'Sensor fails to communicate with host
    Awaiting_Comm   As Boolean          'Waiting for sensor response packet
    Package         As String           ' CHANGED Package Text from the Textbox
End Type

Public SensorInfos(1 To LAST_SENSOR)    As SensorInfo
Public SensorColors(1 To LAST_SENSOR)   As Long
Public CurrentAverage(LAST_SENSOR)      As Double       'store current floatAverage
Public CurrentMean(LAST_SENSOR)         As Double       'store current mean
Public CurrentCv(LAST_SENSOR)           As Double       'store current cv
Public ReportType                       As String
Public Const IFILE_NAMES_LIMIT          As Integer = 4
Public FileNames(1 To IFILE_NAMES_LIMIT) As String
Public FileNamesNumber          As Integer      'FiberTrack
'Public FileNumber               As Integer      'Parameters
Public FileNumber               As Integer
Public FormLoadCount            As Integer      'fibertrack form_load switch

'Public comPort(0 To 1)          As String
Public ComPortIndex             As Integer

Public DenierRange(0 To 3)      As Integer
Public DenierRangeIndex         As Integer

Public IsOpenDataFile           As Boolean
Public IsMain                   As Boolean
Public IsNormal                 As Boolean
Public PlotView(0 To 1)         As String
Public PlotViewIndex            As Integer

'Public menuIndexString          As String
'Public menuIndex                As Integer

'************************************************************************
'  @doc    GUI-Functions
'  @func   Checks if a file exists.
'  @rdesc  Boolean - Returns True if file exists.
'  @parm   String  |File    |Name of File to check for
'  @comm   <f DoesFileExist> This function verifies if a file exists or not.
'  @devnote    KevinD 10/27/97 11:40:00AM
'  @xref   <f WriteToDisk>, <f SizeOfFile>
'
'************************************************************************
Public Function DoesFileExist(file As String) As Boolean
    Dim fileExist As String
    fileExist = Dir(file)
    If fileExist <> "" Or Null Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If
End Function

Public Sub Main()
    'Debug.Assert runUnitTests
    FormLoadCount = 1
    FiberTrack.Show
End Sub
