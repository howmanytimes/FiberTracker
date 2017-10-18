VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FiberTrack 
   BackColor       =   &H00808000&
   Caption         =   "FiberTrack"
   ClientHeight    =   8310
   ClientLeft      =   135
   ClientTop       =   705
   ClientWidth     =   11190
   FillColor       =   &H000000FF&
   Icon            =   "FiberTrack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7850.461
   ScaleMode       =   0  'User
   ScaleWidth      =   11190
   Tag             =   " "
   Begin VB.Frame fraApplicationIni 
      Caption         =   "Application Ini Settings"
      Height          =   3015
      Left            =   3720
      TabIndex        =   100
      Top             =   1560
      Width           =   7095
      Begin VB.TextBox txtNameOfThreadLine 
         Height          =   375
         Left            =   1920
         TabIndex        =   110
         Text            =   "NameOfThreadLine"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtAppendDateFormatToFile 
         Height          =   375
         Left            =   1920
         TabIndex        =   108
         Text            =   "AppendDateFormatToFile"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton cmdSaveIniSettings 
         Caption         =   "Save"
         Height          =   375
         Left            =   600
         TabIndex        =   106
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtExcelPathIniSettings 
         Height          =   375
         Left            =   1920
         MaxLength       =   250
         TabIndex        =   103
         Text            =   "Excel Path"
         Top             =   840
         Width           =   5055
      End
      Begin VB.CheckBox chkIncludeCvIniSettings 
         Caption         =   "Include Cv Column"
         Height          =   255
         Left            =   360
         TabIndex        =   102
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkOpenApplicationMaxIniSettings 
         Caption         =   "Open Application Max"
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelIniSettings 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   105
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblNameOfThreadLine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of ThreadLine:"
         Height          =   195
         Left            =   360
         TabIndex        =   109
         Top             =   1920
         Width           =   1530
      End
      Begin VB.Label lblAppendDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Format to File:"
         Height          =   195
         Left            =   360
         TabIndex        =   107
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label lblExcelPathIniSettings 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Path (exe):"
         Height          =   195
         Left            =   360
         TabIndex        =   104
         Top             =   960
         Width           =   1200
      End
   End
   Begin VB.Frame fraDefDet 
      Caption         =   "Defect Detection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3720
      TabIndex        =   88
      Top             =   1440
      Width           =   3375
      Begin VB.TextBox txtCurValueDD 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   94
         Text            =   "FiberTrack.frx":08CA
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdParDD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Level 1 Defect %"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   93
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanDD 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1920
         TabIndex        =   92
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdOkDD 
         Caption         =   "OK"
         Height          =   375
         Left            =   480
         TabIndex        =   91
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblCurValDD 
         Alignment       =   2  'Center
         Caption         =   "Current Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   90
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblParDD 
         Alignment       =   2  'Center
         Caption         =   "Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   89
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraAppPar 
      Caption         =   "Application Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3720
      TabIndex        =   80
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtCurValueAp 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   86
         Text            =   "FiberTrack.frx":08D3
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdCanAP 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   85
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdOkAP 
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         TabIndex        =   84
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdLineSpeed 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Line &Speed"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   83
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLineSpeed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "m/min"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   87
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblCurValueAP 
         Alignment       =   2  'Center
         Caption         =   "Current Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   82
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblParAP 
         Alignment       =   2  'Center
         Caption         =   "Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   81
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraDenierPar 
      Caption         =   "Orientation Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3720
      TabIndex        =   62
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton cmdDoneDP 
         Caption         =   "Done"
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdCanDP 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdOkDP 
         Caption         =   "OK"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtDenPar 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   66
         Text            =   $"FiberTrack.frx":08D8
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDenPar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orientation &Range"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblParValue 
         Alignment       =   2  'Center
         Caption         =   "Current Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPara 
         Alignment       =   2  'Center
         Caption         =   "Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraPlotChan 
      Caption         =   "Plot Channels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3720
      TabIndex        =   75
      Top             =   720
      Width           =   2175
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   15
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   135
      End
      Begin VB.CommandButton cmdCanPC 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdOkPC 
         Caption         =   "OK"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkPlotChan 
         Caption         =   "Channel &1"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   76
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraComPort 
      Caption         =   "Com Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   57
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton optCom2 
         Caption         =   "Com &2"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optCom1 
         Caption         =   "Com &1"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdCanCom 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdOKCom 
         Caption         =   "OK"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame fraDenRange 
      Caption         =   "Orientation Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6840
      TabIndex        =   70
      Top             =   960
      Width           =   3855
      Begin VB.CommandButton cmdCanDR 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdOkDR 
         Caption         =   "OK"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optDenRange 
         Caption         =   "200 to 2000 Denier"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fraSetup 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   960
      TabIndex        =   50
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdDone 
         Cancel          =   -1  'True
         Caption         =   "Done"
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdComPort 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&ComPort                          Alt+C"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About FiberTrack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2520
      TabIndex        =   98
      Top             =   1560
      Width           =   5835
      Begin VB.Label lblCopy 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   99
         Top             =   480
         Width           =   5355
      End
   End
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2400
      TabIndex        =   95
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdExitHlp 
         Caption         =   "Exit"
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdAbout 
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog comDialog3 
      Left            =   8400
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBarChart 
      Caption         =   "&Bar Chart"
      Height          =   615
      Left            =   4440
      TabIndex        =   48
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame PrintStuff 
      Caption         =   "Print"
      Height          =   1215
      Left            =   4440
      TabIndex        =   45
      Top             =   5160
      Width           =   1215
      Begin VB.CommandButton cmdCurrent 
         Caption         =   "&Current Report"
         Height          =   495
         Left            =   360
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Screen"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog comDialog2 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "FiberTrack.frx":08E9
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   7800
      TabIndex        =   38
      Text            =   " 7"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "FiberTrack.frx":08ED
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "FiberTrack.frx":08EF
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "FiberTrack.frx":08F1
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "FiberTrack.frx":08F5
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "FiberTrack.frx":08F9
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "FiberTrack.frx":08FD
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   6480
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5880
      TabIndex        =   27
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlotData 
      Caption         =   "Plot &Data"
      Height          =   615
      Left            =   2280
      TabIndex        =   26
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Timer tmrZeroCal 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   6600
   End
   Begin VB.CommandButton StartTImer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Timer"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton RunTimer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Timer"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmrIntegration 
      Left            =   8880
      Top             =   6600
   End
   Begin VB.OptionButton Led1 
      BackColor       =   &H80000001&
      Caption         =   "Updating"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton ZeroCal 
      BackColor       =   &H0080FF80&
      Caption         =   "Zero Calibration"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton GainCal 
      BackColor       =   &H0080FFFF&
      Caption         =   "Gain Calibration"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   2295
      Index           =   8
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "FiberTrack.frx":0901
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   7
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "FiberTrack.frx":090B
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   6
      Left            =   5880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "FiberTrack.frx":0915
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   5
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FiberTrack.frx":091F
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   4
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "FiberTrack.frx":0929
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "FiberTrack.frx":0933
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "FiberTrack.frx":093D
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000C&
      Caption         =   "Sensor 8"
      Height          =   375
      Index           =   8
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 7"
      Height          =   375
      Index           =   7
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 6"
      Height          =   375
      Index           =   6
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 5"
      Height          =   375
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 4"
      Height          =   375
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 3"
      Height          =   375
      Index           =   3
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000C&
      Caption         =   "Sensor 2"
      Height          =   375
      Index           =   2
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808080&
      Caption         =   "Sensor 1"
      Height          =   375
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      InputMode       =   1
   End
   Begin VB.Timer tmr2Sec 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   6600
   End
   Begin VB.CommandButton StartStop 
      BackColor       =   &H000000FF&
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Init 
      BackColor       =   &H00FFFF80&
      Caption         =   "Initialize"
      Height          =   375
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin CWUIControlsLib.CWGraph CWGraph1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   9495
      _Version        =   196608
      _ExtentX        =   16748
      _ExtentY        =   5106
      _StockProps     =   71
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   196608
      Graph_0         =   1
      ClassName_1     =   "CCWGraphFrame"
      opts_1          =   30
      Bindings_1      =   2
      ClassName_2     =   "CCWBindingHolderArray"
      Editor_2        =   3
      ClassName_3     =   "CCWBindingHolderArrayEditor"
      Owner_3         =   1
      C[0]_1          =   12632256
      C[1]_1          =   8421504
      C[2]_1          =   8421376
      Event_1         =   4
      ClassName_4     =   "CCWGFPlotEvent"
      Owner_4         =   1
      Plots_1         =   5
      ClassName_5     =   "CCWDataPlots"
      Array_5         =   1
      Editor_5        =   6
      ClassName_6     =   "CCWGFPlotArrayEditor"
      Owner_6         =   1
      Array[0]_5      =   7
      ClassName_7     =   "CCWDataPlot"
      opts_7          =   4194335
      Name_7          =   "Plot-1"
      Bindings_7      =   0
      C[0]_7          =   8388608
      C[1]_7          =   255
      C[2]_7          =   16711680
      C[3]_7          =   16776960
      Event_7         =   4
      X_7             =   8
      ClassName_8     =   "CCWAxis"
      opts_8          =   543
      Name_8          =   "Time"
      Bindings_8      =   0
      Orientation_8   =   2946
      format_8        =   9
      ClassName_9     =   "CCWFormat"
      Format_9        =   "."
      Scale_8         =   10
      ClassName_10    =   "CCWScale"
      opts_10         =   24576
      Bindings_10     =   0
      rMin_10         =   49
      rMax_10         =   621
      dMax_10         =   20
      discInterval_10 =   0.0017
      Radial_8        =   0
      Enum_8          =   11
      ClassName_11    =   "CCWEnum"
      Editor_11       =   12
      ClassName_12    =   "CCWEnumArrayEditor"
      Owner_12        =   8
      Font_8          =   13
      ClassName_13    =   "CCWFont"
      bFont_13        =   -1  'True
      BeginProperty Font_13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      tickopts_8      =   1679
      major_8         =   10
      minor_8         =   10
      Caption_8       =   14
      ClassName_14    =   "CCWDrawObj"
      opts_14         =   30
      Bindings_14     =   0
      C[0]_14         =   -2147483640
      Image_14        =   15
      ClassName_15    =   "CCWTextImage"
      Bindings_15     =   0
      szText_15       =   "Elapsed Time Window(Minutes)"
      font_15         =   0
      Animator_14     =   0
      Blinker_14      =   0
      Y_7             =   16
      ClassName_16    =   "CCWAxis"
      opts_16         =   543
      Name_16         =   "Denier"
      Bindings_16     =   0
      Orientation_16  =   2323
      format_16       =   17
      ClassName_17    =   "CCWFormat"
      Scale_16        =   18
      ClassName_18    =   "CCWScale"
      opts_18         =   57344
      Bindings_18     =   0
      rMin_18         =   29
      rMax_18         =   149
      dMax_18         =   200
      discInterval_18 =   1
      Radial_16       =   0
      Enum_16         =   19
      ClassName_19    =   "CCWEnum"
      Editor_19       =   20
      ClassName_20    =   "CCWEnumArrayEditor"
      Owner_20        =   16
      Font_16         =   21
      ClassName_21    =   "CCWFont"
      bFont_21        =   -1  'True
      BeginProperty Font_21 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      tickopts_16     =   1679
      major_16        =   10
      minor_16        =   5
      Caption_16      =   22
      ClassName_22    =   "CCWDrawObj"
      opts_22         =   30
      Bindings_22     =   0
      C[0]_22         =   -2147483640
      Image_22        =   23
      ClassName_23    =   "CCWTextImage"
      Bindings_23     =   0
      szText_23       =   "Denier"
      style_23        =   1
      font_23         =   0
      Animator_22     =   0
      Blinker_22      =   0
      LineStyle_7     =   1
      LineWidth_7     =   1
      BasePlot_7      =   0
      DefaultXInc_7   =   1
      DefaultPlotPerRow_7=   -1  'True
      Axes_1          =   24
      ClassName_24    =   "CCWAxes"
      Array_24        =   2
      Editor_24       =   25
      ClassName_25    =   "CCWGFAxisArrayEditor"
      Owner_25        =   1
      Array[0]_24     =   8
      Array[1]_24     =   16
      DefaultPlot_1   =   26
      ClassName_26    =   "CCWDataPlot"
      opts_26         =   4194335
      Name_26         =   "[Template]"
      Bindings_26     =   0
      C[0]_26         =   8421376
      C[1]_26         =   255
      C[2]_26         =   16711680
      C[3]_26         =   16776960
      Event_26        =   4
      X_26            =   8
      Y_26            =   16
      LineStyle_26    =   4
      LineWidth_26    =   1
      BasePlot_26     =   0
      DefaultXInc_26  =   1
      DefaultPlotPerRow_26=   -1  'True
      Cursors_1       =   27
      ClassName_27    =   "CCWCursors"
      Array_27        =   3
      Editor_27       =   28
      ClassName_28    =   "CCWGFCursorArrayEditor"
      Owner_28        =   1
      Array[0]_27     =   29
      ClassName_29    =   "CCWCursor"
      opts_29         =   31
      Name_29         =   "Target_Denier"
      Bindings_29     =   0
      C[0]_29         =   65280
      Event_29        =   4
      X_29            =   8
      Y_29            =   16
      XPos_29         =   1
      YPos_29         =   100
      PointIndex_29   =   -1
      ChrosshairStyle_29=   6
      LockPlot_29     =   0
      Array[1]_27     =   30
      ClassName_30    =   "CCWCursor"
      opts_30         =   31
      Name_30         =   "Plus_Tol"
      Bindings_30     =   0
      C[0]_30         =   255
      Event_30        =   4
      X_30            =   8
      Y_30            =   16
      XPos_30         =   24.8
      YPos_30         =   50
      PointIndex_30   =   -1
      ChrosshairStyle_30=   6
      LockPlot_30     =   0
      Array[2]_27     =   31
      ClassName_31    =   "CCWCursor"
      opts_31         =   31
      Name_31         =   "Minus_Tol"
      Bindings_31     =   0
      C[0]_31         =   255
      Event_31        =   4
      X_31            =   8
      Y_31            =   16
      XPos_31         =   36.7
      YPos_31         =   150
      PointIndex_31   =   -1
      ChrosshairStyle_31=   6
      LockPlot_31     =   0
      TrackMode_1     =   2
      GraphBackground_1=   0
      GraphFrame_1    =   32
      ClassName_32    =   "CCWDrawObj"
      opts_32         =   30
      Bindings_32     =   0
      C[0]_32         =   8421504
      C[1]_32         =   8421504
      Image_32        =   33
      ClassName_33    =   "CCWPictImage"
      opts_33         =   1280
      Bindings_33     =   0
      Rows_33         =   1
      Cols_33         =   1
      F_33            =   8421504
      B_33            =   8421504
      ColorReplaceWith_33=   8421504
      ColorReplace_33 =   8421504
      Tolerance_33    =   2
      Animator_32     =   0
      Blinker_32      =   0
      PlotFrame_1     =   34
      ClassName_34    =   "CCWDrawObj"
      opts_34         =   30
      Bindings_34     =   0
      C[0]_34         =   8421504
      C[1]_34         =   12632256
      Image_34        =   35
      ClassName_35    =   "CCWPictImage"
      opts_35         =   1280
      Bindings_35     =   0
      Rows_35         =   1
      Cols_35         =   1
      Pict_35         =   1
      F_35            =   8421504
      B_35            =   12632256
      ColorReplaceWith_35=   8421504
      ColorReplace_35 =   8421504
      Tolerance_35    =   2
      Animator_34     =   0
      Blinker_34      =   0
      Caption_1       =   36
      ClassName_36    =   "CCWDrawObj"
      opts_36         =   30
      Bindings_36     =   0
      C[0]_36         =   -2147483640
      Image_36        =   37
      ClassName_37    =   "CCWTextImage"
      Bindings_37     =   0
      szText_37       =   "Online plot of average Orientation"
      font_37         =   0
      Animator_36     =   0
      Blinker_36      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
   End
   Begin VB.PictureBox picSetup 
      BackColor       =   &H8000000C&
      Height          =   3015
      Left            =   5280
      ScaleHeight     =   2955
      ScaleWidth      =   2595
      TabIndex        =   52
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdCancel1 
         BackColor       =   &H8000000C&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdOk1 
         BackColor       =   &H8000000C&
         Caption         =   "OK"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdSetup 
         BackColor       =   &H8000000C&
         Caption         =   "&Com Port                           Alt+C"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblSetup 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Setup"
         Height          =   255
         Left            =   720
         TabIndex        =   53
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Index           =   1
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "FiberTrack.frx":0947
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblSetUpFile 
      Alignment       =   2  'Center
      Caption         =   "Setup  File Name"
      Height          =   255
      Left            =   3840
      TabIndex        =   49
      Top             =   5280
      Width           =   5295
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Threadline"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   44
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   7
      Left            =   6960
      TabIndex        =   43
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   42
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   41
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   40
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   32
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   31
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblThreadLine 
      Alignment       =   2  'Center
      Caption         =   "Threadline"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   29
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Elapsed Time"
      Height          =   255
      Left            =   8280
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Start Time"
      Height          =   255
      Left            =   7200
      TabIndex        =   24
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Settings..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Settings As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Settings..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenDataFile 
         Caption         =   "Open Data File..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Open Data File In &Excel..."
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu Restart 
         Caption         =   "&Restart"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNormal 
         Caption         =   "&Normal"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuExpanded 
         Caption         =   "&Expanded"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup..."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuMain 
         Caption         =   "&Main"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPlotDataFile 
         Caption         =   "&Plot Data File"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBarGraph 
         Caption         =   "&Bar Graph"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About STC FiberTrack..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "FiberTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------
'--
'-- File name    :  FiberTrack
'-- Title        :  Smart denier sensor host code.
'-- Library      :  WORK
'--              :
'-- Purpose      :
'--              :
'-- Created On   : 11/11/97 20:49:20
'-- Last Updated : 02/29/2014
'-- Comments     :
'--              :
'-- Assumptions  : none
'-- Limitations  : none
'-- Known Errors : none
'-- Developers   : Jim West
'--              : Jason Sweatt
'-- Notes        :
'-- --------------------------------------------------------------------
'-- >>>>>>>>>>>>>>>>>>>>>>> COPYRIGHT NOTICE <<<<<<<<<<<<<<<<<<<<<<<<<<<
'-- ---------------------------------------------------------------------
'-- Copyright 1995-2014 (c) Scientific Technologies Inc.
'--
'-- YourCompanyName$ owns the sole copyright to this software. Under
'-- international copyright laws you (1) may not make a copy of this software
'-- except for the purposes of maintaining a single archive copy, (2) may not
'-- derive works herefrom, (3) may not distribute this work to others. These
'-- rights are provided for information clarification, other restrictions of
'-- rights may apply as well.
'--
'-- This is an unpublished work.
'-- ----------------------------------------------------------------------
'-- >>>>>>>>>>>>>>>>>>>>>>>>>>>>> Warrantee <<<<<<<<<<<<<<<<<<<<<<<<<<<<
'-- ----------------------------------------------------------------------
'-- SCIENTIFIC TECHNOLOGIES MAKES NO WARRANTY OF ANY KIND WITH REGARD TO THE
'-- USE OF THIS SOFTWARE, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT
'-- LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY OR FITNESS FOR A
'-- PARTICULAR PURPOSE.
'-- ----------------------------------------------------------------------
'--   Version No:| Author    | Changes Made:           | Mod. Date:
'--     v1.5     | Jim West  | Automatically Generated | 11/11/97
'-- ----------------------------------------------------------------------
'-- Revision History :
'-- Rev 1.4 J. West 4/23/98 - Added incrementing data file names and made
'-- several other minor enhancements.
'-- Rev 1.5 J. West 5/13/98 - Added an input window for entering parameters.
'-- Rev 1.6 J. West 8/11/98 - Made cv a compile option. Fixed cv problems.
'-- Rev 1.7 J. West 9/8/98  - Incorporated separate calibration denier setting.
'-- Rev 2.5 Jason Sweatt 02/16/2012  - Renamed Denier to Orientation within the application
'-- Rev 2.5 Jason Sweatt 02/25/2014  - Fixing runtime '6' error
'-- ----------------------------------------------------------------------

Dim RealTemp                        As Single

#If SCALE_DIAMETER Then
Dim IntTemp                         As Single
Dim iTempMinDen                     As Single
Dim iTempMaxDen                     As Single
#Else
Dim IntTemp                         As Integer
Dim iTempMinDen                     As Integer
Dim iTempMaxDen                     As Integer
#End If

Dim iPLotChannels(1 To LAST_SENSOR) As Integer
Dim iPCCnt                          As Integer
Dim iPlotDenier                     As Variant
' Dim iPLotDenier(1 To 4, 0 To 0) As Integer

Dim iPLot                           As Integer
Dim iMaxPlotChannels                As Integer
Dim iTempOptDenRangeIndex           As Integer
Dim iOptDenRangeIndex               As Integer
Dim iTempLevel1Slub                 As Long
Dim iTempLevel1Tol                  As Long
Dim iTempLevel2Slub                 As Long
Dim iTempLevel2Tol                  As Long

Dim mo_Comm                         As FTComm

Const YAXIS                         As Integer = 2
Const XAXIS                         As Integer = 1
#If SCALE_DIAMETER Then
Const SETUP_FILENAME                As String = "setupd.svd"
#Else
Const SETUP_FILENAME                As String = "setup.svd"
#End If

Const SENSOR_HIGHEST_DEFAULT = -32000
Const SENSOR_LOWEST_DEFAULT = 32000

'Click About in Help frame
Private Sub cmdAbout_Click()
    fraAbout.Visible = True
End Sub

'cmdBarChart is hidden from users and is used for debugging purposes
Private Sub cmdBarChart_Click()
    BarChart.Visible = True
End Sub

'Cancel Application Parameters setup screen
Private Sub cmdCanAP_Click()
    txtCurValueAp(0).Text = LineSpeed
    txtCurValueAp(1).Text = Integration_time
    txtCurValueAp(2).Text = ZeroCal_Interval / 60
    
    fraAppPar.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

Private Sub cmdCancelIniSettings_Click()
    fraApplicationIni.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

Private Sub cmdSaveIniSettings_Click()
    MsgBox "For some of these application ini settings to take into affect, please exit and re-open the application.", vbOKOnly, "Warning"
    Dim setValue As String
    If chkOpenApplicationMaxIniSettings = vbChecked Then
        setValue = "True"
    Else
        setValue = "False"
    End If
    If PutIniSetting("Application", "OpenMainFormMax", setValue) Then
    
    End If
    If chkIncludeCvIniSettings = vbChecked Then
        setValue = "True"
    Else
        setValue = "False"
    End If
    
    If PutIniSetting("Application", "IncludeCvColumn", setValue) Then
    End If
    
    If PutIniSetting("Application", "ExcelPath", txtExcelPathIniSettings.Text) Then
    End If

    If PutIniSetting("Application", "FileDateFormatAppend", txtAppendDateFormatToFile.Text) Then
    End If
    
    If PutIniSetting("Application", "NameOfThreadLine", txtNameOfThreadLine.Text) Then
    End If
    
    fraApplicationIni.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

'Cancel COM Port setup screen
Private Sub cmdCanCom_Click()
    fraComPort.Visible = False
    
    If SystemStatus.ComPort = 1 Then
        optCom1.value = True
    Else
        optCom2.value = True
    End If
   
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

'Cancel Defect Detection setup screen
Private Sub cmdCanDD_Click()
    txtCurValueDD(0).Text = Level1_slub_tol & " %"
    txtCurValueDD(1).Text = Level1_length & " mm"
    txtCurValueDD(2).Text = Level2_slub_tol & " %"
    txtCurValueDD(3).Text = Level2_length & " mm"
    
    fraDefDet.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

'Cancel Denier Parameter setup screen
Private Sub cmdCanDP_Click()
    fraDenierPar.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
    
    txtDenPar(1).Text = Target_Denier
    txtDenPar(2).Text = Calibration_Denier
    txtDenPar(3).Text = Target_denier_tol
End Sub

'Cancel Denier Range setup screen
Private Sub cmdCanDR_Click()
    fraDenRange.Visible = False
    cmdDoneDP.Visible = True
    cmdDoneDP.Enabled = True
#If SCALE_DIAMETER Then
    txtDenPar(0).Text = FormatDiameter(iMin_Denier) & " to " & FormatDiameter(Max_Denier)
#Else
    txtDenPar(0).Text = iMin_Denier & " to " & Max_Denier
#End If
    optDenRange(iOptDenRangeIndex).value = True
    cmdDenPar(1).Enabled = True
    cmdDenPar(2).Enabled = True
    cmdDenPar(3).Enabled = True
End Sub

'Cancel Plot Channels setup screen
Private Sub cmdCanPC_Click()
    Dim index As Integer
    For index = 0 To 7
      If iPLotChannels(index + 1) = 1 Then
        chkPlotChan(index).value = 1
      Else
        chkPlotChan(index).value = 0
      End If
    Next index
    fraPlotChan.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

Private Sub cmdComPort_Click(index As Integer)
    Dim i As Integer
    Dim response As Integer
    Dim messageText As String
    Dim style As Integer
    Dim currentTitle As String
On Error GoTo cmdComPort_err_hdr
    fraSetup.Enabled = False
    Select Case index
        Case 0          ' Com Port
            fraComPort.Visible = True
            If SystemStatus.ComPort = 1 Then
                optCom1.value = True
            Else
                optCom2.value = True
            End If
            cmdDone.Enabled = False
        Case 1          ' Denier Parameters
            fraDenierPar.Visible = True
            cmdOkDP.Visible = False
            cmdCanDP.Visible = False
            cmdDone.Enabled = False
            cmdDoneDP.Visible = True
            cmdDoneDP.Enabled = True
            optDenRange(iOptDenRangeIndex).value = True
            txtDenPar(0).Enabled = False
            txtDenPar(1).Enabled = False
            txtDenPar(2).Enabled = False
            txtDenPar(3).Enabled = False
        Case 2          'Plot Channels
            fraPlotChan.Visible = True
            cmdDone.Enabled = False
        Case 3          'Application Parameters
            fraAppPar.Visible = True
            cmdDone.Enabled = False
        Case 4          'Defect Detection
            fraDefDet.Visible = True
            cmdDone.Enabled = False
        Case 5 'Applciation Ini Settings
            If GetIniSetting("Application", "OpenMainFormMax") = "True" Then
                chkOpenApplicationMaxIniSettings.value = 1
            Else
                chkOpenApplicationMaxIniSettings.value = 0
            End If
            If GetIniSetting("Application", "IncludeCvColumn") = "True" Then
                chkIncludeCvIniSettings.value = 1
            Else
                chkIncludeCvIniSettings.value = 0
            End If
            txtExcelPathIniSettings.Text = GetIniSetting("Application", "ExcelPath")
            txtAppendDateFormatToFile.Text = GetIniSetting("Application", "FileDateFormatAppend")
            txtNameOfThreadLine.Text = GetIniSetting("Application", "NameOfThreadLine")
            fraApplicationIni.Visible = True
            cmdDone.Enabled = False
        Case 6          'Save Settings
            mnuSave_Click
            fraSetup.Enabled = True
            cmdDone.Enabled = True
        Case 7          'Save Settings as Default
            messageText = "If you save these settings as default," & vbCrLf & "they will be used at startup in all" & vbCrLf & "future sessions."
            style = vbOKCancel Or vbDefaultButton2 Or vbExclamation
            currentTitle = "WARNING!"
            response = MsgBox(messageText, style, currentTitle)
            If response = vbOK Then
                SaveSetupFile App.Path & "\" & SETUP_FILENAME
            End If
            fraSetup.Enabled = True
            cmdDone.Enabled = True
        Case Else       'Error exit
            SetStatusText "Other button pressed"
    End Select
    Exit Sub
cmdComPort_err_hdr:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Save Settings File Error, cmdComPort"
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

Private Sub cmdCurrent_Click()
On Error GoTo Current_Err_Hdr
    Dim sensorIndex As Integer
    Dim iTemp As Integer
    Dim iTempP1 As Integer
    Dim response As Integer
    Dim localFileNamesNumber As Integer
    comDialog3.CancelError = True
    localFileNamesNumber = FreeFile
    comDialog3.DialogTitle = "Save Current Report"
    comDialog3.Filter = "Text File (*.txt)| *.txt"
    comDialog3.FilterIndex = 1
    comDialog3.Flags = cdlOFNExplorer Or cdlOFNExtensionDifferent Or cdlOFNNoChangeDir Or cdlOFNHideReadOnly
    comDialog3.ShowOpen
    Open comDialog2.FileName For Output As #localFileNamesNumber
  
    Write #FileNamesNumber, "           STC Denier Monitoring Current Report  - "; Now
    Write #FileNamesNumber, "           ---------------------------------------------------------"
    Write #FileNamesNumber,
    Write #FileNamesNumber, "Integration Time ="; Integration_time; "Sec.   "; "Target Denier ="; Target_Denier;
    Write #FileNamesNumber, "   Line Speed ="; LineSpeed; "Meters/sec."
    Write #FileNamesNumber,
    
    For sensorIndex = 1 To LAST_SENSOR
        If SensorInfos(sensorIndex).Enabled = True Then
            ' Print Current Information
            Write #FileNamesNumber, "Sensor "; sensorIndex; " "; GetIniSetting("Application", "NameOfThreadLine"); " "; SensorInfos(sensorIndex).Package
            Write #FileNamesNumber, Tab(10); "Current Value = "; Format(CurrentAverage(sensorIndex), "#00.0")
            Write #FileNamesNumber, Tab(10); "Maximum Value = "; SensorInfos(sensorIndex).Highest
            Write #FileNamesNumber, Tab(10); "Mean Value = "; Format(CurrentMean(sensorIndex), "#00.0")
            Write #FileNamesNumber, Tab(10); "Minimum Value = "; SensorInfos(sensorIndex).Lowest
            Write #FileNamesNumber, Tab(10); "CV Value = "; Format(CurrentCv(sensorIndex), "#00.0")
            Write #FileNamesNumber, Tab(10); "Level 1 Defects = "; SensorInfos(sensorIndex).Level1_Slub
            Write #FileNamesNumber, Tab(10); "Level 2 Defects = "; SensorInfos(sensorIndex).Level2_Slub
            ' Write the diferences for crossover
            If sensorIndex + 1 <= LAST_SENSOR Then
                Write #FileNamesNumber, Tab(10); "Delta a-b = "; CurrentAverage(sensorIndex) - CurrentAverage(sensorIndex + 1)
            End If
        End If
    Next sensorIndex
  
    Close #localFileNamesNumber
    Exit Sub
Current_Err_Hdr:
    MsgBox "There was a problem printing to your printer.", vbOKOnly, "FiberTrack cmdCurrent"
    Exit Sub
End Sub

Private Sub cmdDenPar_Click(index As Integer)
    Select Case index
        Case 0
            fraDenRange.Visible = True
            cmdOkDP.Visible = False
            cmdOkDP.Enabled = False
            
            cmdCanDP.Visible = False
            cmdCanDP.Enabled = False
            
            cmdDoneDP.Visible = True
            cmdDoneDP.Enabled = False
            cmdDenPar(1).Enabled = False
            cmdDenPar(2).Enabled = False
            cmdDenPar(3).Enabled = False
            
            txtDenPar(1).Enabled = False
            txtDenPar(2).Enabled = False
            txtDenPar(3).Enabled = False
        Case 1
            txtDenPar(index).Enabled = True
            txtDenPar(index).SetFocus
            cmdOkDP.Visible = True
            cmdOkDP.Enabled = True
            
            cmdCanDP.Visible = True
            cmdCanDP.Enabled = True
            
            cmdDoneDP.Enabled = False
            cmdDoneDP.Visible = False
            
            txtDenPar(0).Enabled = False
            txtDenPar(2).Enabled = False
            txtDenPar(3).Enabled = False
        Case 2
            txtDenPar(index).Enabled = True
            txtDenPar(index).SetFocus
            cmdOkDP.Visible = True
            cmdOkDP.Enabled = True
            
            cmdCanDP.Visible = True
            cmdCanDP.Enabled = True
            
            cmdDoneDP.Enabled = False
            cmdDoneDP.Visible = False
            
            txtDenPar(0).Enabled = False
            txtDenPar(1).Enabled = False
            txtDenPar(3).Enabled = False
        Case 3
            txtDenPar(index).Enabled = True
            txtDenPar(index).SetFocus
            cmdOkDP.Visible = True
            cmdOkDP.Enabled = True
            
            cmdCanDP.Visible = True
            cmdCanDP.Enabled = True
            
            cmdDoneDP.Enabled = False
            cmdDoneDP.Visible = False
            
            txtDenPar(0).Enabled = False
            txtDenPar(1).Enabled = False
            txtDenPar(2).Enabled = False
        Case Else
          SetStatusText "Other button pressed"
    End Select
End Sub

Private Sub cmdDone_Click()
    fraSetup.Visible = False
End Sub

Private Sub cmdDoneDP_Click()
    fraDenierPar.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

'Hidden button
Private Sub cmdExit_Click()
    Unload Me
End Sub

'Exit from Help Frame
Private Sub cmdExitHlp_Click()
    fraAbout.Visible = False
    fraHelp.Visible = False
End Sub

Private Sub cmdLineSpeed_Click(index As Integer)
    txtCurValueAp(index).SetFocus
End Sub

'Click OK in Application Parameters setup screen
Private Sub cmdOkAP_Click()
    Dim iAPErr As Integer
    iAPErr = 0
    If Val(txtCurValueAp(0).Text) <= 6000 And Val(txtCurValueAp(0).Text) >= 1 Then
        LineSpeed = Val(txtCurValueAp(0).Text)
    Else
        txtCurValueAp(0).Text = LineSpeed
        iAPErr = iAPErr + 1
    End If
    
    If Val(txtCurValueAp(1).Text) <= 30 And Val(txtCurValueAp(1).Text) >= 1 Then
        Integration_time = Val(txtCurValueAp(1).Text)
    Else
        txtCurValueAp(1).Text = Integration_time
        iAPErr = iAPErr + 1
    End If
    
    If Val(txtCurValueAp(2).Text) <= 30 And Val(txtCurValueAp(2).Text) >= 1 Then
        ZeroCal_Interval = Val(txtCurValueAp(2).Text) * 60
    Else
        txtCurValueAp(2).Text = ZeroCal_Interval / 60
        iAPErr = iAPErr + 1
    End If
        
    If iAPErr = 0 Then
        fraAppPar.Visible = False
        fraSetup.Enabled = True
        cmdDone.Enabled = True
    End If
End Sub

'Click OK in COM Port setup screen
Private Sub cmdOKCom_Click()
    SystemStatus.ComPort = TempComPort
    fraComPort.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
End Sub

'Click OK in Defect Detection setup screen
Private Sub cmdOkDD_Click()
    Dim iDDErr As Integer
    
    iDDErr = 0
    
    If Val(txtCurValueDD(0).Text) <= 99 And Val(txtCurValueDD(0).Text) >= 3 Then
        Level1_slub_tol = Val(txtCurValueDD(0).Text)
    Else
        txtCurValueDD(0).Text = Level1_slub_tol & " %"
        iDDErr = iDDErr + 1
    End If
    
    If Val(txtCurValueDD(1).Text) <= 100 And Val(txtCurValueDD(1).Text) >= 1 Then
        Level1_length = Val(txtCurValueDD(1).Text)
    Else
        txtCurValueDD(1).Text = Level1_length & " mm"
        iDDErr = iDDErr + 1
    End If
    
    If Val(txtCurValueDD(2).Text) <= 99 And Val(txtCurValueDD(2).Text) >= 3 Then
        Level2_slub_tol = Val(txtCurValueDD(2).Text)
    Else
        txtCurValueDD(2).Text = Level2_slub_tol & " %"
        iDDErr = iDDErr + 1
    End If
    
    If Val(txtCurValueDD(3).Text) <= 100 And Val(txtCurValueDD(3).Text) >= 1 Then
        Level2_length = Val(txtCurValueDD(3).Text)
    Else
        txtCurValueDD(3).Text = Level2_length & " mm"
        iDDErr = iDDErr + 1
    End If
    
    If iDDErr = 0 Then
        fraDefDet.Visible = False
        fraSetup.Enabled = True
        cmdDone.Enabled = True
    End If
End Sub

'Click OK in Denier Parameters setup screen
' 02/15/2012 Changed: Denier to Orientation Range Click OK setup screen
Private Sub cmdOkDP_Click()
    Dim iDPErr As Integer
    
    iDPErr = 0
    
    If Val(txtDenPar(1).Text) <= Max_Denier And Val(txtDenPar(1).Text) >= iMin_Denier Then
        Target_Denier = Val(txtDenPar(1).Text)
    Else
        txtDenPar(1).Text = Target_Denier
        iDPErr = iDPErr + 1
    End If
    
    If Val(txtDenPar(2).Text) <= Max_Denier And Val(txtDenPar(2).Text) >= iMin_Denier Then
        Calibration_Denier = Val(txtDenPar(2).Text)
    Else
        txtDenPar(2).Text = Calibration_Denier
        iDPErr = iDPErr + 1
    End If
    
    If Val(txtDenPar(3).Text) <= 20 And Val(txtDenPar(3).Text) >= 0 Then
        Target_denier_tol = Val(txtDenPar(3).Text)
    Else
        txtDenPar(3).Text = Target_denier_tol & " %"
        iDPErr = iDPErr + 1
    End If
    
    If iDPErr = 0 Then
        CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier  'Initialize plot setup
        mnuNormal.Checked = Not mnuNormal.Checked
        mnuExpanded.Checked = Not mnuExpanded.Checked
        IsNormal = True
        PlotViewIndex = 0
        
        CWGraph1.Cursors.Item(1).YPosition = Target_Denier
        CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
        CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)

        fraDenierPar.Visible = False
        fraSetup.Enabled = True
        cmdDone.Enabled = True
    End If
End Sub

'Click OK in Denier Range setup
' 02/15/2012 Changed: Denier to Orientation Range Click OK setup
Private Sub cmdOkDR_Click()
    fraDenRange.Visible = False
    cmdDoneDP.Visible = True
    cmdDoneDP.Enabled = True
    iMin_Denier = iTempMinDen
    Max_Denier = iTempMaxDen
    iOptDenRangeIndex = iTempOptDenRangeIndex
    DenierRangeIndex = iTempOptDenRangeIndex
    
    CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier  'Initialize plot setup
    CWGraph1.Cursors.Item(1).YPosition = Target_Denier
    CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
    
    cmdDenPar(1).Enabled = True
    cmdDenPar(2).Enabled = True
    cmdDenPar(3).Enabled = True
End Sub

'Click OK in Plot Channels setup
Private Sub cmdOkPC_Click()
On Error GoTo cmdOkPC_err_hdr
    Dim iPC As Integer
    iPCCnt = 0
    For iPC = 0 To 7
        iPLotChannels(iPC + 1) = 0
    Next iPC
    For iPC = 0 To 7
        If chkPlotChan(iPC).value = 1 Then
            iPLotChannels(iPC + 1) = 1
            iPCCnt = iPCCnt + 1
        End If
    Next iPC
    If iPCCnt = 0 Then
        MsgBox "Please select the channels to be plotted!", vbOKOnly Or vbExclamation, "cmdOkPC"
        Exit Sub
    End If
    If iPCCnt > iMaxPlotChannels Then
        MsgBox "The number of channels checked exceeds the maximum!", vbOKOnly, "cmdOkPC"
        For iPC = 0 To 7
            iPLotChannels(iPC + 1) = 0
        Next iPC
        Exit Sub
    End If
    
    ReDim iPlotDenier(1 To iPCCnt, 0 To 0)
    fraPlotChan.Visible = False
    fraSetup.Enabled = True
    cmdDone.Enabled = True
    Exit Sub
cmdOkPC_err_hdr:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Error, cmdOkPC FiberTrack"
End Sub

Private Sub cmdParDD_Click(index As Integer)
    txtCurValueDD(index).SetFocus
End Sub

'cmdPlotData is hidden from users and is used for debugging purposes
Private Sub cmdPlotData_Click()
    FiberTrack.Hide
    PlotData.Show       ' vbModal
    PlotData.cboomPlotData.ListIndex = 0
    PlotData.cmdReturn.SetFocus
End Sub

'cmdPrint is hidden from users and is used for debugging purposes
Private Sub cmdPrint_Click()
    FiberTrack.PrintForm
End Sub

'Click on "Sensor n" button at top of sensor column.
Private Sub Command_Click(index As Integer)
    'Only allow these buttons before Start button is pressed.
    Dim i As Integer
    If SystemStatus.Enabled = False Then
        SetStatusText "Set the operating parameters such as Integration time and Line speed, then click on the Initialize button."
        Init.Enabled = True                         'Enable initialization button
        Init.Visible = True
        For i = 1 To LAST_SENSOR
            If i = index Then
                If SensorInfos(i).Enabled = False Then
                    SensorInfos(i).Enabled = True
                    Command(i).BackColor = &H80FF80
                    Text1(i).Text = "Enabled"
                Else
                    SensorInfos(i).Enabled = False
                    Command(i).BackColor = &H808080
                    Text1(i).Text = "Disabled"
                End If
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    MsgBox "Fibertrack Key = " & KeyAscii, , "Fibertrack_KeyPress"
End Sub

Private Sub Form_Load()                 'Main initialization routine.
    Dim i As Integer
    Dim iY As Integer
    Dim iI As Integer
    Const iMax As Integer = 5
    Dim message(1 To iMax) As String
    Dim iInit As Integer
    Dim iLast As Integer
    Dim iIndex As Integer
    Dim setupCaption(0 To 7) As String
    Dim sDenParCaption(0 To 3) As String
    Dim sDenParValue(0 To 3) As String
    Dim sDenRge(0 To 3) As String
    Dim sAppPar(0 To 2) As String
    Dim sCurValueAP(0 To 2) As String
    Dim sCurValLblAP(0 To 2) As String
    Dim sDefDet(0 To 3) As String
    Dim sCurValueDD(0 To 3) As String
    Dim copy(0 To 3) As String
    Dim add(0 To 3) As String
    Dim telephone(0 To 3) As String
    Dim warning(0 To 3) As String
    Dim logo(0 To 3) As String
    
On Error GoTo fibertrack_err_rtn
      
#If SCALE_DIAMETER Then
    fraDenRange.Caption = "Diameter Range"
    fraDenRange.Width = fraDenRange.Width + 1000
    fraDenRange.Left = fraDenRange.Left + 480
    optDenRange(0).Width = optDenRange(0).Width + 600
    
    cmdCanDR.Left = cmdCanDR.Left + 600
    cmdOkDR.Left = cmdOkDR.Left + 600
    
    fraDenierPar.Caption = "Diameter Parameters"
    fraDenierPar.Width = fraDenierPar.Width + 600
    txtDenPar(0).Left = txtDenPar(0).Left + 240
    txtDenPar(0).Width = txtDenPar(0).Width + 240
    cmdDenPar(0).Width = cmdDenPar(0).Width + 120
    lblParValue.Left = lblParValue.Left + 360
    
    fraSetup.Width = fraSetup.Width + 100
    cmdComPort(0).Width = cmdComPort(0).Width + 100
    
    CWGraph1.Axes(YAXIS).DiscreteInterval = 0.001
    CWGraph1.Axes(YAXIS).Maximum = 0.5
#End If
    
    Set mo_Comm = New FTComm
    Set mo_Comm.MSComm = Me.MSComm1
    
    Init.Enabled = False
    StartStop.Enabled = False
    ZeroCal.Enabled = False
    GainCal.Enabled = False
    
    With Me
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        
        If .Left < 0 Then
            .Left = 0
        End If
        
        If .Top < 0 Then
            .Top = 0
        End If
        
        If .Height > Screen.Height Then
            .Height = Screen.Height
        End If
        
        If .Width > Screen.Width Then
            .Width = Screen.Width
        End If
        If CBool(GetIniSetting("Application", "OpenMainFormMax")) Then
            .WindowState = vbMaximized
        Else
            .WindowState = vbNormal
        End If
    End With
    
    fraSetup.Visible = False
    picSetup.Visible = False
    fraComPort.Visible = False
    fraDenierPar.Visible = False
    fraDenRange.Visible = False
    fraPlotChan.Visible = False
    fraAppPar.Visible = False
    fraDefDet.Visible = False
    fraApplicationIni.Visible = False
    fraHelp.Visible = False
    fraAbout.Visible = False
    
    IsMain = True
    IsNormal = True
    iPLot = 1
    
    cmdAbout.Caption = StringFormat("&About {0} Alt+A", GetIniSetting("Constants", "Name"))
    iMaxPlotChannels = 4
    
    ' 02/15/2012 Changed: Denier to Orientation Range Click OK setup screen
    setupCaption(0) = "&Com Port Alt+C"
#If SCALE_DIAMETER Then
    setupCaption(1) = "&Diameter Parameters Alt+D"
#Else
    setupCaption(1) = "Denier Range Alt+O"
#End If
    setupCaption(2) = "&Plot Channels Alt+P"
    setupCaption(3) = "&Application Parameters Alt+A"
    setupCaption(4) = "De&fect Detection Alt+F"
    setupCaption(5) = "Application &Ini Settings Alt+I"
    setupCaption(6) = "&Save Settings Alt+S"
    setupCaption(7) = "Sa&ve settings as Default Alt+V"
    
    copy(0) = StringFormat("{0} - {1} - {2}{3}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"), vbCrLf)
    copy(1) = StringFormat("Copyright © 2004-{0} Sensatus Technologies Corporated.{1}", Year(Now), vbCrLf)
    copy(2) = StringFormat("All rights reserved.{0}", vbCrLf)
    copy(3) = Date
    logo(0) = copy(0) & copy(1) & copy(2) & copy(3)
    
    add(0) = "Sensatus Technologies Corporation" & vbCrLf
    add(1) = "450 Edgell Road" & vbCrLf
    add(2) = "Framingham, MA 01701" & vbCrLf
    add(3) = "USA"
    logo(1) = add(0) & add(1) & add(2) & add(3)
    
    telephone(0) = "Telephone: 781-555-0000" & vbCrLf
    ' telephone(2) = "Fax: 510-744-1442" & vbCrLf
    ' telephone(3) = "Web site: www.stifibermonitoring.com" & vbCrLf
    telephone(1) = "E-mail: jpiso@aol.com"
    logo(2) = telephone(0) & telephone(1) ' & telephone(2) & telephone(3)
    
    warning(0) = "Warning!  This computer program is protected by copyright" & vbCrLf
    warning(1) = "law and international treaties.  Unauthorized reproduction or" & vbCrLf
    warning(2) = "distribution of this program is illegal and may result in" & vbCrLf
    warning(3) = "criminal and civil penalties."
    logo(3) = warning(0) & warning(1) & warning(2) & warning(3)
    
    lblPara.Font.Underline = True
    lblPara.Caption = "Parameter"
    lblParValue.Font.Underline = True
    lblParValue.Caption = "Current Value"
    
    cmdComPort(0).Caption = setupCaption(0)
    iInit = 1
    iLast = 7
        
    For iIndex = iInit To iLast
        Load cmdComPort(iIndex)
        cmdComPort(iIndex).Top = cmdComPort(iIndex - 1).Top + cmdComPort(iIndex - 1).Height
        cmdComPort(iIndex).Visible = True
        cmdComPort(iIndex).Caption = setupCaption(iIndex)
    Next iIndex
    
    iInit = 1
    iLast = 7
    cmdSetup(0).Caption = setupCaption(0)
    
    For iIndex = iInit To iLast
        Load cmdSetup(iIndex)
        cmdSetup(iIndex).Top = cmdSetup(iIndex - 1).Top + cmdSetup(iIndex - 1).Height
        cmdSetup(iIndex).Visible = True
        cmdSetup(iIndex).Caption = setupCaption(iIndex)
    Next iIndex
        
#If SCALE_DIAMETER Then
    sDenParCaption(0) = "Diameter &Range"
    sDenParCaption(1) = "Target &Diameter"
    sDenParCaption(2) = "Target &Gain"
    sDenParCaption(3) = "Diameter &Tolerance"

    iMin_Denier = 0
    Max_Denier = 0.5
    Target_Denier = 0.25
    Calibration_Denier = 0.25
    Target_denier_tol = 20
#Else
    sDenParCaption(0) = "Denier &Range"
    sDenParCaption(1) = "&Target Denier"
    sDenParCaption(2) = "Target &Gain"
    sDenParCaption(3) = "Denier T&olerance"

    iMin_Denier = 0
    Max_Denier = 1#
    Target_Denier = 100
    Calibration_Denier = 100
    Target_denier_tol = 5
#End If

    iOptDenRangeIndex = 0
    
    sDenParValue(0) = iMin_Denier & " to " & Max_Denier
    sDenParValue(1) = Target_Denier
    sDenParValue(2) = Calibration_Denier
    sDenParValue(3) = Target_denier_tol & " %"
        
    Font.Underline = True
    cmdDenPar(0).Caption = sDenParCaption(0)
    txtDenPar(0).Text = sDenParValue(0)
        
    iInit = 1
    iLast = 3
    
    For iIndex = iInit To iLast
        Load txtDenPar(iIndex)
        txtDenPar(iIndex).Top = txtDenPar(iIndex - 1).Top + txtDenPar(iIndex - 1).Height
        txtDenPar(iIndex).Visible = True
        txtDenPar(iIndex).Text = sDenParValue(iIndex)
        
        Load cmdDenPar(iIndex)
        cmdDenPar(iIndex).Top = cmdDenPar(iIndex - 1).Top + 40 + cmdDenPar(iIndex - 1).Height
        cmdDenPar(iIndex).Visible = True
        cmdDenPar(iIndex).Caption = sDenParCaption(iIndex)
    Next iIndex
    
    'sDenRge(0) = "  50 to   500 Denier"
    'sDenRge(1) = "100 to 1000 Denier"
    'sDenRge(2) = "150 to 1500 Denier"
    'sDenRge(3) = "200 to 2000 Denier"
    'CLM1112
    
#If SCALE_DIAMETER Then
    sDenRge(0) = "0.000 to 0.500  Diameter"
    sDenRge(1) = "0.100 to 1.000  Diameter"
    iInit = 1
    iLast = 1
#Else
    'sDenRge(0) = "    0 to   500 Denier"
    'sDenRge(1) = "100 to 1000 Denier"
    'sDenRge(2) = "200 to 2000 Denier"
    'sDenRge(3) = "500 to 5000 Denier"
    sDenRge(0) = "  0 to 1.00      Denier"
    sDenRge(1) = "  0 to 100       Centi-Newtons"
    sDenRge(2) = "100 to 1000     Centi-Newtons"
    sDenRge(3) = "500 to 5000     Denier"
    
    iInit = 1
    iLast = 3
#End If
    
    optDenRange(0).Caption = sDenRge(0)
    
    For iIndex = iInit To iLast
        Load optDenRange(iIndex)
        optDenRange(iIndex).Top = optDenRange(iIndex - 1).Top + optDenRange(iIndex - 1).Height
        optDenRange(iIndex).Visible = True
        optDenRange(iIndex).Caption = sDenRge(iIndex)
    Next iIndex
    
    iInit = 1
    iLast = 7
    chkPlotChan(0).Caption = "Channel &1"
    
    For iIndex = iInit To iLast
        Load chkPlotChan(iIndex)
        chkPlotChan(iIndex).Top = chkPlotChan(iIndex - 1).Top + chkPlotChan(iIndex - 1).Height
        chkPlotChan(iIndex).Visible = True
        chkPlotChan(iIndex).Caption = "Channel " & "&" & iIndex + 1
    Next iIndex
    
    sAppPar(0) = "Line &Speed               "
    sAppPar(1) = "Integration &Time         "
    sAppPar(2) = "&Zero Calibration Interval"
    
    LineSpeed = 250                        'Use these system settings for now
    Integration_time = 2                    'until the input routines are ready
    ZeroCal_Interval = 15 * 60              'Initialize to 15 minute intervals
    
    sCurValueAP(0) = LineSpeed
    sCurValueAP(1) = Integration_time
    sCurValueAP(2) = ZeroCal_Interval / 60
    sCurValLblAP(0) = "m/min"
    sCurValLblAP(1) = "sec"
    sCurValLblAP(2) = "min"
        
    iInit = 1
    iLast = 2
    txtCurValueAp(0).Text = sCurValueAP(0)
    cmdLineSpeed(0).Caption = sAppPar(0)
    lblLineSpeed(0).Caption = sCurValLblAP(0)
    lblLineSpeed(0).Alignment = 0
    For iIndex = iInit To iLast
        Load txtCurValueAp(iIndex)
        txtCurValueAp(iIndex).Top = txtCurValueAp(iIndex - 1).Top + txtCurValueAp(iIndex - 1).Height
        txtCurValueAp(iIndex).Visible = True
        txtCurValueAp(iIndex).Text = sCurValueAP(iIndex)
        
        Load cmdLineSpeed(iIndex)
        cmdLineSpeed(iIndex).Top = cmdLineSpeed(iIndex - 1).Top + 40 + cmdLineSpeed(iIndex - 1).Height
        cmdLineSpeed(iIndex).Visible = True
        cmdLineSpeed(iIndex).Caption = sAppPar(iIndex)
        
        Load lblLineSpeed(iIndex)
        lblLineSpeed(iIndex).Top = lblLineSpeed(iIndex - 1).Top + 40 + lblLineSpeed(iIndex - 1).Height
        lblLineSpeed(iIndex).Visible = True
        lblLineSpeed(iIndex).Caption = sCurValLblAP(iIndex)
    Next iIndex
    
    cmdLineSpeed(0).Width = 1865
    cmdLineSpeed(1).Width = 1865
    cmdLineSpeed(2).Width = 1865
    
    iInit = 1
    iLast = 3
    lblCopy(0) = logo(0)
    
    For iIndex = iInit To iLast
        Load lblCopy(iIndex)
        lblCopy(iIndex).Top = lblCopy(iIndex - 1).Top + lblCopy(iIndex - 1).Height + 200
        lblCopy(iIndex).Visible = True
        lblCopy(iIndex).Caption = logo(iIndex)
    Next iIndex
    
    sDefDet(0) = "Level 1 Defect %     "
    sDefDet(1) = "Level 1 Defect Length"
    sDefDet(2) = "Level 2 Defect %     "
    sDefDet(3) = "Level 2 Defect Length"
    Level1_slub_tol = 5
    Level1_length = 10
    Level2_slub_tol = 7
    Level2_length = 5
    
    sCurValueDD(0) = Level1_slub_tol & " %"
    sCurValueDD(1) = Level1_length & " mm"
    sCurValueDD(2) = Level2_slub_tol & " %"
    sCurValueDD(3) = Level2_length & " mm"
        
    iInit = 1
    iLast = 3
    txtCurValueDD(0).Text = sCurValueDD(0)
    cmdParDD(0).Caption = sDefDet(0)
        
    For iIndex = iInit To iLast
        Load txtCurValueDD(iIndex)
        txtCurValueDD(iIndex).Top = txtCurValueDD(iIndex - 1).Top + txtCurValueDD(iIndex - 1).Height
        txtCurValueDD(iIndex).Visible = True
        txtCurValueDD(iIndex).Text = sCurValueDD(iIndex)
        
        Load cmdParDD(iIndex)
        cmdParDD(iIndex).Top = cmdParDD(iIndex - 1).Top + 40 + cmdParDD(iIndex - 1).Height
        cmdParDD(iIndex).Visible = True
        cmdParDD(iIndex).Caption = sDefDet(iIndex)
    Next iIndex
    
    cmdParDD(0).Width = 1815
    cmdParDD(1).Width = 1815
    cmdParDD(2).Width = 1815
    cmdParDD(3).Width = 1815
    
    Me.Height = 7350
    
    FormLoadCount = FormLoadCount + 1
    
    Me.Caption = StringFormat("{0} - {1} - {2} Denier", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"))
    cmdExit.ToolTipText = "Click here to exit this program."
    cmdPlotData.ToolTipText = "Click here to plot a line chart of saved data files."
    cmdBarChart.ToolTipText = "Click here to plot a bar chart of the current data."
    cmdPrint.ToolTipText = "Click here to print the current screen image."
    cmdCurrent.ToolTipText = "Click here to print the current run results."
    
    SensorColors(1) = RGB(0, 0, 255)    ' Blue
    SensorColors(2) = RGB(204, 0, 0)    ' Red
    SensorColors(3) = RGB(204, 255, 0)  ' Yellow
    SensorColors(4) = RGB(0, 102, 0)    ' Green
    SensorColors(5) = RGB(0, 255, 204)  ' Cyan
    SensorColors(6) = RGB(204, 0, 153)  ' Magenta
    SensorColors(7) = RGB(204, 102, 0)  ' Orange
    SensorColors(8) = RGB(0, 51, 0)     ' Darker Green
    For iY = 1 To 8
        lblThreadLine(iY).BackColor = SensorColors(iY)
        lblThreadLine(iY).Caption = GetIniSetting("Application", "NameOfThreadLine") ' CHANGED
    Next iY
    
    SystemStatus.Enabled = False           'Initialize system status word
    SystemStatus.Gain_cal = False
    SystemStatus.IsStartup = False
    SystemStatus.Zero_cal = False
    SystemStatus.Do_ZeroCal = False
    SystemStatus.IsRunning = False
    'Initialize control buttons
    
    If Not SystemStatus.IsFormInit Then     'Do this only the 1st time
        CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier  'Initialize plot setup
        'Debug.Print Target_Denier, Calibration_Denier, Target_denier_tol
        'Target_Denier = 100
        'Calibration_Denier = 100
        'Target_denier_tol = 5
        
        Label1.Visible = False
        Label2.Visible = False
        RunTimer.Visible = False
        StartTImer.Visible = False
        SystemStatus.ComPort = 1            'Default Comm Port
      
        Max_Denier = 1#
        For i = 1 To LAST_SENSOR              'Disable all sensors
            SensorInfos(i).Enabled = False
            SensorInfos(i).Online = False
            SensorInfos(i).Out_to_Lunch = False
            SensorInfos(i).Awaiting_Comm = False
            SensorInfos(i).Cal_Factor = 1
            SensorInfos(i).Zero_Value = 0
            SensorInfos(i).Lowest = SENSOR_LOWEST_DEFAULT
            SensorInfos(i).Highest = SENSOR_HIGHEST_DEFAULT
            SensorInfos(i).SumOfAverages = 0
            Command(i).BackColor = &H808080
            Call Command_Click(i)
        Next i
    End If
    
    For i = 1 To LAST_SENSOR                'Disable these flags
        SensorInfos(i).Awaiting_Comm = False
        SensorInfos(i).Lowest = SENSOR_LOWEST_DEFAULT
        SensorInfos(i).Highest = SENSOR_HIGHEST_DEFAULT
        SensorInfos(i).SumOfAverages = 0
    Next i
    
    SystemStatus.PlotChannel = 1          'Initialize graphing channel #
    iPLotChannels(1) = 1
    chkPlotChan(0).value = 1
    
    ReDim iPlotDenier(1 To 1, 0 To 0)
    
    LineSpeed = 250                        'Use these system settings for now
    Integration_time = 2                    'until the input routines are ready
    Level1_slub_tol = 5
    Level1_length = 10
    Level2_slub_tol = 7
    Level2_length = 5
    Led1.Visible = False
    ZeroCal_Interval = 15 * 60              'Initialize to 15 minute intervals
    tmrIntegration.Interval = Integration_time * 1000
    tmrIntegration.Enabled = False
    tmrZeroCal.Interval = 1000                  'One second cal. timer
    tmrZeroCal.Enabled = False
    
    CWGraph1.TrackMode = cwGTrackDragCursor 'Setup plot
    
    CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
    
    CWGraph1.Cursors.Item(1).YPosition = Target_Denier
    CWGraph1.Axes.Item(XAXIS).Caption = "Elapsed Time (Minutes)"
    
    CWGraph1.Axes.Item(YAXIS).Caption = GetIniSetting("Constants", "GraphYCaption")
    CWGraph1.Caption = GetIniSetting("Constants", "GraphCaption")
   
    SystemStatus.IsFormInit = True                 'Ready to go
    
    SetStatusText "Enable individual sensors by clicking on the appropriate sensor button with the mouse."
                         
    PlotData.Hide
    FiberTrack.Show
    
    lblSetUpFile.Visible = DEBUG_CONTROLS
    PrintStuff.Visible = DEBUG_CONTROLS
    cmdExit.Visible = DEBUG_CONTROLS
    cmdBarChart.Visible = DEBUG_CONTROLS
    cmdPlotData.Visible = DEBUG_CONTROLS
    Text17.Visible = DEBUG_CONTROLS
    
    Plot_Interval = 10
    ReportType = "Current"
    
    ' * * * * * Load default settings for use at startup
    If Not ReadSetupFile(App.Path & "\" & SETUP_FILENAME) Then
        'New logic: DO NOTHING!!!
        
        '
        'mnuOpen_Click
        'Exit Sub
    End If
    
    If SystemStatus.ComPort = 1 Then
        optCom1.value = True
    Else
        optCom2.value = True
    End If
    
    txtCurValueAp(0).Text = LineSpeed
    txtCurValueAp(1).Text = Integration_time
    txtCurValueAp(2).Text = ZeroCal_Interval / 60
    
    txtCurValueDD(0).Text = Level1_slub_tol & " %"
    txtCurValueDD(1).Text = Level1_length & " mm"
    txtCurValueDD(2).Text = Level2_slub_tol & " %"
    txtCurValueDD(3).Text = Level2_length & " mm"
        
#If SCALE_DIAMETER Then
    Select Case Max_Denier
        Case 1#
            iMin_Denier = 0.1
            txtDenPar(0).Text = optDenRange(1).Caption
            iOptDenRangeIndex = 1
        Case Else       'Also covers case 0.5
            Max_Denier = 0.5
            iMin_Denier = 0#
            txtDenPar(0).Text = optDenRange(0).Caption
            iOptDenRangeIndex = 0
    End Select
#Else
    If Max_Denier = 1500 Then
        MsgBox "Warning: Setup file specifies setting of 150 to 1500 Denier.  This parameter is no longer available.  200 to 2000 Denier will be used instead.", vbInformation, "Form_Load"
        Max_Denier = 2000
    End If
    
    Select Case Max_Denier
        Case 1000
            iMin_Denier = 100
            txtDenPar(0).Text = optDenRange(1).Caption
            iOptDenRangeIndex = 1
        Case 2000
            iMin_Denier = 200
            txtDenPar(0).Text = optDenRange(2).Caption
            iOptDenRangeIndex = 2
        Case 5000
            iMin_Denier = 500
            txtDenPar(0).Text = optDenRange(3).Caption
            iOptDenRangeIndex = 3
        Case Else           'Also covers case 500
            iMin_Denier = 0
            Max_Denier = 1#
            txtDenPar(0).Text = optDenRange(0).Caption
            iOptDenRangeIndex = 0
    End Select
#End If
    optDenRange(iOptDenRangeIndex).value = True
    
    txtDenPar(1).Text = Target_Denier
    txtDenPar(2).Text = Calibration_Denier
    txtDenPar(3).Text = Target_denier_tol & " %"
            
    For i = 1 To LAST_SENSOR              'Restore sensor enable buttons
        If SensorInfos(i).Enabled = True Then
            Command(i).BackColor = &H80FF80
            Text1(i).Text = "Enabled"
        Else:
            Command(i).BackColor = &H808080
            Text1(i).Text = "Disabled"
        End If
    Next i
            
    CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier  'Restore plot setup
    CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(1).YPosition = Target_Denier
     
    Select Case PlotViewIndex
        Case 0
            SetStatusText "Normal plot window."
            CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier
            CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(1).YPosition = Target_Denier
    
            ' Turn check mark on menu items on and off.
            mnuExpanded.Checked = False
            mnuNormal.Checked = True
            IsNormal = True
        Case 1
            SetStatusText "Expanded plot window."
            RealTemp = (Target_Denier * (Target_denier_tol / 100)) * 1.5
            IntTemp = RealTemp
            CWGraph1.Axes.Item(YAXIS).SetMinMax Target_Denier - IntTemp, Target_Denier + IntTemp
            CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(1).YPosition = Target_Denier
            
            ' Turn check mark on menu items on and off.
            mnuExpanded.Checked = True
            mnuNormal.Checked = False
            IsNormal = False
        Case Else
    End Select
    If SystemStatus.Enabled = False Then
        SetStatusText "Operating parameters have been restored from save file, now ready to run."
        Init.Enabled = True                 'Enable initialization button
        Init.Visible = True
    End If
    Exit Sub
fibertrack_err_rtn:
    If Err.Number = 53 Then
        Exit Sub            '53 = file not found
    End If
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open File Error, Form_Load"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuFile
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim iTempT2L As Integer
    Dim iTempF1L As Integer
    Dim iTempT2W As Integer
    Dim iTempF1W As Integer
    Dim commandWidth As Integer
    Dim commandLeft As Integer
    Dim threadLineWidth As Integer
    Dim offset As Integer
    Dim offsetLeft As Integer
    offset = 100
    offsetLeft = 100
    
    CWGraph1.Width = FiberTrack.Width - 335
    commandWidth = Int(FiberTrack.Width / 9)
    commandLeft = FiberTrack.ScaleLeft + Int(Command(1).Width / 2)
    Command(1).Width = commandWidth
    Command(1).Left = commandLeft

    For i = 2 To LAST_SENSOR
        Command(i).Left = Command(i - 1).Left + Command(1).Width
        Command(i).Width = commandWidth
    Next i

    threadLineWidth = Int(Command(1).Width / 5 * 4) - offset
    lblThreadLine(1).Width = threadLineWidth - offset
    lblThreadLine(1).Left = commandLeft
    Text2(1).Width = commandWidth - lblThreadLine(1).Width + offset
    Text2(1).Left = lblThreadLine(1).Left + lblThreadLine(1).Width - offsetLeft

    For i = 2 To LAST_SENSOR
        lblThreadLine(i).Left = Command(i).Left + 20
        lblThreadLine(i).Width = lblThreadLine(1).Width - 20 - offset
    Next i
    
    For j = 2 To LAST_SENSOR
        Text2(j).Left = lblThreadLine(j).Left + lblThreadLine(j).Width - offsetLeft
        Text2(j).Width = Text2(1).Width + offset
    Next j

    Text1(1).Width = Command(1).Width
    Text1(1).Left = Command(1).Left
    
    For i = 2 To LAST_SENSOR
        Text1(i).Left = Command(i).Left
        Text1(i).Width = Command(1).Width
    Next i
 
    Init.Left = Command(1).Left
    GainCal.Left = Command(1).Left
    StartStop.Left = Command(1).Left
    StartStop.Top = FiberTrack.Height - FiberTrack.Height / 6 - StartStop.Height
    GainCal.Top = StartStop.Top - StartStop.Height
    Init.Top = GainCal.Top - GainCal.Height
    
    ZeroCal.Left = Init.Left + Init.Width + Int(FiberTrack.Width / 40)
    ZeroCal.Top = Init.Top
    
    RunTimer.Left = Command(LAST_SENSOR).Left + Command(1).Width - RunTimer.Width
    StartTImer.Left = RunTimer.Left - StartTImer.Width
    RunTimer.Top = FiberTrack.Height - FiberTrack.Height / 6 - RunTimer.Height
    StartTImer.Top = RunTimer.Top
    
    Label2.Width = RunTimer.Width
    Label2.Left = Command(LAST_SENSOR).Left + Command(1).Width - Label2.Width
    Label2.Top = RunTimer.Top - Label2.Height
    
    Label1.Width = Label2.Width
    Label1.Left = Label2.Left - Label1.Width - 50
    Label1.Top = Label2.Top
     
    CWGraph1.Height = Init.Top - Text1(1).Top - Text1(1).Height - FiberTrack.Height / 20
    
    FiberTrack.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo form_unload_err_hdr
    
    FileNamesNumber = FreeFile
          
    Open "FiberNames.fna" For Output As #FileNamesNumber
    For FileNumber = 1 To IFILE_NAMES_LIMIT
        If FileNames(FileNumber) <> "" Then
            Print #FileNamesNumber, FileNames(FileNumber)
        End If
    Next FileNumber
    
    Close #FileNamesNumber
    
    Set mo_Comm = Nothing
    
    End
    
    Exit Sub
    
form_unload_err_hdr:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Form_Unload DataFile Error, FiberTrack"
    End
    Exit Sub
End Sub

Private Sub fraAbout_Click()
    TopLevelMenuClick
End Sub

Private Sub GainCal_Click()
    Dim index As Integer, Dummy
    Dim iY As Integer
    Dim iCalValue(1 To LAST_SENSOR, 1 To 5) As Integer
    Dim fCalFactor(1 To LAST_SENSOR, 1 To 5) As Single
    Dim iV As Integer
    Dim iSe As Integer

    If SystemStatus.Enabled = True Then
        SystemStatus.Gain_cal = True         'Set this status flag
        'Send the Start gain calibration packet
        
        SetStatusText "Performing gain calibration sequence. Using target gain value of " & Calibration_Denier & " Denier."    'V1.7
        Screen.MousePointer = vbHourglass
  
        For iV = 1 To 5
            For iSe = 1 To LAST_SENSOR
                iCalValue(iSe, iV) = 0
                fCalFactor(iSe, iV) = 0#
            Next iSe
        Next iV
  
        For iY = 1 To 5
            Command_Packet(3) = CAL_COMMAND
            Command_Packet(2) = ALL_CHANNELS        'Sensor #
            PacketChecksum Command_Packet           'compute packet checksum
            mo_Comm.WriteData Command_Packet        'Xmit command packet
            
            'Wait for a period = integratation time, then collect sensor data.
            SystemStatus.IsWait2Seconds = True        'Set this flag
            tmr2Sec.Interval = 2000 * Integration_time 'Initialize timer
            tmr2Sec.Enabled = True
            
            Do Until SystemStatus.IsWait2Seconds = False 'Wait on timer
                Dummy = DoEvents()                  'Suspend task - Give time to OS
            Loop
    
            'Poll each sensor for data update
            Command_Packet(3) = POLL_COMMAND
            
            For index = 1 To LAST_SENSOR
                If SensorInfos(index).Online Then
                    Command_Packet(2) = index           ' Sensor #
                    PacketChecksum Command_Packet       ' compute packet checksum
                    Call InitReceiveCvFunction(index)   ' Set up for sensor comm packet
                    mo_Comm.WriteData Command_Packet    ' Xmit command packet
                    SystemStatus.IsWait2Seconds = True  ' Set this flag
                    tmr2Sec.Interval = 1000             ' Enable 1sec timer
                    tmr2Sec.Enabled = True
                    
                    'Wait up to 2 sec. for response.
                    'Expecting 26 byte packet from sensor.
                    mo_Comm.WaitForPacket
                    
                    tmr2Sec.Enabled = False              'Disable timer
                    SystemStatus.IsWait2Seconds = False
                    SensorInfos(index).Awaiting_Comm = False
                    
                    If RxComm.done = True And RxComm.error = False Then
                        Call ProcessCalData(index)          'Process packet
            
                        iCalValue(index, iY) = SensorInfos(index).Cal_Value     ' i = sensor number
                        fCalFactor(index, iY) = SensorInfos(index).Cal_Factor
                    Else
                        SetStatusText "Comm failure on sensor #" & index
                    End If
                End If
            Next index
            
            SystemStatus.Gain_cal = False          'Clear this flag
        Next iY

        For index = 1 To LAST_SENSOR
            If SensorInfos(index).Online = True Then
                SensorInfos(index).Cal_Value = (iCalValue(index, 3) + iCalValue(index, 4) + iCalValue(index, 5)) / 3
                SensorInfos(index).Cal_Factor = (fCalFactor(index, 3) + fCalFactor(index, 4) + fCalFactor(index, 5)) / 3
                Text1(index).Text = "Cal Factor = " & Format(SensorInfos(index).Cal_Factor, "##.0000")
            End If
        Next index
    End If

    SetStatusText "Finished gain calibration sequence."
    GainCal.Enabled = False
    Screen.MousePointer = vbArrow
End Sub

Private Sub Init_Click()
On Error GoTo init_err_rtn
    Dim index As Integer
    Dim Dummy

    If SystemStatus.Enabled = False Then   'Ignore startup if system enabled
        If Not mo_Comm.OpenPort(CInt(SystemStatus.ComPort)) Then Exit Sub
        
        SetStatusText "Performing Sensor initialization, please wait."
        
        ' Change mouse pointer to hourglass.
        Screen.MousePointer = vbHourglass
  
        SystemStatus.IsStartup = True          'Flag startup sequence
        SystemStatus.Do_ZeroCal = False      'No zerocals unless operator enabled
        
        'Insert initialization message into comm xmit buffer
        Host_send(0) = FF_BYTE                'Comm packet starts with "0"
        Host_send(1) = 13                     'Packet byte count
        Host_send(2) = ALL_CHANNELS
        Host_send(3) = INIT_COMMAND
  
        InsertWord Host_send, LineSpeed, 4     'Insert linespeed into comm buffer
#If SCALE_DIAMETER Then
        InsertWord Host_send, CInt(Target_Denier * 1000), 6
#Else
        InsertWord Host_send, Target_Denier, 6
#End If
        Host_send(8) = Integration_time
        Host_send(9) = Level1_slub_tol
        'JW 5/25/00 Length gets scaled by 10
        Host_send(10) = Level1_length * 10
        Host_send(11) = Level2_slub_tol
        Host_send(12) = Level2_length * 10
  
        PacketChecksum Host_send                'Compute packet checksum
        mo_Comm.WriteData Host_send             'Xmit packet
        
        SystemStatus.IsWait2Seconds = True          'Set this flag
        tmr2Sec.Interval = 5000                 'Initialize 5 sec. timer
        tmr2Sec.Enabled = True
        Do Until SystemStatus.IsWait2Seconds = False 'Wait 5 sec.
            Dummy = DoEvents()                  'Suspend task - Give time to OS
        Loop
        
        'Poll each sensor for status update
        Command_Packet(0) = FF_BYTE             'Build a poll command packet
        Command_Packet(1) = 4                   'packet count
        Command_Packet(3) = POLL_COMMAND
        For index = 1 To LAST_SENSOR
            If SensorInfos(index).Enabled = True Then
                SetStatusText "Now polling sensor " & index & " - will timeout in 1 second."
                
                Command_Packet(2) = index           ' Sensor #
                PacketChecksum Command_Packet       ' compute packet checksum
                Call InitReceiveCvFunction(index)   ' Set up for sensor comm packet
                mo_Comm.WriteData Command_Packet    ' Xmit command packet

                SystemStatus.IsWait2Seconds = True  ' Set this flag
                tmr2Sec.Interval = 1000             ' Enable 1sec timer
                tmr2Sec.Enabled = True
                
                'Wait up to 2 sec. for response.
                'Expecting 26 byte packet from sensor.
                mo_Comm.WaitForPacket

                tmr2Sec.Enabled = False              'Disable timer
                SystemStatus.IsWait2Seconds = False
                SensorInfos(index).Awaiting_Comm = False
                '      Debug.Print "Sensor # " & i & " received " & MSComm1.InBufferCount & " bytes."
                
                If RxComm.done = True And RxComm.error = False Then
                    SetStatusText "Received correct length packet from Sensor # " & index
                    Call ProcessSensorInit(index)          'Process packet
                Else
                    Text1(index).Text = "No Response"
                    Command(index).BackColor = RGB(255, 75, 75)     ' Red
                    SensorInfos(index).Highest = Calibration_Denier
                    SensorInfos(index).Lowest = Calibration_Denier
                    SensorInfos(index).Enabled = False
                    BarChart.lblName(index - 1).Visible = False
                End If
            End If
        Next index
        
        SetStatusText "Finished initialization of all sensors. Now perform a Zero and Gain calibration sequence."
        SystemStatus.Enabled = True
        
        'Disable this button
        Init.Enabled = False
        'Enable buttons for next steps
        StartStop.Enabled = True
        GainCal.Enabled = True
        ZeroCal.Enabled = True
        'Reset mouse
        Screen.MousePointer = vbArrow
    End If

    Exit Sub
    '*******************Sensors initialized********************
init_err_rtn:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Initialization Error, Initialization FiberTrack"
    Exit Sub
End Sub

Private Sub lblCopy_Click(index As Integer)
    fraAbout_Click
End Sub

'Click on ThreadLine.
Private Sub lblThreadLine_Click(index As Integer)
On Error GoTo lblThreadLine_err_hdr
    comDialog2.CancelError = True
    comDialog2.Color = SensorColors(index)
    
    comDialog2.Flags = cdlCCPreventFullOpen & cdlCCRGBInit
    
    comDialog2.ShowColor
    
    ' SensorColors goes from 1 to 8
    SensorColors(index) = comDialog2.Color
    lblThreadLine(index).BackColor = comDialog2.Color
    BarChart.lblName(index - 1).BackColor = comDialog2.Color
    
    Exit Sub

lblThreadLine_err_hdr:
    'cdlCancel = Cancel button clicked
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "FiberTrack, lblThreadLine_Click"
End Sub

Private Sub mnuAbout_Click()
    'fraAbout.Visible = True
    frmAbout.Show vbModal
    Exit Sub
    fraSetup.Visible = False
    fraHelp.Visible = True
    'Escape button will fire cmdExitHlp_Click
    cmdExitHlp.Cancel = True
    cmdAbout_Click
End Sub

Private Sub mnuBarGraph_Click()
    If IsMain Then
        SetStatusText "Bar Graph window."
        ' Turn check mark on menu items on and off.
        mnuMain.Checked = Not mnuMain.Checked
        mnuBarGraph.Checked = Not mnuBarGraph.Checked
        IsMain = False
        FiberTrack.Hide
        ' BarChart.WindowState = FiberTrack.WindowState
        BarChart.Show
     End If
End Sub

Private Sub mnuExcel_Click()
    Load PlotData
    PlotData.ChooseExcelFile
    Unload PlotData
End Sub

Private Sub Restart_Click()
    Call Shell(App.Path & "\" & App.EXEName & ".exe", vbNormalFocus)
    Unload Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'See also mnuNormal_Click for Normal plot window.
Private Sub mnuExpanded_Click()
    If IsNormal Then
        SetStatusText "Expanded plot window."
        RealTemp = (Target_Denier * (Target_denier_tol / 100)) * 1.5
        IntTemp = RealTemp
        CWGraph1.Axes.Item(YAXIS).SetMinMax Target_Denier - IntTemp, Target_Denier + IntTemp
#If SCALE_DIAMETER Then
        CWGraph1.Cursors.Item(1).YPosition = Target_Denier
#Else
        CWGraph1.Cursors.Item(1).YPosition = Target_Denier - Target_Denier * 0.001
#End If
        ' Turn check mark on menu items on and off.
        mnuExpanded.Checked = True
        mnuNormal.Checked = False
        IsNormal = False
        PlotViewIndex = 1
    End If
End Sub

Private Sub mnuFile_Click()
    TopLevelMenuClick
End Sub

Private Sub mnuHelp_Click()
    TopLevelMenuClick
End Sub

Private Sub mnuMain_Click()
    If Not IsMain Then
        SetStatusText "Main window."
        ' Turn check mark on menu items on and off.
        mnuMain.Checked = Not mnuMain.Checked
        mnuPlotDataFile.Checked = Not mnuPlotDataFile.Checked
        IsMain = True
        
        FiberTrack.Show
    
        PlotData.Hide       ' vbModal
    End If
End Sub

'See also Extended Plot window in mnuExpanded_Click.
Private Sub mnuNormal_Click()
    If Not IsNormal Then
        SetStatusText "Normal plot window."
        CWGraph1.Axes.Item(YAXIS).SetMinMax MIN_DENIER, Max_Denier
#If SCALE_DIAMETER Then
        CWGraph1.Cursors.Item(1).YPosition = Target_Denier
#Else
        CWGraph1.Cursors.Item(1).YPosition = Target_Denier - Target_Denier * 0.05
#End If

        ' Turn check mark on menu items on and off.
        mnuNormal.Checked = True
        mnuExpanded.Checked = False
        IsNormal = True
        PlotViewIndex = 0
    End If
End Sub

'Open setup file
Private Sub mnuOpen_Click()
    Dim i           As Integer
    Dim iTemp       As Integer
    Dim iTempP1     As Integer
    Dim iResponse   As Integer

On Error GoTo mnuOpen_err_hdr

    comDialog2.CancelError = True

    comDialog2.DialogTitle = "Open Settings"
    comDialog2.Filter = "FiberTrack Setup File (*.sup)|*.sup|FiberTrack Default Settings (" & SETUP_FILENAME & ")|" & SETUP_FILENAME & "|All Files (*.*)|*.*"
    comDialog2.FilterIndex = 1
    comDialog2.Flags = cdlOFNExplorer Or cdlOFNExtensionDifferent Or cdlOFNNoChangeDir Or cdlOFNHideReadOnly

    comDialog2.ShowOpen
    
    lblSetUpFile.Caption = "SetUp file -> " & comDialog2.FileName
    
    ReadSetupFile comDialog2.FileName
    
    txtCurValueAp(0).Text = LineSpeed
    txtCurValueAp(1).Text = Integration_time
    txtCurValueAp(2).Text = ZeroCal_Interval / 60

    optDenRange(DenierRangeIndex).value = True
    iOptDenRangeIndex = DenierRangeIndex

    'iMin_Denier = Max_Denier / 10
    'CLM1112
    If Max_Denier = 1# Then
        iMin_Denier = 0
    Else
        iMin_Denier = Max_Denier / 10
    End If
    
    txtDenPar(1).Text = Target_Denier
    txtDenPar(2).Text = Calibration_Denier
    txtDenPar(3).Text = Target_denier_tol & " %"
    
    txtCurValueDD(0).Text = Level1_slub_tol & " %"
    txtCurValueDD(1).Text = Level1_length & " mm"
    txtCurValueDD(2).Text = Level2_slub_tol & " %"
    txtCurValueDD(3).Text = Level2_length & " mm"
    
    If SystemStatus.ComPort = 1 Then
        optCom1.value = True
    Else
        optCom2.value = True
    End If

    For i = 1 To LAST_SENSOR              'Restore sensor enable buttons
       If SensorInfos(i).Enabled = True Then
             Command(i).BackColor = &H80FF80
             Text1(i).Text = "Enabled"
        Else:
             Command(i).BackColor = &H808080
             Text1(i).Text = "Disabled"
        End If
    Next i
    
    CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier  'Restore plot setup
    CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
    CWGraph1.Cursors.Item(1).YPosition = Target_Denier
 
    Select Case PlotViewIndex
        Case 0
            SetStatusText "Normal plot window."
            CWGraph1.Axes.Item(YAXIS).SetMinMax iMin_Denier, Max_Denier
            CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(1).YPosition = Target_Denier
            
            ' Turn check mark on menu items on and off.
            mnuExpanded.Checked = False
            mnuNormal.Checked = True
            IsNormal = True

        Case 1
            SetStatusText "Expanded plot window."
            RealTemp = (Target_Denier * (Target_denier_tol / 100)) * 1.5
            IntTemp = RealTemp
            CWGraph1.Axes.Item(YAXIS).SetMinMax Target_Denier - IntTemp, Target_Denier + IntTemp
            CWGraph1.Cursors.Item(2).YPosition = Target_Denier + (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(3).YPosition = Target_Denier - (Target_denier_tol / 100 * Target_Denier)
            CWGraph1.Cursors.Item(1).YPosition = Target_Denier
            
            ' Turn check mark on menu items on and off.
            mnuExpanded.Checked = True
            mnuNormal.Checked = False
            IsNormal = False
        Case Else
    End Select
 
    If SystemStatus.Enabled = False Then
        SetStatusText "Operating parameters have been restored from save file, now ready to run."
        Init.Enabled = True                 'Enable initialization button
        Init.Visible = True
    End If
    
    Exit Sub
 
mnuOpen_err_hdr:

    If Err.Number = cdlCancel Then
        Exit Sub
    End If
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open Setup File Error, mnuOpen FiberTrack"
End Sub

Private Sub mnuOpenDataFile_Click()
    mnuPlotDataFile_Click
    PlotData.ChoosePlotDataFile
End Sub

Private Sub mnuOptions_Click()
    TopLevelMenuClick
End Sub

Private Sub mnuPlotDataFile_Click()
    If IsMain Then
        SetStatusText "Plot Data File window."
        
        ' Turn check mark on menu items on and off.
        mnuMain.Checked = Not mnuMain.Checked
        mnuPlotDataFile.Checked = Not mnuPlotDataFile.Checked
        IsMain = False
        
        FiberTrack.Hide
    
        PlotData.Show
        'PlotData.WindowState = FiberTrack.WindowState
        PlotData.cboomPlotData.ListIndex = 0
        PlotData.cmdReturn.SetFocus
    End If
End Sub

Private Sub mnuPrint_Click()

On Error GoTo mnuPrint_err_hdr

    comDialog2.CancelError = True
    comDialog2.Flags = &H100000 + &H4
    '   &H80000  disables Print to File box
    '   &H100000 hides the Print to File box
    '   &H4      disables the Selection option
    
    comDialog2.ShowPrinter
    PrintSetupFile
    
    Exit Sub
    
mnuPrint_err_hdr:

    If Err.Number = cdlCancel Then  'cdlCancel = Cancel button clicked
        Exit Sub
    End If

    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Print error, mnuPrint"

End Sub

'Save setup file
Private Sub mnuSave_Click()
    Dim i           As Integer
    Dim iResponse   As Integer

On Error GoTo mnuSave_err_hdr

    comDialog2.CancelError = True

    comDialog2.DialogTitle = "Save Settings"
    comDialog2.Filter = "FiberTrack Setup File (*.sup)| *.sup"
    comDialog2.FilterIndex = 1
    
    ' Save Flags
    '   cdlOFNExplorer              = &H80000&
    '   cdlOFNCreatePrompt          = &H02000&
    '   cdlOFNNoChangeDir           = &H00008&
    '   cdlOFNHideReadOnly          = &H00004&
    '   cdlOFNOverwritePrompt       = &H00002&

    comDialog2.Flags = &H80200E
    ' comDialog2.Flags = cdlOFNCreatePrompt & cdlOFNExplorer & _
            cdlOFNOverwritePrompt & cdlOFNNoChangeDir

    comDialog2.ShowSave
    SaveSetupFile comDialog2.FileName
    
    SetStatusText "Saving operating parameters to file named " & comDialog2.FileName & "."
    
    '* lblSetUpFile.Visible = True
    lblSetUpFile.Caption = "SetUp file -> " & comDialog2.FileName
    Exit Sub

mnuSave_err_hdr:
    If Err.Number = cdlCancel Then Exit Sub     'Cancel button in Common Dialog

    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Save Setup File Error, mnuSave"

End Sub

'Setup from top-level menubar
Private Sub mnuSetup_Click()
    fraHelp.Visible = False
    fraSetup.Visible = True
    'Escape button will fire cmdDone_Click
    cmdDone.Cancel = True
End Sub

Private Sub mnuView_Click()
    TopLevelMenuClick
End Sub

Private Sub mnuWindow_Click()
    TopLevelMenuClick
End Sub

Private Sub MSComm1_OnComm()
    mo_Comm.HandleEvent
End Sub

Private Sub optCom1_Click()
    TempComPort = 1
End Sub

Private Sub optCom2_Click()
    TempComPort = 2
End Sub

Private Sub optDenRange_Click(index As Integer)
#If SCALE_DIAMETER Then
    Dim iMinDen(0 To 1) As Single
    Dim iMaxDen(0 To 1) As Single
    
    iMinDen(0) = 0#
    iMaxDen(0) = 0.5
    
    iMinDen(1) = 0.1
    iMaxDen(1) = 1#
#Else
    Dim iMinDen(0 To 3) As Integer
    Dim iMaxDen(0 To 3) As Integer
    
    iMinDen(0) = 0
    iMaxDen(0) = 1#
    
    iMinDen(1) = 0
    iMaxDen(1) = 100
    
    iMinDen(2) = 100
    iMaxDen(2) = 1000
    
    iMinDen(3) = 500
    iMaxDen(3) = 5000
#End If

    txtDenPar(0).Text = optDenRange(index).Caption
    iTempMinDen = iMinDen(index)
    iTempMaxDen = iMaxDen(index)
    iTempOptDenRangeIndex = index
End Sub

Private Sub picSetup_KeyPress(KeyAscii As Integer)
    MsgBox "picSetup Key = " & KeyAscii, , "picSetup_KeyPress"
End Sub

Private Sub StartStop_Click()
    Dim index           As Integer, Dummy
    Dim longHours       As Long
    Dim longMinutes     As Long
    Dim longSeconds     As Long
    Dim sensorIndex     As Integer
    Dim response        As Integer
    Dim dataFile        As FTDataFile
    Dim dataFileLine    As FTDataLine
    Dim currentSensorReportData As SensorReportData
    Dim temp

    If SystemStatus.Enabled And SystemStatus.IsRunning Then
        response = MsgBox("Do you really want to stop?", vbYesNo Or vbExclamation, "FiberTrack, StartStop, StopButton")
        If response = vbNo Then
        '    SystemStatus.IsRunning = False
            Exit Sub
        End If
    End If

    If SystemStatus.Enabled And Not SystemStatus.IsRunning Then
        CWGraph1.ClearData                        'clear graph data display
        'JW
        CWGraph1.Axes.Item(XAXIS).SetMinMax 0, 20  'Initialize plot setup
        Set dataFile = New FTDataFile
        dataFile.FileFormat = Text
        'CM::Replaced append with automatic incrementing of file #
        For index = 1 To LAST_SENSOR
            dataFile.SensorOnline(index) = SensorInfos(index).Online
            SensorInfos(index).Package = Trim$(Text2(index).Text)
            If Not SensorInfos(index).Online Then
                SensorInfos(index).Highest = Calibration_Denier
                SensorInfos(index).Lowest = Calibration_Denier
                BarChart.lblName(index - 1).Visible = False
            End If
        Next index
        
        dataFile.DenierTarget = Target_Denier
        dataFile.IntegrationTime = Integration_time
        dataFile.LineSpeed = LineSpeed
        
        If Not dataFile.OpenFile Then Exit Sub
        
        dataFile.WriteHeader
        Me.Caption = StringFormat("{0} - {1} - {2} - {3}", GetIniSetting("Contants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"), dataFile.FileName)
        SetStatusText StringFormat("Temporary data file is named {0}.", dataFile.FileName)
        StartStop.Caption = "Stop"
        SystemStatus.IsRunning = True
        Led1.Caption = "Updating"
        RunTimer.Visible = True
        StartTImer.Visible = True
        cmdExit.Enabled = False
        Label1.Visible = True
        Label2.Visible = True
        ZeroCal.Enabled = False                 'Disable these buttons
        GainCal.Enabled = False
        Init.Enabled = False
        StartTImer.Caption = Now                'Set start time
        
        'Send the Start data collection packet
        For index = 1 To LAST_SENSOR
            SensorInfos(index).Num_Cycles = 0         'Clear integration count
        Next index
        
        Command_Packet(3) = STARTDATA_COMMAND
        Command_Packet(2) = ALL_CHANNELS        'Sensor #
        PacketChecksum Command_Packet           'compute packet checksum
        mo_Comm.WriteData Command_Packet        'Xmit command packet

        SystemStatus.IsWaitInitTime = True      'Will wait for tmrIntegration
        SystemStatus.StartTime = Now           'Used to calculate runtime
        SystemStatus.Next_Zerocal = ZeroCal_Interval  'seconds to next zero cal
        tmrIntegration.Enabled = True
        tmrZeroCal.Enabled = True
        
        'JW - 2/27/00
        Command_Packet2(0) = FF_BYTE            'Comm packet starts with "0"
        Command_Packet2(1) = 16                 'Packet byte count

        '***********Main data acquisition loop******************************
        'Wait for a period = integratation time, then collect sensor data.
        Do Until SystemStatus.Enabled = False
            Call SetTime(longHours, longMinutes, longSeconds)   'Update front panel elapsed time
            RunTimer.Caption = StringFormat("{0}:{1}:{2}", Format(RunInHours, "00"), Format(RunInMinutes, "00"), Format(RunInSeconds, "00"))
            Do Until SystemStatus.IsWaitInitTime = False 'Wait on timer
                Dummy = DoEvents()                    'Suspend task - Give time to OS
            Loop
            
            SystemStatus.IsWaitInitTime = True      'reenable timer flag
            
            'Print system time in output file
            Set dataFileLine = New FTDataLine
            dataFileLine.Hours = RunInHours
            dataFileLine.Minutes = RunInMinutes
            dataFileLine.Seconds = RunInSeconds
            
            'Poll each sensor for data update
            '    Led1.Visible = True                     'Turn on sample led
            Call SetTime(RunInHours, RunInMinutes, RunInSeconds)   'Update run time
            
            'JW - 2/28/00
            Command_Packet2(3) = POLL_COMMAND
            iPLot = 1
            For index = 1 To LAST_SENSOR
                If SensorInfos(index).Online = True Then
                    'JW - 2/27/00
                    Command_Packet2(2) = index              'Sensor #
                    'JW 2/27/00 More information has to be added to command buffer.
                    Call CommBufferFix(index)
                    PacketChecksum Command_Packet2      'compute packet checksum
                    Call InitReceiveCvFunction(index)   'Set up for sensor comm packet
                    mo_Comm.WriteData Command_Packet2   'Xmit command packet
                    SystemStatus.IsWait2Seconds = True  'Set this flag
                    tmr2Sec.Interval = 500              'Enable 1/2sec timer
                    tmr2Sec.Enabled = True
                    
                    'Wait up to 2 sec. for response.
                    'Expecting 26 byte packet from sensor.
                    mo_Comm.WaitForPacket
                    
                    tmr2Sec.Enabled = False              'Disable timer
                    SystemStatus.IsWait2Seconds = False
                    SensorInfos(index).Awaiting_Comm = False
                    If RxComm.done = True And RxComm.error = False Then
                        SensorInfos(index).Num_Cycles = SensorInfos(index).Num_Cycles + 1
                        'JW - 5/25/00 Delay processing of sensor data for several cycles.
                        If SensorInfos(index).Num_Cycles > 2 Then
                            Set currentSensorReportData = GetSensorReportData(index)
                            dataFileLine.Avg(index) = currentSensorReportData.Average 'Process packet
                            dataFileLine.Cv(index) = currentSensorReportData.Cv
                        Else
                            temp = mo_Comm.ReadData()      ' Flush comm buffer contents
                            'Debug.Assert False            ' Just checking
                        End If
                    Else
                        SetStatusText "Comm failure on sensor #" & index
                    End If
                End If
            Next index
            
            dataFile.WriteDataLine dataFileLine
            CWGraph1.ChartY iPlotDenier, 1 / (60 / Integration_time)
            iPLot = 1
            Plot_Data
            Led1.Visible = False                     'Turn off sample led
            'Check if time for next Zero cal cycle
            If SystemStatus.Do_ZeroCal = True And SystemStatus.Next_Zerocal = 0 Then
                SystemStatus.Next_Zerocal = ZeroCal_Interval  'setup next zero cal
                Do_ZeroCal                        'Perform Zero calibration
            End If
            'JW - Fixed Integration period timer 11/26/99
            tmrIntegration.Interval = Integration_time * 1000 'Reinitialize integration timer
            SystemStatus.IsWaitInitTime = True      'Will wait for next integration period
        Loop
        '********************End of main data acquisition loop********************
        'Will get here if Stop button depressed
        
        dataFile.WriteFooter
        mo_Comm.ClosePort
        dataFile.CloseFile            'close the output file
        '  Reset
        ' <<<<<<<<<< print summary report
        response = MsgBox("Do you want to print the summary report?", vbYesNo Or vbQuestion, GetIniSetting("Constants", "Name"))
        If response = vbYes Then
            ReportType = "Summary"
            SaveReport
            ReportType = "Current"
        End If              'response = Yes
        ' >>>>>>>>>> print summary report

        ' >>>>>>>>>> save data file
        response = MsgBox(StringFormat("Do you want to save the data file ({0})?", dataFile.FileName), vbYesNo Or vbQuestion, GetIniSetting("Constants", "Name"))
        If response = vbNo Then
            dataFile.EraseFile
        End If

        Me.Caption = StringFormat("{0} - {1} - {2}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"))
        ' >>>>>>>>>> save data file
        cmdExit.Enabled = True
        '3.29.99  Call Form_Load                        'reinitialize system
        SystemStatus.Enabled = False           'Initialize system status word
        SystemStatus.Gain_cal = False
        SystemStatus.IsStartup = False
        SystemStatus.Zero_cal = False
        SystemStatus.Do_ZeroCal = False
        SystemStatus.IsRunning = False
    
        'Initialize control buttons
        StartStop.Enabled = False
        StartStop.Caption = "Start"
        ZeroCal.Enabled = False                 'Disable cal buttons
        GainCal.Enabled = False
        Init.Enabled = True

        For index = 1 To LAST_SENSOR                'Disable these flags
            SensorInfos(index).Awaiting_Comm = False
            SensorInfos(index).Lowest = SENSOR_LOWEST_DEFAULT
            SensorInfos(index).Highest = SENSOR_HIGHEST_DEFAULT
            SensorInfos(index).SumOfAverages = 0
            SensorInfos(index).Online = False
        Next index

        CWGraph1.ClearData
        SystemStatus.IsFormInit = True                 'Ready to go
        SetStatusText "Enable individual sensors by clicking on the appropriate sensor button with the mouse."
    ElseIf SystemStatus.Enabled And SystemStatus.IsRunning Then
        'This will disable system
        SystemStatus.Enabled = False
    End If
End Sub

'If this is for 5 sec. timer, will clear status
Private Sub tmr2Sec_Timer()
    If SystemStatus.IsWait2Seconds = True Then
        SystemStatus.IsWait2Seconds = False
        tmr2Sec.Enabled = False
    End If
End Sub

'This is for Integration interval timer, will clear status
Private Sub tmrIntegration_Timer()
    If SystemStatus.IsWaitInitTime = True Then
        SystemStatus.IsWaitInitTime = False
    End If
End Sub

'This is for Zero calibration interval timer.
Private Sub tmrZeroCal_Timer()
    If SystemStatus.Do_ZeroCal = True Then
        If SystemStatus.Next_Zerocal > 0 Then     'Decrement timer until zero
            SystemStatus.Next_Zerocal = SystemStatus.Next_Zerocal - 1
        End If
    End If
End Sub

'Click to perform Zero Calibration
Private Sub ZeroCal_Click()
    Dim index As Integer
    Dim Dummy As Variant
    
    If SystemStatus.Enabled = True Then
        SystemStatus.Zero_cal = True         'Set this status flag
        SystemStatus.Do_ZeroCal = True       'Signal for periodic zerocal
        
        'Send the Start zero calibration packet
        SetStatusText "Performing zero calibration sequence."
        Screen.MousePointer = vbHourglass
  
        Command_Packet(2) = ALL_CHANNELS      'Sensor #
        Command_Packet(3) = ZERO_COMMAND
        PacketChecksum Command_Packet           'compute packet checksum
        mo_Comm.WriteData Command_Packet        'Xmit command packet

        'Wait for a period = integratation time, then collect sensor data.
        SystemStatus.IsWait2Seconds = True        'Set this flag
        tmr2Sec.Interval = 2000 * Integration_time  'Initialize timer
        tmr2Sec.Enabled = True
        
        Do Until SystemStatus.IsWait2Seconds = False 'Wait on timer
            Dummy = DoEvents()                  'Suspend task - Give time to OS
        Loop

        'Poll each sensor for data update
        Command_Packet(3) = POLL_COMMAND
        For index = 1 To LAST_SENSOR
            If SensorInfos(index).Online = True Then
                Command_Packet(2) = index           ' Sensor #
                PacketChecksum Command_Packet       ' compute packet checksum
                Call InitReceiveCvFunction(index)   ' Set up for sensor comm packet
                mo_Comm.WriteData Command_Packet    ' Xmit command packet
                SystemStatus.IsWait2Seconds = True  ' Set this flag
                tmr2Sec.Interval = 1000             ' Enable 1sec timer
                tmr2Sec.Enabled = True
                
                'Wait up to 2 sec. for response; expecting 26 byte packet from sensor.
                mo_Comm.WaitForPacket
                
                tmr2Sec.Enabled = False             'Disable timer
                SystemStatus.IsWait2Seconds = False
                SensorInfos(index).Awaiting_Comm = False
                
                If RxComm.done = True And RxComm.error = False Then
                    Call ProcessZeroData(index)          'Process packet
                Else
                    SetStatusText "Comm failure on sensor #" & index
                End If
            End If
        Next index
        
        SystemStatus.Zero_cal = False          'Clear this flag
    End If
    
    SetStatusText "Finished Zero Calibration sequence."
    ZeroCal.Enabled = False
    Screen.MousePointer = vbArrow
End Sub

Private Sub TopLevelMenuClick()
    HideAllFrames
End Sub

Private Sub HideAllFrames()
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is VB.Frame Then
            ctl.Visible = False
        End If
    Next ctl
End Sub


'Jw 5/25/00 This is added to place the correct objects into the command buffer.
'The last average is used to compute the new slub values.
Private Sub CommBufferFix(sensorIndex As Integer)
    Dim temp1 As Double
    Dim temp2, temp3 As Long
    
    'Do level 1 thick and thin spots
    temp1 = CurrentAverage(sensorIndex) / SensorInfos(sensorIndex).Cal_Factor
    temp2 = temp1 + SensorInfos(sensorIndex).Zero_Value
    temp1 = temp1 * (Level1_slub_tol / 100#)
    temp3 = temp2 + temp1                   'temp3 is slub target value
    InsertWord Command_Packet2, temp3, 6    'place in comm buffer
    temp3 = temp2 - temp1                   'temp3 is thin spot target value
    InsertWord Command_Packet2, temp3, 8
    'Do level 2 thick and thin spots
    temp1 = CurrentAverage(sensorIndex) / SensorInfos(sensorIndex).Cal_Factor
    temp2 = temp1 + SensorInfos(sensorIndex).Zero_Value
    temp1 = temp1 * (Level2_slub_tol / 100#)
    temp3 = temp2 + temp1                   'temp3 is slub target value
    InsertWord Command_Packet2, temp3, 11   'place in comm buffer
    temp3 = temp2 - temp1                   'temp3 is thin spot target value
    InsertWord Command_Packet2, temp3, 13
End Sub

'Perform the cv computations - V1.6
Private Sub ComputeCv(sensorIndex As Integer, inputCv As Double)
    Log StringFormat("ComputeCv Start values for cv = {0} and sensor = {1}", inputCv, sensorIndex)
    Dim cvSummary As Double
    cvSummary = Host_rcv(26 * (sensorIndex - 1) + S_CVS3) * 16777216#  'Compute cv from sensor data
    cvSummary = cvSummary + Host_rcv(26 * (sensorIndex - 1) + S_CVS2) * 65536#
    cvSummary = cvSummary + Host_rcv(26 * (sensorIndex - 1) + S_CVS1) * 256#
    cvSummary = cvSummary + Host_rcv(26 * (sensorIndex - 1) + S_CVS0)
    
    '128 sample periods per second
    inputCv = cvSummary / (128 * Integration_time)
    
    'Scale the result
    Log StringFormat("before scale cv = {0} and Cal_Factor = {1}", inputCv, SensorInfos(sensorIndex).Cal_Factor)
    inputCv = inputCv * SensorInfos(sensorIndex).Cal_Factor
    Log StringFormat("after scale cv = {0} and Cal_Factor = {1}", inputCv, SensorInfos(sensorIndex).Cal_Factor)
    
    Log StringFormat("cv = {0} and Cal_Factor = {1}", inputCv, SensorInfos(sensorIndex).Cal_Factor)
    If SensorInfos(sensorIndex).AverageDiameter > 1 Then        'Check for math problem V1.7
        Log StringFormat("cv = {0} and AverageDiameter = {1}", inputCv, SensorInfos(sensorIndex).AverageDiameter)
        inputCv = inputCv / SensorInfos(sensorIndex).AverageDiameter      'JW Compute as % of average denier
        Log StringFormat("After Calculations cv = {0} and AverageDiameter = {1}", inputCv, SensorInfos(sensorIndex).AverageDiameter)
    Else
        inputCv = 0
        Log "AverageDiameter = 0 and CV = 0"
    End If
    
    'save cv value for printing
    CurrentCv(sensorIndex) = inputCv
End Sub

'Display text instead of graph when BarChart can't be plotted
'if inappropriate values are assigned to important setup parameters.
'Called from Plot_Data in a loop, approx. 1x/2sec       'CM 2001-11-23
Private Sub DisableBarChart()
    'Neither Target_denier_tol nor Calibration_Denier can be zero:
    'Leads to error on scaleHeight below
    BarChart.Caption = "Bar Graph Disabled - Bad Denier Parameters"    'Note "Bar Graph" terminology similar to menu
    
    With BarChart.MyChart
        'This RESETS values that may have been set if Gain Target and
        'Denier Tolerance were valid (i.e. non-zero) earlier
        .Visible = True
        .ScaleMode = vbTwips
        .ScaleHeight = .Height
        .ScaleWidth = .Width
        .ForeColor = vbBlack
       
        'Display nasty explanatory note about Denier Parameters.
        .Cls        'Must clear due to looping; resets "line number"
        .FontBold = True
        .FontUnderline = True
        BarChart.MyChart.Print "Bar Graph Display Disabled."
        
        .FontBold = False
        .FontUnderline = False
        .FontSize = .FontSize - 2
        BarChart.MyChart.Print "Zero value is not expected for Target Gain and/or Denier Tolerance."
        .FontSize = .FontSize + 2
    End With
End Sub

'Perform a Zero calibration cycle.
Private Sub Do_ZeroCal()
    SetStatusText "Will now perform Zero calibration cycle."
    
    ZeroCal_Click
    
    'Must now send a Start data collection command to the sensors
    Command_Packet(2) = ALL_CHANNELS        'Sensor #
    Command_Packet(3) = STARTDATA_COMMAND
    PacketChecksum Command_Packet           'compute packet checksum
    mo_Comm.WriteData Command_Packet        'Xmit command packet
    
    SetStatusText "Zero calibration completed."
End Sub

Private Function FormatDiameter(singleValue As Double) As String
    FormatDiameter = Format$(singleValue, "0.000")
End Function

'Initialize the receive function.
Private Sub InitReceiveCvFunction(sensorIndex As Integer)
    Dim index As Integer
    For index = 0 To RcvBUF_LENGTH - 1              'Clear receive buffer
        Host_rcv((sensorIndex - 1) * 24 + index) = 0
    Next index
    
    SensorInfos(sensorIndex).Awaiting_Comm = True    'Set this channel flag
    RxComm.index = 0                            'Initialize RxComm parameters
    RxComm.Sensor = sensorIndex
    RxComm.error = False
    RxComm.done = False
    mo_Comm.ClearInputBuffer                    'Clear Rxcomm1 input buffer
End Sub

'This routine is executed in a loop (once per data acquisition, or every 2 seconds)
Private Sub Plot_Data()
    Dim i As Integer
    Dim iLineNumber As Integer
    Dim lColorMatrix(1 To 8) As Long
    Dim fMax(1 To 8) As Single
    Dim fMin(1 To 8) As Single
    Dim fBase As Single
    Dim fTolPlus As Single  'includes fBase
    Dim fTolMinus As Single 'includes fBase
    Dim fTol As Single      'per cent of fBase
    Dim fScaleTop As Single
    Dim fScaleHeight As Single
    
    fBase = Calibration_Denier
    fTol = Target_denier_tol / 100 * fBase
    
    If fTol = 0 Then
        'Inappropriate value for fTol; can't continue with plot routine
        DisableBarChart
        Exit Sub
    End If
        
    'Otherwise,
    BarChart.Caption = StringFormat("Bar Chart({0}) - {1} - {2}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"))
    
    fTolPlus = fBase + Target_denier_tol / 100 * fBase
    fTolMinus = fBase - Target_denier_tol / 100 * fBase
    fScaleTop = fBase + 2 * fTol
    
    fScaleHeight = 4 * fTol
    
    For i = 1 To 8
        'CM 3/19/02 Handle uninitialized sensor data,
        'i.e. when Num_Cycles > 2 is false and Process_SensorData isn't called
        If SensorInfos(i).Highest = SENSOR_HIGHEST_DEFAULT Then
            fMax(i) = Calibration_Denier
        Else
            fMax(i) = SensorInfos(i).Highest
        End If
        
        If SensorInfos(i).Lowest = SENSOR_LOWEST_DEFAULT Then
            fMin(i) = Calibration_Denier
        Else
            fMin(i) = SensorInfos(i).Lowest
        End If
        
'        fMax(i) = SensorInfos(i).Highest
'        fMin(i) = SensorInfos(i).Lowest
    Next i
    
    BarChart.MyChart.ScaleLeft = -5
    BarChart.MyChart.ScaleTop = fScaleTop
    
    BarChart.MyChart.ScaleWidth = 23
    BarChart.MyChart.ScaleHeight = -fScaleHeight
            
    lColorMatrix(1) = SensorColors(1) 'vbBlue
    lColorMatrix(2) = SensorColors(2) 'vbGreen
    lColorMatrix(3) = SensorColors(3) 'vbYellow
    lColorMatrix(4) = SensorColors(4) 'vbMagenta
    lColorMatrix(5) = SensorColors(5) 'vbCyan
    lColorMatrix(6) = SensorColors(6) 'vbWhite
    lColorMatrix(7) = SensorColors(7) 'vbRed
    lColorMatrix(8) = SensorColors(8) 'vbYellow
   
    BarChart.MyChart.CurrentY = fBase - 1.2 * fTol
    For iLineNumber = 1 To 8
        BarChart.MyChart.CurrentX = iLineNumber
        BarChart.MyChart.ForeColor = lColorMatrix(iLineNumber)
        BarChart.MyChart.Print iLineNumber;
    Next iLineNumber
    
    BarChart.MyChart.ForeColor = vbRed
    BarChart.MyChart.CurrentX = 1.5
    BarChart.MyChart.CurrentY = fBase - 1.5 * fTol
    BarChart.MyChart.Print "Sensor Numbers"
    
    BarChart.MyChart.ForeColor = vbBlack
    BarChart.MyChart.CurrentX = 11
    BarChart.MyChart.CurrentY = fBase + 1.15 * fTol
    BarChart.MyChart.Print "+ Tolerance"
    BarChart.MyChart.ForeColor = vbBlack
    BarChart.MyChart.CurrentX = 11
    BarChart.MyChart.CurrentY = fBase + 0.15 * fTol
    BarChart.MyChart.Print "  Target Gain"   'Denier"
    BarChart.MyChart.ForeColor = vbBlack
    BarChart.MyChart.CurrentX = 11
    BarChart.MyChart.CurrentY = fBase - 0.85 * fTol
    BarChart.MyChart.Print "- Tolerance"
    
    BarChart.MyChart.ForeColor = vbBlue
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase - 0.7 * fTol
    BarChart.MyChart.Print "r"
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase - 0.35 * fTol
    BarChart.MyChart.Print "e"
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase
    BarChart.MyChart.Print "i"
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase + 0.35 * fTol
    BarChart.MyChart.Print "n"
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase + 0.7 * fTol
    BarChart.MyChart.Print "e"
    BarChart.MyChart.CurrentX = -3
    BarChart.MyChart.CurrentY = fBase + 1.05 * fTol
    BarChart.MyChart.Print "D"
    
    BarChart.MyChart.ForeColor = vbRed
    BarChart.MyChart.Line (0, fTolPlus)-(10, fTolPlus)
    BarChart.MyChart.ForeColor = vbBlack
    BarChart.MyChart.Line (0, fBase)-(10, fBase)
    BarChart.MyChart.ForeColor = vbRed
    BarChart.MyChart.Line (0, fTolMinus)-(10, fTolMinus)
    
    For i = 1 To 8
'        Debug.Print i, fMax(i), i + 1, fMin(i), lColorMatrix(i)
        'BF is Fill Box.
        BarChart.MyChart.Line (i, fMax(i))-(i + 1, fMin(i)), lColorMatrix(i), BF
    Next i
End Sub

'Print Setup File (formerly mnuText_Click)
Private Sub PrintSetupFile()
    Dim i           As Integer
    Dim iTemp       As Integer
    Dim iTempP1     As Integer
    Dim iResponse   As Integer

On Error GoTo mnuText_err_hdr

    ReadSetupFile comDialog2.FileName
    Printer.Print
    Printer.Print
    Printer.Print "Setup File Name ........... " & comDialog2.FileName
    Printer.Print
    Printer.Print "Parameters:"
    Printer.Print
    Printer.Print "Line Speed .................... " & LineSpeed & " meters per minute"
    Printer.Print "Target Denier ................. " & Target_Denier & " Denier"
    Printer.Print "Target Gain ................... " & Calibration_Denier & " Denier"
    Printer.Print "Maximum Denier ................ " & Max_Denier & " Denier"
    Printer.Print "Integration time .............. " & Integration_time & " seconds"
    Printer.Print "Target denier tolerance ........" & Target_denier_tol & " per cent"
    Printer.Print "Zero Calibration Interval ..... " & ZeroCal_Interval / 60 & " minutes"
    Printer.Print "Plot Interval ................. " & Plot_Interval & " minutes"
    Printer.Print "Plot view ..................... " & PlotView(PlotViewIndex)
    ' Printer.Print "Plot channel .............. " & SystemStatus.PlotChannel
    Printer.Print "Communications Port ........ COM" & SystemStatus.ComPort
    Printer.Print "Level 1 defect tolerance ...... " & Level1_slub_tol & " per cent"
    Printer.Print "Level 1 defect length ..... " & Level1_length & " millimeters"
    Printer.Print "Level 2 defect tolerance .. " & Level2_slub_tol & " per cent"
    Printer.Print "Level 2 defect length ..... " & Level2_length & " millimeters"
    Printer.Print
    Printer.Print
    For i = 1 To LAST_SENSOR
        Printer.Print "Sensor " & i & " " & GetIniSetting("Application", "NameOfThreadLine") & " ............ " & SensorInfos(i).Package
        Printer.Print "Sensor " & i & " Enabled ............ " & SensorInfos(i).Enabled
        Printer.Print "Sensor " & i & " Zero Value ......... " & SensorInfos(i).Zero_Value & " Denier"
        Printer.Print "Sensor " & i & " Calibration Factor . " & SensorInfos(i).Cal_Factor
        Printer.Print "Sensor " & i & " Calibration Value .. " & SensorInfos(i).Cal_Value & " Denier"
        Printer.Print
    Next i
    Printer.EndDoc
    Exit Sub
mnuText_err_hdr:
    If Err.Number = 75 Then
        MsgBox "Please Open Setup File.  Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open Setup File Error, PrintSetupFile"
        Exit Sub
    End If
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open Setup File Error, PrintSetupFile"
End Sub

'Copy comm buffer contents to sensor data array and verify the Sensor data packet.
Private Sub ProcessCalData(sensorIndex As Integer)
    'Use Calibration_Denier instead of Target Denier V1.7
    Dim index As Integer
    Dim counter As Integer
    Dim value As Double
    Dim temp
    temp = mo_Comm.ReadData                         'Xfer comm buffer contents to temp buffer
    If temp(S_PIC) = sensorIndex And temp(S_PKT) = CAL_PACKET Then
        For index = 0 To 25
            Host_rcv(26 * (sensorIndex - 1) + index) = temp(index)
        Next index
        'Compute gain calibration value
        counter = Host_rcv(26 * (sensorIndex - 1) + C_DAT_HI) * 256
        counter = counter + Host_rcv(26 * (sensorIndex - 1) + C_DAT_LO)
        counter = counter - SensorInfos(sensorIndex).Zero_Value    'Correct for zero offset
        SensorInfos(sensorIndex).Cal_Value = counter
        If counter <> 0 Then
            value = Calibration_Denier / counter            'Compute calibration factor V1.7
            SensorInfos(sensorIndex).Cal_Factor = value
            'JW 5/25/00 Format calibration display
            Text1(sensorIndex).Text = "Cal factor = " & Format(value, "##.0000")
        Else
            Text1(sensorIndex).Text = "Cal Error"
        End If
    Else
        Text1(sensorIndex).Text = "Comm Err"
    End If
End Sub

'Copy comm buffer contents to sensor data array and verify the Sensor data packet.
'Return Average value (float), which is written to file.
Private Function GetSensorReportData(sensorIndex As Integer) As SensorReportData
On Error Resume Next

    Dim index           As Integer
#If SCALE_DIAMETER Then
    Dim localAverage    As Double
#Else
    Dim localAverage    As Long
#End If

    Dim floatAverage    As Double
    Dim cvSummary       As Double
    Dim cvValue         As Double 'V1.6
    Dim temp
    temp = mo_Comm.ReadData             'Xfer comm buffer contents to temp buffer
    If temp(1) = sensorIndex And temp(2) = SEND_DATA Then
        For index = 0 To 25
            Host_rcv(26 * (sensorIndex - 1) + index) = temp(index)
        Next index
  
        ' accumulate the level1 and level2 slub counts
        With SensorInfos(sensorIndex)
            .Level1_Slub = .Level1_Slub + temp(S_L1S)
            .Level2_Slub = .Level2_Slub + temp(S_L2S)
        End With
        
        'display average diameter, min & max denier, and cv
        'Integer and floating-point averages are computed in order to save compute time.
        localAverage = (Host_rcv(26 * (sensorIndex - 1) + S_DAT_HI) * 256) + Host_rcv(26 * (sensorIndex - 1) + S_DAT_LO)
        'JW 5/25/00 Store raw sensor average denier
        SensorInfos(sensorIndex).Raw_Denier = localAverage
        localAverage = localAverage - SensorInfos(sensorIndex).Zero_Value         'Correct for offset
        floatAverage = localAverage * SensorInfos(sensorIndex).Cal_Factor   'Scale the result
        localAverage = floatAverage
        'save floatAverage for current printing
        CurrentAverage(sensorIndex) = floatAverage
        SensorInfos(sensorIndex).AverageDiameter = localAverage         'save the average diameter V1.6
        SensorInfos(sensorIndex).SumOfAverages = SensorInfos(sensorIndex).SumOfAverages + localAverage
        
        If localAverage < SensorInfos(sensorIndex).Lowest Then       'Update min/max values
            SensorInfos(sensorIndex).Lowest = localAverage
        End If
        
        If localAverage > SensorInfos(sensorIndex).Highest Then
            SensorInfos(sensorIndex).Highest = localAverage
        End If
        
        If COMPUTE_CV Then                                     'V1.6 start cv code
            Call ComputeCv(sensorIndex, cvValue)
        End If                                                 'V1.6 end code
    
        #If SCALE_DIAMETER Then
        Text1(sensorIndex).Text = "Now = " & FormatDiameter(floatAverage) & vbCrLf & _
                                  "Max. = " & FormatDiameter(SensorInfos(sensorIndex).Highest) & vbCrLf & _
                                  "Min. = " & FormatDiameter(SensorInfos(sensorIndex).Lowest) & vbCrLf & _
                                  "Mean = " & FormatDiameter(SensorInfos(sensorIndex).SumOfAverages / (SensorInfos(sensorIndex).Num_Cycles - 2)) & vbCrLf & _
                                  "CV = " & Format(cvValue, "#0.0") & "%" & vbCrLf & _
                                  "L1slub = " & temp(S_L1S) & vbCrLf & _
                                  "L2slub = " & temp(S_L2S) & vbCrLf & _
                                  " DELTA = " & FormatDiameter(floatAverage)
                                
        #Else
        'JW Corrected calculations by subtracting 2 from the number of int. periods
        Text1(sensorIndex).Text = "Now = " & Format(floatAverage, "#00.0") & vbCrLf & _
                                  "Max. = " & SensorInfos(sensorIndex).Highest & vbCrLf & _
                                  "Min. = " & SensorInfos(sensorIndex).Lowest & vbCrLf & _
                                  "Mean = " & Format(SensorInfos(sensorIndex).SumOfAverages / (SensorInfos(sensorIndex).Num_Cycles - 2), "#00.0") & vbCrLf & _
                                  "CV = " & Format(cvValue, "#0.0") & "%" & vbCrLf & _
                                  "L1slub = " & temp(S_L1S) & vbCrLf & _
                                  "L2slub = " & temp(S_L2S) & vbCrLf & _
                                  "DELTA = " & Format(floatAverage, "#00.0")
        #End If
        
        'save current mean value for printing
        CurrentMean(sensorIndex) = SensorInfos(sensorIndex).SumOfAverages / (SensorInfos(sensorIndex).Num_Cycles - 2)
        
        ' <<<<<<<<<< Test Only !!! >>>>>>>>>>
        '  Print #1, Tab; Format(floatAverage, "####.0");
        '  Print #1, Tab; Format(floatAverage, "####.0");
        ' <<<<<<<<<< Test Only !!! >>>>>>>>>>
    
        'Graph the current channel
        If iPLotChannels(sensorIndex) = 1 Then
            iPlotDenier(iPLot, 0) = localAverage
            CWGraph1.Plots.Item(iPLot).LineColor = SensorColors(sensorIndex)
            iPLot = iPLot + 1
        End If
  
        'Check limits V1.7
        If (localAverage <= Target_Denier + Target_Denier * (Target_denier_tol / 100)) And (localAverage >= Target_Denier - Target_Denier * (Target_denier_tol / 100)) Then
            Text1(sensorIndex).BackColor = RGB(216, 216, 216)   ' RGB_SILVER
            Text1(sensorIndex).ForeColor = RGB(12, 12, 12)      ' Black
        Else
            Text1(sensorIndex).BackColor = RGB(255, 75, 75)     ' Red
            Text1(sensorIndex).BackColor = RGB(242, 242, 242)   ' light gray
        End If
    Else
        Text1(sensorIndex).Text = "Comm Err"
    End If
    'Debug.Print "floatAverage " & CStr(floatAverage)
    'Debug.Print "cvValue " & CStr(cvValue)
    
    Dim currentSensorReportData As New SensorReportData
    With currentSensorReportData
        .Average = floatAverage
        .Cv = cvValue
    End With
    
    Debug.Print "floatAverage " & CStr(currentSensorReportData.Average)
    Debug.Print "cvValue " & CStr(currentSensorReportData.Cv)
    Set GetSensorReportData = currentSensorReportData
End Function

'Copy comm buffer contents to sensor data array and
'verify the Sensor initialization packet.
Private Sub ProcessSensorInit(sensorIndex As Integer)
    Dim index As Integer
    Dim temp
    temp = mo_Comm.ReadData                 'Xfer comm buffer contents to temp buffer
    If temp(1) = sensorIndex And temp(2) = INIT_COMPLETE Then
        For index = 0 To 25
            Host_rcv(26 * (sensorIndex - 1) + index) = temp(index)
        Next index
        SensorInfos(sensorIndex).Online = True
        Text1(sensorIndex).Text = "Online"       'Update display button
        Text1(sensorIndex).BackColor = RGB(216, 216, 216)   ' RGB_SILVER
        Text1(sensorIndex).ForeColor = RGB(12, 12, 12)      ' Black
    Else
        SensorInfos(sensorIndex).Online = False
        Text1(sensorIndex).Text = "No Comm"
        Command(sensorIndex).BackColor = RGB(255, 75, 75)     ' Red
        Text1(sensorIndex).ForeColor = RGB(242, 242, 242)    ' light gray
    End If
End Sub

'Copy comm buffer contents to sensor data array and
'verify the Sensor data packet.
Private Sub ProcessZeroData(sensorIndex As Integer)
    Dim index As Integer
    Dim offset As Integer
    Dim temp
    temp = mo_Comm.ReadData                 'Xfer comm buffer contents to temp buffer
    If temp(S_PIC) = sensorIndex And temp(S_PKT) = ZERO_PACKET Then
        For index = 0 To 25
            Host_rcv(26 * (sensorIndex - 1) + index) = temp(index)
        Next index
        
        'display Zero value
        offset = Host_rcv(26 * (sensorIndex - 1) + Z_DAT_HI) * 256
        offset = offset + Host_rcv(26 * (sensorIndex - 1) + Z_DAT_LO)
        SensorInfos(sensorIndex).Zero_Value = offset
        Text1(sensorIndex).Text = "Zero offset = " & offset
    Else
        Text1(sensorIndex).Text = "Comm Err"
    End If
End Sub

'Read setup information from disk in s_FileName.  See FTSetup.
Private Function ReadSetupFile(inputFileName As String) As Boolean
    Dim setupFile As New FTSetup
    Dim index As Integer
    If Not setupFile.ReadFromFile(inputFileName) Then
        ReadSetupFile = False
        Exit Function
    End If
    
    'Read information from FTSetup class into global variables
    With setupFile
        SystemStatus.ComPort = .ComPort
        ComPortIndex = .ComPortIndex
        Calibration_Denier = .DenierCalibration
        Max_Denier = .DenierMax
        DenierRangeIndex = .DenierRangeIndex
        Target_Denier = .DenierTarget
        Target_denier_tol = .DenierTargetTolerance
        Integration_time = .IntegrationTime
        Level1_length = CLng(.Level1Length)
        Level1_slub_tol = CLng(.Level1SlubTolerance)
        Level2_length = CLng(.Level2Length)
        Level2_slub_tol = CLng(.Level2SlubTolerance)
        LineSpeed = .LineSpeed
        SystemStatus.PlotChannel = .PlotChannel
        Plot_Interval = .PlotInterval
        PlotViewIndex = .PlotViewIndex
        ZeroCal_Interval = .ZeroCalInterval
        
        For index = 1 To LAST_SENSOR
            SensorInfos(index).Cal_Factor = .SensorCalFactor(index)
            SensorInfos(index).Cal_Value = .SensorCalValue(index)
            SensorInfos(index).Enabled = .SensorEnabled(index)
            SensorInfos(index).Zero_Value = .SensorZeroValue(index)
        Next index
    End With

    ReadSetupFile = True
End Function

'Save current report (formerly mnuCurrentReport_Click)
Private Sub SaveReport()
    Dim iSensor         As Integer
    Dim i               As Integer
    Dim iTemp           As Integer
    Dim iTempP1         As Integer
    Dim iResponse       As Integer
    Dim iFileNamesNbr   As Integer

On Error GoTo Current_Err_Hdr

    comDialog3.CancelError = True

    iFileNamesNbr = FreeFile

    comDialog3.DialogTitle = "Print " & ReportType & " Report"
    comDialog3.Filter = "Text File (*.txt)| *.txt"
    comDialog3.FilterIndex = 1
    ' Save Flags
    '   cdlOFNExplorer              = &H80000&
    '   cdlOFNCreatePrompt          = &H02000&
    '   cdlOFNNoChangeDir           = &H00008&
    '   cdlOFNHideReadOnly          = &H00004&
    '   cdlOFNOverwritePrompt       = &H00002&
    
    comDialog3.Flags = &H80200E
    ' comDialog3.Flags = cdlOFNCreatePrompt & cdlOFNExplorer & _
             cdlOFNOverwritePrompt & cdlOFNNoChangeDir
         
    comDialog3.ShowSave
    Open comDialog3.FileName For Output As #iFileNamesNbr
    
    Print #iFileNamesNbr, "    STS Denier Monitoring " & ReportType & " Report  - "; Now
    Print #iFileNamesNbr, "    ----------------------------------------------------------------"
    Print #iFileNamesNbr,
    Print #iFileNamesNbr, "IntegrationTime ="; Integration_time; "Sec "; "Target Denier ="; Target_Denier;
    Print #iFileNamesNbr, "LineSpeed ="; LineSpeed; "Meters/sec"
    Print #iFileNamesNbr,
    For iSensor = 1 To LAST_SENSOR
        If SensorInfos(iSensor).Enabled = True Then
            Print #iFileNamesNbr, "Sensor "; iSensor
            If ReportType = "Current" Then
                Print #iFileNamesNbr, Tab(10); "Current Value = "; Format(CurrentAverage(iSensor), "#00.0")
            End If
            
            Print #iFileNamesNbr, Tab(10); "Maximum Value = "; SensorInfos(iSensor).Highest
            Print #iFileNamesNbr, Tab(10); "Mean Value = "; Format(CurrentMean(iSensor), "#00.0")
            Print #iFileNamesNbr, Tab(10); "Minimum Value = "; SensorInfos(iSensor).Lowest
            Print #iFileNamesNbr, Tab(10); "CV Value = "; Format(CurrentCv(iSensor), "#00.0")
            Print #iFileNamesNbr, Tab(10); "Level 1 Defects = "; SensorInfos(iSensor).Level1_Slub
            Print #iFileNamesNbr, Tab(10); "Level 2 Defects = "; SensorInfos(iSensor).Level2_Slub
        End If
    Next iSensor
  
    Close #iFileNamesNbr
    
    Exit Sub
Current_Err_Hdr:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source & ".  Filename = " & comDialog2.FileName, vbExclamation, "File Error FiberTrack mnuCurrentReport"
    Exit Sub
End Sub

'Save setup information to disk as fileName.  See FTSetup
Private Function SaveSetupFile(FileName As String) As Boolean
    Dim currentSetup As New FTSetup
    Dim index As Integer
    
    'Move information from global variables to FTSetup class
    With currentSetup
        .ComPort = SystemStatus.ComPort
        .ComPortIndex = ComPortIndex
        .DenierCalibration = Calibration_Denier
        .DenierMax = Max_Denier
        .DenierRangeIndex = DenierRangeIndex
        .DenierTarget = Target_Denier
        .DenierTargetTolerance = Target_denier_tol
        .IntegrationTime = Integration_time
        .Level1Length = CLng(Level1_length)
        .Level1SlubTolerance = CLng(Level1_slub_tol)
        .Level2Length = CLng(Level2_length)
        .Level2SlubTolerance = CLng(Level2_slub_tol)
        .LineSpeed = LineSpeed
        .PlotChannel = SystemStatus.PlotChannel
        .PlotInterval = Plot_Interval
        .PlotViewIndex = PlotViewIndex
        .ZeroCalInterval = ZeroCal_Interval
        
        For index = 1 To LAST_SENSOR
            .SensorCalFactor(index) = SensorInfos(index).Cal_Factor
            .SensorCalValue(index) = SensorInfos(index).Cal_Value
            .SensorEnabled(index) = SensorInfos(index).Enabled
            .SensorZeroValue(index) = SensorInfos(index).Zero_Value
        Next index
    End With

    'Tell FTSetup class to write itself to disk
    SaveSetupFile = currentSetup.WriteToFile(FileName)
End Function

'Wrapper for status information display.
Public Sub SetStatusText(statusText As String)
    Text17.Text = statusText
    Debug.Print "STATUS: " & statusText
End Sub

'Updates the system runtime to the front-panel display
Public Sub SetTime(inputHours As Long, inputMinutes As Long, inputSeconds As Long)
    Dim temp
    temp = DateDiff("s", SystemStatus.StartTime, Now)      'Computes elapsed seconds
    inputHours = temp \ 3600
    inputMinutes = (temp Mod 3600) \ 60
    inputSeconds = (temp Mod 3600) Mod 60
End Sub
