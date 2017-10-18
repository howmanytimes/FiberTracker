VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PlotData 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Plot Data"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11250
   Icon            =   "PlotData.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSensorCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Please check "
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   9840
      TabIndex        =   36
      ToolTipText     =   "Please check the sensors to plot."
      Top             =   960
      Width           =   1335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "O&K"
         Height          =   375
         Left            =   360
         TabIndex        =   45
         Top             =   3240
         Width           =   495
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   44
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkSensor 
         Caption         =   "Sensor 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
   End
   Begin CWUIControlsLib.CWGraph cwgLineChart 
      Height          =   3375
      Left            =   840
      TabIndex        =   30
      Top             =   960
      Width           =   8895
      _Version        =   196608
      _ExtentX        =   15690
      _ExtentY        =   5953
      _StockProps     =   71
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
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
      C[0]_1          =   16777215
      Event_1         =   4
      ClassName_4     =   "CCWGFPlotEvent"
      Owner_4         =   1
      Plots_1         =   5
      ClassName_5     =   "CCWDataPlots"
      Array_5         =   4
      Editor_5        =   6
      ClassName_6     =   "CCWGFPlotArrayEditor"
      Owner_6         =   1
      Array[0]_5      =   7
      ClassName_7     =   "CCWDataPlot"
      opts_7          =   4194335
      Name_7          =   "Plot-1"
      Bindings_7      =   0
      C[0]_7          =   255
      C[1]_7          =   65280
      C[2]_7          =   16711680
      C[3]_7          =   16711680
      Event_7         =   4
      X_7             =   8
      ClassName_8     =   "CCWAxis"
      opts_8          =   543
      Name_8          =   "Time - Seconds"
      Bindings_8      =   0
      Orientation_8   =   2944
      format_8        =   9
      ClassName_9     =   "CCWFormat"
      Scale_8         =   10
      ClassName_10    =   "CCWScale"
      opts_10         =   90112
      Bindings_10     =   0
      rMin_10         =   25
      rMax_10         =   582
      dMax_10         =   20
      discInterval_10 =   1
      Radial_8        =   0
      Enum_8          =   11
      ClassName_11    =   "CCWEnum"
      Editor_11       =   12
      ClassName_12    =   "CCWEnumArrayEditor"
      Owner_12        =   8
      Font_8          =   0
      tickopts_8      =   1679
      base_8          =   2
      major_8         =   10
      minor_8         =   5
      Caption_8       =   13
      ClassName_13    =   "CCWDrawObj"
      opts_13         =   30
      Bindings_13     =   0
      C[0]_13         =   -2147483640
      Image_13        =   14
      ClassName_14    =   "CCWTextImage"
      Bindings_14     =   0
      font_14         =   0
      Animator_13     =   0
      Blinker_13      =   0
      Y_7             =   15
      ClassName_15    =   "CCWAxis"
      opts_15         =   1567
      Name_15         =   "Denier"
      Bindings_15     =   0
      Orientation_15  =   2067
      format_15       =   16
      ClassName_16    =   "CCWFormat"
      Scale_15        =   17
      ClassName_17    =   "CCWScale"
      opts_17         =   122880
      Bindings_17     =   0
      rMin_17         =   11
      rMax_17         =   198
      dMax_17         =   10
      discInterval_17 =   1
      Radial_15       =   0
      Enum_15         =   18
      ClassName_18    =   "CCWEnum"
      Editor_18       =   19
      ClassName_19    =   "CCWEnumArrayEditor"
      Owner_19        =   15
      Font_15         =   0
      tickopts_15     =   1679
      major_15        =   10
      minor_15        =   5
      Caption_15      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   30
      Bindings_20     =   0
      C[0]_20         =   -2147483640
      Image_20        =   21
      ClassName_21    =   "CCWTextImage"
      Bindings_21     =   0
      font_21         =   0
      Animator_20     =   0
      Blinker_20      =   0
      LineStyle_7     =   1
      LineWidth_7     =   2
      BasePlot_7      =   0
      DefaultXInc_7   =   1
      DefaultPlotPerRow_7=   -1  'True
      Array[1]_5      =   22
      ClassName_22    =   "CCWDataPlot"
      opts_22         =   4194335
      Name_22         =   "Plot-2"
      Bindings_22     =   0
      C[0]_22         =   16711680
      C[1]_22         =   255
      C[2]_22         =   16711680
      C[3]_22         =   16776960
      Event_22        =   4
      X_22            =   8
      Y_22            =   15
      LineStyle_22    =   1
      LineWidth_22    =   2
      BasePlot_22     =   0
      DefaultXInc_22  =   1
      DefaultPlotPerRow_22=   -1  'True
      Array[2]_5      =   23
      ClassName_23    =   "CCWDataPlot"
      opts_23         =   4194335
      Name_23         =   "Plot-3"
      Bindings_23     =   0
      C[0]_23         =   255
      C[1]_23         =   255
      C[2]_23         =   16711680
      C[3]_23         =   16776960
      Event_23        =   4
      X_23            =   8
      Y_23            =   15
      LineStyle_23    =   1
      LineWidth_23    =   2
      BasePlot_23     =   0
      DefaultXInc_23  =   1
      DefaultPlotPerRow_23=   -1  'True
      Array[3]_5      =   24
      ClassName_24    =   "CCWDataPlot"
      opts_24         =   4194335
      Name_24         =   "Plot-4"
      Bindings_24     =   0
      C[0]_24         =   255
      C[1]_24         =   65280
      C[2]_24         =   16711680
      C[3]_24         =   16711680
      Event_24        =   4
      X_24            =   8
      Y_24            =   15
      LineStyle_24    =   1
      LineWidth_24    =   2
      BasePlot_24     =   0
      DefaultXInc_24  =   1
      DefaultPlotPerRow_24=   -1  'True
      Axes_1          =   25
      ClassName_25    =   "CCWAxes"
      Array_25        =   2
      Editor_25       =   26
      ClassName_26    =   "CCWGFAxisArrayEditor"
      Owner_26        =   1
      Array[0]_25     =   8
      Array[1]_25     =   15
      DefaultPlot_1   =   27
      ClassName_27    =   "CCWDataPlot"
      opts_27         =   4194335
      Name_27         =   "[Template]"
      Bindings_27     =   0
      C[0]_27         =   255
      C[1]_27         =   65280
      C[2]_27         =   16711680
      C[3]_27         =   16711680
      Event_27        =   4
      X_27            =   8
      Y_27            =   15
      LineStyle_27    =   1
      LineWidth_27    =   2
      BasePlot_27     =   0
      DefaultXInc_27  =   1
      DefaultPlotPerRow_27=   -1  'True
      Cursors_1       =   28
      ClassName_28    =   "CCWCursors"
      Array_28        =   2
      Editor_28       =   29
      ClassName_29    =   "CCWGFCursorArrayEditor"
      Owner_29        =   1
      Array[0]_28     =   30
      ClassName_30    =   "CCWCursor"
      opts_30         =   31
      Name_30         =   "Cursor-1"
      Bindings_30     =   0
      C[0]_30         =   255
      Event_30        =   4
      X_30            =   8
      Y_30            =   15
      XPos_30         =   2
      YPos_30         =   1
      PointIndex_30   =   -1
      ChrosshairStyle_30=   8
      LockPlot_30     =   0
      Array[1]_28     =   31
      ClassName_31    =   "CCWCursor"
      opts_31         =   31
      Name_31         =   "Cursor-2"
      Bindings_31     =   0
      C[0]_31         =   16711680
      Event_31        =   4
      X_31            =   8
      Y_31            =   15
      XPos_31         =   4
      YPos_31         =   2
      PointIndex_31   =   -1
      ChrosshairStyle_31=   8
      LockPlot_31     =   0
      TrackMode_1     =   6
      GraphBackground_1=   0
      GraphFrame_1    =   32
      ClassName_32    =   "CCWDrawObj"
      opts_32         =   30
      Bindings_32     =   0
      Image_32        =   33
      ClassName_33    =   "CCWPictImage"
      opts_33         =   1280
      Bindings_33     =   0
      Rows_33         =   1
      Cols_33         =   1
      F_33            =   -2147483633
      B_33            =   -2147483633
      ColorReplaceWith_33=   8421504
      ColorReplace_33 =   8421504
      Tolerance_33    =   2
      Animator_32     =   0
      Blinker_32      =   0
      PlotFrame_1     =   34
      ClassName_34    =   "CCWDrawObj"
      opts_34         =   30
      Bindings_34     =   0
      C[1]_34         =   16777215
      Image_34        =   35
      ClassName_35    =   "CCWPictImage"
      opts_35         =   1280
      Bindings_35     =   0
      Rows_35         =   1
      Cols_35         =   1
      Pict_35         =   1
      F_35            =   -2147483633
      B_35            =   16777215
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
      font_37         =   0
      Animator_36     =   0
      Blinker_36      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
   End
   Begin VB.ComboBox cboomPlotData 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   7560
      Width           =   735
   End
   Begin MSComDlg.CommonDialog comDialog1 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "Print S&creen"
      Height          =   495
      Left            =   4440
      TabIndex        =   46
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlotData 
      Caption         =   "&Plot Data File"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "&Zoom Out"
      Height          =   495
      Left            =   5640
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame fraCursorMeas 
      Caption         =   "Cursor Measurements"
      Height          =   975
      Left            =   7560
      TabIndex        =   25
      Top             =   5040
      Width           =   1935
      Begin VB.TextBox PeriodVal 
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox AmplitudeVal 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Amplitude:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Period:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.TextBox YPosition 
      Height          =   285
      Left            =   8640
      TabIndex        =   24
      Top             =   6840
      Width           =   735
   End
   Begin VB.Frame fraCursorInfo 
      Caption         =   "Cursor Information"
      Height          =   975
      Left            =   7560
      TabIndex        =   20
      Top             =   6240
      Width           =   1935
      Begin VB.TextBox XPosition 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "X Position:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Y Position:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraGraphOps 
      Caption         =   "Graph Operations"
      Height          =   1335
      Left            =   4920
      TabIndex        =   18
      Top             =   5880
      Width           =   2415
      Begin VB.OptionButton optCursorCood 
         Caption         =   "Cursor Coodinates"
         Height          =   195
         Left            =   360
         TabIndex        =   50
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optPan 
         Caption         =   "Pan"
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "Zoom"
         Height          =   195
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Width           =   735
      End
      Begin CWUIControlsLib.CWSlide GraphMode 
         Height          =   975
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   1720
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Reset_0         =   0   'False
         CompatibleVers_0=   196608
         Slider_0        =   1
         ClassName_1     =   "CCWSlider"
         opts_1          =   2078
         Bindings_1      =   2
         ClassName_2     =   "CCWBindingHolderArray"
         Editor_2        =   3
         ClassName_3     =   "CCWBindingHolderArrayEditor"
         Owner_3         =   1
         C[0]_1          =   -2147483643
         BGImg_1         =   4
         ClassName_4     =   "CCWDrawObj"
         opts_4          =   30
         Bindings_4      =   0
         Image_4         =   5
         ClassName_5     =   "CCWPictImage"
         opts_5          =   1280
         Bindings_5      =   0
         Rows_5          =   1
         Cols_5          =   1
         Pict_5          =   286
         F_5             =   -2147483633
         B_5             =   -2147483633
         ColorReplaceWith_5=   8421504
         ColorReplace_5  =   8421504
         Tolerance_5     =   2
         Animator_4      =   0
         Blinker_4       =   0
         BFImg_1         =   6
         ClassName_6     =   "CCWDrawObj"
         opts_6          =   62
         Bindings_6      =   0
         Image_6         =   7
         ClassName_7     =   "CCWPictImage"
         opts_7          =   1280
         Bindings_7      =   0
         Rows_7          =   1
         Cols_7          =   1
         Pict_7          =   286
         F_7             =   -2147483633
         B_7             =   -2147483633
         ColorReplaceWith_7=   8421504
         ColorReplace_7  =   8421504
         Tolerance_7     =   2
         Animator_6      =   0
         Blinker_6       =   0
         Label_1         =   8
         ClassName_8     =   "CCWDrawObj"
         opts_8          =   30
         Bindings_8      =   0
         C[0]_8          =   -2147483640
         Image_8         =   9
         ClassName_9     =   "CCWTextImage"
         Bindings_9      =   0
         style_9         =   15878208
         font_9          =   0
         Animator_8      =   0
         Blinker_8       =   0
         Border_1        =   10
         ClassName_10    =   "CCWDrawObj"
         opts_10         =   28
         Bindings_10     =   0
         Image_10        =   11
         ClassName_11    =   "CCWPictImage"
         opts_11         =   1280
         Bindings_11     =   0
         Rows_11         =   1
         Cols_11         =   1
         Pict_11         =   25
         F_11            =   -2147483633
         B_11            =   -2147483633
         ColorReplaceWith_11=   8421504
         ColorReplace_11 =   8421504
         Tolerance_11    =   2
         Animator_10     =   0
         Blinker_10      =   0
         FillBound_1     =   12
         ClassName_12    =   "CCWGuiObject"
         opts_12         =   28
         Bindings_12     =   0
         FillTok_1       =   13
         ClassName_13    =   "CCWGuiObject"
         opts_13         =   30
         Bindings_13     =   0
         Axis_1          =   14
         ClassName_14    =   "CCWAxis"
         opts_14         =   1055
         Name_14         =   "Axis"
         Bindings_14     =   0
         Orientation_14  =   133523
         format_14       =   15
         ClassName_15    =   "CCWFormat"
         Scale_14        =   16
         ClassName_16    =   "CCWScale"
         opts_16         =   24576
         Bindings_16     =   0
         rMin_16         =   10
         rMax_16         =   54
         dMin_16         =   1
         dMax_16         =   3
         discInterval_16 =   1
         Radial_14       =   0
         Enum_14         =   17
         ClassName_17    =   "CCWEnum"
         Array_17        =   3
         Editor_17       =   18
         ClassName_18    =   "CCWEnumArrayEditor"
         Owner_18        =   14
         Array[0]_17     =   19
         ClassName_19    =   "CCWEnumElt"
         opts_19         =   1
         Name_19         =   "Zoom"
         Bindings_19     =   0
         DrawList_19     =   0
         varVarType_19   =   2
         Array[1]_17     =   20
         ClassName_20    =   "CCWEnumElt"
         opts_20         =   1
         Name_20         =   "Pan"
         Bindings_20     =   0
         DrawList_20     =   0
         varVarType_20   =   2
         var_Val_20      =   1
         Array[2]_17     =   21
         ClassName_21    =   "CCWEnumElt"
         opts_21         =   1
         Name_21         =   "Cursor Coordinates"
         Bindings_21     =   0
         DrawList_21     =   0
         varVarType_21   =   2
         var_Val_21      =   2
         Font_14         =   0
         tickopts_14     =   2718
         Caption_14      =   22
         ClassName_22    =   "CCWDrawObj"
         opts_22         =   30
         Bindings_22     =   0
         C[0]_22         =   -2147483640
         Image_22        =   23
         ClassName_23    =   "CCWTextImage"
         Bindings_23     =   0
         font_23         =   0
         Animator_22     =   0
         Blinker_22      =   0
         DrawLst_1       =   24
         ClassName_24    =   "CDrawList"
         count_24        =   10
         list[10]_24     =   10
         list[9]_24      =   25
         ClassName_25    =   "CCWThumb"
         opts_25         =   31
         Name_25         =   "Pointer-1"
         Bindings_25     =   0
         C[0]_25         =   8388608
         C[1]_25         =   8388608
         C[2]_25         =   -2147483635
         Image_25        =   26
         ClassName_26    =   "CCWPictImage"
         opts_26         =   1280
         Bindings_26     =   0
         Rows_26         =   1
         Cols_26         =   1
         Pict_26         =   213
         F_26            =   8388608
         B_26            =   8388608
         ColorReplaceWith_26=   8421504
         ColorReplace_26 =   8421504
         Tolerance_26    =   2
         Animator_25     =   0
         Blinker_25      =   0
         style_25        =   1
         Value_25        =   1
         list[8]_24      =   14
         list[7]_24      =   8
         list[6]_24      =   13
         list[5]_24      =   6
         list[4]_24      =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   30
         Bindings_27     =   0
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
         Bindings_28     =   0
         Rows_28         =   1
         Cols_28         =   1
         Pict_28         =   7
         F_28            =   -2147483633
         B_28            =   -2147483633
         ColorReplaceWith_28=   8421504
         ColorReplace_28 =   8421504
         Tolerance_28    =   2
         Animator_27     =   0
         Blinker_27      =   0
         list[3]_24      =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   30
         Bindings_29     =   0
         Image_29        =   30
         ClassName_30    =   "CCWPictImage"
         opts_30         =   1280
         Bindings_30     =   0
         Rows_30         =   1
         Cols_30         =   1
         Pict_30         =   96
         F_30            =   -2147483633
         B_30            =   -2147483633
         ColorReplaceWith_30=   8421504
         ColorReplace_30 =   8421504
         Tolerance_30    =   2
         Animator_29     =   0
         Blinker_29      =   0
         list[2]_24      =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   30
         Bindings_31     =   0
         Image_31        =   32
         ClassName_32    =   "CCWPictImage"
         opts_32         =   1280
         Bindings_32     =   0
         Rows_32         =   1
         Cols_32         =   1
         Pict_32         =   95
         F_32            =   -2147483633
         B_32            =   -2147483633
         ColorReplaceWith_32=   8421504
         ColorReplace_32 =   8421504
         Tolerance_32    =   2
         Animator_31     =   0
         Blinker_31      =   0
         list[1]_24      =   4
         IncDec_1        =   0
         Ptrs_1          =   33
         ClassName_33    =   "CCWPointerArray"
         Array_33        =   1
         Editor_33       =   34
         ClassName_34    =   "CCWPointerArrayEditor"
         Owner_34        =   1
         Array[0]_33     =   25
         Stats_1         =   35
         ClassName_35    =   "CCWStats"
         Bindings_35     =   0
         doInc_1         =   31
         doDec_1         =   29
         doFrame_1       =   27
      End
   End
   Begin VB.CommandButton cmdBarChart 
      Caption         =   "&Bar Chart"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdLineChart 
      Caption         =   "&Line Chart"
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin CWUIControlsLib.CWGraph cwgBarChart 
      Height          =   2295
      Left            =   1440
      TabIndex        =   15
      Top             =   840
      Width           =   6615
      _Version        =   196608
      _ExtentX        =   11668
      _ExtentY        =   4048
      _StockProps     =   71
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
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
      C[0]_1          =   16777215
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
      opts_7          =   7340063
      Name_7          =   "Plot-1"
      Bindings_7      =   0
      C[0]_7          =   16776960
      C[1]_7          =   255
      C[2]_7          =   16711680
      C[3]_7          =   16776960
      Event_7         =   4
      X_7             =   8
      ClassName_8     =   "CCWAxis"
      opts_8          =   1567
      Name_8          =   "Time - Seconds"
      Bindings_8      =   0
      Orientation_8   =   2944
      format_8        =   9
      ClassName_9     =   "CCWFormat"
      Scale_8         =   10
      ClassName_10    =   "CCWScale"
      opts_10         =   90112
      Bindings_10     =   0
      rMin_10         =   25
      rMax_10         =   430
      dMax_10         =   10
      discInterval_10 =   1
      Radial_8        =   0
      Enum_8          =   11
      ClassName_11    =   "CCWEnum"
      Editor_11       =   12
      ClassName_12    =   "CCWEnumArrayEditor"
      Owner_12        =   8
      Font_8          =   0
      tickopts_8      =   1679
      base_8          =   2
      major_8         =   10
      minor_8         =   5
      Caption_8       =   13
      ClassName_13    =   "CCWDrawObj"
      opts_13         =   30
      Bindings_13     =   0
      C[0]_13         =   -2147483640
      Image_13        =   14
      ClassName_14    =   "CCWTextImage"
      Bindings_14     =   0
      font_14         =   0
      Animator_13     =   0
      Blinker_13      =   0
      Y_7             =   15
      ClassName_15    =   "CCWAxis"
      opts_15         =   1567
      Name_15         =   "Denier"
      Bindings_15     =   0
      Orientation_15  =   2067
      format_15       =   16
      ClassName_16    =   "CCWFormat"
      Scale_15        =   17
      ClassName_17    =   "CCWScale"
      opts_17         =   122880
      Bindings_17     =   0
      rMin_17         =   11
      rMax_17         =   126
      dMax_17         =   10
      discInterval_17 =   1
      Radial_15       =   0
      Enum_15         =   18
      ClassName_18    =   "CCWEnum"
      Editor_18       =   19
      ClassName_19    =   "CCWEnumArrayEditor"
      Owner_19        =   15
      Font_15         =   0
      tickopts_15     =   1679
      major_15        =   10
      minor_15        =   5
      Caption_15      =   20
      ClassName_20    =   "CCWDrawObj"
      opts_20         =   30
      Bindings_20     =   0
      C[0]_20         =   -2147483640
      Image_20        =   21
      ClassName_21    =   "CCWTextImage"
      Bindings_21     =   0
      style_21        =   1
      font_21         =   0
      Animator_20     =   0
      Blinker_20      =   0
      LineStyle_7     =   2
      LineWidth_7     =   1
      BasePlot_7      =   0
      DefaultXInc_7   =   1
      DefaultPlotPerRow_7=   -1  'True
      Axes_1          =   22
      ClassName_22    =   "CCWAxes"
      Array_22        =   2
      Editor_22       =   23
      ClassName_23    =   "CCWGFAxisArrayEditor"
      Owner_23        =   1
      Array[0]_22     =   8
      Array[1]_22     =   15
      DefaultPlot_1   =   24
      ClassName_24    =   "CCWDataPlot"
      opts_24         =   7340063
      Name_24         =   "[Template]"
      Bindings_24     =   0
      C[0]_24         =   16776960
      C[1]_24         =   255
      C[2]_24         =   16711680
      C[3]_24         =   16776960
      Event_24        =   4
      X_24            =   8
      Y_24            =   15
      LineStyle_24    =   2
      LineWidth_24    =   1
      BasePlot_24     =   0
      DefaultXInc_24  =   1
      DefaultPlotPerRow_24=   -1  'True
      Cursors_1       =   25
      ClassName_25    =   "CCWCursors"
      Array_25        =   2
      Editor_25       =   26
      ClassName_26    =   "CCWGFCursorArrayEditor"
      Owner_26        =   1
      Array[0]_25     =   27
      ClassName_27    =   "CCWCursor"
      opts_27         =   31
      Name_27         =   "Cursor-1"
      Bindings_27     =   0
      C[0]_27         =   255
      Event_27        =   4
      X_27            =   8
      Y_27            =   15
      XPos_27         =   1
      YPos_27         =   1
      PointIndex_27   =   -1
      ChrosshairStyle_27=   8
      LockPlot_27     =   0
      Array[1]_25     =   28
      ClassName_28    =   "CCWCursor"
      opts_28         =   31
      Name_28         =   "Cursor-2"
      Bindings_28     =   0
      C[0]_28         =   16711680
      Event_28        =   4
      X_28            =   8
      Y_28            =   15
      XPos_28         =   2
      YPos_28         =   2
      PointIndex_28   =   -1
      ChrosshairStyle_28=   8
      LockPlot_28     =   0
      TrackMode_1     =   6
      GraphBackground_1=   0
      GraphFrame_1    =   29
      ClassName_29    =   "CCWDrawObj"
      opts_29         =   30
      Bindings_29     =   0
      Image_29        =   30
      ClassName_30    =   "CCWPictImage"
      opts_30         =   1280
      Bindings_30     =   0
      Rows_30         =   1
      Cols_30         =   1
      F_30            =   -2147483633
      B_30            =   -2147483633
      ColorReplaceWith_30=   8421504
      ColorReplace_30 =   8421504
      Tolerance_30    =   2
      Animator_29     =   0
      Blinker_29      =   0
      PlotFrame_1     =   31
      ClassName_31    =   "CCWDrawObj"
      opts_31         =   30
      Bindings_31     =   0
      C[1]_31         =   16777215
      Image_31        =   32
      ClassName_32    =   "CCWPictImage"
      opts_32         =   1280
      Bindings_32     =   0
      Rows_32         =   1
      Cols_32         =   1
      Pict_32         =   1
      F_32            =   -2147483633
      B_32            =   16777215
      ColorReplaceWith_32=   8421504
      ColorReplace_32 =   8421504
      Tolerance_32    =   2
      Animator_31     =   0
      Blinker_31      =   0
      Caption_1       =   33
      ClassName_33    =   "CCWDrawObj"
      opts_33         =   30
      Bindings_33     =   0
      C[0]_33         =   -2147483640
      Image_33        =   34
      ClassName_34    =   "CCWTextImage"
      Bindings_34     =   0
      font_34         =   0
      Animator_33     =   0
      Blinker_33      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
   End
   Begin VB.CommandButton cmdViewFile 
      Caption         =   "&View Data File"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearGraph 
      Caption         =   "Clear &Graph"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Data File"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtDataFile 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Text            =   " "
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label lblPlotData 
      Alignment       =   2  'Center
      Caption         =   "X axis in minutes"
      Height          =   255
      Left            =   1920
      TabIndex        =   51
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblSensor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   35
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblSensor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   34
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblSensor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   33
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblSensor 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   32
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblXAxis 
      Alignment       =   2  'Center
      Caption         =   "TIME - Seconds"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "R"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "I"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "N"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "E"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "D"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label lblDataFile 
      Alignment       =   2  'Center
      Caption         =   "Data File Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Data File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Open Data File In &Excel"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Data File"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompare 
         Caption         =   "&Compare Data Files"
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintScreen 
         Caption         =   "Print &Screen"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileStep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Return"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuGraphOp 
      Caption         =   "&Graph Operations"
      Begin VB.Menu mnuGraphZoom 
         Caption         =   "&Zoom"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuGraphPan 
         Caption         =   "&Pan"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGraphCur 
         Caption         =   "&Cursor Coordinates"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "PlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mi_XMax                 As Integer
Dim mi_XMin                 As Integer
Dim mi_YMax                 As Integer
Dim mi_YMin                 As Integer

Dim mb_BarChart             As Boolean
Dim mb_LineChart            As Boolean

Dim ms_Sensor(1 To 8)       As String
Dim mi_ChkSensor(1 To 8)    As Integer
Dim mi_ActualSensor(1 To 4) As Integer
Dim mi_NumChkSensor         As Integer
Dim ml_SensorColors(1 To 8) As Long

Dim mo_DataFile             As FTDataFile

Private Sub SetDatafileDialog(cdlData As MSComDlg.CommonDialog)
    With cdlData
        .CancelError = True
        .DialogTitle = "Select Data File To Open"
        .Filter = "All Readable File Types (*.dat;*.csv;*.xls)|*.dat;*.csv;*.xls|FiberTrack Data Files (*.dat)|*.dat|Excel Files (*.xls)|*.xls|Comma-Separated Values Files (*.csv)|*.csv|All Files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNExplorer Or _
                 cdlOFNFileMustExist Or _
                 cdlOFNNoChangeDir Or _
                 cdlOFNHideReadOnly
    End With
End Sub

Public Sub ChooseExcelFile()
On Error GoTo ErrHan
    Dim o_Controller As New FTExcelController
    SetDatafileDialog comDialog1
    With comDialog1
        .DialogTitle = "Select Data File to View In Excel"
        .InitDir = App.Path
        .ShowOpen
        o_Controller.ViewDataFileInExcel .FileName
    End With
    Exit Sub
ErrHan:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, "FiberTrack PlotData"
    Err.Clear
End Sub

Public Sub ChoosePlotDataFile()
On Error GoTo ErrHan
    SetDatafileDialog comDialog1
    With comDialog1
        .DialogTitle = "Select Data File to Plot"
        .InitDir = App.Path
        .ShowOpen
        OpenDataFile .FileName
    End With
    Exit Sub
ErrHan:
    If Err.Number = cdlCancel Then Exit Sub     'Cancel
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "FiberTrack Open PlotData File Error"
    If Not mo_DataFile Is Nothing Then
        mo_DataFile.CloseFile
        Set mo_DataFile = Nothing
    End If
    IsOpenDataFile = False
    'Disable unavailable functionality
    mnuView.Enabled = False
    cmdPlotData.Enabled = False
End Sub

'Invisible button
Private Sub cmdBarChart_Click()
    ' hide line chart
    mb_LineChart = False
    ' show bar chart
'    BarChart.Visible = True
    cwgBarChart.Visible = True
    mb_BarChart = True
    lblCaption.Caption = ""
    cwgBarChart.ClearData
    cwgBarChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgBarChart.Axes.Item(2).SetMinMax 0, mi_YMax
    GraphMode.ValuePairIndex = 1                ' 1 = Zoom
    
    cwgBarChart.TrackMode = cwGTrackZoomRectXY
    cwgBarChart.Cursors(1).Visible = False
    cwgBarChart.Cursors(2).Visible = False
    XPosition.Text = ""
    YPosition.Text = ""
    
    Dim i As Integer
    For i = lblSensor.LBound To lblSensor.UBound
        lblSensor(i).Visible = False
        lblSensor(i).Caption = ""
    Next i
End Sub

Private Sub cmdBarChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdBarChart.ToolTipText = "Click here to plot the current data as a bar chart."
End Sub

Private Sub cmdCancel_Click()
    Dim uResponse   As VbMsgBoxResult
    Dim i           As Integer
    uResponse = MsgBox("Do you wish to cancel the data plot?", vbYesNo Or vbQuestion, "FiberTrack - PlotData")
    
    If uResponse = vbYes Then
        lblCaption.Caption = ""
        fraSensorCheck.Visible = False
        
        For i = chkSensor.LBound To chkSensor.UBound
            chkSensor(i).Visible = False
            chkSensor(i).Caption = ""
            chkSensor(i).value = 0
        Next i
        mi_NumChkSensor = 0
        IsOpenDataFile = False
        'Disable unavailable functionality
        cmdPlotData.Enabled = False
    Else
        'cmdOK_Click
    End If
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdCancel.ToolTipText = "Click here to discard the above choices."
End Sub

Private Sub cmdClearGraph_Click()
    cwgLineChart.ClearData
    cwgLineChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgLineChart.Axes.Item(2).SetMinMax 0, mi_YMax
    
    cwgBarChart.ClearData
    cwgBarChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgBarChart.Axes.Item(2).SetMinMax 0, mi_YMax
        
    GraphMode.ValuePairIndex = 1                ' 1 = Zoom
    cwgBarChart.TrackMode = cwGTrackZoomRectXY
    cwgBarChart.Cursors(1).Visible = False
    cwgBarChart.Cursors(2).Visible = False
    
    XPosition.Text = ""
    YPosition.Text = ""
    txtDataFile.Text = ""
    lblCaption.Caption = ""
    
    Dim i As Integer
    For i = lblSensor.LBound To lblSensor.UBound
        lblSensor(i).Visible = False
        lblSensor(i).Caption = ""
    Next i
End Sub

Private Sub cmdClearGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdClearGraph.ToolTipText = "Click here to clear the data plot."
End Sub

'Invisible button
Private Sub cmdLineChart_Click()
    ' show line chart
    cwgLineChart.Visible = True     ' line chart
    mb_LineChart = True
    ' hide bar chart
    cwgBarChart.Visible = False     ' bar chart
    mb_BarChart = False
    cwgLineChart.ClearData
    cwgLineChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgLineChart.Axes.Item(2).SetMinMax 0, mi_YMax
    lblCaption.Caption = ""
    GraphMode.ValuePairIndex = 1                ' 1 = Zoom
    cwgLineChart.TrackMode = cwGTrackZoomRectXY
    cwgLineChart.Cursors(1).Visible = False
    cwgLineChart.Cursors(2).Visible = False
    XPosition.Text = ""
    YPosition.Text = ""
    Dim i As Integer
    For i = lblSensor.LBound To lblSensor.UBound
        lblSensor(i).Visible = False
        lblSensor(i).Caption = ""
    Next i
    txtDataFile.Text = ""
    txtDataFile.SetFocus
End Sub

Private Sub cmdLineChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdLineChart.ToolTipText = "Click here to plot data from saved data files as a line chart."
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    mi_NumChkSensor = 0
    For i = 1 To mo_DataFile.NumSensors
        'set mi_ChkSensor()
        mi_ChkSensor(i) = chkSensor(i - 1).value
        If mi_ChkSensor(i) = vbChecked Then
            mi_NumChkSensor = mi_NumChkSensor + 1
            If mi_NumChkSensor <= UBound(mi_ActualSensor) Then
                mi_ActualSensor(i) = chkSensor(i - 1).Tag
            End If
        End If
    Next i
    If mi_NumChkSensor = 0 Then
        MsgBox "Please check the sensors to plot.", vbInformation Or vbOKOnly, "PlotData"
        Exit Sub
    End If
    If mi_NumChkSensor > 4 Then
        MsgBox "Please check only from one to four sensors for plotting.", vbInformation Or vbOKOnly, "PlotData"
        Exit Sub
    End If
    fraSensorCheck.Visible = False
    For i = chkSensor.LBound To chkSensor.UBound
        chkSensor(i).Visible = False
        chkSensor(i).Caption = ""
    Next i
    cboomPlotData.SetFocus
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOK.ToolTipText = "Click here to accept the above choices."
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOpen.ToolTipText = "Click here to open the data file and display the heading information above the chart."
End Sub

Private Sub cmdPlotData_Click()
    PlotDataFile
End Sub         'cmdPlotData

Private Sub cmdPlotData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPlotData.ToolTipText = "Click here to plot the data file."
End Sub

Private Sub cmdPrintForm_Click()
    PlotData.PrintForm
End Sub

Private Sub cmdPrintForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrintForm.ToolTipText = "Click here to print the current screen image."
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub cmdReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdReturn.ToolTipText = "Click here to return to the main FiberTrack screen."
End Sub

'Invisible button
Private Sub cmdViewFile_Click()
On Error GoTo viewfile_err_hdr
    Dim returnValue As Variant
    Dim response As Integer
    If mo_DataFile.FileFormat = Excel Then
        response = MsgBox("Do you wish to view the data file with Excel?", vbYesNo Or vbQuestion Or vbDefaultButton2, "PlotData, cmdViewFile")
        If response = vbYes Then
            returnValue = Shell("excel.exe " & txtDataFile.Text, vbNormalFocus)
            Exit Sub
        End If
    End If
    returnValue = Shell("notepad.exe " & txtDataFile.Text, vbNormalFocus)
    Exit Sub
viewfile_err_hdr:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open DataFile Error, cmdViewFile"
    Exit Sub
End Sub

Private Sub cmdViewFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdViewFile.ToolTipText = "Click here to view the data file with either the Windows application Excel or Notepad."
End Sub

Private Sub cmdZoomOut_Click()
    cwgLineChart.Axes.Item(1).SetMinMax mi_XMin, mi_XMax
    cwgLineChart.Axes.Item(2).SetMinMax mi_YMin, mi_YMax
    
    cwgBarChart.Axes.Item(1).SetMinMax mi_XMin, mi_XMax
    cwgBarChart.Axes.Item(2).SetMinMax mi_YMin, mi_YMax
End Sub

Private Sub cmdZoomOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdZoomOut.ToolTipText = "Click here to return the data plot to the original mode."
End Sub

Private Sub cwgLineChart_CursorChange(CursorIndex As Long, XPos As Variant, YPos As Variant, bTracking As Boolean)
    'Display cursor position on user interface
    If CursorIndex = 1 Then
        XPosition.Text = cwgLineChart.Cursors(1).XPosition ' data from cursor
        YPosition.Text = YPos ' data from event handler
    End If
    
    If GraphMode.value = 3 Then             'GraphMode is hidden and there is no index of 3
        Dim amplitude As Double
        Dim period As Double
        amplitude = Abs(cwgLineChart.Cursors(2).YPosition - cwgLineChart.Cursors(1).YPosition)
        period = Abs(cwgLineChart.Cursors(2).XPosition - cwgLineChart.Cursors(1).XPosition)
        AmplitudeVal = amplitude
        PeriodVal = period
    End If
End Sub

Private Sub cwgBarChart_CursorChange(CursorIndex As Long, XPos As Variant, YPos As Variant, bTracking As Boolean)
    'Display cursor position on user interface
    If CursorIndex = 1 Then
        XPosition.Text = cwgBarChart.Cursors(1).XPosition ' data from cursor
        YPosition.Text = YPos ' data from event handler
    End If
    
    If GraphMode.value = 3 Then             'GraphMode is hidden and there is no index of 3
        Dim amplitude As Double
        Dim period As Double
        amplitude = Abs(cwgBarChart.Cursors(2).YPosition - cwgBarChart.Cursors(1).YPosition)
        period = Abs(cwgBarChart.Cursors(2).XPosition - cwgBarChart.Cursors(1).XPosition)
        AmplitudeVal = amplitude
        PeriodVal = period
    End If
End Sub

Private Sub Form_Load()
    '12.8.2004 move graph operations to menu
    fraGraphOps.Visible = False
    cboomPlotData.AddItem 5
    cboomPlotData.AddItem 10
    cboomPlotData.AddItem 15
    cboomPlotData.AddItem 20
    cboomPlotData.AddItem 25
    cboomPlotData.AddItem 30
    With Me
        .Caption = StringFormat("Plot Data({0}) - {1} - {2}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"))
        Debug.Print .Left, .Top, .Height, .Width
        .Height = 7065
        .Width = 9570
        
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
        Debug.Print .Left, .Top, .Height, .Width
    End With

    cwgLineChart.Axes.Item(2).Caption = GetIniSetting("Constants", "GraphYCaption")
    cwgLineChart.Axes.Item(1).Caption = "Elapsed Time (Minutes)"
    
    IsOpenDataFile = False
    'Disable unavailable functionality
    mnuView.Enabled = False
    cmdPlotData.Enabled = False
    GraphMode.Visible = False
    cmdViewFile.Visible = False
    cmdOpen.Visible = False
    cmdBarChart.Visible = False
    cmdLineChart.Visible = False
    txtDataFile.Visible = False
    lblDataFile.Visible = False
    
    lblXAxis.Visible = False
    lblYAxis(0).Visible = False
    lblYAxis(1).Visible = False
    lblYAxis(2).Visible = False
    lblYAxis(3).Visible = False
    lblYAxis(4).Visible = False
    lblYAxis(5).Visible = False
    
    mi_XMax = 150
    mi_YMax = 150
    
    fraCursorMeas.Visible = False
    
    lblSensor(0).Visible = False
    lblSensor(1).Visible = False
    lblSensor(2).Visible = False
    lblSensor(3).Visible = False
    
    ml_SensorColors(1) = RGB(0, 0, 255)    ' Blue
    ml_SensorColors(2) = RGB(204, 0, 0)    ' Red
    ml_SensorColors(3) = RGB(204, 255, 0)  ' Yellow
    ml_SensorColors(4) = RGB(0, 102, 0)    ' Green
    ml_SensorColors(5) = RGB(0, 255, 204)  ' Cyan
    ml_SensorColors(6) = RGB(204, 0, 153)  ' Magenta
    ml_SensorColors(7) = RGB(204, 102, 0)  ' Orange
    ml_SensorColors(8) = RGB(0, 51, 0)     ' Darker Green
    
    cwgLineChart.Enabled = True         ' cwgLineChart = line chart
    cwgBarChart.Enabled = True         ' cwgBarChart = bar chart
    
    cwgLineChart.Visible = True        ' cwgLineChart = line chart
    mb_LineChart = True               ' cwgLineChart = line chart
    
    cwgBarChart.Visible = False         ' cwgBarChart = bar chart
    mb_BarChart = False                ' cwgBarChart = bar chart
     
    cwgLineChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgLineChart.Axes.Item(2).SetMinMax 0, mi_YMax
    cwgBarChart.Axes.Item(1).SetMinMax 0, mi_XMax
    cwgBarChart.Axes.Item(2).SetMinMax 0, mi_YMax

    GraphMode.ValuePairIndex = 1                ' 1 = Zoom
    cwgLineChart.TrackMode = cwGTrackZoomRectXY
    cwgLineChart.Cursors(1).Visible = False
    cwgLineChart.Cursors(2).Visible = False
    
    cwgBarChart.TrackMode = cwGTrackZoomRectXY
    cwgBarChart.Cursors(1).Visible = False
    cwgBarChart.Cursors(2).Visible = False
    
    fraSensorCheck.Visible = False
    
    Dim i As Integer
    For i = chkSensor.LBound To chkSensor.UBound
        chkSensor(i).Visible = False
        chkSensor(i).Caption = ""
        chkSensor(i).value = 0
    Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PlotData.PopupMenu mnuFile
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Debug.Print "In PlotData.Form_QueryUnload"
'
End Sub

Private Sub Form_Terminate()
    Debug.Print "In PlotData.Form_Terminate" '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "In PlotData.Form_Unload"
    If Not mo_DataFile Is Nothing Then
        mo_DataFile.CloseFile
        Set mo_DataFile = Nothing
    End If
    If Not FiberTrack.Visible Then
        FiberTrack.Show
        FiberTrack.SetStatusText "Main window."
        ' Turn check mark on menu items on and off.
        FiberTrack.mnuMain.Checked = Not FiberTrack.mnuMain.Checked
        FiberTrack.mnuPlotDataFile.Checked = Not FiberTrack.mnuPlotDataFile.Checked
        IsMain = True
    End If
End Sub

Private Sub fraSensorCheck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraSensorCheck.ToolTipText = "Please check the sensors to plot."
End Sub

Private Sub GraphMode_PointerValueChanged(ByVal Pointer As Long, value As Variant)
    ' Change TrackMode of graph
    ' Hide and show cursors
    Select Case value
        Case 0      'Zoom
            If mb_LineChart Then
                cwgLineChart.TrackMode = cwGTrackZoomRectXY     ' line chart
                cwgLineChart.Cursors(1).Visible = False
                cwgLineChart.Cursors(2).Visible = False
            Else
                cwgBarChart.TrackMode = cwGTrackZoomRectXY     ' bar chart
                cwgBarChart.Cursors(1).Visible = False
                cwgBarChart.Cursors(2).Visible = False
            End If
        Case 1      'Pan
            If mb_LineChart Then
                cwgLineChart.Cursors(1).Visible = False         ' line chart
                cwgLineChart.Cursors(2).Visible = False
                cwgLineChart.TrackMode = cwGTrackPanPlotAreaXY
            Else
                cwgBarChart.Cursors(1).Visible = False         ' bar chart
                cwgBarChart.Cursors(2).Visible = False
                cwgBarChart.TrackMode = cwGTrackPanPlotAreaXY
            End If
        Case 2      'Cursor Coordinates
            If mb_LineChart Then
                cwgLineChart.TrackMode = cwGTrackDragCursor     ' line chart
                cwgLineChart.Cursors(1).Visible = True
                cwgLineChart.Cursors(2).Visible = False
                cwgLineChart.Cursors(1).XPosition = (cwgLineChart.Axes(1).Minimum + cwgLineChart.Axes(1).Maximum) / 2
                cwgLineChart.Cursors(1).YPosition = (cwgLineChart.Axes(2).Minimum + cwgLineChart.Axes(2).Maximum) / 2
            Else
                cwgBarChart.TrackMode = cwGTrackDragCursor     ' bar chart
                cwgBarChart.Cursors(1).Visible = True
                cwgBarChart.Cursors(2).Visible = False
                cwgBarChart.Cursors(1).XPosition = (cwgBarChart.Axes(1).Minimum + cwgBarChart.Axes(1).Maximum) / 2
                cwgBarChart.Cursors(1).YPosition = (cwgBarChart.Axes(2).Minimum + cwgBarChart.Axes(2).Maximum) / 2
            End If
        Case 3      'Unknown
            If mb_LineChart Then
                cwgLineChart.TrackMode = cwGTrackDragCursor     ' line chart
                cwgLineChart.Cursors(1).Visible = True
                cwgLineChart.Cursors(2).Visible = True
            Else
                cwgBarChart.TrackMode = cwGTrackDragCursor     ' bar chart
                cwgBarChart.Cursors(1).Visible = True
                cwgBarChart.Cursors(2).Visible = True
            End If
    End Select
End Sub

Private Sub mnuCompare_Click()
On Error GoTo ErrHan
    Dim dataFile1 As New FTDataFile
    Dim dataFile2 As New FTDataFile
    Dim isRead As Boolean
    SetDatafileDialog comDialog1
    With comDialog1
        .DialogTitle = "Select 1st File To Compare"
        .InitDir = App.Path
        .ShowOpen
        
        isRead = dataFile1.ReadFile(.FileName)
        If Not isRead Then Exit Sub
        
        .DialogTitle = "Select 2nd File To Compare"
        .FileName = ""      'Clear file name selection box in dialog
        .ShowOpen
        
        isRead = dataFile2.ReadFile(.FileName)
        If Not isRead Then Exit Sub
    End With
    
    Dim outputFile As FTDataFile
    Set outputFile = dataFile1.CreateCompareFile(dataFile2, 1, 1)
    If outputFile Is Nothing Then Exit Sub
    
    'MsgBox "Comparison file has been saved to " & o_OutputFile.FileName, vbOKOnly Or vbInformation, "Compare Files"
    OpenDataFile outputFile.FileName
    
    Exit Sub
ErrHan:
    If Err.Number = cdlCancel Then Exit Sub
    Debug.Assert False
End Sub

Private Sub mnuExcel_Click()
    ChooseExcelFile
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGraphCur_Click()
Call optCursorCood_Click
End Sub

Private Sub mnuGraphPan_Click()
Call optPan_Click
End Sub

Private Sub mnuGraphZoom_Click()
Call optZoom_Click
End Sub

Private Sub mnuOpen_Click()
    ChoosePlotDataFile
End Sub

Private Sub mnuView_Click()
On Error GoTo ErrHan
    Dim returnValue As Variant
    Dim currentFileName As String
    If mo_DataFile Is Nothing Then
        currentFileName = ""
    Else
        currentFileName = mo_DataFile.FileName
    End If
    If Trim$(currentFileName) = "" Then
        Beep
    Else
        returnValue = Shell("notepad.exe " & currentFileName, vbNormalFocus)
    End If
    Exit Sub
ErrHan:
    MsgBox "Error Number = " & str(Err.Number) & ",  " & Err.Description & ", " & Err.Source, vbExclamation, "Open DataFile Error, mnuViewDataFile"
    Exit Sub
End Sub

Private Sub optCursorCood_Click()
    If mb_LineChart Then
        cwgLineChart.TrackMode = cwGTrackDragCursor     ' line chart
        cwgLineChart.Cursors(1).Visible = True
        cwgLineChart.Cursors(2).Visible = False
        fraCursorInfo.Visible = True
        YPosition.Visible = True
        cwgLineChart.Cursors(1).XPosition = (cwgLineChart.Axes(1).Minimum + cwgLineChart.Axes(1).Maximum) / 2
        cwgLineChart.Cursors(1).YPosition = (cwgLineChart.Axes(2).Minimum + cwgLineChart.Axes(2).Maximum) / 2
    Else
        cwgBarChart.TrackMode = cwGTrackDragCursor     ' bar chart
        cwgBarChart.Cursors(1).Visible = True
        cwgBarChart.Cursors(2).Visible = False
        fraCursorInfo.Visible = True
        YPosition.Visible = True
        cwgBarChart.Cursors(1).XPosition = (cwgBarChart.Axes(1).Minimum + cwgBarChart.Axes(1).Maximum) / 2
        cwgBarChart.Cursors(1).YPosition = (cwgBarChart.Axes(2).Minimum + cwgBarChart.Axes(2).Maximum) / 2
    End If
End Sub

Private Sub optPan_Click()
    If mb_LineChart Then
        cwgLineChart.Cursors(1).Visible = False         ' line chart
        cwgLineChart.Cursors(2).Visible = False
        cwgLineChart.TrackMode = cwGTrackPanPlotAreaXY
    Else
        cwgBarChart.Cursors(1).Visible = False         ' bar chart
        cwgBarChart.Cursors(2).Visible = False
        cwgBarChart.TrackMode = cwGTrackPanPlotAreaXY
    End If
End Sub

Private Sub optZoom_Click()
    If mb_LineChart Then
        cwgLineChart.TrackMode = cwGTrackZoomRectXY     ' line chart
        cwgLineChart.Cursors(1).Visible = False
        cwgLineChart.Cursors(2).Visible = False
    Else
        cwgBarChart.TrackMode = cwGTrackZoomRectXY     ' bar chart
        cwgBarChart.Cursors(1).Visible = False
        cwgBarChart.Cursors(2).Visible = False
    End If
End Sub

Private Sub OpenDataFile(inputFileName As String)
    Dim index As Integer
    Dim counter As Integer
    Set mo_DataFile = New FTDataFile
    mnuView.Enabled = True
    txtDataFile.Text = inputFileName
    If Not mo_DataFile.ReadFile(inputFileName) Then Exit Sub
    IsOpenDataFile = True
    cmdPlotData.Enabled = True
    For index = chkSensor.LBound To chkSensor.UBound
        chkSensor(index).Visible = False
        chkSensor(index).Caption = ""
        chkSensor(index).value = vbUnchecked
        chkSensor(index).Tag = ""
    Next index
    For index = LBound(mi_ChkSensor) To UBound(mi_ChkSensor)
        mi_ChkSensor(index) = vbUnchecked
    Next index
    For index = LBound(ms_Sensor) To UBound(ms_Sensor)
        ms_Sensor(index) = ""
    Next index
    If mo_DataFile.CompareMode Then
        ms_Sensor(1) = "Sensor 1"
        ms_Sensor(2) = "Sensor 2"
        ms_Sensor(3) = "S1 - S2"
        ms_Sensor(4) = "S1 / S2"
    End If
    'Determine which sensor feeds are available
    counter = 0
    For index = 1 To LAST_SENSOR
        If mo_DataFile.SensorOnline(index) Then
            counter = counter + 1
            If counter > UBound(ms_Sensor) Then Exit For
            If Not mo_DataFile.CompareMode Then
                ms_Sensor(counter) = "Sensor " & CStr(index)
            End If
            If (counter - 1) <= chkSensor.UBound Then
                chkSensor(counter - 1).Tag = index            'Actual Sensor
            End If
        End If
    Next index
    For index = 1 To mo_DataFile.NumSensors
        ' chkSensor goes from 0 to 7
        ' ms_Sensor goes from 1 to 8
        chkSensor(index - 1).Visible = True
        chkSensor(index - 1).Caption = ms_Sensor(index)
    Next index
    mi_YMax = mo_DataFile.DenierTarget + mo_DataFile.DenierTarget * 0.1
    mi_YMin = 0
    mi_XMin = 0
    mi_XMax = 0
    cwgLineChart.Axes.Item(1).SetMinMax mi_XMin, cboomPlotData.Text
    cwgLineChart.Axes.Item(2).SetMinMax mi_YMin, mi_YMax
    cwgBarChart.Axes.Item(1).SetMinMax mi_XMin, mi_XMax
    cwgBarChart.Axes.Item(2).SetMinMax mi_YMin, mi_YMax
    lblCaption.Caption = mo_DataFile.Title & vbCrLf & mo_DataFile.Info & vbCrLf & mo_DataFile.FileName
    fraSensorCheck.Visible = True
    cmdCancel.SetFocus
    ' <<<<<<<<<< Exit here and leave the rest to cmdPlotData >>>>>>>>>>
    ' <<<<<<<<<< The Input Data File is still Opened !!! >>>>>>>>>>
End Sub

Private Sub PlotDataFile()
    'General bounds information:
        'lblSensor:             0 To 3
        'chkSensor:             0 To 7
        'mi_ChkSensor:          1 To 8
        'ms_Sensor:             1 To 8
        'iY:                    1 To 3
        'cwgLineChart.Plots.Item:   1 To 4
On Error GoTo ErrHan
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim n               As Long
    Dim iSeconds        As Integer      'Integration Time
    Dim iRangeMin       As Integer
    Dim iRangeMax       As Integer
    Dim iColumns        As Integer
    Dim iRows           As Integer
    Dim f2PlotArray()   As Double
    Dim f2TimeArray()   As Single

    cwgLineChart.ClearData
    Dim o_Plot As CWPlot
    For Each o_Plot In cwgLineChart.Plots
        If Not (o_Plot Is Nothing) Then
            o_Plot.ClearData
        End If
    Next o_Plot
    
    If Not IsOpenDataFile Then
        Debug.Assert False
        MsgBox "Please open data file.", vbInformation Or vbOKOnly, "PlotData, cmdPlotData"
        Exit Sub
    End If
    
    If mi_NumChkSensor = 0 Then
        MsgBox "Please check the sensors to plot.", vbInformation Or vbOKOnly, "PlotData, cmdPlotData"
        Exit Sub
    End If

    IsOpenDataFile = False
    'Don't disable mnuView here
    cmdPlotData.Enabled = False
    fraGraphOps.Visible = True
    fraCursorInfo.Visible = True
    YPosition.Visible = True
    
    mi_XMin = mi_XMax
    mi_XMax = mi_XMax + cboomPlotData.Text
    cwgLineChart.Axes.Item(1).SetMinMax mi_XMin, mi_XMax
    cwgLineChart.ImmediateUpdates = True

    j = 0
    For i = LBound(mi_ChkSensor) To UBound(mi_ChkSensor)
        If mi_ChkSensor(i) = vbChecked Then
            cwgLineChart.Plots.Item(j + 1).LineColor = ml_SensorColors(i)
            cwgLineChart.Plots.Item(j + 1).LineColor = ml_SensorColors(i)
            lblSensor(j).Visible = True
            lblSensor(j).Caption = ms_Sensor(i)
            lblSensor(j).BackColor = ml_SensorColors(i)
            j = j + 1
        End If
    Next i

    iColumns = 1500      ' with 2 sec interval this is 5 minutes of data
                         ' position 0 for X-axis

    iRows = mo_DataFile.NumSensors
    ReDim f2PlotArray(1 To mi_NumChkSensor)   ' number of sensor to plot
    ReDim f2TimeArray(1 To mi_NumChkSensor)
    cwgLineChart.ChartStyle = cwChartScope
    iSeconds = CInt(mo_DataFile.IntegrationTime)
    iRangeMin = 0
    iRangeMax = 60 / iSeconds * cboomPlotData.Text
    
    If iRangeMax > mo_DataFile.NumValues Then
        iRangeMax = CInt(mo_DataFile.NumValues)
    End If
    
    'Set chart buffer; otherwise older data will get deleted and
    'it will look like the chart is running off stage right
    cwgLineChart.ChartLength = mo_DataFile.NumValues    'Max version
    'cwgLineChart.ChartLength = iRangeMax - iRangeMin   'Min version
    For n = iRangeMin To iRangeMax
        If (n - iRangeMin) >= iColumns Then Exit For
        k = 0
        For j = 1 To LAST_SENSOR
            Debug.Assert j >= LBound(mi_ChkSensor) And j <= UBound(mi_ChkSensor)
            If mi_ChkSensor(j) = vbChecked Then
                k = k + 1
                
                Debug.Assert k >= LBound(f2PlotArray) And k <= UBound(f2PlotArray)
                f2PlotArray(k) = mo_DataFile.DenierValues(SensorFromChkSensor(j), n)
            End If
        Next j
        If mb_LineChart Then
            If mo_DataFile.DataLine(n).TimeAsMinutes > cboomPlotData.Text Then Exit For
            
            For i = LBound(f2PlotArray) To UBound(f2PlotArray)
                cwgLineChart.Plots(i).ChartXvsY mo_DataFile.DataLine(n).TimeAsMinutes, f2PlotArray(i)
            Next i
        ElseIf mb_BarChart Then
            cwgBarChart.ChartY f2PlotArray, iSeconds, False
        End If
    Next n
    mi_NumChkSensor = 0
    cwgLineChart.Refresh
    Exit Sub
ErrHan:
    MsgBox "Error occurred plotting data file." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation Or vbOKOnly, "FiberTrack - Error Plotting Data File"
    Exit Sub
End Sub

'Returns actual sensor number from mi_ChkSensor index
Private Function SensorFromChkSensor(sensorIndex As Integer) As Integer
    Debug.Assert sensorIndex >= LBound(mi_ActualSensor) And sensorIndex <= UBound(mi_ActualSensor)
    SensorFromChkSensor = mi_ActualSensor(sensorIndex)
End Function
