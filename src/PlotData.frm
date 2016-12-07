VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PlotData 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Plot Data"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   1575
   ClientWidth     =   9480
   Icon            =   "PlotData.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSensorCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Please check "
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   8160
      TabIndex        =   36
      ToolTipText     =   "Please check the sensors to plot."
      Top             =   840
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
      Height          =   2295
      Left            =   840
      TabIndex        =   30
      Top             =   960
      Width           =   8415
      _Version        =   524288
      _ExtentX        =   14843
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
      CompatibleVers_0=   524288
      Graph_0         =   1
      ClassName_1     =   "CCWGraphFrame"
      opts_1          =   30
      C[0]_1          =   16777215
      Event_1         =   2
      ClassName_2     =   "CCWGFPlotEvent"
      Owner_2         =   1
      Plots_1         =   3
      ClassName_3     =   "CCWDataPlots"
      Array_3         =   4
      Editor_3        =   4
      ClassName_4     =   "CCWGFPlotArrayEditor"
      Owner_4         =   1
      Array[0]_3      =   5
      ClassName_5     =   "CCWDataPlot"
      opts_5          =   4194335
      Name_5          =   "Plot-1"
      C[0]_5          =   255
      C[1]_5          =   65280
      C[2]_5          =   16711680
      C[3]_5          =   16711680
      Event_5         =   2
      X_5             =   6
      ClassName_6     =   "CCWAxis"
      opts_6          =   543
      Name_6          =   "Time - Seconds"
      Orientation_6   =   2944
      format_6        =   7
      ClassName_7     =   "CCWFormat"
      Scale_6         =   8
      ClassName_8     =   "CCWScale"
      opts_8          =   90112
      rMin_8          =   25
      rMax_8          =   550
      dMax_8          =   20
      discInterval_8  =   1
      Radial_6        =   0
      Enum_6          =   9
      ClassName_9     =   "CCWEnum"
      Editor_9        =   10
      ClassName_10    =   "CCWEnumArrayEditor"
      Owner_10        =   6
      Font_6          =   0
      tickopts_6      =   1679
      base_6          =   2
      major_6         =   10
      minor_6         =   5
      Caption_6       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   30
      C[0]_11         =   -2147483640
      Image_11        =   12
      ClassName_12    =   "CCWTextImage"
      font_12         =   0
      Animator_11     =   0
      Blinker_11      =   0
      Y_5             =   13
      ClassName_13    =   "CCWAxis"
      opts_13         =   1567
      Name_13         =   "Denier"
      Orientation_13  =   2067
      format_13       =   14
      ClassName_14    =   "CCWFormat"
      Scale_13        =   15
      ClassName_15    =   "CCWScale"
      opts_15         =   122880
      rMin_15         =   11
      rMax_15         =   126
      dMax_15         =   10
      discInterval_15 =   1
      Radial_13       =   0
      Enum_13         =   16
      ClassName_16    =   "CCWEnum"
      Editor_16       =   17
      ClassName_17    =   "CCWEnumArrayEditor"
      Owner_17        =   13
      Font_13         =   0
      tickopts_13     =   1679
      major_13        =   10
      minor_13        =   5
      Caption_13      =   18
      ClassName_18    =   "CCWDrawObj"
      opts_18         =   30
      C[0]_18         =   -2147483640
      Image_18        =   19
      ClassName_19    =   "CCWTextImage"
      font_19         =   0
      Animator_18     =   0
      Blinker_18      =   0
      LineStyle_5     =   1
      LineWidth_5     =   2
      BasePlot_5      =   0
      DefaultXInc_5   =   1
      DefaultPlotPerRow_5=   -1  'True
      Array[1]_3      =   20
      ClassName_20    =   "CCWDataPlot"
      opts_20         =   4194335
      Name_20         =   "Plot-2"
      C[0]_20         =   16711680
      C[1]_20         =   255
      C[2]_20         =   16711680
      C[3]_20         =   16776960
      Event_20        =   2
      X_20            =   6
      Y_20            =   13
      LineStyle_20    =   1
      LineWidth_20    =   2
      BasePlot_20     =   0
      DefaultXInc_20  =   1
      DefaultPlotPerRow_20=   -1  'True
      Array[2]_3      =   21
      ClassName_21    =   "CCWDataPlot"
      opts_21         =   4194335
      Name_21         =   "Plot-3"
      C[0]_21         =   255
      C[1]_21         =   255
      C[2]_21         =   16711680
      C[3]_21         =   16776960
      Event_21        =   2
      X_21            =   6
      Y_21            =   13
      LineStyle_21    =   1
      LineWidth_21    =   2
      BasePlot_21     =   0
      DefaultXInc_21  =   1
      DefaultPlotPerRow_21=   -1  'True
      Array[3]_3      =   22
      ClassName_22    =   "CCWDataPlot"
      opts_22         =   4194335
      Name_22         =   "Plot-4"
      C[0]_22         =   255
      C[1]_22         =   65280
      C[2]_22         =   16711680
      C[3]_22         =   16711680
      Event_22        =   2
      X_22            =   6
      Y_22            =   13
      LineStyle_22    =   1
      LineWidth_22    =   2
      BasePlot_22     =   0
      DefaultXInc_22  =   1
      DefaultPlotPerRow_22=   -1  'True
      Axes_1          =   23
      ClassName_23    =   "CCWAxes"
      Array_23        =   2
      Editor_23       =   24
      ClassName_24    =   "CCWGFAxisArrayEditor"
      Owner_24        =   1
      Array[0]_23     =   6
      Array[1]_23     =   13
      DefaultPlot_1   =   25
      ClassName_25    =   "CCWDataPlot"
      opts_25         =   4194335
      Name_25         =   "[Template]"
      C[0]_25         =   255
      C[1]_25         =   65280
      C[2]_25         =   16711680
      C[3]_25         =   16711680
      Event_25        =   2
      X_25            =   6
      Y_25            =   13
      LineStyle_25    =   1
      LineWidth_25    =   2
      BasePlot_25     =   0
      DefaultXInc_25  =   1
      DefaultPlotPerRow_25=   -1  'True
      Cursors_1       =   26
      ClassName_26    =   "CCWCursors"
      Array_26        =   2
      Editor_26       =   27
      ClassName_27    =   "CCWGFCursorArrayEditor"
      Owner_27        =   1
      Array[0]_26     =   28
      ClassName_28    =   "CCWCursor"
      opts_28         =   31
      Name_28         =   "Cursor-1"
      C[0]_28         =   255
      Event_28        =   2
      X_28            =   6
      Y_28            =   13
      XPos_28         =   2
      YPos_28         =   1
      PointIndex_28   =   -1
      ChrosshairStyle_28=   8
      LockPlot_28     =   0
      Array[1]_26     =   29
      ClassName_29    =   "CCWCursor"
      opts_29         =   31
      Name_29         =   "Cursor-2"
      C[0]_29         =   16711680
      Event_29        =   2
      X_29            =   6
      Y_29            =   13
      XPos_29         =   4
      YPos_29         =   2
      PointIndex_29   =   -1
      ChrosshairStyle_29=   8
      LockPlot_29     =   0
      TrackMode_1     =   6
      GraphBackground_1=   0
      GraphFrame_1    =   30
      ClassName_30    =   "CCWDrawObj"
      opts_30         =   30
      Image_30        =   31
      ClassName_31    =   "CCWPictImage"
      opts_31         =   1280
      Rows_31         =   1
      Cols_31         =   1
      F_31            =   -2147483633
      B_31            =   -2147483633
      ColorReplaceWith_31=   8421504
      ColorReplace_31 =   8421504
      Tolerance_31    =   2
      Animator_30     =   0
      Blinker_30      =   0
      PlotFrame_1     =   32
      ClassName_32    =   "CCWDrawObj"
      opts_32         =   30
      C[1]_32         =   16777215
      Image_32        =   33
      ClassName_33    =   "CCWPictImage"
      opts_33         =   1280
      Rows_33         =   1
      Cols_33         =   1
      Pict_33         =   1
      F_33            =   -2147483633
      B_33            =   16777215
      ColorReplaceWith_33=   8421504
      ColorReplace_33 =   8421504
      Tolerance_33    =   2
      Animator_32     =   0
      Blinker_32      =   0
      Caption_1       =   34
      ClassName_34    =   "CCWDrawObj"
      opts_34         =   30
      C[0]_34         =   -2147483640
      Image_34        =   35
      ClassName_35    =   "CCWTextImage"
      font_35         =   0
      Animator_34     =   0
      Blinker_34      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
      Bindings_1      =   36
      ClassName_36    =   "CCWBindingHolderArray"
      Editor_36       =   37
      ClassName_37    =   "CCWBindingHolderArrayEditor"
      Owner_37        =   1
      Annotations_1   =   38
      ClassName_38    =   "CCWAnnotations"
      Editor_38       =   39
      ClassName_39    =   "CCWAnnotationArrayEditor"
      Owner_39        =   1
      AnnotationTemplate_1=   40
      ClassName_40    =   "CCWAnnotation"
      opts_40         =   63
      Name_40         =   "[Template]"
      Plot_40         =   25
      Text_40         =   "[Template]"
      TextXPoint_40   =   13.4
      TextYPoint_40   =   13.4
      TextColor_40    =   16777215
      TextFont_40     =   41
      ClassName_41    =   "CCWFont"
      bFont_41        =   -1  'True
      BeginProperty Font_41 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShapeXPoints_40 =   42
      ClassName_42    =   "CDataBuffer"
      Type_42         =   5
      m_cDims;_42     =   1
      m_cElts_42      =   1
      Element[0]_42   =   6.6
      ShapeYPoints_40 =   43
      ClassName_43    =   "CDataBuffer"
      Type_43         =   5
      m_cDims;_43     =   1
      m_cElts_43      =   1
      Element[0]_43   =   6.6
      ShapeFillColor_40=   16777215
      ShapeLineColor_40=   16777215
      ShapeLineWidth_40=   1
      ShapeLineStyle_40=   1
      ShapePointStyle_40=   10
      ShapeImage_40   =   44
      ClassName_44    =   "CCWDrawObj"
      opts_44         =   62
      Image_44        =   45
      ClassName_45    =   "CCWPictImage"
      opts_45         =   1280
      Rows_45         =   1
      Cols_45         =   1
      Pict_45         =   7
      F_45            =   -2147483633
      B_45            =   -2147483633
      ColorReplaceWith_45=   8421504
      ColorReplace_45 =   8421504
      Tolerance_45    =   2
      Animator_44     =   0
      Blinker_44      =   0
      ArrowVisible_40 =   -1  'True
      ArrowColor_40   =   16777215
      ArrowWidth_40   =   1
      ArrowLineStyle_40=   1
      ArrowHeadStyle_40=   1
   End
   Begin VB.ComboBox cboomPlotData 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   5040
      Width           =   735
   End
   Begin MSComDlg.CommonDialog comDialog1 
      Left            =   600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "Print S&creen"
      Height          =   495
      Left            =   4320
      TabIndex        =   46
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlotData 
      Caption         =   "&Plot Data File"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "&Zoom Out"
      Height          =   495
      Left            =   5640
      TabIndex        =   31
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame fraCursorMeas 
      Caption         =   "Cursor Measurements"
      Height          =   975
      Left            =   7080
      TabIndex        =   25
      Top             =   3960
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
      Left            =   8160
      TabIndex        =   24
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame fraCursorInfo 
      Caption         =   "Cursor Information"
      Height          =   975
      Left            =   7080
      TabIndex        =   20
      Top             =   5160
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
      Left            =   4440
      TabIndex        =   18
      Top             =   4800
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
         _Version        =   524288
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
         CompatibleVers_0=   524288
         Slider_0        =   1
         ClassName_1     =   "CCWSlider"
         opts_1          =   2078
         C[0]_1          =   -2147483643
         BGImg_1         =   2
         ClassName_2     =   "CCWDrawObj"
         opts_2          =   30
         Image_2         =   3
         ClassName_3     =   "CCWPictImage"
         opts_3          =   1280
         Rows_3          =   1
         Cols_3          =   1
         Pict_3          =   286
         F_3             =   -2147483633
         B_3             =   -2147483633
         ColorReplaceWith_3=   8421504
         ColorReplace_3  =   8421504
         Tolerance_3     =   2
         Animator_2      =   0
         Blinker_2       =   0
         BFImg_1         =   4
         ClassName_4     =   "CCWDrawObj"
         opts_4          =   62
         Image_4         =   5
         ClassName_5     =   "CCWPictImage"
         opts_5          =   1280
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
         Label_1         =   6
         ClassName_6     =   "CCWDrawObj"
         opts_6          =   30
         C[0]_6          =   -2147483640
         Image_6         =   7
         ClassName_7     =   "CCWTextImage"
         style_7         =   15878208
         font_7          =   0
         Animator_6      =   0
         Blinker_6       =   0
         Border_1        =   8
         ClassName_8     =   "CCWDrawObj"
         opts_8          =   28
         Image_8         =   9
         ClassName_9     =   "CCWPictImage"
         opts_9          =   1280
         Rows_9          =   1
         Cols_9          =   1
         Pict_9          =   25
         F_9             =   -2147483633
         B_9             =   -2147483633
         ColorReplaceWith_9=   8421504
         ColorReplace_9  =   8421504
         Tolerance_9     =   2
         Animator_8      =   0
         Blinker_8       =   0
         FillBound_1     =   10
         ClassName_10    =   "CCWGuiObject"
         opts_10         =   28
         FillTok_1       =   11
         ClassName_11    =   "CCWGuiObject"
         opts_11         =   30
         Axis_1          =   12
         ClassName_12    =   "CCWAxis"
         opts_12         =   1055
         Name_12         =   "Axis"
         Orientation_12  =   133523
         format_12       =   13
         ClassName_13    =   "CCWFormat"
         Scale_12        =   14
         ClassName_14    =   "CCWScale"
         opts_14         =   24576
         rMin_14         =   10
         rMax_14         =   54
         dMin_14         =   1
         dMax_14         =   3
         discInterval_14 =   1
         Radial_12       =   0
         Enum_12         =   15
         ClassName_15    =   "CCWEnum"
         Array_15        =   3
         Editor_15       =   16
         ClassName_16    =   "CCWEnumArrayEditor"
         Owner_16        =   12
         Array[0]_15     =   17
         ClassName_17    =   "CCWEnumElt"
         opts_17         =   1
         Name_17         =   "Zoom"
         DrawList_17     =   0
         varVarType_17   =   2
         Array[1]_15     =   18
         ClassName_18    =   "CCWEnumElt"
         opts_18         =   1
         Name_18         =   "Pan"
         DrawList_18     =   0
         varVarType_18   =   2
         var_Val_18      =   1
         Array[2]_15     =   19
         ClassName_19    =   "CCWEnumElt"
         opts_19         =   1
         Name_19         =   "Cursor Coordinates"
         DrawList_19     =   0
         varVarType_19   =   2
         var_Val_19      =   2
         Font_12         =   0
         tickopts_12     =   2718
         Caption_12      =   20
         ClassName_20    =   "CCWDrawObj"
         opts_20         =   30
         C[0]_20         =   -2147483640
         Image_20        =   21
         ClassName_21    =   "CCWTextImage"
         font_21         =   0
         Animator_20     =   0
         Blinker_20      =   0
         DrawLst_1       =   22
         ClassName_22    =   "CDrawList"
         count_22        =   10
         list[10]_22     =   8
         list[9]_22      =   23
         ClassName_23    =   "CCWThumb"
         opts_23         =   31
         Name_23         =   "Pointer-1"
         C[0]_23         =   8388608
         C[1]_23         =   8388608
         C[2]_23         =   -2147483635
         Image_23        =   24
         ClassName_24    =   "CCWPictImage"
         opts_24         =   1280
         Rows_24         =   1
         Cols_24         =   1
         Pict_24         =   213
         F_24            =   8388608
         B_24            =   8388608
         ColorReplaceWith_24=   8421504
         ColorReplace_24 =   8421504
         Tolerance_24    =   2
         Animator_23     =   0
         Blinker_23      =   0
         style_23        =   1
         Value_23        =   1
         Fill_23         =   25
         ClassName_25    =   "CCWDrawObj"
         opts_25         =   62
         Image_25        =   26
         ClassName_26    =   "CCWPictImage"
         opts_26         =   1280
         Rows_26         =   1
         Cols_26         =   1
         Pict_26         =   286
         F_26            =   -2147483633
         B_26            =   -2147483633
         ColorReplaceWith_26=   8421504
         ColorReplace_26 =   8421504
         Tolerance_26    =   2
         Animator_25     =   0
         Blinker_25      =   0
         list[8]_22      =   12
         list[7]_22      =   6
         list[6]_22      =   11
         list[5]_22      =   4
         list[4]_22      =   27
         ClassName_27    =   "CCWDrawObj"
         opts_27         =   30
         Image_27        =   28
         ClassName_28    =   "CCWPictImage"
         opts_28         =   1280
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
         list[3]_22      =   29
         ClassName_29    =   "CCWDrawObj"
         opts_29         =   30
         Image_29        =   30
         ClassName_30    =   "CCWPictImage"
         opts_30         =   1280
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
         list[2]_22      =   31
         ClassName_31    =   "CCWDrawObj"
         opts_31         =   30
         Image_31        =   32
         ClassName_32    =   "CCWPictImage"
         opts_32         =   1280
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
         list[1]_22      =   2
         IncDec_1        =   0
         Ptrs_1          =   33
         ClassName_33    =   "CCWPointerArray"
         Array_33        =   1
         Editor_33       =   34
         ClassName_34    =   "CCWPointerArrayEditor"
         Owner_34        =   1
         Array[0]_33     =   23
         Bindings_1      =   35
         ClassName_35    =   "CCWBindingHolderArray"
         Editor_35       =   36
         ClassName_36    =   "CCWBindingHolderArrayEditor"
         Owner_36        =   1
         Stats_1         =   37
         ClassName_37    =   "CCWStats"
         doInc_1         =   31
         doDec_1         =   29
         doFrame_1       =   27
      End
   End
   Begin VB.CommandButton cmdBarChart 
      Caption         =   "&Bar Chart"
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdLineChart 
      Caption         =   "&Line Chart"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1095
   End
   Begin CWUIControlsLib.CWGraph cwgBarChart 
      Height          =   2295
      Left            =   1440
      TabIndex        =   15
      Top             =   840
      Width           =   6615
      _Version        =   524288
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
      CompatibleVers_0=   524288
      Graph_0         =   1
      ClassName_1     =   "CCWGraphFrame"
      opts_1          =   30
      C[0]_1          =   16777215
      Event_1         =   2
      ClassName_2     =   "CCWGFPlotEvent"
      Owner_2         =   1
      Plots_1         =   3
      ClassName_3     =   "CCWDataPlots"
      Array_3         =   1
      Editor_3        =   4
      ClassName_4     =   "CCWGFPlotArrayEditor"
      Owner_4         =   1
      Array[0]_3      =   5
      ClassName_5     =   "CCWDataPlot"
      opts_5          =   7340063
      Name_5          =   "Plot-1"
      C[0]_5          =   16776960
      C[1]_5          =   255
      C[2]_5          =   16711680
      C[3]_5          =   16776960
      Event_5         =   2
      X_5             =   6
      ClassName_6     =   "CCWAxis"
      opts_6          =   1567
      Name_6          =   "Time - Seconds"
      Orientation_6   =   2944
      format_6        =   7
      ClassName_7     =   "CCWFormat"
      Scale_6         =   8
      ClassName_8     =   "CCWScale"
      opts_8          =   90112
      rMin_8          =   25
      rMax_8          =   430
      dMax_8          =   10
      discInterval_8  =   1
      Radial_6        =   0
      Enum_6          =   9
      ClassName_9     =   "CCWEnum"
      Editor_9        =   10
      ClassName_10    =   "CCWEnumArrayEditor"
      Owner_10        =   6
      Font_6          =   0
      tickopts_6      =   1679
      base_6          =   2
      major_6         =   10
      minor_6         =   5
      Caption_6       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   30
      C[0]_11         =   -2147483640
      Image_11        =   12
      ClassName_12    =   "CCWTextImage"
      font_12         =   0
      Animator_11     =   0
      Blinker_11      =   0
      Y_5             =   13
      ClassName_13    =   "CCWAxis"
      opts_13         =   1567
      Name_13         =   "Denier"
      Orientation_13  =   2067
      format_13       =   14
      ClassName_14    =   "CCWFormat"
      Scale_13        =   15
      ClassName_15    =   "CCWScale"
      opts_15         =   122880
      rMin_15         =   11
      rMax_15         =   126
      dMax_15         =   10
      discInterval_15 =   1
      Radial_13       =   0
      Enum_13         =   16
      ClassName_16    =   "CCWEnum"
      Editor_16       =   17
      ClassName_17    =   "CCWEnumArrayEditor"
      Owner_17        =   13
      Font_13         =   0
      tickopts_13     =   1679
      major_13        =   10
      minor_13        =   5
      Caption_13      =   18
      ClassName_18    =   "CCWDrawObj"
      opts_18         =   30
      C[0]_18         =   -2147483640
      Image_18        =   19
      ClassName_19    =   "CCWTextImage"
      style_19        =   1
      font_19         =   0
      Animator_18     =   0
      Blinker_18      =   0
      LineStyle_5     =   2
      LineWidth_5     =   1
      BasePlot_5      =   0
      DefaultXInc_5   =   1
      DefaultPlotPerRow_5=   -1  'True
      Axes_1          =   20
      ClassName_20    =   "CCWAxes"
      Array_20        =   2
      Editor_20       =   21
      ClassName_21    =   "CCWGFAxisArrayEditor"
      Owner_21        =   1
      Array[0]_20     =   6
      Array[1]_20     =   13
      DefaultPlot_1   =   22
      ClassName_22    =   "CCWDataPlot"
      opts_22         =   7340063
      Name_22         =   "[Template]"
      C[0]_22         =   16776960
      C[1]_22         =   255
      C[2]_22         =   16711680
      C[3]_22         =   16776960
      Event_22        =   2
      X_22            =   6
      Y_22            =   13
      LineStyle_22    =   2
      LineWidth_22    =   1
      BasePlot_22     =   0
      DefaultXInc_22  =   1
      DefaultPlotPerRow_22=   -1  'True
      Cursors_1       =   23
      ClassName_23    =   "CCWCursors"
      Array_23        =   2
      Editor_23       =   24
      ClassName_24    =   "CCWGFCursorArrayEditor"
      Owner_24        =   1
      Array[0]_23     =   25
      ClassName_25    =   "CCWCursor"
      opts_25         =   31
      Name_25         =   "Cursor-1"
      C[0]_25         =   255
      Event_25        =   2
      X_25            =   6
      Y_25            =   13
      XPos_25         =   1
      YPos_25         =   1
      PointIndex_25   =   -1
      ChrosshairStyle_25=   8
      LockPlot_25     =   0
      Array[1]_23     =   26
      ClassName_26    =   "CCWCursor"
      opts_26         =   31
      Name_26         =   "Cursor-2"
      C[0]_26         =   16711680
      Event_26        =   2
      X_26            =   6
      Y_26            =   13
      XPos_26         =   2
      YPos_26         =   2
      PointIndex_26   =   -1
      ChrosshairStyle_26=   8
      LockPlot_26     =   0
      TrackMode_1     =   6
      GraphBackground_1=   0
      GraphFrame_1    =   27
      ClassName_27    =   "CCWDrawObj"
      opts_27         =   30
      Image_27        =   28
      ClassName_28    =   "CCWPictImage"
      opts_28         =   1280
      Rows_28         =   1
      Cols_28         =   1
      F_28            =   -2147483633
      B_28            =   -2147483633
      ColorReplaceWith_28=   8421504
      ColorReplace_28 =   8421504
      Tolerance_28    =   2
      Animator_27     =   0
      Blinker_27      =   0
      PlotFrame_1     =   29
      ClassName_29    =   "CCWDrawObj"
      opts_29         =   30
      C[1]_29         =   16777215
      Image_29        =   30
      ClassName_30    =   "CCWPictImage"
      opts_30         =   1280
      Rows_30         =   1
      Cols_30         =   1
      Pict_30         =   1
      F_30            =   -2147483633
      B_30            =   16777215
      ColorReplaceWith_30=   8421504
      ColorReplace_30 =   8421504
      Tolerance_30    =   2
      Animator_29     =   0
      Blinker_29      =   0
      Caption_1       =   31
      ClassName_31    =   "CCWDrawObj"
      opts_31         =   30
      C[0]_31         =   -2147483640
      Image_31        =   32
      ClassName_32    =   "CCWTextImage"
      font_32         =   0
      Animator_31     =   0
      Blinker_31      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
      Bindings_1      =   33
      ClassName_33    =   "CCWBindingHolderArray"
      Editor_33       =   34
      ClassName_34    =   "CCWBindingHolderArrayEditor"
      Owner_34        =   1
      Annotations_1   =   35
      ClassName_35    =   "CCWAnnotations"
      Editor_35       =   36
      ClassName_36    =   "CCWAnnotationArrayEditor"
      Owner_36        =   1
      AnnotationTemplate_1=   37
      ClassName_37    =   "CCWAnnotation"
      opts_37         =   63
      Name_37         =   "[Template]"
      Plot_37         =   22
      Text_37         =   "[Template]"
      TextXPoint_37   =   6.7
      TextYPoint_37   =   6.7
      TextColor_37    =   16777215
      TextFont_37     =   38
      ClassName_38    =   "CCWFont"
      bFont_38        =   -1  'True
      BeginProperty Font_38 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShapeXPoints_37 =   39
      ClassName_39    =   "CDataBuffer"
      Type_39         =   5
      m_cDims;_39     =   1
      m_cElts_39      =   1
      Element[0]_39   =   3.3
      ShapeYPoints_37 =   40
      ClassName_40    =   "CDataBuffer"
      Type_40         =   5
      m_cDims;_40     =   1
      m_cElts_40      =   1
      Element[0]_40   =   3.3
      ShapeFillColor_37=   16777215
      ShapeLineColor_37=   16777215
      ShapeLineWidth_37=   1
      ShapeLineStyle_37=   1
      ShapePointStyle_37=   10
      ShapeImage_37   =   41
      ClassName_41    =   "CCWDrawObj"
      opts_41         =   62
      Image_41        =   42
      ClassName_42    =   "CCWPictImage"
      opts_42         =   1280
      Rows_42         =   1
      Cols_42         =   1
      Pict_42         =   7
      F_42            =   -2147483633
      B_42            =   -2147483633
      ColorReplaceWith_42=   8421504
      ColorReplace_42 =   8421504
      Tolerance_42    =   2
      Animator_41     =   0
      Blinker_41      =   0
      ArrowVisible_37 =   -1  'True
      ArrowColor_37   =   16777215
      ArrowWidth_37   =   1
      ArrowLineStyle_37=   1
      ArrowHeadStyle_37=   1
   End
   Begin VB.CommandButton cmdViewFile 
      Caption         =   "&View Data File"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearGraph 
      Caption         =   "Clear &Graph"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Data File"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtDataFile 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   " "
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblPlotData 
      Alignment       =   2  'Center
      Caption         =   "X axis in minutes"
      Height          =   255
      Left            =   1320
      TabIndex        =   51
      Top             =   5520
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
      Left            =   3120
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "R"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   11
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "I"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "N"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "E"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblYAxis 
      Alignment       =   2  'Center
      Caption         =   "D"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label lblDataFile 
      Alignment       =   2  'Center
      Caption         =   "Data File Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   5280
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
    
    ml_SensorColors(1) = vbBlue
    ml_SensorColors(2) = vbRed
    ml_SensorColors(3) = vbYellow
    ml_SensorColors(4) = vbGreen
    ml_SensorColors(5) = vbCyan
    ml_SensorColors(6) = vbMagenta
    ml_SensorColors(7) = vbBlack
    ml_SensorColors(8) = vbGrayText
    
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
