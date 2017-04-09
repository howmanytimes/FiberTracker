VERSION 5.00
Begin VB.Form BarChart 
   Caption         =   "BarChart"
   ClientHeight    =   5835
   ClientLeft      =   1650
   ClientTop       =   1725
   ClientWidth     =   7710
   Icon            =   "BarChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Return"
      Default         =   -1  'True
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.PictureBox MyChart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sensor 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "BarChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
'This example creates a grid in a PictureBox control
'and sets coordinates for the upper-left corner
'to -1, -1 instead of 0, 0.
'Every 0.25 second, dots are randomly plotted from the
'upper-left corner to the lower-right corner.
'To try this example, paste the code into the
'Declarations section of a form that contains a
'large PictureBox and a Timer control, and then press F5.
'    Timer1.Interval = 250   ' Set Timer interval.
'    Picture1.ScaleTop = -1  ' Set scale for top of grid.
'    Picture1.ScaleLeft = -1 ' Set scale for left of grid.
'    Picture1.ScaleWidth = 2 ' Set scale (-1 to 1).
'    Picture1.ScaleHeight = 2
'    Picture1.ForeColor = vbRed
'    Picture1.Line (-1, 0)-(1, 0)    ' Draw horizontal line.
'    Picture1.ForeColor = vbBlue
'    Picture1.Line (0, -1)-(0, 1)    ' Draw vertical line.
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyR And _
        (Shift And vbCtrlMask = vbCtrlMask) Then
        cmdReturn_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim iIndex As Integer
    Dim iInit As Integer
    Dim lColorMatrix(0 To 7) As Long
    
    With Me
        .Caption = StringFormat("Bar Chart({0}) - {1} - {2}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"))
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
    End With
    
    cmdReturn.Left = (Me.Width / 2) - (cmdReturn.Width / 2)
    MyChart.Left = (Me.Width / 2) - (MyChart.Width / 2)
    
    PlotData.Hide
    
    SensorColors(1) = RGB(0, 0, 255)    ' Blue
    SensorColors(2) = RGB(204, 0, 0)    ' Red
    SensorColors(3) = RGB(204, 255, 0)  ' Yellow
    SensorColors(4) = RGB(0, 102, 0)    ' Green
    SensorColors(5) = RGB(0, 255, 204)  ' Cyan
    SensorColors(6) = RGB(204, 0, 153)  ' Magenta
    SensorColors(7) = RGB(204, 102, 0)  ' Orange
    SensorColors(8) = RGB(0, 51, 0)     ' Darker Green

    lColorMatrix(0) = SensorColors(1) 'vbBlue
    lColorMatrix(1) = SensorColors(2) 'vbGreen
    lColorMatrix(2) = SensorColors(3) 'vbYellow
    lColorMatrix(3) = SensorColors(4) 'vbMagenta
    lColorMatrix(4) = SensorColors(5) 'vbCyan
    lColorMatrix(5) = SensorColors(6) 'vbWhite
    lColorMatrix(6) = SensorColors(7) 'vbRed
    lColorMatrix(7) = SensorColors(8) 'vbYellow

    iInit = 0
    lblName(iInit).BackColor = lColorMatrix(iInit)
    For iIndex = iInit + 1 To 7
        ' sensor names
        Load lblName(iIndex)
        lblName(iIndex).Top = lblName(iInit).Top
        lblName(iIndex).Left = (iIndex - 1) * (lblName(iInit).Width + 15) + 875
        lblName(iIndex).Caption = "Sensor " & iIndex + 1
        lblName(iIndex).BackColor = lColorMatrix(iIndex)
        lblName(iIndex).Visible = True
    Next iIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BarChart.Hide
    FiberTrack.Show
    FiberTrack.SetStatusText "Main window."
    ' Turn check mark on menu items on and off.
    FiberTrack.mnuMain.Checked = Not FiberTrack.mnuMain.Checked
    FiberTrack.mnuBarGraph.Checked = Not FiberTrack.mnuBarGraph.Checked
    IsMain = True
End Sub
