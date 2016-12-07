VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About STI FiberTrack"
   ClientHeight    =   5295
   ClientLeft      =   -420
   ClientTop       =   795
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Height          =   195
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label lblEmail 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      ToolTipText     =   "Click this link to send e-mail to STC Online Fiber Monitoring System"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblTelephone 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblWarning 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Warning!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   5
      Top             =   4320
      Width           =   5415
   End
   Begin VB.Label lblBuildNum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Build 2.05.00.02"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblSubtitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Online Fiber Monitoring System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lblProduct 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STI FiberTrack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'cmdCancel is below the bottom edge of the Form.
'Its Visible property is set to True, as is its Cancel property.
'Thus, the user can press Escape to close the About dialog.
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim copy(0 To 3)        As String
    Dim address(0 To 3)     As String
    Dim warning(0 To 3)     As String
    Dim logo(0 To 3)        As String
    Dim iSavedMousePointer  As Integer
    iSavedMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Me.Picture = GetCustomPicture(5001)        '17KB Jpeg vs. 489KB Bmp
    Dim name As String
    name = GetIniSetting("Constants", "Name")
    Me.Caption = StringFormat("About {0} {1}", name, GetIniSetting("Constants", "ProductName"))
    lblSubtitle.Caption = GetIniSetting("Constants", "ProductName")

    Screen.MousePointer = iSavedMousePointer
    lblProduct.Caption = name
    lblProduct.Caption = StringFormat("{0} - {1}", name, GetIniSetting("Constants", "Version"))
    
    ' See Project -> Properties
    lblBuildNum.Caption = StringFormat("Build {0}.{1}.{2}.{3}", App.Major, App.Minor, App.Revision, GetIniSetting("Application", "Build"))
    copy(0) = StringFormat("{0} - {1} - {2}{3}", GetIniSetting("Constants", "Name"), GetIniSetting("Constants", "ProductName"), GetIniSetting("Constants", "Version"), vbCrLf)
    copy(1) = StringFormat("Copyright © 2004-{0} Sensatus Technologies Corporated.{1}", Year(Now), vbCrLf)
    copy(2) = StringFormat("All rights reserved.{0}", vbCrLf)
    copy(3) = Date
    lblCopyright.Caption = copy(0) & copy(1) & copy(2) & copy(3)
    
    address(0) = "Sensatus Technologies Corporation" & vbCrLf
    address(1) = "450 Edgell Road" & vbCrLf
    address(2) = "Framingham, MA 01701" & vbCrLf
    address(3) = "USA"
        
    warning(0) = "Warning!  This computer program is protected by copyright" & vbCrLf
    warning(1) = "law and international treaties.  Unauthorized reproduction or" & vbCrLf
    warning(2) = "distribution of this program is illegal and may result in" & vbCrLf
    warning(3) = "criminal and civil penalties."
    
    lblAddress.Caption = address(0) & address(1) & address(2) & address(3)
    
    lblTelephone.Caption = "Telephone: 781-555-0000"
    lblTelephone.Caption = "E-mail: jpiso@aol.com"
    lblWarning.Caption = warning(0) & warning(1) & warning(2) & warning(3)
    '"Warning!  This computer program is protected by copyright law and international treaties.  Unauthorized reproduction or distribution of this program is illegal and may result in criminal and civil penalties."

End Sub

Private Sub lblEmail_Click()
    OpenEmail "jpiso@aol.com"
End Sub
