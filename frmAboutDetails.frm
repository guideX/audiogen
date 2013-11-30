VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmAboutDetails 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Audiogen"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   3555
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3615
      Begin VB.Image imgRecord 
         Height          =   705
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   705
      End
      Begin VB.Label lblAppTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Audiogen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "<Design mode>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   4000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4000
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "© 2003-2004 Team Nexgen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   4590
      Width           =   3615
      Begin VB.CheckBox chkShowOnStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show on startup"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin OsenXPCntrl.OsenXPButton OsenXPButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   2480
         TabIndex        =   12
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmAboutDetails.frx":1708A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   -120
      X2              =   3880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   3880
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Label lblProtected 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is protected by copyright law and international treaties"
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblNexgen 
      BackStyle       =   0  'Transparent
      Caption         =   "Team Nexgen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmAboutDetails.frx":170A6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblPublisher 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblKnightFal 
      BackStyle       =   0  'Transparent
      Caption         =   "KnightFal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmAboutDetails.frx":171F8
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblGuideX 
      BackStyle       =   0  'Transparent
      Caption         =   "|guideX|"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmAboutDetails.frx":1734A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblProgrammingTeam 
      BackStyle       =   0  'Transparent
      Caption         =   "Development Team"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblAppInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutDetails.frx":1749C
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
End
Attribute VB_Name = "frmAboutDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
imgRecord.Picture = frmGraphics.imgAudiogenLogo.Picture
'Me.Icon = frmGraphics.Icon
lblAppTitle.Caption = App.Title
lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & " build " & App.Revision
SetCheckBoxValue chkShowOnStartup, lSettings.sShowAboutDetailsOnStartup
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case chkShowOnStartup.value
Case 0
    lSettings.sShowAboutDetailsOnStartup = False
    WriteINI lIniFiles.iSettings, "Settings", "ShowAboutDetailsOnStartup", lSettings.sShowAboutDetailsOnStartup
Case 1
    lSettings.sShowAboutDetailsOnStartup = True
    WriteINI lIniFiles.iSettings, "Settings", "ShowAboutDetailsOnStartup", lSettings.sShowAboutDetailsOnStartup
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgRecord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgRecord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblAppInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAppInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblAppTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAppTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCopyright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCopyright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblGuideX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "mailto:guidex@team-nexgen.com", frmAboutDetails.hwnd
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblGuideX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblKnightFal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "mailto:knightfal@team-nexgen.com", frmAboutDetails.hwnd
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblKnightFal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblNexgen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.com", frmAboutDetails.hwnd
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblNexgen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblProgrammingTeam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblProgrammingTeam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblProtected_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblProtected_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPublisher_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPublisher_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub OsenXPButton1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub OsenXPButton1_Click()"
End Sub
