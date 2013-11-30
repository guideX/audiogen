VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmErrorOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Audiogen?"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrorOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDoEverytime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Always run audiogen this way"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3015
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
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
      MICON           =   "frmErrorOptions.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   -30
      Width           =   4215
      Begin VB.OptionButton optNormalMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run in normal mode"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optBetaTestingMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Debug (Beta testing) mode"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Would you like to run Audiogen in Beta testing mode?"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   5000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5000
      Y1              =   975
      Y2              =   975
   End
End
Attribute VB_Name = "frmErrorOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If optNormalMode.value = True Then
    lSettings.sDebugMode = False
    WriteINI lIniFiles.iErrorLog, "Settings", "DebugMode", "False"
ElseIf optBetaTestingMode.value = True Then
    lSettings.sDebugMode = True
    WriteINI lIniFiles.iErrorLog, "Settings", "DebugMode", "True"
End If
If chkDoEverytime.value = 1 Then
    WriteINI lIniFiles.iErrorLog, "Settings", "DisplayErrorOptions", "False"
ElseIf chkDoEverytime.value = 0 Then
    WriteINI lIniFiles.iErrorLog, "Settings", "DisplayErrorOptions", "True"
End If
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim b As Boolean
Me.Icon = frmGraphics.Icon
b = ReadINI(lIniFiles.iErrorLog, "Settings", "DebugMode", True)
If b = True Then
    optBetaTestingMode.value = True
    optNormalMode.value = False
ElseIf b = False Then
    optBetaTestingMode.value = False
    optNormalMode.value = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub
