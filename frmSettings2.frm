VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Audiogen - Settings"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGeneralOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtPassword 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   3460
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   3100
         Width           =   2175
      End
      Begin VB.CheckBox chkNormalize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Normalize"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboCDDrive 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2355
         Width           =   2175
      End
      Begin VB.CheckBox chkFullScreenVideo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Full Screen Video"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox chkFirstRun 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "First Run"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkFinalize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Finalize"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboBitrate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmSettings2.frx":1708A
         Left            =   960
         List            =   "frmSettings2.frx":17106
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   660
         Width           =   2175
      End
      Begin VB.CheckBox chkDebugMode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Debug mode"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkCheckTaskbarStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check Taskbar Status"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkAutoEject 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto eject"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Always on top"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CD Drive:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "frmSettings2.frx":174AF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4080
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
