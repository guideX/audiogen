VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmWinVer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen - Choose Windows Version"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWinVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdSelect 
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Select"
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
      MICON           =   "frmWinVer.frx":1708A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboWindowsVersion 
      Height          =   315
      ItemData        =   "frmWinVer.frx":170A6
      Left            =   120
      List            =   "frmWinVer.frx":170A8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmWinVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sWinVer = cboWindowsVersion.ListIndex
WriteINI lIniFiles.iSettings, "Settings", "WinVer", lSettings.sWinVer
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdSelect_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
cboWindowsVersion.Clear
cboWindowsVersion.AddItem "Other"
cboWindowsVersion.AddItem "Windows 95"
cboWindowsVersion.AddItem "Windows 98"
cboWindowsVersion.AddItem "Windows ME"
cboWindowsVersion.AddItem "Windows NT"
cboWindowsVersion.AddItem "Windows 2000"
cboWindowsVersion.AddItem "Windows XP"
cboWindowsVersion.AddItem "Windows 2003"
cboWindowsVersion.ListIndex = Int(lSettings.sWinVer)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub
