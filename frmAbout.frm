VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen - Splash Screen"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdWebsite 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Website"
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
      MICON           =   "frmAbout.frx":1708A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdHistory 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "History"
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
      MICON           =   "frmAbout.frx":170A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdErrorOptions 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Options"
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
      MICON           =   "frmAbout.frx":170C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmrEnable 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   6840
   End
   Begin VB.CheckBox chkShowSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show on startup"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   2055
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose2 
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmAbout.frx":170DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   6960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Start"
      ENAB            =   0   'False
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
      MICON           =   "frmAbout.frx":170FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdRegister 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Register"
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
      MICON           =   "frmAbout.frx":17116
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdDetails 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "About"
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
      MICON           =   "frmAbout.frx":17132
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   7650
      Left            =   0
      Top             =   0
      Width           =   6270
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Menu"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSep32896392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu mnuSep32789237263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lBackwards As Boolean
Dim lClosed As Boolean

Private Sub chkShowSplash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: Show on startup - If left unchecked, this window will no longer show up on startup."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub chkShowSplash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrEnable.Enabled = False
lSettings.sProcessScripts = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "ProcessScripts", False)
lClosed = True
Unload Me
frmMain.Show
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdClose_Click()"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: Start - Begins the Audiogen algorithm."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdClose2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
End
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdDetails_Click()"
End Sub

Private Sub cmdClose2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: Close will hide this window, after which Audiogen will close down"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdClose2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdDetails_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmAboutDetails.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdDetails_Click()"
End Sub

Private Sub cmdDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: About - Shows the About window which contains information about the authors of this program and other pertinant information."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdErrorOptions_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmErrorOptions.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdErrorOptions_Click()"
End Sub

Private Sub cmdErrorOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: Options - Displays the error options window, which can be used to change the way Audiogen deals with errors."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdErrorOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdHistory_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmReleaseHistory.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdHistory_Click()"
End Sub

Private Sub cmdHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: History - A log of Audiogen versions and their release dates."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdRegister_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrEnable.Enabled = False
AlwaysOnTop Me, False
frmRegister.Show 1
If lSettings.sRegistered = True Then
    With frmAbout
        .cmdRegister.Visible = False
        .tmrEnable.Enabled = False
        .cmdClose.Enabled = True
        .chkShowSplash.Enabled = True
    End With
    lblDescription.Caption = "Help: You have just registered your copy of Audiogen. As a result, you will no longer see and delays when this Window is loaded. You also now have the option to turn off this window."
End If
AlwaysOnTop Me, True
tmrEnable.Enabled = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdRegister_Click()"
End Sub

Private Sub cmdRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: Register - Displays the register window, which will enable you to purchase a full version of Audiogen."
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub cmdWebsite_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.com/", Me.hwnd
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdWebsite_Click()"
End Sub

Private Sub cmdWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblDescription.Caption = "Help: View the Team Nexgen Website"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_Load()
lSettings.sHandleErrors = ReadINI(App.Path & "\inis\errorlog.ini", "Settings", "HandleErrors", True)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, m As Boolean
Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
'Me.Icon = frmGraphics.Icon
Image1.Picture = frmGraphics.imgSplash.Picture
lIniFiles.iSettings = App.Path & "\inis\a_settings.ini"
lSettings.sName = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "Name", "")
lSettings.sPassword = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "Password", "")
lSettings.sShowSplash = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "ShowSplash", True)
lSettings.sShowAboutDetailsOnStartup = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "ShowAboutDetailsOnStartup", False)
If Len(lSettings.sName) <> 0 And Len(lSettings.sPassword) <> 0 And KeyGen(lSettings.sName, "pickles", 1) = lSettings.sPassword Then
    If lSettings.sShowSplash = True Then
        cmdDetails.Enabled = True
        cmdClose.Enabled = True
        chkShowSplash.Enabled = True
        chkShowSplash.value = 1
        cmdRegister.Visible = False
    Else
        lClosed = True
        Unload Me
        frmMain.Show
        Exit Sub
    End If
Else
    lSettings.sShowSplash = True
    chkShowSplash.Enabled = False
    chkShowSplash.value = 1
    tmrEnable.Enabled = True
End If
AlwaysOnTop Me, True
Me.SetFocus
If lSettings.sShowAboutDetailsOnStartup = True Then
    frmAboutDetails.Show
    AlwaysOnTop frmAboutDetails, True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lClosed = False Then
    Dim m As VbMsgBoxResult
    m = MsgBox("Are you sure you wish to end?", vbYesNo + vbQuestion)
    If m = vbYes Then
        End
    Else
        Cancel = 1
        Exit Sub
    End If
End If
If chkShowSplash.value = 0 Then
    WriteINI App.Path & "\inis\a_settings.ini", "Settings", "ShowSplash", False
Else
    WriteINI App.Path & "\inis\a_settings.ini", "Settings", "ShowSplash", True
End If
AlwaysOnTop Me, False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    FormDrag Me
Else
    PopupMenu mnuHidden
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(lblDescription.Caption) <> 0 Then lblDescription.Caption = ""
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    FormDrag Me
Else
    PopupMenu mnuHidden
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDescription_Click()"
End Sub

Private Sub mnuAbout_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmAboutDetails.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuAbout_Click()"
End Sub

Private Sub mnuClose_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
End
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuClose_Click()"
End Sub

Private Sub mnuHistory_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmReleaseHistory.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuHistory_Click()"
End Sub

Private Sub mnuOptions_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
frmErrorOptions.Show 1
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuOptions_Click()"
End Sub

Private Sub mnuRegister_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrEnable.Enabled = False
AlwaysOnTop Me, False
frmRegister.Show 1
If lSettings.sRegistered = True Then
    With frmAbout
        .cmdRegister.Visible = False
        .tmrEnable.Enabled = False
        .cmdClose.Enabled = True
        .chkShowSplash.Enabled = True
    End With
    lblDescription.Caption = "Help: You have just registered your copy of Audiogen. As a result, you will no longer see and delays when this Window is loaded. You also now have the option to turn off this window."
End If
AlwaysOnTop Me, True
tmrEnable.Enabled = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuRegister_Click()"
End Sub

Private Sub mnuStart_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrEnable.Enabled = False
lSettings.sProcessScripts = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "ProcessScripts", False)
lClosed = True
Unload Me
frmMain.Show
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuStart_Click()"
End Sub

Private Sub mnuWebsite_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Surf "http://www.team-nexgen.com/", Me.hwnd
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuWebsite_Click()"
End Sub

Private Sub tmrEnable_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
tmrEnable.Enabled = False
cmdClose.Enabled = True
cmdClose.SetFocus
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrEnable_Timer()"
End Sub
