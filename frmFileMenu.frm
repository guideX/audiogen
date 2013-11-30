VERSION 5.00
Begin VB.Form frmFileMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   96
      X2              =   96
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1
      X2              =   1
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Image imgDecodeArrow 
      Height          =   195
      Left            =   1155
      Picture         =   "frmFileMenu.frx":000C
      Top             =   600
      Width           =   195
   End
   Begin VB.Image imgConvertArrow 
      Height          =   195
      Left            =   1155
      Picture         =   "frmFileMenu.frx":0256
      Top             =   345
      Width           =   195
   End
   Begin VB.Image imgOpenArrow 
      Height          =   195
      Left            =   1155
      Picture         =   "frmFileMenu.frx":04A0
      Top             =   60
      Width           =   195
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   90
      Y1              =   89
      Y2              =   89
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   8
      X2              =   90
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Image imgExit 
      Height          =   210
      Left            =   3960
      Picture         =   "frmFileMenu.frx":06EA
      Top             =   2160
      Width           =   210
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   141
      X2              =   227
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   141
      X2              =   227
      Y1              =   269
      Y2              =   269
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   8
      X2              =   90
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   90
      Y1              =   21
      Y2              =   21
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   -8
      X2              =   128
      Y1              =   107
      Y2              =   107
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   136
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        &Exit"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   2055
   End
   Begin VB.Label lblRip 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        &Rip"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   1065
      Width           =   2055
   End
   Begin VB.Label lblDecode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        &Effects"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   585
      Width           =   2055
   End
   Begin VB.Label lblBurn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        &Burn"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   825
      Width           =   2055
   End
   Begin VB.Image imgOpen 
      Height          =   210
      Left            =   4080
      Picture         =   "frmFileMenu.frx":085E
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblNormalize 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        &Normalize"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4080
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        &Search"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4080
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblPlay 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        &Play"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   5160
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblPause 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        Pause"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4200
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label lblStop 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        Stop"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4080
      TabIndex        =   9
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label lblMerge 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        &Make Album"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   5160
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Convert"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   345
      Width           =   2055
   End
   Begin VB.Label lblOpen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        &Open"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   45
      Width           =   2055
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndfilemenu_startup).txt")
AlwaysOnTop Me, True
Me.Icon = frmGraphics.Icon
lMenus.mFileMenuVisible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
RefreshMenus
lMenus.mFileMenuIndex = 0
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
lMenus.mFileMenuVisible = False
frmMain.imgFile.Visible = False
frmMain.imgFileOver.Visible = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblExit.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    lblExit.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblExit.BackColor = &H8000000F
    Unload Me
    End
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblOpen.BackColor = vbWhite
    imgOpen.Picture = frmGraphics.imgOpen3.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 2
If Button = 0 Then
    imgOpen.Picture = frmGraphics.imgOpen2.Picture
    RefreshMenus
    lblOpen.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    'frmMain.imgFile.Picture = frmMain.imgFile.Picture
    lblOpen.BackColor = &H8000000F
    Me.Visible = False
    PromptOpen
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblExit.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 100
If lMenus.mOpenMenuVisible = True Then
    Unload frmOpenMenu
End If
If lMenus.mConvertMenuVisible = True Then
    Unload frmMenuConvert
End If
'CheckMenus
ResetFileMenuArrows
If Button = 0 Then
    lblExit.BackColor = &HC0C0C0
    RefreshMenus
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblExit.BackColor = &H8000000F
    Unload Me
    EndProgram
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMerge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblMerge.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMerge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMerge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 11
If Button = 0 Then
    RefreshMenus
    lblMerge.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMerge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMerge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblMerge.BackColor = &H8000000F
    MergeFiles
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMerge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblRip.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 1
'If imgRipArrow.Picture <> frmGraphics.imgMenuArrow3.Picture Then
'    ResetFileMenuArrows
'    imgRipArrow.Picture = frmGraphics.imgMenuArrow3.Picture
'End If
If Button = 0 Then
    RefreshMenus
    lblRip.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblRip.BackColor = &H8000000F
    StartRip
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblBurn.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblBurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 4
ResetFileMenuArrows
If LabelMouseMove(frmFileMenu.lblBurn, Button) = True Then
    If lMenus.mEffectsMenuVisible Then Unload frmMenuEffects
End If
If Button = 0 Then
    RefreshMenus
    lblBurn.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblBurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblBurn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblDecode.BackColor = &H8000000F
    Unload Me
    AdjustStatus sBurn
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblBurn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDecode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If LabelMouseDown(lblDecode, Button) = True Then
        imgDecodeArrow.Picture = frmGraphics.imgMenuArrow2.Picture
        If lMenus.mConvertMenuVisible = True Then Unload frmMenuConvert
        frmMenuEffects.left = frmFileMenu.left + frmFileMenu.Width
        frmMenuEffects.Visible = True
        lblDecode.BackColor = vbWhite
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDecode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 3
If imgDecodeArrow.Picture <> frmGraphics.imgMenuArrow3.Picture Then
    ResetFileMenuArrows
    imgDecodeArrow.Picture = frmGraphics.imgMenuArrow3.Picture
End If
If LabelMouseMove(frmFileMenu.lblDecode, Button) = True Then
    If lMenus.mConvertMenuVisible = True Then Unload frmMenuConvert
    frmMenuEffects.Visible = True
    frmMenuEffects.top = frmMain.top + 950
    frmMenuEffects.left = frmFileMenu.left + frmFileMenu.Width
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDecode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, i As Integer
If Button = 1 Then
    lblDecode.BackColor = &H8000000F
    'Unload Me
    msg = frmMain.tvwToBurn.SelectedItem.Text
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If Len(lFiles.fFile(i).fFilename) <> 0 Then
            If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
                msg2 = msg
                msg3 = msg2
                msg3 = left(msg2, Len(msg2) - 4) & ".wav"
                msg = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg2))
                DecodeFile msg, msg3, msg2
            End If
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblNormalize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblNormalize.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblNormalize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblNormalize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 5
If Button = 0 Then
    RefreshMenus
    lblNormalize.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblNormalize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblNormalize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If Button = 1 Then
    lblNormalize.BackColor = &H8000000F
    msg = frmMain.tvwToBurn.SelectedItem.Text
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If i <> 0 Then
            If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
                NormalizeFile lFiles.fFile(i).fFilename
            End If
        End If
    End If
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblNormalize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lMenus.mConvertMenuVisible = True Then
    Unload frmMenuConvert
End If
If LabelMouseDown(lblOpen, Button) = True Then
    imgOpenArrow.Picture = frmGraphics.imgMenuArrow2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblOpen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lMenus.mConvertMenuVisible = True Then
    Unload frmMenuConvert
End If
lMenus.mFileMenuIndex = 2
If Button <> 1 And imgOpenArrow.Picture <> frmGraphics.imgMenuArrow3.Picture Then
    ResetFileMenuArrows
    imgOpenArrow.Picture = frmGraphics.imgMenuArrow3.Picture
End If
If LabelMouseMove(lblOpen, Button) = True Then
'If Button = 0 Then
'    RefreshMenus
'    imgOpen.Picture = frmGraphics.imgOpen2.Picture
    'If lblOpen.BackColor <> &HC0C0C0 Then
    'lblOpen.BackColor = &HC0C0C0
    'Pause 0.2
    frmOpenMenu.top = frmFileMenu.top
    frmOpenMenu.left = frmFileMenu.left + frmFileMenu.Width
    'frmOpenMenu.SetFocus
    '    frmOpenMenu.SetFocus
    'End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If Button = 1 Then
'    lblOpen.BackColor = &H8000000F
If LabelMouseUp(lblOpen, Button) = True Then
    frmOpenMenu.Visible = True
    frmOpenMenu.top = frmMain.top + frmFileMenu.top + 10
    frmOpenMenu.left = frmMain.left + frmFileMenu.left + 1450
    'frmOpenMenu.SetFocus
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblPause.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 8
If Button = 0 Then
    RefreshMenus
    lblPause.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblPause.BackColor = &H8000000F
    AdjustStatus sPaused
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblPlay.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 7
If Button = 0 Then
    RefreshMenus
    lblPlay.BackColor = &HC0C0C0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblPlay.BackColor = &H8000000F
    QuickPlay lFiles.fFile(FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text)).fFilename
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRandom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
'    If LabelMouseDown(lblRandom, Button) = True Then
    imgConvertArrow.Picture = frmGraphics.imgMenuArrow2.Picture
    If lMenus.mOpenMenuVisible = True Then Unload frmOpenMenu
    frmMenuConvert.top = frmMain.top + 750
    frmMenuConvert.left = frmFileMenu.left + frmFileMenu.Width
    frmMenuConvert.Visible = True
    lblRandom.BackColor = vbWhite
'    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRandom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRandom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 10
frmMain.lblStatus.Caption = lMenus.mEffectsMenuVisible
If lMenus.mEffectsMenuVisible = True Then
    Unload frmMenuEffects
End If
If imgConvertArrow.Picture <> frmGraphics.imgMenuArrow3.Picture Then
    ResetFileMenuArrows
    imgConvertArrow.Picture = frmGraphics.imgMenuArrow3.Picture
End If
If LabelMouseMove(frmFileMenu.lblRandom, Button) = True Then
    If lMenus.mOpenMenuVisible = True Then Unload frmOpenMenu
    frmMenuConvert.Visible = True
    frmMenuConvert.top = frmMain.top + 750
    frmMenuConvert.left = frmFileMenu.left + frmFileMenu.Width
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRandom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRandom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblRandom.BackColor = &H8000000F
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRandom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblSearch.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 6
If Button = 0 Then
    lblSearch.BackColor = &HC0C0C0
    RefreshMenus
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblSearch.BackColor = &H8000000F
    frmSearch.Show
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblStop.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mFileMenuIndex = 9
If Button = 0 Then
    lblStop.BackColor = &HC0C0C0
    RefreshMenus
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblStop.BackColor = &H8000000F
    AdjustStatus sStop
    AdjustStatus sIdle
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub
