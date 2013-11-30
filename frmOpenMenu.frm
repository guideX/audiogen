VERSION 5.00
Begin VB.Form frmOpenMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   128
      X2              =   128
      Y1              =   0
      Y2              =   192
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   6
      X2              =   123
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   6
      X2              =   123
      Y1              =   21
      Y2              =   21
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   -8
      X2              =   128
      Y1              =   52
      Y2              =   52
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   112
      Y2              =   -8
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   128
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblVideo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Video"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   555
      Width           =   2175
   End
   Begin VB.Label lblSupportedTypes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Supported Types"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label lblAudio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Audio"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2175
   End
End
Attribute VB_Name = "frmOpenMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lMenus.mUsingSubMenu = True
AlwaysOnTop Me, True
lMenus.mOpenMenuVisible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
AlwaysOnTop Me, False
lMenus.mOpenMenuVisible = False
lMenus.mUsingSubMenu = False
End Sub

Private Sub lblAudio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
LabelMouseDown lblAudio, Button
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAudio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblAudio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mOpenMenuIndex = 2
LabelMouseMove lblAudio, Button
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAudio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblAudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblAudio, Button) = True Then
    Unload frmOpenMenu
    Unload frmFileMenu
    PromptOpen True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSupportedTypes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblSupportedTypes.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSupportedTypes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSupportedTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mOpenMenuIndex = 1
If Button = 0 Then
    RefreshMenus
    If lblSupportedTypes.BackColor <> &HC0C0C0 Then
        lblSupportedTypes.BackColor = &HC0C0C0
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSupportedTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSupportedTypes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblSupportedTypes, Button) = True Then
    Unload frmOpenMenu
    Unload frmFileMenu
    PromptOpen
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSupportedTypes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblVideo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
LabelMouseDown lblVideo, Button
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblVideo_Click()"
End Sub

Private Sub lblVideo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mOpenMenuIndex = 3
LabelMouseMove lblVideo, Button
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblVideo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblVideo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblVideo, Button) = True Then
    Unload frmOpenMenu
    Unload frmFileMenu
    PromptOpen False, True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblVideo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub
