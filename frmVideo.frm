VERSION 5.00
Begin VB.Form frmVideo 
   Caption         =   "Audiogen"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   Icon            =   "frmVideo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8235
   WindowState     =   2  'Maximized
   Begin VB.Frame fraVideo 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
Form_Resize
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
fraVideo.Width = Me.ScaleWidth
fraVideo.Height = Me.ScaleHeight
If lPlayer.pStatus = sPlay Then
    PutMultimedia fraVideo.hwnd, lblAliasname.Caption, 0, 0, ScaleX(Val(fraVideo.Width), 1, 3), ScaleX(Val(fraVideo.Height), 1, 3)
    fraVideo.Refresh
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Resize()"
End Sub

Private Sub fraVideo_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub fraVideo_DblClick()"
End Sub
