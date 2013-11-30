VERSION 5.00
Begin VB.Form frmTaskbar 
   Caption         =   "Audiogen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaskbar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sAlwaysOnTop = False Then
    AlwaysOnTop frmMain, True
    AlwaysOnTop frmMain, False
End If
lTaskbar = True
If frmMain.WindowState = vbMinimized Then frmMain.WindowState = Normal
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_GotFocus()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.left = -2000
Me.top = -10000
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_LostFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lTaskbar = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_LostFocus()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
EndProgram
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_LostFocus()"
End Sub
