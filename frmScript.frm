VERSION 5.00
Begin VB.Form frmScript 
   Caption         =   "Script Editor"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5340
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtScript 
      ForeColor       =   &H00404040&
      Height          =   3975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSep8302789372 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuSep893928723 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "Execute"
      End
      Begin VB.Menu mnuSep38927890362 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
txtScript.Width = Me.ScaleWidth
txtScript.Height = Me.ScaleHeight
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Resize()"
End Sub

Private Sub mnuExecute_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then frmMain.ctlScript.ExecuteStatement txtScript.Text
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuExecute_Click()"
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuNew_Click()
txtScript.Text = ""
End Sub

Private Sub mnuOpen_Click()
Dim msg As String, msg2 As String
msg = OpenDialog(frmMain, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", "Open Script", App.Path & "\a_script\")
msg2 = ReadFile(msg)
If Len(msg2) <> 0 Then
    txtScript.Text = msg2
End If
End Sub

Private Sub mnuSaveAs_Click()
Dim msg As String, msg2 As String
msg2 = SaveDialog(frmScript, "Text Files (*.txt)|*.txt|", "Save as ...", CurDir)
msg2 = left(msg2, Len(msg2) - 1)
If Len(msg2) <> 0 Then
    msg = txtScript.Text
    SaveFile msg2, msg
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuSaveAs_Click()"
End Sub
