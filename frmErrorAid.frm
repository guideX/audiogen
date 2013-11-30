VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmErrorAid 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Error Aid - Audiogen"
   ClientHeight    =   1620
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrorAid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   393217
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSep836732659872 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmErrorAid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
TreeView1.Width = Me.ScaleWidth - 300
TreeView1.Height = Me.ScaleHeight - 200
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
End Sub

Private Sub mnuSaveAs_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String
For i = 0 To TreeView1.Nodes.Count
    If TreeView1.Nodes(i).Text <> 0 Then
        msg = msg & vbCrLf & TreeView1.Nodes(i).Text
    End If
Next i
msg2 = SaveDialog(Me, "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*|", "Save as ...", CurDir)
If Len(msg2) <> 0 Then
    msg2 = left(msg2, Len(msg2) - 1) & ".log"
    SaveFile msg2, msg
End If
End Sub

