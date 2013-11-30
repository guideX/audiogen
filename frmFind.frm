VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4830
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdFindNext 
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Find Next"
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
      MICON           =   "frmFind.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   500
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Cancel"
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
      MICON           =   "frmFind.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optSelect 
      Appearance      =   0  'Flat
      Caption         =   "Search in treeview"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
   End
   Begin VB.OptionButton optSelect 
      Appearance      =   0  'Flat
      Caption         =   "Search playlist"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblFind 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdFindNext_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, f As Integer, msg2 As String
If Len(txtFind.Text) <> 0 Then
    msg = txtFind.Text
    If optSelect(0).value = True Then
        frmMain.ResetFileTreeView
        For i = 0 To lFiles.fCount
            msg2 = lFiles.fFile(i).fFilename
            msg2 = GetFileTitle(msg2)
            If Len(msg2) <> 0 Then
                If InStr(1, LCase(msg2), LCase(msg), vbTextCompare) Then
                    frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2, 3
                End If
            End If
        Next i
    ElseIf optSelect(1).value = True Then
        For i = 1 To frmMain.tvwFiles.Nodes.Count
            If Len(frmMain.tvwFiles.Nodes(i).Text) <> 0 Then
                If InStr(1, LCase(frmMain.tvwFiles.Nodes(i).Text), LCase(msg), vbTextCompare) Then
                    frmMain.tvwFiles.Nodes(i).Selected = True
                    Unload Me
                    frmMain.tvwFiles.SetFocus
                    Exit For
                End If
            End If
        Next i
    ElseIf optSelect(2).value = True Then
    ElseIf optSelect(3).value = True Then
    ElseIf optSelect(4).value = True Then
    ElseIf optSelect(5).value = True Then
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdFindNext_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
optSelect(lSettings.sFindSelectIndex).value = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WriteINI lIniFiles.iSettings, "Settings", "FindSelectIndex", lSettings.sFindSelectIndex
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub optSelect_Click(Index As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If optSelect(Index).value = True Then lSettings.sFindSelectIndex = Index
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub optSelect_Click(Index As Integer)"
End Sub
