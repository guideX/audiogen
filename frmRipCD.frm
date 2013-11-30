VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRipCD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen - CD Ripper"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4890
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
   Icon            =   "frmRipCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdNext 
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Next >"
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
      MICON           =   "frmRipCD.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdFinish 
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Finish"
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
      MICON           =   "frmRipCD.frx":0028
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
      Height          =   375
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      MICON           =   "frmRipCD.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraRipper 
      BorderStyle     =   0  'None
      Caption         =   "Step 1"
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtAlbum 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtArtist 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Step 1: Artist and Album names"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label lblArtist 
         Caption         =   "Artist:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblAlbum 
         Caption         =   "Album:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame fraRipper 
      BorderStyle     =   0  'None
      Caption         =   "Step 2"
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtCopyTo 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   4005
      End
      Begin VB.ComboBox cboBitrate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmRipCD.frx":0060
         Left            =   600
         List            =   "frmRipCD.frx":00DC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   4005
      End
      Begin VB.ComboBox cboFormat 
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmRipCD.frx":0485
         Left            =   600
         List            =   "frmRipCD.frx":048F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   4005
      End
      Begin OsenXPCntrl.OsenXPButton cmdCopyTo 
         Height          =   375
         Left            =   600
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Select"
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
         MICON           =   "frmRipCD.frx":04B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblSaveLocation 
         Caption         =   "Path:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Step 2: Format, bitrate, and save to location"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label lblBitrate 
         Caption         =   "Bitrate:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFormat 
         Caption         =   "Format:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame fraRipper 
      BorderStyle     =   0  'None
      Caption         =   "Step 3"
      ForeColor       =   &H00404040&
      Height          =   2295
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      Begin OsenXPCntrl.OsenXPButton cmdChange 
         Height          =   375
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Change"
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
         MICON           =   "frmRipCD.frx":04D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.TreeView tvwToRip 
         Height          =   1575
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2778
         _Version        =   393217
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Step 3: Select Filenames"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Step 2: Select Format, bitrate and save to location"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
End
Attribute VB_Name = "frmRipCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBitrate_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtCopyTo.Text) <> 0 And Len(cboBitrate.Text) <> 0 Then
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = False
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboBitrate_Change()"
End Sub

Private Sub cmdCancel_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdCancel_Click()"
End Sub

Private Sub cmdChange_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim ibox As String
ibox = InputBox("Please select a new filename", "Select Filename", tvwToRip.SelectedItem.Text)
If Len(ibox) <> 0 Then
    If Right(LCase(ibox), 4) <> ".wav" Then ibox = ibox & ".wav"
    tvwToRip.SelectedItem.Text = ibox
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdChange_Click()"
End Sub

Private Sub cmdCopyTo_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmSelectDir.Show 1
txtCopyTo.Text = lSettings.sSelectDirReturnValue
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdCopyTo_Click()"
End Sub

Private Sub cmdFinish_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, b As Integer
msg = txtCopyTo.Text
For i = 1 To tvwToRip.Nodes.Count
    If Len(tvwToRip.Nodes(i).Text) <> 0 And tvwToRip.Nodes(i).Checked = True Then
        b = b + 1
        lBurnQue.bFiles(b) = msg & tvwToRip.Nodes(i).Text
    End If
Next i
lBurnQue.bCount = b
DoEvents
MkDir msg
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdFinish_Click()"
End Sub

Private Sub cmdNext_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If fraRipper(0).Visible = True Then
    fraRipper(0).Visible = False
    fraRipper(1).Visible = True
    txtCopyTo.Text = App.Path & "\cdcopy\" & txtArtist.Text & " - " & txtAlbum.Text & "\"
    cboBitrate.ListIndex = 8
    cmdNext.Enabled = True
ElseIf fraRipper(1).Visible = True Then
    fraRipper(0).Visible = False
    fraRipper(1).Visible = False
    fraRipper(2).Visible = True
    cmdNext.Enabled = False
    cmdFinish.Enabled = True
    Dim i As Integer
ElseIf fraRipper(2).Visible = True Then

End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdNext_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'InitRipper
Me.Icon = frmGraphics.Icon
cboFormat.ListIndex = 0
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub txtAlbum_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtArtist.Text) And Len(txtAlbum.Text) <> 0 Then
    cmdNext.Enabled = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub txtAlbum_Change()"
End Sub

Private Sub txtArtist_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtArtist.Text) And Len(txtAlbum.Text) <> 0 Then
    cmdNext.Enabled = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub txtArtist_Change()"
End Sub

Private Sub txtCopyTo_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(txtCopyTo.Text) <> 0 And Len(cboBitrate.Text) <> 0 Then
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = False
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub txtCopyTo_Change()"
End Sub
