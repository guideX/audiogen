VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDFCF4A3-AD96-11D4-9959-0050BACD4F4C}#1.0#0"; "MDec.ocx"
Object = "{34B82A63-9874-11D4-9E66-0020780170C6}#1.0#0"; "MEnc.ocx"
Begin VB.Form frmKHZConverter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen - Converting Bitrate"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKHZConverter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrKillWave 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1200
      Top             =   2160
   End
   Begin VB.Timer tmrStartEncode 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   720
      Top             =   2160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MDECLib.MDec MDec1 
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin MENCLib.MEnc MEnc1 
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.TextBox txtFilename 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CheckBox chkStep2 
      Appearance      =   0  'Flat
      Caption         =   "Complete"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   930
      Width           =   1215
   End
   Begin VB.CheckBox chkStep1 
      Appearance      =   0  'Flat
      Caption         =   "Complete"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   690
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   -10
      Width           =   3615
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate Conversion in progress"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image imgRecord 
         Height          =   465
         Left            =   120
         Stretch         =   -1  'True
         Top             =   80
         Width           =   465
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   3880
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   -120
      X2              =   3880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblFilename 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Step 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Step 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmKHZConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ConvertKHZ()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = txtFilename.Text
If Len(msg) <> 0 Then
    If Right(LCase(msg), 4) = ".wav" Then
        chkStep1.value = 1
        msg = left(msg, Len(msg) - 4) & ".mp3"
        EncodeKHZ txtFilename.Text, msg
    ElseIf Right(LCase(msg), 4) = ".mp3" Then
        msg = left(msg, Len(msg) - 4) & ".wav"
        DecodeKHZ txtFilename.Text, msg
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdConvert_Click()"
End Sub

Private Sub DecodeKHZ(lInputFile As String, lOutputFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult, i As Integer, d As Integer, msg As String, msg2 As String, lPath As String
If Len(lInputFile) <> 0 And Len(lOutputFile) <> 0 Then
    If DoesFileExist(lInputFile) = True Then
        If DoesFileExist(lOutputFile) = True Then
            Kill lOutputFile
            DoEvents
        End If
        Select Case Right(LCase(lInputFile), 4)
        Case ".mp3"
            MDec1.OPENFILENAME = lInputFile
            MDec1.savefilename = lOutputFile
            MDec1.Decode
        End Select
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub DecodeKHZ(lInputFile As String, lOutputFile As String)"
End Sub

Private Sub EncodeKHZ(lInputFile As String, lOutputFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, mbox As VbMsgBoxResult, b As Long, i As Long
If Len(lInputFile) <> 0 And Len(lOutputFile) <> 0 Then
    If DoesFileExist(lInputFile) = True Then
        If DoesFileExist(lOutputFile) = True Then
            Kill lOutputFile
            DoEvents
        End If
        msg = lInputFile
        msg = GetFileTitle(msg)
        msg2 = lOutputFile
        msg2 = GetFileTitle(msg2)
        b = 320000
        i = frmMain.cboBitrate.ListIndex
        i = i * 8000
        b = b - i
        MEnc1.bitrate = b
        MEnc1.channels = 0
        MEnc1.OPENFILENAME = lInputFile
        MEnc1.savefilename = lOutputFile
        MEnc1.Encode
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function EncodeFile(lInputFile As String, lOutputFile As String) As String"
End Sub

Private Sub cmdConvert_Click()
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
imgRecord.Picture = frmGraphics.imgRecord.Picture
lSettings.sConvertingKHZ = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lSettings.sConvertingKHZ = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub MDec1_PercentDone(ByVal nPercent As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProgressBar1.value = nPercent
If nPercent = 100 Then
    chkStep1.value = 1
    tmrStartEncode.Enabled = True
    ProgressBar1.value = 0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub MDec1_PercentDone(ByVal nPercent As Long)"
End Sub

Private Sub MEnc1_PercentDone(ByVal nPercent As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProgressBar1.value = nPercent
If nPercent = 100 Then
    chkStep2.value = 1
    Pause 0.2
    tmrKillWave.Enabled = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub MEnc1_PercentDone(ByVal nPercent As Long)"
End Sub

Private Sub tmrKillWave_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Kill left(txtFilename.Text, Len(txtFilename.Text) - 4) & ".wav"
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrKillWave_Timer()"
End Sub

Private Sub tmrStartEncode_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
tmrStartEncode.Enabled = False
msg = txtFilename.Text
msg = left(txtFilename.Text, Len(txtFilename.Text) - 4) & ".wav"
Pause 0.2
EncodeKHZ msg, txtFilename.Text
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrStartEncode_Timer()"
End Sub
