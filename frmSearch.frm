VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton OsenXPButton2 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Save"
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
      MICON           =   "frmSearch.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton1 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Add All"
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
      MICON           =   "frmSearch.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Add"
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
      MICON           =   "frmSearch.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton ok 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmSearch.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstResults 
      ForeColor       =   &H00404040&
      Height          =   1980
      IntegralHeight  =   0   'False
      Left            =   1200
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox cboSearchIn 
      ForeColor       =   &H00404040&
      Height          =   315
      ItemData        =   "frmSearch.frx":007C
      Left            =   1200
      List            =   "frmSearch.frx":0086
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin OsenXPCntrl.OsenXPButton btnSearch 
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Search"
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
      MICON           =   "frmSearch.frx":00A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSearchText 
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4320
      Y1              =   3975
      Y2              =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4320
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4320
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search In:"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For:"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSearch_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, msg As String, msg2 As String, f As Integer
Dim cCol1 As tSearch
Select Case cboSearchIn.ListIndex
Case 0
    msg = txtSearchText.Text
    If Len(msg) <> 0 Then
        lstResults.Clear
        For i = 0 To lFiles.fCount
            msg2 = lFiles.fFile(i).fFilename
            msg2 = GetFileTitle(msg2)
            If Len(msg2) <> 0 Then
                If InStr(1, msg2, msg, vbTextCompare) Then
                    lstResults.AddItem msg2
                End If
            End If
        Next i
    Else
        MsgBox "You must input a search string", vbExclamation, "Error"
        Exit Sub
    End If
Case 1
    For i = 0 To lDrives.dCount
        If lDrives.dDrive(i).dDriveType = dHardDrive Then
            If Len(lDrives.dDrive(i).dDriveLetter) <> 0 Then
                GetFiles lDrives.dDrive(i).dDriveLetter & "\", txtSearchText.Text, vbDirectory, cCol1
                For f = 0 To cCol1.Count
                    If Len(cCol1.Path(f)) <> 0 Then
                        lstResults.AddItem cCol1.Path(f)
                    End If
                Next f
            End If
        End If
    Next i
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub btnSearch_Click()"
End Sub

Private Sub cboSearchIn_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case cboSearchIn.ListIndex
Case 0
    If Len(txtSearchText.Text) = 0 Or txtSearchText.Text = "*.mp3" Then txtSearchText.Text = ""
Case 1
    If Len(txtSearchText.Text) = 0 Then txtSearchText.Text = "*.mp3"
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboSearchIn_Click()"
End Sub

Private Sub cmdOK_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmMain.tvwToBurn.Nodes.Add , , , lstResults.Text
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdOK_Click()"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
cboSearchIn.ListIndex = 0
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub ok_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ok_Click()"
End Sub

Private Sub OsenXPButton1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
For i = 0 To lstResults.ListCount
    If Len(lstResults.List(i)) <> 0 Then
        frmMain.tvwToBurn.Nodes.Add , , , lstResults.List(i)
    End If
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub OsenXPButton1_Click()"
End Sub

Private Sub OsenXPButton2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If lstResults.ListCount <> 0 Then
    For i = 0 To lstResults.ListCount
        If Len(lstResults.List(i)) <> 0 Then
            msg = lstResults.List(i) & vbCrLf & msg & vbCrLf
        End If
    Next i
End If
msg = Trim(msg)
If Len(msg) <> 0 Then
    msg2 = SaveDialog(Me, "Playlist Files (*.m3u)|*.m3u|All Files (*.*)|*.*|", "Save playlist as ...", CurDir)
    msg2 = left(msg2, Len(msg2) - 1) & ".m3u"
    If Len(msg2) <> 0 Then SaveFile msg2, msg
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub OsenXPButton2_Click()"
End Sub

