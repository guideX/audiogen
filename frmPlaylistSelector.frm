VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmPlaylistSelector 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen - Previous Filenames"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6780
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
   ScaleHeight     =   2805
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   14
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "frmPlaylistSelector.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5535
   End
   Begin VB.ListBox lstFilenames 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmPlaylistSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
frmTagEditor.txtFilename.Text = txtFilename.Text
lTag.tFile = txtFilename.Text
GetTagInfo
DoEvents
DisplayTag
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To lFiles.fCount
    If Len(lFiles.fFile(i).fFilename) <> 0 Then
        lstFilenames.AddItem lFiles.fFile(i).fFilename
    End If
Next i
End Sub

Private Sub lstFilenames_Click()
txtFilename.Text = lstFilenames.Text
End Sub
