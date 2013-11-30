VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReleaseHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen - Release History"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReleaseHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Close"
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
      MICON           =   "frmReleaseHistory.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6588
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image imgRecord 
      Height          =   465
      Left            =   120
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmReleaseHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmGraphics.Icon
imgRecord.Picture = frmGraphics.imgRecord.Picture
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , , "Pre-Beta 1.0"
TreeView1.Nodes.Add , , , "Beta 1.0"
TreeView1.Nodes.Add , , , "Version 1.0"
TreeView1.Nodes.Add 1, tvwChild, , "AG 1.0 PB1 (08/01/2003)"
TreeView1.Nodes.Add 1, tvwChild, , "AG 1.0 PB2 (08/10/2003)"
TreeView1.Nodes.Add 1, tvwChild, , "AG 1.0 PB3 (08/22/2003)"
TreeView1.Nodes.Add 1, tvwChild, , "AG 1.0 PB4 (08/27/2003)"
TreeView1.Nodes.Add 1, tvwChild, , "AG 1.0 PB5 (09/02/2003)"
TreeView1.Nodes.Add 2, tvwChild, , "AG 1.0 B1 (09/13/2003)"
TreeView1.Nodes.Add 2, tvwChild, , "AG 1.0 B2 (09/26/2003)"
TreeView1.Nodes.Add 3, tvwChild, , "AG 1.0 59 (11/02/2003)"
TreeView1.Nodes.Add 3, tvwChild, , "AG 1.0 80 (12.01.2003)"
TreeView1.Nodes.Add 3, tvwChild, , "AG 1.0 90 (04.29.2004)"
'TreeView1.Nodes.Add 3, tvwChild, , "AG " & App.Major & "." & App.Minor & " " & App.Revision & " (" & Date & ")"
TreeView1.Nodes(1).Expanded = True
TreeView1.Nodes(2).Expanded = True
TreeView1.Nodes(3).Expanded = True
End Sub
