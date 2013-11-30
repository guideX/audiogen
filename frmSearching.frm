VERSION 5.00
Begin VB.Form frmSearching 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4650
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
   Icon            =   "frmSearching.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait, searching for media ..."
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgRecord 
      Height          =   465
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmSearching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
imgRecord.Picture = frmGraphics.imgRecord.Picture
AlwaysOnTop Me, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

