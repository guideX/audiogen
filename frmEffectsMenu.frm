VERSION 5.00
Begin VB.Form frmEffectsMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
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
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   96
      Y1              =   161
      Y2              =   161
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   8
      X2              =   96
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   104
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   104
      X2              =   104
      Y1              =   0
      Y2              =   216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   136
      Y1              =   1.333
      Y2              =   1.333
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1
      X2              =   0
      Y1              =   0
      Y2              =   216
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       FadeIN"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Echo"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Distortion"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   735
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       CFilter"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Chorus"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   1
      Top             =   285
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Amplitude"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   0
      Top             =   50
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Shifting"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Reverb"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblEffect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &FadeOUT"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblNormalize 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Normalize"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblPlay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Play"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2895
      Width           =   1575
   End
   Begin VB.Label lblSaveAs 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Save As ..."
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2655
      Width           =   1575
   End
   Begin VB.Label lblShowForm 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       &Show Window"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2430
      Width           =   1575
   End
End
Attribute VB_Name = "frmEffectsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, True
lMenus.mDecodeMenuVisible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_LostFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Unload Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_LostFocus()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
lMenus.mDecodeMenuVisible = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lblEffect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblEffect(Index), Button) = True Then
    lEffectsPresets.eEffectQueIndex = Index
    frmFileMenu.Visible = False
    frmEffects.Show
    Unload frmFileMenu
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblEffect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblEffect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseMove(lblEffect(Index), Button) = True Then
    lMenus.mDecodeMenuIndex = Index
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblEffect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblShowForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblShowForm, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblShowForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblShowForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseMove(lblShowForm, Button) = True Then
    lMenus.mDecodeMenuIndex = 11
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblEffect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub
