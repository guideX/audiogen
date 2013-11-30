VERSION 5.00
Begin VB.Form frmMenuConvert 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
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
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   112
      X2              =   112
      Y1              =   192
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   216
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1
      X2              =   1
      Y1              =   0
      Y2              =   208
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblMP3ToWave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Decode MP3"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   255
      Width           =   1815
   End
   Begin VB.Label lblWaveToMP3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Encode MP3"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1815
   End
   Begin VB.Label lblWaveToWMA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Encode WMA"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label lblDecodeWMA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "       Decode WMA"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   705
      Width           =   1695
   End
End
Attribute VB_Name = "frmMenuConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuVisible = True
lMenus.mUsingSubMenu = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuVisible = False
lMenus.mUsingSubMenu = False
AlwaysOnTop frmFileMenu, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblDecodeWMA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblDecodeWMA, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecodeWMA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDecodeWMA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuIndex = 2
If LabelMouseMove(lblDecodeWMA, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecodeWMA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDecodeWMA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblDecodeWMA, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDecodeWMA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMP3ToWave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblMP3ToWave, Button) = True Then
    Dim msg As String, msg2 As String, msg3 As String, i As Integer, f As Integer
    msg = frmMain.tvwFiles.SelectedItem.Text
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If i <> 0 Then
            If Len(lFiles.fFile(i).fFilename) <> 0 Then
                msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
                msg3 = left(msg, Len(msg) - 4) & ".wav"
                DecodeFile msg2, msg3, msg
            End If
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMP3ToWave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMP3ToWave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuIndex = 2
If LabelMouseMove(lblMP3ToWave, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMP3ToWave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblMP3ToWave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblMP3ToWave, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblMP3ToWave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblWaveToMP3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblWaveToMP3, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblWaveToMP3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblWaveToMP3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuIndex = 1
If LabelMouseMove(lblWaveToMP3, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblWaveToMP3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblWaveToMP3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseUp(lblWaveToMP3, Button) = True Then
    DisplayTag
    EncodeWaveToMp3FromTreeview frmMain.tvwFiles
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblWaveToMP3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblWaveToWMA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If LabelMouseDown(lblWaveToWMA, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblWaveToWMA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblWaveToWMA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mConvertMenuIndex = 3
If LabelMouseMove(lblWaveToWMA, Button) = True Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblWaveToWMA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub
