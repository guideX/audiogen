Attribute VB_Name = "mdlMenus"
Option Explicit

Public Sub ResetFileMenuArrows()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
With frmFileMenu
    .imgOpenArrow.Picture = frmGraphics.imgMenuArrow1.Picture
    .imgConvertArrow.Picture = frmGraphics.imgMenuArrow1.Picture
    .imgDecodeArrow.Picture = frmGraphics.imgMenuArrow1.Picture
End With
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ResetFileMenuArrows()"
End Sub

Public Sub RefreshMenus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lMenus.mFileMenuVisible = True Then
    If lMenus.mFileMenuIndex <> 1 And frmFileMenu.lblRip.BackColor <> &HE0E0E0 Then frmFileMenu.lblRip.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 2 And frmFileMenu.lblOpen.BackColor <> &HE0E0E0 Then frmFileMenu.lblOpen.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 2 And frmFileMenu.imgOpen.Picture <> frmGraphics.imgOpen1.Picture Then frmFileMenu.imgOpen.Picture = frmGraphics.imgOpen1.Picture
    If lMenus.mFileMenuIndex <> 3 And frmFileMenu.lblDecode.BackColor <> &HE0E0E0 Then frmFileMenu.lblDecode.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 4 And frmFileMenu.lblBurn.BackColor <> &HE0E0E0 Then frmFileMenu.lblBurn.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 5 And frmFileMenu.lblNormalize.BackColor <> &HE0E0E0 Then frmFileMenu.lblNormalize.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 6 And frmFileMenu.lblSearch.BackColor <> &HE0E0E0 Then frmFileMenu.lblSearch.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 7 And frmFileMenu.lblPlay.BackColor <> &HE0E0E0 Then frmFileMenu.lblPlay.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 8 And frmFileMenu.lblPause.BackColor <> &HE0E0E0 Then frmFileMenu.lblPause.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 9 And frmFileMenu.lblStop.BackColor <> &HE0E0E0 Then frmFileMenu.lblStop.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 10 And frmFileMenu.lblRandom.BackColor <> &HE0E0E0 Then frmFileMenu.lblRandom.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 11 And frmFileMenu.lblMerge.BackColor <> &HE0E0E0 Then frmFileMenu.lblMerge.BackColor = &HE0E0E0
    If lMenus.mFileMenuIndex <> 100 And frmFileMenu.lblExit.BackColor <> &HE0E0E0 Then frmFileMenu.lblExit.BackColor = &HE0E0E0
End If
If lMenus.mEditMenuVisible = True Then
    'If lMenus.mEditMenuIndex <> 100 And frmFileMenu.lblExit.BackColor <> &HE0E0E0 Then frmFileMenu.lblExit.BackColor = &HE0E0E0
End If
If lMenus.mOpenMenuVisible = True Then
    If lMenus.mOpenMenuIndex <> 1 And frmOpenMenu.lblSupportedTypes.BackColor <> &HE0E0E0 Then frmOpenMenu.lblSupportedTypes.BackColor = &HE0E0E0
    If lMenus.mOpenMenuIndex <> 2 And frmOpenMenu.lblAudio.BackColor <> &HE0E0E0 Then frmOpenMenu.lblAudio.BackColor = &HE0E0E0
    If lMenus.mOpenMenuIndex <> 3 And frmOpenMenu.lblVideo.BackColor <> &HE0E0E0 Then frmOpenMenu.lblVideo.BackColor = &HE0E0E0
End If
If lMenus.mConvertMenuVisible = True Then
    If lMenus.mConvertMenuIndex <> 1 And frmMenuConvert.lblWaveToMP3.BackColor <> &HE0E0E0 Then frmMenuConvert.lblWaveToMP3.BackColor = &HE0E0E0
    If lMenus.mConvertMenuIndex <> 2 And frmMenuConvert.lblMP3ToWave.BackColor <> &HE0E0E0 Then frmMenuConvert.lblMP3ToWave.BackColor = &HE0E0E0
    If lMenus.mConvertMenuIndex <> 3 And frmMenuConvert.lblWaveToWMA.BackColor <> &HE0E0E0 Then frmMenuConvert.lblWaveToWMA.BackColor = &HE0E0E0
    If lMenus.mConvertMenuIndex <> 4 And frmMenuConvert.lblDecodeWMA.BackColor <> &HE0E0E0 Then frmMenuConvert.lblDecodeWMA.BackColor = &HE0E0E0
End If
If lMenus.mEffectsMenuVisible = True Then
    For i = 1 To 10
        If lMenus.mEffectsMenuIndex <> i And frmMenuEffects.lblEffect(i).BackColor <> &HE0E0E0 Then frmMenuEffects.lblEffect(i).BackColor = &HE0E0E0
    Next i
    If lMenus.mEffectsMenuIndex <> 11 And frmMenuEffects.lblShowForm.BackColor <> &HE0E0E0 Then frmMenuEffects.lblShowForm.BackColor = &HE0E0E0
    If lMenus.mEffectsMenuIndex <> 11 And frmMenuEffects.lblSaveAs.BackColor <> &HE0E0E0 Then frmMenuEffects.lblSaveAs.BackColor = &HE0E0E0
    If lMenus.mEffectsMenuIndex <> 11 And frmMenuEffects.lblPlay.BackColor <> &HE0E0E0 Then frmMenuEffects.lblPlay.BackColor = &HE0E0E0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub RefreshMenus()"
End Sub

Public Function LabelMouseUp(lLabel As Label, lButton As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 Then
    lLabel.BackColor = &H8000000F
    LabelMouseUp = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LabelMouseUp(lLabel As Label, lButton As Integer)"
End Function

Public Function LabelMouseDown(lLabel As Label, lButton As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 1 Then
    LabelMouseDown = True
    lLabel.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSupportedTypes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Function

Public Function LabelMouseMove(lLabel As Label, lButton As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 0 Then
    RefreshMenus
    If lLabel.BackColor <> &HC0C0C0 Then
        lLabel.BackColor = &HC0C0C0
        LabelMouseMove = True
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LabelMouseMove(lLabel As Label)"
End Function
