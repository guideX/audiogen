Attribute VB_Name = "mdlPlayback"
Option Explicit

Public Function PromptOpen(Optional lAudioOnly As Boolean, Optional lVideoOnly As Boolean, Optional lMP3Only As Boolean, Optional lWaveOnly As Boolean) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String, f As Integer
If lMP3Only = True Then
    msg = OpenDialog(frmMain, "Mpeg Layer 3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*", "Open Mpeg Layer 3 Files", CurDir)
    If Len(msg) <> 0 Then
        AddToFiles msg, False
        frmMain.ResetFileTreeView
        frmMain.tvwFiles.Nodes.Add 1, tvwChild, , GetFileTitle(msg), 5
    End If
    GoTo ErrorChk
End If
If lAudioOnly = True Then
    msg = OpenDialog(frmMain, "Supported Audio|" & lSettings.sSupportedAudio & "|", "Open (Audio Only) ...", CurDir)
    If Len(msg) <> 0 Then
        AddToFiles msg, False
        frmMain.ResetFileTreeView
        frmMain.tvwFiles.Nodes.Add 1, tvwChild, , GetFileTitle(msg), 5
    End If
    GoTo ErrorChk
End If
If lVideoOnly = True Then
    msg = OpenDialog(frmMain, "Supported Video|" & lSettings.sSupportedVideo & "|", "Open Play (Video Only) ...", CurDir)
    If Len(msg) <> 0 Then
        AddToFiles msg, False
        frmMain.ResetFileTreeView
        frmMain.tvwFiles.Nodes.Add 1, tvwChild, , GetFileTitle(msg), 5
    End If
    GoTo ErrorChk
End If
msg = OpenDialog(frmMain, "Supported Media|" & lSettings.sSupportedMedia & "|", "Open Play ...", CurDir)
If Len(msg) <> 0 Then
    AddToFiles msg, False
    frmMain.ResetFileTreeView
    frmMain.tvwFiles.Nodes.Add 1, tvwChild, , GetFileTitle(msg), 5
End If
ErrorChk:
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
        PromptOpen = i
        OpenContainingFolder msg2
        For f = 1 To frmMain.tvwFiles.Nodes.Count
            If Len(frmMain.tvwFiles.Nodes(f).Text) <> 0 Then
                If LCase(frmMain.tvwFiles.Nodes(f).Text) = LCase(msg) Then
                    frmMain.tvwFiles.SetFocus
                    frmMain.tvwFiles.Nodes(f).Selected = True
                    Exit For
                End If
            End If
        Next f
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub PromptOpen()"
End Function

Public Sub StartPlayback(lFile As String, lIntro As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, f As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(startplayback).txt")
If Len(lFile) <> 0 Then
    i = FindFileIndex(lFile)
    If i = 0 Then
        i = AddToFiles(lFile, False, 0)
    End If
    lFiles.fIndex = i
    lPlayer.pStatus = sPlay
    lPlayer.pFilename = lFile
    lPlayer.pFileType = fAllFileTypes
    frmMain.ResizeMain
    DisplayMediaItem lFile, frmMain.tvwFiles
    msg = PlayMultimedia(frmMain.lblAliasname.Caption, 0, GetTotalframes(frmMain.lblAliasname.Caption))
    If Right(LCase(lFile), 4) = ".mpg" Or Right(LCase(lFile), 4) = ".wmv" Or Right(LCase(lFile), 5) = ".mpeg" Or Right(LCase(lFile), 4) = ".avi" Or Right(LCase(lFile), 4) = ".mov" And frmMain.fraVideo.Visible = False Then frmMain.SwitchView
    frmMain.lblStatus = msg
    'If frmMain.imgBurn.Picture <> frmGraphics.imgAbort1.Picture Then frmMain.imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture
    'frmMain.imgCdCopy.Picture = frmGraphics.imgCDCopyDisabled.Picture
    If msg <> "Success" Then
        MsgBox "There was an error starting playback" & vbCrLf & msg, vbExclamation
    Else
        SetPlaybackObjects lFile
        frmMain.tmrPosition.Enabled = True
        If lIntro = True Then
            PutMultimedia frmVideo.fraVideo.hwnd, frmMain.lblAliasname.Caption, 0, 0, frmVideo.ScaleX(Val(frmVideo.fraVideo.Width), 1, 3), frmVideo.ScaleX(Val(frmVideo.fraVideo.Height), 1, 3)
            frmVideo.fraVideo.Refresh
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub StartPlayback(lFile As String)"
End Sub

Public Sub GoBackward()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(gobackward).txt")
Dim i As Integer
If lFiles.fCount = 0 Then Exit Sub
If lFiles.fCount = 1 Then
    i = 1
    GoTo StartPlayback2
End If
i = lFiles.fIndex
If i = 0 Then
    i = lFiles.fCount
CheckIndex:
    If Len(lFiles.fFile(i).fFilename) = 0 Then
        i = i - 1
        GoTo CheckIndex
    End If
Else
    i = lFiles.fIndex - 1
End If
StartPlayback2:
QuickPlay lFiles.fFile(i).fFilename
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub GoBackward()"
End Sub

Public Sub GoForward()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(goforward).txt")
If lFiles.fCount = 0 Then Exit Sub
If lFiles.fCount = 1 Then
    i = 1
    GoTo StartPlayback2
End If
i = lFiles.fIndex
If i = 0 Then
    i = lFiles.fCount
CheckIndex:
    If Len(lFiles.fFile(i).fFilename) = 0 Then
        i = i + 1
        GoTo CheckIndex
    End If
Else
    i = lFiles.fIndex + 1
End If
StartPlayback2:
AdjustStatus sSelectFile, lFiles.fFile(i).fFilename
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub GoForward()"
End Sub

Public Function QuickPlay(lFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(quickplay).txt")
If Len(lFile) = 0 Then
    'lFile = OpenDialog(frmMain, "Supported Media|*.avi;*.qt;*.mov;*.dat;*.snd;*.mpg;*.mpa;*.mpv;*.enc;*.m1v;*.mp2;*.mp3;*.mpe;*.mpeg;*.mpm;*.au;*.snd;*.aif;*.aiff;*.aifc;*.wav;*.wmv;*.wma;*.avi|", "Quick Play ...", CurDir)
    lFile = OpenDialog(frmMain, "Supported Media |" & lSettings.sSupportedMedia & "|", "Quick Play ...", CurDir)
    If Len(lFile) = 0 Then Exit Function
End If
If Right(LCase(lFile), 4) = ".cda" Then
    lPlayer.pFileType = fCDAudio
    AdjustStatus sSelectFile, lFile
Else
    If DoesFileExist(lFile) = True Then
        msg = lFile
        msg = GetFileTitle(msg)
        lEvents.eCurrentFile = msg
        lPlayer.pFileType = fAllFileTypes
        AdjustStatus sSelectFile, lFile, ""
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function QuickPlay(lFile As String) As String"
End Function

Public Sub InitPlayback()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(initplayback).txt")
If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
    SetDefaultDevice "MPEGVideo", "mciqtz.drv"
End If
If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
    SetDefaultDevice "sequencer", "mciseq.drv"
End If
If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
    SetDefaultDevice "avivideo", "mciavi.drv"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub InitPlayback()"
End Sub

Public Sub StopPlayback()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(stopplayback).txt")
Select Case lPlayer.pFileType
Case fCDAudio
    frmMain.ctlRipper.StopPlaying
    frmMain.imgProgress.left = 2000
    frmMain.imgProgressYellow.Width = 1
    frmMain.imgProgress.Picture = frmGraphics.imgSlider3.Picture
    
    frmMain.imgPlay.Picture = frmGraphics.imgPlay1.Picture
    frmMain.imgStop.Picture = frmGraphics.imgStop4.Picture
    AdjustStatus sIdle
Case fAllFileTypes
    StopMultimedia frmMain.lblAliasname.Caption
    CloseMultimedia frmMain.lblAliasname.Caption
    frmMain.fraVideo.Refresh
    frmMain.imgProgress.left = 2000
    frmMain.imgProgressYellow.Width = 1
    frmMain.imgProgress.Picture = frmGraphics.imgSlider3.Picture
    If frmMain.fraVideo.Visible = True Then frmMain.SwitchView
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub StopPlayback()"
End Sub

Public Sub SetPlaybackObjects(lFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(setplaybackobjects).txt")
If Len(lFile) <> 0 Then
    lPlayer.pFilename = lFile
    msg = lFile
    msg = GetFileTitle(msg)
    If Len(msg) <> 0 Then
        frmMain.lblFormat.Caption = Right(LCase(msg), 3)
        frmMain.imgPlay.Picture = frmGraphics.imgPlay4.Picture
        frmMain.lblStatus.Caption = "Playing " & LCase(msg)
        frmMain.imgStop.Picture = frmGraphics.imgStop1.Picture
        frmMain.imgProgress.Picture = frmGraphics.imgSlider1.Picture
        lPlayer.pStatus = sPlay
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetPlaybackObjects(lFile As String)"
End Sub

