VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Menus (Hidden)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuQue 
      Caption         =   "Hidden (Que)"
      Begin VB.Menu mnuProporties2 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuSep93782963972 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuEncode 
         Caption         =   "Encode"
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Decode"
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Hidden (Files)"
      Begin VB.Menu mnuAddtoQue 
         Caption         =   "Add to Burn Que"
      End
      Begin VB.Menu mnuSaveAsPlaylist 
         Caption         =   "Create Playlist"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowFolder 
         Caption         =   "Containing Folder"
      End
      Begin VB.Menu mnuSep3970273 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "Encode"
         Begin VB.Menu mnuEncode2 
            Caption         =   "Wave -> Mp3"
         End
         Begin VB.Menu mnuEncodeWMA 
            Caption         =   "Wave -> Wma"
         End
      End
      Begin VB.Menu mnu3289083dh28hi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay2 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuDecode2 
         Caption         =   "Decode"
      End
      Begin VB.Menu mnuNormalize12 
         Caption         =   "Normalize"
      End
      Begin VB.Menu mnuConvertBitrate 
         Caption         =   "Convert Bitrate"
      End
      Begin VB.Menu mnuRip378 
         Caption         =   "Rip"
      End
      Begin VB.Menu mnuEffects34 
         Caption         =   "Effects"
      End
      Begin VB.Menu mnuSep93728762389 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep39720308927 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddtoPlaylist 
         Caption         =   "Add to Playlist"
      End
      Begin VB.Menu mnuRemove2 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuDelete2 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuSep7937 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTag 
         Caption         =   "View Tag"
      End
      Begin VB.Menu mnuProporties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Hidden (File)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Begin VB.Menu mnuSupportedTypes 
            Caption         =   "Supported Types"
         End
         Begin VB.Menu mnuSep389236789269 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAudioOnly 
            Caption         =   "Audio Only"
         End
         Begin VB.Menu mnuVideoOnly 
            Caption         =   "Video Only"
         End
         Begin VB.Menu mnuMp3Only 
            Caption         =   "MP3 Only"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSep89329786378926 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
         Begin VB.Menu mnuWaveToMp31 
            Caption         =   "Wave to MP3"
         End
         Begin VB.Menu mnuMP3ToWave 
            Caption         =   "MP3 to Wave"
         End
         Begin VB.Menu mnuSep3787289639263 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWaveToWMA1 
            Caption         =   "Wave to Windows Media Audio"
         End
         Begin VB.Menu mnuWMAtowave 
            Caption         =   "Windows Media Audio to Wave"
         End
      End
      Begin VB.Menu mnuEffects 
         Caption         =   "Effects"
         Begin VB.Menu mnuShowEffects 
            Caption         =   "Show Effects"
         End
         Begin VB.Menu mnuSep8378926297836 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNormalize 
            Caption         =   "Normalize"
         End
         Begin VB.Menu mnuSep382896392786 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAmplitude 
            Caption         =   "Amplitude"
         End
         Begin VB.Menu mnuChorus 
            Caption         =   "Chorus"
         End
         Begin VB.Menu mnuCFilter 
            Caption         =   "Click Filter"
         End
         Begin VB.Menu mnuDistortion 
            Caption         =   "Distortion"
         End
         Begin VB.Menu mnuEcho 
            Caption         =   "Echo"
         End
         Begin VB.Menu mnuFadeIN 
            Caption         =   "Fade In"
         End
         Begin VB.Menu mnuFadeOut 
            Caption         =   "Fade Out"
         End
         Begin VB.Menu mnuReverb 
            Caption         =   "Reverb"
         End
         Begin VB.Menu mnuShifting 
            Caption         =   "Shifting"
         End
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Playback"
         Begin VB.Menu mnuQuickPlay 
            Caption         =   "Quick Play"
         End
         Begin VB.Menu mnuSep832863826486 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlay1 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPause1 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuStop1 
            Caption         =   "Stop"
         End
      End
      Begin VB.Menu mnuSep78938926978236 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakeMP3Album 
         Caption         =   "Save As Album"
      End
      Begin VB.Menu mnuSep9378926389 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Hidden (Edit)"
      Begin VB.Menu mnuCut1 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy1 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste1 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSep38897269386 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRemove1 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuSep3892789469782 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Me.Icon = frmGraphics.Icon
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub mnuAddtoPlaylist_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
If Len(frmMain.tvwSources.SelectedItem.Text) <> 0 Then
    If FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text) = 0 Then
        msg2 = DecodeLocation(frmMain.tvwSources.SelectedItem.Text, frmMain.tvwSources.SelectedItem.FullPath) & "\" & frmMain.tvwFiles.SelectedItem.Text
        AddToFiles msg2, False
    Else
        MsgBox "File '" & frmMain.tvwFiles.SelectedItem.Text & "' exists in playlist", vbExclamation
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuAddtoPlaylist_Click()"
End Sub

Private Sub mnuAddtoQue_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
msg = frmMain.tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    frmMain.tvwToBurn.Nodes.Add , , , GetFileTitle(msg)
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuAddtoQue_Click()"
End Sub

Private Sub mnuAmplitude_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

AddAmplitude
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuAmplitude_Click()"
End Sub

Private Sub mnuAudioOnly_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptOpen True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuAudioOnly_Click()"
End Sub

Private Sub mnuDecode_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, i As Integer, f As Integer
msg = frmMain.tvwToBurn.SelectedItem.Text
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
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuDecode_Click()"
End Sub

Private Sub mnuDecode2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
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
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuDecode2_Click()"
End Sub

Private Sub mnuDelete2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, mbox As VbMsgBoxResult, f As Integer
msg = frmMain.tvwFiles.SelectedItem.Text
mbox = MsgBox("Are you sure you wish to delete the file '" & msg & "'?", vbYesNo + vbQuestion, "Delete Confirmation")
If mbox = vbYes Then
    f = frmMain.tvwFiles.SelectedItem.Index
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If i <> 0 And Len(lFiles.fFile(i).fFilename) <> 0 Then
            If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
                Kill lFiles.fFile(i).fFilename
                frmMain.tvwFiles.Nodes.Remove f
            End If
        End If
    End If
Else
    Exit Sub
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuDelete2_Click()"
End Sub

Private Sub mnuEdit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuVisible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuEdit_Click()"
End Sub

Private Sub mnuEffects34_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmEffects.Show
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuEffects34_Click()"
End Sub

Private Sub mnuEncode_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lext As String, msg As String, msg2 As String, i As Integer, msg3 As String
msg = frmMain.tvwToBurn.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        lext = Right(LCase(msg), 4)
        Select Case lext
        Case ".mp3"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        Case ".wav"
            msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
            msg3 = left(msg, Len(msg) - 4) & ".mp3"
            EncodeFile lFiles.fFile(i).fFilename, msg2 & msg3
        Case ".wma"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        End Select
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuEncode_Click()"
End Sub

Private Sub mnuEncode2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lext As String, msg As String, msg2 As String, i As Integer, msg3 As String
msg = frmMain.tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        lext = Right(LCase(msg), 4)
        Select Case lext
        Case ".mp3"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        Case ".wav"
            msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
            msg3 = left(msg, Len(msg) - 4) & ".mp3"
            EncodeFile lFiles.fFile(i).fFilename, msg2 & msg3
        Case ".wma"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        End Select
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuEncode2_Click()"
End Sub

Private Sub mnuEncodeWMA_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lext As String, msg As String, msg2 As String, i As Integer, msg3 As String
msg = frmMain.tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        lext = Right(LCase(msg), 4)
        Select Case lext
        Case ".mp3"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        Case ".rm"
            MsgBox "Audiogen does not support Real Media"
        Case ".ogg"
            MsgBox "Audiogen does not yet support OGG.", vbCritical
        Case ".wav"
            msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
            msg3 = left(msg, Len(msg) - 4) & ".wma"
            EncodeWMA lFiles.fFile(i).fFilename, msg2 & msg3
        Case ".wma"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Sub
        End Select
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuEncodeWMA_Click()"
End Sub

Private Sub mnuExit_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
EndProgram
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuExit_Click()"
End Sub

Private Sub mnuFile_Click()
lMenus.mFileMenuVisible = True
frmMain.imgFile.Visible = True
frmMain.imgFileOver.Visible = False
End Sub

Private Sub mnuMp3Only_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptOpen False, False, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuMp3Only_Click()"
End Sub

Private Sub mnuMP3ToWave_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
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
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuMP3ToWave_Click()"
End Sub

Private Sub mnuNormalize_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(frmMain.tvwFiles.SelectedItem.Text) <> 0 Then
    NormalizeFile lFiles.fFile(FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text)).fFilename
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuNormalize_Click()"
End Sub

Private Sub mnuNormalize12_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
QuickNormalize frmMain.tvwFiles.SelectedItem.Text
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuPlay_Click()"
End Sub

Private Sub mnuPause1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuPlay_Click()"
End Sub

Private Sub mnuPlay_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
QuickPlay lFiles.fFile(FindFileIndexByFilename(frmMain.tvwToBurn.SelectedItem.Text)).fFilename
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuPlay_Click()"
End Sub

Private Sub mnuPlay1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next

If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuPlay1_Click()"
End Sub

Private Sub mnuPlay2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
msg = frmMain.tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        If Len(lFiles.fFile(i).fFilename) <> 0 Then
            QuickPlay lFiles.fFile(i).fFilename
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuPlay2_Click()"
End Sub

Private Sub mnuProporties_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call ShowFileProperties(frmMain.hwnd, lFiles.fFile(FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text)).fFilename)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuProporties_Click()"
End Sub

Private Sub mnuProporties2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Call ShowFileProperties(frmMain.hwnd, lFiles.fFile(FindFileIndexByFilename(frmMain.tvwToBurn.SelectedItem.Text)).fFilename)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuProporties2_Click()"
End Sub

Private Sub mnuQuickPlay_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
QuickPlay lFiles.fFile(PromptOpen(True)).fFilename
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuQuickPlay_Click()"
End Sub

Private Sub mnuRemove_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, f As Integer
msg = frmMain.tvwToBurn.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 And Len(lFiles.fFile(i).fFilename) <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
        If RemovePlaylistEntry(i) = True Then
            f = frmMain.tvwToBurn.SelectedItem.Index
            If f <> 0 Then frmMain.tvwToBurn.Nodes.Remove f
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuRemove_Click()"
End Sub

Private Sub mnuRemove2_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, f As Integer
msg = frmMain.tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 And Len(lFiles.fFile(i).fFilename) <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
        If RemovePlaylistEntry(i) = True Then
            f = frmMain.tvwFiles.SelectedItem.Index
            If f <> 0 Then frmMain.tvwFiles.Nodes.Remove f
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuRemove2_Click()"
End Sub

Private Sub mnuRename_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg2 = frmMain.tvwFiles.SelectedItem.Text
msg = InputBox("Select new filename:", App.Title, msg2)
If Len(msg) = 0 Then Exit Sub
msg2 = GetFileTitle(msg2)
i = FindFileIndexByFilename(msg2)
If i <> 0 Then
    If DoesFileExist(lFiles.fFile(i).fFilename) = True Then
        lFiles.fFile(i).fFilename = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg2)) & msg
        frmMain.tvwFiles.SelectedItem.Text = msg
        
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuRename_Click()"
End Sub

Private Sub mnuRip378_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmRipCD.Show 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuShowEffects_Click()"
End Sub

Private Sub mnuShowEffects_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmEffects.Show
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuShowEffects_Click()"
End Sub

Private Sub mnuShowFolder_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If Len(frmMain.tvwFiles.SelectedItem.Text) <> 0 Then
    msg = lFiles.fFile(FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text)).fFilename
    msg2 = msg
    msg2 = GetFileTitle(msg2)
    msg = left(msg, Len(msg) - Len(msg2))
    DisplayDirectory msg
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuRemove2_Click()"
End Sub

Private Sub mnuSupportedTypes_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptOpen
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuSupportedTypes_Click()"
End Sub

Private Sub mnuVideoOnly_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PromptOpen False, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuVideoOnly_Click()"
End Sub

Private Sub mnuViewTag_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayTag
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuViewTag_Click()"
End Sub

Private Sub mnuWaveToMp31_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayTag
EncodeWaveToMp3FromTreeview frmMain.tvwFiles
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuWaveToMp31_Click()"
End Sub

Private Sub mnuWaveToWMA1_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
EncodeWaveToMp3FromTreeview frmMain.tvwFiles, True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub mnuWaveToWMA1_Click()"
End Sub

