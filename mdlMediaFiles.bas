Attribute VB_Name = "mdlMediaFiles"
Option Explicit

Public Sub DisplayMediaItem(lFullFile As String, lTreeView As TreeView)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(displaymediaitem).txt")
If lPlayer.pStatus <> sPlay Then
    frmMain.lblFormat.Caption = ""
    frmMain.lblBitrate.Caption = ""
    frmMain.lblTimeDisplay.Caption = ""
    frmMain.lblTitle.Caption = ""
    frmMain.lblKHZ.Caption = ""
    CheckMenus
    If Len(lFullFile) <> 0 Then
        msg = lFullFile
        msg = GetFileTitle(msg)
        i = FindFileIndexByFilename(msg)
        frmMain.lblFormat.Caption = Right(LCase(lFiles.fFile(i).fFilename), 3)
        lTag.tFile = lFiles.fFile(i).fFilename
        SetAddress left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
        If Right(LCase(lFiles.fFile(i).fFilename), 4) = ".mp3" Then
            GetTagInfo
            GetMP3Info
            DoEvents
            frmMain.lblKHZ.Caption = lTag.tFreqChan
            frmMain.lblBitrate.Caption = lTag.tBitrate & " kbps"
            If Len(lTag.tLength) <> 0 Then
                If Len(Trim(lTag.tTitle)) <> 0 Then
                    frmMain.lblTitle.Caption = lTag.tTitle
                    frmMain.lblTimeDisplay.Caption = Format(lTag.tLength / 0.6, "##:##")
                Else
                    frmMain.lblTitle.Caption = lTreeView
                    frmMain.lblTimeDisplay.Caption = Format(lTag.tLength / 0.6, "##:##")
                End If
            End If
        End If
        If Len(frmMain.lblTitle.Caption) = 0 Then frmMain.lblTitle.Caption = left(msg, Len(msg) - 4)
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DisplayMediaItem(lFullFile As String)"
End Sub

Public Function RipTrack(lTrack As Integer, lFilename As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult, msg As String, b As Boolean, f As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(riptrack).txt")
If lTrack <> 0 Then
    msg = lFilename
    msg = GetFileTitle(msg)
    If DoesFileExist(lFilename) = True Then
        mbox = MsgBox("The file '" & msg & "' already exists. Delete file and continue?", vbYesNo + vbQuestion, "Overwrite Confirmation")
        If mbox = vbYes Then
            Kill lFilename
        End If
    End If
    frmMain.lblTransferStatus.Caption = "Copy"
    frmMain.lblTrackNumber.Caption = "Track " & Str(lTrack)
    f = frmMain.ctlRipper.ReadTrack(lTrack, lFilename)
    If f <> 0 Then
        ProcessCDDriveError f
    Else
        lEvents.eCurrentFile = msg
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function RipTrack(lTrack As Integer, lFilename As String) As String"
End Function

Public Function StartRip() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, f As Integer, msg As String, msg2 As String, msg3 As String, l As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(startrip).txt")
If lEvents.eProcessing = True Then Exit Function
frmRipCD.tvwToRip.Nodes.Clear
For l = 1 To frmMain.ctlRipper.TrackCount
    frmRipCD.tvwToRip.Nodes.Add , , , "Track " & l & ".wav"
Next l
frmRipCD.Show 1
SelectCDDriveByCombo
For i = 1 To lBurnQue.bCount + 1
    If Len(lBurnQue.bFiles(i)) <> 0 Then
        msg = lBurnQue.bFiles(i)
        msg2 = msg
        msg2 = GetFileTitle(msg2)
        If Len(msg) <> 0 And Len(msg2) <> 0 Then
            AddEvent eEncode, msg, left(msg, Len(msg) - 4) & ".mp3"
            AddEvent eRip, Str(i), msg
            frmMain.lblStatus.Caption = "Ripping CD-Audio"
        End If
    End If
Next i
frmMain.tmrCheckEvents.Enabled = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function StartRip() As String"
End Function

Public Function EncodeWMA(lInputFile As String, lOutputFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, mbox As VbMsgBoxResult, b As Long, i As Long
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(encodewma).txt")
If Len(lInputFile) <> 0 And Len(lOutputFile) <> 0 Then
    If DoesFileExist(lInputFile) = True Then
        If DoesFileExist(lOutputFile) = True Then
            mbox = MsgBox("File exists, would you like to delete the existing file first?", vbYesNo + vbQuestion, "Overwrite confirmation")
            If mbox = vbYes Then
                Kill lOutputFile
            ElseIf mbox = vbNo Then
                MsgBox "Encode canceled!", vbExclamation
                Exit Function
            End If
        End If
        msg = lInputFile
        msg = GetFileTitle(msg)
        msg2 = lOutputFile
        msg2 = GetFileTitle(msg2)
        b = 320000
        i = frmMain.cboBitrate.ListIndex
        i = i * 8000
        b = b - i
        With frmMain
            .prgSpaceLeft.Max = 100
            .wmaEnc.bitrate = b
            .wmaEnc.Encode lInputFile, lOutputFile
            frmMain.lblStatus.Caption = "Encoding file " & msg
            frmMain.lblTimeLeft.Caption = "Encode File:"
            lEvents.eCurrentFile = lOutputFile
        End With
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function EncodeWMA(lInputFile As String, lOutputFile As String) As String"
End Function

Public Function EncodeFile(lInputFile As String, lOutputFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, mbox As VbMsgBoxResult, b As Long, i As Long
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(encodefile).txt")
If Len(lInputFile) <> 0 And Len(lOutputFile) <> 0 Then
    If DoesFileExist(lInputFile) = True Then
        If DoesFileExist(lOutputFile) = True Then
            mbox = MsgBox("File exists, would you like to delete the existing file first?", vbYesNo + vbQuestion, "Overwrite confirmation")
            If mbox = vbYes Then
                Kill lOutputFile
            ElseIf mbox = vbNo Then
                MsgBox "Encode canceled!", vbExclamation
                Exit Function
            End If
        End If
        msg = lInputFile
        msg = GetFileTitle(msg)
        msg2 = lOutputFile
        msg2 = GetFileTitle(msg2)
        b = 320000
        i = frmMain.cboBitrate.ListIndex
        i = i * 8000
        b = b - i
        frmMain.prgSpaceLeft.Max = 100
        frmMain.ctlMP3Enc.bitrate = b
        frmMain.ctlMP3Enc.channels = 0
        frmMain.ctlMP3Enc.OPENFILENAME = lInputFile
        frmMain.ctlMP3Enc.savefilename = lOutputFile
        frmMain.lblStatus.Caption = "Encoding file " & msg
        frmMain.lblTimeLeft.Caption = "Encode File:"
        frmMain.ctlMP3Enc.Encode
        lEvents.eCurrentFile = lOutputFile
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function EncodeFile(lInputFile As String, lOutputFile As String) As String"
End Function

Public Sub QuickNormalize(lFilename As String, Optional lPath As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If Len(lFilename) <> 0 Then
    If Len(lPath) <> 0 Then
        If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
        i = FindFileIndex(lPath & lFilename)
        If i = 0 Then
            MsgBox "File not found", vbExclamation
            Exit Sub
        End If
    Else
        i = FindFileIndexByFilename(lFilename)
        If i = 0 Then
            MsgBox "File not found", vbExclamation
            Exit Sub
        End If
    End If
    If i <> 0 Then
        msg = lFiles.fFile(i).fFilename
        msg = GetFileTitle(msg)
        lPath = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
        msg = left(msg, Len(msg) - 4) & " (Normalized).wav"
        NormalizeFile lFiles.fFile(i).fFilename, lPath & msg
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub QuickNormalize(lFilename As String, Optional lPath As String)"
End Sub

Public Function NormalizeFile(lFile As String, Optional lOutputFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String, msg3 As String, msg4 As String, mbox As VbMsgBoxResult
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(normalizefile).txt")
If Len(lFile) <> 0 Then
    If DoesFileExist(lFile) = True Then
        msg = lFile
        msg = GetFileTitle(msg)
        msg2 = left(lFile, Len(lFile) - Len(msg))
        msg3 = "temp " & GetRnd(1000) & ".wav"
        msg4 = left(msg, Len(msg) - 4) & " (Normalized).wav"
        If Len(msg) <> 0 And Len(msg2) <> 0 And Len(msg3) <> 0 Then
            Select Case Right(LCase(msg), 4)
            Case ".mp3"
                DecodeFile msg2, msg3, msg: DoEvents
                msg = msg3
            Case ".wma"
                DecodeFile msg2, msg3, msg: DoEvents
                msg = msg3
            End Select
        End If
        DoEvents
        Pause 0.2
        If DoesFileExist(msg2 & msg) = True Then
            If Len(lOutputFile) <> 0 Then
                Dim fg As String
                If DoesFileExist(lOutputFile) = True Then
                    mbox = MsgBox("The file '" & lOutputFile & "' already exists. Would you like to delete that file and continue?", vbYesNo + vbQuestion, "Overwrite confirmation")
                    If mbox = vbYes Then
                        Kill lOutputFile
                        DoEvents
                    ElseIf mbox = vbNo Then
                        lEvents.eProcessing = False
                        NormalizeFile = lOutputFile
                        Exit Function
                    End If
                End If
                fg = lOutputFile
                fg = GetFileTitle(fg)
                lEvents.eCurrentFile = fg
                lBurnQue.bNormOutFile = lOutputFile
            Else
                If DoesFileExist(msg2 & msg) = True Then
                    mbox = MsgBox("The file '" & msg2 & msg & "' already exists. Would you like to delete that file and continue?", vbYesNo + vbQuestion, "Overwrite confirmation")
                    If mbox = vbYes Then
                        Kill msg2 & msg
                        DoEvents
                    ElseIf mbox = vbNo Then
                        lEvents.eProcessing = False
                        NormalizeFile = msg2 & msg
                        Exit Function
                    End If
                End If
                lEvents.eCurrentFile = msg
                lBurnQue.bNormOutFile = msg2 & msg
            End If
            lBurnQue.bNormInFile = lFile
            If Len(lBurnQue.bNormInFile) <> 0 And Len(lBurnQue.bNormOutFile) <> 0 Then
                DoEvents
                Pause 0.2
                frmMain.tmrCheckNormalize.Enabled = True
            Else
                MsgBox "Normalize on the file '" & msg & "' has been canceled because of an error", vbExclamation
            End If
            NormalizeFile = lBurnQue.bNormOutFile
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function NormalizeFile(lFile As String, Optional lOutputFile As String) As String"
End Function

Public Function MergeFiles() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, X As Long, SavedSpot As Long, theByte() As Byte, Length As Long, f As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(mergefiles).txt")
SavedSpot = 1
If lBurnQue.bCount = 0 Then Exit Function
lBurnQue.bMergeFilename = SaveDialog(frmMain, "Mpeg Layer 3 (*.mp3)|*.mp3|Windows Media Audio (*.wma)|*.wma|Wave Audio Files (*.wav)|*.wav|", "Save album as...", CurDir)
lBurnQue.bMergeFilename = left(lBurnQue.bMergeFilename, Len(lBurnQue.bMergeFilename) - 1)
frmMain.lblStatus.Caption = "Making Album"
Select Case Int(frmMain.Tag)
Case 1
    lBurnQue.bMergeFilename = lBurnQue.bMergeFilename & ".mp3"
Case 2
    lBurnQue.bMergeFilename = lBurnQue.bMergeFilename & ".wma"
Case 3
    lBurnQue.bMergeFilename = lBurnQue.bMergeFilename & ".wav"
End Select
frmMain.lblTimeLeft.Caption = "Make Album:"
frmMain.lblStatus.Caption = "Making album"
For i = 1 To lBurnQue.bCount
    frmMain.lblStatus.Caption = "Opened File(s) " & i & " of " & lBurnQue.bCount
    Pause 1
    If DoesFileExist(lBurnQue.bFiles(i)) = False Then
        MsgBox "File not found, merge aborted!", vbExclamation
        Exit For
    End If
    Length = FileLen(lBurnQue.bFiles(i))
    ReDim theByte(Length - 1)
    Open lBurnQue.bFiles(i) For Binary Access Read As #1
        Get #1, , theByte()
    Close #1
    Open lBurnQue.bMergeFilename For Binary As #1
        Put #1, SavedSpot, theByte()
    Close #1
    f = Int((100 / lBurnQue.bCount) * i + 1)
    If f < 100 And f > -1 Then frmMain.prgSpaceLeft.value = Int((100 / lBurnQue.bCount) * i + 1)
    SavedSpot = SavedSpot + Length
    DoEvents
Next i
frmMain.prgSpaceLeft.value = 0
AddToFiles lBurnQue.bMergeFilename, False
frmMain.ResetFileTreeView
frmMain.tvwFiles.Nodes.Add 1, tvwChild, , GetFileTitle(lBurnQue.bMergeFilename), 9
lBurnQue.bMergeFilename = ""
frmMain.lblStatus.Caption = "Album Complete"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function MergeFiles() As String"
End Function

Public Function InitCDBurner() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, e As Integer, msg As String, j As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(initcdburner).txt")
e = frmMain.ctlBurn.Init
If e = 0 Then
    For i = 1 To frmMain.ctlBurn.GetBurnerCount
        frmMain.ctlBurn.SelBurnerByIndex i - 1
        If frmMain.ctlBurn.DetermineDriver > 0 And frmMain.ctlBurn.DetermineDriver < 100 Then
            frmMain.ctlBurn.SelectDriver frmMain.ctlBurn.DetermineDriver
            i = 100
        End If
    Next i
    If frmMain.imgAutoEject.Picture = frmGraphics.imgAutoEject2.Picture Then
        frmMain.ctlBurn.EjectAfterWrite = True
    Else
        frmMain.ctlBurn.EjectAfterWrite = False
    End If
    
    frmMain.ctlBurn.SetProcessPriority 3
    If lSettings.sTestMode = True Then
        frmMain.ctlBurn.TestMode = True
    Else
        frmMain.ctlBurn.TestMode = False
    End If
    If lSettings.sFinalize = True Then
        frmMain.ctlBurn.Finalize = True
    Else
        frmMain.ctlBurn.Finalize = False
    End If
    frmMain.ctlBurn.BurnSpeed = 16
    frmMain.ctlBurn.FifoBufferCount = 50
    frmMain.ctlBurn.SectorsPerBuffer = 17
Else
    ProcessCDDriveError e
    'MsgBox "Audiogen was unable to initialize your CD Burner." & vbCrLf & "Error number: " & e, vbCritical
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function InitCDBurner() As Boolean"
End Function

Public Function RemovePlaylistEntry(lIndex As Integer) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(removeplaylistentry).txt")
If lIndex <> 0 Then
    lFiles.fFile(lIndex).fFilename = ""
    WriteINI lIniFiles.iPlaylists, Str(lIndex), "", ""
    RemovePlaylistEntry = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function RemovePlaylistEntry(lIndex As Long) As Boolean"
End Function

Public Sub SpaceLeft(lFilename As String, lSubtract As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim j As Integer, m As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(spaceleft).txt")
If LCase(Trim(lTag.tFile)) <> LCase(Trim(lFilename)) Then
    lTag.tFile = lFilename
    ClearInputs
    GetMP3Info
    DoEvents
End If
Pause 0.2
If Len(lTag.tLength) <> 0 And lTag.tLength <> "0" Then
    If lSubtract = True Then
        lSettings.sTimeSelected = lSettings.sTimeSelected + lTag.tLength
    Else
        lSettings.sTimeSelected = lSettings.sTimeSelected - lTag.tLength
    End If
    j = lSettings.sTimeSelected / 0.6
    m = 7400 - j
    frmMain.prgSpaceLeft.Max = 7400
    frmMain.prgSpaceLeft.value = m
    frmMain.lblTimeLeft.Caption = "Time Left:"
    frmMain.lblSpaceLeft.Caption = Format(m, "##:##") & " Minutes left"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SpaceLeft(lFilename As String, lSubtract As Boolean)"
End Sub

Public Function AddToBurnQue(lFile As String, lPath As String, Optional lNode As node) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, f As Integer, s As Integer, msg As String, j As Integer, m As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(addtoburnque).txt")
If Len(lFile) <> 0 And Len(lPath) <> 0 Then
    If DoesFileExist(lPath & lFile) = True Then
        For i = 1 To lBurnQue.bCount
            If Len(lBurnQue.bFiles(i)) = 0 Then
                f = i
                Exit For
            End If
        Next i
        SpaceLeft lPath & lFile, True
        If f = 0 Then
            If Right(LCase(lFile), 4) = ".mp3" Then
                If left(lTag.tFreqChan, 5) = "44100" Then
                    lBurnQue.bCount = lBurnQue.bCount + 1
                    lBurnQue.bFiles(lBurnQue.bCount) = lPath & lFile
                    AddToBurnQue = lBurnQue.bCount
                Else
                    Dim mbox As VbMsgBoxResult
                    If lSettings.sConvertKHZ = True Then
                        ChangeKHZ lPath & lFile
                        If DoesFileExist(lPath & lFile) = True Then
                            ClearInputs
                            lTag.tFile = lPath & lFile
                            GetTagInfo
                            GetMP3Info
                            DoEvents
                            Pause 0.2
                            If left(lTag.tFreqChan, 5) = "44100" Then
                                AddToFiles lPath & lFile, False
                                lBurnQue.bCount = lBurnQue.bCount + 1
                                lBurnQue.bFiles(lBurnQue.bCount) = lPath & lFile
                                AddToBurnQue = lBurnQue.bCount
                                Exit Function
                            Else
                                MsgBox "Convert KHZ Error occured", vbCritical
                            End If
                        Else
                            MsgBox "File Doesn't Exist"
                        End If
                    Else
                        frmMain.tvwToBurn.Nodes.Remove frmMain.tvwToBurn.SelectedItem.Index
                        mbox = MsgBox("The file '" & lFile & "' could not be added to the burn que because it is not 44100 KHZ, Audiogen will continue using the other files.", vbQuestion + vbYesNo, "Error Adding File")
                        If mbox = vbNo Then
                            ClearBurnQue
                            Exit Function
                        End If
                    End If
                End If
            ElseIf Right(LCase(lFile), 4) = ".wav" Then
                If LCase(GetWaveKHZ(lPath & lFile)) = "44 khz" Then
                    lBurnQue.bCount = lBurnQue.bCount + 1
                    lBurnQue.bFiles(lBurnQue.bCount) = lPath & lFile
                    AddToBurnQue = lBurnQue.bCount
                Else
                    frmMain.tvwToBurn.Nodes.Remove frmMain.tvwToBurn.SelectedItem.Index
                    MsgBox "The file '" & lFile & "' could not be added to the burn que because it is not 44100 KHZ.", vbExclamation, "Error Adding File"
                End If
            End If
        Else
            lBurnQue.bFiles(f) = lPath & lFile
            AddToBurnQue = f
        End If
    Else
        If Right(LCase(lFile), 4) = ".cda" Then
            msg = left(lFile, Len(lFile) - 4) & ".wav"
            s = FindBurnQueIndexByFilename(msg)
            If s <> 0 Then
                AddToBurnQue = s
                Exit Function
            End If
            For i = 1 To lBurnQue.bCount
                If Len(lBurnQue.bFiles(i)) = 0 Then
                    f = i
                    Exit For
                End If
            Next i
            If f = 0 Then
                lBurnQue.bCount = lBurnQue.bCount + 1
                lBurnQue.bFiles(f) = lPath & msg
                AddToBurnQue = lBurnQue.bCount
            Else
                lBurnQue.bFiles(f) = lPath & msg
                AddToBurnQue = f
            End If
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function AddToBurnQue(lFile As String, lPath As String) As Integer"
End Function

Public Sub ClearBurnQue()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(addtoburnque).txt")
For i = 0 To 99
    lBurnQue.bFiles(i) = ""
Next i
frmMain.prgSpaceLeft.value = 0
frmMain.lblTimeLeft.Caption = ""
frmMain.lblSpaceLeft.Caption = ""
lBurnQue.bCount = 0
lBurnQue.bMergeFilename = ""
lBurnQue.bNormInFile = ""
lBurnQue.bNormOutFile = ""
lBurnQue.bTrackIndex = 0
lSettings.sTimeSelected = 0
frmMain.lblStatus.Caption = "Burn que cleared"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ClearBurnQue()"
End Sub

Public Sub DeleteBurnQueEntry(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim j As Integer, m As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(deleteburnqueentry).txt")
SpaceLeft lBurnQue.bFiles(lIndex), False
lBurnQue.bFiles(lIndex) = ""
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DeleteBurnQueEntry(lIndex As Integer)"
End Sub

Public Function FindBurnQueIndexByFilename(lFile As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(findburnqueindexbyfilename).txt")
For i = 0 To 99
    msg = lBurnQue.bFiles(i)
    If Len(msg) <> 0 Then
        msg = GetFileTitle(msg)
        If LCase(Trim(lFile)) = LCase(Trim(msg)) Then
            FindBurnQueIndexByFilename = i
            Exit For
        End If
    End If
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindBurnQueIndexByFilename(lFile As String) As Integer"
End Function

Public Function InitRipper() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim f As Integer, msg As VbMsgBoxResult
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(initripper).txt")
f = frmMain.ctlRipper.Init
If f <> 0 Then
    ProcessCDDriveError f
    If frmMain.ctlRipper.IsAspiLoaded = False Then
        msg = MsgBox("Your ASPI layer is not ready, would you like to install an ASPI layer now?", vbYesNo + vbQuestion, "Audiogen")
        If msg = vbYes Then
            Shell App.Path & "\external\aspiupd.exe", vbNormalFocus
            End
        ElseIf msg = vbNo Then
            Exit Function
        End If
    Else
        MsgBox "An error occured loading your ripper. You may have another program open reading from your ripper. Quit all other programs that may be accessing your ripper and try again", vbCritical
    End If
Else
    InitRipper = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub InitRipper()"
End Function

Public Sub StartBurn()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String, msg3 As String, msg4 As String, khz As String, mbox As VbMsgBoxResult
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(startburn).txt")
frmMain.prgSpaceLeft.value = 0
frmMain.prgSpaceLeft.Max = 100
frmMain.prgPercentDone.value = 0
If lBurnQue.bCount <> 0 Then
    frmMain.imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture
    AddEvent eStartBurn, "", ""
    frmMain.tmrCheckEvents.Enabled = True
    For i = 0 To 99
        If Len(lBurnQue.bFiles(i)) <> 0 Then
            If DoesFileExist(lBurnQue.bFiles(i)) = True Then
                msg = lBurnQue.bFiles(i)
                msg2 = msg
                msg = GetFileTitle(msg)
                msg2 = left(msg2, Len(msg2) - Len(msg))
                msg3 = left(msg, Len(msg) - 4) & ".wav"
                msg4 = left(msg, Len(msg) - 4) & " (Normalized).wav"
                Select Case Right(LCase(lBurnQue.bFiles(i)), 3)
                Case "mp3"
                    lTag.tFile = msg2 & msg3
                    GetMP3Info
                    If left(lTag.tFreqChan, 5) = 22050 Then
                        mbox = MsgBox("The file '" & lBurnQue.bFiles(i) & "' is not 44 KHZ, and will not be burned. Would you like to contine with the burn process?", vbYesNo + vbQuestion, "Continue with Burn?")
                        If mbox = vbNo Then
                            ClearBurnQue
                            frmMain.tmrCheckEvents.Enabled = False
                            MsgBox "Burn was canceled! Your disc will be useable for another burn!", vbInformation
                            frmMain.ctlBurn.Reset
                            ClearBurnQue
                            InitCDBurner
                            AdjustStatus sIdle
                            frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture
                            frmMain.tvwToBurn.Nodes.Clear
                            Exit Sub
                        Else
                            DeleteBurnQueEntry i
                        End If
                    End If
                    lBurnQue.bFiles(i) = msg2 & msg3
                    'If frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks2.Picture Then AddEvent eNormalize, msg2 & msg3, msg2 & msg4
                    If lSettings.sNormalize = True Then AddEvent eNormalize, msg2 & msg3, msg2 & msg4
                    If DoesFileExist(msg2 & msg3) = True And FileSystem.FileLen(msg2 & msg3) = 0 Then
                        Kill msg2 & msg3
                        DoEvents
                    End If
                    If DoesFileExist(msg2 & msg3) = False Then
                        AddEvent eDecode, msg2 & msg, msg2 & msg3
                    End If
                Case "wma"
                    lBurnQue.bFiles(i) = msg2 & msg3
                    If frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks2.Picture Then AddEvent eNormalize, msg2 & msg3, msg2 & msg4
                    If DoesFileExist(msg2 & msg3) = False Then
                        AddEvent eDecode, msg2 & msg, msg2 & msg3
                    End If
                Case "wav"
                    If LCase(GetWaveKHZ(msg2 & msg)) = "44 khz" Then
                        lBurnQue.bFiles(i) = msg2 & msg
                    Else
                        mbox = MsgBox("The file '" & lBurnQue.bFiles(i) & "' is not 44 KHZ, and will not be burned. Would you like to contine with the burn process?", vbYesNo + vbQuestion, "Continue with Burn?")
                        If mbox = vbNo Then
                            ClearBurnQue
                            frmMain.tmrCheckEvents.Enabled = False
                            MsgBox "Burn was canceled! Your disc will be useable for another burn!", vbInformation
                            frmMain.ctlBurn.Reset
                            ClearBurnQue
                            InitCDBurner
                            AdjustStatus sIdle
                            frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture
                            frmMain.tvwToBurn.Nodes.Clear
                            Exit Sub
                        End If
                    End If
                End Select
            End If
        End If
    Next i
    DoEvents
    frmMain.tmrCheckEvents.Enabled = True
    frmMain.lblTransferStatus.Caption = "Transfer status:"
    PlayWav App.Path & "\audio\a_burncomencing.wav", SND_ASYNC
Else
    MsgBox "Unable to start the burn process. No files in que. Please select some files and try again", vbCritical, "Empty Que"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub StartBurn()"
End Sub

Public Function DecodeFile(lPath As String, lOutputFile As String, lInputFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult, i As Integer, d As Integer, msg As String, msg2 As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(decodefile).txt")
If Len(lPath) <> 0 And Len(lInputFile) <> 0 And Len(lOutputFile) <> 0 Then
    If DoesFileExist(lPath & lInputFile) = True Then
        If DoesFileExist(lPath & lOutputFile) = True Then
            mbox = MsgBox("The file '" & lOutputFile & "' exists, overwrite?", vbQuestion + vbYesNo, "Overwrite?")
            If mbox = vbYes Then
                Kill lPath & lOutputFile
            Else
                DecodeFile = lPath & lOutputFile
                Exit Function
            End If
        End If
        ChDir lPath
        msg = lPath & lInputFile
        msg2 = lPath & lOutputFile
        frmMain.lblTimeLeft.Caption = "Decode File:"
        frmMain.lblStatus.Caption = "Decoding " & lInputFile
        frmMain.prgSpaceLeft.Max = 100
        Select Case Right(LCase(msg), 4)
        Case ".mp3"
            frmMain.ctlMP3Decode.OPENFILENAME = msg
            frmMain.ctlMP3Decode.savefilename = msg2
            frmMain.ctlMP3Decode.Decode
            DoEvents
            lEvents.eCurrentFile = msg2
        Case ".wav"
        Case ".wma"
            frmMain.ctlWMADecode.OPENFILENAME = msg
            frmMain.ctlWMADecode.savefilename = msg2
            frmMain.ctlWMADecode.Decode
            DoEvents
            lEvents.eCurrentFile = msg2
        End Select
        AddToFiles msg, False, 0
        DecodeFile = msg
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function DecodeFile(lPath As String, lOutputFile As String, lInputFile As String) As String"
End Function

Public Function FindNextDecodeEventIndex() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(findnextdecodeeventindex).txt")
For i = 0 To lEvents.eCount
    If lEvents.eEvent(i).eType = eDecode Then
        FindNextDecodeEventIndex = i
        Exit For
    End If
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindNextDecodeEventIndex() As Integer"
End Function

Public Function DecodeEventPresentInQue() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(decodeeventpresentinuque).txt")
For i = 0 To lEvents.eCount
    If lEvents.eEvent(i).eType = eDecode Then
        DecodeEventPresentInQue = True
        Exit For
    End If
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function DecodeEventPresentInQue() As Boolean"
End Function

Public Sub AdjustStatus(lStatus As gStatus, Optional lExtended As String, Optional lExtended2 As String, Optional lIntro As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
Dim typeDevice As String, Result As String
Dim lcap As String
lcap = "Audiogen - "
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(adjuststatus).txt")
frmMain.lblSpaceLeft.Caption = ""
Select Case lStatus
Case sDoneBurning
    frmTaskbar.Caption = lcap & "Done burning"
    frmMain.lblStatus.Caption = "Done burning"
    frmMain.imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture
    frmMain.ResetFileTreeView
Case sIdle
    i = GetRnd(10)
    Select Case i
    Case 1
        frmMain.lblStatus.Caption = "Audiogen is ready"
    Case 2
        frmMain.lblStatus.Caption = "Audiogen is idle"
    Case 3
        frmMain.lblStatus.Caption = "Ready for commands"
    Case 4
        frmMain.lblStatus.Caption = "Standing by"
    Case 5
        frmMain.lblStatus.Caption = "Awaiting orders"
    Case 6
        frmMain.lblStatus.Caption = "Audiogen is loaded"
    Case 7
        frmMain.lblStatus.Caption = "Awaiting human interface"
    Case 8
        frmMain.lblStatus.Caption = "Prepared to process audio"
    Case 9
        frmMain.lblStatus.Caption = "Awaiting Input"
    Case 10
        frmMain.lblStatus.Caption = "Go ahead user"
    End Select
    frmTaskbar.Caption = lcap & frmMain.lblStatus.Caption
    lPlayer.pStatus = sIdle
    frmMain.imgStop.Picture = frmGraphics.imgStop4.Picture
    frmMain.imgPlay.Picture = frmGraphics.imgPlay1.Picture
Case sStop
    frmTaskbar.Caption = lcap & "playback stopped"
    If VerifyStatus(sStop) = True Then
        lPlayer.pStatus = sStop
        StopPlayback
    End If
Case sPaused
    frmTaskbar.Caption = lcap & "Paused"
    If VerifyStatus(sPaused) = True Then
        lPlayer.pStatus = sPaused
        Select Case lPlayer.pFileType
        Case fAllFileTypes
            frmMain.lblStatus = PauseMultimedia(frmMain.lblAliasname)
        End Select
    End If
Case sSelectFile
    If VerifyStatus(sSelectFile) = True Then
        If lPlayer.pFileType = fAllFileTypes Or lPlayer.pFileType = fCDAudio And Len(lExtended) <> 0 Then
            CloseAll
            If GetStatusMultimedia(LCase(Trim(frmMain.lblAliasname.Caption))) = "playing" Then StopPlayback
            frmMain.lblAliasname.Caption = Time$ & Date$ & GetRnd(10000)
            If Right(LCase(lExtended), 4) = ".avi" Then
                typeDevice = "AviVideo"
            ElseIf Right(LCase(lExtended), 4) = ".rmi" Or Right(LCase(lExtended), 4) = ".mid" Then
                typeDevice = "sequencer"
            ElseIf Right(LCase(lExtended), 4) = ".vob" Then
                typeDevice = "MPEGVideo"
            ElseIf Right(LCase(lExtended), 4) = ".cda" Then
                msg = Right(Trim(lExtended), 6)
                msg = left(msg, 2)
                If Int(msg) <> 0 Then
                    lPlayer.pStatus = sPlay
                    lPlayer.pFilename = msg
                    lPlayer.pFileType = fCDAudio
                    frmMain.lblStatus = "Playing " & msg
                    SetPlaybackObjects msg
                    frmMain.ctlRipper.Play Int(msg), 0, frmMain.ctlRipper.GetTrackLength(Int(msg))
                End If
                Exit Sub
            Else
                typeDevice = "MPEGVideo"
            End If
            If lSettings.sFullScreenVideo = False Then
                Result = OpenMultimedia(frmMain.fraVideo.hwnd, frmMain.lblAliasname.Caption, lExtended, typeDevice)
            ElseIf lSettings.sFullScreenVideo = True Then
                frmVideo.Show
                frmVideo.Width = Screen.Width
                frmVideo.Height = Screen.Height
                frmVideo.fraVideo.Width = Screen.Width
                frmVideo.fraVideo.Height = Screen.Height
                Result = OpenMultimedia(frmVideo.fraVideo.hwnd, frmMain.lblAliasname.Caption, lExtended, typeDevice)
            End If
            frmMain.lblStatus = Result
            If Result = "Success" Then
                StartPlayback lExtended, lIntro
            ElseIf left(LCase(Result), 18) = "a problem occurred" Then
                MsgBox "There was a problem playing the file '" & lExtended & "'", vbExclamation
            End If
        End If
    End If
    Dim m As String
    m = lPlayer.pFilename
    m = GetFileTitle(m)
    frmTaskbar.Caption = lcap & "Play: " & m
Case sDecode
    frmTaskbar.Caption = lcap & "Decode"
    If VerifyStatus(sDecode) = True Then
        If lPlayer.pFileType = fMpegLayer3 And Len(lExtended2) <> 0 Then
            AddEvent eDecode, lExtended, lExtended2
            lPlayer.pFilename = lExtended2
        End If
    End If
Case sBurn
    frmTaskbar.Caption = lcap & "Burn"
    If VerifyStatus(sSelectFile) = True Then
        StartBurn
    End If
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub AdjustStatus(lStatus As gStatus, Optional lExtended As String, Optional lExtended2 As String)"
End Sub

Public Sub LoadFiles()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, n As Integer, msg2 As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(loadfiles).txt")
lFiles.fCount = ReadINI(lIniFiles.iPlaylists, "Settings", "Count", 0)
If lFiles.fCount <> 0 Then
    For i = 1 To lFiles.fCount
        msg = ReadINI(lIniFiles.iPlaylists, Str(i), "Filename", ""): DoEvents
        If Len(msg) <> 0 Then
            If DoesFileExist(msg) = True Then
                n = n + 1
                lFiles.fFile(n).fFilename = msg
                msg2 = lFiles.fFile(n).fFilename
                msg2 = GetFileTitle(msg2): DoEvents
                If Len(msg2) <> 0 Then frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2, 7
            End If
        End If
    Next i
    lFiles.fCount = n
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LoadFiles()"
End Sub

Public Sub LoadSettings()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(loadsettings).txt")
lIniFiles.iCDAudioTracks = App.Path & "\inis\a_cdaudiotracks.ini"
lIniFiles.iPlaylists = App.Path & "\inis\a_playlists.ini"
lIniFiles.iRecientPlaylist = App.Path & "\inis\a_recient.ini"
lIniFiles.iSettings = App.Path & "\inis\a_settings.ini"
lIniFiles.iEffects = App.Path & "\inis\a_effects.ini"
lIniFiles.iErrorLog = App.Path & "\inis\a_errorlog.ini"
lIniFiles.iSkins = App.Path & "\inis\a_skins.ini"
lDrives.dCurrentDrive = ReadINI(lIniFiles.iSettings, "Settings", "LastDrive", "")
lSettings.sFullScreenVideo = ReadINI(lIniFiles.iSettings, "Settings", "FullScreenVideo", True)
lSettings.sSupportedMedia = ReadINI(lIniFiles.iSettings, "Settings", "SupportedMedia", "*.avi;*.qt;*.mov;*.snd;*.mpg;*.mpa;*.mpv;*.enc;*.m1v;*.mp2;*.mp3;*.mpe;*.mpeg;*.mpm;*.au;*.snd;*.aif;*.aiff;*.aifc;*.wav;*.wmv;*.wma")
lSettings.sSupportedAudio = ReadINI(lIniFiles.iSettings, "Settings", "SupportedAudio", "*.wav;*.mp3;*.wma;*.wmv;*.snd;*.au;*.ogg")
lSettings.sSupportedVideo = ReadINI(lIniFiles.iSettings, "Settings", "SupportedVideo", "*.avi;*.qt;*.mov;*.mpg;*.mpv;*.enc;*.m1v;*.mpe;*.mpeg;*.wmv")
lSettings.sAutoEject = ReadINI(lIniFiles.iSettings, "Settings", "AutoEject", True)
lSettings.sFinalize = ReadINI(lIniFiles.iSettings, "Settings", "Finalize", True)
lSettings.sNormalize = ReadINI(lIniFiles.iSettings, "Settings", "Normalize", False)
lSettings.sShowSplash = ReadINI(lIniFiles.iSettings, "Settings", "ShowSplash", True)
lSettings.sTestMode = ReadINI(lIniFiles.iSettings, "Settings", "TestMode", False)
lSettings.sInitialHeight = Int(ReadINI(lIniFiles.iSettings, "Settings", "InitialHeight", 8700))
lSettings.sInitialWidth = Int(ReadINI(lIniFiles.iSettings, "Settings", "InitialWidth", 8800))
lSettings.sInitialTop = Int(ReadINI(lIniFiles.iSettings, "Settings", "InitialTop", 100))
lSettings.sShowAboutDetailsOnStartup = ReadINI(lIniFiles.iSettings, "Settings", "ShowAboutDetailsOnStartup", False)
lSettings.sInitialLeft = Int(ReadINI(lIniFiles.iSettings, "Settings", "InitialLeft", 100))
lSettings.sFirstRun = ReadINI(lIniFiles.iSettings, "Settings", "FirstRun", True)
lSettings.sFindSelectIndex = ReadINI(lIniFiles.iSettings, "Settings", "FindSelectIndex", 0)
lSettings.sConvertKHZ = ReadINI(lIniFiles.iSettings, "Settings", "ConvertKHZ", True)
lSettings.sName = ReadINI(lIniFiles.iSettings, "Settings", "Name", "")
lSettings.sPassword = ReadINI(lIniFiles.iSettings, "Settings", "Password", "")
lSettings.sProcessScripts = ReadINI(lIniFiles.iSettings, "Settings", "ProcessScripts", False)
lSettings.sLastCDDrive = ReadINI(lIniFiles.iSettings, "Settings", "LastCDDrive", 0)
lSettings.sBitrate = ReadINI(lIniFiles.iSettings, "Settings", "Bitrate", 8)
lSettings.sCheckTaskbarStatus = ReadINI(lIniFiles.iSettings, "Settings", "CheckTaskbarStatus", False)
lSettings.sAlwaysOnTop = ReadINI(lIniFiles.iSettings, "Settings", "AlwaysOnTop", False)
frmMain.cboBitrate.ListIndex = lSettings.sBitrate
If lSettings.sAlwaysOnTop = True Then AlwaysOnTop frmMain, True
If lSettings.sNormalize = False Then
    frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks1.Picture
Else
    frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks2.Picture
End If
If lSettings.sAutoEject = False Then
    frmMain.imgAutoEject.Picture = frmGraphics.imgAutoEject1.Picture
Else
    frmMain.imgAutoEject.Picture = frmGraphics.imgAutoEject2.Picture
End If
frmMain.Width = lSettings.sInitialWidth
frmMain.Height = lSettings.sInitialHeight
frmMain.top = lSettings.sInitialTop
frmMain.left = lSettings.sInitialLeft
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LoadSettings()"
End Sub

Public Function AddToFiles(lFilename As String, lLoadedFromPlaylist As Boolean, Optional lIndex As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(addtofiles).txt")
If Len(lFilename) <> 0 Then
    msg = lFilename
    msg = GetFileTitle(msg)
    msg2 = left(lFilename, Len(lFilename) - Len(msg))
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        AddToFiles = i
        Exit Function
    End If
    If Len(msg) <> 0 And Len(msg2) <> 0 Then
        If lLoadedFromPlaylist = True Then
            lFiles.fFile(lIndex).fFilename = lFilename
            AddToFiles = lIndex
        Else
            If lLoadedFromPlaylist = False Then
                lFiles.fCount = lFiles.fCount + 1
                lFiles.fFile(lFiles.fCount).fFilename = lFilename
                If LCase(ReadINI(lIniFiles.iPlaylists, Str(lFiles.fCount), "Filename", "")) <> LCase(lFilename) Then
                    WriteINI lIniFiles.iPlaylists, "Settings", "Count", Str(lFiles.fCount)
                    WriteINI lIniFiles.iPlaylists, Str(lFiles.fCount), "Filename", lFilename
                End If
            End If
            AddToFiles = lFiles.fCount
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function AddToFiles(lFilename As String, lLoadedFromPlaylist As Boolean, Optional lIndex As Integer) As Integer"
End Function

Public Function FindFileIndexByFilename(lFilename As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, msg2 As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(findfileindexbyfilename).txt")
msg = lFilename
If Len(lFilename) <> 0 Then
    For i = 0 To lFiles.fCount
        msg2 = lFiles.fFile(i).fFilename
        If Len(msg2) <> 0 Then
            If InStr(1, LCase(msg2), LCase(msg)) Then
                'If IsFileFuzzyMatchByFilename(lFilename) = True Then
                '    MsgBox "TRUE"
                'End If
                FindFileIndexByFilename = i
                Exit For
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindFileIndexByFilename(lFilename As String) As Integer"
End Function

Public Function FindFileIndex(lFullPath As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(findfileindex).txt")
If Len(lFullPath) <> 0 Then
    For i = 0 To lFiles.fCount
        If LCase(lFiles.fFile(i).fFilename) = LCase(lFullPath) Then
            FindFileIndex = i
            Exit For
        Else
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindFileIndex(lFullPath As String) As Integer"
End Function

Public Function GetWaveKHZ(lFilename As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim Buf As String * 58, beg As Byte
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getwavekhz).txt")
If DoesFileExist(lFilename) = True Then
    Open lFilename For Binary As #1
        Get #1, 1, Buf
    Close #1
    beg = InStr(1, Buf, "WAVE")
    If beg <> 0 Then
        GetWaveKHZ = Sredi(Mid(Buf, 25, 1)) / 17 * 11 & " KHz"
    Else
        GetWaveKHZ = "44 khz"
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetWaveKHZ(lFilename As String) As Long"
End Function

Private Function Sredi(ByVal accStr As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Sredi = Trim(Str(Asc(accStr)))
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Function Sredi(ByVal accStr As String) As String"
End Function

Public Sub AddtoRecient(lFile As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(addtorecient).txt")
If Len(lFile) <> 0 Then
    If DoesFileExist(lFile) = True Then
        i = Int(ReadINI(lIniFiles.iRecientPlaylist, "Settings", "Count", 0)) + 1
        WriteINI lIniFiles.iRecientPlaylist, "Settings", "Count", Str(i)
        WriteINI lIniFiles.iRecientPlaylist, Str(i), "Filename", lFile
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub AddtoRecient(lFile As String)"
End Sub
