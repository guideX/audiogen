Attribute VB_Name = "mdlFunctions"
Option Explicit

Public Function ExchangeWinVer(lReturnInteger As Boolean, lInt As Integer, Optional lString As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lReturnInteger = True Then
    If Len(lString) <> 0 Then
        Select Case lString
        Case "Other"
            ExchangeWinVer = "0"
        Case "Windows 95"
            ExchangeWinVer = "1"
        Case "Windows 98"
            ExchangeWinVer = "2"
        Case "Windows ME"
            ExchangeWinVer = "3"
        Case "Windows NT"
            ExchangeWinVer = "4"
        Case "Windows 2000"
            ExchangeWinVer = "5"
        Case "Windows XP"
            ExchangeWinVer = "6"
        Case "Windows 2003"
            ExchangeWinVer = "7"
        End Select
    Else
        Exit Function
    End If
Else
    Select Case lInt
    Case 0
        ExchangeWinVer = "Other"
    Case 1
        ExchangeWinVer = "Windows 95"
    Case 2
        ExchangeWinVer = "Windows 98"
    Case 3
        ExchangeWinVer = "Windows ME"
    Case 4
        ExchangeWinVer = "Windows NT"
    Case 5
        ExchangeWinVer = "Windows 2000"
    Case 6
        ExchangeWinVer = "Windows XP"
    Case 7
        ExchangeWinVer = "Windows 2003"
    End Select
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ExchangeWinVer(lReturnInteger As Boolean, lInt As Integer, Optional lString As String)"
End Function

Public Function ParseString(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(parsestring).txt")
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
ParseString = left(Str1, Len(Str1) - Len(Str2) - 1)
Err = 0
ErrHandler:
    If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ParseString(lWhole As String, lStart As String, lEnd As String)"
End Function

Public Function FindComoboxIndex(lCombo As ComboBox, lText As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lCombo.ListCount
        If LCase(lCombo.List(i)) = LCase(lText) Then
            FindComoboxIndex = i
            Exit For
            Exit Function
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindComoboxIndex(lCombo As ComboBox, lText As String) As Integer"
End Function

Public Function ReturnDirCompliant(lText As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Again:
If InStr(lText, "/") Or InStr(lText, "\") Or InStr(lText, "*") Or InStr(lText, ":") Or InStr(lText, Chr(34)) Or InStr(lText, "<") Or InStr(lText, ">") Or InStr(lText, "|") Or InStr(lText, "?") Then
    If InStr(lText, "/") Then
        lText = Replace(lText, "/", "_")
    ElseIf InStr(lText, "\") Then
        lText = Replace(lText, "\", "_")
    ElseIf InStr(lText, "*") Then
        lText = Replace(lText, "*", "_")
    ElseIf InStr(lText, ":") Then
        lText = Replace(lText, ":", "_")
    ElseIf InStr(lText, Chr(34)) Then
        lText = Replace(lText, Chr(34), "_")
    ElseIf InStr(lText, "<") Then
        lText = Replace(lText, "<", "_")
    ElseIf InStr(lText, ">") Then
        lText = Replace(lText, ">", "_")
    ElseIf InStr(lText, "?") Then
        lText = Replace(lText, "?", "_")
    ElseIf InStr(lText, "|") Then
        lText = Replace(lText, "|", "_")
    End If
Else
    ReturnDirCompliant = lText
    Exit Function
End If
GoTo Again
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ReturnDirCompliant(lText As String) As String"
End Function

Public Function VerifyStatus(lRequestedStatus As gStatus) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(verifystatus).txt")
If lRequestedStatus <> 0 Then
    Select Case lRequestedStatus
    Case sBurn
        If lPlayer.pStatus = sPlay Then
            mbox = MsgBox("You are currently burning an audio track. Would you like to stop the burn process so you can play?", vbQuestion + vbYesNo, "Stop burning?")
            If mbox = vbYes Then
                frmMain.ctlBurn.Abort
                frmMain.prgPercentDone.value = 0
                frmMain.prgSpaceLeft.value = 0
                frmMain.lblSpaceLeft.Caption = ""
                frmMain.lblTimeLeft.Caption = ""
                frmMain.lblTrackNumber.Caption = ""
                frmMain.lblStatus.Caption = "Canceled burn process"
                Pause 0.2
                AdjustStatus sIdle
                VerifyStatus = True
            Else
                VerifyStatus = False
            End If
        End If
    Case sPlay
        If lPlayer.pStatus = sPlay Or lPlayer.pStatus = sSelectFile Then
            StopPlayback
            VerifyStatus = True
        ElseIf lPlayer.pStatus = sBurn Then
            mbox = MsgBox("You are currently burning. Would you like to abort the burning process so you can play an audio or video file?", vbQuestion + vbYesNo, "Stop burning?")
            If mbox = vbYes Then
                frmMain.ctlBurn.Abort
                VerifyStatus = True
            Else
                VerifyStatus = False
            End If
        Else
            VerifyStatus = True
        End If
    Case sSelectFile
        If lPlayer.pStatus = sSelectFile Or lPlayer.pStatus = sPlay Then
            StopPlayback
            VerifyStatus = True
        Else
            VerifyStatus = True
        End If
    Case sDecode
        VerifyStatus = True
    Case sNormalize
        VerifyStatus = True
    Case sStop
        VerifyStatus = True
    Case sIdle
        VerifyStatus = True
    Case sPaused
        VerifyStatus = True
    Case sRip
        VerifyStatus = True
    End Select
Else
    VerifyStatus = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function VerifyStatus(lRequestedStatus As eEventType) As Boolean"
End Function

Public Function FindTreeViewIndex(lText As String, lTreeView As TreeView) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(findtreeviewindex).txt")
If Len(lText) <> 0 Then
    For i = 1 To lTreeView.Nodes.Count
        If InStr(1, lText, lTreeView.Nodes(i).Text, vbTextCompare) Then
            FindTreeViewIndex = i
            Exit For
        End If
    Next
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindTreeViewIndex(lText As String, lTreeView As TreeView) As Integer"
End Function

Public Function ReadFile(lFile As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim o As Integer, msg As String
o = FreeFile
If DoesFileExist(lFile) = True Then
    Open lFile For Input As #o
        ReadFile = StrConv(InputB(LOF(o), o), vbUnicode)
    Close #o
End If
End Function

Public Function ErrorAid(lErrorNumber As Long, lDescription As String, lSource As String, Optional lDoNotShow As Boolean, Optional lDoNotLog As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, tNode As node
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(erroraid).txt")
Err.Number = 0
If lDoNotShow = False Then
    If lSettings.sDebugMode = True Then
        i = FindTreeViewIndex(lSource, frmErrorAid.TreeView1)
        frmErrorAid.Show
        If i <> 0 Then
            frmErrorAid.TreeView1.Nodes.Add FindTreeViewIndex(lSource, frmErrorAid.TreeView1), tvwChild, , lDescription & " - " & Time & " - " & Date & " - " & lErrorNumber
        Else
            frmErrorAid.TreeView1.Nodes.Add , , , lSource
            frmErrorAid.TreeView1.Nodes.Add FindTreeViewIndex(lSource, frmErrorAid.TreeView1), tvwChild, , lDescription & " - " & Time & " - " & Date & " - " & lErrorNumber
        End If
    Else
        Exit Function
    End If
End If
If lDoNotLog = False Then
    If Len(lIniFiles.iErrorLog) <> 0 Then
        WriteINI lIniFiles.iErrorLog, lSource & " - " & Time & " - " & Date, "Description", lErrorNumber & ": " & lDescription
    Else
        WriteINI App.Path & "\inis\a_errorlog", lSource & " - " & Time & " - " & Date, "Description", lErrorNumber & ": " & lDescription
    End If
End If
If Err.Number <> 0 Then Err.Number = 0
End Function

Public Function GetRnd(Num As Integer) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getrnd).txt")
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetRnd(Num As Integer) As Integer"
End Function

Public Function SaveFile(lFilename As String, lText As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(savefile).txt")
If Len(lFilename) <> 0 And Len(lText) <> 0 Then
    Open lFilename For Output As #1
    Print #1, lText
    Close #1
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function SaveFile(lFilename As String, lText As String) As Boolean"
End Function

Public Function AddEvent(lEventType As eEventType, lInputFile As String, lOutputFile As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(addevent).txt")
Dim i As Integer
i = lEvents.eCount + 1
If i <> 0 Then
    With lEvents.eEvent(i)
        .eInputFile = lInputFile
        .eOutputFile = lOutputFile
        .eType = lEventType
    End With
    lEvents.eCount = i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function AddEvent(lEventType As eEventType, lInputFile As String, lOutputFile As String) As Integer"
End Function

Public Function ProcessNextEvent() As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(processnextevent).txt")
ProcessEvent lEvents.eCount
ProcessNextEvent = lEvents.eCount
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ProcessNextEvent() As Integer"
End Function

Public Function LeftZeroPad(s As String, n As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Len(s) < n Then
    LeftZeroPad = String$(n - Len(s), "0") & s
Else
    LeftZeroPad = s
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function LeftZeroPad(s As String, n As Integer) As String"
End Function

Public Function ProcessEvent(lIndex As Integer) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, msg4 As String, e As Integer, i As Integer, mbox As VbMsgBoxResult
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(processevent).txt")
If lEvents.eProcessing = True Then Exit Function
If lIndex <> 0 Then
    With lEvents.eEvent(lIndex)
        msg = .eInputFile
        msg = GetFileTitle(msg)
        msg2 = left(.eInputFile, Len(.eInputFile) - Len(msg))
        msg3 = .eOutputFile
        msg3 = GetFileTitle(msg3)
        Select Case lEvents.eEvent(lIndex).eType
        Case eRip
            lEvents.eProcessing = True
            ProcessEvent = RipTrack(Int(.eInputFile), .eOutputFile)
            RemoveEvent lIndex
        Case eEncode
            lEvents.eProcessing = True
            ProcessEvent = EncodeFile(.eInputFile, .eOutputFile)
            RemoveEvent lIndex
        Case eDecode
            If Len(msg) <> 0 And Len(msg2) <> 0 And Len(msg3) <> 0 Then
                lEvents.eProcessing = True
                ProcessEvent = DecodeFile(msg2, msg3, msg): DoEvents
                RemoveEvent lIndex
            End If
        Case eNormalize
            lEvents.eProcessing = True
            ProcessEvent = NormalizeFile(.eInputFile, .eOutputFile)
            RemoveEvent lIndex
        Case eStartBurn
            If lBurnQue.bCount <> 0 Then
                frmMain.lblStatus.Caption = "Burning Audio tracks"
                frmMain.imgBurn.Picture = frmGraphics.imgAbort1.Picture
                For i = 1 To lBurnQue.bCount
                    If DoesFileExist(lBurnQue.bFiles(i)) = True And Right(LCase(lBurnQue.bFiles(i)), 4) = ".wav" Then
                        DoEvents
                        If LCase(GetWaveKHZ(lBurnQue.bFiles(i))) <> "44 khz" Then
                        Else
                            e = frmMain.ctlBurn.AddAudioTrack(Trim(lBurnQue.bFiles(i)), 0)
                            If e <> 0 Then
                                mbox = MsgBox("The File: '" & lBurnQue.bFiles(i) & "' could not be used, would you like to abort the burn process?", vbCritical + vbYesNo)
                                If mbox = vbYes Then
                                    MsgBox "Burn was canceled! Your disc will be useable for another burn!", vbInformation
                                    frmMain.ctlBurn.Reset
                                    ClearBurnQue
                                    InitCDBurner
                                    AdjustStatus sIdle
                                    frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture
                                    frmMain.tvwToBurn.Nodes.Clear
                                    Exit Function
                                Else
                                    lBurnQue.bFiles(i) = ""
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
            lEvents.eProcessing = True
            While frmMain.ctlBurn.TstReady <> 0
                mbox = MsgBox("Insert a CD-R or CD-RW now, then click ok", vbExclamation + vbOKCancel)
                If mbox = vbCancel Then GoTo Fuck
            Wend
Fuck:
            e = 0
            frmMain.tmrCheckNormalize.Enabled = False
            e = frmMain.ctlBurn.WriteTracks
            If e <> 0 Then
                MsgBox "Error Burning: " & e
            End If
            ProcessEvent = "."
            RemoveEvent lIndex
        End Select
    End With
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ProcessEvent(lIndex As Integer) As String"
End Function

Public Function ResetMainForm(Optional lClearTrees As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(resetmainform).txt")
With frmMain
    .prgPercentDone.value = 0
    .prgSpaceLeft.value = 0
    .lblFormat.Caption = ""
    .lblBitrate.Caption = ""
    .lblTransferStatus.Caption = ""
    .lblTrackNumber.Caption = ""
    .lblSpaceLeft.Caption = ""
    .lblTimeDisplay.Caption = ""
    .lblTimeLeft.Caption = ""
    .lblTitle.Caption = ""
    .lblStatus.Caption = ""
    .lblTrackNumber.Caption = ""
    .ResetFileTreeView
    If lClearTrees = True Then
        .tvwToBurn.Nodes.Clear
        .tvwFiles.Nodes.Clear
    End If
End With
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ResetMainForm()"
End Function
