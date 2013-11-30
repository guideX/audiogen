Attribute VB_Name = "mdlSubs"
Option Explicit
Private Type SHELLEXECUTEINFO
     cbSize As Long
     fMask As Long
     hwnd As Long
     lpVerb As String
     lpFile As String
     lpParameters As String
     lpDirectory As String
     nShow As Long
     hInstApp As Long
     lpIDList As Long
     lpClass As String
     hkeyClass As Long
     dwHotKey As Long
     hIcon As Long
     hProcess As Long
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Public Sub DisplayPlaylist()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, f(7)
frmMain.tvwPlaylist.Visible = True
frmMain.tvwPlaylist.Nodes.Clear
frmMain.fraVideo.Visible = False
frmMain.fraPlaylist.left = 200
frmMain.fraPlaylist.top = 2300
frmMain.tvwPlaylist.Nodes.Add , , , "Wave Audio (*.wav) (0 files)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "Windows Media Audio (*.wma) (0 files)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "Mpeg Layer 3 (*.mp3) (0 files)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "OGG (*.ogg) (0 files)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "MPEG Video (*.mpeg)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "DivX and AVI (*.avi)", 5
frmMain.tvwPlaylist.Nodes.Add , , , "Quick Time Movies (*.mov)", 5
frmMain.tvwPlaylist.Nodes(1).Visible = False
frmMain.tvwPlaylist.Nodes(2).Visible = False
frmMain.tvwPlaylist.Nodes(3).Visible = False
frmMain.tvwPlaylist.Nodes(4).Visible = False
frmMain.tvwPlaylist.Nodes(5).Visible = False
frmMain.tvwPlaylist.Nodes(6).Visible = False
frmMain.tvwPlaylist.Nodes(7).Visible = False
For i = 0 To lFiles.fCount
    If Len(lFiles.fFile(i).fFilename) <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
        msg = lFiles.fFile(i).fFilename
        msg = GetFileTitle(msg)
        Select Case LCase(Right(msg, 4))
        Case ".wav"
            f(1) = f(1) + 1
            frmMain.tvwPlaylist.Nodes(1).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 1, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(1).Text = "Wave Audio (*.wav) (" & f(1) & " files)"
        Case ".wma"
            f(2) = f(2) + 1
            frmMain.tvwPlaylist.Nodes(2).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 2, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(2).Text = "Windows Media Audio (*.wma) (" & f(2) & " files)"
        Case ".mp3"
            f(3) = f(3) + 1
            frmMain.tvwPlaylist.Nodes(3).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 3, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(3).Text = "Mpeg Layer 3 (*.mp3) (" & f(3) & " files)"
        Case ".ogg"
            f(4) = f(4) + 1
            frmMain.tvwPlaylist.Nodes(4).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 4, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(4).Text = "OGG (*.ogg) (" & f(4) & " files)"
        Case "mpeg"
            f(5) = f(5) + 1
            frmMain.tvwPlaylist.Nodes(5).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 5, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(5).Text = "Mpeg Video (*.mpeg) (" & f(5) & " files)"
        Case ".avi"
            f(6) = f(6) + 1
            frmMain.tvwPlaylist.Nodes(6).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 6, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(6).Text = "DivX and AVI (*.avi) (" & f(6) & " files)"
        Case ".mov"
            f(7) = f(7) + 1
            frmMain.tvwPlaylist.Nodes(7).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 7, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(7).Text = "Quick Time Movies (*.mov) (" & f(7) & " files)"
        Case ".mpg"
            f(5) = f(5) + 1
            frmMain.tvwPlaylist.Nodes(5).Visible = True
            frmMain.tvwPlaylist.Nodes.Add 5, tvwChild, , msg, 7
            frmMain.tvwPlaylist.Nodes(5).Text = "Mpeg Video (*.mpeg) (" & f(5) & " files)"
        End Select
    End If
Next i
frmMain.fraPlaylist.Visible = True
frmMain.tvwPlaylist.Visible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DisplayPlaylist()"
End Sub

Public Function EncodeWaveToMp3FromTreeview(lTreeView As TreeView, Optional lWMA As Boolean) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lext As String, msg As String, msg2 As String, i As Integer, msg3 As String
msg = lTreeView.SelectedItem.Text
try:
If Len(msg) <> 0 Then
    i = FindFileIndexByFilename(msg)
    If i <> 0 Then
        lext = Right(LCase(msg), 4)
        Select Case lext
        Case ".mp3"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Function
        Case ".wav"
            msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
            If lWMA = True Then
                msg3 = left(msg, Len(msg) - 4) & ".wma"
                EncodeWMA lFiles.fFile(i).fFilename, msg2 & msg3
            Else
                msg3 = left(msg, Len(msg) - 4) & ".mp3"
                EncodeFile lFiles.fFile(i).fFilename, msg2 & msg3
            End If
        Case ".wma"
            MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
            Exit Function
        End Select
    Else
        MsgBox "Cannot find file!", vbExclamation
    End If
Else
    msg = OpenDialog(frmMain, lSettings.sSupportedAudio, "Open Audio ...", CurDir)
    lext = Right(LCase(msg), 4)
    Select Case lext
    Case ".mp3"
        MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
        Exit Function
    Case ".wav"
        msg2 = left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
        If lWMA = True Then
            msg3 = left(msg, Len(msg) - 4) & ".wma"
            EncodeWMA lFiles.fFile(i).fFilename, msg2 & msg3
        Else
            msg3 = left(msg, Len(msg) - 4) & ".mp3"
            EncodeFile lFiles.fFile(i).fFilename, msg2 & msg3
        End If
    Case ".wma"
        MsgBox "This file is already encoded. Please select a Wave audio file (*.wav)", vbExclamation + vbSystemModal
        Exit Function
    End Select
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function EncodeWaveToMp3(Optional lFilename, Optional lTreeview As TreeView) As Boolean"
End Function

Public Sub SetCheckBoxValue(lCheckBox As CheckBox, lValue As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lValue = True Then
    lCheckBox.value = 1
Else
    lCheckBox.value = 0
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetCheckBoxValue(lCheckBox As CheckBox, lValue As Boolean)"
End Sub

Public Sub CheckKeyboardCommands(lCommands As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lCommands = 115 And lAlt = True Then
    EndProgram
End If
If lCommands = 18 Then
   lAlt = True
Else
    lAlt = False
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub CheckKeyboardCommands(lCommands As String)"
End Sub

Public Sub ProcessScript(lFile As String, Optional lForced As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, lTitle As String
If Len(lFile) <> 0 And lSettings.sProcessScripts = True Then
    msg = Trim(ReadFile(lFile))
    lTitle = lFile
    lTitle = GetFileTitle(lTitle)
    If Len(msg) <> 0 And Len(msg) <> 1 Then
        frmMain.ctlScript.ExecuteStatement lFile
        lScripts.sCurrentScript = lTitle
        Pause 0.01
        DoEvents
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ProcessScript(lFile As String)"
End Sub

Public Sub ChangeKHZ(lFilename As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(changekhz).txt")
frmKHZConverter.txtFilename.Text = lFilename
frmKHZConverter.ConvertKHZ
frmKHZConverter.Show 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ChangeKHZ()"
End Sub

Public Sub DisplayTestInformation()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = ReadFile(App.Path & "\documentation\testing.txt")
If Len(msg) <> 0 Then
    frmMain.ResetFileTreeView
    For i = 0 To Len(msg)
        If Len(msg) <> 0 Then
            msg2 = left(msg, Int(InStr(1, msg, vbCrLf)))
            If Len(msg2) <> 0 Then
                msg = Right(msg, Len(msg) - Len(msg2))
                msg2 = left(msg2, Len(msg2) - 1)
                msg2 = Right(msg2, Len(msg2) - 1)
                'If left(msg2, 11) = "Description" Then msg2 = "  " & msg2
                msg2 = "  " & msg2
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DisplayTestInformation()"
End Sub

Public Sub DisplayErrorInformation()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer
msg = ReadFile(lIniFiles.iErrorLog)
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(displayerrorinformation).txt")
If Len(msg) <> 0 Then
    frmMain.ResetFileTreeView
    For i = 0 To Len(msg)
        If Len(msg) <> 0 Then
            msg2 = left(msg, Int(InStr(1, msg, vbCrLf)))
            If Len(msg2) <> 0 Then
                msg = Right(msg, Len(msg) - Len(msg2))
                msg2 = left(msg2, Len(msg2) - 1)
                msg2 = Right(msg2, Len(msg2) - 1)
                If left(msg2, 11) = "Description" Then msg2 = "  " & msg2
                msg2 = "  " & msg2
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DisplayErrorInformation()"
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lFlag As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(alwaysontop).txt")
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hwnd, lFlag, myfrm.left / Screen.TwipsPerPixelX, myfrm.top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)"
End Sub

Public Sub SetAddress(lText As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(setaddress).txt")
lSettings.sSettingAddress = True
If Len(lText) <> 0 Then
    For i = 0 To frmMain.cboAddress.ListCount
        If LCase(frmMain.cboAddress.List(i)) = LCase(lText) Then
            frmMain.cboAddress.ListIndex = i
            Exit Sub
        End If
    Next i
    frmMain.cboAddress.AddItem lText, 0
    frmMain.cboAddress.ListIndex = 0
End If
lSettings.sSettingAddress = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetAddress(lText As String)"
End Sub

Public Sub Surf(lUrl As String, lHwnd As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As Long
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(surf).txt")
msg = ShellExecute(lHwnd, vbNullString, lUrl, vbNullString, "c:\", SW_SHOWNORMAL)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub Surf(lUrl As String)"
End Sub

Public Sub PlayWav(strPath As String, sndVal As sndConst)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(playwav).txt")
sndPlaySound strPath, sndVal
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub PlayWav(strPath As String, sndVal As sndConst)"
End Sub

Public Sub CheckMenus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lMenus.mEditMenuVisible = True Then Unload frmMenuEdit
If lMenus.mOpenMenuVisible = True Then Unload frmOpenMenu
If lMenus.mConvertMenuVisible = True Then Unload frmMenuConvert
If lMenus.mEffectsMenuVisible = True Then Unload frmMenuEffects
If lMenus.mFileMenuVisible = True Then Unload frmFileMenu
If lMenus.mFileMenuVisible = False Then
    frmMain.imgFile.Visible = False
    frmMain.imgFileOver.Visible = False
End If
If lMenus.mEditMenuVisible = False Then
    frmMain.imgEdit.Visible = False
    frmMain.imgEdit2.Visible = False
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub CheckMenus()"
End Sub

Public Sub ShowFileProperties(FormHwnd As Long, sFileName As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim udtSEI As SHELLEXECUTEINFO
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(showfileproporties).txt")
With udtSEI
       .cbSize = Len(udtSEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
       .hwnd = FormHwnd
       .lpVerb = "properties"
       .lpFile = sFileName
       .lpParameters = vbNullChar
       .lpDirectory = vbNullChar
       .nShow = 0
       .hInstApp = 0
       .lpIDList = 0
End With
Call ShellExecuteEX(udtSEI)
If udtSEI.hInstApp <= 32 Then MsgBox sFileName & "not found, There is an error", vbCritical, "Error"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ShowFileProperties(FormHwnd As Long, sFileName As String)"
End Sub

Public Function EndProgram() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbox As VbMsgBoxResult
Call Shell_NotifyIcon(NIM_DELETE, try)
frmMain.Visible = False
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(end).txt")
If lPlayer.pStatus = sBurn Then
    mbox = MsgBox("You are currently in a waiting mode or burning mode. Exiting will surely cause the burn to fail, and your rewritable disc to possibly be unusable. Are you sure you wish to continue?", vbYesNoCancel + vbQuestion, "Warning!")
    If mbox = vbNo Then
        EndProgram = False
        Exit Function
    ElseIf mbox = vbCancel Then
        EndProgram = False
        Exit Function
    End If
End If
CloseAll
WriteINI lIniFiles.iSettings, "Settings", "InitialWidth", frmMain.Width
WriteINI lIniFiles.iSettings, "Settings", "InitialHeight", frmMain.Height
WriteINI lIniFiles.iSettings, "Settings", "InitialTop", frmMain.top
WriteINI lIniFiles.iSettings, "Settings", "InitialLeft", frmMain.left
If frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks1.Picture Then
    WriteINI lIniFiles.iSettings, "Settings", "Normalize", "False"
ElseIf frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks2.Picture Then
    WriteINI lIniFiles.iSettings, "Settings", "Normalize", "True"
End If
If frmMain.imgAutoEject.Picture = frmGraphics.imgAutoEject1.Picture Then
    WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "False"
ElseIf frmMain.imgAutoEject.Picture = frmGraphics.imgAutoEject2.Picture Then
    WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "True"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub EndProgram()"
EndProgram = True
DoEvents
End
End Function

Public Sub SetProgress(lPercent As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lProgressClicked = True Then Exit Sub
If lPercent = 0 Then
    frmMain.imgProgress.left = 2000
Else
    frmMain.imgProgress.left = ((lPercent * frmMain.imgProgressChange.Width * 119) / 12000) + 2000
    frmMain.imgProgressYellow.Width = ((lPercent * frmMain.imgProgressChange.Width * 119) / 12000)
    If lPercent = 100 Then
        frmMain.imgProgress.left = 2000
        frmMain.imgProgressYellow.Width = 0
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetProgress(lPercent As Integer)"
End Sub

Public Sub Pause(lInterval)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim Current
Current = Timer
Do While Timer - Current < Val(lInterval)
DoEvents
Loop
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub Pause(lInterval)"
End Sub

Public Sub FormDrag(lFormname As Form)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ReleaseCapture
Call SendMessage(lFormname.hwnd, &HA1, 2, 0&)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub FormDrag(lFormname As Form)"
End Sub

Public Sub NodeCheck(ByVal node As MSComctlLib.node)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, f As Integer, j As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(nodecheck).txt")
msg = Trim(node.Text)
If Len(msg) <> 0 Then
    If Right(LCase(msg), 4) = ".cda" Then
        j = FindBurnQueIndexByFilename(msg)
        If node.Checked = True Then
            If j = 0 Then
                AddToBurnQue msg, App.Path & "\cdcopy\", node
                frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture
                'frmMain.imgCdCopy.Picture = frmGraphics.imgCDCopyDisabled.Picture
            End If
        ElseIf node.Checked = False Then
            If j <> 0 Then
                DeleteBurnQueEntry j
            End If
        End If
    Else
        f = FindFileIndexByFilename(msg)
        If f <> 0 Then
            If node.Checked = True Then
                If j = 0 Then
                    AddToBurnQue msg, left(lFiles.fFile(f).fFilename, Len(lFiles.fFile(f).fFilename) - Len(msg)), node
                End If
            ElseIf node.Checked = False Then
                j = FindBurnQueIndexByFilename(msg)
                If j <> 0 Then
                    DeleteBurnQueEntry j
                End If
            End If
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub NodeCheck(ByVal node As MSComctlLib.node)"
End Sub

Public Sub ImageBoxMouseMove(lButton As Integer, lImage As Image, lImage1 As Image, lImage2 As Image, lX As Single, lY As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lButton = 0 Then
    lX = frmMain.ScaleX(lX) * 1.8
    lY = frmMain.ScaleY(lY) * 1.8
    If lImage.Picture = lImage2.Picture Then
        If lX > lImage.Width Or lX < -1 Or lY > lImage.Height Or lY < -1 Then lImage.Picture = lImage1.Picture
    ElseIf lImage.Picture = lImage1.Picture Then
        If lX < lImage.Width And lX > -1 And lY < lImage.Height And lY > -1 Then lImage.Picture = lImage2.Picture
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ImageBoxMouseMove(lButton As Integer, lImage As Image, lImage1 As Image, lImage2 As Image, lX As Single, lY As Single)"
End Sub

Public Sub RemoveEvent(lIndex As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(removeevent).txt")
With lEvents.eEvent(lIndex)
    .eInputFile = ""
    .eOutputFile = ""
    .eType = 0
End With
lEvents.eCount = lEvents.eCount - 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub RemoveEvent(lIndex As Integer)"
End Sub

Public Sub ClearEvents()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(clearevents).txt")
For i = 0 To 100
    RemoveEvent i
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ClearEvents()"
End Sub
