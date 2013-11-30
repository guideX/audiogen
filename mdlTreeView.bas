Attribute VB_Name = "mdlTreeView"
Option Explicit

Public Sub DisplayTreeviewFunction(lText As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(displaytreeviewfunction).txt")
If Len(lText) <> 0 Then
    Select Case LCase(lText)
    Case "playlist"
        frmMain.ResetFileTreeView
        For i = 0 To lFiles.fCount
            If Len(lFiles.fFile(i).fFilename) And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
                msg = lFiles.fFile(i).fFilename
                msg = GetFileTitle(msg)
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg, 7
            End If
        Next i
    Case "settings"
        If frmMain.tvwSources.SelectedItem.Children <> 0 Then
        Else
            With frmMain.tvwSources
                Dim f As Integer
                f = FindTreeViewIndex("Settings", frmMain.tvwSources)
                .Nodes.Add f, tvwChild, , "Always On Top"
                .Nodes.Add f, tvwChild, , "Auto Eject"
                .Nodes.Add f, tvwChild, , "Check Taskbar Status"
                .Nodes.Add f, tvwChild, , "Convert KHZ"
                .Nodes.Add f, tvwChild, , "Debug Mode"
                .Nodes.Add f, tvwChild, , "Finalize Disc"
                .Nodes.Add f, tvwChild, , "Full Screen Video"
                .Nodes.Add f, tvwChild, , "Handle Errors"
                .Nodes.Add f, tvwChild, , "Name"
                .Nodes.Add f, tvwChild, , "Normalize"
                .Nodes.Add f, tvwChild, , "Password"
                .Nodes.Add f, tvwChild, , "Process Scripts"
                .Nodes.Add f, tvwChild, , "Show Splash"
                .Nodes.Add f, tvwChild, , "Supported Media"
                .Nodes.Add f, tvwChild, , "Test Mode"
            End With
        End If
    Case "errors"
        DisplayErrorInformation
    End Select
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub DisplayTreeviewFunction(lText As String)"
End Sub

Public Sub DisplayDirectory(lPath As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, cCol1 As tSearch, cCol2 As tSearch, msg As String, f As Integer, msg2 As String, k As Integer, n As Integer, cdrom As Boolean, b As Integer, msg3 As String, lToc As String
Dim lDataExists As Boolean
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(displaydirectory).txt")
If GetDriveType(left(lPath, 2)) = dCDDrive Then cdrom = True
If cdrom = False Then
    frmMain.Tag = lPath
    frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture
    'frmMain.imgCdCopy.Picture = frmGraphics.imgCDCopyDisabled.Picture
    frmMain.ResetFileTreeView
CheckFiles:
    lTreeviewText = frmMain.tvwSources.SelectedItem.Text
    If Len(lPath) <> 0 Then GetDirs lPath, vbDirectory, cCol1: DoEvents
    If frmMain.tvwSources.SelectedItem.Children = 0 Then
        If cCol1.Count <> 0 Then
            For i = 1 To cCol1.Count
                msg = cCol1.Path(i)
                msg2 = GetFileTitle(msg)
                If Len(msg) <> 0 Then
                    frmMain.tvwSources.Nodes.Add frmMain.tvwSources.SelectedItem.Index, tvwChild, , msg2, 1
                End If
            Next i
        End If
    End If
Else
    frmMain.imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture
    frmMain.imgCdCopy.Picture = frmGraphics.imgCdCopy.Picture
    SelectCurrentCDDrive lDrives.dCurrentDrive
    DoEvents
    frmMain.ResetFileTreeView
End If
SetAddress lPath
GetFiles lPath, "*.*", vbNormal, cCol2: DoEvents
If cCol2.Count <> 0 Then
    For i = 1 To cCol2.Count
        If Len(cCol2.Path(i)) <> 0 Then
            msg = Right(cCol2.Path(i), Len(cCol2.Path(i)) - Len(lPath))
            If left(msg, 1) = "\" Then msg = Right(msg, Len(msg) - 1)
            If Right(msg, 1) = "\" Then msg = left(msg, Len(msg) - 1)
            If cdrom = False Then
                If InStr(1, ".qt;.au", Right(LCase(cCol2.Path(i)), 3), vbTextCompare) Or InStr(1, ".mov;.dat;.snd;.mpg;.mpa;.mpv;.enc;.m1v;.mp2;.mp3;.mpe;.mpm;.snd;.aif;.wav;.wmv;.wma", Right(LCase(cCol2.Path(i)), 4), vbTextCompare) Or InStr(1, ".mpeg;.aiff;.aifc", Right(LCase(cCol2.Path(i)), 5), vbTextCompare) Then
                    n = AddToFiles(cCol2.Path(i), False)
                    frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg, 7
                End If
            Else
                n = AddToFiles(cCol2.Path(i), False)
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg, 7
            End If
        End If
    Next i
End If
ErrorCheck:
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub TreeView1_DblClick()", True, False
End Sub

Public Function DecodeLocation(lText As String, lFullPath As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(decodelocation).txt")
If left(LCase(lFullPath), 12) = "my documents" Then
    If left(lFullPath, 1) = "\" Then
        DecodeLocation = GetMyDocumentsDir & "My Documents" & Right(lFullPath, Len(lFullPath) - 12)
    Else
        If lFullPath = "My Documents" Then
            DecodeLocation = GetMyDocumentsDir & "My Documents\"
        Else
            DecodeLocation = GetMyDocumentsDir & "My Documents\" & Right(lFullPath, Len(lFullPath) - 13)
        End If
    End If
ElseIf left(LCase(lFullPath), 7) = "desktop" Then
    If left(lFullPath, 1) = "\" Then
        DecodeLocation = GetMyDocumentsDir & "My Documents" & Right(lFullPath, Len(lFullPath) - 7)
    Else
        If lFullPath = "My Documents" Then
            DecodeLocation = GetMyDocumentsDir & "Desktop\"
        Else
            If LCase(lFullPath) = "desktop" Then
                DecodeLocation = GetMyDocumentsDir & "Desktop\"
            Else
                DecodeLocation = GetMyDocumentsDir & "Desktop\" & Right(lFullPath, Len(lFullPath) - 8)
            End If
        End If
    End If
ElseIf left(LCase(lFullPath), 8) = "my music" Then
    If left(lFullPath, 1) = "\" Then
        DecodeLocation = GetMyDocumentsDir & "My Documents\My Music" & Right(lFullPath, Len(lFullPath) - 8)
    Else
        If lFullPath = "My Music" Then
            DecodeLocation = GetMyDocumentsDir & "My Documents\My Music"
        Else
            DecodeLocation = GetMyDocumentsDir & "My Documents\My Music\" & Right(lFullPath, Len(lFullPath) - 9)
        End If
    End If
ElseIf left(LCase(lFullPath), 15) = "copied cd-audio" Then
    If left(lFullPath, 1) = "\" Then
        DecodeLocation = App.Path & "\cdcopy" & Right(lFullPath, Len(lFullPath) - 15)
    Else
        If lFullPath = "Copied CD-Audio" Then
            DecodeLocation = App.Path & "\cdcopy\"
        Else
            DecodeLocation = App.Path & "\cdcopy\" & Right(lFullPath, Len(lFullPath) - 16)
        End If
    End If
Else
    DecodeLocation = lFullPath
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function DecodeLocation(lText As String) As String"
End Function

Public Function HasLocation(lText As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(haslocation).txt")
HasLocation = True
Select Case LCase(lText)
Case "settings"
    HasLocation = False
Case "errors"
    HasLocation = False
Case "playlist"
    HasLocation = False
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function HasLocation(lText As String) As Boolean"
End Function

Public Sub OpenContainingFolder(lFilepath As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(opencontainingfolder).txt")
frmMain.tvwFiles.Nodes.Clear
If Len(lFilepath) <> 0 Then
    msg = lFilepath
    msg = GetFileTitle(msg)
    msg = left(lFilepath, Len(lFilepath) - Len(msg))
    DisplayDirectory msg
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub OpenContainingFolder(lDirectory As String)"
End Sub
