Attribute VB_Name = "mdlCDDrives"
Option Explicit
Dim FS

Public Sub LoadDrives()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim d, lToc As String, lCDDriveLoaded As Boolean
lDrives.dCount = 0
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(loaddrives).txt")
Set FS = CreateObject("scripting.filesystemobject")
For Each d In FS.Drives
    Select Case d.DriveType
        Case 2
            lDrives.dCount = lDrives.dCount + 1
            lDrives.dDrive(lDrives.dCount).dDriveLetter = d
            lDrives.dDrive(lDrives.dCount).dDriveType = dHardDrive
            frmMain.tvwSources.Nodes.Add , , , lDrives.dDrive(lDrives.dCount).dDriveLetter, 10
        Case 4
            lDrives.dCDCount = lDrives.dCDCount + 1
            lDrives.dCount = lDrives.dCount + 1
            lDrives.dDrive(lDrives.dCount).dDriveLetter = d
            lDrives.dDrive(lDrives.dCount).dDriveType = dCDDrive
            lDrives.dDrive(lDrives.dCount).dDriveNumber = lDrives.dCDCount
            If Len(lDrives.dCurrentDrive) = 0 Then
                lCDDriveLoaded = True
                SelectCurrentCDDrive lDrives.dDrive(lDrives.dCount).dDriveLetter
            End If
            frmMain.tvwSources.Nodes.Add , , , lDrives.dDrive(lDrives.dCount).dDriveLetter, 9
            frmMain.cboDrives.AddItem lDrives.dDrive(lDrives.dCount).dDriveLetter & " - (" & Trim(frmMain.ctlRipper.DriveStringByNumber(lDrives.dDrive(lDrives.dCount).dDriveNumber)) & ")"
            If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LoadDrives() (Trim(frmMain.ctlRipper.DriveStringByNumber(lDrives.dDrive(lDrives.dCount).dDriveNumber)))"
            'frmMain.lblStatus.Caption = "CD Devices Detected: " & lDrives.dCDCount
            If InitMediaToc(d) = True Then
                lDrives.dDrive(lDrives.dCount).dToc = GetTOC
            End If
    End Select
Next
If Len(lDrives.dCurrentDrive) <> 0 And lCDDriveLoaded = False Then SelectCurrentCDDrive lDrives.dCurrentDrive
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LoadDrives()"
End Sub

Public Function SelectCurrentCDDrive(lDriveLetter As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, f As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(selectcurrentcddrive).txt")
If Len(lDriveLetter) <> 0 Then
    i = frmMain.ctlRipper.Init
    If i = 0 Then
        f = frmMain.ctlRipper.OpenDriveByLetter(lDriveLetter)
        If f <> 0 Then
            ProcessCDDriveError f
            Exit Function
        End If
        lDrives.dCurrentDrive = lDriveLetter
        lTracks.tTrackCount = frmMain.ctlRipper.TrackCount
        WriteINI lIniFiles.iSettings, "Settings", "LastDrive", lDriveLetter
        SelectCurrentCDDrive = True
    Else
        ProcessCDDriveError i
    End If
Else
    MsgBox "No drive letter specified", vbExclamation, "SelectCurrentCDDrive"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function SelectCurrentCDDrive(lDriveLetter As String, lDriveNumber As Integer) As Boolean"
End Function

Public Sub ProcessCDDriveError(lNumber As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(processcddriveerror).txt")
Select Case lNumber

Case 687
    MsgBox "Error: Cannot complete request, reached parameter out of range error", vbExclamation, "CD Drive"
Case 690
    MsgBox "Error: Please insert an audio cd and try again", vbExclamation, "CD Drive"
Case 684
    Dim mbox As VbMsgBoxResult
    mbox = MsgBox("You do not have ASPI Driver installed. Without these drivers you can not either burn nor rip. Would you like to install an ASPI layer? (Requires reboot)", vbYesNo + vbQuestion, "Audiogen")
    If mbox = vbYes Then
        Shell App.Path & "\external\aspiupd.exe", vbNormalFocus
        End
    End If
    'MsgBox "Error: Can not load your aspi drivers!", vbExclamation, "CD Drive"
Case 690
    MsgBox "Error: Audiogen can not find a cd in drive '" & lDrives.dCurrentDrive & "' rip process can not continue!", vbExclamation, App.Title
Case 691
    MsgBox "Error: CD has changed!", vbExclamation, "CD Drive"
Case 692
    MsgBox "Error: Not ready!", vbExclamation, "CD Drive"
Case 693
    MsgBox "Error: Seek error!", vbExclamation, "CD Drive"
Case 694
    MsgBox "Error: Read error!", vbExclamation, "CD Drive"
Case 695
    MsgBox "Error: No CD Detected!", vbExclamation, "CD Drive"
Case 696
    MsgBox "Error: General error!", vbExclamation, "CD Drive"
Case 697
    MsgBox "Error: Illegal CD Change!", vbExclamation, "CD Drive"
Case 698
    MsgBox "Error: Drive not found!", vbExclamation, "CD Drive"
Case 699
    MsgBox "Error: DAC Unable!", vbExclamation, "CD Drive"
Case 700
    MsgBox "Error: ASPI error!", vbExclamation, "CD Drive"
Case 701
    MsgBox "Error: User break!", vbExclamation, "CD Drive"
Case 702
    MsgBox "Error: CD time out!", vbExclamation, "CD Drive"
Case 703
    MsgBox "Error: Out of memory!", vbExclamation, "CD Drive"
Case 704
    MsgBox "Error: Sector not found!", vbExclamation, "CD Drive"
Case 712
    MsgBox "Error: Free hard disk space error!", vbExclamation, "CD Drive"
Case 713
    MsgBox "Error: Device not found!", vbExclamation, "CD Drive"
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ProcessCDDriveError(lNumber As Integer)"
End Sub

Public Function SelectCDDriveByCombo() As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lDriveLetter As String, lDriveNumber
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(selectcddrivebycombo).txt")
frmMain.ctlRipper.Init
lDriveLetter = left(frmMain.cboDrives.Text, 2)
If Len(lDriveLetter) <> 0 Then
    SelectCurrentCDDrive lDriveLetter
Else
    MsgBox "No drive specified", vbExclamation, "SelectCurrentCDDriveByNumber"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function SelectCDDriveByCombo() As Boolean"
End Function

Public Function GetDriveType(lDrive As String) As gDriveType
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getdrivetype).txt")
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
            GetDriveType = lDrives.dDrive(i).dDriveType
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetDriveType(lDrive As String) As gDriveType"
End Function

Public Function GetDriveIndex(lDrive As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getdriveindex).txt")
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
            GetDriveIndex = i
            Exit For
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetDriveIndex(lDrive As String) As gDriveType"
End Function

Public Function GetCDDriveIndex(lDrive As String) As Integer
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, m As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getcddriveindex).txt")
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If lDrives.dDrive(i).dDriveType = dCDDrive Then
            m = m + 1
            If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
                GetCDDriveIndex = m
                Exit For
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetCDDriveIndex(lDrive As String) As Integer"
End Function
