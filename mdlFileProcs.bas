Attribute VB_Name = "mdlFileProcs"
Option Explicit

Public Function DoesFileExist(lFilename As String) As Boolean
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, dr As String
If Len(lFilename) <> 0 Then
    dr = left(lFilename, 3)
    For i = 0 To lDrives.dCount
        If left(LCase(dr), 2) = left(LCase(lDrives.dDrive(i).dDriveLetter), 2) And lDrives.dDrive(i).dDriveType = dCDDrive Then
            DoesFileExist = False
            Exit Function
        End If
    Next i
    msg = Dir(lFilename)
    If msg <> "" Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function DoesFileExist(lFilename As String) As Boolean"
End Function

Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, lDefault As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, RetVal As String, Worked As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(readini).txt")
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), lFile)
If Worked = 0 Then
    ReadINI = lDefault
Else
    ReadINI = left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, lDefault As String)"
End Function

Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal value As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(writeini).txt")
WritePrivateProfileString Section, Key, value, lFile
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)"
End Sub

Public Function GetFileTitle(lFilename As String) As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getfiletitle).txt")
If Len(lFilename) <> 0 Then
Again:
    If InStr(lFilename, "\") Then
        lFilename = Right(lFilename, Len(lFilename) - InStr(lFilename, "\"))
        If InStr(lFilename, "\") Then
            GoTo Again
        Else
            GetFileTitle = lFilename
        End If
    Else
        GetFileTitle = lFilename
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetFileTitle(lFilename As String) As String"
End Function

Public Function GetMyDocumentsDir() As String
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim sPath As String, IDL As Long, strPath As String, lngPos As Long
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getmydocumentsdir).txt")
If SHGetSpecialFolderLocation(0, sfidPROGRAMS, IDL) = NOERROR Then
    sPath = String$(255, 0)
    SHGetPathFromIDListA IDL, sPath
    lngPos = InStr(sPath, Chr(0))
    If lngPos > 0 Then
        strPath = left$(sPath, lngPos - 1)
    End If
End If
GetMyDocumentsDir = left(strPath, Len(strPath) - 19)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetMyDocumentsDir() As String"
End Function
