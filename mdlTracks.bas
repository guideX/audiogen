Attribute VB_Name = "mdlTracks"
Option Explicit

Public Function GetCDTracks(lToc As String) As Boolean
On Local Error Resume Next
Dim i As Integer
If Len(lToc) <> 0 Then
'    If ReadINI(lIniFiles.iCD, lToc, "Enabled", False) = True Then
'        lTracks.tDiscLen = ReadINI(lIniFiles.iCD, lToc, "DiscLen", "")
'        lTracks.tArtist = ReadINI(lIniFiles.iCD, lToc, "Artist", "")
'        lTracks.tTitle = ReadINI(lIniFiles.iCD, lToc, "Title", "")
'        lTracks.tGenre = ReadINI(lIniFiles.iCD, lToc, "Genre", "")
'        lTracks.tLabel = ReadINI(lIniFiles.iCD, lToc, "Label", "")
'        lTracks.tYear = ReadINI(lIniFiles.iCD, lToc, "Year", "")
'        If lTracks.tCount > 300 Then lTracks.tCount = 300
'        If Len(lTracks.tArtist) <> 0 And Len(lTracks.tTitle) <> 0 Then
'            For i = 1 To lTracks.tCount
'                lTracks.tTrack(i).tName = ReadINI(lIniFiles.iCD, lToc, str(i), "")
'                lTracks.tTrack(i).tLength = ReadINI(lIniFiles.iCD, lToc, str(i) & "L", "")
'            Next i
'        End If
'        GetCDTracks = True
'    Else
'        GetCDTracks = False
'        Exit Function
'    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetCDTracks(lToc As String) As Boolean"
End Function

Public Function GetSimpleTracks()
On Local Error Resume Next
'Dim i As Integer, X As Integer, Z As Integer, h As Boolean, msg As String
'SelectCDDrive
'frmMain.Ripper.Init: DoEvents
'frmMain.Ripper.OpenDriveByLetter lRipperSettings.eDriveLetter
'DoEvents
'Z = frmMain.Ripper.TrackCount
'frmMain.lblInfo.Caption = "Please wait ..."
'Pause 0.2
'lTracks.tCount = Z
'If Z <> 0 Then
'    lRipperSettings.eAvailable = True
'    InitMediaToc lRipperSettings.eDriveLetter
'    msg = GetTOC
'    lRipperSettings.eDiscID = msg
'    If Len(lRipperSettings.eDiscID) <> 0 Then
'        If GetCDTracks(lRipperSettings.eDiscID) = False Then
'            If lEvents.eSettings.iFreeDB.cEnabled = True Then
'                frmMain.wskFreeDB.Close
'                frmMain.wskFreeDB.LocalPort = GetRnd(1500)
'                frmMain.wskFreeDB.Connect lEvents.eSettings.iFreeDB.cServer, 8880
'            End If
'        Else
'            frmMain.lblInfo.Caption = lTracks.tArtist & " \ " & lTracks.tTitle
'            Pause 0.2
'        End If
'    End If
'Else
'    lRipperSettings.eAvailable = False
'End If
End Function

Public Function ReturnFreeDBQueryString(lToc As String) As String
On Error GoTo errChk
Dim strTocData() As String, sum As Long, tmp As Long, idx As Integer, msg As String, msg2 As String, lTrackNum As String, mediaID As String
msg = Trim$(lToc)
If (msg = "" Or InStr(1, msg, " ") = 0) Then Exit Function
strTocData = Split(msg, " ", 100, vbTextCompare)
lTrackNum = UBound(strTocData)
For idx = 1 To lTrackNum
''''''''''''''    lTracks.tTrack(idx).tLength = (Val(strTocData(idx)) - Val(strTocData(idx - 1))) \ 75
    'colTrackTimes.Add (Val(strTocData(idx)) - Val(strTocData(idx - 1))) \ 75
Next idx
'm_AlbumSeconds = (Val(strTocData(m_Tracks)) \ 75) - (Val(strTocData(0)) \ 75)
'lTracks.tDiscLen = (Val(strTocData(lTrackNum)) \ 75) - (Val(strTocData(0)) \ 75)
For idx = 0 To lTrackNum - 1
    tmp = Val(strTocData(idx)) \ 75
    Do While tmp > 0
        sum = sum + (tmp Mod 10)
        tmp = tmp \ 10
    Loop
Next idx
''''''''mediaID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(lTracks.tDiscLen), 4) & LeftZeroPad(Hex$(lTrackNum), 2))
msg2 = mediaID & "+" & lTrackNum
For idx = 0 To lTrackNum - 1
    msg2 = msg2 & "+" & strTocData(idx)
Next
msg2 = msg2 & "+" & (Val(strTocData(lTrackNum)) \ 75)
ReturnFreeDBQueryString = msg2
Exit Function
errChk:
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function ReturnFreeDBQueryString(lToc As String) As String"
End Function
