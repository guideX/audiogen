Attribute VB_Name = "mdlFreeDB"
Option Explicit

Public Function GetFreeDBQueryString(lToc As String) As String
Dim strTocData() As String, sum As Long, tmp As Long, idx As Integer, msg As String, msg2 As String, lTrackNum As String, mediaID As String
On Error GoTo errChk
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(getfreedbquerystring).txt")
msg = Trim$(lToc)
If (msg = "" Or InStr(1, msg, " ") = 0) Then Exit Function
strTocData = Split(msg, " ", 100, vbTextCompare)
lTrackNum = UBound(strTocData)
For idx = 1 To lTrackNum
    lTracks.tTrack(idx).tLength = (Val(strTocData(idx)) - Val(strTocData(idx - 1))) \ 75
Next idx
lTracks.tDiscLen = (Val(strTocData(lTrackNum)) \ 75) - (Val(strTocData(0)) \ 75)
For idx = 0 To lTrackNum - 1
    tmp = Val(strTocData(idx)) \ 75
    Do While tmp > 0
        sum = sum + (tmp Mod 10)
        tmp = tmp \ 10
    Loop
Next idx
mediaID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(lTracks.tDiscLen), 4) & LeftZeroPad(Hex$(lTrackNum), 2))
msg2 = mediaID & "+" & lTrackNum
For idx = 0 To lTrackNum - 1
    msg2 = msg2 & "+" & strTocData(idx)
Next
msg2 = msg2 & "+" & (Val(strTocData(lTrackNum)) \ 75)
GetFreeDBQueryString = msg2
Exit Function
errChk:
    If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function GetFreeDBQueryString(lToc As String) As String"
End Function

Public Sub CloseFreeDB(lWinsock As Winsock)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(closefreedb).txt")
lWinsock.Close
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub CloseFreeDB(lWinsock As Winsock)"
End Sub

Public Sub ConnectToFreeDBServer(lWinsock As Winsock, lServer As String, lPort As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(connecttofreedbserver).txt")
With lWinsock
    .Close
    .LocalPort = GetRnd(1500)
    .Connect lServer, lPort
End With
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ConnectToFreeDBServer(lServer As String, lPort As Long)"
End Sub

Public Sub ConnectFreeDB(lWinsock As Winsock)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(connectfreedb).txt")
lWinsock.SendData "cddb hello guidex@team-nexgen.com " & lWinsock.LocalHostName & " Audiogen 1." & App.Minor & "." & App.Revision & vbCrLf: DoEvents
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ConnectFreeDB(lWinsock As Winsock)"
End Sub

Public Sub ProcessFreeDBData(lWinsock As Winsock)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, lefty As String, msg4 As String, i As Integer, j As Integer, lGenre As String, msg5 As String
Dim lCat As String, DiscID As String
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(processfreedbdata).txt")
lWinsock.GetData msg, vbString
If left(msg, 3) = "200" Then
    If InStr(LCase(msg), "hello and welcome") Then
        lWinsock.SendData "cddb query " & Replace(GetFreeDBQueryString(GetTOC), "+", " ") & vbCrLf
    Else
        msg2 = Right(msg, Len(msg) - 4)
        lefty = left(msg2, 1)
        lCat = lefty & ParseString(msg2, left(msg2, 1), " ")
        msg2 = Right(msg2, Len(msg2) - Len(lCat) - 1)
        lefty = left(msg2, 1)
        DiscID = lefty & ParseString(msg2, left(msg2, 1), " ")
        lWinsock.SendData "cddb read " & lCat & " " & DiscID & vbCrLf
        lTracks.tArtist = Trim(ParseString(msg2, " ", "/"))
        msg4 = Right(msg2, Len(msg2) - Len(ParseString(msg2, left(msg2, 1), "/")) - 3)
        lTracks.tTitle = left(msg4, Len(msg4) - 2)
    End If
ElseIf left(msg, 3) = "210" Then
    j = lTracks.tCount
    lGenre = ParseString(msg, "210 ", "CD ")
    lGenre = Right(lGenre, Len(lGenre) - 3)
    lefty = UCase(left(lGenre, 1))
    lGenre = lefty & ParseString(lGenre, left(lGenre, 1), " ")
    lTracks.tGenre = lGenre
    For i = 0 To j
        msg5 = "EXTD="
        msg3 = "TTITLE" & i & "="
        msg4 = "TTITLE" & i + 1 & "="
        If InStr(msg, msg3) And InStr(msg, msg3) Then
            msg2 = ParseString(msg, msg3, msg4)
            DoEvents
            If Len(msg2) <> 0 Then
                msg2 = Right(msg2, Len(msg2) - Len(msg3) + 1)
                msg2 = left(msg2, Len(msg2) - 2)
                lTracks.tTrack(i + 1).tName = msg2
            End If
        Else
            lTracks.tTrack(i).tName = "Track " & i
        End If
        DoEvents
    Next i
    msg3 = "TTITLE" & j - 1 & "="
    msg2 = ParseString(msg, msg3, msg5)
    msg2 = Right(msg2, Len(msg2) - Len(msg3) + 1)
    lTracks.tTrack(j).tName = left(msg2, Len(msg2) - 2)
    lWinsock.Close
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ProcessFreeDBData(lWinsock As Winsock)"
End Sub
