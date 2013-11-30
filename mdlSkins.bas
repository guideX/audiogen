Attribute VB_Name = "mdlSkins"
Option Explicit

Enum eCombineMode
    cRgn_None = 0
    cRgn_And = 1
    cRgn_Or = 2
    cRgn_XOr = 3
    cRgn_Diff = 4
    cRgn_Copy = 5
End Enum
Enum eShapeTypes
    sOther = 0
    sRectRgn = 1
    sEllipce = 2
    sRoundRectRgn = 3
End Enum
Private Type gRegions
    rRgn As Long
    X1 As Long
    X2 As Long
    X3 As Long
    Y1 As Long
    Y2 As Long
    Y3 As Long
End Type
Private Type gWindowPos
    wTitleBarHeight As Integer
    wWindowBorder As Integer
End Type
Private Type gWindowSize
    sWidth As Long
    sHeight As Long
    sLeft As Long
    sTop As Long
End Type
Private Type gShape
    sName As String
    sType As eShapeTypes
    sRgn As gRegions
    sCombineMode As eCombineMode
    sDestRgn As Integer
    sSrcRgn1 As Integer
    sSrcRgn2 As Integer
    sEnabled As Boolean
End Type
Private Type gSkin
    sEnabled As Boolean
    sName As String
    sShape(20) As gShape
    sSkinSettings As gWindowSize
    sShapeCount As Integer
    sFilename As String
    sFilepath As String
    sGraphic As String
End Type
Private Type gSkins
    sSkinIndex As Integer
    sSkin(15) As gSkin
    sCount As Integer
End Type
Global lSkins As gSkins
Global lWindowSize As gWindowSize
Global lWindowPos As gWindowPos

Public Function OpenSkin(lFilename As String) As Integer
On Local Error Resume Next
Dim i As Integer, X As Integer, msg As String, f As Integer, msg2 As String ', A As Integer
If Len(lFilename) = 0 Then Exit Function
msg2 = lFilename
msg2 = GetFileTitle(msg2)
For i = 1 To lSkins.sCount
    If LCase(lSkins.sSkin(i).sFilename) = LCase(msg2) Then
        OpenSkin = i
        Exit Function
    End If
Next i
i = 0
If Len(lFilename) <> 0 Then
    With lSkins
        i = .sCount + 1
        .sCount = i
        .sSkin(i).sSkinSettings.sWidth = ReadINI(lFilename, "Settings", "Width", 200)
        .sSkin(i).sSkinSettings.sHeight = ReadINI(lFilename, "Settings", "Height", 200)
        .sSkin(i).sGraphic = ReadINI(lFilename, "Settings", "Graphic", "")
        .sSkin(i).sName = ReadINI(lFilename, "Settings", "Name", "Default Skin")
        .sSkin(i).sShapeCount = ReadINI(lFilename, "Settings", "ShapeCount", 0)
        .sSkin(i).sFilename = msg2
        .sSkin(i).sFilepath = left(lFilename, Len(lFilename) - Len(.sSkin(i).sFilename))
        If Len(.sSkin(i).sName) <> 0 Then .sSkin(i).sEnabled = True
        If .sSkin(i).sShapeCount <> 0 Then
            For X = 1 To .sSkin(i).sShapeCount
                msg = "rgn" & X
                .sSkin(i).sShape(X).sEnabled = ReadINI(lFilename, msg, "enabled", "")
                .sSkin(i).sShape(X).sName = ReadINI(lFilename, msg, "name", "")
                .sSkin(i).sShape(X).sDestRgn = ReadINI(lFilename, msg, "destrgn", 0)
                .sSkin(i).sShape(X).sSrcRgn1 = ReadINI(lFilename, msg, "srcrgn1", 0)
                .sSkin(i).sShape(X).sSrcRgn2 = ReadINI(lFilename, msg, "srcrgn2", 0)
                .sSkin(i).sShape(X).sCombineMode = ReadINI(lFilename, msg, "combinemode", 0)
                .sSkin(i).sShape(X).sRgn.X1 = ReadINI(lFilename, msg, "x1", 0)
                .sSkin(i).sShape(X).sRgn.X2 = ReadINI(lFilename, msg, "x2", 0)
                .sSkin(i).sShape(X).sRgn.X3 = ReadINI(lFilename, msg, "x3", 0)
                .sSkin(i).sShape(X).sRgn.Y1 = ReadINI(lFilename, msg, "y1", 0)
                .sSkin(i).sShape(X).sRgn.Y2 = ReadINI(lFilename, msg, "y2", 0)
                .sSkin(i).sShape(X).sRgn.Y3 = ReadINI(lFilename, msg, "y3", 0)
                .sSkin(i).sShape(X).sType = ReadINI(lFilename, msg, "type", 1)
            Next X
        End If
    End With
    OpenSkin = i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function OpenSkin(lFilename As String) As Integer"
End Function

Public Sub ApplySkin(lForm As Form, lIndex As Integer)
On Local Error Resume Next
lForm.Width = lSkins.sSkin(lIndex).sSkinSettings.sWidth
lForm.Height = lSkins.sSkin(lIndex).sSkinSettings.sHeight
lSkins.sSkinIndex = lIndex
LoadShape lForm, lIndex
If DoesFileExist(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sGraphic) = True Then lForm.Picture = LoadPicture(lSkins.sSkin(lIndex).sFilepath & lSkins.sSkin(lIndex).sGraphic)
lForm.Visible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ApplySkin(lForm As Form, lIndex As Integer)"
End Sub

Public Sub GetWindowSettings(lHandle As Long)
On Local Error Resume Next
Dim lClientPos As RECT, i As Long, lWSize As RECT, lBorderWidth As Long
i = GetWindowRect(lHandle, lWSize)
i = GetClientRect(lHandle, lClientPos)
lWindowPos.wTitleBarHeight = lWSize.Bottom - lWSize.top - lClientPos.Bottom - lBorderWidth
lWindowPos.wWindowBorder = lWSize.Right - lWSize.left - lClientPos.Right - 2
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub GetWindowSettings(lHandle As Long)"
End Sub

Public Function FindSkinIndexByFilename(lFilename As String) As Integer
On Local Error Resume Next
Dim i As Integer
If Len(lFilename) <> 0 Then
    For i = 1 To lSkins.sCount
        If LCase(lFilename) = LCase(lSkins.sSkin(i).sFilename) Then FindSkinIndexByFilename = i
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindSkinIndexByFilename(lFilename As String) As Integer"
End Function

Public Function FindSkinIndex(lName As String) As Integer
On Local Error Resume Next
Dim i As Integer
If Len(lName) <> 0 Then
    For i = 1 To lSkins.sCount
        If LCase(lName) = LCase(lSkins.sSkin(i).sName) Then FindSkinIndex = i
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Function FindSkinIndex(lName As String) As Integer"
End Function

Public Sub SetImageBox(lImageBox As Image, lImageBox1 As Image, lImageBox2 As Image, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long)
On Local Error Resume Next
If Len(lImg1) <> 0 Then
    lImageBox1.Picture = LoadPicture(lImg1)
    With lImageBox
        .Picture = lImageBox1.Picture
        .left = lLeft
        .top = lTop
    End With
End If
If Len(lImg2) <> 0 Then lImageBox2.Picture = LoadPicture(lImg2)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetImageBox(lImageBox As Image, lImageBox1 As Image, lImageBox2 As Image, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long)"
End Sub

Public Sub SetPictureBox(lPictureBox As PictureBox, lPictureBox1 As PictureBox, lPictureBox2 As PictureBox, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long)
On Local Error Resume Next
If Len(lImg1) <> 0 And Len(lImg1) <> 1 Then
    lPictureBox.left = lLeft
    lPictureBox.top = lTop
    lPictureBox.Picture = LoadPicture(lImg1)
    lPictureBox1.Picture = LoadPicture(lImg1)
    If Len(lImg2) <> 0 Then lPictureBox2.Picture = LoadPicture(lImg2)
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetPictureBox(lPictureBox As PictureBox, lPictureBox1 As PictureBox, lPictureBox2 As PictureBox, lImg1 As String, lImg2 As String, lLeft As Long, lTop As Long)"
End Sub

Public Sub SetLabel(lLabel As Label, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
On Local Error Resume Next
lLabel.left = lLeft
lLabel.top = lTop
lLabel.Width = lWidth
lLabel.Height = lHeight
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SetLabel(lLabel As Label, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)"
End Sub

Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)
On Local Error Resume Next
Dim i As Integer, X As Long, Y As Long, tmp As Long
GetWindowSettings lForm.hWnd
X = lWindowPos.wWindowBorder
Y = lWindowPos.wTitleBarHeight
With lSkins.sSkin(lSkins.sSkinIndex)
    For i = 1 To .sShapeCount
        If .sShape(i).sEnabled = True Then
            Select Case .sShape(i).sType
            Case 1
                .sShape(i).sRgn.rRgn = CreateRectRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 2
                .sShape(i).sRgn.rRgn = CreateEllipticRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2)
            Case 3
                .sShape(i).sRgn.rRgn = CreateRoundRectRgn(X + .sShape(i).sRgn.X1, Y + .sShape(i).sRgn.Y1, X + .sShape(i).sRgn.X2, Y + .sShape(i).sRgn.Y2, .sShape(i).sRgn.X3, .sShape(i).sRgn.Y3)
            End Select
        End If
    Next i
    For i = 1 To .sShapeCount
        If .sShape(i).sEnabled = True Then
            If .sShape(i).sCombineMode <> 0 And .sShape(i).sDestRgn <> 0 And .sShape(i).sSrcRgn1 <> 0 And .sShape(i).sSrcRgn2 <> 0 Then
                tmp = CombineRgn(.sShape(.sShape(i).sDestRgn).sRgn.rRgn, .sShape(.sShape(i).sSrcRgn1).sRgn.rRgn, .sShape(.sShape(i).sSrcRgn2).sRgn.rRgn, .sShape(i).sCombineMode)
            End If
        End If
    Next i
    SetWindowRgn lForm.hWnd, .sShape(1).sRgn.rRgn, True
End With
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub LoadShape(lForm As Form, lSkinIndex As Integer)"
End Sub

Public Sub PictureBoxMouseMove(lButton As Integer, lPictureBox As Image, lPic1 As Image, lPic2 As Image, lX As Single, lY As Single)
On Local Error Resume Next
If lButton = 0 Then
ElseIf lButton = 1 Then
    lX = frmMain.ScaleX(lX) * 1.8
    lY = frmMain.ScaleY(lY) * 1.8
    If lPictureBox.Picture = lPic2.Picture Then
        If lX > lPictureBox.Width Or lX < -1 Or lY > lPictureBox.Height Or lY < -1 Then lPictureBox.Picture = lPic1.Picture
    ElseIf lPictureBox.Picture = lPic1.Picture Then
        If lX < lPictureBox.Width And lX > -1 And lY < lPictureBox.Height And lY > -1 Then lPictureBox.Picture = lPic2.Picture
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub PictureBoxMouseMove(lButton As Integer, lPictureBox As Image, lPic1 As Image, lPic2 As Image, lX As Single, lY As Single)"
End Sub
