VERSION 5.00
Begin VB.Form frmEditMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   88
      X2              =   88
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   5
      X2              =   80
      Y1              =   113
      Y2              =   113
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   5
      X2              =   80
      Y1              =   114
      Y2              =   114
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   5.333
      X2              =   80
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   5.333
      X2              =   80
      Y1              =   21
      Y2              =   21
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   136
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   136
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblCut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Cut"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblCopy 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Copy"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblRemove 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Remove"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Delete"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblSelectAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Select All"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblFind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Find"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblFindNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Find Next"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   5
      X2              =   80
      Y1              =   74
      Y2              =   74
   End
   Begin VB.Label lblPaste 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Paste"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   5.333
      X2              =   80
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Label lblUndo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "        Undo"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RefreshEditMenu()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lMenus.mEditMenuIndex <> 1 And lblUndo.BackColor <> &HE0E0E0 Then lblUndo.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 2 And lblCut.BackColor <> &HE0E0E0 Then lblCut.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 3 And lblCopy.BackColor <> &HE0E0E0 Then lblCopy.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 4 And lblPaste.BackColor <> &HE0E0E0 Then lblPaste.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 5 And lblDelete.BackColor <> &HE0E0E0 Then lblDelete.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 6 And lblRemove.BackColor <> &HE0E0E0 Then lblRemove.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 7 And lblSelectAll.BackColor <> &HE0E0E0 Then lblSelectAll.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 8 And lblFind.BackColor <> &HE0E0E0 Then lblFind.BackColor = &HE0E0E0
If lMenus.mEditMenuIndex <> 9 And lblFindNext.BackColor <> &HE0E0E0 Then lblFindNext.BackColor = &HE0E0E0
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub RefreshEditMenu()"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii = 27 Then
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_startup).txt")
AlwaysOnTop Me, True
Me.Icon = frmGraphics.Icon
lMenus.mEditMenuVisible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
AlwaysOnTop Me, False
lMenus.mEditMenuVisible = False
frmMain.imgEdit.Visible = False
frmMain.imgEdit2.Visible = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub lblCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblCopy.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 3
If lblCopy.ForeColor = &H404040 Then
    If Button = 0 Then
        lblCopy.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblcopy_mouseup).txt")
If Button = 1 Then
    lblCopy.BackColor = &H8000000F
    msg = frmMain.tvwFiles.SelectedItem.Text
    If Err.Number <> 0 Then
        frmMain.lblStatus.Caption = "Could not copy " & msg
        MsgBox "This file could not be found!", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If i <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
            lClipboard.cCopyOnly = True
            lClipboard.cFileToCopy = lFiles.fFile(i).fFilename
            frmMain.lblStatus.Caption = "File copied to clipboard"
            lblPaste.ForeColor = &H404040
        Else
            lClipboard.cCopyOnly = True
            lClipboard.cFileToCopy = ""
            frmMain.lblStatus.Caption = "Could not copy " & msg
            MsgBox "This file could not be found!", vbExclamation, App.Title
        End If
    End If
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblCut.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 2
If lblCut.ForeColor = &H404040 Then
    If Button = 0 Then
        lblCut.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblCut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblcut_mouseup).txt")
If Button = 1 Then
    lblCut.BackColor = &H8000000F
    msg = frmMain.tvwFiles.SelectedItem.Text
    If Err.Number <> 0 Then
        frmMain.lblStatus.Caption = "Could not cut " & msg
        MsgBox "This file could not be found!", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
    If Len(msg) <> 0 Then
        i = FindFileIndexByFilename(msg)
        If i <> 0 And DoesFileExist(lFiles.fFile(i).fFilename) = True Then
            lClipboard.cCopyOnly = False
            lClipboard.cFileToCopy = lFiles.fFile(i).fFilename
            frmMain.lblStatus.Caption = "File cut"
            lblPaste.ForeColor = &H404040
        Else
            lClipboard.cCopyOnly = True
            lClipboard.cFileToCopy = ""
            frmMain.lblStatus.Caption = "Could not cut " & msg
            MsgBox "This file could not be found!", vbExclamation, App.Title
        End If
    End If
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblCut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblDelete.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 5
If lblDelete.ForeColor = &H404040 Then
    If Button = 0 Then
        lblDelete.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lbldelete_mouseup).txt")
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblFind.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 8
If lblFind.ForeColor = &H404040 Then
    If Button = 0 Then
        lblFind.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblfind_mouseup).txt")
    Unload Me
    frmFind.Show
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFindNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblFindNext.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFindNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 9
If lblFindNext.ForeColor = &H404040 Then
    If Button = 0 Then
        lblFindNext.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFindNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFindNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblPaste.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 4
If lblPaste.ForeColor = &H404040 Then
    If Button = 0 Then
        lblPaste.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblPaste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, msg2 As String, msg3 As String, lext As String, i As Integer
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblpaste_mouseup).txt")
If Button = 1 Then
    If Len(lClipboard.cFileToCopy) <> 0 Then
        msg = lClipboard.cFileToCopy
        lext = Right(msg, 4)
        msg = left(msg, Len(msg) - 4)
        msg2 = Right(msg, 3)
        If lClipboard.cCopyOnly = True Then
            If Right(msg2, 1) = ")" And left(msg2, 1) = "(" Then
                msg2 = left(msg2, Len(msg2) - 1)
                msg2 = Right(msg2, Len(msg2) - 1)
                i = Int(msg2) + 1
                lClipboard.cNewFilename = left(msg, Len(msg) - 3) & "(" & Trim(Str(i)) & ")" & lext
                frmSearching.Show: DoEvents
                msg3 = lClipboard.cNewFilename
                msg3 = GetFileTitle(msg3)
                frmSearching.Label1.Caption = "Copying file " & msg3
                FileCopy lClipboard.cFileToCopy, lClipboard.cNewFilename
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg3, 7
                AddToFiles lClipboard.cNewFilename, False
                lClipboard.cFileToCopy = ""
                lClipboard.cNewFilename = ""
                lClipboard.cCopyOnly = False
                Unload frmSearching
            Else
                frmSearching.Show: DoEvents
                msg2 = lClipboard.cFileToCopy
                msg3 = Right(msg2, 4)
                msg2 = left(msg2, Len(msg2) - 4)
                lClipboard.cNewFilename = msg2 & " (1)" & msg3
                msg3 = lClipboard.cNewFilename
                msg3 = GetFileTitle(msg3)
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg3, 7
                frmSearching.Label1.Caption = "Copying file " & msg3
                FileCopy lClipboard.cFileToCopy, lClipboard.cNewFilename
                AddToFiles lClipboard.cNewFilename, False
                lClipboard.cFileToCopy = ""
                lClipboard.cNewFilename = ""
                lClipboard.cCopyOnly = False
                Unload frmSearching
            End If
        End If
    Else
        lClipboard.cFileToCopy = ""
        lClipboard.cNewFilename = ""
        lClipboard.cCopyOnly = False
    End If
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblPaste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRemove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblRemove.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRemove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 6
If lblRemove.ForeColor = &H404040 Then
    If Button = 0 Then
        lblRemove.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblRemove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Button = 1 Then
    lblRemove.BackColor = &H8000000F
    RemovePlaylistEntry FindFileIndexByFilename(frmMain.tvwFiles.SelectedItem.Text)
    frmMain.tvwFiles.Nodes.Remove frmMain.tvwFiles.SelectedItem.Index
    If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblremove_mouseup).txt")
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblRemove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblSelectAll.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 7
If lblSelectAll.ForeColor = &H404040 Then
    If Button = 0 Then
        lblSelectAll.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Button = 1 Then
    lblSelectAll.BackColor = &H8000000F
    Unload Me
    For i = 1 To frmMain.tvwFiles.Nodes.Count
        If Len(frmMain.tvwFiles.Nodes(i).Text) <> 0 Then
            frmMain.tvwFiles.Nodes(i).Selected = True
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblUndo.BackColor = vbWhite
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblUndo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lMenus.mEditMenuIndex = 1
If lblUndo.ForeColor = &H404040 Then
    If Button = 0 Then
        lblUndo.BackColor = &HC0C0C0
        RefreshEditMenu
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblUndo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lblUndo.BackColor = &H8000000F
    If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndeditmenu_lblundo_mouseup).txt")
    Unload Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub
