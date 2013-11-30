VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hidden"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   Visible         =   0   'False
   Begin VB.Image imgMenuArrow3 
      Height          =   195
      Left            =   4080
      Picture         =   "frmGraphics.frx":000C
      Top             =   5160
      Width           =   195
   End
   Begin VB.Image imgMenuArrow1 
      Height          =   195
      Left            =   4080
      Picture         =   "frmGraphics.frx":0256
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image imgMenuArrow2 
      Height          =   195
      Left            =   4080
      Picture         =   "frmGraphics.frx":04A0
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image imgAudiogenLogo 
      Height          =   495
      Left            =   3120
      Top             =   5160
      Width           =   615
   End
   Begin VB.Image imgSearch 
      Height          =   375
      Left            =   2520
      Top             =   5280
      Width           =   375
   End
   Begin VB.Image imgLine 
      Height          =   375
      Left            =   2520
      Top             =   4800
      Width           =   375
   End
   Begin VB.Image imgMenuBoxDown 
      Height          =   255
      Left            =   840
      Top             =   5880
      Width           =   375
   End
   Begin VB.Image imgMenuBox 
      Height          =   255
      Left            =   840
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgRecord 
      Height          =   465
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   465
   End
   Begin VB.Image imgAdd1 
      Height          =   255
      Left            =   120
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image imgAdd2 
      Height          =   255
      Left            =   120
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image imgClear1 
      Height          =   255
      Left            =   480
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image imgClear2 
      Height          =   255
      Left            =   480
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image imgTN 
      Height          =   375
      Left            =   1560
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBottomRight 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Top             =   3480
      Width           =   360
   End
   Begin VB.Image imgBottomMid 
      Height          =   375
      Left            =   1560
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgBottomLeft 
      Height          =   375
      Left            =   1560
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgMidRight 
      Height          =   375
      Left            =   1560
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image imgMidLeft2 
      Height          =   375
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image imgMidLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1560
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgTopRight 
      Height          =   255
      Left            =   1560
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgTopMid 
      Height          =   255
      Left            =   1560
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgTopLeft 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1560
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgAudiogen 
      Height          =   255
      Left            =   120
      Top             =   3720
      Width           =   255
   End
   Begin VB.Image imgStop4 
      Height          =   255
      Left            =   1200
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgStop3 
      Height          =   255
      Left            =   840
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgStop2 
      Height          =   255
      Left            =   480
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgStop1 
      Height          =   255
      Left            =   120
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image imgPlay4 
      Height          =   255
      Left            =   1200
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgPlay3 
      Height          =   255
      Left            =   840
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgPlay2 
      Height          =   255
      Left            =   480
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgPlay1 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image imgExit2 
      Height          =   255
      Left            =   480
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgExit1 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgSplash 
      Height          =   4380
      Left            =   2040
      Top             =   120
      Width           =   6000
   End
   Begin VB.Image imgCDCopyDisabled 
      Height          =   255
      Left            =   1200
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgForward2 
      Height          =   255
      Left            =   480
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image imgForward3 
      Height          =   255
      Left            =   840
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image imgForward1 
      Height          =   255
      Left            =   120
      Top             =   3360
      Width           =   255
   End
   Begin VB.Image imgBack3 
      Height          =   255
      Left            =   840
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image imgBack4 
      Height          =   255
      Left            =   1200
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image imgBack2 
      Height          =   255
      Left            =   480
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image imgBack1 
      Height          =   255
      Left            =   120
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image imgMax2 
      Height          =   255
      Left            =   480
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgMax1 
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgSlider4 
      Height          =   255
      Left            =   1200
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image imgSlider3 
      Height          =   255
      Left            =   840
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image imgSlider2 
      Height          =   255
      Left            =   480
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image imgSlider1 
      Height          =   255
      Left            =   120
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image imgAbort3 
      Height          =   255
      Left            =   840
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image imgAbort2 
      Height          =   255
      Left            =   480
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image imgAbort1 
      Height          =   255
      Left            =   120
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image imgAutoEject2 
      Height          =   255
      Left            =   1200
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image imgAutoEject1 
      Height          =   255
      Left            =   840
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image imgNormalizeTracks2 
      Height          =   255
      Left            =   480
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image imgNormalizeTracks1 
      Height          =   255
      Left            =   120
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image imgBurnDisabled 
      Height          =   255
      Left            =   1200
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgCdCopyDown 
      Height          =   255
      Left            =   840
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgCdCopyOver 
      Height          =   255
      Left            =   480
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgCdCopy 
      Height          =   255
      Left            =   120
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgBurnDown 
      Height          =   255
      Left            =   840
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgBurnOver 
      Height          =   255
      Left            =   480
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgBurn 
      Height          =   255
      Left            =   120
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image imgOpen3 
      Height          =   210
      Left            =   120
      Picture         =   "frmGraphics.frx":06EA
      Top             =   6000
      Width           =   255
   End
   Begin VB.Image imgOpen2 
      Height          =   210
      Left            =   120
      Picture         =   "frmGraphics.frx":076A
      Top             =   5760
      Width           =   255
   End
   Begin VB.Image imgOpen1 
      Height          =   210
      Left            =   120
      Picture         =   "frmGraphics.frx":07EA
      Top             =   5520
      Width           =   255
   End
   Begin VB.Image imgMinimize1 
      Height          =   255
      Left            =   120
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgMinimize2 
      Height          =   255
      Left            =   480
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim lPath As String, msg As String
'lIniFiles.iGFX = ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "GFX", App.Path & "\skins\audiogen1\gfx.ini")
lIniFiles.iGFX = App.Path & "\" & ReadINI(App.Path & "\inis\a_settings.ini", "Settings", "SkinGFX", "")
msg = lIniFiles.iGFX
msg = GetFileTitle(msg)
lPath = left(lIniFiles.iGFX, Len(lIniFiles.iGFX) - Len(msg))
Me.Icon = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Icon", ""))
imgLine.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Line", ""))
imgSplash.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Splash", ""))
imgExit1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Exit1", ""))
imgExit2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Exit2", ""))
imgMinimize1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Minimize1", ""))
imgMinimize2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Minimize2", ""))
imgMax1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Maximize1", ""))
imgMax2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Maximize2", ""))
imgAudiogenLogo.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "AudiogenLogo", ""))
imgPlay1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Play1", ""))
imgPlay2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Play2", ""))
imgPlay3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Play3", ""))
imgPlay4.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Play4", ""))
imgStop1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Stop1", ""))
imgStop2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Stop2", ""))
imgStop3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Stop3", ""))
imgStop4.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Stop4", ""))
imgBurn.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Burn1", ""))
imgBurnOver.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Burn2", ""))
imgBurnDown.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Burn3", ""))
imgBurnDisabled.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Burn4", ""))
imgCdCopy.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "CDCopy1", ""))
imgCdCopyOver.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "CDCopy2", ""))
imgCdCopyDown.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "CDCopy3", ""))
imgCDCopyDisabled.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "CDCopy4", ""))
imgAbort1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Abort1", ""))
imgAbort2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Abort2", ""))
imgAbort3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Abort3", ""))
imgBack1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Back1", ""))
imgBack2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Back2", ""))
imgBack3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Back3", ""))
imgBack4.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Back4", ""))
imgForward1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Forward1", ""))
imgForward2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Forward2", ""))
imgForward3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Forward3", ""))
imgAudiogen.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Audiogen", ""))
imgTopLeft.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "TopLeft", ""))
imgTopMid.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "TopMid", ""))
imgTopRight.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "TopRight", ""))
imgMidLeft.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "MidLeft", ""))
imgMidLeft2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "MidLeft2", ""))
imgMidRight.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "MidRight", ""))
imgBottomLeft.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "BottomLeft", ""))
imgBottomMid.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "BottomMid", ""))
imgBottomRight.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "BottomRight", ""))
imgSlider1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Slider1", ""))
imgSlider2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Slider2", ""))
imgSlider3.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Slider3", ""))
imgSlider4.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Slider4", ""))
imgNormalizeTracks1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Normalize1", ""))
imgNormalizeTracks2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Normalize2", ""))
imgAutoEject1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "AutoEject1", ""))
imgAutoEject2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "AutoEject2", ""))
imgTN.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "TN", ""))
imgAdd1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Add1", ""))
imgAdd2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Add2", ""))
imgClear1.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Clear1", ""))
imgClear2.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Clear2", ""))
imgRecord.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Record", ""))
imgMenuBox.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "MenuBox1", ""))
imgMenuBoxDown.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "MenuBox2", ""))
imgSearch.Picture = LoadPicture(lPath & ReadINI(lIniFiles.iGFX, "Settings", "Search", ""))
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub
