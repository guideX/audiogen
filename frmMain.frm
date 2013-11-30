VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{72C32F06-6CAA-11D2-A800-0000E8545063}#1.0#0"; "ACDWRITE.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7314ED99-8643-4E82-A4F8-5E9F4DEC14BE}#1.0#0"; "VolumeControl.ocx"
Object = "{6701F563-8C33-11D5-A844-0080AE000001}#1.0#0"; "dgpnorm.ocx"
Object = "{A6FC7BFB-24EE-11D7-BEB4-444553540000}#2.0#0"; "WaDec.ocx"
Object = "{FDFCF4A3-AD96-11D4-9959-0050BACD4F4C}#1.0#0"; "MDec.ocx"
Object = "{9F5F61C6-83A0-11D2-A800-00A0CC20D781}#1.0#0"; "ACD.OCX"
Object = "{34B82A63-9874-11D4-9E66-0020780170C6}#1.0#0"; "MEnc.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{350143A3-863D-11D5-A844-0080AE000001}#2.0#0"; "WAEnc.ocx"
Object = "{60819404-3CCE-11D2-A800-008048E89E3E}#1.0#0"; "Effect.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   -30
   ClientWidth     =   8610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "v"
   Visible         =   0   'False
   Begin VB.Frame fraVideo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   480
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraPlaylist 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Job Chooser"
      Height          =   855
      Left            =   7200
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
      Begin MSComctlLib.TreeView tvwPlaylist 
         Height          =   735
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         _Version        =   393217
         Indentation     =   44
         LabelEdit       =   1
         Style           =   1
         ImageList       =   "imgPics"
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmMain.frx":1708A
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdScript 
      Height          =   315
      Left            =   6600
      TabIndex        =   37
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Script"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMain.frx":171EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdSwitch 
      Height          =   315
      Left            =   5640
      TabIndex        =   36
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Switch"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMain.frx":17208
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Search for media within playlist"
      Top             =   1080
      Width           =   300
   End
   Begin VB.ComboBox cboDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   390
   End
   Begin VB.ComboBox cboBitrate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   315
      ItemData        =   "frmMain.frx":17224
      Left            =   3240
      List            =   "frmMain.frx":172A0
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   390
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      ItemData        =   "frmMain.frx":17649
      Left            =   7560
      List            =   "frmMain.frx":1764B
      TabIndex        =   5
      Top             =   720
      Width           =   510
   End
   Begin MSComCtl2.FlatScrollBar fsbChannels 
      Height          =   300
      Left            =   7440
      TabIndex        =   9
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   393216
      Arrows          =   65536
      Min             =   1
      Max             =   4
      Orientation     =   1179649
      Value           =   1
   End
   Begin MSComCtl2.FlatScrollBar fsbVolume 
      Height          =   300
      Left            =   7800
      TabIndex        =   10
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   393216
      Arrows          =   65536
      Max             =   100
      Orientation     =   1179649
   End
   Begin MSComctlLib.ProgressBar prgSpaceLeft 
      Height          =   300
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Max             =   7400
   End
   Begin MSComctlLib.ProgressBar prgPercentDone 
      Height          =   300
      Left            =   4920
      TabIndex        =   13
      Top             =   720
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ListBox lstEvents 
      Height          =   645
      Left            =   120
      TabIndex        =   27
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraControls 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Controls"
      Height          =   1095
      Left            =   4320
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   4815
      Begin EFFECTLib.Effect ctlEffects 
         Height          =   495
         Left            =   0
         TabIndex        =   39
         Top             =   600
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin WAENCLib.WAEnc wmaEnc 
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         Top             =   0
         Width           =   255
         _Version        =   131072
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin MSWinsockLib.Winsock wskUpdate 
         Left            =   3720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskFreeDB 
         Left            =   3240
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrAddtoBurnQueDelay 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   1560
         Top             =   480
      End
      Begin MDECLib.MDec ctlMP3Decode 
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   0
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin MENCLib.MEnc ctlMP3Enc 
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   0
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin VB.Timer tmrResetProcessingVar 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   600
         Top             =   480
      End
      Begin WADECLib.WaDec ctlWMADecode 
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   0
         Width           =   255
         _Version        =   131072
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin MSScriptControlCtl.ScriptControl ctlScript 
         Left            =   2640
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Timer tmrPosition 
         Interval        =   600
         Left            =   1080
         Top             =   480
      End
      Begin VB.Timer tmrCheckEvents 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   600
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   1560
         Top             =   0
      End
      Begin VB.Timer tmrCheckNormalize 
         Interval        =   2000
         Left            =   1080
         Top             =   0
      End
      Begin ACDLib.ACD ctlRipper 
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin ACDWRITELib.ACDWRITE ctlBurn 
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   0
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin MSComctlLib.ImageList imgPics 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1764D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CE3F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22631
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28E93
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E685
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33E77
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3A6D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":51773
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":57FD5
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58427
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58879
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6F913
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":75105
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin DGPNORMLib.DGPNorm ctlNormalize 
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   0
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   0
      End
      Begin VolControl.VolumeControl VolumeControl1 
         Left            =   2040
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         Volume          =   26
      End
   End
   Begin MSComctlLib.TreeView tvwSources 
      Height          =   660
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      _Version        =   393217
      Indentation     =   44
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   5
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgPics"
      Appearance      =   0
   End
   Begin MSComctlLib.TreeView tvwFiles 
      Height          =   660
      Left            =   1320
      TabIndex        =   14
      Top             =   1080
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   5
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgPics"
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.TreeView tvwToBurn 
      Height          =   660
      Left            =   2040
      TabIndex        =   26
      Top             =   1080
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   5
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Edit"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgEdit2 
      Height          =   300
      Left            =   3120
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgEdit 
      Height          =   300
      Left            =   3480
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgSearch 
      Height          =   300
      Left            =   3840
      ToolTipText     =   "Search for media within playlist"
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgProgress 
      Height          =   300
      Left            =   6360
      MouseIcon       =   "frmMain.frx":7A8F7
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&File"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgFileOver 
      Height          =   300
      Left            =   2760
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgFile 
      Height          =   300
      Left            =   2400
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgAdd 
      Height          =   300
      Left            =   4200
      MouseIcon       =   "frmMain.frx":7AA49
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Label lblItem 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   2680
      Width           =   1335
   End
   Begin VB.Label lblSpaceLeft 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6840
      TabIndex        =   33
      Top             =   1950
      Width           =   1695
   End
   Begin VB.Label lblFormat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   5640
      TabIndex        =   30
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Audiogen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      MouseIcon       =   "frmMain.frx":7AB9B
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   1440
      Width           =   675
   End
   Begin VB.Image imgMidRight 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   600
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgMidLeft2 
      Height          =   300
      Left            =   960
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblAliasname 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgTN 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmMain.frx":7ACED
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgProgressChange 
      Height          =   300
      Left            =   6000
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgAudiogen 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7800
      MouseIcon       =   "frmMain.frx":7AE3F
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblChannels 
      BackStyle       =   0  'Transparent
      Caption         =   "Channels:"
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4920
      TabIndex        =   25
      Top             =   1080
      Width           =   300
   End
   Begin VB.Image imgForward 
      Height          =   300
      Left            =   4920
      MouseIcon       =   "frmMain.frx":7AF91
      MousePointer    =   99  'Custom
      ToolTipText     =   "Go forward a track"
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgBack 
      Height          =   300
      Left            =   4560
      MouseIcon       =   "frmMain.frx":7B0E3
      MousePointer    =   99  'Custom
      ToolTipText     =   "Go backward a track"
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblBitrate 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6240
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.Image imgMaximize 
      Height          =   300
      Left            =   7080
      MouseIcon       =   "frmMain.frx":7B235
      MousePointer    =   99  'Custom
      ToolTipText     =   "Maximize"
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgClear 
      Height          =   300
      Left            =   600
      MouseIcon       =   "frmMain.frx":7B387
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgCdCopy 
      Height          =   300
      Left            =   5640
      MouseIcon       =   "frmMain.frx":7B4D9
      MousePointer    =   99  'Custom
      ToolTipText     =   "Copy (rip) cdaudio to your hard drive"
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgBurn 
      Height          =   300
      Left            =   5280
      MouseIcon       =   "frmMain.frx":7B62B
      MousePointer    =   99  'Custom
      ToolTipText     =   "Burn Audio files to a disc in cda format"
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblVolume 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4560
      TabIndex        =   23
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label lblTrackNumber 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4080
      TabIndex        =   22
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label lblTimeDisplay 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image imgStop 
      Height          =   345
      Left            =   4200
      MouseIcon       =   "frmMain.frx":7B77D
      MousePointer    =   99  'Custom
      ToolTipText     =   "Stop playback"
      Top             =   360
      Width           =   345
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   7720
      Width           =   4575
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Left            =   3840
      MouseIcon       =   "frmMain.frx":7B8CF
      MousePointer    =   99  'Custom
      ToolTipText     =   "Play"
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   200
      Left            =   360
      TabIndex        =   19
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label lblTransferStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   5280
      TabIndex        =   18
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label lblTimeLeft 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image imgMinimize 
      Height          =   300
      Left            =   6720
      MouseIcon       =   "frmMain.frx":7BA21
      MousePointer    =   99  'Custom
      ToolTipText     =   "Minimize"
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgClose 
      Height          =   300
      Left            =   7440
      MouseIcon       =   "frmMain.frx":7BB73
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgAutoEject 
      Height          =   300
      Left            =   1680
      MouseIcon       =   "frmMain.frx":7BCC5
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgNormalize 
      Height          =   300
      Left            =   1320
      MouseIcon       =   "frmMain.frx":7BE17
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgBottomLeft 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgProgressYellow 
      Height          =   300
      Left            =   960
      MouseIcon       =   "frmMain.frx":7BF69
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgBottomMid 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgBottomRight 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblKHZ 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgTopRight 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2400
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgTopLeft 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgMidLeft 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3120
      Top             =   360
      Width           =   300
   End
   Begin VB.Image imgTopMid 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DraggedKnot As node
Dim SelectedKnot As node

Public Sub SwitchView()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If fraVideo.Visible = True Then
    fraVideo.Visible = False
ElseIf fraPlaylist.Visible = True Then
    fraPlaylist.Visible = False
Else
    DisplayPlaylist
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub SwitchView()"
End Sub

Public Sub ResizeMain()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If WindowState = vbNormal Then
    lSettings.sMinimized = False
ElseIf WindowState = vbMinimized Then
    Exit Sub
End If
If Height < 8700 Then
    Height = 8700
    Exit Sub
End If
If Width < 8800 Then
    Width = 8800
    Exit Sub
End If
If Len(lScripts.sMain_Resize) <> 0 Then ctlScript.ExecuteStatement lScripts.sMain_Resize
If lPlayer.pStatus = sPlay Then
    PutMultimedia fraVideo.hwnd, lblAliasname.Caption, 0, 0, ScaleX(Val(fraVideo.Width), 1, 3), ScaleX(Val(fraVideo.Height), 1, 3)
    fraVideo.Refresh
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Resize()"
End Sub

Public Sub UpdateImages()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Select Case lPlayer.pStatus
Case sBurn
    If imgStop.Picture <> frmGraphics.imgStop4.Picture Then imgStop.Picture = frmGraphics.imgStop4.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay1.Picture Then imgPlay.Picture = frmGraphics.imgPlay1.Picture
Case sIdle
    If imgStop.Picture <> frmGraphics.imgStop4.Picture Then imgStop.Picture = frmGraphics.imgStop4.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay1.Picture Then imgPlay.Picture = frmGraphics.imgPlay1.Picture
Case sPlay
    If imgStop.Picture <> frmGraphics.imgStop1.Picture Then imgStop.Picture = frmGraphics.imgStop1.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay4.Picture Then imgPlay.Picture = frmGraphics.imgPlay4.Picture
Case sPaused
    If imgStop.Picture <> frmGraphics.imgStop1.Picture Then imgStop.Picture = frmGraphics.imgStop1.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay1.Picture Then imgPlay.Picture = frmGraphics.imgPlay1.Picture
Case sSelectFile
    If imgStop.Picture <> frmGraphics.imgStop4.Picture Then imgStop.Picture = frmGraphics.imgStop4.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay4.Picture Then imgPlay.Picture = frmGraphics.imgPlay1.Picture
Case sDecode
    If imgStop.Picture <> frmGraphics.imgStop4.Picture Then imgStop.Picture = frmGraphics.imgStop4.Picture
    If imgPlay.Picture <> frmGraphics.imgPlay4.Picture Then imgPlay.Picture = frmGraphics.imgPlay4.Picture
End Select
If imgBurn.Picture <> frmGraphics.imgBurnDisabled.Picture Then
    If imgBurn.Picture = frmGraphics.imgAbort1.Picture Or imgBurn.Picture = frmGraphics.imgAbort2.Picture Or imgBurn.Picture = frmGraphics.imgAbort3.Picture Then
        imgBurn.Picture = frmGraphics.imgAbort1.Picture
    Else
        imgBurn.Picture = frmGraphics.imgBurn.Picture
    End If
End If
If frmMain.imgFileOver.Visible = True Then frmMain.imgFileOver.Visible = False
If frmMain.imgEdit2.Visible = True Then frmMain.imgEdit2.Visible = False
If imgProgress.Picture <> frmGraphics.imgSlider3.Picture Then
    imgProgress.Picture = frmGraphics.imgSlider1.Picture
End If
If imgBack.Picture <> frmGraphics.imgBack1.Picture Then
    frmMain.imgBack.Picture = frmGraphics.imgBack1.Picture
End If
If imgForward.Picture <> frmGraphics.imgForward1.Picture Then
    frmMain.imgForward.Picture = frmGraphics.imgForward1.Picture
End If
imgMaximize.Picture = frmGraphics.imgMax1.Picture
imgClose.Picture = frmGraphics.imgExit1.Picture
imgMinimize.Picture = frmGraphics.imgMinimize1.Picture
If imgCdCopy.Picture <> frmGraphics.imgCDCopyDisabled.Picture Then imgCdCopy.Picture = frmGraphics.imgCdCopy.Picture
If lSettings.sCheckTaskbarStatus = True Then
    If lTaskbar = False Then frmTaskbar.SetFocus
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub UpdateImages()"
End Sub

Public Sub ResetFileTreeView()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
tvwFiles.Scroll = False
tvwFiles.Nodes.Clear
tvwFiles.Nodes.Add , , , "Results", 5
tvwFiles.Nodes(1).Expanded = True
tvwFiles.Scroll = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Public Sub ResetFileTreeView()"
End Sub

Private Sub cboAddress_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sSettingAddress = True Then Exit Sub
If LCase(frmMain.cboAddress.Text) = LCase(frmMain.Tag) Then Exit Sub
Dim msg As String, lBase As tSearch, i As Long, msg2 As String
msg = frmMain.cboAddress.Text
If Len(msg) <> 0 Then
    If left(LCase(msg), 3) = "www" Or left(LCase(msg), 7) = "http://" Then
        Surf msg, frmMain.hwnd
        Exit Sub
    End If
    frmMain.ResetFileTreeView
    GetFiles msg, lSettings.sSupportedMedia, vbNormal, lBase
    DoEvents
    If lBase.Count = 0 Then Exit Sub
    For i = 1 To lBase.Count
        If Len(lBase.Path(i)) <> 0 Then
            msg2 = lBase.Path(i)
            If Len(Trim(msg2)) <> 0 Then
                msg2 = lBase.Path(i)
                msg2 = GetFileTitle(msg2)
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2, 7
            End If
        End If
    Next i
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboAddress_Click()"
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckKeyboardCommands KeyCode
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, lBase As tSearch, i As Long, msg2 As String
If KeyAscii = 13 Then
    msg = frmMain.cboAddress.Text
    If Len(msg) <> 0 Then
        If left(LCase(msg), 3) = "www" Or left(LCase(msg), 7) = "http://" Then
            Surf msg, frmMain.hwnd
            Exit Sub
        End If
        frmMain.ResetFileTreeView
        GetFiles msg, lSettings.sSupportedMedia, vbNormal, lBase
        DoEvents
        Pause 0.2
        If lBase.Count = 0 Then Exit Sub
        For i = 1 To lBase.Count
            If Len(lBase.Path(i)) <> 0 Then
                msg2 = lBase.Path(i)
                If Len(Trim(msg2)) <> 0 Then
                    msg2 = lBase.Path(i)
                    msg2 = GetFileTitle(msg2)
                    frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg2, 7
                End If
            End If
        Next i
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboAddress_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub cboBitrate_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndmain_cbobitrate_switch).txt")
lSettings.sBitrate = cboBitrate.ListIndex
WriteINI lIniFiles.iSettings, "Settings", "Bitrate", Int(lSettings.sBitrate)
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboBitrate_Change()"
End Sub

Private Sub cboDrives_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lSettings.sProcessScripts = True Then ProcessScript (App.Path & "\a_script\sub(wndmain_cbodrives_switch).txt")
If imgCdCopy.Picture = frmGraphics.imgCdCopy.Picture Then
    SelectCDDriveByCombo
    WriteINI lIniFiles.iSettings, "Settings", "LastCDDrive", cboDrives.ListIndex
ElseIf frmMain.imgBurn.Picture = frmGraphics.imgBurn.Picture Then
    
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboDrives_Change()"
End Sub

Private Sub cmdScript_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessScript App.Path & "\a_script\sub(wndmain_cmdscript_click).txt"
frmScript.Show
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdScript_Click()"
End Sub

Private Sub cmdSwitch_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessScript App.Path & "\a_script\sub(wndmain_switch_click).txt"
SwitchView
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cmdSwitch_Click()"
End Sub

Private Sub ctlBurn_ActTrack(ByVal Track As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lBurnQue.bTrackIndex = Track
lblTrackNumber.Caption = "Track " & Track
If lBurnQue.bCount <> 0 Then
    prgSpaceLeft.Max = 100
    prgSpaceLeft.value = Track * 100 / lBurnQue.bCount
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_ActTrack(ByVal Track As Integer)"
End Sub

Private Sub ctlBurn_ASPIEvent(ByVal ErrorCode As Integer, ByVal ErrorString As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As VbMsgBoxResult
msg = MsgBox("An ASPI Event error occured. This may mean your ASPI drivers are out of date, or do not exist." & vbCrLf & ErrorString, vbYesNo + vbCritical, "Audiogen")
If msg = vbYes Then
    Shell App.Path & "\aspiupd.exe", vbNormalFocus
    End
ElseIf msg = vbNo Then
    Exit Sub
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_ASPIEvent(ByVal ErrorCode As Integer, ByVal ErrorString As String)"
End Sub

Private Sub ctlBurn_BurningProcessComplete(ByVal Success As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ClearBurnQue
ResetMainForm
PlayWav App.Path & "\audio\a_burncomplete.wav", SND_ASYNC
lblStatus.Caption = "Burn Complete"
Pause 1
AdjustStatus sDoneBurning
lEvents.eProcessing = False
imgBurn.Picture = frmGraphics.imgBurn.Picture
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_BurningProcessComplete(ByVal Success As Long)"
End Sub

Private Sub ctlBurn_Failure(ByVal ErrCode As Long, ByVal ErrorString As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
PlayWav App.Path & "\audio\a_burnfailure.wav", SND_ASYNC
ResetMainForm
lblStatus.Caption = "Burn Failed - " & ErrorString
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_Failure(ByVal ErrCode As Long, ByVal ErrorString As String)"
End Sub

Private Sub ctlBurn_Fixating()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblStatus.Caption = "Finalizing Compact Disc"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_Fixating()"
End Sub

Private Sub ctlBurn_InfoMessage(ByVal MsgCode As Long, ByVal MsgString As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblStatus.Caption = MsgString
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_InfoMessage(ByVal MsgCode As Long, ByVal MsgString As String)"
End Sub

Private Sub ctlBurn_LunNotReady()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblStatus.Caption = "Not ready!"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_LunNotReady()"
End Sub

Private Sub ctlBurn_StatusChange(ByVal status As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lBurnQue.bStatus = status
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_StatusChange(ByVal Status As Integer)"
End Sub

Private Sub ctlBurn_TrackPercent(ByVal percent As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If prgPercentDone.Visible = False Then prgPercentDone.Visible = True
prgPercentDone.value = percent
lblTransferStatus.Caption = "Transfer status: " & percent & "%"
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_TrackPercent(ByVal percent As Integer)"
End Sub

Private Sub ctlBurn_UserStop()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lblStatus.Caption = "User has canceled"
AdjustStatus sIdle
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Burner_UserStop()"
End Sub

Private Sub ctlEffects_OnActionPosition(ByVal ActionPosition As Integer)
lblStatus.Caption = "Effects: " & ActionPosition & "%"
End Sub

Private Sub ctlMP3Decode_PercentDone(ByVal nPercent As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
prgSpaceLeft.value = nPercent
frmMain.lblTimeLeft.Caption = "Decode: " & nPercent & "%"
If nPercent = 100 Then
    ctlMP3Decode.Stop
    AddToFiles lEvents.eCurrentFile, False
    ResetMainForm
    OpenContainingFolder lEvents.eCurrentFile
    Pause 0.2
    tmrResetProcessingVar.Enabled = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub MP3Decode_PercentDone(ByVal nPercent As Long)"
End Sub

Private Sub ctlMP3Enc_PercentDone(ByVal nPercent As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
prgSpaceLeft.value = nPercent
If nPercent = 100 Then
    Dim msg As String
    prgSpaceLeft.Max = 7400
    ResetMainForm False
    lEvents.eProcessing = False
    AddToFiles lEvents.eCurrentFile, False
    OpenContainingFolder lEvents.eCurrentFile
    If lSettings.sConvertingKHZ = True Then
        frmKHZConverter.chkStep1.value = 1
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub MP3Enc_PercentDone(ByVal nPercent As Long)"
End Sub

Private Sub ctlNormalize_PercentDone(ByVal nPercent As Double)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
prgSpaceLeft.value = nPercent
lblStatus.Caption = "Normalize " & nPercent & "%"
If nPercent = 100 Then
    lEvents.eProcessing = False
    ResetMainForm
    OpenContainingFolder lEvents.eCurrentFile
    Exit Sub
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlNormalize_PercentDone(ByVal nPercent As Double)"
End Sub

Private Sub ctlRipper_ActPosition(ByVal Position As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
i = Position / ctlRipper.GetTrackLength(Right(lblTrackNumber.Caption, Len(lblTrackNumber.Caption) - 6)) * 100 / 1.48
'If prgPercentDone.Visible = False Then prgPercentDone.Visible = True
prgPercentDone.value = i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlRipper_ActPosition(ByVal Position As Long)"
End Sub

Private Sub ctlRipper_CopyStart()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'If prgPercentDone.Visible = False Then prgPercentDone.Visible = True
prgPercentDone.value = 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlRipper_CopyStart()"
End Sub

Private Sub ctlRipper_CopyStop()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lEvents.eProcessing = False
prgPercentDone.value = 0
lblTrackNumber.Caption = ""
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlRipper_CopyStop()"
End Sub

Private Sub ctlRipper_Failure(ByVal ErrorCode As Long, ByVal ErrorString As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ResetMainForm
lblStatus.Caption = ErrorString
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlRipper_Failure(ByVal ErrorCode As Long, ByVal ErrorString As String)"
End Sub

Private Sub ctlRipper_ReadComplete()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lPlayer.pStatus = sPlay And lPlayer.pFileType = fCDAudio Then Exit Sub
ResetMainForm
lblStatus.Caption = "Read Complete"
OpenContainingFolder lEvents.eCurrentFile
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlRipper_ReadComplete()"
End Sub

Private Sub ctlScript_Error()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As VbMsgBoxResult
lblStatus.Caption = ctlScript.Error.Description
msg = MsgBox("A Script Error Occured" & vbCrLf & vbCrLf & "File: " & lScripts.sCurrentScript & vbCrLf & "Line: " & ctlScript.Error.Line & vbCrLf & "Description: " & ctlScript.Error.Description & vbCrLf & vbCrLf & "Would you like to edit this file?", vbYesNo + vbCritical, "Audiogen")
If msg = vbYes Then
    frmScript.Show
    frmScript.txtScript.Text = ReadFile(App.Path & "\a_script\" & lScripts.sCurrentScript)
End If
PlayWav App.Path & "\audio\cdremove.wav", SND_ASYNC
If Err.Number <> 0 Then ErrorAid ctlScript.Error.Number, ctlScript.Error.Description, "Private Sub ctlScript_Error()"
End Sub

Private Sub ctlWMADecode_PercentDone(ByVal nPercent As Double)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
prgSpaceLeft.value = nPercent
If nPercent = 100 Then
    lEvents.eProcessing = False
    ResetMainForm False
    OpenContainingFolder lEvents.eCurrentFile
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub ctlWMADecode_PercentDone(ByVal nPercent As Double)"
End Sub

Private Sub Form_GotFocus()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_GotFocus()"
End Sub

Private Sub Form_Load()
lIniFiles.iErrorLog = App.Path & "\inis\a_errorlog.ini"
lSettings.sHandleErrors = ReadINI(lIniFiles.iErrorLog, "Settings", "HandleErrors", True)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String, cCol1 As tSearch, cCol2 As tSearch, j As Integer, cCol3 As tSearch, f As Integer
frmTaskbar.Show
try.cbSize = Len(try)
try.hwnd = Me.hwnd
try.uId = vbNull
try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
try.uCallBackMessage = WM_MOUSEMOVE
try.hIcon = Me.Icon
try.szTip = Trim(App.Title) & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - Team Nexgen - http://www.team-nexgen.com"
Call Shell_NotifyIcon(NIM_ADD, try)
Call Shell_NotifyIcon(NIM_MODIFY, try)
If ReadINI(lIniFiles.iErrorLog, "Settings", "DisplayErrorOptions", True) = True Then
    frmErrorOptions.Show 1
End If
If lSettings.sHandleErrors = False Then MsgBox "WARNING: Error handling has been turned off. This can cause abrupt and abnormal termination of Audiogen at any time during your use of the program. Please be cautious and observant of bugs.", vbCritical, "Error Aid"
CloseAll
If LCase(Trim(Command$)) = "debugmode" Then lSettings.sDebugMode = True
With ctlScript
    .AddObject "wndErrorAid", frmErrorAid, True
    .AddObject "wndScript", frmScript, True
    .AddObject "wndSearch", frmSearch, True
    .AddObject "wndFileMenu", frmFileMenu, True
    .AddObject "wndGraphics", frmGraphics, True
    .AddObject "wndMain", frmMain, True
End With
frmMain.lblAbout.Caption = "Audiogen v1.0 Build " & App.Revision
ctlEffects.Authorize "Leon J Aiossa", "1081841574"
ctlRipper.Authorize "Leon Aiossa", "698070606"
ctlWMADecode.MyKey = "¥DI-SUD+FU_YTFÍ-iguDFJD-SDdefNMRR-SsrfEDS¿"
wmaEnc.MyKey = "§RDEH-YRD_WODJT_PEUFNGO-JSIDFJD-SDOENMRR-SDFSEDS_WDFSF¿"
ctlBurn.Authorize "Leon Aiossa", "517097936"
LoadDrives
LoadSettings
InitRipper
InitCDBurner
VolumeControl1.DeviceToControl = mWave
fsbVolume.value = VolumeControl1.Volume
With tvwSources
    .Nodes.Add , , , "My Documents", 4
    .Nodes.Add , , , "My Music", 5
    .Nodes.Add , , , "Desktop", 2
    DoEvents
    LoadFiles
    .Nodes.Add , , , "Playlist", 5
    .Nodes.Add , , , "Copied CD-Audio", 13
    If lSettings.sDebugMode = True Then .Nodes.Add , , , "Errors", 11
    f = FindTreeViewIndex("Playlist", tvwSources)
    .Nodes.Add , , , "Settings", 4
End With
lScripts.sCurrentScript = "sub(wndmain_startup).txt"
ctlScript.ExecuteStatement ReadFile(App.Path & "\a_script\sub(wndmain_startup).txt")
lScripts.sMain_Init = ReadFile(App.Path & "\a_script\sub(wndmain_initialstate).txt"): DoEvents
lScripts.sMain_Resize = ReadFile(App.Path & "\a_script\sub(wndmain_resize).txt"): DoEvents
lScripts.sCurrentScript = "sub(wndmain_initialstate).txt"
ctlScript.ExecuteStatement lScripts.sMain_Init
lScripts.sCurrentScript = "sub(wndmain_resize).txt"
ctlScript.ExecuteStatement lScripts.sMain_Resize
If lSettings.sFirstRun = True Then
    WriteINI lIniFiles.iSettings, "Settings", "FirstRun", "False"
    QuickPlay App.Path & "\audio\a_intro.mp3"
    DisplayMediaItem App.Path & "\audio\a_intro.mp3", frmMain.tvwFiles
End If
cboDrives.ListIndex = lSettings.sLastCDDrive
AdjustStatus sIdle
If lSettings.sShowSplash = False And lSettings.sShowAboutDetailsOnStartup = True Then
    frmAboutDetails.Show 1
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    CheckMenus
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Form_Resize()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ResizeMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
EndProgram
End Sub

Private Sub fsbChannels_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim r As String
Select Case fsbChannels.value
Case 1
    r = ChannelsControl(frmMain.lblAliasname.Caption, "left", "on")
    r = ChannelsControl(frmMain.lblAliasname.Caption, "right", "on")
    lblStatus.Caption = "Both channels on"
Case 2
    r = ChannelsControl(frmMain.lblAliasname.Caption, "left", "on")
    r = ChannelsControl(frmMain.lblAliasname.Caption, "right", "off")
    lblStatus.Caption = "Left channel only"
Case 3
    r = ChannelsControl(frmMain.lblAliasname.Caption, "left", "off")
    r = ChannelsControl(frmMain.lblAliasname.Caption, "right", "on")
    lblStatus.Caption = "Right channel only"
Case 4
    r = ChannelsControl(frmMain.lblAliasname.Caption, "left", "off")
    r = ChannelsControl(frmMain.lblAliasname.Caption, "right", "off")
    lblStatus.Caption = "All audio channels off"
End Select
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub fsbChannels_Change()"
End Sub

Private Sub fsbVolume_Change()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lblVolume.Caption <> "Vol: " & fsbVolume.value & "%" Then VolumeControl1.Volume = fsbVolume.value
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub fsbVolume_Change()"
End Sub

Private Sub fsbVolume_Scroll()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
VolumeControl1.Volume = fsbVolume.value
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub fsbVolume_Scroll()"
End Sub

Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessScript App.Path & "\a_script\sub(imgadd_mousedown).txt"
If Button = 1 Then
    imgAdd.Picture = frmGraphics.imgAdd2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, msg2 As String
If Button = 1 Then
    imgAdd.Picture = frmGraphics.imgAdd1.Picture
    msg = tvwFiles.SelectedItem.Text
    i = FindTreeViewIndex(msg, tvwToBurn)
    If i <> 0 Then Exit Sub
    If Len(msg) <> 0 Then
        msg2 = tvwToBurn.Nodes.Add(, , , msg)
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgAudiogen_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmAboutDetails.Show 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgAudiogen_DblClick()"
End Sub

Private Sub imgAudiogen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgAudiogen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgAutoEject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If imgAutoEject.Picture = frmGraphics.imgAutoEject1.Picture Then
        imgAutoEject.Picture = frmGraphics.imgAutoEject2
        WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "True"
        lSettings.sAutoEject = True
        If lSettings.sTreeviewSource = "Auto Eject" Then
            ResetFileTreeView
            tvwFiles.Nodes.Add 1, tvwChild, , "True"
        End If
    Else
        imgAutoEject.Picture = frmGraphics.imgAutoEject1
        WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "False"
        lSettings.sAutoEject = False
        If lSettings.sTreeviewSource = "Auto Eject" Then
            ResetFileTreeView
            tvwFiles.Nodes.Add 1, tvwChild, , "False"
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgAutoEject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    imgBack.Picture = frmGraphics.imgBack3.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Button = 0 Then
    imgBack.Picture = frmGraphics.imgBack2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    ProcessScript App.Path & "\a_script\sub(wndmain_imgback_mouseup).txt"
    GoBackward
    imgBack.Picture = frmGraphics.imgBack1.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If imgBurn.Picture = frmGraphics.imgAbort1.Picture Or imgBurn.Picture = frmGraphics.imgAbort2.Picture Or imgBurn.Picture = frmGraphics.imgAbort3.Picture Then
    If Button = 1 Then
        imgBurn.Picture = frmGraphics.imgAbort2.Picture
    End If
ElseIf imgBurn.Picture = frmGraphics.imgBurn.Picture Or imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture Or imgBurn.Picture = frmGraphics.imgBurnOver.Picture Then
    If Button = 1 Then
        If imgBurn.Picture <> frmGraphics.imgBurnDisabled.Picture Then
            imgBurn.Picture = frmGraphics.imgBurnDown.Picture
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If imgBurn.Picture = frmGraphics.imgAbort1.Picture Or imgBurn.Picture = frmGraphics.imgAbort2.Picture Or imgBurn.Picture = frmGraphics.imgAbort3.Picture Then
    If Button = 0 Then
        imgBurn.Picture = frmGraphics.imgAbort3.Picture
    End If
ElseIf imgBurn.Picture = frmGraphics.imgBurn.Picture Or imgBurn.Picture = frmGraphics.imgBurnDisabled.Picture Or imgBurn.Picture = frmGraphics.imgBurnOver.Picture Then
    If Button = 0 Then
        If imgBurn.Picture <> frmGraphics.imgBurnDisabled.Picture Then
            imgBurn.Picture = frmGraphics.imgBurnOver.Picture
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgBurn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If imgBurn.Picture = frmGraphics.imgAbort1.Picture Or imgBurn.Picture = frmGraphics.imgAbort2.Picture Or imgBurn.Picture = frmGraphics.imgAbort3.Picture Then
    ctlBurn.Abort
ElseIf imgBurn.Picture = frmGraphics.imgBurn.Picture Or imgBurn.Picture = frmGraphics.imgBurnDown.Picture Or imgBurn.Picture = frmGraphics.imgBurnOver.Picture Then
    If Button = 1 Then
        If imgBurn.Picture <> frmGraphics.imgBurnDisabled.Picture Then
            imgBurn.Picture = frmGraphics.imgBurn.Picture
            AdjustStatus sBurn
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgBurn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgCdCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If imgCdCopy.Picture <> frmGraphics.imgCDCopyDisabled.Picture Then
        imgCdCopy.Picture = frmGraphics.imgCdCopyDown.Picture
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgCdCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgCdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    If imgCdCopy.Picture <> frmGraphics.imgCDCopyDisabled.Picture Then
        imgCdCopy.Picture = frmGraphics.imgCdCopyOver.Picture
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgCdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgCdCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If imgCdCopy.Picture <> frmGraphics.imgCDCopyDisabled.Picture Then
        ProcessScript App.Path & "\a_script\sub(wndmain_cdcopy_mouseup).txt"
        imgCdCopy.Picture = frmGraphics.imgCdCopy.Picture
        StartRip
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgCdCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    imgClear.Picture = frmGraphics.imgClear2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If Button = 1 Then
    frmMain.tvwToBurn.Nodes.Clear
    imgClear.Picture = frmGraphics.imgClear1.Picture
    ClearBurnQue
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
EndProgram
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
imgClose.Picture = frmGraphics.imgExit2.Picture
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    imgForward.Picture = frmGraphics.imgForward3.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Button = 0 Then
    imgForward.Picture = frmGraphics.imgForward2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgForward_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    imgForward.Picture = frmGraphics.imgForward1.Picture
    GoForward
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMaximize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMaximized
End If
imgMaximize.Picture = frmGraphics.imgMax1.Picture
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMaximize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMaximize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    imgMaximize.Picture = frmGraphics.imgMax2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMaximize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMidLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMidLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMidLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMidLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMidRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMidRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
imgMinimize.Picture = frmGraphics.imgMinimize2.Picture
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
frmTaskbar.WindowState = vbMinimized
Me.WindowState = vbMinimized
lSettings.sMinimized = True
imgMinimize.Picture = frmGraphics.imgMinimize1.Picture
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgNormalize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If imgNormalize.Picture = frmGraphics.imgNormalizeTracks1.Picture Then
        imgNormalize.Picture = frmGraphics.imgNormalizeTracks2
        lSettings.sNormalize = True
        WriteINI lIniFiles.iSettings, "Settings", "Normalize", "True"
    Else
        imgNormalize.Picture = frmGraphics.imgNormalizeTracks1
        lSettings.sNormalize = False
        WriteINI lIniFiles.iSettings, "Settings", "Normalize", "False"
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgNormalize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lPlayer.pStatus <> sPlay Then imgPlay.Picture = frmGraphics.imgPlay2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    If lPlayer.pStatus <> sPlay Then imgPlay.Picture = frmGraphics.imgPlay3.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lPlayer.pStatus <> sPlay And lPlayer.pStatus <> sBurn Then
        If Len(lTag.tFile) <> 0 Then
            ProcessScript App.Path & "\a_script\sub(wndmain_imgplay_mouseup).txt"
            QuickPlay lTag.tFile
        End If
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgProgressChange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lPlayer.pStatus <> sPlay Then
    FormDrag Me
    Exit Sub
End If
If Button = 1 Then
    If imgProgress.Picture = frmGraphics.imgSlider3.Picture Then Exit Sub
    lProgressClicked = True
    imgProgress.Picture = frmGraphics.imgSlider2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgProgressChange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgProgressChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If imgProgress.Picture = frmGraphics.imgSlider3.Picture Then Exit Sub
If Button = 0 Then
    imgProgress.Picture = frmGraphics.imgSlider4.Picture
ElseIf Button = 1 Then
    i = X + 2000
    If i < imgProgressChange.Width + 2000 And i > 2000 Then
        imgProgress.left = i
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgProgressChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgProgressChange_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long, f As Long
If imgProgress.Picture = frmGraphics.imgSlider3.Picture Then Exit Sub
If Button = 1 Then
    imgProgress.Picture = frmGraphics.imgSlider1.Picture
    If lPlayer.pStatus = sPlay Then
        Select Case lPlayer.pFileType
        Case fAllFileTypes
            i = (X * GetTotalframes(frmMain.lblAliasname.Caption) / imgProgressChange.Width)
            frmMain.lblStatus.Caption = MoveMultimedia(frmMain.lblAliasname.Caption, i)
        End Select
    Else
        FormDrag Me
    End If
    lProgressClicked = False
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgProgressChange_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgProgressYellow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    lProgressClicked = True
    imgProgress.Picture = frmGraphics.imgSlider2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgProgressYellow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgProgressYellow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Long
If Button = 0 Then
    imgProgress.Picture = frmGraphics.imgSlider4.Picture
ElseIf Button = 1 Then
    i = X + 2000
    If i < imgProgressChange.Width + 2000 And i > 2000 Then
        imgProgress.left = i
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgProgressYellow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lPlayer.pStatus <> sIdle Then imgStop.Picture = frmGraphics.imgStop2.Picture
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If imgStop.Picture <> frmGraphics.imgStop4.Picture Then
    UpdateImages
    If Button = 0 Then
        If lPlayer.pStatus <> sIdle Then imgStop.Picture = frmGraphics.imgStop3.Picture
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then
    If lPlayer.pStatus <> sIdle Then
        ProcessScript App.Path & "\a_script\sub(wndmain_stop_mouseup).txt"
        AdjustStatus sStop
        AdjustStatus sIdle
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 1 Then FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub imgTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
frmAboutDetails.Show 1
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblBitrate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
CheckMenus
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblBitrate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblChannels_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblChannels_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
frmMenuEdit.top = frmMain.top + lblFile.top + lblFile.Height + 50
frmMenuEdit.left = frmMain.left + lblFile.left + 540
frmMenuEdit.Visible = True
imgEdit.Visible = True
imgEdit2.Visible = False
imgFile.Visible = False
imgFileOver.Visible = False
lMenus.mEditMenuVisible = True
lMenus.mFileMenuVisible = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    imgFile.Visible = False
    imgFileOver.Visible = False
    If lMenus.mFileMenuVisible = True Then
        CheckMenus
        frmMenuEdit.top = frmMain.top + lblFile.top + lblFile.Height + 50
        frmMenuEdit.left = frmMain.left + lblFile.left + 540
        frmFileMenu.Visible = False
        imgEdit2.Visible = False
        imgEdit.Visible = True
        lMenus.mEditMenuVisible = True
        lMenus.mFileMenuVisible = False
    End If
    If lMenus.mEditMenuVisible = False Then
        imgEdit2.Visible = True
        imgEdit.Visible = False
    Else
        imgEdit2.Visible = False
        imgEdit.Visible = True
    End If
Else
    imgEdit.Visible = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
frmFileMenu.top = frmMain.top + lblFile.top + lblFile.Height + 50
frmFileMenu.left = frmMain.left + lblFile.left + 90
frmFileMenu.Visible = True
imgFileOver.Visible = False
imgFile.Visible = True
imgEdit.Visible = False
imgEdit2.Visible = False
lMenus.mEditMenuVisible = False
lMenus.mFileMenuVisible = True
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If Button = 0 Then
    imgEdit.Visible = False
    imgEdit2.Visible = False
    If lMenus.mEditMenuVisible = True Then
        CheckMenus
        frmFileMenu.top = frmMain.top + lblFile.top + lblFile.Height + 50
        frmFileMenu.left = frmMain.left + lblFile.left + 90
        frmFileMenu.Visible = True
        imgFileOver.Visible = False
        imgFile.Visible = True
        lMenus.mEditMenuVisible = False
        lMenus.mFileMenuVisible = True
    End If
    lMenus.mEditMenuVisible = False
    If lMenus.mFileMenuVisible = False Then
        imgFileOver.Visible = True
        imgFile.Visible = False
    Else
        imgFileOver.Visible = False
        imgFile.Visible = True
    End If
Else
    imgFileOver.Visible = True
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblFormat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblFormat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblSpaceLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblSpaceLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblTimeDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblTimeDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblTimeLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblTimeLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblTrackNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Button = 1 Then
    FormDrag Me
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblTrackNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblTransferStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblTransferStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub lblVolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
FormDrag Me
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub lblVolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub prgPercentDone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub prgPercentDone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub prgSpaceLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckMenus
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub prgSpaceLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub Timer1_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
lstEvents.Clear
lstEvents.AddItem lEvents.eProcessing
For i = 1 To lEvents.eCount
    lstEvents.AddItem lEvents.eEvent(i).eType & " - " & lEvents.eEvent(i).eInputFile & " - " & lEvents.eEvent(i).eOutputFile
Next i
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub Timer1_Timer()"
End Sub

Private Sub tmrAddtoBurnQueDelay_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
tmrAddtoBurnQueDelay.Enabled = False
i = FindFileIndexByFilename(frmMain.tvwToBurn.SelectedItem.Text)
If i <> 0 Then
    msg = lFiles.fFile(i).fFilename
    msg = GetFileTitle(msg)
    AddToBurnQue msg, left(lFiles.fFile(i).fFilename, Len(lFiles.fFile(i).fFilename) - Len(msg))
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrAddtoBurnQueDelay_Timer()"
End Sub

Private Sub tmrCheckEvents_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If lEvents.eProcessing = False And lEvents.eCount <> 0 Then
    ProcessNextEvent
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrCheckEvents_Timer()"
End Sub

Private Sub tmrCheckNormalize_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If prgSpaceLeft.value = 0 Then
    If Len(lBurnQue.bNormInFile) <> 0 Then
        msg = lBurnQue.bNormInFile
        msg = GetFileTitle(msg)
        lblStatus.Caption = "Normalize " & msg
        lblTimeLeft.Caption = "Normalize File:"
        Pause 2
        ctlNormalize.Normalize lBurnQue.bNormInFile, lBurnQue.bNormOutFile, 95
    End If
Else
    tmrCheckNormalize.Enabled = False
    Exit Sub
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrCheckNormalize_Timer()"
End Sub

Private Sub tmrResetProcessingVar_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
lEvents.eProcessing = False
tmrResetProcessingVar.Enabled = False
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrResetProcessingVar_Timer()"
End Sub

Private Sub tmrPosition_Timer()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, p As Long
If lPlayer.pStatus = sPlay Then
    Select Case lPlayer.pFileType
    Case fAllFileTypes
        p = GetPercent(lblAliasname.Caption)
        If p <> -1 Then
            SetProgress Str(p)
            If p = 100 Then
                AdjustStatus sStop
                AdjustStatus sIdle
                tmrPosition.Enabled = False
                If fraVideo.Visible = True Then SwitchView
            ElseIf p > 101 Then
                AdjustStatus sStop
                AdjustStatus sIdle
                tmrPosition.Enabled = False
                If fraVideo.Visible = True Then SwitchView
            Else
                If Right(LCase(lPlayer.pFilename), 4) = ".mp3" Or Right(LCase(lPlayer.pFilename), 4) = ".wav" Or Right(LCase(lPlayer.pFilename), 4) = ".wma" Or Right(LCase(lPlayer.pFilename), 4) = ".snd" Or Right(LCase(lPlayer.pFilename), 3) = ".au" Then
                    Dim m As String
                    lblTimeDisplay.Caption = Format(GetCurrentMultimediaPos(frmMain.lblAliasname.Caption) / Val(GetFramesPerSecond(frmMain.lblAliasname.Caption)), "00:00") & " / " & Format(GetTotalTimeByMS(frmMain.lblAliasname.Caption) / 1000, "00:00")
                    lblStatus.Caption = "Play: " & GetPercent(frmMain.lblAliasname.Caption) & "%"
                    m = lPlayer.pFilename
                    m = GetFileTitle(m)
                    frmTaskbar.Caption = "Audiogen - " & m & " - " & lblStatus.Caption
                Else
                    lblTimeDisplay.Caption = GetPercent(frmMain.lblAliasname.Caption) & "%"
                End If
            End If
        End If
    End Select
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tmrPosition_Timer()"
End Sub

Private Sub tvwPlaylist_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
i = FindFileIndexByFilename(tvwPlaylist.SelectedItem.Text)
If i <> 0 Then
    fraPlaylist.Visible = False
    frmMain.Refresh
    Pause 0.2
    QuickPlay lFiles.fFile(i).fFilename
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwPlaylist_DblClick()"
End Sub

Private Sub tvwSources_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwSources_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub tvwSources_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, i As Integer, mbox As VbMsgBoxResult
CheckMenus
lSettings.sTreeviewSource = tvwSources.SelectedItem.Text
If Button = 1 Then
    frmMain.lblItem.Caption = tvwSources.SelectedItem.Text
    Select Case LCase(frmMain.tvwSources.SelectedItem.Parent.Text)
    Case "settings"
        frmMain.ResetFileTreeView
        With frmMain.tvwSources
            Select Case LCase(.SelectedItem.Text)
            Case "always on top"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sAlwaysOnTop))
            Case "check taskbar status"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sCheckTaskbarStatus))
            Case "full screen video"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sFullScreenVideo))
            Case "auto eject"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sAutoEject))
            Case "convert khz"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sConvertKHZ))
            Case "debug mode"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sDebugMode))
            Case "finalize disc"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sFinalize))
            Case "handle errors"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sHandleErrors))
            Case "name"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(lSettings.sName)
            Case "normalize"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sNormalize))
            Case "password"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(lSettings.sPassword)
            Case "show splash"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sShowSplash))
            Case "supported media"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(lSettings.sSupportedMedia)
            Case "test mode"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sTestMode))
            Case "process scripts"
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , Trim(Str(lSettings.sProcessScripts))
            End Select
        End With
        Exit Sub
    End Select
    If tvwSources.SelectedItem.Expanded = True Then
        Pause 0.1
        tvwSources.SelectedItem.Expanded = False
        'If tvwSources.SelectedItem.Parent <> Nothing Then
        'If LCase(tvwSources.SelectedItem.Parent.Text) = "playlist" Then
            'If tvwSources.SelectedItem.Text = "Audio" Then
                '*.snd;*.mpa;*.enc;*.m1v;*.mp2;*.mp3;*.mpe;*.mpm;*.au;*.snd;*.aif;*.aiff;*.aifc;*.wav;*.wma;
                '*.qt;*.mov;*.dat;*.mpg;*.mpv;*.mpeg;*.wmv;*.avi;
            'ElseIf tvwSources.SelectedItem.Text = "Video" Then
            
            'End If
        'End If
        If tvwSources.SelectedItem.Text <> "Playlist" Then
            DisplayDirectory DecodeLocation(frmMain.tvwSources.SelectedItem.FullPath, frmMain.tvwSources.SelectedItem.FullPath)
        End If
    Else
        If tvwSources.SelectedItem.Children = 0 Then
            If HasLocation(tvwSources.SelectedItem.Text) = True Then
                DisplayDirectory DecodeLocation(frmMain.tvwSources.SelectedItem.FullPath, frmMain.tvwSources.SelectedItem.FullPath)
            Else
                DisplayTreeviewFunction frmMain.tvwSources.SelectedItem.Text
            End If
            Pause 0.1
            tvwSources.SelectedItem.Expanded = True
        Else
            Pause 0.1
            tvwSources.SelectedItem.Expanded = True
            If tvwSources.SelectedItem.Text = "Playlist" Then
                DisplayTreeviewFunction frmMain.tvwSources.SelectedItem.Text
            End If
        End If
    End If
Else
    If HasLocation(frmMain.tvwSources.SelectedItem.Text) = True Then
        Dim cCol1 As tSearch
        GetSubFiles DecodeLocation(frmMain.tvwSources.SelectedItem.Text, frmMain.tvwSources.SelectedItem.FullPath), lSettings.sSupportedMedia, vbDirectory, vbNormal, cCol1: DoEvents
        frmMain.ResetFileTreeView
        For i = 1 To cCol1.Count
            If Len(cCol1.Path(i)) And DoesFileExist(cCol1.Path(i)) <> 0 Then
                msg = cCol1.Path(i)
                msg = GetFileTitle(msg)
                frmMain.tvwFiles.Nodes.Add 1, tvwChild, , msg, 7
                AddToFiles cCol1.Path(i), False
            End If
        Next i
        Unload frmSearching
    Else
        Select Case LCase(frmMain.tvwSources.SelectedItem.Text)
        Case "errors"
            Surf lIniFiles.iErrorLog, frmMain.hwnd
        Case "playlist"
            SwitchView
        End Select
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwSources_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub


Private Sub tvwFiles_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim mbool As String
Select Case LCase(tvwSources.SelectedItem.Parent.Text)
Case "settings"
    mbool = Trim(frmMain.tvwFiles.SelectedItem.Text)
    Select Case LCase(tvwSources.SelectedItem.Text)
    Case "full screen video"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sFullScreenVideo = False
            WriteINI lIniFiles.iSettings, "Settings", "FullScreenVideo", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sFullScreenVideo = True
            WriteINI lIniFiles.iSettings, "Settings", "FullScreenVideo", "True"
        End If
    Case "auto eject"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sAutoEject = False
            WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "False"
            imgAutoEject.Picture = frmGraphics.imgAutoEject1.Picture
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sAutoEject = True
            WriteINI lIniFiles.iSettings, "Settings", "AutoEject", "True"
            imgAutoEject.Picture = frmGraphics.imgAutoEject2.Picture
        End If
    Case "always on top"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            AlwaysOnTop frmMain, False
            lSettings.sAlwaysOnTop = False
            WriteINI lIniFiles.iSettings, "Settings", "AlwaysOnTop", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            AlwaysOnTop frmMain, True
            lSettings.sAlwaysOnTop = True
            WriteINI lIniFiles.iSettings, "Settings", "AlwaysOnTop", "True"
        End If
    Case "check taskbar status"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sCheckTaskbarStatus = False
            WriteINI lIniFiles.iSettings, "Settings", "CheckTaskbarStatus", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sCheckTaskbarStatus = True
            WriteINI lIniFiles.iSettings, "Settings", "CheckTaskbarStatus", "True"
        End If
    Case "convert khz"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sConvertKHZ = False
            WriteINI lIniFiles.iSettings, "Settings", "ConvertKHZ", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sConvertKHZ = True
            WriteINI lIniFiles.iSettings, "Settings", "ConvertKHZ", "True"
        End If
    Case "debug mode"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sDebugMode = False
            WriteINI lIniFiles.iSettings, "Settings", "DebugMode", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sDebugMode = True
            WriteINI lIniFiles.iSettings, "Settings", "DebugMode", "True"
        End If
    Case "finalize disc"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sFinalize = False
            WriteINI lIniFiles.iSettings, "Settings", "Finalize", "False"
            frmMain.ctlBurn.Finalize = False
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sFinalize = True
            WriteINI lIniFiles.iSettings, "Settings", "Finalize", "True"
            frmMain.ctlBurn.Finalize = True
        End If
    Case "handle errors"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sHandleErrors = False
            WriteINI lIniFiles.iErrorLog, "Settings", "Finalize", "False"
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sHandleErrors = True
            WriteINI lIniFiles.iErrorLog, "Settings", "Finalize", "True"
        End If
    Case "name"
        frmRegister.Show 1
    Case "normalize"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sNormalize = False
            WriteINI lIniFiles.iSettings, "Settings", "Normalize", "False"
            frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks1.Picture
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sNormalize = True
            WriteINI lIniFiles.iSettings, "Settings", "Normalize", "True"
            frmMain.imgNormalize.Picture = frmGraphics.imgNormalizeTracks2.Picture
        End If
    Case "password"
        frmRegister.Show 1
    Case "show splash"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            lSettings.sShowSplash = False
            WriteINI lIniFiles.iSettings, "Settings", "ShowSplash", False
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            lSettings.sShowSplash = True
            WriteINI lIniFiles.iSettings, "Settings", "ShowSplash", True
        End If
    Case "supported media"
        lSettings.sSupportedMedia = InputBox("Enter supported media (Use with caution):", "Supported media", lSettings.sSupportedMedia)
        WriteINI lIniFiles.iSettings, "Settings", "SupportedMedia", lSettings.sSupportedMedia
    Case "test mode"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            WriteINI lIniFiles.iSettings, "Settings", "TestMode", "False"
            lSettings.sTestMode = False
            frmMain.ctlBurn.TestMode = False
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            WriteINI lIniFiles.iSettings, "Settings", "TestMode", "True"
            lSettings.sTestMode = True
            frmMain.ctlBurn.TestMode = True
        End If
    Case "process scripts"
        ResetFileTreeView
        If mbool = "True" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "False"
            WriteINI lIniFiles.iSettings, "Settings", "ProcessScripts", "False"
            lSettings.sProcessScripts = False
        ElseIf mbool = "False" Then
            frmMain.tvwFiles.Nodes.Add 1, tvwChild, , "True"
            WriteINI lIniFiles.iSettings, "Settings", "ProcessScripts", "True"
            lSettings.sTestMode = True
        End If
    End Select
End Select
DisplayMediaItem lFiles.fFile(FindFileIndexByFilename(tvwFiles.SelectedItem.Text)).fFilename, tvwFiles
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwFiles_Click()"
End Sub

Private Sub tvwFiles_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
msg = tvwFiles.SelectedItem.Text
If Len(msg) <> 0 Then
    If lPlayer.pStatus = sBurn Then
        MsgBox "You can't play a video or audio file while your burning audio", vbOKCancel + vbExclamation
        Exit Sub
    End If
    i = FindFileIndexByFilename(msg)
    If Len(lFiles.fFile(i).fFilename) <> 0 Then
        QuickPlay lFiles.fFile(i).fFilename
    Else
        If Right(LCase(msg), 4) = ".cda" Then QuickPlay msg
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwFiles_DblClick()"
End Sub

Private Sub tvwFiles_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
If KeyAscii <> 0 Then
    Select Case KeyAscii
    Case 13
        If Len(tvwFiles.SelectedItem.Text) <> 0 Then
            QuickPlay lFiles.fFile(FindFileIndexByFilename(tvwFiles.SelectedItem.Text)).fFilename
        End If
    End Select
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwFiles_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub tvwFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim knot As node
Set knot = tvwFiles.HitTest(X, Y)
If Not (knot Is Nothing) Then
    knot.Selected = True
    Set DraggedKnot = knot
End If
If Button = 2 Then
    lblStatus.Caption = tvwFiles.SelectedItem.Text
    If Err.Number <> 0 Then Exit Sub
    If LCase(lblStatus.Caption) = "results" Then Exit Sub
    tvwFiles.Nodes.Item(tvwFiles.Index).Selected = True
    Pause 0.2
    DoEvents
    PopupMenu frmMenus.mnuFiles
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)"
End Sub

Private Sub tvwFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
'UpdateImages
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub tvwToBurn_Click()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
DisplayMediaItem lFiles.fFile(FindFileIndexByFilename(tvwToBurn.SelectedItem.Text)).fFilename, tvwToBurn
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwToBurn_Click()", True, True
End Sub

Private Sub tvwToBurn_DblClick()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, msg As String
If lPlayer.pStatus = sPlay Then
    AdjustStatus sStop
    DoEvents
End If
msg = tvwToBurn.SelectedItem.Text
i = FindFileIndexByFilename(msg)
If i <> 0 Then QuickPlay lFiles.fFile(i).fFilename
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwToBurn_DblClick()"
End Sub

Private Sub tvwToBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
CheckMenus
If Button = 2 Then
    msg = tvwToBurn.SelectedItem.Text
    If Len(msg) <> 0 Then
        PopupMenu frmMenus.mnuQue
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwToBurn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", True, False
End Sub

Private Sub tvwToBurn_NodeCheck(ByVal node As MSComctlLib.node)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer, j As String
lblFormat.Caption = ""
lblBitrate.Caption = ""
lblTimeDisplay.Caption = ""
lblTitle.Caption = ""
CheckMenus
If Len(tvwToBurn.SelectedItem.Text) <> 0 Then
    i = FindFileIndexByFilename(node.Text)
    lblFormat.Caption = Right(LCase(lFiles.fFile(i).fFilename), 3)
    lTag.tFile = lFiles.fFile(i).fFilename
    If Right(LCase(lFiles.fFile(i).fFilename), 4) = ".mp3" Then
        GetTagInfo
        GetMP3Info
        DoEvents
        lblKHZ.Caption = lTag.tFreqChan
        lblBitrate.Caption = lTag.tBitrate & " kbps"
        If Len(lTag.tLength) <> 0 Then
            j = lTag.tLength / 0.6
            If Len(Trim(lTag.tTitle)) <> 0 Then
                lblTitle.Caption = lTag.tTitle
                lblTimeDisplay.Caption = Format(j, "##:##")
            Else
                lblTitle.Caption = node.Text
                lblTimeDisplay.Caption = Format(j, "##:##")
            End If
        End If
    End If
    If Len(lblTitle.Caption) = 0 Then lblTitle.Caption = left(node.Text, Len(node.Text) - 4)
End If
Pause 0.2
node.Selected = True
NodeCheck node
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwToBurn_NodeCheck(ByVal node As MSComctlLib.node)"
End Sub

Private Sub tvwToBurn_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim i As Integer
If LCase(DraggedKnot.Text) = "results" Then
    For i = 1 To tvwFiles.Nodes.Count
        If Len(tvwFiles.Nodes(i).Text) <> 0 And LCase(tvwFiles.Nodes(i).Text) <> "results" Then
            tvwToBurn.Nodes.Add , , , tvwFiles.Nodes(i).Text
        End If
    Next i
End If
If FindTreeViewIndex(DraggedKnot.Text, tvwToBurn) <> 0 Then Exit Sub
If Not (SelectedKnot Is Nothing) Then
    If Right(LCase(DraggedKnot.Text), 4) <> ".cda" Then
        i = FindFileIndexByFilename(DraggedKnot.Text)
        If i <> 0 Then
            If Right(LCase(DraggedKnot.Text), 4) = ".mp3" Or Right(LCase(DraggedKnot.Text), 4) = ".wma" Or Right(LCase(DraggedKnot.Text), 4) = ".wav" Then
                tvwToBurn.Nodes.Add SelectedKnot.Parent, tvwChild, DraggedKnot.Key, DraggedKnot.Text
            End If
        End If
    Else
        tvwToBurn.Nodes.Add SelectedKnot.Parent, tvwChild, DraggedKnot.Key, DraggedKnot.Text
    End If
Else
    If Right(LCase(DraggedKnot.Text), 4) <> ".cda" Then
        i = FindFileIndexByFilename(DraggedKnot.Text)
        If i <> 0 Then
            If Right(LCase(DraggedKnot.Text), 4) = ".mp3" Or Right(LCase(DraggedKnot.Text), 4) = ".wma" Or Right(LCase(DraggedKnot.Text), 4) = ".wav" Then
                tvwToBurn.Nodes.Add , tvwNext, DraggedKnot.Key, DraggedKnot.Text
            End If
        End If
    Else
        tvwToBurn.Nodes.Add , tvwNext, DraggedKnot.Key, DraggedKnot.Text
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub tvwToBurn_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CheckKeyboardCommands KeyCode
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String
If KeyAscii = 13 Then
    ProcessScript App.Path & "\a_script\sub(wndmain_txtsearch_keypress).txt"
    KeyAscii = 0
    msg = txtSearch.Text
    If Len(msg) <> 0 Then
        Dim i As Long, msg2 As String, f As Integer, cCol1 As tSearch
        ResetFileTreeView
        For i = 0 To lFiles.fCount
            msg2 = lFiles.fFile(i).fFilename
            msg2 = GetFileTitle(msg2)
            If Len(msg2) <> 0 Then
                If InStr(1, msg2, msg, vbTextCompare) Then
                    tvwFiles.Nodes.Add 1, tvwChild, , msg2, 7
                End If
            End If
        Next i
    Else
        MsgBox "You must input a search string", vbExclamation, "Error"
        Exit Sub
    End If
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub txtSearch_KeyPress(KeyAscii As Integer)"
End Sub

Private Sub wmaEnc_PercentDone(ByVal nPercent As Double)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
prgSpaceLeft.value = nPercent
If nPercent = 100 Then
    lEvents.eProcessing = False
    ResetMainForm False
    OpenContainingFolder lEvents.eCurrentFile
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wmaEnc_PercentDone(ByVal nPercent As Double)"
End Sub

Private Sub wskFreeDB_Close()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
CloseFreeDB wskFreeDB
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskFreeDB_Close()"
End Sub

Private Sub wskFreeDB_Connect()
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ConnectFreeDB wskFreeDB
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskFreeDB_Connect()"
End Sub

Private Sub wskFreeDB_DataArrival(ByVal bytesTotal As Long)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ProcessFreeDBData wskFreeDB
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskFreeDB_DataArrival(ByVal bytesTotal As Long)"
End Sub

Private Sub wskFreeDB_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
ErrorAid Str(Number), Description, "Private Sub wskFreeDB_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)"
End Sub

Private Sub wskUpdate_Close()
On Local Error Resume Next
Dim msg As String, msg2 As String
If Len(lSettings.sLatestVersion) <> 0 Then
    wskUpdate.Close: wskUpdate.Tag = "CLOSED"
    If lSettings.sLatestVersion <> App.Major & "." & App.Minor Then
        'update available
    End If
Else
    wskUpdate.Close: wskUpdate.Tag = "CLOSED"
End If
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskUpdate_Close()"
End Sub

Private Sub wskUpdate_Connect()
On Local Error Resume Next
Dim getString As String, ShortWebSite As String
wskUpdate.Tag = "OPEN"
ShortWebSite = "http://www.team-nexgen.com/agupdate.ini"
getString = "GET " + ShortWebSite + " HTTP/1.0" + vbCrLf
getString = getString + "Accept: */*" + vbCrLf
getString = getString + "Accept: text/html" + vbCrLf
getString = getString + vbCrLf
wskUpdate.SendData getString
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskUpdate_Connect()"
End Sub

Private Sub wskUpdate_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Dim Buffer As String
If wskUpdate.Tag = "OPEN" Then wskUpdate.GetData Buffer
lSettings.sLatestVersion = lSettings.sLatestVersion & Buffer
If Err.Number <> 0 Then ErrorAid Err.Number, Err.Description, "Private Sub wskUpdate_DataArrival(ByVal bytesTotal As Long)"
End Sub
