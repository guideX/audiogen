VERSION 5.00
Begin VB.Form frmMoreMenus 
   Caption         =   "More Menus (Hidden)"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenType 
         Caption         =   "Open"
         Begin VB.Menu mnuOpenFile 
            Caption         =   "All Supported Media"
         End
         Begin VB.Menu mnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpenAudioTypes 
            Caption         =   "Audio Types"
         End
         Begin VB.Menu mnuOpenVideoTypes 
            Caption         =   "Video Types"
         End
         Begin VB.Menu mnuSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpenMp3 
            Caption         =   "Mpeg Layer 3 (*.mp3)"
         End
         Begin VB.Menu mnuOpenWav 
            Caption         =   "Wave Audio (*.wav)"
         End
         Begin VB.Menu mnuOpenWMA 
            Caption         =   "Windows Media Audio (*.wma)"
         End
         Begin VB.Menu mnuOpenOGG 
            Caption         =   "Ogg Vorbis (*.ogg)"
         End
         Begin VB.Menu mnuOpenQuickTime 
            Caption         =   "Quick Time (*.qt)"
         End
         Begin VB.Menu mnuOpenSND 
            Caption         =   "Sound Files (*.snd)"
         End
         Begin VB.Menu mnuOpenAU 
            Caption         =   "Au Files (*.au)"
         End
         Begin VB.Menu mnuOpenDat 
            Caption         =   "Unfinished KaZaA Downloads (*.dat)"
         End
         Begin VB.Menu mnuSep8368926392 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpenMpeg 
            Caption         =   "Mpeg Video (*.mpeg;*.mpg)"
         End
         Begin VB.Menu mnuOpenAVI 
            Caption         =   "AVI Video (*.avi)"
         End
         Begin VB.Menu mnuOpenWMV 
            Caption         =   "Windows Media Video (*.wmv;*.wmx;*.wm)"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRemoveFromPlaylist 
         Caption         =   "Remove from Playlist"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSep389093862 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFunctionsAndFiles 
         Caption         =   "Functions and Files"
      End
      Begin VB.Menu mnuVideoPlayer 
         Caption         =   "Video Player"
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "&Play"
   End
End
Attribute VB_Name = "frmMoreMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

