Attribute VB_Name = "mdlTypes"
Option Explicit

Enum gWinVer
    wUnknown = 0
    w95_98 = 1
    wME = 2
    wNT = 3
    w2000 = 4
    wXP = 5
    w2003 = 6
End Enum
Enum gStatus
    sIdle = 0
    sPlay = 1
    sPaused = 2
    sSelectFile = 3
    sDecode = 4
    sBurn = 5
    sRip = 6
    sStop = 7
    sNormalize = 8
    sDoneBurning = 9
End Enum
Enum gFileType
    fMpegLayer3 = 1
    fCDAudio = 2
    fAllFileTypes = 4
End Enum
Enum gDriveType
    dHardDrive = 0
    dBurner = 1
    dCDDrive = 2
End Enum
Private Type gFile
    fFilename As String
End Type
Private Type gFiles
    fIndex As Integer
    fFile(14000) As gFile
    fCount As Long
End Type
Private Type gINIFiles
    iSettings As String
    iPlaylists As String
    iRecientPlaylist As String
    iErrorLog As String
    iCDAudioTracks As String
    iGFX As String
    iEffects As String
    iSkins As String
End Type
Private Type gDrive
    dDriveType As gDriveType
    dDriveNumber As Integer
    dDriveLetter As String
    dToc As String
End Type
Private Type gDrives
    dCount As Integer
    dDrive(32) As gDrive
    dCDCount As Integer
    dCurrentDrive As String
End Type
Enum eEventType
    eDecode = 2
    eNormalize = 3
    eStartBurn = 4
    ePlay = 5
    eRip = 6
    eEncode = 7
End Enum
Private Type gEvent
    eType As eEventType
    eInputFile As String
    eOutputFile As String
End Type
Private Type gEvents
    eEvent(100) As gEvent
    eCount As Integer
    eProcessing As Boolean
    eCurrentFile As String
End Type
Private Type gBurnQue
    bFiles(99) As String
    bCount As Integer
    bStatus As Integer
    bTrackIndex As Integer
    bNormInFile As String
    bNormOutFile As String
    bMergeFilename As String
End Type
Private Type gMenus
    mUsingSubMenu As Boolean
    mFileMenuVisible As Boolean
    mFileMenuIndex As Integer
    mEditMenuVisible As Boolean
    mEditMenuIndex As Integer
    mOpenMenuVisible As Boolean
    mOpenMenuIndex As Integer
    mConvertMenuIndex As Integer
    mConvertMenuVisible As Boolean
    mEffectsMenuIndex As Integer
    mEffectsMenuVisible As Boolean
End Type
Private Type gPlayer
    pEnabled As Boolean
    pStatus As gStatus
    pFileType As gFileType
    pFilename As String
End Type
Private Type gSettings
    sLatestVersion As String
    sAlwaysOnTop As Boolean
    sConvertKHZ As Boolean
    sTestMode As Boolean
    sFinalize As Boolean
    sAutoEject As Boolean
    sNormalize As Boolean
    sBitrate As Integer
    sHandleErrors As Boolean
    sShowSplash As Boolean
    sInitialWidth As Integer
    sInitialHeight As Integer
    sInitialTop As Integer
    sInitialLeft As Integer
    sTimeSelected As Long
    sDebugMode As Boolean
    sCheckTaskbarStatus As Boolean
    sMinimized As Boolean
    sFirstRun As Boolean
    sFindSelectIndex As Integer
    sConvertingKHZ As Boolean
    sSelectDirReturnValue As String
    sSupportedMedia As String
    sSupportedAudio As String
    sSupportedVideo As String
    sSettingAddress As Boolean
    sName As String
    sPassword As String
    sRegistered As Boolean
    sProcessScripts As Boolean
    sTreeviewSource As String
    sLastCDDrive As Integer
    sShowAboutDetailsOnStartup As Boolean
    sFullScreenVideo As Boolean
End Type
Private Type gTrack
    tFile As String
    tTitle As String
    tArtist As String
    tAlbum As String
    tYear As String
    tBitrate As Long
    tRip As Boolean
    tLength As String
    tName As String
End Type
Private Type gTracks
    tTrackCount As Integer
    tTrack(99) As gTrack
    tDiscLen As String
    tArtist As String
    tTitle As String
    tCount As Integer
    tGenre As Integer
End Type
Private Type gClipboard
    cFileToCopy As String
    cNewFilename As String
    cCopyOnly As Boolean
End Type
Private Type gScripts
    sMain_Resize As String
    sMain_Init As String
    sCurrentScript As String
End Type
Global lScripts As gScripts
Global lClipboard As gClipboard
Global lTracks As gTracks
Global lSettings As gSettings
Global lPlayer As gPlayer
Global lMenus As gMenus
Global lDrives As gDrives
Global lBurnQue As gBurnQue
Global lIniFiles As gINIFiles
Global lEvents As gEvents
Global lProgressClicked As Boolean
Global lTreeviewText As String
Global lTreeviewExpanded As Boolean
Global lSearchCanceled As Boolean
Global lTagVisible As Boolean
Global lFiles As gFiles
Global lAlt As Boolean
Global lTaskbar As Boolean
