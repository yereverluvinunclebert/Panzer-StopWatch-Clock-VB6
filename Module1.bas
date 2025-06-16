Attribute VB_Name = "Module1"
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
' gaugeForm_BubblingEvent ' leaving that here so I can copy/paste to find it

'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Module for declaring any public and private constants, APIs and types used by the functions therein.
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
'constants used to choose a font via the system dialog window
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE As Integer = 32
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
Private Const CF_SCREENFONTS As Long = &H1

'type declaration used to choose a font via the system dialog window
Private Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  Flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'APIs used to choose a font via the system dialog window
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' API and enums for acquiring the special folder paths
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Public Enum FolderEnum ' has to be public
    feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
    feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
    feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
    feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
    feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
    feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
    feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
    feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
    feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
    feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
    feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
    feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
    feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feProgramFiles = 38 ' \Program Files
    feProgramFilesCommon = 43 ' \Program Files\Common Files
    'feRecycleBin = 10 ' ???
    feUser = 40 ' \Docs & Settings\User
    feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
    feUserAppData = 26 ' \Docs & Settings\User\Application Data
    feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feUserCookies = 33 ' \Docs & Settings\User\Cookies
    feUserDesktop = 16 ' \Docs & Settings\User\Desktop
    feUserDocs = 5 ' \Docs & Settings\User\My Documents
    feUserFavorites = 6 ' \Docs & Settings\User\Favorites
    feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
    feUserNetHood = 19 ' \Docs & Settings\User\NetHood
    feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
    feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
    feUserRecent = 8 ' \Docs & Settings\User\Recent
    feUserSendTo = 9 ' \Docs & Settings\User\SendTo
    feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
    feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
    feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
    feUserTemplates = 21 ' \Docs & Settings\User\Templates
    feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
    feWindows = 36 ' \Windows
    feWindowFonts = 20 ' \Windows\Fonts
    feWindowsResources = 56 ' \Windows\Resources
    feWindowsSystem = 37 ' \Windows\System32
End Enum
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs for useful functions START
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' APIs for useful functions END
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Constants and APIs for playing sounds
Public Const SND_ASYNC As Long = &H1             '  play asynchronously
Public Const SND_FILENAME  As Long = &H20000     '  name is a file name
'Public Const SND_MEMORY = &H4
'Public Const SND_NODEFAULT = &H2

Public Declare Function PlaySound Lib "WINMM.DLL" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'API Functions to read/write information from INI File
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'constants and APIs defined for querying the registry
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
'------------------------------------------------------ ENDS





'------------------------------------------------------ STARTS
' Enums defined for opening a common dialog box to select files without OCX dependencies
Private Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum

' Types defined for opening a common dialog box to select files without OCX dependencies
Private Type OPENFILENAME
    lStructSize As Long    'The size of this struct (Use the Len function)
    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    Flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long 'LPCITEMIDLIST
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long  'BFFCALLBACK
    lParam As Long
    iImage As Long
End Type

' vars defined for opening a common dialog box to select files without OCX dependencies
Private x_OpenFilename As OPENFILENAME

' APIs declared for opening a common dialog box to select files without OCX dependencies
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
'Private Declare Function SHBrowseForFolderA Lib "Shell32.dll" (bInfo As BROWSEINFO) As Long
'Private Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl As Long, ByVal szPath As String) As Long
'Private Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'------------------------------------------------------ ENDS


'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As RECT) As Long
'Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As RECT) As Long
'
'Public Type RECT
'  Left As Long
'  Top As Long
'  Right As Long ' This is +1 (right - left = width)
'  Bottom As Long ' This is +1 (bottom - top = height)
'End Type

'------------------------------------------------------ STARTS
' APIs, constants and types defined for determining the OS version
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' stored vars read from settings.ini
'
' general
Public gblStartup As String
Public gblGaugeFunctions As String
'Public gblAnimationInterval As String
Public gblSmoothSecondHand As String


Public gblClockFaceSwitchPref As String
Public gblMainGaugeTimeZone As String
Public gblMainDaylightSaving As String
Public gblSecondaryGaugeTimeZone As String
Public gblSecondaryDaylightSaving As String

' config
Public gblEnableTooltips As String
Public gblEnablePrefsTooltips As String
Public gblEnableBalloonTooltips As String

Public gblShowTaskbar As String
Public gblDpiAwareness As String

Public gblGaugeSize As String
Public gblScrollWheelDirection As String

' position
Public gblAspectHidden As String
Public gblWidgetPosition As String
Public gblWidgetLandscape As String
Public gblWidgetPortrait As String
Public gblLandscapeFormHoffset As String
Public gblLandscapeFormVoffset As String
Public gblPortraitHoffset As String
Public gblPortraitYoffset As String
Public gblvLocationPercPrefValue As String
Public gblhLocationPercPrefValue As String

' sounds
Public gblEnableSounds  As String

' development
Public gblDebug As String
Public gblDblClickCommand As String
Public gblOpenFile As String
Public gblDefaultEditor As String
       
' font
Public gblClockFont As String
Public gblPrefsFont As String
Public gblPrefsFontSizeHighDPI As String
Public gblPrefsFontSizeLowDPI As String
Public gblPrefsFontItalics  As String
Public gblPrefsFontColour  As String

' window
Public gblWindowLevel As String
Public gblPreventDragging As String
Public gblOpacity  As String
Public gblWidgetHidden  As String
Public gblHidingTime  As String
Public gblIgnoreMouse  As String
Public gblFirstTimeRun  As String

' General storage variables declared
Public gblSettingsDir As String
Public gblSettingsFile As String

Public gblTrinketsDir      As String
Public gblTrinketsFile      As String

Public gblClockHighDpiXPos As String
Public gblClockHighDpiYPos As String
Public gblClockLowDpiXPos As String
Public gblClockLowDpiYPos As String
Public gblLastSelectedTab As String
Public gblSkinTheme As String
Public gblUnhide As String

' vars stored for positioning the prefs form
Public gblFormHighDpiXPosTwips As String
Public gblFormHighDpiYPosTwips As String

Public gblFormLowDpiXPosTwips As String
Public gblFormLowDpiYPosTwips As String


'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' General variables declared
'Public toolSettingsFile  As String
Public classicThemeCapable As Boolean
Public storeThemeColour As Long
Public windowsVer As String

' vars to obtain correct screen width (to correct VB6 bug)
Public screenWidthTwips As Long
Public screenHeightTwips As Long
Public screenHeightPixels As Long
Public screenWidthPixels As Long
Public oldScreenHeightPixels As Long
Public oldScreenWidthPixels As Long

' key presses
Public CTRL_1 As Boolean
Public SHIFT_1 As Boolean

' other globals
Public debugFlg As Integer
Public minutesToHide As Integer
Public aspectRatio As String
  
Public oldgblSettingsModificationTime  As Date

Public Const visibleAreaWidth As Long = 648 ' this is the width of the rightmost visible point of the widget - ie. the surround
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
     
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                            lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public gblWindowLevelWasChanged As Boolean

'------------------------------------------------------ ENDS
                            



' Flag for debug mode '.06 DAEB 19/04/2021 common.bas moved to the common area so that it can be used by each of the utilities
Private mbDebugMode As Boolean ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without

Public tzDelta As Long
Public tzDelta1 As Long

Public msgBoxADynamicSizingFlg As Boolean

Public gblStopWatchStartTime As Date ' these to be changed to properties
Public gblStopWatchState As Integer
Public gblStopWatchZeroed As Boolean

' These need to be changed to properties? or gbl variables
'Public rotationDegreesPerInterval  As Long ' rotationDegreesPerInterval = 0


'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : RobDog888 https://www.vbforums.com/member.php?17511-RobDog888
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    On Error GoTo fFExists_Error
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        fFExists = True
    Else
        fFExists = False
    End If

   On Error GoTo 0
   Exit Function

fFExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of Module Module1"
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : zeezee https://www.vbforums.com/member.php?90054-zeezee
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByVal pstrFolder As String) As Boolean
   On Error GoTo fDirExists_Error

    fDirExists = (PathFileExists(pstrFolder) = 1)
    If fDirExists Then fDirExists = (PathIsDirectory(pstrFolder) <> 0)

   On Error GoTo 0
   Exit Function

fDirExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of Module Module1"
End Function
''---------------------------------------------------------------------------------------
'' Procedure : fFExists
'' Author    : beededea
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function fFExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'    On Error GoTo fFExists_Error
'   'If debugflg = 1  Then Debug.Print "%fFExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fFExists = FS.FileExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'fFExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of module module1"
'End Function


'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Function fDirExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'    On Error GoTo fDirExists_Error
'   '''If debugflg = 1  Then msgBox "%fDirExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fDirExists = FS.FolderExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'fDirExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of module module1"
'End Function




'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
'Public Function fExtractSuffix(ByVal strPath As String) As String
'
'
'    Dim stringBits() As String ' string array
'    Dim upperBit As Integer: upperBit = 0
'
'    On Error GoTo fExtractSuffix_Error
'    '''If debugflg = 1  Then DebugPrint "%" & "fnExtractSuffix"
'
'    If strPath = vbNullString Then
'        fExtractSuffix = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        stringBits = Split(strPath, ".")
'        upperBit = UBound(stringBits)
'        fExtractSuffix = stringBits(upperBit)
'    Else
'        fExtractSuffix = strPath
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fExtractSuffix_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractSuffix of module module1"
'End Function
'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffixWithDot
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
'Public Function fExtractSuffixWithDot(ByVal strPath As String) As String
'
'    Dim stringBits() As String ' string array
'    Dim upperBit As Integer:    upperBit = 0
'
'    On Error GoTo fExtractSuffixWithDot_Error
'    '''If debugflg = 1  Then DebugPrint "%" & "fExtractSuffixWithDot"
'
'    If strPath = vbNullString Then
'        fExtractSuffixWithDot = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        stringBits = Split(strPath, ".")
'        upperBit = UBound(stringBits)
'        fExtractSuffixWithDot = "." & stringBits(upperBit)
'    Else
'        fExtractSuffixWithDot = vbNullString
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fExtractSuffixWithDot_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractSuffixWithDot of module module1"
'End Function

'---------------------------------------------------------------------------------------
' Procedure : fExtractFileNameNoSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the filename without a suffix
'---------------------------------------------------------------------------------------
'
'Public Function fExtractFileNameNoSuffix(ByVal strPath As String) As String
'
'    Dim stringBits() As String ' string array
'    Dim lowerBit As Integer:    lowerBit = 0
'
'    On Error GoTo fExtractFileNameNoSuffix_Error
'    '''If debugflg = 1  Then DebugPrint "%" & "fnExtractFileNameNoSuffix"
'
'    If strPath = vbNullString Then
'        fExtractFileNameNoSuffix = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        stringBits = Split(strPath, ".")
'        lowerBit = LBound(stringBits)
'        fExtractFileNameNoSuffix = stringBits(lowerBit)
'    Else
'        fExtractFileNameNoSuffix = strPath
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fExtractFileNameNoSuffix_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractFileNameNoSuffix of module module1"
'End Function
'
'---------------------------------------------------------------------------------------
' Procedure : fLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Function fLicenceState() As Integer
    Dim slicence As String: slicence = "0"
    
    On Error GoTo fLicenceState_Error
    ''If debugflg = 1  Then DebugPrint "%" & "fLicenceState"
    
    fLicenceState = 0
    ' read the tool's own settings file
    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = fGetINISetting("Software\PzStopWatch", "licence", gblSettingsFile)
        ' if the licence state is not already accepted then display the licence form
        If slicence = "1" Then fLicenceState = 1
    End If

   On Error GoTo 0
   Exit Function

fLicenceState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fLicenceState of Form common"

End Function
'---------------------------------------------------------------------------------------
' Procedure : showLicence
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Sub showLicence(ByVal licenceState As Integer)
    Dim slicence As String: slicence = "0"
    On Error GoTo showLicence_Error
    ''If debugflg = 1  Then DebugPrint "%" & "showLicence"
    
    ' if the licence state is not already accepted then display the licence form
    If licenceState = 0 Then
        'Call LoadFileToTB(frmLicence.txtLicenceTextBox, App.Path & "\Resources\txt\licence.txt", False)
        Call licenceSplash
    End If

   On Error GoTo 0
   Exit Sub

showLicence_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLicence of Form common"

End Sub
    
    
    

'---------------------------------------------------------------------------------------
' Procedure : setDPIaware
' Author    : beededea
' Date      : 29/10/2023
' Purpose   : This sets DPI awareness for the whole program incl. native VB6 forms, requires a program hard restart.
'---------------------------------------------------------------------------------------
'
Public Sub setDPIaware()
    On Error GoTo setDPIaware_Error
    
'    Cairo.SetDPIAwareness ' for debugging
'    msgBoxADynamicSizingFlg = True
    
    If gblDpiAwareness = "1" Then
        If Not InIDE Then
            Cairo.SetDPIAwareness ' this way avoids the VB6 IDE shrinking (sadly, VB6 has a high DPI unaware IDE)
            msgBoxADynamicSizingFlg = True
        End If
    End If


    On Error GoTo 0
    Exit Sub

setDPIaware_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setDPIaware of Module modMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : testDPIAndSetInitialAwareness
' Author    : beededea
' Date      : 29/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub testDPIAndSetInitialAwareness()
    On Error GoTo testDPIAndSetInitialAwareness_Error

    If fPixelsPerInchX() > 96 Then ' only DPI aware by default when greater than 'standard'
        gblDpiAwareness = "1"
        Call setDPIaware
    End If

    On Error GoTo 0
    Exit Sub

testDPIAndSetInitialAwareness_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testDPIAndSetInitialAwareness of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : beededea
' Date      : 26/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadFileToTB(ByVal TxtBox As Object, ByVal FilePath As String, Optional ByVal Append As Boolean = False)
    'PURPOSE: Loads file specified by FilePath into textcontrol
    '(e.g., Text Box, Rich Text Box) specified by TxtBox
    
    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    'Returns: True if Successful, false otherwise
    
    Dim iFile As Integer: iFile = 0
    Dim s As String: s = vbNullString
    
    On Error GoTo LoadFileToTB_Error

   ''If debugflg = 1  Then msgbox "%" & LoadFileToTB

    If Dir$(FilePath) = vbNullString Then Exit Sub
    
    On Error GoTo ErrorHandler:
    s = TxtBox.Text
    
    iFile = FreeFile
    Open FilePath For Input As #iFile
    s = Input(LOF(iFile), #iFile)
    If Append Then
        TxtBox.Text = TxtBox.Text & s
    Else
        TxtBox.Text = s
    End If
    
    'LoadFileToTB = True
    
ErrorHandler:
    If iFile > 0 Then Close #iFile

   On Error GoTo 0
   Exit Sub

LoadFileToTB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFileToTB of Form common"

End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : fGetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function fGetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo fGetINISetting_Error
    Const cparmLen As Integer = 500 ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen ' not going to initialise this with a 500 char string
    Dim sDefault As String * cparmLen
    Dim lLength As Long: lLength = 0

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    fGetINISetting = Mid$(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

fGetINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetINISetting of module module1"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : sPutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save INI Setting in the File
'---------------------------------------------------------------------------------------
'
Public Sub sPutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo sPutINISetting_Error

    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    unusedReturnValue = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

sPutINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sPutINISetting of module module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : savestring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savestring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long: keyhand = 0
    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    On Error GoTo savestring_Error

    unusedReturnValue = RegCreateKey(hKey, strPath, keyhand)
    unusedReturnValue = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    unusedReturnValue = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

savestring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savestring of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    : si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : Returns the path to the specified special folder (AppData etc)
'---------------------------------------------------------------------------------------
'
Public Function fSpecialFolder(ByVal pfe As FolderEnum) As String
    Const MAX_PATH As Integer = 260
    Dim strPath As String: strPath = vbNullString
    Dim strBuffer As String: strBuffer = vbNullString
    
   On Error GoTo fSpecialFolder_Error

    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    fSpecialFolder = strPath

   On Error GoTo 0
   Exit Function

fSpecialFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fSpecialFolder of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : addTargetfile
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Public Sub addTargetFile(ByVal fieldValue As String, ByRef retFileName As String)
    Dim FilePath As String: FilePath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    Const x_MaxBuffer As Integer = 256
    
    ''If debugflg = 1  Then Debug.Print "%" & "addTargetfile"
    
    On Error Resume Next
    
    ' fix default folder bug opening in wrong folder
    dialogInitDir = App.path
        
    ' set the default folder to the existing reference
    If Not fieldValue = vbNullString Then
        If fFExists(fieldValue) Then
            ' extract the folder name from the string
            FilePath = fGetDirectory(fieldValue)
            ' set the default folder to the existing reference
            dialogInitDir = FilePath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(fieldValue) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = fieldValue 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = App.path 'start dir, might be "C:\" or so also
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With
  

  Call obtainOpenFileName(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

   On Error GoTo 0
   
   Exit Sub

'addTargetfile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addTargetfile of module module1.bas"
 
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fGetDirectory
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function fGetDirectory(ByRef path As String) As String

   On Error GoTo fGetDirectory_Error
   ''If debugflg = 1  Then DebugPrint "%" & "fnGetDirectory"

    If InStrRev(path, "\") = 0 Then
        fGetDirectory = vbNullString
        Exit Function
    End If
    fGetDirectory = Left$(path, InStrRev(path, "\") - 1)

   On Error GoTo 0
   Exit Function

fGetDirectory_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetDirectory of module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : obtainOpenFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   : using GetOpenFileName API rturns file name and title, the filename will be buffered to 256 bytes
'---------------------------------------------------------------------------------------
'
Public Sub obtainOpenFileName(ByRef retFileName As String, ByRef retfileTitle As String)
   On Error GoTo obtainOpenFileName_Error
   ''If debugflg = 1  Then Debug.Print "%obtainOpenFileName"

  If GetOpenFileName(x_OpenFilename) <> 0 Then
'    If x_OpenFilename.lpstrFile = "*.*" Then
'        'txtTarget.Text = savLblTarget
'    Else
        retfileTitle = x_OpenFilename.lpstrFileTitle
        retFileName = x_OpenFilename.lpstrFile
'    End If
  'Else
    'The CANCEL button was pressed
    'MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

obtainOpenFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainOpenFileName of module module1.bas"
End Sub





'
'---------------------------------------------------------------------------------------
' Procedure : GetWindowsVersion
' Author    :
' Date      : 28/05/2023
' Purpose   : Returns the version of Windows that the user is running
'---------------------------------------------------------------------------------------
'
Public Function GetWindowsVersion() As String
    Dim OSV As OSVERSIONINFO
    
    On Error GoTo GetWindowsVersion_Error

    OSV.OSVSize = Len(OSV)

    If GetVersionEx(OSV) = 1 Then
        Select Case OSV.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"
                
                Select Case OSV.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case OSV.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case OSV.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista"
                            Case 1
                                GetWindowsVersion = "Windows 7"
                            Case 2
                                GetWindowsVersion = "Windows 8"
                            Case 3
                                GetWindowsVersion = "Windows 8.1"
                            Case 10
                                GetWindowsVersion = "Windows 10"
                        End Select
                End Select
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case OSV.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If

   On Error GoTo 0
   Exit Function

GetWindowsVersion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWindowsVersion of Module Module1"
End Function




'----------------------------------------
'Name: TestWinVer
'Description: Tests the multiplicity of Windows versions and returns some values
'----------------------------------------
Public Function fTestClassicThemeCapable() As Boolean

    '=================================
    '2000 / XP / NT / 7 / 8 / 10
    '=================================
    On Error GoTo fTestClassicThemeCapable_Error

    Dim ProgramFilesDir As String: ProgramFilesDir = vbNullString
    Dim strString As String: strString = vbNullString
    'Dim shortWindowsVer As String: shortWindowsVer = vbNullString
    
    fTestClassicThemeCapable = False
    windowsVer = vbNullString
    
    ' other variable assignments
    strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    windowsVer = strString

    ' note that when running in compatibility mode the o/s will respond with "Windows XP"
    ' The IDE runs in compatibility mode so it may report the wrong working folder

    'Get the value of "ProgramFiles", or "ProgramFilesDir"
        
    windowsVer = GetWindowsVersion
    
    Select Case windowsVer
    Case "Windows NT 4.0"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 2000"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows XP"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    Case "Windows Server 2003"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows Vista"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 7"
        fTestClassicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case Else ' windows 8/10/11+
        fTestClassicThemeCapable = False
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    End Select

    ProgramFilesDir = strString
    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems
    If Not fDirExists(ProgramFilesDir) Then
        ProgramFilesDir = "c:\program files" ' 32 bit systems
    End If
   
    On Error GoTo 0: Exit Function

fTestClassicThemeCapable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTestClassicThemeCapable of module module1"

End Function




'---------------------------------------------------------------------------------------
' Procedure : fGetstring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : get a string from the registry
'---------------------------------------------------------------------------------------
'
Public Function fGetstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long: keyhand = 0
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim unusedReturnValue As Integer: unusedReturnValue = 0

    Dim lValueType As Variant

    On Error GoTo fGetstring_Error

    unusedReturnValue = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                fGetstring = Left$(strBuf, intZeroPos - 1)
            Else
                fGetstring = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

fGetstring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetstring of module module1"
End Function



' select a font for the fnt form
'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub changeFont(ByVal frm As Form, ByVal fntNow As Boolean, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)
    
   On Error GoTo changeFont_Error

    fntWeight = 0
    fntStyle = False
    'fntColour = 0
    'fntBold = False
    'fntUnderline = False
    fntFontResult = False
    
    'If debugflg = 1  Then Debug.Print "%mnuFont_Click"

    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub
'
'    If fntWeight > 700 Then
'        'fntBold = True
'    Else
'        'fntBold = False
'    End If
    
    If fntFont <> vbNullString And fntNow = True Then
        Call changeFormFont(frm, fntFont, Val(fntSize), fntWeight, fntStyle, fntItalics, fntColour)
    End If
    
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Module Module1"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : select a font for the supplied form
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(ByRef currFont As String, ByRef currSize As Integer, ByRef currWeight As Integer, ByVal currStyle As Boolean, ByRef currColour As Long, ByRef currItalics As Boolean, ByRef currUnderline As Boolean, ByRef fontResult As Boolean)

    Dim thisFont As FormFontInfo

    On Error GoTo displayFontSelector_Error

    With thisFont
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = fDialogFont(thisFont)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If thisFont.Name = vbNullString Then thisFont.Name = "times new roman"
    If thisFont.Name = vbNullString Then Exit Sub
    
    With thisFont
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of module module1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : changeFormFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   : change the font throughout the whole form
'---------------------------------------------------------------------------------------
'
Public Sub changeFormFont(ByVal formName As Object, ByVal suppliedFont As String, ByVal suppliedSize As Integer, ByVal suppliedWeight As Integer, ByVal suppliedStyle As Boolean, ByVal suppliedItalics As Boolean, ByVal suppliedColour As Long)
    On Error GoTo changeFormFont_Error
        
    Dim Ctrl As Control
      
    ' loop through all the controls and identify the labels and text boxes
    For Each Ctrl In formName.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
            If Ctrl.Name <> "lblDragCorner" Then
                If suppliedFont <> vbNullString Then Ctrl.Font.Name = suppliedFont
                If suppliedSize > 0 Then Ctrl.Font.Size = suppliedSize
                Ctrl.Font.Italic = suppliedItalics
            End If
            Select Case True
                Case (TypeOf Ctrl Is CommandButton)
                    ' stupif fecking VB6 will not let you change the font of the forecolour on a button!
                    'Ctrl.ForeColor = suppliedColour
                    ' do nothing
                Case Else
                    Ctrl.ForeColor = suppliedColour
            End Select
        End If
    Next
     
   On Error GoTo 0
   Exit Sub

changeFormFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFormFont of module module1"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : fDialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : display the default windows dialog box that allows the user to select a font
'---------------------------------------------------------------------------------------
'
Public Function fDialogFont(ByRef f As FormFontInfo) As Boolean
      
    Dim logFnt As LOGFONT
    Dim ftStruc As FONTSTRUC
    Dim lLogFontAddress As Long: lLogFontAddress = 0
    Dim lMemHandle As Long: lMemHandle = 0
    Dim hWndAccessApp As Long: hWndAccessApp = 0
    
    Const LOGPIXELSY As Integer = 90        '  Logical pixels/inch in Y

    On Error GoTo fDialogFont_Error
    
    logFnt.lfWeight = f.Weight
    logFnt.lfItalic = f.Italic * -1
    logFnt.lfUnderline = f.UnderLine * -1
    logFnt.lfHeight = -fMulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
    f.Name = "Centurion Light SF"
    Call StringToByte(f.Name, logFnt.lfFaceName()) ' HERE
    ftStruc.rgbColors = f.Color
    ftStruc.lStructSize = Len(ftStruc)
    
    lMemHandle = GlobalAlloc(GHND, Len(logFnt))
    If lMemHandle = 0 Then
      fDialogFont = False
      Exit Function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
      fDialogFont = False
      Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, logFnt, Len(logFnt)
    ftStruc.lpLogFont = lLogFontAddress
    'ftStruc.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    ftStruc.Flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT
    If ChooseFont(ftStruc) = 1 Then
      CopyMemory logFnt, ByVal lLogFontAddress, Len(logFnt)
      f.Weight = logFnt.lfWeight
      f.Italic = CBool(logFnt.lfItalic)
      f.UnderLine = CBool(logFnt.lfUnderline)
      f.Name = fByteToString(logFnt.lfFaceName())
      f.Height = CLng(ftStruc.iPointSize / 10)
      f.Color = ftStruc.rgbColors
      fDialogFont = True
    Else
      fDialogFont = False
    End If

   On Error GoTo 0
   Exit Function

fDialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDialogFont of Module module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fMulDiv
' Author    :
' Date      : 21/08/2020
' Purpose   :  fMulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third 32-bit value.
'---------------------------------------------------------------------------------------
'
Private Function fMulDiv(ByVal In1 As Long, ByVal In2 As Long, ByVal In3 As Long) As Long
        
    Dim lngTemp As Long: lngTemp = 0
    On Error GoTo fMulDiv_Error
    
    On Error GoTo fMulDiv_err
    If In3 <> 0 Then
        lngTemp = In1 * In2
        lngTemp = lngTemp / In3
    Else
        lngTemp = -1
    End If

    fMulDiv = lngTemp
    Exit Function
fMulDiv_err:
    lngTemp = -1
    Resume fMulDiv_err

   On Error GoTo 0
   Exit Function

fMulDiv_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fMulDiv of Module module1"
End Function



'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a provided string to a byte array
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(ByVal InString As String, ByRef ByteArray() As Byte)
    
    Dim intLbound As Integer: intLbound = 0
    Dim intUbound As Integer: intUbound = 0
    Dim intLen As Integer: intLen = 0
    Dim intX As Integer: intX = 0
    
    On Error GoTo StringToByte_Error

    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
        ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
    Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StringToByte of Module module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fByteToString
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a byte array provided to a string
'---------------------------------------------------------------------------------------
'
Private Function fByteToString(ByRef aBytes() As Byte) As String
      
    Dim dwBytePoint As Long: dwBytePoint = 0
    Dim dwByteVal As Long: dwByteVal = 0
    Dim szOut As String: szOut = vbNullString
    
    On Error GoTo fByteToString_Error

    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes) ' whileing and wending my way through the bytearrays >sigh<
      dwByteVal = aBytes(dwBytePoint)
      If dwByteVal = 0 Then
        fByteToString = szOut
        Exit Function
      Else
        szOut = szOut & Chr$(dwByteVal)
      End If
      dwBytePoint = dwBytePoint + 1
    Wend
    fByteToString = szOut

   On Error GoTo 0
   Exit Function

fByteToString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fByteToString of Module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : aboutClickEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub aboutClickEvent()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo aboutClickEvent_Error
    
    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    ' The RC forms are measured in pixels so the positioning needs to pre-convert the twips into pixels
   
    fMain.aboutForm.Top = (screenHeightPixels / 2) - (fMain.aboutForm.Height / 2)
    fMain.aboutForm.Left = (screenWidthPixels / 2) - (fMain.aboutForm.Width / 2)
     
    fMain.aboutForm.Load
    fMain.aboutForm.Show
    
    'aboutWidget.opacity = 0
    aboutWidget.ShowMe = True
    aboutWidget.Widget.Refresh
    
    'fMain.aboutForm.Load
    'fMain.aboutForm.show
      
    If (fMain.aboutForm.WindowState = 1) Then
        fMain.aboutForm.WindowState = 0
    End If

   On Error GoTo 0
   Exit Sub

aboutClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure aboutClickEvent of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpSplash
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub helpSplash()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo helpSplash_Error

    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    fMain.helpForm.Top = (screenHeightPixels / 2) - (fMain.helpForm.Height / 2)
    fMain.helpForm.Left = (screenWidthPixels / 2) - (fMain.helpForm.Width / 2)
     
    'helpWidget.MyOpacity = 0
    helpWidget.ShowMe = True
    helpWidget.Widget.Refresh
    
    fMain.helpForm.Load
    fMain.helpForm.Show
    
     If (fMain.helpForm.WindowState = 1) Then
         fMain.helpForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

helpSplash_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpSplash of Form menuForm"
     
End Sub
'---------------------------------------------------------------------------------------
' Procedure : licenceSplash
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub licenceSplash()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo licenceSplash_Error

    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    
    fMain.licenceForm.Top = (screenHeightPixels / 2) - (fMain.licenceForm.Height / 2)
    fMain.licenceForm.Left = (screenWidthPixels / 2) - (fMain.licenceForm.Width / 2)
     
    'licenceWidget.opacity = 0
    'opacityflag = 0
    licenceWidget.ShowMe = True
    licenceWidget.Widget.Refresh
    
    fMain.licenceForm.Load
    fMain.licenceForm.Show

    ' the btnDecline_Click and btnAccept_Click are in modmain.bas
    
     If (fMain.licenceForm.WindowState = 1) Then
         fMain.licenceForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

licenceSplash_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure licenceSplash of Form menuForm"
     
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuCoffee_ClickEvent()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo mnuCoffee_ClickEvent_Error
    
    answer = vbYes
    answerMsg = " Help support the creation of more widgets like this, DO send us a coffee! This button opens a browser window and connects to the Kofi donate page for this widget). Will you be kind and proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Donate a Kofi", True, "mnuCoffeeClickEvent")

    If answer = vbYes Then
        Call ShellExecute(menuForm.hwnd, "Open", "https://www.ko-fi.com/yereverluvinunclebert", vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuCoffee_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuSupport_ClickEvent()

    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo mnuSupport_ClickEvent_Error
    
    answer = vbYes
    answerMsg = "Visiting the support page - this button opens a browser window and connects to our Github issues page where you can send us a support query. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Contact Support", True, "mnuSupportClickEvent")

    If answer = vbYes Then
        Call ShellExecute(menuForm.hwnd, "Open", "https://github.com/yereverluvinunclebert/Panzer-StopWatch-Clock-VB6/issues", vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuSupport_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuLicence_ClickEvent()

   On Error GoTo mnuLicence_ClickEvent_Error

    'Call LoadFileToTB(frmLicence.txtLicenceTextBox, App.Path & "\Resources\txt\licence.txt", False)
    
    Call licenceSplash

   On Error GoTo 0
   Exit Sub

mnuLicence_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setMainTooltips
' Author    : beededea
' Date      : 15/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setMainTooltips()
   On Error GoTo setMainTooltips_Error

    If gblEnableTooltips = "1" Then

        overlayWidget.Widget.ToolTip = "Use CTRL+mouse scrollwheel up/down to resize."
        helpWidget.Widget.ToolTip = "Click on me to make me go away."
        aboutWidget.Widget.ToolTip = "Click on me to make me go away."
        
        fAlpha.gaugeForm.Widgets("housing/tickbutton").Widget.ToolTip = "Choose smooth movement or regular ticks"
        fAlpha.gaugeForm.Widgets("housing/helpbutton").Widget.ToolTip = "Press for a little help"
        fAlpha.gaugeForm.Widgets("housing/startbutton").Widget.ToolTip = "Press to restart (when stopped)"
        fAlpha.gaugeForm.Widgets("housing/stopbutton").Widget.ToolTip = "Press to stop clock operation."
        fAlpha.gaugeForm.Widgets("housing/switchfacesbutton").Widget.ToolTip = "Press to do nothing at all."
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.ToolTip = "Press to lock the widget in place"
        fAlpha.gaugeForm.Widgets("housing/prefsbutton").Widget.ToolTip = "Press to open the widget preferences"
        fAlpha.gaugeForm.Widgets("housing/surround").Widget.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
        
    Else
        overlayWidget.Widget.ToolTip = vbNullString
        helpWidget.Widget.ToolTip = vbNullString
        aboutWidget.Widget.ToolTip = vbNullString
        
        fAlpha.gaugeForm.Widgets("housing/tickbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/helpbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/startbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/stopbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/switchfacesbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/prefsbutton").Widget.ToolTip = vbNullString
        fAlpha.gaugeForm.Widgets("housing/surround").Widget.ToolTip = vbNullString
        
    End If
    
    Call ChangeToolTipWidgetDefaultSettings(Cairo.ToolTipWidget.Widget)

   On Error GoTo 0
   Exit Sub

setMainTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMainTooltips of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ChangeToolTipWidgetDefaultSettings
' Author    : beededea
' Date      : 20/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ChangeToolTipWidgetDefaultSettings(ByRef My_Widget As cWidgetBase)

   On Error GoTo ChangeToolTipWidgetDefaultSettings_Error

    With My_Widget
    
        .FontName = gblClockFont
        .FontSize = Val(gblPrefsFontSizeLowDPI)
    
    End With

   On Error GoTo 0
   Exit Sub

ChangeToolTipWidgetDefaultSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChangeToolTipWidgetDefaultSettings of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 01/03/2023
' Purpose   : ' adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to
'               monitor/resolution changes.
'---------------------------------------------------------------------------------------
'
Public Sub makeVisibleFormElements()

    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0
    Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo makeVisibleFormElements_Error

    'NOTE that when you position a widget you are positioning the form it is drawn upon.

    If gblDpiAwareness = "1" Then
        formLeftPixels = Val(gblClockHighDpiXPos)
        formTopPixels = Val(gblClockHighDpiYPos)
    Else
        formLeftPixels = Val(gblClockLowDpiXPos)
        formTopPixels = Val(gblClockLowDpiYPos)
    End If
    
    ' The RC forms are measured in pixels, whereas the native forms are in twips, do remember that...

    monitorCount = fGetMonitorCount
    If monitorCount > 1 Then
        Call SetFormOnMonitor(fAlpha.gaugeForm.hwnd, formLeftPixels, formTopPixels)
    Else
        fAlpha.gaugeForm.Left = formLeftPixels
        fAlpha.gaugeForm.Top = formTopPixels
    End If
    
    fAlpha.gaugeForm.Show

    On Error GoTo 0
    Exit Sub

makeVisibleFormElements_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeVisibleFormElements of Module Module1"
            Resume Next
          End If
    End With
        
End Sub



'
''---------------------------------------------------------------------------------------
'' Procedure : fTwipsPerPixelX
'' Author    : Elroy from Vbforums
'' Date      : 23/01/2022
'' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
''---------------------------------------------------------------------------------------
''
'Public Function fTwipsPerPixelX() As Single
'    Dim hdc As Long: hdc = 0
'    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
'
'    Const LOGPIXELSX As Integer = 88        '  Logical pixels/inch in X
'    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
'    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
'    '
'    On Error GoTo fTwipsPerPixelX_Error
'
'    ' 23/01/2022 .01 monitorModule.bas DAEB added if then else if you can't get device context
'    hdc = GetDC(0)
'    If hdc <> 0 Then
'        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
'        ReleaseDC 0, hdc
'        fTwipsPerPixelX = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
'    Else
'        fTwipsPerPixelX = Screen.TwipsPerPixelX
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fTwipsPerPixelX_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelX of Module Module1"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : fTwipsPerPixelY
'' Author    : Elroy from Vbforums
'' Date      : 23/01/2022
'' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
''---------------------------------------------------------------------------------------
''
'Public Function fTwipsPerPixelY() As Single
'    Dim hdc As Long: hdc = 0
'    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
'
'    Const LOGPIXELSY As Integer = 90         '  Logical pixels/inch in Y
'    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
'    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
'
'   On Error GoTo fTwipsPerPixelY_Error
'
'    ' 23/01/2022 .01 monitorModule.bas DAEB added if then else if you can't get device context
'    hdc = GetDC(0)
'    If hdc <> 0 Then
'        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
'        ReleaseDC 0, hdc
'        fTwipsPerPixelY = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
'    Else
'        fTwipsPerPixelY = Screen.TwipsPerPixelY
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fTwipsPerPixelY_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelY of Module Module1"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
    '36 home
    '40 is down
    '38 is up
    '37 is left
    '39 is right
    '33  Page up
    '34  Page down
    '35  End
    'ctrl 116
    'Shift 16
    'f5 18
'---------------------------------------------------------------------------------------
'
Public Sub getKeyPress(ByVal KeyCode As Integer, ByVal Shift As Integer)

    On Error GoTo getkeypress_Error

    If CTRL_1 Or SHIFT_1 Then
            CTRL_1 = False
            SHIFT_1 = False
    End If
    
    If Shift Then
        SHIFT_1 = True
    End If

    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True
        Case vbKeyShift
            SHIFT_1 = True
        Case 82 ' R
            If Shift = 1 Then Call hardRestart
        Case 116
            Call reloadWidget 'f5 refresh button as per all browsers
    End Select
 
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Module module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : determineScreenDimensions
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : VB6 has a bug - it should return width in twips on my screen but often returns a faulty value when a game runs full screen, changing the resolution
'             the screen width determination is incorrect, the API call below resolves this.
'---------------------------------------------------------------------------------------
'
Public Sub determineScreenDimensions()

   On Error GoTo determineScreenDimensions_Error
   
    'If debugflg = 1 Then msgbox "% sub determineScreenDimensions"

    ' only calling TwipsPerPixelX/Y functions once on startup
    screenTwipsPerPixelX = fTwipsPerPixelX
    screenTwipsPerPixelY = fTwipsPerPixelY
    
    screenHeightPixels = GetDeviceCaps(menuForm.hdc, VERTRES) ' we use the name of any form that we don't mind being loaded at this point
    screenWidthPixels = GetDeviceCaps(menuForm.hdc, HORZRES)

    screenHeightTwips = screenHeightPixels * screenTwipsPerPixelY
    screenWidthTwips = screenWidthPixels * screenTwipsPerPixelX
    
    oldScreenHeightPixels = screenHeightPixels ' will be used to check for orientation changes
    oldScreenWidthPixels = screenWidthPixels
    
   On Error GoTo 0
   Exit Sub

determineScreenDimensions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & " in procedure determineScreenDimensions of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mainScreen
' Author    : beededea
' Date      : 04/05/2023
' Purpose   : Function to move the main_window onto the main screen - called on startup and by timer
'---------------------------------------------------------------------------------------
'
Public Sub mainScreen()
   On Error GoTo mainScreen_Error

    ' check for aspect ratio and determine whether it is in portrait or landscape mode
    If screenWidthPixels > screenHeightPixels Then
        aspectRatio = "landscape"
    Else
        aspectRatio = "portrait"
    End If
    
    ' check if the widget has a lock for the screen type.
    If aspectRatio = "landscape" Then
        If gblWidgetLandscape = "1" Then
            If gblLandscapeFormHoffset <> vbNullString Then
                fAlpha.gaugeForm.Left = Val(gblLandscapeFormHoffset)
                fAlpha.gaugeForm.Top = Val(gblLandscapeFormVoffset)
            End If
        End If
        If gblAspectHidden = "2" Then
            Debug.Print "Hiding the widget for landscape mode"
            fAlpha.gaugeForm.Visible = False
        End If
    End If
    
    ' check if the widget has a lock for the screen type.
    If aspectRatio = "portrait" Then
        If gblWidgetPortrait = "1" Then
            fAlpha.gaugeForm.Left = Val(gblPortraitHoffset)
            fAlpha.gaugeForm.Top = Val(gblPortraitYoffset)
        End If
        If gblAspectHidden = "1" Then
            Debug.Print "Hiding the widget for portrait mode"
            fAlpha.gaugeForm.Visible = False
        End If
    End If

    ' calculate the on screen widget position
    If fAlpha.gaugeForm.Left < 0 Then
        fAlpha.gaugeForm.Left = 10
    End If
    If fAlpha.gaugeForm.Top < 0 Then
        fAlpha.gaugeForm.Top = 0
    End If
    If fAlpha.gaugeForm.Left > screenWidthPixels - 50 Then
        fAlpha.gaugeForm.Left = screenWidthPixels - 150
    End If
    If fAlpha.gaugeForm.Top > screenHeightPixels - 50 Then
        fAlpha.gaugeForm.Top = screenHeightPixels - 150
    End If

    ' calculate the current hlocation in % of the screen
    ' store the current hlocation in % of the screen
    If gblWidgetPosition = "1" Then
        gblhLocationPercPrefValue = Str$(fAlpha.gaugeForm.Left / screenWidthPixels * 100)
        gblvLocationPercPrefValue = Str$(fAlpha.gaugeForm.Top / screenHeightPixels * 100)
    End If

   On Error GoTo 0
   Exit Sub

mainScreen_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mainScreen of Module Module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : thisForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine
'---------------------------------------------------------------------------------------
'
Public Sub thisForm_Unload() ' name follows VB6 standard naming convention
    On Error GoTo Form_Unload_Error
    
    Call savePosition
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 and RC6 forms
'---------------------------------------------------------------------------------------
'
Public Sub unloadAllForms(ByVal endItAll As Boolean)
    
   On Error GoTo unloadAllForms_Error

    ' stop all VB6 timers in the timer form
    frmTimer.revealWidgetTimer.Enabled = False
    frmTimer.rotationTimer.Enabled = False
    frmTimer.settingsTimer.Enabled = False
    'frmTimer.sleepTimer.Enabled = False
    
    ' stop all VB6 timers in the prefs form
    
    panzerPrefs.themeTimer.Enabled = False
    panzerPrefs.positionTimer.Enabled = False
    
    ' stop all RC6 timers
    
    overlayWidget.tmrClock.Enabled = False
    overlayWidget.tmrStopWatch.Enabled = False
    overlayWidget.tmrSWRotation.Enabled = False
    overlayWidget.tmrHandsRotation.Enabled = False
    
    'unload the RC6 widgets on the RC6 forms first
    
    aboutWidget.Widgets.RemoveAll
    helpWidget.Widgets.RemoveAll
    fAlpha.gaugeForm.Widgets.RemoveAll
    
    ' unload the native VB6 forms
    
    Unload panzerPrefs
    Unload frmMessage
    Unload frmTimer
    Unload menuForm

    ' RC6's own method for killing forms

    fMain.aboutForm.Unload
    fMain.helpForm.Unload
    fAlpha.gaugeForm.Unload
    fMain.licenceForm.Unload
    
    ' remove all variable references to each form in turn
    
    Set panzerPrefs = Nothing
    Set frmMessage = Nothing
    Set frmTimer = Nothing
    Set menuForm = Nothing
        
    ' remove all variable references to each VB6 form in turn
    
    Set fMain.aboutForm = Nothing
    Set fMain.helpForm = Nothing
    Set fAlpha.gaugeForm = Nothing
    Set fMain.licenceForm = Nothing
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unloadAllForms of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : reloadWidget
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub reloadWidget()
    
    On Error GoTo reloadWidget_Error
    
    Call savePosition
    
    Call unloadAllForms(False) ' unload forms but do not END
    
    ' this will call the routines as called by sub main() and initialise the program and RELOAD the RC6 forms.
    Call mainRoutine(True) ' sets the restart flag to avoid repriming the RC6 message pump.

    On Error GoTo 0
    Exit Sub

reloadWidget_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reloadWidget of Module Module1"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : savePosition
' Author    : beededea
' Date      : 04/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savePosition()

   On Error GoTo savePosition_Error

    If gblDpiAwareness = "1" Then
        gblClockHighDpiXPos = Str$(fAlpha.gaugeForm.Left) ' saving in pixels
        gblClockHighDpiYPos = Str$(fAlpha.gaugeForm.Top)
        sPutINISetting "Software\PzStopWatch", "clockHighDpiXPos", gblClockHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzStopWatch", "clockHighDpiYPos", gblClockHighDpiYPos, gblSettingsFile
    Else
        gblClockLowDpiXPos = Str$(fAlpha.gaugeForm.Left) ' saving in pixels
        gblClockLowDpiYPos = Str$(fAlpha.gaugeForm.Top)
        sPutINISetting "Software\PzStopWatch", "clockLowDpiXPos", gblClockLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzStopWatch", "clockLowDpiYPos", gblClockLowDpiYPos, gblSettingsFile
    End If
    
    gblGaugeSize = Str$(fAlpha.gaugeForm.WidgetRoot.Zoom * 100)
    sPutINISetting "Software\PzStopWatch", "gaugeSize", gblGaugeSize, gblSettingsFile

   On Error GoTo 0
   Exit Sub

savePosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savePosition of Module Module1"
    
End Sub
    
    

'---------------------------------------------------------------------------------------
' Procedure : makeProgramPreferencesAvailable
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : open the prefs
'---------------------------------------------------------------------------------------
'
Public Sub makeProgramPreferencesAvailable()
    On Error GoTo makeProgramPreferencesAvailable_Error

    If panzerPrefs.IsVisible = False Then
        panzerPrefs.Visible = True
        panzerPrefs.Show  ' show it again
        panzerPrefs.SetFocus

        If panzerPrefs.WindowState = vbMinimized Then
            panzerPrefs.WindowState = vbNormal
        End If

        ' set the current position of the utility according to previously stored positions
        
        Call readPrefsPosition
        Call panzerPrefs.positionPrefsMonitor
    Else
        panzerPrefs.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

makeProgramPreferencesAvailable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeProgramPreferencesAvailable of module module1.bas"
End Sub
    

'---------------------------------------------------------------------------------------
' Procedure : readPrefsPosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : read the form X/Y params from the toolSettings.ini
'---------------------------------------------------------------------------------------
'
Public Sub readPrefsPosition()
            
   On Error GoTo readPrefsPosition_Error

    If gblDpiAwareness = "1" Then
        gblFormHighDpiXPosTwips = fGetINISetting("Software\PzStopWatch", "formHighDpiXPosTwips", gblSettingsFile)
        gblFormHighDpiYPosTwips = fGetINISetting("Software\PzStopWatch", "formHighDpiYPosTwips", gblSettingsFile)
        
'        ' if a current location not stored then position to the middle of the screen
'        If gblFormHighDpiXPosTwips <> "" Then
'            panzerPrefs.Left = Val(gblFormHighDpiXPosTwips)
'        Else
'            panzerPrefs.Left = screenWidthTwips / 2 - panzerPrefs.Width / 2
'        End If
'
'        If gblFormHighDpiYPosTwips <> "" Then
'            panzerPrefs.Top = Val(gblFormHighDpiYPosTwips)
'        Else
'            panzerPrefs.Top = Screen.Height / 2 - panzerPrefs.Height / 2
'        End If
    Else
        gblFormLowDpiXPosTwips = fGetINISetting("Software\PzStopWatch", "formLowDpiXPosTwips", gblSettingsFile)
        gblFormLowDpiYPosTwips = fGetINISetting("Software\PzStopWatch", "formLowDpiYPosTwips", gblSettingsFile)
        
'        ' if a current location not stored then position to the middle of the screen
'        If gblFormLowDpiXPosTwips <> "" Then
'            panzerPrefs.Left = Val(gblFormLowDpiXPosTwips)
'        Else
'            panzerPrefs.Left = screenWidthTwips / 2 - panzerPrefs.Width / 2
'        End If
'
'        If gblFormLowDpiYPosTwips <> "" Then
'            panzerPrefs.Top = Val(gblFormLowDpiYPosTwips)
'        Else
'            panzerPrefs.Top = Screen.Height / 2 - panzerPrefs.Height / 2
'        End If
    End If
   
   On Error GoTo 0
   Exit Sub

readPrefsPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readPrefsPosition of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : writePrefsPosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : save the current X and y position of this form to allow repositioning when restarting
'---------------------------------------------------------------------------------------
'
Public Sub writePrefsPosition()
        
   On Error GoTo writePrefsPosition_Error

    If panzerPrefs.WindowState = vbNormal Then ' when vbMinimised the value = -48000  !
        If gblDpiAwareness = "1" Then
            gblFormHighDpiXPosTwips = CStr(panzerPrefs.Left)
            gblFormHighDpiYPosTwips = CStr(panzerPrefs.Top)
            
            ' now write those params to the toolSettings.ini
            sPutINISetting "Software\PzStopWatch", "formHighDpiXPosTwips", gblFormHighDpiXPosTwips, gblSettingsFile
            sPutINISetting "Software\PzStopWatch", "formHighDpiYPosTwips", gblFormHighDpiYPosTwips, gblSettingsFile
        Else
            gblFormLowDpiXPosTwips = CStr(panzerPrefs.Left)
            gblFormLowDpiYPosTwips = CStr(panzerPrefs.Top)
            
            ' now write those params to the toolSettings.ini
            sPutINISetting "Software\PzStopWatch", "formLowDpiXPosTwips", gblFormLowDpiXPosTwips, gblSettingsFile
            sPutINISetting "Software\PzStopWatch", "formLowDpiYPosTwips", gblFormLowDpiYPosTwips, gblSettingsFile
            
        End If
        
    End If
    
    On Error GoTo 0
   Exit Sub

writePrefsPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writePrefsPosition of Form panzerPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : settingsTimer_Timer
' Author    : beededea
' Date      : 03/03/2023
' Purpose   : Checking the date/time of the settings.ini file meaning that another tool has edited the settings
'---------------------------------------------------------------------------------------
' this has to be in a shared module and not in the prefs form as it will be running in the normal context woithout prefs showing.

Private Sub settingsTimer_Timer()

    Dim timeDifferenceInSecs As Long: timeDifferenceInSecs = 0 ' max 86 years as a LONG in secs
    Dim settingsModificationTime As Date: settingsModificationTime = #1/1/2000 12:00:00 PM#
    
    On Error GoTo settingsTimer_Timer_Error

    If Not fFExists(gblSettingsFile) Then
        MsgBox ("%Err-I-ErrorNumber 13 - FCW was unable to access the dock settings ini file. " & vbCrLf & gblSettingsFile)
        Exit Sub
    End If
    
    ' check the settings.ini file date/time
    settingsModificationTime = FileDateTime(gblSettingsFile)
    timeDifferenceInSecs = Int(DateDiff("s", oldgblSettingsModificationTime, settingsModificationTime))

    ' if the settings.ini has been modified then reload the map
    If timeDifferenceInSecs > 1 Then

        oldgblSettingsModificationTime = settingsModificationTime
        
    End If
    
    On Error GoTo 0
    Exit Sub

settingsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form module1.bas"
            Resume Next
          End If
    End With
End Sub






'---------------------------------------------------------------------------------------
' Procedure : lockWidget
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub lockWidget()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo lockWidget_Error

    fileToPlay = "lock.wav"
    
    If gblPreventDragging = "1" Then
        menuForm.mnuLockWidget.Checked = False
        panzerPrefs.chkPreventDragging.Value = 0
        gblPreventDragging = "0"
        overlayWidget.Locked = False
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockWidget.Checked = True
        panzerPrefs.chkPreventDragging.Value = 1
        overlayWidget.Locked = True
        gblPreventDragging = "1"
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If
    
    fAlpha.gaugeForm.Refresh
    
    sPutINISetting "Software\PzStopWatch", "preventDragging", gblPreventDragging, gblSettingsFile
   
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    On Error GoTo 0
   Exit Sub

lockWidget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockWidget of Module Module1"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : SwitchOff
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SwitchOff()

   On Error GoTo SwitchOff_Error

    overlayWidget.Ticking = False
    menuForm.mnuSwitchOff.Checked = True
    menuForm.mnuTurnFunctionsOn.Checked = False
    
    gblGaugeFunctions = "0"
    sPutINISetting "Software\PzStopWatch", "gaugeFunctions", gblGaugeFunctions, gblSettingsFile

   On Error GoTo 0
   Exit Sub

SwitchOff_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SwitchOff of Module Module1"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : TurnFunctionsOn
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub TurnFunctionsOn()
    Dim fileToPlay As String: fileToPlay = vbNullString

   On Error GoTo TurnFunctionsOn_Error

    fileToPlay = "ting.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    overlayWidget.Ticking = True
    menuForm.mnuSwitchOff.Checked = False
    menuForm.mnuTurnFunctionsOn.Checked = True
    
    gblGaugeFunctions = "1"
    sPutINISetting "Software\PzStopWatch", "gaugeFunctions", gblGaugeFunctions, gblSettingsFile

   On Error GoTo 0
   Exit Sub

TurnFunctionsOn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TurnFunctionsOn of Form menuForm"
End Sub


'''---------------------------------------------------------------------------------------
''' Procedure : IsDLSavings
''' Author    : beededea
''' Date      : 13/08/2023
''' Purpose   :
'''---------------------------------------------------------------------------------------
'''
'Public Function IsDLSavings() As Boolean
'
'    Dim uInfo As TimeZoneInfo, lReturn As Long
'
'    On Error GoTo IsDLSavings_Error
'
'    lReturn = GetTimeZoneInformation(uInfo)
'
'    If lReturn = TIME_ZONE_ID_DAYLIGHT Then
'        IsDLSavings = True
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'IsDLSavings_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsDLSavings of Module Module1"
'
'End Function




'---------------------------------------------------------------------------------------
' Procedure : hardRestart
' Author    : beededea
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub hardRestart()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo hardRestart_Error

    thisCommand = App.path & "\restart.exe"
    
    If fFExists(thisCommand) Then
        ' run the selected program
        Call ShellExecute(panzerPrefs.hwnd, "open", thisCommand, "Panzer Stopwatch Widget.exe", "prefs", 1)
    Else
        'answer = MsgBox(thisCommand & " is missing", vbOKOnly + vbExclamation)
        answerMsg = thisCommand & " is missing"
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Restart Error Notification", False)
    End If

   On Error GoTo 0
   Exit Sub

hardRestart_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure hardRestart of Module Module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : InIDE
' Author    :
' Date      : 09/02/2021
' Purpose   : checks whether the code is running in the VB6 IDE or not
'---------------------------------------------------------------------------------------
'
Public Function InIDE() As Boolean

   On Error GoTo InIDE_Error

    ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        InIDE = True
    End If

   On Error GoTo 0
   Exit Function

InIDE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InIDE of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : InDebugMode
' Author    : beededea
' Date      : 02/03/2021
' Purpose   : ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
'---------------------------------------------------------------------------------------
'
Private Function InDebugMode() As Boolean
   On Error GoTo InDebugMode_Error

    mbDebugMode = True
    InDebugMode = True

   On Error GoTo 0
   Exit Function

InDebugMode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InDebugMode of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : clearAllMessageBoxRegistryEntries
' Author    : beededea
' Date      : 11/04/2023
' Purpose   : Clear all the message box "show again" entries in the registry
'---------------------------------------------------------------------------------------
'
Public Sub clearAllMessageBoxRegistryEntries()
    On Error GoTo clearAllMessageBoxRegistryEntries_Error

    SaveSetting App.EXEName, "Options", "Show message" & "mnuFacebookClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuLatestClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuSweetsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuWidgetsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuCoffeeClickEvent", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuSupportClickEvent", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkDpiAwarenessRestart", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkDpiAwarenessAbnormal", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkEnableTooltipsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "lblGitHubDblClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "sliOpacityClick", 0

    On Error GoTo 0
    Exit Sub

clearAllMessageBoxRegistryEntries_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearAllMessageBoxRegistryEntries of Form dock"
            Resume Next
          End If
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : determineIconWidth
' Author    : beededea
' Date      : 02/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function determineIconWidth(ByRef thisForm As Form, ByVal thisDynamicSizingFlg As Boolean) As Long

    Dim topIconWidth As Long: topIconWidth = 0
    
    On Error GoTo determineIconWidth_Error
    
'    If thisDynamicSizingFlg = False Then
'        'Exit Function
'    End If
    
    If thisForm.Width < 10500 Then
        topIconWidth = 600 '40 pixels
    End If
    
    If thisForm.Width >= 10500 And thisForm.Width < 12000 Then
        topIconWidth = 730
    End If
            
    If thisForm.Width >= 12000 And thisForm.Width < 13500 Then
        topIconWidth = 834
    End If
            
    If thisForm.Width >= 13500 And thisForm.Width < 15000 Then
        topIconWidth = 940
    End If
            
    If thisForm.Width >= 15000 Then
        topIconWidth = 1010
    End If
    'topIconWidth = 2000
    determineIconWidth = topIconWidth
    
    On Error GoTo 0
    Exit Function

determineIconWidth_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineIconWidth of Form panzerPrefs"

End Function

'---------------------------------------------------------------------------------------
' Procedure : ArrayString
' Author    : beededea
' Date      : 09/10/2023
' Purpose   : allows population of a string array from a comma separated string
'             VB6 normally creates a variant when assigning a comma separated string to a var with an undeclared type
'             this avoids that scenario.
'---------------------------------------------------------------------------------------
'
Public Function ArrayString(ParamArray tokens()) As String()
    On Error GoTo ArrayString_Error

    ReDim Arr(UBound(tokens)) As String
    Dim I As Long
    For I = 0 To UBound(tokens)
        Arr(I) = tokens(I)
    Next
    ArrayString = Arr

    On Error GoTo 0
    Exit Function

ArrayString_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ArrayString of Module Module1"
End Function

