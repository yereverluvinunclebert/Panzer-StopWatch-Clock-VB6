Attribute VB_Name = "modMain"

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS

Public fMain As New cfMain
Public aboutWidget As cwAbout

Public revealWidgetTimerCount As Integer
 
Public fAlpha As New cfAlpha
Public overlayWidget As cwOverlay

'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Main()
    On Error GoTo Main_Error
    
    Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : beededea
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)
    Dim extractCommand As String: extractCommand = vbNullString
    Dim chosenDragLayer As String: chosenDragLayer = vbNullString
    Dim thisPSDFullPath As String: thisPSDFullPath = vbNullString

    On Error GoTo main_routine_Error
    
    chosenDragLayer = "stopwatch/face/housing/surround"
    thisPSDFullPath = App.Path & "\Res\tank-clock-mk1.psd"
    fAlpha.FX = 222 'init position- and zoom-values (directly set on Public-Props of the Form-hosting Class)
    fAlpha.FY = 111
    fAlpha.FZ = 0.4
        
    Cairo.SetDPIAwareness ' this is off for the moment
 
    'load the collection for storing the overlay surfaces with its relevant keys direct from the PSD
    If restart = False Then Call loadExcludePathCollection ' no need to reload the collPSDNonUIElements layer name keys
    
    ' start the load of the PSD file using the RC6 PSD-Parser.instance
    Call fAlpha.InitFromPSD(thisPSDFullPath, chosenDragLayer)  ' no optional close layer as 3rd param
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\PzStopwatch", PzGSettingsFile)
        
    ' check first usage and display licence screen
    Call checkLicenceState

    ' initialise and create the main forms on the current display
    Call createAboutFormOnCurrentDisplay
    
    ' set the z-ordering of the main form
    Call setWindowZordering
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
    
    ' run the functions that are also called at reload time.
    Call adjustMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' check the selected monitor properties to determine form placement
    'Call monitorProperties(frmHidden) - might use RC6 for this?
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if a first time run shows prefs
    If PzGFirstTimeRun = "true" Then     'parse the command line
        Call makeProgramPreferencesAvailable
    End If
    
    ' check for first time running
    Call checkFirstTime

    ' configure any global timers here
    Call configureTimers

    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop
  
    'Debug.Print "App-ShutDown (one can buffer these values for the next run):"; fAlpha.FX; fAlpha.FY; fAlpha.FZ
   
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain"
    
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If PzGFirstTimeRun = "true" Then
        PzGFirstTimeRun = "false"
        sPutINISetting "Software\PzStopwatch", "firstTimeRun", PzGFirstTimeRun, PzGSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error

    ' general
    PzGStartup = vbNullString
    PzGGaugeFunctions = vbNullString
'    PzGAnimationInterval = vbNullString

    PzGClockFaceSwitchPref = vbNullString
    PzGMainGaugeTimeZone = vbNullString
    PzGMainDaylightSaving = vbNullString
    PzGSecondaryGaugeTimeZone = vbNullString
    PzGSecondaryDaylightSaving = vbNullString

    ' config
    PzGEnableTooltips = vbNullString
    PzGEnableBalloonTooltips = vbNullString
    PzGShowTaskbar = vbNullString
    
    PzGGaugeSize = vbNullString
    PzGScrollWheelDirection = vbNullString
    
    ' position
    PzGAspectHidden = vbNullString
    PzGWidgetPosition = vbNullString
    PzGWidgetLandscape = vbNullString
    PzGWidgetPortrait = vbNullString
    PzGLandscapeFormHoffset = vbNullString
    PzGLandscapeFormVoffset = vbNullString
    PzGPortraitHoffset = vbNullString
    PzGPortraitYoffset = vbNullString
    PzGvLocationPercPrefValue = vbNullString
    PzGhLocationPercPrefValue = vbNullString
    
    ' sounds
    PzGEnableSounds = vbNullString
    
    ' development
    PzGDebug = vbNullString
    PzGDblClickCommand = vbNullString
    PzGOpenFile = vbNullString
    PzGDefaultEditor = vbNullString
         
    ' font
    PzGPrefsFont = vbNullString
    PzGPrefsFontSize = vbNullString
    PzGPrefsFontItalics = vbNullString
    PzGPrefsFontColour = vbNullString
    
    ' window
    PzGWindowLevel = vbNullString
    PzGPreventDragging = vbNullString
    PzGOpacity = vbNullString
    PzGWidgetHidden = vbNullString
    PzGHidingTime = vbNullString
    PzGIgnoreMouse = vbNullString
    PzGFirstTimeRun = vbNullString
    
    ' general storage variables declared
    PzGSettingsDir = vbNullString
    PzGSettingsFile = vbNullString
    PzGMaximiseFormX = vbNullString
    PzGMaximiseFormY = vbNullString
    PzGLastSelectedTab = vbNullString
    PzGSkinTheme = vbNullString
    
    ' general variables declared
    toolSettingsFile = vbNullString
    classicThemeCapable = False
    storeThemeColour = 0
    windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    screenTwipsPerPixelX = 0
    screenTwipsPerPixelY = 0
    screenWidthTwips = 0
    screenHeightTwips = 0
    screenHeightPixels = 0
    screenWidthPixels = 0
    oldScreenHeightPixels = 0
    oldScreenWidthPixels = 0
    
    ' key presses
    CTRL_1 = False
    SHIFT_1 = False
    
    ' other globals
    debugflg = 0
    minutesToHide = 0
    aspectRatio = vbNullString
    revealWidgetTimerCount = 0
    oldPzGSettingsModificationTime = #1/1/2000 12:00:00 PM#

   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

    Cairo.ImageList.AddImage "about", App.Path & "\Resources\images\about.png"
    
'    'add Resources to the global ImageList
'    Cairo.ImageList.AddImage "surround", App.Path & "\Resources\images\surround.png"
'    Cairo.ImageList.AddImage "switchFacesButton", App.Path & "\Resources\images\switchFacesButton.png"
'    Cairo.ImageList.AddImage "startButton", App.Path & "\Resources\images\startButton.png"
'    Cairo.ImageList.AddImage "stopButton", App.Path & "\Resources\images\stopButton.png"
'    Cairo.ImageList.AddImage "pin", App.Path & "\Resources\images\pin.png"
'    Cairo.ImageList.AddImage "prefs", App.Path & "\Resources\images\prefs01.png"
'    Cairo.ImageList.AddImage "helpButton", App.Path & "\Resources\images\helpButton.png"
'    Cairo.ImageList.AddImage "tickSwitch", App.Path & "\Resources\images\tickSwitch.png"
'
'    For useloop = 1 To 36
'        Cairo.ImageList.AddImage "EarthGlobe" & useloop, App.Path & "\Resources\images\globe\Earth-spinning_" & useloop & ".png"
'    Next useloop
'
'    Cairo.ImageList.AddImage "Ring", App.Path & "\Resources\images\Ring.png", 545, 545
'    Cairo.ImageList.AddImage "Glow", App.Path & "\Resources\images\Glow.png"
'    Cairo.ImageList.AddImage "bigReflection", App.Path & "\Resources\images\bigReflection.png"
'    Cairo.ImageList.AddImage "windowReflection", App.Path & "\Resources\images\windowReflection.png"
    
    Cairo.ImageList.AddImage "frmIcon", App.Path & "\Resources\images\Icon.png"

   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the globe and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
'    'overlayWidget.RotationSpeed = Val(PzGAnimationInterval)
'    'overlayWidget.Zoom = Val(PzGGaugeSize) / 100
'    'overlayWidget.ZoomDirection = PzGScrollWheelDirection
'
'    If 'overlayWidget.Hidden = False Then
'        'overlayWidget.opacity = Val(PzGOpacity) / 100
'        'overlayWidget.Widget.Refresh
'    End If


'PzGClockFaceSwitchPref
'PzGMainGaugeTimeZone
'PzGMainDaylightSaving
'PzGSecondaryGaugeTimeZone
'PzGSecondaryDaylightSaving
    
    If PzGGaugeFunctions = "1" Then
        overlayWidget.Ticking = True
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        overlayWidget.Ticking = False
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If PzGDefaultEditor <> vbNullString And PzGDebug = "1" Then
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & PzGDefaultEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
        
    If PzGPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        overlayWidget.Locked = False
    Else
        menuForm.mnuLockWidget.Checked = True
        overlayWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
    End If
    
    If PzGShowTaskbar = "0" Then
        fAlpha.gaugeForm.ShowInTaskbar = False
    Else
        fAlpha.gaugeForm.ShowInTaskbar = True
    End If
                 
    ' set the z-ordering of the window
    Call setWindowZordering
    
    ' set the tooltips on the main screen
    Call setMainTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime

    If minutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & minutesToHide & " min."

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setWindowZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setWindowZordering()

   On Error GoTo setWindowZordering_Error

'    If Val(PzGWindowLevel) = 0 Then
'        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(PzGWindowLevel) = 1 Then
'        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(PzGWindowLevel) = 2 Then
'        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
'    End If

   On Error GoTo 0
   Exit Sub

setWindowZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setWindowZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal location As String, ByVal PzGSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(PzGSettingsFile) Then
        
        ' general
        PzGStartup = fGetINISetting(location, "startup", PzGSettingsFile)
        PzGGaugeFunctions = fGetINISetting(location, "gaugeFunctions", PzGSettingsFile)
'        PzGAnimationInterval = fGetINISetting(location, "animationInterval", PzGSettingsFile)
        

        PzGClockFaceSwitchPref = fGetINISetting(location, "clockFaceSwitchPref", PzGSettingsFile)
        PzGMainGaugeTimeZone = fGetINISetting(location, "mainGaugeTimeZone", PzGSettingsFile)
        PzGMainDaylightSaving = fGetINISetting(location, "mainDaylightSaving", PzGSettingsFile)
        PzGSecondaryGaugeTimeZone = fGetINISetting(location, "secondaryGaugeTimeZone", PzGSettingsFile)
        PzGSecondaryDaylightSaving = fGetINISetting(location, "secondaryDaylightSaving", PzGSettingsFile)

        ' configuration
        PzGEnableTooltips = fGetINISetting(location, "enableTooltips", PzGSettingsFile)
        PzGEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", PzGSettingsFile)
        PzGShowTaskbar = fGetINISetting(location, "showTaskbar", PzGSettingsFile)
        
        PzGGaugeSize = fGetINISetting(location, "gaugeSize", PzGSettingsFile)
        PzGScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", PzGSettingsFile)
        'PzGWidgetSkew = fGetINISetting(location, "widgetSkew", PzGSettingsFile)
        
        ' position
        PzGAspectHidden = fGetINISetting(location, "aspectHidden", PzGSettingsFile)
        PzGWidgetPosition = fGetINISetting(location, "widgetPosition", PzGSettingsFile)
        PzGWidgetLandscape = fGetINISetting(location, "widgetLandscape", PzGSettingsFile)
        PzGWidgetPortrait = fGetINISetting(location, "widgetPortrait", PzGSettingsFile)
        PzGLandscapeFormHoffset = fGetINISetting(location, "landscapeHoffset", PzGSettingsFile)
        PzGLandscapeFormVoffset = fGetINISetting(location, "landscapeYoffset", PzGSettingsFile)
        PzGPortraitHoffset = fGetINISetting(location, "portraitHoffset", PzGSettingsFile)
        PzGPortraitYoffset = fGetINISetting(location, "portraitYoffset", PzGSettingsFile)
        PzGvLocationPercPrefValue = fGetINISetting(location, "vLocationPercPrefValue", PzGSettingsFile)
        PzGhLocationPercPrefValue = fGetINISetting(location, "hLocationPercPrefValue", PzGSettingsFile)

        ' font
        PzGPrefsFont = fGetINISetting(location, "prefsFont", PzGSettingsFile)
        PzGPrefsFontSize = fGetINISetting(location, "prefsFontSize", PzGSettingsFile)
        PzGPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", PzGSettingsFile)
        PzGPrefsFontColour = fGetINISetting(location, "prefsFontColour", PzGSettingsFile)
        
        ' sound
        PzGEnableSounds = fGetINISetting(location, "enableSounds", PzGSettingsFile)
        
        ' development
        PzGDebug = fGetINISetting(location, "debug", PzGSettingsFile)
        PzGDblClickCommand = fGetINISetting(location, "dblClickCommand", PzGSettingsFile)
        PzGOpenFile = fGetINISetting(location, "openFile", PzGSettingsFile)
        PzGDefaultEditor = fGetINISetting(location, "defaultEditor", PzGSettingsFile)
        
        ' other
        PzGMaximiseFormX = fGetINISetting("Software\PzStopwatch", "maximiseFormX", PzGSettingsFile)
        PzGMaximiseFormY = fGetINISetting("Software\PzStopwatch", "maximiseFormY", PzGSettingsFile)
        PzGLastSelectedTab = fGetINISetting(location, "lastSelectedTab", PzGSettingsFile)
        PzGSkinTheme = fGetINISetting(location, "skinTheme", PzGSettingsFile)
        
        ' window
        PzGWindowLevel = fGetINISetting(location, "windowLevel", PzGSettingsFile)
        PzGPreventDragging = fGetINISetting(location, "preventDragging", PzGSettingsFile)
        PzGOpacity = fGetINISetting(location, "opacity", PzGSettingsFile)
        
        ' we do not want the widget to hide at startup
        'PzGWidgetHidden = fGetINISetting(location, "widgetHidden", PzGSettingsFile)
        PzGWidgetHidden = "0"
        
        PzGHidingTime = fGetINISetting(location, "hidingTime", PzGSettingsFile)
        PzGIgnoreMouse = fGetINISetting(location, "ignoreMouse", PzGSettingsFile)
         
        PzGFirstTimeRun = fGetINISetting(location, "firstTimeRun", PzGSettingsFile)
        
    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
        If PzGGaugeFunctions = vbNullString Then PzGGaugeFunctions = "1" ' always turn
'        If PzGAnimationInterval = vbNullString Then PzGAnimationInterval = "130"
        If PzGStartup = vbNullString Then PzGStartup = "1"
        
        If PzGClockFaceSwitchPref = vbNullString Then PzGClockFaceSwitchPref = "0"
        If PzGMainGaugeTimeZone = vbNullString Then PzGMainGaugeTimeZone = "1"
        If PzGMainDaylightSaving = vbNullString Then PzGMainDaylightSaving = "1"
        If PzGSecondaryGaugeTimeZone = vbNullString Then PzGSecondaryGaugeTimeZone = "1"
        If PzGSecondaryDaylightSaving = vbNullString Then PzGSecondaryDaylightSaving = "1"

        ' Config
        If PzGEnableTooltips = vbNullString Then PzGEnableTooltips = "1"
        If PzGEnableBalloonTooltips = vbNullString Then PzGEnableBalloonTooltips = "1"
        If PzGShowTaskbar = vbNullString Then PzGShowTaskbar = "0"
        
        If PzGGaugeSize = vbNullString Then PzGGaugeSize = "25"
        If PzGScrollWheelDirection = vbNullString Then PzGScrollWheelDirection = "up"
               
        ' fonts
        If PzGPrefsFont = vbNullString Then PzGPrefsFont = "times new roman" 'prefsFont", PzGSettingsFile)
        If PzGPrefsFontSize = vbNullString Then PzGPrefsFontSize = "8" 'prefsFontSize", PzGSettingsFile)
        If PzGPrefsFontItalics = vbNullString Then PzGPrefsFontItalics = "false"
        If PzGPrefsFontColour = vbNullString Then PzGPrefsFontColour = "0"

        ' sounds
        If PzGEnableSounds = vbNullString Then PzGEnableSounds = "1"

        ' position
        If PzGAspectHidden = vbNullString Then PzGAspectHidden = "0"
        If PzGWidgetPosition = vbNullString Then PzGWidgetPosition = "0"
        If PzGWidgetLandscape = vbNullString Then PzGWidgetLandscape = "0"
        If PzGWidgetPortrait = vbNullString Then PzGWidgetPortrait = "0"
        If PzGLandscapeFormHoffset = vbNullString Then PzGLandscapeFormHoffset = vbNullString
        If PzGLandscapeFormVoffset = vbNullString Then PzGLandscapeFormVoffset = vbNullString
        If PzGPortraitHoffset = vbNullString Then PzGPortraitHoffset = vbNullString
        If PzGPortraitYoffset = vbNullString Then PzGPortraitYoffset = vbNullString
        If PzGvLocationPercPrefValue = vbNullString Then PzGvLocationPercPrefValue = vbNullString
        If PzGhLocationPercPrefValue = vbNullString Then PzGhLocationPercPrefValue = vbNullString
                
        ' development
        If PzGDebug = vbNullString Then PzGDebug = "0"
        If PzGDblClickCommand = vbNullString Then PzGDblClickCommand = vbNullString
        If PzGOpenFile = vbNullString Then PzGOpenFile = vbNullString
        If PzGDefaultEditor = vbNullString Then PzGDefaultEditor = vbNullString
        If PzGPreventDragging = vbNullString Then PzGPreventDragging = "0"
        
        ' window
        If PzGWindowLevel = vbNullString Then PzGWindowLevel = "1" 'WindowLevel", PzGSettingsFile)
        If PzGOpacity = vbNullString Then PzGOpacity = "100"
        If PzGWidgetHidden = vbNullString Then PzGWidgetHidden = "0"
        If PzGHidingTime = vbNullString Then PzGHidingTime = "0"
        If PzGIgnoreMouse = vbNullString Then PzGIgnoreMouse = "0"
        
        ' other
        If PzGFirstTimeRun = vbNullString Then PzGFirstTimeRun = "true"
        If PzGLastSelectedTab = vbNullString Then PzGLastSelectedTab = "general"
        If PzGSkinTheme = vbNullString Then PzGSkinTheme = "dark"
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If debugflg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    PzGSettingsDir = fSpecialFolder(feUserAppData) & "\PzStopwatch" ' just for this user alone
    PzGSettingsFile = PzGSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(PzGSettingsDir) Then
        MkDir PzGSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(PzGSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open PzGSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
    oldPzGSettingsModificationTime = FileDateTime(PzGSettingsFile)

    frmTimer.rotationTimer.Enabled = True
    frmTimer.settingsTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If PzGHidingTime = "0" Then minutesToHide = 1
    If PzGHidingTime = "1" Then minutesToHide = 5
    If PzGHidingTime = "2" Then minutesToHide = 10
    If PzGHidingTime = "3" Then minutesToHide = 20
    If PzGHidingTime = "4" Then minutesToHide = 30
    If PzGHidingTime = "5" Then minutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createAboutFormOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createAboutFormOnCurrentDisplay()
    On Error GoTo createAboutFormOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      fMain.initAndShowAboutForm .WorkLeft, .WorkTop, 1000, 1000, "Panzer Earth Gauge"
    End With

    On Error GoTo 0
    Exit Sub

createAboutFormOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createAboutFormOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        PzGUnhide = "true"
        sPutINISetting "Software\PzStopwatch", "unhide", PzGUnhide, PzGSettingsFile
        Call thisForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadExcludePathCollection
' Author    : beededea
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadExcludePathCollection_Error

        With fAlpha.collPSDNonUIElements ' the exclude list
          .Add Empty, "stopwatch/face/swSecondHand" 'arrow-hand-top
          .Add Empty, "stopwatch/face/swMinuteHand" 'arrow-hand-right
          .Add Empty, "stopwatch/face/swHourHand"   'arrow-hand-bottom
          
          .Add Empty, "stopwatch/face/hourShadow"   'clock-hand-hours-shadow
          .Add Empty, "stopwatch/face/hourHand"     'clock-hand-hours
         
          .Add Empty, "stopwatch/face/minuteShadow" 'clock-hand-minutes-shadow
          .Add Empty, "stopwatch/face/minuteHand"   'clock-hand-minutes
    
          .Add Empty, "stopwatch/face/secondShadow" 'clock-hand-seconds-shadow
          .Add Empty, "stopwatch/face/secondHand"   'clock-hand-seconds
     
          .Add Empty, "stopwatch/bigReflection"     'all reflections
          .Add Empty, "stopwatch/windowReflection"

        End With

   On Error GoTo 0
   Exit Sub

loadExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadExcludePathCollection of Module modMain"

End Sub
