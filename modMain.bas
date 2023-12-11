Attribute VB_Name = "modMain"
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
' gaugeForm_BubblingEvent ' leaving that here so I can copy/paste to find it

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Public Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' to set the full window Opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_LAYERED  As Long = &H80000
Private Const GWL_EXSTYLE  As Long = (-20)
Private Const LWA_COLORKEY  As Long = &H1       'to transparent
Private Const LWA_ALPHA  As Long = &H2          'to semi transparent
'------------------------------------------------------ ENDS

Public fMain As New cfMain
Public aboutWidget As cwAbout
Public helpWidget As cwHelp
Public licenceWidget As cwLicence

Public revealWidgetTimerCount As Integer
 
Public fAlpha As New cfAlpha
Public overlayWidget As cwOverlay
Public widgetName As String


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
    Dim thisPSDFullPath As String: thisPSDFullPath = vbNullString
    Dim licenceState As Integer: licenceState = 0

    On Error GoTo main_routine_Error
    
    widgetName = "Panzer StopWatch Gauge"
    thisPSDFullPath = App.path & "\Res\tank-clock-mk1.psd"
    fAlpha.FX = 222 'init position- and zoom-values (directly set on Public-Props of the Form-hosting Class)
    fAlpha.FY = 111
    fAlpha.FZ = 0.4
    
    prefsCurrentWidth = 9075
    prefsCurrentHeight = 16450
    
    tzDelta = 0
    tzDelta1 = 0
    
    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\PzStopWatch", PzGSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' check first usage via licence acceptance value and then set initial DPI awareness
    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

    'load the collection for storing the overlay surfaces with its relevant keys direct from the PSD
    If restart = False Then Call loadExcludePathCollection ' no need to reload the collPSDNonUIElements layer name keys on a reload
    
    ' start the load of the PSD file using the RC6 PSD-Parser.instance
    Call fAlpha.InitFromPSD(thisPSDFullPath)  ' no optional close layer as 3rd param

    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
            
    ' initialise and create the three main RC forms on the current display
    Call createRCFormsOnCurrentDisplay
    
    ' check the selected monitor properties
    Call monitorProperties(fAlpha.gaugeForm)  ' might use RC6 for this?
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' run the functions that are also called at reload time.
    Call adjustMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if the parameter states re-open prefs then shows the prefs
    If extractCommand = "prefs" Then
        Call makeProgramPreferencesAvailable
        extractCommand = vbNullString
    End If
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Load panzerPrefs
    
    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
        
    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop
     
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain at "
    
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running, first time run shows prefs
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If PzGFirstTimeRun = "true" Then
        'MsgBox "checkFirstTime"

        Call makeProgramPreferencesAvailable
        PzGFirstTimeRun = "false"
        sPutINISetting "Software\PzStopWatch", "firstTimeRun", PzGFirstTimeRun, PzGSettingsFile
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
    PzGSmoothSecondHand = vbNullString

    PzGClockFaceSwitchPref = vbNullString
    PzGMainGaugeTimeZone = vbNullString
    PzGMainDaylightSaving = vbNullString
    PzGSecondaryGaugeTimeZone = vbNullString
    PzGSecondaryDaylightSaving = vbNullString

    ' config
    PzGEnableTooltips = vbNullString
    PzGEnablePrefsTooltips = vbNullString
    PzGEnableBalloonTooltips = vbNullString
    PzGShowTaskbar = vbNullString
    PzGDpiAwareness = vbNullString
    
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
    PzGClockFont = vbNullString
    PzGPrefsFont = vbNullString
    PzGPrefsFontSizeHighDPI = vbNullString
    PzGPrefsFontSizeLowDPI = vbNullString
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
    
    PzGTrinketsDir = vbNullString
    PzGTrinketsFile = vbNullString
    PzGClockHighDpiXPos = vbNullString
    PzGClockHighDpiYPos = vbNullString
    
    PzGClockLowDpiXPos = vbNullString
    PzGClockLowDpiYPos = vbNullString
    
    PzGLastSelectedTab = vbNullString
    PzGSkinTheme = vbNullString
    
    ' general variables declared
    'toolSettingsFile = vbNullString
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
    debugFlg = 0
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
    'Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

'    add Resources to the global ImageList that are not being pulled from the PSD directly
    
    Cairo.ImageList.AddImage "about", App.path & "\Resources\images\about.png"
    Cairo.ImageList.AddImage "help", App.path & "\Resources\images\panzergauge-help.png"
    Cairo.ImageList.AddImage "licence", App.path & "\Resources\images\frame.png"
    
    ' prefs icons
    
    Cairo.ImageList.AddImage "about-icon-dark", App.path & "\Resources\images\about-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "about-icon-light", App.path & "\Resources\images\about-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-dark", App.path & "\Resources\images\config-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-light", App.path & "\Resources\images\config-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-light", App.path & "\Resources\images\development-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-dark", App.path & "\Resources\images\development-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-dark", App.path & "\Resources\images\general-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-light", App.path & "\Resources\images\general-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light", App.path & "\Resources\images\sounds-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark", App.path & "\Resources\images\sounds-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-light", App.path & "\Resources\images\windows-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark", App.path & "\Resources\images\windows-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-dark", App.path & "\Resources\images\font-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-light", App.path & "\Resources\images\font-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-light", App.path & "\Resources\images\position-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-dark", App.path & "\Resources\images\position-icon-dark-1010.jpg"
    
    Cairo.ImageList.AddImage "general-icon-dark-clicked", App.path & "\Resources\images\general-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-dark-clicked", App.path & "\Resources\images\config-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-dark-clicked", App.path & "\Resources\images\font-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark-clicked", App.path & "\Resources\images\sounds-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-dark-clicked", App.path & "\Resources\images\position-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-dark-clicked", App.path & "\Resources\images\development-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark-clicked", App.path & "\Resources\images\windows-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-dark-clicked", App.path & "\Resources\images\about-icon-dark-600-clicked.jpg"
    
    Cairo.ImageList.AddImage "general-icon-light-clicked", App.path & "\Resources\images\general-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-light-clicked", App.path & "\Resources\images\config-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-light-clicked", App.path & "\Resources\images\font-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light-clicked", App.path & "\Resources\images\sounds-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-light-clicked", App.path & "\Resources\images\position-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-light-clicked", App.path & "\Resources\images\development-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-light-clicked", App.path & "\Resources\images\windows-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-light-clicked", App.path & "\Resources\images\about-icon-light-600-clicked.jpg"
    
   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
    
    
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    fAlpha.AdjustZoom Val(PzGGaugeSize) / 100
    
'    overlayWidget.ZoomDirection = PzGScrollWheelDirection

'PzGClockFaceSwitchPref
'PzGMainGaugeTimeZone
'PzGMainDaylightSaving
'PzGSecondaryGaugeTimeZone
'PzGSecondaryDaylightSaving
    
    If PzGClockFaceSwitchPref = "0" Then
        overlayWidget.FaceMode = "0"
        fAlpha.gaugeForm.Widgets("stopwatchface").Widget.Alpha = Val(PzGOpacity) / 100
        fAlpha.gaugeForm.Widgets("clockface").Widget.Alpha = 0
    Else
        overlayWidget.FaceMode = "1"
        fAlpha.gaugeForm.Widgets("clockface").Widget.Alpha = Val(PzGOpacity) / 100
        fAlpha.gaugeForm.Widgets("stopwatchface").Widget.Alpha = 0
    End If
    
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
    
    
    If PzGShowTaskbar = "0" Then
        fAlpha.gaugeForm.ShowInTaskbar = False
    Else
        fAlpha.gaugeForm.ShowInTaskbar = True
    End If
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fAlpha.gaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
     
    With fAlpha.gaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAlpha.gaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAlpha.gaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fAlpha.gaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fAlpha.gaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fAlpha.gaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100

    End With
    
    If PzGSmoothSecondHand = "0" Then
        overlayWidget.SmoothSecondHand = False
        fAlpha.gaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        overlayWidget.SmoothSecondHand = True
        fAlpha.gaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
    End If
        
    If PzGPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        overlayWidget.Locked = False
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        menuForm.mnuLockWidget.Checked = True
        overlayWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fAlpha.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    ' determine the time bias
    If PzGMainDaylightSaving <> "0" Then
        tzDelta = fObtainDaylightSavings("Main") ' determine the time bias
        panzerPrefs.txtMainBias = tzDelta
    End If
    
    ' determine the time bias
'    If PzGSecondaryDaylightSaving <> "0" Then
'        tzDelta1 = fObtainDaylightSavings("Secondary") ' determine the time bias
'        panzerPrefs.txtSecondaryBias = tzDelta1
'    End If
   
    ' set the z-ordering of the window
    Call setAlphaFormZordering
    
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
' Procedure : setAlphaFormZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setAlphaFormZordering()

   On Error GoTo setAlphaFormZordering_Error

    If Val(PzGWindowLevel) = 0 Then
        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 1 Then
        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 2 Then
        Call SetWindowPos(fAlpha.gaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setAlphaFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlphaFormZordering of Module modMain"
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
        PzGSmoothSecondHand = fGetINISetting(location, "smoothSecondHand", PzGSettingsFile)
        

        PzGClockFaceSwitchPref = fGetINISetting(location, "clockFaceSwitchPref", PzGSettingsFile)
        PzGMainGaugeTimeZone = fGetINISetting(location, "mainGaugeTimeZone", PzGSettingsFile)
        PzGMainDaylightSaving = fGetINISetting(location, "mainDaylightSaving", PzGSettingsFile)
        PzGSecondaryGaugeTimeZone = fGetINISetting(location, "secondaryGaugeTimeZone", PzGSettingsFile)
        PzGSecondaryDaylightSaving = fGetINISetting(location, "secondaryDaylightSaving", PzGSettingsFile)

        ' configuration
        PzGEnableTooltips = fGetINISetting(location, "enableTooltips", PzGSettingsFile)
        PzGEnablePrefsTooltips = fGetINISetting(location, "enablePrefsTooltips", PzGSettingsFile)
        PzGEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", PzGSettingsFile)
        PzGShowTaskbar = fGetINISetting(location, "showTaskbar", PzGSettingsFile)
        PzGDpiAwareness = fGetINISetting(location, "dpiAwareness", PzGSettingsFile)
        
        
        PzGGaugeSize = fGetINISetting(location, "gaugeSize", PzGSettingsFile)
        PzGScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", PzGSettingsFile)
        
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
        PzGClockFont = fGetINISetting(location, "clockFont", PzGSettingsFile)
        PzGPrefsFont = fGetINISetting(location, "prefsFont", PzGSettingsFile)
        PzGPrefsFontSizeHighDPI = fGetINISetting(location, "prefsFontSizeHighDPI", PzGSettingsFile)
        PzGPrefsFontSizeLowDPI = fGetINISetting(location, "prefsFontSizeLowDPI", PzGSettingsFile)
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
        PzGClockHighDpiXPos = fGetINISetting("Software\PzStopWatch", "clockHighDpiXPos", PzGSettingsFile)
        PzGClockHighDpiYPos = fGetINISetting("Software\PzStopWatch", "clockHighDpiYPos", PzGSettingsFile)
        
        PzGClockLowDpiXPos = fGetINISetting("Software\PzStopWatch", "clockLowDpiXPos", PzGSettingsFile)
        PzGClockLowDpiYPos = fGetINISetting("Software\PzStopWatch", "clockLowDpiYPos", PzGSettingsFile)
        
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
        If PzGSmoothSecondHand = vbNullString Then PzGSmoothSecondHand = "0"
        
        If PzGClockFaceSwitchPref = vbNullString Then PzGClockFaceSwitchPref = "0"
        If PzGMainGaugeTimeZone = vbNullString Then PzGMainGaugeTimeZone = "0"
        If PzGMainDaylightSaving = vbNullString Then PzGMainDaylightSaving = "0"

        If PzGSecondaryGaugeTimeZone = vbNullString Then PzGSecondaryGaugeTimeZone = "0"
        If PzGSecondaryDaylightSaving = vbNullString Then PzGSecondaryDaylightSaving = "0"

        ' Configuration
        If PzGEnableTooltips = vbNullString Then PzGEnableTooltips = "0"
        If PzGEnablePrefsTooltips = vbNullString Then PzGEnablePrefsTooltips = "1"
        If PzGEnableBalloonTooltips = vbNullString Then PzGEnableBalloonTooltips = "1"
        If PzGShowTaskbar = vbNullString Then PzGShowTaskbar = "0"
        If PzGDpiAwareness = vbNullString Then PzGDpiAwareness = "0"
        If PzGGaugeSize = vbNullString Then PzGGaugeSize = "25"
        If PzGScrollWheelDirection = vbNullString Then PzGScrollWheelDirection = "1"
               
        ' fonts
        If PzGPrefsFont = vbNullString Then PzGPrefsFont = "times new roman"
        If PzGClockFont = vbNullString Then PzGClockFont = PzGPrefsFont
        If PzGPrefsFontSizeHighDPI = vbNullString Then PzGPrefsFontSizeHighDPI = "8"
        If PzGPrefsFontSizeLowDPI = vbNullString Then PzGPrefsFontSizeLowDPI = "8"
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
        If PzGDblClickCommand = vbNullString Then PzGDblClickCommand = "%systemroot%\system32\timedate.cpl"
        If PzGOpenFile = vbNullString Then PzGOpenFile = vbNullString
        If PzGDefaultEditor = vbNullString Then PzGDefaultEditor = vbNullString
        
        ' window
        If PzGWindowLevel = vbNullString Then PzGWindowLevel = "1" 'WindowLevel", PzGSettingsFile)
        If PzGOpacity = vbNullString Then PzGOpacity = "100"
        If PzGWidgetHidden = vbNullString Then PzGWidgetHidden = "0"
        If PzGHidingTime = vbNullString Then PzGHidingTime = "0"
        If PzGIgnoreMouse = vbNullString Then PzGIgnoreMouse = "0"
        If PzGPreventDragging = vbNullString Then PzGPreventDragging = "0"
        
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
' Procedure : getTrinketsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's entry in the trinkets settings file and assign the app.path
'---------------------------------------------------------------------------------------
'
Private Sub getTrinketsFile()
    On Error GoTo getTrinketsFile_Error
    
    Dim iFileNo As Integer: iFileNo = 0
    
    PzGTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    PzGTrinketsFile = PzGTrinketsDir & "\" & widgetName & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(PzGTrinketsDir) Then
        MkDir PzGTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(PzGTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open PzGTrinketsFile For Output As #iFileNo
        Write #iFileNo, App.path & "\" & App.EXEName & ".exe"
        Write #iFileNo,
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getTrinketsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getTrinketsFile of Form modMain"

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
    
    PzGSettingsDir = fSpecialFolder(feUserAppData) & "\PzStopWatch" ' just for this user alone
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
' Procedure : createRCFormsOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRCFormsOnCurrentDisplay()
    On Error GoTo createRCFormsOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowAboutForm(.WorkLeft, .WorkTop, 1000, 1000, widgetName)
    End With
    
    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowHelpForm(.WorkLeft, .WorkTop, 1000, 1000, widgetName)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowLicenceForm(.WorkLeft, .WorkTop, 1000, 1000, widgetName)
    End With
    
        On Error GoTo 0
    Exit Sub

createRCFormsOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRCFormsOnCurrentDisplay of Module modMain"
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
        sPutINISetting "Software\PzStopWatch", "unhide", PzGUnhide, PzGSettingsFile
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

        .Add Empty, "swsecondhand" 'arrow-hand-top
        .Add Empty, "swhourhand"   'arrow-hand-bottom
        .Add Empty, "swminutehand" 'arrow-hand-right
        
        .Add Empty, "hourshadow"   'clock-hand-hours-shadow
        .Add Empty, "hourhand"     'clock-hand-hours
        
        .Add Empty, "minuteshadow" 'clock-hand-minutes-shadow
        .Add Empty, "minutehand"   'clock-hand-minutes
        
        .Add Empty, "secondshadow" 'clock-hand-seconds-shadow
        .Add Empty, "secondhand"   'clock-hand-seconds

        .Add Empty, "bigreflection Copy"     'all reflections
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"


    End With

   On Error GoTo 0
   Exit Sub

loadExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadExcludePathCollection of Module modMain"

End Sub


''---------------------------------------------------------------------------------------
'' Procedure : ExportPngs
'' Author    : Olaf
'' Date      : 06/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Sub exportPngs(PSD_FileNameOrByteArray, ByVal pngFolder As String)
'   On Error GoTo ExportPngs_Error
'
'  New_c.FSO.EnsurePath pngFolder 'make sure the PngFolder-Path "materializes itself" in the FileSystem
'  New_c.FSO.EnsurePathEndSep pngFolder 'add a backslash to the PngFolder-param (in case it was missing)
'
'  With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
'    Dim i As Long
'    For i = 0 To .LayersCount - 1 'loop over all the Layers in the PSD
'      If .LayerByteSize(i) Then   'this is an Alpha-Surface-Layer with "meat" (and not a group-specification)
'         .LayerSurface(i).WriteContentToPngFile pngFolder & Replace(.LayerPath(i), "/", "_") & ".png"
'      End If
'    Next
'  End With
'
'   On Error GoTo 0
'   Exit Sub
'
'ExportPngs_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportPngs of Module modMain"
'End Sub







     



' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen, see Steamydock for a potential replacement?
'---------------------------------------------------------------------------------------
' Procedure : msgBoxA
' Author    : beededea
' Date      : 20/05/2022
' Purpose   : ans = msgBoxA("main message", vbOKOnly, "title bar message", False)
'---------------------------------------------------------------------------------------
'
Public Function msgBoxA(ByVal msgBoxPrompt As String, Optional ByVal msgButton As VbMsgBoxResult, Optional ByVal msgTitle As String, Optional ByVal msgShowAgainChkBox As Boolean = False, Optional ByRef msgContext As String = "none") As Integer
     
    ' set the defined properties of a form
    On Error GoTo msgBoxA_Error

    frmMessage.propMessage = msgBoxPrompt
    frmMessage.propTitle = msgTitle
    frmMessage.propShowAgainChkBox = msgShowAgainChkBox
    frmMessage.propButtonVal = msgButton
    frmMessage.propMsgContext = msgContext
    Call frmMessage.Display ' run a subroutine in the form that displays the form

    msgBoxA = frmMessage.propReturnedValue

    On Error GoTo 0
    Exit Function

msgBoxA_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure msgBoxA of Module mdlMain"
            Resume Next
          End If
    End With

End Function






