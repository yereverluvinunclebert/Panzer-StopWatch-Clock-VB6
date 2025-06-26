Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 22/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder


Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Public Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' to set the full window Opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_LAYERED  As Long = &H80000
Private Const GWL_EXSTYLE  As Long = (-20)
Private Const LWA_COLORKEY  As Long = &H1       'to transparent
Private Const LWA_ALPHA  As Long = &H2          'to semi transparent
'------------------------------------------------------ ENDS

' class objects instantiated
Public fMain As New cfMain
Public aboutWidget As cwAbout
Public helpWidget As cwHelp
Public licenceWidget As cwLicence
Public fGauge As New cfGauge
Public overlayWidget As cwOverlay

' any other private vars
Public gblWidgetName As String



'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Program's entry point
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
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    ' = "none"
    gblStartupFlg = True
    gblWidgetName = "UBoat StopWatch"
    thisPSDFullPath = App.path & "\Res\UBoat-StopWatch-VB6.psd"
    
    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    #If TWINBASIC Then
        gblCodingEnvironment = "TwinBasic"
    #Else
        gblCodingEnvironment = "VB6"
    #End If
        
    menuForm.mnuAbout.Caption = "About UBoat StopWatch Cairo " & gblCodingEnvironment & " widget"
       
    ' Load the sounds into numbered buffers ready for playing
    Call loadAsynchSoundFiles
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions

    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    gblClassicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the gauge settings from the new configuration file
    Call readSettingsFile("Software\UBoatStopWatch", gblSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' Set the opacity of the gauge, passing just this one global variable to a public property within the class
    fGauge.Opacity = gblOpacity
    
    ' check first usage via licence acceptance value and then set initial DPI awareness
    Call setAutomaticDPIState(licenceState)

    'load the collection for storing the overlay surfaces with its relevant keys direct from the PSD
    If restart = False Then Call loadExcludePathCollection ' no need to reload the collPSDNonUIElements layer name keys on a reload
    
    ' start the load of the PSD file using the RC6 PSD-Parser.instance
    Call fGauge.InitFromPSD(thisPSDFullPath)  ' no optional close layer as 3rd param
            
    ' initialise and create the three main RC forms (gauge, about and licence) on the current display
    Call createRCFormsOnCurrentDisplay
    
    ' place the form at the saved location and configure all the form elements
    Call makeVisibleFormElements
        
    ' run the functions that are also called at reload time.
    Call adjustMainControls(licenceState) ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Call loadPreferenceForm
    
    ' if the parameter states re-open prefs then shows the prefs
    If extractCommand = "prefs" Then Call makeProgramPreferencesAvailable

    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
    
    ' note the monitor primary at the preferences form_load and store as gblOldgaugeFormMonitorPrimary
    Call identifyPrimaryMonitor
    
    ' make the busy sand timer invisible
    Call hideBusyTimer
    
    ' start the main gauge timer passing the desired status to a public property within the class
    'overlayWidget.TmrGaugeTicking = True
    
    ' end the startup by un-setting the start global flag
    gblStartupFlg = False
        
    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload. Do not move this line.
    #If TWINBASIC Then
        Cairo.WidgetForms.EnterMessageLoop
    #Else
        If restart = False Then Cairo.WidgetForms.EnterMessageLoop
    #End If
        
    ' note: the final act in startup is the form_resize_event that is triggered by the subclassed WM_EXITSIZEMOVE when the form is finally revealed
     
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain at "
    
End Sub
 

'---------------------------------------------------------------------------------------
' Procedure : loadPreferenceForm
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : load the preferences form but don't yet show it, speeds up access to the prefs via the menu
'---------------------------------------------------------------------------------------
'
Private Sub loadPreferenceForm()
        
   On Error GoTo loadPreferenceForm_Error

    If widgetPrefs.IsLoaded = False Then
        Load widgetPrefs
        gblPrefsFormResizedInCode = True
        Call widgetPrefs.PrefsForm_Resize_Event
    End If

   On Error GoTo 0
   Exit Sub

loadPreferenceForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPreferenceForm of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAutomaticDPIState
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : check first usage via licence acceptance value and then set initial DPI awareness
'---------------------------------------------------------------------------------------
'
Private Sub setAutomaticDPIState(ByRef licenceState As Integer)
   On Error GoTo setAutomaticDPIState_Error

    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

   On Error GoTo 0
   Exit Sub

setAutomaticDPIState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAutomaticDPIState of Module modMain"
End Sub
 
'
'---------------------------------------------------------------------------------------
' Procedure : identifyPrimaryMonitor
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : note the monitor primary at the main form_load and store as gblOldgaugeFormMonitorPrimary - will be resampled regularly later and compared
'---------------------------------------------------------------------------------------
'
Private Sub identifyPrimaryMonitor()
    Dim gaugeFormMonitorID As Long: gaugeFormMonitorID = 0
    
    On Error GoTo identifyPrimaryMonitor_Error

    gaugeMonitorStruct = cWidgetFormScreenProperties(fGauge.gaugeForm, gaugeFormMonitorID)
    gblOldgaugeFormMonitorPrimary = gaugeMonitorStruct.IsPrimary

    On Error GoTo 0
    Exit Sub

identifyPrimaryMonitor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyPrimaryMonitor of Module modMain"
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

    If gblFirstTimeRun = "true" Then
        Call makeProgramPreferencesAvailable
        gblFirstTimeRun = "false"
        sPutINISetting "Software\UBoatStopWatch", "firstTimeRun", gblFirstTimeRun, gblSettingsFile
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
    
    gblMonitorCount = 0

    ' general
    gblStartup = vbNullString
    gblGaugeFunctions = vbNullString
    gblGaugeFunctions = vbNullString

    gblSmoothSecondHand = vbNullString
    gblClockFaceSwitchPref = vbNullString
    gblMainGaugeTimeZone = vbNullString
    gblMainDaylightSaving = vbNullString
    gblSecondaryGaugeTimeZone = vbNullString
    gblSecondaryDaylightSaving = vbNullString

    ' config
    gblGaugeTooltips = vbNullString
    gblPrefsTooltips = vbNullString
    'gblEnablePrefsTooltips = vbNullString
    
    gblShowTaskbar = vbNullString
    gblShowHelp = vbNullString
'    gblTogglePendulum = vbNullString
'    gbl24HourGaugeMode = vbNullString
    
    gblDpiAwareness = vbNullString
    
    gblGaugeSize = vbNullString
    gblScrollWheelDirection = vbNullString
'    gblNumericDisplayRotation = vbNullString
    
    ' position
    gblAspectHidden = vbNullString
    gblWidgetPosition = vbNullString
    gblWidgetLandscape = vbNullString
    gblWidgetPortrait = vbNullString
    gblLandscapeFormHoffset = vbNullString
    gblLandscapeFormVoffset = vbNullString
    gblPortraitHoffset = vbNullString
    gblPortraitYoffset = vbNullString
    gblvLocationPercPrefValue = vbNullString
    gblhLocationPercPrefValue = vbNullString
    
    ' sounds
    gblEnableSounds = vbNullString
'    gblEnableTicks = vbNullString
'    gblEnableChimes = vbNullString
'    gblEnableAlarms = vbNullString
'    gblVolumeBoost = vbNullString
    
    ' development
    gblDebug = vbNullString
    gblDblClickCommand = vbNullString
    gblOpenFile = vbNullString
    gblDefaultVB6Editor = vbNullString
    gblDefaultTBEditor = vbNullString
         
    ' font
    gblClockFont = vbNullString
    gblGaugeFont = vbNullString
    gblPrefsFont = vbNullString
    gblPrefsFontSizeHighDPI = vbNullString
    gblPrefsFontSizeLowDPI = vbNullString
    gblPrefsFontItalics = vbNullString
    gblPrefsFontColour = vbNullString
    
    gblDisplayScreenFont = vbNullString
    gblDisplayScreenFontSize = vbNullString
    gblDisplayScreenFontItalics = vbNullString
    gblDisplayScreenFontColour = vbNullString
    
    ' window
    gblWindowLevel = vbNullString
    gblPreventDragging = vbNullString
    gblOpacity = vbNullString
    gblWidgetHidden = vbNullString
    gblHidingTime = vbNullString
    gblIgnoreMouse = vbNullString
    gblMenuOccurred = False ' bool
    gblFirstTimeRun = vbNullString
    gblMultiMonitorResize = vbNullString
    
    ' general storage variables declared
    gblSettingsDir = vbNullString
    gblSettingsFile = vbNullString
    
    gblTrinketsDir = vbNullString
    gblTrinketsFile = vbNullString
    
    gblGaugeHighDpiXPos = vbNullString
    gblGaugeHighDpiYPos = vbNullString
    
    gblGaugeLowDpiXPos = vbNullString
    gblGaugeLowDpiYPos = vbNullString
    
    gblLastSelectedTab = vbNullString
    gblSkinTheme = vbNullString
    
    ' general variables declared
    'toolSettingsFile = vbNullString
    gblClassicThemeCapable = False
    gblStoreThemeColour = 0
    'windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    gblScreenTwipsPerPixelX = 0
    gblScreenTwipsPerPixelY = 0
    gblPhysicalScreenWidthTwips = 0
    gblPhysicalScreenHeightTwips = 0
    gblPhysicalScreenHeightPixels = 0
    gblPhysicalScreenWidthPixels = 0
    
    gblVirtualScreenHeightPixels = 0
    gblVirtualScreenWidthPixels = 0
    
    gblOldPhysicalScreenHeightPixels = 0
    gblOldPhysicalScreenWidthPixels = 0
    
    gblPrefsPrimaryHeightTwips = vbNullString
    gblPrefsSecondaryHeightTwips = vbNullString
    gblGaugePrimaryHeightRatio = vbNullString
    gblGaugeSecondaryHeightRatio = vbNullString
    
    gblMessageAHeightTwips = vbNullString
    gblMessageAWidthTwips = vbNullString
    
    ' key presses
    gblCTRL_1 = False
    gblSHIFT_1 = False
    
    ' other globals
    gblDebugFlg = 0
    gblMinutesToHide = 0
    gblAspectRatio = vbNullString
    gblOldSettingsModificationTime = #1/1/2000 12:00:00 PM#
    gblCodingEnvironment = vbNullString

    'gblTimeAreaClicked = vbNullString
    
   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add image Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    'Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

'    add Resources to the global ImageList that are not being pulled from the PSD directly
    
    Cairo.ImageList.AddImage "about", App.path & "\Resources\images\about.png"
    Cairo.ImageList.AddImage "licence", App.path & "\Resources\images\frame.png"
    Cairo.ImageList.AddImage "help", App.path & "\Resources\images\panzergauge-help.png"
    Cairo.ImageList.AddImage "frmIcon", App.path & "\Resources\images\Icon.png"
    
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain, an image has probably been accidentally deleted from the resources/images folder."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls(Optional ByVal licenceState As Integer)
   Dim thisEditor As String: thisEditor = vbNullString
   Dim bigScreen As Long: bigScreen = 3840
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' initial call just to obtain initial physical screen monitor ID
    Call positionRCFormByMonitorSize
        
    ' if the licenstate is 0 then the program is running for the first time, so pre-size the form to fit larger screens
    If licenceState = 0 Then
        ' the widget displays at 100% at a screen width of 3840 pixels
        If gblPhysicalScreenWidthPixels >= bigScreen Then
            gblGaugeSize = CStr((gblPhysicalScreenWidthPixels / bigScreen) * 100)
        End If
    End If
    
    ' set the initial size
    If gblMonitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
        If gaugeMonitorStruct.IsPrimary = True Then
            Call fGauge.AdjustZoom(Val(gblGaugePrimaryHeightRatio))
        Else
            Call fGauge.AdjustZoom(Val(gblGaugeSecondaryHeightRatio))
        End If
    Else
        fGauge.AdjustZoom Val(gblGaugeSize) / 100
    End If
    
    If gblGaugeFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gblDebug = "1" Then
        #If TWINBASIC Then
            If gblDefaultTBEditor <> vbNullString Then thisEditor = gblDefaultTBEditor
        #Else
            If gblDefaultVB6Editor <> vbNullString Then thisEditor = gblDefaultVB6Editor
        #End If
        
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & thisEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    If gblShowTaskbar = "0" Then
        fGauge.gaugeForm.ShowInTaskbar = False
    Else
        fGauge.gaugeForm.ShowInTaskbar = True
    End If
    
    ' set the visibility and characteristics of the interactive areas
    ' the alpha is already set to zero for all layers found in the PSD, we now turn them back on as we require
    
    gblStopWatchState = 0
    If gblClockFaceSwitchPref = "0" Then
        overlayWidget.FaceMode = "0"
        fGauge.gaugeForm.Widgets("clockface").Widget.Alpha = 0
        fGauge.gaugeForm.Widgets("stopwatchface").Widget.Alpha = Val(gblOpacity) / 100
        fGauge.gaugeForm.Widgets("stopwatchface").Widget.Refresh
    Else
        overlayWidget.FaceMode = "1"
        fGauge.gaugeForm.Widgets("stopwatchface").Widget.Alpha = 0
        fGauge.gaugeForm.Widgets("clockface").Widget.Alpha = Val(gblOpacity) / 100
        fGauge.gaugeForm.Widgets("clockface").Widget.Refresh
        'gblStopWatchState = 0
    End If
        
    If gblGaugeFunctions = "1" Then
        overlayWidget.Ticking = True
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        overlayWidget.Ticking = False
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gblDebug = "1" Then
        #If TWINBASIC Then
            If gblDefaultTBEditor <> vbNullString Then thisEditor = gblDefaultTBEditor
        #Else
            If gblDefaultVB6Editor <> vbNullString Then thisEditor = gblDefaultVB6Editor
        #End If
        
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & thisEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    If gblShowTaskbar = "0" Then
        fGauge.gaugeForm.ShowInTaskbar = False
    Else
        fGauge.gaugeForm.ShowInTaskbar = True
    End If
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fGauge.gaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fGauge.gaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fGauge.gaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fGauge.gaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fGauge.gaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fGauge.gaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fGauge.gaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fGauge.gaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
    
    If gblSmoothSecondHand = "0" Then
        overlayWidget.SmoothSecondHand = False
        fGauge.gaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        overlayWidget.SmoothSecondHand = True
        fGauge.gaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
    End If
        
   ' set the lock state of the gauge
   If gblPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        overlayWidget.Locked = False
        fGauge.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockWidget.Checked = True
        overlayWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fGauge.gaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If
    
    ' determine the time bias
    If gblMainDaylightSaving <> "0" Then
        tzDelta = fObtainDaylightSavings("Main")
        widgetPrefs.txtMainBias = tzDelta
    End If
    
    ' determine the time bias, secondary gauge
    If gblSecondaryDaylightSaving <> "0" Then
        tzDelta1 = fObtainDaylightSavings("Secondary")
        widgetPrefs.txtSecondaryBias = tzDelta1
    End If

    overlayWidget.MyOpacity = Val(gblOpacity) / 100

    ' set the z-ordering of the window
    Call setAlphaFormZordering
    
    ' set the tooltips on the main screen
    Call setRichClientTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime

    If gblMinutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & gblMinutesToHide & " min."
    
    ' refresh the form in order to show the above changes immediately
    fGauge.gaugeForm.Refresh

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain " _
        & " Most likely one of the layers above is named incorrectly."

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

    If Val(gblWindowLevel) = 0 Then
        Call SetWindowPos(fGauge.gaugeForm.hWnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 1 Then
        Call SetWindowPos(fGauge.gaugeForm.hWnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 2 Then
        Call SetWindowPos(fGauge.gaugeForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
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
Public Sub readSettingsFile(ByVal Location As String, ByVal gblSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gblSettingsFile) Then
        
        ' general
        gblStartup = fGetINISetting(Location, "startup", gblSettingsFile)
        gblGaugeFunctions = fGetINISetting(Location, "widgetFunctions", gblSettingsFile)
        gblSmoothSecondHand = fGetINISetting(Location, "smoothSecondHand", gblSettingsFile)
        
 
        gblClockFaceSwitchPref = fGetINISetting(Location, "clockFaceSwitchPref", gblSettingsFile)
        gblMainGaugeTimeZone = fGetINISetting(Location, "mainGaugeTimeZone", gblSettingsFile)
        gblMainDaylightSaving = fGetINISetting(Location, "mainDaylightSaving", gblSettingsFile)
        gblSecondaryGaugeTimeZone = fGetINISetting(Location, "secondaryGaugeTimeZone", gblSettingsFile)
        gblSecondaryDaylightSaving = fGetINISetting(Location, "secondaryDaylightSaving", gblSettingsFile)
       
        ' configuration
        gblGaugeTooltips = fGetINISetting(Location, "gaugeTooltips", gblSettingsFile)
        gblPrefsTooltips = fGetINISetting(Location, "prefsTooltips", gblSettingsFile)
        
        gblShowTaskbar = fGetINISetting(Location, "showTaskbar", gblSettingsFile)
        gblShowHelp = fGetINISetting(Location, "showHelp", gblSettingsFile)
'        gblTogglePendulum = fGetINISetting(Location, "togglePendulum", gblSettingsFile)
'        gbl24HourGaugeMode = fGetINISetting(Location, "24HourGaugeMode", gblSettingsFile)
        gblDpiAwareness = fGetINISetting(Location, "dpiAwareness", gblSettingsFile)
        gblGaugeSize = fGetINISetting(Location, "gaugeSize", gblSettingsFile)
        gblScrollWheelDirection = fGetINISetting(Location, "scrollWheelDirection", gblSettingsFile)
'        gblNumericDisplayRotation = fGetINISetting(Location, "numericDisplayRotation", gblSettingsFile)
        
        ' position
        gblAspectHidden = fGetINISetting(Location, "aspectHidden", gblSettingsFile)
        gblWidgetPosition = fGetINISetting(Location, "widgetPosition", gblSettingsFile)
        gblWidgetLandscape = fGetINISetting(Location, "widgetLandscape", gblSettingsFile)
        gblWidgetPortrait = fGetINISetting(Location, "widgetPortrait", gblSettingsFile)
        gblLandscapeFormHoffset = fGetINISetting(Location, "landscapeHoffset", gblSettingsFile)
        gblLandscapeFormVoffset = fGetINISetting(Location, "landscapeYoffset", gblSettingsFile)
        gblPortraitHoffset = fGetINISetting(Location, "portraitHoffset", gblSettingsFile)
        gblPortraitYoffset = fGetINISetting(Location, "portraitYoffset", gblSettingsFile)
        gblvLocationPercPrefValue = fGetINISetting(Location, "vLocationPercPrefValue", gblSettingsFile)
        gblhLocationPercPrefValue = fGetINISetting(Location, "hLocationPercPrefValue", gblSettingsFile)

        ' font
        gblClockFont = fGetINISetting(Location, "clockFont", gblSettingsFile)
        gblGaugeFont = fGetINISetting(Location, "gaugeFont", gblSettingsFile)
        gblPrefsFont = fGetINISetting(Location, "prefsFont", gblSettingsFile)
        gblPrefsFontSizeHighDPI = fGetINISetting(Location, "prefsFontSizeHighDPI", gblSettingsFile)
        gblPrefsFontSizeLowDPI = fGetINISetting(Location, "prefsFontSizeLowDPI", gblSettingsFile)
        gblPrefsFontItalics = fGetINISetting(Location, "prefsFontItalics", gblSettingsFile)
        gblPrefsFontColour = fGetINISetting(Location, "prefsFontColour", gblSettingsFile)
    
        gblDisplayScreenFont = fGetINISetting(Location, "displayScreenFont", gblSettingsFile)
        gblDisplayScreenFontSize = fGetINISetting(Location, "displayScreenFontSize", gblSettingsFile)
        gblDisplayScreenFontItalics = fGetINISetting(Location, "displayScreenFontItalics", gblSettingsFile)
        gblDisplayScreenFontColour = fGetINISetting(Location, "displayScreenFontColour", gblSettingsFile)
       
        ' sound
'        gblEnableSounds = fGetINISetting(Location, "enableSounds", gblSettingsFile)
'        gblEnableTicks = fGetINISetting(Location, "enableTicks", gblSettingsFile)
'        gblEnableChimes = fGetINISetting(Location, "enableChimes", gblSettingsFile)
'        gblEnableAlarms = fGetINISetting(Location, "enableAlarms", gblSettingsFile)
'        gblVolumeBoost = fGetINISetting(Location, "volumeBoost", gblSettingsFile)
        
        ' development
        gblDebug = fGetINISetting(Location, "debug", gblSettingsFile)
        gblDblClickCommand = fGetINISetting(Location, "dblClickCommand", gblSettingsFile)
        gblOpenFile = fGetINISetting(Location, "openFile", gblSettingsFile)
        gblDefaultVB6Editor = fGetINISetting(Location, "defaultVB6Editor", gblSettingsFile)
        gblDefaultTBEditor = fGetINISetting(Location, "defaultTBEditor", gblSettingsFile)
        
        ' other
        gblGaugeHighDpiXPos = fGetINISetting("Software\UBoatStopWatch", "gaugeHighDpiXPos", gblSettingsFile)
        gblGaugeHighDpiYPos = fGetINISetting("Software\UBoatStopWatch", "gaugeHighDpiYPos", gblSettingsFile)
        gblGaugeLowDpiXPos = fGetINISetting("Software\UBoatStopWatch", "gaugeLowDpiXPos", gblSettingsFile)
        gblGaugeLowDpiYPos = fGetINISetting("Software\UBoatStopWatch", "gaugeLowDpiYPos", gblSettingsFile)
        gblLastSelectedTab = fGetINISetting(Location, "lastSelectedTab", gblSettingsFile)
        gblSkinTheme = fGetINISetting(Location, "skinTheme", gblSettingsFile)
        
        ' window
        gblWindowLevel = fGetINISetting(Location, "windowLevel", gblSettingsFile)
        gblPreventDragging = fGetINISetting(Location, "preventDragging", gblSettingsFile)
        gblOpacity = fGetINISetting(Location, "opacity", gblSettingsFile)
        
        ' we do not want the widget to hide at startup
        gblWidgetHidden = "0"
        
        gblHidingTime = fGetINISetting(Location, "hidingTime", gblSettingsFile)
        gblIgnoreMouse = fGetINISetting(Location, "ignoreMouse", gblSettingsFile)
        gblMultiMonitorResize = fGetINISetting(Location, "multiMonitorResize", gblSettingsFile)
        gblFirstTimeRun = fGetINISetting(Location, "firstTimeRun", gblSettingsFile)

                           
        gblGaugeSecondaryHeightRatio = fGetINISetting(Location, "gaugeSecondaryHeightRatio", gblSettingsFile)
        gblGaugePrimaryHeightRatio = fGetINISetting(Location, "gaugePrimaryHeightRatio", gblSettingsFile)
        
        gblMessageAHeightTwips = fGetINISetting(Location, "messageAHeightTwips", gblSettingsFile)
        gblMessageAWidthTwips = fGetINISetting(Location, "messageAWidthTwips ", gblSettingsFile)
        
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
        If gblGaugeFunctions = vbNullString Then gblGaugeFunctions = "1" ' always turn
'        If gblAnimationInterval = vbNullString Then gblAnimationInterval = "130"
        If gblStartup = vbNullString Then gblStartup = "1"
        
        If gblSmoothSecondHand = vbNullString Then gblSmoothSecondHand = "0"
        If gblClockFaceSwitchPref = vbNullString Then gblClockFaceSwitchPref = "0"
        If gblMainGaugeTimeZone = vbNullString Then gblMainGaugeTimeZone = "0"
        If gblMainDaylightSaving = vbNullString Then gblMainDaylightSaving = "0"

        If gblSecondaryGaugeTimeZone = vbNullString Then gblSecondaryGaugeTimeZone = "0"
        If gblSecondaryDaylightSaving = vbNullString Then gblSecondaryDaylightSaving = "0"

        ' Configuration
        If gblGaugeTooltips = "False" Then gblGaugeTooltips = "0"
        If gblGaugeTooltips = vbNullString Then gblGaugeTooltips = "0"
        
        'If gblEnablePrefsTooltips = vbNullString Then gblEnablePrefsTooltips = "false"
        If gblPrefsTooltips = "False" Then gblPrefsTooltips = "0"
        If gblPrefsTooltips = vbNullString Then gblPrefsTooltips = "0"
        
        If gblShowTaskbar = vbNullString Then gblShowTaskbar = "0"
        If gblShowHelp = vbNullString Then gblShowHelp = "1"
'        If gblTogglePendulum = vbNullString Then gblTogglePendulum = "0"
'        If gbl24HourGaugeMode = vbNullString Then gbl24HourGaugeMode = "1"
'
        If gblDpiAwareness = vbNullString Then gblDpiAwareness = "0"
        If gblGaugeSize = vbNullString Then gblGaugeSize = "100"
        If gblScrollWheelDirection = vbNullString Then gblScrollWheelDirection = "1"
'        If gblNumericDisplayRotation = vbNullString Then gblNumericDisplayRotation = "1"
               
        ' fonts
        If gblPrefsFont = vbNullString Then gblPrefsFont = "times new roman"
        If gblClockFont = vbNullString Then gblClockFont = gblPrefsFont
        If gblPrefsFontSizeHighDPI = vbNullString Then gblPrefsFontSizeHighDPI = "8"
        If gblPrefsFontSizeLowDPI = vbNullString Then gblPrefsFontSizeLowDPI = "8"
        If gblPrefsFontItalics = vbNullString Then gblPrefsFontItalics = "false"
        If gblPrefsFontColour = vbNullString Then gblPrefsFontColour = "0"

        If gblGaugeFont = vbNullString Then gblGaugeFont = gblPrefsFont

        If gblDisplayScreenFont = vbNullString Then gblDisplayScreenFont = "courier new"
        If gblDisplayScreenFont = "Courier  New" Then gblDisplayScreenFont = "courier new"
        If gblDisplayScreenFontSize = vbNullString Then gblDisplayScreenFontSize = "5"
        If gblDisplayScreenFontItalics = vbNullString Then gblDisplayScreenFontItalics = "false"
        If gblDisplayScreenFontColour = vbNullString Then gblDisplayScreenFontColour = "0"

        ' sounds
        
        If gblEnableSounds = vbNullString Then gblEnableSounds = "1"
'        If gblEnableTicks = vbNullString Then gblEnableTicks = "0"
'        If gblEnableChimes = vbNullString Then gblEnableChimes = "0"
'        If gblEnableAlarms = vbNullString Then gblEnableAlarms = "0"
'        If gblVolumeBoost = vbNullString Then gblVolumeBoost = "0"
        
        
        ' position
        If gblAspectHidden = vbNullString Then gblAspectHidden = "0"
        If gblWidgetPosition = vbNullString Then gblWidgetPosition = "0"
        If gblWidgetLandscape = vbNullString Then gblWidgetLandscape = "0"
        If gblWidgetPortrait = vbNullString Then gblWidgetPortrait = "0"
        If gblLandscapeFormHoffset = vbNullString Then gblLandscapeFormHoffset = vbNullString
        If gblLandscapeFormVoffset = vbNullString Then gblLandscapeFormVoffset = vbNullString
        If gblPortraitHoffset = vbNullString Then gblPortraitHoffset = vbNullString
        If gblPortraitYoffset = vbNullString Then gblPortraitYoffset = vbNullString
        If gblvLocationPercPrefValue = vbNullString Then gblvLocationPercPrefValue = vbNullString
        If gblhLocationPercPrefValue = vbNullString Then gblhLocationPercPrefValue = vbNullString
                
        ' development
        If gblDebug = vbNullString Then gblDebug = "0"
        If gblDblClickCommand = vbNullString And gblFirstTimeRun = "True" Then gblDblClickCommand = "mmsys.cpl"
        If gblOpenFile = vbNullString Then gblOpenFile = vbNullString
        If gblDefaultVB6Editor = vbNullString Then gblDefaultVB6Editor = vbNullString
        If gblDefaultTBEditor = vbNullString Then gblDefaultTBEditor = vbNullString
        
        ' window
        If gblWindowLevel = vbNullString Then gblWindowLevel = "1" 'WindowLevel", gblSettingsFile)
        If gblOpacity = vbNullString Then gblOpacity = "100"
        If gblWidgetHidden = vbNullString Then gblWidgetHidden = "0"
        If gblHidingTime = vbNullString Then gblHidingTime = "0"
        If gblIgnoreMouse = vbNullString Then gblIgnoreMouse = "0"
        If gblPreventDragging = vbNullString Then gblPreventDragging = "0"
        If gblMultiMonitorResize = vbNullString Then gblMultiMonitorResize = "0"
        
        
        ' other
        If gblFirstTimeRun = vbNullString Then gblFirstTimeRun = "true"
        If gblLastSelectedTab = vbNullString Then gblLastSelectedTab = "general"
        If gblSkinTheme = vbNullString Then gblSkinTheme = "dark"
        
        
        ' gauge UI element state
        'If gblSetToggleEnabled = vbNullString Then gblSetToggleEnabled = "False"
'        If gblMuteToggleEnabled = vbNullString Then gblMuteToggleEnabled = "False"
'        If gblPendulumToggleEnabled = vbNullString Then gblPendulumToggleEnabled = "False"
'        If gblPendulumEnabled = vbNullString Then gblPendulumEnabled = "False"
        
        
'        If gblWeekdayToggleEnabled = vbNullString Then gblWeekdayToggleEnabled = "True"
'        If gblDisplayScreenToggleEnabled = vbNullString Then gblDisplayScreenToggleEnabled = "True"
'        If gblTimeMachineToggleEnabled = vbNullString Then gblTimeMachineToggleEnabled = "False"
'        If gblBackToggleEnabled = vbNullString Then gblBackToggleEnabled = "True"
'        If gblAlarmClapperEnabled = vbNullString Then gblAlarmClapperEnabled = "True"
'        If gblChimeClapperEnabled = vbNullString Then gblChimeClapperEnabled = "True"
'        If gblChainEnabled = vbNullString Then gblChainEnabled = "True"
'        If gblCrankEnabled = vbNullString Then gblCrankEnabled = "False"
'        If gblAlarmToggle1Enabled = vbNullString Then gblAlarmToggle1Enabled = "False"
'        If gblAlarmToggle2Enabled = vbNullString Then gblAlarmToggle2Enabled = "False"
'        If gblAlarmToggle3Enabled = vbNullString Then gblAlarmToggle3Enabled = "False"
'        If gblAlarmToggle4Enabled = vbNullString Then gblAlarmToggle4Enabled = "False"
'        If gblAlarmToggle5Enabled = vbNullString Then gblAlarmToggle5Enabled = "False"
'
'        If gblAlarm1Date = vbNullString Then gblAlarm1Date = "Alarm not yet set"
'        If gblAlarm2Date = vbNullString Then gblAlarm2Date = "Alarm not yet set"
'        If gblAlarm3Date = vbNullString Then gblAlarm3Date = "Alarm not yet set"
'        If gblAlarm4Date = vbNullString Then gblAlarm4Date = "Alarm not yet set"
'        If gblAlarm5Date = vbNullString Then gblAlarm5Date = "Alarm not yet set"
        
        If gblGaugePrimaryHeightRatio = "" Then gblGaugePrimaryHeightRatio = "1"
        If gblGaugeSecondaryHeightRatio = "" Then gblGaugeSecondaryHeightRatio = "1"
        
        
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
    
    gblTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    gblTrinketsFile = gblTrinketsDir & "\" & gblWidgetName & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblTrinketsDir) Then
        MkDir gblTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblTrinketsFile For Output As #iFileNo
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
    ''If gblDebugFlg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblSettingsDir = fSpecialFolder(feUserAppData) & "\UBoatStopWatch" ' just for this user alone
    gblSettingsFile = gblSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblSettingsDir) Then
        MkDir gblSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblSettingsFile For Output As #iFileNo
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
    
    gblOldSettingsModificationTime = FileDateTime(gblSettingsFile)

    frmTimer.tmrScreenResolution.Enabled = True
    frmTimer.unhideTimer.Enabled = True

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

    If gblHidingTime = "0" Then gblMinutesToHide = 1
    If gblHidingTime = "1" Then gblMinutesToHide = 5
    If gblHidingTime = "2" Then gblMinutesToHide = 10
    If gblHidingTime = "3" Then gblMinutesToHide = 20
    If gblHidingTime = "4" Then gblMinutesToHide = 30
    If gblHidingTime = "5" Then gblMinutesToHide = 60

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
      Call fMain.initAndCreateAboutForm(.WorkLeft, .WorkTop, 1000, 1000, gblWidgetName)
    End With
    
    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateHelpForm(.WorkLeft, .WorkTop, 1000, 1000, gblWidgetName)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateLicenceForm(.WorkLeft, .WorkTop, 1000, 1000, gblWidgetName)
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
        gblUnhide = "true"
        sPutINISetting "Software\UBoatStopWatch", "unhide", gblUnhide, gblSettingsFile
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

    With fGauge.collPSDNonUIElements ' the exclude list

        .Add Empty, "swsecondhand" 'arrow-hand-top
        .Add Empty, "swhourhand"   'arrow-hand-bottom
        .Add Empty, "swminutehand" 'arrow-hand-right
        
        .Add Empty, "hourshadow"   'clock-hand-hours-shadow
        .Add Empty, "hourhand"     'clock-hand-hours
        
        .Add Empty, "minuteshadow" 'clock-hand-minutes-shadow
        .Add Empty, "minutehand"   'clock-hand-minutes
        
        .Add Empty, "secondshadow" 'clock-hand-seconds-shadow
        .Add Empty, "secondhand"   'clock-hand-seconds

        .Add Empty, "bigreflectioncopy"     'all reflections
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"
        '.Add Empty, "tickbutton"
        
    End With
        

   On Error GoTo 0
   Exit Sub

loadExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadExcludePathCollection of Module modMain"

End Sub



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




'---------------------------------------------------------------------------------------
' Procedure : loadAsynchSoundFiles
' Author    : beededea
' Date      : 27/01/2025
' Purpose   : Load the sounds into numbered buffers ready for playing
'---------------------------------------------------------------------------------------
'
Private Sub loadAsynchSoundFiles()

   On Error GoTo loadAsynchSoundFiles_Error
'
'    LoadSoundFile 1, App.path & "\resources\sounds\belltoll-quiet.wav"
'    LoadSoundFile 2, App.path & "\resources\sounds\belltoll.wav"
'    LoadSoundFile 3, App.path & "\resources\sounds\belltollLong-quiet.wav"
'    LoadSoundFile 4, App.path & "\resources\sounds\belltollLong.wav"
'    LoadSoundFile 5, App.path & "\resources\sounds\fullchime-quiet.wav"
'    LoadSoundFile 6, App.path & "\resources\sounds\fullchime.wav"
'    LoadSoundFile 7, App.path & "\resources\sounds\halfchime-quiet.wav"
'    LoadSoundFile 8, App.path & "\resources\sounds\halfchime.wav"
'    LoadSoundFile 9, App.path & "\resources\sounds\quarterchime-quiet.wav"
'    LoadSoundFile 10, App.path & "\resources\sounds\quarterchime.wav"
'    LoadSoundFile 11, App.path & "\resources\sounds\threequarterchime-quiet.wav"
'    LoadSoundFile 12, App.path & "\resources\sounds\threequarterchime.wav"
'    LoadSoundFile 13, App.path & "\resources\sounds\ticktock-quiet.wav"
'    LoadSoundFile 14, App.path & "\resources\sounds\ticktock.wav"
'    LoadSoundFile 15, App.path & "\resources\sounds\zzzz-quiet.wav"
'    LoadSoundFile 16, App.path & "\resources\sounds\zzzz.wav"
'    LoadSoundFile 17, App.path & "\resources\sounds\till-quiet.wav"
'    LoadSoundFile 18, App.path & "\resources\sounds\till.wav"

   On Error GoTo 0
   Exit Sub

loadAsynchSoundFiles_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAsynchSoundFiles of Module modMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : playAsynchSound
' Author    : beededea
' Date      : 27/01/2025
' Purpose   : requires minimal changes to replace playSound code in the rest of the program
'---------------------------------------------------------------------------------------
'
Public Sub playAsynchSound(ByVal SoundFile As String)

     Dim soundindex As Long: soundindex = 0

     On Error GoTo playAsynchSound_Error

'     If SoundFile = App.path & "\resources\sounds\belltoll-quiet.wav" Then soundindex = 1
'     If SoundFile = App.path & "\resources\sounds\belltoll.wav" Then soundindex = 2
'     If SoundFile = App.path & "\resources\sounds\belltollLong-quiet.wav" Then soundindex = 3
'     If SoundFile = App.path & "\resources\sounds\belltollLong.wav" Then soundindex = 4
'     If SoundFile = App.path & "\resources\sounds\fullchime-quiet.wav" Then soundindex = 5
'     If SoundFile = App.path & "\resources\sounds\fullchime.wav" Then soundindex = 6
'     If SoundFile = App.path & "\resources\sounds\halfchime-quiet.wav" Then soundindex = 7
'     If SoundFile = App.path & "\resources\sounds\halfchime.wav" Then soundindex = 8
'     If SoundFile = App.path & "\resources\sounds\quarterchime-quiet.wav" Then soundindex = 9
'     If SoundFile = App.path & "\resources\sounds\quarterchime.wav" Then soundindex = 10
'     If SoundFile = App.path & "\resources\sounds\threequarterchime-quiet.wav" Then soundindex = 11
'     If SoundFile = App.path & "\resources\sounds\threequarterchime.wav" Then soundindex = 12
'     If SoundFile = App.path & "\resources\sounds\ticktock-quiet.wav" Then soundindex = 13
'     If SoundFile = App.path & "\resources\sounds\ticktock.wav" Then soundindex = 14
'     If SoundFile = App.path & "\resources\sounds\zzzz-quiet.wav" Then soundindex = 15
'     If SoundFile = App.path & "\resources\sounds\zzzz.wav" Then soundindex = 16
'     If SoundFile = App.path & "\resources\sounds\till-quiet.wav" Then soundindex = 17
'     If SoundFile = App.path & "\resources\sounds\till.wav" Then soundindex = 18

     Call playSounds(soundindex) ' writes the wav files (previously stored in a memory buffer) and feeds that buffer to the waveOutWrite API

   On Error GoTo 0
   Exit Sub

playAsynchSound_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure playAsynchSound of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : stopAsynchSound
' Author    : beededea
' Date      : 27/01/2025
' Purpose   : requires minimal changes to previous playSound code
'---------------------------------------------------------------------------------------
'
Public Sub stopAsynchSound(ByVal SoundFile As String)

     Dim soundindex As Long: soundindex = 0

     On Error GoTo stopAsynchSound_Error

'     If SoundFile = App.path & "\resources\sounds\belltoll-quiet.wav" Then soundindex = 1
'     If SoundFile = App.path & "\resources\sounds\belltoll.wav" Then soundindex = 2
'     If SoundFile = App.path & "\resources\sounds\belltollLong-quiet.wav" Then soundindex = 3
'     If SoundFile = App.path & "\resources\sounds\belltollLong.wav" Then soundindex = 4
'     If SoundFile = App.path & "\resources\sounds\fullchime-quiet.wav" Then soundindex = 5
'     If SoundFile = App.path & "\resources\sounds\fullchime.wav" Then soundindex = 6
'     If SoundFile = App.path & "\resources\sounds\halfchime-quiet.wav" Then soundindex = 7
'     If SoundFile = App.path & "\resources\sounds\halfchime.wav" Then soundindex = 8
'     If SoundFile = App.path & "\resources\sounds\quarterchime-quiet.wav" Then soundindex = 9
'     If SoundFile = App.path & "\resources\sounds\quarterchime.wav" Then soundindex = 10
'     If SoundFile = App.path & "\resources\sounds\threequarterchime-quiet.wav" Then soundindex = 11
'     If SoundFile = App.path & "\resources\sounds\threequarterchime.wav" Then soundindex = 12
'     If SoundFile = App.path & "\resources\sounds\ticktock-quiet.wav" Then soundindex = 13
'     If SoundFile = App.path & "\resources\sounds\ticktock.wav" Then soundindex = 14
'     If SoundFile = App.path & "\resources\sounds\zzzz-quiet.wav" Then soundindex = 15
'     If SoundFile = App.path & "\resources\sounds\zzzz.wav" Then soundindex = 16
'     If SoundFile = App.path & "\resources\sounds\till-quiet.wav" Then soundindex = 17
'     If SoundFile = App.path & "\resources\sounds\till.wav" Then soundindex = 18
     
     Call StopSound(soundindex)

   On Error GoTo 0
   Exit Sub

stopAsynchSound_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopAsynchSound of Module modMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : stopAllAsynchSounds
' Author    : beededea
' Date      : 04/02/2025
' Purpose   : ONLY stops any WAV files currently playing in asynchronous mode.
'---------------------------------------------------------------------------------------
'
Public Sub stopAllAsynchSounds()
            
   On Error GoTo stopAllAsynchSounds_Error

'    Call StopSound(1)
'    Call StopSound(2)
'    Call StopSound(3)
'    Call StopSound(4)
'    Call StopSound(5)
'    Call StopSound(6)
'    Call StopSound(7)
'    Call StopSound(8)
'    Call StopSound(9)
'    Call StopSound(10)
'    Call StopSound(12)
'    Call StopSound(13)
'    Call StopSound(14)
'    Call StopSound(15)
'    Call StopSound(16)
'    Call StopSound(17)
'    Call StopSound(18)

   On Error GoTo 0
   Exit Sub

stopAllAsynchSounds_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopAllAsynchSounds of Module modMain"

End Sub

