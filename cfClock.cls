VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfGauge
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : This is the class for the gauge form and all the elements upon it.
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
' (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
'  ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' Declare gaugeForm as an RC form to be available throughout by calling fGauge as defined as this class in modMain

Public WithEvents gaugeForm As cWidgetForm
Attribute gaugeForm.VB_VarHelpID = -1

' public vars and objects

Public collPSDNonUIElements As cCollection

Public FX As Long   ' available externally, to allow the gaugeForm to be moved, need to add set/let for these public vars
Public FY As Long
Public FZ As Single

Private mGlowing As Boolean

' do these need to be changed to properties?
Private pauseElapsedTimeSecs As Long ' long means a max of 2,147,483,647 seconds, 24,855 days or 68 years
Private restartElapsedTimeSecs As Long

Private pausedSWSDeg As Long ' pausedSWSDeg = 0

' class private members for property assignment

Private mMuteToggleEnabled As Boolean
Private mTicking As Boolean
Private mShowHelp As Boolean
Private mOpacity As String
Private pvtPSDWidth As Long
Private pvtPSDHeight As Long
Private pvtBusyTimerRotateValue As Integer

' RC widgets

Private WithEvents helpbottom As cWidgetBase
Attribute helpbottom.VB_VarHelpID = -1

'local general vars used only within this class


'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()
   On Error GoTo initialisePrivateVars_Error

    pvtPSDWidth = 0
    pvtPSDHeight = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cfGauge"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
    pvtBusyTimerRotateValue = 1

    Set collPSDNonUIElements = New_c.Collection(False)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfGauge"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt and me
' Date      : 31/07/2023
' Purpose   :  Using a previously populated Cairo imageList with keyed entries  it creates a surface from each and every PSD layer
'              in the PSD file.
'              It adds path X,Y and alpha properties to the excluded collection
'              It then creates widgets with corresponding keynames, locations and sizes as per the original PSD for each layer
'              It creates an instance of the cwOverlay class and populates it with the excluded items that will be rendered in the overlay
'              The overlay comprises items that are non-clickable and will not generate events, ie. animated gauge hands, pendulum &c.
'---------------------------------------------------------------------------------------
'
Public Sub InitFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim I As Long: I = 0
    Dim W As cWidgetBase
    Dim someOpacity As Double: someOpacity = 0
    
    On Error GoTo InitFromPSD_Error
    
    someOpacity = Val(mOpacity) / 100

    'create the Top-Level-Form
    Set gaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), gblWidgetName, True, 1, 1)
        gaugeForm.WidgetRoot.BackColor = -1 ' transparent
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        pvtPSDWidth = .Width
        pvtPSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For I = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(I) Then  ' check this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(I) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(I), .LayerSurface(I)
                    
                    ' check if each layer is in the layer exclude list, if it IS then we add it to a collection for non UI elements (ie. do not create Widgets)
                    If collPSDNonUIElements.Exists(.LayerPath(I)) Then
                        
                        'we add layer info. (used later in cwOverlay) to the excluded layers that will form the overlay.
                        collPSDNonUIElements(.LayerPath(I)) = Array(.LayerX(I), .LayerY(I), someOpacity)  'here we update the so far empty slots with the PSD-offsets
                                    
                    Else
                    
                        'create a widget instance for all layers in the PSD, excluding any layers entered into the exclude-list
                        
                        Set W = gaugeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                        
                        W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget

                        W.Alpha = 0
                    
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblGaugeTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL
                        
                    End If
                End If
            End If
        Next I
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
        
    ' create a overlay here just for a single custom widget that contains just the animated elements.
    Set overlayWidget = gaugeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, pvtPSDWidth, pvtPSDHeight)
    Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for gauge-hand-positioning
        overlayWidget.cyPerc = 0.4727272
    
     ' set the various widgets as WithEvents Variables
'    Set tickbutton = gaugeForm.Widgets("housing/tickbutton").Widget
'    Set helpbutton = gaugeForm.Widgets("housing/helpbutton").Widget
'    Set startbutton = gaugeForm.Widgets("housing/startbutton").Widget
'    Set stopbutton = gaugeForm.Widgets("housing/stopbutton").Widget
'    Set switchfacesbutton = gaugeForm.Widgets("housing/switchfacesbutton").Widget
'    Set lockbutton = gaugeForm.Widgets("housing/lockbutton").Widget
'    Set prefsbutton = gaugeForm.Widgets("housing/prefsbutton").Widget
'    Set surround = gaugeForm.Widgets("housing/surround").Widget
    
    On Error GoTo 0
   Exit Sub

InitFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFromPSD of Class Module cfGauge"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events for all pre-defined PSD layer widgets
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_BubblingEvent(ByRef Sender As Object, ByRef EventName As String, ByRef p1 As Variant, ByRef p2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   
    Dim stdMessage As String: stdMessage = vbNullString
    Dim stdSubject As String: stdSubject = vbNullString
'    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo gaugeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    If Not Sender.Widget.Enabled Then Exit Sub   ' Sender is disabled and bubbling events will be ignored
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
    ' set the balloon tooltips and perform other tasks onMouseMove, RC style tooltips are defined elsewhere
     If EventName = "W_MouseMove" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "housing/lockbutton"
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, gblWidgetName & " Lock Button Help", , , , True
            Case "housing/prefsbutton"
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "Click me to open the preferences utility. The configuration of the widget is performed using the preference utility.", _
                     TTIconInfo, gblWidgetName & " Prefs Button Help", , , , True
            Case "housing/tickbutton"
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "Click here to toggle the tick type from once per second to a more fluid, continuous movement.", _
                     TTIconInfo, gblWidgetName & " Tick Button Help", , , , True
            Case "housing/switchfacesbutton"
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "This button switches the face from the standard clock to the stopwatch and vice versa.", _
                     TTIconInfo, gblWidgetName & " Switch Faces Button Help", , , , True
            Case "housing/stopbutton"
                Call stopbutton_mouseMove(Sender)
            Case "housing/startbutton"
                Call startbutton_mouseMove(Sender)
            Case "housing/helpbutton"
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, gblWidgetName & " Help Button Help", , , , True
            Case Else
                If gblGaugeTooltips = "0" Then CreateToolTip gaugeForm.hWnd, "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. " _
                    & vbCrLf & "You can turn off the balloon tooltips in the preferences.", _
                    TTIconInfo, gblWidgetName & " Configuration Help", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "housing/lockbutton"
                Call lockbutton_mouseDown(Sender)
            Case "housing/prefsbutton"
                Call prefsbutton_mouseDown(Sender)
            Case "housing/tickbutton"
                Call tickbutton_mouseDown(Sender)
            Case "housing/switchfacesbutton"
                 Call switchfacesbutton_mouseDown(Sender)
            Case "housing/stopbutton"
                Call stopbutton_mouseDown(Sender)
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "housing/prefsbutton"
                Call prefsbutton_mouseUp(Sender)
            Case "housing/switchfacesbutton"
                Call switchfacesbutton_mouseUp(Sender)
            Case "housing/stopbutton"
                Call stopbutton_mouseUp(Sender)
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_Click" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "housing/helpbutton"
                 Call helpbutton_click
            Case "housing/startbutton"
                 Call startbutton_click
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
            
   On Error GoTo 0
   Exit Sub

gaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_BubblingEvent of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : startbutton_mouseMove
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_mouseMove(ByRef Sender As Object)
    On Error GoTo startbutton_mouseMove_Error

    If gblGaugeTooltips = "0" Then
        If gblClockFaceSwitchPref = "1" Then
            CreateToolTip gaugeForm.hWnd, "Clock mode, this button restarts all stopped clock functions for this widget, hand movement &c .", _
                TTIconInfo, gblWidgetName & " Start Button Help", , , , True
        Else
        
            If gblStopWatchState = 0 Then
                If gblStopWatchZeroed = False Then
                    CreateToolTip gaugeForm.hWnd, "Stopwatch mode, press to zero the stopwatch dials.", _
                        TTIconInfo, gblWidgetName & " Zero Dials Button Help", , , , True
                Else
                    CreateToolTip gaugeForm.hWnd, "Stopwatch mode, press to start a timing run using the small faces. The main clock hands will be rotated away from the stopwatch faces to avoid obscuring them.", _
                        TTIconInfo, gblWidgetName & " Start Timing Button Help", , , , True
                End If
            ElseIf gblStopWatchState = 1 Then
                CreateToolTip gaugeForm.hWnd, "Stopwatch Mode: Press to pause the stopwatch dials. The timing run continues in the background.", _
                    TTIconInfo, gblWidgetName & " Pause Dials Button Help", , , , True
            ElseIf gblStopWatchState = 2 Then
                CreateToolTip gaugeForm.hWnd, "The stopwatch is now paused: Press to restart the stopwatch dials.", _
                    TTIconInfo, gblWidgetName & " Restart Dials Button Help", , , , True
            End If
        End If
    End If

    On Error GoTo 0
    Exit Sub

startbutton_mouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_mouseMove of Class Module cfGauge"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseMove
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseMove(ByRef Sender As Object)
    On Error GoTo stopbutton_mouseMove_Error

    If gblGaugeTooltips = "0" Then
        If gblClockFaceSwitchPref = "1" Then
            CreateToolTip gaugeForm.hWnd, "Clock mode, this button stops all animated clock functions for this widget, hand movements &c .", _
                TTIconInfo, gblWidgetName & " Stop Button Help", , , , True
        Else
            CreateToolTip gaugeForm.hWnd, "Stopwatch mode, press to continue the timing run. This button is active only when a timing run has been paused.", _
                TTIconInfo, gblWidgetName & " Continue Timing Button Help", , , , True
        End If
    End If

    On Error GoTo 0
    Exit Sub

stopbutton_mouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseMove of Class Module cfGauge"
End Sub

                


'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_click()
   On Error GoTo helpbutton_click_Error

   Call helpSplash

   On Error GoTo 0
   Exit Sub

helpbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfGauge"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : startbutton_click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_click()
    Dim fileToPlay As String: fileToPlay = vbNullString
    
'    Dim pausedSWHDeg As Double: pausedSWHDeg = 0
'    Dim pausedSWMDeg As Double: pausedSWMDeg = 0

    Dim restartSWHDeg As Double: restartSWHDeg = 0
    Dim restartSWMDeg As Double: restartSWHDeg = 0
    Dim newSWSDeg As Double: restartSWHDeg = 0
    Dim newSWSec As Integer: restartSWHDeg = 0
    
    Dim restartTime As Date:  restartTime = #1/1/2000 12:00:00 PM#
    Dim restartSWSDeg As Long: restartSWSDeg = 0
    Dim stopWatchPauseTime As Date: stopWatchPauseTime = #1/1/2000 12:00:00 PM#

    Dim rotationTime As Integer: rotationTime = 0
    Dim rotationSWIntervalMax As Integer: rotationSWIntervalMax = 0
    Dim rotationVarianceSecs As Long: rotationVarianceSecs = 0
    Dim rotationVarianceDegrees As Long: rotationVarianceDegrees = 0
    Dim rotationDegreesPerInterval  As Long: rotationDegreesPerInterval = 0
    
    Dim playCounterTing As Boolean: playCounterTing = False
    Dim playCounter As Boolean: playCounter = False
    Dim playTing As Boolean: playTing = False
    Dim SWRotationDegrees As Long: SWRotationDegrees = 0

    On Error GoTo startbutton_click_Error
    
    rotationTime = 1100 ' ms
    rotationSWIntervalMax = 20 ' ms
    
    If gblClockFaceSwitchPref = "1" Then
        ' if in clock mode just enable the ticking
        Call TurnFunctionsOn
    Else
        Select Case gblStopWatchState
            Case 0 ' start the stopwatch the first time
                playTing = False
                playCounterTing = True
                playCounter = False
                
                gblStopWatchState = 1 ' stopwatch running
                gblStopWatchStartTime = Now()
                overlayWidget.StopWatchTicking = True
                
            Case 1 ' pause the stopwatch
                playTing = True
                playCounterTing = False
                playCounter = False
                
                gblStopWatchState = 2 ' pause the hands in CW_PAINT, stopwatch still running
                
                ' store the current degree and second position values
                stopWatchPauseTime = Now()
                pauseElapsedTimeSecs = DateDiff("s", gblStopWatchStartTime, stopWatchPauseTime)
                pausedSWSDeg = (pauseElapsedTimeSecs * 6) Mod 360

            Case 2 ' restart the stopwatch
                playTing = False
                playCounterTing = False
                playCounter = True
                
                ' back to stopwatch running
                ' but first task will be to zero and rotate the hands

                ' calculate the new position where the hands would be if they were rotating
                restartTime = Now()
                restartElapsedTimeSecs = DateDiff("s", gblStopWatchStartTime, restartTime)
                restartSWSDeg = (restartElapsedTimeSecs * 6) Mod 360
                
                ' rotate from the old stored position to the new one
                ' the amount of degrees variance needs to be covered by the zeroing time sound duration of 1.1 seconds

                rotationVarianceSecs = restartElapsedTimeSecs - pauseElapsedTimeSecs
                rotationVarianceDegrees = (rotationVarianceSecs) Mod 360

                rotationDegreesPerInterval = rotationTime / rotationVarianceDegrees ' ms per degree
                If rotationDegreesPerInterval >= rotationSWIntervalMax Then rotationDegreesPerInterval = rotationSWIntervalMax
                
                overlayWidget.StopWatchTicking = False
                
                If gblEnableSounds = "1" Then
                    fileToPlay = "counter.wav"
                    If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                        playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
                    End If
                End If
            
                If pausedSWSDeg < restartSWSDeg Then
                    SWRotationDegrees = restartSWSDeg - pausedSWSDeg
                Else
                    SWRotationDegrees = 360 - pausedSWSDeg + restartSWSDeg
                End If
                
                ' start the third timer to rotate passing through the interval
                overlayWidget.rotationSWInterval = rotationDegreesPerInterval
                overlayWidget.startSWDegrees = pausedSWSDeg
                overlayWidget.SWRotationDegrees = SWRotationDegrees
                overlayWidget.startSWRotation = True
                
        End Select
    End If
    
    If gblEnableSounds = "1" Then
        If playCounterTing = True Then
            fileToPlay = "counterTing.wav"
            If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
        End If
        
        If playCounter = True Then
            fileToPlay = "counter.wav"
            If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
        End If

        If playTing = True Then
            fileToPlay = "ting.wav"
            If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

startbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseUp(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim stopSWSDeg As Integer: stopSWSDeg = 0
    Dim SWRotationDegrees As Integer: SWRotationDegrees = 0
    
    On Error GoTo stopbutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gblOpacity) / 100
    Sender.Widget.Refresh
        
    If gblEnableSounds = "1" Then
        If gblStopWatchState = 0 Then
            fileToPlay = "stop.wav"
        Else
            fileToPlay = "counterStop.wav"
        End If
        
        If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If
    
    If gblClockFaceSwitchPref = "1" Then
        ' if in clock mode just switch off
        Call SwitchOff
    Else
        
        ' rotate the stopwatch from the old stored position to the new one
        If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
            ' calculate where the hands are now
            stopSWSDeg = overlayWidget.getSWSDeg
     
            overlayWidget.StopWatchTicking = False
            
            If gblEnableSounds = "1" Then
                fileToPlay = "counter.wav"
                If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                    playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
                End If
            End If
        
            SWRotationDegrees = 360 - stopSWSDeg
                            
            ' start the third timer to rotate stopwatch secondhand
            overlayWidget.rotationSWInterval = 20
            overlayWidget.startSWDegrees = stopSWSDeg
            overlayWidget.SWRotationDegrees = SWRotationDegrees
            overlayWidget.startSWRotation = True
    
            gblStopWatchStartTime = Now()
            
            ' rotate the main hands from the stored position to the current time
        
            ' we know the current position, we have the start degrees
            ' calculate the finishing degrees using the current time
            ' calculate the no of degrees to rotate
            ' pass these to the overlay class
            ' trigger the rotation
            
        End If
        

        
    End If

   On Error GoTo 0
   Exit Sub

stopbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseUp of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseDown(ByRef Sender As Object)
   
    On Error GoTo stopbutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh
           
   On Error GoTo 0
   Exit Sub

stopbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : switchfacesbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub switchfacesbutton_mouseDown(ByRef Sender As Object)
    On Error GoTo switchfacesbutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh
    
    If gblClockFaceSwitchPref = "1" Then
        gblClockFaceSwitchPref = "0"
        overlayWidget.FaceMode = "0"
        gaugeForm.Widgets("clockface").Widget.Alpha = 0
        gaugeForm.Widgets("stopwatchface").Widget.Alpha = Val(gblOpacity) / 100
        gaugeForm.Widgets("stopwatchface").Widget.Refresh
    Else
        gblClockFaceSwitchPref = "1"
        overlayWidget.FaceMode = "1"
        gaugeForm.Widgets("stopwatchface").Widget.Alpha = 0
        gaugeForm.Widgets("clockface").Widget.Alpha = Val(gblOpacity) / 100
        gaugeForm.Widgets("clockface").Widget.Refresh
        gblStopWatchState = 0
    End If
    
    sPutINISetting "Software\PanzerStopWatch", "clockFaceSwitchPref", gblClockFaceSwitchPref, gblSettingsFile
    
   On Error GoTo 0
   Exit Sub

switchfacesbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseDown of Class Module cfGauge"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseUp
' Author    : beededea
' Date      : 07/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseUp(ByRef Sender As Object)
    On Error GoTo prefsbutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gblOpacity) / 100
    Sender.Widget.Refresh

    On Error GoTo 0
    Exit Sub

prefsbutton_mouseUp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseUp of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : switchfacesbutton_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub switchfacesbutton_mouseUp(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
    On Error GoTo switchfacesbutton_mouseUp_Error

    Sender.Widget.Alpha = Val(gblOpacity) / 100
    Sender.Widget.Refresh
        
    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   On Error GoTo 0
   Exit Sub

switchfacesbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseUp of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : lockbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_mouseDown(ByRef Sender As Object)
   On Error GoTo lockbutton_mouseDown_Error
    
    If overlayWidget.Locked = False Then
        Sender.Widget.Alpha = 0
    Else
        Sender.Widget.Alpha = Val(gblOpacity) / 100
    End If
    Sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseDown(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo prefsbutton_mouseDown_Error

    Sender.Widget.Alpha = 0
    'Sender.Widget.Visible = False
    Sender.Widget.Refresh ' this removes the delay in the Button depressing
    fileToPlay = "winding.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   'MsgBox "prefsbutton_mouseDown"
   Call makeProgramPreferencesAvailable

   On Error GoTo 0
   Exit Sub

prefsbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfGauge"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : tickbutton_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tickbutton_mouseDown(ByRef Sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo tickbutton_mouseDown_Error
        
    If overlayWidget.SmoothSecondHand = True Then
        overlayWidget.SmoothSecondHand = False
        Sender.Widget.Alpha = Val(gblOpacity) / 100
        gblSmoothSecondHand = "0"
    Else
        overlayWidget.SmoothSecondHand = True
        Sender.Widget.Alpha = 0 ' sender is the referring control
        gblSmoothSecondHand = "1"
        
    End If
    Sender.Widget.Refresh ' this removes the delay in the Button depressing
    
    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    sPutINISetting "Software\PanzerStopWatch", "smoothSecondHand", gblSmoothSecondHand, gblSettingsFile
    
   On Error GoTo 0
   Exit Sub

tickbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_mouseDown of Class Module cfGauge"

End Sub


    



'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : If a double-click command has been selected in the prefs then this will carry it out
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo gaugeForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If gblGaugeFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If gblSHIFT_1 = True Then
        gblSHIFT_1 = False
        Call ShellExecute(gaugeForm.hWnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(gaugeForm.hWnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

gaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_DblClick of Class Module cfGauge"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses and key downs for the main toggles H,A,S,M,P,T,D,W,B & 1-5
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo gaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

gaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_KeyDown of Class Module cfGauge"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :    if this is a multi monitor system then:
'                we note the monitor ID at gaugeForm form_load and store as the gaugeFormMonitorID, for widgetPrefs widgetPrefsMonitorID
'                on gaugeFromMouseUp (drag complete) we sample the monitor ID again -only works, for the prefs form we have to have a 200ms timer that tests for form .left and .top x,y change (no MOVED event and mouse UP does not work on the form itself when the titlebar dragged)
'                if the monitor has changed then sample the physical monitor resolution
'                if the resolution is different then calculate new size proportion
'                alter the size of the form to correspond to the monitor
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    On Error GoTo gaugeForm_MouseUp_Error
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call saveMainRCFormPosition

    Call positionRCFormByMonitorSize
    
    On Error GoTo 0
    Exit Sub

gaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseUp of Class Module cfGauge"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : move the whole form on a mousDown and drag
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo gaugeForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub
    If gblMenuOccurred = True Then
        gblMenuOccurred = False
        Exit Sub
    End If

    If Button = vbLeftButton And Not gaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Left
      FY = (y - y0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Top
      gaugeForm.Move FX, FY
      
      Debug.Print "FX " & FX
      Debug.Print "FY " & FY
      
    End If


   On Error GoTo 0
   Exit Sub

gaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseMove of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : Handles rotation of the mousewheel, with CTRL key pressed, making the widget smaller or larger. By default, rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo gaugeForm_MouseWheel_Error

    Rotation = Rotation / 2

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   'If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)
   If MouseKeys = 8 Then widgetPrefs.GaugeSize = (FZ + RotationDirection / 2400) 'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseWheel of Class Module cfGauge"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Adjust the size of the whole gauge, anchoring the gauge top left, storing the values for later use
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    gaugeForm.WidgetRoot.Zoom = FZ '* gaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    gaugeForm.Move gaugeForm.Left, gaugeForm.Top, pvtPSDWidth * gaugeForm.WidgetRoot.Zoom, pvtPSDHeight * gaugeForm.WidgetRoot.Zoom
    
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblGaugeSize = CStr(NewZoom * 100)
    
    If widgetPrefs.IsLoaded Then widgetPrefs.sliGaugeSize.Value = Val(gblGaugeSize) ' ensure that the prefs size slider matches when zoom is modified without using the slider
    
    If gaugeMonitorStruct.IsPrimary = True Then
        gblGaugePrimaryHeightRatio = CStr(NewZoom)
    Else
        gblGaugeSecondaryHeightRatio = CStr(NewZoom)
    End If

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   : A click on the gauge form, test screen resolution and handle right mouse click for the menu
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo gaugeForm_MouseDown_Error

    'essential code to test for monitor resolution before and after a mouse-down pick up and drag to another monitor
    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.tmrScreenResolution.Interval = 200
    frmTimer.tmrScreenResolution.Enabled = True
    
    If Button = vbRightButton Then
        gblMenuOccurred = True
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gblIgnoreMouse = "1" Then Exit Sub
        ' so we pass it through
    End If

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseDown of Class Module cfGauge"
End Sub

'---------------------------------------------------------------------------------------
' Property  : RotateBusyTimer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : makes previous busy image invisible and next visible simulating rotation
'---------------------------------------------------------------------------------------
'
Public Property Let RotateBusyTimer(ByVal newValue As Boolean)
    On Error GoTo RotateBusyTimerLet_Error
    
    If gblStartupFlg = False Then Exit Property
    
    gaugeForm.Widgets("busy" & pvtBusyTimerRotateValue).Widget.Alpha = 0

    pvtBusyTimerRotateValue = pvtBusyTimerRotateValue + 1
    If pvtBusyTimerRotateValue >= 6 Then pvtBusyTimerRotateValue = 1

    With gaugeForm.Widgets("busy" & pvtBusyTimerRotateValue).Widget
        .Alpha = 1
        .Refresh
    End With
    
   On Error GoTo 0
   Exit Property

RotateBusyTimerLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RotateBusyTimer of Class Module cwoverlay"
End Property



'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
'
' Purpose   : Useful property to determine by value as to whether this form is loaded,
'             allowing external checks to the form to determine whether it is loaded,
'             WITHOUT activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If gaugeForm.WindowState = vbNormal Then
        IsVisible = gaugeForm.Visible
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Class Module cfGauge"
            Resume Next
          End If
    End With
End Property



'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Get Opacity() As String
   On Error GoTo OpacityGet_Error

   Opacity = mOpacity

   On Error GoTo 0
   Exit Property

OpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfGauge"
End Property
'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Let Opacity(ByVal newValue As String)
   On Error GoTo OpacityLet_Error

   If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

OpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfGauge"
End Property





