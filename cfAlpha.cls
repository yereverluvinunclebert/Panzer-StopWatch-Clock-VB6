VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents gaugeForm As cWidgetForm
Attribute gaugeForm.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long
Private DragWidgetKey As String

Private mGlowing As Boolean
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

    mGlowing = False ' when the alpha for the overlays arrives this might be usable
    Set collPSDNonUIElements = New_c.Collection(False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfAlpha"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlay class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitFromPSD(PSD_FileNameOrByteArray, Optional WidgetKeyForDragging As String, Optional WidgetKeyForClose As String)
    
    Const AlphaWithTaskbarEntry = 6
    Const AlphaNoTaskbarEntry = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set gaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), "Panzer Stopwatch Gauge", True, 1, 1)
        gaugeForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
      PSDWidth = .Width
      PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
      DragWidgetKey = LCase$(WidgetKeyForDragging) 'store the passed WidgetKey (for the Widget, which allows Form-Dragging on MouseDown)
    
    For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
      If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
         If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
            
            'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
            Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
            
            If collPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                collPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), Val(PzGOpacity) / 100) 'here we update the so far empty slots with the PSD-offsets
               
            Else 'everything not in the exclude-list will get its own Widget-instance
               Set W = gaugeForm.Widgets.Add(New_W("cwAlphaImg"), LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                   W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                   'W.Alpha = .LayerAlphaPercent(i) 'important (to transport the PSD-LayerAlpha to this Widget-instance as well)
                   ' W.ToolTip = W.Key 'can be left out (or one can set it on specific Widgets, as shown in the Select Case below)
                   W.Alpha = Val(PzGOpacity) / 100
                   
                ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
                
                ' Task: Add all layers that provide a UI function below and assign properties as required wip
                
                ' set the characteristics of the interactive areas
               Select Case LCase$(W.Key) 'demonstrates, how to influence certain Widget-Props specifically (depending on the current LayerPath-Key)
                ' set the characteristics of the indentified drag area
                Case DragWidgetKey 'the "special DragWidgetKey" we've passed from the outside into this init-routine
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_SIZEALL
                    
                    W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."

                Case "stopwatch/face/housing/helpbutton" 'three Widget-Keys, picked for "special click-behaviour"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press for a little help."
                    Else
                        W.ToolTip = vbNullString
                    End If
                    
                Case "stopwatch/face/housing/startbutton"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press to zero."
                    Else
                        W.ToolTip = vbNullString
                    End If
                    
                Case "stopwatch/face/housing/stopbutton"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                   If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press to continue the timing run. This button is active only when a timing run has been paused."
                    Else
                        W.ToolTip = vbNullString
                    End If

                Case "stopwatch/face/housing/switchfacesbutton"
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press to switch the dial faces."
                    Else
                        W.ToolTip = vbNullString
                    End If
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                    
                Case "stopwatch/face/housing/lockbutton"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press to lock the widget in place"
                    Else
                        W.ToolTip = vbNullString
                    End If
                    
                Case "stopwatch/face/housing/prefsbutton"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6

                    W.MousePointer = IDC_HAND
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Press to open the widget preferences"
                    Else
                        W.ToolTip = vbNullString
                    End If
                    
                Case "stopwatch/face/housing/tickbutton"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Choose smooth movement or regular ticks"
                    Else
                        W.ToolTip = vbNullString
                    End If
                 Case "stopwatch/face/watchface"
                    W.HoverColor = -1 ' set the hover colour to grey - ' task: might change this later when new ver or RC6 arrives
                    W.MousePointer = IDC_SIZEALL
                    If PzGEnableTooltips = "1" Then
                        W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                    Else
                        W.ToolTip = vbNullString
                    End If
                   
                Case Else 'all other added Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                    W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                    W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                    W.MousePointer = IDC_SIZEALL
                    
               End Select
               
               ' other (unused) widgets created from the layers shown here for reference
               
'                .Add Empty, "stopwatch/face/mki for vb6"
'                .Add Empty, "stopwatch/face/panzer"
'
'                .Add Empty, "stopwatch/face/swhourface"
'                .Add Empty, "stopwatch/face/swminuteface"
'                .Add Empty, "stopwatch/face/swsecondface"
'
'                .Add Empty, "stopwatch/face/4digit"
'                .Add Empty, "stopwatch/face/2digit"
'                .Add Empty, "stopwatch/face/eagle"
'                .Add Empty, "stopwatch/face/made in germany"
'                .Add Empty, "stopwatch/face/italo fontana"
'
'                .Add Empty, "stopwatch/face/darkfacebackground"
'                .Add Empty, "stopwatch/face/watchface"
               
            End If
         End If
      End If
      
    Next
  End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
  
  Set overlayWidget = gaugeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, PSDWidth, PSDHeight)
  Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
      overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
      overlayWidget.cyPerc = 0.4727272
  
  ' Task: Move these to sub main function
'  gaugeForm.show 'we are through with the initializations - and show the Form
'  gaugeForm.Move FX, FY
'  AdjustZoom IIf(FZ, FZ, 1)

   On Error GoTo 0
   Exit Sub

InitFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFromPSD of Class Module cfAlpha"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
   On Error GoTo gaugeForm_BubblingEvent_Error

    If PzGIgnoreMouse = "1" Then Exit Sub
    If EventName = "W_Click" Then
    Select Case LCase$(Sender.Widget.Key)
    
        Case "stopwatch/face/housing/helpbutton" 'three Widget-Keys, picked for "special click-behaviour"
            Call helpbutton_click_event
        Case "stopwatch/face/housing/startbutton"
             Call startbutton_click_event
        Case "stopwatch/face/housing/stopbutton"
             Call stopbutton_click_event
        Case "stopwatch/face/housing/switchfacesbutton"
             Call switchfacesbutton_click_event
        Case "stopwatch/face/housing/lockbutton"
             Call lockbutton_click_event
        Case "stopwatch/face/housing/prefsbutton"
             Call prefsbutton_click_event
        Case "stopwatch/face/housing/tickbutton"
             Call tickbutton_click_event
        Case "stopwatch/face/watchface"
             Call watchface_click_event
        Case Else
            Debug.Print Sender.Widget.Key
    End Select
  End If

   On Error GoTo 0
   Exit Sub

gaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_BubblingEvent of Class Module cfAlpha"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_click_event()
   On Error GoTo helpbutton_click_event_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub
   Call helpSplash
   

   On Error GoTo 0
   Exit Sub

helpbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click_event of Class Module cfAlpha"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : startbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_click_event()
   On Error GoTo startbutton_click_event_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub
    MsgBox "start!"

   On Error GoTo 0
   Exit Sub

startbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : stopbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_click_event()
   On Error GoTo stopbutton_click_event_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub
    MsgBox "stop!"

   On Error GoTo 0
   Exit Sub

stopbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : switchfacesbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub switchfacesbutton_click_event()
   On Error GoTo switchfacesbutton_click_event_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub

   On Error GoTo 0
   Exit Sub

switchfacesbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : lockbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_click_event()
   On Error GoTo lockbutton_click_event_Error
   
    '    If PzGIgnoreMouse = "1" Then Exit Sub
    
    Call lockWidget
    
   On Error GoTo 0
   Exit Sub

lockbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_click_event()
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo prefsbutton_click_event_Error
    
    '    If PzGIgnoreMouse = "1" Then Exit Sub

    fileToPlay = "winding.wav"
    If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
   Call makeProgramPreferencesAvailable

   On Error GoTo 0
   Exit Sub

prefsbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : tickbutton_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tickbutton_click_event()
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo tickbutton_click_event_Error
    
    '    If PzGIgnoreMouse = "1" Then Exit Sub
    
    If overlayWidget.Ticking = True Then
        fileToPlay = "ting.wav"
        If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        Call SwitchOff
    Else
        Call TurnFunctionsOn
    End If

   On Error GoTo 0
   Exit Sub

tickbutton_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_click_event of Class Module cfAlpha"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : watchface_click_event
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub watchface_click_event()
   On Error GoTo watchface_click_event_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub

    MsgBox "face!"

   On Error GoTo 0
   Exit Sub

watchface_click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure watchface_click_event of Class Module cfAlpha"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub gaugeForm_DblClick()

    On Error GoTo gaugeForm_DblClick_Error

    '    If PzGIgnoreMouse = "1" Then Exit Sub

    If LTrim$(PzGDblClickCommand) = vbNullString Then Exit Sub
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fAlpha.gaugeForm.hwnd, "Open", PzGOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(fAlpha.gaugeForm.hwnd, "Open", PzGDblClickCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

gaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef shift As Integer)
    On Error GoTo gaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, shift)

    On Error GoTo 0
    Exit Sub

gaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the globe/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    On Error GoTo gaugeForm_MouseUp_Error

    'If overlayWidget.Locked = True Then Exit Sub
    '    If PzGIgnoreMouse = "1" Then Exit Sub
    
    ' no point in saving when prefs are visible.
    If panzerPrefs.IsVisible = False Then ' checks without activating the form - important!.
        Call savePosition
    End If

    'Static x0, y0: If Button Then gaugeForm.Move gaugeForm.Left + X - x0, gaugeForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

gaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo gaugeForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If PzGIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not gaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (X - x0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Left
      FY = (Y - y0) * gaugeForm.WidgetRoot.Zoom + gaugeForm.Top
      gaugeForm.Move FX, FY
    End If

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseMove of Class Module cfAlpha"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
' rotate forward positive = smaller
Private Sub gaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo gaugeForm_MouseWheel_Error

    If PzGScrollWheelDirection = "down" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseWheel of Class Module cfAlpha"
End Sub
''
''---------------------------------------------------------------------------------------
'' Procedure : gaugeForm_ResizeWithDimensionsDIP
'' Author    : beededea
'' Date      : 31/07/2023
'' Purpose   : put any other resizing action in here - unused
''---------------------------------------------------------------------------------------
''
'Private Sub gaugeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
'   On Error GoTo gaugeForm_ResizeWithDimensionsDIP_Error
'
'  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form
'
'   On Error GoTo 0
'   Exit Sub
'
'gaugeForm_ResizeWithDimensionsDIP_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_ResizeWithDimensionsDIP of Class Module cfAlpha"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

  FZ = NewZoom
  If FZ < 0.25 Then FZ = 0.25 Else If FZ > 2 Then FZ = 2 'keep FZ within a sane ZoomInterval
  gaugeForm.WidgetRoot.Zoom = FZ '* gaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
  gaugeForm.Move gaugeForm.Left, gaugeForm.Top, PSDWidth * gaugeForm.WidgetRoot.Zoom, PSDHeight * gaugeForm.WidgetRoot.Zoom

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfAlpha"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If PzGIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If


   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Glowing
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Glowing() As Boolean
   On Error GoTo GlowingGet_Error

  Glowing = mGlowing

   On Error GoTo 0
   Exit Property

GlowingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Glowing of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Glowing
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Glowing(ByVal newValue As Boolean)
   On Error GoTo GlowingLet_Error

  If mGlowing <> newValue Then mGlowing = newValue Else Exit Property
    If mGlowing Then
        'Set tmrClock = New_c.Timer(200, True)
    Else
        'Set tmrClock = Nothing
    End If

   On Error GoTo 0
   Exit Property

GlowingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Glowing of Class Module cwOverlay"
End Property ' the earth Glowing flag

