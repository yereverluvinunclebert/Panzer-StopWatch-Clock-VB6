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

', CloseWidgetKey As String, CloseWidget As cWidgetBase

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
 
    ' see Purpose notes above
        
    ' Cairo.ImageList-buffering - and associated Widget-construction (via a loop through all PSD-Layers)
    
    For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
      If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
         If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
            
            'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
            Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
            
            If collPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                collPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), .LayerAlphaPercent(i)) 'here we update the so far empty slots with the PSD-offsets
               
            Else 'everything not in the exclude-list will get its own Widget-instance
               Set W = gaugeForm.Widgets.Add(New_W("cwAlphaImg"), LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                   W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                   W.Alpha = .LayerAlphaPercent(i) 'important (to transport the PSD-LayerAlpha to this Widget-instance as well)
                   'W.ToolTip = W.Key 'can be left out (or one can set it on specific Widgets, as shown in the Select Case below)
                   
                ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
                ' Task: Name the PSD layers correctly
                
                ' Task: Add all layers that provide a UI function below and assign properties as required
                
               Select Case LCase$(W.Key) 'demonstrates, how to influence certain Widget-Props specifically (depending on the current LayerPath-Key)
                ' set the characteristics of the indentified drag area
                Case DragWidgetKey 'the "special DragWidgetKey" we've passed from the outside into this init-routine
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_SIZEALL
                    
                    W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                ' set the characteristics of the interactive areas
                Case "stopwatch/face/housing/helpButton", "stopwatch/face/housing/startButton", "stopwatch/face/housing/stopButton", "stopwatch/face/housing/switchFacesButton" 'three Widget-Keys, picked for "special click-behaviour"
                    W.HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
                    W.MousePointer = IDC_HAND
                Case Else 'all other added Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                    W.HoverColor = -1
               End Select
            End If
         End If
      End If
      
    Next
  End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
  
'  If Len(CloseWidgetKey) Then 'in case the Optional InitParam was set, we create an additional cwAlphaImg Widget for closing the Form
'    Set CloseWidget = gaugeForm.Widgets.Add(New_W("cwAlphaImg"), CloseWidgetKey).Widget
'        CloseWidget.ImageKey = CloseWidgetKey 'this cwAlphaImg-instance will render from a normal ImageList-Key (not a PSD-LayerPath-Key)
'        CloseWidget.HoverColor = vbRed
'        CloseWidget.MousePointer = IDC_HAND
'  End If
  
  Set overlayWidget = gaugeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, PSDWidth, PSDHeight)
  Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
      overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
      overlayWidget.cyPerc = 0.4727272
  
  ' Task: Move these to sub main function
  gaugeForm.show 'we are through with the initializations - and show the Form
  gaugeForm.Move FX, FY
  AdjustZoom IIf(FZ, FZ, 1)

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

  If EventName = "W_Click" Then
    Select Case LCase$(Sender.Widget.Key)
      'Case CloseWidgetKey
      
        ' Task: add each UI element and call a function for each
        
        'gaugeForm.Unload 'only one specific WidgetKey is handled here so far (click-wise)
      'please add your own cases, based on the PSD-LayerPaths we generally use as Keys here (for the ImageList as well as the Widgets)
      'Case Else:  Debug.Print Sender.Widget.Key
    End Select
  End If

   On Error GoTo 0
   Exit Sub

gaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_BubblingEvent of Class Module cfAlpha"
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
    If Button = vbLeftButton And Not gaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)
      'If StrComp(gaugeForm.ActiveWidget.Key, DragWidgetKey, vbTextCompare) Then Exit Sub 'leave, if the ActiveWidget.Key is not the DragWidgetKey
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
'
Private Sub gaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   On Error GoTo gaugeForm_MouseWheel_Error

  If MouseKeys = 8 Then Call AdjustZoom(FZ + Rotation / 1200)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

gaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseWheel of Class Module cfAlpha"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : gaugeForm_ResizeWithDimensionsDIP
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : put any other resizing action in here - unused
'---------------------------------------------------------------------------------------
'
Private Sub gaugeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
   On Error GoTo gaugeForm_ResizeWithDimensionsDIP_Error

  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form

   On Error GoTo 0
   Exit Sub

gaugeForm_ResizeWithDimensionsDIP_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_ResizeWithDimensionsDIP of Class Module cfAlpha"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub AdjustZoom(ByVal NewZoom As Single)
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

