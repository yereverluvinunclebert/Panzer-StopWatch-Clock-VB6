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

Public WithEvents GaugeForm As cWidgetForm
Attribute GaugeForm.VB_VarHelpID = -1
Public PSDExcludePaths As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long, PSDHeight As Long, DragWidgetKey As String ', CloseWidgetKey As String, CloseWidget As cWidgetBase
Attribute PSDWidth.VB_VarHelpID = -1
Private Overlay As cwOverlay
Attribute Overlay.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set PSDExcludePaths = New_c.Collection(False)
End Sub

Public Sub InitFromPSD(PSD_FileNameOrByteArray, Optional WidgetKeyForDragging As String, Optional WidgetKeyForClose As String)
           
  'create the Alpha-capable Top-Level-Form
  Const AlphaWithTaskbarEntry = 6, AlphaNoTaskbarEntry = 7
  Set GaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), "AlphaForm", True, 1, 1)
      GaugeForm.WidgetRoot.BackColor = -1
 
  With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
    PSDWidth = .Width: PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    DragWidgetKey = LCase$(WidgetKeyForDragging) 'store the passed WidgetKey (for the Widget, which allows Form-Dragging on MouseDown)
    'CloseWidgetKey = LCase$(WidgetKeyForClose)   'same here, for the optional CloseWidget (which needs a preloaded ImgResource under the same Key)
 
    'Cairo.ImageList-buffering - and associated Widget-construction (via a loop over all PSD-Layers)
    Dim i As Long, W As cWidgetBase
    For i = 0 To .LayersCount - 1 'loop over all the Layers in the PSD
      If .LayerByteSize(i) Then  'this is an Alpha-Surface-Layer with "meat" (and not a group-specification)
         If .LayerAlphaPercent(i) > 0 Then
            Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i) 'add the current LayerSurface into the global ImageList (using LayerPath as the ImageKey)
            
            If PSDExcludePaths.Exists(.LayerPath(i)) Then 'in case of the "exceptions" (defined in Sub Main), we don't create Widgets, but store only some Layer-Infos (used in cwOverlay)
               PSDExcludePaths(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), .LayerAlphaPercent(i)) 'here we update the so far empty slots with the PSD-offsets
               
            Else 'everything not in the exclude-list will get its own Widget-instance
               Set W = GaugeForm.Widgets.Add(New_W("cwAlphaImg"), LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                   W.ImageKey = W.Key 'W.Key equals ImageList-Key and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                   W.Alpha = .LayerAlphaPercent(i) 'important (to transport the PSD-LayerAlphy to this Widget-instance as well)
                   W.ToolTip = W.Key 'can be left out (or one can set it on specific Widgets, as shown in the Select Case below)
                   
               Select Case LCase$(W.Key) 'demonstrates, how to influence certain Widget-Props specifically (depending on the current LayerPath-Key)
                 Case DragWidgetKey 'the "special DragWidgetKey" we've passed from the outside into this init-routine
                      W.HoverColor = vbGreen
                      W.MousePointer = IDC_SIZEALL
                 Case "set 1/face/set 1/layer 31", "set 1/face/set 1/layer 37", "set 1/face/set 1/layer 37 copy" 'three Widget-Keys, picked for "special click-behaviour"
                      W.HoverColor = vbCyan
                      W.MousePointer = IDC_HAND
      
                 Case Else 'all other added Layer-Widgets will be hovered in vbRed by default
                      'W.HoverColor = -1 'a HoverColor of -1 would make them "non-hoverable" and "fully click-through"
               End Select
            End If
         End If
      End If
    Next
  End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
  
'  If Len(CloseWidgetKey) Then 'in case the Optional InitParam was set, we create an additional cwAlphaImg Widget for closing the Form
'    Set CloseWidget = GaugeForm.Widgets.Add(New_W("cwAlphaImg"), CloseWidgetKey).Widget
'        CloseWidget.ImageKey = CloseWidgetKey 'this cwAlphaImg-instance will render from a normal ImageList-Key (not a PSD-LayerPath-Key)
'        CloseWidget.HoverColor = vbRed
'        CloseWidget.MousePointer = IDC_HAND
'  End If
  
  Set Overlay = GaugeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, PSDWidth, PSDHeight)
  Set Overlay.OverlayImages = PSDExcludePaths 'make the excluded PSD-Surface-Paths known in the Overlay-Widget
      Overlay.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
      Overlay.cyPerc = 0.4727272
  
  GaugeForm.show 'we are through with the initializations - and show the Form
  GaugeForm.Move FX, FY
  AdjustZoom IIf(FZ, FZ, 1)
End Sub

Private Sub GaugeForm_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
  If EventName = "W_Click" Then 'central handling of all Click-Events
    Select Case LCase$(Sender.Widget.Key)
      'Case CloseWidgetKey
        
        'GaugeForm.Unload 'only one specific WidgetKey is handled here so far (click-wise)
      'please add your own cases, based on the PSD-LayerPaths we generally use as Keys here (for the ImageList as well as the Widgets)
      'Case Else:  Debug.Print Sender.Widget.Key
    End Select
  End If
End Sub

Private Sub GaugeForm_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
  If Button = vbLeftButton And Not GaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)
     If StrComp(GaugeForm.ActiveWidget.Key, DragWidgetKey, vbTextCompare) Then Exit Sub 'leave, if the ActiveWidget.Key is not the DragWidgetKey
     FX = (x - x0) * GaugeForm.WidgetRoot.Zoom + GaugeForm.Left
     FY = (y - y0) * GaugeForm.WidgetRoot.Zoom + GaugeForm.Top
     GaugeForm.Move FX, FY
  End If
End Sub

Private Sub GaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
  If MouseKeys = 8 Then AdjustZoom FZ + Rotation / 1200  'change the Zoom only, when the Ctrl-Key is down
End Sub
Private Sub GaugeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form
End Sub

Private Sub AdjustZoom(ByVal NewZoom As Single)
  FZ = NewZoom
  If FZ < 0.25 Then FZ = 0.25 Else If FZ > 2 Then FZ = 2 'keep FZ within a sane ZoomInterval
  GaugeForm.WidgetRoot.Zoom = FZ '* GaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
  GaugeForm.Move GaugeForm.Left, GaugeForm.Top, PSDWidth * GaugeForm.WidgetRoot.Zoom, PSDHeight * GaugeForm.WidgetRoot.Zoom
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub GaugeForm_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwGlobe"
End Sub
