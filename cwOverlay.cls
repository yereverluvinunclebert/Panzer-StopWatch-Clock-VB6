VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'a "fully click-through" Widget, which allows for "all kind of things to render "on top of the PSD-generated Widgets"

Public OverlayImages As cCollection, cxPerc As Single, cyPerc As Single 'these Public Props are set in the Parent (here, cfAlpha)

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrClock As cTimer, HDeg As Double, MDeg As Double, SDeg As Double
Attribute tmrClock.VB_VarHelpID = -1

Private mTicking As Boolean

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.BackColor = -1  'we don't use any BackColor here
      W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
  Set tmrClock = New_c.Timer(200, True)
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property

Private Sub W_HitTest(ByVal X As Single, ByVal Y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
  HitResultHit = False 'for this Overlay-Widget, we make everything "fully click-through"
End Sub


Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   Dim i As Long, Arr, Key As String, Pat As cCairoPattern
   For i = 0 To OverlayImages.Count - 1
     Arr = OverlayImages.ItemByIndex(i)
     Key = OverlayImages.KeyByIndex(i)
     
     If IsArray(Arr) Then
        CC.Save
           Set Pat = Cairo.ImageList(Key).CreateSurfacePattern 'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
           Select Case Key
             Case "set 1/face/layer 40", "set 1/face/layer 32" 'hours-shadow- and hour-hand
                 Set Pat.Matrix = Pat.Matrix.TranslateCoords(19, 160) 'shift to the "rotation-point" within the Pattern
                 CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                 CC.RotateDrawingsDeg -7 + HDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
   
             Case "set 1/face/layer 38", "set 1/face/layer 39" 'minutes-shadow- and minutes-hand
                 Set Pat.Matrix = Pat.Matrix.TranslateCoords(44, 105) 'shift to the "rotation-point" within the Pattern
                 CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                 CC.RotateDrawingsDeg -61 + MDeg
                 
              Case "set 1/face/layer 33", "set 1/face/layer 34" 'seconds-shadow- and seconds-hand
                 Set Pat.Matrix = Pat.Matrix.TranslateCoords(60, 160) 'shift to the "rotation-point" within the Pattern
                 CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                 CC.RotateDrawingsDeg -50 + SDeg
  
             Case Else 'plain output for everything else ... this includes the reflection-layers,
                       'as well as the little arrow-hands which are not dynamically handled currently
                Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
           End Select
           CC.Paint Arr(2), Pat 'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
        CC.Restore
     End If
   Next
End Sub
 
Private Sub tmrClock_Timer() 'update the private "Degree-Variables" for Hour-, Minute- and Second-Hands + Refresh
  Static T As Date: T = Time
  HDeg = 30 * (Hour(T) Mod 12 + Minute(T) / 60)
  MDeg = 6 * (Minute(T) + Second(T) / 60)
  If SDeg <> 6 * Second(T) Then SDeg = 6 * Second(T) Else Exit Sub 'only refresh, when SDeg was changed to a new value
  W.Parent.Refresh
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Ticking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ticking() As Boolean
   On Error GoTo TickingGet_Error

  Ticking = mTicking

   On Error GoTo 0
   Exit Property

TickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Ticking of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
   On Error GoTo TickingLet_Error

  If mTicking <> newValue Then mTicking = newValue Else Exit Property
    If mTicking Then
        Set tmrClock = New_c.Timer(200, True)
    Else
        Set tmrClock = Nothing
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwGlobe"
End Property ' the earth Ticking flag
