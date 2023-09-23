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
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfAlpha
'---------------------------------------------------------------------------------------

Option Explicit

Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Private HDeg As Double
Private MDeg As Double
Private SDeg As Double
Private MsDeg As Double

'------------------------------------------------------ STARTS
Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias                    As Long
    StandardName(0 To 63)   As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(0 To 63)   As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'------------------------------------------------------ ENDS


Private mTicking As Boolean ' property Ticking
Private mSmoothSecondHand As Boolean ' property SmoothSecondHand
Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden

 
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo Class_Initialize_Error

    Set W = Cairo.WidgetBase
      W.BackColor = -1  'we don't use any BackColor here, transparent
      W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    intervalValue = 500
      
    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(intervalValue, mTicking)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : Ticking
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
        If mSmoothSecondHand Then
            intervalValue = 100
        Else
            intervalValue = 500
        End If

        tmrClock.Enabled = True
    Else
        tmrClock.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property ' the earth Ticking flag
'---------------------------------------------------------------------------------------
' Property : SmoothSecondHand
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SmoothSecondHand() As Boolean
   On Error GoTo SmoothSecondHandGet_Error

    SmoothSecondHand = mSmoothSecondHand

   On Error GoTo 0
   Exit Property

SmoothSecondHandGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : SmoothSecondHand
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SmoothSecondHand(ByVal newValue As Boolean)
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo SmoothSecondHandLet_Error

    If mSmoothSecondHand <> newValue Then mSmoothSecondHand = newValue Else Exit Property
    
    If mSmoothSecondHand Then
        intervalValue = 100
    Else
        intervalValue = 500
    End If
            
    tmrClock.Interval = intervalValue
    
   On Error GoTo 0
   Exit Property

SmoothSecondHandLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
End Property ' the earth SmoothSecondHand flag

'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

  Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

  If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwGlobe"
End Property

'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

  Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwGlobe"
End Property
'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : beededea
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal Y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim i As Long: i = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtDayNumber As String: txtDayNumber = vbNullString
    Dim txtMonthNumber As String: txtMonthNumber = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
    Dim thisOpacity As Double: thisOpacity = 0
    
    Const cream = &HFFF5F5F5
   
   On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 18
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    thisOpacity = (Val(PzGOpacity) / 100) * 0.75
    
    txtDayNumber = Day(Now)
    txtMonthNumber = Month(Now)
    
    For i = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfAlpha by duplicating collPSDNonUIElements.
     Arr = cCollOverlayImages.ItemByIndex(i)
     Key = cCollOverlayImages.KeyByIndex(i)
     
     If IsArray(Arr) Then
            CC.Save
            'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
            Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
            
            ' for specific keys perform rotation actions, clock hands &c
            Select Case Key
                Case "stopwatch/face/hourshadow", "stopwatch/face/hourhand" 'hours-shadow- and hour-hand
                    
                    ' draw the date and day in the windows just once to retain opacity
                    If textDrawn = False Then ' has to be here to be underneath  all the hands and only run once
                        CC.DrawText 260, 105, 200, 200, txtDayNumber, , , 2, True, , thisOpacity
                        CC.DrawText 461, 309, 200, 200, txtMonthNumber, , , 2, True, , thisOpacity
                        textDrawn = True
                    End If
                    
                   Set Pat.Matrix = Pat.Matrix.TranslateCoords(19, 160) 'shift to the "rotation-point" within the Pattern
                   CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                   CC.RotateDrawingsDeg -7 + HDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                   
                Case "stopwatch/face/minuteshadow", "stopwatch/face/minutehand" 'minutes-shadow- and minutes-hand
                   Set Pat.Matrix = Pat.Matrix.TranslateCoords(44, 105) 'shift to the "rotation-point" within the Pattern
                   CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                   CC.RotateDrawingsDeg -61 + MDeg
                   
                Case "stopwatch/face/secondshadow", "stopwatch/face/secondhand" 'seconds-shadow- and seconds-hand
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(60, 160) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    CC.RotateDrawingsDeg -50 + SDeg + MsDeg
    
                Case Else 'plain output for everything else ... this includes the reflection-layers,
                    'Task: add the little arrow-hands which are not dynamically handled currently
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
            End Select
            CC.Paint Arr(2), Pat 'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
            
            CC.Restore
     End If
   Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : tmrClock_Timer
' Author    : beededea
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrClock_Timer() 'update the private "Degree-Variables" for Hour-, Minute- and Second-Hands + Refresh
    Dim currentLocaltime As SYSTEMTIME
    Static T As Date: T = Time
    
    On Error GoTo tmrClock_Timer_Error
    
    GetSystemTime currentLocaltime

    HDeg = 30 * (Hour(T) Mod 12 + Minute(T) / 60)
    MDeg = 6 * (Minute(T) + Second(T) / 60)

    If mSmoothSecondHand = True Then
        SDeg = 6 * Second(T)

        With currentLocaltime
            MsDeg = 6 * (.wMilliseconds / 1000)
        End With
        
        ' always refresh
    Else
        If SDeg <> 6 * Second(T) Then
            SDeg = 6 * Second(T)
            
            MsDeg = 0
        Else
            Exit Sub  'only refresh when SDeg has changed to a new value
        End If
    End If
    
    W.Parent.Refresh

    On Error GoTo 0
    Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub


