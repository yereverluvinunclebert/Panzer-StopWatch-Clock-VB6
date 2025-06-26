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
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfGauge
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfGauge)
Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Public WithEvents tmrStopWatch As cTimer
Attribute tmrStopWatch.VB_VarHelpID = -1
Public WithEvents tmrSWRotation As cTimer
Attribute tmrSWRotation.VB_VarHelpID = -1
Public WithEvents tmrHandsRotation As cTimer
Attribute tmrHandsRotation.VB_VarHelpID = -1

Private HDeg As Double
Private MDeg As Double
Private HDeg1 As Double
Private MDeg1 As Double
Private SDeg As Double
Private Sec As Integer
Private MsDeg As Double

Private SWHDeg As Double
Private SWMDeg As Double
Private SWSDeg As Double
Private SWSec As Integer

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
Private mStopWatchTicking As Boolean ' property StopWatchTicking
Private mSmoothSecondHand As Boolean ' property SmoothSecondHand
Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden
Private mFaceMode As String ' property FaceMode
Private storeTxtDayNumber As String

Private mMyOpacity As Single


Private mstartSWRotation As Boolean
Private mstartSWDegrees As Long
Private mSWRotationDegrees As Long
Private mrotationSWInterval As Long

Private rotationTimerCount As Integer ' rotationTimerCount = 0


'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   : initialise vars
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()

   On Error GoTo initialisePrivateVars_Error


   ' pvtCurrentDateToDisplay = vbNullString

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize for cwOverlay
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise) set the inital values of private vars, set up the timers
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    intervalValue = 500
'
' SWHDeg = 0
' SWMDeg = 0
' SWSDeg = 0
' SWSec = 0

    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(intervalValue, mTicking)

    Set tmrStopWatch = Nothing
    Set tmrStopWatch = New_c.Timer(intervalValue, False)

    Set tmrSWRotation = Nothing
    Set tmrSWRotation = New_c.Timer(20, False)

    Set tmrHandsRotation = Nothing
    Set tmrHandsRotation = New_c.Timer(20, False)
    
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
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
        tmrClock.Enabled = True
    Else
        tmrClock.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : getSWSDeg
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get getSWSDeg() As Integer
   On Error GoTo getSWSDegGet_Error

   getSWSDeg = SWSDeg

   On Error GoTo 0
   Exit Property

getSWSDegGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property getSWSDeg of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : StopWatchTicking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get StopWatchTicking() As Boolean
   On Error GoTo StopWatchTickingGet_Error

   StopWatchTicking = mStopWatchTicking

   On Error GoTo 0
   Exit Property

StopWatchTickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StopWatchTicking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : StopWatchTicking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let StopWatchTicking(ByVal newValue As Boolean)
    On Error GoTo StopWatchTickingLet_Error

    If mStopWatchTicking <> newValue Then mStopWatchTicking = newValue Else Exit Property
    
    If mStopWatchTicking Then
        tmrStopWatch.Enabled = True
    Else
        tmrStopWatch.Enabled = False
        SWHDeg = 0
        SWMDeg = 0
        SWSDeg = 0
        SWSec = 0
    End If

   On Error GoTo 0
   Exit Property

StopWatchTickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StopWatchTicking of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : startSWRotation
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get startSWRotation() As Boolean
   On Error GoTo startSWRotationGet_Error

   startSWRotation = mstartSWRotation

   On Error GoTo 0
   Exit Property

startSWRotationGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property startSWRotation of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : startSWRotation
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let startSWRotation(ByVal newValue As Boolean)
    On Error GoTo startSWRotationLet_Error

    If mstartSWRotation <> newValue Then mstartSWRotation = newValue Else Exit Property
    
    If mstartSWRotation = 0 Then
        tmrStopWatch.Enabled = True
        tmrSWRotation.Enabled = False
    Else
        tmrStopWatch.Enabled = False
        tmrSWRotation.Interval = 50 ' mstartSWRotation
        tmrSWRotation.Enabled = True
    End If

   On Error GoTo 0
   Exit Property

startSWRotationLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property startSWRotation of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : startSWDegrees
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get startSWDegrees() As Long
   On Error GoTo startSWDegreesGet_Error

   startSWDegrees = mstartSWDegrees

   On Error GoTo 0
   Exit Property

startSWDegreesGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property startSWDegrees of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : startSWDegrees
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let startSWDegrees(ByVal newValue As Long)
    On Error GoTo startSWDegreesLet_Error
'
'    If mstartSWDegrees = 0 Then
'        mstartSWDegrees = SWSDeg
'        Exit Property
'    End If
    If mstartSWDegrees <> newValue Then mstartSWDegrees = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

startSWDegreesLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property startSWDegrees of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : rotationSWInterval
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get rotationSWInterval() As Long
   On Error GoTo rotationSWIntervalGet_Error

   rotationSWInterval = mrotationSWInterval

   On Error GoTo 0
   Exit Property

rotationSWIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property rotationSWInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : rotationSWInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let rotationSWInterval(ByVal newValue As Long)
    On Error GoTo rotationSWIntervalLet_Error

    If mrotationSWInterval <> newValue Then mrotationSWInterval = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

rotationSWIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property rotationSWInterval of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : SWRotationDegrees
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SWRotationDegrees() As Long
   On Error GoTo SWRotationDegreesGet_Error

   SWRotationDegrees = mSWRotationDegrees

   On Error GoTo 0
   Exit Property

SWRotationDegreesGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SWRotationDegrees of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : SWRotationDegrees
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SWRotationDegrees(ByVal newValue As Long)
    On Error GoTo SWRotationDegreesLet_Error

    If mSWRotationDegrees <> newValue Then mSWRotationDegrees = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

SWRotationDegreesLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SWRotationDegrees of Class Module cwOverlay"
End Property

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
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
End Property




'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property (by value) used to determine the opacity of elements in the overlay
'---------------------------------------------------------------------------------------
'
Public Property Let MyOpacity(ByVal newValue As Double)
   On Error GoTo MyOpacityLet_Error

   If mMyOpacity <> newValue Then mMyOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

MyOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
End Property

'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property (by value) used to determine the opacity of elements in the overlay
'---------------------------------------------------------------------------------------
'
Public Property Get MyOpacity() As Double
   On Error GoTo MyOpacityGet_Error

   MyOpacity = mMyOpacity

   On Error GoTo 0
   Exit Property

MyOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
End Property



'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
End Property

    


'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
End Property

'---------------------------------------------------------------------------------------
' Property  : FaceMode
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FaceMode() As String
   On Error GoTo FaceModeGet_Error

   FaceMode = mFaceMode

   On Error GoTo 0
   Exit Property

FaceModeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FaceMode of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : FaceMode
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FaceMode(ByVal newValue As String)
   On Error GoTo FaceModeLet_Error

    If mFaceMode <> newValue Then mFaceMode = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

FaceModeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FaceMode of Class Module cwOverlay"
End Property
 


'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtDayNumber As String: txtDayNumber = vbNullString
    Dim txtMonthNumber As String: txtMonthNumber = vbNullString
    Dim textDrawn As Boolean: textDrawn = False
    Dim thisOpacity As Double: thisOpacity = 0
    Dim cream As String: cream = vbNullString
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
   
    On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 18
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    thisOpacity = (Val(gblOpacity) / 100) * 0.75
    
    storeTxtDayNumber = txtDayNumber
    txtDayNumber = Day(DateAdd("m", tzDelta, Now))
    If storeTxtDayNumber <> txtDayNumber Then textDrawn = False ' trigger a redraw of the day and month within the boxes
    
    txtMonthNumber = Month(DateAdd("m", tzDelta, Now))

    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfGauge by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      ' derive a Pattern-Obj from each retrieved ImageList-Surface-reference
      If IsArray(Arr) Then
             CC.Save
             
             ' for specific keys perform rotation actions, clock hands &c
             Select Case Key
                 Case "hourshadow", "hourhand" 'hours-shadow- and hour-hand
                 
                    ' the DrawText code MUST be here, before the graphical elements, in order to appear underneath the hands
                    ' if you move this section out of the loop it causes graphical problems
                    If FaceMode = "0" Then  'stopwatch mode
                       If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
                           CC.DrawText 187, 212, 200, 200, Str$(SWSec), , , 2, True, , thisOpacity
                       Else
                           CC.DrawText 187, 212, 200, 200, Str$(Sec), , , 2, True, , thisOpacity
                       End If
                    Else ' clock mode
                        If textDrawn = False Then ' draw the date and day in the windows just once to retain opacity
                            CC.DrawText 260, 103, 200, 200, txtDayNumber, , , 2, True, , thisOpacity
                            CC.DrawText 458, 307, 200, 200, txtMonthNumber, , , 2, True, , thisOpacity
                            textDrawn = True
                        End If
                    End If
                    
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(19, 160) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
                        CC.RotateDrawingsDeg 250 'place the hour hand out of the way of the stopwatch mini-faces
                    Else
                        CC.RotateDrawingsDeg -7 + HDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                    End If
                 
                 Case "minuteshadow", "minutehand" 'minutes-shadow- and minutes-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(44, 105) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
                        CC.RotateDrawingsDeg 185        'place the minute hand out of the way of the stopwatch mini-faces
                    Else
                        CC.RotateDrawingsDeg -61 + MDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                    End If
                    
                 Case "secondshadow", "secondhand" 'seconds-shadow- and seconds-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(60, 160) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area

                    If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
                        CC.RotateDrawingsDeg 185 'place the minute hand out of the way of the stopwatch mini-faces
                    Else
                        CC.RotateDrawingsDeg -50 + SDeg + MsDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                    End If

                 Case "swhourhand"
                     If FaceMode = "0" Then  'stopwatch mode
                        Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(15, 61) 'shift to the "rotation-point" within the Pattern
                        CC.TranslateDrawings 352, 418  'shift to the rotation-point
                        CC.RotateDrawingsDeg SWHDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                     Else ' clock mode
                        Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(15, 61) 'shift to the "rotation-point" within the Pattern
                        CC.TranslateDrawings 432, 240  'shift to the rotation-point
                        CC.RotateDrawingsDeg HDeg1 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                     End If
                 Case "swminutehand"
                     If FaceMode = "0" Then  'stopwatch mode
                         Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         Set Pat.Matrix = Pat.Matrix.TranslateCoords(12, 57) 'shift to the "rotation-point" within the Pattern
                         CC.TranslateDrawings 353, 203  'shift to the rotation-point
                         CC.RotateDrawingsDeg SWMDeg
                     Else ' clock mode
                         Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         Set Pat.Matrix = Pat.Matrix.TranslateCoords(12, 57) 'shift to the "rotation-point" within the Pattern
                         CC.TranslateDrawings 286, 380  'shift to the rotation-point for the clock
                         CC.RotateDrawingsDeg MDeg1
                     End If
                 Case "swsecondhand"
                     If FaceMode = "0" Then  'stopwatch mode
                         ' just show the respective hand static and unmoving
                        If gblStopWatchState = 0 Then
                            Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                            Set Pat.Matrix = Pat.Matrix.TranslateCoords(12, 57) 'shift to the "rotation-point" within the Pattern
                            CC.TranslateDrawings 460, 312  'shift to the rotation-point
                            CC.RotateDrawingsDeg 0
                        End If
                         ' only rotate the stopwatch hands if the top button has been pressed at least once
                        If gblStopWatchState = 1 Or gblStopWatchState = 2 Then
                            Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                            Set Pat.Matrix = Pat.Matrix.TranslateCoords(12, 57) 'shift to the "rotation-point" within the Pattern
                            CC.TranslateDrawings 460, 312  'shift to the rotation-point
                            CC.RotateDrawingsDeg SWSDeg
                        End If
                     End If
                 Case "bigreflectioncopy", "bigreflection", "windowreflection", "helpbutton" 'plain output for everything else ... this includes the reflection-layers,
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
                 Case Else
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
' Author    : olaf schmidt
' Date      : 30/07/2023
' Purpose   : update the degree-Variables for the clock Hour, Minute and Second Hands
'---------------------------------------------------------------------------------------
'
Private Sub tmrClock_Timer()
 
    Dim currentLocaltime As SYSTEMTIME
    Static T As Date: T = Time
    
    On Error GoTo tmrClock_Timer_Error
    
    GetSystemTime currentLocaltime ' API
    
    ' every 15 mins approx. update DaylightSavings and determine the time bias
    If gblMainDaylightSaving <> "0" Then
        If (Minute(T) Mod 15 = 0) And (Second(T) = 0) And (currentLocaltime.wMilliseconds < 100) Then
            If gblMainDaylightSaving <> "0" Then
                 tzDelta = fObtainDaylightSavings("Main") ' determine the time bias
            End If
            If gblSecondaryDaylightSaving <> "0" Then
                 tzDelta1 = fObtainDaylightSavings("Secondary") ' determine the time bias
            End If
        End If
    End If
        
    ' Debug.Print ("%updateTime-I tzDelta " & tzDelta)

    ' time calc for main gauge
    HDeg = 30 * (Hour(T) Mod 12 + (Minute(T) + tzDelta) / 60)
    MDeg = 6 * ((Minute(T) + tzDelta) + (Second(T) / 60))
    
    ' time calc for secondary gauge
    HDeg1 = 30 * (Hour(T) Mod 12 + (Minute(T) + tzDelta1) / 60)
    MDeg1 = 6 * ((Minute(T) + tzDelta1) + (Second(T) / 60))
    
    Sec = Second(T)

    If mSmoothSecondHand = True Then

        SDeg = 6 * Sec

        With currentLocaltime
            MsDeg = 6 * (.wMilliseconds / 1000)
        End With
        
        ' always refresh
    Else
        If SDeg <> 6 * Sec Then
            SDeg = 6 * Sec
            
            MsDeg = 0
        Else
            Exit Sub  'only refresh when SDeg has changed to a new value
        End If
    End If
    
    W.Parent.Refresh ' causes a paint event to fire

   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : tmrStopWatch_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : update the degree-Variables for the stopwatch Hour, Minute and Second Hands
'---------------------------------------------------------------------------------------
'
Private Sub tmrStopWatch_Timer()
    Dim theTime As Date:  theTime = #1/1/2000 12:00:00 PM#
    Dim elapsedTimeSecs As Long ' long means a max of 2,147,483,647 seconds, 24,855 days or 68 years
    
    On Error GoTo tmrStopWatch_Timer_Error
        
    theTime = Now()
    elapsedTimeSecs = DateDiff("s", gblStopWatchStartTime, theTime)

    ' time degree calculation for main gauge
    If gblStopWatchState = 2 Then ' pause
        SWHDeg = SWHDeg
        SWMDeg = SWMDeg
        SWSDeg = SWSDeg
        SWSec = elapsedTimeSecs Mod 60 ' the digital display keeps running
    Else ' stopwatch running
        SWHDeg = (elapsedTimeSecs * 0.001667) Mod 360
        SWMDeg = (elapsedTimeSecs * 0.1) Mod 360
        SWSDeg = (elapsedTimeSecs * 6) Mod 360
        SWSec = elapsedTimeSecs Mod 60
    End If

    W.Parent.Refresh ' no need to refresh as this is done by the main timer

   On Error GoTo 0
   Exit Sub

tmrStopWatch_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrStopWatch_Timer of Class Module cwOverlay"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrSWRotation_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : update the degree-Variables for the stopwatch Second Hand only
'             allowing the gauge Second hand to rotate to new position
'---------------------------------------------------------------------------------------
'
Private Sub tmrSWRotation_Timer()
    Dim degrees As Integer: degrees = 0
    
    On Error GoTo tmrSWRotation_Timer_Error
           
    If rotationTimerCount < mSWRotationDegrees Then
        rotationTimerCount = rotationTimerCount + mrotationSWInterval
        If rotationTimerCount > mSWRotationDegrees Then rotationTimerCount = mSWRotationDegrees
        degrees = mstartSWDegrees + rotationTimerCount
       
        'Case "swsecondhand"
        SWSDeg = degrees
    Else
        tmrStopWatch.Enabled = True
        
        rotationTimerCount = 0
        If gblStopWatchState = 0 Or gblStopWatchState = 2 Then
            gblStopWatchState = 1
        Else
            gblStopWatchState = 0
        End If
        
        startSWRotation = False
        tmrSWRotation.Enabled = False
    End If
                
    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrSWRotation_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSWRotation_Timer of Class Module cwOverlay"
End Sub

'from set location to current time
    
'---------------------------------------------------------------------------------------
' Procedure : tmrHandsRotation_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : update the degree-Variables for the stopwatch Second Hand only
'             allowing the gauge Second hand to rotate to new position
'---------------------------------------------------------------------------------------
'
Private Sub tmrHandsRotation_Timer()
    Dim degrees As Integer: degrees = 0
    
    On Error GoTo tmrHandsRotation_Timer_Error
           
    If rotationTimerCount < mSWRotationDegrees Then
        rotationTimerCount = rotationTimerCount + mrotationSWInterval
        If rotationTimerCount > mSWRotationDegrees Then rotationTimerCount = mSWRotationDegrees
        degrees = mstartSWDegrees + rotationTimerCount
       
        'Case "swsecondhand"
        SWSDeg = degrees
        HDeg = degrees
        MDeg = degrees
        SDeg = degrees
    Else
        tmrStopWatch.Enabled = True
        
        rotationTimerCount = 0
        If gblStopWatchState = 0 Or gblStopWatchState = 2 Then
            gblStopWatchState = 1
        Else
            gblStopWatchState = 0
        End If
        
        startSWRotation = False
        tmrHandsRotation.Enabled = False
    End If
                
    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrHandsRotation_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrHandsRotation_Timer of Class Module cwOverlay"
End Sub



''---------------------------------------------------------------------------------------
'' Property  : TmrGaugeTicking
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property to determine (by value) externally the state of a private timer
''---------------------------------------------------------------------------------------
''
'Public Property Get TmrGaugeTicking() As Boolean
'   On Error GoTo TmrGaugeTickingGet_Error
'
'   TmrGaugeTicking = mTmrGaugeTicking
'
'   On Error GoTo 0
'   Exit Property
'
'TmrGaugeTickingGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrGaugeTicking of Class Module cfGauge"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : TmrGaugeTicking
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property to determine (by value) externally the state of a private timer
''---------------------------------------------------------------------------------------
''
'Public Property Let TmrGaugeTicking(ByVal newValue As Boolean)
'   On Error GoTo TmrGaugeTickingLet_Error
'
'   tmrGauge.Enabled = newValue
'
'   If mTmrGaugeTicking <> newValue Then mTmrGaugeTicking = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'TmrGaugeTickingLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrGaugeTicking of Class Module cfGauge"
'End Property
'
''---------------------------------------------------------------------------------------
'' Property  : TmrGaugeInterval
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property to determine (by value) externally the interval of a private timer
''---------------------------------------------------------------------------------------
''
'Public Property Get TmrGaugeInterval() As Long
'   On Error GoTo TmrGaugeIntervalGet_Error
'
'   TmrGaugeInterval = mTmrGaugeInterval
'
'   On Error GoTo 0
'   Exit Property
'
'TmrGaugeIntervalGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrGaugeInterval of Class Module cfGauge"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : TmrGaugeInterval
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   : property to determine (by value) externally the interval of a private timer
''---------------------------------------------------------------------------------------
''
'Public Property Let TmrGaugeInterval(ByVal newValue As Long)
'   On Error GoTo TmrGaugeIntervalLet_Error
'
'   If mTmrGaugeInterval <> newValue Then mTmrGaugeInterval = newValue Else Exit Property
'
'   tmrGauge.Interval = newValue
'
'   On Error GoTo 0
'   Exit Property
'
'TmrGaugeIntervalLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrGaugeInterval of Class Module cfGauge"
'End Property
'


