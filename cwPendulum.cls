VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwPendulum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwPendulum
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a single fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfGauge
'---------------------------------------------------------------------------------------

Option Explicit

' public vars

Public cCollPendulumImages As cCollection 'collection made available publically, can't easily pass collections via property get/set

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

' public vars external usage

Public sldCurWidth As Double ' need to add set/get for these public vars
Public cxPerc As Single      ' need to add set/get for these public vars
Public cyPerc As Single       ' available externally, to allow the overlay to be centered on the form, need to add set/let for these public vars

' RC timers

Private WithEvents tmrPendulum As cTimer
Attribute tmrPendulum.VB_VarHelpID = -1

' class private members for property assignment

Private mSwingPendulum As Boolean
Private mTmrPendulumTicking As Boolean

' private vars general usage

Private pvtPendulumAcceleration As Double
Private pvtGravity As Double
Private pvtPendulumAngle As Double
Private pvtPendulumVelocity As Double


'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   : initialise vars
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()

   On Error GoTo initialisePrivateVars_Error

    pvtPendulumAcceleration = 0
    pvtGravity = 0
    pvtPendulumAngle = 0
    pvtPendulumVelocity = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cwPendulum"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize for cwPendulum
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
        W.BackColor = -1            ' we don't use any BackColor here, transparent
        W.ImplementsHitTest = True  ' essential

    ' variables set for the pendulum
    ' g = -0.1, a = 0.6, v = 2.5 uses 3% of cpu with a 40ms timer

    pvtGravity = -0.2
    pvtPendulumAcceleration = 0.6
    pvtPendulumVelocity = 3.5
        
    Set tmrPendulum = Nothing
    Set tmrPendulum = New_c.Timer(70, False)
    
    TmrPendulumTicking = False
    
    On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwPendulum"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwPendulum"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwPendulum"
End Property



'---------------------------------------------------------------------------------------
' Property : SwingPendulum
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property (by value) used to toggle the swinging pendulum animation
'---------------------------------------------------------------------------------------
'
Public Property Get SwingPendulum() As Boolean
   On Error GoTo SwingPendulumGet_Error

   SwingPendulum = mSwingPendulum

   On Error GoTo 0
   Exit Property

SwingPendulumGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SwingPendulum of Class Module cfGauge"
End Property
'---------------------------------------------------------------------------------------
' Property  : SwingPendulum
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property (by value) used to toggle the swinging pendulum animation
'---------------------------------------------------------------------------------------
'
Public Property Let SwingPendulum(ByVal newValue As Boolean)

    On Error GoTo SwingPendulumLet_Error
        
    If mSwingPendulum = True Then
        'tmrPendulum.Enabled = True
        TmrPendulumTicking = True
    Else
        TmrPendulumTicking = False
        'tmrPendulum.Enabled = False
    End If
    
    If mSwingPendulum <> newValue Then mSwingPendulum = newValue Else Exit Property
    SwingPendulum = mSwingPendulum
    
    W.Refresh

   On Error GoTo 0
   Exit Property

SwingPendulumLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SwingPendulum of Class Module cfGauge"
End Property



' ----------------------------------------------------------------
' Procedure Name: tmrPendulum_Timer
' Purpose:  timer controlling the pendulum only
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrPendulum_Timer()
    
    On Error GoTo tmrPendulum_Timer_Error

    pvtPendulumAcceleration = pvtGravity * pvtPendulumAngle
    pvtPendulumVelocity = pvtPendulumVelocity + pvtPendulumAcceleration
    pvtPendulumAngle = pvtPendulumAngle + pvtPendulumVelocity

    W.Parent.Refresh ' causes a paint event to fire
    
   On Error GoTo 0
   Exit Sub

tmrPendulum_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPendulum_Timer of Class Module cwPendulum"
End Sub



'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Pendulum-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwPendulum"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : this is run when the graphics are initially placed on the screen, refreshing and when a mouse cursor is interacting
'---------------------------------------------------------------------------------------
'

Private Sub W_Paint(ByRef CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern

    W.Refresh
    
    On Error GoTo W_Paint_Error
  
    For I = 0 To cCollPendulumImages.Count - 1 ' this is the ExcludedPaths collection populated in cfGauge by duplicating collPendulumElements.
      Arr = cCollPendulumImages.ItemByIndex(I)
      Key = cCollPendulumImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
            
            ' for specific keys, perform actions &c
             Select Case Key
 
                
                Case "pendulum"
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(21, 0) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 624, 325

                    CC.RotateDrawingsDeg pvtPendulumAngle 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'gauge)
                
             End Select
    
            CC.Paint W.AlphaInherited * Arr(2), Pat  'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
            CC.Restore
      End If
    Next
       
   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwPendulum"
End Sub


'---------------------------------------------------------------------------------------
' Property  : TmrPendulumTicking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) externally the interval of a private timer
'---------------------------------------------------------------------------------------
'
Public Property Get TmrPendulumTicking() As Boolean
   On Error GoTo TmrPendulumTickingGet_Error

   TmrPendulumTicking = mTmrPendulumTicking

   On Error GoTo 0
   Exit Property

TmrPendulumTickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrPendulumTicking of Class Module cfGauge"
End Property
'---------------------------------------------------------------------------------------
' Property  : TmrPendulumTicking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) externally the interval of a private timer
'---------------------------------------------------------------------------------------
'
Public Property Let TmrPendulumTicking(ByVal newValue As Boolean)
   On Error GoTo TmrPendulumTickingLet_Error

   tmrPendulum.Enabled = newValue
   
   If mTmrPendulumTicking <> newValue Then mTmrPendulumTicking = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

TmrPendulumTickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TmrPendulumTicking of Class Module cfGauge"
End Property

