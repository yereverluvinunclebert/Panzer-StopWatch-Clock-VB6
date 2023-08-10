VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module      : cwhelp
' Author            : beededea
' Date              : 17/05/2023
' Purpose           :
'---------------------------------------------------------------------------------------

Option Explicit
 
Private WithEvents W As cWidgetBase '<- this is required in each cwImplementation...
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrFadeIn As cTimer
Attribute tmrFadeIn.VB_VarHelpID = -1
Private WithEvents tmrFadeOut As cTimer
Attribute tmrFadeOut.VB_VarHelpID = -1

Private helpxo As Integer
Private helpyo As Integer

Private mOpacity As Double
Private mShow As Boolean

'property opacity is a property that populates mOpacity used by RenderSurfaceContent
'property show is a property that populates mShow that triggers a fade in timer

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

   On Error GoTo Class_Initialize_Error

    helpxo = 0
    helpyo = 0
    opacity = 0
    show = False

    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
   
    Set tmrFadeIn = New_c.Timer(20, False)
    Set tmrFadeOut = New_c.Timer(20, False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwhelp"
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets:  Set Widgets = W.Widgets: End Property ' does this really need to be exposed?
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, shift As Integer, ByVal X As Single, ByVal Y As Single)
   On Error GoTo W_MouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    fileToPlay = "ting.wav"
    If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    tmrFadeOut.Enabled = True ' this is not currently making the opacity reduce

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwhelp"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let opacity(ByVal newValue As Double)
   On Error GoTo opacityLet_Error

   If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

opacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get opacity() As Double
   On Error GoTo opacityGet_Error

   opacity = mOpacity

   On Error GoTo 0
   Exit Property

opacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwhelp"
End Property

'---------------------------------------------------------------------------------------
' Procedure : show
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let show(ByVal newValue As Boolean)
    On Error GoTo showLet_Error

    mShow = newValue
    If mShow = True Then tmrFadeIn.Enabled = True
    
    On Error GoTo 0
    Exit Property

showLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure show of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : show
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get show() As Boolean
   On Error GoTo showGet_Error

   show = mShow

   On Error GoTo 0
   Exit Property

showGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure show of Class Module cwhelp"
End Property

'
'---------------------------------------------------------------------------------------
' Procedure : tmrFadeIn_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the help form or help image fade in opacity
'---------------------------------------------------------------------------------------
'
Private Sub tmrFadeIn_Timer()
      
   On Error GoTo tmrFadeIn_Timer_Error
    fMain.helpForm.show
    
    opacity = opacity + 0.05
    If opacity >= 1 Then
        opacity = 1

        tmrFadeIn.Enabled = False

        Exit Sub
    End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrFadeIn_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrFadeIn_Timer of Class Module cwhelp"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : tmrFadeOut_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the help form or help image fade out in opacity
'---------------------------------------------------------------------------------------
' < - this timer is not causing the opacity to fade out
Private Sub tmrFadeOut_Timer()
      
   On Error GoTo tmrFadeOut_Timer_Error

    opacity = opacity - 0.05
    
    If opacity < 0 Then
        opacity = 0
        
        tmrFadeOut.Enabled = False
        fMain.helpForm.Hide
        
        Exit Sub
    End If
    
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrFadeOut_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrFadeOut_Timer of Class Module cwhelp"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   On Error GoTo W_Paint_Error

    Draw CC, dx_Aligned, dy_Aligned, mOpacity

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwhelp"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Draw
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Draw(ByRef CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal mOpacity As Double)
   
    Dim scaleValue As Double: scaleValue = 0.8
    
    On Error GoTo Draw_Error

    If PzGDpiAwareness = "1" Then scaleValue = 1.4

    CC.ScaleDrawings scaleValue, scaleValue

    CC.RenderSurfaceContent "help", helpxo, helpyo, , , , mOpacity
        
    W.Refresh
   

   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwhelp"

End Sub


