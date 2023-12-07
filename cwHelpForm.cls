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
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Class Module      : cwhelp
' Author            : beededea
' Date              : 17/05/2023
' Purpose           :
'---------------------------------------------------------------------------------------

Option Explicit
 
Private WithEvents W As cWidgetBase '<- this is required in each cwImplementation...
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrHelpFadeIn As cTimer
Attribute tmrHelpFadeIn.VB_VarHelpID = -1
Private WithEvents tmrHelpFadeOut As cTimer
Attribute tmrHelpFadeOut.VB_VarHelpID = -1

Private helpxo As Integer
Private helpyo As Integer

Private mMyOpacity As Double
Private mShowMe As Boolean
Private mZoom As Double


'property MyOpacity is a property that populates mMyOpacity used by RenderSurfaceContent
'property ShowMe is a property that populates mShowMe that triggers a fade in timer
'property mZoom

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
    MyOpacity = 0
    ShowMe = False

    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
   
    Set tmrHelpFadeIn = New_c.Timer(20, False)
    Set tmrHelpFadeOut = New_c.Timer(20, False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwhelp"
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets:  Set Widgets = W.Widgets: End Property ' does this really need to be exposed?
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Zoom() As Double
   On Error GoTo ZoomGet_Error

    Zoom = mZoom

   On Error GoTo 0
   Exit Property

ZoomGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Zoom(ByVal newValue As Double)
   On Error GoTo ZoomLet_Error

  If mZoom <> newValue Then mZoom = newValue Else Exit Property
  If mZoom > 3 Then mZoom = 3 Else If mZoom < 0.05 Then mZoom = 0.05
  W.Refresh

   On Error GoTo 0
   Exit Property

ZoomLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal Y As Single)
   On Error GoTo W_MouseDown_Error

    tmrHelpFadeOut.Enabled = True ' this is not currently making the opacity reduce

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwhelp"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let MyOpacity(ByVal newValue As Double)
   On Error GoTo MyOpacityLet_Error

   If mMyOpacity <> newValue Then mMyOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

MyOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get MyOpacity() As Double
   On Error GoTo MyOpacityGet_Error

   MyOpacity = mMyOpacity

   On Error GoTo 0
   Exit Property

MyOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwhelp"
End Property

'---------------------------------------------------------------------------------------
' Procedure : ShowMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ShowMe(ByVal newValue As Boolean)
    On Error GoTo ShowMeLet_Error

    mShowMe = newValue
    If mShowMe = True Then tmrHelpFadeIn.Enabled = True
    
    On Error GoTo 0
    Exit Property

ShowMeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMe of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ShowMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ShowMe() As Boolean
   On Error GoTo ShowMeGet_Error

   ShowMe = mShowMe

   On Error GoTo 0
   Exit Property

ShowMeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMe of Class Module cwhelp"
End Property

'
'---------------------------------------------------------------------------------------
' Procedure : tmrHelpFadeIn_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the help form or help image fade in opacity
'---------------------------------------------------------------------------------------
'
Private Sub tmrHelpFadeIn_Timer()
      
   On Error GoTo tmrHelpFadeIn_Timer_Error
    fMain.helpForm.show
    MyOpacity = MyOpacity + 0.05
    If MyOpacity >= 1 Then
        MyOpacity = 1

        tmrHelpFadeIn.Enabled = False

        Exit Sub
    End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrHelpFadeIn_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrHelpFadeIn_Timer of Class Module cwhelp"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : tmrHelpFadeOut_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the help form or help image fade out in opacity
'---------------------------------------------------------------------------------------
Private Sub tmrHelpFadeOut_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    On Error GoTo tmrHelpFadeOut_Timer_Error

    MyOpacity = MyOpacity - 0.05
    
    If MyOpacity < 0 Then
        MyOpacity = 0
        
        tmrHelpFadeOut.Enabled = False
        fMain.helpForm.Hide
    
        fileToPlay = "till.wav"
        If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        
        Exit Sub
    End If
    
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrHelpFadeOut_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrHelpFadeOut_Timer of Class Module cwhelp"
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

    Draw CC, dx_Aligned, dy_Aligned, mMyOpacity

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
Private Sub Draw(ByRef CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal mMyOpacity As Double)
    On Error GoTo Draw_Error

    CC.ScaleDrawings mZoom, mZoom

    ' Widgets are refreshed "from the ground up" in RC5 and not so in RC6
    ' extra code NOT needed when using RC5 but now required in RC6 - starts
    CC.Operator = CAIRO_OPERATOR_CLEAR
     CC.Paint 'clear the whole background of this Widgets underlying Surface via Clear-Operator
    CC.Operator = CAIRO_OPERATOR_OVER 'switch back to the default-Operator
    ' extra code not needed when using RC5 but now required in RC6 - ends
    
    CC.RenderSurfaceContent "help", helpxo, helpyo, , , , mMyOpacity
        
   ' W.Refresh

   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwhelp"

End Sub

