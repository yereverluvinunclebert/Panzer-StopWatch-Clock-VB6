VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Class Module      : cwAbout
' Author            : beededea
' Date              : 17/05/2023
' Purpose           : puts an about widget on the form
'---------------------------------------------------------------------------------------

Option Explicit
 
Private WithEvents W As cWidgetBase '<- this is required in each cwImplementation...
Attribute W.VB_VarHelpID = -1

' RC timers

Private WithEvents tmrAboutFadeIn As cTimer
Attribute tmrAboutFadeIn.VB_VarHelpID = -1
Private WithEvents tmrAboutFadeOut As cTimer
Attribute tmrAboutFadeOut.VB_VarHelpID = -1

' private members for property assignment

Private mMyOpacity As Double
Private mShowMe As Boolean
Private mZoom As Double

'local general vars used only within this class

Private aboutxo As Integer
Private aboutyo As Integer

'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   : initialise vars
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()

   On Error GoTo initialisePrivateVars_Error

    aboutxo = 0
    aboutyo = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cwAbout"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : Constructor (Initialise)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

   On Error GoTo Class_Initialize_Error

    Call initialisePrivateVars
    
    MyOpacity = 0
    ShowMe = False
    
    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
       
    Set tmrAboutFadeIn = New_c.Timer(20, False)
    Set tmrAboutFadeOut = New_c.Timer(20, False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwAbout"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   : An RC default property
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   : An RC default property
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the zoom value, it being selected by a mousewheel scroll or the slider in the prefs
'---------------------------------------------------------------------------------------
'
Public Property Get Zoom() As Double
   On Error GoTo ZoomGet_Error

    Zoom = mZoom

   On Error GoTo 0
   Exit Property

ZoomGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the zoom value, it being selected by a mousewheel scroll or the slider in the prefs
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwAbout"
End Property

'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : Handle the click event for a click upon the widget
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
   Dim fileToPlay As String: fileToPlay = vbNullString
   On Error GoTo W_MouseDown_Error
    
    ' works better here
'    If gblVolumeBoost = "1" Then
'        fileToPlay = "till.wav"
'    Else
'        fileToPlay = "till-quiet.wav"
'    End If
'
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        playSound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    tmrAboutFadeOut.Enabled = True

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwAbout"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity value used by RenderSurfaceContent
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
' Purpose   : property to determine (by value) the opacity value used by RenderSurfaceContent
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
' Procedure : ShowMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) that triggers a fade in timer
'---------------------------------------------------------------------------------------
'
Public Property Let ShowMe(ByVal newValue As Boolean)
    On Error GoTo ShowMeLet_Error

    mShowMe = newValue
    
    If mShowMe = True Then tmrAboutFadeIn.Enabled = True
    
    On Error GoTo 0
    Exit Property

ShowMeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMe of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ShowMe
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) that triggers a fade in timer
'---------------------------------------------------------------------------------------
'
Public Property Get ShowMe() As Boolean
   On Error GoTo ShowMeGet_Error

   ShowMe = mShowMe

   On Error GoTo 0
   Exit Property

ShowMeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowMe of Class Module cwAbout"
End Property

'
'---------------------------------------------------------------------------------------
' Procedure : tmrAboutFadeIn_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the about form or about image fade in opacity
'---------------------------------------------------------------------------------------
'
Private Sub tmrAboutFadeIn_Timer()
      
   On Error GoTo tmrAboutFadeIn_Timer_Error
    fMain.aboutForm.Show
    MyOpacity = MyOpacity + 0.05
    If MyOpacity >= 1 Then
        MyOpacity = 1

        tmrAboutFadeIn.Enabled = False

        Exit Sub
    End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrAboutFadeIn_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAboutFadeIn_Timer of Class Module cwAbout"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : tmrAboutFadeOut_Timer
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : timer to make the about form or about image fade out in opacity
'---------------------------------------------------------------------------------------
Private Sub tmrAboutFadeOut_Timer()
    On Error GoTo tmrAboutFadeOut_Timer_Error
    
    MyOpacity = MyOpacity - 0.05
    
    If MyOpacity < 0 Then
        MyOpacity = 0
        
        tmrAboutFadeOut.Enabled = False
        fMain.aboutForm.Hide
        Exit Sub
    End If
    
    W.Refresh

   On Error GoTo 0
   Exit Sub

tmrAboutFadeOut_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAboutFadeOut_Timer of Class Module cwAbout"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : Default RC paint routine
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
   On Error GoTo W_Paint_Error

    Draw CC, dx_Aligned, dy_Aligned, mMyOpacity

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwAbout"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Draw
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : Draw routine called by W_PAINT
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
    
    CC.RenderSurfaceContent "about", aboutxo, aboutyo, , , , mMyOpacity
        
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontItalic = False
    W.FontName = gblGaugeFont
    W.ForeColor = vbBlack
    W.FontBold = True
    W.FontSize = 11

    W.SelectFontSettingsInto CC
    
    If gblDpiAwareness = "0" Then
        CC.DrawText 350, 0, dx, dy - 275, "version " & App.Major & "." & App.Minor & "." & App.Revision & " Alpha", , , 2, True, , mMyOpacity
    Else
        
        #If TWINBASIC Then
            CC.DrawText 125, -202, dx, dy, "TwinBasic", , , 2, True, , mMyOpacity
        #Else
            CC.DrawText 168, -202, dx, dy, "VB6", , , 2, True, , mMyOpacity
        #End If
        CC.DrawText 350, 0, dx, dy - 1020, "version " & App.Major & "." & App.Minor & "." & App.Revision & " Alpha", , , 2, True, , mMyOpacity
        
    End If
    W.Refresh

   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwAbout"

End Sub




' the following global variables are used in this class, I should replace each with a property and use that locally
' but I don't think I can be bothered, there are too many of them. This is my code and the gauge class is not going
' to have multiple instances running so it does not really matter. If you want to convert each of these to property GET/LETs
' then please feel free to work your way through this list.

' gblVolumeBoost
' gblEnableSounds





'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

