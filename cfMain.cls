VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfMain
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents aboutForm As cWidgetForm
Attribute aboutForm.VB_VarHelpID = -1
Public WithEvents helpForm As cWidgetForm
Attribute helpForm.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : initAndShowStandardForms
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initAndShowStandardForms(ByVal X As Long, ByVal Y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowStandardForms_Error
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", , 1400, 1200)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
         
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 0, 0, 1100, 970)

    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1400, 1200)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
         
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 0, 0, 1100, 970)


   On Error GoTo 0
   Exit Sub

initAndShowStandardForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowStandardForms of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
'' Procedure : gaugeForm_DblClick
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub gaugeForm_DblClick()
'
'    On Error GoTo gaugeForm_DblClick_Error
'
'    If LTrim$(PzGDblClickCommand) = vbNullString Then Exit Sub
'    If SHIFT_1 = True Then
'        SHIFT_1 = False
'        Call ShellExecute(fAlpha.gaugeForm.hwnd, "Open", PzGOpenFile, vbNullString, App.Path, 1)
'    Else
'        Call ShellExecute(fAlpha.gaugeForm.hwnd, "Open", PzGDblClickCommand, vbNullString, App.Path, 1)
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'gaugeForm_DblClick_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_DblClick of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub
'
'
''Private Sub gaugeForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''        'If pinClicked = True Then MsgBox "X = " & x & " Y = " & y & " pinClicked = " & pinClicked
''
''End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : gaugeForm_MouseMove
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub gaugeForm_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'    On Error GoTo gaugeForm_MouseMove_Error
'
'    'If overlayWidget.Locked = True Then Exit Sub
'    If PzGIgnoreMouse = "1" Then Exit Sub
'
'    Static x0 As Single
'    Static y0 As Single
'    If Button Then gaugeForm.Move gaugeForm.Left + X - x0, gaugeForm.Top + Y - y0 Else x0 = X: y0 = Y
'
'    On Error GoTo 0
'    Exit Sub
'
'gaugeForm_MouseMove_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseMove of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : gaugeForm_KeyDown
'' Author    : beededea
'' Date      : 01/06/2019
'' Purpose   : get F5 and SHIFT keypresses
''---------------------------------------------------------------------------------------
''
'Private Sub gaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef shift As Integer)
'    On Error GoTo gaugeForm_KeyDown_Error
'
'    Call getKeyPress(KeyCode, shift)
'
'    On Error GoTo 0
'    Exit Sub
'
'gaugeForm_KeyDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_KeyDown of Class Module cfMain"
'End Sub
''
''---------------------------------------------------------------------------------------
'' Procedure : gaugeForm_MouseUp
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   : save the form x,y position when ever the globe/form is dragged
''---------------------------------------------------------------------------------------
''
'Private Sub gaugeForm_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'    On Error GoTo gaugeForm_MouseUp_Error
'
'    If overlayWidget.Locked = True Then Exit Sub
'
'    ' no point in saving when prefs are visible.
'    If panzerPrefs.IsVisible = False Then ' checks without activating the form - important!.
'       PzGMaximiseFormX = Str$(gaugeForm.Left) ' saving in pixels
'       PzGMaximiseFormY = Str$(gaugeForm.Top)
'
'       sPutINISetting "Software\PzStopwatch", "maximiseFormX", PzGMaximiseFormX, PzGSettingsFile
'       sPutINISetting "Software\PzStopwatch", "maximiseFormY", PzGMaximiseFormY, PzGSettingsFile
'    End If
'
'    'Static x0, y0: If Button Then gaugeForm.Move gaugeForm.Left + X - x0, gaugeForm.Top + Y - y0 Else x0 = X: y0 = Y
'
'    On Error GoTo 0
'    Exit Sub
'
'gaugeForm_MouseUp_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gaugeForm_MouseUp of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub
