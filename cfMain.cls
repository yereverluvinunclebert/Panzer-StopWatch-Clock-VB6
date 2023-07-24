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



'---------------------------------------------------------------------------------------
' Procedure : InitAndShowAsFreeForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub InitAndShowAsFreeForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo InitAndShowAsFreeForm_Error

    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 750, 675)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
         
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 0, 0, 870, 870)

   On Error GoTo 0
   Exit Sub

InitAndShowAsFreeForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitAndShowAsFreeForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
'' Procedure : GaugeForm_DblClick
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub GaugeForm_DblClick()
'
'    On Error GoTo GaugeForm_DblClick_Error
'
'    If LTrim$(PzGDblClickCommand) = vbNullString Then Exit Sub
'    If SHIFT_1 = True Then
'        SHIFT_1 = False
'        Call ShellExecute(fAlpha.GaugeForm.hwnd, "Open", PzGOpenFile, vbNullString, App.Path, 1)
'    Else
'        Call ShellExecute(fAlpha.GaugeForm.hwnd, "Open", PzGDblClickCommand, vbNullString, App.Path, 1)
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'GaugeForm_DblClick_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GaugeForm_DblClick of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub
'
'
''Private Sub GaugeForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''        'If pinClicked = True Then MsgBox "X = " & x & " Y = " & y & " pinClicked = " & pinClicked
''
''End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : GaugeForm_MouseMove
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub GaugeForm_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'    On Error GoTo GaugeForm_MouseMove_Error
'
'    'If globeWidget.Locked = True Then Exit Sub
'    If PzGIgnoreMouse = "1" Then Exit Sub
'
'    Static x0 As Single
'    Static y0 As Single
'    If Button Then GaugeForm.Move GaugeForm.Left + X - x0, GaugeForm.Top + Y - y0 Else x0 = X: y0 = Y
'
'    On Error GoTo 0
'    Exit Sub
'
'GaugeForm_MouseMove_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GaugeForm_MouseMove of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : GaugeForm_KeyDown
'' Author    : beededea
'' Date      : 01/06/2019
'' Purpose   : get F5 and SHIFT keypresses
''---------------------------------------------------------------------------------------
''
'Private Sub GaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef shift As Integer)
'    On Error GoTo GaugeForm_KeyDown_Error
'
'    Call getKeyPress(KeyCode, shift)
'
'    On Error GoTo 0
'    Exit Sub
'
'GaugeForm_KeyDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GaugeForm_KeyDown of Class Module cfMain"
'End Sub
''
''---------------------------------------------------------------------------------------
'' Procedure : GaugeForm_MouseUp
'' Author    : beededea
'' Date      : 05/05/2023
'' Purpose   : save the form x,y position when ever the globe/form is dragged
''---------------------------------------------------------------------------------------
''
'Private Sub GaugeForm_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'    On Error GoTo GaugeForm_MouseUp_Error
'
'    If globeWidget.Locked = True Then Exit Sub
'
'    ' no point in saving when prefs are visible.
'    If panzerPrefs.IsVisible = False Then ' checks without activating the form - important!.
'       PzGMaximiseFormX = Str$(GaugeForm.Left) ' saving in pixels
'       PzGMaximiseFormY = Str$(GaugeForm.Top)
'
'       sPutINISetting "Software\PzStopwatch", "maximiseFormX", PzGMaximiseFormX, PzGSettingsFile
'       sPutINISetting "Software\PzStopwatch", "maximiseFormY", PzGMaximiseFormY, PzGSettingsFile
'    End If
'
'    'Static x0, y0: If Button Then GaugeForm.Move GaugeForm.Left + X - x0, GaugeForm.Top + Y - y0 Else x0 = X: y0 = Y
'
'    On Error GoTo 0
'    Exit Sub
'
'GaugeForm_MouseUp_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GaugeForm_MouseUp of Class Module cfMain"
'            Resume Next
'          End If
'    End With
'End Sub