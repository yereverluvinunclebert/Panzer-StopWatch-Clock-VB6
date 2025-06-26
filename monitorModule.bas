Attribute VB_Name = "monitorModule"
'---------------------------------------------------------------------------------------
' Module    : monitorModule
' Author    : beededea
' Date      : 13/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context

Option Explicit

'Constants for the return value when finding a monitor
Private Enum dwFlags
    MONITOR_DEFAULTTONULL = &H0       'If the monitor is not found, return 0
    MONITOR_DEFAULTTOPRIMARY& = &H1   'If the monitor is not found, return the primary monitor
    MONITOR_DEFAULTTONEAREST = &H2    'If the monitor is not found, return the nearest monitor
End Enum

Public Const MONITORINFOF_PRIMARY As Integer = 1

Public prefsMonitorStruct As UDTMonitor
Public gaugeMonitorStruct As UDTMonitor


Public Type UDTMonitor
    handle As Long
    Left As Long
    Right As Long
    Top As Long
    Bottom As Long
    
    WorkLeft As Long
    WorkRight As Long
    WorkTop As Long
    Workbottom As Long
    
    Height As Long
    Width As Long
    
    WorkHeight As Long
    WorkWidth As Long
    
    IsPrimary As Boolean
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long ' This is +1 (right - left = width)
  Bottom As Long ' This is +1 (bottom - top = height)
End Type

'Structure for the position of a monitor
Private Type tagMONITORINFO
    cbSize      As Long 'Size of structure
    rcMonitor   As RECT 'Monitor rect
    rcWork      As RECT 'Working area rect
    dwFlags     As Long 'Flags
End Type

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MonitorFromRect Lib "user32" (rc As RECT, ByVal dwFlags As dwFlags) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, MonInfo As tagMONITORINFO) As Long

Private rcMonitors() As RECT 'coordinate array for all monitors
Private rcVS         As RECT 'coordinates for Virtual Screen

' vars to obtain correct screen width (to correct VB6 bug) STARTS
Public Const HORZRES As Integer = 8
Public Const VERTRES As Integer = 10
Public Const DESKTOPHORZRES As Integer = &H76

Public gblScreenTwipsPerPixelX As Long ' .07 DAEB 26/04/2021 common.bas changed to use pixels alone, removed all unnecessary twip conversion
Public gblScreenTwipsPerPixelY As Long ' .07 DAEB 26/04/2021 common.bas changed to use pixels alone, removed all unnecessary twip conversion
'Public physicalScreenWidthTwips As Long
'Public physicalScreenHeightTwips As Long



''---------------------------------------------------------------------------------------
'' Procedure : fPixelsPerInchX
'' Author    : Elroy from Vbforums
'' Date      : 23/01/2021
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function fPixelsPerInchX() As Long
'    Dim hDC As Long: hDC = 0
'    Dim virtualWidth As Long: virtualWidth = 0
'    Dim physicalWidth As Long: physicalWidth = 0
'
'    Const ninetysix As Double = 96
'    'Const LOGPIXELSX As Integer = 88       '  Logical pixels/inch in X
'
'    On Error GoTo fPixelsPerInchX_Error
'
'    hDC = GetDC(0)
'    If hDC <> 0 Then
'        'fPixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX) ' always returns 96DPI
'
'        virtualWidth = GetDeviceCaps(hDC, HORZRES)
'        physicalWidth = GetDeviceCaps(hDC, DESKTOPHORZRES)
'
'        fPixelsPerInchX = (96 * physicalWidth / virtualWidth)
'        ReleaseDC 0, hDC
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'fPixelsPerInchX_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fPixelsPerInchX of Module Module1"
'End Function


'---------------------------------------------------------------------------------------
' Procedure : fTwipsPerPixelX
' Author    : Elroy from Vbforums
' Date      : 23/01/2021
' Purpose   : Calculate the twips per pixel in the X axis, by default does not use Screen.TwipsPerPixelX
'             as when a tablet screen is rotated, the "Screen" object of VB doesn't respond to the change
'             so it has to be done by hand using GetDeviceCaps API.
'---------------------------------------------------------------------------------------
'
Public Function fTwipsPerPixelX() As Single
    Dim hDC As Long: hDC = 0
    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
    
    Const LOGPIXELSX As Integer = 88       '  Logical pixels/inch in X
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    '
    On Error GoTo fTwipsPerPixelX_Error
    
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hDC = GetDC(0)
    If hDC <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
        ReleaseDC 0, hDC
        fTwipsPerPixelX = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        fTwipsPerPixelX = Screen.TwipsPerPixelX
    End If

   On Error GoTo 0
   Exit Function

fTwipsPerPixelX_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelX of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fTwipsPerPixelY
' Author    : Elroy from Vbforums
' Date      : 23/01/2021
' Purpose   : Calculate the twips per pixel in the Y axis, by default does not use Screen.TwipsPerPixelX
'             as when a tablet screen is rotated, the "Screen" object of VB doesn't respond to the change
'             so it has to be done by hand using GetDeviceCaps API.
'---------------------------------------------------------------------------------------
'
Public Function fTwipsPerPixelY() As Single
    Dim hDC As Long: hDC = 0
    Dim lPixelsPerInch As Long: lPixelsPerInch = 0
    
    Const LOGPIXELSY As Integer = 90        '  Logical pixels/inch in Y
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    
   On Error GoTo fTwipsPerPixelY_Error
   
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hDC = GetDC(0)
    If hDC <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY)
        ReleaseDC 0, hDC
        fTwipsPerPixelY = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        fTwipsPerPixelY = Screen.TwipsPerPixelY
    End If

   On Error GoTo 0
   Exit Function

fTwipsPerPixelY_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTwipsPerPixelY of Module Module1"

End Function

'---------------------------------------------------------------------------------------
' Procedure : fGetMonitorCount
' Author    : beededea
' Date      : 17/08/2024
' Purpose   : Return the count of the number of monitors using the EnumDisplayMonitors API to callback to MonitorEnumProc
'---------------------------------------------------------------------------------------
'
Public Function fGetMonitorCount() As Long
   On Error GoTo fGetMonitorCount_Error

    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, fGetMonitorCount

   On Error GoTo 0
   Exit Function

fGetMonitorCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetMonitorCount of Module monitorModule"
End Function

'---------------------------------------------------------------------------------------
' Procedure : MonitorEnumProc
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : Return the count of the number of monitors using the EnumDisplayMonitors API to callback to this function
'---------------------------------------------------------------------------------------
'
Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByRef dwData As Long) As Long
    On Error GoTo MonitorEnumProc_Error

    ReDim Preserve rcMonitors(dwData)
    rcMonitors(dwData) = lprcMonitor
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue

    On Error GoTo 0
    Exit Function

MonitorEnumProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MonitorEnumProc of Module monitorModule"
End Function

'---------------------------------------------------------------------------------------
' Procedure : setFormOnMonitor
' Author    : Hypetia from TekTips https://www.tek-tips.com/userinfo.cfm?member=Hypetia
' Date      : 01/03/2023
' Purpose   : Called on startup - restores the form's saved position and puts it on screen
'             if the form finds itself offscreen due to monitor position/resolution changes.
'---------------------------------------------------------------------------------------
'
Public Sub SetFormOnMonitor(ByRef hWnd As Long, ByVal Left As Long, ByVal Top As Long)

    Dim rc As RECT ' structure that receives the screen coordinate
    Dim hMonitor As Long: hMonitor = 0
    Dim mi As tagMONITORINFO
    
    On Error GoTo setFormOnMonitor_Error

    GetWindowRect hWnd, rc 'obtain the current form's window rectangle co-ords
        
    'move the window rectangle to the previously saved position supplied as two params.
    OffsetRect rc, Left - rc.Left, Top - rc.Top
    
    'find the monitor handle closest to our window rectangle
    hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    ' 3089353   1st monitor
    ' 436805389 2nd monitor
    
    'get monitor co-ordinates and working area
    mi.cbSize = Len(mi)
    GetMonitorInfo hMonitor, mi
    
    'adjust the window rectangle so it fits inside the work area of the monitor
    If rc.Left < mi.rcWork.Left Then OffsetRect rc, mi.rcWork.Left - rc.Left, 0
    If rc.Right > mi.rcWork.Right Then OffsetRect rc, mi.rcWork.Right - rc.Right, 0
    If rc.Top < mi.rcWork.Top Then OffsetRect rc, 0, mi.rcWork.Top - rc.Top
    If rc.Bottom > mi.rcWork.Bottom Then OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
    
    'move the window to new calculated position
    MoveWindow hWnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0

    On Error GoTo 0
    Exit Sub

setFormOnMonitor_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormOnMonitor of Module Module1"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cWidgetFormScreenProperties
' Author    :
' Date      : 23/01/2021
' Purpose   : provides the properties of the monitor upon which the supplied RC form's rectangle sits.
'             User supplies the RC form name.
'---------------------------------------------------------------------------------------
'
Public Function cWidgetFormScreenProperties(ByVal frm As cWidgetForm, ByRef monitorID As Long) As UDTMonitor
    
    Dim hMonitor As Long: hMonitor = 0
    Dim MONITORINFO As tagMONITORINFO
    Dim Frect As RECT
    Dim ad As Double: ad = 0
    
    On Error GoTo cWidgetFormScreenProperties_Error
   
    'If gblDebugFlg = 1 Then MsgBox "%" & " func cWidgetFormScreenProperties"
    
    ' reads the size and position of the user supplied form window
    GetWindowRect frm.hWnd, Frect
    hMonitor = MonitorFromRect(Frect, MONITOR_DEFAULTTOPRIMARY) ' get handle for monitor containing most of Frm
                                                                ' if disconnected return handle (and properties) for primary monitor
    On Error GoTo GetMonitorInformation_Err
    MONITORINFO.cbSize = Len(MONITORINFO)
    GetMonitorInfo hMonitor, MONITORINFO
    
    'Return the properties (in Twips) of the monitor upon which most of Frm is mapped
    With cWidgetFormScreenProperties
        .handle = hMonitor
        'convert all dimensions from pixels to twips
        .Left = MONITORINFO.rcMonitor.Left * gblScreenTwipsPerPixelX
        .Right = MONITORINFO.rcMonitor.Right * gblScreenTwipsPerPixelX
        .Top = MONITORINFO.rcMonitor.Top * gblScreenTwipsPerPixelY
        .Bottom = MONITORINFO.rcMonitor.Bottom * gblScreenTwipsPerPixelY

        .Height = (MONITORINFO.rcMonitor.Bottom - MONITORINFO.rcMonitor.Top) * gblScreenTwipsPerPixelY
        .Width = (MONITORINFO.rcMonitor.Right - MONITORINFO.rcMonitor.Left) * gblScreenTwipsPerPixelX

        .IsPrimary = MONITORINFO.dwFlags And MONITORINFOF_PRIMARY
    End With
    
    monitorID = hMonitor

    Exit Function
GetMonitorInformation_Err:
    Beep
    If Err.Number = 453 Then
        'should be handled if pre win2k compatibility is required
        'Non-Multimonitor OS, return -1
        'GetMonitorInformation = -1
        'etc
    End If

   On Error GoTo 0
   Exit Function

cWidgetFormScreenProperties_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cWidgetFormScreenProperties of Module common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : formScreenProperties
' Author    :
' Date      : 23/01/2021
' Purpose   : provides the properties of the monitor upon which the supplied VB6 form's rectangle sits.
'             User supplies the VB6 form name.
'---------------------------------------------------------------------------------------
'
Public Function formScreenProperties(ByVal frm As Form, ByRef monitorID As Long) As UDTMonitor
    
    Dim hMonitor As Long: hMonitor = 0
    Dim MONITORINFO As tagMONITORINFO
    Dim Frect As RECT
    Dim ad As Double: ad = 0
    
    On Error GoTo formScreenProperties_Error
   
    If gblDebugFlg = 1 Then MsgBox "%" & " func formScreenProperties"
    
    ' reads the size and position of the user supplied form window
    GetWindowRect frm.hWnd, Frect
    hMonitor = MonitorFromRect(Frect, MONITOR_DEFAULTTOPRIMARY) ' get handle for monitor containing most of Frm
                                                                ' if disconnected return handle (and properties) for primary monitor
    On Error GoTo GetMonitorInformation_Err
    MONITORINFO.cbSize = Len(MONITORINFO)
    GetMonitorInfo hMonitor, MONITORINFO
    
    'Return the properties (in Twips) of the monitor upon which most of Frm is mapped
    With formScreenProperties
        .handle = hMonitor
        'convert all dimensions from pixels to twips
        .Left = MONITORINFO.rcMonitor.Left * gblScreenTwipsPerPixelX
        .Right = MONITORINFO.rcMonitor.Right * gblScreenTwipsPerPixelX
        .Top = MONITORINFO.rcMonitor.Top * gblScreenTwipsPerPixelY
        .Bottom = MONITORINFO.rcMonitor.Bottom * gblScreenTwipsPerPixelY

        .Height = (MONITORINFO.rcMonitor.Bottom - MONITORINFO.rcMonitor.Top) * gblScreenTwipsPerPixelY
        .Width = (MONITORINFO.rcMonitor.Right - MONITORINFO.rcMonitor.Left) * gblScreenTwipsPerPixelX

        .IsPrimary = MONITORINFO.dwFlags And MONITORINFOF_PRIMARY
    End With
    
    monitorID = hMonitor

    Exit Function
GetMonitorInformation_Err:
    Beep
    If Err.Number = 453 Then
        'should be handled if pre win2k compatibility is required
        'Non-Multimonitor OS, return -1
        'GetMonitorInformation = -1
        'etc
    End If

   On Error GoTo 0
   Exit Function

formScreenProperties_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure formScreenProperties of Module common"
End Function



'---------------------------------------------------------------------------------------
' Procedure : positionPrefsByMonitorSize
' Author    : beededea
' Date      : 20/08/2024
' Purpose   : at startup obtains monitor ID and characteristics
'             in addition, if there is more than one screen, size the form by a ratio according to the form's physical monitor properties
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsByMonitorSize()

    Static oldWidgetPrefsLeft As Long
    Static oldWidgetPrefsTop As Long
    Static beenMovingFlg As Boolean
    
    'Static oldPrefsFormMonitorID As Long
    Static oldPrefsMonitorStructWidthTwips As Long
    Static oldPrefsMonitorStructHeightTwips As Long
    Static oldPrefsGaugeLeftPixels As Long
        
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    Dim prefsFormMonitorPrimary As Long: prefsFormMonitorPrimary = 0
    Dim monitorStructWidthTwips As Long: monitorStructWidthTwips = 0
    Dim monitorStructHeightTwips As Long: monitorStructHeightTwips = 0
    Dim resizeProportion As Double: resizeProportion = 0
    Dim newPrefsHeight As Single: newPrefsHeight = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    ' calls a routine that tests for a change in the monitor upon which the form sits, if so, resizes
    On Error GoTo positionPrefsByMonitorSize_Error

    ' if just one monitor or the global switch is off then exit
    If gblMonitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
    
        ' turn off the timer that saves the prefs height and position
        widgetPrefs.tmrPrefsMonitorSaveHeight.Enabled = False
        widgetPrefs.tmrWritePosition.Enabled = False
   
        ' populate the OLD vars if empty, to allow valid comparison next run
        If oldWidgetPrefsLeft <= 0 Then oldWidgetPrefsLeft = widgetPrefs.Left
        If oldWidgetPrefsTop <= 0 Then oldWidgetPrefsTop = widgetPrefs.Top

       ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
        prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
        
        ' test whether the current monitor is the primary
        prefsFormMonitorPrimary = prefsMonitorStruct.IsPrimary ' -1 true
        
        ' sample the physical monitor resolution
        monitorStructWidthTwips = prefsMonitorStruct.Width
        monitorStructHeightTwips = prefsMonitorStruct.Height
        
        ' store other values as 'old' vars for latter comparison and usage
        If oldPrefsMonitorStructWidthTwips = 0 Then oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
        If oldPrefsMonitorStructHeightTwips = 0 Then oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
        If oldPrefsGaugeLeftPixels = 0 Then oldPrefsGaugeLeftPixels = widgetPrefs.Left
    
        ' if the monitor ID has changed
        If gblOldPrefsFormMonitorPrimary <> prefsFormMonitorPrimary Then
    
            ' screenWrite ("Prefs Stored monitor primary status = " & CBool(gblOldPrefsFormMonitorPrimary))
            ' screenWrite ("Prefs Current monitor primary status = " & CBool(prefsFormMonitorPrimary))
           
            If LTrim$(gblMultiMonitorResize) = "1" Then
                'if the resolution is different then calculate new size proportion
                If monitorStructWidthTwips <> oldPrefsMonitorStructWidthTwips Or monitorStructHeightTwips <> oldPrefsMonitorStructHeightTwips Then
                    'now calculate the size of the widget according to the screen HeightTwips.
                    resizeProportion = prefsMonitorStruct.Height / oldPrefsMonitorStructHeightTwips
                    newPrefsHeight = widgetPrefs.Height * resizeProportion
                    gblPrefsFormResizedInCode = True
                    widgetPrefs.Height = newPrefsHeight
                End If
            ElseIf LTrim$(gblMultiMonitorResize) = "2" Then
                ' set the widget size according to saved values
                gblPrefsFormResizedInCode = True
                If prefsMonitorStruct.IsPrimary = True Then
                    widgetPrefs.Height = CLng(gblPrefsPrimaryHeightTwips)
                Else
                    widgetPrefs.Height = CLng(gblPrefsSecondaryHeightTwips)
                End If
            End If
            
        End If
        
        ' set the current values as 'old' for comparison on next run
        gblOldPrefsFormMonitorPrimary = prefsFormMonitorPrimary
        
        oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
        oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
        oldPrefsGaugeLeftPixels = widgetPrefs.Left

    End If

    oldWidgetPrefsLeft = widgetPrefs.Left
    oldWidgetPrefsTop = widgetPrefs.Top
    
    ' restart any timers that position the prefs and store position/size values
    widgetPrefs.tmrPrefsMonitorSaveHeight.Enabled = True
    widgetPrefs.tmrWritePosition.Enabled = True

   On Error GoTo 0
   Exit Sub

positionPrefsByMonitorSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsByMonitorSize of Module monitorModule"
    
    End Sub


'Function EnumMonitors(F As Form) As Long
'    Dim N As Long
'    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
'    With F
'        .Move .Left, .Top, (rcVS.Right - rcVS.Left) * 2 + .Width - .ScaleWidth, (rcVS.Bottom - rcVS.Top) * 2 + .Height - .ScaleHeight
'    End With
'    F.Scale (rcVS.Left, rcVS.Top)-(rcVS.Right, rcVS.Bottom)
'    F.Caption = N & " Monitor" & IIf(N > 1, "s", vbNullString)
'    F.lblMonitors(0).Appearance = 0 'Flat
'    F.lblMonitors(0).BorderStyle = 1 'FixedSingle
'    For N = 0 To N - 1
'        If N Then
'            Load F.lblMonitors(N)
'            F.lblMonitors(N).Visible = True
'        End If
'        With rcMonitors(N)
'            F.lblMonitors(N).Move .Left, .Top, .Right - .Left, .Bottom - .Top
'            F.lblMonitors(N).Caption = "Monitor " & N + 1 & vbLf & _
'                .Right - .Left & " x " & .Bottom - .Top & vbLf & _
'                "(" & .Left & ", " & .Top & ")-(" & .Right & ", " & .Bottom & ")"
'        End With
'    Next
'End Function


'---------------------------------------------------------------------------------------
' Procedure : fVirtualScreenWidth
' Author    : beededea
' Date      : 17/08/2024
' Purpose   : Determines the whole screen width including any virtual 'extra' caused by multiple monitor positioning.
'             Called on startup and via tmrScreenResolution_Timer to test whether the width of the current monitor
'             where the form currently sits, has changed.
'---------------------------------------------------------------------------------------
'
Public Function fVirtualScreenWidth(ByRef inPixels As Boolean) As Long
    ' This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
    Dim Pixels As Long: Pixels = 0
    Const SM_CXVIRTUALSCREEN = 78
    '
   On Error GoTo fVirtualScreenWidth_Error

    Pixels = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    If inPixels = True Then
        fVirtualScreenWidth = Pixels
    Else
        fVirtualScreenWidth = Pixels * gblScreenTwipsPerPixelX
    End If

   On Error GoTo 0
   Exit Function

fVirtualScreenWidth_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fVirtualScreenWidth of Module monitorModule"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fVirtualScreenHeight
' Author    : beededea
' Date      : 14/02/2025
' Purpose   : Determines the whole screen height including any virtual 'extra' caused by multiple monitor positioning.
'             Called on startup and via tmrScreenResolution_Timer to test whether the height of the current monitor
'             where the form currently sits, has changed.
'---------------------------------------------------------------------------------------
'
Public Function fVirtualScreenHeight(ByRef inPixels As Boolean, Optional ByRef bSubtractTaskbar As Boolean = False) As Long
    ' This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
    Dim Pixels As Long: Pixels = 0
    Const CYVIRTUALSCREEN = 79
    '
   On Error GoTo fVirtualScreenHeight_Error

    Pixels = GetSystemMetrics(CYVIRTUALSCREEN)
    If bSubtractTaskbar Then
        ' The taskbar is typically 30 pixels or 450 twips, or, at least, this is the assumption made here.
        ' It can actually be multiples of this, or possibly moved to the side or top.
        ' This procedure does not account for these possibilities.
        fVirtualScreenHeight = (Pixels - 30)
    Else
        fVirtualScreenHeight = Pixels
    End If
    
    If inPixels = True Then
        fVirtualScreenHeight = fVirtualScreenHeight
    Else
        fVirtualScreenHeight = fVirtualScreenHeight * gblScreenTwipsPerPixelY
    End If

   On Error GoTo 0
   Exit Function

fVirtualScreenHeight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fVirtualScreenHeight of Module monitorModule"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : positionRCFormByMonitorSize
' Author    : beededea
' Date      : 20/08/2024
' Purpose   : at startup obtains monitor ID and characteristics
'             in addition, if there is more than one screen, size the form by a ratio according to the form's physical monitor properties
'---------------------------------------------------------------------------------------
'
Public Sub positionRCFormByMonitorSize()
    
    Static oldMonitorStructWidthTwips As Long
    Static oldMonitorStructHeightTwips As Long
    Static oldGaugeLeftPixels As Long
        
    Dim gaugeFormMonitorPrimary As Long: gaugeFormMonitorPrimary = 0
    Dim gaugeFormMonitorID As Long: gaugeFormMonitorID = 0
    
    Dim monitorStructWidthTwips As Long: monitorStructWidthTwips = 0
    Dim monitorStructHeightTwips As Long: monitorStructHeightTwips = 0
    Dim resizeProportion As Double: resizeProportion = 0

    On Error GoTo positionRCFormByMonitorSize_Error
  
    If gblMonitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
                    
        ' note the monitor ID at gaugeForm form_load and store as the gaugeFormMonitorID
        gaugeMonitorStruct = cWidgetFormScreenProperties(fGauge.gaugeForm, gaugeFormMonitorID)
        
        gaugeFormMonitorPrimary = gaugeMonitorStruct.IsPrimary
        
        If fGauge.gaugeForm.Left = oldGaugeLeftPixels Then Exit Sub ' this can only work if the reposition is being performed by the timer
        ' we are also calling it on a mouseUP event, so the comparison to original position is lost to us
    
        ' sample the physical monitor resolution
        monitorStructWidthTwips = gaugeMonitorStruct.Width
        monitorStructHeightTwips = gaugeMonitorStruct.Height
                
        If oldMonitorStructWidthTwips = 0 Then oldMonitorStructWidthTwips = monitorStructWidthTwips
        If oldMonitorStructHeightTwips = 0 Then oldMonitorStructHeightTwips = monitorStructHeightTwips
        If oldGaugeLeftPixels = 0 Then oldGaugeLeftPixels = fGauge.gaugeForm.Left
    
        If gblOldgaugeFormMonitorPrimary <> gaugeFormMonitorPrimary Then
            
            ' screenWrite ("Stored monitor primary status = " & CBool(gblOldgaugeFormMonitorPrimary))
            ' screenWrite ("Current monitor primary status = " & CBool(gaugeFormMonitorPrimary))
            
            If LTrim$(gblMultiMonitorResize) = "1" Then
                'if the resolution is different then calculate new size proportion
                If monitorStructWidthTwips <> oldMonitorStructWidthTwips Or monitorStructHeightTwips <> oldMonitorStructHeightTwips Then
                    ' screenWrite ("Resizing by proportion per monitor ")
                    
                    'now calculate the size of the widget according to the screen HeightTwips.
                    resizeProportion = gaugeMonitorStruct.Height / oldMonitorStructHeightTwips
                    resizeProportion = (Val(gblGaugeSize) / 100) * resizeProportion
                    
                    'if  dragging from right to left then reposition
                    If fGauge.gaugeForm.Left > oldGaugeLeftPixels Then
                        fGauge.gaugeForm.Left = fGauge.gaugeForm.Left + fGauge.gaugeForm.Widgets("maincasingsurround").Widget.Left
                    Else
                        fGauge.gaugeForm.Left = fGauge.gaugeForm.Left - fGauge.gaugeForm.Widgets("maincasingsurround").Widget.Left
                    End If
                    fGauge.gaugeForm.Refresh
                    Call fGauge.AdjustZoom(resizeProportion)
                End If
            ElseIf LTrim$(gblMultiMonitorResize) = "2" Then
                ' screenWrite ("Resizing per monitor stored size ")
                If gaugeMonitorStruct.IsPrimary = True Then
                    If gblGaugePrimaryHeightRatio = "" Then gblGaugePrimaryHeightRatio = "1"
                    resizeProportion = Val(gblGaugePrimaryHeightRatio)
                Else
                    If gblGaugeSecondaryHeightRatio = "" Then gblGaugeSecondaryHeightRatio = "1"
                    resizeProportion = Val(gblGaugeSecondaryHeightRatio)
                End If
                
                                    
                'if  dragging from right to left then reposition
                If fGauge.gaugeForm.Left > oldGaugeLeftPixels Then
                    fGauge.gaugeForm.Left = fGauge.gaugeForm.Left + fGauge.gaugeForm.Widgets("maincasingsurround").Widget.Left
                Else
                    fGauge.gaugeForm.Left = fGauge.gaugeForm.Left - fGauge.gaugeForm.Widgets("maincasingsurround").Widget.Left
                End If
                fGauge.gaugeForm.Refresh
                Call fGauge.AdjustZoom(resizeProportion)
            End If
        End If
    
        gblOldgaugeFormMonitorPrimary = gaugeFormMonitorPrimary
        
        oldMonitorStructWidthTwips = monitorStructWidthTwips
        oldMonitorStructHeightTwips = monitorStructHeightTwips
        oldGaugeLeftPixels = fGauge.gaugeForm.Left
    End If

   On Error GoTo 0
   Exit Sub

positionRCFormByMonitorSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionRCFormByMonitorSize of Module Module1"

End Sub




