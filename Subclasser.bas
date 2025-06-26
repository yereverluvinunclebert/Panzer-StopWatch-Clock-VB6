Attribute VB_Name = "Subclasser"
'---------------------------------------------------------------------------------------
' Module    : Subclasser
' Author    : Elroy
' Date      : 16/07/2024
' Purpose   : used to subclass specific named controls
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function GetWindowSubclass Lib "comctl32" Alias "#411" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function NextSubclassProcOnChain Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (ByRef dstObject As Any, ByRef srcObjPtr As Any) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFOSTRUCTURE) As Long

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long ' This is +1 (right - left = width)
    Bottom As Long ' This is +1 (bottom - top = height)
End Type
'
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    ItemAction As Long
    ItemState As Long       ' Bitflags: ODS_COMBOBOXEDIT    = &h1000& (edit control being drawn).
                            '           ODS_SELECTED        = &h0001&
                            '           ODS_DISABLED        = &h0004&
                            '           ODS_FOCUS           = &h0010&
                            '           ODS_NOACCEL         = &h0100&
                            '           ODS_NOFOCUSRECT     = &h0200&
                            '           Others, but they don't apply to combobox.
                            '
    hWndItem As Long        ' hWnd to the ComboBox.
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type
'
Private Type COMBOBOXINFOSTRUCTURE
    cbSize          As Long
    rcItem          As RECT
    rcButton        As RECT
    stateButton     As Long
    hwndCombo       As Long
    hwndEdit        As Long
    hwndList        As Long
End Type

    



Private Sub SubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, Optional dwRefData As Long, Optional uIdSubclass As Long)
    If uIdSubclass = 0& Then uIdSubclass = hWnd
    Call SetWindowSubclass(hWnd, AddressOf_ProcToSubclass, uIdSubclass, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, Optional uIdSubclass As Long)
    If uIdSubclass = 0& Then uIdSubclass = hWnd
    Call RemoveWindowSubclass(hWnd, AddressOf_ProcToSubclass, uIdSubclass)
End Sub

'Public Sub SubclassMouseWheel(CtlHwnd As Long, TheObjPtr As Long)
'    SubclassSomeWindow CtlHwnd, AddressOf MouseWheel_Proc, TheObjPtr
'End Sub

Public Sub SubclassComboBox(CtlHwnd As Long, TheObjPtr As Long)
    SubclassSomeWindow CtlHwnd, AddressOf ComboBox_Proc, TheObjPtr
End Sub

Public Sub SubclassForm(CtlHwnd As Long, TheObjPtr As Long)
    SubclassSomeWindow CtlHwnd, AddressOf Form_Proc, TheObjPtr
End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : MouseWheel_Proc
'' Author    : beededea
'' Date      : 16/07/2024
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Function MouseWheel_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
'
'    Const WM_DESTROY                    As Long = &H2&  ' All other needed constants are declared within the procedures.
'    Const WM_MOUSEWHEEL As Long = &H20A&
'    Dim fra             As Object
'
'   On Error GoTo MouseWheel_Proc_Error
'
'    If uMsg = WM_DESTROY Then
'        UnSubclassSomeWindow hWnd, AddressOf_MouseWheel_Proc, uIdSubclass
'        MouseWheel_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
'        Exit Function
'    End If
'
'    If uMsg = WM_MOUSEWHEEL Then     ' Mouse-Wheel.
'        Set fra = ComObjectFromPtr(dwRefData)
'        On Error Resume Next        ' Protect in case programmer forgot to put in procedure.
'            fra.Parent.MouseMoveOnFrame fra.Name, wParam
'        On Error GoTo 0
'        Set fra = Nothing
'    End If
'
'
'    ' If we fell out, just proceed as normal.
'    MouseWheel_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
'
'   On Error GoTo 0
'   Exit Function
'
'MouseWheel_Proc_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseWheel_Proc of Module Subclasser"
'
'End Function
    

'---------------------------------------------------------------------------------------
' Procedure : ComboBox_Proc
' Author    : Elroy
' Date      : 16/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ComboBox_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Const WM_DESTROY                    As Long = &H2&  ' All other needed constants are declared within the procedures.
   On Error GoTo ComboBox_Proc_Error

    If uMsg = WM_DESTROY Then
        UnSubclassSomeWindow hWnd, AddressOf_ComboBox_Proc, uIdSubclass
        ComboBox_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    '
    Dim uDrawItem       As DRAWITEMSTRUCT
    Dim oBrush          As Long
    Dim oThePen         As Long
    Dim iRet            As Long
    Dim sText           As String
    Static sPrevText    As String
    Dim cbo             As Object
    '
    Const WM_DRAWITEM           As Long = &H2B&
    Const ODT_COMBOBOX          As Long = 3&
    Const DC_PEN                As Long = 19&
    Const DC_BRUSH              As Long = 18&
    Const TRANSPARENT           As Long = 1&
    Const COLOR_WINDOW          As Long = 5&
    Const COLOR_WINDOWTEXT      As Long = 8&
    Const COLOR_HIGHLIGHT       As Long = 13&
    Const COLOR_HIGHLIGHTTEXT   As Long = 14&
    Const CB_GETLBTEXT          As Long = &H148&
    Const CB_GETLBTEXTLEN       As Long = &H149&
    Const DT_SINGLELINE         As Long = &H20&
    Const DT_VCENTER            As Long = &H4&
    Const DT_NOPREFIX           As Long = &H800&
    Const ODS_SELECTED          As Long = &H1&
    Const ODS_COMBOBOXEDIT      As Long = &H1000& ' (edit control being drawn).
    Const WM_SETCURSOR          As Long = &H20&
    '
    If uMsg = WM_SETCURSOR Then     ' Mouse-Move.
        Set cbo = ComObjectFromPtr(dwRefData)
        On Error Resume Next        ' Protect in case programmer forgot to put in procedure.
            cbo.Parent.MouseMoveOnComboText cbo.Name
        On Error GoTo 0
        Set cbo = Nothing
    End If
    '
    ' If we fell out, just proceed as normal.
    ComboBox_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

   On Error GoTo 0
   Exit Function

ComboBox_Proc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComboBox_Proc of Module Subclasser"
End Function




'---------------------------------------------------------------------------------------
' Procedure : Form_Proc
' Author    : Elroy
' Date      : 16/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function Form_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Const WM_DESTROY            As Long = &H2&  ' All other needed constants are declared within the procedures.
    'Const WM_MOVE               As Long = &H3  ' called all during any form move
    Const WM_EXITSIZEMOVE       As Long = &H232 ' called only when all movement is completed
    
    Dim frm As Object
    
    On Error GoTo Form_Proc_Error

    If uMsg = WM_DESTROY Then
        UnSubclassSomeWindow hWnd, AddressOf_Form_Proc, uIdSubclass
        Form_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    '
    If uMsg = WM_EXITSIZEMOVE Then     ' Mouse-Move.
        Set frm = ComObjectFromPtr(dwRefData)
        On Error Resume Next        ' Protect in case programmer forgot to put in procedure.
            frm.Form_Moved frm.Name
        On Error GoTo 0
        Set frm = Nothing
    End If
    
    '
    ' If we fell out, just proceed as normal.
    Form_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)

   On Error GoTo 0
   Exit Function

Form_Proc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Proc of Module Subclasser"
End Function

'Private Function AddressOf_MouseWheel_Proc() As Long
'    AddressOf_MouseWheel_Proc = ProcedureAddress(AddressOf MouseWheel_Proc)
'End Function

Private Function AddressOf_ComboBox_Proc() As Long
    AddressOf_ComboBox_Proc = ProcedureAddress(AddressOf ComboBox_Proc)
End Function

Private Function AddressOf_Form_Proc() As Long
    AddressOf_Form_Proc = ProcedureAddress(AddressOf Form_Proc)
End Function

Private Function ProcedureAddress(AddressOf_TheProc As Long) As Long
    ProcedureAddress = AddressOf_TheProc
End Function

Private Function ComObjectFromPtr(ByVal Ptr As Long) As IUnknown
    ' Ideas for these were initially shown to me by The Trick.
    ' This uses the pointer to an existing (instantiated) COM object and makes another reference to it.
    ' This reference is handled completely correctly in that Addref is performed for it.
    ' Usage: Set ObjAlias = ComObjectFromPtr(ObjPtr(ObjOriginal))
    vbaObjSetAddref ComObjectFromPtr, ByVal Ptr
End Function



Public Function cboEditHwndFromHwnd(ByVal cboHwnd As Long) As Long
    ' Returns the hWnd to the EDIT control within a ComboBox.
    Dim cbi As COMBOBOXINFOSTRUCTURE
    '
    cbi.cbSize = LenB(cbi)
    GetComboBoxInfo cboHwnd, cbi
    cboEditHwndFromHwnd = cbi.hwndEdit
End Function
