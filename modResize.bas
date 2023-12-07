Attribute VB_Name = "modResize"
Option Explicit

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Public Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Public prefsControlPositions() As ControlPositionType
Public prefsCurrentWidth As Double
Public prefsCurrentHeight As Double

Public msgBoxAControlPositions() As ControlPositionType
Public msgBoxACurrentWidth As Double
Public msgBoxACurrentHeight As Double

'---------------------------------------------------------------------------------------
' Procedure : ResizeControls
' Author    : adapted from Rod Stephens @ vb-helper.com
' Date      : 16/04/2021
' Purpose   : Arrange the controls for a new size.
'---------------------------------------------------------------------------------------
'
Public Sub resizeControls(ByRef thisForm As Form, ByRef m_ControlPositions() As ControlPositionType, ByVal m_FormWid As Double, ByVal m_FormHgt As Double, ByVal formFontSize As Long)
    Dim I As Integer: I = 0
    Dim Ctrl As Control
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
        
    On Error GoTo ResizeControls_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt
    
    I = 1

    For Each Ctrl In thisForm.Controls
        With m_ControlPositions(I)
            If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is Image) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is Slider) Then

                If (TypeOf Ctrl Is Image) Then

                    Ctrl.Stretch = True
                    Ctrl.Left = x_scale * .Left
                    Ctrl.Top = y_scale * .Top
                    Ctrl.Width = x_scale * .Width
                    Ctrl.Height = Ctrl.Width ' always square in our case

                    Ctrl.Refresh
                Else
                    Ctrl.Left = x_scale * .Left
                    Ctrl.Top = y_scale * .Top
                    Ctrl.Width = x_scale * .Width
                    If Not (TypeOf Ctrl Is ComboBox) Then
                        ' Cannot change height of ComboBoxes.
                        Ctrl.Height = y_scale * .Height
                    End If
                    On Error Resume Next
                    Ctrl.Font.Size = y_scale * formFontSize
                    Ctrl.Refresh
                    On Error GoTo 0
                End If
            End If
        End With
        I = I + 1
    Next Ctrl
        
   On Error GoTo 0
   Exit Sub

ResizeControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ResizeControls of Form formSoftwareList"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : SaveSizes
' Author    : Rod Stephens vb-helper.com
' Date      : 16/04/2021
' Purpose   : Resize controls to fit when a form resizes
'             Save the form's and controls' dimensions.
' Credit    : Rod Stephens vb-helper.com
'---------------------------------------------------------------------------------------
'
Public Sub SaveSizes(ByVal thisForm As Form, ByRef m_ControlPositions() As ControlPositionType, ByRef m_FormWid As Double, ByRef m_FormHgt As Double)
    Dim I As Integer: I = 0
    Dim Ctrl As Control

    ' Save the controls' positions and sizes.
    On Error GoTo SaveSizes_Error

    ReDim m_ControlPositions(1 To thisForm.Controls.Count)
    I = 1
    For Each Ctrl In thisForm.Controls
        With m_ControlPositions(I)
        
            If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is Image) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is Slider) Then
                .Left = Ctrl.Left
                .Top = Ctrl.Top
                .Width = Ctrl.Width
                .Height = Ctrl.Height
                On Error Resume Next ' cater for any controls that do not have a font property that may cause an error
                .FontSize = Ctrl.Font.Size
                On Error GoTo 0
            End If
        End With
        I = I + 1
    Next Ctrl

    ' Save the form's size.
    m_FormWid = thisForm.ScaleWidth
    m_FormHgt = thisForm.ScaleHeight

   On Error GoTo 0
   Exit Sub

SaveSizes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSizes of Form formSoftwareList"
End Sub
