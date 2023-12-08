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
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
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
Public WithEvents licenceForm As cWidgetForm
Attribute licenceForm.VB_VarHelpID = -1

Public licenceFormX As Long
Public licenceFormY As Long
' Public licenceFormZ As Single

Private WithEvents btnAccept As cwButton
Attribute btnAccept.VB_VarHelpID = -1
Private WithEvents btnDecline As cwButton
Attribute btnDecline.VB_VarHelpID = -1
Private WithEvents textBox As cwTextBox
Attribute textBox.VB_VarHelpID = -1
Private lblDeclineOrNot As cwLabel
Attribute lblDeclineOrNot.VB_VarHelpID = -1



'---------------------------------------------------------------------------------------
' Procedure : initAndShowAboutForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initAndShowAboutForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim aboutFormWidth As Long: aboutFormWidth = 0
    Dim aboutFormHeight As Long: aboutFormHeight = 0
    Dim aboutWidgetWidth As Long: aboutWidgetWidth = 0
    Dim aboutWidgetHeight As Long: aboutWidgetHeight = 0
    Dim aboutWidgetZoomValue As Double: aboutWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowAboutForm_Error
    
    If PzGDpiAwareness = "0" Then
        aboutFormWidth = 900
        aboutFormHeight = 850
        If aboutFormWidth > screenWidthPixels Then aboutFormWidth = screenWidthPixels
        If aboutFormHeight > screenHeightPixels Then aboutFormHeight = screenHeightPixels
    Else
        aboutFormWidth = 1800
        aboutFormHeight = 1600
        If aboutFormWidth > screenWidthPixels Then
            aboutFormWidth = screenWidthPixels
            aboutWidgetZoomValue = 1
        Else
            aboutWidgetZoomValue = 2
        End If
        If aboutFormHeight > screenHeightPixels Then
            aboutFormHeight = screenHeightPixels
            aboutWidgetZoomValue = 1
        Else
            aboutWidgetZoomValue = 2
        End If
    End If
    aboutWidgetWidth = aboutFormWidth - 100
    aboutWidgetHeight = aboutFormHeight - 100
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", True, aboutFormWidth, aboutFormHeight)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
    
    ' create a new about widget with a name, position and size, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 1, 1, aboutWidgetWidth, aboutWidgetHeight)
    
    If PzGDpiAwareness = "0" Then
        aboutWidget.Zoom = aboutWidgetZoomValue
    Else
        aboutWidget.Zoom = aboutWidgetZoomValue
    End If
    
   On Error GoTo 0
   Exit Sub

initAndShowAboutForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowAboutForm of Class Module cfMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initAndShowHelpForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initAndShowHelpForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim helpFormWidth As Long: helpFormWidth = 0
    Dim helpFormHeight As Long: helpFormHeight = 0
    Dim helpWidgetWidth As Long: helpWidgetWidth = 0
    Dim helpWidgetHeight As Long: helpWidgetHeight = 0
    
    Dim helpWidgetZoomValue As Double: helpWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowHelpForm_Error
    
    If PzGDpiAwareness = "0" Then
        helpFormWidth = 900
        helpFormHeight = 800
        If helpFormWidth > screenWidthPixels Then helpFormWidth = screenWidthPixels
        If helpFormHeight > screenHeightPixels Then helpFormHeight = screenHeightPixels
        helpWidgetZoomValue = 0.8
    Else
        helpFormWidth = 1500
        helpFormHeight = 1300
        If helpFormWidth > screenWidthPixels Then
            helpFormWidth = screenWidthPixels
            helpWidgetZoomValue = 1
        Else
            helpWidgetZoomValue = 1.4
        End If
        
        If helpFormHeight > screenHeightPixels Then
            helpFormHeight = screenHeightPixels
            helpWidgetZoomValue = 1
        Else
            helpWidgetZoomValue = 1.4
        End If
        
    End If
    helpWidgetWidth = helpFormWidth - 100
    helpWidgetHeight = helpFormHeight - 100
    
    ' create an invisible form using a Cairo Widget Form with a predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , helpFormWidth, helpFormHeight)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
         
    ' create a new help widget image with a name, position and size
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 0, 0, helpWidgetWidth, helpWidgetHeight)
       
     If PzGDpiAwareness = "0" Then
        helpWidget.Zoom = helpWidgetZoomValue
    Else
        helpWidget.Zoom = helpWidgetZoomValue
    End If

   On Error GoTo 0
   Exit Sub

initAndShowHelpForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowHelpForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : initAndShowLicenceForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initAndShowLicenceForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim licenceFormWidth As Long: licenceFormWidth = 0
    Dim licenceFormHeight As Long: licenceFormHeight = 0
    Dim licenceWidgetWidth As Long: licenceWidgetWidth = 0
    Dim licenceWidgetHeight As Long: licenceWidgetHeight = 0
    
    Dim licenceWidgetZoomValue As Double: licenceWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowLicenceForm_Error
    
    If PzGDpiAwareness = "0" Then
        licenceFormWidth = 650
        licenceFormHeight = 550
        If licenceFormWidth > screenWidthPixels Then licenceFormWidth = screenWidthPixels
        If licenceFormHeight > screenHeightPixels Then licenceFormHeight = screenHeightPixels
        licenceWidgetZoomValue = 0.8
    Else
        licenceFormWidth = 750
        licenceFormHeight = 650
        If licenceFormWidth > screenWidthPixels Then
            licenceFormWidth = screenWidthPixels
            licenceWidgetZoomValue = 0.8
        Else
            licenceWidgetZoomValue = 1.5
        End If
        
        If licenceFormHeight > screenHeightPixels Then
            licenceFormHeight = screenHeightPixels
            licenceWidgetZoomValue = 0.8
        Else
            licenceWidgetZoomValue = 1.5
        End If
        
    End If
    licenceWidgetWidth = licenceFormWidth - 100
    licenceWidgetHeight = licenceFormHeight - 100
           
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set licenceForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, True, licenceFormWidth, licenceFormHeight)
    
    ' make the licence form transparent
    licenceForm.WidgetRoot.BackColor = -1
         
    ' create a new licence class widget with a name, position and size
    Set licenceWidget = licenceForm.Widgets.Add(New cwLicence, "licence", 0, 0, licenceWidgetWidth, licenceWidgetHeight)
    
    If PzGDpiAwareness = "0" Then
        licenceWidget.Zoom = licenceWidgetZoomValue
    Else
        licenceWidget.Zoom = licenceWidgetZoomValue
    End If

    ' add two RC predefined button controls to the licence form
    If PzGDpiAwareness = "1" Then
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 352, 397, 70, 29)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 428, 397, 70, 29)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 50, 60, 440, 280)
        Set lblDeclineOrNot = licenceForm.Widgets.Add(New_W("cwLabel"), "lblDeclineOrNot", 50, 345, 440, 35)
        textBox.Widget.FontSize = 11
        btnDecline.Widget.FontSize = 10
        btnAccept.Widget.FontSize = 10
        lblDeclineOrNot.Widget.FontSize = 8
    Else
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 189, 210, 40, 18)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 228, 210, 40, 18)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 25, 30, 240, 155)
        Set lblDeclineOrNot = licenceForm.Widgets.Add(New_W("cwLabel"), "lblDeclineOrNot", 25, 188, 240, 20)
        textBox.Widget.FontSize = 8
        btnDecline.Widget.FontSize = 7
        btnAccept.Widget.FontSize = 7
        lblDeclineOrNot.Widget.FontSize = 6
    End If
    
    btnDecline.Caption = "Decline"
    btnDecline.Widget.ForeColor = vbWhite
    btnDecline.Widget.Alpha = 0.05
    btnDecline.Widget.ToolTip = "Press this button to decline the terms of the GPL licence."
    
    btnAccept.Caption = "Accept"
    btnAccept.Widget.ForeColor = vbWhite
    btnAccept.Widget.Alpha = 0.05
    btnAccept.Widget.ToolTip = "Press this button to accept the terms of the GPL licence."
    
    textBox.MultiLine = True
    textBox.ScrollBars = vbVertical
    textBox.Widget.BackColor = &H4000&
    textBox.Widget.ForeColor = vbWhite
    
    Call LoadFileToTB(textBox, App.path & "\resources\txt\licence.txt", False)

    textBox.Text = widgetName & vbCrLf & vbCrLf & _
    "Copyright © 1992-2023 Dean Beedell" & vbCrLf & vbCrLf & _
    "This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public Licence as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version." & vbCrLf & vbCrLf & _
        "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." & vbCrLf & vbCrLf & _
        "You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA" & vbCrLf & vbCrLf & _
        "If you use this software in any way whatsoever then that implies acceptance of the licence. If you do not wish to comply with the licence terms then please remove the download, binary and source code from your systems immediately." & vbCrLf & vbCrLf & _
        "You may use any of my own imagery in your own creations but commercially only with my permission. In all other non-commercial cases I require a credit to the original artist using my name or one of my pseudonyms and a link to my site. With regard to the commercial use of incorporated images, permission and a licence would need to be obtained from the original owner and creator, ie. me."

    lblDeclineOrNot.borderWidth = 0
    lblDeclineOrNot.WordWrap = True
    lblDeclineOrNot.Alignment = vbCenter
    lblDeclineOrNot.Widget.BackColor = -1
    lblDeclineOrNot.Widget.ForeColor = vbWhite
    
    lblDeclineOrNot.Caption = "If you do not agree with the terms set forth above, please click the " & _
        "Decline button below and destroy this widget and its documentation"
        
   On Error GoTo 0
   Exit Sub

initAndShowLicenceForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowLicenceForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnAccept_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAccept_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim slicence As String: slicence = "0"
    
    On Error GoTo btnAccept_Click_Error
    
    fileToPlay = "ting.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    answer = msgBoxA("User accepted licence, thankyou.", vbOKOnly + vbExclamation, "Licence Acceptance", True, "btnAcceptClick")
    
    slicence = "1"
    sPutINISetting "Software\PzStopWatch", "licence", slicence, PzGSettingsFile
    
    licenceWidget.ShowMe = False
    licenceWidget.MyOpacity = 0
    licenceWidget.Widget.Refresh

    licenceForm.Visible = False
    
   On Error GoTo 0
   Exit Sub

btnAccept_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAccept_Click of Class Module cfMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDecline_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDecline_Click()
    
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim slicence As String: slicence = "0"
        
    On Error GoTo btnDecline_Click_Error

    'answer = MsgBox("Licence declined, please destroy your copy of this widget and its documentation.", vbYes)
    answer = msgBoxA("Licence declined, please destroy your copy of this widget and its documentation. This widget will now unload.", vbOKOnly + vbExclamation, "Licence Declined", False)
    slicence = "0"
    sPutINISetting "Software\PzStopWatch", "licence", slicence, PzGSettingsFile

    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

btnDecline_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDecline_Click of Class Module cfMain"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : licenceForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub licenceForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo licenceForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If PzGIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not licenceForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      licenceFormX = (x - x0) * licenceForm.WidgetRoot.Zoom + licenceForm.Left
      licenceFormY = (y - y0) * licenceForm.WidgetRoot.Zoom + licenceForm.Top
      licenceForm.Move licenceFormX, licenceFormY
    End If


   On Error GoTo 0
   Exit Sub

licenceForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure licenceForm_MouseMove of Class Module cfAlpha"
End Sub


