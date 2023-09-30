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
Public WithEvents licenceForm As cWidgetForm
Attribute licenceForm.VB_VarHelpID = -1

Public licenceFormX As Long
Public licenceFormY As Long
Public licenceFormZ As Single

Private WithEvents btnAccept As cwButton
Attribute btnAccept.VB_VarHelpID = -1
Private WithEvents btnDecline As cwButton
Attribute btnDecline.VB_VarHelpID = -1
Private WithEvents textBox As cwTextBox
Attribute textBox.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : initAndShowStandardForms
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initAndShowStandardForms(ByVal x As Long, ByVal Y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim aboutLeft As Long: aboutLeft = 0
    Dim aboutTop As Long: aboutTop = 0
    Dim aboutFormWidth As Long: aboutFormWidth = 0
    Dim aboutFormHeight As Long: aboutFormHeight = 0
    Dim aboutWidgetWidth As Long: aboutWidgetWidth = 0
    Dim aboutWidgetHeight As Long: aboutWidgetHeight = 0
        
    Dim helpFormWidth As Long: helpFormWidth = 0
    Dim helpFormHeight As Long: helpFormHeight = 0
    Dim helpWidgetWidth As Long: helpWidgetWidth = 0
    Dim helpWidgetHeight As Long: helpWidgetHeight = 0
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowStandardForms_Error
    
    If PzGDpiAwareness = "0" Then
        aboutFormWidth = 800
        aboutFormHeight = 600
        aboutWidgetWidth = 850
        aboutWidgetHeight = 750
    Else
        aboutFormWidth = 1600
        aboutFormHeight = 1200
        aboutWidgetWidth = 1700
        aboutWidgetHeight = 1500
    End If
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", True, aboutWidgetWidth, aboutWidgetHeight)
       
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
    
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 100, 100, aboutWidgetWidth, aboutWidgetHeight)
    
    ' -oOo-
    
    If PzGDpiAwareness = "0" Then
        helpFormWidth = 800
        helpFormHeight = 700
        helpWidgetWidth = 850
        helpWidgetHeight = 750
    Else
        helpFormWidth = 1400
        helpFormHeight = 1200
        helpWidgetWidth = 1400
        helpWidgetHeight = 1150
    End If
          
    ' create an invisible form using a Cairo Widget Form with a predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , helpFormWidth, helpFormHeight)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
         
    ' create a new help widget image with a name, position and size
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 0, 0, helpWidgetWidth, helpWidgetHeight)

    ' -oOo-
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set licenceForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, True, 580, 490)
    
    ' make the licence form transparent
    licenceForm.WidgetRoot.BackColor = -1
         
    ' create a new licence class widget with a name, position and size
    Set licenceWidget = licenceForm.Widgets.Add(New cwLicence, "licence", 0, 0, 550, 450)
    
    ' add two RC predefined button controls to the licence form
    If PzGDpiAwareness = "1" Then
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 352, 397, 70, 29)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 428, 397, 70, 29)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 50, 60, 440, 280)
        textBox.Widget.FontSize = 11
    Else
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 189, 210, 40, 18)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 228, 210, 40, 18)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 25, 30, 240, 155)
        textBox.Widget.FontSize = 8
    End If
    
    btnDecline.Caption = ""
    btnDecline.Widget.Alpha = 0.05

    btnAccept.Caption = ""
    btnAccept.Widget.Alpha = 0.05
    
    textBox.MultiLine = True
    textBox.ScrollBars = vbVertical
    textBox.Widget.BackColor = &H4000&
    
    textBox.Widget.ForeColor = vbWhite
    textBox.Text = widgetName & vbCrLf & vbCrLf & _
    "Copyright © 1992-2023 Dean Beedell" & vbCrLf & vbCrLf & _
    "This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public Licence as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version." & vbCrLf & vbCrLf & _
        "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." & vbCrLf & vbCrLf & _
        "You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA" & vbCrLf & vbCrLf & _
        "If you use this software in any way whatsoever then that implies acceptance of the licence. If you do not wish to comply with the licence terms then please remove the download, binary and source code from your systems immediately." & vbCrLf & vbCrLf & _
        "You may use any of my own imagery in your own creations but commercially only with my permission. In all other non-commercial cases I require a credit to the original artist using my name or one of my pseudonyms and a link to my site. With regard to the commercial use of incorporated images, permission and a licence would need to be obtained from the original owner and creator, ie. me."

   On Error GoTo 0
   Exit Sub

initAndShowStandardForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowStandardForms of Class Module cfMain"
    
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
    On Error GoTo btnAccept_Click_Error

    answer = msgBoxA("User accepted licence, thankyou.", vbOKOnly, "Licence Acceptance", False)
    
    'MsgBox "User accepted licence, thankyou."
    
    licenceWidget.ShowMe = False
    licenceWidget.opacity = 0
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
    Dim answerMsg As String: answerMsg = vbNullString
        
    On Error GoTo btnDecline_Click_Error

    'answer = MsgBox("Licence declined, please destroy your copy of this widget and its documentation.", vbYes)
    answer = msgBoxA("Licence declined, please destroy your copy of this widget and its documentation. This widget will now unload.", vbOKOnly, "Licence Declined", False)

    If answer = 6 Then ' OK
        Call thisForm_Unload
    End If

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
Private Sub licenceForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo licenceForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If PzGIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not licenceForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      licenceFormX = (x - x0) * licenceForm.WidgetRoot.Zoom + licenceForm.Left
      licenceFormY = (Y - y0) * licenceForm.WidgetRoot.Zoom + licenceForm.Top
      licenceForm.Move licenceFormX, licenceFormY
    End If


   On Error GoTo 0
   Exit Sub

licenceForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure licenceForm_MouseMove of Class Module cfAlpha"
End Sub
