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

Private WithEvents acceptButton As cwButton
Attribute acceptButton.VB_VarHelpID = -1
Private WithEvents declineButton As cwButton
Attribute declineButton.VB_VarHelpID = -1
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
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndShowStandardForms_Error
    
    aboutFormWidth = 1600
    aboutFormHeight = 1200
    aboutWidgetWidth = 1700
    aboutWidgetHeight = 1500
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", True, aboutWidgetWidth, aboutWidgetHeight)
       
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
    
'    aboutLeft = (aboutFormWidth / 2) - (aboutWidgetWidth / 2)
'    aboutTop = (aboutFormHeight / 2) - (aboutWidgetHeight / 2)
    
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 100, 100, aboutWidgetWidth, aboutWidgetHeight)
    
    
    ' -oOo-
    
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1400, 1200)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
         
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 0, 0, 1400, 1150)

    ' -oOo-
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set licenceForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, True, 580, 490)
    
    ' make the licence form transparent
    licenceForm.WidgetRoot.BackColor = -1
         
    ' create a new licence class widget with a name, position and size
    Set licenceWidget = licenceForm.Widgets.Add(New cwLicence, "licence", 0, 0, 550, 450)
    
    ' add two RC predefined button controls to the licence form
    Set declineButton = licenceForm.Widgets.Add(New_W("cwButton"), "declineButton", 352, 397, 70, 29)
        declineButton.Caption = ""
        declineButton.Widget.Alpha = 0.05
    
    Set acceptButton = licenceForm.Widgets.Add(New_W("cwButton"), "acceptButton", 428, 397, 70, 29)
        acceptButton.Caption = ""
        acceptButton.Widget.Alpha = 0.05
        
    Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 50, 60, 440, 280)
        textBox.MultiLine = True
        textBox.ScrollBars = vbVertical
        textBox.Widget.BackColor = &H4000&
        textBox.Widget.FontSize = 11
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
' Procedure : acceptButton_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub acceptButton_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    On Error GoTo acceptButton_Click_Error

    answer = msgBoxA("User accepted licence, thankyou.", vbOKOnly, "Licence Acceptance", False)
    
    'MsgBox "User accepted licence, thankyou."
    
    licenceWidget.ShowMe = False
    licenceWidget.opacity = 0
    licenceWidget.Widget.Refresh

    licenceForm.Visible = False
    
   On Error GoTo 0
   Exit Sub

acceptButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure acceptButton_Click of Class Module cfMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : declineButton_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub declineButton_Click()
    
    Dim answer As VbMsgBoxResult: answer = vbNo
        
    On Error GoTo declineButton_Click_Error

    answer = MsgBox("Licence declined, please destroy your copy of this widget and its documentation.", vbYes)
    If answer = 1 Then
        Call thisForm_Unload
    End If

   On Error GoTo 0
   Exit Sub

declineButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure declineButton_Click of Class Module cfMain"
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
