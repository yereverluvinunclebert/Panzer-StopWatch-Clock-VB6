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
    aboutWidgetWidth = 745
    aboutWidgetHeight = 622
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", True, 1600, 1200)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
    
    aboutLeft = (aboutFormWidth / 2) - (aboutWidgetWidth / 2)
    aboutTop = (aboutFormHeight / 2) - (aboutWidgetHeight / 2)
    
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", aboutLeft, aboutTop, aboutWidgetWidth, aboutWidgetHeight)
    
    
    ' -oOo-
    
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1400, 1200)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
         
    ' create a new globe widget with a name and location, first step of that is to run Class_Initialize
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 0, 0, 1400, 1150)


   On Error GoTo 0
   Exit Sub

initAndShowStandardForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndShowStandardForms of Class Module cfMain"
    
End Sub

