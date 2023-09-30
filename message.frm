VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "SteamyDock Enhanced Icon Settings Tool"
   ClientHeight    =   2100
   ClientLeft      =   4845
   ClientTop       =   4800
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMessage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   -30
      TabIndex        =   2
      Top             =   0
      Width           =   5970
      Begin VB.Frame fraPicVB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   195
         TabIndex        =   4
         Top             =   270
         Width           =   735
         Begin VB.Image picVBInformation 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":0000
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBCritical 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":11EA
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBExclamation 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":23D2
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBQuestion 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":360A
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         Height          =   195
         Left            =   1110
         TabIndex        =   3
         Top             =   570
         Width           =   4455
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton btnButtonTwo 
      Caption         =   "&No"
      Height          =   372
      Left            =   4980
      TabIndex        =   1
      Top             =   1620
      Width           =   972
   End
   Begin VB.CommandButton btnButtonOne 
      Caption         =   "&Yes"
      Height          =   372
      Left            =   3885
      TabIndex        =   0
      Top             =   1620
      Width           =   972
   End
   Begin VB.CheckBox chkShowAgain 
      Caption         =   "&Hide this message for the rest of this session"
      Height          =   420
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3435
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen STARTS
Option Explicit
Private mintLabelHeight As Integer
Private yesNoReturnValue As Integer
Private formMsgContext As String
Private formShowAgainChkBox As Boolean

'Private lastFormHeight As Long

Private msgBoxAdynamicSizingFlg As Boolean
Private Const cMsgBoxAFormHeight As Long = 2565
Private Const cMsgBoxAFormWidth  As Long = 11055

'---------------------------------------------------------------------------------------
' Property : btnButtonTwo_Click
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnButtonTwo_Click()
   On Error GoTo btnButtonTwo_Click_Error

    If formShowAgainChkBox = True Then SaveSetting App.EXEName, "Options", "Show message" & formMsgContext, chkShowAgain.Value
    yesNoReturnValue = 7
    Unload Me

   On Error GoTo 0
   Exit Sub

btnButtonTwo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property btnButtonTwo_Click of Form frmMessage"
End Sub

'---------------------------------------------------------------------------------------
' Property : btnButtonOne_Click
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnButtonOne_Click()
   On Error GoTo btnButtonOne_Click_Error

    Me.Visible = False
    If formShowAgainChkBox = True Then SaveSetting App.EXEName, "Options", "Show message" & formMsgContext, chkShowAgain.Value
    yesNoReturnValue = 6
    Unload Me

   On Error GoTo 0
   Exit Sub

btnButtonOne_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property btnButtonOne_Click of Form frmMessage"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Display
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Display()

    Dim intShow As Integer
    
    On Error GoTo Display_Error

    If formShowAgainChkBox = True Then
    
        chkShowAgain.Visible = True
        ' Returns a key setting value from an application's entry in the Windows registry
        intShow = GetSetting(App.EXEName, "Options", "Show message" & formMsgContext, vbUnchecked)
        
        If intShow = vbUnchecked Then
            Me.show vbModal
        End If
    Else
        Me.show vbModal
    End If

   On Error GoTo 0
   Exit Sub

Display_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Display of Form frmMessage"

End Sub
' property to allow a message to be passed to the form
'---------------------------------------------------------------------------------------
' Property  : propMessage
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propMessage(ByVal strMessage As String)

    Dim intDiff As Integer
    
    On Error GoTo propMessage_Error

    lblMessage.Caption = strMessage
    
    ' Expand the form and move the other controls if the message is too long to show.
      
    intDiff = lblMessage.Height - mintLabelHeight
    
    If PzGDpiAwareness = "1" Then
        Me.Height = 4000
    Else
        Me.Height = Me.Height + intDiff
    End If

    fraMessage.Height = fraMessage.Height + intDiff
    fraPicVB.Top = fraPicVB.Top + (intDiff / 2)
        
    chkShowAgain.Top = chkShowAgain.Top + intDiff
    btnButtonOne.Top = btnButtonOne.Top + intDiff
    btnButtonTwo.Top = btnButtonTwo.Top + intDiff

   On Error GoTo 0
   Exit Property

propMessage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propMessage of Form frmMessage"

End Property

'---------------------------------------------------------------------------------------
' Property  : propTitle
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propTitle(ByVal strTitle As String)
   On Error GoTo propTitle_Error

    If strTitle = "" Then
        frmMessage.Caption = "SteamyDock Icon Enhanced Settings"
    Else
        frmMessage.Caption = strTitle
    End If

   On Error GoTo 0
   Exit Property

propTitle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propTitle of Form frmMessage"
End Property

'---------------------------------------------------------------------------------------
' Property  : propMsgContext
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propMsgContext(ByVal thisContext As String)
   On Error GoTo propMsgContext_Error

    formMsgContext = thisContext

   On Error GoTo 0
   Exit Property

propMsgContext_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propMsgContext of Form frmMessage"
End Property

'---------------------------------------------------------------------------------------
' Property  : propShowAgainChkBox
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propShowAgainChkBox(ByVal showAgainVis As Boolean)
   On Error GoTo propShowAgainChkBox_Error

    formShowAgainChkBox = showAgainVis

   On Error GoTo 0
   Exit Property

propShowAgainChkBox_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propShowAgainChkBox of Form frmMessage"
End Property

'---------------------------------------------------------------------------------------
' Property  : propButtonVal
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propButtonVal(ByVal buttonVal As Integer)
    
    Dim fileToPlay As String: fileToPlay = vbNullString

   On Error GoTo propButtonVal_Error

    btnButtonOne.Visible = False
    btnButtonTwo.Visible = False
    'btnButtonThree.Visible = false

    picVBInformation.Visible = False
    picVBCritical.Visible = False
    picVBExclamation.Visible = False
    picVBQuestion.Visible = False

    btnButtonOne.Left = 3885

    If buttonVal >= 64 Then ' vbInformation
       buttonVal = buttonVal - 64
       picVBInformation.Visible = True
    ElseIf buttonVal >= 48 Then '    vbExclamation
        buttonVal = buttonVal - 48
        picVBExclamation.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "ting.wav"
        If fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    ElseIf buttonVal >= 32 Then '    vbQuestion
        buttonVal = buttonVal - 32
        picVBQuestion.Visible = True
    ElseIf buttonVal >= 20 Then '    vbCritical
        buttonVal = buttonVal - 20
        picVBCritical.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "belltoll01.wav"
        If fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If


    If buttonVal = 0 Then '    vbOKOnly 0
        picVBInformation.Visible = True
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = False
        btnButtonOne.Caption = "OK"
        btnButtonOne.Left = 4620
    End If
    If buttonVal = 1 Then '    vbOKCancel 1
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "OK"
        btnButtonTwo.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 2 Then 'vbAbortRetryIgnore 2
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Abort"
        btnButtonOne.Caption = "Retry"
        'btnButtonThree.Caption = "Ignore"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 3 Then '    vbYesNoCancel 3
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
        'btnButtonThree.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 4 Then '    vbYesNo 4
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 5 Then '    vbRetryCancel 5
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Retry"
        btnButtonTwo.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 6 Then '    vbYes 6
        'btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = ""
        btnButtonTwo.Caption = "Yes"
        picVBQuestion.Visible = True
    End If
    If buttonVal = 7 Then '    vbNo 7
        'btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = ""
        btnButtonTwo.Caption = "No"
        picVBQuestion.Visible = True
    End If

   On Error GoTo 0
   Exit Property

propButtonVal_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propButtonVal of Form frmMessage"
        
End Property

'---------------------------------------------------------------------------------------
' Procedure : propReturnedValue
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propReturnedValue()

   On Error GoTo propReturnedValue_Error

    propReturnedValue = yesNoReturnValue

   On Error GoTo 0
   Exit Property

propReturnedValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propReturnedValue of Form frmMessage"
    
End Property




'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim Ctrl As Control

    On Error GoTo Form_Load_Error

    mintLabelHeight = lblMessage.Height
    
    If PzGDpiAwareness = "1" Then
        msgBoxAdynamicSizingFlg = True
    End If
    
    msgBoxACurrentWidth = cMsgBoxAFormWidth
    msgBoxACurrentHeight = cMsgBoxAFormHeight
    
    If PzGDpiAwareness = "1" Then
        ' save the initial positions of ALL the controls on the msgbox form
        Call SaveSizes(Me, msgBoxAControlPositions(), msgBoxACurrentWidth, msgBoxACurrentHeight)
    End If
        
    ' .TBD DAEB 05/05/2021 frmMessage.frm Added the font mod. here instead of within the changeFont tool
    '                       as each instance of the form is new, the font modification must be here.
    For Each Ctrl In Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
            If PzGPrefsFont <> "" Then Ctrl.Font.Name = PzGPrefsFont
           
            If PzGDpiAwareness = "1" Then
                If Val(Abs(PzGPrefsFontSizeHighDPI)) > 0 Then Ctrl.Font.Size = Val(Abs(PzGPrefsFontSizeHighDPI))
            Else
                If Val(Abs(PzGPrefsFontSizeLowDPI)) > 0 Then Ctrl.Font.Size = Val(Abs(PzGPrefsFontSizeLowDPI))
            End If
            'Ctrl.Font.Italic = CBool(SDSuppliedFontItalics) TBD
           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
        End If
    Next

    chkShowAgain.Visible = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmMessage"
    
End Sub

'---------------------------------------------------------------------------------------
' Property : Form_Resize
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
    Dim currentFont As Integer: currentFont = 0
    Dim ratio As Double: ratio = 0
    
    On Error GoTo Form_Resize_Error
    
    If WindowState = vbMinimized Then Exit Sub

    ratio = cMsgBoxAFormHeight / cMsgBoxAFormWidth
    If PzGDpiAwareness = "1" Then
        currentFont = PzGPrefsFontSizeHighDPI
    Else
        currentFont = PzGPrefsFontSizeLowDPI
    End If
    
    If msgBoxAdynamicSizingFlg = True Then

        Call resizeControls(Me, msgBoxAControlPositions(), msgBoxACurrentWidth, msgBoxACurrentHeight, currentFont)
        
        Me.Width = Me.Height / ratio ' maintain the aspect ratio

        'Call loadHigherResImages
    Else
        If Me.WindowState = 0 Then
            If Me.Width > 9090 Then Me.Width = 9090
            If Me.Width < 6105 Then Me.Width = 6105
            'If lastFormHeight <> 0 Then Me.Height = lastFormHeight
        End If
    End If
    
    'Call resizeControls(Me, msgBoxAControlPositions(), msgBoxACurrentWidth, msgBoxACurrentHeight, currentFont)

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Form_Resize of Form frmMessage"
End Sub

' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen ENDS
Private Sub picVBInformation_Click()

End Sub

