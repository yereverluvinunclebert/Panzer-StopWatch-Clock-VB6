VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "SteamyDock Enhanced Icon Settings Tool"
   ClientHeight    =   2100
   ClientLeft      =   4845
   ClientTop       =   4800
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMessage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   -30
      TabIndex        =   2
      Top             =   0
      Width           =   6000
      Begin VB.Frame fraPicVB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   195
         TabIndex        =   4
         Top             =   75
         Width           =   735
         Begin VB.Image picVBInformation 
            Appearance      =   0  'Flat
            Height          =   720
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBCritical 
            Height          =   720
            Left            =   0
            Top             =   -15
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBExclamation 
            Appearance      =   0  'Flat
            Height          =   720
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBQuestion 
            Height          =   720
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Some message text and more text in addition"
         Height          =   195
         Left            =   1110
         TabIndex        =   3
         Top             =   360
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
      Caption         =   "&Hide this message."
      Height          =   420
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMessage
' Author    : beededea
' Date      : 20/11/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen STARTS
Option Explicit
Private yesNoReturnValue As Integer
Private formMsgContext As String
Private formShowAgainChkBox As Boolean

'Private lastFormHeight As Long

Private Const cMsgBoxAFormHeight As Long = 2565
Private Const cMsgBoxAFormWidth  As Long = 11055

Private mPropMessage As String
Private mPropTitle As String
Private mPropMsgContext As String
Private mPropShowAgainChkBox As Boolean
Private mPropButtonVal As Integer
Private mPropReturnedValue As Integer



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
    
    msgBoxACurrentWidth = cMsgBoxAFormWidth
    msgBoxACurrentHeight = cMsgBoxAFormHeight
    
    'If PzGDpiAwareness = "1" Then
        ' save the initial positions of ALL the controls on the msgbox form
        Call SaveSizes(Me, msgBoxAControlPositions(), msgBoxACurrentWidth, msgBoxACurrentHeight)
    'End If
        
    ' .TBD DAEB 05/05/2021 frmMessage.frm Added the font mod. here instead of within the changeFont tool
    '                       as each instance of the form is new, the font modification must be here.
    For Each Ctrl In Me.Controls
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
    Dim currentFont As Long: currentFont = 0
    Dim ratio As Double: ratio = 0
    
    On Error GoTo Form_Resize_Error
    
    If Me.WindowState = vbMinimized Then Exit Sub

    ratio = cMsgBoxAFormHeight / cMsgBoxAFormWidth
    If PzGDpiAwareness = "1" Then
        currentFont = Val(PzGPrefsFontSizeHighDPI)
    Else
        currentFont = Val(PzGPrefsFontSizeLowDPI)
    End If
    
    If msgBoxADynamicSizingFlg = True Then
        Call setMessageIconImagesLight(1920)
        Call resizeControls(Me, msgBoxAControlPositions(), msgBoxACurrentWidth, msgBoxACurrentHeight, currentFont)
        Me.Width = Me.Height / ratio ' maintain the aspect ratio
    Else
        Call setMessageIconImagesLight(600)
    End If

''    If panzerPrefs.mnuDark.Checked = True Then
''        Call setMessageIconImagesDark(determineIconWidth(Me, msgBoxADynamicSizingFlg))
''    Else
'        Call setMessageIconImagesLight(1920)
''    End If
    
   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Form_Resize of Form frmMessage"
End Sub

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
    Me.Hide

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
    Me.Hide

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
            Me.Show vbModal
        End If
    Else
        Me.Show vbModal
    End If

   On Error GoTo 0
   Exit Sub

Display_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Display of Form frmMessage"

End Sub




'
'---------------------------------------------------------------------------------------
' Property  : propMessage
' Author    : beededea
' Date      : 23/09/2023
' Purpose   : property to allow a message to be passed to the form
'---------------------------------------------------------------------------------------
'
Public Property Let propMessage(ByVal newValue As String)
    
    On Error GoTo propMessage_Error
    
    If mPropMessage <> newValue Then mPropMessage = newValue Else Exit Property

    lblMessage.Caption = mPropMessage
    
    ' Expand the form and move the other controls if the message is too long to show.
          
    If msgBoxADynamicSizingFlg = True Then
        ' this causes a resize event
        Me.Height = (screenHeightTwips / 5.5) '+ intDiff
    Else
        fraPicVB.Top = 285
    End If
   
   On Error GoTo 0
   Exit Property

propMessage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propMessage of Form frmMessage"

End Property

'---------------------------------------------------------------------------------------
' Procedure : propMessage
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propMessage() As String
   On Error GoTo propMessageGet_Error

   propMessage = mPropMessage

   On Error GoTo 0
   Exit Property

propMessageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propMessage of Class Module cwhelp"
End Property

'---------------------------------------------------------------------------------------
' Property  : propTitle
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propTitle(ByVal newValue As String)
   On Error GoTo propTitle_Error
   
    If mPropTitle <> newValue Then mPropTitle = newValue Else Exit Property

    If mPropTitle = "" Then
        Me.Caption = "Panzer Clock Message."
    Else
        Me.Caption = mPropTitle
    End If

   On Error GoTo 0
   Exit Property

propTitle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propTitle of Form frmMessage"
End Property
'---------------------------------------------------------------------------------------
' Procedure : propTitle
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propTitle() As String
   On Error GoTo propTitleGet_Error

   propTitle = mPropTitle

   On Error GoTo 0
   Exit Property

propTitleGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propTitle of Class Module cwhelp"
End Property

'---------------------------------------------------------------------------------------
' Property  : propMsgContext
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propMsgContext(ByVal newValue As String)
   On Error GoTo propMsgContext_Error
   
   If mPropMsgContext <> newValue Then mPropMsgContext = newValue Else Exit Property

   formMsgContext = mPropMsgContext

   On Error GoTo 0
   Exit Property

propMsgContext_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propMsgContext of Form frmMessage"
End Property
'---------------------------------------------------------------------------------------
' Procedure : propMsgContext
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propMsgContext() As String
   On Error GoTo propMsgContextGet_Error

   propMsgContext = mPropMsgContext

   On Error GoTo 0
   Exit Property

propMsgContextGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propMsgContext of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Procedure : propReturnedValue
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propReturnedValue() As Integer
   On Error GoTo propReturnedValue_Error
   
    propReturnedValue = yesNoReturnValue

   On Error GoTo 0
   Exit Property

propReturnedValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propReturnedValue of Form frmMessage"
    
End Property

'---------------------------------------------------------------------------------------
' Property  : propReturnedValue
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propReturnedValue(ByVal newValue As Integer)
   On Error GoTo propReturnedValue_Error
   
    If mPropReturnedValue <> newValue Then mPropReturnedValue = newValue Else Exit Property

    formShowAgainChkBox = mPropReturnedValue

   On Error GoTo 0
   Exit Property

propReturnedValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propReturnedValue of Form frmMessage"
End Property

'---------------------------------------------------------------------------------------
' Property  : propShowAgainChkBox
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propShowAgainChkBox(ByVal newValue As Boolean)
   On Error GoTo propShowAgainChkBox_Error
   
    If mPropShowAgainChkBox <> newValue Then mPropShowAgainChkBox = newValue Else Exit Property

    formShowAgainChkBox = mPropShowAgainChkBox

   On Error GoTo 0
   Exit Property

propShowAgainChkBox_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propShowAgainChkBox of Form frmMessage"
End Property
'---------------------------------------------------------------------------------------
' Procedure : propShowAgainChkBox
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get propShowAgainChkBox() As Boolean
   On Error GoTo propShowAgainChkBoxGet_Error

   propShowAgainChkBox = mPropShowAgainChkBox

   On Error GoTo 0
   Exit Property

propShowAgainChkBoxGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure propShowAgainChkBox of Class Module cwhelp"
End Property
'---------------------------------------------------------------------------------------
' Property  : propButtonVal
' Author    : beededea
' Date      : 23/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let propButtonVal(ByVal newValue As Integer)
    
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo propButtonVal_Error

    If mPropButtonVal <> newValue Then mPropButtonVal = newValue Else Exit Property
    
    btnButtonOne.Visible = False
    btnButtonTwo.Visible = False

    picVBInformation.Visible = False
    picVBCritical.Visible = False
    picVBExclamation.Visible = False
    picVBQuestion.Visible = False
    
    If mPropButtonVal = 0 Then ' vbInformation
       picVBInformation.Visible = True
    ElseIf mPropButtonVal >= 64 Then ' vbInformation
       mPropButtonVal = mPropButtonVal - 64
       picVBInformation.Visible = True
    ElseIf mPropButtonVal >= 48 Then '    vbExclamation
        mPropButtonVal = mPropButtonVal - 48
        picVBExclamation.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "ting.wav"
        If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    ElseIf mPropButtonVal >= 32 Then '    vbQuestion
        mPropButtonVal = mPropButtonVal - 32
        picVBQuestion.Visible = True
    ElseIf mPropButtonVal >= 20 Then '    vbCritical
        mPropButtonVal = mPropButtonVal - 20
        picVBCritical.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "belltoll01.wav"
        If fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If

    If mPropButtonVal = 0 Then '    vbOKOnly 0
        picVBInformation.Visible = True
        btnButtonOne.Visible = False
        btnButtonTwo.Visible = True
        btnButtonTwo.Caption = "OK"
        'btnButtonOne.Left = 4620
    End If
    If mPropButtonVal = 1 Then '    vbOKCancel 1
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "OK"
        btnButtonTwo.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If
    If mPropButtonVal = 2 Then 'vbAbortRetryIgnore 2
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Abort"
        btnButtonOne.Caption = "Retry"
        'btnButtonThree.Caption = "Ignore"
        picVBQuestion.Visible = True
    End If
    If mPropButtonVal = 3 Then '    vbYesNoCancel 3
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
        'btnButtonThree.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If
    If mPropButtonVal = 4 Then '    vbYesNo 4
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
        picVBQuestion.Visible = True
    End If
    If mPropButtonVal = 5 Then '    vbRetryCancel 5
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Retry"
        btnButtonTwo.Caption = "Cancel"
        picVBQuestion.Visible = True
    End If

   On Error GoTo 0
   Exit Property

propButtonVal_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property propButtonVal of Form frmMessage"
        
End Property


'---------------------------------------------------------------------------------------
' Procedure : loadHigherResMessageImages
' Author    : beededea
' Date      : 18/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResMessageImages()

    On Error GoTo loadHigherResMessageImages_Error

    If Me.WindowState = vbMinimized Then Exit Sub

'    If panzerPrefs.mnuDark.Checked = True Then
'        Call setMessageIconImagesDark(determineIconWidth(Me, msgBoxADynamicSizingFlg))
'    Else
        Call setMessageIconImagesLight(1920)
'    End If

   On Error GoTo 0
   Exit Sub

loadHigherResMessageImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResMessageImages of Form panzerPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesLight
' Author    : beededea
' Date      : 22/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setMessageIconImagesLight(ByVal thisIconWidth As Long)
    
    Dim resourcePath As String: resourcePath = vbNullString
    
    On Error GoTo setMessageIconImagesLight_Error
    
    resourcePath = App.path & "\resources\images"
    
    If fFExists(resourcePath & "\windowsInformation" & thisIconWidth & ".jpg") Then Set picVBInformation.Picture = LoadPicture(resourcePath & "\windowsInformation" & thisIconWidth & ".jpg")
    If fFExists(resourcePath & "\windowsOrangeExclamation" & thisIconWidth & ".jpg") Then Set picVBExclamation.Picture = LoadPicture(resourcePath & "\windowsOrangeExclamation" & thisIconWidth & ".jpg")
    If fFExists(resourcePath & "\windowsShieldQMark" & thisIconWidth & ".jpg") Then Set picVBQuestion.Picture = LoadPicture(resourcePath & "\windowsShieldQMark" & thisIconWidth & ".jpg")
    If fFExists(resourcePath & "\windowsCritical" & thisIconWidth & ".jpg") Then Set picVBCritical.Picture = LoadPicture(resourcePath & "\windowsCritical" & thisIconWidth & ".jpg")
    
    picVBInformation.Refresh
    picVBQuestion.Refresh
    picVBExclamation.Refresh
    picVBCritical.Refresh
    
   On Error GoTo 0
   Exit Sub

setMessageIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMessageIconImagesLight of Form frmMessage"

End Sub
