VERSION 5.00
Begin VB.Form menuForm 
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu mnuMainMenu 
      Caption         =   "mainmenu"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Panzer StopWatch Cairo widget"
      End
      Begin VB.Menu mnuBlank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramPreferences 
         Caption         =   "Widget Preferences"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi"
         Index           =   2
      End
      Begin VB.Menu blank7 
         Caption         =   ""
      End
      Begin VB.Menu mnuHelpSplash 
         Caption         =   "Panzer StopWatch Gauge Help"
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuWidgets 
            Caption         =   "See the other widgets"
         End
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version from Github"
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "Contact Support"
         End
         Begin VB.Menu mnuFacebook 
            Caption         =   "Chat about the widget on Facebook"
         End
         Begin VB.Menu mnuHelpHTM 
            Caption         =   "Open Help CHM"
         End
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu blank2 
         Caption         =   ""
      End
      Begin VB.Menu mnuAppFolder 
         Caption         =   "Reveal Widget in Windows Explorer"
      End
      Begin VB.Menu blank4 
         Caption         =   ""
      End
      Begin VB.Menu menuRestart 
         Caption         =   "Reload Widget (F5)"
      End
      Begin VB.Menu mnuEditWidget 
         Caption         =   "Edit Widget using..."
      End
      Begin VB.Menu mnuSwitchOff 
         Caption         =   "Switch off my functions (Ticking hands)"
      End
      Begin VB.Menu mnuTurnFunctionsOn 
         Caption         =   "Turn all functions ON"
      End
      Begin VB.Menu mnuseparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockWidget 
         Caption         =   "Lock Widget"
      End
      Begin VB.Menu mnuHideWidget 
         Caption         =   "Hide Widget"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Close Widget"
      End
   End
End
Attribute VB_Name = "menuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : The main dock won't take a menu when using GDI so we have a separate form for the menu
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Width = 1  ' the menu form is made as small as possible and moved off screen so that it does not show anywhere on the
    Me.Height = 1 ' screen, the menu appearing at the cursor point when it is told to do so by the dock form mousedown.
    'Me.ControlBox = False ' design time properties set in the IDE
    'Me.ShowInTaskbar = False ' set in the IDE ' NOTE: is possible in RC forms at runtime
    'Me.MaxButton = False ' set in the IDE
    'Me.MinButton = False ' set in the IDE
    Me.Visible = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : menuRestart_Click
' Author    : beededea
' Date      : 03/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub menuRestart_Click()

    On Error GoTo menuRestart_Click_Error
   
    Call reloadWidget

    On Error GoTo 0
    Exit Sub

menuRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuRestart_Click of Form menuForm"
End Sub

      

'---------------------------------------------------------------------------------------
' Procedure : mnuAppFolder_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAppFolder_Click()
    Dim folderPath As String: folderPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuAppFolder_Click_Error

    folderPath = App.Path
    If fDirExists(folderPath) Then ' if it is a folder already

        execStatus = ShellExecute(Me.hwnd, "open", folderPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        MsgBox "Having a bit of a problem opening a folder for this widget - " & folderPath & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuAppFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAppFolder_Click of Form menuForm"

End Sub








Private Sub mnublank1_Click()
    'Call mnuAbout_Click
End Sub

Private Sub mnuBlank2_Click()
    'Call mnuIconSettings_Click_Event
End Sub

Private Sub mnublnk_Click()
    'Call mnusettings_Click
End Sub





'---------------------------------------------------------------------------------------
' Procedure : mnuEditWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuEditWidget_Click()
    Dim editorPath As String: editorPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuEditWidget_Click_Error

    editorPath = PzGDefaultEditor
    If fFExists(editorPath) Then ' if it is a folder already
        '''If debugflg = 1  Then msgBox "ShellExecute " & sCommand
        
            ' run the selected program
        execStatus = ShellExecute(Me.hwnd, "open", editorPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open the IDE for this widget failed."
    Else
        MsgBox "Having a bit of a problem opening an IDE for this widgt - " & editorPath & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuEditWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuEditWidget_Click of Form menuForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuHelpHTM_Click
' Author    : beededea
' Date      : 14/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpHTM_Click()
    On Error GoTo mnuHelpHTM_Click_Error

        If fFExists(App.Path & "\help\Help.chm") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\Help.chm", vbNullString, App.Path, 1)
        Else
            MsgBox ("The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

mnuHelpHTM_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpHTM_Click of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuHelpPdf_Click
' Author    : beededea
' Date      : 14/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpSplash_Click()


    Dim fileToPlay As String: fileToPlay = vbNullString
    On Error GoTo mnuHelpPdf_Click_Error

    fileToPlay = "till.wav"
    If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    helpForm.Top = (screenHeightTwips / 2) - (helpForm.Height / 2)
    helpForm.Left = (screenWidthTwips / 2) - (helpForm.Width / 2)
     
    helpForm.show
    
     If (helpForm.WindowState = 1) Then
         helpForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpPdf_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHideWidget_Click
' Author    : beededea
' Date      : 14/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHideWidget_Click()
    On Error GoTo mnuHideWidget_Click_Error
       
    'overlayWidget.Hidden = True
    fAlpha.gaugeForm.Visible = False
    frmTimer.revealWidgetTimer.Enabled = True
    PzGWidgetHidden = "1"
    ' we have to save the value here
    sPutINISetting "Software\PzStopwatch", "widgetHidden", PzGWidgetHidden, PzGSettingsFile

   On Error GoTo 0
   Exit Sub

mnuHideWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHideWidget_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLockWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLockWidget_Click()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo mnuLockWidget_Click_Error

    fileToPlay = "lock.wav"
    If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If PzGPreventDragging = "1" Then
        mnuLockWidget.Checked = False
        PzGPreventDragging = "0"
        overlayWidget.Locked = False
    Else
        mnuLockWidget.Checked = True
        overlayWidget.Locked = 1
        PzGPreventDragging = "1"
    End If

    sPutINISetting "Software\PzStopwatch", "preventDragging", PzGPreventDragging, PzGSettingsFile

   On Error GoTo 0
   Exit Sub

mnuLockWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLockWidget_Click_Error of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuProgramPreferences_Click
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuProgramPreferences_Click()
    
    On Error GoTo mnuProgramPreferences_Click_Error

    Call makeProgramPreferencesAvailable

    On Error GoTo 0
    Exit Sub

mnuProgramPreferences_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuProgramPreferences_Click of Form menuForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuQuit_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuQuit_Click()

    On Error GoTo mnuQuit_Click_Error
    
    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

mnuQuit_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuQuit_Click of Form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(Index As Integer)

    Call mnuCoffee_ClickEvent

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of form menuForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuFacebook_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuFacebook_Click()
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuFacebook_Click_Error
    '''If debugflg = 1  Then msgBox "%" & "mnuFacebook_Click"

    answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFacebook_Click of form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuLatest_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuLatest_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo

    On Error GoTo mnuLatest_Click_Error
    '''If debugflg = 1  Then msgBox "%" & "mnuLatest_Click"
    
    'MsgBox "The download menu option is not yet enabled."

    answer = MsgBox("Download latest version of the program from github - this button opens a browser window and connects to the widget download page where you can check and download the latest SETUP.EXE file). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://github.com/yereverluvinunclebert/Panzer-StopWatch-Clock-VB6", vbNullString, App.Path, 1)
    End If


    On Error GoTo 0
    Exit Sub

mnuLatest_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLatest_Click of form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicence_Click()
    On Error GoTo mnuLicence_Click_Error
    '''If debugflg = 1  Then msgBox "%" & "mnuLicence_Click"
        
    Call mnuLicence_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuLicence_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_Click of form menuForm"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()

    On Error GoTo mnuSupport_Click_Error
    
    Call mnuSupport_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSweets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSweets_Click()
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuSweets_Click_Error
       '''If debugflg = 1  Then msgBox "%" & "mnuSweets_Click"
    
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a Kofi! This button opens a browser window and connects to Kofi donation page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.ko-fi.com/yereverluvinunclebert", vbNullString, App.Path, 1)
    End If
    
    On Error GoTo 0
    Exit Sub

mnuSweets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSweets_Click of form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSwitchOff_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSwitchOff_Click()
   On Error GoTo mnuSwitchOff_Click_Error

    overlayWidget.Ticking = False
    mnuSwitchOff.Checked = True
    mnuTurnFunctionsOn.Checked = False
    
    PzGGaugeFunctions = "0"
    sPutINISetting "Software\PzStopwatch", "gaugeFunctions", PzGGaugeFunctions, PzGSettingsFile

   On Error GoTo 0
   Exit Sub

mnuSwitchOff_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSwitchOff_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTurnFunctionsOn_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTurnFunctionsOn_Click()
    Dim fileToPlay As String: fileToPlay = vbNullString
    On Error GoTo mnuTurnFunctionsOn_Click_Error

    fileToPlay = "ting.wav"
    If PzGEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    overlayWidget.Ticking = True
    mnuSwitchOff.Checked = False
    mnuTurnFunctionsOn.Checked = True
    
    PzGGaugeFunctions = "1"
    sPutINISetting "Software\PzStopwatch", "gaugeFunctions", PzGGaugeFunctions, PzGSettingsFile

   On Error GoTo 0
   Exit Sub

mnuTurnFunctionsOn_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTurnFunctionsOn_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuWidgets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuWidgets_Click()
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuWidgets_Click_Error
    '''If debugflg = 1  Then msgBox "%" & "mnuWidgets_Click"

    answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuWidgets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuWidgets_Click of form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click()
    
    On Error GoTo mnuAbout_Click_Error
    
    Call aboutClickEvent

    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of form menuForm"
End Sub


