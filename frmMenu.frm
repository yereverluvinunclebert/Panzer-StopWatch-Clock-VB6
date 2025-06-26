VERSION 5.00
Begin VB.Form menuForm 
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   4290
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
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
         Caption         =   "About Steampunk Clock Calendar Cairo widget ..."
      End
      Begin VB.Menu mnuBlank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramPreferences 
         Caption         =   "Widget Preferences ..."
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi ..."
         Index           =   2
      End
      Begin VB.Menu blank7 
         Caption         =   ""
      End
      Begin VB.Menu mnuHelpSplash 
         Caption         =   "Steampunk Clock Calendar Help ..."
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options ..."
         Begin VB.Menu mnuWidgets 
            Caption         =   "See the other widgets ..."
         End
         Begin VB.Menu mnuGithubHome 
            Caption         =   "Github Home for this widget ..."
         End
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version from Github ..."
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "Contact Support ..."
         End
         Begin VB.Menu mnuFacebook 
            Caption         =   "Chat about the widget on Facebook ..."
         End
         Begin VB.Menu mnuHelpHTM 
            Caption         =   "Open Help CHM ..."
         End
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement ..."
      End
      Begin VB.Menu blank2 
         Caption         =   ""
      End
      Begin VB.Menu mnuAppFolder 
         Caption         =   "Reveal Widget in Windows Explorer ..."
      End
      Begin VB.Menu blank4 
         Caption         =   ""
      End
      Begin VB.Menu menuReload 
         Caption         =   "Reload Widget (F5 or Shift+F5 for hard restart)"
      End
      Begin VB.Menu mnuEditWidget 
         Caption         =   "Edit Widget using..."
      End
      Begin VB.Menu mnuSwitchOff 
         Caption         =   "Switch off my functions (all animations, sounds, alarms, ticking &&c)"
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
      Begin VB.Menu blank12 
         Caption         =   ""
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
'---------------------------------------------------------------------------------------
' Module    : menuForm
' Author    : beededea
' Date      : 26/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule ModuleWithoutFolder

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : blank12_Click
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : Accidental clicking on a blank menu item next to the close widget menu item
'---------------------------------------------------------------------------------------
'
Private Sub blank12_Click()
   On Error GoTo blank12_Click_Error

    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

blank12_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure blank12_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : We have a separate form for the right click menu. We do not need an on-form menu for the
'               various RC6 forms so a native VB6 menu will do. It looks good in any case as it is
'               merely replicating the Yahoo widget menu.
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
' Procedure : menuReload_Click
' Author    : beededea
' Date      : 03/05/2023
' Purpose   : reload via the menu
'---------------------------------------------------------------------------------------
'
Private Sub menuReload_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo menuReload_Click_Error
    
    answer = vbYes
    answerMsg = "Performing a hard restart now, press OK."
        
    If gblSHIFT_1 = True Then
        gblSHIFT_1 = False
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Performing a hard restart", True, "menuReloadClick")
        
        Call hardRestart
    Else
        Call reloadProgram
    End If
    
    On Error GoTo 0
    Exit Sub

menuReload_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuReload_Click of Form menuForm"
End Sub

      



'---------------------------------------------------------------------------------------
' Procedure : mnuAppFolder_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : menu item to reveal the widget's app folder
'---------------------------------------------------------------------------------------
'
Private Sub mnuAppFolder_Click()
    Dim folderPath As String: folderPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuAppFolder_Click_Error

    folderPath = App.path
    If fDirExists(folderPath) Then ' if it is a folder already

        execStatus = ShellExecute(Me.hWnd, "open", folderPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        MsgBox "Having a bit of a problem opening a folder for this widget - " & folderPath & " It doesn't seem to have a valid working directory set.", "UBoat StopWatch Confirmation Message", vbOKOnly + vbExclamation
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "UBoat StopWatch Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuAppFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAppFolder_Click of Form menuForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuEditWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : menu item to open the VB6/TwinBasic IDE depending upon the editing environment, either the VBP or .TWINPROJ files
'---------------------------------------------------------------------------------------
'
Private Sub mnuEditWidget_Click()
    Dim editorPath As String: editorPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuEditWidget_Click_Error

    #If TWINBASIC Then
        editorPath = gblDefaultTBEditor
    #Else
        editorPath = gblDefaultVB6Editor
    #End If
    
    If fFExists(editorPath) Then ' if it is a folder already
        '''If gblDebugFlg = 1  Then msgBox "ShellExecute " & sCommand
        
        ' run the selected program
        execStatus = ShellExecute(Me.hWnd, "open", editorPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open the IDE for this widget failed."
    Else
        MsgBox "Having a bit of a problem opening an IDE for this widget - " & editorPath & " It doesn't seem to have a valid working directory set."
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "UBoat StopWatch Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuEditWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuEditWidget_Click of Form menuForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuGithubHome_Click
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : menu item to open the github HTML page
'---------------------------------------------------------------------------------------
'
Public Sub mnuGithubHome_Click()

    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo mnuGithubHome_Click_Error
    
    answer = vbYes

    answerMsg = "This button opens a browser window and connects to the widget's HOME page on github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Upgrade", True, "mnuGithubHomeClick")

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://github.com/yereverluvinunclebert/UBoat-StopWatch-" & gblCodingEnvironment, vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuGithubHome_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuGithubHome_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelpHTM_Click
' Author    : beededea
' Date      : 14/05/2023
' Purpose   : menu item to open the help compiled HTML file
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpHTM_Click()
    On Error GoTo mnuHelpHTM_Click_Error

        If fFExists(App.path & "\help\Help.chm") Then
            Call ShellExecute(Me.hWnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
        Else
            MsgBox ("The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

mnuHelpHTM_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpHTM_Click of Form menuForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : mnuHelpSplash_Click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : menu item to open display the splash bitmap
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpSplash_Click()
   On Error GoTo mnuHelpSplash_Click_Error

    Call helpSplash

   On Error GoTo 0
   Exit Sub

mnuHelpSplash_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpSplash_Click of Form menuForm"

End Sub






'---------------------------------------------------------------------------------------
' Procedure : mnuHideWidget_Click
' Author    : beededea
' Date      : 14/05/2023
' Purpose   : menu item to hide the main gauge form
'---------------------------------------------------------------------------------------
'
Private Sub mnuHideWidget_Click()
    On Error GoTo mnuHideWidget_Click_Error
       
    'overlayWidget.Hidden = True
    fGauge.gaugeForm.Visible = False
    frmTimer.revealWidgetTimer.Enabled = True
    gblWidgetHidden = "1"
    ' we have to save the value here
    sPutINISetting "Software\UBoatStopWatch", "widgetHidden", gblWidgetHidden, gblSettingsFile

   On Error GoTo 0
   Exit Sub

mnuHideWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHideWidget_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLockWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : menu item to lock the gaugeForm in place
'---------------------------------------------------------------------------------------
'
Private Sub mnuLockWidget_Click()

    On Error GoTo mnuLockWidget_Click_Error
    
    Call toggleWidgetLock

   On Error GoTo 0
   Exit Sub

mnuLockWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLockWidget_Click_Error of Form menuForm"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuProgramPreferences_Click
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : menu item to lock the gaugeForm in place on the desktop
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
' Purpose   : menu item to quit the program altogether
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
    On Error GoTo mnuCoffee_Click_Error
    
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
' Purpose   : menu item to open the FB HTML page
'---------------------------------------------------------------------------------------
'
Public Sub mnuFacebook_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo mnuFacebook_Click_Error
    
    answer = vbYes

    answerMsg = "Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Visit Facebook Request", True, "mnuFacebookClick")
    'answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.path, 1)
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
' Purpose   : menu item to open the Github HTML page for the widget release page
'---------------------------------------------------------------------------------------
'
Public Sub mnuLatest_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo mnuLatest_Click_Error
    '''If gblDebugFlg = 1  Then msgBox "%" & "mnuLatest_Click"
    
    answer = vbYes

    answerMsg = "Download latest version of the program from github - this button opens a browser window and connects to the widget releases page where you can download the latest installer. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Upgrade", True, "mnuLatestClick")

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://github.com/yereverluvinunclebert/UBoat-StopWatch-" & gblCodingEnvironment & "/releases", vbNullString, App.path, 1)
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
' Purpose   : menu item to open the licence form
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicence_Click()
    On Error GoTo mnuLicence_Click_Error
        
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
' Purpose   : menu item to open the issues page on github
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

''---------------------------------------------------------------------------------------
'' Procedure : mnuSweets_Click
'' Author    : beededea
'' Date      : 13/02/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub mnuSweets_Click()
'    Dim answer As VbMsgBoxResult: answer = vbNo
'    Dim answerMsg As String: answerMsg = vbNullString
'
'    On Error GoTo mnuSweets_Click_Error
'    answer = vbYes
'    answerMsg = " Help support the creation of more widgets like this. Buy me a Kofi! This button opens a browser window and connects to Kofi donation page). Will you be kind and proceed?"
'    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Donate a Kofi", True, "mnuSweetsClick")
'    'answer = MsgBox(" Help support the creation of more widgets like this. Buy me a Kofi! This button opens a browser window and connects to Kofi donation page). Will you be kind and proceed?", vbExclamation + vbYesNo)
'
'    If answer = vbYes Then
'        Call ShellExecute(Me.hwnd, "Open", "https://www.ko-fi.com/yereverluvinunclebert", vbNullString, App.path, 1)
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'mnuSweets_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSweets_Click of form menuForm"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSwitchOff_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : menu item to switch off all of the program's functions
'---------------------------------------------------------------------------------------
'
Private Sub mnuSwitchOff_Click()
   On Error GoTo mnuSwitchOff_Click_Error

    Call SwitchOff
    
   On Error GoTo 0
   Exit Sub

mnuSwitchOff_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSwitchOff_Click of Form menuForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuTurnFunctionsOn_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : menu item to switch on all of the program's functions
'---------------------------------------------------------------------------------------
'
Private Sub mnuTurnFunctionsOn_Click()
   On Error GoTo mnuTurnFunctionsOn_Click_Error

   Call TurnFunctionsOn
   
   On Error GoTo 0
   Exit Sub

mnuTurnFunctionsOn_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTurnFunctionsOn_Click of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuWidgets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu item to open the deviantart widgets page
'---------------------------------------------------------------------------------------
'
Private Sub mnuWidgets_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo mnuWidgets_Click_Error
    answer = vbYes

    answerMsg = " This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to connect to Steampunk widgets", True, "mnuWidgetsClick")
    'answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.path, 1)
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
' Purpose   : menu item to open the about form
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

