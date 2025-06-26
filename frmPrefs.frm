VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form widgetPrefs 
   Caption         =   "Steampunk Clock Calendar Preferences"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10693.53
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   Visible         =   0   'False
   Begin VB.Timer tmrPrefsMonitorSaveHeight 
      Interval        =   5000
      Left            =   -90
      Top             =   5220
   End
   Begin VB.Timer tmrPrefsScreenResolution 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   -90
      Top             =   6420
   End
   Begin VB.Frame fraDevelopmentButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   5490
      TabIndex        =   38
      Top             =   0
      Width           =   1065
      Begin VB.Label lblDevelopment 
         Caption         =   "Development"
         Height          =   240
         Left            =   45
         TabIndex        =   39
         Top             =   855
         Width           =   960
      End
      Begin VB.Image imgDevelopment 
         Height          =   600
         Left            =   150
         Picture         =   "frmPrefs.frx":0ECA
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgDevelopmentClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1482
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer tmrWritePosition 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   -180
      Top             =   6975
   End
   Begin VB.CheckBox chkEnableResizing 
      Caption         =   "Enable Corner Resize"
      Height          =   210
      Left            =   3240
      TabIndex        =   119
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fraAboutButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7695
      TabIndex        =   86
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   87
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgAbout 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1808
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgAboutClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1D90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraConfigButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1215
      TabIndex        =   40
      Top             =   -15
      Width           =   930
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   41
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgConfig 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":227B
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgConfigClicked 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":285A
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraPositionButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4410
      TabIndex        =   36
      Top             =   0
      Width           =   930
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   37
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgPosition 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":2D5F
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgPositionClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3330
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save the changes you have made to the preferences"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Open the help utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.Frame fraSoundsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3315
      TabIndex        =   11
      Top             =   -15
      Width           =   930
      Begin VB.Label lblSounds 
         Caption         =   "Sounds"
         Height          =   240
         Left            =   210
         TabIndex        =   12
         Top             =   870
         Width           =   615
      End
      Begin VB.Image imgSounds 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":36CE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgSoundsClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3C8D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   -90
      Top             =   5835
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close the utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.Frame fraWindowButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   6615
      TabIndex        =   4
      Top             =   0
      Width           =   975
      Begin VB.Label lblWindow 
         Caption         =   "Window"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgWindow 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":415D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgWindowClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":4627
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraFontsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   930
      Begin VB.Label lblFonts 
         Caption         =   "Fonts"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgFonts 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":49D3
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
      Begin VB.Image imgFontsClicked 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":4F29
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame fraGeneralButton 
      Height          =   1140
      Left            =   240
      TabIndex        =   0
      Top             =   -15
      Width           =   930
      Begin VB.Image imgGeneral 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":53C2
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   855
         Width           =   705
      End
      Begin VB.Image imgGeneralClicked 
         Height          =   600
         Left            =   165
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   7875
      Left            =   195
      TabIndex        =   161
      Top             =   1275
      Visible         =   0   'False
      Width           =   7245
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   7350
         Left            =   465
         TabIndex        =   162
         Top             =   300
         Width           =   6750
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Stop Watch Widget at Windows Startup "
            Height          =   465
            Left            =   1995
            TabIndex        =   172
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   5625
            Width           =   4020
         End
         Begin VB.ComboBox cmbMainGaugeTimeZone 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6A0C
            Left            =   1995
            List            =   "frmPrefs.frx":6A0E
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   2040
            Width           =   3720
         End
         Begin VB.ComboBox cmbMainDaylightSaving 
            Height          =   315
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   3090
            Width           =   3720
         End
         Begin VB.TextBox txtMainBias 
            Height          =   285
            Left            =   5895
            Locked          =   -1  'True
            TabIndex        =   169
            Text            =   "0"
            Top             =   3090
            Width           =   720
         End
         Begin VB.ComboBox cmbTickSwitchPref 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6A10
            Left            =   2010
            List            =   "frmPrefs.frx":6A12
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   6165
            Width           =   3720
         End
         Begin VB.CheckBox chkGaugeFunctions 
            Caption         =   "Ticking toggle *"
            Height          =   225
            Left            =   1995
            TabIndex        =   167
            ToolTipText     =   "When checked this box enables the spinning earth functionality. That's it!"
            Top             =   180
            Width           =   3405
         End
         Begin VB.ComboBox cmbClockFaceSwitchPref 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6A14
            Left            =   2010
            List            =   "frmPrefs.frx":6A16
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   1080
            Width           =   3720
         End
         Begin VB.ComboBox cmbSecondaryDaylightSaving 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   4950
            Width           =   3720
         End
         Begin VB.ComboBox cmbSecondaryGaugeTimeZone 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   4395
            Width           =   3720
         End
         Begin VB.TextBox txtSecondaryBias 
            Height          =   285
            Left            =   5895
            Locked          =   -1  'True
            TabIndex        =   163
            Text            =   "0"
            Top             =   4950
            Width           =   720
         End
         Begin VB.Label lblGeneral 
            Caption         =   "When checked this box enables the clock hands - That's it! *"
            Height          =   660
            Index           =   2
            Left            =   2025
            TabIndex        =   187
            Tag             =   "lblEnableSoundsDesc"
            Top             =   540
            Width           =   3615
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   960
            TabIndex        =   186
            Tag             =   "lblRefreshInterval"
            Top             =   5745
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Gauge Functions :"
            Height          =   315
            Index           =   6
            Left            =   510
            TabIndex        =   185
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Choose the timezone for the main clock. Defaults to the system time. *"
            Height          =   660
            Index           =   4
            Left            =   2025
            TabIndex        =   184
            Top             =   2520
            Width           =   3810
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Main Gauge Time Zone :"
            Height          =   480
            Index           =   5
            Left            =   15
            TabIndex        =   183
            Top             =   2115
            Width           =   1950
         End
         Begin VB.Label lblGeneral 
            Caption         =   "If Daylight Saving is used in your region, choose the appropriate region from the drop down list above. *"
            Height          =   540
            Index           =   7
            Left            =   2025
            TabIndex        =   182
            Top             =   3570
            Width           =   3660
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Daylight Saving :"
            Height          =   345
            Index           =   8
            Left            =   600
            TabIndex        =   181
            Top             =   3165
            Width           =   1365
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Bias (mins)"
            Height          =   345
            Index           =   1
            Left            =   5910
            TabIndex        =   180
            Top             =   2775
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Secondhand Movement :"
            Height          =   480
            Index           =   3
            Left            =   0
            TabIndex        =   179
            Top             =   6225
            Width           =   1950
         End
         Begin VB.Label lblGeneral 
            Caption         =   "The movement of the hand can be set to smooth or one-second ticks, the smooth movement uses more CPU."
            Height          =   660
            Index           =   9
            Left            =   2025
            TabIndex        =   178
            Top             =   6600
            Width           =   3810
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Clock Face to Show :"
            Height          =   345
            Index           =   10
            Left            =   390
            TabIndex        =   177
            Top             =   1140
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Choose the default clock face to show when the widget starts."
            Height          =   660
            Index           =   12
            Left            =   2025
            TabIndex        =   176
            Top             =   1515
            Width           =   3810
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Minor Gauge Daylight Savings Time:"
            Height          =   540
            Index           =   13
            Left            =   285
            TabIndex        =   175
            Top             =   4980
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Minor Gauge Zone :"
            Height          =   330
            Index           =   14
            Left            =   360
            TabIndex        =   174
            Top             =   4425
            Width           =   1995
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Bias (mins)"
            Height          =   345
            Index           =   15
            Left            =   5910
            TabIndex        =   173
            Top             =   4635
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About"
      Height          =   8580
      Left            =   240
      TabIndex        =   88
      Top             =   1155
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton btnGithubHome 
         Caption         =   "Github &Home"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   300
         Width           =   1470
      End
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6225
         Left            =   7980
         TabIndex        =   102
         Top             =   2205
         Width           =   420
      End
      Begin VB.TextBox txtAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Text            =   "frmPrefs.frx":6A18
         Top             =   2205
         Width           =   7935
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1425
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   1050
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   675
         Width           =   1470
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2940
         TabIndex        =   106
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3450
         TabIndex        =   105
         Top             =   510
         Width           =   525
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   104
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3090
         TabIndex        =   103
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2715
         TabIndex        =   100
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1050
         TabIndex        =   99
         Top             =   855
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1065
         TabIndex        =   98
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2715
         TabIndex        =   97
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   96
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1050
         TabIndex        =   95
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Windows Vista, 7, 8, 10  && 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2715
         TabIndex        =   94
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblAbout 
         Caption         =   "(32bit WoW64)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3900
         TabIndex        =   93
         Top             =   510
         Width           =   2655
      End
   End
   Begin VB.Frame fraSounds 
      Caption         =   "Sounds"
      Height          =   2280
      Left            =   855
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Frame fraSoundsInner 
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   765
         TabIndex        =   24
         Top             =   285
         Width           =   6420
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable ALL sounds for the whole widget."
            Height          =   225
            Left            =   1485
            TabIndex        =   34
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   4485
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Determine the sound of UI elements, clock tick and chiming volumes. Set the overall volume to loud or quiet."
            Height          =   540
            Index           =   4
            Left            =   885
            TabIndex        =   139
            Tag             =   "lblSharedInputFile"
            Top             =   750
            Width           =   4680
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   85
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
         End
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   8145
      Left            =   225
      TabIndex        =   8
      Top             =   1230
      Width           =   7605
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   7545
         Left            =   435
         TabIndex        =   33
         Top             =   435
         Width           =   6450
         Begin VB.Frame fraClockTooltips 
            BorderStyle     =   0  'None
            Height          =   1110
            Left            =   1785
            TabIndex        =   157
            Top             =   2685
            Width           =   3345
            Begin VB.OptionButton optGaugeTooltips 
               Caption         =   "Disable Clock Tooltips *"
               Height          =   300
               Index           =   2
               Left            =   225
               TabIndex        =   160
               Top             =   780
               Width           =   2790
            End
            Begin VB.OptionButton optGaugeTooltips 
               Caption         =   "Clock - Enable Square Tooltips"
               Height          =   300
               Index           =   1
               Left            =   225
               TabIndex        =   159
               Top             =   450
               Width           =   2790
            End
            Begin VB.OptionButton optGaugeTooltips 
               Caption         =   "Clock - Enable Balloon Tooltips *"
               Height          =   315
               Index           =   0
               Left            =   225
               TabIndex        =   158
               Top             =   120
               Width           =   3060
            End
         End
         Begin vb6projectCCRSlider.Slider sliGaugeSize 
            Height          =   390
            Left            =   1920
            TabIndex        =   154
            Top             =   30
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   5
            Max             =   200
            Value           =   5
            TickFrequency   =   3
            SelStart        =   5
         End
         Begin VB.Frame fraPrefsTooltips 
            BorderStyle     =   0  'None
            Height          =   1125
            Index           =   0
            Left            =   1860
            TabIndex        =   144
            Top             =   3810
            Width           =   3150
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Disable Prefs Tooltips *"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   156
               Top             =   780
               Width           =   2970
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable Balloon Tooltips *"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   146
               Top             =   120
               Width           =   2760
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable SquareTooltips *"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   145
               Top             =   450
               Width           =   2970
            End
         End
         Begin VB.Frame fraMainTooltips 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   1995
            TabIndex        =   143
            Top             =   2775
            Width           =   4095
         End
         Begin VB.CheckBox chkShowHelp 
            Caption         =   "Show Help on Widget Start"
            Height          =   225
            Left            =   2010
            TabIndex        =   138
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   5535
            Width           =   3405
         End
         Begin VB.CheckBox chkDpiAwareness 
            Caption         =   "DPI Awareness Enable *"
            Height          =   285
            Left            =   2010
            TabIndex        =   129
            ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
            Top             =   5880
            Width           =   3405
         End
         Begin VB.CheckBox chkShowTaskbar 
            Caption         =   "Show Widget in Taskbar"
            Height          =   225
            Left            =   2010
            TabIndex        =   127
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   5160
            Width           =   3405
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   80
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1695
            Width           =   2490
         End
         Begin VB.Label lblConfiguration 
            Caption         =   $"frmPrefs.frx":79CF
            Height          =   915
            Index           =   0
            Left            =   1980
            TabIndex        =   130
            Top             =   6240
            Width           =   4335
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   675
            Index           =   6
            Left            =   2025
            TabIndex        =   107
            Top             =   2115
            Width           =   3990
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "160"
            Height          =   315
            Index           =   4
            Left            =   4740
            TabIndex        =   84
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "120"
            Height          =   315
            Index           =   3
            Left            =   3990
            TabIndex        =   83
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   82
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   81
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1740
            Width           =   2055
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2070
            TabIndex        =   79
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   870
            Width           =   3810
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Gauge Size :"
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   78
            Top             =   105
            Width           =   975
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "80"
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   77
            Top             =   555
            Width           =   360
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "200 (%)"
            Height          =   315
            Index           =   5
            Left            =   5385
            TabIndex        =   76
            Top             =   555
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2085
            TabIndex        =   75
            Top             =   555
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   6195
      Left            =   255
      TabIndex        =   9
      Top             =   1230
      Width           =   8280
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   5400
         Left            =   690
         TabIndex        =   25
         Top             =   360
         Width           =   6105
         Begin VB.TextBox txtDisplayScreenFont 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   149
            Text            =   "Courier  New"
            Top             =   3255
            Width           =   3285
         End
         Begin VB.CommandButton btnDisplayScreenFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5010
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   3255
            Width           =   585
         End
         Begin VB.TextBox txtDisplayScreenFontSize 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "8"
            Top             =   3795
            Width           =   510
         End
         Begin VB.CommandButton btnResetMessages 
            Caption         =   "Reset"
            Height          =   300
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   4695
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   120
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1065
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "8"
            ToolTipText     =   "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
            Top             =   1065
            Width           =   510
         End
         Begin VB.CommandButton btnPrefsFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5025
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   90
            Width           =   585
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Times New Roman"
            Top             =   90
            Width           =   3285
         End
         Begin VB.Label lblFontsTab 
            Caption         =   $"frmPrefs.frx":7A83
            Height          =   1710
            Index           =   0
            Left            =   1680
            TabIndex        =   188
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   1545
            Width           =   4455
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in the console display screen on the main clock *"
            Height          =   480
            Index           =   9
            Left            =   2415
            TabIndex        =   152
            Top             =   3780
            Width           =   4035
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Display Console Font :"
            Height          =   300
            Index           =   8
            Left            =   0
            TabIndex        =   151
            Tag             =   "lblPrefsFont"
            Top             =   3270
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Console  Font Size :"
            Height          =   330
            Index           =   5
            Left            =   165
            TabIndex        =   150
            Tag             =   "lblPrefsFontSize"
            Top             =   3825
            Width           =   1590
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Hidden message boxes can be reactivated by pressing this reset button."
            Height          =   480
            Index           =   4
            Left            =   2670
            TabIndex        =   136
            Top             =   4635
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   405
            TabIndex        =   134
            Tag             =   "lblPrefsFont"
            Top             =   4740
            Width           =   1470
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Index           =   10
            Left            =   4920
            TabIndex        =   121
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size *"
            Height          =   480
            Index           =   7
            Left            =   2310
            TabIndex        =   32
            Top             =   1095
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Base Font Size :"
            Height          =   330
            Index           =   3
            Left            =   435
            TabIndex        =   31
            Tag             =   "lblPrefsFontSize"
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Prefs Utility Font :"
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   30
            Tag             =   "lblPrefsFont"
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window, gauge tooltips and message boxes *"
            Height          =   480
            Index           =   6
            Left            =   1695
            TabIndex        =   29
            Top             =   480
            Width           =   4035
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   7935
      Left            =   225
      TabIndex        =   10
      Top             =   1620
      Width           =   8235
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   7500
         Left            =   165
         TabIndex        =   14
         Top             =   345
         Width           =   7005
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   2115
            TabIndex        =   155
            Top             =   4575
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   20
            SmallChange     =   2
            SelStart        =   20
         End
         Begin VB.ComboBox cmbMultiMonitorResize 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   5805
            Width           =   3720
         End
         Begin VB.Frame fraHiding 
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   1395
            TabIndex        =   108
            Top             =   2325
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   1605
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   109
               Top             =   210
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   112
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":7BC1
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   110
               Top             =   600
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   0
            Width           =   3720
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse *"
            Height          =   225
            Left            =   2250
            TabIndex        =   15
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   1500
            Width           =   2535
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Multi-Monitor Resizing :"
            Height          =   255
            Index           =   11
            Left            =   375
            TabIndex        =   141
            Top             =   5835
            Width           =   1830
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   $"frmPrefs.frx":7C64
            Height          =   1140
            Index           =   10
            Left            =   2235
            TabIndex        =   140
            Top             =   6255
            Width           =   4050
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This setting controls the relative layering of this widget. You may use it to place it on top of other windows or underneath. "
            Height          =   660
            Index           =   3
            Left            =   2235
            TabIndex        =   117
            Top             =   570
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level :"
            Height          =   345
            Index           =   0
            Left            =   915
            TabIndex        =   23
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "20%"
            Height          =   315
            Index           =   7
            Left            =   2205
            TabIndex        =   22
            Top             =   5070
            Width           =   345
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "100%"
            Height          =   315
            Index           =   9
            Left            =   5565
            TabIndex        =   21
            Top             =   5070
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "60%"
            Height          =   315
            Index           =   8
            Left            =   3975
            TabIndex        =   20
            Top             =   5070
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity:"
            Height          =   315
            Index           =   6
            Left            =   1470
            TabIndex        =   19
            Top             =   4620
            Width           =   780
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Index           =   5
            Left            =   2250
            TabIndex        =   18
            Top             =   5385
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the program to ignore all mouse events except right click menu interactions."
            Height          =   660
            Index           =   4
            Left            =   2235
            TabIndex        =   17
            Top             =   1890
            Width           =   3810
         End
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   45
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   122
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   124
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   123
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":7D7B
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   128
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":7E42
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   126
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   125
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   53
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   49
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":7EE6
            Left            =   1530
            List            =   "frmPrefs.frx":7EE8
            Style           =   2  'Dropdown List
            TabIndex        =   46
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   55
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   54
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg.  mmsys.cpl to run the sounds utility."
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   52
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   51
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   48
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   47
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   7440
      Left            =   270
      TabIndex        =   42
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   6960
         Left            =   150
         TabIndex        =   43
         Top             =   300
         Width           =   7680
         Begin VB.TextBox txtLandscapeHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   67
            Top             =   4425
            Width           =   2130
         End
         Begin VB.CheckBox chkPreventDragging 
            Caption         =   "Widget Position Locked. *"
            Height          =   225
            Left            =   2265
            TabIndex        =   115
            ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
            Top             =   3465
            Width           =   2505
         End
         Begin VB.TextBox txtPortraitYoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   73
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6465
            Width           =   2130
         End
         Begin VB.TextBox txtPortraitHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   71
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6000
            Width           =   2130
         End
         Begin VB.TextBox txtLandscapeVoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   69
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   4875
            Width           =   2130
         End
         Begin VB.ComboBox cmbWidgetLandscape 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   3930
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPortrait 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Choose the alarm sound."
            Top             =   5505
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   59
            ToolTipText     =   "Choose the alarm sound."
            Top             =   2100
            Width           =   2160
         End
         Begin VB.ComboBox cmbAspectHidden 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   56
            ToolTipText     =   "Choose the alarm sound."
            Top             =   0
            Width           =   2160
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   7
            Left            =   4530
            TabIndex        =   132
            Tag             =   "lblPrefixString"
            Top             =   6495
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   5
            Left            =   4530
            TabIndex        =   131
            Tag             =   "lblPrefixString"
            Top             =   6045
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   118
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   4
            Left            =   4530
            TabIndex        =   114
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   2
            Left            =   4530
            TabIndex        =   113
            Tag             =   "lblPrefixString"
            Top             =   4500
            Width           =   390
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Top Y pos :"
            Height          =   510
            Index           =   17
            Left            =   645
            TabIndex        =   74
            Tag             =   "lblPrefixString"
            Top             =   6480
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Left X pos :"
            Height          =   510
            Index           =   16
            Left            =   660
            TabIndex        =   72
            Tag             =   "lblPrefixString"
            Top             =   6015
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Top Y pos :"
            Height          =   510
            Index           =   15
            Left            =   420
            TabIndex        =   70
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Left X pos :"
            Height          =   510
            Index           =   14
            Left            =   420
            TabIndex        =   68
            Tag             =   "lblPrefixString"
            Top             =   4455
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Landscape :"
            Height          =   435
            Index           =   13
            Left            =   450
            TabIndex        =   66
            Tag             =   "lblAlarmSound"
            Top             =   3975
            Width           =   2115
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":7EEA
            Height          =   3435
            Index           =   12
            Left            =   5145
            TabIndex        =   64
            Tag             =   "lblAlarmSoundDesc"
            Top             =   3480
            Width           =   2520
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Portrait :"
            Height          =   375
            Index           =   11
            Left            =   690
            TabIndex        =   63
            Tag             =   "lblAlarmSound"
            Top             =   5550
            Width           =   2040
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":80BC
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   61
            Tag             =   "lblAlarmSoundDesc"
            Top             =   2550
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent:"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   60
            Tag             =   "lblAlarmSound"
            Top             =   2145
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":815B
            Height          =   3045
            Index           =   6
            Left            =   2265
            TabIndex        =   58
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   5370
         End
         Begin VB.Label lblPosition 
            Caption         =   "Aspect Ratio Hidden Mode :"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   2145
         End
      End
   End
   Begin VB.Label lblDragCorner 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   8700
      TabIndex        =   137
      ToolTipText     =   "drag me"
      Top             =   10350
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   133
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   116
      Top             =   10155
      Width           =   3870
   End
   Begin VB.Menu prefsMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About Panzer Earth Widget"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuClosePreferences 
         Caption         =   "Close Preferences"
      End
   End
End
Attribute VB_Name = "widgetPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder

'---------------------------------------------------------------------------------------
' Module    : widgetPrefs
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : VB6 standard form to display the prefs
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' Constants and APIs to create and subclass the dragCorner
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants defined for setting a theme to the prefs
Private Const COLOR_BTNFACE As Long = 15

' APIs declared for setting a theme to the prefs
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Private Types for determining prefs sizing
Private pvtPrefsDynamicSizingFlg As Boolean
Private pvtLastFormHeight As Long
Private Const pvtcPrefsFormHeight As Long = 11055
Private Const pvtcPrefsFormWidth  As Long = 9090

Private pvtPrefsFormResizedByDrag As Boolean

'    gblPrefsStartWidth = 9075
'    gblPrefsStartHeight = 16450
'------------------------------------------------------ ENDS

Private pvtPrefsStartupFlg As Boolean
Private pvtAllowSizeChangeFlg As Boolean

' module level balloon tooltip variables for subclassed comboBoxes ONLY.
Private pCmbMultiMonitorResizeBalloonTooltip As String
Private pCmbScrollWheelDirectionBalloonTooltip As String
Private pCmbWindowLevelBalloonTooltip As String
Private pCmbHidingTimeBalloonTooltip As String
Private pCmbAspectHiddenBalloonTooltip As String
Private pCmbWidgetPositionBalloonTooltip As String
Private pCmbWidgetLandscapeBalloonTooltip As String
Private pCmbWidgetPortraitBalloonTooltip As String
Private pCmbDebugBalloonTooltip As String




Private mIsLoaded As Boolean ' property
Private mGaugeSize As Single   ' property

Private gblConstraintRatio As Double



Private Sub btnDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDefaultEditor.hWnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGithubHome_Click
' Author    : beededea
' Date      : 22/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGithubHome_Click()
   On Error GoTo btnGithubHome_Click_Error

    Call menuForm.mnuGithubHome_Click

   On Error GoTo 0
   Exit Sub

btnGithubHome_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGithubHome_Click of Form widgetPrefs"
End Sub








'---------------------------------------------------------------------------------------
' Procedure : optGaugeTooltips_Click
' Author    : beededea
' Date      : 19/08/2023
' Purpose   : three options radio buttons for selecting the gauge/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optGaugeTooltips_Click(Index As Integer)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo optGaugeTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button

    If pvtPrefsStartupFlg = False Then
        gblGaugeTooltips = CStr(Index)
    
        optGaugeTooltips(0).Tag = CStr(Index)
        optGaugeTooltips(1).Tag = CStr(Index)
        optGaugeTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\UBoatStopWatch", "gaugeTooltips", gblGaugeTooltips, gblSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "optGaugeTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadProgram
        End If
    End If


   On Error GoTo 0
   Exit Sub

optGaugeTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGaugeTooltips_Click of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkGaugeFunctions_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGaugeFunctions_Click()
    On Error GoTo chkGaugeFunctions_Click_Error

    btnSave.Enabled = True ' enable the save button
    overlayWidget.Ticking = chkGaugeFunctions.Value

    On Error GoTo 0
    Exit Sub

chkGaugeFunctions_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGaugeFunctions_Click of Form widgetPrefs"
End Sub

' Procedure : cmbClockFaceSwitchPref_Click
' Author    : beededea
' Date      : 08/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbClockFaceSwitchPref_Click()
    On Error GoTo cmbClockFaceSwitchPref_Click_Error

    If pvtPrefsStartupFlg = False Then ' don't run this on startup
        btnSave.Enabled = True ' enable the save button
    End If

    On Error GoTo 0
    Exit Sub

cmbClockFaceSwitchPref_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbClockFaceSwitchPref_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbMainDaylightSaving_Click
' Author    : beededea
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbMainDaylightSaving_Click()
    
   On Error GoTo cmbMainDaylightSaving_Click_Error

    If pvtPrefsStartupFlg = False Then ' don't run this on startup
        btnSave.Enabled = True ' enable the save button
        If cmbMainDaylightSaving.ListIndex <> 0 Then
            tzDelta = fObtainDaylightSavings("Main") ' determine the time bias
        End If
    End If
   
   On Error GoTo 0
   Exit Sub

cmbMainDaylightSaving_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMainDaylightSaving_Click of Form widgetPrefs"
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbMainGaugeTimeZone_Click
' Author    : beededea
' Date      : 29/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbMainGaugeTimeZone_Click()
    
    
   On Error GoTo cmbMainGaugeTimeZone_Click_Error

    If pvtPrefsStartupFlg = False Then ' don't run this on startup
        btnSave.Enabled = True ' enable the save button
        tzDelta = fObtainDaylightSavings("Main") ' determine the time bias
    End If

    On Error GoTo 0
    Exit Sub

cmbMainGaugeTimeZone_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMainGaugeTimeZone_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbSecondaryDaylightSaving_Click
' Author    : beededea
' Date      : 08/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSecondaryDaylightSaving_Click()
    
    
    On Error GoTo cmbSecondaryDaylightSaving_Click_Error

    If pvtPrefsStartupFlg = False Then ' don't run this on startup
        btnSave.Enabled = True ' enable the save button
        If cmbSecondaryDaylightSaving.ListIndex <> 0 Then
            tzDelta1 = fObtainDaylightSavings("Secondary") ' determine the time bias to display on the scondary gauge when in clock mode
        End If
    End If

    On Error GoTo 0
    Exit Sub

cmbSecondaryDaylightSaving_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbSecondaryDaylightSaving_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbSecondaryGaugeTimeZone_Click
' Author    : beededea
' Date      : 08/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSecondaryGaugeTimeZone_Click()
    
    
    On Error GoTo cmbSecondaryGaugeTimeZone_Click_Error

    If pvtPrefsStartupFlg = False Then ' don't run this on startup
        btnSave.Enabled = True ' enable the save button
        tzDelta1 = fObtainDaylightSavings("Secondary") ' determine the time bias
    End If

    On Error GoTo 0
    Exit Sub

cmbSecondaryGaugeTimeZone_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbSecondaryGaugeTimeZone_Click of Form widgetPrefs"
End Sub

Private Sub cmbTickSwitchPref_Click()
   btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_Click
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : three options radio buttons for selecting the VB6 preference form tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_Click(Index As Integer)

   On Error GoTo optPrefsTooltips_Click_Error

    If pvtPrefsStartupFlg = False Then
    
        btnSave.Enabled = True ' enable the save button
        gblPrefsTooltips = CStr(Index)
        optPrefsTooltips(0).Tag = CStr(Index)
        optPrefsTooltips(1).Tag = CStr(Index)
        optPrefsTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\UBoatStopWatch", "prefsTooltips", gblPrefsTooltips, gblSettingsFile
        
        ' set the tooltips on the prefs screen
        Call setPrefsTooltips
    End If
     
   On Error GoTo 0
   Exit Sub

optPrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_Click of Form widgetPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkEnableChimes_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableChimes_Click()
   On Error GoTo chkEnableChimes_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableChimes_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableChimes_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkNumericDisplayRotation_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkNumericDisplayRotation_Click()
   On Error GoTo chkNumericDisplayRotation_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkNumericDisplayRotation_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkNumericDisplayRotation_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbMultiMonitorResize_Click
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on.
'---------------------------------------------------------------------------------------
'
Private Sub cmbMultiMonitorResize_Click()
   On Error GoTo cmbMultiMonitorResize_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If pvtPrefsStartupFlg = True Then Exit Sub
    
    gblMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
    
    If cmbMultiMonitorResize.ListIndex = 2 Then
        If prefsMonitorStruct.IsPrimary = True Then
            gblGaugePrimaryHeightRatio = fGauge.gaugeForm.WidgetRoot.Zoom
            sPutINISetting "Software\UBoatStopWatch", "gaugePrimaryHeightRatio", gblGaugePrimaryHeightRatio, gblSettingsFile
            
            'gblPrefsPrimaryHeightTwips = Trim$(cstr(widgetPrefs.Height))
            sPutINISetting "Software\UBoatStopWatch", "prefsPrimaryHeightTwips", gblPrefsPrimaryHeightTwips, gblSettingsFile
        Else
            gblGaugeSecondaryHeightRatio = fGauge.gaugeForm.WidgetRoot.Zoom
            sPutINISetting "Software\UBoatStopWatch", "gaugeSecondaryHeightRatio", gblGaugeSecondaryHeightRatio, gblSettingsFile
            
            'gblPrefsSecondaryHeightTwips = Trim$(cstr(widgetPrefs.Height))
            sPutINISetting "Software\UBoatStopWatch", "prefsSecondaryHeightTwips", gblPrefsSecondaryHeightTwips, gblSettingsFile
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmbMultiMonitorResize_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMultiMonitorResize_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkVolumeBoost_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkVolumeBoost_Click()
   On Error GoTo chkVolumeBoost_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkVolumeBoost_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkVolumeBoost_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableTicks_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTicks_Click()
   On Error GoTo chkEnableTicks_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableTicks_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTicks_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkShowHelp_Click
' Author    : beededea
' Date      : 03/07/2024
' Purpose   : show help on the program startup
'---------------------------------------------------------------------------------------
'
Private Sub chkShowHelp_Click()
   On Error GoTo chkShowHelp_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkShowHelp.Value = 1 Then
        gblShowHelp = "1"
    Else
        gblShowHelp = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowHelp_Click of Form widgetPrefs"
End Sub





' ----------------------------------------------------------------
' Procedure Name: Form_Initialize
' Purpose:
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: beededea
' Date: 05/10/2023
' ----------------------------------------------------------------
Private Sub Form_Initialize()
    On Error GoTo Form_Initialize_Error
    
    ' initialise private variables
    Call initialisePrefsVars

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form widgetPrefs"
    
    End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_Load     WidgetPrefs
' Author    : beededea
' Date      : 25/04/2023
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    Me.Visible = False
    btnSave.Enabled = False ' disable the save button
    Me.mnuAbout.Caption = "About UBoat StopWatch Cairo " & gblCodingEnvironment & " widget"

    pvtPrefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    'pvtPrefsDynamicSizingFlg = False
    IsLoaded = True
    gblWindowLevelWasChanged = False
    gblPrefsStartWidth = pvtcPrefsFormWidth
    gblPrefsStartHeight = pvtcPrefsFormHeight
    pvtPrefsFormResizedByDrag = False
            
    ' subclass ALL forms created by intercepting WM_Create messages, identifying dialog forms to centre them in the middle of the monitor - specifically the font form.
    If Not InIDE Then subclassDialogForms
    
    ' subclass specific WidgetPrefs controls that need additional functionality that VB6 does not provide (scrollwheel/balloon tooltips)
    Call subClassControls
    
    ' set form resizing
    Call setFormResizingVars
    
    ' note the monitor primary at the preferences form_load and store as gblOldgaugeFormMonitorPrimary
    Call identifyPrefsPrimaryMonitor
    
    ' reverts TwinBasic form themeing to that of the earlier classic look and feel
    #If TWINBASIC Then
       Call setVisualStyles
    #End If
       
    ' read the last saved position from the settings.ini
    Call readPrefsPosition
        
    ' determine the frame heights in dynamic sizing or normal mode
    Call setframeHeights
    
    ' set the text in any labels that need a vbCrLf to space the text
    Call setPrefsLabels
    
    ' populate all the comboboxes in the prefs form
    Call populatePrefsComboBoxes
        
    ' adjust all the preferences and main program controls
    Call adjustPrefsControls
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips
    
    ' adjust the theme used by the prefs alone
    Call adjustPrefsTheme
    
    ' size and position the frames and buttons
    Call positionPrefsFramesButtons
    
    ' make the last used tab appear on startup
    Call showLastTab
    
    ' load the about text and load into prefs
    Call loadPrefsAboutText
    
    ' load the preference icons from a previously populated CC imageList
    Call loadHigherResPrefsImages
    
    ' set the height of the whole form not higher than the screen size, cause a form_resize event
    Call setPrefsHeight
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' start the timers
    Call startPrefsTimers
    
    widgetPrefsOldHeight = widgetPrefs.Height
    widgetPrefsOldWidth = widgetPrefs.Width
    
    ' end the startup by un-setting the start global-ish flag
    pvtPrefsStartupFlg = False
    
    btnSave.Enabled = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form widgetPrefs"

End Sub



' ---------------------------------------------------------------------------------------
' Procedure : initialisePrefsVars
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : initialise private variables
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrefsVars()

   On Error GoTo initialisePrefsVars_Error

    pvtPrefsDynamicSizingFlg = False
    pvtLastFormHeight = 0
    pvtPrefsStartupFlg = False
    pvtAllowSizeChangeFlg = False
    pCmbMultiMonitorResizeBalloonTooltip = vbNullString
    pCmbScrollWheelDirectionBalloonTooltip = vbNullString
    pCmbWindowLevelBalloonTooltip = vbNullString
    pCmbHidingTimeBalloonTooltip = vbNullString
    pCmbAspectHiddenBalloonTooltip = vbNullString
    pCmbWidgetPositionBalloonTooltip = vbNullString
    pCmbWidgetLandscapeBalloonTooltip = vbNullString
    pCmbWidgetPortraitBalloonTooltip = vbNullString
    pCmbDebugBalloonTooltip = vbNullString
    pvtPrefsFormResizedByDrag = False
    mIsLoaded = False ' property

   On Error GoTo 0
   Exit Sub

initialisePrefsVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrefsVars of Form widgetPrefs"

End Sub

     
'
'---------------------------------------------------------------------------------------
' Procedure : setFormResizingVars
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set form resizing characteristics
'---------------------------------------------------------------------------------------
'
Private Sub setFormResizingVars()

   On Error GoTo setFormResizingVars_Error

    With lblDragCorner
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
      .Visible = False
    End With
    
    If gblDpiAwareness = "1" Then
        pvtPrefsDynamicSizingFlg = True
        chkEnableResizing.Value = 1
        lblDragCorner.Visible = True
    End If
    
    widgetPrefsOldHeight = widgetPrefs.Height
    widgetPrefsOldWidth = widgetPrefs.Width

   On Error GoTo 0
   Exit Sub

setFormResizingVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormResizingVars of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : identifyPrefsPrimaryMonitor
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : note the monitor primary at the preferences form_load and store as gblOldPrefsFormMonitorPrimary - will be resampled regularly later and compared
'---------------------------------------------------------------------------------------
'
Private Sub identifyPrefsPrimaryMonitor()
    'Dim prefsFormHeight As Long: prefsFormHeight = 0
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo identifyPrefsPrimaryMonitor_Error
    
    'prefsFormHeight = gblPrefsStartHeight

    prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
    gblOldPrefsFormMonitorPrimary = prefsMonitorStruct.IsPrimary ' -1 true

   On Error GoTo 0
   Exit Sub

identifyPrefsPrimaryMonitor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyPrefsPrimaryMonitor of Form widgetPrefs"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setPrefsHeight
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set the height of the whole form not higher than the screen size, cause a form_resize event
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsHeight()

   On Error GoTo setPrefsHeight_Error
   
    ' constrain the height/width ratio
    gblConstraintRatio = pvtcPrefsFormHeight / pvtcPrefsFormWidth

    If gblDpiAwareness = "1" Then
        gblPrefsFormResizedInCode = True
        If gblPrefsPrimaryHeightTwips < gblPhysicalScreenHeightTwips Then
            widgetPrefs.Height = CLng(gblPrefsPrimaryHeightTwips) ' 16450
        Else
            widgetPrefs.Height = gblPhysicalScreenHeightTwips - 1000
        End If
    End If

   On Error GoTo 0
   Exit Sub

setPrefsHeight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsHeight of Form widgetPrefs"
End Sub
   
'---------------------------------------------------------------------------------------
' Procedure : startPrefsTimers
' Author    : beededea
' Date      : 20/02/2025
' Purpose   :  start the timers
'---------------------------------------------------------------------------------------
'
Private Sub startPrefsTimers()

    ' start the timer that records the prefs position every 10 seconds
   On Error GoTo startPrefsTimers_Error

    tmrWritePosition.Enabled = True
    
    ' start the timer that detects a MOVE event on the preferences form
    tmrPrefsScreenResolution.Enabled = True

   On Error GoTo 0
   Exit Sub

startPrefsTimers_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startPrefsTimers of Form widgetPrefs"

End Sub
    

#If TWINBASIC Then
    '---------------------------------------------------------------------------------------
    ' Procedure : setVisualStyles
    ' Author    : beededea
    ' Date      : 13/01/2025
    ' Purpose   : loop through all the controls and identify the labels and text boxes and disable modern styles
    '             reverts TwinBasic form themeing to that of the earlier classic look and feel.
    '---------------------------------------------------------------------------------------
    '
        Private Sub setVisualStyles()
            Dim Ctrl As Control
          
            On Error GoTo setVisualStyles_Error

            For Each Ctrl In widgetPrefs.Controls
                If (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
                    Ctrl.VisualStyles = False
                End If
            Next

       On Error GoTo 0
       Exit Sub

setVisualStyles_Error:

        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setVisualStyles of Form widgetPrefs"
        End Sub
#End If



'---------------------------------------------------------------------------------------
' Procedure : subClassControls
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : sub classing code to capture form movement and intercept messages to the comboboxes to provide missing balloon tooltips functionality
'---------------------------------------------------------------------------------------
'
Private Sub subClassControls()
    
   On Error GoTo subClassControls_Error

    If InIDE Then
        MsgBox "NOTE: Running in IDE so Sub classing is disabled" & vbCrLf & "Mousewheel will not scroll icon maps and balloon tooltips will not display on comboboxes" & vbCrLf & vbCrLf & _
            "In addition, the display screen will not show messages as it currently crashes when run within the IDE."
    Else
        ' sub classing code to intercept messages to the form itself in order to capture WM_EXITSIZEMOVE messages that occur AFTER the form has been resized
        
        Call SubclassForm(widgetPrefs.hWnd, ObjPtr(widgetPrefs))
        
        'now the comboboxes in order to capture the mouseOver and display the balloon tooltips
        
        Call SubclassComboBox(cmbMultiMonitorResize.hWnd, ObjPtr(cmbMultiMonitorResize))
        Call SubclassComboBox(cmbScrollWheelDirection.hWnd, ObjPtr(cmbScrollWheelDirection))
        Call SubclassComboBox(cmbWindowLevel.hWnd, ObjPtr(cmbWindowLevel))
        Call SubclassComboBox(cmbHidingTime.hWnd, ObjPtr(cmbHidingTime))
        
        Call SubclassComboBox(cmbWidgetLandscape.hWnd, ObjPtr(cmbWidgetLandscape))
        Call SubclassComboBox(cmbWidgetPortrait.hWnd, ObjPtr(cmbWidgetPortrait))
        Call SubclassComboBox(cmbWidgetPosition.hWnd, ObjPtr(cmbWidgetPosition))
        Call SubclassComboBox(cmbAspectHidden.hWnd, ObjPtr(cmbAspectHidden))
        Call SubclassComboBox(cmbDebug.hWnd, ObjPtr(cmbDebug))
        
        
    End If

    On Error GoTo 0
    Exit Sub

subClassControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subClassControls of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MouseMoveOnComboText
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Add a balloon tooltip dynamically to combo boxes using subclassing, called by combobox_proc
'             (VB6 will not allow Elroy's advanced tooltips to show on VB6 comboboxes, we must subclass the controls)
'             Note: Each control must also be added to the subClassControls routine
'---------------------------------------------------------------------------------------
'
Public Sub MouseMoveOnComboText(sComboName As String)
    Dim sTitle As String
    Dim sText As String

    On Error GoTo MouseMoveOnComboText_Error
    
    Select Case sComboName
        Case "cmbMultiMonitorResize"
            sTitle = "Help on the Drop Down Icon Filter"
            sText = pCmbMultiMonitorResizeBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbMultiMonitorResize.hWnd, sText, , sTitle, , , , True
        Case "cmbScrollWheelDirection"
            sTitle = "Help on the Scroll Wheel Direction"
            sText = pCmbScrollWheelDirectionBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbScrollWheelDirection.hWnd, sText, , sTitle, , , , True
        Case "cmbWindowLevel"
            sTitle = "Help on the Window Level"
            sText = pCmbWindowLevelBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWindowLevel.hWnd, sText, , sTitle, , , , True
        Case "cmbHidingTime"
            sTitle = "Help on the Hiding Time"
            sText = pCmbHidingTimeBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbHidingTime.hWnd, sText, , sTitle, , , , True
            
        Case "cmbAspectHidden"
            sTitle = "Help on Hiding in Landscape/Portrait Mode"
            sText = pCmbAspectHiddenBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbAspectHidden.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetPosition"
            sTitle = "Help on Widget Position in Landscape/Portrait Modes"
            sText = pCmbWidgetPositionBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWidgetPosition.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetLandscape"
            sTitle = "Help on Widget Locking in Landscape Mode"
            sText = pCmbWidgetLandscapeBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWidgetLandscape.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetPortrait"
            sTitle = "Help on Widget Locking in Portrait Mode"
            sText = pCmbWidgetPortraitBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWidgetPortrait.hWnd, sText, , sTitle, , , , True
        Case "cmbDebug"
            sTitle = "Help on Debug Mode"
            sText = pCmbDebugBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbDebug.hWnd, sText, , sTitle, , , , True
        
    End Select
    
   On Error GoTo 0
   Exit Sub

MouseMoveOnComboText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseMoveOnComboText of Form widgetPrefs"
End Sub


' ---------------------------------------------------------------------------------------
' Procedure : positionPrefsMonitor
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : position the prefs on the current monitor
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsMonitor()

    Dim formLeftTwips As Long: formLeftTwips = 0
    Dim formTopTwips As Long: formTopTwips = 0
    'Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo positionPrefsMonitor_Error
    
    If gblDpiAwareness = "1" Then
        formLeftTwips = Val(gblPrefsHighDpiXPosTwips)
        formTopTwips = Val(gblPrefsHighDpiYPosTwips)
    Else
        formLeftTwips = Val(gblPrefsLowDpiXPosTwips)
        formTopTwips = Val(gblPrefsLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fGauge.gaugeForm.Left + fGauge.gaugeForm.Width) * gblScreenTwipsPerPixelX) + 200 + widgetPrefs.Width > gblPhysicalScreenWidthTwips Then
            widgetPrefs.Left = (fGauge.gaugeForm.Left * gblScreenTwipsPerPixelX) - (widgetPrefs.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    
    If formLeftTwips <> 0 Then
        widgetPrefs.Left = formLeftTwips
    Else
        widgetPrefs.Left = gblPhysicalScreenWidthTwips / 2 - widgetPrefs.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        widgetPrefs.Top = formTopTwips
    Else
        widgetPrefs.Top = Screen.Height / 2 - widgetPrefs.Height / 2
    End If
    
    'monitorCount = fGetMonitorCount
    If gblMonitorCount > 1 Then Call SetFormOnMonitor(Me.hWnd, formLeftTwips / fTwipsPerPixelX, formTopTwips / fTwipsPerPixelY)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        widgetPrefs.Left = 10
    End If
    If Me.Top < 0 Then
        widgetPrefs.Top = 0
    End If
    If Me.Left > gblVirtualScreenWidthTwips - 2500 Then
        widgetPrefs.Left = gblVirtualScreenWidthTwips - 2500
    End If
    If Me.Top > gblVirtualScreenHeightTwips - 2500 Then
        widgetPrefs.Top = gblVirtualScreenHeightTwips - 2500
    End If
    
    
    ' if just one monitor or the global switch is off then exit
    If gblMonitorCount > 1 And LTrim$(gblMultiMonitorResize) = "2" Then

        If prefsMonitorStruct.IsPrimary = True Then
            gblPrefsFormResizedInCode = True
            gblPrefsPrimaryHeightTwips = fGetINISetting("Software\UBoatStopWatch", "prefsPrimaryHeightTwips", gblSettingsFile)
            If Val(gblPrefsPrimaryHeightTwips) <= 0 Then
                widgetPrefs.Height = gblPrefsStartHeight
            Else
                widgetPrefs.Height = CLng(gblPrefsPrimaryHeightTwips)
            End If
        Else
            gblPrefsSecondaryHeightTwips = fGetINISetting("Software\UBoatStopWatch", "prefsSecondaryHeightTwips", gblSettingsFile)
            gblPrefsFormResizedInCode = True
            If Val(gblPrefsSecondaryHeightTwips) <= 0 Then
                widgetPrefs.Height = gblPrefsStartHeight
            Else
                widgetPrefs.Height = CLng(gblPrefsSecondaryHeightTwips)
            End If
        End If
    End If

    '' fGauge.RotateBusyTimer = True
    
    On Error GoTo 0
    Exit Sub

positionPrefsMonitor_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsMonitor of Form widgetPrefs"
End Sub
    
    


'---------------------------------------------------------------------------------------
' Procedure : chkDpiAwareness_Click
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : toggle for setting the DPI awareness
'---------------------------------------------------------------------------------------
'
Private Sub chkDpiAwareness_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo chkDpiAwareness_Click_Error

    btnSave.Enabled = True ' enable the save button
    If pvtPrefsStartupFlg = False Then ' don't run this on startup
                    
        answer = vbYes
        answerMsg = "You must close this widget and HARD restart it, in order to change the widget's DPI awareness (a simple soft reload just won't cut it), do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "DpiAwareness Confirmation", True, "chkDpiAwarenessRestart")
        
        If chkDpiAwareness.Value = 0 Then
            gblDpiAwareness = "0"
        Else
            gblDpiAwareness = "1"
        End If

        sPutINISetting "Software\UBoatStopWatch", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "OK, the widget is still DPI aware until you restart. Some forms may show abnormally."
            answer = msgBoxA(answerMsg, vbOKOnly, "DpiAwareness Notification", True, "chkDpiAwarenessAbnormal")
        
            Exit Sub
        Else

            sPutINISetting "Software\UBoatStopWatch", "dpiAwareness", gblDpiAwareness, gblSettingsFile
            'Call reloadProgram ' this is insufficient, image controls still fail to resize and autoscale correctly
            Call hardRestart
        End If

    End If

   On Error GoTo 0
   Exit Sub

chkDpiAwareness_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDpiAwareness_Click of Form widgetPrefs"
End Sub







'---------------------------------------------------------------------------------------
' Procedure : chkShowTaskbar_Click
' Author    : beededea
' Date      : 19/07/2023
' Purpose   : toggle for showing the program in the taskbar
'---------------------------------------------------------------------------------------
'
Private Sub chkShowTaskbar_Click()

   On Error GoTo chkShowTaskbar_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkShowTaskbar.Value = 1 Then
        gblShowTaskbar = "1"
    Else
        gblShowTaskbar = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowTaskbar_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_Click
' Author    : beededea
' Date      : 01/10/2023
' Purpose   : reset the improved message boxes so that any hidden boxes will reappear
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_Click()

    On Error GoTo btnResetMessages_Click_Error
        
    ' Clear all the message box "show again" entries in the registry
    Call clearAllMessageBoxRegistryEntries
    
    MsgBox "Message boxes fully reset, confirmation pop-ups will continue as normal."

    On Error GoTo 0
    Exit Sub

btnResetMessages_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   : Enabling debug mode - not implemented
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   'If gblDebugFlg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    'mnuDebug_Click
    MsgBox "The debug mode is not yet enabled."

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : Donate button
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuCoffee_ClickEvent

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : FB button
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   'If gblDebugFlg = 1 Then DebugPrint "%btnFacebook_Click"

    Call menuForm.mnuFacebook_Click
    

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnOpenFile_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : button for opening a target file for dblClicking
'---------------------------------------------------------------------------------------
'
Private Sub btnOpenFile_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnOpenFile_Click_Error

    Call addTargetFile(txtOpenFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtOpenFile.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        'answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Create file confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnOpenFile_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnOpenFile_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : auto update button
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   'If gblDebugFlg = 1 Then DebugPrint "%btnUpdate_Click"

    'MsgBox "The update button is not yet enabled."
    menuForm.mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : chkGenStartup_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : Toggle automatic startup by writing to the registry on save
'---------------------------------------------------------------------------------------
'
Private Sub chkGenStartup_Click()
    On Error GoTo chkGenStartup_Click_Error

    btnSave.Enabled = True ' enable the save button

    On Error GoTo 0
    Exit Sub

chkGenStartup_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenStartup_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDefaultEditor_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : button for selecting the default VB6 editor VBP project file
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaultEditor_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnDefaultEditor_Click_Error

    Call addTargetFile(txtDefaultEditor.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtDefaultEditor.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Default Editor Confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnDefaultEditor_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaultEditor_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub




'---------------------------------------------------------------------------------------
' Procedure : chkIgnoreMouse_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : toggle to ignore any mouse clicks
'---------------------------------------------------------------------------------------
'
Private Sub chkIgnoreMouse_Click()
   On Error GoTo chkIgnoreMouse_Click_Error

    If chkIgnoreMouse.Value = 0 Then
        gblIgnoreMouse = "0"
    Else
        gblIgnoreMouse = "1"
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIgnoreMouse_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkPreventDragging_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : lock the program in place, prevent dragging
'---------------------------------------------------------------------------------------
'
Private Sub chkPreventDragging_Click()
    On Error GoTo chkPreventDragging_Click_Error

    btnSave.Enabled = True ' enable the save button
    ' immediately make the widget locked in place
    If chkPreventDragging.Value = 0 Then
        overlayWidget.Locked = False
        gblPreventDragging = "0"
        menuForm.mnuLockWidget.Checked = False
        If gblAspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = vbNullString
            txtLandscapeVoffset.Text = vbNullString
        Else
            txtPortraitHoffset.Text = vbNullString
            txtPortraitYoffset.Text = vbNullString
        End If
    Else
        overlayWidget.Locked = True
        gblPreventDragging = "1"
        menuForm.mnuLockWidget.Checked = True
        If gblAspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = fGauge.gaugeForm.Left
            txtLandscapeVoffset.Text = fGauge.gaugeForm.Top
        Else
            txtPortraitHoffset.Text = fGauge.gaugeForm.Left
            txtPortraitYoffset.Text = fGauge.gaugeForm.Top
        End If
    End If

    On Error GoTo 0
    Exit Sub

chkPreventDragging_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkPreventDragging_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkWidgetHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : toggle to hide the program
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetHidden_Click()
   On Error GoTo chkWidgetHidden_Click_Error

    If chkWidgetHidden.Value = 0 Then
        'overlayWidget.Hidden = False
        fGauge.gaugeForm.Visible = True

        frmTimer.revealWidgetTimer.Enabled = False
        gblWidgetHidden = "0"
    Else
        'overlayWidget.Hidden = True
        fGauge.gaugeForm.Visible = False


        frmTimer.revealWidgetTimer.Enabled = True
        gblWidgetHidden = "1"
    End If
    
    sPutINISetting "Software\UBoatStopWatch", "widgetHidden", gblWidgetHidden, gblSettingsFile
    
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkWidgetHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetHidden_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAspectHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : selector for hiding in portrait/landscape mode
'---------------------------------------------------------------------------------------
'
Private Sub cmbAspectHidden_Click()

   On Error GoTo cmbAspectHidden_Click_Error

    If cmbAspectHidden.ListIndex = 1 And gblAspectRatio = "portrait" Then
        'overlayWidget.Hidden = True
        fGauge.gaugeForm.Visible = False
    ElseIf cmbAspectHidden.ListIndex = 2 And gblAspectRatio = "landscape" Then
        'overlayWidget.Hidden = True
        fGauge.gaugeForm.Visible = False
    Else
        'overlayWidget.Hidden = False
        fGauge.gaugeForm.Visible = True
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAspectHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAspectHidden_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDebug_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : debug selector
'---------------------------------------------------------------------------------------
'
Private Sub cmbDebug_Click()
    On Error GoTo cmbDebug_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbDebug.ListIndex = 0 Then
        txtDefaultEditor.Text = "eg. E:\vb6\UBoat-StopWatch-VB6MkII\UBoat-StopWatch-VB6.vbp"
        txtDefaultEditor.Enabled = False
        lblDebug(7).Enabled = False
        btnDefaultEditor.Enabled = False
        lblDebug(9).Enabled = False
    Else
        #If TWINBASIC Then
            txtDefaultEditor.Text = gblDefaultTBEditor
        #Else
            txtDefaultEditor.Text = gblDefaultVB6Editor
        #End If
        txtDefaultEditor.Enabled = True
        lblDebug(7).Enabled = True
        btnDefaultEditor.Enabled = True
        lblDebug(9).Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbDebug_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDebug_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmbHidingTime_Click
' Author    : beededea
' Date      : 17/02/2025
' Purpose   : enable the save button if a hiding time is selected
'---------------------------------------------------------------------------------------
'
Private Sub cmbHidingTime_Click()
   On Error GoTo cmbHidingTime_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbHidingTime_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbHidingTime_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbScrollWheelDirection_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : selector for resizing using the mouse Scroll Wheel Direction
'---------------------------------------------------------------------------------------
'
Private Sub cmbScrollWheelDirection_Click()
   On Error GoTo cmbScrollWheelDirection_Click_Error

    btnSave.Enabled = True ' enable the save button
    'overlayWidget.ZoomDirection = cmbScrollWheelDirection.List(cmbScrollWheelDirection.ListIndex)

   On Error GoTo 0
   Exit Sub

cmbScrollWheelDirection_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbScrollWheelDirection_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetLandscape_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : option dropdown for locking in landscape mode after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetLandscape_Click()
   On Error GoTo cmbWidgetLandscape_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbWidgetLandscape_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetLandscape_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPortrait_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : option dropdown for locking in portrait mode after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPortrait_Click()
   On Error GoTo cmbWidgetPortrait_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbWidgetPortrait_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPortrait_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPosition_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : option dropdown to position by percent after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPosition_Click()
    On Error GoTo cmbWidgetPosition_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWidgetPosition.ListIndex = 1 Then
        cmbWidgetLandscape.ListIndex = 0
        cmbWidgetPortrait.ListIndex = 0
        cmbWidgetLandscape.Enabled = False
        cmbWidgetPortrait.Enabled = False
        txtLandscapeHoffset.Enabled = False
        txtLandscapeVoffset.Enabled = False
        txtPortraitHoffset.Enabled = False
        txtPortraitYoffset.Enabled = False
        
    Else
        cmbWidgetLandscape.Enabled = True
        cmbWidgetPortrait.Enabled = True
        txtLandscapeHoffset.Enabled = True
        txtLandscapeVoffset.Enabled = True
        txtPortraitHoffset.Enabled = True
        txtPortraitYoffset.Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbWidgetPosition_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPosition_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
 
   On Error GoTo IsLoaded_Error

    IsLoaded = mIsLoaded

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Let IsLoaded(ByVal newValue As Boolean)
 
   On Error GoTo IsLoaded_Error

   If mIsLoaded <> newValue Then mIsLoaded = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
' Purpose   : calling a manual property  by val to a form in this manual property allows external checks to the form to
'             determine whether it is loaded, without also activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If IsLoaded = True Then
        If Me.WindowState = vbNormal Then
            IsVisible = Me.Visible
        Else
            IsVisible = False
        End If
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
            Resume Next
          End If
    End With
End Property

''---------------------------------------------------------------------------------------
'' Procedure : IsVisible
'' Author    : beededea
'' Date      : 16/12/2024
'' Purpose   : property by val to manually determine whether the preference form is visible. It does this without
''             touching any VB6 intrinsic form property which would then load the form itself.
''---------------------------------------------------------------------------------------
''
'Public Property Let IsVisible(ByVal newValue As Boolean)
'
'   On Error GoTo IsVisible_Error
'
'   If mIsVisible <> newValue Then mIsVisible = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'IsVisible_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
'
'End Property

'---------------------------------------------------------------------------------------
' Procedure : showLastTab
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : make the last used tab appear on startup
'---------------------------------------------------------------------------------------
'
Private Sub showLastTab()

   On Error GoTo showLastTab_Error

    If gblLastSelectedTab = "general" Then Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton)  ' was imgGeneralMouseUpEvent
    If gblLastSelectedTab = "config" Then Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton)     ' was imgConfigMouseUpEvent
    If gblLastSelectedTab = "position" Then Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    If gblLastSelectedTab = "development" Then Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    If gblLastSelectedTab = "fonts" Then Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    If gblLastSelectedTab = "sounds" Then Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    If gblLastSelectedTab = "window" Then Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    If gblLastSelectedTab = "about" Then Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)

    '' fGauge.RotateBusyTimer = True
    
  On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLastTab of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : positionPrefsFramesButtons
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : size and position the frames and buttons. Note we are NOT using control
'             arrays so the form can be converted to Cairo forms later.
'---------------------------------------------------------------------------------------
'
Private Sub positionPrefsFramesButtons()
    On Error GoTo positionPrefsFramesButtons_Error

    Dim frameWidth As Integer: frameWidth = 0
    Dim frameTop As Integer: frameTop = 0
    Dim frameLeft As Integer: frameLeft = 0
    Dim buttonTop As Integer:    buttonTop = 0
    'Dim currentFrameHeight As Integer: currentFrameHeight = 0
    Dim rightHandAlignment As Long: rightHandAlignment = 0
    Dim leftHandGutterWidth As Long: leftHandGutterWidth = 0
    
    ' align frames rightmost and leftmost to the buttons at the top
    buttonTop = -15
    frameTop = 1150
    leftHandGutterWidth = 240
    frameLeft = leftHandGutterWidth ' use the first frame leftmost as reference
    rightHandAlignment = fraAboutButton.Left + fraAboutButton.Width ' use final button rightmost as reference
    frameWidth = rightHandAlignment - frameLeft
    fraScrollbarCover.Left = rightHandAlignment - 690
    widgetPrefs.Width = rightHandAlignment + leftHandGutterWidth + 75 ' (not quite sure why we need the 75 twips padding) ' this triggers a resize '
    
    ' 9053 start
    
    ' align the top buttons
    fraGeneralButton.Top = buttonTop
    fraConfigButton.Top = buttonTop
    fraFontsButton.Top = buttonTop
    fraSoundsButton.Top = buttonTop
    fraPositionButton.Top = buttonTop
    fraDevelopmentButton.Top = buttonTop
    fraWindowButton.Top = buttonTop
    fraAboutButton.Top = buttonTop
    
    ' align the frames
    fraGeneral.Top = frameTop
    fraConfig.Top = frameTop
    fraFonts.Top = frameTop
    fraSounds.Top = frameTop
    fraPosition.Top = frameTop
    fraDevelopment.Top = frameTop
    fraWindow.Top = frameTop
    fraAbout.Top = frameTop
    
    fraGeneral.Left = frameLeft
    fraConfig.Left = frameLeft
    fraSounds.Left = frameLeft
    fraPosition.Left = frameLeft
    fraFonts.Left = frameLeft
    fraDevelopment.Left = frameLeft
    fraWindow.Left = frameLeft
    fraAbout.Left = frameLeft
    
    fraGeneral.Width = frameWidth
    fraConfig.Width = frameWidth
    fraSounds.Width = frameWidth
    fraPosition.Width = frameWidth
    fraFonts.Width = frameWidth
    fraWindow.Width = frameWidth
    fraDevelopment.Width = frameWidth
    fraAbout.Width = frameWidth
    
    ' set the base visibility of the frames
    fraGeneral.Visible = True
    fraConfig.Visible = False
    fraSounds.Visible = False
    fraPosition.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraDevelopment.Visible = False
    fraAbout.Visible = False
            
    fraGeneralButton.BorderStyle = 1
    
    #If TWINBASIC Then
        fraGeneralButton.Refresh
    #End If

    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - 50
    btnHelp.Left = frameLeft
    
    '' fGauge.RotateBusyTimer = True

   On Error GoTo 0
   Exit Sub

positionPrefsFramesButtons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsFramesButtons of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to close the prefs form
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    btnSave.Enabled = False ' disable the save button
    Me.Hide
    Me.themeTimer.Enabled = False
    
    Call writePrefsPositionAndSize
    
    Call adjustPrefsControls(True)

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : display the help file
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
    On Error GoTo btnHelp_Click_Error
    
        If fFExists(App.path & "\help\Help.chm") Then
            Call ShellExecute(Me.hWnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BCStr
' Author    : beededea
' Date      : 31/05/2025
' Purpose   : This replacement of CStr is ONLY for non-numeric boolean casts to a string.
'             NOTE: boolean values can be locale-sensitive when converted by CStr returning a local language result
'             It might be best to stick with LTrim$(Str$()) for the moment.
'             It is OK to convert checkboxes using cstr() as the boolean values are stored as 0 and -1
'---------------------------------------------------------------------------------------
'
Function BCStr(ByVal booleanValue As Boolean) As String
   On Error GoTo BCStr_Error

    If booleanValue Then
        BCStr = "True"
    Else
        BCStr = "False"
    End If
    
   On Error GoTo 0
   Exit Function

BCStr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BCStr of Form widgetPrefs"
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : save the values from all the tabs
'             NOTE: boolean values can be locale-sensitive when converted by CStr returning a local language result
'             It might be best to stick with LTrim$(Str$()) for the moment.
'             It is OK to convert checkboxes using cstr() as the boolean values are stored as 0 and -1
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()

    
    On Error GoTo btnSave_Click_Error

    ' configuration
    gblGaugeTooltips = CStr(optGaugeTooltips(0).Tag)
    gblPrefsTooltips = CStr(optPrefsTooltips(0).Tag)
    
    gblShowTaskbar = CStr(chkShowTaskbar.Value)
    gblShowHelp = CStr(chkShowHelp.Value)
    
'    gblTogglePendulum = CStr(chkTogglePendulum.Value)
'    gbl24HourGaugeMode = CStr(chk24HourGaugeMode.Value)
    
    gblDpiAwareness = CStr(chkDpiAwareness.Value)
    gblGaugeSize = CStr(sliGaugeSize.Value)
    gblScrollWheelDirection = CStr(cmbScrollWheelDirection.ListIndex)
        
    ' general
    gblGaugeFunctions = CStr(chkGaugeFunctions.Value)
    gblStartup = CStr(chkGenStartup.Value)
    
    gblClockFaceSwitchPref = cmbClockFaceSwitchPref.ListIndex
    
    gblMainGaugeTimeZone = cmbMainGaugeTimeZone.ListIndex
    gblMainDaylightSaving = cmbMainDaylightSaving.ListIndex
    
    gblSmoothSecondHand = cmbTickSwitchPref.ListIndex
    
    gblSecondaryGaugeTimeZone = cmbSecondaryGaugeTimeZone.ListIndex
    gblSecondaryDaylightSaving = cmbSecondaryDaylightSaving.ListIndex
    
'
    ' sounds
    gblEnableSounds = CStr(chkEnableSounds.Value)
'    gblEnableTicks = CStr(chkEnableTicks.Value)
'    gblEnableChimes = CStr(chkEnableChimes.Value)
'    gblEnableAlarms = CStr(chkEnableAlarms.Value)
'    gblVolumeBoost = CStr(chkVolumeBoost.Value)
    
    'development
    gblDebug = CStr(cmbDebug.ListIndex)
    gblDblClickCommand = txtDblClickCommand.Text
    gblOpenFile = txtOpenFile.Text
    #If TWINBASIC Then
        gblDefaultTBEditor = txtDefaultEditor.Text
    #Else
        gblDefaultVB6Editor = txtDefaultEditor.Text
    #End If
    
    ' position
    gblAspectHidden = CStr(cmbAspectHidden.ListIndex)
    gblWidgetPosition = CStr(cmbWidgetPosition.ListIndex)
    gblWidgetLandscape = CStr(cmbWidgetLandscape.ListIndex)
    gblWidgetPortrait = CStr(cmbWidgetPortrait.ListIndex)
    gblLandscapeFormHoffset = txtLandscapeHoffset.Text
    gblLandscapeFormVoffset = txtLandscapeVoffset.Text
    gblPortraitHoffset = txtPortraitHoffset.Text
    gblPortraitYoffset = txtPortraitYoffset.Text
    
'    gblvLocationPercPrefValue
'    gblhLocationPercPrefValue

    ' fonts
    gblPrefsFont = txtPrefsFont.Text
    gblGaugeFont = gblPrefsFont
    
    gblDisplayScreenFont = txtDisplayScreenFont.Text
    gblDisplayScreenFontSize = txtDisplayScreenFontSize.Text
    
'    gblDisplayScreenFontSize
'    gblDisplayScreenFontItalics
'    gblDisplayScreenFontColour

    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If gblDpiAwareness = "1" Then
'        gblPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        gblPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'gblPrefsFontItalics = txtFontSize.Text

    ' Windows
    gblWindowLevel = CStr(cmbWindowLevel.ListIndex)
    gblPreventDragging = CStr(chkPreventDragging.Value)
    gblOpacity = CStr(sliOpacity.Value)
    gblWidgetHidden = CStr(chkWidgetHidden.Value)
    gblHidingTime = CStr(cmbHidingTime.ListIndex)
    gblIgnoreMouse = CStr(chkIgnoreMouse.Value)
    
    gblMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
     
            
    If gblStartup = "1" Then
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "UBoatStopWatch", """" & App.path & "\" & "UBoat-StopWatch-VB6.exe""")
    Else
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "UBoatStopWatch", vbNullString)
    End If

    ' save the values from the general tab
    If fFExists(gblSettingsFile) Then
        sPutINISetting "Software\UBoatStopWatch", "gaugeTooltips", gblGaugeTooltips, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsTooltips", gblPrefsTooltips, gblSettingsFile

        sPutINISetting "Software\UBoatStopWatch", "showTaskbar", gblShowTaskbar, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "showHelp", gblShowHelp, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        
        sPutINISetting "Software\UBoatStopWatch", "gaugeSize", gblGaugeSize, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "scrollWheelDirection", gblScrollWheelDirection, gblSettingsFile
                
        sPutINISetting "Software\UBoatStopWatch", "widgetFunctions", gblGaugeFunctions, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "smoothSecondHand", gblSmoothSecondHand, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "clockFaceSwitchPref", gblClockFaceSwitchPref, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "mainGaugeTimeZone", gblMainGaugeTimeZone, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "mainDaylightSaving", gblMainDaylightSaving, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "secondaryGaugeTimeZone", gblSecondaryGaugeTimeZone, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "secondaryDaylightSaving", gblSecondaryDaylightSaving, gblSettingsFile
        
              
        sPutINISetting "Software\UBoatStopWatch", "aspectHidden", gblAspectHidden, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "widgetPosition", gblWidgetPosition, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "widgetLandscape", gblWidgetLandscape, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "widgetPortrait", gblWidgetPortrait, gblSettingsFile

        sPutINISetting "Software\UBoatStopWatch", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "gaugeFont", gblGaugeFont, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontColour", gblPrefsFontColour, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFont", gblDisplayScreenFont, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontSize", gblDisplayScreenFontSize, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontItalics", gblDisplayScreenFontItalics, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontColour", gblDisplayScreenFontColour, gblSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\UBoatStopWatch", "windowLevel", gblWindowLevel, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "preventDragging", gblPreventDragging, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "opacity", gblOpacity, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "widgetHidden", gblWidgetHidden, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "hidingTime", gblHidingTime, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "ignoreMouse", gblIgnoreMouse, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "multiMonitorResize", gblMultiMonitorResize, gblSettingsFile
        
        
        sPutINISetting "Software\UBoatStopWatch", "startup", gblStartup, gblSettingsFile

        sPutINISetting "Software\UBoatStopWatch", "enableSounds", gblEnableSounds, gblSettingsFile
'        sPutINISetting "Software\UBoatStopWatch", "enableTicks", gblEnableTicks, gblSettingsFile
'        sPutINISetting "Software\UBoatStopWatch", "enableChimes", gblEnableChimes, gblSettingsFile
'        sPutINISetting "Software\UBoatStopWatch", "enableAlarms", gblEnableAlarms, gblSettingsFile
'        sPutINISetting "Software\UBoatStopWatch", "volumeBoost", gblVolumeBoost, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "debug", gblDebug, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "dblClickCommand", gblDblClickCommand, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "openFile", gblOpenFile, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "defaultVB6Editor", gblDefaultVB6Editor, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "defaultTBEditor", gblDefaultTBEditor, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "gaugeHighDpiXPos", gblGaugeHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "gaugeHighDpiYPos", gblGaugeHighDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\UBoatStopWatch", "gaugeLowDpiXPos", gblGaugeLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "gaugeLowDpiYPos", gblGaugeLowDpiYPos, gblSettingsFile
                       
    End If
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

    ' sets the characteristics of the gauge and menus immediately after saving
    Call adjustMainControls(1)
    
    Me.SetFocus
    btnSave.Enabled = False ' disable the save button showing it has successfully saved
    
    ' reload here if the gblWindowLevel Was Changed
    If gblWindowLevelWasChanged = True Then
        gblWindowLevelWasChanged = False
        Call reloadProgram
    End If
    
   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkEnableSounds_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : toggle to enable/disable sounds on save
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableSounds_Click()
   On Error GoTo chkEnableSounds_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableSounds_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableSounds_Click of Form widgetPrefs"
End Sub


' ----------------------------------------------------------------
' Procedure Name: cmbWindowLevel_Click
' Purpose: option to determine the windows Z order of the main program (not the prefs form)
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/05/2024
' ----------------------------------------------------------------
Private Sub cmbWindowLevel_Click()
    On Error GoTo cmbWindowLevel_Click_Error
    btnSave.Enabled = True ' enable the save button
    If pvtPrefsStartupFlg = False Then gblWindowLevelWasChanged = True
    
    On Error GoTo 0
    Exit Sub

cmbWindowLevel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWindowLevel_Click, line " & Erl & "."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnPrefsFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to select the font dialog
'---------------------------------------------------------------------------------------
'
Private Sub btnPrefsFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnPrefsFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gblPrefsFont
    ' gblGaugeFont
    
    If gblDpiAwareness = "1" Then
        fntSize = Val(gblPrefsFontSizeHighDPI)
    Else
        fntSize = Val(gblPrefsFontSizeLowDPI)
    End If
    
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(gblPrefsFontItalics)
    fntColour = CLng(gblPrefsFontColour)
        
    Call changeFont(widgetPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    
    gblPrefsFont = CStr(fntFont)
    gblGaugeFont = gblPrefsFont
    
    If gblDpiAwareness = "1" Then
        gblPrefsFontSizeHighDPI = CStr(fntSize)
        Call Form_Resize
    Else
        gblPrefsFontSizeLowDPI = CStr(fntSize)
    End If
    
    gblPrefsFontItalics = CStr(fntItalics)
    gblPrefsFontColour = CStr(fntColour)
    
    ' changes the displayed font to an adjusted base font size after a resize
    Call PrefsForm_Resize_Event

    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\UBoatStopWatch", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "gaugeFont", gblGaugeFont, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "PrefsFontColour", gblPrefsFontColour, gblSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "arial"
    txtPrefsFont.Text = fntFont
    txtPrefsFont.Font.Name = fntFont
    'txtPrefsFont.Font.Size = fntSize
    txtPrefsFont.Font.Italic = fntItalics
    txtPrefsFont.ForeColor = fntColour
    
    txtPrefsFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnPrefsFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrefsFont_Click of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnDisplayScreenFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to select the font dialog for the display console
'---------------------------------------------------------------------------------------
'
Private Sub btnDisplayScreenFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnDisplayScreenFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gblDisplayScreenFont
    
    fntSize = Val(gblDisplayScreenFontSize)
    If fntSize = 0 Then fntSize = 5
    fntItalics = CBool(gblDisplayScreenFontItalics)
    fntColour = CLng(gblDisplayScreenFontColour)
    
    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub
            
    gblDisplayScreenFont = CStr(fntFont)
    gblDisplayScreenFontSize = CStr(fntSize)
    gblDisplayScreenFontItalics = CStr(fntItalics)
    gblDisplayScreenFontColour = CStr(fntColour)
    
    If gblFGaugeAvailable = True Then
        fGauge.gaugeForm.Widgets("lblTerminalText").Widget.FontSize = gblDisplayScreenFontSize
        fGauge.gaugeForm.Widgets("lblTerminalText").Widget.FontName = gblDisplayScreenFont
    End If

    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFont", gblDisplayScreenFont, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontSize", gblDisplayScreenFontSize, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontItalics", gblDisplayScreenFontItalics, gblSettingsFile
        sPutINISetting "Software\UBoatStopWatch", "displayScreenFontColour", gblDisplayScreenFontColour, gblSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "courier new"
    txtDisplayScreenFont.Text = fntFont
    txtDisplayScreenFont.Font.Name = fntFont
    txtDisplayScreenFont.Font.Italic = fntItalics
    txtDisplayScreenFont.ForeColor = fntColour
    txtDisplayScreenFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnDisplayScreenFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDisplayScreenFont_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls(Optional ByVal restartState As Boolean)
    
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim sliGaugeSizeOldValue As Long: sliGaugeSizeOldValue = 0
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo adjustPrefsControls_Error
    
    ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
    'prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
    
    'widgetPrefs.Height = CLng(gblPrefsPrimaryHeightTwips)
            
    ' general tab
    chkGaugeFunctions.Value = Val(gblGaugeFunctions)
    chkGenStartup.Value = Val(gblStartup)
            
    cmbClockFaceSwitchPref.ListIndex = Val(gblClockFaceSwitchPref)
    
    'set the choice for four timezone comboboxes that were populated from file.
    cmbMainGaugeTimeZone.ListIndex = Val(gblMainGaugeTimeZone)
    cmbMainDaylightSaving.ListIndex = Val(gblMainDaylightSaving)
'
    txtMainBias.Text = tzDelta
'
    cmbTickSwitchPref.ListIndex = Val(gblSmoothSecondHand)
'
    cmbSecondaryGaugeTimeZone.ListIndex = Val(gblSecondaryGaugeTimeZone)
    
    cmbSecondaryDaylightSaving.ListIndex = Val(gblSecondaryDaylightSaving)
    ' configuration tab
   
    ' ' fGauge.RotateBusyTimer = True
    
    ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliGaugeSizeOldValue = sliGaugeSize.Value
    sliGaugeSize.Value = Val(gblGaugeSize)
    If sliGaugeSize.Value <> sliGaugeSizeOldValue Then
        btnSave.Visible = True
    End If
        
    cmbScrollWheelDirection.ListIndex = Val(gblScrollWheelDirection)
    
    optGaugeTooltips(CStr(gblGaugeTooltips)).Value = True
    optGaugeTooltips(0).Tag = CStr(gblGaugeTooltips)
    optGaugeTooltips(1).Tag = CStr(gblGaugeTooltips)
    optGaugeTooltips(2).Tag = CStr(gblGaugeTooltips)
        
    optPrefsTooltips(CStr(gblPrefsTooltips)).Value = True
    optPrefsTooltips(0).Tag = CStr(gblPrefsTooltips)
    optPrefsTooltips(1).Tag = CStr(gblPrefsTooltips)
    optPrefsTooltips(2).Tag = CStr(gblPrefsTooltips)
    
    chkShowTaskbar.Value = Val(gblShowTaskbar)
    chkShowHelp.Value = Val(gblShowHelp)
'    chkTogglePendulum.Value = Val(gblTogglePendulum)
'    chk24HourGaugeMode.Value = Val(gbl24HourGaugeMode)
    
    chkDpiAwareness.Value = Val(gblDpiAwareness)
    'chkNumericDisplayRotation.Value = Val(gblNumericDisplayRotation)
        
    ' sounds tab
    ' ' fGauge.RotateBusyTimer = True

    chkEnableSounds.Value = Val(gblEnableSounds)
'    chkEnableTicks.Value = Val(gblEnableTicks)
'    chkEnableChimes.Value = Val(gblEnableChimes)
'    chkEnableAlarms.Value = Val(gblEnableAlarms)
'    chkVolumeBoost.Value = Val(gblVolumeBoost)
    
    
    ' development
    ' ' fGauge.RotateBusyTimer = True
    
    cmbDebug.ListIndex = Val(gblDebug)
    txtDblClickCommand.Text = gblDblClickCommand
    txtOpenFile.Text = gblOpenFile
    #If TWINBASIC Then
        txtDefaultEditor.Text = gblDefaultTBEditor
    #Else
        txtDefaultEditor.Text = gblDefaultVB6Editor
    #End If
    
    lblGitHub.Caption = "You can find the code for the UBoat StopWatch on github, visit by double-clicking this link https://github.com/yereverluvinunclebert/ UBoat-StopWatch"
     
     
     If Not restartState = True Then
        ' fonts tab
        If gblPrefsFont <> vbNullString Then
            txtPrefsFont.Text = gblPrefsFont
            If gblDpiAwareness = "1" Then
                Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeHighDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
                txtPrefsFontSize.Text = gblPrefsFontSizeHighDPI
            Else
                Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeLowDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
                txtPrefsFontSize.Text = gblPrefsFontSizeLowDPI
            End If
        End If
        
        txtDisplayScreenFontSize.Text = gblDisplayScreenFontSize
    
        txtDisplayScreenFont.Font.Name = gblDisplayScreenFont
        'txtDisplayScreenFont.Font.Size = Val(gblDisplayScreenFont)
    End If
    
    ' position tab
    ' ' fGauge.RotateBusyTimer = True
    
    cmbAspectHidden.ListIndex = Val(gblAspectHidden)
    cmbWidgetPosition.ListIndex = Val(gblWidgetPosition)
        
    If gblPreventDragging = "1" Then
        If gblAspectRatio = "landscape" Then
'            txtLandscapeHoffset.Text = fGauge.gaugeForm.Left
'            txtLandscapeVoffset.Text = fGauge.gaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblGaugeHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblGaugeHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblGaugeLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblGaugeLowDpiYPos & "px"
            End If
        Else
'            txtPortraitHoffset.Text = fGauge.gaugeForm.Left
'            txtPortraitYoffset.Text = fGauge.gaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblGaugeHighDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblGaugeHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblGaugeLowDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblGaugeLowDpiYPos & "px"
            End If
        End If
    End If
    
    'cmbWidgetLandscape
    ' ' fGauge.RotateBusyTimer = True
    
    cmbWidgetLandscape.ListIndex = Val(gblWidgetLandscape)
    cmbWidgetPortrait.ListIndex = Val(gblWidgetPortrait)
    txtLandscapeHoffset.Text = gblLandscapeFormHoffset
    txtLandscapeVoffset.Text = gblLandscapeFormVoffset
    txtPortraitHoffset.Text = gblPortraitHoffset
    txtPortraitYoffset.Text = gblPortraitYoffset

    ' Windows tab
    ' ' fGauge.RotateBusyTimer = True
    
    cmbWindowLevel.ListIndex = Val(gblWindowLevel)
    chkIgnoreMouse.Value = Val(gblIgnoreMouse)
    chkPreventDragging.Value = Val(gblPreventDragging)
    sliOpacity.Value = Val(gblOpacity)
    chkWidgetHidden.Value = Val(gblWidgetHidden)
    cmbHidingTime.ListIndex = Val(gblHidingTime)
    cmbMultiMonitorResize.ListIndex = Val(gblMultiMonitorResize)
    
    If gblMonitorCount > 1 Then
        cmbMultiMonitorResize.Visible = True
        lblWindowLevel(10).Visible = True
        lblWindowLevel(11).Visible = True
    Else
        cmbMultiMonitorResize.Visible = False
        lblWindowLevel(10).Visible = False
        lblWindowLevel(11).Visible = False
    End If
    
    ' ' fGauge.RotateBusyTimer = True
    
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form widgetPrefs on line " & Erl

End Sub

'---------------------------------------------------------------------------------------

'
'---------------------------------------------------------------------------------------
' Procedure : populatePrefsComboBoxes
' Author    : beededea
' Date      : 10/09/2022
' Purpose   : all combo boxes in the prefs are populated here with default values
'           : done by preference here rather than in the IDE
'---------------------------------------------------------------------------------------

Private Sub populatePrefsComboBoxes()

    Dim useloop As Integer: useloop = 0
    Dim minString As String: minString = vbNullString
    
    On Error GoTo populatePrefsComboBoxes_Error
    
'    With fGauge.gaugeForm.Widgets("busy1").Widget
'        .Alpha = 1
'        .Refresh
'    End With
    
    cmbScrollWheelDirection.AddItem "up", 0
    cmbScrollWheelDirection.ItemData(0) = 0
    cmbScrollWheelDirection.AddItem "down", 1
    cmbScrollWheelDirection.ItemData(1) = 1
    
    ' ' fGauge.RotateBusyTimer = True
    
    cmbAspectHidden.AddItem "none", 0
    cmbAspectHidden.ItemData(0) = 0
    cmbAspectHidden.AddItem "portrait", 1
    cmbAspectHidden.ItemData(1) = 1
    cmbAspectHidden.AddItem "landscape", 2
    cmbAspectHidden.ItemData(2) = 2

    ' ' fGauge.RotateBusyTimer = True
    
    cmbWidgetPosition.AddItem "disabled", 0
    cmbWidgetPosition.ItemData(0) = 0
    cmbWidgetPosition.AddItem "enabled", 1
    cmbWidgetPosition.ItemData(1) = 1
    
    ' ' fGauge.RotateBusyTimer = True
    
    cmbWidgetLandscape.AddItem "disabled", 0
    cmbWidgetLandscape.ItemData(0) = 0
    cmbWidgetLandscape.AddItem "enabled", 1
    cmbWidgetLandscape.ItemData(1) = 1
    
    cmbWidgetPortrait.AddItem "disabled", 0
    cmbWidgetPortrait.ItemData(0) = 0
    cmbWidgetPortrait.AddItem "enabled", 1
    cmbWidgetPortrait.ItemData(1) = 1
    
    cmbDebug.AddItem "Debug OFF", 0
    cmbDebug.ItemData(0) = 0
    cmbDebug.AddItem "Debug ON", 1
    cmbDebug.ItemData(1) = 1
    
    ' ' fGauge.RotateBusyTimer = True
    
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.ItemData(0) = 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.ItemData(1) = 1
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    cmbWindowLevel.ItemData(2) = 2

    ' ' fGauge.RotateBusyTimer = True
    
    ' populate the hiding timer combobox
    cmbHidingTime.AddItem "1 minute", 0
    cmbHidingTime.ItemData(0) = 1
    cmbHidingTime.AddItem "5 minutes", 1
    cmbHidingTime.ItemData(1) = 5
    cmbHidingTime.AddItem "10 minutes", 2
    cmbHidingTime.ItemData(2) = 10
    cmbHidingTime.AddItem "20 minutes", 3
    cmbHidingTime.ItemData(3) = 20
    cmbHidingTime.AddItem "30 minutes", 4
    cmbHidingTime.ItemData(4) = 30
    cmbHidingTime.AddItem "I hour", 5
    cmbHidingTime.ItemData(5) = 60
    
    ' ' fGauge.RotateBusyTimer = True
    
    ' populate the multi monitor combobox
    cmbMultiMonitorResize.AddItem "Disabled", 0
    cmbMultiMonitorResize.ItemData(0) = 0
    cmbMultiMonitorResize.AddItem "Automatic Resizing Enabled", 1
    cmbMultiMonitorResize.ItemData(1) = 1
    cmbMultiMonitorResize.AddItem "Manual Sizing Stored Per Monitor", 2
    cmbMultiMonitorResize.ItemData(2) = 2
    
    ' populate the clock face to show
    cmbClockFaceSwitchPref.AddItem "stopwatch", 0
    cmbClockFaceSwitchPref.AddItem "standard", 1
 
    'populate one timezone combobox from file.
    Call readFileWriteComboBox(cmbMainGaugeTimeZone, App.path & "\Resources\txt\timezones.txt")
    
    'populate one timezone combobox from file.
    Call readFileWriteComboBox(cmbMainDaylightSaving, App.path & "\Resources\txt\DLScodesWin.txt")

    'populate one timezone combobox from file.
    Call readFileWriteComboBox(cmbSecondaryGaugeTimeZone, App.path & "\Resources\txt\timezones.txt")
    
    'populate one timezone combobox from file.
    Call readFileWriteComboBox(cmbSecondaryDaylightSaving, App.path & "\Resources\txt\DLScodesWin.txt")

    cmbTickSwitchPref.AddItem "Tick", 0
    cmbTickSwitchPref.ItemData(0) = 0
    cmbTickSwitchPref.AddItem "Smooth", 1
    cmbTickSwitchPref.ItemData(1) = 1

    
    On Error GoTo 0
    Exit Sub

populatePrefsComboBoxes_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populatePrefsComboBoxes of Form widgetPrefs"
            Resume Next
          End If
    End With
                
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : readFileWriteComboBox
' Author    : beededea
' Date      : 28/07/2023
' Purpose   : Open and load the Array with the timezones text File
'---------------------------------------------------------------------------------------
'
Private Sub readFileWriteComboBox(ByRef thisComboBox As Control, ByVal thisFileName As String)
    Dim strArr() As String
    Dim lngCount As Long: lngCount = 0
    Dim lngIdx As Long: lngIdx = 0
        
    On Error GoTo readFileWriteComboBox_Error

    If fFExists(thisFileName) = True Then
       ' the files must be DOS CRLF delineated
       Open thisFileName For Input As #1
           strArr() = Split(Input(LOF(1), 1), vbCrLf)
       Close #1
    
       lngCount = UBound(strArr)
    
       thisComboBox.Clear
       For lngIdx = 0 To lngCount
           thisComboBox.AddItem strArr(lngIdx)
       Next lngIdx
    End If

   On Error GoTo 0
   Exit Sub

readFileWriteComboBox_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readFileWriteComboBox of Form widgetPrefs"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : clearBorderStyle
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : removes all styling from the icon frames and makes the major frames below invisible too, not using control arrays.
'---------------------------------------------------------------------------------------
'
Private Sub clearBorderStyle()

   On Error GoTo clearBorderStyle_Error

    fraGeneral.Visible = False
    fraConfig.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraPosition.Visible = False
    fraDevelopment.Visible = False
    fraSounds.Visible = False
    fraAbout.Visible = False

    fraGeneralButton.BorderStyle = 0
    fraConfigButton.BorderStyle = 0
    fraDevelopmentButton.BorderStyle = 0
    fraPositionButton.BorderStyle = 0
    fraFontsButton.BorderStyle = 0
    fraWindowButton.BorderStyle = 0
    fraSoundsButton.BorderStyle = 0
    fraAboutButton.BorderStyle = 0
    
    #If TWINBASIC Then
        fraGeneralButton.Refresh
        fraConfigButton.Refresh
        fraDevelopmentButton.Refresh
        fraPositionButton.Refresh
        fraFontsButton.Refresh
        fraWindowButton.Refresh
        fraSoundsButton.Refresh
        fraAboutButton.Refresh
    #End If

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearBorderStyle of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 30/05/2023
' Purpose   : IMPORTANT: Called at every twip of resising, goodness knows what interval, we barely use this, instead we subclass and look for WM_EXITSIZEMOVE
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()

    pvtPrefsFormResizedByDrag = True
         
    ' do not call the resizing function when the form is resized by dragging the border
    ' only call this if the resize is done in code
        
    If InIDE Or gblPrefsFormResizedInCode = True Then
        Call PrefsForm_Resize_Event
    End If
            
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrefsForm_Resize_Event
' Author    : beededea
' Date      : 10/10/2024
' Purpose   : Called mostly by WM_EXITSIZEMOVE, the subclassed (intercepted) message that indicates that the window has just been moved.
'             (and on a mouseUp during a bottom-right drag of the additional corner indicator). Also, in code as specifcally required with an indicator flag.
'             This prevents a resize occurring during every twip movement and the controls resizing themselves continuously.
'             They now only resize when the form resize has completed.
'
'---------------------------------------------------------------------------------------
'
Public Sub PrefsForm_Resize_Event()

    Dim currentFontSize As Long: currentFontSize = 0
    
    On Error GoTo PrefsForm_Resize_Event_Error

    ' When minimised and a resize is called then simply exit.
    If Me.WindowState = vbMinimized Then Exit Sub
        
    If pvtPrefsDynamicSizingFlg = True And pvtPrefsFormResizedByDrag = True Then
    
        widgetPrefs.Width = widgetPrefs.Height / gblConstraintRatio ' maintain the aspect ratio, note: this change calls this routine again...
        
        If gblDpiAwareness = "1" Then
            currentFontSize = gblPrefsFontSizeHighDPI
        Else
            currentFontSize = gblPrefsFontSizeLowDPI
        End If

        'make tab frames invisible so that the control resizing is not apparent to the user
        Call makeFramesInvisible
        Call resizeControls(Me, prefsControlPositions(), gblPrefsStartWidth, gblPrefsStartHeight, currentFontSize)

        Call tweakPrefsControlPositions(Me, gblPrefsStartWidth, gblPrefsStartHeight)
        'Call loadHigherResPrefsImages ' if you want higher res icons then load them here, current max. is 1010 twips or 67 pixels
        Call makeFramesVisible
        
    Else
        If Me.WindowState = 0 Then ' normal
            If widgetPrefs.Width > 9090 Then widgetPrefs.Width = 9090
            If widgetPrefs.Width < 9085 Then widgetPrefs.Width = 9090
            If pvtLastFormHeight <> 0 Then
               gblPrefsFormResizedInCode = True
               widgetPrefs.Height = pvtLastFormHeight
            End If
        End If
    End If
    
    gblPrefsFormResizedInCode = False
    pvtPrefsFormResizedByDrag = False

   On Error GoTo 0
   Exit Sub

PrefsForm_Resize_Event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrefsForm_Resize_Event of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeFramesInvisible
' Author    : beededea
' Date      : 23/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeFramesInvisible()
    
   On Error GoTo makeFramesInvisible_Error

            If gblLastSelectedTab = "general" Then
                fraGeneral.Visible = False
                fraGeneralButton.Visible = False
            End If
            If gblLastSelectedTab = "config" Then
                fraConfig.Visible = False
                fraConfigButton.Visible = False
            End If
            If gblLastSelectedTab = "position" Then
                fraPosition.Visible = False
                fraPositionButton.Visible = False
            End If
                
            If gblLastSelectedTab = "development" Then
                fraDevelopment.Visible = False
                fraDevelopmentButton.Visible = False
            End If

            If gblLastSelectedTab = "fonts" Then
                fraFonts.Visible = False
                fraFontsButton.Visible = False
            End If

            If gblLastSelectedTab = "sounds" Then
                fraSounds.Visible = False
                fraSoundsButton.Visible = False
            End If

            If gblLastSelectedTab = "window" Then
                fraWindow.Visible = False
                fraWindowButton.Visible = False
            End If

            If gblLastSelectedTab = "about" Then
                fraAbout.Visible = False
                fraAboutButton.Visible = False
            End If

   On Error GoTo 0
   Exit Sub

makeFramesInvisible_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeFramesInvisible of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeFramesVisible
' Author    : beededea
' Date      : 23/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeFramesVisible()
    
   On Error GoTo makeFramesVisible_Error

            If gblLastSelectedTab = "general" Then
                fraGeneral.Visible = True
                fraGeneralButton.Visible = True
            End If
            If gblLastSelectedTab = "config" Then
                fraConfig.Visible = True
                fraConfigButton.Visible = True
            End If
            If gblLastSelectedTab = "position" Then
                fraPosition.Visible = True
                fraPositionButton.Visible = True
            End If
                
            If gblLastSelectedTab = "development" Then
                fraDevelopment.Visible = True
                fraDevelopmentButton.Visible = True
            End If

            If gblLastSelectedTab = "fonts" Then
                fraFonts.Visible = True
                fraFontsButton.Visible = True
            End If

            If gblLastSelectedTab = "sounds" Then
                fraSounds.Visible = True
                fraSoundsButton.Visible = True
            End If

            If gblLastSelectedTab = "window" Then
                fraWindow.Visible = True
                fraWindowButton.Visible = True
            End If

            If gblLastSelectedTab = "about" Then
                fraAbout.Visible = True
                fraAboutButton.Visible = True
            End If

   On Error GoTo 0
   Exit Sub

makeFramesVisible_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeFramesVisible of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_Moved
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Non VB6-standard event caught by subclassing and intercepting the WM_EXITSIZEMOVE (WM_MOVED) event
'---------------------------------------------------------------------------------------
'
Public Sub Form_Moved(sForm As String)

    On Error GoTo Form_Moved_Error
        
    'passing a form name as it allows us to potentially subclass another form's movement
    Select Case sForm
        Case "widgetPrefs"
            ' call a resize of all controls only when the form resize (by dragging) has completed (mouseUP)
            If pvtPrefsFormResizedByDrag = True Then
            
                ' test the current form height and width, if the same then it is a form_moved and not a form_resize.
                If widgetPrefs.Height = widgetPrefsOldHeight And widgetPrefs.Width = widgetPrefsOldWidth Then
                    Exit Sub
                Else
                    widgetPrefsOldHeight = widgetPrefs.Height
                    widgetPrefsOldWidth = widgetPrefs.Width
                    
                    Call PrefsForm_Resize_Event
                    pvtPrefsFormResizedByDrag = False
                    
                End If
            End If
            
            ' call the procedure to resize the form automatically if it now resides on a different sized monitor
            Call positionPrefsByMonitorSize
           
        Case Else
    End Select
    
   On Error GoTo 0
   Exit Sub

Form_Moved_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Moved of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tweakPrefsControlPositions
' Author    : beededea
' Date      : 22/09/2023
' Purpose   : final tweak the bottom frame top and left positions
'---------------------------------------------------------------------------------------
'
Private Sub tweakPrefsControlPositions(ByVal thisForm As Form, ByVal m_FormWid As Single, ByVal m_FormHgt As Single)

    ' not sure why but the resizeControls routine can lead to incorrect positioning of frames and buttons
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
    
    On Error GoTo tweakPrefsControlPositions_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt

    fraGeneral.Left = fraGeneralButton.Left
    fraConfig.Left = fraGeneralButton.Left
    fraSounds.Left = fraGeneralButton.Left
    fraPosition.Left = fraGeneralButton.Left
    fraFonts.Left = fraGeneralButton.Left
    fraDevelopment.Left = fraGeneralButton.Left
    fraWindow.Left = fraGeneralButton.Left
    fraAbout.Left = fraGeneralButton.Left
         
    'fraGeneral.Top = fraGeneralButton.Top
    fraConfig.Top = fraGeneral.Top
    fraSounds.Top = fraGeneral.Top
    fraPosition.Top = fraGeneral.Top
    fraFonts.Top = fraGeneral.Top
    fraDevelopment.Top = fraGeneral.Top
    fraWindow.Top = fraGeneral.Top
    fraAbout.Top = fraGeneral.Top
    
    ' final tweak the bottom button positions
    
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (100 * y_scale)
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnHelp.Top
    
    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - (150 * x_scale)
    btnHelp.Left = fraGeneral.Left

    txtPrefsFontCurrentSize.Text = y_scale * txtPrefsFontCurrentSize.FontSize
    
    lblAsterix.Top = btnSave.Top + 50
    
   On Error GoTo 0
   Exit Sub

tweakPrefsControlPositions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tweakPrefsControlPositions of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : standard form unload
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

    'gblPrefsLoadedFlg = False
    
    ' Release the subclass hook for dialog forms
    If Not InIDE Then ReleaseHook
    
    IsLoaded = False
    
    Call writePrefsPositionAndSize
    
    Call DestroyToolTip

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : various _MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : setting the balloon tooltip text for several controls
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : optGaugeTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : setting the tooltip text for the specific radio button for selecting the gauge/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optGaugeTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString
    On Error GoTo optGaugeTooltips_MouseMove_Error

    If gblPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optGaugeTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the GUI", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the RichClient square tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips."
            CreateToolTip optGaugeTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on RichClient Tooltips on the GUI", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optGaugeTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the GUI", , , , True
        End If
    
    End If

   On Error GoTo 0
   Exit Sub

optGaugeTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGaugeTooltips_MouseMove of Form widgetPrefs"
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_MouseMove
' Author    : beededea
' Date      : 01/10/2023
' Purpose   : reset message boxes mouseOver
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo btnResetMessages_MouseMove_Error

    If gblPrefsTooltips = "0" Then CreateToolTip btnResetMessages.hWnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
                  TTIconInfo, "Help on the message reset button", , , , True

    On Error GoTo 0
    Exit Sub

btnResetMessages_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_MouseMove of Form widgetPrefs"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True
End Sub

Private Sub chkEnableResizing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableResizing.hWnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
                  TTIconInfo, "Help on Resizing", , , , True
End Sub

Private Sub txtMainBias_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblGaugeTooltips = "0" Then CreateToolTip txtMainBias.hWnd, "This field displays the current total bias that will be applied to the current time based upon the timezone and daylight savings selection made, (read only). ", _
                  TTIconInfo, "Help on the bias field.", , , , True
End Sub

Private Sub txtPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFont.hWnd, "This is a read-only text box. It displays the current font as set when you click the font selector button. This is in operation for informational purposes only. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change.  My preferred font for this utility is Centurion Light SF at 8pt size.", _
                  TTIconInfo, "Help on the Currently Selected Font", , , , True
End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFontCurrentSize.hWnd, "This is a read-only text box. It displays the current font size as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
                  TTIconInfo, "Help on Setting the Font size Dynamically", , , , True
End Sub

Private Sub txtPrefsFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFontSize.hWnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size", , , , True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseMove
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseMove_Error

    lblDragCorner.MousePointer = 8

    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form widgetPrefs"
   
End Sub


Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnAboutDebugInfo.hWnd, "Here you can switch on Debug mode, not yet functional for this widget.", _
                  TTIconInfo, "Help on the Debug Info. Buttton", , , , True
End Sub



Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnClose.hWnd, "Close the Preference Utility", _
                  TTIconInfo, "Help on the Close Buttton", , , , True
End Sub

Private Sub btnDefaultVB6Editor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDefaultEditor.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select a Visual Basic Project (VBP) file for opening via the right click menu edit option. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the VBP File Explorer Button", , , , True
End Sub

Private Sub btnDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDisplayScreenFont.hWnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. When resizing the main gauge the display screen font size will change in relation to gauge size. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on the Font Selector Button", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDonate.hWnd, "Here you can visit my KofI page and donate a Coffee if you like my creations.", _
                  TTIconInfo, "Help on the Donate Buttton", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnFacebook.hWnd, "Here you can visit the Facebook page for the steampunk Widget community.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnGithubHome_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnGithubHome.hWnd, "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnHelp.hWnd, "Opens the help document, this will open as a compiled HTML file.", _
                  TTIconInfo, "Help on the Help Buttton", , , , True
End Sub



Private Sub btnOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnOpenFile.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select any file you would like to execute on a shift+DBlClick. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the shift+DBlClick File Explorer Button", , , , True
End Sub

Private Sub btnPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnPrefsFont.hWnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. Centurion Light SF is a good one and my personal favourite. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on Setting the Font Selector Button", , , , True
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnSave.hWnd, "Save the changes you have made to the preferences", _
                  TTIconInfo, "Help on the Save Buttton", , , , True
End Sub

Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnUpdate.hWnd, "Here you can able to download a new version of the program from github, when you click the button it will open a browser window and take you to the github page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub


Private Sub chkDpiAwareness_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkDpiAwareness.hWnd, "Check here to make the program DPI aware. NOT required on small to medium screens that are less than 1920 bytes wide. Try it and see which suits your system. RESTART required.", _
                  TTIconInfo, "Help on DPI Awareness Mode", , , , True
End Sub



Private Sub chkEnableSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableSounds.hWnd, "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI, as well as all other chimes, tick sounds &c.", _
                  TTIconInfo, "Help on Enabling/Disabling Sounds", , , , True
End Sub



Private Sub chkGenStartup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkGenStartup.hWnd, "Check this box to enable the automatic start of the program when Windows is started.", _
                  TTIconInfo, "Help on the Widget Automatic Start Toggle", , , , True
End Sub

Private Sub chkIgnoreMouse_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkIgnoreMouse.hWnd, "Checking this box causes the program to ignore all mouse events. A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Ignore Mouse optGaugeTooltips", , , , True
End Sub

Private Sub chkPreventDragging_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkPreventDragging.hWnd, "Checking this box causes the program to lock in place and ignore all attempts to move it with the mouse. " & vbCrLf & vbCrLf & _
        "The widget can be locked into a certain position in either landscape/portrait mode, ensuring that the widget always appears exactly where you want it to.  " & vbCrLf & vbCrLf & _
        "Using the fields adjacent, you can assign a default x/y position for both Landscape or Portrait mode.  " & vbCrLf & vbCrLf & _
        "When the widget is locked in place (using the Widget Position Locked option in the Window Tab), this value is set automatically.  " & vbCrLf & vbCrLf & _
        "A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Lock in Place option", , , , True
End Sub

Private Sub chkShowHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkShowHelp.hWnd, "Checking this box causes the rather attractive help canvas to appear every time the widget is started.", _
                  TTIconInfo, "Help on the Ignore Mouse option", , , , True
End Sub

Private Sub chkShowTaskbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkShowTaskbar.hWnd, "Check the box to show the widget in the Windows taskbar. A typical user may have multiple desktop widgets and it makes no sense to fill the taskbar with taskbar entries, this option allows you to enable a single one or two at your whim.", _
                  TTIconInfo, "Help on the Showing Entries in the Taskbar", , , , True
End Sub




Private Sub chkGaugeFunctions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkGaugeFunctions.hWnd, "When checked this box enables the spinning earth functionality. Any adjustment takes place instantly.", _
                  TTIconInfo, "Help on the Widget Function Toggle", , , , True
End Sub

Private Sub chkWidgetHidden_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkWidgetHidden.hWnd, "Checking this box causes the program to hide for a certain number of minutes. More useful from the widget's right click menu where you can hide the widget at will. Seemingly, a strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Hidden option", , , , True
End Sub
Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True
    If gblPrefsTooltips = "0" Then CreateToolTip fraAbout.hWnd, "The About tab tells you all about this program and its creation using " & gblCodingEnvironment & ".", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraConfigInner.hWnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraConfig.hWnd, "The configuration panel is the location for important configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True
End Sub
Private Sub fraDevelopment_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraDevelopment.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub
Private Sub fraDefaultVB6Editor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H80000012
End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraDevelopmentInner.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraFonts.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub

Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraFontsInner.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub


Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraGeneral.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraGeneralInner.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraPosition.hWnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub
Private Sub fraPositionInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraPositionInner.hWnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub

Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False

End Sub
Private Sub fraSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If gblPrefsTooltips = "0" Then CreateToolTip fraSounds.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraSoundsInner.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraWindow.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraWindowInner.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub


Private Sub fraGeneralButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraGeneralButton.hWnd, "Clicking on the General icon reveals the General Tab where the essential items can be configured, alarms, startup &c.", _
                  TTIconInfo, "Help on the General Tab Icon", , , , True
End Sub

Private Sub fraConfigButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraConfigButton.hWnd, "Clicking on the Config icon reveals the Configuration Tab where the optional items can be configured, DPI, tooltips &c.", _
                  TTIconInfo, "Help on the Configuration Tab Icon", , , , True
End Sub

Private Sub fraFontsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraFontsButton.hWnd, "Clicking on the Fonts icon reveals the Fonts Tab where the font related items can be configured, size, type, popups &c.", _
                  TTIconInfo, "Help on the Font Tab Icon", , , , True
End Sub
Private Sub fraSoundsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraSoundsButton.hWnd, "Clicking on the Sounds icon reveals the Sounds Tab where sound related items can be configured, volume, type &c.", _
                  TTIconInfo, "Help on the Sounds Tab Icon", , , , True
End Sub
Private Sub fraPositionButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraPositionButton.hWnd, "Clicking on the Position icon reveals the Position Tab where items related to Positioning can be configured, aspect ratios, landscape, &c.", _
                  TTIconInfo, "Help on the Position Tab Icon", , , , True
End Sub
Private Sub fraDevelopmentButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraDevelopmentButton.hWnd, "Clicking on the Development icon reveals the Development Tab where items relating to Development can be configured, debug, VBP location, &c.", _
                  TTIconInfo, "Help on the Development Tab Icon", , , , True
End Sub
Private Sub fraWindowButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraWindowButton.hWnd, "Clicking on the Window icon reveals the Window Tab where items relating to window sizing and layering can be configured &c.", _
                  TTIconInfo, "Help on the Window Tab Icon", , , , True
End Sub
Private Sub fraAboutButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraAboutButton.hWnd, "Clicking on the About icon reveals the About Tab where information about this desktop widget can be revealed.", _
                  TTIconInfo, "Help on the About Tab Icon", , , , True
End Sub
Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H8000000D
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : series of radio buttons to set the tooltip type for the prefs utility
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString

    On Error GoTo optPrefsTooltips_MouseMove_Error

    If gblPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the Preference Utility", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the standard Windows-style square tooltips for elements within the Steampunk GUI. These tooltips are single-line and the font size is limited to the Windows font size."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on  VB6 Native Tooltips on the Preference Utility", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the Preference Utility", , , , True
        End If
    End If

   On Error GoTo 0
   Exit Sub

optPrefsTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_MouseMove of Form widgetPrefs"
End Sub

Private Sub sliGaugeSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliGaugeSize.hWnd, "Adjust to a percentage of the original size. Any adjustment in size made here takes place instantly (you can also use Ctrl+Mousewheel when hovering over the gauge itself).", _
                  TTIconInfo, "Help on the Size Slider", , , , True
End Sub

Private Sub sliOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliOpacity.hWnd, "Sliding this causes the program's opacity to change from solidly opaque to fully transparent or some way in-between. Seemingly, a strange option for a windows program, a useful left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Opacity Slider", , , , True

End Sub
Private Sub txtAboutText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False
End Sub



Private Sub txtDblClickCommand_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDblClickCommand.hWnd, "Field to hold the any double click command that you have assigned to this widget. For example: taskmgr or %systemroot%\syswow64\ncpa.cpl", _
                  TTIconInfo, "Help on the Double Click Command", , , , True
End Sub

Private Sub txtDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDefaultEditor.hWnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

Private Sub txtDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDisplayScreenFont.hWnd, "This is a read-only text box. It displays the current font - as set when you click the font selector button. This field is in operation for informational purposes only. When resizing the main gauge (CTRL+ mouse scroll wheel) the font size will change in relation to gauge size. The base font determines the initial size, the resulting resized font will dynamically change. My preferred font for the display screen is Courier New at 6pt size.", _
                  TTIconInfo, "Help on the Display Screen Font", , , , True
End Sub

Private Sub txtDisplayScreenFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDisplayScreenFontSize.hWnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size for Display Screen", , , , True
End Sub

Private Sub txtLandscapeHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtLandscapeHoffset.hWnd, "Field to hold the horizontal offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape X Horizontal Field", , , , True
End Sub

Private Sub txtLandscapeVoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtLandscapeVoffset.hWnd, "Field to hold the vertical offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape Y Vertical Field", , , , True
End Sub

Private Sub txtOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtOpenFile.hWnd, "Field to hold the path to a file you would like to execute on a shift+DBlClick, if you select the adjacent button a file explorer will appear allowing you to select any file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the shift+DBlClick Field", , , , True
End Sub

Private Sub txtPortraitHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPortraitHoffset.hWnd, "Field to hold the horizontal offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait X Horizontal Field", , , , True
End Sub

Private Sub txtPortraitYoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPortraitYoffset.hWnd, "Field to hold the vertical offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait Y Vertical Field", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : General _MouseDown events to generate menu pop-ups across the form
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : due to a bug/difference with TwinBasic versus VB6
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : standard form down event to generate the menu across the board
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then

        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
        
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseDown
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : the label corner mouse down
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseDown_Error
    
    If Button = vbLeftButton Then
        pvtPrefsFormResizedByDrag = True
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form widgetPrefs"

End Sub

Private Sub fraFonts_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
'
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraPositionInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub
Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgAbout.Visible = False
    imgAboutClicked.Visible = True
End Sub
Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDevelopment.Visible = False
    imgDevelopmentClicked.Visible = True
End Sub
Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgFonts.Visible = False
    imgFontsClicked.Visible = True
End Sub
Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgConfig.Visible = False
    imgConfigClicked.Visible = True
End Sub
Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPosition.Visible = False
    imgPositionClicked.Visible = True
End Sub
Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgSounds.Visible = False
    imgSoundsClicked.Visible = True
End Sub
Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgWindow.Visible = False
    imgWindowClicked.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtAboutText_MouseDown
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : make a pop up menu appear on the text box by being tricky and clever
'---------------------------------------------------------------------------------------
'
Private Sub txtAboutText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo txtAboutText_MouseDown_Error

    If Button = vbRightButton Then
        txtAboutText.Enabled = False
        txtAboutText.Enabled = True
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtAboutText_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAboutText_MouseDown of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblGitHub_dblClick
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : label to allow a link to github to be clicked
'---------------------------------------------------------------------------------------
'
Private Sub lblGitHub_dblClick()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo lblGitHub_dblClick_Error
    
    If gblGaugeFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub

    answer = vbYes
    answerMsg = "This option opens a browser window and take you straight to Github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Proceed to Github? ", True, "lblGitHubDblClick")
    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://github.com/yereverluvinunclebert/UBoat-StopWatch-" & gblCodingEnvironment, vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

lblGitHub_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGitHub_dblClick of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsScreenResolution_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : for handling rotation of the screen in tablet mode or a resolution change
'             possibly due to an old game in full screen mode.
            ' when the timer frequency is reduced caused some weird async. effects, 500 ms seems fine, disable the timer first
'---------------------------------------------------------------------------------------
'
Private Sub tmrPrefsScreenResolution_Timer()

'    Static oldWidgetPrefsLeft As Long
'    Static oldWidgetPrefsTop As Long
'    Static beenMovingFlg As Boolean
'
'    Static oldPrefsFormMonitorID As Long
''    Static oldPrefsFormMonitorPrimary As Long
'    Static oldPrefsMonitorStructWidthTwips As Long
'    Static oldPrefsMonitorStructHeightTwips As Long
'    Static oldPrefsGaugeLeftPixels As Long
'
'    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
'    'Dim prefsFormMonitorPrimary As Long: prefsFormMonitorPrimary = 0
'    Dim monitorStructWidthTwips As Long: monitorStructWidthTwips = 0
'    Dim monitorStructHeightTwips As Long: monitorStructHeightTwips = 0
'    Dim resizeProportion As Double: resizeProportion = 0
'
'    Dim answer As VbMsgBoxResult: answer = vbNo
'    Dim answerMsg As String: answerMsg = vbNullString
'
'    On Error GoTo tmrPrefsScreenResolution_Timer_Error
'
'    ' calls a routine that tests for a change in the monitor upon which the form sits, if so, resizes
'    If widgetPrefs.IsVisible = False Then Exit Sub
'
'    ' prefs hasn't moved at all
'    If widgetPrefs.Left = oldWidgetPrefsLeft Then Exit Sub  ' this can only work if the reposition is being performed by the timer
'    ' we are also hopefully calling this routine on a mouseUP event after a form move, where the above line will not apply.
'
'    ' if just one monitor or the global switch is off then exit
'    If monitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
'
'        ' turn off the timer that saves the prefs height and position
'        tmrPrefsMonitorSaveHeight.Enabled = False
'        tmrWritePosition.Enabled = False
'        tmrPrefsScreenResolution.Enabled = False ' turn off this very timer here
'
'        ' populate the OLD vars if empty, to allow valid comparison next run
'        If oldWidgetPrefsLeft <= 0 Then oldWidgetPrefsLeft = widgetPrefs.Left
'        If oldWidgetPrefsTop <= 0 Then oldWidgetPrefsTop = widgetPrefs.Top
'
'        ' test whether the form has been moved (VB6 has no FORM_MOVING nor FORM_MOVED EVENTS)
'        If widgetPrefs.Left <> oldWidgetPrefsLeft Or widgetPrefs.Top <> oldWidgetPrefsTop Then
'
'               ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
'                prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
'
'                'prefsFormMonitorPrimary = prefsMonitorStruct.IsPrimary ' -1 true
'
'                ' sample the physical monitor resolution
'                monitorStructWidthTwips = prefsMonitorStruct.Width
'                monitorStructHeightTwips = prefsMonitorStruct.Height
'
'                'if the old monitor ID has not been stored already (form load) then do so now
'                If oldPrefsFormMonitorID = 0 Then oldPrefsFormMonitorID = prefsFormMonitorID
'
'                ' same with other 'old' vars
'                If oldPrefsMonitorStructWidthTwips = 0 Then oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                If oldPrefsMonitorStructHeightTwips = 0 Then oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                If oldPrefsGaugeLeftPixels = 0 Then oldPrefsGaugeLeftPixels = widgetPrefs.Left
'
'                ' if the monitor ID has changed
'                If oldPrefsFormMonitorID <> prefsFormMonitorID Then
'                'If oldPrefsFormMonitorPrimary <> prefsFormMonitorPrimary Then
'
''                    ' screenWrite ("Prefs Stored monitor primary status = " & CBool(oldPrefsFormMonitorPrimary))
''                    ' screenWrite ("Prefs Current monitor primary status = " & CBool(prefsFormMonitorPrimary))
'
'                    If LTrim$(gblMultiMonitorResize) = "1" Then
'                        'if the resolution is different then calculate new size proportion
'                        If monitorStructWidthTwips <> oldPrefsMonitorStructWidthTwips Or monitorStructHeightTwips <> oldPrefsMonitorStructHeightTwips Then
'                            'now calculate the size of the widget according to the screen HeightTwips.
'                            resizeProportion = prefsMonitorStruct.Height / oldPrefsMonitorStructHeightTwips
'                            newPrefsHeight = widgetPrefs.Height * resizeProportion
'                            widgetPrefs.Height = newPrefsHeight
'                        End If
'                    ElseIf LTrim$(gblMultiMonitorResize) = "2" Then
'                        ' set the size according to saved values
'                        If prefsMonitorStruct.IsPrimary = True Then
'                            widgetPrefs.Height = Val(gblPrefsPrimaryHeightTwips)
'                        Else
'                            'gblPrefsSecondaryHeightTwips = "15000"
'                            widgetPrefs.Height = Val(gblPrefsSecondaryHeightTwips)
'                        End If
'                    End If
'
'                End If
'
'                ' set the current values as 'old' for comparison on next run
'                'oldPrefsFormMonitorPrimary = prefsFormMonitorPrimary
'                oldPrefsFormMonitorID = prefsFormMonitorID
'                oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                oldPrefsGaugeLeftPixels = widgetPrefs.Left
'            End If
'
'    End If
'
'    oldWidgetPrefsLeft = widgetPrefs.Left
'    oldWidgetPrefsTop = widgetPrefs.Top
'
'    tmrPrefsScreenResolution.Enabled = True
'    tmrPrefsMonitorSaveHeight.Enabled = True
'    tmrWritePosition.Enabled = True
    
    On Error GoTo 0
    Exit Sub

tmrPrefsScreenResolution_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsScreenResolution_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : General _MouseUp events to generate menu pop-ups across the form
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : due to a bug/difference with TwinBasic versus VB6
'---------------------------------------------------------------------------------------
#If TWINBASIC Then
    Private Sub imgAboutClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
    End Sub
#Else
    Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgDevelopmentClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    End Sub
#Else
    Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgFontsClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    End Sub
#Else
    Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgConfigClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
    End Sub
#Else
    Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgPositionClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    End Sub
#Else
    Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgSoundsClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    End Sub
#Else
    Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgWindowClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    End Sub
#Else
    Private Sub imgWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgGeneralClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
    End Sub
#Else
    Private Sub imgGeneral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
    End Sub
#End If


Private Sub sliGaugeSize_GotFocus()
    pvtAllowSizeChangeFlg = True
End Sub

Private Sub sliGaugeSize_LostFocus()
    pvtAllowSizeChangeFlg = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Various Change events below
' Author    : beededea
' Date      : 15/08/2023
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 15/08/2023
' Purpose   : save the slider opacity values as they change
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo sliOpacity_Change_Error

    btnSave.Enabled = True ' enable the save button

    If pvtPrefsStartupFlg = False Then
        gblOpacity = CStr(sliOpacity.Value)
    
        sPutINISetting "Software\UBoatStopWatch", "opacity", gblOpacity, gblSettingsFile
        
        'Call setOpacity(sliOpacity.Value) ' this works but reveals the background form itself
        
        answer = vbYes
        answerMsg = "You must perform a hard reload on this widget in order to change the widget's opacity, do you want me to do it for you now?"
        answer = msgBoxA(answerMsg, vbYesNo, "Hard Reload Request", True, "sliOpacityClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call hardRestart
        End If
    End If

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : slider to change opacity of the whole gauge.
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Change()
   On Error GoTo sliOpacity_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliGaugeSize_Change
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : slider to change the size of the whole gauge.
'---------------------------------------------------------------------------------------
'
Public Sub sliGaugeSize_Change()
    On Error GoTo sliGaugeSize_Change_Error

    btnSave.Enabled = True ' enable the save button
    
    'If pvtAllowSizeChangeFlg = True Then Call fGauge.AdjustZoom(sliGaugeSize.Value / 100)
    If pvtAllowSizeChangeFlg = True Then Me.GaugeSize = sliGaugeSize.Value / 100
    
    Call saveMainRCFormSize

    On Error GoTo 0
    Exit Sub

sliGaugeSize_Change_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGaugeSize_Change of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Property  : GaugeSize
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the GaugeSize of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Get GaugeSize() As Single
   On Error GoTo gaugeSizeGet_Error

   GaugeSize = mGaugeSize

   On Error GoTo 0
   Exit Property

gaugeSizeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property GaugeSize of Class Module cwoverlay"
End Property

'---------------------------------------------------------------------------------------
' Property  : GaugeSize
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value) the GaugeSize value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let GaugeSize(ByVal newValue As Single)
   On Error GoTo gaugeSizeLet_Error

    If mGaugeSize <> newValue Then mGaugeSize = newValue Else Exit Property
        
    Call fGauge.AdjustZoom(mGaugeSize)

   On Error GoTo 0
   Exit Property

gaugeSizeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property GaugeSize of Class Module cwoverlay"
End Property


Private Sub txtDblClickCommand_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeVoffset_Change()
    btnSave.Enabled = True ' enable the save button
End Sub
Private Sub txtOpenFile_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtPortraitHoffset_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtPortraitYoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPrefsFont_Change()
    btnSave.Enabled = True ' enable the save button
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : right click about option from the pop-up menu
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


'---------------------------------------------------------------------------------------
' Procedure : setPrefsTooltips
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsTooltips()

   On Error GoTo setPrefsTooltips_Error
    
    ' here we set the variables used for the comboboxes, each combobox has to be sub classed and these variables are used during that process
    
    If optPrefsTooltips(0).Value = True Then
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        pCmbMultiMonitorResizeBalloonTooltip = "This option will only appear on multi-monitor systems. This dropdown has three choices that affect the automatic sizing of both the main gauge and the preference utility. " & vbCrLf & vbCrLf & _
            "For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on. The automatic option resizes according to the relative proportions of the two screens.  " & vbCrLf & vbCrLf & _
            "The manual option resizes according to sizes that you set manually. Just resize the gauge on the monitor of your choice and the program will store it. This option only works for no more than TWO monitors."
   
        pCmbScrollWheelDirectionBalloonTooltip = "This option will allow you to change the direction of the mouse scroll wheel when resizing the gauge gauge. IF you want to resize the gauge on your desktop, hold the CTRL key along with moving the scroll wheel UP/DOWN. Some prefer scrolling UP rather than DOWN. You configure that here."
        pCmbWindowLevelBalloonTooltip = "You can determine the window level here. You can keep it above all other windows or you can set it to bottom to keep the widget below all other windows."
        pCmbHidingTimeBalloonTooltip = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        
        pCmbWidgetLandscapeBalloonTooltip = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        pCmbWidgetPortraitBalloonTooltip = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        pCmbWidgetPositionBalloonTooltip = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        pCmbAspectHiddenBalloonTooltip = "Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        
        pCmbDebugBalloonTooltip = "Here you can set debug mode. This will enable the editor field and allow you to assign a VBP/TwinProj file for the " & gblCodingEnvironment & " IDE editor"
        
    Else
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        
        pCmbMultiMonitorResizeBalloonTooltip = vbNullString
        pCmbScrollWheelDirectionBalloonTooltip = vbNullString
        pCmbWindowLevelBalloonTooltip = vbNullString
        pCmbHidingTimeBalloonTooltip = vbNullString
        
        pCmbWidgetLandscapeBalloonTooltip = vbNullString
        pCmbWidgetPortraitBalloonTooltip = vbNullString
        pCmbWidgetPositionBalloonTooltip = vbNullString
        pCmbAspectHiddenBalloonTooltip = vbNullString
        pCmbDebugBalloonTooltip = vbNullString
        
        ' for some reason, the balloon tooltip on the checkbox used to dismiss the balloon tooltips does not disappear, this forces it go away.
        CreateToolTip optPrefsTooltips(0).hWnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(1).hWnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(2).hWnd, "", _
                  TTIconInfo, "Help", , , , True
                  
    End If
    
    
    ' next we just do the native VB6 tooltips
     If optPrefsTooltips(1).Value = True Then
       
       imgConfig.ToolTipText = "Opens the configuration tab"
        imgConfigClicked.ToolTipText = "Opens the configuration tab"
        imgDevelopment.ToolTipText = "Opens the Development tab"
        imgDevelopmentClicked.ToolTipText = "Opens the Development tab"
        imgPosition.ToolTipText = "Opens the Position tab"
        imgPositionClicked.ToolTipText = "Opens the Position tab"
        btnSave.ToolTipText = "Save the changes you have made to the preferences"
        btnHelp.ToolTipText = "Open the help utility"
        imgSounds.ToolTipText = "Opens the Sounds tab"
        imgSoundsClicked.ToolTipText = "Opens the Sounds tab"
        btnClose.ToolTipText = "Close the utility"
        imgWindow.ToolTipText = "Opens the Window tab"
        imgWindowClicked.ToolTipText = "Opens the Window tab"
        lblWindow.ToolTipText = "Opens the Window tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFontsClicked.ToolTipText = "Opens the Fonts tab"
        imgGeneral.ToolTipText = "Opens the general tab"
        imgGeneralClicked.ToolTipText = "Opens the general tab"
        lblPosition(6).ToolTipText = "Tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this the widget widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE."
        chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
        chkGaugeFunctions.ToolTipText = "When checked this box enables the spinning earth functionality. Any adjustment takes place instantly. "
        
        txtPortraitYoffset.ToolTipText = "Field to hold the vertical offset for the widget position in portrait mode."
        txtPortraitHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in portrait mode."
        txtLandscapeVoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        txtLandscapeHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        cmbWidgetLandscape.ToolTipText = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        cmbWidgetPortrait.ToolTipText = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        cmbWidgetPosition.ToolTipText = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        cmbAspectHidden.ToolTipText = " Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        chkEnableSounds.ToolTipText = "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI as well as all other chimes, tick sounds."
'        chkEnableTicks.ToolTipText = "Enables or disables just the sound of the gauge ticking."
'        chkEnableChimes.ToolTipText = "Enables or disables just the gauge chimes."
'        chkEnableAlarms.ToolTipText = "Enables or disables the gauge alarm chimes. Please note disabling this means your alarms will not alert you audibly!"
        
'        chkVolumeBoost.ToolTipText = "Sets the volume of the various sound elements, you can boost from quiet to loud."
        btnDefaultEditor.ToolTipText = "Click to select the .vbp file to edit the program - You need to have access to the source!"
        txtDblClickCommand.ToolTipText = "Enter a Windows command for the gauge to operate when double-clicked."
        btnOpenFile.ToolTipText = "Click to select a particular file for the gauge to run or open when double-clicked."
        txtOpenFile.ToolTipText = "Enter a particular file for the gauge to run or open when double-clicked."
        cmbDebug.ToolTipText = "Choose to set debug mode."
        
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
        btnPrefsFont.ToolTipText = "The Font Selector."
        txtPrefsFont.ToolTipText = "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size via the font selector that fits the text boxes"
        
        lblFontsTab(0).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(6).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(7).ToolTipText = "Choose a font size that fits the text boxes"
        
        txtDisplayScreenFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the gauge display screen only"
        btnDisplayScreenFont.ToolTipText = "The Font Selector."
        txtDisplayScreenFont.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the gauge display screen only"
        
        cmbWindowLevel.ToolTipText = "You can determine the window position here. Set to bottom to keep the widget below other windows."
        cmbHidingTime.ToolTipText = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        cmbMultiMonitorResize.ToolTipText = "When you have a multi-monitor set-up, the widget can auto-resize on a smaller secondary monitor. Here you determine the proportion of the resize."
        
        chkEnableResizing.ToolTipText = "Provides an alternative method of supporting high DPI screens."
        chkPreventDragging.ToolTipText = "Checking this box turns off the ability to drag the program with the mouse. The locking in position effect takes place instantly."
        chkIgnoreMouse.ToolTipText = "Checking this box causes the program to ignore all mouse events."
        sliOpacity.ToolTipText = "Set the transparency of the program. Any change in opacity takes place instantly."
        cmbScrollWheelDirection.ToolTipText = "To change the direction of the mouse scroll wheel when resizing the gauge gauge."
        
        optGaugeTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        optGaugeTooltips(1).ToolTipText = "Check the box to enable RichClient square tooltips for all controls on the main program"
        optGaugeTooltips(2).ToolTipText = "Check the box to disable tooltips for all controls on the main program"
        
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        chkShowHelp.ToolTipText = "Check the box to show the help page on startup"
        
        sliGaugeSize.ToolTipText = "Adjust to a percentage of the original size. Any adjustment in size takes place instantly (you can also use Ctrl+Mousewheel hovering over the gauge itself)."
        btnFacebook.ToolTipText = "This will link you to the our Steampunk/Dieselpunk program users Group."
        imgAbout.ToolTipText = "Opens the About tab"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        btnDonate.ToolTipText = "Buy me a Kofi! This button opens a browser window and connects to Kofi donation page"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs."
        
        btnGithubHome.ToolTipText = "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page."

        txtPrefsFontCurrentSize.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        'lblCurrentFontsTab.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        
        chkDpiAwareness.ToolTipText = " Check the box to make the program DPI aware. RESTART required."
        'optEnablePrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
        
        optPrefsTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls within this Preference Utility. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size."
        optPrefsTooltips(1).ToolTipText = "Check the box to enable Windows-style square tooltips for all controls within this Preference Utility. Note that their font size will match the Windows system font size."
        optPrefsTooltips(2).ToolTipText = "This setting enables/disables the tooltips for all elements within this Preference Utility."

        btnResetMessages.ToolTipText = "This button restores the pop-up messages to their original visible state."
        

    Else
    

    
        imgConfig.ToolTipText = vbNullString
        imgConfigClicked.ToolTipText = vbNullString
        imgDevelopment.ToolTipText = vbNullString
        imgDevelopmentClicked.ToolTipText = vbNullString
        imgPosition.ToolTipText = vbNullString
        imgPositionClicked.ToolTipText = vbNullString
        btnSave.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        imgSounds.ToolTipText = vbNullString
        imgSoundsClicked.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        imgWindow.ToolTipText = vbNullString
        imgWindowClicked.ToolTipText = vbNullString
        imgFonts.ToolTipText = vbNullString
        imgFontsClicked.ToolTipText = vbNullString
        imgGeneral.ToolTipText = vbNullString
        imgGeneralClicked.ToolTipText = vbNullString
        chkGenStartup.ToolTipText = vbNullString
        chkGaugeFunctions.ToolTipText = vbNullString
                
        txtPortraitYoffset.ToolTipText = vbNullString
        txtPortraitHoffset.ToolTipText = vbNullString
        txtLandscapeVoffset.ToolTipText = vbNullString
        txtLandscapeHoffset.ToolTipText = vbNullString
        cmbWidgetLandscape.ToolTipText = vbNullString
        cmbWidgetPortrait.ToolTipText = vbNullString
        cmbWidgetPosition.ToolTipText = vbNullString
        cmbAspectHidden.ToolTipText = vbNullString
        chkEnableSounds.ToolTipText = vbNullString
'        chkEnableTicks.ToolTipText = vbNullString
'        chkEnableChimes.ToolTipText = vbNullString
'        chkEnableAlarms.ToolTipText = vbNullString
        'chkVolumeBoost.ToolTipText = vbNullString
        
        btnDefaultEditor.ToolTipText = vbNullString
        txtDblClickCommand.ToolTipText = vbNullString
        btnOpenFile.ToolTipText = vbNullString
        txtOpenFile.ToolTipText = vbNullString
        cmbDebug.ToolTipText = vbNullString
        txtPrefsFontSize.ToolTipText = vbNullString
        btnPrefsFont.ToolTipText = vbNullString
        txtPrefsFont.ToolTipText = vbNullString
        txtPrefsFontCurrentSize.ToolTipText = vbNullString
        
        
        txtDisplayScreenFontSize.ToolTipText = vbNullString
        btnDisplayScreenFont.ToolTipText = vbNullString
        txtDisplayScreenFont.ToolTipText = vbNullString
        
        cmbWindowLevel.ToolTipText = vbNullString
        cmbHidingTime.ToolTipText = vbNullString
        cmbMultiMonitorResize.ToolTipText = vbNullString
        
        chkEnableResizing.ToolTipText = vbNullString
        chkPreventDragging.ToolTipText = vbNullString
        chkIgnoreMouse.ToolTipText = vbNullString
        sliOpacity.ToolTipText = vbNullString
        cmbScrollWheelDirection.ToolTipText = vbNullString
        
        optGaugeTooltips(0).ToolTipText = vbNullString
        optGaugeTooltips(1).ToolTipText = vbNullString
        optGaugeTooltips(2).ToolTipText = vbNullString
        
        chkShowTaskbar.ToolTipText = vbNullString
        chkShowHelp.ToolTipText = vbNullString

        sliGaugeSize.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        imgAbout.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        btnGithubHome.ToolTipText = vbNullString
        
        chkDpiAwareness.ToolTipText = vbNullString
        'optEnablePrefsTooltips.ToolTipText = vbNullString
        
        optPrefsTooltips(0).ToolTipText = vbNullString
        optPrefsTooltips(1).ToolTipText = vbNullString
        optPrefsTooltips(2).ToolTipText = vbNullString
        
        btnResetMessages.ToolTipText = vbNullString
    
    End If

   On Error GoTo 0
   Exit Sub

setPrefsTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsTooltips of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setPrefsLabels
' Author    : beededea
' Date      : 27/09/2023
' Purpose   : set the text in any labels that need a vbCrLf to space the text
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsLabels()

    On Error GoTo setPrefsLabels_Error

    lblFontsTab(0).Caption = "When resizing the form (drag bottom right) the font size will in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change." & vbCrLf & vbCrLf & _
        "My preferred font for this utility is Centurion Light SF at 8pt size."

    On Error GoTo 0
    Exit Sub

setPrefsLabels_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsLabels of Form widgetPrefs"
        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DestroyToolTip
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : It's not a bad idea to put this in the Form_Unload event just to make sure.
'---------------------------------------------------------------------------------------
'
Public Sub DestroyToolTip()
    
   On Error GoTo DestroyToolTip_Error

    If hwndTT <> 0& Then DestroyWindow hwndTT
    hwndTT = 0&

   On Error GoTo 0
   Exit Sub

DestroyToolTip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DestroyToolTip of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadPrefsAboutText
' Author    : beededea
' Date      : 12/03/2020
' Purpose   : The text for the about page is stored here
'---------------------------------------------------------------------------------------
'
Private Sub loadPrefsAboutText()
    On Error GoTo loadPrefsAboutText_Error
    'If gblDebugFlg = 1 Then Debug.Print "%loadPrefsAboutText"
    
    lblMajorVersion.Caption = App.Major
    lblMinorVersion.Caption = App.Minor
    lblRevisionNum.Caption = App.Revision
    
    lblAbout(1).Caption = "(32bit WoW64 using " & gblCodingEnvironment & ")"
    
    Call LoadFileToTB(txtAboutText, App.path & "\resources\txt\about.txt", False)
    
    ' ' fGauge.RotateBusyTimer = True
    
   On Error GoTo 0
   Exit Sub

loadPrefsAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPrefsAboutText of Form widgetPrefs"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : picButtonMouseUpEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : capture the icon button clicks avoiding creating a control array
'---------------------------------------------------------------------------------------
'
Private Sub picButtonMouseUpEvent(ByVal thisTabName As String, ByRef thisPicName As Image, ByRef thisPicNameClicked As Image, ByRef thisFraName As Frame, Optional ByRef thisFraButtonName As Frame)
    
    On Error GoTo picButtonMouseUpEvent_Error
    
    Dim padding As Long: padding = 0
    Dim BorderWidth As Long: BorderWidth = 0
    Dim captionHeight As Long: captionHeight = 0
    Dim y_scale As Single: y_scale = 0
    
    thisPicNameClicked.Visible = False
    thisPicName.Visible = True
      
    btnSave.Visible = False
    btnClose.Visible = False
    btnHelp.Visible = False
    
    Call clearBorderStyle

    gblLastSelectedTab = thisTabName
    sPutINISetting "Software\UBoatStopWatch", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile

    thisFraName.Visible = True
    
    thisFraButtonName.BorderStyle = 1

    #If TWINBASIC Then
        thisFraButtonName.Refresh
    #End If

    ' Get the form's current scale factors.
    y_scale = Me.ScaleHeight / gblPrefsStartHeight
    
    If gblDpiAwareness = "1" Then
        btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (100 * y_scale)
    Else
        btnHelp.Top = thisFraName.Top + thisFraName.Height + (200 * y_scale)
    End If
    
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnSave.Top
    
    btnSave.Visible = True
    btnClose.Visible = True
    btnHelp.Visible = True
    
    lblAsterix.Top = btnSave.Top + 50
    lblSize.Top = lblAsterix.Top - 300
    
    chkEnableResizing.Top = btnSave.Top + 50
    'chkEnableResizing.Left = lblAsterix.Left
    
    BorderWidth = (widgetPrefs.Width - Me.ScaleWidth) / 2
    captionHeight = widgetPrefs.Height - Me.ScaleHeight - BorderWidth
        
    ' under windows 10+ the internal window calcs are all wrong due to the bigger title bars
    If pvtPrefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        pvtLastFormHeight = btnHelp.Top + btnHelp.Height + captionHeight + BorderWidth + padding
        gblPrefsFormResizedInCode = True
        widgetPrefs.Height = pvtLastFormHeight
    End If
    
    If gblDpiAwareness = "0" Then
        If thisTabName = "about" Then
            lblAsterix.Visible = False
            chkEnableResizing.Visible = True
        Else
            lblAsterix.Visible = True
            chkEnableResizing.Visible = False
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

picButtonMouseUpEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picButtonMouseUpEvent of Form widgetPrefs"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : a timer to apply a theme automatically
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    Dim SysClr As Long: SysClr = 0

    On Error GoTo themeTimer_Timer_Error
    
    If widgetPrefs.IsVisible = False Then Exit Sub

    SysClr = GetSysColor(COLOR_BTNFACE)

    If SysClr <> gblStoreThemeColour Then
        Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click()
    On Error GoTo mnuCoffee_Click_Error
    
    Call mnuCoffee_ClickEvent

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form widgetPrefs"
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : menu option to show licence from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
    On Error GoTo mnuLicenceA_Click_Error

    Call mnuLicence_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuLicenceA_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicenceA_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open support page from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
    
    On Error GoTo mnuSupport_Click_Error

    Call mnuSupport_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuClosePreferences_Click
' Author    : beededea
' Date      : 06/09/2024
' Purpose   : right click close option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuClosePreferences_Click()
   On Error GoTo mnuClosePreferences_Click_Error

    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuClosePreferences_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClosePreferences_Click of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click auto theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
            mnuAuto.Checked = False
            
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            mnuAuto.Checked = True
            
            themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click dark theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enabled"
    mnuLight.Caption = "Light Theme Enable"
    themeTimer.Enabled = False
    
    gblSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click light theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error
    
    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enable"
    mnuLight.Caption = "Light Theme Enabled"
    themeTimer.Enabled = False
    
    gblSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form widgetPrefs"
End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 06/05/2023
' Purpose   : set the theme shade, Windows classic dark/new lighter theme colours from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)
    
    Dim Ctrl As Control
    
    On Error GoTo setThemeShade_Error

    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    
    ' all buttons must be set to graphical
    For Each Ctrl In Me.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        'classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
        
        Call setPrefsIconImagesDark
        
    Else
        'classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
        
        Call setPrefsIconImagesLight
                
    End If
    
    'now change the color of the sliders.
'    widgetPrefs.sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    'widgetPrefs.'sliWidgetSkew.BackColor = RGB(redC, greenC, blueC)
    sliGaugeSize.BackColor = RGB(redC, greenC, blueC)
    sliOpacity.BackColor = RGB(redC, greenC, blueC)
    txtAboutText.BackColor = RGB(redC, greenC, blueC)
    
    sPutINISetting "Software\UBoatStopWatch", "skinTheme", gblSkinTheme, gblSettingsFile

    On Error GoTo 0
    Exit Sub

setThemeShade_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeShade of Module Module1"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Private Sub setThemeColour()
    
    Dim SysClr As Long: SysClr = 0
    
   On Error GoTo setThemeColour_Error
   'If gblDebugFlg = 1  Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        gblSkinTheme = "dark"
        
        mnuDark.Caption = "Dark Theme Enabled"
        mnuLight.Caption = "Light Theme Enable"

    Else
        Call setModernThemeColours
        mnuDark.Caption = "Dark Theme Enable"
        mnuLight.Caption = "Light Theme Enabled"
    End If

    gblStoreThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsTheme
' Author    : beededea
' Date      : 25/04/2023
' Purpose   : adjust the theme used by the prefs alone
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsTheme()
   On Error GoTo adjustPrefsTheme_Error

    If gblSkinTheme <> vbNullString Then
        If gblSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If gblClassicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            themeTimer.Enabled = True
        Else
            gblSkinTheme = "light"
            Call setModernThemeColours
        End If
    End If
    
    ' ' fGauge.RotateBusyTimer = True

   On Error GoTo 0
   Exit Sub

adjustPrefsTheme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsTheme of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setModernThemeColours
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : by 'modern theme' we mean the very light indeed almost white background to standard Windows forms...
'---------------------------------------------------------------------------------------
'
Private Sub setModernThemeColours()
         
    Dim SysClr As Long: SysClr = 0
    
    On Error GoTo setModernThemeColours_Error
    
    'the widgetPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

    'MsgBox "Windows Alternate Theme detected"
    SysClr = GetSysColor(COLOR_BTNFACE)
    If SysClr = 13160660 Then
        Call setThemeShade(212, 208, 199)
        gblSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        gblSkinTheme = "light"
    End If

   On Error GoTo 0
   Exit Sub

setModernThemeColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setModernThemeColours of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadHigherResPrefsImages
' Author    : beededea
' Date      : 18/06/2023
' Purpose   : load the images for the classic or high brightness themes
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResPrefsImages()
    
    On Error GoTo loadHigherResPrefsImages_Error
      
    If Me.WindowState = vbMinimized Then Exit Sub
    
'
'    If dynamicSizingFlg = False Then
'        Exit Sub
'    End If
'
'    If Me.Width < 10500 Then
'        topIconWidth = 600
'    End If
'
'    If Me.Width >= 10500 And Me.Width < 12000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 730
'    End If
'
'    If Me.Width >= 12000 And Me.Width < 13500 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 834
'    End If
'
'    If Me.Width >= 13500 And Me.Width < 15000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 940
'    End If
'
'    If Me.Width >= 15000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 1010
'    End If
        
    If mnuDark.Checked = True Then
        Call setPrefsIconImagesDark
    Else
        Call setPrefsIconImagesLight
    End If
    
    ' ' fGauge.RotateBusyTimer = True
    
   On Error GoTo 0
   Exit Sub

loadHigherResPrefsImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : tmrWritePosition_Timer
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : periodically read the prefs form position and store
'---------------------------------------------------------------------------------------
'
Private Sub tmrWritePosition_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    On Error GoTo tmrWritePosition_Timer_Error
   
    If widgetPrefs.IsVisible = True Then Call writePrefsPositionAndSize

   On Error GoTo 0
   Exit Sub

tmrWritePosition_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrWritePosition_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnableResizing_Click
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : toggle to enable sizing when in low DPI aware mode
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableResizing_Click()
   On Error GoTo chkEnableResizing_Click_Error

    If chkEnableResizing.Value = 1 Then
        pvtPrefsDynamicSizingFlg = True
        txtPrefsFontCurrentSize.Visible = True
        'lblCurrentFontsTab.Visible = True
        'Call writePrefsPositionAndSize
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        pvtPrefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        'lblCurrentFontsTab.Visible = False
        Unload widgetPrefs
        Me.Show
        Call readPrefsPosition
        chkEnableResizing.Caption = "Enable Corner Resizing"
    End If
    
    Call setframeHeights

   On Error GoTo 0
   Exit Sub

chkEnableResizing_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableResizing_Click of Form widgetPrefs"

End Sub


 



'---------------------------------------------------------------------------------------
' Procedure : setframeHeights
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : set the frame heights to manual sizes for the low DPI mode as per the YWE prefs
'---------------------------------------------------------------------------------------
'
Private Sub setframeHeights()
   On Error GoTo setframeHeights_Error

    If pvtPrefsDynamicSizingFlg = True Then
        fraGeneral.Height = fraAbout.Height
        fraFonts.Height = fraAbout.Height
        fraConfig.Height = fraAbout.Height
        fraSounds.Height = fraAbout.Height
        fraPosition.Height = fraAbout.Height
        fraDevelopment.Height = fraAbout.Height
        fraWindow.Height = fraAbout.Height
        
        fraGeneral.Width = fraAbout.Width
        fraFonts.Width = fraAbout.Width
        fraConfig.Width = fraAbout.Width
        fraSounds.Width = fraAbout.Width
        fraPosition.Width = fraAbout.Width
        fraDevelopment.Width = fraAbout.Width
        fraWindow.Width = fraAbout.Width
    
        'If gblDpiAwareness = "1" Then
            ' save the initial positions of ALL the controls on the prefs form
            Call SaveSizes(widgetPrefs, prefsControlPositions(), gblPrefsStartWidth, gblPrefsStartHeight)
        'End If
    Else
        fraGeneral.Height = 7737
        fraConfig.Height = 8259
        fraSounds.Height = 3985
        fraPosition.Height = 7544
        fraFonts.Height = 4533
        
        ' the lowest window controls are not displayed on a single monitor
        If gblMonitorCount > 1 Then
            fraWindow.Height = 8138
            fraWindowInner.Height = 7500
        Else
            fraWindow.Height = 6586
            fraWindowInner.Height = 6000
        End If

        fraDevelopment.Height = 6297
        fraAbout.Height = 8700
    End If
    
   On Error GoTo 0
   Exit Sub

setframeHeights_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setframeHeights of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesDark
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : set the bright images for the grey classic theme
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesDark()
    
    On Error GoTo setPrefsIconImagesDark_Error
    
'    #If TWINBASIC Then
'
'        Set imgGeneral.Picture = LoadPicture(App.path & "\Resources\images\general-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgFonts.Picture = LoadPicture(App.path & "\Resources\images\font-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgSounds.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgPosition.Picture = LoadPicture(App.path & "\Resources\images\position-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgDevelopment.Picture = LoadPicture(App.path & "\Resources\images\development-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgWindow.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgAbout.Picture = LoadPicture(App.path & "\Resources\images\about-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'    '
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgGeneralClicked.Picture = LoadPicture(App.path & "\Resources\images\general-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgConfigClicked.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgFontsClicked.Picture = LoadPicture(App.path & "\Resources\images\font-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgSoundsClicked.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgPositionClicked.Picture = LoadPicture(App.path & "\Resources\images\position-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgDevelopmentClicked.Picture = LoadPicture(App.path & "\Resources\images\development-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgWindowClicked.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgAboutClicked.Picture = LoadPicture(App.path & "\Resources\images\about-icon-dark-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'
'    #Else
        
        Set imgGeneral.Picture = Cairo.ImageList("general-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
    
        Set imgConfig.Picture = Cairo.ImageList("config-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
'        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
        
        Set imgFonts.Picture = Cairo.ImageList("font-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgSounds.Picture = Cairo.ImageList("sounds-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgPosition.Picture = Cairo.ImageList("position-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgDevelopment.Picture = Cairo.ImageList("development-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgWindow.Picture = Cairo.ImageList("windows-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgAbout.Picture = Cairo.ImageList("about-icon-dark").Picture
        ' fGauge.RotateBusyTimer = True
        
    '
        Set imgGeneralClicked.Picture = Cairo.ImageList("general-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-dark-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
    
'    #End If

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesDark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesDark of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesLight
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : set the bright images for the bright 'modern' theme
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesLight()
    
    On Error GoTo setPrefsIconImagesLight_Error
    
'    #If TWINBASIC Then
'
'        Set imgGeneral.Picture = LoadPicture(App.path & "\Resources\images\general-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgFonts.Picture = LoadPicture(App.path & "\Resources\images\font-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgSounds.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgPosition.Picture = LoadPicture(App.path & "\Resources\images\position-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgDevelopment.Picture = LoadPicture(App.path & "\Resources\images\development-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgWindow.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgAbout.Picture = LoadPicture(App.path & "\Resources\images\about-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgGeneralClicked.Picture = LoadPicture(App.path & "\Resources\images\general-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgConfigClicked.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgFontsClicked.Picture = LoadPicture(App.path & "\Resources\images\font-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgSoundsClicked.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgPositionClicked.Picture = LoadPicture(App.path & "\Resources\images\position-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgDevelopmentClicked.Picture = LoadPicture(App.path & "\Resources\images\development-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgWindowClicked.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'        Set imgAboutClicked.Picture = LoadPicture(App.path & "\Resources\images\about-icon-light-600-clicked.jpg")
'        ' fGauge.RotateBusyTimer = True
'
'
'    #Else
        
        Set imgGeneral.Picture = Cairo.ImageList("general-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgConfig.Picture = Cairo.ImageList("config-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        
'        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-1010.jpg")
'        ' fGauge.RotateBusyTimer = True
        
        Set imgFonts.Picture = Cairo.ImageList("font-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgSounds.Picture = Cairo.ImageList("sounds-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgPosition.Picture = Cairo.ImageList("position-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgDevelopment.Picture = Cairo.ImageList("development-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgWindow.Picture = Cairo.ImageList("windows-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgAbout.Picture = Cairo.ImageList("about-icon-light").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgGeneralClicked.Picture = Cairo.ImageList("general-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
        
        Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-light-clicked").Picture
        ' fGauge.RotateBusyTimer = True
            
'    #End If
        
   On Error GoTo 0
   Exit Sub

setPrefsIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form widgetPrefs"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsMonitorSaveHeight_Timer
' Author    : beededea
' Date      : 26/08/2024
' Purpose   : save the current height of this form to allow resizing when restarting or placing on another monitor
'---------------------------------------------------------------------------------------
'
Private Sub tmrPrefsMonitorSaveHeight_Timer()

    'Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo tmrPrefsMonitorSaveHeight_Timer_Error
    
    If widgetPrefs.IsVisible = False Then Exit Sub

    If LTrim$(gblMultiMonitorResize) <> "2" Then Exit Sub
 
    If prefsMonitorStruct.IsPrimary = True Then
        gblPrefsPrimaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
        sPutINISetting "Software\UBoatStopWatch", "prefsPrimaryHeightTwips", gblPrefsPrimaryHeightTwips, gblSettingsFile
    Else
        gblPrefsSecondaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
        sPutINISetting "Software\UBoatStopWatch", "prefsSecondaryHeightTwips", gblPrefsSecondaryHeightTwips, gblSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

tmrPrefsMonitorSaveHeight_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsMonitorSaveHeight_Timer of Form widgetPrefs"

End Sub





'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\




