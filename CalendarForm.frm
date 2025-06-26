VERSION 5.00
Begin VB.Form CalendarForm 
   Caption         =   "Select Date"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2745
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CalendarForm
' Developed by Trevor Eyre
' trevoreyre@gmail.com
' v1.5.2 - 1.7.2016
'
' This custom date picker can be used by importing the CalendarForm.frm file into
' your VBA project. It is called exclusively through the GetDate function. For
' instructions on how to call on the CalendarForm, skip to the GetDate function
' documentation after the Global Variables section.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Contributors
'
' Graham Mayor, graham@gmayor.com, www.gmayor.com
'   -Fix for userform sizing incorrectly in Word 2013
'
' Greg Maxey, gmaxey@mvps.org, gregmaxey.mvps.org/word_tips.htm
'   -Fix for leap year bug on years divisible by 100
'   -Moved all initialization code to separate sub to simplify GetDate function
'   -Various code reorganization and optimizations
'
' Marc Meketon, marc.meketon@gmail.com
'   -Fully qualify "Control" declarations as "MSForms.Control" for compatibility
'       with Access
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Overview
'
' The goal in creating this form was first and foremost to overcome the monstrosity that
' is the Microsoft MonthView control. If you're reading this, you probably already know
' what I'm talking about. Many others have been in my place and have come up with their
' own date pickers to solve this problem. So why yet another custom date picker?
'
' I was most interested in the following features:
'   -Ease of use. I wanted a completely self-contained form that could be imported into
'       any VBA project and used without any additional coding.
'   -Simple, attractive design. While a lot of custom date pickers on the internet look
'       good and work well, none of them quite nailed it for me in terms of style and
'       UI design.
'   -Fully customizable functionality and look. I tried to include as many of the
'       options from the MonthView control as I could, without getting too messy.
'
' Since none of the date pickers I have been able to find in all my searching have quite
' completed my checklist, here we are! Now my hope is that some other tired soul may
' also benefit from my labors.
'
' If you encounter any bugs, or have any great ideas or feature requests that could
' improve this bad boy, please send me an email.
'
' What's new in v1.5.2:
'   -Bug fix: Userform not sizing properly in Word 2013
'   -Bug fix: Minimum font size not being preserved correctly
'   -Bug fix: Replaced WorksheetFunction.Max with custom Max function for compatibility
'       with other Office programs
'
' (For a full changelog, known bugs, and future enhancements, scroll to the bottom)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Global Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'These two Enums are used in the GetDate function for the user to select the start day
'of the week, and the behavior of the week numbers. These are used in place of the
'Excel constants vbDayOfWeek and vbFirstWeekOfYear in order to avoid dealing with
'system time, which is an option in both of those. Otherwise the values are identical.
Public Enum calDayOfWeek
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum

Public Enum calFirstWeekOfYear      'Controls how the week numbers are calculated and displayed
    FirstJan1 = 1                   'The week with January 1st is always counted as week 1
    FirstFourDays = 2               'The first week in January that has at least four days in it is
                                        'counted as week 1. This calculation will change depending
                                        'on the setting used for first day of the week. The ISO
                                        'standard is calculating week 1 as the first week in January
                                        'with four days with Monday being the first day of the week.
    FirstFullWeek = 3               'The first week in January with a full week is counted as week 1.
                                        'Like the FirstFourDays setting, this calculation will change
                                        'depending on the first day of the week used.
End Enum

Private UserformEventsEnabled As Boolean    'Controls userform events
Private DateOut As Date                     'The date returned from the CalendarForm
Private SelectedDateIn As Date              'The initial selected date, as well as the date currently selected by the
                                                'user if the Okay button is enabled
Private OkayEnabled As Boolean              'Stores whether Okay button is enabled
Private TodayEnabled As Boolean             'Stores whether Today button is enabled
Private MinDate As Date                     'Minimum date set by user
Private MaxDate As Date                     'Maximum date set by user
Private cmbYearMin As Long                  'Current lower bounds of year combobox. Not necessarily restricted to this min
Private cmbYearMax As Long                  'Current upper bounds of year combobox. Not necessarily restricted to this max
Private StartWeek As VbDayOfWeek            'First day of week in calendar
Private WeekOneOfYear As VbFirstWeekOfYear  'First week of year when setting week numbers
Private HoverControlName As String          'Name of the date label that is currently being hovered over. Used when returning
                                                'the hovered control to its original color
Private HoverControlColor As Long           'Original color of the date label that is currently being hovered over
Private RatioToResize As Double             'Ratio to resize elements of userform. This is set by the DateFontSize argument
                                                'in the GetDate function
Private bgDateColor As Long                 'Color of date label backgrounds
Private bgDateHoverColor As Long            'Color of date label backgrounds when hovering over
Private bgDateSelectedColor As Long         'Color of selected date label background
Private lblDateColor As Long                'Font color of date labels
Private lblDatePrevMonthColor As Long       'Font color of trailing month date labels
Private lblDateTodayColor As Long           'Font color of today's date
Private lblDateSatColor As Long             'Font color of Saturday date labels
Private lblDateSunColor As Long             'Font color of Sunday date labels


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetDate
'
' This function is the point of entry into the CalendarForm. It controls EVERYTHING.
' Every argument is optional, meaning your function call can be as simple as:
'
'   MyDateVariable = CalendarForm.GetDate
'
' That's all there is to it. The calendar initializes, pops up, the user selects a date,
' the selection is received by your variable, and the calendar unloads.
'
' From there, you can use as many or as few arguments as you want in order to get the
' desired calendar that suits your needs. All default values are also set in this
' function, so if you want to change default colors or behavior without having to
' explicitly do so in every function call, you can set those in the argument list
' here.
'
' Below is a list of all arguments, their data type, and their function:
'   SelectedDate (Date) - This is the initial selected date on the calendar. Used to
'       show the users last selection. If this value is set, the calendar will
'       initialize to the month and year of the SelectedDate. If not, it will
'       initialize to today's date (with no selection).
'   FirstDayOfWeek (calDayOfWeek) - Sets which day to use as first day of the week.
'   MinimumDate (Date) - Restricts the selection of any dates below this date.
'   MaximumDate (Date) - Restricts the selection of any dates above this date.
'   RangeOfYears (Long) - Sets the range of years to show in the year combobox in
'       either direction from the initial SelectedDate. For example, if the
'       SelectedDate is in 2014, and the RangeOfYears is set to 10 (the default value),
'       the year combobox will show 10 years below 2014 to 10 years above 2014, so it
'       will have a range of 2004-2024. Note that if this range falls outside the bounds
'       set by the MinimumDate or MaximumDate, it will be overridden. Also, this
'       range does NOT limit the years that a user can select. If the upper limit of
'       the year combobox is 2024, and the user clicks the month spinner to surpass
'       December 2024, it will keep right on going to 2025 and beyond (and those
'       years will be added to the year combobox).
'   DateFontSize (Long) - Controls the size of the CalendarForm. This value cannot
'       be set below 9 (the default). To make the form bigger, set this value larger,
'       and everything else in the userform will be resized to fit.
'   TodayButton (Boolean) - Controls whether or not the Today button is visible.
'   OkayButton (Boolean) - Controls whether or not the Okay button is visible. If the
'       Okay button is enabled, when the user selects a date, it is highlighted, but
'       is not returned until they click Okay. If the Okay button is disabled,
'       clicking a date will automatically return that date and unload the form.
'   ShowWeekNumbers (Boolean) - Controls the visibility of the week numbers.
'   FirstWeekOfYear (calFirstWeekOfYear) - Sets the behavior of the week numbers. See
'       the calFirstWeekOfYear Enum in the Global Variables section to see the possible
'       values and their behavior.
'   PositionTop (Long) - Sets the top position of the CalendarForm. If no value is
'       assigned, the CalendarForm is set to position 1 - CenterOwner. Note that
'       PositionTop and PositionLeft must BOTH be set in order to override the default
'       center position.
'   PositionLeft (Long) - Sets the left position of the CalendarForm. If no value is
'       assigned, the CalendarForm is set to position 1 - CenterOwner. Note that
'       PositionTop and PositionLeft must BOTH be set in order to override the default
'       center position.
'   BackgroundColor (Long) - Sets the background color of the CalendarForm.
'   HeaderColor (Long) - Sets the background color of the header. The header is the
'       month and year label at the top.
'   HeaderFontColor (Long) - Sets the color of the header font.
'   SubHeaderColor (Long) - Sets the background color of the subheader. The subheader
'       is the day of week labels under the header (Su, Mo, Tu, etc).
'   SubHeaderFontColor (Long) - Sets the color of the subheader font.
'   DateColor (Long) - Sets the background color of the individual date labels.
'   DateFontColor (Long) - Sets the font color of the individual date labels.
'   SaturdayFontColor (Long) - Sets the font color of Saturday date labels.
'   SundayFontColor (Long) - Sets the font color of Sunday date labels.
'   DateBorder (Boolean) - Controls whether or not the date labels have borders.
'   DateBorderColor (Long) - Sets the color of the date label borders. Note that the
'       argument DateBorder must be set to True for this setting to take effect.
'   DateSpecialEffect (fmSpecialEffect) - Sets a special effect for the date labels.
'       This can be set to bump, etched, flat (default value), raised, or sunken.
'       This can be used to make the date labels look like buttons if you desire.
'       Note that this setting overrides any date border settings you have made.
'   DateHoverColor (Long) - Sets the background color when hovering the mouse over
'       a date label.
'   DateSelectedColor (Long) - Sets the background color of the selected date.
'   TrailingMonthFontColor (Long) - Sets the color of the date labels in trailing
'       months. Trailing months are the date labels from last month at the top of the
'       calendar and from next month at the bottom of the calendar.
'   TodayFontColor (Long) - Sets the font color of today's date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDate(Optional SelectedDate As Date = 0, _
    Optional FirstDayOfWeek As calDayOfWeek = Sunday, _
    Optional MinimumDate As Date = 0, _
    Optional MaximumDate As Date = 0, _
    Optional RangeOfYears As Long = 10, _
    Optional DateFontSize As Long = 9, _
    Optional TodayButton As Boolean = False, Optional OkayButton As Boolean = False, _
    Optional ShowWeekNumbers As Boolean = False, Optional FirstWeekOfYear As calFirstWeekOfYear = FirstJan1, _
    Optional PositionTop As Long = -5, Optional PositionLeft As Long = -5, _
    Optional BackgroundColor As Long = 16777215, _
    Optional HeaderColor As Long = 15658734, _
    Optional HeaderFontColor As Long = 0, _
    Optional SubHeaderColor As Long = 16448250, _
    Optional SubHeaderFontColor As Long = 8553090, _
    Optional DateColor As Long = 16777215, _
    Optional DateFontColor As Long = 0, _
    Optional SaturdayFontColor As Long = 0, _
    Optional SundayFontColor As Long = 0, _
    Optional DateBorder As Boolean = False, Optional DateBorderColor As Long = 15658734, _
    Optional DateSpecialEffect As fmSpecialEffect = fmSpecialEffectFlat, _
    Optional DateHoverColor As Long = 15658734, _
    Optional DateSelectedColor As Long = 14277081, _
    Optional TrailingMonthFontColor As Long = 12566463, _
    Optional TodayFontColor As Long = 15773696) As Date
    
    'Set global variables
    DateFontSize = Max(DateFontSize, 9) 'Font size cannot be below 9
    OkayEnabled = OkayButton
    TodayEnabled = TodayButton
    RatioToResize = DateFontSize / 9
    bgDateColor = DateColor
    lblDateColor = DateFontColor
    lblDateSatColor = SaturdayFontColor
    lblDateSunColor = SundayFontColor
    bgDateHoverColor = DateHoverColor
    bgDateSelectedColor = DateSelectedColor
    lblDatePrevMonthColor = TrailingMonthFontColor
    lblDateTodayColor = TodayFontColor
    StartWeek = FirstDayOfWeek
    WeekOneOfYear = FirstWeekOfYear
    
    'Initialize userform
    UserformEventsEnabled = False
    Call InitializeUserform(SelectedDate, MinimumDate, MaximumDate, RangeOfYears, PositionTop, PositionLeft, _
        DateFontSize, ShowWeekNumbers, BackgroundColor, HeaderColor, HeaderFontColor, SubHeaderColor, _
        SubHeaderFontColor, DateBorder, DateBorderColor, DateSpecialEffect)
    UserformEventsEnabled = True
    
    'Show userform, return selected date, and unload
    Me.Show
    GetDate = DateOut
    Unload Me
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' InitializeUserform
'
' This sub initializes the size and positions of every element on the userform.
' Everything is sized based on the RatioToResize variable. RatioToResize is calculated
' based on the ratio of the font size passed to the GetDate function to the default
' font size.
'
' The visibility of the Okay button, Today button, and week numbers is also set here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeUserform(SelectedDate As Date, MinimumDate As Date, MaximumDate As Date, _
    RangeOfYears As Long, _
    PositionTop As Long, PositionLeft As Long, _
    SizeFont As Long, bWeekNumbers As Boolean, _
    BackgroundColor As Long, _
    HeaderColor As Long, _
    HeaderFontColor As Long, _
    SubHeaderColor As Long, _
    SubHeaderFontColor As Long, _
    DateBorder As Boolean, DateBorderColor As Long, _
    DateSpecialEffect As fmSpecialEffect)
    
    Dim TempDate As Date                        'Used to set selected date, if none has been provided
    Dim SelectedYear As Long                    'Year of selected date
    Dim SelectedMonth As Long                   'Month of selected date
    Dim SelectedDay As Long                     'Day of seledcted date (if applicable)
    Dim TempDayOfWeek As Long                   'Used to set day labels in subheader
    Dim BorderSpacing As Double                 'Padding between the outermost elements of userform and edge of userform
    Dim HeaderDefaultFontSize As Long           'Default font size of the header labels (month and year)
    Dim bgHeaderDefaultHeight As Double         'Default height of the background behind header labels
    Dim lblMonthYearDefaultHeight As Double     'Default height of the month and year header labels
    Dim scrlMonthDefaultHeight As Double        'Default height of the month scroll bar
    Dim bgDayLabelsDefaultHeight As Double      'Default height of the background behind the subheader day of week labels
    Dim bgDateDefaultHeight As Double           'Default height of the background behind each date label
    Dim bgDateDefaultWidth As Double            'Default width of the background behind each date label
    Dim lblDateDefaultHeight As Double          'Default height of each date label
    Dim cmdButtonDefaultHeight As Double        'Default height of Today and Okay command buttons
    Dim cmdButtonDefaultWidth As Double         'Default width of Today and Okay command buttons
    Dim cmdButtonsCombinedWidth As Double       'Combined width of Today and Okay buttons. Used to center on userform
    Dim cmdButtonsMaxHeight As Double           'Maximum height of command buttons and month scroll bar
    Dim cmdButtonsMaxWidth As Double            'Maximum width of command buttons
    Dim cmdButtonsMaxFontSize As Long           'Maximum font size of command buttons
    Dim bgControl As MSForms.Control            'Stores current date label background in loop to initialize various settings
    Dim lblControl As MSForms.Control           'Stores current date label in loop to initialize various settings
    Dim HeightOffset As Double                  'Difference between form height and inside height, to account for toolbar
    Dim i As Long                               'Used for loops
    Dim j As Long                               'Used for loops
    
    'Initialize default values
    BorderSpacing = 6 * RatioToResize
    HeaderDefaultFontSize = 11
    bgHeaderDefaultHeight = 30
    lblMonthYearDefaultHeight = 13.5
    scrlMonthDefaultHeight = 18
    bgDayLabelsDefaultHeight = 18
    bgDateDefaultHeight = 18
    bgDateDefaultWidth = 18
    lblDateDefaultHeight = 10.5
    cmdButtonDefaultHeight = 24
    cmdButtonDefaultWidth = 60
    cmdButtonsMaxHeight = 36
    cmdButtonsMaxWidth = 90
    cmdButtonsMaxFontSize = 14

    
    'Set MinDate and MaxDate. If no MinimumDate or MaximumDate are provided, set the
    'MinDate to 1/1/1900 and the MaxDate to 12/31/9999. If MaxDate is less than
    'MinDate, it will default to the MinDate.
    If MinimumDate <= 0 Then
        MinDate = CDate("1/1/1900")
    Else
        MinDate = MinimumDate
    End If
    If MaximumDate = 0 Then
        MaxDate = CDate("12/31/9999")
    Else
        MaxDate = MaximumDate
    End If
    If MaxDate < MinDate Then MaxDate = MinDate
    
    'If today's date falls outside min/max, make sure Today button is disabled
    If Date < MinDate Or Date > MaxDate Then TodayEnabled = False

    'Initialize userform position. Initial value of top and left is -5. Check
    'this value to see if a different value has been passed. If not, position
    'to CenterOwner. Must set both top and left positions to override center position
    If PositionTop <> -5 And PositionLeft <> -5 Then
        Me.StartUpPosition = 0
        Me.Top = PositionTop
        Me.Left = PositionLeft
    Else
        Me.StartUpPosition = 1
    End If
    
    'Size header elements - header background, month scroll bar, scroll cover (which is just
    'a blank label which sits on top of the month scroll bar to make it look like two spin
    'buttons), month/year labels in header, and the month and year comboboxes
    With bgHeader
        .Height = bgHeaderDefaultHeight * RatioToResize
        'The header width depends on whether week numbers are visible or not
        If bWeekNumbers Then
            .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
        Else
            .Width = 7 * (bgDateDefaultWidth * RatioToResize)
        End If
        .Left = BorderSpacing
        .Top = BorderSpacing
    End With
    'Month scroll bar. I set a maximum height for the scroll bar, because as it gets
    'larger, the width of the scroll buttons never increases, so eventually it ends
    'up looking really tall and skinny and weird.
    With scrlMonth
        .Width = bgHeader.Width - (2 * BorderSpacing)
        .Left = bgHeader.Left + BorderSpacing
        .Height = scrlMonthDefaultHeight * RatioToResize
        If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
        .Top = bgHeader.Top + ((bgHeader.Height - .Height) / 2)
    End With
    'Cover over month scroll bar
    With bgScrollCover
        .Height = scrlMonth.Height
        .Width = scrlMonth.Width - 25 '25 is the width of the actual scroll buttons,
                                      'which need to remain visible
        .Left = scrlMonth.Left + 12.5
        .Top = scrlMonth.Top
    End With
    'The .left position of the month and year labels in the header will be set
    'in the function SetMonthYear, as it changes based on the selected month/year.
    'So only the top needs to be positioned now
    With lblMonth
        .AutoSize = False
        .Height = lblMonthYearDefaultHeight * RatioToResize
        .Font.Size = HeaderDefaultFontSize * RatioToResize
        .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
    End With
    With lblYear
        .AutoSize = False
        .Height = lblMonthYearDefaultHeight * RatioToResize
        .Font.Size = HeaderDefaultFontSize * RatioToResize
        .Top = bgScrollCover.Top + ((bgScrollCover.Height - .Height) / 2)
    End With
    cmbMonth.Top = lblMonth.Top + (lblMonth.Height - cmbMonth.Height)
    cmbYear.Top = lblYear.Top + (lblYear.Height - cmbYear.Height)

    'Size subheader elements - the subheader bacgkround (bgDayLabels), the day of
    'week labels themselves, and the week number subheader label, if applicable
    With bgDayLabels
        .Height = bgDayLabelsDefaultHeight * RatioToResize
        'The width depends on whether week numbers are visible or not
        If bWeekNumbers Then
            .Width = 8 * (bgDateDefaultWidth * RatioToResize) + BorderSpacing
        Else
            .Width = 7 * (bgDateDefaultWidth * RatioToResize)
        End If
        .Left = BorderSpacing
        .Top = bgHeader.Top + bgHeader.Height
    End With
    'Week number subheader label
    If Not bWeekNumbers Then
        lblWk.Visible = False
    Else
        With lblWk
            .AutoSize = False
            .Height = lblDateDefaultHeight * RatioToResize
            .Font.Size = SizeFont
            .Width = bgDateDefaultWidth * RatioToResize
            .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
            .Left = BorderSpacing
        End With
    End If
    'Day of week subheader labels
    For i = 1 To 7
        With Me("lblDay" & CStr(i))
            .AutoSize = False
            .Height = lblDateDefaultHeight * RatioToResize
            .Font.Size = SizeFont
            .Width = bgDateDefaultWidth * RatioToResize
            .Top = bgDayLabels.Top + ((bgDayLabels.Height - .Height) / 2)
            If i = 1 Then
                'Left position of first label depends on whether week numbers are visible
                If bWeekNumbers Then
                    .Left = lblWk.Left + lblWk.Width + BorderSpacing
                Else
                    .Left = BorderSpacing
                End If
            Else 'All other labels placed directly next to preceding label
                .Left = Me("lblDay" & CStr(i - 1)).Left + Me("lblDay" & CStr(i - 1)).Width
            End If
        End With
    Next i
    
    'Size all date labels and backgrounds
    For i = 1 To 6 'Rows
        'First set position and visibility of week number label
        If Not bWeekNumbers Then
            Me("lblWeek" & CStr(i)).Visible = False
        Else
            With Me("lblWeek" & CStr(i))
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.Size = SizeFont
                .Width = bgDateDefaultWidth * RatioToResize
                .Left = BorderSpacing
                If i = 1 Then
                    .Top = bgDayLabels.Top + bgDayLabels.Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                Else
                    .Top = Me("bgDate" & CStr(i - 1) & "1").Top + Me("bgDate" & CStr(i - 1) & "1").Height + (((bgDateDefaultHeight * RatioToResize) - .Height) / 2)
                End If
            End With
        End If
        
        'Now set position of each date label in current row
        For j = 1 To 7
            Set bgControl = Me("bgDate" & CStr(i) & CStr(j))
            Set lblControl = Me("lblDate" & CStr(i) & CStr(j))
            'The date label background is sized and placed first. Then the actual date label is simply
            'set to the same position and centered vertically.
            With bgControl
                .Height = bgDateDefaultHeight * RatioToResize
                .Width = bgDateDefaultWidth * RatioToResize
                If j = 1 Then
                    'Left position of first label in row depends on whether week numbers are visible
                    If bWeekNumbers Then
                        .Left = Me("lblWeek" & CStr(i)).Left + Me("lblWeek" & CStr(i)).Width + BorderSpacing
                    Else
                        .Left = BorderSpacing
                    End If
                Else 'All other labels placed directly next to preceding label in row
                    .Left = Me("bgDate" & CStr(i) & CStr(j - 1)).Left + Me("bgDate" & CStr(i) & CStr(j - 1)).Width
                End If
                If i = 1 Then
                    .Top = bgDayLabels.Top + bgDayLabels.Height
                Else
                    .Top = Me("bgDate" & CStr(i - 1) & CStr(j)).Top + Me("bgDate" & CStr(i - 1) & CStr(j)).Height
                End If
            End With
            'Size and position actual date label
            With lblControl
                .AutoSize = False
                .Height = lblDateDefaultHeight * RatioToResize
                .Font.Size = SizeFont
                .Width = bgControl.Width
                .Left = bgControl.Left
                .Top = bgControl.Top + ((bgControl.Height - .Height) / 2)
            End With
        Next j
    Next i
    
    'Set userform width. Height set later, since it depends on Today and Okay buttons
    frameCalendar.Width = bgDate67.Left + bgDate67.Width + BorderSpacing
    'Make sure userform is large enough to show entire calendar
    If Me.InsideWidth < (frameCalendar.Left + frameCalendar.Width) Then
        Me.Width = Me.Width + ((frameCalendar.Left + frameCalendar.Width) - Me.InsideWidth)
    End If

    'Set size and visibility of Okay button and date selection labels
    If Not OkayEnabled Then
        cmdOkay.Visible = False
        lblSelection.Visible = False
        lblSelectionDate.Visible = False
    Else
        'Okay button. I set a maximum and width, for the same reason as the month
        'scroll bar. Eventually, the gigantic buttons just start looking weird.
        With cmdOkay
            .Visible = True
            .Height = cmdButtonDefaultHeight * RatioToResize
            If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
            .Width = cmdButtonDefaultWidth * RatioToResize
            If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
            If SizeFont > cmdButtonsMaxFontSize Then
                .Font.Size = cmdButtonsMaxFontSize
            Else
                .Font.Size = SizeFont
            End If
            .Top = bgDate61.Top + bgDate61.Height + bgDayLabels.Height + BorderSpacing
        End With
        'The "Selection" label
        With lblSelection
            .Visible = True
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            .Width = frameCalendar.Width
            .Font.Size = HeaderDefaultFontSize * RatioToResize
            .AutoSize = True
            .Top = (bgDate61.Top + bgDate61.Height) + ((bgDayLabels.Height + BorderSpacing - .Height) / 2)
        End With
        'The actual selected date label
        With lblSelectionDate
            .Visible = True
            .AutoSize = False
            .Height = lblMonthYearDefaultHeight * RatioToResize
            .Width = frameCalendar.Width - lblSelection.Width
            .Font.Size = HeaderDefaultFontSize * RatioToResize
            .Top = lblSelection.Top
        End With
    End If
    
    'Set size and visibility of Today button. Make sure it is within max bounds.
    'Top is not set for Today button yet, because it depends on whether Okay button
    'is enabled. Therefore, it is set farther down.
    If Not TodayEnabled Then
        cmdToday.Visible = False
    Else
        With cmdToday
            .Visible = True
            .Height = cmdButtonDefaultHeight * RatioToResize
            If .Height > cmdButtonsMaxHeight Then .Height = cmdButtonsMaxHeight
            .Width = cmdButtonDefaultWidth * RatioToResize
            If .Width > cmdButtonsMaxWidth Then .Width = cmdButtonsMaxWidth
            If SizeFont > cmdButtonsMaxFontSize Then
                .Font.Size = cmdButtonsMaxFontSize
            Else
                .Font.Size = SizeFont
            End If
        End With
    End If
    
    'Position Okay and Today buttons, depending on which ones are enabled
    If OkayEnabled And TodayEnabled Then 'Both buttons enabled.
        cmdToday.Top = cmdOkay.Top
        cmdButtonsCombinedWidth = cmdToday.Width + cmdOkay.Width
        cmdToday.Left = ((frameCalendar.Width - cmdButtonsCombinedWidth) / 2) - (BorderSpacing / 2)
        cmdOkay.Left = cmdToday.Left + cmdToday.Width + BorderSpacing
    ElseIf OkayEnabled Then 'Only Okay button enabled
        cmdOkay.Left = (frameCalendar.Width - cmdOkay.Width) / 2
    ElseIf TodayEnabled Then 'Only Today button enabled
        cmdToday.Top = bgDate61.Top + bgDate61.Height + BorderSpacing
        cmdToday.Left = (frameCalendar.Width - cmdToday.Width) / 2
    End If
    
    'Set userform height, depending on which buttons are enabled
    HeightOffset = Me.Height - Me.InsideHeight
    If OkayEnabled Then
        frameCalendar.Height = cmdOkay.Top + cmdOkay.Height + HeightOffset + BorderSpacing
    ElseIf TodayEnabled Then 'Only Today button enabled
        frameCalendar.Height = cmdToday.Top + cmdToday.Height + HeightOffset + BorderSpacing
    Else 'Neither button enabled
        frameCalendar.Height = bgDate61.Top + bgDate61.Height + HeightOffset + BorderSpacing
    End If
    
    'Make sure userform is large enough to show entire calendar
    If Me.InsideHeight < (frameCalendar.Top + frameCalendar.Height) Then
        Me.Height = Me.Height + ((frameCalendar.Top + frameCalendar.Height) - Me.InsideHeight - HeightOffset)
    End If
    
    'Check if SelectedDateIn was set by user, and ensure it is within min/max range
    If SelectedDate > 0 Then
        If SelectedDate < MinDate Then
            SelectedDate = MinDate
        ElseIf SelectedDate > MaxDate Then
            SelectedDate = MaxDate
        End If
        SelectedDateIn = SelectedDate
        SelectedYear = Year(SelectedDateIn)
        SelectedMonth = Month(SelectedDateIn)
        SelectedDay = Day(SelectedDateIn)
        Call SetSelectionLabel(SelectedDateIn)
    Else 'No SelectedDate provided, default to today's date
        cmdOkay.Enabled = False
        TempDate = Date
        If TempDate < MinDate Then
            TempDate = MinDate
        ElseIf TempDate > MaxDate Then
            TempDate = MaxDate
        End If
        SelectedYear = Year(TempDate)
        SelectedMonth = Month(TempDate)
        SelectedDay = 0 'Don't want to highlight a 'selected date,' since user supplied no date
        Call SetSelectionLabel(Empty)
    End If
    
    'Initialize month and year comboboxes, as well as month scroll bar. Make sure
    'years are within range of 1900 to 9999. If year combobox falls outside bounds
    'of MinDate and MaxDate, it will be overridden.
    Call SetMonthCombobox(SelectedYear, SelectedMonth)
    scrlMonth.value = SelectedMonth
    cmbYearMin = SelectedYear - RangeOfYears
    cmbYearMax = SelectedYear + RangeOfYears
    If cmbYearMin < Year(MinDate) Then
        cmbYearMin = Year(MinDate)
    End If
    If cmbYearMax > Year(MaxDate) Then
        cmbYearMax = Year(MaxDate)
    End If
    For i = cmbYearMin To cmbYearMax
        cmbYear.AddItem i
    Next i
    cmbYear.value = SelectedYear
    
    'Set userform colors and effects
    Me.BackColor = BackgroundColor
    frameCalendar.BackColor = BackgroundColor
    bgHeader.BackColor = HeaderColor
    bgScrollCover.BackColor = HeaderColor
    lblMonth.ForeColor = HeaderFontColor
    lblYear.ForeColor = HeaderFontColor
    lblSelection.ForeColor = SubHeaderFontColor
    lblSelectionDate.ForeColor = SubHeaderFontColor
    bgDayLabels.BackColor = SubHeaderColor
    For i = 1 To 7
        Me("lblDay" & CStr(i)).ForeColor = SubHeaderFontColor
    Next i
    If bWeekNumbers Then
        lblWk.ForeColor = SubHeaderFontColor
        For i = 1 To 6
            Me("lblWeek" & CStr(i)).ForeColor = SubHeaderFontColor
        Next i
    End If
    For i = 1 To 6
        For j = 1 To 7
            With Me("bgDate" & CStr(i) & CStr(j))
                If DateBorder Then
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DateBorderColor
                End If
                .SpecialEffect = DateSpecialEffect
            End With
        Next j
    Next i
    
    'Initialize subheader day labels, based on selected first day of week
    TempDayOfWeek = StartWeek
    For i = 1 To 7
        Me("lblDay" & CStr(i)).Caption = Choose(TempDayOfWeek, "Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
        TempDayOfWeek = TempDayOfWeek + 1
        If TempDayOfWeek = 8 Then TempDayOfWeek = 1
    Next i
            
    'Set month and year labels in header, as well as date labels
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmdOkay_Click
'
' When the Okay button is clicked, DateOut is set, and the CalendarForm is hidden to
' return control to the GetDate function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOkay_Click()
    DateOut = SelectedDateIn
    Me.Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmdToday_Click
'
' The functionality of the Today button changes depending on whether the Okay button is
' enabled or not. If the Okay button is enabled, clicking the Today button jumps to
' today's date and selects it.
'
' If the Okay button is disabled, clicking the Today button jumps to today's date, but
' nothing is selected.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdToday_Click()
    Dim SelectedMonth As Long           'Month of selected date
    Dim SelectedYear As Long            'Year of selected date
    Dim SelectedDay As Long             'Day of selected date, if applicable
    Dim TodayDate As Date               'Today's date
    
    UserformEventsEnabled = False
    SelectedDay = 0
    TodayDate = Date
    
    'If Okay button is enabled, set SelectedDateIn, and the selection labels
    If OkayEnabled Then
        cmdOkay.Enabled = True
        SelectedDateIn = TodayDate
        Call SetSelectionLabel(TodayDate)
        SelectedDay = Day(TodayDate)
    End If
    
    'Get the month, day, and year, and set month scroll bar
    SelectedMonth = Month(TodayDate)
    SelectedYear = Year(TodayDate)
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    scrlMonth.value = SelectedMonth
    
    'Set month/year labels and date labels
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserForm_QueryClose
'
' I originally included this sub to override when the user cancelled the
' CalendarForm using the X button, in order to avoid receiving an invalid date value
' back from the userform (1/0/1900 12:00:00 AM). This sub sets DateOut to currently
' selected Date, or to the initial SelectedDate passed to the GetDate function if user
' has not changed the selection, or the Okay button is not enabled.
'
' Note that it is still possible for the CalendarForm to return an invalid date value
' if no initial SelectedDate is set, the user does not make any selection, and then
' cancels the userform.
'
' I ended up removing the sub, because I like being able to detect if the user has
' cancelled the userform by testing the date from it. For instance, if user selects
' a date, but then changes their mind and cancels the userform, you wouldn't want to
' still return that date to your variable. You would want to revert to their previous
' selection, or do some error handling, if necessary.
'
' If you want the functionality described above, of returning the selected date or
' initial date if the user cancels, you can un-comment this sub.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = 0 Then
'        Cancel = True
'        DateOut = SelectedDateIn
'        Me.Hide
'    End If
'End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ClickControl
'
' This sub handles the event of clicking on one of the date label controls. Every date
' label has a click event which passes that label to this sub.
'
' If the Okay button is enabled, clicking a date selects that date, but does not return.
' If Okay button is disabled, clicking a date hides the userform and returns that date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClickControl(ctrl As MSForms.Control)
    Dim SelectedMonth As Long           'Month of selected date
    Dim SelectedYear As Long            'Year of selected date
    Dim SelectedDay As Long             'Day of selected date
    Dim SelectedDate As Date            'Date that the user has selected
    Dim RowIndex As Long                'Row index of the clicked date label
    Dim ColumnIndex As Long             'Column index of the clicked date label
    
    'Get selected day/year from scroll bar and combobox
    SelectedMonth = scrlMonth.value
    SelectedYear = cmbYear.value
    
    'Get indices of date label from label name and selected day from caption
    RowIndex = CLng(Left(Right(ctrl.Name, 2), 1))
    ColumnIndex = CLng(Right(ctrl.Name, 1))
    SelectedDay = CLng(ctrl.Caption)
    
    'Selection is from previous month. The largest day that could exist in
    'the first row from the current month is 6, so if the day is larger than
    'that, we know it came from the previous month, in which case we need
    'to decrement the selected month
    If RowIndex = 1 And SelectedDay > 7 Then
        SelectedMonth = SelectedMonth - 1
        'Handle January
        If SelectedMonth = 0 Then
            SelectedYear = SelectedYear - 1
            SelectedMonth = 12
        End If
    
    'Selection is from next month. The trailing dates from next month can
    'show up in rows 5 and 6. The smallest day that could exist in these rows
    'from the current month is about 23, so if the day is smaller than that,
    'we know it came from next month.
    ElseIf RowIndex >= 5 And SelectedDay < 20 Then
        SelectedMonth = SelectedMonth + 1
        'Handle December
        If SelectedMonth = 13 Then
            SelectedYear = SelectedYear + 1
            SelectedMonth = 1
        End If
    End If
    
    SelectedDate = DateSerial(SelectedYear, SelectedMonth, SelectedDay)
    
    'If Okay button is disabled, click will automatically hide form to return selected
    'date. If Okay button is enabled, click will select date, but will not return until
    'Okay is clicked
    If Not OkayEnabled Then
        DateOut = SelectedDate
        Me.Hide
    Else
        UserformEventsEnabled = False
            cmdOkay.Enabled = True
            SelectedDateIn = SelectedDate
            scrlMonth.value = SelectedMonth
            Call SetSelectionLabel(SelectedDate)
            Call SetMonthYear(SelectedMonth, SelectedYear)
            Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
        UserformEventsEnabled = True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HoverControl
'
' This sub handles the event of hovering over one of the date label controls. Every date
' label has a MouseMove event which passes that label to this sub.
'
' This sub returns the last hovered date label to its original color, sets the currently
' hovered date label to the bgDateHoverColor, and stores its name and original color
' to global variables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HoverControl(ctrl As MSForms.Control)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
    HoverControlName = ctrl.Name
    HoverControlColor = ctrl.BackColor
    ctrl.BackColor = bgDateHoverColor
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' lblMonth_Click / lblYear_Click
'
' The month and year labels in the header have invisible comboboxes behind them. These
' two subs show the combobox drop downs when you click on the labels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lblMonth_Click()
    cmbMonth.DropDown
End Sub
Private Sub lblYear_Click()
    cmbYear.DropDown
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmbMonth_Change / cmbYear_Change
'
' The month and year comboboxes both call the cmbMonthYearChange sub when the user makes
' a selection. The year combobox also resets the month combobox, in case the user
' selects a year that is limited by a minimum or maximum date, to make sure the month
' combobox doesn't end up with selections that shouldn't be available.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbMonth_Change()
    Call cmbMonthYearChange
End Sub
Private Sub cmbYear_Change()
    If Not UserformEventsEnabled Then Exit Sub
    
    UserformEventsEnabled = False
    Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    UserformEventsEnabled = True
    
    Call cmbMonthYearChange
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' cmbMonthYearChange
'
' This sub handles the user making a selection from either the month or year combobox.
' It gets the selected month and year from the comboboxes, sets the value of the month
' scroll bar to match, and resets the calendar date labels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmbMonthYearChange()
    Dim SelectedMonth As Long           'Month of selected date
    Dim SelectedYear As Long            'Year of selected date
    Dim SelectedDay As Long             'Day of selected date
    
    If Not UserformEventsEnabled Then Exit Sub
    UserformEventsEnabled = False
    
    'Get selected month and year. If the selected year has a minimum date set, then
    'the month combobox might not contain all the months of the year. In this case
    'the combobox index has to be offset by the month of the minimum date. No
    'calculation is necessary if the selected year has a maximum date set, because
    'the indices of the months in the combobox are still going to be the same in
    'either case.
    SelectedYear = cmbYear.value
    If SelectedYear = Year(MinDate) Then
        SelectedMonth = cmbMonth.ListIndex + Month(MinDate)
    Else
        SelectedMonth = cmbMonth.ListIndex + 1
    End If
    
    'Get selected day, set the value of the month scroll bar, and reset all
    'date labels on the userform
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    scrlMonth.value = SelectedMonth
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' scrlMonth_Change
'
' This sub handles the user clicking the scroll bar to increment or decrement the month.
' It checks to keep the month within the bounds set by the minimum or maximum date,
' and resets all the labels of the userform to the new month.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub scrlMonth_Change()
    Dim TempYear As Long        'Temporarily store selected year to test min and max dates
    Dim MinMonth As Long        'Sets lower limit of scroll bar
    Dim MaxMonth As Long        'Sets upper limit of scroll bar
    Dim SelectedMonth As Long   'Month of selected date
    Dim SelectedYear As Long    'Year of selected date
    Dim SelectedDay As Long     'Day of selected date
    
    If Not UserformEventsEnabled Then Exit Sub
    UserformEventsEnabled = False
    
    'Default lower and upper limit of scroll bar to allow full range of months
    MinMonth = 0
    MaxMonth = 13
    
    'If the current year is the min or max year, set min or max months
    TempYear = cmbYear.value
    If TempYear = Year(MinDate) Then MinMonth = Month(MinDate)
    If TempYear = Year(MaxDate) Then MaxMonth = Month(MaxDate)
    
    'Keep scroll bar within range of min and max dates
    If scrlMonth.value < MinMonth Then scrlMonth.value = scrlMonth.value + 1
    If scrlMonth.value > MaxMonth Then scrlMonth.value = scrlMonth.value - 1
    
    'If user goes down one month from January, scroll bar will have value of
    '0. In this case, reset scroll bar back to December and decrement year
    'by 1.
    If scrlMonth.value = 0 Then
        scrlMonth.value = 12
        cmbYear.value = cmbYear.value - 1
        'If new year is outside range of combobox, add it to combobox
        If cmbYear.value < cmbYearMin Then
            cmbYear.AddItem cmbYear.value, 0
            cmbYearMin = cmbYear.value
        End If
        Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    'If user goes up one month from December, scroll bar will have value of
    '13. Reset to January and increment year.
    ElseIf scrlMonth.value = 13 Then
        scrlMonth.value = 1
        cmbYear.value = cmbYear.value + 1
        'If new year is outside range of combobox, add it to combobox
        If cmbYear.value > cmbYearMax Then
            cmbYear.AddItem cmbYear.value, cmbYear.ListCount
            cmbYearMax = cmbYear.value
        End If
        Call SetMonthCombobox(cmbYear.value, scrlMonth.value)
    End If
    
    'Get selected month, year, and day, and reset all userform labels
    SelectedMonth = scrlMonth.value
    SelectedYear = cmbYear.value
    SelectedDay = GetSelectedDay(SelectedMonth, SelectedYear)
    Call SetMonthYear(SelectedMonth, SelectedYear)
    Call SetDays(SelectedMonth, SelectedYear, SelectedDay)
    
    UserformEventsEnabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetMonthCombobox
'
' This sub clears the list in the month combobox and resets it. This is done every time
' the month changes to make sure the months displayed in the combobox don't ever fall
' outside the bounds set by the minimum or maximum date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMonthCombobox(YearIn As Long, MonthIn As Long)
    Dim YearMinDate As Long             'Year of the minimum date
    Dim YearMaxDate As Long             'Year of the maximum date
    Dim MonthMinDate As Long            'Month of the minimum date
    Dim MonthMaxDate As Long            'Month of the maximum date
    Dim i As Long                       'Used for looping
    
    'Get month and year of minimum and maximum dates and clear combobox
    YearMinDate = Year(MinDate)
    YearMaxDate = Year(MaxDate)
    MonthMinDate = Month(MinDate)
    MonthMaxDate = Month(MaxDate)
    cmbMonth.Clear

    'Both minimum and maximum dates occur in selected year
    If YearIn = YearMinDate And YearIn = YearMaxDate Then
        For i = MonthMinDate To MonthMaxDate
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn < MonthMinDate Then MonthIn = MonthMinDate
        If MonthIn > MonthMaxDate Then MonthIn = MonthMaxDate
        cmbMonth.ListIndex = MonthIn - MonthMinDate
    
    'Only minimum date occurs in selected year
    ElseIf YearIn = YearMinDate Then
        For i = MonthMinDate To 12
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn < MonthMinDate Then MonthIn = MonthMinDate
        cmbMonth.ListIndex = MonthIn - MonthMinDate
    
    'Only maximum date occurs in selected year
    ElseIf YearIn = YearMaxDate Then
        For i = 1 To MonthMaxDate
            cmbMonth.AddItem Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        Next i
        If MonthIn > MonthMaxDate Then MonthIn = MonthMaxDate
        cmbMonth.ListIndex = MonthIn - 1
    
    'No minimum or maximum date in selected year. Add all months to combobox
    Else
        cmbMonth.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        cmbMonth.ListIndex = MonthIn - 1
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetMonthYear
'
' This sub sets the month and year comboboxes to keep them in sync with any changes
' made to the selected month or year. It also sets the month and year labels in the
' header, and positions them in the center of the month scroll bar.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetMonthYear(MonthIn As Long, YearIn As Long)
    Dim ExtraSpace As Double                'Space between month and year labels
    Dim CombinedLabelWidth As Double        'Combined width of both month and year labels
    
    ExtraSpace = 4 * RatioToResize
    
    'Set value of comboboxes
    If YearIn = Year(MinDate) Then
        cmbMonth.ListIndex = MonthIn - Month(MinDate)
    Else
        cmbMonth.ListIndex = MonthIn - 1
    End If
    cmbYear.value = YearIn
    
    'Set labels and position to center of scroll buttons. Labels are first
    'set to the width of the userform to avoid overflow, and then autosized
    'to fit to the text before being centered
    With lblMonth
        .AutoSize = False
        .Width = frameCalendar.Width
        .Caption = cmbMonth.value
        .AutoSize = True
    End With
    With lblYear
        .AutoSize = False
        .Width = frameCalendar.Width
        .Caption = cmbYear.value
        .AutoSize = True
    End With
    
    'Get combined width of labels and center to scroll bar
    CombinedLabelWidth = lblMonth.Width + lblYear.Width
    With lblMonth
        .Left = ((frameCalendar.Width - CombinedLabelWidth) / 2) - (ExtraSpace / 2)
    End With
    With lblYear
        .Left = lblMonth.Left + lblMonth.Width + ExtraSpace
    End With
    
    'Reposition comboboxes to line up with labels
    cmbMonth.Left = lblMonth.Left - (cmbMonth.Width - lblMonth.Width) - ExtraSpace - 2
    cmbYear.Left = lblYear.Left
    
    'Clear hover control name, so labels in new month don't revert to
    'colors from previously selected month
    HoverControlName = vbNullString
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetDays
'
' This sub sets the caption, visibility, and colors of all the date labels on the
' userform, as well as the week number labels. If a selected day is passed to the
' sub, it will highlight that date accordingly. Otherwise, no selected date will be
' highlighted.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDays(MonthIn As Long, YearIn As Long, Optional DayIn As Long)
    Dim PrevMonth As Long               'Month preceding selected month. Used for trailing dates
    Dim NextMonth As Long               'Month following selected month. Used for trailing dates
    Dim Today As Date                   'Today's date
    Dim TodayDay As Long                'Day number of today's date
    Dim StartDayOfWeek  As Long         'Stores the weekday number of the first day in selected month
    Dim LastDayOfMonth As Long          'Last day of the month
    Dim LastDayOfPrevMonth As Long      'Last day of preceding month. Used for trailing dates
    Dim CurrentDay As Long              'Tracks current day in the month while setting labels
    Dim TempCurrentDay As Long          'Tracks the current day for previous month without incrementing actual CurrentDay
    Dim WeekNumber As Long              'Stores week number for week number labels
    Dim StartDayOfWeekDate As Date      'Stores first date in the week. Used to calculate week numbers
    Dim SaturdayIndex As Long           'Column index of Saturdays. Used to set color of Saturday labels, if applicable
    Dim SundayIndex As Long             'Column index of Sundays
    Dim MinDay As Long                  'Stores lower limit of days if minimum date falls in selected month
    Dim MaxDay As Long                  'Stores upper limit of days if maximum date falls in selected month
    Dim PrevMonthMinDay As Long         'Stores lower limit of days if minimum date falls in preceding month
    Dim NextMonthMaxDay As Long         'Stores upper limit of days if maximum date falls in next month
    Dim lblControl As MSForms.Control   'Stores current date label while changing settings
    Dim bgControl As MSForms.Control    'Stores current date label background while changing settings
    Dim i As Long                       'Used for looping
    Dim j As Long                       'Used for looping
    
    'Set min and max day, if applicable. If not, min and max day are set to 0 and 32,
    'respectively, since dates will never fall outside those bounds
    MinDay = 0
    MaxDay = 32
    If YearIn = Year(MinDate) And MonthIn = Month(MinDate) Then MinDay = Day(MinDate)
    If YearIn = Year(MaxDate) And MonthIn = Month(MaxDate) Then MaxDay = Day(MaxDate)
    
    'Find previous month and next month. Handle January
    'and December appropriately
    PrevMonth = MonthIn - 1
    If PrevMonth = 0 Then PrevMonth = 12
    NextMonth = MonthIn + 1
    If NextMonth = 13 Then NextMonth = 1
    
    'Set min and max days for previous month and next month, if applicable
    PrevMonthMinDay = 0
    NextMonthMaxDay = 32
    If YearIn = Year(MinDate) And PrevMonth = Month(MinDate) Then PrevMonthMinDay = Day(MinDate)
    If YearIn = Year(MaxDate) And NextMonth = Month(MaxDate) Then NextMonthMaxDay = Day(MaxDate)

    'Find last day of selected month and previous month. Find first weekday
    'in current month, and index of Saturday and Sunday relative to first weekday
    LastDayOfMonth = Day(DateSerial(YearIn, MonthIn + 1, 0))
    LastDayOfPrevMonth = Day(DateSerial(YearIn, MonthIn, 0))
    StartDayOfWeek = Weekday(DateSerial(YearIn, MonthIn, 1), StartWeek)
    If StartWeek = 1 Then SundayIndex = 1 Else SundayIndex = 9 - StartWeek
    SaturdayIndex = 8 - StartWeek

    'If user is viewing current month/year, we want to highlight today's date. If
    'not, TodayDay is set to 0, since that value will never be encountered
    Today = Date
    If YearIn = Year(Today) And MonthIn = Month(Today) Then
        TodayDay = Day(Today)
    Else
        TodayDay = 0
    End If
    
    'Loop through all date labels and set captions and colors
    CurrentDay = 1
    For i = 1 To 6 'Rows
    
        'Set week number first, as it happens only once per row
        'Entire first row is last month
        If StartDayOfWeek = 1 And i = 1 Then
            'Calculate day number of first day in the week
            TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek + 5))
            If PrevMonth <> 12 Then
                StartDayOfWeekDate = DateSerial(YearIn, PrevMonth, TempCurrentDay)
            Else
                StartDayOfWeekDate = DateSerial(YearIn - 1, PrevMonth, TempCurrentDay)
            End If
            
        'Previous month, but entire row is not last month. In this
        'case just use first of month. This is done because when using
        'the DatePart function to calculate week number, the last week
        'in December can be calculated incorrectly, so we want to default
        'to January 1st instead, which is always correct
        ElseIf i = 1 Then
            StartDayOfWeekDate = DateSerial(YearIn, MonthIn, 1)
        
        Else
            'Current month
            If CurrentDay <= LastDayOfMonth Then
                TempCurrentDay = CurrentDay
                StartDayOfWeekDate = DateSerial(YearIn, MonthIn, TempCurrentDay)
            
            'Next month
            Else
                TempCurrentDay = CLng(CurrentDay - LastDayOfMonth)
                If NextMonth <> 1 Then
                    StartDayOfWeekDate = DateSerial(YearIn, NextMonth, TempCurrentDay)
                Else
                    StartDayOfWeekDate = DateSerial(YearIn + 1, NextMonth, TempCurrentDay)
                End If
            End If
        End If
        WeekNumber = DatePart("ww", StartDayOfWeekDate, StartWeek, WeekOneOfYear)
        
        'Address DatePart function bug of sometimes incorrectly returning week 53
        'for last week in December when it should be week 1 of new year. If we get
        '53, but January 1st resides in the week we are calculating (any time the
        'first day of the week is greater than Dec 25th), we want to calculate based
        'off January 1st, instead of date in December.
        If WeekNumber > 52 And TempCurrentDay > 25 Then
            WeekNumber = DatePart("ww", DateSerial(YearIn + 1, 1, 1), StartWeek, WeekOneOfYear)
        End If
        Me("lblWeek" & CStr(i)).Caption = WeekNumber
        
        'Set date labels
        For j = 1 To 7 'Columns
            Set lblControl = Me("lblDate" & CStr(i) & CStr(j))
            Set bgControl = Me("bgDate" & CStr(i) & CStr(j))
            With lblControl
                
                'Previous month dates. If month starts on first day of week, entire
                'first row will be previous month
                If StartDayOfWeek = 1 And i = 1 Then
                    'If minimum date is in current month, then previous month shouldn't be visible
                    If MinDay <> 0 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek + 6 - j))
                        'Make sure previous month dates don't go beyond minimum date
                        If TempCurrentDay < PrevMonthMinDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    
                'Previous month dates if month DOESN'T start on first day of week
                ElseIf i = 1 And j < StartDayOfWeek Then
                    'If minimum date is in current month, then previous month shouldn't be visible
                    If MinDay <> 0 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(LastDayOfPrevMonth - (StartDayOfWeek - 1 - j))
                        'Make sure previous month dates don't go beyond minimum date
                        If TempCurrentDay < PrevMonthMinDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            .Enabled = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If

                'Next month dates
                ElseIf CurrentDay > LastDayOfMonth Then
                    'If maximum date is in current month, then next month shouldn't be visible
                    If MaxDay <> 32 Then
                        .Visible = False
                        bgControl.Visible = False
                    Else
                        TempCurrentDay = CLng(CurrentDay - LastDayOfMonth)
                        'Make sure next month dates don't go beyond maximum date
                        If TempCurrentDay > NextMonthMaxDay Then
                            .Visible = False
                            bgControl.Visible = False
                        Else
                            .Visible = True
                            .Enabled = True
                            bgControl.Visible = True
                            .ForeColor = lblDatePrevMonthColor
                            .Caption = CStr(TempCurrentDay)
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    CurrentDay = CurrentDay + 1
                    
                'Current month dates
                Else
                    'Disable any dates outside bounds of minimum or maximum dates.
                    'Background of date label is set to invisible, so it doesn't
                    'hover, and the date label itself is disabled so it can't be clicked
                    If CurrentDay < MinDay Or CurrentDay > MaxDay Then
                        .Visible = True
                        .Enabled = False
                        bgControl.Visible = False
                    Else 'Within bounds. Enable and set colors
                        .Visible = True
                        .Enabled = True
                        bgControl.Visible = True
                        'Set text color
                        If CurrentDay = TodayDay Then
                            .ForeColor = lblDateTodayColor
                        ElseIf j = SaturdayIndex Then
                            .ForeColor = lblDateSatColor
                        ElseIf j = SundayIndex Then
                            .ForeColor = lblDateSunColor
                        Else
                            .ForeColor = lblDateColor
                        End If
                        
                        'Set background color
                        If CurrentDay = DayIn Then
                            bgControl.BackColor = bgDateSelectedColor
                        Else
                            bgControl.BackColor = bgDateColor
                        End If
                    End If
                    .Caption = CStr(CurrentDay)
                    CurrentDay = CurrentDay + 1
                End If
            End With
        Next j
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetSelectionLabel
'
' This sub sets the caption and position of the labels that show the user's current
' selection if the Okay button is enabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetSelectionLabel(DateIn As Date)
    Dim CombinedLabelWidth As Double        'Combined width of both labels, used to center
    Dim ExtraSpace As Double                'Space between the two labels
    
    ExtraSpace = 3 * RatioToResize
    
    'If there is no selected date set yet, selected date label should be null
    If DateIn = 0 Then
        lblSelectionDate.Caption = vbNullString
        lblSelection.Left = frameCalendar.Left + ((frameCalendar.Width - lblSelection.Width) / 2)
    Else 'A selection has been made. Set caption and center
        With lblSelectionDate
            .AutoSize = False
            .Width = frameCalendar.Width
            .Caption = Format(DateIn, "mm/dd/yyyy")
            .AutoSize = True
        End With
    
        CombinedLabelWidth = lblSelection.Width + lblSelectionDate.Width
        lblSelection.Left = ((frameCalendar.Width - CombinedLabelWidth) / 2) - (ExtraSpace / 2)
        lblSelectionDate.Left = lblSelection.Left + lblSelection.Width + ExtraSpace
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSelectedDay
'
' This function checks the current month and year to see if they match the selected
' date. If so, it returns the day number of the selected date. If not, it returns 0.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSelectedDay(MonthIn As Long, YearIn As Long) As Long
    GetSelectedDay = 0
    
    'Check if a selected date was provided by the user
    If SelectedDateIn <> 0 Then
        If MonthIn = Month(SelectedDateIn) And YearIn = Year(SelectedDateIn) Then
            GetSelectedDay = Day(SelectedDateIn)
        End If
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Min / Max
'
' Get the min/max of an arbitrary number of arguments
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Min(ParamArray values() As Variant) As Variant
   Dim minValue As Variant
   Dim value As Variant
   minValue = values(0)
   For Each value In values
       If value < minValue Then minValue = value
   Next
   Min = minValue
End Function
Private Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue As Variant
   Dim value As Variant
   maxValue = values(0)
   For Each value In values
       If value > maxValue Then maxValue = value
   Next
   Max = maxValue
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The following subs all call the ClickControl sub, passing the date label that has been
' clicked. It could have saved some lines of code to create a class module which handled
' the functionality of hovering and clicking on the different controls, then simply
' declaring each date label as an object of that class. However, that would have
' necessitated the inclusion of another module in order to make the CalendarForm function
' properly. Since the main goal of this project was to have this userform be completely
' self-contained, I opted for this route.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'User clicked on the background of the date label
Private Sub bgDate11_Click(): ClickControl lblDate11: End Sub
Private Sub bgDate12_Click(): ClickControl lblDate12: End Sub
Private Sub bgDate13_Click(): ClickControl lblDate13: End Sub
Private Sub bgDate14_Click(): ClickControl lblDate14: End Sub
Private Sub bgDate15_Click(): ClickControl lblDate15: End Sub
Private Sub bgDate16_Click(): ClickControl lblDate16: End Sub
Private Sub bgDate17_Click(): ClickControl lblDate17: End Sub
Private Sub bgDate21_Click(): ClickControl lblDate21: End Sub
Private Sub bgDate22_Click(): ClickControl lblDate22: End Sub
Private Sub bgDate23_Click(): ClickControl lblDate23: End Sub
Private Sub bgDate24_Click(): ClickControl lblDate24: End Sub
Private Sub bgDate25_Click(): ClickControl lblDate25: End Sub
Private Sub bgDate26_Click(): ClickControl lblDate26: End Sub
Private Sub bgDate27_Click(): ClickControl lblDate27: End Sub
Private Sub bgDate31_Click(): ClickControl lblDate31: End Sub
Private Sub bgDate32_Click(): ClickControl lblDate32: End Sub
Private Sub bgDate33_Click(): ClickControl lblDate33: End Sub
Private Sub bgDate34_Click(): ClickControl lblDate34: End Sub
Private Sub bgDate35_Click(): ClickControl lblDate35: End Sub
Private Sub bgDate36_Click(): ClickControl lblDate36: End Sub
Private Sub bgDate37_Click(): ClickControl lblDate37: End Sub
Private Sub bgDate41_Click(): ClickControl lblDate41: End Sub
Private Sub bgDate42_Click(): ClickControl lblDate42: End Sub
Private Sub bgDate43_Click(): ClickControl lblDate43: End Sub
Private Sub bgDate44_Click(): ClickControl lblDate44: End Sub
Private Sub bgDate45_Click(): ClickControl lblDate45: End Sub
Private Sub bgDate46_Click(): ClickControl lblDate46: End Sub
Private Sub bgDate47_Click(): ClickControl lblDate47: End Sub
Private Sub bgDate51_Click(): ClickControl lblDate51: End Sub
Private Sub bgDate52_Click(): ClickControl lblDate52: End Sub
Private Sub bgDate53_Click(): ClickControl lblDate53: End Sub
Private Sub bgDate54_Click(): ClickControl lblDate54: End Sub
Private Sub bgDate55_Click(): ClickControl lblDate55: End Sub
Private Sub bgDate56_Click(): ClickControl lblDate56: End Sub
Private Sub bgDate57_Click(): ClickControl lblDate57: End Sub
Private Sub bgDate61_Click(): ClickControl lblDate61: End Sub
Private Sub bgDate62_Click(): ClickControl lblDate62: End Sub
Private Sub bgDate63_Click(): ClickControl lblDate63: End Sub
Private Sub bgDate64_Click(): ClickControl lblDate64: End Sub
Private Sub bgDate65_Click(): ClickControl lblDate65: End Sub
Private Sub bgDate66_Click(): ClickControl lblDate66: End Sub
Private Sub bgDate67_Click(): ClickControl lblDate67: End Sub
'User clicked on the actual date label itself
Private Sub lblDate11_Click(): ClickControl lblDate11: End Sub
Private Sub lblDate12_Click(): ClickControl lblDate12: End Sub
Private Sub lblDate13_Click(): ClickControl lblDate13: End Sub
Private Sub lblDate14_Click(): ClickControl lblDate14: End Sub
Private Sub lblDate15_Click(): ClickControl lblDate15: End Sub
Private Sub lblDate16_Click(): ClickControl lblDate16: End Sub
Private Sub lblDate17_Click(): ClickControl lblDate17: End Sub
Private Sub lblDate21_Click(): ClickControl lblDate21: End Sub
Private Sub lblDate22_Click(): ClickControl lblDate22: End Sub
Private Sub lblDate23_Click(): ClickControl lblDate23: End Sub
Private Sub lblDate24_Click(): ClickControl lblDate24: End Sub
Private Sub lblDate25_Click(): ClickControl lblDate25: End Sub
Private Sub lblDate26_Click(): ClickControl lblDate26: End Sub
Private Sub lblDate27_Click(): ClickControl lblDate27: End Sub
Private Sub lblDate31_Click(): ClickControl lblDate31: End Sub
Private Sub lblDate32_Click(): ClickControl lblDate32: End Sub
Private Sub lblDate33_Click(): ClickControl lblDate33: End Sub
Private Sub lblDate34_Click(): ClickControl lblDate34: End Sub
Private Sub lblDate35_Click(): ClickControl lblDate35: End Sub
Private Sub lblDate36_Click(): ClickControl lblDate36: End Sub
Private Sub lblDate37_Click(): ClickControl lblDate37: End Sub
Private Sub lblDate41_Click(): ClickControl lblDate41: End Sub
Private Sub lblDate42_Click(): ClickControl lblDate42: End Sub
Private Sub lblDate43_Click(): ClickControl lblDate43: End Sub
Private Sub lblDate44_Click(): ClickControl lblDate44: End Sub
Private Sub lblDate45_Click(): ClickControl lblDate45: End Sub
Private Sub lblDate46_Click(): ClickControl lblDate46: End Sub
Private Sub lblDate47_Click(): ClickControl lblDate47: End Sub
Private Sub lblDate51_Click(): ClickControl lblDate51: End Sub
Private Sub lblDate52_Click(): ClickControl lblDate52: End Sub
Private Sub lblDate53_Click(): ClickControl lblDate53: End Sub
Private Sub lblDate54_Click(): ClickControl lblDate54: End Sub
Private Sub lblDate55_Click(): ClickControl lblDate55: End Sub
Private Sub lblDate56_Click(): ClickControl lblDate56: End Sub
Private Sub lblDate57_Click(): ClickControl lblDate57: End Sub
Private Sub lblDate61_Click(): ClickControl lblDate61: End Sub
Private Sub lblDate62_Click(): ClickControl lblDate62: End Sub
Private Sub lblDate63_Click(): ClickControl lblDate63: End Sub
Private Sub lblDate64_Click(): ClickControl lblDate64: End Sub
Private Sub lblDate65_Click(): ClickControl lblDate65: End Sub
Private Sub lblDate66_Click(): ClickControl lblDate66: End Sub
Private Sub lblDate67_Click(): ClickControl lblDate67: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The following subs all call the HoverControl sub, passing the background of the date
' label that has been hovered over.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'User hovered over the date background
Private Sub bgDate11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate11: End Sub
Private Sub bgDate12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate12: End Sub
Private Sub bgDate13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate13: End Sub
Private Sub bgDate14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate14: End Sub
Private Sub bgDate15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate15: End Sub
Private Sub bgDate16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate16: End Sub
Private Sub bgDate17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate17: End Sub
Private Sub bgDate21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate21: End Sub
Private Sub bgDate22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate22: End Sub
Private Sub bgDate23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate23: End Sub
Private Sub bgDate24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate24: End Sub
Private Sub bgDate25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate25: End Sub
Private Sub bgDate26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate26: End Sub
Private Sub bgDate27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate27: End Sub
Private Sub bgDate31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate31: End Sub
Private Sub bgDate32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate32: End Sub
Private Sub bgDate33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate33: End Sub
Private Sub bgDate34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate34: End Sub
Private Sub bgDate35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate35: End Sub
Private Sub bgDate36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate36: End Sub
Private Sub bgDate37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate37: End Sub
Private Sub bgDate41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate41: End Sub
Private Sub bgDate42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate42: End Sub
Private Sub bgDate43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate43: End Sub
Private Sub bgDate44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate44: End Sub
Private Sub bgDate45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate45: End Sub
Private Sub bgDate46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate46: End Sub
Private Sub bgDate47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate47: End Sub
Private Sub bgDate51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate51: End Sub
Private Sub bgDate52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate52: End Sub
Private Sub bgDate53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate53: End Sub
Private Sub bgDate54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate54: End Sub
Private Sub bgDate55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate55: End Sub
Private Sub bgDate56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate56: End Sub
Private Sub bgDate57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate57: End Sub
Private Sub bgDate61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate61: End Sub
Private Sub bgDate62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate62: End Sub
Private Sub bgDate63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate63: End Sub
Private Sub bgDate64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate64: End Sub
Private Sub bgDate65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate65: End Sub
Private Sub bgDate66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate66: End Sub
Private Sub bgDate67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate67: End Sub
'User hovered over the actual date label
Private Sub lblDate11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate11: End Sub
Private Sub lblDate12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate12: End Sub
Private Sub lblDate13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate13: End Sub
Private Sub lblDate14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate14: End Sub
Private Sub lblDate15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate15: End Sub
Private Sub lblDate16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate16: End Sub
Private Sub lblDate17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate17: End Sub
Private Sub lblDate21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate21: End Sub
Private Sub lblDate22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate22: End Sub
Private Sub lblDate23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate23: End Sub
Private Sub lblDate24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate24: End Sub
Private Sub lblDate25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate25: End Sub
Private Sub lblDate26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate26: End Sub
Private Sub lblDate27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate27: End Sub
Private Sub lblDate31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate31: End Sub
Private Sub lblDate32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate32: End Sub
Private Sub lblDate33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate33: End Sub
Private Sub lblDate34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate34: End Sub
Private Sub lblDate35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate35: End Sub
Private Sub lblDate36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate36: End Sub
Private Sub lblDate37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate37: End Sub
Private Sub lblDate41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate41: End Sub
Private Sub lblDate42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate42: End Sub
Private Sub lblDate43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate43: End Sub
Private Sub lblDate44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate44: End Sub
Private Sub lblDate45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate45: End Sub
Private Sub lblDate46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate46: End Sub
Private Sub lblDate47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate47: End Sub
Private Sub lblDate51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate51: End Sub
Private Sub lblDate52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate52: End Sub
Private Sub lblDate53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate53: End Sub
Private Sub lblDate54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate54: End Sub
Private Sub lblDate55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate55: End Sub
Private Sub lblDate56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate56: End Sub
Private Sub lblDate57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate57: End Sub
Private Sub lblDate61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate61: End Sub
Private Sub lblDate62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate62: End Sub
Private Sub lblDate63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate63: End Sub
Private Sub lblDate64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate64: End Sub
Private Sub lblDate65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate65: End Sub
Private Sub lblDate66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate66: End Sub
Private Sub lblDate67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): HoverControl bgDate67: End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserForm_MouseMove / frameCalendar_MouseMove / bgDayLabels_MouseMove
'
' These three subs restore the last hovered date label to its original color when user is
' no longer hovering over any date labels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub
Private Sub frameCalendar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub
Private Sub bgDayLabels_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If HoverControlName <> vbNullString Then
        Me.Controls(HoverControlName).BackColor = HoverControlColor
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Known Bugs
'
' -If today button falls outside of years in combobox, it is possible for it to add years
'   to the combobox out of order. IE if combobox holds 2016-2026 and user clicks 'Today'
'   in 2014, combobox could then hold 2014, 2016, 2017, etc...
' -December 9999 generates an error when trying to calculate last day of month,  because
'   January 10000 is not a valid date in Excel
' -Occasionally, the month or year label is truncated. Cannot reproduce consistently
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Changelog
'
' v1.5.2
' -Bug fix: Userform not sizing properly in Word 2013
' -Bug fix: Minimum font size not being preserved correctly
' -Bug fix: Replaced WorksheetFunction.Max with custom Max function for compatibility
'   with other Office programs
'
' v1.5.1
' -Move all initialization code from GetDate and SetUserformSize to InitializeUserform
' -Fully qualify "Control" declarations as "MSForms.Control" for compatibility with Access
' -Bug fix: Eliminated FindLastDayOfMonth function, which contained a leap year bug
' -Bug fix: Calendar frame not setting background color correctly
' -Bug fix: Hover over calendar frame clears hovered control
'
' v1.5.0
' -Added a frame around all calendar elements. Calendar now positions and sizes itself
'   relative to its frame, rather than the userform as a whole. This way, the frame
'   can be placed anywhere within a larger userform to use it as an embedded calendar
'   rather than a popup. If you size the userform larger than the calendar, it will
'   remain that size, so you can add other controls.
'
' v1.4.0
' -Initial public release
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Future enhancements
'
' -Calculate all userform colors off one color argument, to reduce the wall of
'   arguments in GetDate function
' -Combine DateBorder and DateSpecialEffect arguments to one enumeration, since they
'   cancel eachother out
' -Remove userform toolbar (credit: Flemming Vadet, fv@smartoffice.dk, www.smartoffice.dk)
' -Remove extra row of trailing dates for months that have only 5 rows of dates, making
'   sure to handle special case of months with 4 rows, like Feb 2015 (credit: Greg Maxey,
'   gmaxey@mvps.org, gregmaxey.mvps.org/word_tips.htm)
' -Today button selects date and closes if Okay disabled
' -Add Cancel button
' -Better diferrentiation between disabled dates and trailing month dates (credit: Greg
'   Maxey, gmaxey@mvps.org, gregmaxey.mvps.org/word_tips.htm)
' -Move selected day calculation to SetDays function only, to avoid having to
'   redundantly calculate it in so many different functions
' -Add option to hide weekends (credit: Don Gray, don@rania.co.uk, www.rania.co.uk/ST)
' -Change cursor when hovering selectable controls
' -Month/Year in header change color on hover
' -Change buttons to flat labels w/ icons
' -Add tooltip when hovering over a date
' -Add worksheet to explain how to import/export userform
' -Add documentation explaining how to use with different date formats
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

