Attribute VB_Name = "modDaylightSavings"
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : modDaylightSavings
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : coverting some .js routines to VB6, converting manually, will look for some
'             native vb6 methods of doing the same and use those to test the results.
'---------------------------------------------------------------------------------------

Option Explicit


'Private BiasAdjust As Boolean
'
'' results UDT
'Private Type TZ_LOOKUP_DATA
'   TimeZoneName As String
'   bias As Long
'   IsDST As Boolean
'End Type
'
'Private tzinfo() As TZ_LOOKUP_DATA
'
''holds the correct key for the OS version
'Private sTzKey As String
'
''windows Constants And declares
'Private Const TIME_ZONE_ID_UNKNOWN As Long = 1
'Private Const TIME_ZONE_ID_STANDARD As Long = 1
'Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
'Private Const VER_PLATFORM_WIN32_NT = 2
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'
''registry Constants
'Private Const SKEY_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
'Private Const SKEY_9X = "SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones"
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const ERROR_SUCCESS = 0
''Private Const REG_SZ As Long = 1
''Private Const REG_BINARY = 3
''Private Const REG_DWORD As Long = 4
'Private Const STANDARD_RIGHTS_READ As Long = &H20000
'Private Const KEY_QUERY_VALUE As Long = &H1
'Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
'Private Const KEY_NOTIFY As Long = &H10
'Private Const SYNCHRONIZE As Long = &H100000
'Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or _
'                                   KEY_QUERY_VALUE Or _
'                                   KEY_ENUMERATE_SUB_KEYS Or _
'                                   KEY_NOTIFY) And _
'                                   (Not SYNCHRONIZE))
'
Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type REG_TIME_ZONE_INFORMATION
'   bias As Long
'   StandardBias As Long
'   DaylightBias As Long
'   StandardDate As SYSTEMTIME
'   DaylightDate As SYSTEMTIME
'End Type
'
'
Private Type TIME_ZONE_INFORMATION
    bias                    As Long
    StandardName(0 To 63)   As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(0 To 63)   As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type
'
'Private Type OSVERSIONINFO
'   OSVSize As Long
'   dwVerMajor As Long
'   dwVerMinor As Long
'   dwBuildNumber As Long
'   PlatformID As Long
'   szCSDVersion As String * 128
'End Type
'
'' APIs for determining the timezone
'
'Private Declare Function GetVersionEx Lib "kernel32" _
'() '   Alias "GetVersionExA" _
'  (lpVersionInformation As OSVERSIONINFO) As Long

Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeFormat& Lib "kernel32" Alias "GetTimeFormatA" _
(ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, _
ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long)

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'------------------------------------------------------ ENDS




'---------------------------------------------------------------------------------------
' Procedure : obtainDaylightSavings
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub obtainDaylightSavings()
    Dim DLSrules() As String
    
    On Error GoTo obtainDaylightSavings_Error
            
    'Debug.Print ("%DST func obtainDaylightSavings")
    
    ' From DLSRules.txt - assign all rules in this file to an array
    DLSrules = getDLSrules(App.path & "\Resources\txt\DLSRules.txt")

    'calculate the timezone bias
    panzerPrefs.txtBias = updateDLS(DLSrules)
    
    On Error GoTo 0
    Exit Sub

obtainDaylightSavings_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainDaylightSavings of Module modDaylightSavings"
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : updateDLS
' Author    : beededea
' Date      : 10/10/2023
' Purpose   : calculate the timezone bias
'---------------------------------------------------------------------------------------
'
Private Function updateDLS(ByRef DLSrules() As String) As Long

    Dim remoteGMTOffset1 As Long: remoteGMTOffset1 = 0
    Dim thisRule As String: thisRule = vbNullString
    Dim chosenTimeZone As String: chosenTimeZone = vbNullString
    Dim dlsRule() As String
    Dim separator As String: separator = vbNullString
    Dim localGMTOffset As Long
    
    separator = (" - ")
    
    On Error GoTo updateDLS_Error
    
    ''Debug.Print ("%DST func updateDLS")
        
    ' From timezones.txt take the offset from the selected timezone in the prefs
    chosenTimeZone = panzerPrefs.cmbMainGaugeTimeZone.List(panzerPrefs.cmbMainGaugeTimeZone.ListIndex)
    If chosenTimeZone = "System Time" Then
        tzDelta = 0
        Exit Function
    End If
    
    remoteGMTOffset1 = getRemoteOffset(chosenTimeZone) ' returns a long containing number of minutes

    ' From DSLcodesWin.txt, extract the current rule contents from the selected rule in the prefs
    thisRule = panzerPrefs.cmbMainDaylightSaving.List(panzerPrefs.cmbMainDaylightSaving.ListIndex)
    dlsRule = Split(thisRule, separator)
    
    ' read the first component of the split rule
    thisRule = dlsRule(0)
    
    tzDelta1 = theDLSdelta(DLSrules, thisRule, remoteGMTOffset1) ' return
    
    'Debug.Print ("%DST-I thisRule " & thisRule)
    'Debug.Print ("%DST-I remoteGMTOffset1 " & remoteGMTOffset1)
    'Debug.Print ("%DST-O tzDelta1 " & tzDelta1)

    localGMTOffset = fGetTimeZoneOffset ' returns a long in minutes // for UK this would be 0, for India it would be -330
    
    'Debug.Print ("%updateTime-I localGMTOffset " & localGMTOffset)    ' //-60
    'Debug.Print ("%updateTime-I remoteGMTOffset1 " & remoteGMTOffset1) ' //0
    'Debug.Print ("%updateTime-I localGMTOffset + remoteGMTOffset1 " & localGMTOffset + remoteGMTOffset1) ' // -600
    
    tzDelta = localGMTOffset + remoteGMTOffset1
    tzDelta = tzDelta + tzDelta1
    
'    Debug.Print ("%updateTime-I tzDelta " & tzDelta)
'    Debug.Print ("%updateTime-I tzDelta1 " & tzDelta1)
    
    updateDLS = tzDelta
    
    On Error GoTo 0
    Exit Function

updateDLS_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   updateDLS of Module modDaylightSavings"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fGetTimeZoneOffset
' Author    :
' Date      : 17/10/2023
' Purpose   : obtain the difference in mins between local time and system time
'---------------------------------------------------------------------------------------
'
Public Function fGetTimeZoneOffset() As Long
    Dim myTime As SYSTEMTIME
    Dim stringBuffer As String: stringBuffer = vbNullString
    Dim retVal As Long: retVal = 0
    Dim localTime As Date
    Dim standardTime As Date
    Dim paddingLocation As Integer: paddingLocation = 0
    Dim thisTimeDeviation As Long: thisTimeDeviation = 0
    
    On Error GoTo fGetTimeZoneOffset_Error

    GetLocalTime myTime
    stringBuffer = String$(255, Chr$(0))
    
    ' GetTimeFormat function formats a time as a time string for a specified locale
    retVal = GetTimeFormat&(LOCALE_SYSTEM_DEFAULT, 0, myTime, 0, stringBuffer, 254) ' fills first 8 chars in stringBuffer with system time
    paddingLocation = InStr(1, stringBuffer, Chr(0)) ' find the first bit of padding
    standardTime = Mid(stringBuffer, 1, paddingLocation - 1)
    
    GetSystemTime myTime
    stringBuffer = String$(255, Chr$(0))
    
    retVal = GetTimeFormat&(LOCALE_SYSTEM_DEFAULT, 0, myTime, 0, stringBuffer, 254) ' fills first 8 chars in stringBuffer with local time
    paddingLocation = InStr(1, stringBuffer, Chr(0))
    localTime = Mid(stringBuffer, 1, paddingLocation - 1)
    
    thisTimeDeviation = DateDiff("s", standardTime, localTime) / 60
    
    fGetTimeZoneOffset = thisTimeDeviation

    On Error GoTo 0
    Exit Function

fGetTimeZoneOffset_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetTimeZoneOffset of Module modDaylightSavings"
End Function
'---------------------------------------------------------------------------------------
' Function  : getRemoteOffset
' Author    : beededea
' Date      : 10/10/2023
' Purpose   : returns number of minutes
'---------------------------------------------------------------------------------------
'
 Private Function getRemoteOffset(ByVal entry As String) As Long

    Dim found As Boolean: found = False
    Dim thisValue As Long: thisValue = 0
    Dim foundGMT As Boolean: foundGMT = False
    Dim foundNeg As Boolean: foundNeg = False
    Dim foundString As Boolean: foundString = False
    Dim foundHrs As Boolean: foundHrs = False
    Dim foundMins As Boolean: foundMins = False
    Dim subString As String: subString = vbNullString
    Dim hoursOffset As Integer: hoursOffset = 0
    Dim minsOffset As Integer: minsOffset = 0
    
    On Error GoTo getRemoteOffset_Error
    
    ''Debug.Print ("%DST func getRemoteOffset")
    ''Debug.Print ("%DST-I entry " & entry)
    
    ' check for GMT 1-3
    subString = Left$(entry, 3)
    foundGMT = InStr(subString, "GMT")
    
    ' check for  +/- at pos. 5
    subString = Mid$(entry, 5, 1)
    If InStr(subString, "-") = 1 Then
        foundNeg = True
    Else
        foundNeg = False
    End If
    
    ' check for a string at 13 - end
    subString = Mid$(entry, 13, Len(entry))
    If subString <> vbNullString Then foundString = True
    
    ' check for a valid time at pos. 7-11
    subString = Mid$(entry, 7, 5)
    If IsNumeric(Mid$(subString, 1, 2)) Then
        hoursOffset = Val(Mid$(subString, 1, 2))
        foundHrs = True
    End If
    If IsNumeric(Mid$(subString, 4, 2)) Then
        minsOffset = Val(Mid$(subString, 4, 2))
        foundMins = True
    End If
    
    ' check all tests have passed
    If foundGMT = True And foundString = True And _
        foundHrs = True And _
        foundMins = True Then
        found = True
    Else
        found = False
        getRemoteOffset = thisValue
        ''Debug.Print ("%DST-O getRemoteOffset " & getRemoteOffset)
        Exit Function
    End If
        
    If (found = True) Then
        thisValue = minsOffset + (60 * hoursOffset)
        If foundNeg = True Then
            getRemoteOffset = thisValue - thisValue * 2
            ''Debug.Print ("%DST-O getRemoteOffset " & getRemoteOffset)
            Exit Function
        Else
            getRemoteOffset = thisValue
            ''Debug.Print ("%DST-O getRemoteOffset " & getRemoteOffset)
            Exit Function
        End If
    End If
    
    getRemoteOffset = Null 'return null;
    ''Debug.Print ("%DST-O abnormal getRemoteOffset " & getRemoteOffset)
    
    On Error GoTo 0
    Exit Function

getRemoteOffset_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getRemoteOffset of Module modDaylightSavings"
 End Function


'---------------------------------------------------------------------------------------
' Function  : getDLSrules
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : read the rule list from file into an intermediate variant via a split, then
'             into a string array
' ["US", "Apr", "Sun>=1", "120", "60", "Oct", "lastSun", "60"]
'---------------------------------------------------------------------------------------
'
Public Function getDLSrules(ByVal path As String) As String()
    
    Dim ruleList() As String
    Dim rules() As String
    Dim iFile As Integer: iFile = 0
    Dim I As Variant
    Dim lFileLen As Long: lFileLen = 0
    Dim sBuffer As String: sBuffer = vbNullString
    Dim useloop As Integer: useloop = 0
    Dim arraySize As Integer: arraySize = 0
    
    On Error GoTo getDLSrules_Error
    
    ''Debug.Print ("%DST func getDLSrules")
    ''Debug.Print ("%DST-I path " & path)

    If Dir$(path) = vbNullString Then
        Exit Function
    End If
    
    On Error GoTo ErrorHandler:
    
    iFile = FreeFile
    Open path For Binary Access Read As #iFile
    lFileLen = LOF(iFile)
    If lFileLen Then
        'Create output buffer
        sBuffer = String(lFileLen, " ")
        'Read contents of file
        Get iFile, 1, sBuffer
        'Split the file contents into an array
        ruleList = Split(sBuffer, vbCrLf)
    End If

    ' set the output rules array size to match the number of rules found
    arraySize = UBound(ruleList)
    ReDim rules(arraySize)

    ' convert the intermediate variant readinfg from ruleList to strings in output rules
    For Each I In ruleList ' for each requires a variant as I
        ' Note: to replicate the .js we should .split the rule by comma and read the contents into
        ' a 2-dimensional rules array but we run into VB6 Redim problems on 2 dimensional arrays
        ' instead we will parse the rules string when we need it - later.
        rules(useloop) = CStr(I)
        useloop = useloop + 1
    Next I
    
ErrorHandler:
    If iFile > 0 Then Close #iFile
    
    getDLSrules = rules ' return
    'Debug.Print "%DST-O getDLSrules(eg.) " & rules(1)
    
    On Error GoTo 0
    Exit Function

getDLSrules_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getDLSrules of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Function   : getNumberOfMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get the number of the month given a month name
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfMonth(ByVal thisMonth As String, ByVal utcFlag As Boolean) As Integer
    
    On Error GoTo getNumberOfMonth_Error
    
    ''Debug.Print ("%DST func getNumberOfMonth")
    ''Debug.Print ("%DST-I thisMonth " & thisMonth)
    
    getNumberOfMonth = Month(CDate(thisMonth & "/1/2000"))
    If utcFlag = True Then getNumberOfMonth = getNumberOfMonth - 1 ' convert 'normal month starting number of 1 to starting with 0 UTC

    If getNumberOfMonth < 0 Or getNumberOfMonth > 11 Then
        MsgBox ("getNumberOfMonth: " & thisMonth & " is not a valid month name")
        getNumberOfMonth = -1 ' return invalid
        
        ''Debug.Print ("%DST-O abnormal getNumberOfMonth " & getNumberOfMonth)
    End If
    
    ''Debug.Print ("%DST-O getNumberOfMonth " & getNumberOfMonth)
    
    On Error GoTo 0
    Exit Function

getNumberOfMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getNumberOfMonth of Module modDaylightSavings"

End Function

'---------------------------------------------------------------------------------------
' Function   : getNumberOfDay
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get the number of the day given a day name
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfDay(ByVal thisDay As String) As Integer
    Dim daysString As String: daysString = vbNullString
    Dim dayArray() As String
    Dim days(6) As String
    Dim I As Variant
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo getNumberOfDay_Error
    
    ''Debug.Print ("%DST func getNumberOfDay")
    ''Debug.Print ("%DST-I thisDay " & thisDay)

    daysString = "Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6"
    dayArray = Split(daysString, ",")
    
    For Each I In dayArray ' for each requires a variant as I
        days(useloop) = CStr(I)
        If InStr(days(useloop), thisDay) > 0 Then
            getNumberOfDay = Val(LTrim$(Mid$(days(useloop), 6, Len(days(useloop))))) ' return
            
            ''Debug.Print ("%DST-O getNumberOfDay " & getNumberOfDay)
            Exit Function
        End If
        useloop = useloop + 1
    Next I

    MsgBox ("getNumberOfDay: " & thisDay & " is not a valid day name")
    getNumberOfDay = 99 ' return invalid
    
    ''Debug.Print ("%DST-O Abnormal getNumberOfDay " & getNumberOfDay)

    On Error GoTo 0
    Exit Function

getNumberOfDay_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getNumberOfDay of Module modDaylightSavings"

End Function



'---------------------------------------------------------------------------------------
' Function  : getDaysInMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get the number of natural full days in a given month
'---------------------------------------------------------------------------------------
'
Public Function getDaysInMonth(ByVal thisMonth As Integer, ByVal thisYear As Integer) As Integer
    Dim monthDaysString As String: monthDaysString = vbNullString
    Dim monthDaysArray() As String
    'Dim useloop As Integer: useloop = 0
    
    On Error GoTo getmonthsIn_Error
    
    ''Debug.Print ("%DST func getDaysInMonth")
    ''Debug.Print ("%DST-I thisMonth " & thisMonth)
    ''Debug.Print ("%DST-I thisYear " & thisYear)
    
    If thisMonth < 0 And thisMonth > 11 Then
        MsgBox ("getDaysInMonth: " & thisMonth & " is not a valid month number")
        getDaysInMonth = 99 ' return invalid
        
        'Debug.Print "%DST-O Abnormal getDaysInMonth " & getDaysInMonth
        Exit Function
    End If

    monthDaysString = "31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31"
    monthDaysArray = Split(monthDaysString, ",")
    
    If thisMonth <> 1 Then ' all except Feb
        getDaysInMonth = Val(LTrim$(monthDaysArray(thisMonth))) ' return
        ''Debug.Print ("%DST-O getDaysInMonth " & getDaysInMonth)
        Exit Function
    End If
    
    If thisYear Mod 4 <> 0 Then
        getDaysInMonth = 28 ' return
        ''Debug.Print ("%DST-O getDaysInMonth " & getDaysInMonth)
        Exit Function
    End If
    
    If thisYear Mod 400 <> 0 Then
        getDaysInMonth = 29 ' return
        ''Debug.Print ("%DST-O getDaysInMonth " & getDaysInMonth)
        Exit Function
    End If
    
    If thisYear Mod 100 <> 0 Then
        getDaysInMonth = 28 ' return
        ''Debug.Print ("%DST-O getDaysInMonth " & getDaysInMonth)
        Exit Function
    End If

    getDaysInMonth = 29 ' return
    ''Debug.Print ("%DST-O getDaysInMonth " & getDaysInMonth)

    On Error GoTo 0
    Exit Function

getmonthsIn_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getmonthsIn of Module modmonthlightSavings"

End Function
    

'---------------------------------------------------------------------------------------
' Function  : getDateOfFirst
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :  get Date (1..31) Of First dayName (Sun..Sat) after date (1..31) of monthName (Jan..Dec) of year (2004..)
'              dayName:     Sun, Mon, Tue, Wed, Thu, Fr, Sat
'              monthName:   Jan, Feb, etc.
'---------------------------------------------------------------------------------------
'
Public Function getDateOfFirst(ByVal dayName As String, ByVal thisDayNumber As Integer, ByVal monthName As String, ByVal thisYear As Integer) As Integer

    Dim tDay As Integer: tDay = 0
    Dim tMonth As Integer: tMonth = 0
    Dim last As Integer: last = 0
    Dim d As Date
    Dim lastDay As Long: lastDay = 0

    On Error GoTo getDateOfFirst_Error
    
    ''Debug.Print ("%DST func getDateOfFirst")
    ''Debug.Print ("%DST-I dayName " & dayName)
    ''Debug.Print ("%DST-I thisDayNumber " & thisDayNumber)
    ''Debug.Print ("%DST-I monthName " & monthName)
    ''Debug.Print ("%DST-I thisYear " & thisYear)

    tDay = getNumberOfDay(dayName)
    tMonth = getNumberOfMonth(monthName, True)
    
    If tDay = 99 Or tMonth = 99 Then
        getDateOfFirst = 99 ' return invalid
        'Debug.Print "%DST-O Abnormal getDateOfFirst " & getDateOfFirst
        Exit Function
    End If
    
    last = thisDayNumber + 6
    
    ' convert starting with 0 UTC to normal month starting number of 1 for the VB6 CDate function to cope with
    d = CDate(last & "/" & tMonth + 1 & "/" & thisYear)
    
    lastDay = Weekday(d, vbSunday) - 1
        
    getDateOfFirst = last - (lastDay - tDay + 7) Mod 7 'return
    ''Debug.Print ("%DST-O getDateOfFirst " & getDateOfFirst)
    
    On Error GoTo 0
    Exit Function

getDateOfFirst_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getDateOfFirst of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Function  : getDateOfLast
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get Date (1..31) Of Last dayName (Sun..Sat) of monthName (Jan..Dec) of year (2004..)
'             dayName:     Sun, Mon, Tue, Wed, Thu, Fr, Sat
'             monthName:   Jan, Feb, etc.
'---------------------------------------------------------------------------------------
'
Public Function getDateOfLast(ByVal dayName As String, ByVal monthName As String, ByVal thisYear As Integer) As Integer
    Dim tDay As Integer: tDay = 0
    Dim tMonth As Integer: tMonth = 0
    Dim last As Integer: last = 0
    Dim d As Date
    Dim lastDay As Long: lastDay = 0
    
    On Error GoTo getDateOfLast_Error
    
    ''Debug.Print ("%DST func getDateOfLast")
    ''Debug.Print ("%DST-I dayName " & dayName)
    ''Debug.Print ("%DST-I monthName " & monthName)
    ''Debug.Print ("%DST-I thisYear " & thisYear)

    tDay = getNumberOfDay(dayName)
    ''Debug.Print ("%DST-I tDay " & tDay)

    tMonth = getNumberOfMonth(monthName, True)
    ''Debug.Print ("%DST-I tMonth " & tMonth)
    
    If tDay = 99 Or tMonth = 99 Then
        getDateOfLast = 99 ' return invalid
        ''Debug.Print ("%DST-O Abnormal getDateOfLast " & getDateOfLast)
        Exit Function
    End If
    
    last = getDaysInMonth(tMonth, thisYear)
    ''Debug.Print ("%DST-I last " & last)
    
    ' convert starting with 0 UTC to normal month starting number of 1 for the VB6 CDate cast to cope with
    d = CDate(last & "/" & tMonth + 1 & "/" & thisYear)
    ''Debug.Print ("%DST-I d " & d)
    
    'lastDayDate = DateSerial(thisYear, tMonth, last)
    lastDay = Weekday(d, vbSunday) - 1
    ''Debug.Print ("%DST-I lastDay " & lastDay)
    
    getDateOfLast = last - (lastDay - tDay + 7) Mod 7 'return
    ''Debug.Print ("%DST-O getDateOfLast " & getDateOfLast)
    
    On Error GoTo 0
    Exit Function

getDateOfLast_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   getDateOfLast of Module modDaylightSavings"

End Function


'---------------------------------------------------------------------------------------
' Function   : dayOfMonth
' Author    : beededea
' Date      : 09/10/2023
' Purpose   : get day of the month
'---------------------------------------------------------------------------------------
'
Public Function dayOfMonth(ByVal monthName As String, ByVal dayRule As String, ByVal thisYear As Integer) As Integer
    Dim dayName As String: dayName = vbNullString
    Dim thisDate As String: thisDate = vbNullString

    On Error GoTo dayOfMonth_Error
    
    ''Debug.Print ("%DST func dayOfMonth")
    ''Debug.Print ("%DST-I monthName " & monthName)
    ''Debug.Print ("%DST-I dayRule " & dayRule)
    ''Debug.Print ("%DST-I thisYear " & thisYear)

    If IsNumeric(dayRule) Then
        dayOfMonth = CInt(dayRule)
        ''Debug.Print ("%DST-O dayOfMonth " & dayOfMonth)
        Exit Function
    End If

    ' dayRule of form lastThu or Sun>=15
    If InStr(dayRule, "last") = 1 Then '    // dayRule of form lastThu
        dayName = Mid$(dayRule, 5)
        dayOfMonth = getDateOfLast(dayName, monthName, thisYear)
        ''Debug.Print ("%DST-O dayOfMonth " & dayOfMonth)
        
        Exit Function
    End If
    
'    // dayRule of form Sun>=15
    dayName = Left$(dayRule, 3)
    thisDate = Val(Mid$(dayRule, 6))
    dayOfMonth = getDateOfFirst(dayName, thisDate, monthName, thisYear)
    
    ''Debug.Print ("%DST-O dayOfMonth " & dayOfMonth)
        
    On Error GoTo 0
    Exit Function

dayOfMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   dayOfMonth of Module modDaylightSavings"
End Function



'---------------------------------------------------------------------------------------
' Function  : theDLSdelta
' Author    : beededea
' Date      : 09/10/2023
' Purpose   :
' parameter 1 all the rules of the type: ["US","Apr","Sun>=1","120","60","Oct","lastSun","60"]
' parameter 2 prefs selected rule eg. EU - Europe - European Union
' parameter 3 remote GMT Offset
'---------------------------------------------------------------------------------------
'
Public Function theDLSdelta(ByRef DLSrules() As String, ByVal rule As String, ByVal cityTimeOffset As Long) As Long

    On Error GoTo theDLSdelta_Error
    
'   set up variables
    Dim monthName() As String
    Dim startMonth As String: startMonth = vbNullString
    Dim startDay As String: startDay = vbNullString
    Dim startTimeDeviationInMins As Integer: startTimeDeviationInMins = 0
    Dim delta As Long: delta = 0
    Dim endMonth  As String: endMonth = vbNullString
    Dim endDay As String:  endDay = vbNullString
    Dim endTimeDeviationInMins As Integer: endTimeDeviationInMins = 0
    Dim useUTC As Boolean: useUTC = False
    Dim theDate As Date
    Dim startYear As Integer: startYear = 0
    Dim endYear As Integer: endYear = 0
    Dim currentMonth As String: currentMonth = vbNullString
    Dim newMonthNumber As Integer: newMonthNumber = 0
    Dim startDate As Integer: startDate = 0
    Dim endDate As Integer: endDate = 0
    Dim stdTime As Date
    Dim theGMTOffset As Long: theGMTOffset = 0
    Dim startHour As Integer: startHour = 0
    Dim startMin As Integer: startMin = 0
    Dim theStart As Date
    Dim endHour As Integer: endHour = 0
    Dim endMin As Integer: endMin = 0
    Dim theEnd As Date
    Dim dlsRule() As String
    
    Dim useloop As Integer: useloop = 0
    Dim arrayElementPresent As Boolean: arrayElementPresent = False
    Dim arrayNumber As Integer: arrayNumber = 0
    Dim ruleString As String: ruleString = vbNullString
    Dim buildDate As String: buildDate = vbNullString
    Dim numberOfMonth As Integer: numberOfMonth = 0
    Dim separator As String: separator = vbNullString
    Dim dateDiff1 As Double: dateDiff1 = 0
    Dim dateDiff2 As Double: dateDiff2 = 0
    
    separator = (""",""")
    monthName = ArrayString("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    Debug.Print ("%DST func theDLSdelta")
    Debug.Print ("%DST-I  DLSrules(eg.) " & DLSrules(0))
    Debug.Print ("%DST-I  rule " & rule)
    Debug.Print ("%DST-I  cityTimeOffset " & cityTimeOffset)
'
'     check whether DLS is in operation
'
    If rule = "NONE" Then
        theDLSdelta = 0 ' return abnormal
        Debug.Print ("%DST-O theDLSdelta = 0 Abnormal ")
        Exit Function
    End If
    
    arrayElementPresent = False
    
    ' find at least one matching rule in the list
    For useloop = 0 To UBound(DLSrules)

        dlsRule = Split(DLSrules(useloop), separator)
        ruleString = Mid$(dlsRule(0), 3, Len(dlsRule(0)))  '
        
        If ruleString = rule Then
            arrayElementPresent = True
            arrayNumber = useloop
            Exit For
        End If
    Next useloop
    
    Debug.Print ("%DST   DLSrules(" & arrayNumber & ") " & DLSrules(arrayNumber))

    If arrayElementPresent = False Then
        Debug.Print ("%DST-O Abnormal DLSdelta: " & rule & " is not in the list of DLS rules.")
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If

'    // extract the current rule from the rules array using the arrayNumber
'
    dlsRule = Split(DLSrules(arrayNumber), separator)
'
'    // read the various components of the split rule
'    ["AR","Oct","Sun>=15","0","60","Mar","Sun>=15","-60"]
'    ["US", "Apr", "Sun>=1", "120", "60","Oct", "lastSun", "60"]
'
    startMonth = dlsRule(1)
    startDay = dlsRule(2)
    startTimeDeviationInMins = Val(dlsRule(3))
    delta = Val(dlsRule(4))
    endMonth = dlsRule(5)
    endDay = dlsRule(6)
    endTimeDeviationInMins = Val(Left$(dlsRule(7), Len(dlsRule(7)) - 2))

'    negative times for UTC transitions (GMT starts a mid-day)
'
    useUTC = (Val(startTimeDeviationInMins) < 0) And (Val(endTimeDeviationInMins) < 0)
'
    If (useUTC) Then
        startTimeDeviationInMins = 0 - startTimeDeviationInMins
        endTimeDeviationInMins = 0 - endTimeDeviationInMins
    End If
    
    Debug.Print ("%DST   Rule:       " & rule)
    Debug.Print ("%DST   startMonth: " & startMonth)
    Debug.Print ("%DST   startDay:   " & startDay)
    Debug.Print ("%DST   startTimeDeviationInMins:  " & startTimeDeviationInMins)
    Debug.Print ("%DST   delta:      " & delta)
    Debug.Print ("%DST   endMonth:   " & endMonth)
    Debug.Print ("%DST   endDay:     " & endDay)
    Debug.Print ("%DST   endTimeDeviationInMins:    " & endTimeDeviationInMins)
    Debug.Print ("%DST   useUTC:     " & useUTC)

    Debug.Print ("*****************************")

    theDate = Now()
    Debug.Print ("%DST-I  theDate " & theDate)
    
    startYear = Year(theDate)
    Debug.Print ("%DST-I  startYear " & startYear)
        
    endYear = startYear
    Debug.Print ("%DST-I  endYear " & endYear)
    
    If getNumberOfMonth(startMonth, True) >= 6 Then          ' Southern Hemisphere
        currentMonth = Month(theDate)
        If currentMonth >= 6 Then
            endYear = endYear + 1
        Else
            startYear = startYear - 1
        End If
    End If

    If startTimeDeviationInMins < 0 Then
        startTimeDeviationInMins = 0 - startTimeDeviationInMins
    End If  ' ignore invalid sign

    startDate = dayOfMonth(startMonth, startDay, startYear)
    If startDate = 0 Then
        theDLSdelta = 0 ' return
        Debug.Print ("%DST   theDLSdelta " & theDLSdelta)
        Exit Function
    End If
    
    endDate = dayOfMonth(endMonth, endDay, endYear)
    If endDate = 0 Then
        theDLSdelta = 0 ' return
        Debug.Print ("%DST   theDLSdelta " & theDLSdelta)
        Exit Function
    End If
    
    If Val(endTimeDeviationInMins) < 0 Then ' transition on previous day in standard time
        endTimeDeviationInMins = 0 - endTimeDeviationInMins
        endDate = endDate - 1
        endTimeDeviationInMins = 1440 - endTimeDeviationInMins
        If (endDate = 0) Then
            newMonthNumber = getNumberOfMonth(endMonth, False)  ' dean
            endMonth = monthName(newMonthNumber)
            endDate = getDaysInMonth(newMonthNumber, endYear)
        End If
    End If
    
    Debug.Print ("%DST   startDate:  " & startMonth & " " & startDate & "," & startYear)
    Debug.Print ("%DST   startTimeDeviationInMins:  " & (startTimeDeviationInMins - startTimeDeviationInMins Mod 60) / 60 & ":" & startTimeDeviationInMins Mod 60)
    Debug.Print ("%DST   endDate:    " & endMonth & " " & endDate & "," & endYear)
    Debug.Print ("%DST   endTimeDeviationInMins:    " & (endTimeDeviationInMins - endTimeDeviationInMins Mod 60) / 60 & ":" & endTimeDeviationInMins Mod 60)

    theGMTOffset = 60 * cityTimeOffset
    
    Debug.Print ("%DST   cityTimeOffset:    " & cityTimeOffset)
    Debug.Print ("%DST   theGMTOffset:    " & theGMTOffset)
 
    theDate = Now()
    stdTime = GetCurrentGMTDate

    startHour = Int(startTimeDeviationInMins / 60)
    startMin = startTimeDeviationInMins Mod 60
    
    numberOfMonth = getNumberOfMonth(startMonth, False) '<<<
    
    Debug.Print ("%DST   ----")
    Debug.Print ("%DST   startYear=" & startYear)
    Debug.Print ("%DST   numberOfMonth=" & (numberOfMonth - 1))
    Debug.Print ("%DST   startDate=" & startDate)
    Debug.Print ("%DST   startHour=" & startHour)
    Debug.Print ("%DST   startMin=" & startMin)
    
    buildDate = Str$(startDate) & "/" & numberOfMonth & "/" & Str$(startYear) & " " & Str$(startHour) & ":" & Str$(startMin)
    theStart = CDate(buildDate)
    
    If useUTC = False Then
        theStart = DateAdd("s", -theGMTOffset, theStart)
    End If

    Debug.Print ("%DST   theStart= " & theStart)

    endHour = Int(endTimeDeviationInMins / 60)
    endMin = endTimeDeviationInMins Mod 60
    
    numberOfMonth = getNumberOfMonth(endMonth, False)

    Debug.Print ("%DST   ----")
    Debug.Print ("%DST   endYear=" & endYear)
    Debug.Print ("%DST   numberOfMonth=" & numberOfMonth - 1)
    Debug.Print ("%DST   endDate=" & endDate)
    Debug.Print ("%DST   endHour=" & endHour)
    Debug.Print ("%DST   endMin=" & endMin)
    
    buildDate = Str$(endDate) & "/" & numberOfMonth & "/" & Str$(endYear) & " " & Str$(endHour) & ":" & Str$(endMin)
    theEnd = CDate(buildDate)

    If useUTC = False Then
        theEnd = theEnd - theGMTOffset
    End If
    
    Debug.Print ("%DST   stdTime=" & stdTime)
    Debug.Print ("%DST   theStart=" & theStart)
    Debug.Print ("%DST   theEnd=" & theEnd)
    
'    If stdTime < theStart Then Debug.Print ("Standard time is less than the Start time")
'    If stdTime < theEnd Then Debug.Print ("Standard time is less than the Start time")
    
    dateDiff1 = DateDiff("s", stdTime, theStart)
    dateDiff2 = DateDiff("s", stdTime, theEnd)

    If (stdTime < theStart) Then
        Debug.Print ("%DST   DLS starts in " & Int(dateDiff1 / 60) & " minutes.")
    ElseIf (stdTime < theEnd) Then
        Debug.Print ("%DST   DLS ends in   " & Int(dateDiff2 / 60) & " minutes.")
    End If
'
'    If theStart <= stdTime Then Debug.Print ("the Start time is less than Standard ")
'    If stdTime < theEnd Then Debug.Print ("Standard time is less than the end time")


    If (theStart <= stdTime) And (stdTime < theEnd) Then
        theDLSdelta = delta ' return
        Debug.Print ("%DST-O theDLSdelta 1 " & theDLSdelta)
        Exit Function
    Else
        theDLSdelta = 0 ' return
        Debug.Print ("%DST-O  theDLSdelta 2 " & theDLSdelta)
        Exit Function
    End If

    On Error GoTo 0
    Exit Function

theDLSdelta_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function theDLSdelta of Module modDaylightSavings"
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetCurrentGMTDate
' Author    :
' Date      : 15/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetCurrentGMTDate() As Date

   Dim tzi As TIME_ZONE_INFORMATION
   Dim gmt As Date
   Dim dwBias As Long
   Dim tmp As String

    On Error GoTo GetCurrentGMTDate_Error

   Select Case GetTimeZoneInformation(tzi)
   Case TIME_ZONE_ID_DAYLIGHT
      dwBias = tzi.bias + tzi.DaylightBias
   Case Else
      dwBias = tzi.bias + tzi.StandardBias
   End Select

   gmt = DateAdd("n", dwBias, Now)
   tmp = Format$(gmt, "dd mmm yyyy hh:mm:ss")

   GetCurrentGMTDate = CDate(tmp)

    On Error GoTo 0
    Exit Function

GetCurrentGMTDate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetCurrentGMTDate of Module modDaylightSavings"

End Function
