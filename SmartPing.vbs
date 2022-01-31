Option Explicit
'**************************************************************************************************
' Traditional PING covers all the screen with repetitive info.
' This script will only use a single line on your screen unless something has changed.
'**************************************************************************************************
Const intFramesPerSecond = 10
Const intBarWidth = 40 
Const intDelay = 1 ' seconds to wait between pings

Class TimeInterval

    Private strTIInitialYear
    Private strTIInitialMonth
    Private strTIInitialDay
    Private strTIInitialHour
    Private strTIInitialMinute
    Private strTIInitialSecond

    Private strTIFinalYear
    Private strTIFinalMonth
    Private strTIFinalDay
    Private strTIFinalHour
    Private strTIFinalMinute
    Private strTIFinalSecond

    Private Sub Class_Initialize()
        SetInitialNow
    End Sub

    Public Function GetFixedDigits(intDigits)
        Dim objRegExp
        Set objRegExp = New RegExp
            objRegExp.Pattern = "^[0-9]$" 
        If  objRegExp.Test(intDigits) Then 
            GetFixedDigits = "0" & CStr(intDigits) 
        Else 
            GetFixedDigits = CStr(intDigits) 
        End If
    End Function

    Public Sub SetInitialNow
        strTIInitialYear   = Year( Now )
        strTIInitialMonth  = GetFixedDigits( Month(   Now ) ) 
        strTIInitialDay    = GetFixedDigits( Day(     Now ) ) 
        strTIInitialHour   = GetFixedDigits( Hour(    Now ) ) 
        strTIInitialMinute = GetFixedDigits( Minute ( Now ) ) 
        strTIInitialSecond = GetFixedDigits( Second ( Now ) ) 
        SetFinalNow
    End Sub

    Public Sub SetFinalNow
        strTIFinalYear   = Year(Now)
        strTIFinalMonth  = GetFixedDigits( Month(  Now ) ) 
        strTIFinalDay    = GetFixedDigits( Day(    Now ) ) 
        strTIFinalHour   = GetFixedDigits( Hour(   Now ) ) 
        strTIFinalMinute = GetFixedDigits( Minute( Now ) ) 
        strTIFinalSecond = GetFixedDigits( Second( Now ) ) 
    End Sub

    Public Function GetFormatedTime(intTimeDifference)
        Const SecondsPerMinute = 60
        Const SecondsPerHour   = 3600  ' 60*60 = 3600
        Const SecondsPerDay    = 86400 ' 60*60*24 = 86400

        Dim intDays
            intDays = Int( intTimeDifference / SecondsPerDay ) 
        intTimeDifference = intTimeDifference - ( SecondsPerDay * intDays )
        
        Dim intHours
            intHours = Int( intTimeDifference / SecondsPerHour ) 
        intTimeDifference = intTimeDifference - ( SecondsPerHour * intHours )
        
        Dim intMinutes
            intMinutes = Int( intTimeDifference / SecondsPerMinute ) 

        Dim intSeconds
            intSeconds = intTimeDifference - ( SecondsPerMinute * intMinutes )
                
        If intDays > 0 Then
            GetFormatedTime = GetFixedDigits( intDays  ) & "d " &_
                              GetFixedDigits( intHours ) & ":" &_
                              GetFixedDigits(intMinutes) & ":" &_
                              GetFixedDigits(intSeconds) 
        Else
            GetFormatedTime = GetFixedDigits( intHours ) & ":" &_
                              GetFixedDigits(intMinutes) & ":" &_
                              GetFixedDigits(intSeconds) 
        End If
    End Function

    Public Function GetInitial
        If SameDay Then
            GetInitial = strTIInitialHour   & ":" &_
                         strTIInitialMinute & ":" &_
                         strTIInitialSecond
        Else
            GetInitial = strTIInitialYear   & "-" &_
                         strTIInitialMonth  & "-" &_
                         strTIInitialDay    & " " &_
                         strTIInitialHour   & ":" &_
                         strTIInitialMinute & ":" &_
                         strTIInitialSecond
        End If
    End Function
    
    Public Function GetFinal
        If SameDay Then
            GetFinal = strTIFinalHour   & ":" &_
                       strTIFinalMinute & ":" &_
                       strTIFinalSecond
        Else
            GetFinal = strTIFinalYear   & "-" &_
                       strTIFinalMonth  & "-" &_
                       strTIFinalDay    & " " &_
                       strTIFinalHour   & ":" &_
                       strTIFinalMinute & ":" &_
                       strTIFinalSecond
        End If
    End Function

    Public Function SameDay
        SameDay = (strTIInitialYear  = strTIFinalYear)   AND _
                  (strTIInitialMonth = strTIFinalMonth) AND _
                  (strTIInitialDay   = strTIFinalDay) 
    End Function

    Public Function GetDuration
        Dim FirstDate, LastDate, TimeIntervalInSeconds
        FirstDate = FormatDateTime(GetInitial)
        LastDate  = FormatDateTime(GetFinal)
        TimeIntervalInSeconds = DateDiff("s",FirstDate,LastDate)
        GetDuration = GetFormatedTime(TimeIntervalInSeconds)
    End Function

End Class ' TimeInterval

Dim objDownTime
Set objDownTime = New TimeInterval
    
Dim strMarker 
Dim intAverageResponseTime ' in miliseconds
Dim intStringPosition
Dim intLChars
Dim intRChars
Dim strRemoteTarget
Dim intIndex
Dim objPing
Dim strResultChar
Dim strResultString
    strResultString = Space(intBarWidth)
Dim blnTerminate
    blnTerminate = False

'**************************************************************************************************
If WScript.Arguments.Count > 0 Then
    strRemoteTarget = WScript.Arguments(0)
Else
    strRemoteTarget = "localhost"
End If

WScript.StdOut.Write ("Pinging " & strRemoteTarget)

Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strRemoteTarget & "'")

If NOT IsIPv4(strRemoteTarget) AND (objPing.ProtocolAddress<>"") Then
    WScript.StdOut.Write " (" & objPing.ProtocolAddress & ")"
End If

WScript.StdOut.Write vbCrLf


Do
    Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strRemoteTarget & "'")

    If NOT IsNull(objPing.StatusCode) Then
        If objPing.StatusCode = 0 Then
            If strResultChar = "." Then 
                WScript.Echo
                strResultString = Space(intBarWidth)
                intStringPosition = 0
            End If
            strResultChar = "+"
            intAverageResponseTime = Round((intAverageResponseTime + objPing.ResponseTime)/2)
            objDownTime.SetInitialNow
        Else 
            strResultChar = "."
            intStringPosition = 0
            objDownTime.SetFinalNow
        End If
    Else
        WScript.Echo
        WScript.Echo "Error: Name not found."
        WScript.Quit
    End If

    intStringPosition = intStringPosition + 1
    If intStringPosition > intBarWidth Then intStringPosition = 1
    intLChars = intStringPosition - 1
    intRChars = intBarWidth - intStringPosition

    For intIndex = 1 to intFramesPerSecond
        If intIndex = intFramesPerSecond Then
            strResultString =   Left(strResultString,intLChars) &_
                                    strResultChar &_ 
                                    Right(strResultString,intRChars)
        Else
            strResultString =   Left(strResultString,intLChars) &_
                                GetMarker &_
                                Right(strResultString,intRChars)
        End If
        If strResultChar = "." Then
            PrintOnTheSameLine  "No response since [" & objDownTime.GetInitial &_
                                "] until [" & objDownTime.GetFinal &_
                                "] Downtime: [" &_
                                objDownTime.GetDuration & "]" & Space(18)
        Else
            PrintOnTheSameLine "[" & strResultString & "] " &_
                                "TTL: " & objPing.ResponseTimeToLive & " " &_
                                "  Average latency: " &_ 
                                intAverageResponseTime & "ms" & Space(18)
        End If
        
        WScript.Sleep Round(intDelay*1000/intFramesPerSecond)
    Next
Loop Until blnTerminate

'**************************************************************************************************

Sub PrintOnTheSameLine(strText)
    WScript.StdOut.Write vbCr
    WScript.StdOut.Write strText
End Sub

'**************************************************************************************************

Function GetMarker()
    Select Case strMarker
        Case "|"
            strMarker = "/"
        Case "/"
            strMarker = "-"
        Case "-"
            strMarker = "\"
        Case Else
            strMarker = "|"
    End Select
    GetMarker = strMarker
End Function

Function IsIPv4(strTarget)
    ' IPv4 has 4 octects separated by .
    ' Each octect has to be between 0 and 255.
    Dim arrOctects
        arrOctects = Split(strTarget,".")
    Dim blnValid
    If UBound(arrOctects) = 3 Then
        blnValid = IsValidOctect(arrOctects(0))
        blnValid = blnValid AND IsValidOctect(arrOctects(1))
        blnValid = blnValid AND IsValidOctect(arrOctects(2))
        blnValid = blnValid AND IsValidOctect(arrOctects(3))
    Else
        blnValid = False
    End If
    IsIPv4 = blnValid
End Function

Function IsValidOctect(intOctect)
    If IsNumeric(intOctect) Then
        IsValidOctect = (intOctect >= 0 ) AND (intOctect <= 255)
    Else
        IsValidOctect = False
    End If
End Function