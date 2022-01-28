Option Explicit
'**************************************************************************************************
' Traditional PING covers all the screen with repetitive info.
' This script will only use a single line on your screen.
'**************************************************************************************************
Const intFramesPerSecond = 10
Const intBarWidth = 40 
Const intDelay = 1 ' seconds to wait between pings

' Space for time classes

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
Dim intLastSuccessfulTime
    intLastSuccessfulTime = Timer 'The Timer function returns the number of seconds since 12:00 AM.
Dim strLastSuccessfulTime
    strLastSuccessfulTime = Time  'The Time function returns the current system time.
Dim strLastSuccessfulMoment
    strLastSuccessfulMoment = Now 'The Now function returns the current date and time according to the setting of your computer's system date and time.
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
            intLastSuccessfulTime = Timer      
            strLastSuccessfulTime = Time  
        Else 
            strResultChar = "."
            intStringPosition = 0
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
            PrintOnTheSameLine  "No response since [" & strLastSuccessfulTime &_
                                "] until [" & Time &_
                                "] Downtime: [" &_
                                GetElapsedTime(Timer - intLastSuccessfulTime) & "]" & Space(18)
        Else
            PrintOnTheSameLine "[" & strResultString & "] Average latency: " &_ 
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

Function GetElapsedTime(intDelta)
    Dim intHours
        intHours = Int( intDelta / 3600 ) 
    Dim intMins
        intMins  = Int( (intDelta - (intHours * 3600)) / 60 ) 
    Dim intSecs
        intSecs  = Int( (intDelta - (intHours * 3600)) - (intMins * 60)) 
    GetElapsedTime = GetFixedDigits(intHours) & ":" &_
                     GetFixedDigits(intMins) & ":" &_
                     GetFixedDigits(intSecs) 
End Function

Function GetFixedDigits(intDigits)
    If intDigits < 10 Then 
        GetFixedDigits = "0" & CStr(intDigits)
    Else
        GetFixedDigits = CStr(intDigits)
    End If
End Function

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
    ' Each octect has to be between 0 and 255, with some exceptions.
    '  - 0.0.0.0 is undefined.
    '  - 255.255.255.255 is a broadcast address.
    '  - x.x.x.0 is highly probable a subnet address.
    '  - x.x.x.255 is the broadcast address of that network.
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

Function IsIPv4Private(strTarget)
    ' The private ranges are:
    ' 10.x.x.x/8
    ' 172.16.x.x/12
    ' 192.168.x.x/16
    ' 169.254.x.x/16
End Function
