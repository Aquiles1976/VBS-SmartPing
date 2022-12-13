Option Explicit

'https://www.robvanderwoude.com/vbstech_network_ping.php
'Win32_PingStatus
'Requirements:
'Windows version: 	XP, Server 2003, or Vista
'Network: 	TCP/IP
'Client software: 	N/A
'Script Engine: 	any
'Summarized: 	Works in Windows XP or later.
'Doesn't work in Windows 95, 98, ME, NT 4 or 2000.
WScript.Echo "www.robvanderwoude.com on-line: " & Ping( "www.robvanderwoude.com" )

Function Ping( myHostName )
' This function returns True if the specified host could be pinged.
' myHostName can be a computer name or IP address.
' The Win32_PingStatus class used in this function requires Windows XP or later.
' This function is based on the TestPing function in a sample script by Don Jones
' http://www.scriptinganswers.com/vault/computer%20management/default.asp#activedirectoryquickworkstationinventorytxt

    ' Standard housekeeping
    Dim colPingResults, objPingResult, strQuery

    ' Define the WMI query
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & myHostName & "'"

    ' Run the WMI query
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

    ' Translate the query results to either True or False
    For Each objPingResult In colPingResults
        If Not IsObject( objPingResult ) Then
            Ping = False
        ElseIf objPingResult.StatusCode = 0 Then
            Ping = True
        Else
            Ping = False
        End If
    Next

    Set colPingResults = Nothing
End Function

'System Scripting Runtime
'Requirements:
'Windows version: 	Windows 98, ME, NT 4, 2000, XP, Server 2003 or Vista
'Network: 	TCP/IP
'Client software: 	System Scripting Runtime
'Script Engine: 	any
'Summarized: 	Works in Windows 98 and later with System Scripting Runtime is installed, with any script engine.

WScript.Echo "www.robvanderwoude.com on-line: " & PingSSR( "www.robvanderwoude.com" )

Function PingSSR( myHostName )
' This function returns True if the specified host could be pinged.
' myHostName can be a computer name or IP address.
' This function requires the System Scripting Runtime by Franz Krainer
' http://www.netal.com/ssr.htm

    ' Standard housekeeping
    Dim objIP

    Set objIP = CreateObject( "SScripting.IPNetwork" )

    If objIP.Ping( myHostName ) = 0 Then
        PingSSR = True
    Else
        PingSSR = False
    End If

    Set objIP = Nothing
End Function

'https://www.activexperts.com/network-monitor/scripts/vbscript/ping

' ///////////////////////////////////////////////////////////////////////////////
' // ActiveXperts Network Monitor  - VBScript based checks
' // For more information about ActiveXperts Network Monitor and VBScript, visit
' // http://www.activexperts.com/support/network-monitor/online/vbscript/
' ///////////////////////////////////////////////////////////////////////////////

Option Explicit

' Declaration of global variables
Dim   SYSDATA, SYSEXPLANATION   ' SYSDATA is displayed in the 'Data' column in the Manager; SYSEXPLANATION in the 'LastResponse' column

' Constants - return values
Const retvalUnknown = 1         ' ActiveXperts Network Monitor functions should always return True (-1, Success), False (0, Error) or retvalUnknown (1, Uncertain)

' // To test a function outside Network Monitor (e.g. using CSCRIPT from the
' // command line), remove the comment character (') in the following lines:
' Dim bResult
' bResult = Ping( "www.activexperts.com", 160 )
' WScript.Echo "Return value: [" & bResult & "]"
' WScript.Echo "SYSDATA: [" & SYSDATA & "]"
' WScript.Echo "SYSEXPLANATION: [" & SYSEXPLANATION & "]"


Function Ping( strHost, nMaxTimeout )
' Description: 
'     Ping a remote host. 
'     This function uses ActiveXperts Network Component.
'     ActiveXperts Network Component is automatically licensed when ActiveXperts Network Monitor is purchased
'     For more information about ActiveXperts Network Component, see: www.activexperts.com/network-component
' Parameters:
'     1) strHost As String - Hostname or IP address of the computer you want to ping
'     2) nmaxTimeOut - Timeout in milliseconds
' Usage:
'     Ping( "<Hostname | IP>", <Timeout_MSecs> )
' Sample:
'     Ping( "www.activexperts.com", 160 )

  Dim objIcmp

  Ping              = retvalUnknown  ' Default return value
  SYSDATA           = ""             ' Will hold the response time in milliseconds
  SYSEXPLANATION    = ""             ' Set initial value

  Set objIcmp = CreateObject( "AxNetwork.Icmp" )

  objIcmp.Ping strHost, 3000 ' Maximum. timeout: 3000 ms
  If( objIcmp.LastError <> 0 ) Then
    Ping            = False
    SYSDATA         = ""
    SYSEXPLANATION  = objIcmp.GetErrorDescription( objIcmp.LastError )
    Exit Function
  End If

  If( objIcmp.LastDuration > nMaxTimeout ) Then
    Ping            = False
    SYSDATA         = objIcmp.LastDuration
    SYSEXPLANATION  = "Request from [" & strHost & "] timed out, time=[" & objIcmp.LastDuration & "ms] (>" & nMaxTimeOut & "ms)"
  Else
    Ping            = True
    SYSDATA         = objIcmp.LastDuration
    SYSEXPLANATION  = "Reply from " & strHost & ", time=[" & objIcmp.LastDuration & "ms], TTL=[" & objIcmp.LastTTL & "]"
  End If

End Function




'https://morgantechspace.com/2015/07/vbscript-check-if-machine-is-online-by-ping.html

Dim hostname
hostname = "Your-PC"
Set WshShell = WScript.CreateObject("WScript.Shell")
Ping = WshShell.Run("ping -n 1 " & hostname, 0, True)
Select Case Ping
Case 0 
   WScript.Echo "The machine '" & hostname & "' is Online"
Case 1 
   WScript.Echo "The machine '" & hostname & "' is Offline"
End Select






'https://customerfx.com/article/ping-a-remote-server-using-vbscript/

Function Ping(Target)
Dim results

    On Error Resume Next

    Set shell = CreateObject("WScript.Shell")

    ' Send 1 echo request, waiting 2 seconds for result
    Set exec = shell.Exec("ping -n 1 -w 2000 " & Target)
    results = LCase(exec.StdOut.ReadAll)

    Ping = (InStr(results, "reply from") > 0)
End Function

'Now to use that code you simply pass the IP address or hostname to the Ping function and check the boolean result

If Ping("192.168.1.100") Then
    ' Do something to access the resource
End If

'You can optionally adjust the number of times the ping request is sent and the timeout interval to wait for the result.





'https://github.com/scripting-samples/scripts/blob/master/Ping.vbs
On Error Resume Next
 
strTarget = "nuchita" 'IP address or hostname
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ping -n 2 -w 1000 " & strTarget)
strPingResults = LCase(objExec.StdOut.ReadAll)
If InStr(strPingResults, "reply from") Then
  WScript.Echo strTarget & " responded to ping."
  Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strTarget & "\root\cimv2")
  Set colCompSystems = objWMIService.ExecQuery("SELECT * FROM " & _
   "Win32_ComputerSystem")
  For Each objCompSystem In colCompSystems
    WScript.Echo "Host Name: " & LCase(objCompSystem.Name)
  Next
Else
  WScript.Echo strTarget & " did not respond to ping."
End If




'https://www.itprotoday.com/devops-and-software-development/how-can-i-use-vbscript-script-ping-machine
Dim strHost

' Check that all arguments required have been passed.
If Wscript.Arguments.Count <> 0 then
    Ping = False
            'WScript.Echo "Status code is " & objRetStatus.StatusCode
        else
            Ping = True
            'Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
            'Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
            'Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
        end if
    next
End Function 