Option Explicit
'E:\MyProjects\Netlogon\WUpdate>ping localhost

'Haciendo ping a P22535.mnpsa.com.cu [::1] con 32 bytes de datos:
'Respuesta desde ::1: tiempo<1m
'Respuesta desde ::1: tiempo<1m
'Respuesta desde ::1: tiempo<1m
'Respuesta desde ::1: tiempo<1m

'Estadísticas de ping para ::1:
'    Paquetes: enviados = 4, recibidos = 4, perdidos = 0
'    (0% perdidos),
'Tiempos aproximados de ida y vuelta en milisegundos:
'    Mínimo = 0ms, Máximo = 0ms, Media = 0ms

'Parametros importantes:
' - Nombre DNS y direccion IP del host al que se le hace ping.
' - tamaño de los paquetes de datos (por defecto 32 bytes)
' - Tiempo de respuesta en ms.
' - Resultado de la solicitud.
' - Estadisticas de paquetes perdidos y tiempos de respuesta.

Dim strAddress
'strAddress = "localhost"
'strAddress = "moa-wsus"
'strAddress = "172.16.20.31"
'strAddress = "10.10.10.10"
'strAddress = "172.16.20.240"
If WScript.Arguments.Count > 0 Then
    strAddress = WScript.Arguments(0)
Else
    strAddress = "localhost"
End If

WScript.Echo "Pinging " & strAddress

If Left(GetOsVersion, 2) <> "6." Then
    WScript.Echo "Unsupported Operating System: " & GetOsVersion
    WScript.Echo "This script requires Windows Vista or Windows Server 2008 or later"
    WScript.Echo "More info: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus"
    WScript.Quit
End If

Dim objPing
Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strAddress & "'")

With objPing
    Wscript.Echo "Address: " & .Address

    If .BufferSize <> 32 Then ' Buffer size in Bytes sent with the ping command. The default value is 32.
        Wscript.Echo "Buffer size: " & .BufferSize & " Bytes"
    End If

    If .NoFragmentation Then 'If TRUE, "Do not Fragment" is marked on the packets sent. The default is FALSE.
        Wscript.Echo "No Fragmentation: " & .NoFragmentation
    End If

    If .PrimaryAddressResolutionStatus <> 0 then 'Status of the address resolution process. If successful, the value is 0 (zero). Any other value indicates an unsuccessful address resolution.
        Wscript.Echo "PrimaryAddressResolutionStatus: " & .PrimaryAddressResolutionStatus
    End If

    If .ProtocolAddress <> "" then 'Address that the destination used to reply. The default is "".
        Wscript.Echo "ProtocolAddress: " & .ProtocolAddress
    End If

    If .ProtocolAddressResolved <> "" Then 'Resolved address corresponding to the ProtocolAddress property. The default is "".
        Wscript.Echo "ProtocolAddressResolved: " & .ProtocolAddressResolved
    End If

    If .RecordRoute <> 0 then 'How many hops should be recorded while the packet is in route. The default is 0 (zero).
        Wscript.Echo "RecordRoute: " & .RecordRoute
    End If

    If .ReplyInconsistency then 'Inconsistent reply data is reported.
        Wscript.Echo "ReplyInconsistency: " & .ReplyInconsistency
    End If

    If .ReplySize <> 32 then 'Represents the size of the buffer returned.
        Wscript.Echo "ReplySize: " & .ReplySize
    End If

    If NOT .ResolveAddressNames Then 'Command resolves address names of output address values. The default is FALSE, which indicates no resolution.
        Wscript.Echo "ResolveAddressNames: " & .ResolveAddressNames
    End If

    If NOT IsNull(.ResponseTime) Then
        Wscript.Echo "ResponseTime: " & .ResponseTime & " ms" 'Time elapsed to handle the request.
    End If

    If .ResponseTimeToLive <> 0 Then
        Select Case .ResponseTimeToLive
            Case 127
                Wscript.Echo "ResponseTimeToLive: " & .ResponseTimeToLive & " (Probably a Windows host)"
            Case 63
                Wscript.Echo "ResponseTimeToLive: " & .ResponseTimeToLive & " (Probably a Linux host)"
            Case Else
                Wscript.Echo "ResponseTimeToLive: " & .ResponseTimeToLive 
        End Select
        
        'Time to live from the moment the request is received.
    End If

    If NOT IsNull (.RouteRecord) Then 'Record of intermediate hops.
        Wscript.Echo "RouteRecord: " & Join (.RouteRecord, "; ")
    End If
  
    If NOT IsNull (.RouteRecordResolved) Then 'Resolved address that corresponds to the RouteRecord value.
        Wscript.Echo "RouteRecordResolved: " & Join (.RouteRecordResolved, "; ")
    End If

    If .SourceRoute <> "" then 'Comma-separated list of valid Source Routes. The default is "".
        Wscript.Echo "SourceRoute: " & .SourceRoute
    End If

    If .SourceRouteType <> 0 Then 
        'Type of source route option to be used on the host list specified in the SourceRoute property. 
        'If a value outside of the ValueMap is specified, then 0 (zero) is assumed. 
        'The default is 0 (zero).
        Wscript.Echo "SourceRouteType: " & GetSourceRouteType(.SourceRouteType)
    End If

    If IsNull(.StatusCode) Then
        Wscript.Echo "Status code: Computer name does not exist." 'Ping command status codes.
    Else
        Wscript.Echo "Status code: " & GetStatusCode(.StatusCode) 'Ping command status codes.
    End If

    If .StatusCode <> 0 Then
        Wscript.Echo "Timeout: " & .TimeOut & " ms"
        'Time-out value in milliseconds. 
        'If a response is not received in this time, no response is assumed. 
        'The default is 1000 milliseconds, but in my test on Windows 10 it is 4000ms.
    End If

    If NOT IsNull (.TimeStampRecord) Then 'Record of time stamps for intermediate hops.
        Wscript.Echo "TimeStampRecord: " & Join (.TimeStampRecord, "; ")
    End If
  
    If NOT IsNull (.TimeStampRecordAddress) Then 'Intermediate hop that corresponds to the TimeStampRecord value.
        Wscript.Echo "TimeStampRecordAddress: " & Join (.TimeStampRecordAddress, "; ")
    End If
  
    If NOT IsNull (.TimeStampRecordAddressResolved) Then 'Resolved address that corresponds to the TimeStampRecordAddress value.
        Wscript.Echo "TimeStampRecordAddressResolved: " & Join (.TimeStampRecordAddressResolved, "; ")
    End If

    If .TimeStampRoute <> 0 then 
        Wscript.Echo "TimeStampRoute: " & .TimeStampRoute 
        'How many hops should be recorded with time stamp information while the packet is in route. 
        'A time stamp is the number of milliseconds that have passed since midnight Universal Time (UT). 
        'If the time is not available in milliseconds or cannot be provided with respect to midnight UT, 
        'then any time may be inserted as a time stamp, 
        'provided the high order bit of the Timestamp property is set to 1 (one) to indicate the use of a nonstandard value. 
        'The default is 0 (zero).
    End If

    Wscript.Echo "TimeToLive: " & .TimeToLive * 1000 & " ms"
    'Life span of the ping packet in seconds. 
    'The value is treated as an upper limit. 
    'All routers must decrement this value by 1 (one). 
    'When this value becomes 0 (zero), the packet is dropped by the router. 
    'The default value is 80 seconds. 
    'The hops between routers rarely take this amount of time.

    If .TypeOfService <> 0 Then
        Wscript.Echo "TypeOfService: " & GetTypeOfService(.TypeOfService)
        'Type of service that is used. 
        'The default value is 0 (zero), meaning "Normal".
    End If

End With



' ___________________
Function GetOsVersion
    Dim strCurrentVersion
    strCurrentVersion = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion"
    Dim strCurrentBuild
    strCurrentBuild = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild"
    GetOsVersion = GetReg( strCurrentVersion ) & "." & GetReg( strCurrentBuild )
    'https://www.gaijin.at/en/infos/windows-version-numbers
    'Operating system           CurrentVersion  CurrentBuild  Release date
    '-------------------------  --------------  ------------  ------------
    'Windows 11, Ver 21H2      	6.3	            22000	        2021-10-04
    'Windows Server 2022       	6.3	            20348	        2021-08-18
    'Windows 10, Ver 21H1      	6.3	            19043	        2021-05-18
    'Windows Server,Ver 20H2   	6.3	            19042	        2020-10-20
    'Windows 10, Ver 20H2      	6.3	            19042	        2020-10-20
    'Windows 10, Ver 2004      	6.3	            19041	        2020-05-27
    'Windows Server, Ver 1909  	6.3	            18363	        2019-11-12
    'Windows 10, Ver 1909      	6.3	            18363	        2019-11-12
    'Windows 10, Ver 1903      	6.3	            18362	        2019-05-21
    'Windows 10, Ver 1809      	6.3	            17763	        2018-11-13    
    'W-Server 2019, Ver 1809   	6.3	            17763	        2018-11-13
    'Windows 10, Ver 1803      	6.3	            17134	        2018-04-30
    'Windows 10, Ver 1709      	6.3	            16299	        2017-10-17
    'Windows 10, ver 1703      	6.3	            15063	        2017-04-05
    'W-Server 2016, Ver 1607   	6.3	            14393	        2016-10-15
    'Windows 10, ver 1607      	6.3	            14393	        2016-08-02
    'Windows 10, Ver 1511      	6.3	            10586	        2015-11-10
    'Windows 10, ver 1507      	6.3	            10240	        2015-07-29
    'Windows 8.1               	6.3	            9600	        2013-10-17  
    'Windows Server 2012 R2    	6.3	            9600	        2013-10-18
    'Windows 8                 	6.2	            9200	        2012-10-26
    'Windows Server 2012       	6.2	            9200	        2012-09-04            
    'W-Server 2008 R2 SP1      	6.1	            7601	        2011-02-22
    'Windows 7 SP1             	6.1	            7601	        2011-02-22
    'W-Server 2008 R2          	6.1	            7600	        2009-10-22
    'Windows 7                 	6.1	            7600	        2009-10-22
    'Windows Server 2008 SP2   	6.0	            6003	        2019-03-19
    'W-Server 2008 SP2         	6.0	            6002	        2009-05-26
    'Windows Vista SP2         	6.0	            6002	        2009-05-26
    'W-Server 2008             	6.0	            6001	        2008-02-27
    'Windows Vista SP1         	6.0	            6001	        2008-02-04
    'Windows Vista             	6.0	            6000	        2007-01-30
    'Windows Server 2003 SP1   	5.2	            3790.118	    2005-03-30
    'Windows Server 2003 R2    	5.2	            3790	        2005-12-06                      
    'Windows Server 2003 SP2   	5.2	            3790	        2007-03-13
    'Windows Server 2003       	5.2	            3790	        2003-04-24
    'Windows XP SP1            	5.1	            2600.1105-1106	2002-09-09            
    'Windows XP SP2            	5.1	            2600.218	    2004-08-25
    'Windows XP SP3            	5.1	            2600	        2008-04-21
    'Windows XP                	5.1	            2600	        2001-10-25
    'Windows 2000              	5.0	            2195	        2000-02-17             
    'Windows ME                	4.9	            3000	        2000-09-14           
    'Windows 98 Second Edition 	4.1	            2222	        1999-05-05
    'Windows 98                	4.1	            1998	        1998-05-15
    'Windows 95 OEM 2.5        	4.0	            950 C         	1997-11-26
    'Windows 95 OEM 2.1        	4.0	            950 B         	1997-08-27
    'Windows 95 OEM 2.0        	4.0	            950 B         	1996-08-24
    'Windows 95 OEM 1.0        	4.0	            950 A         	1996-02-14
    'Windows NT 4              	4.0	            1381	        1996-08-24
    'Windows 95                	4.0	            950	            1995-08-24

End Function

Function GetReg(strRegString)
    Dim objShell
    Set objShell = CreateObject( "WScript.Shell" )
    GetReg = objShell.RegRead( strRegString )
End Function

' ____________________________________________
Function GetSourceRouteType (intSourceRouteType)

    Dim strType

    Select Case intSourceRouteType
        case 1
            strType = "Loose Source Routing"
        case 2
            strType = "Strict Source Routing"
        case Else
            ' Default - 0 - or any other value.
            strType = intSourceRouteType & " - None"
    End Select

    GetSourceRouteType = strType

End Function

' ______________________________________
Function GetTypeOfService (intServiceType)

    Dim strType

    Select Case intServiceType
        case 2
            strType = "Minimize Monetary Cost"
        case 4
            strType = "Maximize Reliability"
        case 8
            strType = "Maximize Throughput"
        case 16
            strType = "Minimize Delay"
        Case Else
            ' Default - 0 - or any other value.
            strType = intServiceType & " - Normal"
    End Select
    GetTypeOfService = strType

End Function

' ____________________________
Function GetStatusCode (intCode)

    Dim strStatus

    Select Case intCode
        case  0
            strStatus = "Success"
        case  11001
            strStatus = "Buffer Too Small"
        case  11002
            strStatus = "Destination Net Unreachable"
        case  11003
            strStatus = "Destination Host Unreachable"
        case  11004
            strStatus = "Destination Protocol Unreachable"
        case  11005
            strStatus = "Destination Port Unreachable"
        case  11006
            strStatus = "No Resources"
        case  11007
            strStatus = "Bad Option"
        case  11008
            strStatus = "Hardware Error"
        case  11009
            strStatus = "Packet Too Big"
        case  11010
            strStatus = "Request Timed Out"
        case  11011
            strStatus = "Bad Request"
        case  11012
            strStatus = "Bad Route"
        case  11013
            strStatus = "TimeToLive Expired Transit"
        case  11014
            strStatus = "TimeToLive Expired Reassembly"
        case  11015
            strStatus = "Parameter Problem"
        case  11016
            strStatus = "Source Quench"
        case  11017
            strStatus = "Option Too Big"
        case  11018
            strStatus = "Bad Destination"
        case  11032
            strStatus = "Negotiating IPSEC"
        case  11050
            strStatus = "General Failure"
        case Else
            strStatus = intCode & " - Unknown"
    End Select

    GetStatusCode = strStatus

End Function