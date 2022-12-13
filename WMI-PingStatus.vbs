On Error Resume Next
Dim strComputer
Dim objWMIService
Dim propValue
Dim objItem
Dim SWBemlocator
Dim UserName
Dim Password
Dim colItems

strComputer = "."
UserName = ""
Password = ""
Set SWBemlocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = SWBemlocator.ConnectServer(strComputer,"root\CIMV2",UserName,Password)
Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus",,48)
For Each objItem in colItems
	WScript.Echo "Address: " & objItem.Address
	WScript.Echo "BufferSize: " & objItem.BufferSize
	WScript.Echo "NoFragmentation: " & objItem.NoFragmentation
	WScript.Echo "PrimaryAddressResolutionStatus: " & objItem.PrimaryAddressResolutionStatus
	WScript.Echo "ProtocolAddress: " & objItem.ProtocolAddress
	WScript.Echo "ProtocolAddressResolved: " & objItem.ProtocolAddressResolved
	WScript.Echo "RecordRoute: " & objItem.RecordRoute
	WScript.Echo "ReplyInconsistency: " & objItem.ReplyInconsistency
	WScript.Echo "ReplySize: " & objItem.ReplySize
	WScript.Echo "ResolveAddressNames: " & objItem.ResolveAddressNames
	WScript.Echo "ResponseTime: " & objItem.ResponseTime
	WScript.Echo "ResponseTimeToLive: " & objItem.ResponseTimeToLive
	for each propValue in objItem.RouteRecord
		WScript.Echo "RouteRecord: " & propValue
	next
	for each propValue in objItem.RouteRecordResolved
		WScript.Echo "RouteRecordResolved: " & propValue
	next
	WScript.Echo "SourceRoute: " & objItem.SourceRoute
	WScript.Echo "SourceRouteType: " & objItem.SourceRouteType
	WScript.Echo "StatusCode: " & objItem.StatusCode
	WScript.Echo "Timeout: " & objItem.Timeout
	for each propValue in objItem.TimeStampRecord
		WScript.Echo "TimeStampRecord: " & propValue
	next
	for each propValue in objItem.TimeStampRecordAddress
		WScript.Echo "TimeStampRecordAddress: " & propValue
	next
	for each propValue in objItem.TimeStampRecordAddressResolved
		WScript.Echo "TimeStampRecordAddressResolved: " & propValue
	next
	WScript.Echo "TimestampRoute: " & objItem.TimestampRoute
	WScript.Echo "TimeToLive: " & objItem.TimeToLive
	WScript.Echo "TypeofService: " & objItem.TypeofService
Next
