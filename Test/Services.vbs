' Service.vbs
' Sample script to List services N-Z
' www.computerperformance.co.uk/
' Author Guy Thomas http://computerperformance.co.uk/
' Version 1.5 December 2010
' -------------------------------------------------------' 
Option Explicit
Dim objWMIService, objItem, objService, strServiceList
Dim colListOfServices, strComputer, strService

'On Error Resume Next

' ---------------------------------------------------------
' Pure WMI commands
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery _
("Select * from Win32_Service ")

' WMI and VBScript loop
For Each objService in colListOfServices
If UCase(Left(objService.name,1)) >"N" then
strServiceList = strServiceList & vbCr & _
objService.name

End if
Next

WScript.Echo strServiceList

' End of Example WMI script to list services