strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer& "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_Service")
For Each objService in colServices
 intPadding = 300 - Len(objService.DisplayName)
 intPadding2 = 10 - Len(objService.StartMode)
 strDisplayName = objService.DisplayName & Space(intPadding)
 strStartMode = objService.StartMode & Space(intPadding2)
 Wscript.Echo strDisplayName & strStartMode & objService.State & objService.StartMode & objService.StartName
Next