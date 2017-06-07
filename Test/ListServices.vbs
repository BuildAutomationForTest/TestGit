strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery _
        ("Select * from Win32_Service")
For Each objService in colListOfServices
    WScript.Echo objService.SystemName& ","  
    WScript.Echo objService.Name& "," 
    WScript.Echo  objService.ServiceType& "," 
    WScript.Echo  objService.State& "," 
    WScript.Echo  objService.ExitCode& "," 
    WScript.Echo  objService.ProcessID& "," 
    WScript.Echo  objService.AcceptPause& "," 
    WScript.Echo  objService.AcceptStop& "," 
    WScript.Echo  objService.Caption& "," 
    WScript.Echo  objService.Description& "," 
    WScript.Echo  objService.DesktopInteract& "," 
    WScript.Echo  objService.DisplayName& "," 
    WScript.Echo  objService.ErrorControl& "," 
    WScript.Echo  objService.PathName& "," 
    WScript.Echo  objService.Started& "," 
    WScript.Echo  objService.StartMode& "," 
    WScript.Echo  objService.StartName& "," 
Next
