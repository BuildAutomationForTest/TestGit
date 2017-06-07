Dim arrResultsArray()
i = 0

Set objCmdLib = _
CreateObject("Microsoft.CmdLib")
Set objCmdLib.ScriptingHost = WScript.Application
arrHeader = Array("Display Name", "State", "Start Mode")
arrMaxLength = Array(52, 12, 12)
strFormat = "Table"
blnPrintHeader = True
arrBlnHide = Array(False, False, False)

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" _
  & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
  ("Select * FROM Win32_Service")

For Each objService In colServices
  ReDim Preserve arrResultsArray(i)
  arrResultsArray(i) = _
  Array(objService.DisplayName, _
    objService.State,objService.StartMode)
  i = i + 1
Next
objCmdLib.ShowResults arrHeader, _
  arrResultsArray, arrMaxLength, _
  strFormat, blnPrintHeader, arrBlnHide
