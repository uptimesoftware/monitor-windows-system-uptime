dim WshShell, objstdout
dim HOSTNAME, PASSWORD, USERNAME

' Create an instance of the wshShell object
' Gather all incoming variables

Set oShell = CreateObject("WScript.Shell")
Set oWshProcessEnv = oShell.Environment("process")

HOSTNAME = WScript.Arguments.Item(0)
USERNAME = WScript.Arguments.Item(1)
PASSWORD = WScript.Arguments.Item(2)

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(HOSTNAME, _
     "root\CIMV2", _
     USERNAME, _
     PASSWORD, _
     "MS_409")
 objSWbemServices.Security_.ImpersonationLevel = 3
Set colItems = objSWbemServices.ExecQuery("Select * From Win32_PerfFormattedData_PerfOS_System", _
     "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
	intSystemUptime = Int(objItem.SystemUpTime / 60 / 1440)
    WScript.Echo "SystemUpTime " & intSystemUptime
Next