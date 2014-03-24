dim WshShell, objstdout
dim HOSTNAME, PASSWORD, USERNAME
On Error Resume Next

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

if objSWbemServices Is Nothing Then
	WScript.Echo "Access Denied! - Check your username & password"
	WScript.Quit 1
End If

Set colItems = objSWbemServices.ExecQuery("Select * From Win32_PerfFormattedData_PerfOS_System", _
     "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	 
if colItems is Nothing Then
	WScript.Echo "Unable to Retrieve Windows SystemUpTime via WMI"
	WScript.Quit 1
End If

For Each objItem In colItems
	intSystemUptime = Int(objItem.SystemUpTime / 60 / 1440)
    WScript.Echo "SystemUpTime " & intSystemUptime
Next