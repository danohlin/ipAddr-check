' Server IP address Check - check.vbs
' Written by Dan Ohlin
' May 8, 2013
' version 1.1 May 9, 2013

' Place server names in servers.txt, one per line.
' Script will pull all IP addresses for each server, 
' using currently logged on credentials. 
' See output in results.txt.


On Error Resume Next
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
const ForAppending = 8

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set WSHShell = wscript.createObject("wscript.shell")
'open the data file
Set oTextStream = oFSO.OpenTextFile("servers.txt")
'make an array from the data file
arrComputers = Split(oTextStream.ReadAll, vbNewLine)
'close the file
oTextStream.Close

Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile("results.txt", ForAppending, True)
'tf.WriteLine("result output")


For Each strComputer In arrComputers



Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")

For Each IPConfig In IPConfigSet
 If Not IsNull(IPConfig.IPAddress) Then
 For i = LBound(IPConfig.IPAddress) To UBound(IPConfig.IPAddress)
  If Not InStr(IPConfig.IPAddress(i), ":") > 0 Then  ' only get IP4 addresses

   'tf.WriteLine("==========================================")
   tf.WriteLine(strComputer & "," & IPConfig.IPAddress(i))

  End If
 Next
 End If
Next

Set objWMIService = Nothing
Set IPConfigSet = Nothing


Next


tf.Close

WScript.Echo "Complete"

