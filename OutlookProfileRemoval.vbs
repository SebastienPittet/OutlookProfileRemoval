' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨
'    Author: sebastien at pittet dot org
'    Goal  : Reset MAPI Profile for Office 365 migration
'  Version : Office365PRF.vbs 1.1
' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

On Error Resume Next

'@@@@@@@@@@@@@
' MAIN PROGRAM
'@@@@@@@@@@@@@

Const PRF_LOCATION = "\\FQDN_ServerName\NETLOGON\Office365\office365.prf"

If Not RegRead("HKCU\Software\PITTET\Office365\ProfileMigrated") = "Migrated" Then

	'Delete all existing MAPI Profile
	Call ShellRun("reg delete ""HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"" /f")
	Call WriteLogEvent("4", "Removed all existing MAPI profile.")

	If GetSystemVariable("ProgramFiles(x86)") = "C:\Program Files (x86)" Then 'CPU architecture x64
		Call ShellRun("""C:\Program Files (x86)\Microsoft Office\Office14\Outlook.exe""" & " /importprf " & PRF_LOCATION)
	Else 'CPU architecture 32bits
		Call ShellRun("""C:\Program Files\Microsoft Office\Office14\Outlook.exe""" & " /importprf " & PRF_LOCATION)
	End If	'CPU architecture

	'Set the new profile as default MAPI Profile
	Call RegWrite("HKCU\Software\Microsoft\Exchange\Client\Options\PickLogonProfile","0")
	
		'Log the migration
		Call RegWrite("HKCU\Software\PITTET\Office365\ProfileMigrated","Migrated")
		Call RegWrite("HKCU\Software\PITTET\Office365\MigrationDate",Now())
		Call WriteLogEvent("4", "New Outlook profile created for Office 365 migration. User=" & GetUserName & ". Computer=" & GetComputerName & ".")
Else
	Call WriteLogEvent("4", "Migration already done for this user.")
End if


'@@@@@@@@@@@@@@@
'Functions & Sub
'@@@@@@@@@@@@@@@


'Add an event in the computer's envent log
   Sub WriteLogEvent(EventType, EventText)
      Const SUCCESS = 0
      Const Error = 1
      Const WARNING = 2
      Const INFORMATION = 4
      Const AUDIT_SUCCESS = 8
      Const AUDIT_FAILURE = 16

      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.LogEvent EventType, EventText
   End Sub
'--------------------------------------------------------------------
'Write a Registry Key
   Sub RegWrite(RegKey, Value)
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.RegWrite RegKey, Value
   End Sub
   '--------------------------------------------------------------------
'Read a Registry Key
   Function RegRead(RegKey)
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      RegRead = WshShell.RegRead(RegKey)
   End Function
'--------------------------------------------------------------------
'Execute a shell command
   Sub ShellRun(CommandLine)
      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.Run CommandLine,2,True
   End Sub
'--------------------------------------------------------------------

'Return the valume of an environment variable
   Function GetSystemVariable(Variable)
      'Taper "set" dans le prompt MS-DOS pour obtenir la liste des variables
      Set WshShell = WScript.CreateObject( "WScript.Shell" )
      GetSystemVariable = WshShell.ExpandEnvironmentStrings("%" & Variable & "%")
   End Function
'--------------------------------------------------------------------
   Function GetComputerName
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetComputerName = WshNetwork.ComputerName
   End Function
'--------------------------------------------------------------------
   Function GetUserName
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetUserName = WshNetwork.UserName
   End Function
   '--------------------------------------------------------------------

