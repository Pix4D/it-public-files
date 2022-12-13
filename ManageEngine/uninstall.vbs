'Manage Engine Desktopcentral Agent .

'Script to Clean up ManageEngine Desktop Central Agent from Add remove programs .
'================================================================================

On Error Resume Next

'Removing the Agent from Add Remove Programs (if uninstallation failed)
'=====================================================================
Err.Clear
Set WshShell = WScript.CreateObject("WScript.Shell")

   WshShell.RegRead("HKEY_CLASSES_ROOT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\")
   WshShell.RegDelete "HKEY_CLASSES_ROOT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\"
   WshShell.RegDelete "HKEY_CLASSES_ROOT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Media\"
   WshShell.RegDelete "HKEY_CLASSES_ROOT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\"
   WshShell.RegDelete "HKEY_CLASSES_ROOT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\"
   
   WshShell.RegRead ("HKEY_CLASSES_ROOT\Installer\Features\F1322DA684FF95D4CA6204A5AF2ED37B\")
   WshShell.RegDelete "HKEY_CLASSES_ROOT\Installer\Features\F1322DA684FF95D4CA6204A5AF2ED37B\"
   
   
   WshShell.RegRead("HKEY_CURRENT_USER\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\")
   WshShell.RegDelete "HKEY_CURRENT_USER\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\"
   WshShell.RegDelete "HKEY_CURRENT_USER\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Media\"
   WshShell.RegDelete "HKEY_CURRENT_USER\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\"
   WshShell.RegDelete "HKEY_CURRENT_USER\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\"

   WshShell.RegRead("HKEY_USERS\.DEFAULT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\")
   WshShell.RegDelete "HKEY_USERS\.DEFAULT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Net\"
   WshShell.RegDelete "HKEY_USERS\.DEFAULT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\Media\"
   WshShell.RegDelete "HKEY_USERS\.DEFAULT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\SourceList\"
   WshShell.RegDelete "HKEY_USERS\.DEFAULT\Installer\Products\F1322DA684FF95D4CA6204A5AF2ED37B\"

   WshShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Usage\")
   WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Usage\"
   WshShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Patches\")
   WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Patches\"
   WshShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\InstallProperties\")
   WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\InstallProperties\"
   WshShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Features\")
   WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\Features\"
   WshShell.RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\")
   WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\F1322DA684FF95D4CA6204A5AF2ED37B\"
   
  
  WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6AD2231F-FF48-4D59-AC26-405AFAE23DB7}\")
  WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{6AD2231F-FF48-4D59-AC26-405AFAE23DB7}\"


   'msgbox "Manage Engine Desktop Central Agent Remove Add Remove Programs Entry Status " & Err.Number

'Uninstall Agent and Remote Control 6 Services (if already avilable )
'====================================================================
Err.Clear
WshShell.Run "%windir%\system32\sc stop "   &Chr(34)& "ManageEngine Desktop Central 6 - Agent"  &Chr(34),0,True
WshShell.Run "%windir%\system32\sc delete " &Chr(34)& "ManageEngine Desktop Central 6 - Agent"  & Chr(34),0,True
WshShell.Run "%windir%\system32\sc stop "   &Chr(34)& "ManageEngine Desktop Central 6 - Remote Control"  &Chr(34),0,True
WshShell.Run "%windir%\system32\sc delete " &Chr(34)& "ManageEngine Desktop Central 6 - Remote Control"  & Chr(34),0,True
   
'msgbox "Manage Engine Desktop Central Agent uninstallation 6 service " & Err.Number

'Uninstall Agent and Remote Control Service (if uninstallation failed)
'=====================================================================
Err.Clear

WshShell.Run "%windir%\system32\sc stop "   &Chr(34)& "ManageEngine Desktop Central - Agent"  &Chr(34),0,True
WshShell.Run "%windir%\system32\sc delete " &Chr(34)& "ManageEngine Desktop Central - Agent"  & Chr(34),0,True
WshShell.Run "%windir%\system32\sc stop "   &Chr(34)& "ManageEngine Desktop Central - Remote Control"  &Chr(34),0,True
WshShell.Run "%windir%\system32\sc delete " &Chr(34)& "ManageEngine Desktop Central - Remote Control"  & Chr(34),0,True

'msgbox "Manage Engine Desktop Central Agent uninstallation 7 service " & Err.Number

'Get the Agent Installed directory and Registry location details
'===============================================================
Err.Clear
checkOSArch = WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")

'Wscript.Echo checkOSArch 

if Err Then
	Err.Clear
	'WScript.Echo "The OS Architecture is unable to find ,so it was assumed to be 32 bit"
	regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\AdventNet\DesktopCentral\DCAgent"
	subKey = "SOFTWARE\AdventNet\DesktopCentral\DCAgent"
else
	if checkOSArch = "x86" Then
		'Wscript.Echo "The OS Architecture is 32 bit"
		regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\AdventNet\DesktopCentral\DCAgent"
		subKey = "SOFTWARE\AdventNet\DesktopCentral\DCAgent"
	else
		'Wscript.Echo "The OS Architecture is 64 bit"
		regkey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\AdventNet\DesktopCentral\DCAgent"
		subKey = "SOFTWARE\Wow6432Node\AdventNet\DesktopCentral\DCAgent"
	End IF
End If

KillProcess "dcagenttrayicon.exe"

'To kill dcagent trayicon exe 
'============================

Sub KillProcess(strProcessToKill)
	strComputer = "."

	SET objWMIService = GETOBJECT("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 

	SET colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

	count = 0
	FOR EACH objProcess in colProcess
		objProcess.Terminate()
		count = count + 1
	NEXT 
End Sub

'Delete Desktop Central Agent Directories
'========================================
Err.Clear

Set objFSO = CreateObject("Scripting.FileSystemObject")

agentInstalledDir = WshShell.RegRead(regkey&"\DCAgentInstallDir")

'msgbox "Manage Engine Desktop Central Agent installed Directory " & agentInstalledDir

if(objFSO.FolderExists(agentInstalledDir) = False) Then
	'msgbox "DesktopCentral Agent folder already deleted!"
else
	set folder = objFSO.GetFolder(agentInstalledDir)
	folder.Delete
	if(objFSO.FolderExists(agentInstalledDir) = False) Then
	'msgbox "DesktopCentral Agent folder deleted successfully"
	else
	'msgbox "Problem in deleting Agent folder: " & agentInstalledDir
	end if
end if

'msgbox "Manage Engine Desktop Central Agent folder cleanup " & Err.Number

'Removing Agent Registry location details
'==========================================

const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

DeleteSubkeys subKey 

Sub DeleteSubkeys(strKeyPath) 
    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys 

    If IsArray(arrSubkeys) Then 
        For Each strSubkey In arrSubkeys 
            DeleteSubkeys strKeyPath & "\" & strSubkey 
        Next 
    End If 

    objReg.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath 
End Sub


'msgbox "Manage Engine Desktop Central Agent registry cleanup " & Err.Number




