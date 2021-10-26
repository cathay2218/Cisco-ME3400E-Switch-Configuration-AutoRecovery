# $language = "VBScript"
# $interface = "1.0"

'Create by Cheng Yeh, Chiang; 2021/10/16(Sat)
'Project: Taipei Circular Line - Data Communication System
'Solution: Cisco ME 3400E Series Switch Automatic Configuration Recovery Script (VisualBasic Script Based)

Dim ConfigFilePath
const SCRIPT_TITLE = "Cisco ME 3400E Switch Configuration Recovery Script"

Function Main
	'Warning: This Script ONLY FOR CISCO ME 3400E Series Switch
	
	'Continue to Run?
	Resp_toRun = MsgBox("This Recovery Script is ONLY FOR Cisco ME 3400E Series" & vbCr & "Continue??", 36, SCRIPT_TITLE) 'vbYesNo = 4, Warning Query icon = 32
	If Resp_toRun = vbNo Then
		MsgBox "Script Stop by User!!", 64, SCRIPT_TITLE 'vbOKOnly = 0, Information Message icon = 64
		Exit Function
	End If
	
	'Check Switch Execute Mode
	If SwitchMode = False Then
		MsgBox "UNABLE to Confirm Switch Execute Level, Script Exit!!", 48, SCRIPT_TITLE 'vbOKOnly = 0, Warning Message icon = 48
		Exit Function
	End If
	
	'Switch Model Check
	If CheckModel = False Then
		MsgBox "This Switch is NOT Cisco ME 3400E Series, Script Exit!!", 48, SCRIPT_TITLE 'vbOKOnly = 0, Warning Message icon = 48
		Exit Function
	End If
	
	'Select Switch Configuration File
	SelectConfig
	If ConfigFilePath = "" Then
		MsgBox "Configuration File Path is Empty, Script Exit!!", 48, SCRIPT_TITLE 'vbOKOnly = 0, Warning Message icon = 48
		Exit Function
	End If
	
	'Check ConfigFilePath is Correct
	Resp_Path = MsgBox("Config File: " & ConfigFilePath & vbCr & vbCr & "Correct??", 36, SCRIPT_TITLE) 'vbYesNo = 4, Warning Query icon = 32
	If Resp_Path = vbNo Then
		MsgBox "Script Stop by User!!", 64, SCRIPT_TITLE 'vbOKOnly = 0, Information Message icon = 64
		Exit Function
	End If
	
	'Upload Switch Configuration File and Reboot
	Recovery
End Function

Function SwitchMode
	crt.Screen.Send "" & vbCr
	
	'Read Current Line
	Dim CURRENT_ROW_CONTENT
	CURRENT_ROW_CONTENT = crt.Screen.Get(crt.Screen.CurrentRow, 0, crt.Screen.CurrentRow, 500)
	
	'Check Execute Level
	If InStr(CURRENT_ROW_CONTENT, ">") <> 0 then
		'ExecLevel = "UserMode"
		'Nothing to Do
	ElseIf InStr(CURRENT_ROW_CONTENT, "(config)#") <> 0 then
		'ExecLevel = "ConfigMode"
		crt.Screen.Send "end" & vbCr
	ElseIf InStr(CURRENT_ROW_CONTENT, "(config-if)#") <> 0 then
		'ExecLevel = "ConfigMode-Interface"
		crt.Screen.Send "end" & vbCr
	ElseIf InStr(CURRENT_ROW_CONTENT, "(config-vlan)#") <> 0 then
		'ExecLevel = "ConfigMode-VLAN"
		crt.Screen.Send "end" & vbCr
	ElseIf InStr(CURRENT_ROW_CONTENT, "(config-line)#") <> 0 then
		'ExecLevel = "ConfigMode-Line"
		crt.Screen.Send "end" & vbCr
	ElseIf InStr(CURRENT_ROW_CONTENT, "#") <> 0 then
		'ExecLevel = "PrivilegeMode"
		'Nothing to Do
	Else
		'ExecLevel = "Unknown"
		SwitchMode = False
		Exit Function
	End If
	
	SwitchMode = True
End Function

Function CheckModel
	const MODEL = "ME-3400EG-12CS-M"
	const IOS_IMAGE = "ME340x-METROIPACCESSK9-M"
	const IOS_IMAGE_FILENAME = "me340x-metroipaccessk9-mz.122-60.EZ4.bin"
	
	'Check Switch Model Number
	crt.Screen.Send "sh ver | inc Model number" & vbCr
	If (crt.Screen.WaitForString (MODEL, 1) = False) Then 
		CheckModel = False
		Exit Function
	End If
	
	'Check Switch Model Name
	crt.Screen.Send "sh ver | inc Software" & vbCr
	If (crt.Screen.WaitForString (IOS_IMAGE, 1) = False) Then 
		CheckModel = False
		Exit Function
	End If
	
	'Check Switch Image File Name
	crt.Screen.Send "sh ver | inc image file" & vbCr
	If (crt.Screen.WaitForString (IOS_IMAGE_FILENAME, 1) = False) Then 
		CheckModel = False
		Exit Function
	End If
	
	CheckModel = True
End Function

Function SelectConfig
	'Select Upload Configuration File	 
	'Open File Dialog via MS HTML Application
	
	Set winShell = CreateObject("WScript.Shell")
	Set objExec = winShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
	ConfigFilePath = objExec.StdOut.ReadLine
End Function

Function Recovery	
	'Enter Privilege Mode
	crt.Screen.Send "" & vbCr
	crt.Screen.Send "enable" & vbCr
	crt.Screen.Send "enable" & vbCr
	crt.Screen.Send "enable" & vbCr
	
	'Clear Original Configuration
		'A. Clear VLAN Database
	crt.Screen.Send "delete vlan.dat" & vbCr
	crt.Screen.WaitForString("Delete filename [vlan.dat]?")
	crt.Screen.Send "" & vbCr
	crt.Screen.WaitForString("Delete flash:/vlan.dat?")
	crt.Screen.Send "" & vbCr
	
		'B. Clear Startup Configuration
	crt.Screen.Send "erase startup-config" & vbCr
	crt.Screen.WaitForString("Erasing the nvram filesystem will remove all configuration files! Continue?")
	crt.Screen.Send "" & vbCr
	crt.Screen.WaitForString("%SYS-7-NV_BLOCK_INIT: Initialized the geometry of nvram")
	crt.Screen.Send "" & vbCr
	crt.Screen.Send "" & vbCr
	
	'Upload Configuration via Xmodem Protocal(interface: RS-232)
	crt.Screen.Send "copy xmodem: startup-config" & vbCr
	crt.Screen.WaitForString("Destination filename [startup-config]")
	crt.Screen.Send "" & vbCr
	
	crt.Screen.WaitForString("Begin the Xmodem or Xmodem-1K transfer now")
	crt.Sleep 2000
	crt.FileTransfer.SendXmodem ConfigFilePath
	
	'Configuration upload complete, Restart Switch
	crt.Screen.WaitForString("bytes copied in")
	crt.Screen.Send "reload" & vbCr
	crt.Screen.WaitForString("Proceed with reload")
	crt.Screen.Send "" & vbCr
	
	'Show Interact Message
	crt.Screen.WaitForString("Reload Reason: Reload command")
	MsgBox "Switch has been Restart" & vbCr & "Configuration Recovery Complete!!", 64, SCRIPT_TITLE 'vbOKOnly = 0, Information Message icon = 64
End Function

'MsgBox Button and Icon Code
'0 = OK button only
'1 = OK and Cancel buttons
'2 = Abort, Retry, and Ignore buttons
'3 = Yes, No, and Cancel buttons
'4 = Yes and No buttons
'5 = Retry and Cancel buttons
'16 = Critical Message icon
'32 = Warning Query icon
'48 = Warning Message icon
'64 = Information Message icon
'0 = First button is default
'256 = Second button is default
'512 = Third button is default
'768 = Fourth button is default