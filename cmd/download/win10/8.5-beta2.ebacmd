'EBA Command Center 8.5 | Installer
'Copyright EBA Tools 2021
'NOTICE: This script is the installer. This script will download EBA Command Center from the internet.
Option Explicit
On Error Resume Next

'Define Variables
Dim app,backup1,backup2,cmd,connectRetry,count(4),curConnectRetry,curVer,data,dataLoc,defaultShutdown,desktop,download,eba,enableEndOp,enableLegacyEndOp,endOpFail,exeValue,exeValueExt,fileDir,forVar,forVar1,forVar2,forVar3,forVar4,fs,htmlContent,https,importData,isAdmin,isInstalled,line,lines(5),loadedPlugins(9),logData,logDir,logging,logIn,logInType,missFiles,nowDate,nowTime,os,pluginCount,prog,programLoc,pWord,regLoc,saveLogin,scriptDir,scriptLoc,short,shutdownTimer,skipDo,skipExe,startMenu,startup,startupType,status,stream,sys,temp(9),title,uName,user,userType,ver,WMI,XML

'Set variables
Set app = CreateObject("Shell.Application")
Set cmd = CreateObject("Wscript.Shell")
connectRetry = 5
count(0) = 0
count(4) = 0
curConnectRetry = 1
dataLoc = cmd.ExpandEnvironmentStrings("%AppData%") & "\EBA"
defaultShutdown = "shutdown"
desktop = cmd.SpecialFolders("AllUsersDesktop")
Set download = CreateObject("Microsoft.XMLHTTP")
enableEndOp = 1
enableLegacyEndOp = False
endOpFail = False
exeValue = "eba.null"
exeValueExt = "eba.null"
Set fs = CreateObject("Scripting.FileSystemObject")
Set https = CreateObject("msxml2.xmlhttp.3.0")
isAdmin = True
isInstalled = False
line = vblf & "---------------------------------------" & vblf
logDir = dataLoc & "\EBA.log"
logging = False
missFiles = False
pluginCount = 0
prog = 0
programLoc = "C:\Program Files (x86)\EBA"
regLoc = "HKLM\SOFTWARE\EBA-Cmd"
saveLogin = False
scriptDir = fs.GetParentFolderName(scriptLoc)
scriptLoc = Wscript.ScriptFullName
shutdownTimer = 10
skipDo = False
skipExe = false
startMenu = cmd.SpecialFolders("AllUsersStartMenu") & "\Programs"
startup = cmd.SpecialFolders("Startup")
startupType = "install"
status = "EBA Cmd"
Set stream = CreateObject("Adodb.Stream")
title = "EBA Installer " & ver & " | Debug"
user = "false"
userType = "false"
ver = 8.5
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set XML = CreateObject("Microsoft.XMLDOM")

'Dependencies
Set os = WMI.ExecQuery("Select * from Win32_OperatingSystem")

Call checkWScript

'Beginning Operations
Call clearCounts
Call clearLines
Call clearTemps
Call readSettings

Call checkWScript

'Set Object Settings
XML.Async = "False"


'Check Admin
Call checkWScript
cmd.RegRead("HKEY_USERS\s-1-5-19\")
If Err.Number <> 0 Then
	isAdmin = False
Else
	isAdmin = True
End If
Err.Clear

'On Error GoTo 0

'Check OS
temp(0) = LCase(checkOS())
If InStr(temp(0),"microsoft") Then
	If InStr(temp(0),"windows") Then
		If InStr(temp(0),"11") or InStr(temp(0),"10") or InStr(temp(0),"7") or InStr(temp(0),"8") or InStr(temp(0),"vista") Then
			Call clearTemps
		Else
			Error checkOS & " does not support EBA Cmd.","INVALID_WINDOWS_VERSION"
			Call endOp("c")
		End If
	Else
		Error "Windows Recovery Environment does not support EBA Cmd.","Windows_RE"
		Call endOp("c")
	End If
Else
	Error checkOS & " does not support EBA Cmd.","INVALID_OS"
	Call endOp("c")
End If


'Get Startup Type
If fExists(programLoc & "\EBA.vbs") Then
	startupType = "update"
Else
	startupType = "install"
End If

'Check Uninstallation
If fExists(cmd.SpecialFolders("Startup") & "\uninstallEBA.vbs") And startupType <> "uninstall" Then
	Error "EBA Command Center is set to uninstall. EBA Command Center cannot start, install, update, refresh, or repair right now. Please restart your PC to finalize or cancel uninstallation.","UNINSTALLION_SCHEDULED"
	Call endOp("c")
End If

'Get Retry Count
If fExists(dataLoc & "\connect.ebacmd") Then
	Call read(dataLoc & "\connect.ebacmd","l")
	curConnectRetry = CInt(data)
End If

Call checkWScript

'Check Imports
Call checkImports

'Check if EBA-Cmd is running
If scriptRunning() Then
	Error "Cannot start EBA Cmd","EBA_ALREADY_RUNNING"
	Call endOp("s")
End If
If checkCScript() Then
	Error "EBA Command Center runs on WScript, not CScript.","USE_WSCRIPT_NOT_CSCRIPT"
	Call endOp("s")
End If

'Launch
Do
	If startupType = "install" Then
		Call modeInstall
	Elseif startupType = "update" Then
		Call modeUpdate
	Else
		eba = msgbox("Warning:" & line & "The startup type " & startupType & " was not recognized by EBA Command Center. Want to reset it?",4+48,title)
		If eba = vbYes Then
			Call write(dataLoc & "\startupType.ebacmd","normal")
		End If
		Call endOp("s")
	End If
Loop

'Modes
Sub modeInstall
	title = "EBA Installer " & ver & " | Installation"
	Call checkWScript
	If isAdmin = False Then Call endOp("fa")
	
	Call clearTemps
	
	'Installation
	eba = msgbox("EBA Command Center was not found on this device. You can chose to install a new copy, or search for an existing copy to update to EBA " & ver & ". Continue with installation?",4+64,title)
	If eba = vbNo Then
		Note("The EBA Installer will now close.")
		Call endOp("c")
	End If
	
	'Install Directory
	programLoc = inputbox("Please enter the location where EBA Command Center should be installed. If the location given contains an installation of EBA Command Center, we'll update for you. Please note that the contents of the folder will be deleted if you install a new copy.",title,programLoc)
	programLoc = Replace(programLoc,"""","")
	If Not foldExists(fs.GetParentFolderName(programLoc)) Then
		Error "The directory does not exist: " & fs.GetParentFolderName(programLoc),"DIRECTORY_NOT_FOUND"
		Call endOp("c")
	End If
	If fExists(programLoc & "\EBA.vbs") Then
		startupType = "update"
		Exit Sub
	End If
	
	'Confirm
	eba = msgbox("Confirm the installation:" & line & "Install directory: " & programLoc,4+32,title)
	If eba = vbNo Then Call endOp("c")
	
	'Registry
	cmd.Regwrite regLoc, ""
	cmd.Regwrite regLoc & "\enableOperationCompletedMenu", enableEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\enableLegacyOperationCompletedMenu", enableLegacyEndOp, "REG_DWORD"
	cmd.Regwrite "HKLM\SOFTWARE\EBA-Cmd\installDir", programLoc, "REG_SZ"
	cmd.Regwrite "HKLM\SOFTWARE\EBA-Cmd\timesToAutoRetryInternetConnection", connectRetry, "REG_DWORD"
	
	'Folders
	delete("C:\EBA")
	delete("C:\EBA-Installer")
	delete(programLoc)
	delete(dataLoc)
	newFolder(programLoc)
	newFolder(programLoc & "\Commands")
	newFolder(dataLoc)
	newFolder(dataLoc & "\Users")
	newFolder(dataLoc & "\Commands")
	newFolder(dataLoc & "\Settings")
	newFolder(dataLoc & "\Plugins")
	newFolder(dataLoc & "\PluginData")
	Call createPlugdatFolder
	
	Call update(dataLoc & "\startupType.ebacmd","firstrun","overwrite")
	
	'Create Commands
	Call updateCommands
	
	'Data Files
	Call update(dataLoc & "\isLoggedIn.ebacmd","" & vblf & "","")
	Call update(dataLoc & "\settings\logging.ebacmd","true","")
	Call update(dataLoc & "\settings\saveLogin.ebacmd","false","")
	Call update(dataLoc & "\settings\shutdownTimer.ebacmd","10","")
	Call update(dataLoc & "\settings\defaultShutdown.ebacmd","shutdown","")
	Call update(dataLoc & "\secureShutdown.ebacmd","true","")
	
	'Apply Setup
	If Not fExists(logDir) Then Call log("Created Log File")
	Call log("Installation | Installed EBA Cmd " & ver)
	
	'Create Icons
	Call createShortcut(desktop & "\EBA Command Center.lnk")
	Call createShortcut(startMenu & "\EBA Command Center.lnk")
	
	'Installed!
	eba = msgbox("EBA Command Center finished installing! Do you want to launch EBA Command Center, and perform Initial Setup now?",4+32,title)
	If eba = vbYes Then Call endOp("r")
	Call endOp("c")
End Sub
Sub modeUpdate
	title = "EBA Installer " & ver & " | Update"
	Call checkWScript
	If isAdmin = False Then Call endOp("fa")
	
	eba = msgbox("EBA Command Center is installed at " & programLoc & line & "Do you want to update EBA Command Center now?",4+32,title)
	If eba = vbNo Then Call endOp("c")
	
	'Registry
	cmd.Regwrite regLoc, ""
	cmd.Regwrite regLoc & "\enableOperationCompletedMenu", enableEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\enableLegacyOperationCompletedMenu", enableLegacyEndOp, "REG_DWORD"
	cmd.Regwrite "HKLM\SOFTWARE\EBA-Cmd\installDir", programLoc, "REG_SZ"
	cmd.Regwrite "HKLM\SOFTWARE\EBA-Cmd\timesToAutoRetryInternetConnection", connectRetry, "REG_DWORD"
	
	'Folders
	newFolder(programLoc)
	newFolder(programLoc & "\Commands")
	newFolder(dataLoc)
	newFolder(dataLoc & "\Users")
	newFolder(dataLoc & "\Commands")
	newFolder(dataLoc & "\Settings")
	newFolder(dataLoc & "\Plugins")
	newFolder(dataLoc & "\PluginData")
	delete(programLoc & "\Plugins")
	Call createPlugdatFolder
	
	'Create Commands
	Call updateCommands
	
	'Data Files
	Call update(dataLoc & "\isLoggedIn.ebacmd","" & vblf & "","")
	Call update(dataLoc & "\settings\logging.ebacmd","true","")
	Call update(dataLoc & "\settings\saveLogin.ebacmd","false","")
	Call update(dataLoc & "\settings\shutdownTimer.ebacmd","10","")
	Call update(dataLoc & "\settings\defaultShutdown.ebacmd","shutdown","")
	delete(dataLoc & "\ebaKey.ebacmd")
	
	'Apply Setup
	If Not fExists(logDir) Then Call log("Created Log File")
	Call log("Installation | Updated to EBA Cmd " & ver)
	
	'Create Icons
	Call createShortcut(desktop & "\EBA Command Center.lnk")
	Call createShortcut(startMenu & "\EBA Command Center.lnk")
	
	'Update Complete
	Note("EBA Command Center was updated to version " & ver)
	
	Call endOp("s")
End Sub




'Subroutines
Sub addError
	count(3) = count(3) + 1
End Sub
Sub addNote
	count(1) = count(1) + 1
End Sub
Sub addWarn
	count(2) = count(2) + 1
End Sub
Sub append(dir,writeData)
	If fExists(dir) Then
		Set sys = fs.OpenTextFile (dir, 8)
		sys.writeLine writeData
		sys.Close
	Elseif foldExists(fs.GetParentFolderName(dir)) Then
		Set sys = fs.CreateTextFile (dir, 8)
		sys.writeLine writeData
		sys.Close
	End If
End Sub
Sub checkImports
	If LCase(Right(importData, 10)) = ".ebaimport" Or LCase(Right(importData, 10)) = ".ebabackup" Or LCase(Right(importData, 10)) = ".ebaplugin" Then
		
		If LCase(Right(importData, 10)) = ".ebaimport" Then
			Call readLines(importData,1)
			
			If LCase(lines(1)) = "type: startup key" Then
				Call readLines(importData,2)
				
				If LCase(lines(2)) = "data: eba.recovery" Then
					eba = msgbox("Start EBA Command Center in recovery mode?",4+32,title)
					If eba = vbYes Then startupType = "recover"
					
				Else
					Error "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Startup Key" & vblf & "Data: " & lines(2),"UNKNOWN_STARTUP_KEY"
				End If
				
			Elseif lines(1) = "Type: Command" Then
				Call readLines(importData,5)
				
				If fExists(dataLoc & "\Commands\" & lines(2) & ".ebacmd") Or fExists(programLoc & "\Commands\" & lines(2) & ".ebacmd") Then
					Error "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Command" & vblf & "Error: Command with same name already exists: " & lines(2),"FILE_ALREADY_EXISTS"
				Else
					
					eba = msgbox("Do you want to import this command?" & line & "Name: " & lines(2) & vblf & "Type: " & lines(3) & vblf & "Target: " & lines(4) & vblf & "Require Login: " & lines(5),4+32,title)
					If eba = vbYes Then
						fileDir = dataLoc & "\Commands\" & lines(2) & ".ebacmd"
						Call append(fileDir,lines(4))
						Call append(fileDir,lines(3))
						Call append(fileDir,lines(5))
						Call endOp("n")
					End If
				End If
				
			Else
				Error "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Unknown","INVALID_IMPORT_FILE"
			End If
			
		Elseif LCase(Right(importData, 10)) = ".ebabackup" Then
			
			eba = msgbox("Do you want to import the contents of this backup file?", 4+32, title)
			If eba = vbYes Then
				
				'Get Type
				newFolder(dataLoc & "\tmp")
				fs.CopyFile importData, dataLoc & "\tmp\temp.zip"
				importData = dataLoc & "\tmp\temp.zip"
				Set backup1 = app.NameSpace(dataLoc & "\tmp")
				Set backup2 = app.NameSpace(importData)
				backup1.CopyHere(backup2.Items)
				temp(0) = False
				temp(1) = True
				Call checkWScript
				If fExists(dataLoc & "\tmp\host.txt") Then
					Call read(dataLoc & "\tmp\host.txt","l")
					If data = "user" or data = "cmd" or data = "settings" or data = "plug" Then
						temp(0) = data
					Else
						temp(1) = False
					End If
				Else
					temp(1) = False
				End If
				Call checkWScript
				delete(dataLoc & "\tmp")
				
				If temp(1) = False Then
					eba = LCase(inputbox("EBA Command Center could not figure out this backup file type. What is it?" & line & "'USER': Backed up user accounts." & vblf & "'CMD': Backed up commands." & vblf & "'SETTINGS': Backed up settings." & vblf & "'PLUG': Backed up plugins.",title))
					If eba = "user" or eba = "cmd" or eba = "settings" or eba = "plug" Then
						temp(0) = data
					Else
						Warn("Argument not valid.")
					End If
				End If
				If temp(0) <> False Then
					fs.CopyFile importData, dataLoc & "\tmp\temp" & ".zip"
					importData = dataLoc & "\tmp\temp" & ".zip"
					
					If temp(0) = "user" Then
						Set backup1 = App.NameSpace(dataLoc & "\Users")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Users")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						Call checkWScript
						
					Elseif eba = "cmd" Then
						Set backup1 = App.NameSpace(dataLoc & "\Commands")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Commands")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						Call checkWScript
						
					Elseif eba = "settings" Then
						Set backup1 = App.NameSpace(dataLoc & "\Settings")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Settings")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						Call checkWScript
					Elseif eba = "plug" Then
						Set backup1 = App.NameSpace(dataLoc & "\Plugins")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Plugins")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						Call checkWScript
					End If
					delete(dataLoc & "\tmp")
				End If
			End If
			
		Elseif LCase(Right(importData, 10)) = ".ebaplugin" Then
			eba = msgbox("Do you want to install this plugin? Make sure you trust the source of this plugin.", 4+32, title)
			If eba = vbYes Then
				Call checkWScript
				fs.CopyFile importData, dataLoc & "\tmp\temp.zip"
				importData = dataLoc & "\tmp\temp.zip"
				Set backup1 = App.NameSpace(dataLoc & "\Plugins")
				Set backup2 = App.NameSpace(importData)
				backup1.CopyHere(backup2.Items)
				If Err.Number = 0 Then
					Note("Plugin has been installed. Please restart EBA Command Center.")
				Else
					Error "Plugin failed to install. See WScript Error for more info.","WS/" & Err.Number
				End If
				Call checkWScript
				delete(dataLoc & "\tmp")
			End If
		End If
	Elseif importData = "" Then
		importData = False
	Else
		Error "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Unknown" & vblf & "Error: FileEXT not recognized my EBA Cmd." & lines(2),"FILEEXT_NOT_KNOWN"
	End If
End Sub
Sub checkWScript
	temp(8) = Err.Number
	temp(9) = Err.Description
	temp(7) = Err.Description
	If Not temp(8) = 0 Then
		If Err.Number = -2147024894 Then
			temp(9) = "Something went wrong accessing a file/registry key on your system."
		Elseif Err.Number = -2147024891 Then
			temp(9) = "Failed to access system registry."
		Elseif Err.Number = -2147483638 Then
			temp(9) = "Failed to download data from the EBA Website."
		Elseif Err.Number = -2146697211 Then
			temp(9) = "The installer failed to download critical EBA Command Centef files. Check your internet connection and try again."
		Elseif Err.Number = 70 Then
			temp(9) = "EBA Command Center failed to access a file because your system denied access. The file might be in use."
		Else
			temp(9) = temp(9) & " (EBA Cmd did not recognize this error)."
		End If
		Error "A WScript Error occurred during operation " & (count(0) + 1) & line & "Description: " & temp(9) & line & "Dev Description: " & temp(7),"WS/" & temp(8)
	End If
	Err.Clear
End Sub
Sub clearCounts
	For forVar = 1 to 3
		count(forVar) = 0
	Next
End Sub
Sub clearLines
	For forVar = 0 to 5
		lines(forVar) = False
	Next
End Sub
Sub clearTemps
	For forVar = 0 to 9
		temp(forVar) = False
	Next
	exeValue = "eba.null"
	exeValueExt = "eba.null"
End Sub
Sub createPlugdatFolder
	newFolder(dataLoc & "\PluginData\Commands")
	newFolder(dataLoc & "\PluginData\Scripts")
	newFolder(dataLoc & "\PluginData\Scripts\Startup")
	newFolder(dataLoc & "\PluginData\Scripts\EndOp")
	newFolder(dataLoc & "\PluginData\Scripts\Shutdown")
End Sub
Sub createShortcut(target)
	Set Short = cmd.CreateShortcut(target)
	If fExists(programLoc & "\icon.ico") Then
		With Short
			.TargetPath = programLoc & "\EBA.vbs"
			.IconLocation = programLoc & "\icon.ico"
			.Save
		End With
	Else
		With Short
			.TargetPath = programLoc & "\EBA.vbs"
			.IconLocation = "C:\Windows\System32\cmd.exe"
			.Save
		End With
	End If
End Sub
Sub dataExists(dir)
	If Not fExists(dir) Then
		missFiles = dir
	End If
End Sub
Sub endOp(arg)
	'Crash
	If arg = "c" Then
		Call log("EBA Command Center crashed.")
		wscript.quit
	End If
	
	Call checkWScript
	
	'Force Shutdown
	If arg = "f" Then
		Call log("EBA Command Center was forced to shut down")
		wscript.quit
	End If
	
	'Force Restart as Admin
	If arg = "fa" Then
		app.ShellExecute "wscript.exe", DblQuote(scriptLoc), "", "runas", 1
		wscript.quit
	End If
	
	'Force Restart at Directory
	If arg = "fd" Then
		cmd.run DblQuote(scriptLoc)
		wscript.quit
	End If
	
	'Operation Complete
	count(0) = count(0) + 1
	If enableEndOp = 1 Then
		If endOpFail = false Then
			If enableLegacyEndOp = 1 Then
				msgbox "Operation " & count(0) & " Completed with " & count(3) & " errors, " & count(2) & " warnings, and " & count(1) & " notices.",64,title
			Else
				msgbox "Operation " & count(0) & " Completed:" & line & "Errors: " & count(3) & vblf & "Warnings: " & count(2) & vblf & "Notices: " & count(1),64,title
			End If
		Else
			If enableLegacyEndOp = 1 Then
				msgbox "Operation " & count(0) & " Failed with " & count(3) & " errors, " & count(2) & " warnings, and " & count(1) & " notices.",48,title
			Else
				msgbox "Operation " & count(0) & " Failed:" & line & "Errors: " & count(3) & vblf & "Warnings: " & count(2) & vblf & "Notices: " & count(1),48,title
			End If
		End If
	End If
	Call clearCounts
	Call clearLines
	Call clearTemps
	endOpFail = False
	
	'Shutdown
	If arg = "s" Then
		Call log("EBA Command Center was shut down.")
		wscript.quit
	End If
	
	'Restart
	If arg = "r" Then
		Call log("EBA Command Center restarted.")
		cmd.run DblQuote(programLoc & "\EBA.vbs")
		wscript.quit
	End If
	
	'Restart as Admin
	If arg = "ra" Then
		Call endOp("fa")
	End If
	
	'Restart At Directory
	If arg = "rd" Then
		cmd.run DblQuote(scriptLoc)
		Wscript.quit
	End If
End Sub
Sub getTime
	nowDate = Right(0 & DatePart("m",Date),2) & "/" & Right(0 & DatePart("d",Date),2) & "/" & Right(0 & DatePart("yyyy",Date),2)
	nowTime = Right(0 & Hour(Now),2) & ":" & Right(0 & Minute(Now),2) & ":" & Right(0 & Second(Now),2)
End Sub
Sub loadPlugins(plugDir)
	If pluginCount > 9 Then
		warn "Failed to load plugin: " & plugDir & line & "The maximum number of plugins (10) has been reached."
	Else
		loadedPlugins(pluginCount) = plugDir
		pluginCount = pluginCount + 1
	End If
End Sub
Sub log(logInput)
	If logging = "true" Then
		Call getTime
		logData = "[" & nowTime & " - " & nowDate & "] " & logInput
		Call append(logDir, logData)
	End If
End Sub
Sub preparePlugins
	Call checkWScript
	For Each forVar In fs.GetFolder(dataLoc & "\Plugins").Subfolders
		If fExists(forVar & "\meta.xml") Then
			XML.load(forVar & "\meta.xml")
			Call checkWScript
			For Each forVar1 In XML.selectNodes("/Meta/Format")
				Call checkWScript
				If forVar1.text = "1" Then
					For Each forVar2 In XML.selectNodes("/Meta/License/ID")
						Call checkWScript
						For Each forVar3 In XML.selectNodes("/Meta/Version/Name")
							Call checkWScript
							For Each forVar4 In XML.selectNodes("/Meta/Version/Version")
								Call checkWScript
								temp(2) = forVar3.text
								temp(0) = goOnline("https://ethanblaisalarms.github.io/cmd/plugin/" & forVar2.text & ".txt")
								temp(0) = Left(temp(0), Len(temp(0)) - 1)
								temp(1) = goOnline("https://ethanblaisalarms.github.io/cmd/plugin/ver/" & forVar2.text & ".txt")
								temp(1) = Left(temp(1), Len(temp(1)) - 1)
								If temp(0) = temp(2) Then
									If CDbl(forVar4.text) <= CDbl(temp(1)) Then
										Call loadPlugins(forVar)
									Else
										Call addWarn
										eba = msgbox("Warning:" & line & "The plugin at " & forVar & " is an experimental version. Load anyways?",4+48,title)
										If eba = vbYes Then Call loadPlugins(forVar)
									End If
								Else
									Call addWarn
									eba = msgbox("Warning:" & line & "The plugin at " & forVar & " is not licensed. This means EBA has not validated this plugin. Loading it could be risky. Load anyways?",4+48,title)
									If eba = vbYes Then
										Call loadPlugins(forVar)
									End If
								End If
							Next
						Next
					Next
				Else
					Error "The plugin at " & forVar & " contains an invalid META.XML file, and will be skipped.","UNKNOWN_FORMAT_VERSION"
				End If
			Next
			
		Else
			warn "The plugin at " & forVar & " is missing META.XML, and will be skipped."
		End If
	Next
End Sub
Sub readCommands(baseDir)
	Call readLines(baseDir & "\Commands\" & eba & ".ebacmd",3)
	If LCase(lines(2)) = "short" Then
		eba = lines(1)
		If fExists(dataLoc & "\Commands\" & lines(1) & ".ebacmd") Then
			Call readLines(dataLoc & "\Commands\" & lines(1) & ".ebacmd",3)
		Elseif fExists(programLoc & "\Commands\" & lines(1) & ".ebacmd") Then
			Call readLines(programLoc & "\Commands\" & lines(1) & ".ebacmd",3)
		Elseif fExists(dataLoc & "\PluginData\Commands\" & lines(1) & ".ebacmd") Then
			Call readLines(dataLoc & "\PluginData\Commands\" & lines(1) & ".ebacmd",3)
		Else
			Error "That shortcut command points to a command that does not exist: " & lines(1),"INVALID_COMMAND"
		End If
	End If
	If LCase(lines(3)) = "no" Then
		temp(0) = True
	Elseif logInType = "admin" or logInType = "owner" Then
		temp(0) = True
	Else
		temp(0) = False
	End If
	If LCase(lines(2)) = "exe" Then
		If temp(0) = True Then
			If InStr(lines(1)," ") Then
				exeValue = LCase(Left(lines(1),InStr(lines(1)," ")-1))
				exeValueExt = LCase(Replace(lines(1),exeValue & " ",""))
			Else
				exeValue = LCase(lines(1))
			End If
		Else
			Error "That command requires a quick login to an administrator account. You can do so by running 'login'.","LOGIN_REQUIRED"
			eba = msgbox("Do you want to login now?",4+32,title)
			If eba = vbYes Then
				skipExe = "eba.login"
			End If
		End If
	Elseif LCase(lines(2)) = "cmd" Then
		If temp(0) = True Then
			cmd.run lines(1)
		Else
			Error "That command requires a quick login to an administrator account. You can do so by running 'login'.","LOGIN_REQUIRED"
			eba = msgbox("Do you want to login now?",4+32,title)
			If eba = vbYes Then
				skipExe = "eba.login"
			End If
		End If
	Elseif LCase(lines(2)) = "file" Then
		If temp(0) = True Then
			cmd.run DblQuote(lines(1))
		Else
			Error "That command requires a quick login to an administrator account. You can do so by running 'login'.","LOGIN_REQUIRED"
			eba = msgbox("Do you want to login now?",4+32,title)
			If eba = vbYes Then
				skipExe = "eba.login"
			End If
		End If
	Elseif LCase(lines(2)) = "url" Then
		Set short = cmd.CreateShortcut(dataLoc & "\temp.url")
		With short
			.TargetPath = lines(1)
			.Save
		End With
		cmd.run DblQuote(dataLoc & "\temp.url")
	Elseif LCase(lines(2)) = "script" Then
		If fExists(dataLoc & "\PluginData\Scripts\" & lines(1)) Then
			cmd.run dataLoc & "\PluginData\Scripts\" & lines(1)
		Else
			Error "The command references a script that does not exist.","FILE_NOT_FOUND"
		End If
	Else
		Error "That command contains invalid data or is corrupt.","INVALID_COMMAND"
	End If
End Sub
Sub readLines(dir,lineInt)
	If fExists(dir) Then
		Set sys = fs.OpenTextFile (dir, 1)
		For forVar = 1 to lineInt
			lines(forVar) = sys.readLine
		Next
		sys.Close
	Else
		Error "Given file not found: " & dir,"BAD_FILE_DIRECTORY"
	End If
End Sub
Sub readSettings
	Call checkWScript
	
	programLoc = "C:\Program Files\EBA"
	
	'Registry Call read
	programLoc = cmd.Regread(regLoc & "\installDir")
	enableEndOp = cmd.Regread(regLoc & "\enableOperationCompletedMenu")
	connectRetry = cmd.Regread(regLoc & "\timesToAutoRetryInternetConnection")
	enableLegacyEndOp = cmd.Regread(regLoc & "\enableLegacyOperationCompletedMenu")
	
	'Conversion
	enableEndOp = CInt(enableEndOp)
	connectRetry = CInt(connectRetry)
	enableLegacyEndOp = CInt(enableLegacyEndOp)
	Err.Clear
	
	'Read Files
	If fExists(dataLoc & "\settings\logging.ebacmd") Then
		Call read(dataLoc & "\settings\logging.ebacmd","l")
		logging = data
	Else
		logging = "true"
	End If
	
	If fExists(dataLoc & "\settings\saveLogin.ebacmd") Then
		Call read(dataLoc & "\settings\saveLogin.ebacmd","l")
		saveLogin = data
	Else
		saveLogin = "false"
	End If
	
	If fExists(dataLoc & "\settings\shutdownTimer.ebacmd") Then
		Call read(dataLoc & "\settings\shutdownTimer.ebacmd","l")
		shutdownTimer = CDbl(data)
	Else
		shutdownTimer = 10
	End If
	
	If fExists(dataLoc & "\settings\defaultShutdown.ebacmd") Then
		Call read(dataLoc & "\settings\defaultShutdown.ebacmd","l")
		defaultShutdown = data
	Else
		defaultShutdown = "shutdown"
	End If
	
	Err.Clear
End Sub
Sub read(dir,arg)
	If fExists(dir) Then
		Dim tempVal
		Set sys = fs.OpenTextFile (dir,1)
		tempVal = sys.readAll
		tempVal = Left(tempVal, Len(tempVal)	- 2)
		sys.Close
		If arg = "l" Then tempVal = LCase(tempVal)
		If arg = "u" Then tempVal = UCase(tempVal)
		data = tempVal
	Else
		Error "Given file not found: " & dir,"BAD_FILE_DIRECTORY"
	End If
End Sub
Sub runPlugins
	Call createPlugdatFolder
	Call clearTemps
	For forVar = 0 to 9
		temp(0) = loadedPlugins(forVar)
		If foldExists(temp(0) & "\Commands") Then
			For Each forVar1 In fs.GetFolder(temp(0) & "\Commands").Files
				XML.load(temp(0) & "\Commands\" & forVar1.name)
				For Each forVar2 In XML.selectNodes("/Command/Format")
					If forVar2.text = "1" Then
						For Each forVar3 In XML.selectNodes("/Command/Target")
							temp(1) = forVar3.Text
						Next
						For Each forVar3 In XML.selectNodes("/Command/Type")
							temp(2) = forVar3.text
						Next
						For Each forVar3 In XML.selectNodes("/Command/Login")
							temp(3) = forVar3.text
						Next
						Call write(dataLoc & "\PluginData\Commands\" & Replace(forVar1.name,".xml","") & ".ebacmd",temp(1) & vblf & temp(2) & vblf & temp(3) & vblf & "no")
					Else
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Commands\" & forVar1.name & vblf & "Error Generated: <Command>/<Format>***ERR_INVAL***</Format>\</Command>" & vblf & "What this means: The value at /Command/Format is invalid." & line & "This XML will be skipped.")
					End If
				Next
			Next
		End If
		If foldExists(temp(0) & "\Scripts.vbs") Then
			If foldExists(temp(0) & "\Scripts.vbs\Startup") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.vbs\Startup").Files
					If LCase(Right(forVar1, 4)) = ".vbs" Then
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Startup\" & forVar1.Name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.vbs\Startup\" & forVar1.name & vblf & "Error Generated: ScriptVBSEncounteredNonVBS" & vblf & "What this means: The script could not be loaded at startup by EBA Command Center because Script.vbs only supports VBS files." & line & "This script will not execute on startup.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.vbs\OperationComplete") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.vbs\OperationComplete").Files
					If LCase(Right(forVar1, 4)) = ".vbs" Then
						newFolder(dataLoc & "\PluginData\Scripts\EndOp")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\EndOp\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.vbs\OperationCompleted\" & forVar1.name & vblf & "Error Generated: ScriptVBSEncounteredNonVBS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.vbs only supports VBS files." & line & "This script will not execute after EndOp.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.vbs\Shutdown") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.vbs\Shutdown").Files
					If LCase(Right(forVar1, 4)) = ".vbs" Then
						newFolder(dataLoc & "\PluginData\Scripts\Shutdown")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Shutdown\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.vbs\Shutdown\" & forVar1.name & vblf & "Error Generated: ScriptVBSEncounteredNonVBS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.vbs only supports VBS files." & line & "This script will not execute on shutdown.")
					End If
				Next
			End If
			For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.vbs").Files
				If LCase(Right(forVar1, 4)) = ".vbs" Then
					newFolder(dataLoc & "\PluginData\Scripts")
					fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\" & forVar1.name
				Elseif LCase(forVar1.name) <> "desktop.ini" Then
					Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.vbs\" & forVar1.name & vblf & "Error Generated: ScriptVBSEncounteredNonVBS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.vbs only supports VBS files." & line & "This script will not execute when referenced.")
				End If
			Next
		End If
		If foldExists(temp(0) & "\Scripts.js") Then
			newFolder(dataLoc & "\PluginData\Scripts")
			If foldExists(temp(0) & "\Scripts.js\Startup") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.js\Startup").Files
					If LCase(Right(forVar1, 3)) = ".js" Then
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Startup\" & forVar1.Name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.js\Startup\" & forVar1.name & vblf & "Error Generated: ScriptJSEncounteredNonJS" & vblf & "What this means: The script could not be loaded at startup by EBA Command Center because Script.js only supports JS files." & line & "This script will not execute on startup.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.js\OperationComplete") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.js\OperationComplete").Files
					If LCase(Right(forVar1, 3)) = ".js" Then
						newFolder(dataLoc & "\PluginData\Scripts\EndOp")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\EndOp\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.js\OperationCompleted\" & forVar1.name & vblf & "Error Generated: ScriptJSEncounteredNonJS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.js only supports JS files." & line & "This script will not execute after EndOp.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.js\Shutdown") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.js\Shutdown").Files
					If LCase(Right(forVar1, 3)) = ".js" Then
						newFolder(dataLoc & "\PluginData\Scripts\Shutdown")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Shutdown\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.js\Shutdown\" & forVar1.name & vblf & "Error Generated: ScriptJSEncounteredNonJS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.js only supports JS files." & line & "This script will not execute on shutdown.")
					End If
				Next
			End If
			For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.js").Files
				If LCase(Right(forVar1, 3)) = ".js" Then
					newFolder(dataLoc & "\PluginData\Scripts")
					fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\" & forVar1.name
				Elseif LCase(forVar1.name) <> "desktop.ini" Then
					Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.js\" & forVar1.name & vblf & "Error Generated: ScriptJSEncounteredNonJS" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.js only supports JS files." & line & "This script will not execute when referenced.")
				End If
			Next
		End If
		If foldExists(temp(0) & "\Scripts.exe") Then
			newFolder(dataLoc & "\PluginData\Scripts")
			If foldExists(temp(0) & "\Scripts.exe\Startup") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.exe\Startup").Files
					If LCase(Right(forVar1, 4)) = ".exe" Then
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Startup\" & forVar1.Name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.exe\Startup\" & forVar1.name & vblf & "Error Generated: ScriptEXEEncounteredNonEXE" & vblf & "What this means: The script could not be loaded at startup by EBA Command Center because Script.exe only supports EXE files." & line & "This script will not execute on startup.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.exe\OperationComplete") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.exe\OperationComplete").Files
					If LCase(Right(forVar1, 4)) = ".exe" Then
						newFolder(dataLoc & "\PluginData\Scripts\EndOp")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\EndOp\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.exe\OperationCompleted\" & forVar1.name & vblf & "Error Generated: ScriptEXEEncounteredNonEXE" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.exe only supports EXE files." & line & "This script will not execute after EndOp.")
					End If
				Next
			End If
			If foldExists(temp(0) & "\Scripts.exe\Shutdown") Then
				For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.exe\Shutdown").Files
					If LCase(Right(forVar1, 4)) = ".exe" Then
						newFolder(dataLoc & "\PluginData\Scripts\Shutdown")
						fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\Shutdown\" & forVar1.name
					Elseif LCase(forVar1.name) <> "desktop.ini" Then
						Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.exe\Shutdown\" & forVar1.name & vblf & "Error Generated: ScriptEXEEncounteredNonEXE" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.exe only supports EXE files." & line & "This script will not execute on shutdown.")
					End If
				Next
			End If
			For Each forVar1 In fs.GetFolder(temp(0) & "\Scripts.exe").Files
				If LCase(Right(forVar1, 4)) = ".exe" Then
					newFolder(dataLoc & "\PluginData\Scripts")
					fs.CopyFile forVar1, dataLoc & "\PluginData\Scripts\" & forVar1.name
				Elseif LCase(forVar1.name) <> "desktop.ini" Then
					Internal("Internal Exception in Plugin " & temp(0) & line & "Location: Scripts.exe\" & forVar1.name & vblf & "Error Generated: ScriptEXEEncounteredNonEXE" & vblf & "What this means: The script could not be loaded by EBA Command Center because Script.exe only supports EXE files." & line & "This script will not execute when referenced.")
				End If
			Next
		End If
		If foldExists(temp(0) & "\Files") Then
			newFolder(dataLoc & "\PluginData\Files")
			For Each forVar1 In fs.GetFolder(temp(0) & "\Files").Files
				If LCase(forVar1.name) <> "desktop.ini" Then fs.CopyFile forVar1, dataLoc & "\PluginData\Files\" & forVar1.name
			Next
		End If
	Next
	For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\Startup").Files
		cmd.run DblQuote(forVar)
	Next
End Sub
Sub shutdown(shutdownMethod)
	If shutdownMethod = "shutdown" Then
		cmd.run "shutdown /s /t " & shutdownTimer & " /f /c ""You requested a system shutdown in EBA Command Center."""
		Warn("Your PC will shut down in " & shutdownTimer & " seconds. Press OK to cancel.")
	Elseif shutdownMethod = "restart" Then
		cmd.run "shutdown /r /t " & shutdownTimer & " /f /c ""You requested a system restart in EBA Command Center."""
		Warn("Your PC will restart in " & shutdownTimer & " seconds. Press OK to cancel.")
	Elseif shutdownMethod = "hibernate" Then
		cmd.run "shutdown /h"
	Else
		cmd.run "shutdown /s /t 15 /f /c ""There was an issue with the shutdown method, so EBA Cmd will shutdown your PC in 15 seconds."""
		Warn("Your PC will shutdown in 15 seconds (due to an error with the shutdownMethod). Press OK to cancel.")
	End If
	cmd.run "shutdown /a"
End Sub
Sub update(dir,writeData,arg)
	If LCase(arg) = "overwrite" Then
		Call write(dir,writeData)
	Elseif LCase(arg) = "append" Then
		Call append(dir,writeData)
	Else
		If Not fExists(dir) Then
			Call write(dir,writeData)
		End If
	End If
End Sub
Sub updateCommands
	dwnld "https://eba-tools.github.io/data/cmd/EBA-8.5-beta2.vbs"
	If fExists(programLoc & "\tmp.ebacmd") Then
		delete(programLoc & "\EBA.vbs")
		fs.CopyFile programLoc & "\tmp.ebacmd", programLoc & "\EBA.vbs"
		delete(programLoc & "\tmp.ebacmd")
	Else
		error "The installer failed to download the requested version of EBA Command Center. Please check your connection to the internet and try again."
		Call endOp("c")
	End If
	dwnld "https://eba-tools.github.io/data/icon.ico"
	If fExists(programLoc & "\tmp.ebacmd") Then
		delete(programLoc & "\icon.ico")
		fs.CopyFile programLoc & "\tmp.ebacmd", programLoc & "\icon.ico"
		delete(programLoc & "\tmp.ebacmd")
	Else
		error "The installer failed to download the requested version of EBA Command Center. Please check your connection to the internet and try again."
		Call endOp("c")
	End If
	If Err.Number <> 0 Then
		error "The installer failed to download the requested version of EBA Command Center. Please check your connection to the internet and try again."
		Call endOp("c")
	End If
	
	fileDir = programLoc & "\Commands\admin.ebacmd"
	Call update(fileDir,"eba.admin","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\admin.ebacmd")
	
	fileDir = programLoc & "\Commands\backup.ebacmd"
	Call update(fileDir,"eba.backup","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\backup.ebacmd")
	
	fileDir = programLoc & "\Commands\config.ebacmd"
	Call update(fileDir,"eba.config","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"yes","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\config.ebacmd")
	
	fileDir = programLoc & "\Commands\crash.ebacmd"
	Call update(fileDir,"eba.crash","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\crash.ebacmd")
	
	fileDir = programLoc & "\Commands\dev.ebacmd"
	Call update(fileDir,"eba.dev","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\dev.ebacmd")
	
	fileDir = programLoc & "\Commands\end.ebacmd"
	Call update(fileDir,"eba.end","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\end.ebacmd")
	
	fileDir = programLoc & "\Commands\error.ebacmd"
	Call update(fileDir,"eba.error","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\error.ebacmd")
	
	fileDir = programLoc & "\Commands\export.ebacmd"
	Call update(fileDir,"eba.export","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\export.ebacmd")
	
	fileDir = programLoc & "\Commands\help.ebacmd"
	Call update(fileDir,"eba.help","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\help.ebacmd")
	
	fileDir = programLoc & "\Commands\import.ebacmd"
	Call update(fileDir,"eba.import","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\import.ebacmd")
	
	fileDir = programLoc & "\Commands\login.ebacmd"
	Call update(fileDir,"eba.login","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\login.ebacmd")
	
	fileDir = programLoc & "\Commands\logout.ebacmd"
	Call update(fileDir,"eba.logout","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\logout.ebacmd")
	
	fileDir = programLoc & "\Commands\logs.ebacmd"
	Call update(fileDir,logDir,"overwrite")
	Call update(fileDir,"file","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\logs.ebacmd")
	
	fileDir = programLoc & "\Commands\plugin.ebacmd"
	Call update(fileDir,"eba.plugin","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	
	fileDir = programLoc & "\Commands\read.ebacmd"
	Call update(fileDir,"eba.read","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\read.ebacmd")
	
	fileDir = programLoc & "\Commands\refresh.ebacmd"
	Call update(fileDir,"eba.refresh","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"yes","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\refresh.ebacmd")
	
	fileDir = programLoc & "\Commands\restart.ebacmd"
	Call update(fileDir,"eba.restart","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\restart.ebacmd")
	
	fileDir = programLoc & "\Commands\run.ebacmd"
	Call update(fileDir,"sys.run","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\run.ebacmd")
	
	fileDir = programLoc & "\Commands\shutdown.ebacmd"
	Call update(fileDir,"sys.shutdown","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\shutdown.ebacmd")
	
	fileDir = programLoc & "\Commands\uninstall.ebacmd"
	Call update(fileDir,"eba.uninstall","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"yes","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\uninstall.ebacmd")
	
	fileDir = programLoc & "\Commands\update.ebacmd"
	Call update(fileDir,"https://ethanblaisalarms.github.io/cmd","overwrite")
	Call update(fileDir,"url","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\update.ebacmd")
	
	fileDir = programLoc & "\Commands\upgrade.ebacmd"
	Call update(fileDir,"eba.upgrade","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"yes","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\upgrade.ebacmd")
	
	fileDir = programLoc & "\Commands\ver.ebacmd"
	Call update(fileDir,"eba.version","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\ver.ebacmd")
	
	fileDir = programLoc & "\Commands\version.ebacmd"
	Call update(fileDir,"eba.version","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\version.ebacmd")
	
	fileDir = programLoc & "\Commands\write.ebacmd"
	Call update(fileDir,"eba.write","overwrite")
	Call update(fileDir,"exe","append")
	Call update(fileDir,"no","append")
	Call update(fileDir,"builtin","append")
	delete(dataLoc & "\Commands\write.ebacmd")
End Sub
Sub write(dir,writeData)
	If fExists(dir) Then
		Set sys = fs.OpenTextFile (dir, 2)
		sys.writeLine writeData
		sys.Close
	Elseif foldExists(fs.GetParentFolderName(dir)) Then
		Set sys = fs.CreateTextFile (dir, 2)
		sys.writeLine writeData
		sys.Close
	Else
		Error "Given file not found: " & dir,"BAD_FILE_DIRECTORY"
	End If
End Sub

'Functions
Function alert(msg)
	Call addWarn
	alert = msgbox("Alert:" & line & msg,48,title)
End Function
Function checkCScript()
	WMI.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & Replace(scriptLoc,"\","\\") & "%' AND CommandLine LIKE '%CScript%'")
End Function
Function checkOS()
	For Each forVar in os
		checkOS = forVar.Caption
	Next
End Function
Function critical(msg,code)
	Call addError
	critical = msgbox("Critical:" & line & msg & line & "Error code: " & code,16,title)
End Function
Function DblQuote(str)
	DblQuote = Chr(34) & str & Chr(34)
End Function
Function db(msg)
	db = msgbox("Debug message:" & line & msg,64,"EBA Command Center | Debug")
End Function
Function delete(dir)
	If fExists(dir) Then
		fs.DeleteFile(dir)
	Elseif foldExists(dir) Then
		fs.DeleteFolder(dir)
	End If
End Function
Function dwnld(url)
	download.open "get", url, False
	download.send
	With stream
		.type = 1
		.open
		.write download.responseBody
		.savetofile programLoc & "\tmp.ebacmd"
		.close
	End With
End Function
Function error(msg,code)
	Call addError
	error = msgbox("Error:" & line & msg & line & "Error code: " & code,16,title)
End Function
Function fExists(dir)
	fExists = fs.FileExists(dir)
End Function
Function foldExists(dir)
	foldExists = fs.FolderExists(dir)
End Function
Function goOnline(url)
	https.open "get", url, False
	https.send
	goOnline = https.responseText
End Function
Function internal(msg,code)
	Call addError
	internal = msgbox("Internal Exception:" & line & msg & line & "Error code: " & code,48,title)
End Function
Function newFolder(dir)
	If Not foldExists(dir) Then
		If foldExists(fs.GetParentFolderName(dir)) Then
			newFolder = fs.CreateFolder(dir)
		End If
	End If
End Function
Function note(msg)
	Call addNote
	note = msgbox("Notice:" & line & msg,64,title)
End Function
Function scriptRunning()
	WMI.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & Replace(scriptLoc,"\","\\") & "%' AND CommandLine LIKE '%WScript%'")
End Function
Function warn(msg)
	Call addWarn
	warn = msgbox("Warning:" & line & msg,48,title)
End Function