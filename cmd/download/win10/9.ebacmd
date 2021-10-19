'EBA Command Center 9
'Copyright EBA Tools 2019 - 2021
Option Explicit
On Error Resume Next

'Define Variables
'  Objects
Dim app,backup1,backup2,cmd,download,fs,https,os,short,stream,sys,WMI,XML
Set app = CreateObject("Shell.Application")
Set backup1 = Nothing
Set backup2 = Nothing
Set cmd = CreateObject("Wscript.Shell")
Set download = CreateObject("Microsoft.XMLHTTP")
Set fs = CreateObject("Scripting.FileSystemObject")
Set https = CreateObject("msxml2.xmlhttp.3.0")
Set short = Nothing
Set stream = CreateObject("Adodb.Stream")
Set sys = Nothing
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set os = WMI.ExecQuery("Select * from Win32_OperatingSystem")
Set XML = CreateObject("Microsoft.XMLDOM")

'  Variables
Dim allowDebugMessages,count(4),curVer,data,dataLoc,dataLocTmp,defaultShutdown,desktop,errHandle,eba,enableEndOp,enableLegacyEndOp,endOpFail,exeValue,exeValueExt,fileDir,forVar,forVar1,forVar2,forVar3,forVar4,globTmp,htmlContent,importData,isAdmin,isDev,isInstalled,line,lines(5),logDir,logging,logIn,logInType,maxPlugin,missFiles,nowDate,nowTime,opDetails(7),pcName,pluginCount,prog,programLoc,pWord,regLoc,saveLogin,scriptDir,scriptLoc,shutdownTimer,skipDo,skipExe,startMenu,startup,startupType,status,temp(9),title,uName,user,userType,ver
count(0) = 1
count(4) = 0
dataLoc = cmd.ExpandEnvironmentStrings("%AppData%") & "\EBA"
dataLocTmp = dataLoc & ".tmp"
defaultShutdown = "shutdown"
desktop = cmd.SpecialFolders("AllUsersDesktop")
enableEndOp = 1
enableLegacyEndOp = False
endOpFail = False
exeValue = "eba.null"
exeValueExt = "eba.null"
isAdmin = True
isDev = False
isInstalled = False
line = vblf & "---------------------------------------" & vblf
logDir = dataLoc & "\EBA.log"
logging = False
missFiles = False
pcName = cmd.ExpandEnvironmentStrings("%ComputerName%")
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
title = setTitle("Debug")
user = "false"
userType = "false"
ver = 9

Call checkWS

newFolder dataLocTmp

Call checkWS

'Modifiable Startup Settings
allowDebugMessages = False
errHandle = True
maxPlugin = 10
title = setTitle("Debug")

If allowDebugMessages = True Then
	db "Debug messages have been enabled in the scripts ""Modifiable Startup Settings"", which is not recommended. We recommend only using this option for debug usage. To disable debug messages, set ""allowDebugMessages"" to False under Modifiable Startup Settings. If you need help, please contact us."
End If
If maxPlugin <> 10 Then
	db "Maximum plugins is not set to the recommended value of 10 in the scripts ""Modifiable Startup Settings"", which is not recommended. We recommend only using this option for debug usage. To set maximum plugins to 10, set ""maxPlugin"" to 10 under Modifiable Startup Settings. If you need help, please contact us."
End If
If maxPlugin > 20 Then
	critical "Maximum plugins is higher than the maximum value of 20 in the scripts ""Modifiable Startup Settings"", which is not recommended. We recommend only using this option for debug usage. To change maximum plugins, set ""maxPlugin"" to any value under Modifiable Startup Settings. If you need help, please contact us.","Plugins_Too_High"
	Call endOp("c")
End If
Dim loadedPlugins(19)
If errHandle = False Then
	db "Error handling has been disabled in the scripts ""Modifiable Startup Settings"", which is not recommended. We recommend only using this option for debug usage. To enable error handling, set ""errHandle"" to True under Modifiable Startup Settings. If you need help, please contact us."
End If

'Beginning Operations
Call clearCounts
Call clearLines
Call clearTemps
Call readSettings
Call checkPCName
Call checkWS

'Set Object Settings
XML.Async = "False"

'Check Admin
Call checkWS
cmd.RegRead("HKEY_USERS\s-1-5-19\")
If Err.Number <> 0 Then
	isAdmin = False
	dbm "EBA Command Center is not being ran as administrator."
Else
	isAdmin = True
	dbm "EBA Command Center is being ran as administrator."
End If
Err.Clear

If errHandle = False Then On Error GoTo 0

checkWS

'Check OS
temp(0) = LCase(checkOS())
dbm "OS Detected: " & checkOS
If InStr(temp(0),"microsoft") Then
	If InStr(temp(0),"windows") Then
		If InStr(temp(0),"11") or InStr(temp(0),"10") or InStr(temp(0),"7") or InStr(temp(0),"8") or InStr(temp(0),"vista") Then
			Call clearTemps
		Else
			critical checkOS & " does not support EBA Cmd.","Outdated_Windows_Version"
			Call endOp("c")
		End If
	Else
		critical "Windows Recovery Environment does not support EBA Cmd.","Windows_Recovery"
		Call endOp("c")
	End If
Else
	critical	checkOS & " does not support EBA Cmd.","Operating_System_Not_Supported"
	Call endOp("c")
End If

checkWS

'Get Startup Type
If fExists(dataLoc & "\startupType.ebacmd") Then
	Call read(dataLoc & "\startupType.ebacmd","l")
	startupType = data
Else
	If foldExists(dataLoc) Then
		startupType = "normal"
	Else
		startupType = "firstrun"
	End If
End If
If fExists(programLoc & "\EBA.vbs") Then
	If scriptLoc <> programLoc & "\EBA.vbs" Then startupType = "update"
Else
	startupType = "install"
End If
If LCase(scriptLoc) = LCase(startup & "\uninstallEBA.vbs") Then
	startupType = "uninstall"
End If
dbm "Startup type: " & startupType

checkWS

'Check Uninstallation
If fExists(cmd.SpecialFolders("Startup") & "\uninstallEBA.vbs") And startupType <> "uninstall" Then
	critical "EBA Command Center is set to uninstall. EBA Command Center cannot start, install, update, refresh, or repair right now. Please restart your PC to finalize or cancel uninstallation.","UNINSTALLION_SCHEDULED"
	Call endOp("c")
End If
dbm "EBA Command Center is not set to uninstall."

checkWS

'Get Imports
For Each forVar In Wscript.Arguments
	importData = forVar
Next
dbm "Imported Data: " & importData

checkWS

'Prep Plugins
Call preparePlugins

checkWS

'Check Imports
Call checkImports

'Launch
Do
	If startupType = "firstrepair" Then
		Call modeFirstrepair
	Elseif startupType = "firstrun" Then
		Call modeFirstrun
	Elseif startupType = "install" Then
		Call modeInstall
	Elseif startupType = "normal" Then
		Call modeNormal
	Elseif startupType = "recover" Then
		Call modeRecover
	Elseif startupType = "refresh" Then
		Call modeRefresh
	Elseif startupType = "repair" Then
		Call modeRepair
	Elseif startupType = "uninstall" Then
		Call modeUninstall
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
Sub modeFirstrepair
	title = setTitle("Recovery")
	Call checkWS
	
	Note("Hello!")
	Note("EBA Command Center is almost done repairing.")
	Note("All thats left to do is check if your User Account is functional.")
	Do
		eba = inputbox("Check your user accounts below. Afterwards, press Cancel to stop checking." & line & "Enter your Username:",title)
		If eba = "" Then
			Exit Do
		Elseif fExists(dataLoc & "\Users\" & eba & ".ebacmd") Then
			Call readLines(dataLoc & "\Users\" & eba & ".ebacmd",2)
			If LCase(lines(2)) = "owner" Then
				Note("That User Account exists on this device, and has administrator permissions.")
			Elseif LCase(lines(2)) = "admin" Then
				Note("That User Account exists on this device, and has administrator permissions. It was not created during inital setup.")
			Elseif LCase(lines(2)) = "general" Then
				Warn("That User Account exists, but the account does not have administrator permissions.")
			Else
				Warn "That User Account exists, but the account is corrupt."
			End If
		Else
			Warn("That User Account does not exist!")
		End If
	Loop
	
	eba = msgbox("Do you need to re-add an Administrator User Account?",4+32,title)
	If eba = vbYes Then
		Note("EBA Command Center will launch Initial Setup.")
		startupType = "firstrun"
		Exit Sub
	End If
	Note("EBA Command Center will restart.")
	Call endOp("r")
End Sub
Sub modeFirstrun
	title = setTitle("Initial Setup")
	Call checkWS
	
	Note("Welcome!")
	Note("Thanks for choosing EBA Command Center!")
	Note("We're about to perform initial setup.")
	Note("If this is your first time using EBA Command Center, we recommend checking out the EBA Wiki (on our website).")
	Note("Ok, enough chit-chat. Lets begin setup!")
	
	'Username
	Note("Lets begin with a User Account. Your account is stored locally on your PC.")
	
	prog = 1
	Do while prog = 1
		uName = inputbox("Type the username you want on the account:",title)
		If uName = "" Then
			eba = msgbox("Want to exit Initial Setup?",4+48,title)
			If eba = vbYes Then Call endOp("s")
		Elseif Len(uName) < 3 Then
			Warn("Too short! Usernames must be at least 3 characters long!")
		Elseif Len(uName) > 15 Then
			Warn("Too long! Usernames cannot be longer than 15 characters.")
		Else
			If inStr(1,uName,"\") > 0 Then
				Warn("Back-slash(\) is not allowed in usernames!")
			Elseif inStr(1,uName,"/") > 0 Then
				Warn("Slash(/) is not allowed in usernames!")
			Elseif inStr(1,uName,":") > 0 Then
				Warn("Colon(:) is not allowed in usernames!")
			Elseif inStr(1,uName,"*") > 0 Then
				Warn("Asterisk(*) is not allowed in usernames!")
			Elseif inStr(1,uName,"?") > 0 Then
				Warn("Question-mark(?) is not allowed in usernames!")
			Elseif inStr(1,uName,"""") > 0 Then
				Warn("Quote("") is not allowed in usernames!")
			Elseif inStr(1,uName,"<") > 0 Then
				Warn("Less-than(<) is not allowed in usernames!")
			Elseif inStr(1,uName,">") > 0 Then
				Warn("Greater-than(>) is not allowed in usernames!")
			Elseif inStr(1,uName,"|") > 0 Then
				Warn("Vertical-line(|) is not allowed in usernames!")
			Else
				prog = 2
			End If
		End If
	Loop
	
	'Password
	Do while prog = 2
		pWord = inputbox("Create a password for " & uName, title)
		If pWord = "" Then
			eba = msgbox("Continue without a password?", 4+48, title)
			If eba = vbYes Then
				prog = 3
			End If
		Else
			temp(0) = inputbox("Confirm password:", title)
			If temp(0) = pword Then
				prog = 3
			Else
				Warn("Passwords did not match.")
			End If
		End If
	Loop
	
	'Config
	Note("Your User Account has been set up! Now lets take a look at your preferences.")
	Call write(dataLoc & "\settings\logging.ebacmd","true")
	
	eba = msgbox("Do you want to enable this option?" & line & "SaveLogin | Saves your login status when you exit EBA Command Center.",4+32,title)
	If eba = vbYes Then
		Call write(dataLoc & "\settings\saveLogin.ebacmd","true")
	Else
		Call write(dataLoc & "\settings\saveLogin.ebacmd","false")
	End If
	
	Note("You can edit more settings in the Config menu. Advanced settings can be found in the Windows Registry (at " & regLoc & "). Be sure to check out the EBA Wiki for more details.")
	
	Call write(dataLoc & "\Users\" & uName & ".ebacmd",pWord & vblf & "owner")
	Call log("Critical Alert | New Admin Account created: " & uName)
	Call write(dataLoc & "\startupType.ebacmd","normal")
	Note("EBA Command Center has been set up! EBA Command Center will now load.")
	Call endOp("r")
End Sub
Sub modeInstall
	title = setTitle("Installation")
	Call checkWS
	If isAdmin = False Then Call endOp("fa")
	
	Call clearTemps
	
	'Installation
	eba = msgbox("EBA Command Center was not found on this device. You can chose to install a new copy, or search for an existing copy to update to EBA " & ver & ". Continue?",4+64,title)
	If eba = vbNo Then
		Note("The EBA Installer will now close.")
		Call endOp("c")
	End If
	
	'Install Directory
	programLoc = inputbox("Please enter the location where EBA Command Center should be installed. If the location given contains an installation of EBA Command Center, we'll update for you. Please note that the contents of the folder will be deleted if you install a new copy.",title,programLoc)
	programLoc = Replace(programLoc,"""","")
	If Not foldExists(fs.GetParentFolderName(programLoc)) Then
		Error "The directory does not exist: " & fs.GetParentFolderName(programLoc),"Parent_Folder_Not_Found"
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
	cmd.Regwrite regLoc & "\enableOperationCompletedMenu", enableEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\enableLegacyOperationCompletedMenu", enableLegacyEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\installDir", programLoc, "REG_SZ"
	
	'Folders
	delete("C:\EBA")
	delete("C:\EBA-Installer")
	delete(programLoc)
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
	Call downloadFiles
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
	fs.CopyFile scriptLoc, programLoc & "\EBA.vbs"
	
	'Create Icons
	Call createShortcut(desktop & "\EBA Command Center.lnk")
	Call createShortcut(startMenu & "\EBA Command Center.lnk")
	
	'Installed!
	eba = msgbox("EBA Command Center finished installing! Do you want to launch EBA Command Center, and perform Initial Setup now?",4+32,title)
	If eba = vbYes Then Call endOp("r")
	Call endOp("c")
End Sub
Sub modeNormal
	title = setTitle("Debug")
	Call checkWS
	
	
	curVer = CDbl(goOnline("https://eba-tools.github.io/cmd/version.txt"))
	
	dbm "Current Version: " & ver & "..." & vblf & "Latest Version: " & curVer & "..."
	
	title = Replace(setTitle("")," |","")
	If ver < curVer Then
		eba = msgbox("Notice:" & vblf & "An update for EBA Command Center is available! Do you want to update now?",4+64,title)
		If eba = vbYes Then
			delete dataLocTmp & "\installer.vbs"
			dwnld "https://eba-tools.github.io/data/cmd/EBA-" & curVer & ".vbs",dataLocTmp & "\installer.vbs"
			If fExists(dataLocTmp & "\installer.vbs") Then
				Note "The installer was downloaded from the internet. Click OK to continue with the update."
				cmd.run DblQuote(dataLocTmp & "\installer.vbs")
				Call endOp("f")
			Else
				Error "Failed to download the latest version of EBA Command Center. Please check your internet connection and try again."
			End If
		End If
	Elseif ver > curVer Then
		Note("Thanks for using EBA Command Center Beta!" & line & "Current Version: " & ver)
	End If
	
	'Data File Checks
	Call dataExists(programLoc & "\EBA.vbs")
	Call dataExists(programLoc & "\Commands\config.ebacmd")
	Call dataExists(programLoc & "\Commands\end.ebacmd")
	Call dataExists(programLoc & "\Commands\login.ebacmd")
	
	If Not missFiles = False Then
		skipDo = True
		eba = msgbox("EBA Command Center didn't start correctly." & line & "'ABORT': Exit EBA Command Center." & vblf & "'RETRY': Restart EBA Cmd." & vblf & "'IGNORE': Continue to recovery.",2+16,"EBA Cmd " & ver & " | StartFail")
		If eba = vbAbort Then Call endOp("c")
		If eba = vbRetry Then Call endOp("r")
		If eba = vbIgnore Then
			eba = LCase(inputbox("Select recovery options:" & line & "'START': Bypass this menu and start EBA Command Center" & vblf & "'RETRY': Restart EBA Command Center" & vblf & "'RECOVERY': Start EBA Command Center in Recovery Mode." & vblf & "'AUTO': Start automatic repair.",title))
			If eba = "retry" Then
				Call endOp("r")
			Elseif eba = "recovery" Then
				startupType = "recover"
				skipDo = True
			Elseif eba = "auto" Then
				startupType = "repair"
				skipDo = True
			Elseif eba = "start" Then
				eba = msgbox("Warning:" & line & "EBA Command Center didnt start correctly. We recommend running recovery options instead of starting. Continue anyways?",4+48,title)
				If eba = vbYes Then skipDo = False
			End If
		End If
	End If
	
	If skipDo = False Then		
		Call checkWS
		Call clearTemps
		
		Call write(dataLoc & "\connect.ebacmd",1)
		
		If Not fExists(logDir) Then Call log("Log File Created")
		
		If saveLogin = "false" Then Call write(dataLoc & "\isLoggedIn.ebacmd",vblf)
		
		delete(dataLoc & "\susActivity.ebacmd")
		
		Call runPlugins
		
		eba = msgbox("Start EBA Command Center " & ver & "?",4+32,title)
		If eba = vbNo Then Call endOp("c")
		Call log(title & " was launched.")
		'Call write(dataLoc & "\secureShutdown.ebacmd","false")
	End If
	
	Call checkWS
	
	Do
		If skipDo = True Then Exit Do
		If Not Err.Number = 0 Then
			Error "A critical error occurred within EBA Cmd. Crashing...","WS/" & Err.Number & "?Mode=CriticalError"
			Call endOp("c")
		End If
		
		Call dataExists(programLoc & "\EBA.vbs")
		Call dataExists(programLoc & "\Commands\config.ebacmd")
		Call dataExists(programLoc & "\Commands\end.ebacmd")
		Call dataExists(programLoc & "\Commands\login.ebacmd")
		
		If Not missFiles = False Then
			eba = msgbox("A critical error occurred within EBA Command Center. We recommend closing EBA Command Center. Close now?",4+16,title)
			If eba = vbYes Then Call endOp("c")
		End If
		
		Call readLines(dataLoc & "\isLoggedIn.ebacmd",2)
		logIn = lines(1)
		logInType = lines(2)
		If logIn = "" Then
			status = "Not Logged In"
		Else
			status = "Logged In: " & logIn
		End If
		
		'User Input
		If skipExe = false Then
			eba = LCase(inputbox("Enter Command Below:" & vblf & "EBA" & line & status, title))
			exeValue = "eba.null"
			If eba = "" Then eba = "end"
			If fExists(dataLoc & "\Commands\" & eba & ".ebacmd") Then
				Call readCommands(dataLoc)
			Elseif fExists(programLoc & "\Commands\" & eba & ".ebacmd") Then
				Call readCommands(programLoc)
			Elseif fExists(dataLoc & "\PluginData\Commands\" & eba & ".ebacmd") Then
				Call readCommands(dataLoc & "\PluginData")
			Else
				Error "That command could not be found or is corrupt.","INVALID_COMMAND"
			End If
			Call log("Command Executed: " & eba)
		Else
			exeValue = skipExe
			skipExe = false
		End If
		
		'Execution Values
		If exeValue = "eba.admin" Then
			If isAdmin = False Then
				Call endOp("ra")
			End If
			Note("EBA Command Center is already running as administrator.")
		Elseif exeValue = "eba.backup" Then
			eba = msgbox("Your backup will be saved to " & dataLoc & "\backup.ebabackup" & line & "Note that the file at that location will be overwrote. Continue?",4+32,title)
			If eba = vbYes Then
				eba = LCase(inputbox("What type of backup do you want to run?" & line & "'USER': Backs up user accounts." & vblf & "'CMD': Backs up commands." & vblf & "'SETTINGS': Backs up settings." & vblf & "'PLUG': Backs up plugins.",title))
				If eba = "user" or eba = "cmd" or eba = "settings" or eba = "plug" Then
					If fExists(dataLoc & "\backup.ebabackup") Then fs.DeleteFile(dataLoc & "\backup.ebabackup")
					Call checkWS
					If Not fExists(dataLoc & "\backup.zip") Then Call write(dataLoc & "\backup.zip", Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0))
					temp(0) = fs.GetAbsolutePathName(dataLoc & "\backup.zip")
					If eba = "user" Then
						Set backup1 = objApp.NameSpace(temp(0))
						temp(1) = fs.GetAbsolutePathName(dataLoc & "\Users")
						Set backup2 = objApp.NameSpace(temp(1))
						backup1.CopyHere backup2.items, 4
						If Err.Number = 0 Then
							Note("Backed up all files in " & dataLoc & "\Users")
						Else
							Error "Backup failed. See WScript Error for more info.","WS/" & Err.Number & "?Mode=BackupError"
						End If
						Call checkWS
					Elseif eba = "cmd" Then
						Set backup1 = objApp.NameSpace(dataLoc & "\backup.zip")
						temp(1) = fs.GetAbsolutePathName(dataLoc & "\Commands")
						Set backup2 = objApp.NameSpace(temp(1))
						backup1.CopyHere backup2.items, 4
						If Err.Number = 0 Then
							Note("Backed up all files in " & dataLoc & "\Commands")
						Else
							Error "Backup failed. See WScript Error for more info.","WS/" & Err.Number & "?Mode=BackupError"
						End If
						Call checkWS
					Elseif eba = "settings" Then
						Set backup1 = objApp.NameSpace(dataLoc & "\backup.zip")
						temp(1) = fs.GetAbsolutePathName(dataLoc & "\Settings")
						Set backup2 = objApp.NameSpace(temp(1))
						backup1.CopyHere backup2.items, 4
						If Err.Number = 0 Then
							Note("Backed up all files in " & dataLoc & "\Settings")
						Else
							Error "Backup failed. See WScript Error for more info.","WS/" & Err.Number & "?Mode=BackupError"
						End If
						Call checkWS
					Elseif eba = "plug" Then
						Set backup1 = objApp.NameSpace(dataLoc & "\backup.zip")
						temp(1) = fs.GetAbsolutePathName(dataLoc & "\Plugins")
						Set backup2 = objApp.NameSpace(temp(1))
						backup1.CopyHere backup2.items, 4
						If Err.Number = 0 Then
							Note("Backed up all files in " & dataLoc & "\Plugins")
						Else
							Error "Backup failed. See WScript Error for more info.","WS/" & Err.Number & "?Mode=BackupError"
						End If
						Call checkWS
					End If
					If fExists(dataLoc & "\backup.zip") Then fs.MoveFile dataLoc & "\backup.zip", dataLoc & "\backup.ebabackup"
				Else
					Warn("Invalid argument.")
				End If
			End If
		Elseif exeValue = "eba.config" Then
			If exeValueExt = "eba.cmd" Then
				eba = "cmd"
			Elseif exeValueExt = "eba.cmdnew" Then
				eba = "cmd"
			Elseif exeValueExt = "eba.cmdedit" Then
				eba = "cmd"
			Elseif exeValueExt = "eba.acc" Then
				eba = "acc"
			Elseif exeValueExt = "eba.accnew" Then
				eba = "acc"
			Elseif exeValueExt = "eba.accedit" Then
				eba = "acc"
			Elseif exeValueExt = "eba.defaultshutdown" Then
				eba = "defaultshutdown"
			Elseif exeValueExt = "eba.logs" Then
				eba = "logs"
			Elseif exeValueExt = "eba.savelogin" Then
				eba = "savelogin"
			Elseif exeValueExt = "eba.shutdowntimer" Then
				eba = "shutdowntimer"
			Elseif exeValueExt = "eba.adv" Then
				eba = "advanced"
			Elseif exeValueExt = "eba.null" Then
				eba = LCase(inputbox("EBA Config:" & vblf & "EBA > Config" & line & status, title))
			Else
				Error "Unknown Exe Value Extension." & vblf & exeValueExt,"INVALID_EXE_EXT"
			End If
			If eba = "cmd" Then
				If exeValueExt = "eba.cmd" or exeValueExt = "eba.null" Then
					eba = LCase(inputbox("Modify Commands:" & vblf & "EBA > Config > Commands" & line & status, title))
				Elseif exeValueExt = "eba.cmdnew" Then
					eba = "new"
				Elseif exeValueExt = "eba.cmdedit" Then
					eba = "edit"
				Else
					Error "Unknown Error","INVALID_EXE_EXT"
				End If
				If eba = "new" Then
					status = "This is what you will type to execute the command."
					eba = LCase(inputbox("Create Command Below:" & vblf & "EBA > Config > Commands > New" & line & status, title))
					If fExists(dataLoc & "\Commands\" & eba & ".ebacmd") or fExists(programLoc & "\Commands\" & eba & ".ebacmd") Then
						Error "That command already exists.","DUPLICATE_FILE_DETECTED"
					ElseIf inStr(1,eba,"\") > 0 Then
						Warn("""\"" is not allowed in command names!")
					Elseif inStr(1,eba,"/") > 0 Then
						Warn("""/"" is not allowed in command names!")
					Elseif inStr(1,eba,":") > 0 Then
						Warn(""":"" is not allowed in command names!")
					Elseif inStr(1,eba,"*") > 0 Then
						Warn("""*"" is not allowed in command names!")
					Elseif inStr(1,eba,"?") > 0 Then
						Warn("""?"" is not allowed in command names!")
					Elseif inStr(1,eba,"""") > 0 Then
						Warn("' "" ' is not allowed in command names!")
					Elseif inStr(1,eba,"<") > 0 Then
						Warn("""<"" is not allowed in command names!")
					Elseif inStr(1,eba,">") > 0 Then
						Warn(""">"" is not allowed in command names!")
					Elseif inStr(1,eba,"|") > 0 Then
						Warn("""|"" is not allowed in command names!")
					Else
						temp(0) = false
						temp(3) = eba
						eba = LCase(inputbox("What is the type?" & line & "'CMD': Execute a command" & vblf & "'FILE': Execute a file" & vblf & "'URL': Web shortcut" & vblf & "'SHORT': Shortcut to another command", title))
						If eba = "cmd" Then
							temp(0) = True
							temp(1) = "cmd"
							temp(2) = LCase(inputbox("Type the command to execute:",title))
						Elseif eba = "file" Then
							temp(1) = "file"
							temp(2) = LCase(inputbox("Type the target file/folder:",title))
							temp(2) = Replace(temp(2),"""","")
							If fExists(temp(2)) or foldExists(temp(2)) Then
								temp(0) = True
							Else
								Error "The target file was not found.","BAD_DIRECTORY"
							End If
						Elseif eba = "url" Then
							temp(0) = True
							temp(1) = "url"
							temp(2) = LCase(inputbox("Type the URL below. Include https://",title,"https://example.com"))
						Elseif eba = "short" Then
							temp(1) = "short"
							temp(2) = LCase(inputbox("Type the target command below:",title))
							If fExists(dataLoc & "\Commands\" & temp(2) & ".ebacmd") Then
								temp(0) = True
							Else
								Error "The target command was not found or is corrupt.","INVALID COMMAND"
							End If
						Elseif eba = "exe" Then
							temp(0) = True
							temp(1) = "exe"
							temp(2) = LCase(inputbox("Type the execution value below:",title))
						End If
						If temp(0) = False Then
							Warn("The command could not be created.")
						Else
							If temp(1) = "short" Then
								temp(4) = "no"
							Else
								eba = msgbox("Require administrator login to execute?",4+32,title)
								If eba = vbNo Then
									temp(4) = "no"
								Else
									temp(4) = "yes"
								End If
							End If
							eba = msgbox("Confirm the command:" & line & "Name: " & temp(3) & vblf & "Type: " & temp(1) & vblf & "Target: " & temp(2) & vblf & "Login Required: " & temp(4),4+32,title)
							If eba = vbNo Then
								Warn("Creation of command canceled.")
							Else
								Call log("Command Created: " & temp(3))
								Call write(dataLoc & "\Commands\" & temp(3) & ".ebacmd",temp(2) & vblf & temp(1) & vblf & temp(4) & vblf & "no")
							End If
						End If
					End If
				Elseif eba = "edit" Then
					eba = LCase(inputbox("Enter Command Below:" & vblf & "EBA > Config > Commands > Modify" & line & status, title))
					If fExists(dataLoc & "\Commands\" & eba & ".ebacmd") Then
						temp(1) = eba
						Call readLines(dataLoc & "\Commands\" & eba & ".ebacmd",4)
						temp(0) = True
						If LCase(lines(4)) = "builtin" Then
							eba = msgbox("Warning:" & line & "That is a built-in command. If you modify this command, it could mess up EBA Command Center. Continue?",4+48,title)
							If eba = vbNo Then temp(0) = False
						End If
						If temp(0) = True Then
							eba = LCase(inputbox("What do you want to modify?" & line & "'TARGET': Edit the target" & vblf & "'NAME': Rename the command" & vblf & "'LOGIN': Change login requirements" & vblf & "'DELETE': Delete the command.",title))
							If eba = "target" Then
								temp(2) = "target"
								temp(3) = LCase(inputbox("Enter new target:",title,lines(1)))
								lines(1) = temp(3)
								temp(4) = True
							Elseif eba = "name" Then
								temp(2) = "name"
								temp(3) = LCase(inputbox("Enter new name:",title,temp(1)))
								temp(4) = True
							Elseif eba = "login" Then
								temp(2) = "login"
								temp(3) = msgbox("Require login to execute?",4+32,title)
								If temp(3) = vbNo Then
									temp(3) = "no"
								Else
									temp(3) = "yes"
								End If
								lines(3) = temp(3)
								temp(4) = True
							Elseif eba = "delete" Then
								temp(2) = "delete"
								eba = msgbox("Warning:" & line & "Deleting a command cannot be undone. Delete anyways?",4+48,title)
								If eba = vbYes Then
									fs.DeleteFile(dataLoc & "\Commands\" & temp(1) & ".ebacmd")
									Call log("Command deleted: " & temp(1))
									temp(4) = True
								End If
							End If
							If temp(4) = True Then
								If Not temp(2) = "delete" Then
									eba = msgbox("Confirm command modification:" & line & "Modification: " & temp(2) & vblf & "New Value: " & temp(3),4+32,title)
									If eba = vbYes Then
										If temp(2) = "name" Then
											fs.MoveFile dataLoc & "\Commands\" & temp(1) & ".ebacmd", dataLoc & "\Commands\" & temp(3) & ".ebacmd"
											Call log("Command renamed from " & temp(1) & " to " & temp(3))
										Else
											Call write(dataLoc & "\Commands\" & temp(1) & ".ebacmd",lines(1) & vblf & lines(2) & vblf & lines(3) & vblf & lines(4))
											Call log("Command Modified: " & temp(1))
										End If
									End If
								End If
							Else
								Warn("The command could not be modified.")
							End If
						End If
					Else
						Error "Either that command does not exist, or it is a built-in command.","INVALID_COMMAND"
					End If
				Else
					Error "Config option not found.","INVALID_ARGUMENT"
				End If
			Elseif eba = "acc" or eba = "account" Then
				If exeValueExt = "eba.acc" or exeValueExt = "eba.null" Then
					eba = LCase(inputbox("Modify Accounts:" & vblf & "EBA > Config > Accounts" & line & status, title))
				Elseif exeValueExt = "eba.accnew" Then
					eba = "new"
				Elseif exeValueExt = "eba.accedit" Then
					eba = "edit"
				Else
					Error "Unknown Error","UNKNOWN_ERROR"
				End If
				If eba = "new" Then
					temp(0) = fs.GetFolder(dataLoc & "\Users").Files.Count
					If temp(0) < 100 Then
						eba = inputbox("You are using " & temp(0) & " of " & temp(1) & " accounts." & line & "Create a username:",title)
						uName = eba
						If fExists(dataLoc & "\Users\" & eba & ".ebacmd") Then
							Error "That user already exists.","DUPLICATE_FILE_DETECTED"
						Elseif Len(uName) < 3 Then
							Warn("That username is too short!")
						Elseif Len(uName) > 15 Then
							Warn("That username is too long!")
						Elseif inStr(1,uName,"\") > 0 Then
							Warn("""\"" is not allowed in usernames!")
						Elseif inStr(1,uName,"/") > 0 Then
							Warn("""/"" is not allowed in usernames!")
						Elseif inStr(1,uName,":") > 0 Then
							Warn(""":"" is not allowed in usernames!")
						Elseif inStr(1,uName,"*") > 0 Then
							Warn("""*"" is not allowed in usernames!")
						Elseif inStr(1,uName,"?") > 0 Then
							Warn("""?"" is not allowed in usernames!")
						Elseif inStr(1,uName,"""") > 0 Then
							Warn("' "" ' is not allowed in usernames!")
						Elseif inStr(1,uName,"<") > 0 Then
							Warn("""<"" is not allowed in usernames!")
						Elseif inStr(1,uName,">") > 0 Then
							Warn(""">"" is not allowed in usernames!")
						Elseif inStr(1,uName,"|") > 0 Then
							Warn("""|"" is not allowed in usernames!")
						Else
							pWord = inputbox("Create a password for " & uName,title)
							If pWord = "" Then
								eba = msgbox("Continue without a password?",4+48,title)
								If eba = vbYes Then
									eba = msgbox("Make this an administrator account?",4+32+256,title)
									If eba = vbYes Then
										Call write(dataLoc & "\Users\" & uName & ".ebacmd",pWord & vblf & "admin")
										Call log("New administrator account created: " & uName)
									Else
										Call write(dataLoc & "\Users\" & uName & ".ebacmd",pword & vblf & "general")
										Call log("New account created: " & uName)
									End If
								End If
							Elseif Len(pWord) < 8 Then
								Warn("Password is too short.")
							Elseif Len(pWord) > 30 Then
								Warn("Password is too long.")
							Elseif InStr(1,lcase(pWord),"password") > 0 or Instr(1,pword,"1234") > 0 Then
								Warn("Ok, really..." & nls & "Ok, that password is too easy to figure out. Choose a strong password with a mix of letters and numbers.")
							Else
								eba = inputbox("Confirm password:",title)
								If eba = pWord Then
									eba = msgbox("Make this an administrator account?",4+32+256,title)
									If eba = vbYes Then
										Call write(dataLoc & "\Users\" & uName & ".ebacmd",pWord & vblf & "admin")
										Call log("New administrator account created: " & uName)
									Else
										Call write(dataLoc & "\Users\" & uName & ".ebacmd",pword & vblf & "general")
										Call log("New account created: " & uName)
									End If
								Else
									Error "Passwords do not match.","PASSWORDS_NO_MATCH"
								End If
							End If
						End If
					Else
						Error "EBA Command Center has an account limit of 100. You are using " & temp(0) & " accounts, and cannot add more.","TOO_MANY_ACCOUNTS"
					End If
				Elseif eba = "edit" Then
					eba = inputbox("Enter the username:",title)
					If fExists(dataLoc & "\Users\" & eba & ".ebacmd") Then
						Call readLines(dataLoc & "\Users\" & eba & ".ebacmd",2)
						temp(0) = eba
						eba = LCase(inputbox("What do you want to modify?" & line & "'PWORD': Change password" & vblf & "'ADMIN': Change admin status" & vblf & "'DELETE': Delete account",title))
						If eba = "pword" Then
							eba = inputbox("Enter current password:",title)
							If eba = lines(1) Then
								pWord = inputbox("Create new password:",title)
								If pWord = "" Then
									eba = msgbox("Continue without a password?",4+48,title)
									If eba = vbYes Then
										Call write(dataLoc & "\Users\" & temp(0) & ".ebacmd",pWord & vblf & lines(2))
										Call log("Password changed for " & temp(0))
									End If
								Elseif Len(pWord) < 8 Then
									Warn("Password is too short.")
								Elseif Len(pWord) > 30 Then
									Warn("Password is too long.")
								Elseif InStr(1,lcase(pWord),"password") > 0 or Instr(1,pword,"1234") > 0 Then
									Warn("Ok, really..." & nls & "Ok, that password is too easy to figure out. Choose a strong password with a mix of letters and numbers.")
								Else
									eba = inputbox("Confirm password:",title)
									If eba = pWord Then
										Call write(dataLoc & "\Users\" & temp(0) & ".ebacmd",pWord & vblf & lines(2))
										Call log("Password changed for " & temp(0))
									Else
										Error "Passwords did not match.","PASSWORD_NO_MATCH"
									End If
								End If
							Else
								Error "Incorrect password.","INCORRECT_PASSWORD"
							End If
						Elseif eba = "admin" Then
							If lines(2) = "owner" Then
								Warn("That modification cannot be applied to this account. This is the account that was created on setup.")
							Else
								eba = msgbox("Make this account an administrator?",4+32+256,title)
								If eba = vbNo Then
									Call write(dataLoc & "\Users\" & temp(0) & ".ebacmd",lines(1) & vblf & "general")
									Call log("Made " & temp(0) & " a general account.")
								Else
									Call write(dataLoc & "\Users\" & temp(0) & ".ebacmd",lines(1) & vblf & "admin")
									Call log("Made " & temp(0) & " an administrator.")
								End If
							End If
						Elseif eba = "delete" Then
							If lines(2) = "owner" Then
								Warn("That modification cannot be applied to this account. This is the account that was created on setup.")
							Else
								eba = msgbox("Confirm delete?",4+32+256,title)
								If eba = vbYes Then
									fs.DeleteFile(dataLoc & "\Users\" & temp(0) & ".ebacmd")
									Call log("Account deleted: " & temp(0))
								End If
							End If
						Else
							Error "Config option not found.","INVALID_ARGUMENT"
						End If
					Else
						Error "Username not found.","FILE_NOT_FOUND"
					End If
				Else
					Error "Config option not found.","INVALID_ARGUMENT"
				End If
			Elseif eba = "logs" Then
				eba = msgbox("Logs are set to " & logging & ". Would you like to enable EBA Logs? (EBA Command Center will restart)", 4+32, title)
				If eba = vbYes Then
					Call write(dataLoc & "\settings\logging.ebacmd","true")
					Call log("Logging enabled by " & logIn)
				Else
					Call write(dataLoc & "\settings\logging.ebacmd","false")
					Call log("Logging disabled by " & logIn)
				End If
				Call endOp("r")
			Elseif eba = "savelogin" Then
				eba = msgbox("Save Login are set to " & saveLogin & ". Would you like to enable Save Login? (EBA Command Center will restart)", 4+32, title)
				If eba = vbYes Then
					Call write(dataLoc & "\settings\saveLogin.ebacmd","true")
					Call log("Save Login enabled by " & logIn)
				Else
					Call write(dataLoc & "\settings\saveLogin.ebacmd","false")
					Call log("Save Login disabled by " & logIn)
				End If
				Call endOp("r")
			Elseif eba = "shutdowntimer" Then
				eba = inputbox("Shutdown Timer is currently set to " & shutdownTimer & ". Please set a new value (must be at least 0, and must be an integer). EBA Command Center will restart.",title,10)
				If eba = "" Then eba = 0
				Call checkWS
				If CInt(eba) > -1 Then
					If Err.Number = 0 Then
						Call write(dataLoc & "\settings\shutdownTimer.ebacmd",eba)
						Call endOp("r")
					Else
						Warn("A WScript Error occurred while converting that value to an integer. Your settings were not changed.")
					End If
				Else
					Warn("That value didnt work. " & eba & " is not a positive integer.")
				End If
			Elseif eba = "defaultshutdown" Then
				eba = LCase(inputbox("Default Shutdown Method is currently set to " & defaultShutdown & ". Please set a new value:" & line & "'SHUTDOWN', 'RESTART', or 'HIBERNATE'. EBA Command Center will restart.",title,"shutdown"))
				If eba = "" Then eba = "shutdown"
				If eba = "shutdown" or eba = "restart" or eba = "hibernate" Then
					Call write(dataLoc & "\settings\defaultShutdown.ebacmd",eba)
					Call endOp("r")
				Else
					Error "That value is not valid. Nothing was changed.","INVALID_ARGUMENT"
				End If
			Elseif eba = "advanced" or eba = "adv" Then
				eba = msgbox("You can modify advanced EBA Command Center settings using Registry Editor. When you open Registry Editor, navigate to HKEY_LOCAL_MACHINE/SOFTWARE/EBA-Cmd to find advanced settings. Be careful, modifying the registry incorrectly could break EBA Command Center or other apps." & vblf & "Open Registry Editor?",4+48,title)
				If eba = vbYes Then
					cmd.run "regedit.exe"
				End If
			Else
				Error "Config option not found.","INVALID_ARGUMENT"
			End If
		Elseif exeValue = "eba.crash" Then
			wscript.sleep 2500
			msgbox "EBA Command Center just crashed! Please restart EBA Command Center.",16,"EBA Crash Handler"
			Call endOp("c")
		Elseif exeValue = "eba.dev" Then
			If isDev = true Then
				isDev = false
				Call log("Dev mode disabled")
				Warn("Developer Mode has been disabled. EBA Command Center will now restart.")
				Call endOp("r")
			ElseIf isDev = false Then
				isDev = true
				title = "EBA Command Center - Developer Mode"
				Call log("Dev mode enabled")
				Warn("Developer Mode has been enabled.")
			End If
		Elseif exeValue = "eba.end" Then
			eba = msgbox("Exit EBA Command Center?",4+32,title)
			If eba = vbYes Then Call endOp("s")
		Elseif exeValue = "eba.error" Then
			Warn("WScript Errors have been enabled. If you encounter a WScript error, EBA Command Center will crash. To disable WScript Errors, restart EBA Command Center.")
			On Error GoTo 0
		Elseif exeValue = "eba.export" Then
			eba = LCase(inputbox("EBA Config:" & vblf & "EBA > Export" & line & status, title))
			If fExists(dataLoc & "\Commands\" & eba & ".ebacmd") Then
				temp(0) = eba
				eba = inputbox("Where do you want the exported file?",title,desktop)
				eba = Replace(eba,"""","")
				If foldExists(eba) Then
					Call readLines(dataLoc & "\Commands\" & temp(0) & ".ebacmd",3)
					Call write(eba & "\EBA_Export.ebaimport","Type: Command" & vblf & temp(0) & vblf & lines(2) & vblf & lines(1) & vblf & lines(3))
					Call log("Command Exported: " & temp(0))
				Else
					Error "Cannot export to the given location.","BAD_DIRECTORY"
				End If
			Else
				Error "Command does not exist.","INVALID_COMMAND"
			End If
		Elseif exeValue = "eba.help" Then
			Note("If you need help, please contact us:" & vblf & "https://ethanblaisalarms.github.io/cmd/contactus")
		Elseif exeValue = "eba.import" Then
			importData = inputbox("Enter the path of the file you want to import.",title)
			importData = Replace(importData,"""","")
			If fExists(importData) Then
				Call checkImports
			Else
				Error "Path not found.","FILE_NOT_FOUND"
			End If
		Elseif exeValue = "eba.login" Then
			uName = inputbox("Enter your username:",title)
			If fExists(dataLoc & "\Users\" & uName & ".ebacmd") Then
				Call readLines(dataLoc & "\Users\" & uName & ".ebacmd",2)
				If Not lines(1) = "" Then
					pWord = inputbox("Enter the password:",title)
					If pWord = lines(1) Then
						Call log("Logged in: " & uName)
						Note("Logged in as " & uName)
						Call write(dataLoc & "\isLoggedIn.ebacmd",uName & vblf & lines(2))
					Else
						Call log("Failed to log in: " & uName)
						Error "Incorrect Password.","INCORRECT_PASSWORD"
					End If
				Else
					Call log("Logged in: " & uName)
					Note("Logged in as " & uName)
					Call write(dataLoc & "\isLoggedIn.ebacmd",uName & vblf & lines(2))
				End If
			Else
				Error "Username not found.","USERNAME_NOT_FOUND"
			End If
		Elseif exeValue = "eba.logout" Then
			Call write(dataLoc & "\isLoggedIn.ebacmd","" & vblf & "")
			Call log("Logged out all accounts")
			Note("Logged out.")
		Elseif exeValue = "eba.null" Then
			exeValue = "eba.null"
		Elseif exeValue = "eba.plugin" Then
			temp(0) = "Currently Loaded Plugins:" & line
			For Each forVar In fs.GetFolder(dataLoc & "\Plugins").Subfolders
				If fExists(forVar & "\meta.xml") Then
					XML.load(forVar & "\meta.xml")
					Call checkWS
					For Each forVar1 In XML.selectNodes("/Meta/Format")
						Call checkWS
						If forVar1.text = "1" Then
							For Each forVar2 In XML.selectNodes("/Meta/Version/DisplayName")
								temp(0) = temp(0) & forVar2.text & vblf
							Next
						Else
							Error "The plugin at " & forVar & " contains an invalid META.XML file, and will not be displayed.","UNKNOWN_FORMAT_VERSION"
						End If
					Next
					
				Else
					Warn("The plugin at " & forVar & " is missing META.XML, and will not be displayed.")
				End If
			Next
			msgbox temp(0),64,title
		Elseif exeValue = "eba.read" Then
			If isDev = false Then
				Error "This command can only be ran in EBA Developer Mode!","DEV_DISABLE"
			Else
				eba = inputbox("EBA > Call read", title)
				eba = Replace(eba,"""","")
				If fExists(eba) Then
					Call read(eba,"n")
					Call log("File Call read: " & eba)
					msgbox "EBA > Call read > " & eba & line & data,0,title
				Else
					Call log("Failed to Call read " & eba)
					Error "File " & eba & " not found!","FILE_NOT_FOUND"
				End If
			End If
		Elseif exeValue = "eba.refresh" Then
			If isDev = false Then
				Error "This command can only be used in EBA Developer Mode!","DEV_DISABLED"
			Else
				eba = msgbox("EBA Command Center will restart and open in reinstall mode.", 48, title)
				Call write(dataLoc & "\startupType.ebacmd","refresh")
				Call endOp("r")
			End If
		Elseif exeValue = "eba.restart" Then
			Call endOp("r")
		Elseif exeValue = "eba.reset" Then
			eba = msgbox("Are you sure you want to reset your PC?",4+48,title)
			If eba = vbYes Then
				eba = msgbox("This cannot be undone. Resetting your PC will uninstall all apps, reset all settings, and delete your files! Proceed?",4+48,title)
				If eba = vbYes Then
					cmd.run "systemreset"
					Note("Your PC is being reset. Follow all on-screen prompts. Press OK to cancel.")
				End If
			End If
		Elseif exeValue = "sys.run" Then
			eba = inputbox("Please enter the file, folder, or command you would like to execute:", title)
			temp(1) = Replace(eba,"""","")
			If fExists(temp(1)) Then
				cmd.run DblQuote(temp(1))
				Call log("File Executed: " & eba)
			Elseif foldExists(temp(1)) Then
				cmd.run DblQuote(temp(1))
				Call log("Folder Opened: " & eba)
			Else
				cmd.run eba
				Call log("Command Executed: " & eba)
			End If
		Elseif exeValue = "sys.shutdown" Then
			If exeValueExt = "eba.null" Or exeValueExt = "eba.default" Then
				eba = msgbox("Are you sure you want to " & defaultShutdown & " your PC? Make sure you save any unsaved data first!", 4+32, title)
				If eba = vbYes Then
					Call shutdown(defaultShutdown)
				End If
			Elseif exeValueExt = "eba.shutdown" Then
				eba = msgbox("Are you sure you want to shutdown your PC? All unsaved data will be lost!", 4+32, title)
				If eba = vbYes Then
					Call shutdown("shutdown")
				End If
			Elseif exeValueExt = "eba.restart" Then
				eba = msgbox("Are you sure you want to restart your PC? All unsaved data will be lost!", 4+32, title)
				If eba = vbYes Then
					Call shutdown("restart")
				End If
			Elseif exeValueExt = "eba.hibernate" Then
				eba = msgbox("Are you sure you want to hibernate your PC? We recommend saving unsaved data first!", 4+32, title)
				If eba = vbYes Then
					Call shutdown("hibernate")
				End If
			Else
				Error "Unknown Exe Value Extension.","UNKNOWN_ERROR"
			End If
		Elseif exeValue = "eba.uninstall" Then
			If isDev = false Then
				Error "This command can only be ran in EBA Developer Mode!","UNKNOWN_ERROR"
			Else
				eba = msgbox("Warning:" & line & "This will unistall EBA Command Center completely! Your EBA Command Center data will be erased! Uninstallation will require a system restart. Continue?", 4+48, title)
				Call addWarn
				If eba = vbYes Then
					fs.CopyFile scriptLoc, startup & "\uninstallEBA.vbs"
					Warn("EBA Command Center has been uninstalled. You will need to restart your PC to finish uninstallation")
					Call endOp("c")
				End If
				Note("Uninstallation canceled!")
			End If
		Elseif exeValue = "eba.upgrade" Then
			Note("EBA Keys have been phased out of EBA Command Center. There is now only one edition, which is 100% free! No need to upgrade.")
		Elseif exeValue = "eba.version" Then
			msgbox "EBA Command Center:" & line & "Version: " & ver & vblf & "Installed in: " & programLoc,64,title
		Elseif exeValue = "eba.write" Then
			If isDev = false Then
				Error "This command can only be ran in EBA Developer Mode!",""
			Else
				eba = inputbox("EBA > Call write", title)
				eba = Replace(eba,"""","")
				If fExists(eba) Then
					temp(0) = eba
					eba = inputbox("EBA > Call write > " & eba,title)
					If Lcase(eba) = "cancel" Then
						Note("Operation Canceled")
					Else
						Call log("Wrote data to " & temp(0) & ": " & eba)
						Call write(temp(0),eba)
					End If
				Else
					Call log("Failed to Call write to " & eba)
					Error "File " & eba & " not found!",""
				End If
			End If
		Else
			Error "The Execution Value is not valid." & vblf & exeValue,"INVALID_EXE_VALUE"
		End If
		
		If skipExe = false Then Call endOp("n")
	Loop
End Sub
Sub modeRecover
	title = setTitle("Recovery")
	Call checkWS
	
	Warn("EBA Command Center has launched into Recovery Mode.")
	
	temp(9) = enableLegacyEndOp
	enableLegacyEndOp = 1
	
	Do
		eba = LCase(inputbox("Enter Command Below:" & line & "Path: EBA > Recovery" & vblf & "Not Logged In",title))
		If eba = "repair" Then
			startupType = "repair"
			Exit Do
		Elseif eba = "startup" Then
			eba = LCase(inputbox("Enter a startupType:",title))
			Call write(dataLoc & "\startupType.ebacmd",eba)
		Elseif eba = "auto" Then
			Note "EBA Automatic Repair is now called EBA Repair. It is now accessed by typing 'repair' in the recovery menu."
		Elseif eba = "normal" Then
			startupType = "normal"
			Exit Do
		Elseif eba = "refresh" Then
			startupType = "refresh"
			Call write(dataLoc & "\startupType.ebacmd","refresh")
			Exit Sub
		Elseif eba = "" Then
			eba = msgbox("Exit EBA Cmd?",4+32,title)
			If eba = vbYes Then
				Call endOp("f")
			End If
		Else
			Error "Unrecognized command: " & eba,"INVALID_RECOVERY_COMMAND"
		End If
		Call endOp("n")
	Loop
	enableLegacyEndOp = temp(9)
End Sub
Sub modeRefresh
	title = setTitle("Installation")
	Call checkWS
	If isAdmin = False Then Call endOp("fa")
	
	eba = msgbox("You are about to refresh EBA Command Center. Refreshing will create a clean install of EBA Command Center. You can choose what data you would like to keep on the next screen. Continue?",4+48,title)
	If eba = vbNo Then
		Call write(dataLoc & "\startupType.ebacmd","normal")
		Call endOp("rd")
	End If
	
	temp(0) = False
	temp(1) = False
	temp(2) = False
	temp(3) = False
	
	eba = msgbox("Do you want to keep this data:" & line & "Commands",4+32,title)
	If eba = vbNo Then
		temp(0) = False
	Else
		temp(0) = True
	End If
	
	eba = msgbox("Do you want to keep this data:" & line & "Users",4+32,title)
	If eba = vbNo Then
		temp(1) = False
	Else
		temp(1) = True
	End If
	
	eba = msgbox("Do you want to keep this data:" & line & "Settings",4+32,title)
	If eba = vbNo Then
		temp(2) = False
	Else
		temp(2) = True
	End If
	
	eba = msgbox("Do you want to keep this data:" & line & "Plugins",4+32,title)
	If eba = vbNo Then
		temp(3) = False
	Else
		temp(3) = True
	End If
	
	eba = msgbox("Data you selected to keep:" & line & "EBA Cmd: True" & vblf & "EBA Registry: " & temp(2) & vblf & "Commands: " & temp(0) & vblf & "Users: " & temp(1) & vblf & "Settings: " & temp(2) & vblf & "Plugins: " & temp(3) & vblf & "Other: False" & line & "Are you sure you want to refresh EBA Command Center using the settings above? This cannot be undone!",4+48,title)
	If eba = vbNo Then
		Call write(dataLoc & "\startupType.ebacmd","normal")
		Call endOp("rd")
	End If
	
	Do
		temp(4) = inputbox("Where do you want to install EBA Command Center?",programLoc)
		temp(4) = Replace(temp(4),"""","")
		If Not foldExists(fs.GetParentFolderName(temp(4))) Then
			Error "The directory does not exist: " & fs.GetParentFolderName(programLoc),"DIRECTORY_NOT_FOUND"
		Else
			Exit Do
		End If
	Loop
	
	'Prepare to refresh
	fs.MoveFile scriptLoc, "C:\eba.temp"
	delete(programLoc)
	programLoc = temp(4)
	
	newFolder(programLoc)
	fs.MoveFile "C:\eba.temp", programLoc & "\EBA.vbs"
	
	'Customized
	If temp(0) = False Then
		delete(dataLoc & "\Commands")
	End If
	
	If temp(1) = False Then
		delete(dataLoc & "\Users")
	End If
	
	If temp(2) = False Then
		cmd.Regwrite regLoc, ""
		cmd.Regwrite regLoc & "\enableOperationCompletedMenu", 1, "REG_DWORD"
		cmd.Regwrite regLoc & "\enableLegacyOperationCompletedMenu", 0, "REG_DWORD"
		cmd.Regwrite regLoc & "\installDir", programLoc, "REG_SZ"
		
		delete(dataLoc & "\Settings")
	End If
	
	If temp(3) = False Then
		delete(dataLoc & "\Plugins")
	End If
	
	'Folders
	newFolder(programLoc & "\Commands")
	newFolder(dataLoc)
	newFolder(dataLoc & "\Users")
	newFolder(dataLoc & "\Commands")
	newFolder(dataLoc & "\Settings")
	newFolder(dataLoc & "\Plugins")
	newFolder(dataLoc & "\PluginData")
	Call createPlugdatFolder
	
	'Create Command Files
	Call updateCommands
	
	'Data Files
	Call update(dataLoc & "\isLoggedIn.ebacmd","" & vblf & "","overwrite")
	Call update(dataLoc & "\settings\logging.ebacmd","true","")
	Call update(dataLoc & "\settings\saveLogin.ebacmd","false","")
	Call update(dataLoc & "\settings\shutdownTimer.ebacmd","10","")
	Call update(dataLoc & "\settings\defaultShutdown.ebacmd","shutdown","")
	Call update(dataLoc & "\secureShutdown.ebacmd","true","overwrite")
	Call update(dataLoc & "\ebaKey.ebacmd",ebaKey,"")
	
	'Apply Setup
	If Not fExists(logDir) Then Call log("Log File Created")
	Call log("Critical Alert | EBA Command Center was refreshed.")
	
	'Create Icons
	Call createShortcut(desktop & "\EBA Command Center.lnk")
	Call createShortcut(startMenu & "\EBA Command Center.lnk")
	
	If temp(1) = False Then
		Call update(dataLoc & "\startupType.ebacmd","firstrun","overwrite")
		Note("EBA Command Center was refreshed. You'll need to run Initial Setup again (user accounts were erased!)")
		Call endOp("c")
	Else
		Call update(dataLoc & "\startupType.ebacmd","normal","overwrite")
		Note("EBA Command Center was refreshed.")
		Call endOp("c")
	End If
End Sub
Sub modeRepair
	title = setTitle("Recovery")
	Call checkWS
	
	temp(9) = enableLegacyEndOp
	enableLegacyEndOp = 1
	
	eba = msgbox("Are you sure you want to perform EBA Repair? This will reset your preferences.",4+48,title)
	
	If eba = vbNo Then
		Call endOp("r")
	Else
		If programLoc = scriptDir Then
			newFolder(dataLoc)
			newFolder(dataLoc & "\Users")
			newFolder(dataLoc & "\Commands")
			newFolder(dataLoc & "\Settings")
			newFolder(dataLoc & "\Plugins")
			newFolder(dataLoc & "\PluginData")
			If foldExists(dataLoc) Then
				Call updateCommands
				Call update(dataLoc & "\isLoggedIn.ebacmd","" & vblf & "","overwrite")
				Call update(dataLoc & "\settings\logging.ebacmd","true","overwrite")
				Call update(dataLoc & "\settings\saveLogin.ebacmd","false","overwrite")
				Call update(dataLoc & "\settings\shutdownTimer.ebacmd","10","overwrite")
				Call update(dataLoc & "\settings\defaultShutdown.ebacmd","shutdown","overwrite")
				Call update(dataLoc & "\secureShutdown.ebacmd","true","overwrite")
				Call update(dataLoc & "\startupType.ebacmd","firstrepair","overwrite")
				Note("EBA Repair has completed. EBA Command Center will now restart.")
				Call endOp("r")
			Else
				Error "EBA Repair failed for an unknown reason. Please try again later.","EBA_REPAIR_FAILED_TO_CREATE_OR_FIND_APPDATA_FOLDER"
				Call endOp("r")
			End If
		Else
			Error "EBA Repair failed because EBA Command Center is running from the installer.","RUNNING_FROM_INSTALLER"
			Call endOp("r")
		End If
	End If
	
	enableLegacyEndOp = temp(9)
End Sub
Sub modeUninstall
	title = setTitle("Recovery")
	Call checkWS
	
	If isAdmin = False Then
		Warn("To continue with uninstallation, EBA Command Center will run as administrator.")
		Call endOp("fa")
	End If
	
	eba = msgbox("EBA Command Center is Call ready to uninstall. Do you want to uninstall now? This cannot be undone, and your data will be lost!",4+48,title)
	If eba = vbNo Then
		Note("Your EBA Command Center data has been restored. EBA Command Center will now close.")
	Else
		delete(programLoc)
		delete(dataLoc)
		cmd.RegDelete("HKLM\SOFTWARE\EBA-Cmd")
		
		Note("EBA Command Center has been uninstalled.")
	End If
	delete(scriptLoc)
	
	enableLegacyEndOp = 1
	Call endOp("n")
	Call endOp("c")
End Sub
Sub modeUpdate
	title = setTitle("Installation")
	Call checkWS
	If isAdmin = False Then Call endOp("fa")
	
	eba = msgbox("Your installation of EBA Command Center will be updated. Continue?",4+32,title)
	If eba = vbNo Then Call endOp("c")
	
	'Registry
	cmd.Regwrite regLoc & "\enableOperationCompletedMenu", enableEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\enableLegacyOperationCompletedMenu", enableLegacyEndOp, "REG_DWORD"
	cmd.Regwrite regLoc & "\installDir", programLoc, "REG_SZ"
	
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
	Call downloadFiles
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
	fs.CopyFile scriptLoc, programLoc & "\EBA.vbs"
	
	'Create Icons
	Call createShortcut(desktop & "\EBA Command Center.lnk")
	Call createShortcut(startMenu & "\EBA Command Center.lnk")
	
	'Update Complete
	Note("EBA Command Center was updated to version " & ver)
	
	Call endOp("s")
End Sub








'Functions and Subroutines
Sub addDetails(str)
	'0: db, 1:note, 2: alert, 3: warn, 4: internal, 5: error, 6: ws, 7: critical
	If str = "crit" Then
		opDetails(7) = opDetails(7) + 1
	Elseif str = "ws" Then
		opDetails(6) = opDetails(6) + 1
	Elseif str = "err" Then
		opDetails(5) = opDetails(5) + 1
	Elseif str = "in" Then
		opDetails(4) = opDetails(4) + 1
	Elseif str = "wrn" Then
		opDetails(3) = opDetails(3) + 1
	Elseif str = "alr" Then
		opDetails(2) = opDetails(2) + 1
	Elseif str = "nt" Then
		opDetails(1) = opDetails(1) + 1
	Elseif str = "db" Then
		opDetails(0) = opDetails(0) + 1
	Else
		internal "EBA Command Center tried to add details to the Operation Complete menu, but the detail type (" & str & ") was invalid."
	End If
End Sub
Function alert(str)
	Call addDetails("alr")
	alert = msgbox("Alert:" & line & str,48,title)
End Function
Sub append(strA,strB)
	If fExists(strA) Then
		Set sys = fs.OpenTextFile (strA, 8)
		sys.writeLine strB
		sys.Close
	Elseif foldExists(fs.GetParentFolderName(strA)) Then
		Set sys = fs.CreateTextFile (strA, 8)
		sys.writeLine strB
		sys.Close
	Else
		db "System tried to write to the file at " & strA & ", but the system failed." & line & "Response Code: ParentDirectoryDoesNotExist"
	End If
End Sub
Sub checkImports
	If LCase(Right(importData, 10)) = ".ebaimport" Or LCase(Right(importData, 10)) = ".ebabackup" Or LCase(Right(importData, 10)) = ".ebaplugin" Then
		dbm "Imported file detected. Scanning file..."
		If LCase(Right(importData, 10)) = ".ebaimport" Then
			Call readLines(importData,1)
			If LCase(lines(1)) = "type: startup key" Then
				Call readLines(importData,2)
				If LCase(lines(2)) = "data: eba.recovery" Then
					eba = msgbox("Start EBA Command Center in recovery mode?",4+32,title)
					If eba = vbYes Then startupType = "recover"
				Else
					internal "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Startup Key" & vblf & "Data: " & lines(2),"StartupKey_Not_Known"
				End If
			Elseif lines(1) = "Type: Command" Then
				Call readLines(importData,5)
				
				If fExists(dataLoc & "\Commands\" & lines(2) & ".ebacmd") Or fExists(programLoc & "\Commands\" & lines(2) & ".ebacmd") Then
					internal "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Command" & vblf & "Error: Command with same name already exists: " & lines(2),"File_Already_Exists"
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
				internal "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Unknown","Unknown_Importable_Type"
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
				checkWS
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
				checkWS
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
						checkWS
						
					Elseif eba = "cmd" Then
						Set backup1 = App.NameSpace(dataLoc & "\Commands")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Commands")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						checkWS
						
					Elseif eba = "settings" Then
						Set backup1 = App.NameSpace(dataLoc & "\Settings")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Settings")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						checkWS
					Elseif eba = "plug" Then
						Set backup1 = App.NameSpace(dataLoc & "\Plugins")
						Set backup2 = App.NameSpace(importData)
						backup1.CopyHere(backup2.Items)
						If Err.Number = 0 Then
							Note("Restored files to " & dataLoc & "\Plugins")
						Else
							Error "Restore failed. See WScript Error for more info.","WS/" & Err.Number
						End If
						checkWS
					End If
				End If
			End If
			delete(dataLoc & "\tmp")
		Elseif LCase(Right(importData, 10)) = ".ebaplugin" Then
			eba = msgbox("Do you want to install this plugin? Make sure you trust the source of this plugin.", 4+32, title)
			If eba = vbYes Then
				checkWS
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
				checkWS
				delete(dataLoc & "\tmp")
			End If
		End If
	Elseif importData = "" Then
		dbm "No imported file detected. Skipping import."
		importData = False
	Else
		internal "There is a problem with the imported file. Details are shown below:" & line & "File: " & importData & vblf & "Type: Unknown" & vblf & "Error: FileEXT not recognized my EBA Cmd." & lines(2),"FileExt_Not_Known"
		dbm "Imported file detected. Unable to scan file."
	End If
End Sub
Function checkOS
	For Each forVar in os
		checkOS = forVar.Caption
	Next
End Function
Sub checkPCName
	delete dataLocTmp & "\blockedPC.ebacmd"
	dwnld "https://eba-tools.github.io/data/blockedPCList.xml",dataLocTmp & "\blockedPC.ebacmd"
	Call read(dataLocTmp & "\blockedPC.ebacmd","l")
	If InStr(data,"<name>" & LCase(pcName) & "</name>") Then
		Call log("[Critical] EBA Command Center could not start because the computer was on the EBA Blacklist. Error code PC_Blacklist_" & pcName)
		critical "Your PC is not allowed to run EBA Command Center because it is on the EBA Blacklist. Please contact EBA Tools to have your PC removed from the blacklist. Make sure to send a copy of the error code below:","PC_Blacklist_" & pcName
		Call endOp("s")
	End If
End Sub
Sub checkWS
	If Not Err.Number = 0 Then
		globTmp = ""
		globTmp = Err.Description & " (EBA Command Center does not know what caused this error)"
		errCode -2147483638, "Failed to download data from the EBA Website. Maybe check your internet connection?"
		errCode -2147024894, "Something went wrong accessing a file/registry key on your system."
		errCode -2147024891, "Failed to access system registry."
		errCode 70, "EBA Command Center cannot access a requested file because the file is in use, or EBA Command Center does not have permission."
		errCode 76, "The system could not locate the file that was specified by the system."
		errCode 500, "This is an internal error. If you are using an unmodified copy of EBA Command Center, please contact us. If this is a modified copy, please contact the modification developer. If you created this modification, view the dev details below."
		errCode 3004, "EBA Command Center failed to save a downloaded file."
		Call addDetails("ws")
		critical "A scripting error occurred during Operation " & count(0) & line & "Error Cause: " & globTmp & line & "Dev Details: " & Err.Description & line & "If you continue to receive this error, please contact us.", "WS/" & Err.Number
		dbm "Checked WScript. Found error: " & Err.Number
	Else
		dbm "Checked WScript. Found no errors."
	End If
	Err.Clear
End Sub
Sub clearCounts
	For forVar = 1 to 3
		count(forVar) = 0
	Next
	For forVar = 0 to 7
		opDetails(forVar) = 0
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
Sub createShortcut(str)
	Set Short = cmd.CreateShortcut(str)
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
Function critical(strA,strB)
	Call addDetails("crit")
	critical = msgbox("Critical Error:" & line & strA & line & "Error code: " & strB,16,title)
End Function
Sub dataExists(str)
	If Not fExists(str) Then
		missFiles = str
	End If
End Sub
Function db(str)
	Call addDetails("db")
	db = msgbox("Debug message:" & line & str,64,setTitle("Debug"))
End Function
Function DblQuote(str)
	DblQuote = """" & str & """"
End Function
Function dbm(str)
	If allowDebugMessages = True Then
		Call addDetails("db")
		dbm = msgbox("Debug message:" & line & str,64,setTitle("Debug"))
	End If
End Function
Function delete(str)
	If fExists(str) Then
		fs.DeleteFile(str)
	Elseif foldExists(str) Then
		fs.DeleteFolder(str)
	End If
End Function
Sub downloadFiles
	dwnld "https://eba-tools.github.io/data/icon.ico", programLoc & "\tmp.ebacmd"
	If fExists(programLoc & "\tmp.ebacmd") Then
		delete programLoc & "\icon.ico"
		fs.MoveFile programLoc & "\tmp.ebacmd", programLoc & "\icon.ico"
	Else
		error "The installer failed to download critical EBA Command Center files. Check your connection to the internet and try again."
		Call endOp("c")
	End If
End Sub
Function dwnld(strA,strB)
	download.open "get", strA, False
	download.send
	With stream
		.type = 1
		.open
		.write download.responseBody
		.savetofile strB
		.close
	End With
End Function
Sub endOp(str)
	Call addDetails("nt")
	delete dataLocTmp
	dbm "Deleted EBA.Tmp."
	
	'Crash
	If str = "c" Then
		Call log("EBA Command Center crashed.")
		quit
	End If
	
	Call checkWS
	
	'Force Shutdown
	If str = "f" Then
		Call log("EBA Command Center was forced to shut down")
		quit
	End If
	
	'Force Restart as Admin
	If str = "fa" Then
		app.ShellExecute "wscript.exe", DblQuote(scriptLoc), "", "runas", 1
		quit
	End If
	
	'Force Restart at Directory
	If str = "fd" Then
		cmd.run DblQuote(scriptLoc)
		quit
	End If
	If fExists(dataLoc & "\PluginData\Scripts\EndOp") Then
		For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\EndOp").Files
			cmd.run forVar
		Next
	End If
	
	'Operation Complete
	If enableEndOp = 1 Then
		Dim tmp1,tmp2
		If Not opDetails(7) = 0 Then tmp1 = tmp1 & "Critical Errors: " & opDetails(7) & vblf
		If Not opDetails(6) = 0 Then tmp1 = tmp1 & "WS Errors: " & opDetails(6) & vblf
		If Not opDetails(5) = 0 Then tmp1 = tmp1 & "Errors: " & opDetails(5) & vblf
		If Not opDetails(4) = 0 Then tmp1 = tmp1 & "Internal Exceptions: " & opDetails(4) & vblf
		If Not opDetails(3) = 0 Then tmp1 = tmp1 & "Warnings: " & opDetails(3) & vblf
		If Not opDetails(2) = 0 Then tmp1 = tmp1 & "Alerts: " & opDetails(2) & vblf
		If Not opDetails(1) = 0 Then tmp1 = tmp1 & "Notices: " & opDetails(1) & vblf
		If Not opDetails(0) = 0 Then tmp1 = tmp1 & "Debug Messages: " & opDetails(0) & vblf
		If Not opDetails(7) = 0 Then tmp2 = tmp2 & opDetails(7) & " Critical Errors,"
		If Not opDetails(6) = 0 Then tmp2 = tmp2 & opDetails(6) & " WS Errors,"
		If Not opDetails(5) = 0 Then tmp2 = tmp2 & opDetails(5) & " Errors,"
		If Not opDetails(4) = 0 Then tmp2 = tmp2 & opDetails(4) & " Internal Exceptions,"
		If Not opDetails(3) = 0 Then tmp2 = tmp2 & opDetails(3) & " Warnings,"
		If Not opDetails(2) = 0 Then tmp2 = tmp2 & opDetails(2) & " Alerts,"
		If Not opDetails(1) = 0 Then tmp2 = tmp2 & opDetails(1) & " Notices,"
		If Not opDetails(0) = 0 Then tmp2 = tmp2 & opDetails(0) & " Debug Messages,"
		tmp1 = Left(tmp1, Len(tmp1) - 1)
		tmp2 = Left(tmp2, Len(tmp2) - 1)
		If enableLegacyEndOp = 1 Then
			note "Operation " & count(0) & " Completed with " & tmp2
		Else
			note "Operation " & count(0) & " Completed:" & line & tmp1
		End If
	End If
	
	count(0) = count(0) + 1
	Call clearCounts
	Call clearLines
	Call clearTemps
	endOpFail = False
	
	'Shutdown
	If str = "s" Then
		If foldExists(dataLoc & "\PluginData\Scripts\Shutdown") Then
			For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\Shutdown").Files
				cmd.run forVar
			Next
		End If
		Call log("EBA Command Center was shut down.")
		quit
	End If
	
	'Restart
	If str = "r" Then
		If foldExists(dataLoc & "\PluginData\Scripts\Shutdown") Then
			For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\Shutdown").Files
				cmd.run forVar
			Next
		End If
		Call log("EBA Command Center restarted.")
		cmd.run DblQuote(programLoc & "\EBA.vbs")
		quit
	End If
	
	'Restart as Admin
	If str = "ra" Then
		For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\Shutdown").Files
			cmd.run forVar
		Next
		Call endOp("fa")
	End If
	
	'Restart At Directory
	If str = "rd" Then
		For Each forVar In fs.GetFolder(dataLoc & "\PluginData\Scripts\Shutdown").Files
			cmd.run forVar
		Next
		cmd.run DblQuote(scriptLoc)
		quit
	End If
	
	newFolder dataLocTmp
	dbm "Created EBA.Tmp."
End Sub
Function error(strA,strB)
	Call addDetails("err")
	error = msgbox("Error:" & line & strA & line & "Error code: " & strB,16,title)
End Function
Function errCode(int,str)
	If Err.Number = int Then globTmp = str
End Function
Function fExists(str)
	fExists = fs.FileExists(str)
End Function
Function foldExists(str)
	foldExists = fs.FolderExists(str)
End Function
Sub getTime
	nowDate = Right(0 & DatePart("m",Date),2) & "/" & Right(0 & DatePart("d",Date),2) & "/" & Right(0 & DatePart("yyyy",Date),2)
	nowTime = Right(0 & Hour(Now),2) & ":" & Right(0 & Minute(Now),2) & ":" & Right(0 & Second(Now),2)
End Sub
Function goOnline(str)
	https.open "get", str, False
	https.send
	If Err.Number = 0 Then
		goOnline = https.responseText
	Else
		warn "We had trouble connecting to the internet."
	End If
End Function
Function internal(strA,strB)
	Call addDetails("in")
	internal = msgbox("Internal Exception:" & line & strA & line & "Error code: " & strB,16,title)
End Function
Sub loadPlugins(str)
	If pluginCount > (maxPlugin - 1) Then
		warn "Failed to load plugin: " & str & line & "The maximum number of plugins (" & maxPlugin & ") has been reached."
	Else
		loadedPlugins(pluginCount) = str
		pluginCount = pluginCount + 1
	End If
End Sub
Sub log(str)
	If logging = "true" Then
		Dim tmp
		Call getTime
		tmp = "[" & nowTime & " - " & nowDate & "] " & str
		Call append(logDir, tmp)
	End If
End Sub
Function newFolder(str)
	If Not foldExists(str) Then
		If foldExists(fs.GetParentFolderName(str)) Then
			newFolder = fs.CreateFolder(str)
		End If
	End If
End Function
Function note(str)
	Call addDetails("nt")
	note = msgbox("Notice:" & line & str,64,title)
End Function
Sub preparePlugins
	dbm "Preparing plugins."
	If foldExists(dataLoc & "\PluginData") Then
		For Each forVar In fs.GetFolder(dataLoc & "\PluginData").SubFolders
			dbm "Deleting old plugin data: " & forVar
			delete(forVar)
			dbm "Deleted old plugin data."
		Next
	Else
		dbm "Old plugin data does not exist. Skipping step."
	End If
	Call createPlugdatFolder
	If foldExists(dataLoc & "\Plugins") Then
		For Each forVar In fs.GetFolder(dataLoc & "\Plugins").Subfolders
			If fExists(forVar & "\meta.xml") Then
				XML.load(forVar & "\meta.xml")
				checkWS
				For Each forVar1 In XML.selectNodes("/Meta/Format")
					checkWS
					If forVar1.text = "1" Then
						For Each forVar2 In XML.selectNodes("/Meta/License/ID")
							checkWS
							For Each forVar3 In XML.selectNodes("/Meta/Version/Name")
								checkWS
								For Each forVar4 In XML.selectNodes("/Meta/Version/Version")
									checkWS
									temp(2) = forVar3.text
									temp(0) = goOnline("https://ethanblaisalarms.github.io/cmd/plugin/" & forVar2.text & ".txt")
									temp(0) = Left(temp(0), Len(temp(0)) - 1)
									temp(1) = goOnline("https://ethanblaisalarms.github.io/cmd/plugin/ver/" & forVar2.text & ".txt")
									temp(1) = Left(temp(1), Len(temp(1)) - 1)
									If temp(0) = temp(2) Then
										If CDbl(forVar4.text) <= CDbl(temp(1)) Then
											Call loadPlugins(forVar)
											dbm "Plugin prepared: " & forVar
										Else
											Call addWarn
											eba = msgbox("Warning:" & line & "The plugin at " & forVar & " is an experimental version. Load anyways?",4+48,title)
											If eba = vbYes Then
												Call loadPlugins(forVar)
												dbm "Plugin prepared: " & forVar
											Else
												dbm "Plugin skipped: " & forVar
											End If
										End If
									Else
										Call addWarn
										eba = msgbox("Warning:" & line & "The plugin at " & forVar & " is not licensed. This means EBA has not validated this plugin. Loading it could be risky. Load anyways?",4+48,title)
										If eba = vbYes Then
											Call loadPlugins(forVar)
											dbm "Plugin prepared: " & forVar
										Else
											dbm "Plugin skipped: " & forVar
										End If
									End If
								Next
							Next
						Next
					Else
						internal "The plugin at " & forVar & " contains an invalid META.XML file, and will be skipped.","Unknown_Format_Version"
						dbm "Plugin skipped: " & forVar
					End If
				Next
			Else
				internal "The plugin at " & forVar & " is missing META.XML, and will be skipped.","Meta.xml_Missing"
				dbm "Plugin skipped: " & forVar
			End If
		Next
	Else
		dbm "No plugins found on the system. Skipping step."
	End If
	dbm "Plugins prepared."
End Sub
Function quit
	quit = wscript.quit
End Function
Sub read(strA,strB)
	If fExists(strA) Then
		Dim tmp
		Set sys = fs.OpenTextFile (strA,1)
		tmp = sys.readAll
		tmp = Left(tmp, Len(tmp)	- 2)
		sys.Close
		If strB = "l" Then tmp = LCase(tmp)
		If strB = "u" Then tmp = UCase(tmp)
		data = tmp
	Else
		db "System tried to read file at " & strA & ", but the system failed." & line & "Response Code: FileNotFound"
	End If
End Sub
Sub readCommands(str)
	Call readLines(str & "\Commands\" & eba & ".ebacmd",3)
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

Sub readLine(str,int)
	If fExists(str) Then
		Set sys = fs.OpenTextFile (str, 1)
		For forVar = 1 to int
			data = sys.readLine
		Next
		sys.Close
	Else
		db "System tried to read file at " & strA & ", but the system failed." & line & "Response Code: FileNotFound"
	End If
End Sub
Sub readLines(str,int)
	If fExists(str) Then
		Set sys = fs.OpenTextFile (str, 1)
		For forVar = 1 to int
			lines(forVar) = sys.readLine
		Next
		sys.Close
	Else
		db "System tried to read file at " & strA & ", but the system failed." & line & "Response Code: FileNotFound"
	End If
End Sub
Sub readSettings
	On Error Resume Next
	Call checkWS
	
	programLoc = "C:\Program Files\EBA"
	
	'Registry read
	programLoc = cmd.Regread(regLoc & "\installDir")
	enableEndOp = cmd.Regread(regLoc & "\enableOperationCompletedMenu")
	enableLegacyEndOp = cmd.Regread(regLoc & "\enableLegacyOperationCompletedMenu")
	
	'Conversion
	enableEndOp = CInt(enableEndOp)
	enableLegacyEndOp = CInt(enableLegacyEndOp)
	
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
	
	dbm "Settings were read. If any settings are missing, they will have been skipped."
	
	Err.Clear
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
Function setTitle(str)
	setTitle = "EBA Command Center " & ver & " | " & str
End Function
Sub shutdown(str)
	If str = "shutdown" Then
		cmd.run "shutdown /s /t " & shutdownTimer & " /f /c ""You requested a system shutdown in EBA Command Center."""
		Warn("Your PC will shut down in " & shutdownTimer & " seconds. Press OK to cancel.")
	Elseif str = "restart" Then
		cmd.run "shutdown /r /t " & shutdownTimer & " /f /c ""You requested a system restart in EBA Command Center."""
		Warn("Your PC will restart in " & shutdownTimer & " seconds. Press OK to cancel.")
	Elseif str = "hibernate" Then
		cmd.run "shutdown /h"
	Else
		cmd.run "shutdown /s /t 15 /f /c ""There was an issue with the shutdown method, so EBA Cmd will shutdown your PC in 15 seconds."""
		Warn("Your PC will shutdown in 15 seconds (due to an error with the shutdownMethod). Press OK to cancel.")
	End If
	cmd.run "shutdown /a"
End Sub
Sub update(strA,strB,strC)
	If LCase(strC) = "overwrite" Then
		Call write(strA,strB)
	Elseif LCase(strC) = "append" Then
		Call append(strA,strB)
	Else
		If Not fExists(strA) Then
			Call write(strA,strB)
		End If
	End If
End Sub
Sub updateCommands
	
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
Function warn(str)
	Call addDetails("wrn")
	warn = msgbox("Warning:" & line & str,48,title)
End Function
Sub write(strA,strB)
	If fExists(strA) Then
		Set sys = fs.OpenTextFile (strA, 2)
		sys.writeLine strB
		sys.Close
	Elseif foldExists(fs.GetParentFolderName(strA)) Then
		Set sys = fs.CreateTextFile (strA, 2)
		sys.writeLine strB
		sys.Close
	Else
		db "System tried to write to the file at " & strA & ", but the system failed." & line & "Response Code: ParentDirectoryDoesNotExist"
	End If
End Sub

Call endOp("n")