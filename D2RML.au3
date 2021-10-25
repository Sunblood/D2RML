#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.15.3 (Beta)
 Author:         Sunblood

 Script Function:
	Automates closing D2R multi-process watchdog handle and webtoken for logging into multiple clients

#ce ----------------------------------------------------------------------------
#include <AutoitConstants.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

FileInstall("handle64.exe", "handle64.exe", 0)
If Not FileExists("handle64.exe") Then
	MsgBox(262144 + 16, "Error", "Error: Unable to locate handle64.exe. Download it from Github and place it in the same folder as D2RML.")
	Exit
EndIf
If @Compiled Then
	SplashTextOn("D2RML", "Preloading handle64. Please wait.", 300, 100, Default, Default, 32)
	ShellExecuteWait("handle64.exe", "", "", "", @SW_HIDE)
	SplashOff()
EndIf

#Region ### START Koda GUI section ### Form=C:\Jack\Programs\D2\Multilaunch\guiMain.kxf
$guiMain = GUICreate("D2RML", 365, 306, -1, -1)
GUISetBkColor(0xC0DCC0)
$buttonAdd = GUICtrlCreateButton("Add Token", 8, 40, 75, 25)
$listViewMain = GUICtrlCreateListView("Account|Token Date|Region", 8, 72, 250, 150, BitOR($LVS_REPORT,$LVS_SINGLESEL,$WS_VSCROLL), BitOR($WS_EX_CLIENTEDGE,$LVS_EX_CHECKBOXES,$LVS_EX_FULLROWSELECT))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 50)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 50)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 50)
$buttonLaunch = GUICtrlCreateButton("Launch Selected", 264, 72, 91, 25)
$buttonRefresh = GUICtrlCreateButton("Refresh Token", 264, 104, 91, 25)
$buttonRemove = GUICtrlCreateButton("Remove Token", 264, 136, 91, 25)
$labelHelp = GUICtrlCreateLabel("How does this work?", 256, 18, 103, 17)
GUICtrlSetFont(-1, 8, 400, 4, "MS Sans Serif")
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor (-1, 0)
GUICtrlCreateLabel("Sun's D2R Multilauncher", 8, 8, 236, 31)
GUICtrlSetFont(-1, 15, 400, 0, "@Microsoft YaHei UI")
$checkboxArgs = GUICtrlCreateCheckbox("Game cmdline:", 8, 232, 89, 17)
$inputArgs = GUICtrlCreateInput("", 97, 230, 160, 21)
$checkboxSkipIntro = GUICtrlCreateCheckbox("Skip intro videos", 8, 256, 97, 17)
$checkboxChangeTitle = GUICtrlCreateCheckbox("Change game title to match token name", 8, 280, 209, 17)
;~ GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global Const $accountRegKey[] = ["HKEY_CURRENT_USER\SOFTWARE\Blizzard Entertainment\Battle.net\Launch Options\OSI", "WEB_TOKEN"]
Global Const $gameInstallRegKey[] = ["HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Diablo II Resurrected", "InstallLocation"]
Global Const $bnetInstallRegKey[] = ["HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Battle.net", "InstallLocation"]
Global Const $gameClass = "[CLASS:OsWindow]"
Global Const $bnetLauncherClass = "[CLASS:Qt5QWindowIcon]"
Global Const $bnetClientClass = "[CLASS:Chrome_WidgetWin_0]"
Global Const $settingsFile = "D2RML.ini"
Global Const $version = "0.0.4"

Global $tokenInProgress = 0

OnAutoItExitRegister("SaveSettings")

WinSetTitle($guiMain, "", "D2RML v" & $version)
LoadSettings()
LoadAccounts()
GUISetState()

CheckVersion()

While 1
	GuiMessages()

WEnd

Func DisableButtons()
	GUICtrlSetState($buttonAdd,$GUI_DISABLE)
	GUICtrlSetState($buttonLaunch,$GUI_DISABLE)
	GUICtrlSetState($buttonRefresh,$GUI_DISABLE)
	GUICtrlSetState($buttonRemove,$GUI_DISABLE)
EndFunc
Func EnableButtons()
	GUICtrlSetState($buttonAdd,$GUI_ENABLE)
	GUICtrlSetState($buttonLaunch,$GUI_ENABLE)
	GUICtrlSetState($buttonRefresh,$GUI_ENABLE)
	GUICtrlSetState($buttonRemove,$GUI_ENABLE)
EndFunc

Func GuiMessages()
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $buttonAdd
			Setup()
		Case $buttonLaunch
			DisableButtons()
			$count = _GUICtrlListView_GetItemCount($listViewMain)
			For $i = 0 To $count
				If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
					$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
					LaunchWithAccount($name)
				EndIf
			Next
			LoadAccounts()
			EnableButtons()
		Case $buttonRefresh
			DisableButtons()
			$count = _GUICtrlListView_GetItemCount($listViewMain)
			For $i = 0 To $count
				If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
					$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
					Setup($name)
				EndIf
			Next
			EnableButtons()
		Case $buttonRemove
			DisableButtons()
			If MsgBox(262144 + 36, "Remove account", "Are you sure you want to remove the checked account(s)?") = 6 Then
				$count = _GUICtrlListView_GetItemCount($listViewMain)
				For $i = 0 To $count
					If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
						$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
						FileDelete($name & ".bin")
					EndIf
				Next
				LoadAccounts()
			EndIf
			EnableButtons()
		Case $labelHelp
			MsgBox(262144 + 64, $version, "Sun's D2R Multilauncher (D2RML) saves D2R login tokens from the registry into .bin files that are then restored on-demand." & @CRLF & _
					"It also allows multiple copies of D2R to be launched simultaneously. No additional software is required." & @CRLF & @CRLF & _
					"Initial setup requires that you log into an account at least once in order to save the token. " & _
					"Login tokens are only valid once, after which they expire and a new token is generated. D2RML does this automatically as long as you *always* use this app to launch the game." & @CRLF & _
					"Logging in through normal means will invalidate the saved token for the account and you will be unable to connect. Use the 'Refresh Token' button to redo the setup and generate a new token.")
	EndSwitch
EndFunc   ;==>GuiMessages

Func CheckVersion()
	$source = BinaryToString(InetRead("https://raw.githubusercontent.com/Sunblood/D2RML/main/D2RML.au3", 1))
	$a = StringSplit($source,@CRLF)
	For $i = 1 to $a[0]
		If StringLeft($a[$i],21) = "Global Const $version" Then
			$b = StringSplit($a[$i],'"')
			$v = $b[2]
			If $v > $version Then
				If MsgBox(36,"Update","A new version of D2RML is available: "&$v&@CRLF&"Open download page?") = 6 Then
					ShellExecute("https://github.com/Sunblood/D2RML")
				EndIf
			EndIf
		EndIf
	Next
EndFunc

Func Setup($name = "")
	If ProcessExists("D2R.exe") Or ProcessExists("Battle.net.exe") Or ProcessExists("Diablo II Resurrected Launcher.exe") Then
		If MsgBox(262144 + 5 + 16, "Error", "D2R and Battle.net launcher(s) must be closed.") = 4 Then ;retry
			Setup($name)
			Return
		Else
			Return
		EndIf
	EndIf

	If $name = "" Then
		$name = InputBox("Setup", "Enter a name for the token:")
		If $name = "" Then Return
		$region = InputBox("Setup", "Enter the server region for the token" & @CRLF & @CRLF & "Acceptable values are NA / EU / KR", "NA")
		If $region = "" Then Return
	EndIf

	WinClose($bnetClientClass)
	WinClose($bnetLauncherClass)

	ToolTip("Creating Tokens: " & $name & @CRLF & "Log into the Launcher with the desired account and press PLAY", 0, 0)
	LaunchLauncher()

	ProcessWait("D2R.exe")
	ToolTip("Creating Tokens: " & $name & @CRLF & "Generating token 1 of 2...", 0, 0)
	WaitForNewKey()
	ToolTip("Creating Tokens: " & $name & @CRLF & "Generating token 2 of 2...", 0, 0)
	WaitForNewKey()
	;Key changes twice - second key is what we want

	ExportRegKey($name & ".bin")
;~ 	WinClose($gameClass)
	WinClose($bnetClientClass)
	WinClose($bnetLauncherClass)

	ToolTip("")
	LoadAccounts()
	If GUICtrlRead($checkboxChangeTitle) = $GUI_CHECKED Then
		WinSetTitle(GetGameWindowHandle("D2R.exe"), "", "D2R:" & $name)
	EndIf
	MsgBox(262144 + 64, "Finished", "Successfully saved token: " & $name)
EndFunc   ;==>Setup

Func LaunchWithAccount($name)
	ToolTip("Launching token: " & $name, 0, 0)
	WriteRegKey($name & ".bin")
	$curKey = RegRead($accountRegKey[0], $accountRegKey[1])
	$gamePID = LaunchGame()

	ToolTip("Launching token: " & $name & @CRLF & "Refreshing token 1 of 2...", 0, 0)
	If WaitForNewKey() = 0 Then
		MsgBox(262144 + 16, "Error", "Error obtaining new tokens. Game process closed.")
		ToolTip("")
		Return
	EndIf
	ToolTip("Launching token: " & $name & @CRLF & "Refreshing token 2 of 2...", 0, 0)
	If WaitForNewKey() = 0 Then
		MsgBox(262144 + 16, "Error", "Error obtaining new tokens. Game process closed.")
		ToolTip("")
		Return
	EndIf

	ToolTip("")

	If ProcessExists($gamePID) Then
		If GUICtrlRead($checkboxChangeTitle) = $GUI_CHECKED Then
			WinSetTitle(GetGameWindowHandle($gamePID), "", "D2R:" & $name)
		EndIf
		ExportRegKey($name & ".bin")
		CloseMultiProcessHandle($gamePID)
	Else
		MsgBox(262144 + 16, "Error", "Error obtaining new tokens. You may need to refresh the token you're trying to use.")
	EndIf
EndFunc   ;==>LaunchWithAccount

Func CloseMultiProcessHandle($pid = "D2R.exe")
	WriteLog("Closing multi-process handle for " & $pid)
	Local $result, $handle

	$getHandle = ComspecGetOutput("handle64.exe -a -p " & $pid & " Instances")
	WriteLog($getHandle)
	$a = StringSplit($getHandle, @CRLF)
	If IsArray($a) Then
		For $i = 1 To $a[0]
			If StringInStr($a[$i], "D2R.exe") Then
				$result = $a[$i]
				WriteLog("Result: " & $i & " - " & $a[$i])
			EndIf
		Next
	Else
		Return 0
		WriteLog("ERROR: Output is not array")
	EndIf
	If $result <> "" Then
		$b = StringSplit($result, ": ")
		If IsArray($b) Then
			$handle = $b[$b[0] - 6]
			WriteLog("Handle: " & $handle)
		EndIf
	EndIf

	If $handle > 0 Then
		WriteLog("Closing process: " & RunWait("handle64.exe -c " & $handle & " -p " & $pid & " -y", @ScriptDir, @SW_HIDE))
	Else
		WriteLog("ERROR: Unknown handle " & $handle & " - PID: " & $pid)
		MsgBox(262144 + 16, "Error", "Error closing multi-instance handle for PID " & $pid & @CRLF & "Handle: " & $handle)
	EndIf
EndFunc   ;==>CloseMultiProcessHandle

Func LoadAccounts()
	_GUICtrlListView_DeleteAllItems($listViewMain)
	$list = _FileSearch("*.bin", 0)
	For $i = 1 To $list[0]
		$date = FileGetTime($list[$i], 0, 0)
		$dateString = $date[0] & "/" & $date[1] & "/" & $date[2] & " " & $date[3] & ":" & $date[4] & ":" & $date[5]
		GUICtrlCreateListViewItem(StringTrimRight(GetFilename($list[$i]), 4) & "|" & $dateString, $listViewMain)
	Next
	_GUICtrlListView_SetColumnWidth($listViewMain, 0, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($listViewMain, 1, $LVSCW_AUTOSIZE)
EndFunc   ;==>LoadAccounts

Func ExportRegKey($keyfile)
	$f = FileOpen($keyfile, 2 + 16)
	FileWrite($f, RegRead($accountRegKey[0], $accountRegKey[1]))
	FileClose($f)
EndFunc   ;==>ExportRegKey
Func WriteRegKey($keyfile)
	$f = FileOpen($keyfile, 16)
	RegWrite($accountRegKey[0], $accountRegKey[1], "REG_BINARY", FileRead($f))
	FileClose($f)
EndFunc   ;==>WriteRegKey
Func WaitForNewKey()
	$curKey = RegRead($accountRegKey[0], $accountRegKey[1])
	$sendTimer = TimerInit()
	$queueTimer = TimerInit()
	Do
		$newKey = RegRead($accountRegKey[0], $accountRegKey[1])
		If TimerDiff($sendTimer) > 500 And TimerDiff($queueTimer) < 15000 And GUICtrlRead($checkboxSkipIntro) = $GUI_CHECKED Then
			ControlSend($gameClass, "", "", "{SPACE}")
			$sendTimer = TimerInit()
		EndIf
		If Not ProcessExists("D2R.exe") Then Return 0
		GuiMessages()
	Until $newKey <> $curKey
	Return 1
EndFunc   ;==>WaitForNewKey
Func LaunchGame()
	$path = RegRead($gameInstallRegKey[0], $gameInstallRegKey[1])
	If GUICtrlRead($checkboxArgs) = $GUI_CHECKED Then
		Return ShellExecute($path & "\D2R.exe", GUICtrlRead($inputArgs))
	Else
		Return ShellExecute($path & "\D2R.exe")
	EndIf
EndFunc   ;==>LaunchGame
Func LaunchLauncher() ;hehe
	Local $bnpath = RegRead($bnetInstallRegKey[0], $bnetInstallRegKey[1])
	Local $gamepath = RegRead($gameInstallRegKey[0], $gameInstallRegKey[1])
	Return ShellExecute($bnpath & "\Battle.net.exe", '--game=osi "--gamepath=' & $gamepath & '"')
EndFunc   ;==>LaunchLauncher
Func GetGameWindowHandle($pid)
	$p = _ProcessGetWindow($pid)
	If IsArray($p) Then
		If $p[0] = 3 Then
			Return $p[1]
		EndIf
	EndIf
	Return 0
EndFunc   ;==>GetGameWindowHandle

Func SaveSettings()
	IniWrite($settingsFile, "Main", "argsEnabled", GUICtrlRead($checkboxArgs))
	IniWrite($settingsFile, "Main", "args", GUICtrlRead($inputArgs))
	IniWrite($settingsFile, "Main", "skipIntro", GUICtrlRead($checkboxSkipIntro))
	IniWrite($settingsFile, "Main", "changeTitle", GUICtrlRead($checkboxChangeTitle))
EndFunc   ;==>SaveSettings
Func LoadSettings()
	GUICtrlSetState($checkboxArgs, IniRead($settingsFile, "Main", "argsEnabled", 4))
	GUICtrlSetData($inputArgs, IniRead($settingsFile, "Main", "args", ""))
	GUICtrlSetState($checkboxSkipIntro, IniRead($settingsFile, "Main", "skipIntro", 4))
	GUICtrlSetState($checkboxChangeTitle, IniRead($settingsFile, "Main", "changeTitle", 4))
EndFunc   ;==>LoadSettings

Func GetFilename($file)
	Local $a = StringSplit($file, "\")
	Return $a[$a[0]]
EndFunc   ;==>GetFilename

Func WriteLog($text)
	If Not @Compiled Then
		FileWriteLine("d2rml.log", @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " - " & $text)
	EndIf
EndFunc   ;==>WriteLog

Func ComspecGetOutput($command, $cmd = 1)
	If $cmd Then
		$iPID = Run(@ComSpec & " /c " & $command, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	Else
		$iPID = Run($command, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	EndIf
	ProcessWaitClose($iPID)
	Return StdoutRead($iPID)
EndFunc   ;==>ComspecGetOutput

Func _FileSearch($szMask, $nOption)
	Local $szRoot = ""
	Local $hFile = 0
	Local $szBuffer = ""
	Local $szReturn = ""
	Local $szPathList = "*"
	Dim $aNULL[1]
	If Not StringInStr($szMask, "\") Then
		$szRoot = @ScriptDir & "\"
	Else
		While StringInStr($szMask, "\")
			$szRoot = $szRoot & StringLeft($szMask, StringInStr($szMask, "\"))
			$szMask = StringTrimLeft($szMask, StringInStr($szMask, "\"))
		WEnd
	EndIf
	If $nOption = 0 Then
		_FileSearchUtil($szRoot, $szMask, $szReturn)
	Else
		While 1
			$hFile = FileFindFirstFile($szRoot & "*.*")
			If $hFile >= 0 Then
				$szBuffer = FileFindNextFile($hFile)
				While Not @error
					If $szBuffer <> "." And $szBuffer <> ".." And _
							StringInStr(FileGetAttrib($szRoot & $szBuffer), "D") Then _
							$szPathList = $szPathList & $szRoot & $szBuffer & "*"
					$szBuffer = FileFindNextFile($hFile)
				WEnd
				FileClose($hFile)
			EndIf
			_FileSearchUtil($szRoot, $szMask, $szReturn)
			If $szPathList == "*" Then ExitLoop
			$szPathList = StringTrimLeft($szPathList, 1)
			$szRoot = StringLeft($szPathList, StringInStr($szPathList, "*") - 1) & "\"
			$szPathList = StringTrimLeft($szPathList, StringInStr($szPathList, "*") - 1)
		WEnd
	EndIf
	If $szReturn = "" Then
		$aNULL[0] = 0
		Return $aNULL
	Else
		Return StringSplit(StringTrimRight($szReturn, 1), "*")
	EndIf
EndFunc   ;==>_FileSearch
Func _FileSearchUtil(ByRef $ROOT, ByRef $MASK, ByRef $RETURN)
	$hFile = FileFindFirstFile($ROOT & $MASK)
	If $hFile >= 0 Then
		$szBuffer = FileFindNextFile($hFile)
		While Not @error
			If $szBuffer <> "." And $szBuffer <> ".." Then _
					$RETURN = $RETURN & $ROOT & $szBuffer & "*"
			$szBuffer = FileFindNextFile($hFile)
		WEnd
		FileClose($hFile)
	EndIf
EndFunc   ;==>_FileSearchUtil


; ===============================================================================================================================
; <_RunWithReducedPrivileges.au3>
;
; Function to run a program with reduced privileges.
;	Useful when running in a higher privilege mode, but need to start a program with reduced privileges.
;	- A common problem this fixes is drag-and-drop not working, and misc functions (sendmessage, etc) not working.
;
; Functions:
;	_RunWithReducedPrivileges()		; runs a process with reduced privileges if currently running in a higher privilege mode
;
; INTERNAL Functions:
;	_RWRPCleanup()		; Helper function for the above
;
; Reference:
;	See 'Creating a process with Medium Integration Level from the process with High Integration Level in Vista'
;		@ http://www.codeproject.com/KB/vista-security/createprocessexplorerleve.aspx
;	  See Elmue's comment 'Here the cleaned and bugfixed code'
;	Also see: 'High elevation can be bad for your application: How to start a non-elevated process at the end of the installation'
;		@ http://www.codeproject.com/KB/vista-security/RunNonElevated.aspx
;	  (Elmue has the same code here too in his response to FaxedHead's comment ('Another alternative to this method'))
;	Another alternative using COM methods:
;	  'Getting the shell to run an application for you - Part 2:How | BrandonLive'
;		@ http://brandonlive.com/2008/04/27/getting-the-shell-to-run-an-application-for-you-part-2-how/
;
; Author: Ascend4nt, based on code by Elmue's fixed version of Alexey Gavrilov's code
; ===============================================================================================================================


; ===================================================================================================================
; Func _RWRPCleanup($hProcess,$hToken,$hDupToken,$iErr=0,$iExt=0)
;
; INTERNAL: Helper function for _RunWithReducedPrivileges()
;
; Author: Ascend4nt
; ===================================================================================================================

Func _RWRPCleanup($hProcess, $hToken, $hDupToken, $iErr = 0, $iExt = 0)
	Local $aHandles[3] = [$hToken, $hDupToken, $hProcess] ; order is important
	For $i = 0 To 2
		If $aHandles[$i] <> 0 Then DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aHandles[$i])
	Next
	Return SetExtended($iExt, $iErr)
EndFunc   ;==>_RWRPCleanup


; ===================================================================================================================
; Func _RunWithReducedPrivileges($sPath,$sCmd='',$sFolder='',$iShowFlag=@SW_SHOWNORMAL,$bWait=False)
;
; Function to run a program with reduced privileges.
;	Useful when running in a higher privilege mode, but need to start a program with reduced privileges.
;	- A common problem this fixes is drag-and-drop not working, and misc functions (sendmessage, etc) not working.
;
; $sPath = Path to executable
; $sCmd = Command-line (optional)
; $sFolder = Folder to start in (optional)
; $iShowFlag = how the program should appear on startup. Default is @SW_SHOWNORMAL.
;	All the regular @SW_SHOW* macros should work here
; $bWait = If True, waits for the process to finish before returning with an exit code
;	If False, it returns without waiting for the process to finish, with the process ID #
;
; Returns:
;	Success: If $bWait=True, the exit code of the Process. If $bWait=False, then the Process ID # of the process
;	Failure: 0, with @error set:
;		@error = 2 = DLLCall error. @extended contains the DLLCall error code (see AutoIt Help)
;		@error = 3 = API returned failure. Call 'GetLastError' API function to get more info.
;
; Author: Ascend4nt, based on code by Elmue's fixed version of Alexey Gavrilov's code
; ===================================================================================================================

Func _RunWithReducedPrivileges($sPath, $sCmd = '', $sFolder = '', $iShowFlag = @SW_SHOWNORMAL, $bWait = False)
	Local $aRet, $iErr, $iRet = 1, $hProcess, $hToken, $hDupToken, $stStartupInfo, $stProcInfo
	Local $sCmdType = "wstr", $sFolderType = "wstr"

;~ 	Run normally if not in an elevated state, or if pre-Vista O/S
	If Not IsAdmin() Or StringRegExp(@OSVersion, "_(XP|200(0|3))") Then   ; XP, XPe, 2000, or 2003?
		If $bWait Then Return RunWait($sPath & ' ' & $sCmd, $sFolder)
		Return Run($sPath & ' ' & $sCmd, $sFolder)
	EndIf

;~ 	Check Parameters and adjust DLLCall types accordingly
	If Not IsString($sCmd) Or $sCmd = '' Then
		$sCmdType = "ptr"
		$sCmd = 0
	EndIf
	If Not IsString($sFolder) Or $sFolder = '' Then
		$sFolderType = "ptr"
		$sFolder = 0
	EndIf
	#cs
		; STARTUPINFOW struct: cb,lpReserved,lpDesktop,lpTitle,dwX,dwY,dwXSize,dwYSize,dwXCountChars,dwYCountChars,dwFillAttribute,
		;	dwFlags,wShowWindow,cbReserved2,lpReserved2,hStdInput,hStdOutput,hStdError
		;	NOTE: This is for process creation info. Also, not sure if the Std I/O can be redirected..?
	#ce
	$stStartupInfo = DllStructCreate("dword;ptr[3];dword[7];dword;word;word;ptr;handle[3]")
	DllStructSetData($stStartupInfo, 1, DllStructGetSize($stStartupInfo))
	DllStructSetData($stStartupInfo, 4, 1)  ; STARTF_USESHOWWINDOW
	DllStructSetData($stStartupInfo, 5, $iShowFlag)

	; PROCESS_INFORMATION struct: hProcess, hThread, dwProcessId, dwThreadId
	;	This is for *receiving* info
	$stProcInfo = DllStructCreate("handle;handle;dword;dword")

;~ 	Open a handle to the Process
	; Explorer runs under a lower privilege, so it is the basis for our security info.
	;	Open the process with PROCESS_QUERY_INFORMATION (0x0400) access
	$aRet = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", 0x0400, "bool", False, "dword", ProcessExists("explorer.exe"))
	If @error Then Return SetError(2, @error, 0)
	If Not $aRet[0] Then Return SetError(3, 0, 0)
	$hProcess = $aRet[0]

;~ 	Open a handle to the Process's token (for duplication)
	; TOKEN_DUPLICATE = 0x0002
	$aRet = DllCall("advapi32.dll", "bool", "OpenProcessToken", "handle", $hProcess, "dword", 2, "handle*", 0)
	If @error Then Return SetError(_RWRPCleanup($hProcess, 0, 0, 2, @error), @extended, 0)
	If $aRet[0] = 0 Then Return SetError(_RWRPCleanup($hProcess, 0, 0, 3), @extended, 0)
	$hToken = $aRet[3]

;~ 	Duplicate the token handle
	; TOKEN_ALL_ACCESS = 0xF01FF, SecurityImpersonation = 2, TokenPrimary = 1,
	$aRet = DllCall("advapi32.dll", "bool", "DuplicateTokenEx", "handle", $hToken, "dword", 0xF01FF, "ptr", 0, "int", 2, "int", 1, "handle*", 0)
	If @error Then Return SetError(_RWRPCleanup($hProcess, $hToken, 0, 2, @error), @extended, 0)
	If Not $aRet[0] Then Return SetError(_RWRPCleanup($hProcess, $hToken, 0, 3), @extended, 0)
	$hDupToken = $aRet[6]

;~ 	Create the process using 'CreateProcessWithTokenW' (Vista+ O/S function)
	$aRet = DllCall("advapi32.dll", "bool", "CreateProcessWithTokenW", "handle", $hDupToken, "dword", 0, "wstr", $sPath, $sCmdType, $sCmd, _
			"dword", 0, "ptr", 0, $sFolderType, $sFolder, "ptr", DllStructGetPtr($stStartupInfo), "ptr", DllStructGetPtr($stProcInfo))
	$iErr = @error
	_RWRPCleanup($hProcess, $hToken, $hDupToken, 2, @error)
	If $iErr Then Return SetError(2, $iErr, 0)
	If Not $aRet[0] Then Return SetError(3, 0, 0)

;~ 	MsgBox(0,"Info","Process info data: Process handle:"&DllStructGetData($stProcInfo,1)&", Thread handle:"&DllStructGetData($stProcInfo,2)& _
;~ 		", Process ID:"&DllStructGetData($stProcInfo,3)&", Thread ID:"&DllStructGetData($stProcInfo,4)&@CRLF)

	$iRet = DllStructGetData($stProcInfo, 3) ; Process ID

;~ 	If called in 'RunWait' style, wait for the process to close
	If $bWait Then
		ProcessWaitClose($iRet)
		$iRet = @extended                  ; Exit code
	EndIf

;~ 	Close Thread and then Process handles (order here is important):
	_RWRPCleanup(0, DllStructGetData($stProcInfo, 2), DllStructGetData($stProcInfo, 1), 0)

	Return $iRet
EndFunc   ;==>_RunWithReducedPrivileges


; #FUNCTION# ============================================================================================================================
; Name...........: _ProcessGetWindow
;
; Description ...: Returns an array of HWNDs containing all windows owned by the process $p_PID, or optionally a single "best guess."
;
; Syntax.........: _ProcessGetWindow( $p_PID [, $p_ReturnBestGuess = False ])
;
; Parameters ....: $p_PID - The PID of the process you want the Window for.
;                  $p_ReturnBestGuess - If True, function will return only 1 reult on a best-guess basis.
;                                           The "Best Guess" is the VISIBLE window owned by $p_PID with the longest title.
;
; Return values .: Success      - Return $_array containing HWND info.
;                                       $_array[0] = Number of results
;                                       $_array[n] = HWND of Window n
;
;                  Failure      - Returns 0
;
;                  Error        - Returns -1 and sets @error
;                                            1 - Requires a non-zero number.
;                                            2 - Process does not exist
;                                            3 - WinList() Error
;
; Author ........: Andrew Bobulsky, contact: RulerOf <at that public email service provided by Google>.
; Remarks .......: The reverse of WinGetProcess()
; =======================================================================================================================================

Func _ProcessGetWindow($p_PID, $p_ReturnBestGuess = False)

	Local $p_ReturnVal[1] = [0]

	Local $p_WinList = WinList()

	If @error Then ;Some Error handling
		SetError(3)
		Return -1
	EndIf

	If $p_PID = 0 Then ;Some Error handling
		SetError(1)
		Return -1
	EndIf

	If ProcessExists($p_PID) = 0 Then ;Some Error handling
		ConsoleWrite("_ProcessGetWindow: Process " & $p_PID & " doesn't exist!" & @CRLF)
		SetError(2)
		Return -1
	EndIf

	For $i = 1 To $p_WinList[0][0] Step 1
		Local $w_PID = WinGetProcess($p_WinList[$i][1])

		; ConsoleWrite("Processing Window: " & Chr(34) & $p_WinList[$i][0] & Chr(34) & @CRLF & " with HWND: " & $p_WinList[$i][1] & @CRLF & " and PID: " & $w_PID & @CRLF)

		If $w_PID = $p_PID Then
			;ConsoleWrite("Match: HWND " & $p_WinList[$i][1] & @CRLF)
			$p_ReturnVal[0] += 1
			_ArrayAdd($p_ReturnVal, $p_WinList[$i][1])
		EndIf
	Next

	If $p_ReturnVal[0] > 1 Then

		If $p_ReturnBestGuess Then

			Do

				Local $i_State = WinGetState($p_ReturnVal[2])
				Local $i_StateLongest = WinGetState($p_ReturnVal[1])

				Select
					Case BitAND($i_State, 2) And BitAND($i_StateLongest, 2) ;If they're both visible
						If StringLen(WinGetTitle($p_ReturnVal[2])) > StringLen(WinGetTitle($p_ReturnVal[1])) Then ;And the new one has a longer title
							_ArrayDelete($p_ReturnVal, 1) ;Delete the "loser"
							$p_ReturnVal[0] -= 1 ;Decrement counter
						Else
							_ArrayDelete($p_ReturnVal, 2) ;Delete the failed challenger
							$p_ReturnVal[0] -= 1
						EndIf

					Case BitAND($i_State, 2) And Not BitAND($i_StateLongest, 2) ;If the new one's visible and the old one isn't
						_ArrayDelete($p_ReturnVal, 1) ;Delete the old one
						$p_ReturnVal[0] -= 1 ;Decrement counter

					Case Else ;Neither window is visible, let's just keep the first one.
						_ArrayDelete($p_ReturnVal, 2)
						$p_ReturnVal[0] -= 1

				EndSelect

			Until $p_ReturnVal[0] = 1

		EndIf

		Return $p_ReturnVal

	ElseIf $p_ReturnVal[0] = 1 Then
		Return $p_ReturnVal ;Only 1 window.
	Else
		Return 0 ;Window not found.
	EndIf
EndFunc   ;==>_ProcessGetWindow
