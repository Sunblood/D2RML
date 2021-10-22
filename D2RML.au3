#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.15.3 (Beta)
 Author:         Sunblood

 Script Function:
	Automates closing D2R multi-process watchdog handle and webtoken for logging into multiple clients

#ce ----------------------------------------------------------------------------
#RequireAdmin
#include <AutoitConstants.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

FileInstall("handle64.exe", "handle64.exe", 1)
If not FileExists("handle64.exe") Then
	MsgBox(262144+16,"Error","Error: Unable to locate handle64.exe. Download it from Github and place it in the same folder as D2RML.")
	Exit
EndIf

#Region ### START Koda GUI section ### Form=C:\Jack\Programs\D2\Multilaunch\guiMain.kxf
$guiMain = GUICreate("D2RML", 365, 259, -1, -1)
GUISetBkColor(0xC0DCC0)
$buttonAdd = GUICtrlCreateButton("Add Token", 8, 40, 75, 25)
$listViewMain = GUICtrlCreateListView("Account|Token Date", 8, 72, 250, 150, BitOR($LVS_REPORT,$LVS_SINGLESEL,$WS_VSCROLL), BitOR($WS_EX_CLIENTEDGE,$LVS_EX_CHECKBOXES,$LVS_EX_FULLROWSELECT))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 50)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 50)
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
;~ GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global Const $accountRegKey[] = ["HKEY_CURRENT_USER\SOFTWARE\Blizzard Entertainment\Battle.net\Launch Options\OSI", "WEB_TOKEN"]
Global Const $gameInstallRegKey[] = ["HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Diablo II Resurrected", "InstallLocation"]
Global Const $gameClass = "[CLASS:OsWindow]"
Global Const $bnetLauncherClass = "[CLASS:Qt5QWindowIcon]"
Global Const $bnetClientClass = "[CLASS:Chrome_WidgetWin_0]"
Global Const $version = "0.0.2"

WinSetTitle($guiMain, "", "D2RML v" & $version)
LoadAccounts()
GUISetState()

While 1

	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $buttonAdd
			Setup()
		Case $buttonLaunch
			$count = _GUICtrlListView_GetItemCount($listViewMain)
			For $i = 0 To $count
				If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
					$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
					LaunchWithAccount($name)
				EndIf
			Next
			LoadAccounts()
		Case $buttonRefresh
			$count = _GUICtrlListView_GetItemCount($listViewMain)
			For $i = 0 To $count
				If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
					$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
					Setup($name)
				EndIf
			Next
		Case $buttonRemove
			If MsgBox(262144+36, "Remove account", "Are you sure you want to remove the checked account(s)?") = 6 Then
				$count = _GUICtrlListView_GetItemCount($listViewMain)
				For $i = 0 To $count
					If _GUICtrlListView_GetItemChecked($listViewMain, $i) Then
						$name = _GUICtrlListView_GetItemText($listViewMain, $i, 0)
						FileDelete($name & ".bin")
					EndIf
				Next
				LoadAccounts()
			EndIf
		Case $labelHelp
			MsgBox(262144+64, $version, "Sun's D2R Multilauncher (D2RML) saves D2R login tokens from the registry into .bin files that are then restored on-demand." & @CRLF & _
					"It also allows multiple copies of D2R to be launched simultaneously. No additional software is required." & @CRLF & @CRLF & _
					"Initial setup requires that you log into an account at least once in order to save the token. " & _
					"Login tokens are only valid once, after which they expire and a new token is generated. D2RML does this automatically as long as you *always* use this app to launch the game." & @CRLF & _
					"Logging in through normal means will invalidate the saved token for the account and you will be unable to connect. Use the 'Refresh Token' button to redo the setup and generate a new token.")
	EndSwitch
WEnd

Func Setup($name = "")
	If ProcessExists("D2R.exe") Or ProcessExists("Battle.net.exe") Or ProcessExists("Diablo II Resurrected Launcher.exe") Then
		If MsgBox(262144+5 + 16, "Error", "D2R and Battle.net launcher(s) must be closed.") = 4 Then ;retry
			Setup($name)
			Return
		Else
			Return
		EndIf
	EndIf

	If $name = "" Then
		$name = InputBox("Setup", "Enter a name for the token:")
		If $name = "" Then Return
	EndIf

	WinClose($bnetClientClass)
	WinClose($bnetLauncherClass)

	ToolTip("Creating Tokens: " & $name & @CRLF & "Log into the Launcher with the desired account"&@CRLF&"Waiting for first launcher to open", 0, 0)
	LaunchLauncher()
	WinWait($bnetLauncherClass)
	ToolTip("Creating Tokens: " & $name & @CRLF & "Log into the Launcher with the desired account"&@CRLF&"Waiting for first launcher to close"&@CRLF&"Window title: "&WinGetTitle($bnetLauncherClass), 0, 0)
	WinWaitClose($bnetLauncherClass)
	ToolTip("Creating Tokens: " & $name & @CRLF & "Press 'Play' and connect to B.Net"&@CRLF&"Waiting for game to open", 0, 0)

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
	MsgBox(262144+64, "Finished", "Successfully saved token: " & $name)
EndFunc   ;==>Setup

Func LaunchWithAccount($name)
	ToolTip("Launching token: " & $name, 0, 0)
	WriteRegKey($name & ".bin")
	$curKey = RegRead($accountRegKey[0], $accountRegKey[1])
	$pid = LaunchGame()

	ToolTip("Launching token: " & $name & @CRLF & "Refreshing token 1 of 2...", 0, 0)
	WaitForNewKey()
	ToolTip("Launching token: " & $name & @CRLF & "Refreshing token 2 of 2...", 0, 0)
	WaitForNewKey()

	ToolTip("")

	If ProcessExists("D2R.exe") Then
		ExportRegKey($name & ".bin")
		CloseMultiProcessHandle($pid)
	Else
		MsgBox(262144+16, "Error", "Error obtaining new tokens. You may need to refresh the token you're trying to use.")
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
		MsgBox(262144+16, "Error", "Error closing multi-instance handle for PID " & $pid & @CRLF & "Handle: " & $handle)
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
		If TimerDiff($sendTimer) > 500 and TimerDiff($queueTimer) < 15000 Then
			ControlSend($gameClass, "", "", "{SPACE}")
			$sendTimer = TimerInit()
		EndIf
		If Not ProcessExists("D2R.exe") Then Return
	Until $newKey <> $curKey
EndFunc   ;==>WaitForNewKey

Func LaunchGame()
	$path = RegRead($gameInstallRegKey[0], $gameInstallRegKey[1])
	If GUICtrlRead($checkboxArgs) = $GUI_CHECKED Then
		Return ShellExecute($path & "\D2R.exe",GUICtrlRead($inputArgs))
	Else
		Return ShellExecute($path & "\D2R.exe")
	EndIf
EndFunc   ;==>LaunchGame
Func LaunchLauncher() ;hehe
	$path = RegRead($gameInstallRegKey[0], $gameInstallRegKey[1])
	Return ShellExecute($path & "\Diablo II Resurrected Launcher.exe")
EndFunc   ;==>LaunchLauncher

Func GetFilename($file)
	$a = StringSplit($file, "\")
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
	$szRoot = ""
	$hFile = 0
	$szBuffer = ""
	$szReturn = ""
	$szPathList = "*"
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
