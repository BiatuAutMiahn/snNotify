#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ServiceNow.ico
#AutoIt3Wrapper_Outfile_x64=..\_.rc\snNotify.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=ServiceNow Notifier
#AutoIt3Wrapper_Res_Fileversion=23.322.1203.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_Fileversion_Use_Template=%YY.%MO%DD.%HH%MI.%SE
#AutoIt3Wrapper_Res_ProductName=snNotify
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
Opt("TrayMenuMode", 3)

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]

 Script Function:
    Notify User when new task/incident is assigned.
#ce ----------------------------------------------------------------------------
#include <Debug.au3>
#include <File.au3>
#include <WinAPISys.au3>
#include <Array.au3>
#include <TrayConstants.au3>
#include <Timers.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>

#include "Includes\Common\Base64.au3"
#include "Includes\Common\Toast.au3"
#include "Includes\Common\_StringInPixels.au3"
#include "Includes\Au3ServiceNow\_snCommon.au3"

$g_bSingleInstance=True
Global Const $VERSION = "23.322.1203.1"

Global $g_AddWatch_idInput, $g_AddWatch_idAdd, $g_AddWatch_hWnd
Global $fToast_OpenTik = False, $hToast_OpenTik
Global $oError, $sMagic = _RandStr()
Global $sSaveSect
Global $bOptRMH = False
Global $bOptBAC = False
Global $sAlias
Global $gbQuery = False
Global $gsQuery = ""
If $bOptRMH Then
	$sSaveSect = "RMH"
	$sAlias = "snNotify (RMH) v" & $VERSION
ElseIf $bOptBAC Then
	$sSaveSect = "BAC"
	$sAlias = "snNotify (BAC) v" & $VERSION
Else
	$sSaveSect = "User"
	$sAlias = "snNotify v" & $VERSION
EndIf

; Error Handling and Logging.
$oError = ObjEvent("AutoIt.Error", "onError")
Global $bCompiled = @Compiled
Func onError()
	$sErr = $oError.windescription & " (" & $oError.number & ")"
	If @Compiled Then
		_snLog($sErr)
	Else
		ConsoleWrite($sErr & @CRLF)
	EndIf
	Return SetError(1, $oError.number, $sErr)
EndFunc   ;==>onError

Func _snLog($sStr, $sFunc = "Main")
	Local $sStamp = @YEAR & '.' & @MON & '.' & @MDAY & ',' & @HOUR & ':' & @MIN & ':' & @SEC & ':' & @MSEC
	Local $sErr = "+[" & $sMagic & '|' & $sStamp & "|" & @ComputerName & "|" & @UserName & "|" & @ScriptName & "|" & $sFunc & "]: " & $sStr
	Local $sdDir = @AppDataDir & "\OhioHealth"
	If Not DirCreate($sdDir) Then
		MsgBox(16, "Error", "Cannot Create Data Dir...Exiting.")
		Exit 0
	EndIf
	If @Compiled Then
		If Not FileWriteLine($sdDir & '\' & @ScriptName & ".log", $sErr) Then
			MsgBox(16, "Error", "Cannot write to log...Exiting.")
			Exit 0
		EndIf
	Else
		ConsoleWrite($sErr & @CRLF)
	EndIf
	Return
EndFunc   ;==>_snLog


#cs

snNotify, Notifies when new tasks are assigned.

-Notify when task has been added
-task will breach soon 48 hours
-task has breached

On First run, get User's sys_id with @username. If fail, prompt end user for UserId.
Tray menu, Change User|Exit


Array of tasks
$sNumber as been assigned to you.
Priority: $sPriority
Description: $sDesc

#ce

Global $aFields = StringSplit("number|sys_class_name|active|state|priority|opened_by|opened_at|short_description|sys_id|assigned_to|assignment_group", '|')
Global $aFieldNames = StringSplit("Number|Class Name|Active|State|Priority|Opened By|Opened At|Short Description|Sys id|Assigned To|Assignment Group", '|')

Global $aWatchFields[] = [4, 3, 4, 9, 10] ;%|state|priority|assigned_to|assignment_group
Global $abOptNotify[] = [4, 1, 1, 1, 1]
Global $aidNotify[] = [0]
Global $bWsLockLast, $bWsLock
Global $iUserIdleTime = 0
Global $bUserIdleLast, $bUserIdle
Global $aPurge[0]
Global $aTasks[1][$aFields[0]], $aTasksLast[1][$aFields[0]]
Local $iTimer, $bFirstRun = True
Local $sMsg, $aRet[2]
$sDataDir = @AppDataDir & "\OhioHealth"
DirCreate($sDataDir)
$sConfigFile = $sDataDir & "\snNotify.ini"
_Toast_Set(0, -1, -1, -1, -1, -1, "Consolas", 125, 125)

; For when we need users to close app for update.
_CheckMaint()

;Purge log file if it gets too big.
If FileGetSize($sDataDir & '\' & @ScriptName & ".log") >= 1048576 Then
	FileDelete($sDataDir & '\' & @ScriptName & ".log")
	_snLog("Log has been purged >=1MB")
EndIf

;Load saved data.
If FileExists($sConfigFile) Then
	IniWrite($sConfigFile, $sSaveSect, "NotifyOpts", _Base64Encode(_ArrayToString($abOptNotify, Chr(0xFE), 0, Default, Chr(0xFD))))
	$sNotifyOptData = IniRead($sConfigFile, $sSaveSect, "NotifyOpts", "")
	If $sNotifyOptData <> "" Then
		$sb64NotifyOpts = BinaryToString(_Base64Decode($sNotifyOptData))
		$abOptNotify = _ArrayFromString($sb64NotifyOpts, Chr(0xFE), Chr(0xFD))
	EndIf
	$sTaskData = IniRead($sConfigFile, $sSaveSect, "Cache", "")
	$sTaskLastData = IniRead($sConfigFile, $sSaveSect, "CacheLast", "")
	If $sTaskData <> "" Then
		_snLog('Found Cached Task Data')
		$sb64Tasks = BinaryToString(_Base64Decode($sTaskData))
		If $sTaskLastData <> "" Then
			_snLog('Found Cached TaskLast Data')
			$sb64TasksLast = BinaryToString(_Base64Decode($sTaskLastData))
			$aTasksLast = _ArrayFromString($sb64TasksLast, Chr(0xFE), Chr(0xFD))
		Else
			_snLog('Cached TaskLast Data No Found')
			$aTasksLast = _ArrayFromString($sb64Tasks, Chr(0xFE), Chr(0xFD))
		EndIf
		$aTasks = _ArrayFromString($sb64Tasks, Chr(0xFE), Chr(0xFD))
		If UBound($aTasks, 2) < $aFields[0] Then
			For $i = UBound($aTasks, 2) To $aFields[0] - 1
				_ArrayColInsert($aTasks, $i)
			Next
		EndIf
		If UBound($aTasksLast, 2) < $aFields[0] Then
			For $i = UBound($aTasksLast, 2) To $aFields[0] - 1
				_ArrayColInsert($aTasksLast, $i)
			Next
		EndIf
	EndIf
	_snLog("Loaded Tasks:" & @CRLF & _ArrayToString($aTasks) & @CRLF)
EndIf

; Tray Config
Opt("TrayIconHide", 0)
TraySetToolTip($sAlias)
Local $bExit = False
$idTrayAddWatch = TrayCreateItem("Watch Custom INC/SCTASK")
TrayCreateItem("")
$idNotifyMenu = TrayCreateMenu("NotifyOn")
For $i = 1 To $aWatchFields[0]
	$aidNotify[0] = UBound($aidNotify, 1)
	ReDim $aidNotify[$aidNotify[0] + 1]
	$aidNotify[$aidNotify[0]] = TrayCreateItem($aFieldNames[$aWatchFields[$i] + 1], $idNotifyMenu)
	TrayItemSetState($aidNotify[$aidNotify[0]], $abOptNotify[$i] ? $TRAY_CHECKED : $TRAY_UNCHECKED)
	_snLog("OptNotify," & $aFields[$aWatchFields[$i] + 1] & ':' & $abOptNotify[$i])
Next
TrayCreateItem("")
$idTrayExit = TrayCreateItem("Exit")
Func _TrayEvent()
	$iTrayMsg = TrayGetMsg()
	Switch $iTrayMsg
		Case $idTrayExit
			AdlibUnRegister("_TrayEvent")
			$bExit = True
			_Toast_Hide()
			_Exit()
        Case $idTrayAddWatch
            _AddWatch()
	EndSwitch
	Local $iNewState = 0, $iTrayCheck = BitOR($TRAY_ENABLE, $TRAY_CHECKED), $iTrayUncheck = BitOR($TRAY_ENABLE, $TRAY_UNCHECKED)
	For $i = 1 To $aidNotify[0]
		If $iTrayMsg == $aidNotify[$i] Then
			$iState = TrayItemGetState($aidNotify[$i])
			If $iState == $iTrayCheck Then
				$iNewState = $TRAY_UNCHECKED
				$abOptNotify[$i] = 0
			ElseIf $iState == $iTrayUncheck Then
				$iNewState = $TRAY_CHECKED
				$abOptNotify[$i] = 1
			EndIf
			TrayItemSetState($aidNotify[$i], $iNewState)
			IniWrite($sConfigFile, $sSaveSect, "NotifyOpts", _Base64Encode(_ArrayToString($abOptNotify, Chr(0xFE), 0, Default, Chr(0xFD))))
			_snLog("ModOptNotify," & $aFields[$aWatchFields[$i]] & ':' & $abOptNotify[$i])
		EndIf
	Next
EndFunc   ;==>_TrayEvent
AdlibRegister("_TrayEvent", 20)

If $bExit Then
	_Exit()
EndIf

If FileExists($sConfigFile) Then
	$aRet = _Toast_ShowMod(0, $sAlias, "Welcome Back, watching for new tickets...", -5)
Else
	$sMsg = "Welcome! snNotify will watch your ServiceNow" & @CRLF
	$sMsg &= "Task queue for new Tasks/Incidents." & @CRLF
	$sMsg &= "On first run you will recieve notifications" & @CRLF
	$sMsg &= "for existing tickets." & @CRLF
	$sMsg &= @CRLF
	$sMsg &= "To Exit or Configure, right click the tray" & @CRLF
	$sMsg &= "icon by the clock and select exit. There, you" & @CRLF
	$sMsg &= "can also change how your notified." & @CRLF
	$sMsg &= @CRLF

	$sMsg &= "To Continue, click the X at the top of this" & @CRLF
	$sMsg &= "notification, or wait 60 seconds."
	$aRet = _Toast_ShowMod(0, $sAlias, $sMsg, -60)
EndIf
_Toast_Hide()

$oQueryUsers = _queryUserId(@UserName)
$sUserId = __snSoapGetAttr($oQueryUsers, "sys_id")
$sUserName = __snSoapGetAttr($oQueryUsers, "name")

TraySetToolTip($sAlias & " - Waiting for Tickets...")
$iTimer = TimerInit()
Local $bDev = @Compiled ? False : True
;If $bDev Then _ArrayDisplay($aTasksLast)
While 1
	$iUserIdleTime = _Timer_GetIdleTime()
	$bUserIdle = $iUserIdleTime >= 15000 ? True : False
	$bWsLock = _isWindowsLocked()
	If $bUserIdleLast <> $bUserIdle Then
		$bUserIdleLast = $bUserIdle
		If $bUserIdle Then
			_snLog("UserIdle")
		Else
			_Toast_Hide()
			_snLog("UserNotIdle")
			Sleep(10000)
		EndIf
	EndIf
	If $bWsLockLast <> $bWsLock Then
		If $bWsLock Then
			_snLog("WorkstationLocked")
		Else
			_Toast_Hide()
			_snLog("WorkstationNotLocked")
			Sleep(10000)
		EndIf
		$bWsLockLast = $bWsLock
	EndIf
	If $bWsLock Or $bUserIdle Then
		ContinueLoop
	EndIf
	If TimerDiff($iTimer) >= 10000 Or $bFirstRun Then
		_CheckMaint()
		$bFirstRun = False
		$bNew = False
		$iTaskTimer = TimerInit()
		$aUserTasks = _getUserTasks()
		; Build Query of Existing Tasks
		$sQuery = ""
		For $i = 1 To $aTasks[0][0]
			$sQuery &= 'sys_id=' & $aTasks[$i][8]
			If $i < $aTasks[0][0] Then $sQuery &= '^OR'
		Next
		$aQuery = _getUserTasks($sQuery)
		;_ArrayDisplay($aQuery)
		;MsgBox(64,"",$sQuery)
		_snLog("Retrieval Time: " & TimerDiff($iTaskTimer) & "ms")
		_snLog("Refreshing Existing Tasks...")
		;_ArrayDisplay($aTasks)
		;_ArrayDisplay($aTasksLast)
		For $i = 1 To $aTasks[0][0]
			$k = 0
			For $l = 1 To $aQuery[0][0]
				If $aQuery[$l][8] == $aTasks[$i][8] Then
					For $j = 1 To $aFields[0]
						$aTasks[$i][$j - 1] = $aQuery[$l][$j - 1]
					Next
				EndIf
			Next
			;number|sys_class_name|active|state|priority|opened_by|opened_at|short_description|sys_id|assigned_to|assignment_group
			_Toast_Set(0, -1, -1, -1, -1, -1, "Consolas", 125, 125)
			Local $bShowToast = False, $iTaskMods = 0, $bModState = False, $sTitle = ""
			Local $sMsg = "Type:             " & $aTasks[$i][1] & @CRLF
			$sMsg &= "Number:           " & $aTasks[$i][0] & @CRLF
			Local $sWatchFields
			Local $sTmp = ""
            If Not Int($aTasks[$i][2]) Then
                $iMax=UBound($aPurge,1)
                ReDim $aPurge[$iMax+1]
                $aPurge[$iMax]=$i
            EndIf
			For $j = 1 To $aFields[0]
				If StringCompare($aTasks[$i][$j - 1], $aTasksLast[$i][$j - 1]) <> 0 Then
					$sWatchFields = "state|priority|assigned_to|assignment_group"
					_snLog("snTaskMod [" & $aFields[$j] & ']: "' & $aTasksLast[$i][$j - 1] & '" -> "' & $aTasks[$i][$j - 1] & '"')
					If StringInStr($sWatchFields, $aFields[$j]) Then
						Switch $aFields[$j]
							Case "state"
								$sMsg &= "State:            "
							Case "priority"
								$sMsg &= "Priority:         "
							Case "assigned_to"
								$sMsg &= 'Assigned To:      '
							Case "assignment_group"
								$sMsg &= 'Assignment Group: '
						EndSwitch
						If $aTasksLast[$i][$j - 1] == "" Then
							$sTmp = 'Empty -> ' & $aTasks[$i][$j - 1] & @CRLF
						ElseIf $aTasks[$i][$j - 1] == "" Then
							$sTmp = $aTasksLast[$i][$j - 1] & ' -> Empty' & @CRLF
						Else
							$sTmp = $aTasksLast[$i][$j - 1] & ' -> ' & $aTasks[$i][$j - 1] & @CRLF
						EndIf
						If StringLen($sTmp) > 40 Then $sMsg &= @CRLF
						$sMsg &= $sTmp
						If StringLen($sTmp) > 40 Then $sMsg &= @CRLF
						$iTaskMods += 1
						$bShowToast = True
						$bModState = True
					EndIf
				Else
					;If Not ($bOptRMH And $bOptBAC) Then
					;	$sWatchFields = "state|priority|assignment_group"
					;Else
					;   $sWatchFields = "state|priority|assigned_to|assignment_group"
					;EndIf
					If StringInStr($sWatchFields, $aFields[$j]) Then
						Switch $aFields[$j]
							Case "state"
								$sMsg &= "State:            "
							Case "priority"
								$sMsg &= "Priority:         "
							Case "assigned_to"
								$sMsg &= 'Assigned To:      '
							Case "assignment_group"
								$sMsg &= 'Assignment Group: '
						EndSwitch
						If StringLen($aTasks[$i][$j - 1]) > 40 Then $sMsg &= @CRLF
                        If $aTasks[$i][$j - 1]=="" Then
                            $sMsg &= "Empty"
                        Else
                            $sMsg &= $aTasks[$i][$j - 1]
                        EndIf
                        $sMsg &= @CRLF
						If StringLen($aTasks[$i][$j - 1]) > 40 Then $sMsg &= @CRLF
					EndIf
				EndIf
			Next
			;[3] = state
            If $aTasks[$i][3]<>$aTasksLast[$i][3] And _isOptNotify(3) Then
				If StringCompare($aTasks[$i][1],"Catalog Task")==0 Then
					If $aTasks[$i][3] == "Closed Complete" And $aTasksLast[$i][3] <> "Closed Complete" Then
						$sTitle = "Task Completed!"
                    Else
                        $sTitle = "Task Changed State!"
					EndIf
                    $bShowToast = True
                    $bModState = True
				ElseIf StringCompare($aTasks[$i][1],"Incident")==0 Then
					If $aTasks[$i][3] == "Closed" And $aTasksLast[$i][3] <> "Closed" Then
						$sTitle = "Incident Closed!"
					ElseIf StringInStr($aTasks[$i][3], "Resolved") And Not StringInStr($aTasksLast[$i][3], "Resolved") Then
						$sTitle = "Incident Resolved!"
                    ElseIf StringCompare($aTasks[$i][3],"On Hold")==0 Then
                        $sTitle = "Incident Placed on Hold!"
                    Else
                        $sTitle = "Incident Changed State!"
                    EndIf
                    $bShowToast = True
                    $bModState = True
				Else
					_snLog("Unhandled sys_class_name: " & $aTasks[$i][1])
				EndIf
            EndIf
            If $iTaskMods > 0 Then
                If $sTitle == "" Then $sTitle = $aTasks[$i][1] & " Modified!"
                ;[4] = priority
                If $aTasks[$i][4] <> $aTasksLast[$i][4] And _isOptNotify(4) Then
                    $sTitle = $aTasks[$i][1] & " Reprioritized!"
                    $bShowToast = True
                EndIf
                ;[9] = assigned_to
                If ($aTasks[$i][9] <> $aTasksLast[$i][9]) Or ($aTasks[$i][10] <> $aTasksLast[$i][10]) And (_isOptNotify(9) Or _isOptNotify(9)) Then
                    If $aTasksLast[$i][9] == $sUserName And $aTasks[$i][9] <> $sUserName Then;
                        $sTitle = $aTasks[$i][1] & " Unassigned!"
                    Else
                        $sTitle = $aTasks[$i][1] & " Reassigned!"
                    EndIf
                    $bShowToast = True
                EndIf
            EndIf
;~ 			If _isOptNotify(3) Then
;~ 				If $aTasks[$i][1] == "Catalog Task" Then
;~ 					If $aTasks[$i][3] == "Closed Complete" And $aTasksLast[$i][3] <> "Closed Complete" Then
;~ 						$sTitle = "Task Completed!"
;~ 						$bShowToast = True
;~ 						$bModState = True
;~ 					EndIf
;~ 				ElseIf $aTasks[$i][1] == "Incident" Then
;~ 					If $aTasks[$i][3] == "Closed" And $aTasksLast[$i][3] <> "Closed" Then
;~ 						$sTitle = "Incident Closed!"
;~ 						$bShowToast = True
;~ 						$bModState = True
;~ 					ElseIf StringInStr($aTasks[$i][3], "Resolved") And Not StringInStr($aTasksLast[$i][3], "Resolved") Then
;~ 						$sTitle = "Incident Resolved!"
;~ 						$bShowToast = True
;~ 						$bModState = True
;~ 					EndIf
;~ 				Else
;~ 					_snLog("Unhandled sys_class_name: " & $aTasks[$i][1])
;~ 				EndIf
;~ 			EndIf
			$sMsg &= "Description: " & @CRLF & @CRLF
			$sMsg &= $aTasks[$i][7]
			If $bShowToast Then
				$bShowToast = False
				$aRet = _Toast_ShowMod(0, $sTitle, $sMsg, -40, True, True)
                ;Exit
				If $fToast_OpenTik Then
					If StringCompare($aTasks[$i][1],"Catalog Task")==0 Then
						ShellExecute("https://ohiohealth.service-now.com/nav_to.do?uri=sc_task.do?sys_id=" & $aTasks[$i][8])
					ElseIf StringCompare($aTasks[$i][1],"Incident")==0 Then
						ShellExecute("https://ohiohealth.service-now.com/nav_to.do?uri=incident.do?sys_id=" & $aTasks[$i][8])
					EndIf
				EndIf
				$fToast_OpenTik = False
				_Toast_Hide()
				$bNew = True
				$iTimer = TimerInit()
			EndIf
;~             Else
;~                 _snLog("RefreshTask, oTask is not an Object: "&VarGetType($oTask)&','&$aTasks[$i][1]&","&$aTasks[$i][8])
;~             EndIf

		Next
		_snLog("Refreshing Existing Tasks...Done")
;~         _snLog("Syncing TasksLast")
;~         For $i=1 To $aTasksLast[0][0]
;~             For $j=1 To $aFields[0]
;~                 $aTasksLast[$i][$j-1]=$aTasks[$i][$j-1]
;~             Next
;~         Next
;~         If Not $bDev Then
;~             _snLog("Saving Data")
;~             $sTasksLast=_ArrayToString($aTasks,Chr(0xFE),0,Default,Chr(0xFD))
;~             $sb64TasksLast=_Base64Encode($sTasksLast)
;~             $sTasks=_ArrayToString($aTasks,Chr(0xFE),0,Default,Chr(0xFD))
;~             $sb64Tasks=_Base64Encode($sTasks)
;~             IniWrite($sConfigFile,$sSaveSect,"Cache",$sb64Tasks)
;~             IniWrite($sConfigFile,$sSaveSect,"CacheLast",$sb64TasksLast)
;~         EndIf
		For $i = 1 To $aUserTasks[0][0]
			For $j = 1 To $aTasks[0][0]
				If $aUserTasks[$i][0] == $aTasks[$j][0] Then
					; Task Already Exists
					ContinueLoop 2
				EndIf
			Next
			Local $aTaskData[$aFields[0]]
			$bNew = True
			; Task is new, append.
			$aTasks[0][0] = UBound($aTasks, 1)
			ReDim $aTasks[$aTasks[0][0] + 1][$aFields[0]]
			ReDim $aTasksLast[$aTasks[0][0] + 1][$aFields[0]]
			For $j = 1 To $aFields[0]
				$aTasks[$aTasks[0][0]][$j - 1] = $aUserTasks[$i][$j - 1]
				$aTasksLast[$aTasks[0][0]][$j - 1] = $aUserTasks[$i][$j - 1]
				$aTaskData[$j - 1] = $aUserTasks[$i][$j - 1]
			Next
			_Toast_Set(0, -1, -1, -1, -1, -1, "Consolas", 125, 125)
			$sTitle = "New Ticket!"
			$sMsg = "Type:                " & $aTasks[$aTasks[0][0]][1] & @CRLF
			$sMsg &= "Number:              " & $aTasks[$aTasks[0][0]][0] & @CRLF
			$sMsg &= "Priority:            " & $aTasks[$aTasks[0][0]][4] & @CRLF
			;If $bOptRMH Or $bOptBAC Then
			$sMsg &= "Assigned To:         " & $aTasks[$aTasks[0][0]][9] & @CRLF
			;EndIf
			$sMsg &= "Assignment Group:    " & $aTasks[$aTasks[0][0]][10] & @CRLF
			$sMsg &= "Description:         " & @CRLF & @CRLF & $aTasks[$aTasks[0][0]][7]
			_snLog("New Task:" & @CRLF & _ArrayToString($aTaskData))
            ;If $bDev Then _ArrayDisplay($aTasksLast)
			$aRet = _Toast_ShowMod(0, $sTitle, $sMsg, -40, True, True)
			;_snLog(@CRLF&$sMsg)
			If $fToast_OpenTik Then
				If StringCompare($aTasks[$aTasks[0][0]][1],"Catalog Task")==0 Then
					ShellExecute("https://<redacted>.service-now.com/nav_to.do?uri=sc_task.do?sys_id=" & $aTasks[$aTasks[0][0]][8])
				ElseIf StringCompare($aTasks[$aTasks[0][0]][1],"Incident")==0 Then
					ShellExecute("https://<redacted>.service-now.com/nav_to.do?uri=incident.do?sys_id=" & $aTasks[$aTasks[0][0]][8])
				EndIf
			EndIf
			$fToast_OpenTik = False
			_Toast_Hide()
			If $bExit Then
				_Exit()
			EndIf
		Next
        ; Purge Inactive Items.
        If UBound($aPurge,1)>0 Then
            Local $aNew[1][$aFields[0]],$iMax=-1
            For $x=1 To $aTasks[0][0]
                For $y=0 To UBound($aPurge,1)-1
                    If $x==$aPurge[$y] Then ContinueLoop 2
                Next
                $iMax=UBound($aNew,1)
                ReDim $aNew[$iMax+1][$aFields[0]]
                For $z=1 To $aFields[0]
                    $aNew[$iMax][$z-1]=$aTasks[$x][$z-1]
                Next
            Next
            $aNew[0][0]=$iMax
            $aTasks=$aNew
            Dim $aPurge[0]
            Dim $aNew[1][$aFields[0]]
        EndIf
		If $bNew Then
			If Not $bDev Then
				_snLog("Saving Data")
				$sTasksLast = _ArrayToString($aTasks, Chr(0xFE), 0, Default, Chr(0xFD))
				$sb64TasksLast = _Base64Encode($sTasksLast)
				$sTasks = _ArrayToString($aTasks, Chr(0xFE), 0, Default, Chr(0xFD))
				$sb64Tasks = _Base64Encode($sTasks)
				IniWrite($sConfigFile, $sSaveSect, "Cache", $sb64Tasks)
				IniWrite($sConfigFile, $sSaveSect, "CacheLast", $sb64TasksLast)
			EndIf

			;If $bDev Then MsgBox(64,"","snNotify")
;~             If $aTasks[$i][8]=="" Then
;~                 MsgBox(64,$sTitle,$sMsg)
;~                 MsgBox(64,$bShowToast,$iTaskMods)
;~                 _DebugArrayDisplay($aTasks)
;~                 _DebugArrayDisplay($aTasksLast)
;~             EndIf
;~             If $aTasks[$i][8]=="" Then
;~                 _DebugArrayDisplay($aTasks)
;~                 _DebugArrayDisplay($aTasksLast)
;~             EndIf
		EndIf
		_snLog("Syncing TasksLast")
		$aTasksLast = $aTasks
;~         For $i=1 To $aTasksLast[0][0]
;~             For $j=1 To $aFields[0]
;~                 $aTasksLast[$i][$j-1]=$aTasks[$i][$j-1]
;~             Next
;~         Next
		$iTimer = TimerInit()
	EndIf
	If $bExit Then
		_Exit()
	EndIf
	Sleep(125)
WEnd

;~ Global $aWatchFields[]=[4,4,5,10,11] ;%|state|priority|assigned_to|assignment_group
;~ Global $abOptNotify[]=[4,0,1,1,1]
;~ Global $aidNotify[]=[0]
Func _isOptNotify($iIdx)
	For $i = 1 To $aWatchFields[0]
		If $iIdx <> $aWatchFields[$i] Then ContinueLoop
		If $abOptNotify[$i] == 1 Then Return True
	Next
	Return False
EndFunc   ;==>_isOptNotify

Func _isWindowsLocked()
	If _WinAPI_OpenInputDesktop() Then Return False
	Return True
EndFunc   ;==>_isWindowsLocked

Func _CheckMaint()
	If FileExists(@ScriptDir & "\InfinityHook.ini") Then
		If IniRead(@ScriptDir & "\InfinityHook.ini", "snNotify", "MaintMode", "") == "True" Then
			_Toast_Set(0, -1, -1, -1, -1, -1, "Consolas", 125, 125)
			$aRet = _Toast_ShowMod(0, $sAlias, "Developer requires snNotify to Exit for maintenance! To exit now press the [x] or wait 60 seconds.", -60)
			_Exit()
		EndIf
	EndIf
EndFunc   ;==>_CheckMaint

Func _Exit()
	_Toast_Set(0, -1, -1, -1, -1, -1, "Consolas", 125, 125)
	$aRet = _Toast_ShowMod(0, $sAlias, "Exiting...                        ", -2)
	_Toast_Hide()
	_snLog('Exit')
	Exit
EndFunc   ;==>_Exit

Exit

Func _getUserTasks($sQuery = Default)
	Local $aReturn[1][$aFields[0]]
	If $sQuery == Default Then
		If $gbQuery Then
			$sQuery = $gsQuery
		ElseIf $bOptRMH Then
			$sQuery = "active=true^assignment_group=b8572ded1bd8cc500210ed776e4bcbe8^ORassignment_group=f4572ded1bd8cc500210ed776e4bcbe6^state!=-5^ORnumberSTARTSWITHprjtask^state!=6"
		ElseIf $bOptBAC Then
			$sQuery = "active=true^assignment_group=6e4769ed1bd8cc500210ed776e4bcb18^state!=-5^ORnumberSTARTSWITHprjtask"
		Else
			$sQuery = "active=true^assigned_to=" & $sUserId & "^state!=-5^state!=6"
		EndIf
	EndIf

	$aTasksQuery = _queryTasks($sQuery)
	$aIncQuery = _queryIncidents($sQuery)
	;_ArrayDisplay($aTasksQuery,"$aTasksQuery")
	;_ArrayDisplay($aIncQuery,"$aIncQuery")
	_snLog("_ProcTasks,$aTasksQuery")
	_ProcTasks($aReturn, $aTasksQuery)
	If @error Then
		_snLog("@Error," & @error & ',$aTasksQuery,' & IsArray($aReturn) & ',' & IsArray($aTasksQuery) & ':' & @CRLF & _ArrayToString($aTasksQuery), "_getUserTasks")
	EndIf
	_snLog("_ProcTasks,$aIncQuery")
	_ProcTasks($aReturn, $aIncQuery)
	If @error Then
		_snLog("@Error," & @error & ',$aIncQuery,' & IsArray($aReturn) & ',' & IsArray($aTasksQuery) & ':' & @CRLF & _ArrayToString($aIncQuery), "_getUserTasks")
	EndIf
	Return $aReturn
EndFunc   ;==>_getUserTasks

Func _ProcTasks(ByRef $aReturn, ByRef $aArr)
	If Not IsArray($aArr) Then Return SetError(1, 0, False)
	If UBound($aArr, 1) = 0 Then Return SetError(2, 0, False)
	For $i = 1 To $aArr[0]
		$sNumber = __snSoapGetAttr($aArr[$i], "number")
		For $j = 1 To $aReturn[0][0]
			If $aReturn[$j][0] == $sNumber Then
				; Check if New
				ContinueLoop 2
			EndIf
		Next
		$aReturn[0][0] = UBound($aReturn, 1)
		ReDim $aReturn[$aReturn[0][0] + 1][$aFields[0]]
		For $k = 1 To $aFields[0]
			$aReturn[$aReturn[0][0]][$k - 1] = __snSoapGetAttr($aArr[$i], $aFields[$k])
		Next
;~         $aReturn[$aReturn[0][0]][0]=$sNumber
;~         $aReturn[$aReturn[0][0]][1]=_soapGetAttr($aArr[$i],"sys_class_name")
;~         $aReturn[$aReturn[0][0]][2]=_soapGetAttr($aArr[$i],"active")
;~         $aReturn[$aReturn[0][0]][3]=_soapGetAttr($aArr[$i],"state")
;~         $aReturn[$aReturn[0][0]][4]=_soapGetAttr($aArr[$i],"priority")
;~         $aReturn[$aReturn[0][0]][5]=_soapGetAttr($aArr[$i],"opened_by")
;~         $aReturn[$aReturn[0][0]][6]=_soapGetAttr($aArr[$i],"opened_at")
;~         $aReturn[$aReturn[0][0]][7]=_soapGetAttr($aArr[$i],"short_description")
;~         $aReturn[$aReturn[0][0]][8]=_soapGetAttr($aArr[$i],"sys_id")
;~         $aReturn[$aReturn[0][0]][9]=_soapGetAttr($aArr[$i],"assigned_to")
	Next
EndFunc   ;==>_ProcTasks

Func _queryRequests($sQuery)
	Return _snQuery($sQuery, "sc_request_list")
EndFunc   ;==>_queryRequests

Func _queryComputers($sQuery)
	Return _snQuery($sQuery, "cmdb_ci_computer")
EndFunc   ;==>_queryComputers

Func _queryUserId($sUserId)
	Return _queryUsers('user_name=' & $sUserId)
EndFunc   ;==>_queryUserId

Func _queryUsers($sQuery)
	Return _snQuery($sQuery, "sys_user_list")
EndFunc   ;==>_queryUsers

Func _queryPrinters()
	Return _snQuery($sQuery, "cmdb_ci_printer_list")
EndFunc   ;==>_queryPrinters

Func _queryItems()
	Return _snQuery($sQuery, "sc_req_item_list")
EndFunc   ;==>_queryItems

Func _queryTasks($sQuery)
	Return _snQuery($sQuery, "sc_task_list")
EndFunc   ;==>_queryTasks

Func _queryIncidents($sQuery)
	Return _snQuery($sQuery, "incident_list")
EndFunc   ;==>_queryIncidents

Func _getIncident($sId)
	Return _snGetById($sId, "incident")
EndFunc   ;==>_getIncident

Func _getRequest($sId)
	Return _snGetById($sId, "sc_request")
EndFunc   ;==>_getRequest
Func _getItem($sId)
	Return _snGetById($sId, "sc_req_item")
EndFunc   ;==>_getItem
Func _getComputer($sId)
	Return _snGetById($sId, "cmdb_ci_computer")
EndFunc   ;==>_getComputer
Func _getUser($sId)
	;& '		<__encoded_query>assigned_to=51f2b21f87629110e2f852083cbb35ab</__encoded_query>' & @CRLF _
	;GOTOuser_name=<redacted>
	Return _snGetById($sId, "sys_user")
EndFunc   ;==>_getUser
Func _getTask($sId)
	;sc_task
	$SoapMsg = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:inc="http://<redacted>.service-now.com/sc_task">' & @CRLF _
			 & '   <soapenv:Header/>' & @CRLF _
			 & '   <soapenv:Body>' & @CRLF _
			 & '	  <inc:get>' & @CRLF _
			 & '		<sys_id>' & $sId & '</sys_id>' & @CRLF _
			 & '	  </inc:get>' & @CRLF _
			 & '   </soapenv:Body>' & @CRLF _
			 & '</soapenv:Envelope>'
	_snLog(@CRLF & $SoapMsg, '_getTask')
	$sQuery = __snSoapQuery("sc_task", $SoapMsg)
	$oXml = ObjCreate("Msxml2.DOMDocument.3.0")
	;ClipPut($sQuery)
	$oXml.loadXML($sQuery)
	$oEnvelope = __snSoapGetNode($oXml, "SOAP-ENV:Envelope")
	$oBody = __snSoapGetNode($oEnvelope, "SOAP-ENV:Body")
	$oRecordsResponse = __snSoapGetNode($oBody, "getResponse")
	;$oRecordsResults=__snSoapGetNode($oRecordsResponse,"getRecordsResult")
	Return $oRecordsResponse
EndFunc   ;==>_getTask
Func _getPrinter()
	;cmdb_ci_printer
EndFunc   ;==>_getPrinter

;-------------Scratch Space Blow-------------

;~ ClipPut($sQuery)

;~ $SoapMsg = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:inc="http://<redacted>.service-now.com/sc_req_item">' & @CRLF _
;~                 & '   <soapenv:Header/>' & @CRLF _
;~                 & '   <soapenv:Body>' & @CRLF _
;~                 & '	  <inc:get>' & @CRLF _
;~                 & '		<sys_id>f04be3e19784ad54e35ffdf3a253affa</sys_id>' & @CRLF _
;~                 & '	  </inc:get>' & @CRLF _
;~                 & '   </soapenv:Body>' & @CRLF _
;~                 & '</soapenv:Envelope>'
;~



;~ $SoapMsg = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:inc="http://<redacted>.service-now.com/sc_request">' & @CRLF _
;~                 & '   <soapenv:Header/>' & @CRLF _
;~                 & '   <soapenv:Body>' & @CRLF _
;~                 & '	  <inc:get>' & @CRLF _
;~                 & '		<sys_id>7c4be3e19784ad54e35ffdf3a253aff9</sys_id>' & @CRLF _
;~                 & '	  </inc:get>' & @CRLF _
;~                 & '   </soapenv:Body>' & @CRLF _
;~                 & '</soapenv:Envelope>'




;
; Toast Mod
;
;
; #FUNCTION# =========================================================================================================
; Name...........: _Toast_Show
; Description ...: Shows a slice message from the systray
; Syntax.........: _Toast_ShowMod($vIcon, $sTitle, $sMessage, [$iDelay [, $fWait [, $fRaw ]]])
; Parameters ....: $vIcon    - 0 - No icon, 8 - UAC, 16 - Stop, 32 - Query, 48 - Exclamation, 64 - Information
;                              The $MB_ICON constant can also be used for the last 4 above
;                              If set to the name of an ico or exe file, the main icon within will be displayed
;                                  If another icon from the file is required, add a trailing "|" followed by the icon index
;                              If set to the name of an image file, that image will be displayed
;                              Any other value returns -1, error 1
;                  $sTitle   - Text to display on Title bar
;                  $sMessage - Text to display in Toast body
;                  $iDelay   - The delay in seconds before the Toast retracts or script continues (Default = 0)
;                              If negative, an [X] is added to the title bar. Clicking [X] retracts/continues immediately
;                  $fWait    - True  - Script waits for delay time before continuing and Toast remains visible
;                              False - Script continues and Toast retracts automatically after delay time
;                  $fRaw     - True  - Message is not wrapped and Toast expands to show full width
;                            - False - Message is wrapped if over max preset Toast width
; Requirement(s).: v3.3.1.5 or higher - AdlibRegister/Unregister used in _Toast_Show
; Return values .: Success:      Returns 3-element array: [Toast width, Toast height, Text line height]
;                  Intermediate: Returns 0 New Toast with previous Toast retracting
;                  Failure:	     Returns -1 and sets @error as follows:
;                                       1 = Toast GUI creation failed
;                                       2 = Taskbar not found
;                                       3 = StringSize error
;                                       4 = When using Raw, the Toast is too wide for the display
; Author ........: Melba23, based on some original code by GioVit for the Toast
; Notes .........; Any visible Toast is retracted by a subsequent _Toast_Hide or _Toast_Show, or clicking a visible [X].
;                  If previous Toast is retracting then new Toast creation delayed until retraction complete
; Example........; Yes
;=====================================================================================================================
Func _Toast_ShowMod($vIcon, $sTitle, $sMessage, $iDelay = 0, $fWait = True, $bisTicket = False, $fRaw = False)
	$fToast_OpenTik = False
	; If previous Toast retracting must wait until process is completed
	If $fToast_Retracting Then
		; Store parameters
		$vIcon_Retraction = $vIcon
		$sTitle_Retraction = $sTitle
		$sMessage_Retraction = $sMessage
		$iDelay_Retraction = $iDelay
		$fWait_Retraction = $fWait
		$fRaw_Retraction = $fRaw
		; Keep looking to see if previous Toast retracted
		AdlibRegister("__Toast_Retraction_Check", 100)
		; Explain situation to user
		Return SetError(5, 0, -1)
	EndIf

	; Store current GUI mode and set Message mode
	Local $nOldOpt = Opt('GUIOnEventMode', 0)

	; Retract any Toast already in place
	If $hToast_Handle <> 0 Then _Toast_Hide()

	; Reset non-reacting Close [X] ControlID
	$hToast_Close_X = 9999

	; Set default auto-sizing Toast widths
	Local $iToast_Width_max = 500
	Local $iToast_Width_min = 150

	; Check for icon
	Local $iIcon_Style = 0
	Local $iIcon_Reduction = 36
	Local $sDLL = "user32.dll"
	Local $sImg = ""
	If StringIsDigit($vIcon) Then
		Switch $vIcon
			Case 0
				$iIcon_Reduction = 0
			Case 8
				$sDLL = "imageres.dll"
				$iIcon_Style = 78
			Case 16 ; Stop
				$iIcon_Style = -4
			Case 32 ; Query
				$iIcon_Style = -3
			Case 48 ; Exclam
				$iIcon_Style = -2
			Case 64 ; Info
				$iIcon_Style = -5
			Case Else
				Return SetError(1, 0, -1)
		EndSwitch
	Else
		If StringInStr($vIcon, "|") Then
			$iIcon_Style = StringRegExpReplace($vIcon, "(.*)\|", "")
			$sDLL = StringRegExpReplace($vIcon, "\|.*$", "")
		Else
			Switch StringLower(StringRight($vIcon, 3))
				Case "exe", "ico"
					$sDLL = $vIcon
				Case "bmp", "jpg", "gif", "png"
					$sImg = $vIcon
			EndSwitch
		EndIf
	EndIf

	; Determine max message width
	Local $iMax_Label_Width = $iToast_Width_max - 20 - $iIcon_Reduction
	If $fRaw = True Then $iMax_Label_Width = 0

	; Get message label size
	Local $aLabel_Pos = _StringSize($sMessage, $iToast_Font_Size, Default, Default, $sToast_Font_Name, $iMax_Label_Width)
	If @error Then
		$nOldOpt = Opt('GUIOnEventMode', $nOldOpt)
		Return SetError(3, 0, -1)
	EndIf

	; Reset text to match rectangle
	$sMessage = $aLabel_Pos[0]

	;Set line height for this font
	Local $iLine_Height = $aLabel_Pos[1]

	; Set label size
	Local $iLabelwidth = $aLabel_Pos[2]
	Local $iLabelheight = $aLabel_Pos[3]

	; Set Toast size
	Local $iToast_Width = $iLabelwidth + 20 + $iIcon_Reduction
	; Check if Toast will fit on screen
	If $iToast_Width > @DesktopWidth - 20 Then
		$nOldOpt = Opt('GUIOnEventMode', $nOldOpt)
		Return SetError(4, 0, -1)
	EndIf
	; Increase if below min size
	If $iToast_Width < $iToast_Width_min + $iIcon_Reduction Then
		$iToast_Width = $iToast_Width_min + $iIcon_Reduction
		$iLabelwidth = $iToast_Width_min - 20
	EndIf

	; Set title bar height - with minimum for [X]
	Local $iTitle_Height = 0
	If $sTitle = "" Then
		If $iDelay < 0 Then $iTitle_Height = 6
	Else
		$iTitle_Height = $iLine_Height + 2
		If $iDelay < 0 Then
			If $iTitle_Height < 17 Then $iTitle_Height = 17
		EndIf
	EndIf

	; Set Toast height as label height + title bar + bottom margin
	Local $iToast_Height = $iLabelheight + $iTitle_Height + 20
	If $bisTicket Then $iToast_Height += $iTitle_Height

	; Ensure enough room for icon if displayed
	If $iIcon_Reduction Then
		If $iToast_Height < $iTitle_Height + 42 Then $iToast_Height = $iTitle_Height + 47
	EndIf

	; Get Toast starting position and direction
	Local $aToast_Data = __Toast_Locate($iToast_Width, $iToast_Height)

	; Create Toast slice with $WS_POPUPWINDOW, $WS_EX_TOOLWINDOW style and $WS_EX_TOPMOST extended style
	$hToast_Handle = GUICreate("", $iToast_Width, $iToast_Height, $aToast_Data[0], $aToast_Data[1], 0x80880000, BitOR(0x00000080, 0x00000008))
	If @error Then
		$nOldOpt = Opt('GUIOnEventMode', $nOldOpt)
		Return SetError(1, 0, -1)
	EndIf
	GUISetFont($iToast_Font_Size, Default, Default, $sToast_Font_Name)
	GUISetBkColor($iToast_Message_BkCol)

	; Set centring parameter
	Local $iLabel_Style = 0 ; $SS_LEFT
	If BitAND($iToast_Style, 1) = 1 Then
		$iLabel_Style = 1 ; $SS_CENTER
	ElseIf BitAND($iToast_Style, 2) = 2 Then
		$iLabel_Style = 2 ; $SS_RIGHT
	EndIf

	; Check installed fonts
	Local $sX_Font = "WingDings"
	Local $sX_Char = "x"
	Local $i = 1
	While 1
		Local $sInstalled_Font = RegEnumVal("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", $i)
		If @error Then ExitLoop
		If StringInStr($sInstalled_Font, "WingDings 2") Then
			$sX_Font = "WingDings 2"
			$sX_Char = "T"
		EndIf
		$i += 1
	WEnd

	; Create title bar if required
	If $sTitle <> "" Then

		; Create disabled background strip
		GUICtrlCreateLabel("", 0, 0, $iToast_Width, $iTitle_Height)
		GUICtrlSetBkColor(-1, $iToast_Header_BkCol)
		GUICtrlSetState(-1, 128) ; $GUI_DISABLE

		If $bisTicket Then
			GUICtrlCreateLabel("", 0, $iToast_Height - $iTitle_Height, $iToast_Width, $iTitle_Height)
			GUICtrlSetBkColor(-1, $iToast_Header_BkCol)
			GUICtrlSetState(-1, 128) ; $GUI_DISABLE
		EndIf

		; Set title bar width to offset text
		Local $iTitle_Width = $iToast_Width - 10

		; Create closure [X] if needed
		If $iDelay < 0 Then
			; Create [X]
			Local $iX_YCoord = Int(($iTitle_Height - 17) / 2)
			$hToast_Close_X = GUICtrlCreateLabel($sX_Char, $iToast_Width - 18, $iX_YCoord, 17, 17)
			GUICtrlSetFont(-1, 14, Default, Default, $sX_Font)
			GUICtrlSetBkColor(-1, -2) ; $GUI_BKCOLOR_TRANSPARENT
			GUICtrlSetColor(-1, $iToast_Header_Col)
			; Reduce title bar width to allow [X] to activate
			$iTitle_Width -= 18
		EndIf

		; Create Title label with bold text, centred vertically in case bar is higher than line
		GUICtrlCreateLabel($sTitle, 10, 0, $iTitle_Width, $iTitle_Height, 0x0200) ; $SS_CENTERIMAGE
		GUICtrlSetBkColor(-1, $iToast_Header_BkCol)
		GUICtrlSetColor(-1, $iToast_Header_Col)
		If BitAND($iToast_Style, 4) = 4 Then GUICtrlSetFont(-1, $iToast_Font_Size, 600)

	Else

		If $iDelay < 0 Then
			; Only need [X]
			$hToast_Close_X = GUICtrlCreateLabel($sX_Char, $iToast_Width - 18, 0, 17, 17)
			GUICtrlSetFont(-1, 14, Default, Default, $sX_Font)
			GUICtrlSetBkColor(-1, -2) ; $GUI_BKCOLOR_TRANSPARENT
			GUICtrlSetColor(-1, $iToast_Message_Col)
		EndIf

	EndIf

	; Create icon
	If $iIcon_Reduction Then
		Switch StringLower(StringRight($sImg, 3))
			Case "bmp", "jpg", "gif"
				GUICtrlCreatePic($sImg, 10, 10 + $iTitle_Height, 32, 32)
			Case "png"
				__Toast_ShowPNG($sImg, $iTitle_Height)
			Case Else
				GUICtrlCreateIcon($sDLL, $iIcon_Style, 10, 10 + $iTitle_Height)
		EndSwitch
	EndIf

	; Create Message label
	GUICtrlCreateLabel($sMessage, 10 + $iIcon_Reduction, 10 + $iTitle_Height, $iLabelwidth, $iLabelheight)
	GUICtrlSetStyle(-1, $iLabel_Style)
	If $iToast_Message_Col <> Default Then GUICtrlSetColor(-1, $iToast_Message_Col)

	$hToast_OpenTik = Null
	If $bisTicket Then
		Local $sBtn = "Open Ticket"
		Local $aBtn_Pos = _StringSize($sBtn, $iToast_Font_Size, Default, Default, $sToast_Font_Name, $iToast_Width)
		$hToast_OpenTik = GUICtrlCreateLabel($sBtn, ($iToast_Width / 2) - ($aBtn_Pos[2] / 2), $iToast_Height - $iTitle_Height, $aBtn_Pos[2], $iTitle_Height)
		;GUICtrlSetFont(-1, 14, Default, Default, $sX_Font)
		GUICtrlSetBkColor(-1, -2) ; $GUI_BKCOLOR_TRANSPARENT
		GUICtrlSetColor(-1, $iToast_Header_Col)
	EndIf

	; Slide Toast Slice into view from behind systray and activate
	DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $hToast_Handle, "int", $iToast_Time_Out, "long", $aToast_Data[2])

	; Activate Toast without stealing focus
	GUISetState(@SW_SHOWNOACTIVATE, $hToast_Handle)

	; If script is to pause
	If $fWait = True Then

		; Clear message queue
		Do
		Until GUIGetMsg() = 0

		; Begin timeout counter
		Local $iTimeout_Begin = TimerInit()

		; Wait for timeout or closure
		Local $iMsg
		While Sleep(10)
			$iMsg = GUIGetMsg()
			If $iMsg == $hToast_Close_X Or TimerDiff($iTimeout_Begin) / 1000 >= Abs($iDelay) Then
				ExitLoop
			ElseIf $iMsg == $hToast_OpenTik Then
				$fToast_OpenTik = True
				ExitLoop
			EndIf
		WEnd

		; If script is to continue and delay has been set
	ElseIf Abs($iDelay) > 0 Then

		; Store timer info
		$iToast_Timer = Abs($iDelay * 1000)
		$iToast_Start = TimerInit()

		; Register Adlib function to run timer
		AdlibRegister("__Toast_Timer_Check", 100)
		; Register message handler to check for [X] click
		GUIRegisterMsg(0x0021, "__Toast_WM_EVENTSMod") ; $WM_MOUSEACTIVATE

	EndIf

	; Reset original mode
	$nOldOpt = Opt('GUIOnEventMode', $nOldOpt)

	; Create array to return Toast dimensions
	Local $aToast_Data[3] = [$iToast_Width, $iToast_Height, $iLine_Height]

	Return $aToast_Data

EndFunc   ;==>_Toast_ShowMod

Func __Toast_WM_EVENTSMod($hWnd, $Msg, $wParam, $lParam)
	#forceref $wParam, $lParam
	If $hWnd = $hToast_Handle Then
		If $Msg = 0x0021 Then ; $WM_MOUSEACTIVATE
			; Check mouse position
			Local $aPos = GUIGetCursorInfo($hToast_Handle)
			If $aPos[4] = $hToast_Close_X Then $fToast_Close = True
			If $aPos[4] = $hToast_OpenTik Then
				$fToast_OpenTik = True
				$fToast_Close = True
			EndIf
		EndIf
	EndIf
	Return 'GUI_RUNDEFMSG'
EndFunc   ;==>__Toast_WM_EVENTSMod

Func _AddWatch()
    Local $aTicket[$aFields[0]]
    Local $iHeight = 72
    Local $iWidth = 256+64
    Local $iMargin = 8
    Local $sRead
    Local $aoQuery
    $g_AddWatch_hWnd = GUICreate("ohNotify - Add to Watch", $iWidth, $iHeight)
    GUISetFont(10,400,"Normal","COnsolas",$g_AddWatch_hWnd)
    $g_AddWatch_idInput = GUICtrlCreateInput("", $iMargin, $iMargin, $iWidth-($iMargin*2), 20)
    _GUICtrlEdit_SetCueBanner(GUICtrlGetHandle($g_AddWatch_idInput),"Enter an SCTASK/INC...",True)
    Local $idCancel = GUICtrlCreateButton("Cancel", ($iWidth/2)-$iMargin-72, 40, 72, 24)
    $g_AddWatch_idAdd = GUICtrlCreateButton("Add", ($iWidth/2)+$iMargin, 40, 72, 24)
    GUICtrlSetState($g_AddWatch_idAdd,$GUI_DISABLE)
    $g_bEnAdd = False
    GUIRegisterMsg($WM_COMMAND,"_WM_COMMAND_AddWatch")
    GUISetState(@SW_SHOW)
    While 1
        $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $idCancel
                GUIRegisterMsg($WM_COMMAND,"")
                GUISetState(@SW_HIDE)
                GUIDelete($g_AddWatch_hWnd)
                Return SetError(0,1,0)
            Case $g_AddWatch_idAdd
                $sRead=StringLower(GUICtrlRead($g_AddWatch_idInput))
                If StringRegExp($sRead,"^inc\d{7,}$") Then
                    $aoQuery=_queryIncidents("number="&$sRead)
                ElseIf StringRegExp($sRead,"^sctask\d{7,}$") Then
                    $aoQuery=_queryTasks("number="&$sRead)
                EndIf
                If Not IsObj($aoQuery) Then
                    MsgBox(48,"ohNotify - Add to Watch","Warning: Invalid incident number, query returned no results.")
                    ContinueLoop
                EndIf
                If IsArray($aoQuery) Then
                    MsgBox(48,"ohNotify - Add to Watch","Warning: Query returned too many results.")
                    ContinueLoop
                EndIf
                For $k = 1 To $aFields[0]
                    $aTicket[$k - 1] = __snSoapGetAttr($aoQuery, $aFields[$k])
                Next
                ;_ArrayDisplay($aTicket)
                Local $bExists=False
                Local $iMax=UBound($aTasksLast,1)
                For $i=1 To $iMax-1
                    If $aTasksLast[$i][8]==$aTicket[8] Then
                        $bExists=True
                        ExitLoop
                    EndIf
                Next
                If $bExists Then
                    MsgBox(48,"ohNotify - Add to Watch","Warning: "&$sRead&" Already exists.")
                    ContinueLoop
                EndIf
                ReDim $aTasksLast[$iMax+1][$aFields[0]]
                For $k = 1 To $aFields[0]
                    $aTasksLast[$iMax][$k - 1] = $aTicket[$k - 1]
                Next
                MsgBox(64,"ohNotify - Add to Watch",$sRead&" has been added to the watchlist.")
                GUIRegisterMsg($WM_COMMAND,"")
                GUISetState(@SW_SHOW)
                GUIDelete($g_AddWatch_hWnd)
                Return SetError(0,0,1)
        EndSwitch
    WEnd
EndFunc

Func _WM_COMMAND_AddWatch($hWnd, $iMsg, $wParam, $lParam)
    If $hWnd<>$g_AddWatch_hWnd Then Return $GUI_RUNDEFMSG
    Local $iCode,$inID,$bMod=False,$sRead
    If BitAND($wParam, 0xFFFF)<>$g_AddWatch_idInput Or BitShift($wParam, 16)<>$EN_CHANGE Then Return $GUI_RUNDEFMSG
    $sRead=StringLower(GUICtrlRead($g_AddWatch_idInput))
    If $sRead=="" Then
        GuiCtrlSetState($g_AddWatch_idAdd,$GUI_DISABLE)
        Return $GUI_RUNDEFMSG
    EndIf
    If Not StringRegExp($sRead,"^(?:inc|sctask)\d{7,}$") Then
        GuiCtrlSetState($g_AddWatch_idAdd,$GUI_DISABLE)
        Return $GUI_RUNDEFMSG
    EndIf
    GuiCtrlSetState($g_AddWatch_idAdd,$GUI_ENABLE)
    Return $GUI_RUNDEFMSG
EndFunc

;
; Generate Random 16 digit Alphanumeric String
; UEZ, modified by Biatu
;
Func _RandStr()
    Local $sRet = "", $aTmp[3], $iLen = 16
    For $i = 1 To $iLen
        $aTmp[0] = Chr(Random(65, 90, 1)) ;A-Z
        $aTmp[1] = Chr(Random(97, 122, 1)) ;a-z
        $aTmp[2] = Chr(Random(48, 57, 1)) ;0-9
        $sRet &= $aTmp[Random(0, 2, 1)]
    Next
    Return $sRet
EndFunc
