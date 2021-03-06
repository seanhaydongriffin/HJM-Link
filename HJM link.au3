#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#include <AutoItConstants.au3>
#Include "Json.au3"
#Include "Settings tab.au3"
#Include "Harvest tab.au3"
#Include "Jira tab.au3"
#Include "Metronome tab.au3"


;$t = "1w"

;$x = StringSplit($t, " ")
;Local $estimate_seconds = 0

;for $i = 1 to $x[0]

;	$arr = StringRegExp($x[$i], "(\d)w", 1)

;	if @error = 0 Then

;		$estimate_seconds = $estimate_seconds + (1 * 5 * 8 * 60 * 60)
;	EndIf
;Next


;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $estimate_seconds = ' & $estimate_seconds & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;Exit


;$r = "0:20"

;$r = HourAndMinutesToHours($r)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

;$r = "0:20"

;$r = HourAndMinutesToHours($r)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;$r = HoursToHourAndMinutes($r)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;$r = HourAndMinutesToHours($r)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;$r = HoursToHourAndMinutes($r)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;Exit

HJM_Link_Startup()

; Main gui

$main_gui = 																	MainGUICreate($tab, 5, 50, 840-10, 720-80)
$status_input = 																GUICtrlCreateStatusInput("Hint - hover mouse over controls for help", 5, 720 - 25, 830, 20)

; Main gui tabs

Settings_tab_setup()
Harvest_tab_setup()
Jira_tab_setup()
Metronome_tab_setup()

; Child guis

Harvest_tab_child_gui_setup()

$current_gui = $main_gui
GUISetState(@SW_SHOW, $main_gui)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_SIZING, "WM_SIZING")
_TipDisplayLen(30000)
_GUICtrlTab_SetCurFocus($tab, Number(IniRead($ini_filename, "Global", "Tab", 1)))

Harvest_tab_event_handler($timesheet_refresh_button)

While True

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_RESIZED

			if $current_gui = $main_gui Then

				; below is a workaround for docking rich edit controls ($status_input)
				$aSize = WinGetClientSize($main_gui)
				_WinAPI_SetWindowPos($status_input, $HWND_TOP, 5, $aSize[1] - 25, $aSize[0] - 10, 20, $SWP_SHOWWINDOW)
			EndIf

		Case $GUI_EVENT_CLOSE

			if $current_gui = $main_gui Then

				GUISetState(@SW_ENABLE, $main_gui)
				GUISetState(@SW_HIDE, $current_gui)
				GUIDelete($current_gui)
				ExitLoop
			EndIf

	EndSwitch

	Settings_tab_event_handler($msg)
	Harvest_tab_event_handler($msg)
	Jira_tab_event_handler($msg)
	Metronome_tab_event_handler($msg)

WEnd

HJM_Link_Shutdown()




Func WM_SIZING($hWnd, $iMsg, $iwParam, $ilParam)
    #forceref $hWnd, $iMsg, $iwParam, $ilParam

	; TAB specific WM_COMMAND handlers ...

	Harvest_tab_WM_SIZING_handler()

	; below is a workaround for docking rich edit controls ($status_input)
	$aSize = WinGetClientSize($main_gui)
	_WinAPI_SetWindowPos($status_input, $HWND_TOP, 5, $aSize[1] - 25, $aSize[0] - 10, 20, $SWP_SHOWWINDOW)

    ; A pointer to a RECT structure with the screen coordinates of the drag rectangle
    ; To change the size or position of the drag rectangle, an application must change the members of this structure.
;    Local $tRECT = DllStructCreate("int;int;int;int", $ilParam) ; $tagRECT
;    Local $iLeft, $iTop, $iRight, $iBottom
;    $iLeft = DllStructGetData($tRECT, 1)
;    $iTop = DllStructGetData($tRECT, 2)
;    $iRight = DllStructGetData($tRECT, 3)
;    $iBottom = DllStructGetData($tRECT, 4)

;~     ;  Uncomment this line have the window stretched to desktop width
;~     DllStructSetData($tRECT, 1, 0) ;left
;~     DllStructSetData($tRECT, 3, @DesktopWidth) ;right

;    $tRECT = 0

    Return $GUI_RUNDEFMSG
EndFunc


Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam

	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	Local $iCode = DllStructGetData($tNMHDR, "Code")

	; TAB specific WM_NOTIFY handlers ...

	Settings_tab_WM_NOTIFY_handler($hWndFrom, $iCode)
	Harvest_tab_WM_NOTIFY_handler($hWndFrom, $iCode)
	Jira_tab_WM_NOTIFY_handler($hWndFrom, $iCode)
	Metronome_tab_WM_NOTIFY_handler($hWndFrom, $iCode)

	; Global WM_NOTIFY handler ...

	Switch $hWndFrom

		Case GUICtrlGetHandle($tab)

			Switch $iCode

				Case $NM_CLICK

;					if _GUICtrlTab_GetCurSel($tab) = 1 Then

;						GUICtrlSetState($scrape_auto_join_upload_button, $GUI_DEFBUTTON)
;					EndIf

;					if _GUICtrlTab_GetCurSel($tab) = 2 Then

;						GUICtrlSetState($scrape_manual_join_upload_button, $GUI_DEFBUTTON)
;					EndIf

					IniWrite($ini_filename, "Global", "Tab", _GUICtrlTab_GetCurSel($tab))

			EndSwitch

	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY



Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom = $lParam
    Local $iCode = BitShift($wParam, 16) ; Hi Word

	; TAB specific WM_COMMAND handlers ...

	Harvest_tab_WM_COMMAND_handler($hWndFrom, $iCode)
;	Scrape_Metadata_tab_WM_COMMAND_handler($hWndFrom, $iCode)
;	Scrape_Images_with_Manual_Join_tab_WM_COMMAND_handler($hWndFrom, $iCode)

	; Global WM_COMMAND handler ...

;    Switch $hWndFrom

 ;       Case GUICtrlGetHandle($system_combo)

;			Switch $iCode

 ;               Case $CBN_SELCHANGE



;			EndSwitch



;    EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

