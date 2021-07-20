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

;$LVN_ITEMACTIVATE
;$LVN_ITEMCHANGED
;$LVN_ITEMCHANGING
;$NM_CLICK
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $NM_CLICK = ' & $NM_CLICK & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $LVN_ITEMCHANGING = ' & $LVN_ITEMCHANGING & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $LVN_ITEMCHANGED = ' & $LVN_ITEMCHANGED & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $LVN_ITEMACTIVATE = ' & $LVN_ITEMACTIVATE & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

HJM_Link_Startup()

; Main gui

$main_gui = 																	MainGUICreate($tab, 5, 50, 840-10, 720-80, $GUI_DOCKVCENTER + $GUI_DOCKBORDERS)
$status_input = 																GUICtrlCreateStatusInput("Hint - hover mouse over controls for help", 5, 720 - 25, 830, 20)

; Main gui tabs

Settings_tab_setup()
Harvest_tab_setup()
Jira_tab_setup()
Metronome_tab_setup()


$current_gui = $main_gui
GUISetState(@SW_SHOW, $main_gui)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
_TipDisplayLen(30000)


While True

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

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

;	Switch $hWndFrom

;		Case GUICtrlGetHandle($tab)

;			Switch $iCode

;				Case $NM_CLICK


;			EndSwitch

;	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY



Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom = $lParam
    Local $iCode = BitShift($wParam, 16) ; Hi Word

	; TAB specific WM_COMMAND handlers ...

;	Scrape_Images_with_Auto_Join_tab_WM_COMMAND_handler($hWndFrom, $iCode)
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

