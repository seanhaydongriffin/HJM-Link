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
;GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
;GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
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


WEnd

HJM_Link_Shutdown()

