#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"

Func Metronome_tab_setup()

	GUICtrlCreateTabItemEx("Metronome")
;	GUICtrlCreateGroupEx  ("Harvest", 20, 300, 780, 160)
;	$harvest_account_id_input = 												GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "HarvestAccountID", ""), 120, 320, 660, 20, $harvest_accound_id_label, "Account ID", 30, 320, 100, 20)
;	$harvest_access_token_input = 												GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "HarvestAccessToken", ""), 120, 340, 660, 20, $harvest_access_token_label, "Access Token", 30, 340, 100, 20)
;	$settings_save_button = 													GUICtrlCreateImageButton("save.ico", 30, 80, 36, "Save these settings", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

EndFunc

Func Metronome_tab_child_gui_setup()
EndFunc


Func Metronome_tab_event_handler($msg)

	Switch $msg

		Case $settings_save_button

			depress_button_and_disable_gui($msg, -1, 100)
			GUICtrlStatusInput_SetText($status_input, "Saving Settings ...")
			IniWrite($ini_filename, "Global", "HarvestAccountID", GUICtrlRead($harvest_account_id_input))
			IniWrite($ini_filename, "Global", "HarvestAccessToken", GUICtrlRead($harvest_access_token_input))
			GUICtrlStatusInput_SetText($status_input, "")
			raise_button_and_enable_gui($msg)

	EndSwitch

EndFunc


Func Metronome_tab_WM_NOTIFY_handler($hWndFrom, $iCode)

;	Switch $hWndFrom


;		Case GUICtrlGetHandle($image_compression_quality_slider)

;			Switch $iCode
;				Case $NM_RELEASEDCAPTURE ; The control is releasing mouse capture


;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $NM_RELEASEDCAPTURE = ' & $NM_RELEASEDCAPTURE & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			EndSwitch


;	EndSwitch

EndFunc

