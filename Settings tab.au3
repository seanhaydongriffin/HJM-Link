#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"
#include <Crypt.au3>

Func Settings_tab_setup()

	GUICtrlCreateTabItemEx("Settings")
	$settings_save_button = 													GUICtrlCreateImageButton("save.ico", 30, 80, 36, "Save these settings", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlCreateGroupEx  ("Harvest", 20, 150, 780, 100)
	$harvest_account_id_input = 												GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "HarvestAccountID", ""), 120, 170, 660, 20, $harvest_accound_id_label, "Account ID", 30, 170, 100, 20)
	$harvest_access_token_input = 												GUICtrlCreatePasswordWithLabel("", 120, 190, 660, 20, $harvest_access_token_label, "Access Token", 30, 190, 100, 20)
	GUICtrlCreateGroupEx  ("Jira", 20, 300, 780, 160)
	$jira_username_input = 														GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "JiraUsername", ""), 120, 320, 660, 20, $jira_username_label, "Account ID", 30, 320, 100, 20)
	$jira_api_token_input = 													GUICtrlCreatePasswordWithLabel("", 120, 340, 660, 20, $jira_api_token_label, "Jira API Token", 30, 340, 100, 20)
	GUICtrlCreateGroupEx  ("Metronome", 20, 500, 780, 160)
	$metronome_email_input = 													GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "MetronomeEmail", ""), 120, 520, 660, 20, $harvest_accound_id_label, "Email", 30, 520, 100, 20)
	$metronome_password_input = 												GUICtrlCreatePasswordWithLabel("", 120, 540, 660, 20, $harvest_access_token_label, "Password", 30, 540, 100, 20)

	Global $harvest_access_token_encrypted = IniRead($ini_filename, "Global", "HarvestAccessToken", "")
	Global $harvest_access_token_decrypted = _Crypt_DecryptData($harvest_access_token_encrypted, "hotdog", $CALG_AES_256)
	$harvest_access_token_decrypted = BinaryToString($harvest_access_token_decrypted)
	GUICtrlSetData($harvest_access_token_input, $harvest_access_token_decrypted)

	Global $jira_api_token_encrypted = IniRead($ini_filename, "Global", "JiraAPIToken", "")
	Global $jira_api_token_decrypted = _Crypt_DecryptData($jira_api_token_encrypted, "hotdog", $CALG_AES_256)
	$jira_api_token_decrypted = BinaryToString($jira_api_token_decrypted)
	GUICtrlSetData($jira_api_token_input, $jira_api_token_decrypted)

	Global $metronome_password_encrypted = IniRead($ini_filename, "Global", "MetronomePassword", "")
	Global $metronome_password_decrypted = _Crypt_DecryptData($metronome_password_encrypted, "hotdog", $CALG_AES_256)
	$metronome_password_decrypted = BinaryToString($metronome_password_decrypted)
	GUICtrlSetData($metronome_password_input, $metronome_password_decrypted)


EndFunc

Func Settings_tab_child_gui_setup()
EndFunc


Func Settings_tab_event_handler($msg)

	Switch $msg

		Case $settings_save_button

			depress_button_and_disable_gui($msg, -1, 100)
			GUICtrlStatusInput_SetText($status_input, "Saving Settings ...")

			$harvest_access_token_encrypted = _Crypt_EncryptData(GUICtrlRead($harvest_access_token_input), "hotdog", $CALG_AES_256)
			$jira_api_token_encrypted = _Crypt_EncryptData(GUICtrlRead($jira_api_token_input), "hotdog", $CALG_AES_256)
			$metronome_password_encrypted = _Crypt_EncryptData(GUICtrlRead($metronome_password_input), "hotdog", $CALG_AES_256)

			IniWrite($ini_filename, "Global", "HarvestAccountID", GUICtrlRead($harvest_account_id_input))
			IniWrite($ini_filename, "Global", "HarvestAccessToken", $harvest_access_token_encrypted)
			IniWrite($ini_filename, "Global", "JiraUsername", GUICtrlRead($jira_username_input))
			IniWrite($ini_filename, "Global", "JiraAPIToken", $jira_api_token_encrypted)
			IniWrite($ini_filename, "Global", "MetronomeEmail", GUICtrlRead($metronome_email_input))
			IniWrite($ini_filename, "Global", "MetronomePassword", $metronome_password_encrypted)
			GUICtrlStatusInput_SetText($status_input, "")
			raise_button_and_enable_gui($msg)

	EndSwitch

EndFunc


Func Settings_tab_WM_NOTIFY_handler($hWndFrom, $iCode)

;	Switch $hWndFrom


;		Case GUICtrlGetHandle($image_compression_quality_slider)

;			Switch $iCode
;				Case $NM_RELEASEDCAPTURE ; The control is releasing mouse capture


;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $NM_RELEASEDCAPTURE = ' & $NM_RELEASEDCAPTURE & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			EndSwitch


;	EndSwitch

EndFunc

