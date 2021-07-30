#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"
#Include "JSON.au3"

Func Jira_tab_setup()

	GUICtrlCreateTabItemEx("Jira")
;	GUICtrlCreateGroupEx  ("Harvest", 20, 300, 780, 160)
;	$harvest_account_id_input = 												GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "HarvestAccountID", ""), 120, 320, 660, 20, $harvest_accound_id_label, "Account ID", 30, 320, 100, 20)
;	$harvest_access_token_input = 												GUICtrlCreateInputWithLabel(IniRead($ini_filename, "Global", "HarvestAccessToken", ""), 120, 340, 660, 20, $harvest_access_token_label, "Access Token", 30, 340, 100, 20)
	$jira_tmp_button = 															GUICtrlCreateImageButton("save.ico", 30, 80, 36, "get work logs", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$jira_tmp2_button = 														GUICtrlCreateImageButton("save.ico", 30, 120, 36, "add work log", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$jira_tmp3_button = 														GUICtrlCreateImageButton("save.ico", 30, 160, 36, "edit estimate", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

EndFunc

Func Jira_tab_child_gui_setup()
EndFunc


Func Jira_tab_event_handler($msg)

	Switch $msg

		case $jira_tmp_button

			; get work logs for jira ticket

;			$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/QA-3285/worklog -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			ProcessWaitClose($iPID)
;			Local $json = StdoutRead($iPID)
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			Local $decoded_json = Json_Decode($json)


			$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/QA-3285 -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $decoded_json = Json_Decode($json)

			$original_estimate = Json_Get($decoded_json, '.fields.timetracking.originalEstimate')
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $original_estimate = ' & $original_estimate & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;			$total_worklogs = Json_Get($decoded_json, '.total')
			$total_worklogs = Json_Get($decoded_json, '.fields.worklog.total')
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $total_worklogs = ' & $total_worklogs & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			for $worklog_index = 0 to ($total_worklogs - 1)

;				Local $worklog_created = Json_Get($decoded_json, '.worklogs[' & $worklog_index & '].created')
				Local $worklog_created = Json_Get($decoded_json, '.fields.worklog.worklogs[' & $worklog_index & '].created')
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $worklog_created = ' & $worklog_created & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;				Local $worklog_timespent = Json_Get($decoded_json, '.worklogs[' & $worklog_index & '].timeSpent')
				Local $worklog_timespent = Json_Get($decoded_json, '.fields.worklog.worklogs[' & $worklog_index & '].timeSpent')
				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $worklog_timespent = ' & $worklog_timespent & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Next


		case $jira_tmp2_button

			; add work log for jira ticket

			$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/QA-3285/worklog -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X POST -d "{\"timeSpentSeconds\": 12000, \"started\": \"2021-07-26T00:00:00.000+0000\"}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			;Local $decoded_json = Json_Decode($json)

		case $jira_tmp3_button

			; edit estimate for jira ticket

			$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/QA-3285 -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X PUT -d "{\"update\": {\"timetracking\": [{\"edit\": {\"originalEstimate\": \"1w 1d\"}}]}}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			;Local $decoded_json = Json_Decode($json)






	EndSwitch

EndFunc


Func Jira_tab_WM_NOTIFY_handler($hWndFrom, $iCode)

;	Switch $hWndFrom


;		Case GUICtrlGetHandle($image_compression_quality_slider)

;			Switch $iCode
;				Case $NM_RELEASEDCAPTURE ; The control is releasing mouse capture


;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $NM_RELEASEDCAPTURE = ' & $NM_RELEASEDCAPTURE & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			EndSwitch


;	EndSwitch

EndFunc

