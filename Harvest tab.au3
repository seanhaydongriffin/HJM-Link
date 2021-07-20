#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"
#Include "JSON.au3"


Func Harvest_tab_setup()

	GUICtrlCreateTabItemEx("Harvest")
	;GUICtrlCreateGroupEx  ("", 20, 140, 250, 500)
	$timesheet_listview = 														GUICtrlCreateListViewEx(30, 230, 760, 360, "Date", 90, "Project", 160, "Task", 160, "Notes", 160, "Hours", 160)
	$timesheet_refresh_button = 												GUICtrlCreateImageButton("refresh.ico", 30, 80, 36, "Get your Harvest times", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

EndFunc

Func Harvest_tab_child_gui_setup()
EndFunc



Func Harvest_tab_event_handler($msg)

	Switch $msg



		case $timesheet_refresh_button

			depress_button_and_disable_gui($msg, -1, 100)

			GUICtrlStatusInput_SetText($status_input, "Getting your Harvest times ...")
			Local $iPID = Run('curl -k "https://api.harvestapp.com/v2/time_entries" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			GUICtrlStatusInput_SetText($status_input, "")

			Local $decoded_json = Json_Decode($json)
			_GUICtrlListView_DeleteAllItems($timesheet_listview)
			_GUICtrlListView_BeginUpdate($timesheet_listview)

			for $i = 0 to 99

				Local $spent_date = Json_Get($decoded_json, '.time_entries[' & $i & '].spent_date')
				Local $project = Json_Get($decoded_json, '.time_entries[' & $i & '].project.name')
				Local $task = Json_Get($decoded_json, '.time_entries[' & $i & '].task.name')
				Local $notes = Json_Get($decoded_json, '.time_entries[' & $i & '].notes')
				Local $hours = Json_Get($decoded_json, '.time_entries[' & $i & '].hours')

				Local $index = _GUICtrlListView_AddItem($timesheet_listview, $spent_date)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $project, 1)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $task, 2)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $notes, 3)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $hours, 4)
			Next

			_GUICtrlListView_EndUpdate($timesheet_listview)
			raise_button_and_enable_gui($msg)


	EndSwitch

EndFunc


Func Harvest_tab_WM_NOTIFY_handler($hWndFrom, $iCode)

;	Switch $hWndFrom


;		Case GUICtrlGetHandle($image_compression_quality_slider)

;			Switch $iCode
;				Case $NM_RELEASEDCAPTURE ; The control is releasing mouse capture


;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $NM_RELEASEDCAPTURE = ' & $NM_RELEASEDCAPTURE & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			EndSwitch


;	EndSwitch

EndFunc
