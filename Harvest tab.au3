#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"
#Include "JSON.au3"
#include <Date.au3>

Global $update_tasks = False

Func Harvest_tab_setup()

	GUICtrlCreateTabItemEx("Harvest")
	;GUICtrlCreateGroupEx  ("", 20, 140, 250, 500)
	$timesheet_listview = 														GUICtrlCreateListViewEx(30, 230, 760, 360, "Date", 90, "Project", 160, "Task", 160, "Notes", 160, "Hours", 160)
	_GUICtrlListView_EnableGroupView($timesheet_listview)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 1, "Monday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 2, "Tuesday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 3, "Wednesday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 4, "Thursday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 5, "Friday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 6, "Saturday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 7, "Sunday")
	$timesheet_refresh_button = 												GUICtrlCreateImageButton("refresh.ico", 30, 80, 36, "Get your Harvest times", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$timesheet_add_button = 													GUICtrlCreateImageButton("add.ico", 70, 80, 28, "Add a new Time Entry")
	$timesheet_delete_button = 													GUICtrlCreateImageButton("delete.ico", 110, 80, 28, "Delete the selected Time Entry")

EndFunc

Func Harvest_tab_child_gui_setup()

	$add_time_entry_gui = 														ChildGUICreate($app_name & " - Add Time Entry", 640, 480, $main_gui)
;	GUICtrlCreateGroupEx  ("----> RetroPie (/boot/config.txt)", 5, 5, 180, 40)
	$add_time_entry_project_listview = 											GUICtrlCreateListViewEx(10, 10, 550, 100, "Project", 160, "ID", 160)
	$add_time_entry_task_listview = 											GUICtrlCreateListViewEx(10, 130, 550, 100, "Task", 160, "ID", 160)
	$add_time_entry_add_button = 												GUICtrlCreateButton("Load", 10, 20, 80, 20)
	$add_time_entry_cancel_button = 											GUICtrlCreateButton("Save", 100, 20, 80, 20)
;	GUICtrlCreateGroupEx  ("----> PC", 200, 5, 180, 40)
;	$boot_config_open_button = 													GUICtrlCreateButton("Open", 205, 20, 80, 20)
;	$boot_config_save_as_button = 												GUICtrlCreateButton("Save As", 295, 20, 80, 20)
;	$boot_config_edit = 														GUICtrlCreateEdit("", 10, 50, 620, 400)
;	$boot_config_status_input = 												GUICtrlCreateInput("", 10, 480 - 25, 640 - 20, 20, $ES_READONLY, $WS_EX_STATICEDGE)
EndFunc


Func Harvest_tab_event_handler($msg)

	Switch $msg

		Case $GUI_EVENT_CLOSE

			if $current_gui = $add_time_entry_gui Then

				GUISetState(@SW_ENABLE, $main_gui)
				GUISetState(@SW_HIDE, $current_gui)
				$current_gui = $main_gui
				raise_button_and_enable_gui($timesheet_add_button)
			EndIf



		case $timesheet_refresh_button

			depress_button_and_disable_gui($msg, -1, 100)

			for $i = 1 to 3

				Local $iPID = Run('curl -k https://api.harvestapp.com/v2/users/me/project_assignments?page=' & $i & '&per_page=100 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
				ProcessWaitClose($iPID)
				Local $json = StdoutRead($iPID)
				Local $decoded_json = Json_Decode($json)

				for $i = 0 to 99

					Local $project_name = Json_Get($decoded_json, '.project_assignments[' & $i & '].project.name')

					if StringLen($project_name) > 0 Then

						Local $task_names = ""

						for $j = 0 to 99

							Local $task_name = Json_Get($decoded_json, '.project_assignments[' & $i & '].task_assignments[' & $j & '].task.name')

							if StringLen($task_name) < 1 Then ExitLoop

							$task_name = $task_name & "|"

							if StringLen($task_names) > 0 Then $task_names = $task_names & @CRLF

							$task_names = $task_names & $task_name
						Next

						$timesheet_project_assignments_dict.Add($project_name, $task_names)
					EndIf
				Next
			Next



			GUICtrlStatusInput_SetText($status_input, "Getting your Harvest times ...")
;			Local $iPID = Run('curl -k "https://api.harvestapp.com/v2/time_entries" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)" -d "{\"from\":\"2021-07-19\"}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?from="2021-07-19"&to="2021-07-20" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?page=1&per_page=10&project_id=25655657&from=2021-06-30 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?project_id=25655657&from=2021-06-30 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?from=2021-07-12&to=2021-07-18 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			GUICtrlStatusInput_SetText($status_input, "")

			Local $decoded_json = Json_Decode($json)
			_GUICtrlListView_DeleteAllItems($timesheet_listview)
			_GUICtrlListView_BeginUpdate($timesheet_listview)

			for $i = 99 to 0 step -1

				Local $spent_date = Json_Get($decoded_json, '.time_entries[' & $i & '].spent_date')

				if StringLen($spent_date) > 0 Then

					Local $spent_date_day_to_week_index = _DateToDayOfWeek(StringLeft($spent_date, 4), StringMid($spent_date, 6, 2), StringRight($spent_date, 2))
					$spent_date = _DateDayOfWeek($spent_date_day_to_week_index)
					Local $project = Json_Get($decoded_json, '.time_entries[' & $i & '].project.name')
					Local $task = Json_Get($decoded_json, '.time_entries[' & $i & '].task.name')
					Local $notes = Json_Get($decoded_json, '.time_entries[' & $i & '].notes')
					Local $hours = Json_Get($decoded_json, '.time_entries[' & $i & '].hours')

					Local $index = _GUICtrlListView_AddItem($timesheet_listview, $spent_date)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $project, 1)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $task, 2)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $notes, 3)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $hours, 4)
					_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, $spent_date_day_to_week_index - 1)
				EndIf
			Next

			_GUICtrlListView_EndUpdate($timesheet_listview)
			raise_button_and_enable_gui($msg)

		Case $timesheet_add_button

			depress_button_and_disable_gui($msg)
			GUISetState(@SW_DISABLE, $main_gui)
			GUISetState(@SW_SHOW, $add_time_entry_gui)
			$current_gui = $add_time_entry_gui

			_GUICtrlListView_DeleteAllItems($add_time_entry_project_listview)
			_GUICtrlListView_BeginUpdate($add_time_entry_project_listview)

			For $vKey In $timesheet_project_assignments_dict

				Local $index = _GUICtrlListView_AddItem($add_time_entry_project_listview, $vKey)

;				ConsoleWrite($vKey & " - " & $timesheet_project_assignments_dict.Item($vKey) & @CRLF)
			Next

			_GUICtrlListView_EndUpdate($add_time_entry_project_listview)

	EndSwitch

	if $update_tasks = True Then

		$update_tasks = False
		Local $selected_project = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)

		_GUICtrlListView_DeleteAllItems($add_time_entry_task_listview)
		_GUICtrlListView_BeginUpdate($add_time_entry_task_listview)

		Local $task_names = $timesheet_project_assignments_dict.Item($selected_project)

		;Local $task_name_arr = StringSplit($task_names, "|")
		Local $task_name_arr = _StringSplit2d($task_names, "|")
		_GUICtrlListView_AddArray($add_time_entry_task_listview, $task_name_arr)


	EndIf
EndFunc


Func Harvest_tab_WM_NOTIFY_handler($hWndFrom, $iCode)


	Switch $hWndFrom


		Case GUICtrlGetHandle($add_time_entry_project_listview)

			Switch $iCode

				Case $LVN_ITEMCHANGED

					$update_tasks = True

			EndSwitch


	EndSwitch


EndFunc
