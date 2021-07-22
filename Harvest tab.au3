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
Global $favourites_path = $app_data_dir & "\favourites.txt"
Global $project_assignments_loaded = False

Func Harvest_tab_setup()

	GUICtrlCreateTabItemEx("Harvest")
	;GUICtrlCreateGroupEx  ("", 20, 140, 250, 500)
	$timesheet_listview = 														GUICtrlCreateListViewEx(30, 230, 760, 360, "Project", 240, "Task", 160, "Notes", 260, "Hours", 50)
	_GUICtrlListView_EnableGroupView($timesheet_listview)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 1, "Monday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 2, "Tuesday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 3, "Wednesday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 4, "Thursday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 5, "Friday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 6, "Saturday")
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 7, "Sunday")
	$timesheet_refresh_button = 												GUICtrlCreateImageButton("refresh.ico", 30, 80, 36, "Get your Harvest times", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$timesheet_add_button = 													GUICtrlCreateImageButton("add.ico", 70, 80, 36, "Add a new Time Entry")
	$timesheet_delete_button = 													GUICtrlCreateImageButton("delete.ico", 110, 80, 36, "Delete the selected Time Entry")
	$timesheet_edit_button = 													GUICtrlCreateImageButton("edit.ico", 150, 80, 36, "Edit the selected Time Entry")

EndFunc

Func Harvest_tab_child_gui_setup()

	$add_time_entry_gui = 														ChildGUICreate($app_name & " - Add Time Entry", 640, 640, $main_gui)
	GUICtrlCreateGroupEx ("Project", 5, 5, 560, 180)
	$add_time_entry_project_listview = 											GUICtrlCreateListViewEx(10, 25, 550, 150, "Name", 600, "ID", 160)
	GUICtrlCreateGroupEx ("Task", 5, 195, 560, 340)
	$add_time_entry_task_listview = 											GUICtrlCreateListViewEx(10, 215, 200, 310, "Name", 200, "ID", 160)
	$add_time_entry_hour_input = 												GUICtrlCreateInput("", 10, 540, 40, 20)
    $add_time_entry_half_hour_radio =											GUICtrlCreateRadioEx("0.5", 60, 540, 40, 20, False, "half hour")
    $add_time_entry_one_hour_radio =											GUICtrlCreateRadioEx("1.0", 110, 540, 40, 20, True, "one hour")
    $add_time_entry_one_half_hour_radio =										GUICtrlCreateRadioEx("1.5", 160, 540, 40, 20, False, "one and half hour")
    $add_time_entry_two_hour_radio =											GUICtrlCreateRadioEx("2.0", 210, 540, 40, 20, False, "two hour")
    $add_time_entry_two_half_hour_radio =										GUICtrlCreateRadioEx("2.5", 260, 540, 40, 20, False, "two and half hour")
    $add_time_entry_three_hour_radio =											GUICtrlCreateRadioEx("3.0", 310, 540, 40, 20, False, "three hour")
    $add_time_entry_three_half_hour_radio =										GUICtrlCreateRadioEx("3.5", 360, 540, 40, 20, False, "three and half hour")
    $add_time_entry_four_hour_radio =											GUICtrlCreateRadioEx("4.0", 410, 540, 40, 20, False, "four hour")
    $add_time_entry_four_half_hour_radio =										GUICtrlCreateRadioEx("4.5", 460, 540, 40, 20, False, "four and half hour")
    $add_time_entry_five_hour_radio =											GUICtrlCreateRadioEx("5.0", 510, 540, 40, 20, False, "five hour")
	$add_time_entry_save_button = 												GUICtrlCreateImageButton("save.ico", 10, 640 - 70, 36, "Save this new Time Entry")
	;$add_time_entry_cancel_button = 											GUICtrlCreateImageButton("cancel.ico", 50, 640 - 70, 36, "Cancel this Time Entry")
	GUICtrlCreateGroupEx ("Favourites", 240, 210, 200, 240)
	$add_time_entry_favourites_list = 											GUICtrlCreateList("", 250, 230, 180, 170, BitOR($GUI_SS_DEFAULT_LIST, $WS_HSCROLL))
	$add_time_entry_favourites_add_button = 									GUICtrlCreateImageButton("add.ico", 250, 410, 36, "Add a new Favourite Task")
	$add_time_entry_favourites_delete_button = 									GUICtrlCreateImageButton("delete.ico", 290, 410, 36, "Delete the selected Favourite Task")
	$add_time_entry_status_input = 												GUICtrlCreateStatusInput("", 10, 640 - 25, 640 - 20, 20)
;	GUICtrlCreateGroupEx  ("----> PC", 200, 5, 180, 40)
;	$boot_config_open_button = 													GUICtrlCreateButton("Open", 205, 20, 80, 20)
;	$boot_config_save_as_button = 												GUICtrlCreateButton("Save As", 295, 20, 80, 20)
;	$boot_config_edit = 														GUICtrlCreateEdit("", 10, 50, 620, 400)
;	$boot_config_status_input = 												GUICtrlCreateInput("", 10, 640 - 25, 640 - 20, 20, $ES_READONLY, $WS_EX_STATICEDGE)

	$edit_time_entry_gui = 														ChildGUICreate($app_name & " - Edit Time Entry", 640, 640, $main_gui)
;	GUICtrlCreateGroupEx  ("----> RetroPie (/boot/config.txt)", 5, 5, 180, 40)
	$edit_time_entry_project_listview = 										GUICtrlCreateListViewEx(10, 10, 550, 100, "Project", 160, "ID", 160)
	$edit_time_entry_task_listview = 											GUICtrlCreateListViewEx(10, 130, 550, 100, "Task", 160, "ID", 160)
	$edit_time_entry_save_button = 												GUICtrlCreateImageButton("save.ico", 10, 640 - 70, 36, "Save this Time Entry")
	;$edit_time_entry_cancel_button = 											GUICtrlCreateImageButton("cancel.ico", 50, 640 - 70, 36, "Cancel this edit")
	$edit_time_entry_status_input = 											GUICtrlCreateStatusInput("", 10, 640 - 25, 640 - 20, 20)


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

			Local $days_from_today
			Local $last_monday_date
			Local $date_part

			for $days_from_today = 0 to -7 step -1

				$last_monday_date = _DateAdd('d', $days_from_today, _NowCalcDate())
				$date_part = StringSplit($last_monday_date, "/", 3)
				Local $spent_date_day_to_week_index = _DateToDayOfWeek($date_part[0], $date_part[1], $date_part[2])

				if $spent_date_day_to_week_index = 2 Then ExitLoop
			Next

			_GUICtrlListView_SetGroupInfo($timesheet_listview, 1, "Monday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 1, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 2, "Tuesday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 2, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 3, "Wednesday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 3, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 4, "Thursday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 4, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 5, "Friday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 5, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 6, "Saturday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))
			$date_part = StringSplit(_DateAdd('d', $days_from_today + 6, _NowCalcDate()), "/", 3)
			_GUICtrlListView_SetGroupInfo($timesheet_listview, 7, "Sunday " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME))

			Local $this_sunday_date = _DateAdd('d', 6, $last_monday_date)

			GUICtrlStatusInput_SetText($status_input, "Please Wait. Getting your Harvest times ...")
			Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?from=' & StringReplace($last_monday_date, "/", "-") & '&to=' & StringReplace($this_sunday_date, "/", "-") & ' -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $this_sunday_date = ' & $this_sunday_date & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $last_monday_date = ' & $last_monday_date & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			;Local $iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?from=2021-07-19&to=2021-07-25 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			GUICtrlStatusInput_SetText($status_input, "")

			Local $decoded_json = Json_Decode($json)
			_GUICtrlListView_DeleteAllItems($timesheet_listview)
			_GUICtrlListView_BeginUpdate($timesheet_listview)
			Local $times_exist_for_day[7] = [False, False, False, False, False, False, False]

			for $i = 99 to 0 step -1

				Local $spent_date = Json_Get($decoded_json, '.time_entries[' & $i & '].spent_date')

				if StringLen($spent_date) > 0 Then

					Local $spent_date_day_to_week_index = _DateToDayOfWeek(StringLeft($spent_date, 4), StringMid($spent_date, 6, 2), StringRight($spent_date, 2))
					$spent_date = _DateDayOfWeek($spent_date_day_to_week_index)
					Local $project = Json_Get($decoded_json, '.time_entries[' & $i & '].project.name')
					Local $task = Json_Get($decoded_json, '.time_entries[' & $i & '].task.name')
					Local $notes = Json_Get($decoded_json, '.time_entries[' & $i & '].notes')
					Local $hours = Json_Get($decoded_json, '.time_entries[' & $i & '].hours')

					Local $index = _GUICtrlListView_AddItem($timesheet_listview, $project)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $task, 1)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $notes, 2)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, $hours, 3)
					_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, $spent_date_day_to_week_index - 1)
					$times_exist_for_day[$spent_date_day_to_week_index - 2] = True
				EndIf
			Next

			for $i = 0 to (UBound($times_exist_for_day) - 1)

				if $times_exist_for_day[$i] = False Then

					Local $index = _GUICtrlListView_AddItem($timesheet_listview, "<click here then add button above>")
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 1)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 2)
					_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 3)
					_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, $i + 1)
				EndIf
			Next

			_GUICtrlListView_EndUpdate($timesheet_listview)
			raise_button_and_enable_gui($msg)

		Case $timesheet_add_button

			depress_button_and_disable_gui($msg)
			GUISetState(@SW_DISABLE, $main_gui)
			GUISetState(@SW_SHOW, $add_time_entry_gui)
			$current_gui = $add_time_entry_gui
			GUICtrlListBoxFromFile($add_time_entry_favourites_list, $favourites_path)
			GUICtrlSetState($add_time_entry_save_button, $GUI_DEFBUTTON)

			if $project_assignments_loaded = False Then

				_GUICtrlListView_DeleteAllItems($add_time_entry_project_listview)

				Local $total_pages = 1
				Local $page_num = 0

				while True

					$page_num = $page_num + 1

					if $page_num > $total_pages or $page_num > 10 Then

						ExitLoop
					EndIf

					if $page_num <= 1 Then

						GUICtrlStatusInput_SetText($add_time_entry_status_input, "Please Wait. Getting your Harvest projects and tasks (page " & $page_num & ") ...")
					Else

						GUICtrlStatusInput_SetText($add_time_entry_status_input, "Please Wait. Getting your Harvest projects and tasks (page " & $page_num & " of " & $total_pages & ") ...")
					EndIf

					Local $iPID = Run('curl -k https://api.harvestapp.com/v2/users/me/project_assignments?page=' & $page_num & '&per_page=100 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
					ProcessWaitClose($iPID)
					Local $json = StdoutRead($iPID)
					Local $decoded_json = Json_Decode($json)

					$total_pages = Json_Get($decoded_json, '.total_pages')

					for $project_index = 0 to 99

						Local $project_name = Json_Get($decoded_json, '.project_assignments[' & $project_index & '].project.name')
						Local $project_id = Json_Get($decoded_json, '.project_assignments[' & $project_index & '].project.id')

						if StringLen($project_name) > 0 Then

							Local $task_names = ""

							for $task_index = 0 to 99

								Local $task_id = Json_Get($decoded_json, '.project_assignments[' & $project_index & '].task_assignments[' & $task_index & '].task.id')
								Local $task_name = Json_Get($decoded_json, '.project_assignments[' & $project_index & '].task_assignments[' & $task_index & '].task.name')

								if StringLen($task_name) < 1 Then ExitLoop

								$task_name = $task_name & "|" & $task_id

								if StringLen($task_names) > 0 Then $task_names = $task_names & @CRLF

								$task_names = $task_names & $task_name
							Next

							$timesheet_project_id_dict.Add($project_name, $project_id)
							$timesheet_project_assignments_dict.Add($project_name, $task_names)
						EndIf
					Next
				WEnd

				GUICtrlStatusInput_SetText($add_time_entry_status_input, "")
				_GUICtrlListView_BeginUpdate($add_time_entry_project_listview)

				For $vKey In $timesheet_project_assignments_dict

					Local $index = _GUICtrlListView_AddItem($add_time_entry_project_listview, $vKey)
					_GUICtrlListView_AddSubItem($add_time_entry_project_listview, $index, $timesheet_project_id_dict.Item($vKey), 1)
				Next

				_GUICtrlListView_EndUpdate($add_time_entry_project_listview)
				_GUICtrlListView_SetItemSelected($add_time_entry_project_listview, 0, true, true)
				GUICtrlSetState($add_time_entry_project_listview, $GUI_FOCUS)
				$project_assignments_loaded = True
				$update_tasks = True
			EndIf

		Case $add_time_entry_save_button

			Local $group_info = _GUICtrlListView_GetGroupInfo($timesheet_listview, _GUICtrlListView_GetItemGroupID($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))))
			Local $time_entry_date = $group_info[0]
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $time_entry_date = ' & $time_entry_date & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $time_entry_date_part = StringSplit($group_info[0], " ", 3)
			$time_entry_date_part[2] = _ConvertMonth($time_entry_date_part[2])
			Local $selected_project_name = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)
			Local $selected_project_id = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 1)
			Local $selected_task_name = _GUICtrlListView_GetItemText($add_time_entry_task_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_task_listview)), 0)
			Local $selected_task_id = _GUICtrlListView_GetItemText($add_time_entry_task_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_task_listview)), 1)

			GUICtrlStatusInput_SetText($add_time_entry_status_input, "Please Wait. Saving the time entry ...")

			Local $iPID = Run('curl -k "https://api.harvestapp.com/v2/time_entries?project_id=' & $selected_project_id & '&task_id=' & $selected_task_id & '&spent_date=' & @YEAR & '-' & $time_entry_date_part[2] & '-' & $time_entry_date_part[1] & '&hours=1.0" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)" -X POST -H "Content-Type: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $decoded_json = Json_Decode($json)

			GUICtrlStatusInput_SetText($add_time_entry_status_input, "")

			_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $selected_project_name, 0)
			_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $selected_task_name, 1)
			_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "", 2)
			_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "1", 3)

			GUISetState(@SW_ENABLE, $main_gui)
			GUISetState(@SW_HIDE, $current_gui)
			$current_gui = $main_gui
			raise_button_and_enable_gui($timesheet_add_button)

		Case $add_time_entry_favourites_add_button

			depress_button_and_disable_gui($msg, $current_gui)
			$result = InputBox($app_name, "Enter a favourite task name", "", "", 240, 140, Default, Default, 0, $main_gui)

			if StringLen($result) > 0 Then

				_GUICtrlListBox_AddString($add_time_entry_favourites_list, $result)
			EndIf
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_favourites_list, $favourites_path)
			SelectFavouriteTask()

		Case $add_time_entry_favourites_delete_button

			depress_button_and_disable_gui($msg, $current_gui, 100)
			_GUICtrlListBox_DeleteString($add_time_entry_favourites_list, _GUICtrlListBox_GetCurSel($add_time_entry_favourites_list))
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_favourites_list, $favourites_path)
			SelectFavouriteTask()

	EndSwitch

	if $update_tasks = True Then

		$update_tasks = False
		Local $selected_project = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)

		_GUICtrlListView_DeleteAllItems($add_time_entry_task_listview)

		Local $task_names = $timesheet_project_assignments_dict.Item($selected_project)

		;Local $task_name_arr = StringSplit($task_names, "|")
		Local $task_name_arr = _StringSplit2d($task_names, "|")

		if UBound($task_name_arr) > 0 Then

			;_ArrayDisplay($task_name_arr)

			;$rr = $task_name_arr[UBound($task_name_arr) - 1]
			;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $rr = ' & $rr & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			if StringLen($task_name_arr[UBound($task_name_arr) - 1][0]) < 1 Then _ArrayDelete($task_name_arr, UBound($task_name_arr) - 1)

			_GUICtrlListView_BeginUpdate($add_time_entry_task_listview)
			_GUICtrlListView_AddArray($add_time_entry_task_listview, $task_name_arr)
			_GUICtrlListView_EndUpdate($add_time_entry_task_listview)

			_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, 0, true, False)
			SelectFavouriteTask()
		EndIf

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

Func GUICtrlListBoxToFile($list, $path)

	FileDelete($path)
	Local $hFileOpen = FileOpen($path, $FO_APPEND)

	for $i = 0 to (_GUICtrlListBox_GetCount($list) - 1)

		FileWriteLine($hFileOpen, _GUICtrlListBox_GetText($list, $i))
	Next

	FileClose($hFileOpen)
EndFunc

Func GUICtrlListBoxFromFile($list, $path)

	if FileExists($path) = True Then

		Local $arr
		_FileReadToArray($path, $arr)

		if UBound($arr) > 0 Then

			_GUICtrlListBox_ResetContent($list)
			_GUICtrlListBox_BeginUpdate($list)

			for $i = 1 to $arr[0]

				_GUICtrlListBox_AddString($list, $arr[$i])
			Next

			_GUICtrlListBox_EndUpdate($list)
		EndIf
	EndIf
EndFunc

Func SelectFavouriteTask()

	for $i = 0 to (_GUICtrlListBox_GetCount($add_time_entry_favourites_list) - 1)

		$result = _GUICtrlListView_FindText($add_time_entry_task_listview, _GUICtrlListBox_GetText($add_time_entry_favourites_list, $i), -1, False)

		if $result > -1 Then

			_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, $result, true, False)
			ExitLoop
		EndIf
	Next

EndFunc

