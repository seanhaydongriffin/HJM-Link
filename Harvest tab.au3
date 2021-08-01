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
Global $project_filters_path = $app_data_dir & "\project filters.txt"
Global $task_filters_path = $app_data_dir & "\task filters.txt"
Global $favourites_path = $app_data_dir & "\favourites.txt"
Global $project_assignments_loaded = False
Global $tmp = -1
Global $favourite[0][5]

Func Harvest_tab_setup()

	GUICtrlCreateTabItemEx("Harvest")
	;GUICtrlCreateGroupEx  ("", 20, 140, 250, 500)
	$timesheet_week_total_label = 												GUICtrlCreateLabelEx("Week Starting", 40, 110, 120, 20, "", $GUI_DOCKALL)
	$timesheet_week_combo = 													GUICtrlCreateComboEx(130, 105, 100, 20, "", $GUI_DOCKALL)


	$startjuldate = _DateToDayValue(@YEAR - 1,1,1)
	$endjuldate = _DateToDayValue(@YEAR,12,31)
	Global $iYear, $iMonth, $iDay

	For $x =  $startjuldate To $endjuldate
		_DayValueToDate ( $x, $iYear,$iMonth, $iDay )
		if _DateToDayOfWeek($iYear,$iMonth,$iDay) = 2 Then

			_GUICtrlComboBox_AddString($timesheet_week_combo, $iDay & "/" & $iMonth & "/" & $iYear)
		EndIf
	Next

	_GUICtrlComboBox_SelectString($timesheet_week_combo, GetLastMondayDate("dd/MM/yyyy"))

	$timesheet_this_week_button = 												GUICtrlCreateImageButton("week.ico", 250, 90, 36, "View this week", $GUI_DOCKALL)
	$timesheet_refresh_button = 												GUICtrlCreateImageButton("refresh.ico", 290, 90, 36, "Get your Harvest times (ALT+R)", $GUI_DOCKALL, False, "&R")
	$timesheet_add_button = 													GUICtrlCreateImageButton("add.ico", 330, 90, 36, "Add a new Time Entry (ALT+A)", $GUI_DOCKALL, False, "&A")
	$timesheet_edit_button = 													GUICtrlCreateImageButton("edit.ico", 370, 90, 36, "Edit the selected Time Entry (ALT+E)", -1, False, "&E")
	$timesheet_delete_button = 													GUICtrlCreateImageButton("delete.ico", 410, 90, 36, "Delete the selected Time Entry (ALT+D)", $GUI_DOCKALL, False, "&D")
	$timesheet_sync_to_jira_checkbox = 											GUICtrlCreateCheckboxEx("Synchronise timesheet to Jira", 450, 70, 150, 20, True, "", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$timesheet_week_total_label = 												GUICtrlCreateLabelEx("Week Total = 0:00", 685, 110, 400, 20, "", $GUI_DOCKRIGHT + $GUI_DOCKWIDTH + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)



	$timesheet_listview = 														GUICtrlCreateListViewEx(30, 130, 775, 500, $GUI_DOCKBORDERS, "Project", 240, "Task", 160, "Notes", 260, "Hours", 90, "ID", 0)
	_GUICtrlListView_JustifyColumn($timesheet_listview, 3, 1)
	_GUICtrlListView_EnableGroupView($timesheet_listview)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 1, "Mon", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 2, "Tue", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 3, "Wed", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 4, "Thu", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 5, "Fri", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 6, "Sat", 2)
	_GUICtrlListView_InsertGroup($timesheet_listview, -1, 7, "Sun", 2)
;	$timesheet_tmp_button = 													GUICtrlCreateImageButton("edit.ico", 190, 80, 36, "Edit the selected Time Entry")

EndFunc

Func Harvest_tab_child_gui_setup()

	$add_time_entry_gui = 														ChildGUICreate($app_name & " - Add Time Entry", 640, 640, $main_gui)

	GUICtrlCreateGroupEx ("Favourites", 5, 5, 630, 60, "", $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
	$add_time_entry_favourites_combo = 											GUICtrlCreateComboEx(10, 25, 500, 20, "", $GUI_DOCKALL)

	Global $favourite[0][5]	; clear the array
	_FileReadToArray($favourites_path, $favourite, 0, chr(29))
	_ArraySort($favourite)

	for $i = 0 to (UBound($favourite) - 1)

		_GUICtrlComboBox_AddString($add_time_entry_favourites_combo, $favourite[$i][0])
	Next

	GUICtrlCreateGroupEx ("Project", 5, 65, 630, 200)
	$add_time_entry_project_listview = 											GUICtrlCreateListViewEx(10, 85, 400, 175, $GUI_DOCKBORDERS, "Name", 600, "ID", 0)
	GUICtrlCreateGroupEx ("Task", 5, 275, 630, 200, "", $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	$add_time_entry_task_listview = 											GUICtrlCreateListViewEx(10, 295, 400, 170, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT, "Name", 600, "ID", 0)
	GUICtrlCreateGroupEx ("Notes", 5, 485, 630, 50, "", $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	$add_time_entry_notes_input = 												GUICtrlCreateInputEx("", 10, 500, 610, 20, "", $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	$add_time_entry_hour_input = 												GUICtrlCreateInputEx("", 40, 540, 40, 20, "", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_hour_input_radio =											GUICtrlCreateRadioEx("", 20, 540, 20, 20, False, "", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_half_hour_radio =											GUICtrlCreateRadioEx("0.5", 90, 540, 40, 20, False, "half hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_one_hour_radio =											GUICtrlCreateRadioEx("1.0", 140, 540, 40, 20, True, "one hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_one_half_hour_radio =										GUICtrlCreateRadioEx("1.5", 190, 540, 40, 20, False, "one and half hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_two_hour_radio =											GUICtrlCreateRadioEx("2.0", 240, 540, 40, 20, False, "two hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_two_half_hour_radio =										GUICtrlCreateRadioEx("2.5", 290, 540, 40, 20, False, "two and half hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_three_hour_radio =											GUICtrlCreateRadioEx("3.0", 340, 540, 40, 20, False, "three hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_three_half_hour_radio =										GUICtrlCreateRadioEx("3.5", 390, 540, 40, 20, False, "three and half hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_four_hour_radio =											GUICtrlCreateRadioEx("4.0", 440, 540, 40, 20, False, "four hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_four_half_hour_radio =										GUICtrlCreateRadioEx("4.5", 490, 540, 40, 20, False, "four and half hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
    $add_time_entry_five_hour_radio =											GUICtrlCreateRadioEx("5.0", 540, 540, 40, 20, False, "five hour", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_save_button = 												GUICtrlCreateImageButton("save.ico", 10, 640 - 70, 36, "Save this new Time Entry", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	;$add_time_entry_cancel_button = 											GUICtrlCreateImageButton("cancel.ico", 50, 640 - 70, 36, "Cancel this Time Entry")

	GUICtrlCreateGroupEx ("Filters", 420, 80, 200, 185, "", $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
	$add_time_entry_project_filters_list = 										GUICtrlCreateListEx(425, 100, 180, 110, "", $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
	$add_time_entry_project_filters_add_button = 								GUICtrlCreateImageButton("add.ico", 425, 220, 36, "Add a new Project Filter", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_project_filters_delete_button = 							GUICtrlCreateImageButton("delete.ico", 465, 220, 36, "Delete the selected Project Filter", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_project_filters_enable_checkbox = 							GUICtrlCreateCheckboxEx("Enable filters", 505, 220, 100, 20, True, "", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlCreateGroupEx ("Filters", 420, 290, 200, 180, "", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_task_filters_list = 										GUICtrlCreateListEx(425, 310, 180, 110, "", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_task_filters_add_button = 									GUICtrlCreateImageButton("add.ico", 425, 430, 36, "Add a new Favourite Task", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_task_filters_delete_button = 								GUICtrlCreateImageButton("delete.ico", 465, 430, 36, "Delete the selected Favourite Task", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_task_filters_enable_checkbox = 								GUICtrlCreateCheckboxEx("Enable filters", 505, 430, 100, 20, True, "", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$add_time_entry_favourites_add_button = 									GUICtrlCreateImageButton("add.ico", 520, 20, 36, "Add a new Favourite with the selections below", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	$add_time_entry_favourites_delete_button = 									GUICtrlCreateImageButton("delete.ico", 560, 20, 36, "Delete the selected Favourite", $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)


	$add_time_entry_status_input = 												GUICtrlCreateStatusInput("", 10, 640 - 25, 640 - 20, 20)
;	GUICtrlCreateGroupEx  ("----> PC", 200, 5, 180, 40)
;	$boot_config_open_button = 													GUICtrlCreateButton("Open", 205, 20, 80, 20)
;	$boot_config_save_as_button = 												GUICtrlCreateButton("Save As", 295, 20, 80, 20)
;	$boot_config_edit = 														GUICtrlCreateEdit("", 10, 50, 620, 400)
;	$boot_config_status_input = 												GUICtrlCreateInput("", 10, 640 - 25, 640 - 20, 20, $ES_READONLY, $WS_EX_STATICEDGE)

	$edit_time_entry_gui = 														ChildGUICreate($app_name & " - Edit Time Entry", 640, 640, $main_gui)
;	GUICtrlCreateGroupEx  ("----> RetroPie (/boot/config.txt)", 5, 5, 180, 40)
	$edit_time_entry_project_listview = 										GUICtrlCreateListViewEx(10, 10, 550, 100, $GUI_DOCKBORDERS, "Project", 160, "ID", 160)
	$edit_time_entry_task_listview = 											GUICtrlCreateListViewEx(10, 130, 550, 100, $GUI_DOCKBORDERS, "Task", 160, "ID", 160)
	$edit_time_entry_save_button = 												GUICtrlCreateImageButton("save.ico", 10, 640 - 70, 36, "Save this Time Entry")
	;$edit_time_entry_cancel_button = 											GUICtrlCreateImageButton("cancel.ico", 50, 640 - 70, 36, "Cancel this edit")
	$edit_time_entry_status_input = 											GUICtrlCreateStatusInput("", 10, 640 - 25, 640 - 20, 20)


EndFunc


Func Harvest_tab_event_handler($msg)

	Switch $msg

		Case $GUI_EVENT_RESIZED

			if $current_gui = $add_time_entry_gui Then

				; below is a workaround for docking rich edit controls ($status_input)
				$aSize = WinGetClientSize($add_time_entry_gui)
				_WinAPI_SetWindowPos($add_time_entry_status_input, $HWND_TOP, 5, $aSize[1] - 25, $aSize[0] - 10, 20, $SWP_SHOWWINDOW)
			EndIf

		Case $GUI_EVENT_CLOSE

			if $current_gui = $add_time_entry_gui Then

				TimeEntrySetState($GUI_ENABLE)
				if StringInStr(WinGetTitle($current_gui), "Add Time Entry") > 0 Then raise_button_and_enable_gui($timesheet_add_button)
				if StringInStr(WinGetTitle($current_gui), "Edit Time Entry") > 0 Then raise_button_and_enable_gui($timesheet_edit_button)
				GUISetState(@SW_ENABLE, $main_gui)
				GUISetState(@SW_HIDE, $current_gui)
				$current_gui = $main_gui
			EndIf



		case $timesheet_refresh_button

			depress_button_and_disable_gui($msg) ;, -1, 100)
			RefreshTimesheet(GetLastMondayDate())
			raise_button_and_enable_gui($msg)
			_GUICtrlListView_SetItemSelected($timesheet_listview, GUICtrlListView_GetTopMostIndex($timesheet_listview), true, true)
			GUICtrlSetState($timesheet_listview, $GUI_FOCUS)


		Case $timesheet_add_button, $timesheet_edit_button

			depress_button_and_disable_gui($msg)
			GUISetState(@SW_DISABLE, $main_gui)
			GUISetState(@SW_SHOW, $add_time_entry_gui)
			$current_gui = $add_time_entry_gui
			GUICtrlSetState($add_time_entry_save_button, $GUI_DEFBUTTON)

			$selected_group_info = _GUICtrlListView_GetGroupInfo($timesheet_listview, _GUICtrlListView_GetItemGroupID($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))))
			Local $selected_timesheet_date_part = StringSplit($selected_group_info[0], " = ", 3)
			$selected_timesheet_date = $selected_timesheet_date_part[0]

			if $msg = $timesheet_add_button Then WinSetTitle($current_gui, "", $app_name & " - Add Time Entry for " & $selected_timesheet_date)
			if $msg = $timesheet_edit_button Then WinSetTitle($current_gui, "", $app_name & " - Edit Time Entry for " & $selected_timesheet_date)

			; pull the filters from file
			GUICtrlListBoxFromFile($add_time_entry_project_filters_list, $project_filters_path)
			GUICtrlListBoxFromFile($add_time_entry_task_filters_list, $task_filters_path)


			if $project_assignments_loaded = False Then

				TimeEntrySetState($GUI_DISABLE)
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

					$iPID = Run('curl -k https://api.harvestapp.com/v2/users/me/project_assignments?page=' & $page_num & '&per_page=100 -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
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
				FilterProject()

				if $msg = $timesheet_add_button Then _GUICtrlListView_SetItemSelected($add_time_entry_project_listview, 0, true, true)

				TimeEntrySetState($GUI_ENABLE)
				$project_assignments_loaded = True
				$update_tasks = True
			EndIf

			if $msg = $timesheet_edit_button Then

				Local $selected_project_name = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 0)
				Local $selected_task_name = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 1)
				Local $selected_notes = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 2)
				Local $selected_hours = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 3)

				_GUICtrlListView_SetItemSelected($add_time_entry_project_listview, _GUICtrlListView_FindText($add_time_entry_project_listview, $selected_project_name, -1, False), True, true)
				_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, _GUICtrlListView_FindText($add_time_entry_task_listview, $selected_task_name, -1, False), True, true)
				GUICtrlSetData($add_time_entry_notes_input, $selected_notes)
				GUICtrlSetState($add_time_entry_hour_input_radio, $GUI_CHECKED)
				GUICtrlSetData($add_time_entry_hour_input, $selected_hours)
			EndIf

			GUICtrlSetState($add_time_entry_favourites_combo, $GUI_FOCUS)

;		Case $timesheet_tmp_button

;			$tmp = $tmp + 1
			;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $tmp = ' & $tmp & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			;_GUICtrlListView_SetItemSelected($timesheet_listview, $tmp, true, True)

			;$yy = _GUICtrlListView_GetNextItem($timesheet_listview, -1, 0, 0)
			;$yy = _GUICtrlListView_GetItemPositionY($timesheet_listview, $tmp)

;			$yy = GUICtrlListView_GetTopMostIndex($timesheet_listview)

;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $yy = ' & $yy & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Case $timesheet_this_week_button

			_GUICtrlComboBox_SelectString($timesheet_week_combo, GetLastMondayDate("dd/MM/yyyy"))

			depress_button_and_disable_gui($msg, -1, 100)
			RefreshTimesheet(GetLastMondayDate())
			raise_button_and_enable_gui($msg)
			_GUICtrlListView_SetItemSelected($timesheet_listview, GUICtrlListView_GetTopMostIndex($timesheet_listview), true, true)
			GUICtrlSetState($timesheet_listview, $GUI_FOCUS)


		Case $add_time_entry_favourites_add_button

			$result = InputBox($app_name, "Enter the name to give this favourite", "", "", 240, 140, Default, Default, 0, $main_gui)

			if StringLen($result) > 0 Then

				Local $selected_project_name = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)
				Local $selected_task_name = _GUICtrlListView_GetItemText($add_time_entry_task_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_task_listview)), 0)
				Local $notes = GUICtrlRead($add_time_entry_notes_input)
				Local $hours = 0

				if GUICtrlRead($add_time_entry_hour_input_radio) = $GUI_CHECKED Then $hours = GUICtrlRead($add_time_entry_hour_input)
				if GUICtrlRead($add_time_entry_half_hour_radio) = $GUI_CHECKED Then $hours = "0.5"
				if GUICtrlRead($add_time_entry_one_hour_radio) = $GUI_CHECKED Then $hours = "1.0"
				if GUICtrlRead($add_time_entry_one_half_hour_radio) = $GUI_CHECKED Then $hours = "1.5"
				if GUICtrlRead($add_time_entry_two_hour_radio) = $GUI_CHECKED Then $hours = "2.0"
				if GUICtrlRead($add_time_entry_two_half_hour_radio) = $GUI_CHECKED Then $hours = "2.5"
				if GUICtrlRead($add_time_entry_three_hour_radio) = $GUI_CHECKED Then $hours = "3.0"
				if GUICtrlRead($add_time_entry_three_half_hour_radio) = $GUI_CHECKED Then $hours = "3.5"
				if GUICtrlRead($add_time_entry_four_hour_radio) = $GUI_CHECKED Then $hours = "4.0"
				if GUICtrlRead($add_time_entry_four_half_hour_radio) = $GUI_CHECKED Then $hours = "4.5"
				if GUICtrlRead($add_time_entry_five_hour_radio) = $GUI_CHECKED Then $hours = "5.0"

				_ArrayAdd($favourite, $result & Chr(29) & $selected_project_name & Chr(29) & $selected_task_name & Chr(29) & $notes & Chr(29) & $hours, 0, Chr(29))
				_ArraySort($favourite)
				FileDelete($favourites_path)
				_FileWriteFromArray($favourites_path, $favourite, Default, Default, Chr(29))
				_GUICtrlComboBox_ResetContent($add_time_entry_favourites_combo)

				for $i = 0 to (UBound($favourite) - 1)

					_GUICtrlComboBox_AddString($add_time_entry_favourites_combo, $favourite[$i][0])
				Next

				_GUICtrlComboBox_SelectString($add_time_entry_favourites_combo, $result)


			EndIf

		Case $add_time_entry_favourites_delete_button

			$result = _ArraySearch($favourite, GUICtrlRead($add_time_entry_favourites_combo), 0, 0, 1, 0, 1, 0)

			if $result > -1 Then

				_ArrayDelete($favourite, $result)
				FileDelete($favourites_path)
				_FileWriteFromArray($favourites_path, $favourite, Default, Default, Chr(29))

			EndIf

			_GUICtrlComboBox_DeleteString($add_time_entry_favourites_combo, _GUICtrlComboBox_GetCurSel($add_time_entry_favourites_combo))

		Case $add_time_entry_save_button

			Local $group_info = _GUICtrlListView_GetGroupInfo($timesheet_listview, _GUICtrlListView_GetItemGroupID($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))))
			Local $time_entry_date = $group_info[0]
			Local $time_entry_date_part = StringSplit($group_info[0], " ", 3)
			$time_entry_date_part[2] = _ConvertMonth($time_entry_date_part[2])
			Local $selected_project_name = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)
			Local $selected_project_id = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 1)
			Local $selected_task_name = _GUICtrlListView_GetItemText($add_time_entry_task_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_task_listview)), 0)
			Local $selected_task_id = _GUICtrlListView_GetItemText($add_time_entry_task_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_task_listview)), 1)
			Local $hours = 0

			if GUICtrlRead($add_time_entry_hour_input_radio) = $GUI_CHECKED Then $hours = GUICtrlRead($add_time_entry_hour_input)
			if GUICtrlRead($add_time_entry_half_hour_radio) = $GUI_CHECKED Then $hours = "0.5"
			if GUICtrlRead($add_time_entry_one_hour_radio) = $GUI_CHECKED Then $hours = "1.0"
			if GUICtrlRead($add_time_entry_one_half_hour_radio) = $GUI_CHECKED Then $hours = "1.5"
			if GUICtrlRead($add_time_entry_two_hour_radio) = $GUI_CHECKED Then $hours = "2.0"
			if GUICtrlRead($add_time_entry_two_half_hour_radio) = $GUI_CHECKED Then $hours = "2.5"
			if GUICtrlRead($add_time_entry_three_hour_radio) = $GUI_CHECKED Then $hours = "3.0"
			if GUICtrlRead($add_time_entry_three_half_hour_radio) = $GUI_CHECKED Then $hours = "3.5"
			if GUICtrlRead($add_time_entry_four_hour_radio) = $GUI_CHECKED Then $hours = "4.0"
			if GUICtrlRead($add_time_entry_four_half_hour_radio) = $GUI_CHECKED Then $hours = "4.5"
			if GUICtrlRead($add_time_entry_five_hour_radio) = $GUI_CHECKED Then $hours = "5.0"

			Local $notes = GUICtrlRead($add_time_entry_notes_input)

			GUICtrlStatusInput_SetText($add_time_entry_status_input, "Please Wait. Saving the time entry ...")

			if StringInStr(WinGetTitle($current_gui), "Add Time Entry") > 0 Or StringCompare(_GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))), "<click here then add button above>") = 0 Then

				$iPID = Run('curl -k "https://api.harvestapp.com/v2/time_entries?project_id=' & $selected_project_id & '&task_id=' & $selected_task_id & '&spent_date=' & @YEAR & '-' & $time_entry_date_part[2] & '-' & $time_entry_date_part[1] & '&hours=' & HourAndMinutesToHours($hours) & '&notes=' & _URIEncode($notes) & '" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)" -X POST -H "Content-Type: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			Else

				; Edit Time Entry

				Local $selected_time_entry_id = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 4)

				$iPID = Run('curl -k "https://api.harvestapp.com/v2/time_entries/' & $selected_time_entry_id & '?project_id=' & $selected_project_id & '&task_id=' & $selected_task_id & '&spent_date=' & @YEAR & '-' & $time_entry_date_part[2] & '-' & $time_entry_date_part[1] & '&hours=' & HourAndMinutesToHours($hours) & '&notes=' & _URIEncode($notes) & '" -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)" -X PATCH -H "Content-Type: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			EndIf

			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $decoded_json = Json_Decode($json)
			Local $id = Json_Get($decoded_json, '.id')
			GUICtrlStatusInput_SetText($add_time_entry_status_input, "")

			if StringInStr(WinGetTitle($current_gui), "Edit Time Entry") Or StringCompare(_GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))), "<click here then add button above>") = 0 Then

				_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $selected_project_name, 0)
				_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $selected_task_name, 1)
				_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $notes, 2)
				_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), HoursToHourAndMinutes($hours, True), 3)
				_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), $id, 4)
			Else

				Local $index = _GUICtrlListView_AddItem($timesheet_listview, $selected_project_name)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $selected_task_name, 1)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $notes, 2)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, HoursToHourAndMinutes($hours, True), 3)
				_GUICtrlListView_AddSubItem($timesheet_listview, $index, $id, 4)
				_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, _GUICtrlListView_GetItemGroupID($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))))
			EndIf

			UpdateTotalHours()

			TimeEntrySetState($GUI_ENABLE)
			if StringInStr(WinGetTitle($current_gui), "Add Time Entry") > 0 Then raise_button_and_enable_gui($timesheet_add_button)
			if StringInStr(WinGetTitle($current_gui), "Edit Time Entry") > 0 Then raise_button_and_enable_gui($timesheet_edit_button)
			GUISetState(@SW_ENABLE, $main_gui)
			GUISetState(@SW_HIDE, $current_gui)
			$current_gui = $main_gui
			GUICtrlSetState($timesheet_listview, $GUI_FOCUS)

		Case $add_time_entry_project_filters_add_button

			depress_button_and_disable_gui($msg, $current_gui)
			$result = InputBox($app_name, "Enter text for filtering project names", "", "", 240, 140, Default, Default, 0, $main_gui)

			if StringLen($result) > 0 Then

				_GUICtrlListBox_AddString($add_time_entry_project_filters_list, $result)
			EndIf
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_project_filters_list, $project_filters_path)
			FilterProject()

		Case $add_time_entry_project_filters_delete_button

			depress_button_and_disable_gui($msg, $current_gui, 100)
			_GUICtrlListBox_DeleteString($add_time_entry_project_filters_list, _GUICtrlListBox_GetCurSel($add_time_entry_project_filters_list))
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_project_filters_list, $project_filters_path)
			FilterProject()

		Case $add_time_entry_task_filters_add_button

			depress_button_and_disable_gui($msg, $current_gui)
			$result = InputBox($app_name, "Enter a favourite task name", "", "", 240, 140, Default, Default, 0, $main_gui)

			if StringLen($result) > 0 Then

				_GUICtrlListBox_AddString($add_time_entry_task_filters_list, $result)
			EndIf
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_task_filters_list, $task_filters_path)
			$update_tasks = True

		Case $add_time_entry_task_filters_delete_button

			depress_button_and_disable_gui($msg, $current_gui, 100)
			_GUICtrlListBox_DeleteString($add_time_entry_task_filters_list, _GUICtrlListBox_GetCurSel($add_time_entry_task_filters_list))
			raise_button_and_enable_gui($msg, $current_gui)

			GUICtrlListBoxToFile($add_time_entry_task_filters_list, $task_filters_path)
			$update_tasks = True

		Case $timesheet_delete_button

			depress_button_and_disable_gui($msg)
			Local $selected_project_name = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 0)
			Local $selected_id = _GUICtrlListView_GetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), 4)

			if StringCompare($selected_project_name, "<click here then add button above>") <> 0 Then

				GUICtrlStatusInput_SetText($status_input, "Please Wait. Deleting the time entry ...")
				$iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries/' & $selected_id & ' -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)" -X DELETE', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
				ProcessWaitClose($iPID)
				Local $json = StdoutRead($iPID)
				GUICtrlStatusInput_SetText($status_input, "")

				Local $selected_group_id = _GUICtrlListView_GetItemGroupID($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)))
				Local $num_items_in_selected_group = 0

				for $i = 0 to (_GUICtrlListView_GetItemCount($timesheet_listview) - 1)

					if _GUICtrlListView_GetItemGroupID($timesheet_listview, $i) = $selected_group_id then $num_items_in_selected_group = $num_items_in_selected_group + 1
				Next

				if $num_items_in_selected_group = 1 Then

					_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "<click here then add button above>", 0)
					_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "", 1)
					_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "", 2)
					_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "", 3)
					_GUICtrlListView_SetItemText($timesheet_listview, Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview)), "", 4)
				Else

					_GUICtrlListView_DeleteItemsSelected($timesheet_listview)
				EndIf
			EndIf

			UpdateTotalHours()
			raise_button_and_enable_gui($msg)

	EndSwitch

	if $update_tasks = True Then

		$update_tasks = False
		FilterTask()

	EndIf
EndFunc


Func Harvest_tab_WM_NOTIFY_handler($hWndFrom, $iCode)


	Switch $hWndFrom


		Case GUICtrlGetHandle($timesheet_listview)

			Switch $iCode

				Case $LVN_ITEMCHANGED

					Local $index = Number(_GUICtrlListView_GetSelectedIndices($timesheet_listview))

				Case $NM_DBLCLK

					$msg = $timesheet_edit_button

			EndSwitch

		Case GUICtrlGetHandle($add_time_entry_project_listview)

			Switch $iCode

				Case $LVN_ITEMCHANGED

					if GUICtrlRead($add_time_entry_task_filters_enable_checkbox) = $GUI_CHECKED Then

						;$update_tasks = True
						FilterTask()

					Else

						UnFilterTask()
					EndIf


			EndSwitch


	EndSwitch


EndFunc


Func Harvest_tab_WM_SIZING_handler()

	; below is a workaround for docking rich edit controls ($status_input)
	$aSize = WinGetClientSize($add_time_entry_gui)
	_WinAPI_SetWindowPos($add_time_entry_status_input, $HWND_TOP, 5, $aSize[1] - 25, $aSize[0] - 10, 20, $SWP_SHOWWINDOW)

EndFunc

Func Harvest_tab_WM_COMMAND_handler($hWndFrom, $iCode)


    Switch $hWndFrom

		Case GUICtrlGetHandle($timesheet_week_combo)

			Switch $iCode

                Case $CBN_SELENDOK ; Sent when the user cancels the selection in a list box

					_GUICtrlComboBox_ShowDropDown($timesheet_week_combo, False)
					Local $week_starting_date = GUICtrlRead($timesheet_week_combo)
					$week_starting_date = _Date_Time_Convert($week_starting_date, "dd/MM/yyyy", "yyyy/MM/dd")

;					depress_button_and_disable_gui($msg, -1, 100)
					RefreshTimesheet($week_starting_date)
;					raise_button_and_enable_gui($msg)
					_GUICtrlListView_SetItemSelected($timesheet_listview, GUICtrlListView_GetTopMostIndex($timesheet_listview), true, true)
					GUICtrlSetState($timesheet_listview, $GUI_FOCUS)

			EndSwitch


		Case GUICtrlGetHandle($add_time_entry_favourites_combo)

			Switch $iCode

                Case $CBN_SELENDOK ; Sent when the user cancels the selection in a list box

					Local $index = _GUICtrlComboBox_GetCurSel($add_time_entry_favourites_combo)
					_GUICtrlListView_SetItemSelected($add_time_entry_project_listview, _GUICtrlListView_FindText($add_time_entry_project_listview, $favourite[$index][1], -1, False), True, true)
					_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, _GUICtrlListView_FindText($add_time_entry_task_listview, $favourite[$index][2], -1, False), True, true)
					GUICtrlSetData($add_time_entry_notes_input, $favourite[$index][3])
					GUICtrlSetState($add_time_entry_hour_input_radio, $GUI_CHECKED)
					GUICtrlSetData($add_time_entry_hour_input, $favourite[$index][4])

			EndSwitch



        Case GUICtrlGetHandle($add_time_entry_project_filters_enable_checkbox)

			Switch $iCode

                Case $BN_CLICKED ; Sent when the user cancels the selection in a list box

					; if checkbox is checked
					if _GUICtrlButton_GetState($hWndFrom) = 521 Then

						FilterProject()
					Else

						UnFilterProject()
					EndIf
			EndSwitch

        Case GUICtrlGetHandle($add_time_entry_task_filters_enable_checkbox)

			Switch $iCode

                Case $BN_CLICKED ; Sent when the user cancels the selection in a list box

					; if checkbox is checked
					if _GUICtrlButton_GetState($hWndFrom) = 521 Then

						FilterTask()
						GUICtrlSetState($add_time_entry_task_listview, $GUI_FOCUS)

					Else

						UnFilterTask()
						GUICtrlSetState($add_time_entry_task_listview, $GUI_FOCUS)
					EndIf
			EndSwitch

        Case GUICtrlGetHandle($add_time_entry_hour_input)

			Switch $iCode

                Case $EN_SETFOCUS

					GUICtrlSetState($add_time_entry_hour_input_radio, $GUI_CHECKED)
					GUICtrlSetState($add_time_entry_half_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_one_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_one_half_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_two_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_two_half_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_three_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_three_half_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_four_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_four_half_hour_radio, $GUI_UNCHECKED)
					GUICtrlSetState($add_time_entry_five_hour_radio, $GUI_UNCHECKED)
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

Func FilterProject()

	_GUICtrlListView_DeleteAllItems($add_time_entry_project_listview)
	_GUICtrlListView_BeginUpdate($add_time_entry_project_listview)

	For $vKey In $timesheet_project_assignments_dict

		if _GUICtrlListBox_GetCount($add_time_entry_project_filters_list) = 0 Then

			Local $index = _GUICtrlListView_AddItem($add_time_entry_project_listview, $vKey)
			_GUICtrlListView_AddSubItem($add_time_entry_project_listview, $index, $timesheet_project_id_dict.Item($vKey), 1)
		EndIf

		for $i = 0 to (_GUICtrlListBox_GetCount($add_time_entry_project_filters_list) - 1)

			if StringInStr($vKey, _GUICtrlListBox_GetText($add_time_entry_project_filters_list, $i)) > 0 Then

				Local $index = _GUICtrlListView_AddItem($add_time_entry_project_listview, $vKey)
				_GUICtrlListView_AddSubItem($add_time_entry_project_listview, $index, $timesheet_project_id_dict.Item($vKey), 1)
				ExitLoop
			EndIf
		Next
	Next

	_GUICtrlListView_EndUpdate($add_time_entry_project_listview)
	_GUICtrlListView_SetItemSelected($add_time_entry_project_listview, 0, true, true)
	GUICtrlSetState($add_time_entry_project_listview, $GUI_FOCUS)

EndFunc

Func UnFilterProject()

	_GUICtrlListView_DeleteAllItems($add_time_entry_project_listview)
	_GUICtrlListView_BeginUpdate($add_time_entry_project_listview)

	For $vKey In $timesheet_project_assignments_dict

		Local $index = _GUICtrlListView_AddItem($add_time_entry_project_listview, $vKey)
		_GUICtrlListView_AddSubItem($add_time_entry_project_listview, $index, $timesheet_project_id_dict.Item($vKey), 1)
	Next

	_GUICtrlListView_EndUpdate($add_time_entry_project_listview)
	_GUICtrlListView_SetItemSelected($add_time_entry_project_listview, 0, true, true)
	GUICtrlSetState($add_time_entry_project_listview, $GUI_FOCUS)

EndFunc

Func FilterTask()

	Local $selected_project = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)

	_GUICtrlListView_DeleteAllItems($add_time_entry_task_listview)

	Local $task_names = $timesheet_project_assignments_dict.Item($selected_project)

	;Local $task_name_arr = StringSplit($task_names, "|")
	Local $task_name_arr = _StringSplit2d($task_names, "|")

	if UBound($task_name_arr) > 0 Then

		;$rr = $task_name_arr[UBound($task_name_arr) - 1]

		if StringLen($task_name_arr[UBound($task_name_arr) - 1][0]) < 1 Then _ArrayDelete($task_name_arr, UBound($task_name_arr) - 1)

		_GUICtrlListView_BeginUpdate($add_time_entry_task_listview)
		_GUICtrlListView_AddArray($add_time_entry_task_listview, $task_name_arr)

		; remove any non filtered tasks

		for $i = 0 to (_GUICtrlListView_GetItemCount($add_time_entry_task_listview) - 1)

			Local $delete_task = True

			for $j = 0 to (_GUICtrlListBox_GetCount($add_time_entry_task_filters_list) - 1)


				if StringInStr(_GUICtrlListView_GetItemText($add_time_entry_task_listview, $i), _GUICtrlListBox_GetText($add_time_entry_task_filters_list, $j)) > 0 Then

					$delete_task = False
					ExitLoop
				EndIf
			Next

			if $delete_task = True Then

				_GUICtrlListView_DeleteItem($add_time_entry_task_listview, $i)

				if ($i + 1) > _GUICtrlListView_GetItemCount($add_time_entry_task_listview) Then ExitLoop
				$i = $i - 1
			EndIf

		Next

		_GUICtrlListView_EndUpdate($add_time_entry_task_listview)

		_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, 0, true, False)
		;FilterTask()
	EndIf

EndFunc

Func UnFilterTask()

	Local $selected_project = _GUICtrlListView_GetItemText($add_time_entry_project_listview, Number(_GUICtrlListView_GetSelectedIndices($add_time_entry_project_listview)), 0)
	_GUICtrlListView_DeleteAllItems($add_time_entry_task_listview)
	Local $task_names = $timesheet_project_assignments_dict.Item($selected_project)
	Local $task_name_arr = _StringSplit2d($task_names, "|")

	if UBound($task_name_arr) > 0 Then

		if StringLen($task_name_arr[UBound($task_name_arr) - 1][0]) < 1 Then _ArrayDelete($task_name_arr, UBound($task_name_arr) - 1)
		_GUICtrlListView_BeginUpdate($add_time_entry_task_listview)
		_GUICtrlListView_AddArray($add_time_entry_task_listview, $task_name_arr)
		_GUICtrlListView_EndUpdate($add_time_entry_task_listview)
		_GUICtrlListView_SetItemSelected($add_time_entry_task_listview, 0, true, False)
	EndIf


EndFunc

Func UpdateTotalHours()

	Local $hours_day_of_week[7] = [0, 0, 0, 0, 0, 0, 0]
	Local $week_total_hours = 0

	for $i = 0 to (_GUICtrlListView_GetItemCount($timesheet_listview) - 1)

		Local $group_id = _GUICtrlListView_GetItemGroupID($timesheet_listview, $i)
		Local $hours = HourAndMinutesToHours(_GUICtrlListView_GetItemText($timesheet_listview, $i, 3))
		$hours_day_of_week[$group_id - 1] = $hours_day_of_week[$group_id - 1] + $hours
		$week_total_hours = $week_total_hours + $hours
	Next

	for $i = 0 to 6

		Local $group_info = _GUICtrlListView_GetGroupInfo($timesheet_listview, $i + 1)
		_GUICtrlListView_SetGroupInfo($timesheet_listview, $i + 1, StringRegExpReplace($group_info[0], " = .*", " = " & HoursToHourAndMinutes($hours_day_of_week[$i], True)), 2)
	Next

	GUICtrlSetData($timesheet_week_total_label, "Week Total = " & HoursToHourAndMinutes($week_total_hours, True))
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $week_total_hours = ' & $week_total_hours & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

EndFunc

Func RefreshTimesheet($week_start_date)


	_GUICtrlListView_DeleteAllItems($timesheet_listview)
	Local $this_sunday_date = _DateAdd('d', 6, $week_start_date)


	GUICtrlStatusInput_SetText($status_input, "Please Wait. Getting your Harvest times ...")
	$iPID = Run('curl -k https://api.harvestapp.com/v2/time_entries?from=' & StringReplace($week_start_date, "/", "-") & '&to=' & StringReplace($this_sunday_date, "/", "-") & ' -H "Authorization: Bearer ' & GUICtrlRead($harvest_access_token_input) & '" -H "Harvest-Account-Id: ' & GUICtrlRead($harvest_account_id_input) & '" -H "User-Agent: MyApp (yourname@example.com)"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	Local $json = StdoutRead($iPID)
	GUICtrlStatusInput_SetText($status_input, "")

	Local $decoded_json = Json_Decode($json)
	_GUICtrlListView_BeginUpdate($timesheet_listview)
	Local $times_exist_for_day[7] = [False, False, False, False, False, False, False]
	Local $hours_day_of_week[7] = [0, 0, 0, 0, 0, 0, 0]

	for $i = 99 to 0 step -1

		Local $spent_date = Json_Get($decoded_json, '.time_entries[' & $i & '].spent_date')

		if StringLen($spent_date) > 0 Then

			Local $spent_date_day_to_week_index = _DateToDayOfWeek(StringLeft($spent_date, 4), StringMid($spent_date, 6, 2), StringRight($spent_date, 2))
			$spent_date = _DateDayOfWeek($spent_date_day_to_week_index)
			Local $project = Json_Get($decoded_json, '.time_entries[' & $i & '].project.name')
			Local $task = Json_Get($decoded_json, '.time_entries[' & $i & '].task.name')
			Local $notes = Json_Get($decoded_json, '.time_entries[' & $i & '].notes')
			Local $hours = Json_Get($decoded_json, '.time_entries[' & $i & '].hours')
			; $hours = $hours + 0.01	; Harvest is storing time to only 2 decimal places so need to add an additional 0.01 hrs convert correctly to hours and minutes
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours = ' & $hours & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $id = Json_Get($decoded_json, '.time_entries[' & $i & '].id')

			Local $index = _GUICtrlListView_AddItem($timesheet_listview, $project)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, $task, 1)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, $notes, 2)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, HoursToHourAndMinutes($hours + 0.01), 3)	; Adding an extra 0.01 hrs because Harvest is storing time to only 2 decimal places
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, $id, 4)
			_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, $spent_date_day_to_week_index - 1)
			$times_exist_for_day[$spent_date_day_to_week_index - 2] = True
			$hours_day_of_week[$spent_date_day_to_week_index - 2] = $hours_day_of_week[$spent_date_day_to_week_index - 2] + $hours
		EndIf
	Next

	for $i = 0 to (UBound($times_exist_for_day) - 1)

		if $times_exist_for_day[$i] = False Then

			Local $index = _GUICtrlListView_AddItem($timesheet_listview, "<click here then add button above>")
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 1)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 2)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 3)
			_GUICtrlListView_AddSubItem($timesheet_listview, $index, "", 4)
			_GUICtrlListView_SetItemGroupID($timesheet_listview, $index, $i + 1)
		EndIf
	Next

	$date_part = StringSplit(_DateAdd('d', 0, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 1, "Mon " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[0] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 1, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 2, "Tue " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[1] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 2, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 3, "Wed " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[2] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 3, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 4, "Thu " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[3] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 4, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 5, "Fri " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[4] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 5, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 6, "Sat " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[5] + 0.01), 2)
	$date_part = StringSplit(_DateAdd('d', 6, $week_start_date), "/", 3)
	_GUICtrlListView_SetGroupInfo($timesheet_listview, 7, "Sun " & $date_part[2] & " " & _DateToMonth($date_part[1], $DMW_SHORTNAME) & " = " & HoursToHourAndMinutes($hours_day_of_week[6] + 0.01), 2)

	GUICtrlSetData($timesheet_week_total_label, "Week Total = " & HoursToHourAndMinutes($hours_day_of_week[0] + $hours_day_of_week[1] + $hours_day_of_week[2] + $hours_day_of_week[3] + $hours_day_of_week[4] + $hours_day_of_week[5] + $hours_day_of_week[6] + 0.01, True))

	$r = $hours_day_of_week[0] + $hours_day_of_week[1] + $hours_day_of_week[2] + $hours_day_of_week[3] + $hours_day_of_week[4] + $hours_day_of_week[5] + $hours_day_of_week[6]
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[6] = ' & $hours_day_of_week[6] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[5] = ' & $hours_day_of_week[5] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[4] = ' & $hours_day_of_week[4] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[3] = ' & $hours_day_of_week[3] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[2] = ' & $hours_day_of_week[2] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[1] = ' & $hours_day_of_week[1] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $hours_day_of_week[0] = ' & $hours_day_of_week[0] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $r = ' & $r & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	_GUICtrlListView_EndUpdate($timesheet_listview)

	; Jira

	if GUICtrlRead($timesheet_sync_to_jira_checkbox) = $GUI_CHECKED Then

		; process estimate updates to Jira tickets

		$timesheet_listview_item_index = GUICtrlListView_GetIndexOrdered($timesheet_listview)
		Local $week_starting = GUICtrlRead($timesheet_week_combo)
		Local $week_starting_part = StringSplit($week_starting, "/", 3)
		Local $monday_millseconds = (_DateDiff("s","1970/01/01 00:00:00", $week_starting_part[2] & "/" & $week_starting_part[1] & "/" & $week_starting_part[0] & " 00:00:00") * 1000) - 1000

		; get an array of all the jira tickets mentioned in the timesheet and remove all worklogs for this current user for this timesheet period

		Local $jira_ticket_in_timesheet[0]

		for $i = 0 to (UBound($timesheet_listview_item_index) - 1)

			Local $note = _GUICtrlListView_GetItemText($timesheet_listview, $timesheet_listview_item_index[$i], 2)
			Local $arr = StringRegExp($note, "(QA-\d+).* est=(\dd)", 1)

			if @error = 0 Then

				if _ArraySearch($jira_ticket_in_timesheet, $arr[0]) < 0 Then _ArrayAdd($jira_ticket_in_timesheet, $arr[0])
			Else

				Local $arr = StringRegExp($note, "(QA-\d+).* done=(\d+)%", 1)

				if @error = 0 Then

					if _ArraySearch($jira_ticket_in_timesheet, $arr[0]) < 0 Then _ArrayAdd($jira_ticket_in_timesheet, $arr[0])
				EndIf
			EndIf
		Next


;		Local $jira_ticket_time_spent_seconds_at_timesheet_start[UBound($jira_ticket_in_timesheet)]

		Local $jira_ticket_decoded_json[UBound($jira_ticket_in_timesheet)]

		Local $jira_ticket_worklogs[UBound($jira_ticket_in_timesheet)]
		; [n][n] = started
		; [n][1] = comment
		; [n][2] = id
		; [n][3] = time spent

		for $i = 0 to (UBound($jira_ticket_in_timesheet) - 1)


			; get the estimate, time spent and worklogs of the jira ticket

			GUICtrlStatusInput_SetText($status_input, "Please Wait. Getting Jira ticket " & $jira_ticket_in_timesheet[$i] & " ...")
			$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/' & $jira_ticket_in_timesheet[$i] & ' -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			ProcessWaitClose($iPID)
			Local $json = StdoutRead($iPID)
			GUICtrlStatusInput_SetText($status_input, "")
			$jira_ticket_decoded_json[$i] = Json_Decode($json)

			$jira_ticket_worklogs[$i] = ""

			; determine what the time spent would be for this jira ticket at the start of this timesheet (Monday), by removing all time spent currently in this timesheet period

;			$jira_ticket_time_spent_seconds_at_timesheet_start[$i] = Json_Get($jira_ticket_decoded_json[$i], '.fields.timetracking.timeSpentSeconds')
			Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.total')

			for $worklog_index = ($total_worklogs - 1) to 0 step -1

				Local $worklog_started2 = StringLeft(Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].started'), 10)
				Local $worklog_time_spent_seconds = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].timeSpentSeconds')

;				if _DateDiff("D", $week_starting_part[2] & "/" & $week_starting_part[1] & "/" & $week_starting_part[0], StringReplace($worklog_started2, "-", "/")) >= 0 then

;					$jira_ticket_time_spent_seconds_at_timesheet_start[$i] = $jira_ticket_time_spent_seconds_at_timesheet_start[$i] - $worklog_time_spent_seconds
;				EndIf

				Local $worklog_comment = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].comment.content[0].content[0].text')
				Local $worklog_id = StringLeft(Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].id'), 10)

				if StringLen($jira_ticket_worklogs[$i]) > 0 Then $jira_ticket_worklogs[$i] = $jira_ticket_worklogs[$i] & chr(30)

				$jira_ticket_worklogs[$i] = $jira_ticket_worklogs[$i] & $worklog_started2 & chr(29) & $worklog_comment & chr(29) & $worklog_id & chr(29) & $worklog_time_spent_seconds

			Next


		Next

		for $i = 0 to (UBound($timesheet_listview_item_index) - 1)

			$selected_group_info = _GUICtrlListView_GetGroupInfo($timesheet_listview, _GUICtrlListView_GetItemGroupID($timesheet_listview, $timesheet_listview_item_index[$i]))
			Local $selected_timesheet_date_part = StringSplit($selected_group_info[0], " = ", 3)
			$selected_timesheet_date = $selected_timesheet_date_part[0]
			Local $selected_timesheet_date_part2 = StringSplit($selected_timesheet_date, " ", 3)
			Local $note = _GUICtrlListView_GetItemText($timesheet_listview, $timesheet_listview_item_index[$i], 2)

			; if the harvest time entry is requesting the estimate of a Jira ticket be updated

			Local $arr = StringRegExp($note, "(QA-\d\d\d\d).* est=(\dd)", 1)

			if @error = 0 Then

				Local $jira_key = $arr[0]
				Local $new_estimate = $arr[1]

				; check if this estimate update is already a worklog in the Jira ticket

				Local $worklog_found = False
				Local $jira_ticket_index = _ArraySearch($jira_ticket_in_timesheet, $jira_key)

				if $jira_ticket_index > -1 Then

					Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.total')

					for $worklog_index = 0 to ($total_worklogs - 1)

						Local $worklog_comment = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].comment.content[0].content[0].text')
						Local $worklog_started = StringLeft(Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].started'), 10)

						if StringCompare($worklog_comment, "original estimate updated to " & $new_estimate) = 0 and StringCompare($worklog_started, @YEAR & "-" & _ConvertMonth($selected_timesheet_date_part2[2]) & "-" & $selected_timesheet_date_part2[1]) = 0 Then

							$worklog_found = True

							; clear the worklog started so we know later not to process this worklog again
							;Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].id', "")
							;Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].comment.content[0].content[0].text', "")
							Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].started', "")

							ExitLoop

						EndIf
					Next

					; if a worklog for this update is not found then create the worklog

					if $worklog_found = False Then

						; update the original estimate of the jira ticket

						GUICtrlStatusInput_SetText($status_input, "Please Wait. Updating the estimate to " & $new_estimate & " for ticket " & $jira_key & " ...")
						$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/' & $jira_key & ' -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X PUT -d "{\"update\": {\"timetracking\": [{\"edit\": {\"originalEstimate\": \"' & $new_estimate & '\"}}]}}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
						ProcessWaitClose($iPID)
						GUICtrlStatusInput_SetText($status_input, "")

						; add a work log to the jira ticket to indicate the original estimate was updated

						GUICtrlStatusInput_SetText($status_input, "Please Wait. Adding a worklog to ticket " & $jira_key & " to indicate the estimate was updated ...")
						$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/QA-3285/worklog -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X POST -d "{\"timeSpentSeconds\": 0, \"comment\": {\"type\": \"doc\", \"version\": 1, \"content\": [{\"type\": \"paragraph\", \"content\": [{\"text\": \"original estimate updated to ' & $new_estimate & '\", \"type\": \"text\"}]}]}, \"started\": \"' & @YEAR & "-" & _ConvertMonth($selected_timesheet_date_part2[2]) & "-" & $selected_timesheet_date_part2[1] & 'T00:00:00.000+0000\"}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
						ProcessWaitClose($iPID)
						GUICtrlStatusInput_SetText($status_input, "")

					EndIf

				EndIf

			EndIf

			Local $arr = StringRegExp($note, "(QA-\d+).* done=(\d+)%", 1)

			if @error = 0 Then

				Local $jira_key = $arr[0]
				Local $done_pcnt = $arr[1]

				; check if this worklog is already in the Jira ticket

				Local $worklog_found = False
				Local $jira_ticket_index = _ArraySearch($jira_ticket_in_timesheet, $jira_key)

				if $jira_ticket_index > -1 Then

					Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.total')

					for $worklog_index = 0 to ($total_worklogs - 1)

						Local $worklog_comment = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].comment.content[0].content[0].text')
						Local $worklog_started = StringLeft(Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].started'), 10)

						if StringCompare($worklog_comment, $note) = 0 and StringCompare($worklog_started, @YEAR & "-" & _ConvertMonth($selected_timesheet_date_part2[2]) & "-" & $selected_timesheet_date_part2[1]) = 0 Then

							$worklog_found = True

							; clear the worklog started so we know later not to process this worklog again
							;Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].id', "")
							;Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].comment.content[0].content[0].text', "")
							Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].started', "")

							ExitLoop
						EndIf
					Next

					; if a worklog for this update is not found then create the worklog

					if $worklog_found = False Then

						Local $ticket_time_spent_seconds = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.timetracking.timeSpentSeconds')
						Local $original_estimate = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.timetracking.originalEstimate')

						; remove all time spent after the date time of this worklog, to arrive at a time spent for the date time of the worklog

						Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.total')

						for $worklog_index = ($total_worklogs - 1) to 0 step -1

							Local $worklog_started = StringLeft(Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].started'), 10)
							Local $worklog_time_spent_seconds = Json_Get($jira_ticket_decoded_json[$jira_ticket_index], '.fields.worklog.worklogs[' & $worklog_index & '].timeSpentSeconds')

							if _DateDiff("D", @YEAR & "/" & _ConvertMonth($selected_timesheet_date_part2[2]) & "/" & $selected_timesheet_date_part2[1], StringReplace($worklog_started, "-", "/")) <= 0 then ExitLoop

							$ticket_time_spent_seconds = $ticket_time_spent_seconds - $worklog_time_spent_seconds
						Next

						; convert Jira week-day-hour estimate back to seconds
						$original_estimate = JiraWeekDayHourToSeconds($original_estimate)

						; calculate how many seconds of work has been done overall
						$ticket_done_seconds = $original_estimate * ($done_pcnt / 100)

						; substract the number of seconds already spent to arrive at the time that needs to be spent in this work log (to match the percent done above)
						Local $worklog_time_spent_seconds = $ticket_done_seconds - $ticket_time_spent_seconds

						if $worklog_time_spent_seconds > 0 Then

							; add work log for jira ticket

							GUICtrlStatusInput_SetText($status_input, "Please Wait. Adding a new worklog to Jira ticket " & $jira_key & " ...")
							$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/' & $jira_key & '/worklog -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X POST -d "{\"timeSpentSeconds\": ' & $worklog_time_spent_seconds & ', \"comment\": {\"type\": \"doc\", \"version\": 1, \"content\": [{\"type\": \"paragraph\", \"content\": [{\"text\": \"' & $note & '\", \"type\": \"text\"}]}]}, \"started\": \"' & @YEAR & "-" & _ConvertMonth($selected_timesheet_date_part2[2]) & "-" & $selected_timesheet_date_part2[1] & 'T00:00:00.000+0000\"}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
							ProcessWaitClose($iPID)
							GUICtrlStatusInput_SetText($status_input, "")
							Local $json = StdoutRead($iPID)
							Local $decoded_json = Json_Decode($json)

							Json_Put($jira_ticket_decoded_json[$jira_ticket_index], '.fields.timetracking.timeSpentSeconds', Number($ticket_time_spent_seconds) + $worklog_time_spent_seconds)

							if StringLen($jira_ticket_worklogs[$jira_ticket_index]) > 0 Then $jira_ticket_worklogs[$jira_ticket_index] = $jira_ticket_worklogs[$jira_ticket_index] & chr(30)

							$jira_ticket_worklogs[$jira_ticket_index] = $jira_ticket_worklogs[$jira_ticket_index] & @YEAR & "-" & _ConvertMonth($selected_timesheet_date_part2[2]) & "-" & $selected_timesheet_date_part2[1] & chr(29) & $note & chr(29) & Json_Get($decoded_json, '.id') & chr(29) & $worklog_time_spent_seconds

						EndIf
					EndIf
				EndIf
			EndIf

		Next


		; iterate through all other worklogs in the affected Jira tickets and delete them

		for $i = 0 to (UBound($jira_ticket_in_timesheet) - 1)

			Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.total')

			for $worklog_index = 0 to ($total_worklogs - 1)

				Local $worklog_id = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].id')
				Local $worklog_started = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].started')

				if StringLen($worklog_started) > 0 Then

					; delete the worklog

					GUICtrlStatusInput_SetText($status_input, "Please Wait. Deleting Jira worklog " & $worklog_id & " for ticket " & $jira_ticket_in_timesheet[$i] & " ...")
					$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/' & $jira_ticket_in_timesheet[$i] & '/worklog/' & $worklog_id & ' -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -X DELETE', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
					ProcessWaitClose($iPID)
					GUICtrlStatusInput_SetText($status_input, "")
				EndIf
			Next
		Next

		; iterate through all worklogs in the affected Jira tickets and check if the time spent is correct, and if not update

		for $i = 0 to (UBound($jira_ticket_in_timesheet) - 1)

			Local $original_estimate = Json_Get($jira_ticket_decoded_json[$i], '.fields.timetracking.originalEstimate')

			; convert Jira week-day-hour estimate back to seconds
			$original_estimate = JiraWeekDayHourToSeconds($original_estimate)

			Local $ticket_time_spent_seconds = 0

			Local $worklog_arr = StringSplit($jira_ticket_worklogs[$i], chr(30), 3)
			_ArraySort($worklog_arr)

			for $j = 0 to (UBound($worklog_arr) - 1)

				Local $worklog_part_arr = StringSplit($worklog_arr[$j], chr(29), 3)
				Local $worklog_started = $worklog_part_arr[0]
				Local $worklog_comment = $worklog_part_arr[1]
				Local $worklog_id = $worklog_part_arr[2]
				Local $worklog_time_spent_seconds = $worklog_part_arr[3]
				Local $total_worklogs = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.total')
				Local $worklog_found = False

				for $worklog_index = ($total_worklogs - 1) to 0 step -1

					Local $worklog_id2 = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].id')
					Local $worklog_time_spent_seconds2 = Json_Get($jira_ticket_decoded_json[$i], '.fields.worklog.worklogs[' & $worklog_index & '].timeSpentSeconds')

					if Number($worklog_id2) = Number($worklog_id) Then

						$worklog_found = True
						Local $arr = StringRegExp($worklog_comment, "done=(\d+)%", 1)

						if @error = 0 Then

							Local $done_pcnt = $arr[0]

							; calculate how many seconds of work has been done overall
							$ticket_done_seconds = $original_estimate * ($done_pcnt / 100)

							; substract the number of seconds already spent to arrive at the time that needs to be spent in this work log (to match the percent done above)
							Local $worklog_time_spent_seconds3 = $ticket_done_seconds - $ticket_time_spent_seconds

							if Number($worklog_time_spent_seconds3) <> Number($worklog_time_spent_seconds2) Then

								; update the time spent for this worklog

								GUICtrlStatusInput_SetText($status_input, "Please Wait. Updating time spent for worklog id " & $worklog_id & " in Jira ticket " & $jira_ticket_in_timesheet[$i] & " ...")
								$iPID = Run('curl -k https://janisoncls.atlassian.net/rest/api/3/issue/' & $jira_ticket_in_timesheet[$i] & '/worklog/' & $worklog_id & ' -u ' & GUICtrlRead($jira_username_input) & ':' & GUICtrlRead($jira_api_token_input) & ' -H "Accept: application/json" -H "Content-Type: application/json" -X PUT -d "{\"timeSpentSeconds\": ' & $worklog_time_spent_seconds3 & '}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
								ProcessWaitClose($iPID)
								GUICtrlStatusInput_SetText($status_input, "")
							EndIf

							$ticket_time_spent_seconds = $ticket_time_spent_seconds + $worklog_time_spent_seconds3
						EndIf

						ExitLoop
					EndIf
				Next

				; if a newly added worklog in this refresh

				if $worklog_found = False Then

					Local $arr = StringRegExp($worklog_comment, "done=(\d+)%", 1)

					if @error = 0 Then

						Local $done_pcnt = $arr[0]

						; calculate how many seconds of work has been done overall
						$ticket_done_seconds = $original_estimate * ($done_pcnt / 100)

						; substract the number of seconds already spent to arrive at the time that needs to be spent in this work log (to match the percent done above)
						Local $worklog_time_spent_seconds3 = $ticket_done_seconds - $ticket_time_spent_seconds

						$ticket_time_spent_seconds = $ticket_time_spent_seconds + $worklog_time_spent_seconds3
					EndIf
				EndIf

			Next

		Next




	EndIf

EndFunc

Func TimeEntrySetState($state)

	if $state = $GUI_DISABLE Then

		GUISetCursor(15, 0, $add_time_entry_gui)
	Else

		GUISetCursor(2, 0, $add_time_entry_gui)
	EndIf

	GUICtrlSetState($add_time_entry_favourites_combo, 					$state)
	GUICtrlSetState($add_time_entry_favourites_add_button, 				$state)
	GUICtrlSetState($add_time_entry_favourites_delete_button, 			$state)
	GUICtrlSetState($add_time_entry_project_listview, 					$state)
	GUICtrlSetState($add_time_entry_project_filters_list, 				$state)
	GUICtrlSetState($add_time_entry_project_filters_add_button, 		$state)
	GUICtrlSetState($add_time_entry_project_filters_delete_button, 		$state)
	GUICtrlSetState($add_time_entry_project_filters_enable_checkbox, 	$state)
	GUICtrlSetState($add_time_entry_task_listview, 						$state)
	GUICtrlSetState($add_time_entry_task_filters_list, 					$state)
	GUICtrlSetState($add_time_entry_task_filters_add_button, 			$state)
	GUICtrlSetState($add_time_entry_task_filters_delete_button, 		$state)
	GUICtrlSetState($add_time_entry_task_filters_enable_checkbox, 		$state)
	GUICtrlSetState($add_time_entry_notes_input,				 		$state)
	GUICtrlSetState($add_time_entry_hour_input_radio,			 		$state)
	GUICtrlSetState($add_time_entry_hour_input,					 		$state)
	GUICtrlSetState($add_time_entry_half_hour_radio,					$state)
	GUICtrlSetState($add_time_entry_one_hour_radio,						$state)
	GUICtrlSetState($add_time_entry_one_half_hour_radio,				$state)
	GUICtrlSetState($add_time_entry_two_hour_radio,						$state)
	GUICtrlSetState($add_time_entry_two_half_hour_radio,				$state)
	GUICtrlSetState($add_time_entry_three_hour_radio,					$state)
	GUICtrlSetState($add_time_entry_three_half_hour_radio,				$state)
	GUICtrlSetState($add_time_entry_four_hour_radio,					$state)
	GUICtrlSetState($add_time_entry_four_half_hour_radio,				$state)
	GUICtrlSetState($add_time_entry_five_hour_radio,					$state)
	GUICtrlSetState($add_time_entry_save_button,						$state)

EndFunc

