
; #INCLUDES# =========================================================================================================
#include-once
#include <GUIConstants.au3>
#include <File.au3>
#include <GuiComboBox.au3>
#include <GuiToolTip.au3>
#include <GuiListView.au3>
#include <GuiListBox.au3>
#include <GuiTab.au3>
#include <GDIPlus.au3>
#include <GuiRichEdit.au3>
#Include <WinAPIEx.au3>
#Include <GUIConstantsEx.au3>
#Include <GUIMenu.au3>
#include <GuiButton.au3>
#include <Date.au3>
#include "DTC.au3"


; #GLOBAL VARIABLES# =================================================================================================

Global $app_name = "HJM Link"
Global $app_data_dir = @AppDataDir & "\" & $app_name
Global $ini_filename = $app_data_dir & "\" & $app_name & ".ini"
Global $log_filename = $app_data_dir & "\" & $app_name & ".log"


Global $local_path = "F:\RetroPie"
Global $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
Global $sDrive1 = "", $sDir1 = "", $sFileName1 = "", $sExtension1 = ""
Global $sDrive2 = "", $sDir2 = "", $sFileName2 = "", $sExtension2 = ""
Global $alphanumeric_arr[36] = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
Global $iStyle = BitOR($TVS_EDITLABELS, $TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES)
Global $current_gui
Global $result = 1
Global $iPID
Global $metronome_token
Global $timesheet_project_id_dict = ObjCreate("Scripting.Dictionary")
Global $timesheet_project_assignments_dict = ObjCreate("Scripting.Dictionary")

; GUIs

; Main gui

Global $main_gui_width = 840
Global $main_gui_height = 720

Global $tooltip = _GUIToolTip_Create(0) ; default style tooltip
_GUIToolTip_SetMaxTipWidth($tooltip, 300)

Global $main_gui
Global $status_input

; Tabs

Global $tab

; Settings tab

Global $settings_save_button
Global $harvest_accound_id_label
Global $harvest_account_id_input
Global $harvest_access_token_label
Global $harvest_access_token_input
Global $metronome_email_input
Global $metronome_password_input

; Harvest tab

Global $timesheet_week_total_label
Global $timesheet_listview
Global $timesheet_refresh_button
Global $timesheet_add_button
Global $timesheet_edit_button
Global $timesheet_delete_button
Global $timesheet_tmp_button
Global $timesheet_week_combo
Global $timesheet_this_week_button

Global $add_time_entry_gui
Global $add_time_entry_favourites_combo
Global $add_time_entry_favourites_add_button
Global $add_time_entry_favourites_delete_button
Global $add_time_entry_project_listview
Global $add_time_entry_project_filters_list
Global $add_time_entry_project_filters_add_button
Global $add_time_entry_project_filters_delete_button
Global $add_time_entry_project_filters_enable_checkbox
Global $add_time_entry_task_listview
Global $add_time_entry_task_filters_list
Global $add_time_entry_task_filters_add_button
Global $add_time_entry_task_filters_delete_button
Global $add_time_entry_task_filters_enable_checkbox
Global $add_time_entry_notes_input
Global $add_time_entry_hour_input_radio
Global $add_time_entry_hour_input
Global $add_time_entry_half_hour_radio
Global $add_time_entry_one_hour_radio
Global $add_time_entry_one_half_hour_radio
Global $add_time_entry_two_hour_radio
Global $add_time_entry_two_half_hour_radio
Global $add_time_entry_three_hour_radio
Global $add_time_entry_three_half_hour_radio
Global $add_time_entry_four_hour_radio
Global $add_time_entry_four_half_hour_radio
Global $add_time_entry_five_hour_radio
Global $add_time_entry_save_button
Global $add_time_entry_cancel_button
Global $add_time_entry_status_input



; Jira tab


; Metronome tab

Global $metronome_refresh_button
Global $metronome_periods_listview
Global $metronome_quarterly_priorities_listview
Global $metronome_action_items_listview
Global $metronome_action_items_add_button
Global $metronome_action_items_delete_button


Func HJM_Link_Startup()

;	_GDIPlus_Startup()


	; Create the app data folder

	if FileExists($app_data_dir) = False Then

		DirCreate($app_data_dir)
	EndIf


	; Erase the log

	if FileExists($log_filename) = true Then

		FileDelete($log_filename)
	EndIf

EndFunc


Func HJM_Link_Shutdown()

;	_GDIPlus_ShutDown ()
EndFunc






Func _TipDisplayLen($time=5000)
    Local $tiphandles, $i
    $tiphandles=WinList('[CLASS:tooltips_class32]')
    for $i=1 to $tiphandles[0][0]
        If WinGetProcess($tiphandles[$i][1]) = @AutoItPID Then _SendMessage($tiphandles[$i][1],0x0403,2,$time)
    Next
EndFunc



Func GUICtrlSetImagePNG($ctrl, $png_path)

	Local $hImage = _GDIPlus_ImageLoadFromFile($png_path)
	Local $hCLSID = _GDIPlus_EncodersGetCLSID("BMP")
	_GDIPlus_ImageSaveToFileEx($hImage, @TempDir & "\test.bmp", $hCLSID)
	GUICtrlSetImage($ctrl, @TempDir & "\test.bmp")
	_GDIPlus_ImageDispose($hImage)
EndFunc

Func GUICtrlUnselect($ctrl)

	Local $aSel = _GUICtrlListBox_GetSelItems($ctrl)

	For $i = 1 To $aSel[0]

		_GUICtrlListBox_SetSel($ctrl, $aSel[$i], False)
	Next
EndFunc

Func Listbox_ItemMoveUD($hLB_ID, $iDir = -1)
    ;Listbox_ItemMoveUD - Up/Down  Move Multi/Single item in a ListBox
    ;$iDir: -1 up, 1 down
    ;Return values -1 nothing to do, 0 nothing moved, >0 performed moves
    Local $iCur, $iNxt, $aCou, $aSel, $i, $m = 0, $y, $slb = 0 ;Current, next, Count, Selection, loop , movecount

    $aSel = _GUICtrlListBox_GetSelItems($hLB_ID) ;Put selected items in an array
    $aCou = _GUICtrlListBox_GetCount($hLB_ID) ;Get total item count of the listbox

    If $aSel[0] = 0 Then
        $y = _GUICtrlListBox_GetCurSel($hLB_ID)
        If $y > -1 Then
            _ArrayAdd($aSel, $y)
            $aSel[0] = 1
            $slb = 1
        EndIf
    EndIf

    ;WinSetTitle($hGUI, "", $aSel[0])                   ;Debugging info

    Select
        Case $iDir = -1 ;Move Up
            For $i = 1 To $aSel[0]
                If $aSel[$i] > 0 Then
                    $iNxt = _GUICtrlListBox_GetText($hLB_ID, $aSel[$i] - 1) ;Save the selection index - 1 text
                    _GUICtrlListBox_ReplaceString($hLB_ID, $aSel[$i] - 1, _GUICtrlListBox_GetText($hLB_ID, $aSel[$i])) ;Replace the index-1 text with the index text
                    _GUICtrlListBox_ReplaceString($hLB_ID, $aSel[$i], $iNxt) ;Replace the selection with the saved var
                    $m = $m + 1
                EndIf
            Next
            For $i = 1 To $aSel[0] ;Restore the selections after moving
                If $aSel[$i] > 0 Then
                    If $slb = 0 Then
                        _GUICtrlListBox_SetSel($hLB_ID, $aSel[$i] - 1, 1)
                    Else
                        _GUICtrlListBox_SetCurSel($hLB_ID, $aSel[$i] - 1)
                    EndIf
                EndIf
            Next
            Return $m
        Case $iDir = 1 ;Move Down
            If $aSel[0] > 0 Then
                For $i = $aSel[0] To 1 Step -1
                    If $aSel[$i] < $aCou - 1 Then
                        $iNxt = _GUICtrlListBox_GetText($hLB_ID, $aSel[$i] + 1)
                        _GUICtrlListBox_ReplaceString($hLB_ID, $aSel[$i] + 1, _GUICtrlListBox_GetText($hLB_ID, $aSel[$i]))
                        _GUICtrlListBox_ReplaceString($hLB_ID, $aSel[$i], $iNxt)
                        $m = $m + 1
                    EndIf
                Next
            EndIf
            For $i = $aSel[0] To 1 Step -1 ;Restore the selections after moving
                If $aSel[$i] < $aCou - 1 Then
                    If $slb = 0 Then
                        _GUICtrlListBox_SetSel($hLB_ID, $aSel[$i] + 1, 1)
                    Else
                        _GUICtrlListBox_SetCurSel($hLB_ID, $aSel[$i] + 1)
                    EndIf
                EndIf
            Next
            Return $m
    EndSelect
    Return -1
EndFunc   ;==>Listbox_ItemMoveUD



Func DirCreateSafe($path)

	If StringLen($path) > 0 and FileExists($path) = False Then

		$result = DirCreate($path)

		if $result <> 1 Then

			MsgBox(262144, $app_name, "failed to create dir " & $path)
			Exit
		EndIf
	EndIf
EndFunc


;Func GUISetImage($gui, $ctrl, $image_filepath)

;	Local $pic_size = GetImageSize($image_filepath, True)
;	WinMove($gui, "", (@DesktopWidth/2)-$pic_size[0]/2, (@DesktopHeight/2)-$pic_size[1]/2, $pic_size[0], $pic_size[1])
;	GUICtrlSetPos($ctrl, 0, 0, $pic_size[0], $pic_size[1])

;	if StringInStr($image_filepath, ".png") > 0 Then

;		GUICtrlSetImagePNG($ctrl, $image_filepath)
;	Else

;		GUICtrlSetImage($ctrl, $image_filepath)
;	EndIf

;EndFunc



Func _StringSplit2d($str, $delimiter)

    ; #FUNCTION# ======================================================================================
    ; Name ................:    _DBG_StringSplit2D($str,$delimiter)
    ; Description .........:    Create 2d array from delimited string
    ; Syntax ..............:    _DBG_StringSplit2D($str, $delimiter)
    ; Parameters ..........:    $str        - EOL (@CR, @LF or @CRLF) delimited string to split
    ;                           $delimiter  - Delimter for columns
    ;                           $showtiming - Display time to create 2D array to the console
    ; Return values .......:    2D array
    ; Author ..............:    kylomas
    ; =================================================================================================

    Local $a1 = StringRegExp($str, '(.*?)(?:\R|$)', 3)

    ;ReDim $a1[UBound($a1) - 1]

    Local $rows = UBound($a1), $cols = 0

    ; determine max number of columns
    For $i = 0 To UBound($a1) - 1
        StringReplace($a1[$i], $delimiter, '')
        $cols = (@extended > $cols ? @extended : $cols)
    Next

    ; define and populate array
    Local $aRET[$rows][$cols + 1]

    For $i = 0 To UBound($a1) - 1
        $a2 = StringSplit($a1[$i], $delimiter, 3)
        For $j = 0 To UBound($a2) - 1
            $aRET[$i][$j] = $a2[$j]
        Next
    Next

    Return $aRET

EndFunc   ;==>_DBG_StringSplit2d



Func MainGUICreate(ByRef $tab, $tab_left, $tab_top, $tab_width, $tab_height)

	Local $gui = GUICreate($app_name & " - Main", $main_gui_width, $main_gui_height, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	$tab = GUICtrlCreateTabEx($tab_left, $tab_top, $tab_width, $tab_height)
	$current_gui = $gui

	Return $gui

EndFunc

Func ChildGUICreate($title, $width, $height, $parent_gui)

	Local $gui = GUICreate($title, $width, $height, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU), $WS_EX_MDICHILD, $parent_gui)
	$current_gui = $gui
	Return $gui

EndFunc

Func GUICtrlCreateComboFromDict($value_dictionary = Null, $left = -1, $top = -1, $width = 80, $height = 20, $resizing = -1)

	local $ctrl = GUICtrlCreateCombo("", $left, $top, $width, $height, BitOR($CBS_DROPDOWNLIST, $CBS_DROPDOWN, $CBS_AUTOHSCROLL, $WS_VSCROLL))
	GUICtrlSetResizing(-1, $GUI_DOCKALL)

	if $value_dictionary <> Null Then

		For $vKey In $value_dictionary

		   _GUICtrlComboBox_AddString($ctrl, $vKey)
		Next
	EndIf

	_GUICtrlComboBox_SetCurSel($ctrl, 0)

	if $resizing > -1 Then

		GUICtrlSetResizing($ctrl, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateComboFromDictWithLabel(ByRef $label, $label_text = "", $label_left = -1, $label_top = -1, $label_width = 80, $label_height = 20, $label_tooltip_text = "", $value_dictionary = Null, $left = -1, $top = -1, $width = 80, $height = 20)

	if StringLen($label_text) > 0 Then

		$label = GUICtrlCreateLabel($label_text, $label_left, $label_top, $label_width, $label_height)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)

		if StringLen($label_tooltip_text) > 0 Then

			_GUIToolTip_AddTool($tooltip, 0, $label_tooltip_text, GUICtrlGetHandle($label))
		EndIf
	EndIf

	local $ctrl = GUICtrlCreateCombo("", $left, $top, $width, $height, BitOR($CBS_DROPDOWNLIST, $CBS_DROPDOWN, $CBS_AUTOHSCROLL, $WS_VSCROLL))
	GUICtrlSetResizing(-1, $GUI_DOCKALL)

	if $value_dictionary <> Null Then

		For $vKey In $value_dictionary

		   _GUICtrlComboBox_AddString($ctrl, $vKey)
		Next
	EndIf

	_GUICtrlComboBox_SetCurSel($ctrl, 0)
	Return $ctrl

EndFunc


Func GUICtrlCreateComboEx($left, $top, $width, $height, $tooltip_text = "", $resizing = -1, $hide = False)

	local $ctrl = GUICtrlCreateCombo("", $left, $top, $width, $height, BitOR($CBS_DROPDOWNLIST, $CBS_SORT))

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle(-1))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $hide = True Then

		GUICtrlSetState(-1, $GUI_HIDE)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateButtonEx($text, $left, $top, $width, $height, $tooltip_text = "", $resizing = -1, $hide = False)

	local $ctrl = GUICtrlCreateButton($text, $left, $top, $width, $height)

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle(-1))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $hide = True Then

		GUICtrlSetState(-1, $GUI_HIDE)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateImageButton($ico_filename, $left, $top, $width_height, $tooltip_text, $resizing = -1, $hide = False)

	local $ctrl = GUICtrlCreateButton("", $left, $top, $width_height, $width_height, $BS_ICON, $WS_EX_DLGMODALFRAME)
	GUICtrlSetImage(-1, @ScriptDir & "\" & $ico_filename)
	_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle(-1))

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $hide = True Then

		GUICtrlSetState(-1, $GUI_HIDE)
	EndIf

	Return $ctrl

EndFunc




Func GUICtrlCreateTabEx($left, $top, $width, $height )

	local $ctrl = GUICtrlCreateTab($left, $top, $width, $height)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
	Return $ctrl

EndFunc

Func GUICtrlCreateInputWithLabel($input_text, $input_left, $input_top, $input_width, $input_height, ByRef $label, $label_text, $label_left, $label_top, $label_width, $label_height, $label_tooltip_text = "", $label2_text = "", $label2_left = -1, $label2_top = -1, $label2_width = -1, $label2_height = -1)

	if StringLen($label_text) > 0 Then

		$label = GUICtrlCreateLabel($label_text, $label_left, $label_top, $label_width, $label_height)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)

		if StringLen($label_tooltip_text) > 0 Then

			_GUIToolTip_AddTool($tooltip, 0, $label_tooltip_text, GUICtrlGetHandle($label))
		EndIf
	EndIf

	local $input = GUICtrlCreateInput($input_text, $input_left, $input_top, $input_width, $input_height)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)

	if StringLen($label2_text) > 0 Then

		GUICtrlCreateLabel($label2_text, $label2_left, $label2_top, $label2_width, $label2_height)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)
	EndIf

	Return $input

EndFunc

Func GUICtrlCreatePasswordWithLabel($input_text, $input_left, $input_top, $input_width, $input_height, ByRef $label, $label_text, $label_left, $label_top, $label_width, $label_height, $label_tooltip_text = "", $label2_text = "", $label2_left = -1, $label2_top = -1, $label2_width = -1, $label2_height = -1)

	if StringLen($label_text) > 0 Then

		$label = GUICtrlCreateLabel($label_text, $label_left, $label_top, $label_width, $label_height)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)

		if StringLen($label_tooltip_text) > 0 Then

			_GUIToolTip_AddTool($tooltip, 0, $label_tooltip_text, GUICtrlGetHandle($label))
		EndIf
	EndIf

	local $input = GUICtrlCreateInput($input_text, $input_left, $input_top, $input_width, $input_height, BitOR($ES_LEFT, $ES_AUTOHSCROLL, $ES_PASSWORD))
	GUICtrlSetResizing(-1, $GUI_DOCKALL)

	if StringLen($label2_text) > 0 Then

		GUICtrlCreateLabel($label2_text, $label2_left, $label2_top, $label2_width, $label2_height)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)
	EndIf

	Return $input

EndFunc

Func GUICtrlCreateStatusInput($text, $left, $top, $width, $height)

	Local $status_input_header = "{\rtf1\ansi\deff0\readprot\annotprot{\fonttbl {\f0 Normal;}}\fs18 " ; Courier or New Times New Roman or roman or Times New Roman Greek
	Local $status_input_footer = "\line "

	local $input = _GUICtrlRichEdit_Create($current_gui, $status_input_header & $text & $status_input_footer, $left, $top, $width, $height, $ES_READONLY)
;	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT + $GUI_DOCKLEFT + $GUI_DOCKRIGHT)
	GUICtrlSetResizing($input, $GUI_DOCKAUTO)
	_GUICtrlRichEdit_SetEventMask($input, $ENM_LINK)
	_GUICtrlRichEdit_AutoDetectURL($input, True)
	Return $input

EndFunc

Func GUICtrlStatusInput_SetText($input, $text)

	Local $status_input_header = "{\rtf1\ansi\deff0\readprot\annotprot{\fonttbl {\f0 Normal;}}\fs18 " ; Courier or New Times New Roman or roman or Times New Roman Greek
	Local $status_input_footer = "\line "
	_GUICtrlRichEdit_SetText($input, $status_input_header & $text & $status_input_footer)
EndFunc


Func GUICtrlCreateSliderEx($left, $top, $width, $height, $resizing, $max, $min, $value)

	local $ctrl = GUICtrlCreateSlider($left, $top, $width, $height)
	GUICtrlSetResizing(-1, $resizing)
	GUICtrlSetLimit(-1, $max, $min)
	GUICtrlSetData(-1, $value)
	Return $ctrl
EndFunc

Func GUICtrlCreateCheckboxEx($text, $left, $top, $width, $height, $checked = False, $tooltip_text = "", $resizing = $GUI_DOCKALL)

	local $ctrl = GUICtrlCreateCheckbox($text, $left, $top, $width, $height)

	if $checked = True Then

		GUICtrlSetState(-1, $GUI_CHECKED)
	Else

		GUICtrlSetState(-1, $GUI_UNCHECKED)
	EndIf

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle($ctrl))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateLabelEx($text, $left, $top, $width, $height, $tooltip_text = "", $resizing = -1)

	local $ctrl = GUICtrlCreateLabel($text, $left, $top, $width, $height)

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle($ctrl))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateInputEx($text, $left, $top, $width, $height, $tooltip_text = "", $resizing = -1)

	local $ctrl = GUICtrlCreateInput($text, $left, $top, $width, $height)

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle($ctrl))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateRadioEx($text, $left, $top, $width, $height, $checked = False, $tooltip_text = "", $resizing = $GUI_DOCKALL, $hide = False)

	local $ctrl = GUICtrlCreateRadio($text, $left, $top, $width, $height)

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $checked = True Then

		GUICtrlSetState(-1, $GUI_CHECKED)
	EndIf

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle($ctrl))
	EndIf

	if $hide = True Then

		GUICtrlSetState(-1, $GUI_HIDE)
	EndIf

	Return $ctrl

EndFunc


Func GUICtrlCreateSingleSelectList($left, $top, $width, $height, $horizontal_scroll_size = -1, $resizing = -1)

	local $ctrl = GUICtrlCreateList("", $left, $top, $width, $height, BitOR($GUI_SS_DEFAULT_LIST, $WS_HSCROLL))

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $horizontal_scroll_size > -1 Then

		GUICtrlSetLimit(-1, $horizontal_scroll_size)
	EndIf

	Return $ctrl

EndFunc


Func GUICtrlCreateListViewEx($left, $top, $width, $height, $resizing = $GUI_DOCKBORDERS, $col_1_name = Null, $col_1_width = Null, $col_2_name = Null, $col_2_width = Null, $col_3_name = Null, $col_3_width = Null, $col_4_name = Null, $col_4_width = Null, $col_5_name = Null, $col_5_width = Null, $row_1_col_1_value = Null, $row_1_col_2_value = Null, $row_1_col_3_value = Null, $row_1_col_4_value = Null, $row_1_col_5_value = Null, $row_2_col_1_value = Null, $row_2_col_2_value = Null, $row_2_col_3_value = Null, $row_2_col_4_value = Null, $row_2_col_5_value = Null, $row_3_col_1_value = Null, $row_3_col_2_value = Null, $row_3_col_3_value = Null, $row_3_col_4_value = Null, $row_3_col_5_value = Null, $row_4_col_1_value = Null, $row_4_col_2_value = Null, $row_4_col_3_value = Null, $row_4_col_4_value = Null, $row_4_col_5_value = Null, $row_5_col_1_value = Null, $row_5_col_2_value = Null, $row_5_col_3_value = Null, $row_5_col_4_value = Null, $row_5_col_5_value = Null)

	Local $col_heading_definition = ""

	if $col_1_name <> Null Then $col_heading_definition = $col_1_name
	if $col_2_name <> Null Then $col_heading_definition = $col_heading_definition & "|" & $col_2_name
	if $col_3_name <> Null Then $col_heading_definition = $col_heading_definition & "|" & $col_3_name
	if $col_4_name <> Null Then $col_heading_definition = $col_heading_definition & "|" & $col_4_name
	if $col_5_name <> Null Then $col_heading_definition = $col_heading_definition & "|" & $col_5_name

	local $ctrl = GUICtrlCreateListView($col_heading_definition, $left, $top, $width, $height, BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_SHOWSELALWAYS, $LVS_SORTASCENDING))

	if $col_1_width <> Null Then _GUICtrlListView_SetColumnWidth(-1, 0, $col_1_width)
	if $col_2_width <> Null Then _GUICtrlListView_SetColumnWidth(-1, 1, $col_2_width)
	if $col_3_width <> Null Then _GUICtrlListView_SetColumnWidth(-1, 2, $col_3_width)
	if $col_4_width <> Null Then _GUICtrlListView_SetColumnWidth(-1, 3, $col_4_width)
	if $col_5_width <> Null Then _GUICtrlListView_SetColumnWidth(-1, 4, $col_5_width)

	Local $i

	if $row_1_col_1_value <> Null Then $i = _GUICtrlListView_AddItem(-1, $row_1_col_1_value, 0)
	if $row_1_col_2_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_1_col_2_value, 1, 0)
	if $row_1_col_3_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_1_col_3_value, 2, 0)
	if $row_1_col_4_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_1_col_4_value, 3, 0)
	if $row_1_col_5_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_1_col_5_value, 4, 0)
	if $row_2_col_1_value <> Null Then $i = _GUICtrlListView_AddItem(-1, $row_2_col_1_value, 0)
	if $row_2_col_2_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_2_col_2_value, 1, 0)
	if $row_2_col_3_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_2_col_3_value, 2, 0)
	if $row_2_col_4_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_2_col_4_value, 3, 0)
	if $row_2_col_5_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_2_col_5_value, 4, 0)
	if $row_3_col_1_value <> Null Then $i = _GUICtrlListView_AddItem(-1, $row_3_col_1_value, 0)
	if $row_3_col_2_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_3_col_2_value, 1, 0)
	if $row_3_col_3_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_3_col_3_value, 2, 0)
	if $row_3_col_4_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_3_col_4_value, 3, 0)
	if $row_3_col_5_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_3_col_5_value, 4, 0)
	if $row_4_col_1_value <> Null Then $i = _GUICtrlListView_AddItem(-1, $row_4_col_1_value, 0)
	if $row_4_col_2_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_4_col_2_value, 1, 0)
	if $row_4_col_3_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_4_col_3_value, 2, 0)
	if $row_4_col_4_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_4_col_4_value, 3, 0)
	if $row_4_col_5_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_4_col_5_value, 4, 0)
	if $row_5_col_1_value <> Null Then $i = _GUICtrlListView_AddItem(-1, $row_5_col_1_value, 0)
	if $row_5_col_2_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_5_col_2_value, 1, 0)
	if $row_5_col_3_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_5_col_3_value, 2, 0)
	if $row_5_col_4_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_5_col_4_value, 3, 0)
	if $row_5_col_5_value <> Null Then _GUICtrlListView_AddSubItem(-1, $i, $row_5_col_5_value, 4, 0)

	_GUICtrlListView_SetExtendedListViewStyle(-1, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreatePicEx($left, $top, $width, $height, $tooltip_text = "", $resizing = -1, $hide = False)

	Local $ctrl = GUICtrlCreatePic("", $left, $top, $width, $height)

	if StringLen($tooltip_text) > 0 Then

		_GUIToolTip_AddTool($tooltip, 0, $tooltip_text, GUICtrlGetHandle($ctrl))
	EndIf

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	if $hide = True Then

		GUICtrlSetState(-1, $GUI_HIDE)
	EndIf

	Return $ctrl
EndFunc

Func GUICtrlCreateTabItemEx($text)

	$text = StringReplace($text, "=", "")
	$text = StringReplace($text, ">", "")
	$text = StringStripWS($text, 3)
	Local $ctrl = GUICtrlCreateTabItem($text)
	Return $ctrl

EndFunc

Func GUICtrlCreateGroupEx($text, $left, $top, $width, $height, $tooltip_text = "", $resizing = $GUI_DOCKBORDERS)

	$text = StringReplace($text, "-", "")
	$text = StringReplace($text, ">", "")
	$text = StringStripWS($text, 3)
	Local $ctrl = GUICtrlCreateGroup($text, $left, $top, $width, $height)

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func GUICtrlCreateListEx($left, $top, $width, $height, $tooltip_text = "", $resizing = $GUI_DOCKBORDERS)

	Local $ctrl = GUICtrlCreateList("", $left, $top, $width, $height, BitOR($GUI_SS_DEFAULT_LIST, $WS_HSCROLL))

	if $resizing > -1 Then

		GUICtrlSetResizing(-1, $resizing)
	EndIf

	Return $ctrl

EndFunc

Func depress_button_and_disable_gui($button, $gui = -1, $delay = 0)

	if $gui = -1 Then $gui = $main_gui

	GUICtrlSetStyle($button, -1, $WS_EX_CLIENTEDGE)
    GUISetCursor(15, 0, $gui)
	GUISetState(@SW_DISABLE, $gui)
	Local $focus_dummy = GUICtrlCreateDummy()
	GUICtrlSetState($focus_dummy, $GUI_FOCUS)

	if $delay > 0 Then

		Sleep($delay)
	EndIf

EndFunc

Func raise_button_and_enable_gui($button, $gui = $main_gui)

	GUICtrlSetStyle($button, -1, $WS_EX_DLGMODALFRAME)
    GUISetCursor(2, 0, $gui)
	GUISetState(@SW_ENABLE, $gui)

EndFunc


Func _GUILock($hWnd, $fLock)

    Local $Data, $State

    If $fLock Then
        GUISetCursor(15, 1, $hWnd)
        $State = $GUI_DISABLE
    Else
        GUISetCursor(2, 1, $hWnd)
        $State = $GUI_ENABLE
    EndIf
    _GUICtrlMenu_EnableMenuItem(_GUICtrlMenu_GetSystemMenu($hWnd), $SC_CLOSE, $fLock, 0)
    $Data = _WinAPI_EnumChildWindows($hWnd)
    If IsArray($Data) Then

    ;_WinAPI_SetWindowLong ($hWnd, $GWL_EXSTYLE, $WS_EX_COMPOSITED)
	;GUISetStyle(BitOR($WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU, $WS_EX_COMPOSITED), -1, $hWnd)
;	GUISetStyle($WS_EX_COMPOSITED, -1, $hWnd)

        For $i = 1 To $Data[0][0]
            GUICtrlSetState(_WinAPI_GetDlgCtrlID($Data[$i][0]), $State)
		;_WinAPI_UpdateWindow(_WinAPI_GetDlgCtrlID($Data[$i][0]))
        Next

	;GUISetStyle(BitOR($WS_MINIMIZEBOX, $WS_MAXIMIZEBOX, $WS_SIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU), -1, $hWnd)

    ;_WinAPI_SetWindowLong ($hWnd, $GWL_EXSTYLE, 0)

    EndIf
EndFunc   ;==>_GUILock


Func _ColorFlip($iColor)
    Return BitAND(BitShift($iColor, -16) + BitAND($iColor, 0xFF00) + BitShift($iColor, 16), 0xFFFFFF)
EndFunc

Func _ConvertMonth($date)
   $date = StringReplace($date, 'Jan', '01')
   $date = StringReplace($date, 'Feb', '02')
   $date = StringReplace($date, 'Mar', '03')
   $date = StringReplace($date, 'Apr', '04')
   $date = StringReplace($date, 'May', '05')
   $date = StringReplace($date, 'Jun', '06')
   $date = StringReplace($date, 'Jul', '07')
   $date = StringReplace($date, 'Aug', '08')
   $date = StringReplace($date, 'Sep', '09')
   $date = StringReplace($date, 'Oct', '10')
   $date = StringReplace($date, 'Nov', '11')
   $date = StringReplace($date, 'Dec', '12')
   Return $date
EndFunc

Func GUICtrlListView_GetTopMostIndex($listview)

	Local $top_most_item_y_pos = 9999
	Local $top_most_item_index = 0

	for $i = 0 to (_GUICtrlListView_GetItemCount($listview) - 1)

		Local $y_pos = _GUICtrlListView_GetItemPositionY($timesheet_listview, $i)

		if $y_pos < $top_most_item_y_pos then

			$top_most_item_y_pos = $y_pos
			$top_most_item_index = $i
		EndIf
	Next

	return $top_most_item_index
EndFunc

Func GetLastMondayDate($format = "")

	Local $days_from_today
	Local $last_monday_date
	Local $date_part

	for $days_from_today = 0 to -7 step -1

		$last_monday_date = _DateAdd('d', $days_from_today, _NowCalcDate())
		$date_part = StringSplit($last_monday_date, "/", 3)
		Local $spent_date_day_to_week_index = _DateToDayOfWeek($date_part[0], $date_part[1], $date_part[2])

		if $spent_date_day_to_week_index = 2 Then ExitLoop
	Next

	if StringLen($format) = 0 Then return $last_monday_date

	return _Date_Time_Convert($last_monday_date, "yyyy/MM/dd", $format)

EndFunc

Func _URIEncode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(BinaryToString(StringToBinary($sData,4),1),"")
    Local $nChar
    $sData=""
    For $i = 1 To $aData[0]
        ; ConsoleWrite($aData[$i] & @CRLF)
        $nChar = Asc($aData[$i])
        Switch $nChar
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $sData &= $aData[$i]
            Case 32
                $sData &= "+"
            Case Else
                $sData &= "%" & Hex($nChar,2)
        EndSwitch
    Next
    Return $sData
EndFunc

Func _URIDecode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(StringReplace($sData,"+"," ",0,1),"%")
    $sData = ""
    For $i = 2 To $aData[0]
        $aData[1] &= Chr(Dec(StringLeft($aData[$i],2))) & StringTrimLeft($aData[$i],2)
    Next
    Return BinaryToString(StringToBinary($aData[1],1),4)
EndFunc

Func HourAndMinutesToHours($hours_and_minutes)

	$time_part = StringSplit($hours_and_minutes, ":", 1)

	if $time_part[0] <> 2 Then Return $hours_and_minutes

	Return $time_part[1] + ($time_part[2] / 60)
EndFunc


Func HoursToHourAndMinutes($hours)

	$time_part = StringSplit($hours, ":", 1)

	if $time_part[0] = 2 Then Return $hours

	Return Int($hours) & ":" & StringFormat("%02d", (($hours - Int($hours)) * 60))
EndFunc
