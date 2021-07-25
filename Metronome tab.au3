#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin
#include-once
#Include "HJM link Ex.au3"
#Include "Json.au3"

Global $update_action_items = False
Global $metronome_user_id = ""

Func Metronome_tab_setup()

	GUICtrlCreateTabItemEx("Metronome")
	$metronome_refresh_button = 												GUICtrlCreateImageButton("refresh.ico", 30, 80, 36, "Get your Metronome data", $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
	GUICtrlCreateGroupEx  ("Periods", 20, 120, 200, 160)
	$metronome_periods_listview = 												GUICtrlCreateListViewEx(30, 150, 160, 120, $GUI_DOCKBORDERS, "Period", 90, "GUID", 160)
	GUICtrlCreateGroupEx  ("Quarterly Priorities", 240, 120, 500, 160)
	$metronome_quarterly_priorities_listview = 									GUICtrlCreateListViewEx(260, 150, 460, 120, $GUI_DOCKBORDERS, "Title", 150, "Status", 50, "Order", 50, "GUID", 300)
	GUICtrlCreateGroupEx  ("Action Items", 20, 300, 720, 210)
	$metronome_action_items_listview = 											GUICtrlCreateListViewEx(40, 330, 680, 120, $GUI_DOCKBORDERS, "Text", 250, "Due Date", 100, "Done", 50, "Order", 50, "GUID", 300)
;	$metronome_action_items_add_button = 										GUICtrlCreateImageButton("add.ico", 40, 460, 28, "Add a new Action Item")
;	$metronome_action_items_delete_button = 									GUICtrlCreateImageButton("delete.ico", 70, 460, 28, "Delete the selected Action Item")

EndFunc

Func Metronome_tab_child_gui_setup()
EndFunc


Func Metronome_tab_event_handler($msg)

	Switch $msg


		Case $metronome_refresh_button

			depress_button_and_disable_gui($msg)
			_GUICtrlListView_DeleteAllItems($metronome_periods_listview)
			_GUICtrlListView_DeleteAllItems($metronome_quarterly_priorities_listview)
			_GUICtrlListView_DeleteAllItems($metronome_action_items_listview)

			; Metronome authenticate
			GUICtrlStatusInput_SetText($status_input, "Metronome authentication ...")
			Metronome_auth(GUICtrlRead($metronome_email_input), GUICtrlRead($metronome_password_input))
			GUICtrlStatusInput_SetText($status_input, "")

			; Get user details
			Local $json = Metronome_cURL("https://metronomesoftware.com/api/user")
			$metronome_user_id = Json_Get($json, '.id')

			; Get quarterly priorities
			GUICtrlStatusInput_SetText($status_input, "Getting your quarterly priorities ...")
			Local $json = Metronome_cURL("https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/priority?id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22&day=2021-07-20&current=true", "", "quarterly_priorities")
			Local $i = -1
			Local $index = -1

			_GUICtrlListView_BeginUpdate($metronome_quarterly_priorities_listview)

			while True

				$i = $i + 1
				Local $id = Json_Get($json, '.quarterly_priorities[' & $i & '].id')

				if StringLen($id) < 1 Then ExitLoop

				Local $id_user = Json_Get($json, '.quarterly_priorities[' & $i & '].id_user')

				if StringCompare($id_user, $metronome_user_id) = 0 Then

					Local $title = Json_Get($json, '.quarterly_priorities[' & $i & '].title')
					Local $status = Json_Get($json, '.quarterly_priorities[' & $i & '].status')
					Local $order = Json_Get($json, '.quarterly_priorities[' & $i & '].order')
					Local $id = Json_Get($json, '.quarterly_priorities[' & $i & '].id')

					$index = _GUICtrlListView_AddItem($metronome_quarterly_priorities_listview, $title)
					_GUICtrlListView_AddSubItem($metronome_quarterly_priorities_listview, $index, $status, 1)
					_GUICtrlListView_AddSubItem($metronome_quarterly_priorities_listview, $index, $order, 2)
					_GUICtrlListView_AddSubItem($metronome_quarterly_priorities_listview, $index, $id, 3)


				EndIf


			WEnd

			if $index > -1 Then

				_GUICtrlListView_SetItemSelected($metronome_quarterly_priorities_listview, 0, true, true)
				GUICtrlSetState($metronome_quarterly_priorities_listview, $GUI_FOCUS)
				$update_action_items = True
			EndIf

			_GUICtrlListView_EndUpdate($metronome_quarterly_priorities_listview)
			GUICtrlStatusInput_SetText($status_input, "")

			raise_button_and_enable_gui($msg)


;			curl "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/priority?id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22&day=2021-07-20&current=true" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826" -H "If-Modified-Since: Mon, 26 Jul 1997 05:00:00 GMT" -H "Cache-Control: no-cache" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLXNCUGpPU2lcL01neW9jN1M1a3lcLzI5dkNRZXJETGdhSmNVM0VxSU5pVTdlZjN2V0N6Q3U3MEhIbU4zcktWeFZva3lqRHh6NEFuWDErYlZNQkZFRENZNHlad294bTZ0ME5GSEdRRkJSYTZpN1U4IiwiaXNzIjoibWV0cm9ub21lLWdyb3d0aCIsImV4cCI6MTYyNjgyMDY2NiwiaWF0IjoxNjI2Nzc3NDY2LCJqdGkiOiIzOGEyMzJhZDcxNjNmMTYyOTE2OGIyMTlmZmYyMDg4NjdmMmNkMTExOTA0ZmMyODgyNzJlZTE4ZmM5ZmUxOGRjNWUxMzBjMDgxNTM4NDE3ODI0MjcxZmQ5ZWRjZmFiNzVjMTNiOWQ1ZWE3NTQxYzJlMmEyZDhjYTA1ZTRkYjBiNzg3MGUxNWY0NWYxZGExNjA2NmFmNDkzMGI2YTE5YmIyZTZkN2E0OWZlYzA3MDBlNGIyZmNjMWIxZTFjZTBiYWY3YjQ2Y2E3OTY0ZmM2NDVmOGI0YTM2ZGEyMjEyZDNjODllNmU2ODgzNzIzMzI0OGEyZDliOTA3NzllODI3ZTc2In0.UcX-JP8asPein4NBbvmX8iTSiC6CYKRne5Y1Pve6tOg" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829"




			;Local $iPID = Run('curl -k -b cookie.txt -c cookie.txt "https://metronomesoftware.com/api/user" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/" -H "If-Modified-Since: Mon, 26 Jul 1997 05:00:00 GMT" -H "Cache-Control: no-cache" -H "X-Auth-Token: ' & $token & '" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			Local $iPID = Run('curl -k -b cookie.txt -c cookie.txt "https://metronomesoftware.com/api/user" --compressed -H "Referer: https://metronomesoftware.com/" -H "X-Auth-Token: ' & $token & '" -H "Connection: keep-alive"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			ProcessWaitClose($iPID)
;			Local $json = StdoutRead($iPID)
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			;Exit


			; curl "https://metronomesoftware.com/api/user" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/" -H "If-Modified-Since: Mon, 26 Jul 1997 05:00:00 GMT" -H "Cache-Control: no-cache" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLXNCUGpPU2lcL01neW9jN1M1a3lcLzI5dkNRZXJETGdhSmNVM0VxSU5pVTdlZjN2V0N6Q3U3MEhIbU4zcktWeFZva3lqRHh6NEFuWDErYlZNQkZFRENZNHlad294bTZ0ME5GSEdRRkJSYTZpN1U4IiwiaXNzIjoibWV0cm9ub21lLWdyb3d0aCIsImV4cCI6MTYyNjgyMDY2NiwiaWF0IjoxNjI2Nzc3NDY2LCJqdGkiOiIzOGEyMzJhZDcxNjNmMTYyOTE2OGIyMTlmZmYyMDg4NjdmMmNkMTExOTA0ZmMyODgyNzJlZTE4ZmM5ZmUxOGRjNWUxMzBjMDgxNTM4NDE3ODI0MjcxZmQ5ZWRjZmFiNzVjMTNiOWQ1ZWE3NTQxYzJlMmEyZDhjYTA1ZTRkYjBiNzg3MGUxNWY0NWYxZGExNjA2NmFmNDkzMGI2YTE5YmIyZTZkN2E0OWZlYzA3MDBlNGIyZmNjMWIxZTFjZTBiYWY3YjQ2Y2E3OTY0ZmM2NDVmOGI0YTM2ZGEyMjEyZDNjODllNmU2ODgzNzIzMzI0OGEyZDliOTA3NzllODI3ZTc2In0.UcX-JP8asPein4NBbvmX8iTSiC6CYKRne5Y1Pve6tOg" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829"


		Case $metronome_action_items_add_button

			;Local $iPID = Run('curl -k "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=99482bef-db36-4e31-8a40-13d28449aaad&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "Content-Type: application/json;charset=utf-8" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLUZlZzVlU004VG10NFo2bUpMSjVHc2VUeGFsYUdoXC94UnNlMGlsUWt0dGI5MW5abmFPK3dKbnljdHpMQVFINVlPMk5iaUpCYUdVZ3V0OVZiWU9CMnZSV1dBNFM3dGVwajhJWHVmRHVFUzBwamoiLCJpc3MiOiJtZXRyb25vbWUtZ3Jvd3RoIiwiZXhwIjoxNjI2ODE3OTI5LCJpYXQiOjE2MjY3NzQ3MjksImp0aSI6IjE2YmJiMzlmNzc2NDY2YzNmZmVhNjBjZjQ5NDZhNmNiYWIyNTc3OGU1OGM0OTgwNTQ5NDdmYmJiMzk5MGI3NDdmOGQwYzg4ZmQzOWI1MTQ1OTQyNjdhZjdmZDZiNWQzZmJjNDg5NjZhMDQzYjE4NDIwZjllMjBmNmI0YzE3YTE4MDllOTVmMmFlMGQ5YjU5OGI4ZGNlOGYxNDk1NzRjNDYwZWY1ZDZlNDc2MjY1ODEyN2EwZGUzMDg3YTNlOTgyZjkzNTVmNTQxMTc3M2VjOWQ4OWJmZGFkMzA2Mzc4MzlhZTZmMDg0MGM3OTVjZGQ0NzFmZTlhYjA3ODkzZTczNWYifQ.HNaEMGgRtE42VisTr817psbmdPH2msVxC6yXdPz0CdI" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Origin: https://metronomesoftware.com" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829" --data-raw "{""id_user_owner"":""ecb3d02c-be28-4415-831b-23c5af9d44ac"",""id_priority"":""99482bef-db36-4e31-8a40-13d28449aaad"",""id_period"":""77ce3bd3-2aa5-4a72-9509-36068e0e2a22"",""txt"":""sean action 3"",""send_notification"":true,""frequency"":""once"",""recurring_days"":[],""notes"":"""",""is_strike"":false,""order"":1}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
			;Local $iPID = Run('curl -k "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=99482bef-db36-4e31-8a40-13d28449aaad&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "Content-Type: application/json;charset=utf-8" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLUZlZzVlU004VG10NFo2bUpMSjVHc2VUeGFsYUdoXC94UnNlMGlsUWt0dGI5MW5abmFPK3dKbnljdHpMQVFINVlPMk5iaUpCYUdVZ3V0OVZiWU9CMnZSV1dBNFM3dGVwajhJWHVmRHVFUzBwamoiLCJpc3MiOiJtZXRyb25vbWUtZ3Jvd3RoIiwiZXhwIjoxNjI2ODE3OTI5LCJpYXQiOjE2MjY3NzQ3MjksImp0aSI6IjE2YmJiMzlmNzc2NDY2YzNmZmVhNjBjZjQ5NDZhNmNiYWIyNTc3OGU1OGM0OTgwNTQ5NDdmYmJiMzk5MGI3NDdmOGQwYzg4ZmQzOWI1MTQ1OTQyNjdhZjdmZDZiNWQzZmJjNDg5NjZhMDQzYjE4NDIwZjllMjBmNmI0YzE3YTE4MDllOTVmMmFlMGQ5YjU5OGI4ZGNlOGYxNDk1NzRjNDYwZWY1ZDZlNDc2MjY1ODEyN2EwZGUzMDg3YTNlOTgyZjkzNTVmNTQxMTc3M2VjOWQ4OWJmZGFkMzA2Mzc4MzlhZTZmMDg0MGM3OTVjZGQ0NzFmZTlhYjA3ODkzZTczNWYifQ.HNaEMGgRtE42VisTr817psbmdPH2msVxC6yXdPz0CdI" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Origin: https://metronomesoftware.com" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829" -d "{\"id_user_owner\":\"ecb3d02c-be28-4415-831b-23c5af9d44ac\",\"id_priority\":\"99482bef-db36-4e31-8a40-13d28449aaad\",\"id_period\":\"77ce3bd3-2aa5-4a72-9509-36068e0e2a22\",\"txt\":\"sean action 3\",\"send_notification\":true,\"frequency\":\"once\",\"recurring_days\":[],\"notes\":\"\",\"is_strike\":false,\"order\":1}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			Local $iPID = Run('curl -k "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=99482bef-db36-4e31-8a40-13d28449aaad&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "Content-Type: application/json;charset=utf-8" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLUZlZzVlU004VG10NFo2bUpMSjVHc2VUeGFsYUdoXC94UnNlMGlsUWt0dGI5MW5abmFPK3dKbnljdHpMQVFINVlPMk5iaUpCYUdVZ3V0OVZiWU9CMnZSV1dBNFM3dGVwajhJWHVmRHVFUzBwamoiLCJpc3MiOiJtZXRyb25vbWUtZ3Jvd3RoIiwiZXhwIjoxNjI2ODE3OTI5LCJpYXQiOjE2MjY3NzQ3MjksImp0aSI6IjE2YmJiMzlmNzc2NDY2YzNmZmVhNjBjZjQ5NDZhNmNiYWIyNTc3OGU1OGM0OTgwNTQ5NDdmYmJiMzk5MGI3NDdmOGQwYzg4ZmQzOWI1MTQ1OTQyNjdhZjdmZDZiNWQzZmJjNDg5NjZhMDQzYjE4NDIwZjllMjBmNmI0YzE3YTE4MDllOTVmMmFlMGQ5YjU5OGI4ZGNlOGYxNDk1NzRjNDYwZWY1ZDZlNDc2MjY1ODEyN2EwZGUzMDg3YTNlOTgyZjkzNTVmNTQxMTc3M2VjOWQ4OWJmZGFkMzA2Mzc4MzlhZTZmMDg0MGM3OTVjZGQ0NzFmZTlhYjA3ODkzZTczNWYifQ.HNaEMGgRtE42VisTr817psbmdPH2msVxC6yXdPz0CdI" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Origin: https://metronomesoftware.com" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829" -d "{\"id_user_owner\":\"ecb3d02c-be28-4415-831b-23c5af9d44ac\",\"id_priority\":\"99482bef-db36-4e31-8a40-13d28449aaad\",\"id_period\":\"77ce3bd3-2aa5-4a72-9509-36068e0e2a22\",\"txt\":\"sean action 3\",\"send_notification\":true,\"frequency\":\"once\",\"recurring_days\":[],\"notes\":\"\",\"is_strike\":false,\"order\":1}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
;			ProcessWaitClose($iPID)
;			Local $json = StdoutRead($iPID)
;			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;			Exit


			; curl "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=99482bef-db36-4e31-8a40-13d28449aaad&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "Content-Type: application/json;charset=utf-8" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLUZlZzVlU004VG10NFo2bUpMSjVHc2VUeGFsYUdoXC94UnNlMGlsUWt0dGI5MW5abmFPK3dKbnljdHpMQVFINVlPMk5iaUpCYUdVZ3V0OVZiWU9CMnZSV1dBNFM3dGVwajhJWHVmRHVFUzBwamoiLCJpc3MiOiJtZXRyb25vbWUtZ3Jvd3RoIiwiZXhwIjoxNjI2ODE3OTI5LCJpYXQiOjE2MjY3NzQ3MjksImp0aSI6IjE2YmJiMzlmNzc2NDY2YzNmZmVhNjBjZjQ5NDZhNmNiYWIyNTc3OGU1OGM0OTgwNTQ5NDdmYmJiMzk5MGI3NDdmOGQwYzg4ZmQzOWI1MTQ1OTQyNjdhZjdmZDZiNWQzZmJjNDg5NjZhMDQzYjE4NDIwZjllMjBmNmI0YzE3YTE4MDllOTVmMmFlMGQ5YjU5OGI4ZGNlOGYxNDk1NzRjNDYwZWY1ZDZlNDc2MjY1ODEyN2EwZGUzMDg3YTNlOTgyZjkzNTVmNTQxMTc3M2VjOWQ4OWJmZGFkMzA2Mzc4MzlhZTZmMDg0MGM3OTVjZGQ0NzFmZTlhYjA3ODkzZTczNWYifQ.HNaEMGgRtE42VisTr817psbmdPH2msVxC6yXdPz0CdI" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Origin: https://metronomesoftware.com" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829" --data-raw "{""id_user_owner"":""ecb3d02c-be28-4415-831b-23c5af9d44ac"",""id_priority"":""99482bef-db36-4e31-8a40-13d28449aaad"",""id_period"":""77ce3bd3-2aa5-4a72-9509-36068e0e2a22"",""txt"":""sean action 1"",""send_notification"":true,""frequency"":""once"",""recurring_days"":[],""notes"":"""",""is_strike"":false,""order"":0}"

			; curl "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=99482bef-db36-4e31-8a40-13d28449aaad&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "Content-Type: application/json;charset=utf-8" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLUZlZzVlU004VG10NFo2bUpMSjVHc2VUeGFsYUdoXC94UnNlMGlsUWt0dGI5MW5abmFPK3dKbnljdHpMQVFINVlPMk5iaUpCYUdVZ3V0OVZiWU9CMnZSV1dBNFM3dGVwajhJWHVmRHVFUzBwamoiLCJpc3MiOiJtZXRyb25vbWUtZ3Jvd3RoIiwiZXhwIjoxNjI2ODE3OTI5LCJpYXQiOjE2MjY3NzQ3MjksImp0aSI6IjE2YmJiMzlmNzc2NDY2YzNmZmVhNjBjZjQ5NDZhNmNiYWIyNTc3OGU1OGM0OTgwNTQ5NDdmYmJiMzk5MGI3NDdmOGQwYzg4ZmQzOWI1MTQ1OTQyNjdhZjdmZDZiNWQzZmJjNDg5NjZhMDQzYjE4NDIwZjllMjBmNmI0YzE3YTE4MDllOTVmMmFlMGQ5YjU5OGI4ZGNlOGYxNDk1NzRjNDYwZWY1ZDZlNDc2MjY1ODEyN2EwZGUzMDg3YTNlOTgyZjkzNTVmNTQxMTc3M2VjOWQ4OWJmZGFkMzA2Mzc4MzlhZTZmMDg0MGM3OTVjZGQ0NzFmZTlhYjA3ODkzZTczNWYifQ.HNaEMGgRtE42VisTr817psbmdPH2msVxC6yXdPz0CdI" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Origin: https://metronomesoftware.com" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829" --data-raw "{""id_user_owner"":""ecb3d02c-be28-4415-831b-23c5af9d44ac"",""id_priority"":""99482bef-db36-4e31-8a40-13d28449aaad"",""id_period"":""77ce3bd3-2aa5-4a72-9509-36068e0e2a22"",""txt"":""sean action 2"",""send_notification"":true,""frequency"":""once"",""recurring_days"":[],""notes"":"""",""is_strike"":false,""order"":1}"

	EndSwitch

	if $update_action_items = True Then

		$update_action_items = False

		Local $arr = _GUICtrlListView_GetItemTextArray($metronome_quarterly_priorities_listview)
		Local $quarterly_priority_guid = $arr[4]

		; Get action items
		GUICtrlStatusInput_SetText($status_input, "Getting your action items ...")
		Local $json = Metronome_cURL("https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems?id_priority=" & $quarterly_priority_guid, "", "action_items")

; curl "https://metronomesoftware.com/api/767cfdf2-8182-49ae-b5ec-b1c84700a826/actionitems?id_priority=99482bef-db36-4e31-8a40-13d28449aaad" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0" -H "Accept: application/json, text/plain, */*" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Referer: https://metronomesoftware.com/priorities/767cfdf2-8182-49ae-b5ec-b1c84700a826?id=ebf48f9f-7d53-4ecb-8efd-e09d0ec71c3f&id_user=ecb3d02c-be28-4415-831b-23c5af9d44ac&id_period=77ce3bd3-2aa5-4a72-9509-36068e0e2a22" -H "If-Modified-Since: Mon, 26 Jul 1997 05:00:00 GMT" -H "Cache-Control: no-cache" -H "X-Auth-Token: eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxLXNCUGpPU2lcL01neW9jN1M1a3lcLzI5dkNRZXJETGdhSmNVM0VxSU5pVTdlZjN2V0N6Q3U3MEhIbU4zcktWeFZva3lqRHh6NEFuWDErYlZNQkZFRENZNHlad294bTZ0ME5GSEdRRkJSYTZpN1U4IiwiaXNzIjoibWV0cm9ub21lLWdyb3d0aCIsImV4cCI6MTYyNjgyMDY2NiwiaWF0IjoxNjI2Nzc3NDY2LCJqdGkiOiIzOGEyMzJhZDcxNjNmMTYyOTE2OGIyMTlmZmYyMDg4NjdmMmNkMTExOTA0ZmMyODgyNzJlZTE4ZmM5ZmUxOGRjNWUxMzBjMDgxNTM4NDE3ODI0MjcxZmQ5ZWRjZmFiNzVjMTNiOWQ1ZWE3NTQxYzJlMmEyZDhjYTA1ZTRkYjBiNzg3MGUxNWY0NWYxZGExNjA2NmFmNDkzMGI2YTE5YmIyZTZkN2E0OWZlYzA3MDBlNGIyZmNjMWIxZTFjZTBiYWY3YjQ2Y2E3OTY0ZmM2NDVmOGI0YTM2ZGEyMjEyZDNjODllNmU2ODgzNzIzMzI0OGEyZDliOTA3NzllODI3ZTc2In0.UcX-JP8asPein4NBbvmX8iTSiC6CYKRne5Y1Pve6tOg" -H "Csrf-Token: f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f" -H "Connection: keep-alive" -H "Cookie: __stripe_mid=5f40101c-ae05-4dbd-a204-08e7469b418cc91390; PLAY_CSRF_TOKEN=f3299242abfceacaeb1c6b78df3ac3089032a74a-1624276892963-aba33c4dcfb9f111ecd8d57f; __stripe_sid=fd9f5cc6-95db-45c6-a9b8-cc669b6e6dcfccd829"

		Local $i = -1
		Local $index = -1

		_GUICtrlListView_DeleteAllItems($metronome_action_items_listview)
		_GUICtrlListView_BeginUpdate($metronome_action_items_listview)

		while True

			$i = $i + 1
			Local $id = Json_Get($json, '.action_items[' & $i & '].id')

			if StringLen($id) < 1 Then ExitLoop

			Local $id_user_owner = Json_Get($json, '.action_items[' & $i & '].id_user_owner')

			if StringCompare($id_user_owner, $metronome_user_id) = 0 Then

				Local $txt = Json_Get($json, '.action_items[' & $i & '].txt')
				Local $due_date = Json_Get($json, '.action_items[' & $i & '].due_date')
				Local $is_done = Json_Get($json, '.action_items[' & $i & '].is_done')
				Local $order = Json_Get($json, '.action_items[' & $i & '].order')
				Local $id = Json_Get($json, '.action_items[' & $i & '].id')

				$index = _GUICtrlListView_AddItem($metronome_action_items_listview, $txt)
				_GUICtrlListView_AddSubItem($metronome_action_items_listview, $index, $due_date, 1)
				_GUICtrlListView_AddSubItem($metronome_action_items_listview, $index, $is_done, 2)
				_GUICtrlListView_AddSubItem($metronome_action_items_listview, $index, $order, 3)
				_GUICtrlListView_AddSubItem($metronome_action_items_listview, $index, $id, 4)

			EndIf


		WEnd

		_GUICtrlListView_EndUpdate($metronome_action_items_listview)
		GUICtrlStatusInput_SetText($status_input, "")

		if $index > -1 Then

			_GUICtrlListView_SetItemSelected($metronome_action_items_listview, 0, true, true)
			GUICtrlSetState($metronome_action_items_listview, $GUI_FOCUS)
		EndIf

	EndIf

EndFunc


Func Metronome_tab_WM_NOTIFY_handler($hWndFrom, $iCode)


	Switch $hWndFrom


		Case GUICtrlGetHandle($metronome_quarterly_priorities_listview)

			Switch $iCode

				Case $LVN_ITEMCHANGED

					$update_action_items = True

			EndSwitch


	EndSwitch



EndFunc


Func Metronome_auth($email, $password)

	Local $iPID = Run('curl -k -c cookie.txt "https://metronomesoftware.com/authenticate/credentials" --compressed -H "Referer: https://metronomesoftware.com/" -H "Content-Type: application/json;charset=utf-8" -H "Connection: keep-alive" -d "{\"email\":\"' & $email & '\",\"password\":\"' & $password & '\",\"rememberMe\":false}"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	Local $json = StdoutRead($iPID)
	Local $decoded_json = Json_Decode($json)
	$metronome_token = Json_Get($decoded_json, '.token')

EndFunc

Func Metronome_cURL($url, $data = "", $json_name_prefix = "")

	Local $iPID = Run('curl -k -b metronome_cookie.txt -c metronome_cookie.txt "' & $url & '" --compressed -H "Referer: https://metronomesoftware.com/" -H "X-Auth-Token: ' & $metronome_token & '" -H "Connection: keep-alive"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	Local $json = StdoutRead($iPID)

	if StringLen($json_name_prefix) > 0 Then

		$json = '{"' & $json_name_prefix & '":' & $json & '}'
	EndIf

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	Local $decoded_json = Json_Decode($json)
	Return $decoded_json

EndFunc
