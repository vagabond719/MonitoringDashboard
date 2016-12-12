;Alerts Dashboard
#include <OutlookEx.au3>
#include <Date.au3>
#include <Array.au3>
#include "EzMySql.au3"
$oError = ObjEvent("AutoIt.Error", "ErrHandler")
$to = "email"
$lastreport = _NowCalcDate() & " " & _NowTime(5)
$path = "C:\Temp\Attachment Drop Point\"
$noemail = 0
;$cc = ""

If Not _EzMySql_Startup() Then
	MsgBox(0, "Error Starting MySql", "Error: " & @error & @CR & "Error string: " & _EzMySql_ErrMsg())
	Exit
EndIf

If Not _EzMySql_Open("", "user", "password", "db", "port") Then
	MsgBox(0, "Error opening Database", "Error: " & @error & @CR & "Error string: " & _EzMySql_ErrMsg())
	Exit
EndIf

$outlook = _OL_Open()

While 1
	$emails = ""
	$emails = _OL_ItemFind($outlook, "MailboxName\Inbox", $olMail, "", "", "", "Subject, Body, ReceivedTime, SenderName, To, EntryID, HTMLBody")
	$folder = ""
	$folder = _OL_FolderAccess($outlook, "MailboxName\Inbox")
	;_ArrayDisplay($folder)
	$nowreport = _NowCalcDate() & " " & _NowTime(5)
;~ 	MsgBox(0,"",$lastreport & " > " & $nowreport)
	If $lastreport < $nowreport And _NowTime(5) > "09:00:00" Then
;~ 		Call("Report")
		$lastreport = _NowCalcDate() & " 09:00:00"
		$lastreport = _DateAdd("D", 1, $lastreport)
;~ 		MsgBox(0,"",$lastreport)
	EndIf

	If IsArray($emails) Then
		$noemail = 0
		If IsArray($folder) Then
			For $c = 1 To $emails[0][0]
				$emails[$c][0] = StringUpper($emails[$c][0])
				If StringInStr($emails[$c][1], "NNM View") = 0 Then
					$emails[$c][1] = StringUpper($emails[$c][1])
				EndIf
				$emails2 = $emails[$c][6]
				$emails[$c][6] = StringUpper($emails[$c][6])
				If StringLeft($emails[$c][0], 3) = "RE:" Or StringLeft($emails[$c][0], 3) = "FW:" Or StringLeft($emails[$c][0], 3) = "ESR" Or StringLeft($emails[$c][0], 3) = "RFC" Or StringLeft($emails[$c][0], 3) = "ISR" Then
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Everythingelse\Communications")
					ContinueLoop
				EndIf
				If $emails[$c][3] = "email" Or $emails[$c][3] = "email" Then
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Everythingelse\Communications")
					ContinueLoop
				EndIf
				$emails[$c][1] = StringReplace($emails[$c][1], "%", "\%")
				$emails[$c][1] = StringReplace($emails[$c][1], "'", "\'")
				$emails[$c][1] = StringReplace($emails[$c][1], ":\", ":")
				If StringInStr($emails[$c][0], "SIM Daily Fault Report") <> 0 And $emails[$c][3] <> "MailboxName"Then
					HPSIMReport()
				ElseIf StringInStr($emails[$c][0], "Normal") <> 0 And $emails[$c][3] = "Alerts, M360 (NBCUniversal)" Then
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\M360")
				ElseIf StringInStr($emails[$c][0], "Severity") <> 0 And $emails[$c][3] = "Alerts, M360 (NBCUniversal)" Then
					Call("M360")
				ElseIf StringInStr($emails[$c][0], "Clariion") <> 0 Then
					Call("Clariion")
				ElseIf StringInStr($emails[$c][0], "Splunk") <> 0 And StringInStr($emails[$c][3], "Splunk") <> 0 Then
					Call("Splunk")
				ElseIf StringInStr($emails[$c][0], "DD Alert") <> 0 And $emails[$c][3] = "@NBC UNI Data Protection" Then
					Call("DataDomain")
				ElseIf ($emails[$c][3] = "@NBC UNI Data Protection" Or $emails[$c][3] = "email" Or $emails[$c][3] = "dave.buonoemail") And _
						($emails[$c][4] = "Data Protection; @NBC Uni NBCU Compute Operations L1; Compute L1 (NBCUniversal)" Or $emails[$c][4] = "Data Protection; Compute L1 (NBCUniversal)") Then
					Call("Netbackup")
				ElseIf $emails[$c][3] = "email" Then
					Call("SIMNYC")
				ElseIf $emails[$c][3] = "UK_SIM_Systememail" Then
					Call("UKSIM")
				ElseIf StringInStr($emails[$c][0], "dfm:") <> 0 And $emails[$c][3] = "NetappOpsMgremail" Then
					Call("Netapp")
				ElseIf StringLeft($emails[$c][0], 23) = "[VMware vCenter - Alarm" Then
					Call("VM")
				ElseIf StringInStr($emails[$c][3], "potomac") <> 0 Then
					Call("Potomac")
				ElseIf StringInStr($emails[$c][3], ".hpsim@") <> 0 Then
					If StringInStr($emails[$c][3], "potomac") <> 0 Then
						_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
						ContinueLoop
					EndIf

					Call("HPSIM")
				ElseIf StringInStr($emails[$c][3], "bin") <> 0 Or StringInStr($emails[$c][3], "daemon") <> 0 Then
					Call("Network")
				Else
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Everythingelse")
				EndIf
			Next
		EndIf
	Else
		$noemail += 1
		If $noemail = 10 Then
			$noemail = 0
			_OL_Close($outlook)
			$outlook = ""
			$emails = ""
			ProcessClose("outlook.exe")
			Sleep(2000)
			Run("C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE")
			Sleep(2000)
			$outlook = _OL_Open()
			Sleep(60000)
			_OL_Wrapper_SendMail($outlook, "email", "", "", "Mailbox Restarted: " & _NowTime(), "", "", $olFormatHTML, $olImportanceNormal)
			ContinueLoop
		EndIf
	EndIf
	Call("Dup")
	Sleep(60000)
WEnd

_OL_Close($outlook)
$outlook = ""
$emails = ""

_EzMySql_Close()
_EzMySql_ShutDown()

Func M360()
	$date = ""
	$device = ""
	$severity = ""
	$monitor = ""
	$monitor2 = ""
	$level = ""
	$FunctionType = ""
	$iapps = ""
	$descriptionD = ""
	$array = StringSplit($emails[$c][1], @CRLF)
	$row = UBound($array)
	$counter = 0
	Dim $array2[$row][2]
	For $y = 1 To UBound($array) - 1
		If StringInStr($array[$y], ":") <> 0 Then
			$array2[$counter][1] = StringStripWS(StringUpper(StringMid($array[$y], StringInStr($array[$y], ":") + 1, StringLen($array[$y]) - StringInStr($array[$y], ":"))), 3)
			$array2[$counter][0] = StringStripWS(StringUpper(StringLeft($array[$y], StringInStr($array[$y], ":") - 1)), 3)
			If $array2[$counter][0] = "SEVERITY" And $counter = 0 Then
				ContinueLoop
			ElseIf $array2[$counter][0] = "FOR FURTHER INFORMATION VISIT THE 360 PORTAL AT HTTP" Or $array2[$counter][0] = "SEND INQUIRES" Then
				ContinueLoop
			ElseIf $array2[$counter][0] = "MONITOR" Then
				$monitor = $array2[$counter][1]
				$monitor2 = $array2[$counter][1]
				If StringInStr($monitor, "LOGICAL DISK DRIVE") <> 0 Then
					$monitor = "LOGICAL DISK DRIVE"
				ElseIf StringInStr($monitor, "DISKSPACE") <> 0 Then
					$monitor = "DISKSPACE"
				ElseIf StringInStr($monitor, "FILE SYSTEM") <> 0 Then
					$monitor = "FILE SYSTEM"
				EndIf
			ElseIf $array2[$counter][0] = "START TIME" Then
				$array2[$counter][1] = StringMid($array2[$counter][1], 5, StringLen($array2[$counter][1]) - 4)
				$monstart = StringMid($array2[$counter][1], 1, 3)
				If $monstart = "JAN" Then
					$mon = "01"
				ElseIf $monstart = "FEB" Then
					$mon = "02"
				ElseIf $monstart = "MAR" Then
					$mon = "03"
				ElseIf $monstart = "APR" Then
					$mon = "04"
				ElseIf $monstart = "MAY" Then
					$mon = "05"
				ElseIf $monstart = "JUN" Then
					$mon = "06"
				ElseIf $monstart = "JUL" Then
					$mon = "07"
				ElseIf $monstart = "AUG" Then
					$mon = "08"
				ElseIf $monstart = "SEP" Then
					$mon = "09"
				ElseIf $monstart = "OCT" Then
					$mon = "10"
				ElseIf $monstart = "NOV" Then
					$mon = "11"
				ElseIf $monstart = "DEC" Then
					$mon = "12"
				EndIf
				$array2[$counter][1] = StringMid($array2[$counter][1], 21, 4) & "/" & $mon & "/" & StringMid($array2[$counter][1], 5, 2) & " " & StringMid($array2[$counter][1], 8, 8)
				$date = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "ALERT DESCRIPTION" Then
				If StringInStr($array2[$counter][1], "ORIGINAL ALERT :") <> 0 Then
					$array2[$counter][1] = StringMid($array2[$counter][1], StringInStr($array2[$counter][1], "ORIGINAL ALERT :", 0, -1) + 17, StringLen($array2[$counter][1]) - StringInStr($array2[$counter][1], "ORIGINAL ALERT :", 0, -1) + 17)
				EndIf
				$descriptionD = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "LEVEL" Then
				$level = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "NODE" Then
				$device = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "FUNCTION TYPE" Then
				$FunctionType = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "SEVERITY" Then
				$severity = $array2[$counter][1]
			ElseIf $array2[$counter][0] = "IMPACTED APPLICATIONS" Then
				$iapps = $array2[$counter][1]
			EndIf
			$counter += 1
		EndIf
	Next
	ReDim $array2[$counter][2]
	If $monitor = "CPU" Or $monitor = "MEMORY" Or $monitor = "SWAP_SPACE" Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "System Poll " & $monitor, "The dashboard has polled " & $device & @CRLF & "\\" & $device & "\c$\Temp\MemPoll", "", $olFormatHTML, $olImportanceNormal)
			$arg1 = StringLeft($device,StringInStr($device,".")-1)
			$arguments = $arg1 & " " & $device & " password"
	;~ 		MsgBox(0,"",$arguments)
			ShellExecute("C:\PollTest\Revisions\PollRunnerREVv1.3.bat", $arguments, "C:\PollTest\Revisions", "", @SW_HIDE)
;~ 			MsgBox(0,"","Pause")
;~ 		EndIf
	ElseIf StringInStr($monitor,"LOGICAL DISK DRIVE") <> 0  And StringLen($monitor) < 23 Then
		$arg1 = StringLeft($device,StringInStr($device,".")-1)
		Local $output = StringRegExp($emails2, '(?<=Logical Disk Drive )\D(?=:)', 3)
		If IsArray($output) Then
			$arguments = $arg1 & " " & $device & " " & $output[0] & " password"
			ShellExecute("C:\PollTest\SpaceChecker\SpaceChkRunnerv1.0.bat", $arguments, "C:\PollTest\SpaceChecker", "", @SW_HIDE)
			_OL_Wrapper_SendMail($outlook, "email", "", "", "System Poll " & $monitor, "The dashboard has checked space on " & $device & @CRLF & "\\server\d$\SpaceReports\", "", $olFormatHTML, $olImportanceNormal)
		EndIf
	EndIf
	$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $date & "', '" & $device & "', 'M360', '" & $severity & "', '" & $monitor & "', 'New', '', '');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$dinsert = "Insert into 360AlertDetails(AlertID, Level, FunctionType, Monitor, ImpactedApps, Description_Details) VALUES ((Select MAX(AlertID) from Alerts), '" & $level & "', '" & $FunctionType & "', '" & $monitor2 & "', '" & $iapps & "', '" & $descriptionD & "');"
	If Not _EzMySql_Exec($dinsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	If Not _EzMySql_Exec($activity) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $activity, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\M360")
EndFunc   ;==>M360

Func M360Norm()

	Local $m360norm[UBound($emails)][12]

	$datepos = StringInStr($emails[$c][1], "Start Time            :")
	$datepos2 = StringInStr($emails[$c][1], "End Time           :")
	$datepos3 = $datepos2 - $datepos

	$datepull = StringMid($emails[$c][1], $datepos + 17, $datepos3 - 20)
	If StringLeft($datepull, 2) = "nd" Then
		$year = StringLeft($emails[$c][2], 4)
		$mon = StringMid($emails[$c][2], 5, 2)
		$day = StringMid($emails[$c][2], 7, 2)

		$m360norm[0][0] = "Date"
		$m360norm[$c][0] = $year & "/" & $mon & "/" & $day & " " & @HOUR & ":" & @MIN & ":" & @SEC
	Else

		$year = StringRight($datepull, 4)
		$monstart = StringLeft($datepull, 3)
		$day = StringMid($datepull, 5, 2)
		$time = StringMid($datepull, 8, 8)

		If $monstart = "Jan" Then
			$mon = "01"
		ElseIf $monstart = "Feb" Then
			$mon = "02"
		ElseIf $monstart = "Mar" Then
			$mon = "03"
		ElseIf $monstart = "Apr" Then
			$mon = "04"
		ElseIf $monstart = "May" Then
			$mon = "05"
		ElseIf $monstart = "Jun" Then
			$mon = "06"
		ElseIf $monstart = "Jul" Then
			$mon = "07"
		ElseIf $monstart = "Aug" Then
			$mon = "08"
		ElseIf $monstart = "Sep" Then
			$mon = "09"
		ElseIf $monstart = "Oct" Then
			$mon = "10"
		ElseIf $monstart = "Nov" Then
			$mon = "11"
		ElseIf $monstart = "Dec" Then
			$mon = "12"
		EndIf

		$m360norm[0][0] = "Date"
		$m360norm[$c][0] = $year & "/" & $mon & "/" & $day & " " & $time
	EndIf

	$serverpos = StringInStr($emails[$c][1], "Node :")
	$serverpos2 = StringInStr($emails[$c][1], "Function Type              :")
	$serverpos3 = $serverpos2 - $serverpos

	$m360norm[0][1] = "Server"
	$m360norm[$c][1] = StringUpper(StringMid($emails[$c][1], $serverpos + 7, $serverpos3 - 7))

	$typepos = StringInStr($emails[$c][1], "Monitor                :")
	$typepos2 = StringInStr($emails[$c][1], "Alert Description            :")
	$typetest = StringInStr($emails[$c][1], "Impacted Applications   :")
	If $typetest <> 0 Then
		$typepos2 = $typetest
	EndIf
	$typepos3 = $typepos2 - $typepos

	$m360norm[0][2] = "Monitor"
	$m360norm[$c][2] = StringStripWS(StringMid($emails[$c][1], $typepos + 10, $typepos3 - 10), 2)

	$time = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC

	$update = "Update Alerts SET Status='Remediated', AssignedTo='System', RemediatedOn='" & $time & "', TicketNumber='000000', Notes='Cleared' WHERE Device='" & StringStripWS(StringUpper($m360norm[$c][1]), 8) & "' AND Description='" & StringUpper($m360norm[$c][2]) & "' AND AssignedTo='';"
	$update2 = "Update Alerts SET Cleared='1' WHERE Device='" & StringStripWS(StringUpper($m360norm[$c][1]), 8) & "' AND Description='" & StringUpper($m360norm[$c][2]) & "' AND AssignedTo <> '' AND AssignedTo <> 'System';"
	$update = $update & $update2

	If Not _EzMySql_Exec($update) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $update, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\M360")

	$m360norm = ""

EndFunc   ;==>M360Norm

Func DataDomain()
	If StringInStr($emails[$c][0], "CLEARED: ") <> 0 Then
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Data Protection-Data Domain")
		Return
	EndIf

	Local $ddalert[UBound($emails)][12]

	$datepos = StringInStr($emails[$c][1], "Time:")
	$datepos2 = StringInStr($emails[$c][1], @CRLF, 0, 1, $datepos)
	$datepos3 = $datepos2 - $datepos

	$datepull = StringMid($emails[$c][1], $datepos, $datepos3)
	$datepull = StringStripWS($datepull, 3)
	$datepull = StringRight($datepull, 20)
;~ 	MsgBox(0,"",$datepull)
	$year = StringRight($datepull, 4)
	$monstart = StringLeft($datepull, 3)
	$day = StringMid($datepull, 5, 2)
	$time = StringMid($datepull, 8, 8)

	If $monstart = "Jan" Then
		$mon = "01"
	ElseIf $monstart = "Feb" Then
		$mon = "02"
	ElseIf $monstart = "Mar" Then
		$mon = "03"
	ElseIf $monstart = "Apr" Then
		$mon = "04"
	ElseIf $monstart = "May" Then
		$mon = "05"
	ElseIf $monstart = "Jun" Then
		$mon = "06"
	ElseIf $monstart = "Jul" Then
		$mon = "07"
	ElseIf $monstart = "Aug" Then
		$mon = "08"
	ElseIf $monstart = "Sep" Then
		$mon = "09"
	ElseIf $monstart = "Oct" Then
		$mon = "10"
	ElseIf $monstart = "Nov" Then
		$mon = "11"
	ElseIf $monstart = "Dec" Then
		$mon = "12"
	EndIf

	$ddalert[0][0] = "Date"
	$ddalert[$c][0] = $year & "/" & $mon & "/" & StringStripWS($day, 1) & " " & $time

	$ddalert[0][1] = "Severity"
	$ddalert[$c][1] = "CRITICAL"

	$serverpos = StringInStr($emails[$c][1], "Hostname:")
	$serverpos2 = StringInStr($emails[$c][1], "Location:")
	$serverpos3 = $serverpos2 - $serverpos

	$ddalert[0][2] = "Server"
	$ddalert[$c][2] = StringUpper(StringMid($emails[$c][1], $serverpos + 11, $serverpos3 - 13))

	$despos = StringInStr($emails[$c][1], "Event Message:")
	$despos2 = StringInStr($emails[$c][1], "Object:")
	If $despos2 = 0 Then
		$despos2 = StringInStr($emails[$c][1], "Event Description:")
		If $despos2 = 0 Then
			$despos2 = StringInStr($emails[$c][1], "Recommended Action:")
		EndIf
	EndIf

	$despos3 = $despos2 - $despos

	$ddalert[0][3] = "Event Message"
	$ddalert[$c][3] = StringMid($emails[$c][1], $despos + 16, $despos3 - 18)

	If StringInStr($ddalert[$c][3], "Space Usage") <> 0 Then
		$ddalert[$c][8] = "SPACE USAGE"
	ElseIf StringInStr($ddalert[$c][3], "disk state") <> 0 Then
		$ddalert[$c][8] = "DISK STATE"
	ElseIf StringInStr($ddalert[$c][3], "Sync-as-of") <> 0 Then
		$ddalert[$c][8] = "SYNC TIME"
	ElseIf StringInStr($ddalert[$c][3], "watchdog") <> 0 Then
		$ddalert[$c][8] = "WATCHDOG"
	ElseIf StringInStr($ddalert[$c][3], "Disk has not") <> 0 Then
		$ddalert[$c][8] = "NO DISK"
	ElseIf StringInStr($ddalert[$c][3], "disk has decreased") <> 0 Then
		$ddalert[$c][8] = "MISSING DISK PATH"
	Else
		$ddalert[$c][8] = "DATADOMAIN"
	EndIf

	$locpos = StringInStr($emails[$c][1], "Location:")
	$locpos2 = StringInStr($emails[$c][1], "SerialNo:")
	$locpos3 = $locpos2 - $locpos

	$ddalert[0][4] = "Location"
	$ddalert[$c][4] = StringMid($emails[$c][1], $locpos + 11, $locpos3 - 13)

	$serialpos = StringInStr($emails[$c][1], "SerialNo:")
	$serialpos2 = StringInStr($emails[$c][1], "Time:")
	$serialpos3 = $serialpos2 - $serialpos

	$ddalert[0][5] = "Serial"
	$ddalert[$c][5] = StringMid($emails[$c][1], $serialpos + 11, $serialpos3 - 13)

	$eventpos = StringInStr($emails[$c][1], "Event Id:")
	$eventpos2 = StringInStr($emails[$c][1], "Event Message:")
	$eventpos3 = $eventpos2 - $eventpos

	$ddalert[0][6] = "Event ID"
	$ddalert[$c][6] = StringMid($emails[$c][1], $eventpos + 16, $eventpos3 - 18)

	$actionpos = StringInStr($emails[$c][1], "Recommended Action:")
	$actionpos2 = StringInStr($emails[$c][1], "---")
	$actionpos3 = $actionpos2 - $actionpos

	$ddalert[0][7] = "Recommended Action"
	$ddalert[$c][7] = StringStripWS(StringMid($emails[$c][1], $actionpos + 23, $actionpos3 - 30), 1)

	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Data Protection-Data Domain")

	$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $ddalert[$c][0] & "', '" & $ddalert[$c][2] & "', 'DataDomain', '" & $ddalert[$c][1] & "', '" & $ddalert[$c][8] & "', 'New', '', '');"
	$dinsert = "Insert into DataDomainAlertDetails(AlertID, Event_Message, Location, Serial, EventID, Recommended_Action) VALUES ((Select MAX(AlertID) from Alerts), '" & $ddalert[$c][3] & "', '" & $ddalert[$c][4] & "', '" & $ddalert[$c][5] & "', '" & $ddalert[$c][6] & "', '" & $ddalert[$c][7] & "');"
	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	$query = "SET SQL_SAFE_UPDATES=0;" & $ainsert & $dinsert & $activity
	If Not _EzMySql_Exec($query) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$ddalert = ""

EndFunc   ;==>DataDomain

Func Netbackup()
	If StringInStr($emails[$c][0], "Backup successful") <> 0 Then
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Data Protection-Netbackup")
		Return
	EndIf

	Local $netbalert[UBound($emails)][12]

	$datepos = StringInStr($emails[$c][1], "Date")
	$datepos2 = StringInStr($emails[$c][1], "Policy")
	$datepos3 = $datepos2 - $datepos

	$datepull = StringMid($emails[$c][1], $datepos + 11, $datepos3 - 15)
	$year = StringRight($datepull, 4)
	$monstart = StringLeft($datepull, 3)
	$day = StringMid($datepull, 5, 2)
	$time = StringMid($datepull, 8, 8)

	If $monstart = "Jan" Then
		$mon = "01"
	ElseIf $monstart = "Feb" Then
		$mon = "02"
	ElseIf $monstart = "Mar" Then
		$mon = "03"
	ElseIf $monstart = "Apr" Then
		$mon = "04"
	ElseIf $monstart = "May" Then
		$mon = "05"
	ElseIf $monstart = "Jun" Then
		$mon = "06"
	ElseIf $monstart = "Jul" Then
		$mon = "07"
	ElseIf $monstart = "Aug" Then
		$mon = "08"
	ElseIf $monstart = "Sep" Then
		$mon = "09"
	ElseIf $monstart = "Oct" Then
		$mon = "10"
	ElseIf $monstart = "Nov" Then
		$mon = "11"
	ElseIf $monstart = "Dec" Then
		$mon = "12"
	Else
		$datepull2 = StringMid($emails[$c][1], $datepos, $datepos3)
		$datepull2 = StringReplace($datepull2, "Date", "")
		$datepull2 = StringStripWS($datepull2, 3)
		$mon = StringLeft($datepull2, StringInStr($datepull2, "/") - 1)
		If StringLen($mon) = 1 Then
			$mon = "0" & $mon
		EndIf
		$year = StringInStr($datepull2, "/", 0, 2)
		$year = StringMid($datepull2, $year + 1, 4)
		$day = StringMid($datepull2, StringInStr($datepull2, "/", 0, 1) + 1, StringInStr($datepull2, "/", 0, 2) - StringInStr($datepull2, "/", 0, 1) - 1)
		If StringLen($day) = 1 Then
			$day = "0" & $day
		EndIf
		$time = StringMid($datepull2, StringInStr($datepull2, " ", 0, 1) + 1, StringInStr($datepull2, " ", 0, 2) - StringInStr($datepull2, " ", 0, 1) - 1)
	EndIf

	$netbalert[0][0] = "Date"
	$netbalert[$c][0] = $year & "/" & $mon & "/" & StringStripWS($day, 1) & " " & $time

	$serverpos = StringInStr($emails[$c][1], "Server")
	$serverpos2 = StringInStr($emails[$c][1], "Date")
	$serverpos3 = $serverpos2 - $serverpos

	$netbalert[0][1] = "Server"
	$netbalert[$c][1] = StringUpper(StringStripWS(StringMid($emails[$c][1], $serverpos + 9, $serverpos3 - 10), 2))

	$statuspos = StringInStr($emails[$c][0], "Backup", 0, 2)
	$statuspos2 = StringInStr($emails[$c][0], "on host")
	$statuspos3 = $statuspos2 - $statuspos

	$netbalert[0][2] = "Severity"
	$netbalert[$c][2] = StringUpper(StringMid($emails[$c][0], $statuspos + 7, $statuspos3 - 8))

	If $netbalert[$c][2] = "SUCCESSFUL" Then
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Data Protection-Netbackup")
		Return
	EndIf

	$detpos = StringInStr($emails[$c][1], "   2 - ")
	If $detpos = 0 Then
		$detpos = StringInStr($emails[$c][1], "Catalog Backup Status")
	Else
		$detpos -= 19
	EndIf
	$detpos2 = StringInStr($emails[$c][1], "(status 0)")
	If $detpos2 = 0 Then
		$detpos2 = StringInStr($emails[$c][1], "Catalog Recovery Media")
	EndIf

	$detpos3 = $detpos2 - $detpos

	$netbalert[0][3] = "Alert Details"
	$netbalert[$c][3] = StringMid($emails[$c][1], $detpos + 24, $detpos3 - 25)
	$netbalert[$c][3] = StringReplace($netbalert[$c][3], @CR, " ")
	$netbalert[$c][3] = StringReplace($netbalert[$c][3], @CRLF, " ")
	$netbalert[$c][3] = StringReplace($netbalert[$c][3], @LF, " ")

	$ainsert = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $netbalert[$c][0] & "', '" & $netbalert[$c][1] & "', 'NetBackup', '" & $netbalert[$c][2] & "', '" & "BACKUP" & "', 'New', '', '');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$dinsert = "SET SQL_SAFE_UPDATES=0;Insert into NetbackupAlertDetails(AlertID, Details) VALUES ((Select MAX(AlertID) from Alerts), '" & $netbalert[$c][3] & "');"
	If Not _EzMySql_Exec($dinsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$activity = "SET SQL_SAFE_UPDATES=0;insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	If Not _EzMySql_Exec($activity) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $activity, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Data Protection-Netbackup")

EndFunc   ;==>Netbackup

Func SIMNYC()

	Local $simnycalert[UBound($emails)][12]

	$datepos = StringInStr($emails[$c][1], "Event received:")
	$datepos2 = StringInStr($emails[$c][1], "Event description:")
	$datepos3 = $datepos2 - $datepos

	$datepull = StringMid($emails[$c][1], $datepos + 16, $datepos3 - 21)
	$time = StringRight($datepull, 8)
	$day = StringLeft($datepull, 2)
	$year = StringMid($datepull, 8, 4)
	$monstart = StringMid($datepull, 4, 3)

	If $monstart = "Jan" Then
		$mon = "01"
	ElseIf $monstart = "Feb" Then
		$mon = "02"
	ElseIf $monstart = "Mar" Then
		$mon = "03"
	ElseIf $monstart = "Apr" Then
		$mon = "04"
	ElseIf $monstart = "May" Then
		$mon = "05"
	ElseIf $monstart = "Jun" Then
		$mon = "06"
	ElseIf $monstart = "Jul" Then
		$mon = "07"
	ElseIf $monstart = "Aug" Then
		$mon = "08"
	ElseIf $monstart = "Sep" Then
		$mon = "09"
	ElseIf $monstart = "Oct" Then
		$mon = "10"
	ElseIf $monstart = "Nov" Then
		$mon = "11"
	ElseIf $monstart = "Dec" Then
		$mon = "12"
	EndIf

	$simnycalert[0][0] = "Date"
	$simnycalert[$c][0] = $year & "/" & $mon & "/" & StringStripWS($day, 1) & " " & $time

	$sevpos = StringInStr($emails[$c][1], "Event Severity:")
	$sevpos2 = StringInStr($emails[$c][1], "Event received:")
	$sevpos3 = $sevpos2 - $sevpos

	$simnycalert[0][1] = "Severity"
	$simnycalert[$c][1] = StringUpper(StringMid($emails[$c][1], $sevpos + 16, $sevpos3 - 18))

	$serverpos = StringInStr($emails[$c][1], "Event originator:")
	$serverpos2 = StringInStr($emails[$c][1], "Event Severity:")
	$serverpos3 = $serverpos2 - $serverpos

	$simnycalert[0][2] = "Server"
	$simnycalert[$c][2] = StringUpper(StringMid($emails[$c][1], $serverpos + 18, $serverpos3 - 20))

	$eventpos = StringInStr($emails[$c][1], "Event Name:")
	$eventpos2 = StringInStr($emails[$c][1], "URL:")
	$eventpos3 = $eventpos2 - $eventpos

	$simnycalert[0][3] = "Event Name"
	$simnycalert[$c][3] = StringMid($emails[$c][1], $eventpos + 12, $eventpos3 - 14)

	If StringInStr($simnycalert[$c][3], "Battery Failed") <> 0 Then
		$simnycalert[$c][5] = "BATTERY"
	ElseIf StringInStr($simnycalert[$c][3], "Status Change") <> 0 Then
		$simnycalert[$c][5] = "CHANGE"
	EndIf

	$despos = StringInStr($emails[$c][1], "Event description:")
	$despos2 = StringInStr($emails[$c][1], "Location:")
	$despos3 = $despos2 - $despos

	$simnycalert[0][4] = "Event Description"
	$simnycalert[$c][4] = StringMid($emails[$c][1], $despos + 19, $despos3 - 24)

	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\SIM-NYC")

	$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $simnycalert[$c][0] & "', '" & $simnycalert[$c][2] & "', 'SIM-NYC', '" & $simnycalert[$c][1] & "', '" & $simnycalert[$c][5] & "', 'New', '', '');"
	$dinsert = "Insert into SIM_NYCalertDetails(AlertID, Event_Name, Description) VALUES ((Select MAX(AlertID) from Alerts), '" & $simnycalert[$c][3] & "', '" & $simnycalert[$c][4] & "');"
	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	$query = "SET SQL_SAFE_UPDATES=0;" & $ainsert & $dinsert & $activity
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf

	$simnycalert = ""
EndFunc   ;==>SIMNYC

Func Netapp()

	Local $netapp[UBound($emails)][12]

	$start = StringInStr($emails[$c][6], "<A HREF=")
	$end = StringInStr($emails[$c][6], "<BR>", "0", "1", $start)
	$result = StringMid($emails[$c][6], $start, $end - $start)
	$netapp[$c][0] = $result

	$serverpos = StringInStr($emails[$c][0], "event on")
	$serverpos2 = StringInStr($emails[$c][0], ":", 0, 2)
	$serverpos3 = StringInStr($emails[$c][0], "(")

	If $serverpos2 <> 0 And $serverpos2 < $serverpos3 Then
		$serverpos4 = $serverpos2 - $serverpos
	Else
		$serverpos4 = $serverpos3 - $serverpos
	EndIf

	$netapp[$c][1] = StringUpper(StringMid($emails[$c][0], $serverpos + 9, $serverpos4 - 9))

	$despos = StringInStr($emails[$c][0], "(")
	$despos2 = StringInStr($emails[$c][0], ")")
	$despos3 = $despos2 - $despos

	$netapp[$c][2] = StringMid($emails[$c][0], $despos + 1, $despos3 - 1)

	If StringInStr($netapp[$c][2], "SnapVault") <> 0 Then
		$netapp[$c][3] = "SNAPVAULT"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "Volume") <> 0 Then
		$netapp[$c][3] = "SPACE"
		If StringInStr($netapp[$c][2], "Full") <> 0 Then
			$netapp[$c][5] = "MAJOR"
		Else
			$netapp[$c][5] = "MINOR"
		EndIf
	ElseIf StringInStr($netapp[$c][2], "Global") <> 0 Then
		$netapp[$c][3] = "GLOBAL"
		If StringInStr($netapp[$c][2], "NonCritical") <> 0 Then
			$netapp[$c][5] = "MINOR"
		Else
			$netapp[$c][5] = "MAJOR"
		EndIf
	ElseIf StringInStr($netapp[$c][2], "Backup") <> 0 Then
		$netapp[$c][3] = "BACKUP"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "Aggregate") <> 0 Then
		$netapp[$c][3] = "AGGREGATE"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "Host") <> 0 Then
		$netapp[$c][3] = "HOST"
		If StringInStr($netapp[$c][2], "Down") <> 0 Then
			$netapp[$c][5] = "CRITICAL"
		Else
			$netapp[$c][5] = "MAJOR"
		EndIf
	ElseIf StringInStr($netapp[$c][2], "Qtree") <> 0 Then
		$netapp[$c][3] = "QTREE"
		$netapp[$c][5] = "MAJOR"
	ElseIf StringInStr($netapp[$c][2], "Disks") <> 0 Then
		$netapp[$c][3] = "DISKS"
		If StringInStr($netapp[$c][2], "No Spares") <> 0 Then
			$netapp[$c][5] = "CRITICAL"
		Else
			$netapp[$c][5] = "MINOR"
		EndIf
	ElseIf StringInStr($netapp[$c][2], "Snapshot") <> 0 Then
		$netapp[$c][3] = "SNAPSHOT"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "SnapMirror") <> 0 Then
		$netapp[$c][3] = "SNAPMIRROR"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "Clock") <> 0 Then
		$netapp[$c][3] = "CLOCK"
		$netapp[$c][5] = "WARNING"
	ElseIf StringInStr($netapp[$c][2], "Inodes") <> 0 Then
		$netapp[$c][3] = "INODES"
		$netapp[$c][5] = "MINOR"
	ElseIf StringInStr($netapp[$c][2], "Dead") <> 0 Then
		$netapp[$c][3] = "DEAD"
		$netapp[$c][5] = "CRITICAL"
	Else
		$netapp[$c][3] = "NETAPP"
		$netapp[$c][5] = "MINOR"
	EndIf

	$year = StringLeft($emails[$c][2], 4)
	$mon = StringMid($emails[$c][2], 5, 2)
	$day = StringMid($emails[$c][2], 7, 2)
	$hour = StringMid($emails[$c][2], 9, 2)
	$min = StringMid($emails[$c][2], 11, 2)
	$sec = StringMid($emails[$c][2], 13, 2)

	$netapp[$c][4] = $year & "/" & $mon & "/" & $day & " " & $hour & ":" & $min & ":" & $sec

	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Storage-Netapp")

	$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $netapp[$c][4] & "', '" & $netapp[$c][1] & "', 'Netapp', '" & $netapp[$c][5] & "', '" & $netapp[$c][3] & "', 'New', '', '');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$dinsert = "Insert into NetappAlertDetails(AlertID, Description, URL, Body) VALUES ((Select MAX(AlertID) from Alerts), '" & $netapp[$c][2] & "', '" & $netapp[$c][0] & "', '" & $emails[$c][1] & "');"
	If Not _EzMySql_Exec($dinsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
;~ 	$query = "SET SQL_SAFE_UPDATES=0;" & $ainsert & $dinsert & $activity
	If Not _EzMySql_Exec($activity) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $activity, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$netapp = ""

EndFunc   ;==>Netapp

Func Isilon()
	Local $isilon[UBound($emails)][200]

	If StringInStr($emails[$c][0], "Cluster") <> 0 And StringInStr($emails[$c][0], "has outstanding") <> 0 Then
		$numpos = StringInStr($emails[$c][0], "(")
		$numpos2 = StringInStr($emails[$c][0], ")")
		$numpos3 = $numpos2 - $numpos

		$alertnum = StringMid($emails[$c][0], $numpos + 1, $numpos3 - 1)

		$pos1 = StringInStr($emails[$c][1], "Alert ID")
		$pos2 = StringInStr($emails[$c][1], ".", 1, -1)
		$pos3 = $pos2 - $pos1

		$start = StringMid($emails[$c][1], $pos1, $pos3)

		$array = StringSplit($start, @CRLF)

		$count = 3

		For $z = 1 To $alertnum
			$count += 1
			$repeat = StringInStr($array[$z + $count], "REPEAT", 1, 1)
			$node = StringInStr($array[$z + $count], "Node", 1, 1)
			$cluster = StringInStr($array[$z + $count], "Cluster", 1, 1)
			$new = StringInStr($array[$z + $count], "NEW", 1, 1)

			If $repeat <> 0 Then
				$alertpos = $repeat
			Else
				$alertpos = $new
			EndIf

			If $node <> 0 Then
				$despos = $node
			Else
				$despos = $cluster
			EndIf

			$isilon[$c][$z - 1] = StringStripWS(StringLeft($array[$z + $count], $alertpos - 1), 8)

			$isiarray = _EzMySql_GetTable2d("select * from IsilonAlertDetails WHERE IsilonAlertID='" & $isilon[$c][$z - 1] & "';")

			If UBound($isiarray) > 1 Then
				_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Storage-Isilon")
				ContinueLoop
			EndIf

			$isilon[$c][$z + $alertnum - 1] = StringMid($array[$z + $count], $despos)

			$serverpos = StringInStr($isilon[$c][$z + $alertnum - 1], " ")
			$serverpos2 = StringInStr($isilon[$c][$z + $alertnum - 1], " ", 1, 2)
			$serverpos3 = $serverpos2 - $serverpos

			If StringInStr($isilon[$c][$z + $alertnum - 1], "Temp") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "TEMP"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "Bays on Node") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "DRIVE HEALTH"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "throttling") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "THROTTLE"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "CPU utilization") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "CPU"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "online") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "ONLINE"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "offline") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "NODE OFFLINE"
			ElseIf StringInStr($isilon[$c][$z + $alertnum - 1], "Redundant Power Supplies") <> 0 Then
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "POWER SUPPLY"
			Else
				$isilon[$c][($z + ($alertnum) * 4) - 1] = "ISILON"
			EndIf

			$year = StringLeft($emails[$c][2], 4)
			$mon = StringMid($emails[$c][2], 5, 2)
			$day = StringMid($emails[$c][2], 7, 2)
			$hour = StringMid($emails[$c][2], 9, 2)
			$min = StringMid($emails[$c][2], 11, 2)
			$sec = StringMid($emails[$c][2], 13, 2)

			$isilon[$c][($z + ($alertnum) * 2) - 1] = $year & "/" & $mon & "/" & $day & " " & $hour & ":" & $min & ":" & $sec

			$isilon[$c][($z + ($alertnum) * 3) - 1] = StringUpper(StringMid($isilon[$c][$z + $alertnum - 1], $serverpos + 1, $serverpos3 - 1))

			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Storage-Isilon")

			$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $isilon[$c][($z + ($alertnum) * 2) - 1] & "', '" & $isilon[$c][($z + ($alertnum) * 3) - 1] & "', 'Isilon', '" & "CRITICAL" & "', '" & $isilon[$c][($z + ($alertnum) * 4) - 1] & "', 'New', '', '');"
			$dinsert = "Insert into IsilonAlertDetails(AlertID, IsilonAlertID, Description, HTML) VALUES ((Select MAX(AlertID) from Alerts), '" & $isilon[$c][$z - 1] & "', '" & $isilon[$c][$z + $alertnum - 1] & "', '" & $emails[$c][1] & "');"
			$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
			$query = "SET SQL_SAFE_UPDATES=0;" & $ainsert & $dinsert & $activity
			If Not _EzMySql_Exec($ainsert) Then
				_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
				_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
				Return
			EndIf
		Next

	ElseIf StringInStr($emails[$c][0], "Node") <> 0 Or StringInStr($emails[$c][0], "throttling critical") <> 0 Then
		$nodepos = StringInStr($emails[$c][0], "Node", 1, 1)
		$clusterpos = StringInStr($emails[$c][0], "Cluster", 1, 1)

		If $clusterpos <> 0 Then
			$nodepos = $clusterpos
		EndIf

		$nodepos2 = StringInStr($emails[$c][0], "[", 1, 1, $nodepos)
		$nodepos3 = $nodepos2 - $nodepos

		$isilon[$c][1] = StringMid($emails[$c][0], $nodepos, $nodepos3)

		$serverpos = StringInStr($isilon[$c][1], " ")
		$serverpos2 = StringInStr($isilon[$c][1], " ", 1, 2)
		$serverpos3 = $serverpos2 - $serverpos

		$isilon[$c][3] = StringUpper(StringMid($isilon[$c][1], $serverpos + 1, $serverpos3 - 1))

		If StringInStr($isilon[$c][1], "Temp") <> 0 Then
			$isilon[$c][4] = "TEMP"
		ElseIf StringInStr($isilon[$c][1], "bays on node") <> 0 Then
			$isilon[$c][4] = "DRIVE HEALTH"
		ElseIf StringInStr($isilon[$c][1], "throttling") <> 0 Then
			$isilon[$c][4] = "THROTTLE"
		ElseIf StringInStr($isilon[$c][1], "CPU utilization") <> 0 Then
			$isilon[$c][4] = "CPU"
		ElseIf StringInStr($isilon[$c][1], "online") <> 0 Then
			$isilon[$c][4] = "ONLINE"
		ElseIf StringInStr($isilon[$c][1], "offline") <> 0 Then
			$isilon[$c][4] = "NODE OFFLINE"
		ElseIf StringInStr($isilon[$c][1], "Redundant Power Supplies") <> 0 Then
			$isilon[$c][4] = "POWER SUPPLY"
		Else
			$isilon[$c][4] = "ISILON"
		EndIf

		$alertidpos = StringInStr($emails[$c][1], "Alert ID")
		$alertidpos2 = StringInStr($emails[$c][1], "Alert Occurred")
		$alertidpos3 = $alertidpos2 - $alertidpos

		$isilon[$c][0] = StringStripWS(StringMid($emails[$c][1], $alertidpos + 17, $alertidpos3 - 17), 8)

		$isiarray = _EzMySql_GetTable2d("select * from IsilonAlertDetails WHERE IsilonAlertID='" & $isilon[$c][0] & "';")

		If UBound($isiarray) > 1 Then
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Storage-Isilon")
			Return
		EndIf

		$year = StringLeft($emails[$c][2], 4)
		$mon = StringMid($emails[$c][2], 5, 2)
		$day = StringMid($emails[$c][2], 7, 2)
		$hour = StringMid($emails[$c][2], 9, 2)
		$min = StringMid($emails[$c][2], 11, 2)
		$sec = StringMid($emails[$c][2], 13, 2)

		$isilon[$c][2] = $year & "/" & $mon & "/" & $day & " " & $hour & ":" & $min & ":" & $sec

		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Storage-Isilon")

		$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $isilon[$c][2] & "', '" & $isilon[$c][3] & "', 'Isilon', '" & "CRITICAL" & "', '" & $isilon[$c][4] & "', 'New', '', '');"
		$dinsert = "Insert into IsilonAlertDetails(AlertID, IsilonAlertID, Description, HTML) VALUES ((Select MAX(AlertID) from Alerts), '" & $isilon[$c][0] & "', '" & $isilon[$c][1] & "', '" & $emails[$c][1] & "');"
		$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
		$query = "SET SQL_SAFE_UPDATES=0;" & $ainsert & $dinsert & $activity
		If Not _EzMySql_Exec($ainsert) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf
	EndIf
EndFunc   ;==>Isilon

Func HPSIM()
	$emails[$c][1] = StringReplace($emails[$c][1], "::", ":")
	$emails[$c][1] = StringReplace($emails[$c][1], @CR, "#")
	$emails[$c][1] = StringReplace($emails[$c][1], @LF, "~")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Severity: ", @CRLF & "Event Severity: ")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event received: ", @CRLF & "Event received: ")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event originator: ", @CRLF & "Event originator: ")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Details:#~#~", "Event Details:")
	$emails[$c][1] = StringReplace($emails[$c][1], "#~", @CRLF)
	$emails[$c][1] = StringReplace($emails[$c][1], "REMOTE INSIGHT/ ", "")
	$emails[$c][1] = StringReplace($emails[$c][1], "REMOTE INSIGHT ", "")
	$hpsim = StringSplit($emails[$c][1], @CRLF)
	;_ArrayDisplay($hpsim)
	$rows = UBound($hpsim)
	$columns = UBound($hpsim, 2) + 1
	Dim $hpsim2[$rows][2]
	$count = 0
	For $y = 1 To UBound($hpsim) - 1
		If $hpsim[$y] <> "" Then
			If $hpsim[$y] <> " " Then
				$hpsim2[$count][0] = StringLeft($hpsim[$y], StringInStr($hpsim[$y], ":") - 1)
				$hpsim2[$count][1] = StringRight($hpsim[$y], StringLen($hpsim[$y]) - StringInStr($hpsim[$y], ":") - 1)
				$hpsim2[$count][0] = StringReplace($hpsim2[$count][0], " ", "_")
				If $hpsim2[$count][0] = "Event_received" Then
					$mon = StringMid($hpsim2[$count][1], 4, 3)
					$year = StringMid($hpsim2[$count][1], 8, 4)
					$day = StringMid($hpsim2[$count][1], 1, 2)
					$time = StringMid($hpsim2[$count][1], 14, 8)

					If $mon = "Jan" Then
						$mon = "01"
					ElseIf $mon = "Feb" Then
						$mon = "02"
					ElseIf $mon = "Mar" Then
						$mon = "03"
					ElseIf $mon = "Apr" Then
						$mon = "04"
					ElseIf $mon = "May" Then
						$mon = "05"
					ElseIf $mon = "Jun" Then
						$mon = "06"
					ElseIf $mon = "Jul" Then
						$mon = "07"
					ElseIf $mon = "Aug" Then
						$mon = "08"
					ElseIf $mon = "Sep" Then
						$mon = "09"
					ElseIf $mon = "Oct" Then
						$mon = "10"
					ElseIf $mon = "Nov" Then
						$mon = "11"
					ElseIf $mon = "Dec" Then
						$mon = "12"
					EndIf
					$hpsim2[$count][1] = $year & "/" & $mon & "/" & $day & " " & $time
				EndIf
				$count += 1
			EndIf
		EndIf
	Next
	ReDim $hpsim2[$count][2]
	If StringLeft($emails[$c][3], 3) = "ecl" Then
		$atype = "ecsimalertdetails"
		$atype2 = "HP SIM EC"
	ElseIf StringLeft($emails[$c][3], 3) = "ec." Then
		$atype = "ecsimalertdetails"
		$atype2 = "HP SIM EC"
	ElseIf StringLeft($emails[$c][3], 3) = "ush" Then
		$atype = "ushsimalertdetails"
		$atype2 = "HP SIM USH"
	ElseIf StringLeft($emails[$c][3], 3) = "aoa" Then
		$atype = "aoasimalertdetails"
		$atype2 = "HP SIM AOA"
	ElseIf StringLeft($emails[$c][3], 3) = "nyc" Then
		$atype = "nycsimalertdetails"
		$atype2 = "HP SIM NYC"
	Else
		$atype = "fail"
		$atype2 = "fail"
	EndIf

	;_ArrayDisplay($hpsim2)

	For $y = 0 To UBound($hpsim2) - 1

		If $y = 0 Then
			$query = "insert into " & $atype & "(alertid"
			$query2 = ") values ((select max(alertid) from alerts)"
		ElseIf $hpsim2[$y][0] = "User_Name" Or $hpsim2[$y][0] = "" Or StringInStr($hpsim2[$y][0], "Association_has_been_removed") <> 0 Or StringInStr($hpsim2[$y][0], "CURRENT_TIME") <> 0 Then
			ContinueLoop
		Else
			$query = $query & ", " & $hpsim2[$y][0]
			$query2 = $query2 & ",'" & $hpsim2[$y][1] & "'"
		EndIf
		If $hpsim2[$y][0] = "Event_Name" Then
			If StringInStr($hpsim2[$y][1], "(SNMP)") <> 0 Then
				$desc = StringMid($hpsim2[$y][1], StringInStr($hpsim2[$y][1], ")") + 2, StringInStr($hpsim2[$y][1], "(", 0, 2) - (StringInStr($hpsim2[$y][1], ")") + 3))
			Else
				$desc = $hpsim2[$y][1]
			EndIf
		ElseIf $hpsim2[$y][0] = "Event_originator" Then
			$device = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Event_Severity" Then
			$severity = StringUpper($hpsim2[$y][1])
;~ 			EndIf
		ElseIf $hpsim2[$y][0] = "Event_received" Then
			$dater = $hpsim2[$y][1]
		EndIf
	Next
	If Not IsDeclared("dater") Then
;~ 		MsgBox(0,"",$dater)
		_OL_Wrapper_SendMail($outlook, "email", "", "", "Incomplete Email", $emails[$c][1], "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$query3 = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $dater & "', '" & $device & "', '" & $atype2 & "', '" & $severity & "', '" & $desc & "', 'New', '', '');"
;~ 	MsgBox(0,"",$query)
	If Not _EzMySql_Exec($query3) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query3, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$ainsert = $query & $query2 & ");"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	;$ainsert = $query3 & $query & $query2 & ");insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	$ainsert = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\HPSIM")
EndFunc   ;==>HPSIM

Func VM()
	$target = ""
	$pstatus = ""
	$nstatus = ""
	$alarmdef = ""
	$cvalues = ""
	$desc = ""
	$emails[$c][1] = StringReplace($emails[$c][1], ": " & @CR & @LF, ":")
	$emails[$c][1] = StringReplace($emails[$c][1], @CR & @LF, @LF)
	$emails[$c][1] = StringReplace($emails[$c][1], @LF & " " & @LF, @LF)
	$emails[$c][1] = StringLeft($emails[$c][1], StringLen($emails[$c][1]) - 1)
	If StringInStr($emails[$c][1], @LF & "Previous Status: ") = 0 Then
		$emails[$c][1] = StringReplace($emails[$c][1], "Previous Status: ", @LF & "Previous Status: ")
	EndIf
	If StringInStr($emails[$c][1], @LF & "New Status: ") = 0 Then
		$emails[$c][1] = StringReplace($emails[$c][1], "New Status: ", @LF & "New Status: ")
	EndIf
	$hpsim = StringSplit($emails[$c][1], @LF)
	$rows = UBound($hpsim)
	Dim $hpsim2[$rows][2]
	For $x = 1 To $rows - 1
		$hpsim2[$x][1] = StringMid($hpsim[$x], (StringInStr($hpsim[$x], ":") + 1), StringLen($hpsim[$x]))
		$hpsim2[$x][0] = StringLeft($hpsim[$x], StringInStr($hpsim[$x], ":") - 1)
		$hpsim2[$x][1] = StringStripWS($hpsim2[$x][1], 3)
		$hpsim2[$x][0] = StringStripWS($hpsim2[$x][0], 3)
		$hpsim2[$x][0] = StringReplace($hpsim2[$x][0], " ", "_")

		If $hpsim2[$x][0] = "Target" Then
			$target = $hpsim2[$x][1]
		ElseIf $hpsim2[$x][0] = "Previous_Status" Then
			$pstatus = $hpsim2[$x][1]
		ElseIf $hpsim2[$x][0] = "New_Status" Then
			$nstatus = $hpsim2[$x][1]
		ElseIf $hpsim2[$x][0] = "Alarm_Definition" Then
			$alarmdef = $hpsim2[$x][1]
		ElseIf $hpsim2[$x][0] = "Current_values_for_metric/state" Then
			$hpsim2[$x][0] = "Current_values_for_metric_state"
			$cvalues = $hpsim2[$x][1]
		ElseIf $hpsim2[$x][0] = "Description" Then
			$desc = $hpsim2[$x][1]
		EndIf

	Next
	$time = StringMid($emails[$c][2], 1, 4) & "/" & StringMid($emails[$c][2], 5, 2) & "/" & StringMid($emails[$c][2], 7, 2) & " " & StringMid($emails[$c][2], 9, 2) & ":" & _
			StringMid($emails[$c][2], 11, 2) & ":" & StringMid($emails[$c][2], 13, 2)
	$ainsert = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES (" & _
			"'" & $time & "', '" & $target & "', 'VM','CRITICAL','ALARM', 'New', '', '');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$ainsert = "insert into vmalertdetails(alertID,Target,Previous_Status,New_Status,Alarm_Definition,Current_values_for_metric_state,Description,Reporting_Server) values " & _
			"((select max(alertid) from alerts),'" & $target & "','" & $pstatus & "','" & $nstatus & "','" & $alarmdef & "','" & $cvalues & "','" & $desc & "','" & _
			$emails[$c][3] & "');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$ainsert = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\VM")
EndFunc   ;==>VM

Func UKSIM()
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Severity", @CRLF & "Event_Severity:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Cleared Status", @CRLF & "Cleared_Status:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Source", @CRLF & "Event_Source:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Associated System Status", @CRLF & "Associated_System_Status:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Associated System", @CRLF & "Associated_System:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Time", @CRLF & "Event_Time:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Description", @CRLF & "Description:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Category", @CRLF & "Event_Category:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Assignee", @CRLF & "Assignee:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Comments", @CRLF & "Comments:")
	$emails[$c][1] = StringReplace($emails[$c][1], "Event Identification and Details ", "")
	$emails[$c][1] = StringReplace($emails[$c][1], @TAB, "")
	$emails[$c][1] = StringReplace($emails[$c][1], @CRLF & @CRLF, @CRLF)
	$hpsim = StringSplit($emails[$c][1], @CRLF)
	$rows = UBound($hpsim)
	$columns = UBound($hpsim, 2) + 1
	Dim $hpsim2[$rows][2]
	$count = 0
	For $y = 1 To UBound($hpsim) - 1
		If $hpsim[$y] <> "" Then
			If $hpsim[$y] <> " " Then
				$hpsim2[$count][0] = StringLeft($hpsim[$y], StringInStr($hpsim[$y], ":") - 1)
				$hpsim2[$count][1] = StringRight($hpsim[$y], StringLen($hpsim[$y]) - StringInStr($hpsim[$y], ":") - 1)
				$hpsim2[$count][1] = StringStripWS($hpsim2[$count][1], 3)
				If $hpsim2[$count][0] = "Event_Severity" Or $hpsim2[$count][0] = "Associated_System_Status" Then
					$hpsim2[$count][1] = StringMid($hpsim2[$count][1], 2, StringLen($hpsim2[$count][1]) - 1)
				ElseIf $hpsim2[$count][0] = "Event_Time" Then
					$hpsim2[$count][1] = StringMid($hpsim2[$count][1], 6, StringLen($hpsim2[$count][1]) - 9)
					$month = StringMid($hpsim2[$count][1], 1, StringInStr($hpsim2[$count][1], "/", 0, 1) - 1)
					If StringLen($month) = 1 Then
						$month = "0" & $month
					EndIf
					$day = StringMid($hpsim2[$count][1], StringInStr($hpsim2[$count][1], "/", 0, 1) + 1, StringInStr($hpsim2[$count][1], "/", 0, 2) - StringInStr($hpsim2[$count][1], "/", 0, 1) - 1)
					$year = StringMid($hpsim2[$count][1], StringInStr($hpsim2[$count][1], "/", 0, 2) + 1, 4)
					$hour = StringMid($hpsim2[$count][1], StringInStr($hpsim2[$count][1], " ") + 1, StringInStr($hpsim2[$count][1], ":") - StringInStr($hpsim2[$count][1], " ") - 1)
					If StringInStr($hpsim2[$count][1], " PM") <> 0 Then
						If $hour <> 12 Then
							$hour = $hour + 12
						EndIf
					EndIf
					If StringLen($hour) = 1 Then
						$hour = "0" & $hour
					EndIf
					$min = StringMid($hpsim2[$count][1], StringInStr($hpsim2[$count][1], ":") + 1, 2)
					$sec = "00"
					$hpsim2[$count][1] = $year & "/" & $month & "/" & $day & " " & $hour & ":" & $min & ":" & $sec
				EndIf
				$hpsim2[$count][0] = StringReplace($hpsim2[$count][0], " ", "_")
				$count += 1
			EndIf
		EndIf
	Next
	ReDim $hpsim2[$count][2]

	For $y = 0 To UBound($hpsim2) - 1
		If $hpsim2[$y][0] = "Event_Severity" Then
			$esev = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Cleared_Status" Then
			$cstatus = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Event_Source" Then
			$esource = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Associated_System" Then
			$target = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Associated_System_Status" Then
			$asstatus = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Event_Time" Then
			$etime = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Description" Then
			$desc = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Event_Category" Then
			$ecat = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Assignee" Then
			$assignee = $hpsim2[$y][1]
		ElseIf $hpsim2[$y][0] = "Comments" Then
			$comments = $hpsim2[$y][1]
		EndIf
	Next
	$ainsert = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES " & _
			"(sysdate(), '" & $target & "', 'HP SIM UK','Critical','Alarm', 'New', '', '');"

	$dinsert = "insert into uksimalertdetails(AlertID,Event_Severity,Cleared_Status,Event_Source,Associated_System,Associated_System_Status,Event_Time,Description,Event_Category,Assignee," & _
			"Comments) values ((select max(alertid) from alerts),'" & $esev & "','" & $cstatus & "','" & $esource & "','" & $target & "','" & $asstatus & "','" & $etime & "','" & $desc & _
			"','" & $ecat & "','" & $assignee & "','" & $comments & "');"

	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	If Not _EzMySql_Exec($dinsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	If Not _EzMySql_Exec($activity) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $activity, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf

	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\HPSIM")
EndFunc   ;==>UKSIM

Func ErrHandler()
	$HexNumber = Hex($oError.number, 8)
	$body = $oError.Description & @CRLF & _
			"Source: " & @TAB & $oError.source & @CRLF & _
			"at Line #: " & $oError.ScriptLine & @TAB & _
			"Last DllError: " & @TAB & $oError.lastdllerror & @CRLF & _
			"Help File: " & @TAB & $oError.helpfile & @TAB & "Context: " & @TAB & $oError.helpcontext & @CRLF & _
			"Subject: " & @TAB & $emails[$c][0]

	;MsgBox(0,"",$body)
	_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $body, "", $olFormatHTML, $olImportanceNormal)
	SetError(1) ; to check for after this function returns
EndFunc   ;==>ErrHandler

Func Dup()
	$array = _EzMySql_GetTable2d("SELECT alertid, date, device, alerttype, severity, description FROM dsbpws01.alerts where status <> 'Remediated'")

;~ 	_ArrayDisplay($array)

	$output = ""

	For $x = 1 To UBound($array) - 1
		For $y = 1 To UBound($array) - 1
			If $array[$x][0] = $array[$y][0] Then
				ContinueLoop
			EndIf
			If $array[$x][1] = $array[$y][1] And $array[$x][2] = $array[$y][2] And $array[$x][3] = $array[$y][3] And $array[$x][5] = $array[$y][5] Then
				If $output = "" Then
					If $array[$x][0] < $array[$y][0] Then
						$output = $output & "'" & $array[$y][0] & "'"
					ElseIf $array[$x][0] > $array[$y][0] Then
						$output = $output & "'" & $array[$x][0] & "'"
					EndIf
				Else
					If $array[$x][0] < $array[$y][0] Then
						If StringInStr($output, $array[$y][0]) = 0 Then
							$output = $output & ",'" & $array[$y][0] & "'"
						EndIf
					ElseIf $array[$x][0] > $array[$y][0] Then
						If StringInStr($output, $array[$x][0]) = 0 Then
							$output = $output & ",'" & $array[$x][0] & "'"
						EndIf
					EndIf
				EndIf
			EndIf
		Next
	Next

	;MsgBox(0, "", $output)
	If $output <> "" Then
		$query = "update alerts set assignedto = 'System', Status='Remediated' where alertid in (" & $output & ")"
		If Not _EzMySql_Exec($query) Then
			;MsgBox(0, "", "FAIL!!")
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query, "", $olFormatHTML, $olImportanceNormal)
			Return
		EndIf
	EndIf
EndFunc   ;==>Dup

Func Potomac()

	$url = ""
	$date = ""
	$device = ""

	If StringInStr($emails[$c][1], "Sign-In ") <> 0 Then
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Potomac")
		Return
;~ 		ContinueLoop
	EndIf
	$potomac = StringSplit($emails[$c][1], @CRLF)
	;_ArrayDisplay($potomac)
	For $x = 1 To UBound($potomac) - 1
		If StringInStr($potomac[$x], "Event received:") <> 0 Then
			$mon = StringMid($potomac[$x], 20, 3)
			$year = StringMid($potomac[$x], 24, 4)
			$day = StringMid($potomac[$x], 17, 2)
			$time = StringMid($potomac[$x], 30, 8)

			If $mon = "Jan" Then
				$mon = "01"
			ElseIf $mon = "Feb" Then
				$mon = "02"
			ElseIf $mon = "Mar" Then
				$mon = "03"
			ElseIf $mon = "Apr" Then
				$mon = "04"
			ElseIf $mon = "May" Then
				$mon = "05"
			ElseIf $mon = "Jun" Then
				$mon = "06"
			ElseIf $mon = "Jul" Then
				$mon = "07"
			ElseIf $mon = "Aug" Then
				$mon = "08"
			ElseIf $mon = "Sep" Then
				$mon = "09"
			ElseIf $mon = "Oct" Then
				$mon = "10"
			ElseIf $mon = "Nov" Then
				$mon = "11"
			ElseIf $mon = "Dec" Then
				$mon = "12"
			EndIf
			$date = $year & "/" & $mon & "/" & $day & " " & $time
			;MsgBox(0,"",$date)
		EndIf
		If StringInStr($potomac[$x], "Event originator:") <> 0 Then
			$device = StringReplace($potomac[$x], "Event originator: ", "")
			$url = "https://" & $device & "/login.html"
		EndIf
	Next
	$ainsert = "Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & $date & "', '" & $device & "', 'POTOMAC', 'CRITICAL', 'FAULT NOTIFICATION', 'New', '', '');"
;~ 	MsgBox(0, "", $ainsert)
	If Not _EzMySql_Exec($ainsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$dinsert = "Insert into potomacalertdetails(AlertID, eventseverity, eventname, eventdescription, url, eventoriginator, eventreceived) VALUES ((Select MAX(AlertID) from Alerts), 'CRITICAL', 'FAULT NOTIFICATION', " & _
			"'This notification is generated by a UCS node whenever a fault is active and the fault state changes.','" & $url & "','" & $device & "','" & $date & "');"
;~ 	MsgBox(0, "", $dinsert)
	If Not _EzMySql_Exec($dinsert) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	$activity = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
;~ 	MsgBox(0, "", $activity)
	If Not _EzMySql_Exec($activity) Then
		_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $activity, "", $olFormatHTML, $olImportanceNormal)
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
		Return
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Potomac")
EndFunc   ;==>Potomac

Func Network()
	If StringInStr($emails[$c][1], "alertclosed") <> 0 Then
		_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Network")
		Return
	EndIf
	$severity = "Normal"
	If StringInStr($emails[$c][1], "On-Air") <> 0 Or StringInStr($emails[$c][1], "ONAIR") <> 0 Then
		$severity = "Critical"
	EndIf
	$emails[$c][6] = StringReplace($emails[$c][6], "'", "`")
	$emails[$c][6] = StringReplace($emails[$c][6], '"', "``")
	$emails[$c][6] = "<html>" & $emails[$c][6] & "</html>"
	$emails[$c][6] = StringReplace($emails[$c][6], "&NBSP;", " ")
	$emails[$c][1] = StringReplace($emails[$c][1], "(Guest)", "")
	$emails[$c][1] = StringReplace($emails[$c][1], "NNM View: ", @LF & "NNM View: ")
	$network = StringSplit($emails[$c][1], @LF)
	$flag = 0
	$counter = 0
	Dim $network2[1][2]
	$date = ""
	$dater = ""
	$device = ""
	$desc = ""
	$values = ""
	$data = ""
	For $y = 1 To UBound($network) - 1


		If $c = "2" And $y = "21" Then
			$network[$y] = StringReplace($network[$y], @CR, "")
		EndIf
		If StringInStr($network[$y], "Message") <> 0 Then
			$flag = 1
			$network[$y - 1] = ""
			$network[$y] = ""
		EndIf
		If $flag = 1 Then
			If StringInStr($network[$y], "------------") <> 0 Then
				$flag = 0
			Else
				$network[$y] = ""
			EndIf
		EndIf
		If StringInStr($network[$y], ":") <> 0 Then
			$network2[$counter][0] = StringMid($network[$y], 1, StringInStr($network[$y], ":") - 1)
			$network2[$counter][1] = StringMid($network[$y], StringInStr($network[$y], ":") + 1, StringLen($network[$y]))
			$network2[$counter][0] = StringReplace($network2[$counter][0], "-", "")
			$network2[$counter][0] = StringStripWS($network2[$counter][0], 3)
			$network2[$counter][0] = StringReplace($network2[$counter][0], " ", "_")
			$network2[$counter][1] = StringStripWS($network2[$counter][1], 3)
			If StringRight($network2[$counter][1], 1) = "-" Then
				$network2[$counter][1] = StringLeft($network2[$counter][1], StringLen($network2[$counter][1]) - 1)
				$network2[$counter][1] = StringStripWS($network2[$counter][1], 3)
			EndIf
			If $network2[$counter][0] = "Occurrence_Time" Then
				$network2[$counter][1] = StringReplace($network2[$counter][1], StringLeft($network2[$counter][1], 4), "")
				$mon = StringLeft($network2[$counter][1], 3)

				If $mon = "Jan" Then
					$mon = "01"
				ElseIf $mon = "Feb" Then
					$mon = "02"
				ElseIf $mon = "Mar" Then
					$mon = "03"
				ElseIf $mon = "Apr" Then
					$mon = "04"
				ElseIf $mon = "May" Then
					$mon = "05"
				ElseIf $mon = "Jun" Then
					$mon = "06"
				ElseIf $mon = "Jul" Then
					$mon = "07"
				ElseIf $mon = "Aug" Then
					$mon = "08"
				ElseIf $mon = "Sep" Then
					$mon = "09"
				ElseIf $mon = "Oct" Then
					$mon = "10"
				ElseIf $mon = "Nov" Then
					$mon = "11"
				ElseIf $mon = "Dec" Then
					$mon = "12"
				EndIf
				$year = StringRight($network2[$counter][1], 4)
				$day = StringMid($network2[$counter][1], 5, 2)
				$time = StringMid($network2[$counter][1], 8, 8)
				$network2[$counter][1] = $year & "/" & $mon & "/" & $day & " " & $time
				$network2[$counter][1] = _DateAdd('h', 3, $network2[$counter][1])
				$dater = $network2[$counter][1]
			ElseIf $network2[$counter][0] = "Alert_Name" Then
				$network2[$counter][1] = StringUpper($network2[$counter][1])
				$desc = $network2[$counter][1]
			ElseIf $network2[$counter][0] = "Source_Node" Then
				$network2[$counter][1] = StringUpper($network2[$counter][1])
				$device = $network2[$counter][1]
			ElseIf $network2[$counter][0] = "Script" Then
				$network2[$counter][0] = ""
				$network2[$counter][1] = ""
			ElseIf $network2[$counter][0] = "NNM" Then
				$network2[$counter][0] = ""
				$network2[$counter][1] = ""
			ElseIf $network2[$counter][0] = "NNM_Admin" Then
				$network2[$counter][0] = ""
				$network2[$counter][1] = ""
			ElseIf $network2[$counter][0] = "NNM_User_Guide" Then
				$network2[$counter][0] = ""
				$network2[$counter][1] = ""
			EndIf
			If $network2[$counter][1] = "null" Then
				$network2[$counter][1] = "NA"
			EndIf
			If $values = "" And $network2[$counter][1] <> "" Then
				$values = $network2[$counter][0]
				$data = $network2[$counter][1]
			ElseIf $network2[$counter][1] <> "" Then
				$values = $values & "," & $network2[$counter][0]
				$data = $data & "','" & $network2[$counter][1]
			EndIf
			$counter += 1
			ReDim $network2[$counter + 1][2]
		EndIf
	Next
	If UBound($network2) > 10 Then
		$values = "(html_source,subject,alertid," & $values & ")"
		$data = "('" & $emails[$c][6] & "','" & $emails[$c][0] & "',(select max(alertid) from alerts),'" & $data & "')"
		$query = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & _
				$dater & "', '" & $device & "', 'Network', '" & $severity & "', '" & $desc & "', 'New', '', '');"
		If Not _EzMySql_Exec($query) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf
		$dinsert = "insert into networkalertdetails " & $values & " values " & $data & ";"
		If Not _EzMySql_Exec($dinsert) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf
		$ainsert = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
		If Not _EzMySql_Exec($ainsert) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf
	EndIf
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Network")
EndFunc   ;==>Network

Func Splunk()
	$message = $emails[$c][6]
	$message = StringReplace($message, "'", "`")
	$emails[$c][6] = StringReplace($emails[$c][6], "<tr><th>_raw</th><th>_time</th><th>host</th><th>index</th><th>linecount</th><th>source</th><th>sourcetype</th><th>splunk_server</th>" & _
			"<th>tag::sourcetype</th></tr>", "")
	$emails[$c][6] = StringReplace($emails[$c][6], "</pre></td><td><pre>", "#")
	$emails[$c][6] = StringReplace($emails[$c][6], "<br>", "|")
	$emails[$c][6] = StringReplace($emails[$c][6], "<tr", "|")
	$emails[$c][6] = StringReplace($emails[$c][6], "<html>" & @CRLF & "<head>" & @CRLF & "</head>" & @CRLF & "<body>" & @CRLF & "saved search results.", "")
	$emails[$c][6] = StringReplace($emails[$c][6], '<table border="1">', "")
	$emails[$c][6] = StringReplace($emails[$c][6], 'Alert was triggered because of', "Alert_Trigger")

	$splunk = StringSplit($emails[$c][6], "|")
	Dim $splunk2[1][11]
	Dim $splunk3[1][14]
	$counter = 0
	For $x = 0 To UBound($splunk) - 1
		$splunk[$x] = StringReplace($splunk[$x], @CR, " ")
		$splunk[$x] = StringReplace($splunk[$x], @CRLF, " ")
		$splunk[$x] = StringReplace($splunk[$x], @LF, " ")
		$splunk[$x] = StringStripWS($splunk[$x], 3)
		$splunk[$x] = StringReplace($splunk[$x], "&quot;", "``")
		$splunk[$x] = StringReplace($splunk[$x], "'", "")
		If $splunk[$x] <> "" Then
			$splunk2[$counter][0] = $splunk[$x]
			If StringInStr($splunk2[$counter][0], "#") <> 0 Then
				$splunk2[$counter][0] = StringReplace($splunk2[$counter][0], 'valign="top"><td><pre>', "")
				$splunk2[$counter][0] = StringReplace($splunk2[$counter][0], '</pre></td></tr>', "")
				$splunk2[$counter][0] = StringReplace($splunk2[$counter][0], '</table></body></html>', "")
				For $y = 1 To 9
					If $y = 9 Then
						$splunk2[$counter][1] = StringLeft($splunk2[$counter][0], StringInStr($splunk2[$counter][0], "#", 0, 1) - 1)
						$splunk2[$counter][0] = "ROW"
						ContinueLoop
					ElseIf $y = 2 Then
						$splunk2[$counter][$y] = StringRight($splunk2[$counter][$y], StringLen($splunk2[$counter][$y]) - 4)
						$splunk2[$counter][$y] = StringReplace($splunk2[$counter][$y], "  ", " ")
						$mon = StringLeft($splunk2[$counter][$y], 3)
						If $mon = "Jan" Then
							$mon = "01"
						ElseIf $mon = "Feb" Then
							$mon = "02"
						ElseIf $mon = "Mar" Then
							$mon = "03"
						ElseIf $mon = "Apr" Then
							$mon = "04"
						ElseIf $mon = "May" Then
							$mon = "05"
						ElseIf $mon = "Jun" Then
							$mon = "06"
						ElseIf $mon = "Jul" Then
							$mon = "07"
						ElseIf $mon = "Aug" Then
							$mon = "08"
						ElseIf $mon = "Sep" Then
							$mon = "09"
						ElseIf $mon = "Oct" Then
							$mon = "10"
						ElseIf $mon = "Nov" Then
							$mon = "11"
						ElseIf $mon = "Dec" Then
							$mon = "12"
						EndIf
						$year = StringRight($splunk2[$counter][$y], 4)
						$day = StringMid($splunk2[$counter][$y], StringInStr($splunk2[$counter][$y], " ") + 1, StringInStr($splunk2[$counter][$y], " ", 0, 2) - StringInStr($splunk2[$counter][$y], " ") - 1)
						If StringLen($day) = 1 Then
							$day = "0" & $day
						EndIf
						$time = StringMid($splunk2[$counter][$y], StringInStr($splunk2[$counter][$y], " ", 0, 2) + 1, 8)
						$splunk2[$counter][$y] = $year & "/" & $mon & "/" & $day & " " & $time
					EndIf
					$start = StringInStr($splunk2[$counter][0], "#", 0, $y) + 1
					$end = StringInStr($splunk2[$counter][0], "#", 0, $y + 1)
					$diff = $end - $start
					$splunk2[$counter][$y + 1] = StringMid($splunk2[$counter][0], $start, $diff)
				Next
			Else
				If StringInStr($splunk[$x], "<a href=") <> 0 Then
					$splunk[$x] = StringReplace($splunk[$x], 'Link to results: <a href="', "")
					$splunk[$x] = StringLeft($splunk[$x], StringInStr($splunk[$x], '"') - 1)
					$splunk2[$counter][0] = "URL"
					$splunk2[$counter][1] = $splunk[$x]
				Else
					$splunk2[$counter][0] = StringMid($splunk[$x], 1, StringInStr($splunk[$x], ":") - 1)
					$splunk2[$counter][1] = StringMid($splunk[$x], StringInStr($splunk[$x], ":") + 1, StringLen($splunk[$x]))
					$splunk2[$counter][0] = StringStripWS($splunk2[$counter][0], 3)
					$splunk2[$counter][1] = StringStripWS($splunk2[$counter][1], 3)
				EndIf
			EndIf
			$counter += 1
			ReDim $splunk2[$counter + 1][10]
		EndIf
	Next
	$staticvalues = ""
	$rowvalues = ""
	$counter = 0
	$counter2 = 10
	For $y = 1 To UBound($splunk2) - 2
		If $splunk2[$y][0] = "ROW" Then
			For $z = 0 To UBound($splunk2, 2) - 1
				$splunk3[$counter][$z] = $splunk2[$y][$z]
				If $z = UBound($splunk2, 2) - 1 Then
					If $splunk3[$counter][10] = "" Then
						$splunk3[$counter][10] = $splunk3[$counter - 1][10]
						$splunk3[$counter][11] = $splunk3[$counter - 1][11]
						$splunk3[$counter][12] = $splunk3[$counter - 1][12]
						$splunk3[$counter][13] = $splunk3[$counter - 1][13]
					EndIf
					$counter += 1
					ReDim $splunk3[$counter + 1][14]
				EndIf
			Next
		Else
			$splunk3[$counter][$counter2] = $splunk2[$y][1]
			$counter2 += 1
		EndIf
	Next
	ReDim $splunk3[$counter][14]
	$dinsert = ""
	For $y = 0 To UBound($splunk3) - 1
;~ 		If $y = 1 Then
		$query = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & _
				$splunk3[$y][2] & "', '" & $splunk3[$y][3] & "', 'Splunk', 'Normal', '" & $splunk3[$y][10] & "', 'New', '', '');"
		If Not _EzMySql_Exec($query) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf

		$dinsert = "insert into splunkalertdetails (alertid,raw,spltime,splhost,splindex,linecount,splsource,sourcetype,splunk_server,tag,description,Query_Terms,Results,Alert_Trigger" & _
				",html_source) values ((select max(alertid) from alerts),'" & $splunk3[$y][1] & "','" & $splunk3[$y][2] & "','" & $splunk3[$y][3] & "','" & $splunk3[$y][4] & _
				"','" & $splunk3[$y][5] & "','" & $splunk3[$y][6] & "','" & $splunk3[$y][7] & "','" & $splunk3[$y][8] & "','" & $splunk3[$y][9] & "','" & $splunk3[$y][10] & "','" & _
				$splunk3[$y][11] & "','" & $splunk3[$y][12] & "','" & $splunk3[$y][13] & "','" & $message & "');"
		If Not _EzMySql_Exec($dinsert) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf

		$ainsert = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
		If Not _EzMySql_Exec($ainsert) Then
			_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
			_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
			Return
		EndIf
	Next
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Splunk")
EndFunc   ;==>Splunk

Func Clariion()
	$emails[$c][1] = StringReplace($emails[$c][1], "Time Stamp", "|" & "Time Stamp")
	$clariion = StringSplit($emails[$c][1], "|")
	Dim $clariion4[2][11]
	$clariion4[0][0] = "X"
	$clariion4[0][1] = "Time Stamp"
	$clariion4[0][2] = "Severity"
	$clariion4[0][3] = "Storage Array"
	$clariion4[0][4] = "Description"
	$clariion4[0][5] = "Event Number"
	$clariion4[0][6] = "Host"
	$clariion4[0][7] = "Device"
	$clariion4[0][8] = "SP"
	$clariion4[0][9] = "C"
	$clariion4[0][10] = "ShortDesc"
	$expandrow = 2
	For $clarx = 1 To UBound($clariion) - 1
		If StringInStr($clariion[$clarx], "Time Stamp") <> 0 Then
			$clariion[$clarx] = StringReplace($clariion[$clarx], " (GMT) ", " ")
			$clariion[$clarx] = StringReplace($clariion[$clarx], "Event Number", @CRLF & "Event Number")
			$clariion[$clarx] = StringReplace($clariion[$clarx], "Host", @CRLF & "Host")
			$clariion[$clarx] = StringReplace($clariion[$clarx], " SP", @CRLF & "SP", 1)
			$clariion[$clarx] = StringReplace($clariion[$clarx], " Device", @CRLF & "Device")
			$clariion[$clarx] = StringReplace($clariion[$clarx], " Storage Array", @CRLF & "Storage Array")
			$clariion[$clarx] = StringReplace($clariion[$clarx], " Severity", @CRLF & "Severity")
			$clariion[$clarx] = StringReplace($clariion[$clarx], " Description", @CRLF & "Description")
			$clariion[$clarx] = StringReplace($clariion[$clarx], "N/A", "NA")
			$clariion[$clarx] = StringReplace($clariion[$clarx], @CRLF & " ", " ")
			$clariion2 = StringSplit($clariion[$clarx], @CRLF)
			$clariion2[0] = "C=" & $c & ", X=" & $clarx
			$rows = UBound($clariion2)
			$details = ""
			Dim $clariion3[1][2]
			$clarz = 0
			$query = ""
			$query2 = ""
			For $clary = 0 To UBound($clariion2) - 1
				If $clariion4[$expandrow - 1][0] = "" Then
					$clariion4[$expandrow - 1][0] = $clarx
					$clariion4[$expandrow - 1][9] = $c
				EndIf
				If $clariion4[$expandrow - 1][0] <> $clarx Then
					$expandrow += 1
					ReDim $clariion4[$expandrow][11]
				EndIf
				If $clariion2[$clary] <> "" Then
					$clarz += 1
					ReDim $clariion3[$clarz][2]
					$clariion3[$clarz - 1][0] = $clariion2[$clary]
					If StringInStr($clariion3[$clarz - 1][0], "Time Stamp") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Time Stamp ", "")
						$clariion3[$clarz - 1][0] = "Time Stamp"
						$clariion4[$expandrow - 1][1] = $clariion3[$clarz - 1][1]
						$year = "20" & StringMid($clariion4[$expandrow - 1][1], 7, 2)
						$mon = StringMid($clariion4[$expandrow - 1][1], 1, 2)
						$day = StringMid($clariion4[$expandrow - 1][1], 4, 2)
						$time = StringMid($clariion4[$expandrow - 1][1], 10, 8)
						$clariion4[$expandrow - 1][1] = $year & "/" & $mon & "/" & $day & " " & $time
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Severity") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Severity ", "")
						$clariion3[$clarz - 1][1] = StringStripWS($clariion3[$clarz - 1][1], 3)
						If $clariion3[$clarz - 1][1] = "Critical Error" Then
							$clariion3[$clarz - 1][1] = "Critical"
						EndIf
						$clariion3[$clarz - 1][0] = "Severity"
						$clariion4[$expandrow - 1][2] = $clariion3[$clarz - 1][1]
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Storage Array") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Storage Array ", "")
						$clariion3[$clarz - 1][0] = "Storage Array"
						$clariion4[$expandrow - 1][3] = $clariion3[$clarz - 1][1]
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Description") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Description ", "")
						$clariion3[$clarz - 1][0] = "Description"
						$clariion4[$expandrow - 1][4] = $clariion3[$clarz - 1][1]
						If StringInStr($clariion4[$expandrow - 1][4], "heartbeat") <> 0 Then
							$clariion4[$expandrow - 1][10] = "Heartbeat"
						ElseIf StringInStr($clariion4[$expandrow - 1][4], "received too many events ") <> 0 Then
							$clariion4[$expandrow - 1][10] = "Excessive Events"
						ElseIf StringInStr($clariion4[$expandrow - 1][4], "(") <> 0 Then
							$clariion4[$expandrow - 1][10] = StringLeft($clariion4[$expandrow - 1][4], StringInStr($clariion4[$expandrow - 1][4], "(") - 1)
						ElseIf StringInStr($clariion4[$expandrow - 1][4], "mutex") <> 0 Then
							$clariion4[$expandrow - 1][10] = "Bad Mutex"
						Else
							$clariion4[$expandrow - 1][10] = $clariion4[$expandrow - 1][4]
						EndIf
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Event Number") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Event Number ", "")
						$clariion3[$clarz - 1][0] = "Event Number"
						$clariion4[$expandrow - 1][5] = $clariion3[$clarz - 1][1]
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Host") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Host ", "")
						$clariion3[$clarz - 1][0] = "Host"
						$clariion4[$expandrow - 1][6] = $clariion3[$clarz - 1][1]
					ElseIf StringInStr($clariion3[$clarz - 1][0], "Device") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][1] = StringReplace($clariion3[$clarz - 1][1], "Device ", "")
						$clariion3[$clarz - 1][0] = "Device"
						$clariion4[$expandrow - 1][7] = $clariion3[$clarz - 1][1]
					ElseIf StringInStr($clariion3[$clarz - 1][0], "SP") <> 0 Then
						$clariion3[$clarz - 1][1] = $clariion3[$clarz - 1][0]
						$clariion3[$clarz - 1][0] = "SP"
						$clariion4[$expandrow - 1][8] = $clariion3[$clarz - 1][1]
					EndIf
				EndIf
			Next
		EndIf
	Next
	$severity = ""
	$description = ""
	For $clary = 1 To UBound($clariion4) - 1
		If $severity = "" Then
			$severity = $clariion4[$clary][2]
		ElseIf $severity = "Error" And $clariion4[$clary][2] = "Critical" Then
			$severity = "Critical"
		ElseIf $severity = "Information" And $clariion4[$clary][2] <> "Information" Then
			$severity = $clariion4[$clary][2]
		EndIf
		If $description = "" And UBound($clariion4) < 3 Then
			$description = $clariion4[$clary][4]
		ElseIf UBound($clariion4) > 2 Then
			$description = "Multiple"
		EndIf
	Next
	For $clary = 1 To UBound($clariion4) - 1
		If $clary = 1 Then
			If $description = "Multiple" Then
				$query = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & _
						$clariion4[$clary][1] & "', '" & $clariion4[$clary][3] & "', 'Clariion', '" & $severity & "', '" & $description & "', 'New', '', '');"
				If Not _EzMySql_Exec($query) Then
					_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query, "", $olFormatHTML, $olImportanceNormal)
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
					Return
				EndIf
			Else
				$query = "SET SQL_SAFE_UPDATES=0;Insert into Alerts(Date, Device, AlertType, Severity, Description, Status, AssignedTo, TicketNumber) VALUES ('" & _
						$clariion4[$clary][1] & "', '" & $clariion4[$clary][3] & "', 'Clariion', '" & $severity & "', '" & $clariion4[$clary][10] & "', 'New', '', '');"
				If Not _EzMySql_Exec($query) Then
					_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $query, "", $olFormatHTML, $olImportanceNormal)
					_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
					Return
				EndIf
			EndIf

			$dinsert = "insert into clariionalertdetails (Time_Stamp,Severity,Storage_Array,Description,Event_Number,Host,Device,SP,alertid) values ('" & _
					$clariion4[$clary][1] & "', '" & $clariion4[$clary][2] & "', '" & $clariion4[$clary][3] & "', '" & $clariion4[$clary][4] & "', '" & $clariion4[$clary][5] & _
					"', '" & $clariion4[$clary][6] & "', '" & $clariion4[$clary][7] & "', '" & $clariion4[$clary][8] & "',(select max(alertid) from alerts));"
			If Not _EzMySql_Exec($dinsert) Then
				_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
				_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
				Return
			EndIf

			$ainsert = "insert into activitylog(AlertID, loggedby, timestamp, action) values ((select max(alertid) from alerts), 'System', sysdate(), 'Received');"
			If Not _EzMySql_Exec($ainsert) Then
				_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $ainsert, "", $olFormatHTML, $olImportanceNormal)
				_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
				Return
			EndIf
		Else
			$dinsert = "insert into clariionalertdetails (Time_Stamp,Severity,Storage_Array,Description,Event_Number,Host,Device,SP,alertid) values ('" & _
					$clariion4[$clary][1] & "', '" & $clariion4[$clary][2] & "', '" & $clariion4[$clary][3] & "', '" & $clariion4[$clary][4] & "', '" & $clariion4[$clary][5] & _
					"', '" & $clariion4[$clary][6] & "', '" & $clariion4[$clary][7] & "', '" & $clariion4[$clary][8] & "',(select max(alertid) from alerts));"
			If Not _EzMySql_Exec($dinsert) Then
				_OL_Wrapper_SendMail($outlook, "email", "", "", "NBC Script Insert Error", $dinsert, "", $olFormatHTML, $olImportanceNormal)
				_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Needs Review")
				Return
			EndIf
		EndIf
	Next
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\Clariion")
EndFunc   ;==>Clariion

Func Report()
	$NewDate = _DateAdd('d', -1, _NowCalcDate())
	$NewDate = StringReplace($NewDate, "/", "-")
	$array = _EzMySql_GetTable2d("SELECT alerttype, count(*) FROM dsbpws01.alerts where assignedto <> 'System' and remediatedon like '" & $NewDate & "%' group by alerttype having count(*)>0")

	$array2 = _EzMySql_GetTable2d("SELECT alerttype, count(*) FROM dsbpws01.alerts where assignedto <> 'System' and remediatedon like '" & $NewDate & "%' and ticketnumber <> " & '""' & _
			" and ticketnumber not like '111%' and ticketnumber not like '0000%' group by alerttype having count(*)>0")

	Dim $array3[UBound($array)][3]

	$z = 0

	For $x = 0 To UBound($array) - 1
		$array3[$z][0] = $array[$x][0]
		$array3[$z][1] = $array[$x][1]
		For $y = 0 To UBound($array2) - 1
			If $array[$x][0] = $array2[$y][0] Then
				$array3[$z][2] = $array2[$y][1]
			EndIf
		Next
		$z += 1
	Next

	$message = '<table border="1" width="100%">' & _
			'<col width="0">' & _
			'<tr bgcolor="#6699FF">' & _
			'<th><center>Alert Type</center></th>' & _
			'<th><center>Alert Total</center></th>' & _
			'<th><center>Ticket Total</center></th>' & _
			'</tr>'

	For $x = 1 To UBound($array3) - 1
		$submessage = '<tr>' & _
				'<td><center>' & $array3[$x][0] & '</center></td>' & _
				'<td><center>' & $array3[$x][1] & '</center></td>' & _
				'<td><center>' & $array3[$x][2] & '</center></td>' & _
				'</tr>'
		$message = $message & $submessage
	Next

	$message = $message & "</table>"

	_OL_Wrapper_SendMail($outlook, "email; email", "", "", "Dashboard Report - " & @MON & "/" & @MDAY & "/" & @YEAR, $message, "", $olFormatHTML, $olImportanceNormal)
EndFunc   ;==>Report

Func HPSIMReport()
	$attachment = _OL_ItemAttachmentGet($outlook, $emails[$c][5])
	$attach = ""
	For $rpty = 1 To UBound($attachment) - 1
		_OL_ItemAttachmentSave($outlook, $emails[$c][5], Default, $rpty, $path & $attachment[$rpty][1])
		If StringInStr($attachment[$rpty][1], ".GIF") = 0 Or StringInStr($emails2, $attachment[$rpty][1]) <> 0 Then
			$attach = $attach & $path & $attachment[$rpty][1]
			If $rpty <> UBound($attachment) - 1 Then
				$attach = $attach & ";"
			EndIf
		EndIf
	Next

	$emails2 = StringReplace($emails2, '</td></thead> <tr', '</td><td>Recent Alerts</td></thead> <tr')
	$emails2 = StringReplace($emails2, '</td></tr> <tr', '</td><td>My Data</td></tr> <tr')
	$emails2 = StringRegExpReplace($emails2, '\"\sALT=\"(.*?)\"', '"')
	$emails2 = StringReplace($emails2, '</td></tr>', '</td><td>My Data</td></tr>', -1)
	Local $aArray = StringRegExp($emails2, '(?<="><td>)(.*?)(?=<\/td><td>)', 3)
	If IsArray($aArray) Then
		For $rpty = 0 To UBound($aArray) - 1
			$results = _EzMySql_GetTable2d("select distinct(description) from alerts where device='" & StringUpper($aArray[$rpty]) & "' ORDER BY alertid DESC LIMIT 10")
			If IsArray($results) Then
				$output = ""
				For $rptz = 1 To UBound($results) - 1
					$output = $output & $results[$rptz][0]
					If $rptz < UBound($results) - 1 Then
						$output = $output & ",<br><br>"
					EndIf
				Next
				$emails2 = StringRegExpReplace($emails2, 'My Data', $output, 1)
			Else
				$emails2 = StringRegExpReplace($emails2, 'My Data', "", 1)
			EndIf
		Next
	EndIf
	_OL_Wrapper_SendMail($outlook, "email", "", "", $emails[$c][0], $emails2, $attach, $olFormatHTML, $olImportanceNormal)
	_OL_ItemMove($outlook, $emails[$c][5], $folder[3], "MailboxName\Inbox\HPSIM Reports")
	FileDelete($path & "*.*")
EndFunc   ;==>HPSIMReport