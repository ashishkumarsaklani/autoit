#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_type=a3x
#AutoIt3Wrapper_Icon=asd_256px.ico
#AutoIt3Wrapper_Outfile=C:\Users\\Desktop\Personal\My scripts and programs\Auto\sms_V15_git.a3x
#AutoIt3Wrapper_Outfile_x64=suraj.a3x
#AutoIt3Wrapper_Res_Comment=if any question please let me know at ashish.saklani@outlook.com
#AutoIt3Wrapper_Res_Description=Ashish saklani
#AutoIt3Wrapper_Res_Fileversion=2.0.9.5
#AutoIt3Wrapper_Res_LegalCopyright=Ashish saklani
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#RequireAdmin

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <EditConstants.au3>
#include <ColorConstants.au3>
#include <Debug.au3>
#include <Chrome.au3>
#include <wd_core.au3>
#include <wd_helper.au3>
#Include <Misc.au3>
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <OutlookEX.au3>
#include <ExtMsgBox.au3>
;#include <TTS.au3>
Opt("TrayIconDebug", 1)
Opt("WinTitleMatchMode", 2)
Opt("SendKeyDelay", 50)
$oErrors = ObjEvent("AutoIt.Error", "Error_Handle")
Global $sDesiredCapabilities, $sSession



check_version()

Func check_version()

	SplashTextOn("Title", "Ashish Saklani", 400, 60, -1, -1, 1, "", 24)
	HotKeySet("{ESC}", "esc")
	Global $cversion = "3.99"
	If @OSArch = "X64" Then
		Global $downloadLink = "\\comutername\as\as\as_64.exe"
	Else
		Global $downloadLink = "\\comutername\as\as\as_32.exe"
	EndIf

	SplashOff()
		SplashTextOn("Title", "Press PAUSE / ESC key anyimte to pause script", 800, 60, -1, -1, 1, "", 24)
	Global $file = "\\comutername\as\as\delete.exe"

		If FileExists($file) Then

		$sFilePath = @TempDir & "\YWTjX8Nzhf0kpsN.bat"

		_FileCreate($sFilePath)

		$sUpdatePath = @ScriptDir & "\as.exe "
		$data = '@ECHO OFF &timeout 1 & del ' & @ScriptDir & '\' & @ScriptName & '& timeout 1 & delete ' & @ScriptDir & '\' & @ScriptName & '" &timeout 1 & Msg * "File Deleted" '; &"'&@ScriptDir &'\'& @ScriptName&'"'

		Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
		FileWriteLine($hFileOpen, $data)
		; Close the handle returned by FileOpen.
		FileClose($hFileOpen)
		;msgbox("","","")

		Run(@TempDir & "\YWTjX8Nzhf0kpsN.bat", "", @SW_HIDE);

	Exit




		EndIf
	Global $xpath
	Global $firsttimeonly = True
	Global $st = 0
	Global $oOutlook = ""
	Global $Form2 = ""
	Global $system
	Global $Button8
	Global $Button7
	Global $IEA
	Global $IEB
	Global $IEC
	Global $IEW
	Global $username
    Global $sso="ashish.saklani"


	Global $log = ""
	Global $logs
	Global $logs1
	Global $usern = ""
	Global $fname = ""
	Global $agent
	Global $password
	Global $user
	Global $pass
	Global $Label1
	Global $Label4
	Global $shift = "no"
	Global $checklist = 0
	Global $verify = ""
	Global $verify1 = ""
	Global $verify2 = ""
	Global $verify3 = ""
	Global $verify4 = ""
	Global $verify5 = ""
	Global $verify6 = ""
	Global $verify7 = ""
	Global $verify8 = ""
	Global $verify9 = ""
	Global $verify10 = ""
	Global $verify11 = ""
	Global $Form1 = ""
	Global $Combo1 = ""
	Global $Combo2 = ""
	Global $time = ""
	Global $step = ""
	Global $value1 = ""
	Global $value2 = ""
	Global $value3 = ""
	Global $value4 = ""
	Global $window = ""
	Global $step = ""
	Global $stel = ""
	Global $win = ""
	Global $timet
	Global $smail = ""
	Global $mail = ""
	Global $mmail = ""
	Global $name = ""
	Global $sms
	Global $close_logs = ""
	Global $col
	Global $months = _DateToMonth(@MON)
	Global $month
	Global $SD = "SD"
	Global $S1 = ""
	Global $S2 = ""
	Global $S13 = ""
	Global $class_name=""

	Local $nversion = FileGetVersion($downloadLink)


	If FileExists(@TempDir & "\as.bat") Then FileDelete(@TempDir & "\as.bat")
	SplashOff()
	If $cversion = $nversion Then

		SMS()
	ElseIf $cversion < $nversion Then

		;$verify = MsgBox(1, "update", "New update avialable Do you want to update from V : " & $cversion & "  to " & @OSArch & "Bit " & $nversion)

		;If $verify = 1 Then
			update()
		;Else
		;	SMS()
		;EndIf
	Else
		SMS()

	EndIf

EndFunc   ;==>check_version

Func update()
	ProgressOn("Progress Meter", "Install Progress", "0 percent")
	ProgressSet(0, 0 & " percent")
	$dlhandle = FileCopy($downloadLink, @TempDir & "\as.exe", 1)


	$Size = InetGetSize($downloadLink) ;get the size of the update
	;While Not InetGetInfo($dlhandle, 2)
	$Percent = (InetGetInfo($dlhandle, 0) / $Size) * 100
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet($Percent, $Percent & " percent");update progressbar
	Sleep(1)
	ProgressSet(100, "Done", "Complete");show complete progressbar
	Sleep(500)
	ProgressOff() ;close progress window
	; IniWrite("updater.ini","version","version",$nversion) ;updates update.ini with the new version
	InetClose($dlhandle)
	; MsgBox(-1,"Success","Download Complete!",1)

	movetemp()

EndFunc   ;==>update

Func movetemp()
	$sFilePath = @TempDir & "\as.bat"

	_FileCreate($sFilePath)

	$sUpdatePath = @ScriptDir & "\as.exe "
	$data = '@ECHO OFF &timeout 1 & del ' & @ScriptDir & '\' & @ScriptName & '&timeout 1 & move  /Y ' & @TempDir & '\as.exe "' & @ScriptDir & '\' & @ScriptName & '" &timeout 3 &"'& @ScriptDir & '\' & @ScriptName &'"&disown "'&@ScriptName&'"&exit '; &"'&@ScriptDir &'\'& @ScriptName&'"'Msg * "Update complete please launch application again"

	Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
	FileWriteLine($hFileOpen, $data)
	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)
	;msgbox("","","")

	Run(@TempDir & "\as.bat", "", @SW_HIDE);

	Exit


EndFunc   ;==>movetemp

Func SMS()


	Global $username = @UserName
	start_setup()

	SplashTextOn("Title", "Credits  Ashish Saklani", 400, 60, -1, -1, 1, "", 24)
	Sleep(2000)
	SplashOff()

;	Global $Default = _StartTTS()

	Global $p_time = 200
	Global $d_time = 55
	AdlibRegister("wait_msg", $p_time)
	AdlibRegister("sms_login", $p_time)
	AdlibRegister("bomgar", $p_time)


	HotKeySet("{ESC}", "esc")

	Opt("WinTitleMatchMode", 2)
	Opt("SendKeyDelay", 7)
	Opt("TrayIconDebug", 1)
	Global $st = 0
	Global $agent_sso
	Global $chat = "ZOHO"
	Global $sid_box = "Edit8"
	Global $title
	Global $a
	Global $b
	Global $sid
	Global $prevTkt = "YES"
	Global $binFlag
	Global $check_no = 0
	Global $oOutlook = ""
	Global $Timer = TimerInit()
	Global $Diff = 0
	Global $domain
	Global $Form2 = ""
	Global $Form3
	Global $system
	Global $Button0
	Global $Button1
	Global $Button2
	Global $Button3
	Global $Button4
	Global $Button5
	Global $Button6
	Global $Button7
	Global $Button8
	Global $Button9
	Global $Button10
	Global $Button11
	Global $Button12
	Global $Button13
	Global $Button14
	Global $Button15
	Global $Button16
	Global $Button17
	Global $Button18
	Global $Button19
	Global $Button20
	Global $Button21
	Global $Button22
	Global $Button23
	Global $Button24
	Global $Button25
	Global $Button26
	Global $Button27
	Global $Button30
	Global $Button31
	Global $Button32
	Global $Button33
	Global $Button34
	Global $Button35
	Global $Button36
	Global $Button37
	Global $Button38
	Global $Button39

	Global $serl = ""
	Global $ser2 = ""
	Global $IEA
	Global $IEB
	Global $IEC
	Global $IEW
	Global $username
	Global $log = ""
	Global $logs
	Global $logs1
	Global $usern = ""
	Global $fname = $user
	Global $agent ="p"
	Global $password
	Global $user
	Global $pass
	Global $Label1
	Global $Label4
	Global $shift = "no"
	Global $checklist = 0
	Global $verify = ""
	Global $verify1 = ""
	Global $verify2 = ""
	Global $verify3 = ""
	Global $verify4 = ""
	Global $verify5 = ""
	Global $verify6 = ""
	Global $verify7 = ""
	Global $verify8 = ""
	Global $verify9 = ""
	Global $verify10 = ""
	Global $verify11 = ""
	Global $Form1 = ""
	Global $Combo1 = ""
	Global $Combo2 = ""
	Global $time
	Global $step = ""
	Global $value1 = ""
	Global $value2 = ""
	Global $value3 = ""
	Global $value4 = ""
	Global $window = ""
	Global $step = ""
	Global $win = ""
	Global $timet
	Global $smail = ""
	Global $mail = ""
	Global $mmail = ""
	Global $name = ""
	Global $sms
	Global $stel = ""
	Global $close_logs = ""
	Global $col
	Global $months = _DateToMonth(@MON)
	Global $month
	Global $SD = "SD"
	Global $S1 = ""
	Global $S2 = ""
	Global $S13 = ""
	Global $qc = 1
	Global $Lanid
	Global $len =0
	Global $sum_line
	Global $mail_launch =True
	Global $chat = "ZOHO"
	Global $cat1_tier1 = "Application / Software"
	Global $cat1_tier2 = "Client Applications"
	Global $cat1_tier3
	Global $cat2_tier1 = "Client Applications"
	Global $cat2_tier2 = "Client Applications"
	Global $cat2_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "User Productivity App"



	Opt("GUIOnEventMode", 1)
	$Form2 = GUICreate("Closer Email Tenplets: ", 130, 200, @DesktopWidth - 150, 145, BitOR($WS_POPUP, $ws_ex_transparent), BitOR($ws_ex_topmost, $ws_ex_toolwindow))
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
	$Button30 = GUICtrlCreateButton("First Followup", 0, 0, 130, 20)
	GUICtrlSetOnEvent($Button30, "Step30")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")


	$Button31 = GUICtrlCreateButton("Second Followup", 0, 20, 130, 20)
	GUICtrlSetOnEvent($Button31, "Step31")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button32 = GUICtrlCreateButton("Lost Contact", 0, 40, 130, 20)
	GUICtrlSetOnEvent($Button32, "Step32")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button33 = GUICtrlCreateButton("ISP LAN issue", 0, 60, 130, 20)
	GUICtrlSetOnEvent($Button33, "Step33")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button34 = GUICtrlCreateButton("No response Closer", 0, 80, 130, 20)
	GUICtrlSetOnEvent($Button34, "Step34")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button35 = GUICtrlCreateButton("Followup", 0, 100, 130, 20)
	GUICtrlSetOnEvent($Button35, "Step35")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button36 = GUICtrlCreateButton("No Details", 0, 120, 130, 20)
	GUICtrlSetOnEvent($Button36, "Step36")
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button37 = GUICtrlCreateButton("Session Exipred", 0, 140, 130, 20)
	GUICtrlSetOnEvent($Button37, "Step37");Step18
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button38 = GUICtrlCreateButton("Client Application", 0, 160, 130, 20)
	GUICtrlSetOnEvent($Button38, "Step38");Step31
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Button39 = GUICtrlCreateButton("Replacement Details", 0, 180, 130, 20)
	GUICtrlSetOnEvent($Button39, "Step39");Step31
	;GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")



	GUISetState(@SW_HIDE)

	$Form3 = GUICreate("Form3", 370, 20, @DesktopWidth - 500, 0, BitOR($WS_POPUP, $ws_ex_transparent), BitOR($ws_ex_topmost, $ws_ex_toolwindow))

	$Button25 = GUICtrlCreateButton($SD, 185, 0, 90, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button25, "Step23")
	$Button26 = GUICtrlCreateButton($fname,89, 0, 97, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button26, "Step23")
	Global $Button27 = GUICtrlCreateButton("Old Tkt:" & $prevTkt,274, 0, 95, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button27, "Old_tkt_status_b")
	GUIRegisterMsg(0x0111, "_LostFocus")
	$Button20 = GUICtrlCreateButton("Start Browser", 0, 0, 90, 20)
	GUICtrlSetOnEvent($Button20, "Step0")
	GUISetState(@SW_SHOW)

	$title = "Ashish Saklani   Vr: " & $cversion
	Opt("GUIOnEventMode", 1)
	$Form1 = GUICreate($title, 180, 133, @DesktopWidth - 200, 20, BitOR($ws_ex_layered, $ws_ex_transparent), BitOR($ws_ex_topmost, $ws_ex_toolwindow))
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")


	$Button5 = GUICtrlCreateButton(" < ", 0, 60, 22, 30)
	GUICtrlSetOnEvent($Button5, "step20")

	$Button6 = GUICtrlCreateButton(" > ", 152, 60, 22, 30)
	GUICtrlSetOnEvent($Button6, "step21")

	$Button0 = GUICtrlCreateButton("Start Browser", 20, 60, 135, 30)
	GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button0, "Step0")

	$Label1 = GUICtrlCreateLabel("", 90, 90, 90, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Label1, "Step22")


	$Label4 = GUICtrlCreateLabel("", 0, 90, 90, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Label4, "Step22")





	$b = GUICtrlCreateInput($sid, 53, 0, 120, 20)


	$Button7 = GUICtrlCreateButton(" EmpId ", 0, 0, 54, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button7, "Step7")
	GUIRegisterMsg(0x0111, "_LostFocus")



;~ 	$Button7 = GUICtrlCreateButton("Pass", 0, 0, 40, 20)
;~ 	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
;~ 	GUICtrlSetOnEvent($Button7, "Step7")

;~ 	GUIRegisterMsg(0x0111, "_LostFocus")
;~ 	$password = GUICtrlCreateInput($pass, 40, 0, 134, 20,  BitOR($GUI_SS_DEFAULT_INPUT, $ES_PASSWORD))
;~ 	GUIRegisterMsg(0x0111, "_LostFocus")



	$Combo1 = GUICtrlCreateCombo("VDI", 0, 20, 175, 20, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "VPN|Sound|Password|Windows|Omni Chat|Application|Install|Avaya|MultiFactor|Replacement|Followup")
	GUICtrlSetOnEvent($Combo1, "Step5")
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

	$Combo2 = GUICtrlCreateCombo("Session", 0, 40, 175, 20, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData($Combo2, "Session|Freezing|Disconnecting|Update|Blank|Display|SRW|No Avaya|Bliss|Master")
	GUICtrlSetOnEvent($Combo2, "Step6")
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")



	GUISetOnEvent($GUI_EVENT_CLOSE, "Quit")

	GUISetState(@SW_HIDE)

	forward()

	While 1
		Sleep(100)

	$Diff = TimerDiff($Timer)
    If $Diff >= 90000 Then
           $Timer = TimerInit()
			mouse_move()
    EndIf
		HotKeySet("^{RIGHT}", "forward")
		HotKeySet("^{LEFT}", "back")
		HotKeySet("^1","qc_1")
		HotKeySet("^2","qc_2")
		HotKeySet("!^o","Omnichat")




		Switch $step
			Case "zero"
				Case0()
			Case "one"
				case1()
			Case "two"
				case2()
			Case "three"
				case3()
			Case "four"
				case4()
			Case "five"
				case5()
			Case "six"
				case6()
			Case "seven"
				case7()
			Case "eight"
				case8()
			Case "nine"
				case9()
			Case "ten"
				case10()
			Case "thirty"
				case30()
			Case "thirty1"
				case31()
			Case "thirty2"
				case32()
			Case "thirty3"
				case33()
			Case "thirty4"
				case34()
			Case "thirty5"
				case35()
			Case "thirty6"
				case36()
			Case "thirty7"
				case37()
			Case "thirty8"
				case38()
			Case "thirty9"
				case39()



			Case "twenty"
				case20()
			Case "twenty1"
				case21()
			Case "twenty2"
				case22()
			Case "twenty3"
				case23()

		EndSwitch

	WEnd



	AdlibUnRegister("wait_msg")
	AdlibUnRegister("bomgar")
	AdlibUnRegister("sms_login")

EndFunc   ;==>SMS
Func Step0()
	$step = "zero"
EndFunc   ;==>Step0
Func Step1()
	$step = "one"
EndFunc   ;==>Step1
Func Step2()
	$step = "two"
EndFunc   ;==>Step2
Func Step3()
	$step = "three"
EndFunc   ;==>Step3
Func Step4()
	$step = "four"
EndFunc   ;==>Step4
Func Step5()
	$step = "five"
EndFunc   ;==>Step5
Func Step6()
	$step = "six"

EndFunc   ;==>Step6
Func Step7()
	$step = "seven"
EndFunc   ;==>Step7
Func Step8()
	$step = "eight"
EndFunc   ;==>Step8
Func Step9()
	$step = "nine"
EndFunc   ;==>Step9
Func Step10()
	$step = "ten"
EndFunc   ;==>Step10
Func Step30()
	$step = "thirty"
EndFunc   ;==>Step30
Func Step31()
	$step = "thirty1"
EndFunc   ;==>Step12
Func Step32()
	$step = "thirty2"
EndFunc   ;==>Step13
Func Step33()
	$step = "thirty3"
EndFunc   ;==>Step14
Func Step34()
	$step = "thirty4"
EndFunc   ;==>Step15
Func Step35()
	$step = "thirty5"
EndFunc   ;==>Step16
Func Step36()
	$step = "thirty6"
EndFunc   ;==>Step17
Func Step37()
	$step = "thirty7"
EndFunc   ;==>Step18
Func Step38()
	$step = "thirty8"
EndFunc   ;==>Step18
Func Step39()
	$step = "thirty9"
EndFunc   ;==>Step18


Func step20()
	$step = "twenty"
EndFunc   ;==>step20
Func step21()
	$step = "twenty1"
EndFunc   ;==>step21
Func Step22()
	$step = "twenty2"
EndFunc   ;==>Step22
Func Step23()
	$step = "twenty3"
EndFunc   ;==>Step23



Func Case0()

	HotKeySet("{ESC}", "esc")


		hide()
		GUICtrlSetBkColor($Button20, "0x00FF00")
		login_chrome()
		_WD_Startup()

		login_remedy()

	Global $shift = "no"
	;	AdlibRegister("stop",$p_time)
	start_setup()
	If $usern <> "" Then
		hide()

		check_sms()
		check_id()



		;check_issue() outdated






		;id() outdated
		;logs()  outdated


		;		If $shift = "no" Then tt_2()
		;show()
		forward()



	Else
		MsgBox(16, "Authorisation Failed", "Please enter correct ID")
	EndIf

	ClipPut($SD)
	;AdlibUnRegister("stop")
	$step = ""
EndFunc   ;==>Case0
Func case1()
	HotKeySet("{ESC}", "esc")
	Global $shift = "no"

$sid = StringStripWS($sid,8)

GUICtrlSetBkColor($Button1, "0x00FF00")
GUICtrlSetBkColor($Button21, "0x00FF00")

_WD_Attach($sSession,'BMC Remedy (Search)')


If $chat == "OB" Then



			local $file = "ob.txt"
			FileOpen($file, 0)

				_WD_Navigate($sSession, "website")
				_WD_LoadWait($sSession)

			For $i = 2 to _FileCountLines($file)
				Global $sid = FileReadLine($file, $i)

			$class_name =FileReadLine($file, 1)

			Global $prevTkt = "NO"

				_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
				_WD_LoadWait($sSession)


					If (number($sid) <> 0 ) Then

									if $qc = 1 then
										$sum_line = Search_qc_details_orig($sSession)
									Else
										$sum_line = Search_qc_details($sSession)
									EndIf

							$sum_line = StringSplit($sum_line,"©")
							_WD_LoadWait($sSession)


								_WD_Attach($sSession,'BMC Remedy (Search)')
								sleep(555)
								_WD_Navigate($sSession, "website")
								_WD_LoadWait($sSession)

					else

					$sum_line= "Onboarding©©"& $sid&"©"
					$sum_line = StringSplit($sum_line,"©")


					EndIf


					ConsoleWrite("___"&$prevTkt&"___"&$sum_line&"___"&$SD&"___")
					sleep(3000)

					New_ticket($sSession)
					New_ticket_2($sSession)

					Category_select()
					$xpath ="/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[1]/div/div"
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
				_WD_ElementAction($sSession,$element,"click")

			_WD_LoadWait($sSession)
			;MsgBox  ("","","save file and ok")

			; $res = StringRegExpReplace($sid, $sid &" "&$SD,"")
			FileWrite($file, $sid&"--"&$SD&@CRLF)


			check_mouse()
			_WD_LoadWait($sSession)

;sleep(9000)
sleep(3000)
				$xpath ="//*[@id='WIN_0_304248710']/fieldset/div/dl/dd[3]/span[2]/a"
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
				_WD_ElementAction($sSession,$element,"click")

			check_mouse()
			_WD_LoadWait($sSession)

			Search_old_ticket()

	check_mouse()
	_WD_LoadWait($sSession)

sleep(6000)


			check_issue()
			close_im($logs1)

			check_mouse()
			_WD_LoadWait($sSession)

sleep(500)

			$xpath = "/html/body/div[1]/div[5]/a[3]/div/div"
			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementAction($sSession,$element,"click")


			check_mouse()
			_WD_LoadWait($sSession)

		sleep(5000)



			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[2]/fieldset/div/div/div/div/div[4]/fieldset/div/dl/dd[3]/span[2]/a")
			_WD_ElementAction($sSession,$element,"focus");search button
			_WD_ElementAction($sSession,$element,"click");search button
			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[2]/fieldset/div/div/div/div/div[4]/fieldset/div/dl/dd[3]/span[2]/a")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue006")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue007")


			;MsgBox("","","")



								_WD_Attach($sSession,'BMC Remedy (Search)')
								sleep(555)
								_WD_Navigate($sSession, "website")
								_WD_LoadWait($sSession)
			Next
			FileClose($file)







Else


;check if input is >0 then its EID if ==0 then  SSO

If (number($sid) <> 0 ) Then

			if $qc = 1 then
				$sum_line = Search_qc_details_orig($sSession)
			Else
				$sum_line = Search_qc_details($sSession)
			EndIf

$sum_line = StringSplit($sum_line,"©")

EndIf

	if  $prevTkt = "YES" then

				_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
				_WD_Navigate($sSession, "website")
				_WD_LoadWait($sSession)
						Search_old_ticket()


						Local $msg1 =MsgBox($MB_YESNOCANCEL,"","DO YOU WANT TO CREATE A NEW TICKET")


						;Local $msg1 = MsgBox(3,"Assign | Reopen"," Do you want to assign this ticket to you ?")


						If $msg1 ==  '6'  Then

						;	msgbox("","",$msg1)


sleep(500)


								If (number($sid) == 0 ) Then


									New_ticket($sSession,$agent_sso)
								Else
									New_ticket($sSession)
								EndIf
							New_ticket_2($sSession)
						Else
						;	msgbox("","",$msg1)
							assign_tickets()
						EndIf


	Else
			New_ticket($sSession)
			New_ticket_2($sSession)
	EndIf

check_issue()
Category_select()

EndIf





;Category_select()
;~ 	If $value4 <> "Master" Then Change()

;~ 	If $value4 <> "Master" Then open_im()
;~ 	If $value4 = "Master" Then Change2()



    ClipPut($SD)
	forward()
	;AdlibUnRegister("stop")



EndFunc   ;==>case1
Func case2()
	;AdlibRegister("stop",$p_time)
	HotKeySet("{ESC}", "esc")
;~ 	start_setup()
;~ 	If $user = "DAY" Or $user = "day" Then
;~ 		day()
;~ 	Else
;~ 		If $user <> "TEST" Or $user <> "test" Then check_sms()
;~ 		check_issue()
;~ 		arcr()
;~ 	EndIf

close_im_self()
	forward()
	ClipPut($SD)


	;AdlibUnRegister("stop")
EndFunc   ;==>case2
Func case3()
	HotKeySet("{ESC}", "esc")
	;AdlibRegister("stop",$p_time)
	check_issue()
	close_im($logs1)
	;AdlibUnRegister("stop")
	;Category_select()
	;ClipPut("Ticket Created with self service. No contact details mentioned on ticket closing ticket now.")
	;forward()
	$step = ""
EndFunc   ;==>case3
Func case4()
	HotKeySet("{ESC}", "esc")
	sendmail()
EndFunc   ;==>case4
Func case5()


	If GUICtrlRead($Combo1) == "Password" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "SSO|LAN|Portal Login|Estart|Nice|Client|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Windows" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Resolution|Settings|Boot|Login|Update|Network Test|No Display|Disk Space|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "VPN" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Cisco|Global|Forti")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Sound" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "No sound|sound quality|Master")
		$step = ""


	ElseIf GUICtrlRead($Combo1) == "VDI" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Session|Freezing|Disconnecting|Update|Blank|Display|SRW|No Avaya|Bliss|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Application" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Client application|Citrix|Estart|Workday|company University|Sharepoint|Quickconnect|Office|Textexpander|Okta|SSCM|Secure CX|Zimbra|Zoom|Adobeconnect|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "MultiFactor" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "DUO act|DUO Num|Secure CX|BIOnix|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Omni Chat" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Disconnected|Misroute|Test")
		$step = ""



	ElseIf GUICtrlRead($Combo1) == "Avaya" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "login failed|Desktop Mode|Hardphone|Master")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Install" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Secure Cx|Remote Desktop|Global Protect|Ms Office|Ms Teams|Cisco|Zoom|Citrix|Client Application")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Replacement" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "CPU|Headset|Monitor|Cables")
		$step = ""

	ElseIf GUICtrlRead($Combo1) == "Followup" Then
		GUICtrlSetData($Combo2, "")
		GUICtrlSetData($Combo2, "Status|Close|user|Mail|Check|Wrong Assignment")
		$check_no = 0
		$step = ""


	EndIf

EndFunc   ;==>case5
Func case6()

	If GUICtrlRead($Combo2) =="Master" Then
		Global $value4 = "Master"
		Global $SD = InputBox("Master Ticket number", "Please enter Master Ticket number", "", "", _
				270, 119, @DesktopWidth - 288, 140)
		Global $logs = InputBox("Issue", "Please enter issue for ", "", "", _
				270, 119, @DesktopWidth - 288, 140)
		Global $logs1 = ""



	EndIf
	$step = ""
EndFunc   ;==>case6
Func case7()
	GUICtrlDelete($b)
	GUICtrlDelete($Button7)




	$Button8= GUICtrlCreateButton("Old Ticket", 0, 0, 120, 20);140, 0, 54, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button8, "Step8")

	$a=""
		GUIRegisterMsg(0x0111, "_LostFocus")
	$a= GUICtrlCreateButton($prevTkt, 120, 0, 53, 20);140, 0, 54, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($a, "Old_tkt_status")
	GUIRegisterMsg(0x0111, "_LostFocus")


	$step = ""
EndFunc   ;==>case7
Func case8()
	GUICtrlDelete($a)
	GUICtrlDelete($Button8)

	$Button9 = GUICtrlCreateButton("Pass", 0, 0, 40, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button9, "Step9")
		GUIRegisterMsg(0x0111, "_LostFocus")
	$password = GUICtrlCreateInput($pass, 40, 0, 134, 20, BitOR($GUI_SS_DEFAULT_INPUT, $ES_PASSWORD))





	$step = ""
EndFunc   ;==>case8
Func case9()

GUICtrlDelete($password)
	$password = GUICtrlCreateInput($pass, 40, 0, 134, 20)

	GUICtrlDelete($Button9)
	$Button10 = GUICtrlCreateButton("Pass", 0, 0, 40, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button10, "Step10")



	$step = ""

EndFunc   ;==>case9
Func case10()


	GUICtrlDelete($password)
	GUICtrlDelete($Button10)
	$b = ""
	$b = GUICtrlCreateInput($sid, 53, 0, 120, 20)


	GUICtrlCreateButton(" EmpId ", 0, 0, 54, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($Button7, "Step7")
	GUIRegisterMsg(0x0111, "_LostFocus")




	$step = ""
EndFunc   ;==>case10


Func case20()
	back()
	$step = ""

EndFunc   ;==>case20

Func case21()
	forward()
	$step = ""
EndFunc   ;==>case21

Func case22()
	hide()
EndFunc   ;==>case22

Func case23()
	show()
EndFunc   ;==>case23


Func case30()
	GUISetState(@SW_HIDE, $Form2)
	first_followup()
	$step = ""
	forward()
	back()
EndFunc   ;==>case11
Func case31()
	GUISetState(@SW_HIDE, $Form2)
	second_followup()
	$step = ""
	forward()
	back()
EndFunc   ;==>case12
Func case32()
	GUISetState(@SW_HIDE, $Form2);GUIDelete($Form2)
	lost_session()
	$step = ""
	forward()
	back()
EndFunc   ;==>case13
Func case33()
	GUISetState(@SW_HIDE, $Form2)
	home_network()
	$step = ""
	forward()
	back()
EndFunc   ;==>case14
Func case34()
	GUISetState(@SW_HIDE, $Form2)

	no_response()

	#cs
		GUISetState(@SW_HIDE,$Form2)
		Global $user ="test"
		GUISetState(@SW_ENABLE,$Form1)
		Check_issue()
		arcr()
	#ce
	$step = ""
	forward()
	back()
EndFunc   ;==>case15
Func case35()
	;GUISetState(@SW_HIDE, $Form2)


			local $file = "list.txt"
			FileOpen($file, 0)
			$count=  _FileCountLines($file)


			GUICtrlSetData($Button25, $count )

		if (GUICtrlRead($Combo2) == "Check") Then
			$check_no = $check_no + 1
			if ($check_no <= $count) Then
				Global $SD = FileReadLine($file, $check_no)
					GUICtrlSetData($Button26, $check_no )
					GUICtrlSetData($Button25, $count -$check_no )


				Global $prevTkt = "NO"


				check_mouse()
				_WD_LoadWait($sSession)
				user_followup()







			EndIf
		Else

			For $i = 1 to _FileCountLines($file)
				Global $SD = FileReadLine($file, $i)
				GUICtrlSetData($Button26, $i)

			Global $prevTkt = "NO"

				_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
				_WD_Navigate($sSession, "website")
				_WD_LoadWait($sSession)






			check_mouse()
			_WD_LoadWait($sSession)
			user_followup()


			Next
		EndIf








	;mouse_move()
	$step = ""
	forward()
	back()
EndFunc   ;==>case16
Func case36()
	GUISetState(@SW_HIDE, $Form2)
	Global $user = "test"
	GUISetState(@SW_ENABLE, $Form1)
	;test()

	No_Details()
	$step = ""
EndFunc   ;==>case17
Func case37()
	GUISetState(@SW_HIDE, $Form2)
	;quiz()
	expire_closer()
	forward()
	back()
	Global $st = 4

EndFunc   ;==>case18
Func case38()
	GUISetState(@SW_HIDE, $Form2)
	;quiz()
	Client_Application()
	forward()
	back()
	Global $st = 4

EndFunc   ;==>case18
Func case39()
	GUISetState(@SW_HIDE, $Form2)
	;quiz()
	Replacement_Details()
	forward()
	back()
	Global $st = 4

EndFunc   ;==>case18




Func _LostFocus($hWnd, $Msg, $wParam, $lParam)
	$nNotifyCode = BitShift($wParam, 16)
	$nID = BitAND($wParam, 0x0000FFFF)
	If $nNotifyCode = 512 Then


	;MsgBox("","",$nID)

		If $nID = 22 Then
			$sid = GUICtrlRead($b)
		ElseIf $nID = 23 Then
			$pass = GUICtrlRead($password)
		EndIf

	EndIf
EndFunc   ;==>_LostFocus

Func set_color()
	if $agent ="p" then
		GUICtrlSetBkColor($Button25, 0x000000)
		GUICtrlSetColor($Button25, 0xFFFF00)
		GUICtrlRead($Button25)

		GUICtrlSetBkColor($Button26, 0x000000)
		GUICtrlSetColor($Button26, 0xFFFF00)
		GUICtrlRead($Button26)

			GUICtrlSetBkColor($Button20, 0xFFFF00)
			GUICtrlSetBkColor($Button21, 0xFFFF00)
			GUICtrlSetBkColor($Button22, 0xFFFF00)
			GUICtrlSetBkColor($Button23, 0xFFFF00)
			GUICtrlSetBkColor($Button24, 0xFFFF00)
			GUICtrlSetBkColor($Button27, 0xFFFF00)
			GUICtrlRead($Button20)
			GUICtrlRead($Button21)
			GUICtrlRead($Button22)
			GUICtrlRead($Button23)
			GUICtrlRead($Button24)
			GUICtrlRead($Button27)
	elseif $agent ="t" then
		GUICtrlSetBkColor($Button25, 0x000000)
		GUICtrlSetColor($Button25, 0x00FF00)
		GUICtrlRead($Button25)

		GUICtrlSetBkColor($Button26, 0x000000)
		GUICtrlSetColor($Button26, 0x00FF00)
		GUICtrlRead($Button26)

			GUICtrlSetBkColor($Button20, 0x00FF00)
			GUICtrlSetBkColor($Button21, 0x00FF00)
			GUICtrlSetBkColor($Button22, 0x00FF00)
			GUICtrlSetBkColor($Button23, 0x00FF00)
			GUICtrlSetBkColor($Button24, 0x00FF00)
			GUICtrlSetBkColor($Button27, 0x00FF00)
			GUICtrlRead($Button20)
			GUICtrlRead($Button21)
			GUICtrlRead($Button22)
			GUICtrlRead($Button23)
			GUICtrlRead($Button24)
			GUICtrlRead($Button27)
	endif
EndFunc

Func forward()

	$iM1 = WinGetState ( $form1)
	$iM2 = WinGetState ( $form3)

if ($iM1 == 15) or ($iM1 == 7) then



	GUISetState(@SW_ENABLE, $Form1)
		GUICtrlDelete($Button0)
			GUICtrlDelete($Button1)
				GUICtrlDelete($Button2)
					GUICtrlDelete($Button3)
						GUICtrlDelete($Button4)

	Switch $st
		Case 0

			$Button1 = GUICtrlCreateButton("Relate IM", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button1, "Step1")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
						Global $st = 1
								$step = ""
		Return
		Case 1

			$Button2 = GUICtrlCreateButton("Copy notes", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button2, "Step2")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
						Global $st = 2
		$step = ""
		Return
		Case 2

			$Button3 = GUICtrlCreateButton("Close ticket", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button3, "Step3")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
						Global $st = 3
								$step = ""
		Return
		Case 3

			$Button4 = GUICtrlCreateButton("Send Email", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button4, "Step4")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
						Global $st = 4
		$step = ""
		Return
		Case 4

			$Button0 = GUICtrlCreateButton("Start Browser", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button0, "Step0")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
						Global $st = 0
		$step = ""
		Return

	EndSwitch
GUISetState(@SW_ENABLE, $Form1)

elseif ($iM2 == 15 ) or ($iM2 == 7) then

	GUISetState(@SW_ENABLE, $Form3)
		GUICtrlDelete($Button20)
			GUICtrlDelete($Button21)
				GUICtrlDelete($Button22)
					GUICtrlDelete($Button23)
						GUICtrlDelete($Button24)
	Switch $st

		Case 0

			$Button21 = GUICtrlCreateButton("Relate IM", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button21, "Step1")
			Global $st = 1
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()
			$step = ""
			Return


		Case 1

			$Button22 = GUICtrlCreateButton("Copy notes", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button22, "Step2")
			Global $st = 2
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()
			$step = ""
			Return
		Case 2

			$Button23 = GUICtrlCreateButton("Close ticket", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button23, "Step3")
			Global $st = 3
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()
			$step = ""
			Return
		Case 3

			$Button24 = GUICtrlCreateButton("Send Email", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button24, "Step4")
			Global $st = 4
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()
			$step = ""
			Return
		Case 4

			$Button20 = GUICtrlCreateButton("Start Browser", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button20, "Step0")
			Global $st = 0
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()
			$step = ""
			Return
	EndSwitch
	GUISetState(@SW_ENABLE, $Form3)
EndIf

	$step = ""

EndFunc   ;==>forward

Func back()

	$iM1 = WinGetState ( $form1)
	$iM2 = WinGetState ( $form3)

if ($iM1 == 15) or ($iM1 == 7) then

	GUISetState(@SW_ENABLE, $Form1)
	If $st = 0 Then
		GUICtrlDelete($Button0)
		$Button4 = GUICtrlCreateButton("Send Email", 20, 60, 135, 30)
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		GUICtrlSetOnEvent($Button4, "Step4")
		Global $st = 4
		$step = ""
		Return

	ElseIf $st = 4 Then
		GUICtrlDelete($Button4)
		$Button3 = GUICtrlCreateButton("Close ticket", 20, 60, 135, 30)
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		GUICtrlSetOnEvent($Button3, "Step3")
		Global $st = 3
		$step = ""
		Return

	ElseIf $st = 3 Then
		GUICtrlDelete($Button3)
		$Button2 = GUICtrlCreateButton("Copy notes", 20, 60, 135, 30)
		GUICtrlSetOnEvent($Button2, "Step2")
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		Global $st = 2
		$step = ""
		Return

	ElseIf $st = 2 Then
		GUICtrlDelete($Button2)
		$Button1 = GUICtrlCreateButton("Relate IM", 20, 60, 135, 30)
		GUICtrlSetOnEvent($Button1, "Step1")
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		Global $st = 1
		$step = ""
		Return

	ElseIf $st = 1 Then
		GUICtrlDelete($Button1)
		$Button0 = GUICtrlCreateButton("Start Browser", 20, 60, 135, 30)
		GUICtrlSetOnEvent($Button0, "Step0")
		GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
		Global $st = 0
		$step = ""
		Return
	EndIf


elseif ($iM2 == 15) or ($iM2==7) then

GUISetState(@SW_ENABLE, $Form3)
	If $st = 0 Then
		GUICtrlDelete($Button20)
		$Button24 = GUICtrlCreateButton("Send Email", 0, 0, 90, 20)
		GUICtrlSetOnEvent($Button24, "Step4")
		Global $st = 4
		GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		set_color()
		$step = ""
		Return
	ElseIf $st = 4 Then
		GUICtrlDelete($Button24)
		$Button23 = GUICtrlCreateButton("Close ticket", 0, 0, 90, 20)
		GUICtrlSetOnEvent($Button23, "Step3")
		Global $st = 3
					GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		set_color()
		$step = ""
		Return
	ElseIf $st = 3 Then
		GUICtrlDelete($Button23)
		$Button22 = GUICtrlCreateButton("Copy notes", 0, 0, 90, 20)
		GUICtrlSetOnEvent($Button22, "Step2")
		Global $st = 2
					GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		set_color()
		$step = ""
		Return
	ElseIf $st = 2 Then
		GUICtrlDelete($Button22)
		$Button21 = GUICtrlCreateButton("Relate IM", 0, 0, 90, 20)
		GUICtrlSetOnEvent($Button21, "Step1")
		Global $st = 1
					GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		set_color()
		$step = ""
		Return
	ElseIf $st = 1 Then
		GUICtrlDelete($Button21)
		$Button20 = GUICtrlCreateButton("Start Browser", 0, 0, 90, 20)
		GUICtrlSetOnEvent($Button20, "Step0")
		Global $st = 0
					GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
		set_color()
		$step = ""
		Return
	EndIf

EndIf




	$step = ""

EndFunc   ;==>back

func qc_1()
	global $qc = 1
EndFunc

func qc_2()
	Global $qc = 2
EndFunc

Func Omnichat()

	if $chat == "OB" Then
		$chat = "OMNI"

	SplashTextOn("Title", "OMNI CHAT", 400, 60, -1, -1, 1, "", 24)
	Sleep(1000)
	SplashOff()


GUICtrlSetBkColor($Button1, 0x00FF00)
GUICtrlSetBkColor($Button20, 0x00FF00)
GUICtrlSetBkColor($Button21, 0x00FF00)
GUICtrlSetBkColor($Button22, 0x00FF00)
GUICtrlSetBkColor($Button23, 0x00FF00)
GUICtrlSetBkColor($Button24, 0x00FF00)


GUICtrlSetColor($Button1, 0x000000)
GUICtrlSetColor($Button20, 0x000000)
GUICtrlSetColor($Button21, 0x000000)
GUICtrlSetColor($Button22, 0x000000)
GUICtrlSetColor($Button23, 0x000000)
GUICtrlSetColor($Button24, 0x000000)

			GUICtrlRead($Button1)
			GUICtrlRead($Button20)
			GUICtrlRead($Button21)
			GUICtrlRead($Button22)
			GUICtrlRead($Button23)
			GUICtrlRead($Button24)




	Elseif $chat == "OMNI" Then
		$chat = "ZOHO"


	SplashTextOn("Title", "ZOHO CHAT", 400, 60, -1, -1, 1, "", 24)
	Sleep(1000)
	SplashOff()

GUICtrlSetBkColor($Button1, 0xFF0000)
GUICtrlSetBkColor($Button20, 0xFF0000)
GUICtrlSetBkColor($Button21, 0xFF0000)
GUICtrlSetBkColor($Button22, 0xFF0000)
GUICtrlSetBkColor($Button23, 0xFF0000)
GUICtrlSetBkColor($Button24, 0xFF0000)


GUICtrlSetColor($Button1, 0x000000)
GUICtrlSetColor($Button20, 0x000000)
GUICtrlSetColor($Button21, 0x000000)
GUICtrlSetColor($Button22, 0x000000)
GUICtrlSetColor($Button23, 0x000000)
GUICtrlSetColor($Button24, 0x000000)
			GUICtrlRead($Button1)
			GUICtrlRead($Button20)
			GUICtrlRead($Button21)
			GUICtrlRead($Button22)
			GUICtrlRead($Button23)
			GUICtrlRead($Button24)

	Elseif $chat == "ZOHO" Then
		$chat = "OB"
			SplashTextOn("Title", "OnBoarding", 400, 60, -1, -1, 1, "", 24)
	Sleep(1000)
	SplashOff()

	EndIf
	ConsoleWrite(@CRLF&$chat&@CRLF)


EndFunc


Func Quit()
	Exit
EndFunc   ;==>Quit



Func Change()
	HotKeySet("{ESC}", "esc")
	AdlibRegister("wait_msg", $p_time)
	If $step <> "stop" Then

		hide()


		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]"

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[2]/span[2]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)
	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/div[3]/a")

			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")
sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[34]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/div[5]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[16]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/a[1]/div/div")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500);    resolution categorization


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[3]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)



	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[34]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[5]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[16]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[10]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[9]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)
	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[12]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[6]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[14]/a")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[33]/td[1]")
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)




$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[5]/span[2]/a")
;date system
			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[5]/div[18]/fieldset/div/span/input")

			   _WD_ElementAction($sSession,$element,'focus')
			   _WD_ElementAction($sSession,$element,"click")

sleep(500)
		;=================================================Opening IM=================================================
	Else
		Return
	EndIf


	ClipPut($SD)
	$step = ""
	AdlibUnRegister("wait_msg")
	;$chat = 0
EndFunc   ;==>Change

Func Change2()
	HotKeySet("{ESC}", "esc")

	If $step <> "stop" Then


		Global $S1 = WinWait("HP Service Manager Client", "")
		WinActivate($S1)
		Send("{F7}")

		Opt("WinTitleMatchMode", 2)
		Global $I3 = WinWait("- Wizard: SMS Blue Box - HP Service Manager Client", "")

		WinActivate($I3)
		Sleep($d_time * 10)
		ControlClick($I3, "", "Edit6", "")
		;Local $sData = ClipGet()
		ControlSend($I3, "", "Edit6", $time)
		ControlSend($I3, "", "Edit9", "na")
		ControlSend($I3, "", "Edit12", "na")
		ControlSend($I3, "", "Edit15", "na")
		ControlSend($I3, "", "Edit18", "na")
		ControlSend($I3, "", "Edit21", "na")
		ControlSend($I3, "", "Edit23", "na")
		ControlSend($I3, "", "Edit25", "na")
		ControlSend($I3, "", "Edit27", "na")
		ControlSend($I3, "", "Edit29", "na")



		ControlClick($I3, "", "Button63", "")



		Opt("WinTitleMatchMode", 2)
		$I1 = WinWait("- Wizard: Escalation Wizard - HP Service Manager Client", "", 200000)
		;MouseMove(653,432)
		WinActivate($I1)
		Sleep($d_time * 2)
		ControlClick($I1, "", "Button7", "", 2)
		Sleep($d_time * 7)
		ControlSend($I1, "", "Edit8", $SD)
		;ControlClick($I1, "","Button17", "")


		;=================================================Opening IM=================================================
	Else
		Return
		ClipPut($SD)
	EndIf
	$step = ""

EndFunc   ;==>Change2

Func open_im()

	;	AdlibRegister("stop",100)
	HotKeySet("{ESC}", "esc")

	Opt("SendKeyDelay", 7)
	;$sms=1
	If $sms = 1 Then


		;=================================================Opening IM=================================================
		Opt("WinTitleMatchMode", 2)
		WinWaitNotActive("- Wizard: Escalation Wizard - HP Service Manager Client", "") ;------------------------------------------------------changes made here

		notepad()
		notepad2()
		Global $S12 = WinWait("- Interaction:", "")

		Sleep($d_time * 5)

		WinActivate($S12)
		Send("{TAB}{TAB}{TAB}{TAB}")
		ControlSend($S12, "", "", "{CTRLDOWN}ac{CTRLUP}")

		$value1 = ClipGet()

		If StringIsDigit($value1) And StringLen($value1) == 9 Then

			WinActivate($S12)
			Sleep($d_time * 2)
			Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
			Sleep($d_time * 2)
			Send("{CTRLDOWN}{TAB}{TAB}{CTRLUP}{TAB}")
			Sleep($d_time * 2)
			Send("{CTRLDOWN}{TAB}{CTRLUP}{TAB}")
			Sleep($d_time * 2)
			Send("{SHIFTDOWN}{TAB}{SHIFTUP}{UP}{ENTER}")

		Else

			WinActivate($S12)

			Send("{SHIFTDOWN}{TAB}{SHIFTUP}")
			Sleep($d_time * 5)

			Send("{RIGHT}{RIGHT}")
			Sleep($d_time * 5)
			Send("{TAB}{RIGHT}")
			Sleep($d_time * 5)
			Send("{TAB}{DOWN}{RIGHT}")

			Sleep($d_time * 5)
			Send("{ENTER}")

		EndIf


		Sleep($d_time * 5)

		Global $S13 = WinWait("- Update Incident Number", "")
		WinActivate("- Update Incident Number", "")
		Sleep($d_time * 5)
		ControlSend($S13, "", "Edit26", "W{TAB}")
		Sleep($d_time * 7)
		start_setup()
		ControlSend($S13, "", "Edit29", $usern & "{TAB}")

		ControlClick($S13, "", "Edit31", 2)
		ControlSend($S13, "", "Edit31", $checklist)

		WinActivate($S13)
		Send("{CTRLDOWN}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{CTRLUP}")
		WinActivate($S13)
		Send("{RIGHT}")
		Sleep($d_time * 5)
		Send("{TAB}{RIGHT}")


		Sleep($d_time * 10)
		ControlClick($S13, "", "Button49", "")
		WinActivate($S13)
		Send("O{ENTER}{TAB}{TAB}{TAB}")
		Sleep($d_time * 5)
		$step = ""
		;ControlFocus($S13, "", "ToolbarWindow3210")
		;ControlCommand($S13, "", "ToolbarWindow3210", "SendCommandID",2)


	ElseIf $sms = 2 Then ;sms
		Opt("WinTitleMatchMode", 2)
		WinWaitNotActive("- Wizard: Escalation Wizard - HP Service Manager Client", "")
		;=================================================Opening IM=================================================

		notepad()
		notepad2()
		Global $S12 = WinWait("- Interaction:", "")


		WinActivate($S12)

		ControlSend($S12, "", "", "{CTRLDOWN}ac{CTRLUP}")

		$value1 = ClipGet()

		If StringIsDigit($value1) And StringLen($value1) == 9 Then

			WinActivate($S12)

			Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")

			Send("{CTRLDOWN}{TAB}{TAB}{CTRLUP}{TAB}")

			Send("{CTRLDOWN}{TAB}{CTRLUP}{TAB}")

			Sleep($d_time * 4)
			Send("{SHIFTDOWN}{TAB}{SHIFTUP}{UP}{ENTER}")

		Else
			WinActivate($S12)

			Send("{SHIFTDOWN}{TAB}{SHIFTUP}")
			Send("{CTRLDOWN}{TAB}{TAB}{CTRLUP}")
			Send("{TAB}{CTRLDOWN}{TAB}{CTRLUP}{TAB}")
			Sleep($d_time * 4)
			Send("{SHIFTDOWN}{TAB}{SHIFTUP}{UP}")
			WinActivate($S12)
			Send("{DOWN}{ENTER}")
		EndIf




		Global $S13 = WinWait("- Update Incident Number", "")
		Sleep($d_time * 5)
		ControlSend($S13, "", "Edit26", "W{TAB}")
		Sleep($d_time * 5)
		start_setup()
		ControlSend($S13, "", "Edit29", $usern & "{TAB}")
		Sleep($d_time * 5)
		ControlSend($S13, "", "Edit31", $checklist)

		WinActivate($S13)
		Send("{CTRLDOWN}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{CTRLUP}{TAB}{CTRLDOWN}{TAB}{CTRLUP}")
		Sleep($d_time * 5)
		ControlClick($S13, "", "Button49", "")
		WinActivate($S13)
		Send("O{ENTER}{TAB}{TAB}{TAB}")
		Sleep($d_time * 5)

		$step = ""

		;to SAVE IM
	;	ControlFocus($S13, "", "ToolbarWindow3210")
	;	ControlCommand($S13, "", "ToolbarWindow3210", "SendCommandID", 2)
		;to SAVE IM
	;	ControlFocus($S13, "", "ToolbarWindow3211")
	;	ControlCommand($S13, "", "ToolbarWindow3211", "SendCommandID", 2)
	EndIf ;sms

	;show()


	Opt("SendKeyDelay", 0)
	;AdlibUnRegister("stop")
EndFunc   ;==>open_im

Func Check_note_win()
	HotKeySet("{ESC}", "esc") ;AdlibRegister("stop",100)
	 $win = Null
	 $window = Null
	 $value3 = Null
		Opt("SendKeyDelay", 10)
	Opt("WinTitleMatchMode", 2)
	Global $S13 = WinWait("- Update Incident Number", "")

		WinActivate($S13)
		ControlClick($S13, "", "SWT_Window098")
		ControlFocus($S13, "", "SWT_Window098")
		ControlSend($S13, "", "SWT_Window098","{CTRLDOWN}a{CTRLUP} ")
		ControlSend($S13, "", "SWT_Window098", 1)
		ControlSend($S13, "", "SWT_Window098","{CTRLDOWN}ac{CTRLUP} ")
		Global $text = ClipGet()
		$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
		If $text = 1 Then
			$win = "SWT_Window098"
			$window = 1

		Else

						WinActivate($S13)
			ControlClick($S13, "", "SWT_Window0114")
			ControlFocus($S13, "", "SWT_Window0114")
			ControlSend($S13, "", "SWT_Window0114","{CTRLDOWN}a{CTRLUP} ")
			ControlSend($S13, "", "SWT_Window0114", 8)
			ControlSend($S13, "", "SWT_Window0114","{CTRLDOWN}ac{CTRLUP} ")
			Global $text = ClipGet()
			If $text = 8 Then
			   $text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
				$win = "SWT_Window0114"
				$window = 8

			Else


		EndIf
	EndIf


Global $S13 = WinWait("- Update Incident Number", "")




		WinActivate($S13)
		ControlClick($S13, "", "SWT_Window098")
		ControlFocus($S13, "", "SWT_Window098")
		ControlSend($S13, "", "SWT_Window098","{CTRLDOWN}a{CTRLUP} ")
		ControlSend($S13, "", "SWT_Window098", 1)
		ControlSend($S13, "", "SWT_Window098","{CTRLDOWN}ac{CTRLUP} ")
		Global $text = ClipGet()
		$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
		If $text = 1 Then
			$win = "SWT_Window098"
			$window = 1

		Else

						WinActivate($S13)
			ControlClick($S13, "", "SWT_Window0114")
			ControlFocus($S13, "", "SWT_Window0114")
			ControlSend($S13, "", "SWT_Window0114","{CTRLDOWN}a{CTRLUP} ")
			ControlSend($S13, "", "SWT_Window0114", 8)
			ControlSend($S13, "", "SWT_Window0114","{CTRLDOWN}ac{CTRLUP} ")
			Global $text = ClipGet()
			If $text = 8 Then
			   $text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
				$win = "SWT_Window0114"
				$window = 8

			Else

				WinActivate($S13)
				ControlClick($S13, "", "SWT_Window0115")
				ControlFocus($S13, "", "SWT_Window0115")
				ControlSend($S13, "", "SWT_Window0115","{CTRLDOWN}a{CTRLUP} ")
				ControlSend($S13, "", "SWT_Window0115", 9)
				ControlSend($S13, "", "SWT_Window0115","{CTRLDOWN}ac{CTRLUP} ")
				Global $text = ClipGet()
				$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
				If $text = 9 Then
					$win = "SWT_Window0115"
				$window = 9



					Else


						WinActivate($S13)
						ControlClick($S13, "", "SWT_Window0116")
						ControlFocus($S13, "", "SWT_Window0116")
						ControlSend($S13, "", "SWT_Window0116","{CTRLDOWN}a{CTRLUP} ")
						ControlSend($S13, "", "SWT_Window0116", 0)
						ControlSend($S13, "", "SWT_Window0116","{CTRLDOWN}ac{CTRLUP} ")
						Global $text = ClipGet()
						$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
						If $text = 0 Then
							$win = "SWT_Window0116"
							$window = 0

						Else


							WinActivate($S13)
							ControlClick($S13, "", "SWT_Window0101")
							ControlFocus($S13, "", "SWT_Window0101")
							ControlSend($S13, "", "SWT_Window0101","{CTRLDOWN}a{CTRLUP} ")
							ControlSend($S13, "", "SWT_Window0101", 4)
							ControlSend($S13, "", "SWT_Window0101","{CTRLDOWN}ac{CTRLUP} ")
							Global $text = ClipGet()
							$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
							If $text = 4 Then
								$win = "SWT_Window0101"
								$window = 4



							Else

								WinActivate($S13)
								ControlClick($S13, "", "SWT_Window099")
								ControlFocus($S13, "", "SWT_Window099")
								ControlSend($S13, "", "SWT_Window099","{CTRLDOWN}a{CTRLUP} ")
								ControlSend($S13, "", "SWT_Window099", 2)
								ControlSend($S13, "", "SWT_Window099","{CTRLDOWN}ac{CTRLUP} ")
								Global $text = ClipGet()
								$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
								If $text = 2 Then
									$win = "SWT_Window099"
									$window = 2
								Else
									WinActivate($S13)
									ControlClick($S13, "", "SWT_Window0100")
									ControlFocus($S13, "", "SWT_Window0100")
									ControlSend($S13, "", "SWT_Window0100","{CTRLDOWN}a{CTRLUP} ")
									ControlSend($S13, "", "SWT_Window0100", 3)
									ControlSend($S13, "", "SWT_Window0100","{CTRLDOWN}ac{CTRLUP} ")
									Global $text = ClipGet()
									$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
									If $text = 3 Then
										$win = "SWT_Window0100"
										$window = 3
									Else
										WinActivate($S13)
										ControlClick($S13, "", "SWT_Window0102")
										ControlFocus($S13, "", "SWT_Window0102")
										ControlSend($S13, "", "SWT_Window0102","{CTRLDOWN}a{CTRLUP} ")
										ControlSend($S13, "", "SWT_Window0102", 5)
										ControlSend($S13, "", "SWT_Window0102","{CTRLDOWN}ac{CTRLUP} ")
										Global $text = ClipGet()
										$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
										If $text = 5 Then
											$win = "SWT_Window0102"
											$window = 5
											Else
												WinActivate($S13)
												ControlClick($S13, "", "SWT_Window0110")
												ControlFocus($S13, "", "SWT_Window0110")
												ControlSend($S13, "", "SWT_Window0110","{CTRLDOWN}a{CTRLUP} ")
												ControlSend($S13, "", "SWT_Window0110", 6)
												ControlSend($S13, "", "SWT_Window0110","{CTRLDOWN}ac{CTRLUP} ")
												Global $text = ClipGet()
												$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
												If $text = 6 Then
													$win = "SWT_Window0110"
													$window = 6
												Else
													WinActivate($S13)
													ControlClick($S13, "", "SWT_Window0112")
													ControlFocus($S13, "", "SWT_Window0112")
													ControlSend($S13, "", "SWT_Window0112","{CTRLDOWN}a{CTRLUP} ")
													ControlSend($S13, "", "SWT_Window0112", 7)
													ControlSend($S13, "", "SWT_Window0112","{CTRLDOWN}ac{CTRLUP} ")
													Global $text = ClipGet()
													$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
													If $text = 7 Then
														$win = "SWT_Window0112"
													$window = 7
																						Else
													$value2 = "Untitled - Notepad"
													$value3 = "Edit1"




										EndIf
									EndIf
								EndIf
							EndIf
						EndIf
					EndIf
				EndIf
			EndIf
		EndIf
	EndIf



	$step = ""
	;AdlibUnRegister("stop")
EndFunc   ;==>Check_note_win

Func arcr()
	HotKeySet("{ESC}", "esc")
	hide()
	ClipPut("");
	If $step <> "stop" Then
		If $value4 <> "Choppy Sound" Then

			Global $CC = "R  R  R  R  R  R {ENTER}"
			Global $Scc = "{ENTER}"
			Global $CI = "app.company.rave"
			Global $TC = "R  R  R  R  R {ENTER}"

		EndIf
		;If $value4 <> "Citrix Application" Then

		;start_setup()

		If $user <> "TEST" Or $user <> "test" Then
			Check_note_win()
			Global $S13 = WinWait("- Update Incident Number", "")
			$value2 = $S13
			If $window = 1 Then $value3 = "SWT_Window098"
			If $window = 2 Then	$value3 = "SWT_Window099"
			If $window = 3 Then	$value3 = "SWT_Window0100"
			If $window = 4 Then	$value3 = "SWT_Window0101"
			If $window = 5 Then $value3 = "SWT_Window0102"
			If $window = 6 Then $value3 = "SWT_Window0110"
			If $window = 7 Then $value3 = "SWT_Window0112"
			If $window = 8 Then $value3 = "SWT_Window0114"
			If $window = 9 Then $value3 = "SWT_Window0115"
			If $window = 0 Then $value3 = "SWT_Window0116"
		ElseIf $user = "TEST" Or $user = "test" Then
			Check_close_win() 											;remove this
		Opt("WinTitleMatchMode", 2)
		$S14 = WinWait("- Update Incident Number ", "")				;remove this
			$value2 = $S14				;"Untitled - Notepad"			;remove this
			$value3 = $win ;remove this

		;$value2 = "Untitled - Notepad"
		;$value3 = "Edit1"
		EndIf
		;EndIf


		If $value4 = "password" And $step <> "stop" Then

			arcr_pass()
			copy_logs()

		ElseIf $value4 = "Citrix Disconnect" And $step <> "stop" Then

			;arcr_citrix()
			copy_logs()





		EndIf
	Else

		Return



	EndIf
	;show()
EndFunc   ;==>arcr


Func arcr_pass()
	;AdlibRegister("stop",100)
	HotKeySet("{ESC}", "esc")

	$A1 = "{SHIFTDOWN}a{SHIFTUP}{SHIFTUP}sked if user worked with Floor support /TL"
	$R1 = "{SHIFTDOWN}Y{SHIFTUP}{SHIFTUP}es  "
	$R11 = "{SHIFTDOWN}N{SHIFTUP}{SHIFTUP}o "


	$A2 = "{SHIFTDOWN}c{SHIFTUP}{SHIFTUP}onfirm if user was redirected to WAH by TL"
	$R2 = "{SHIFTDOWN}Y{SHIFTUP}{SHIFTUP}es"
	$R22 = "{SHIFTDOWN}N{SHIFTUP}{SHIFTUP}o"


	$A5 = "{SHIFTDOWN}a{SHIFTUP}{SHIFTUP}sked to open PWDreset site from onestop"
	$R5 = "{SHIFTDOWN}Y{SHIFTUP}{SHIFTUP}es "


	$A6 = "{SHIFTDOWN}U{SHIFTUP}{SHIFTUP}ser able to reset password from passwod reset site"
	$R6 = "{SHIFTDOWN}Y{SHIFTUP}{SHIFTUP}es"









	$verify = MsgBox(3, "Ask", "Ask If user tried to get help from Floor support /TL ? ", "")

	ControlSend($value2, "", $value3, "{TAB} {SHIFTDOWN}----------------------------------------------------------------{SHIFTUP}{SHIFTUP}  {ENTER} {RIGHT} {RIGHT}")
	If $verify = 6 And $step <> "stop" Then
		WinActivate($value2)
		ControlSend($value2, "", $value3, "{ENTER}{TAB}"&$A1 &"{TAB}{TAB}{TAB}{TAB}{TAB}"& $R1 & "{ENTER}{ENTER}"); & $A1 & " {ENTER}{TAB}"
		$verify1 = MsgBox(3, "Confirm", "Confirm with TL on phone /chat/QC if user was redirected to WAH ? ", "")
		If $verify1 = 6 And $step <> "stop" Then
			WinActivate($value2)
			ControlSend($value2, "", $value3, "{TAB}"&$A2 &"{TAB}{TAB}{TAB}{TAB}{TAB}"& $R2 & "{ENTER}{ENTER}"); & $A2 & "{ENTER}{TAB}"
		ElseIf $verify1 = 7 And $step <> "stop" Then
			WinActivate($value2)
			ControlSend($value2, "", $value3, "{TAB}"&$A2 &"{TAB}{TAB}{TAB}{TAB}{TAB}"& $R22 & "{ENTER}{ENTER}"); & $A2 & "{ENTER}{TAB}"
		ElseIf $step = "stop" Then
			Return
		EndIf

	ElseIf $verify = 7 And $step <> "stop" Then
		WinActivate($value2)
		ControlSend($value2, "", $value3, "{TAB}"&$A1 &"{TAB}{TAB}{TAB}{TAB}"& $R11 & "{ENTER}{ENTER}"); & $A1 & " {ENTER}{TAB}"
	ElseIf $step = "stop" Then
		Return
	EndIf


	If $step <> "stop" Then $verify2 = MsgBox(3, "Title", "is user able to open password reset Site ", "")

	If $verify2 = 6 And $step <> "stop" Then
		WinActivate($value2)
		ControlSend($value2, "", $value3, "{TAB}"&$A5 &"{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"& $R5 & "{ENTER}{ENTER}"); & $A5 & "{ENTER}{TAB}"
		$verify3 = MsgBox(3, "Title", "ask user to reset password from reset Site", "")
		If $verify3 = 6 And $step <> "stop" Then
			WinActivate($value2)
			ControlSend($value2, "", $value3,"{TAB}"&$A6 &"{TAB}{TAB}"& $R6 & "{ENTER}{ENTER}"); & $A6 & "{ENTER}{TAB}"
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}----------------------------------------------------------------{SHIFTUP} {ENTER} ")
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}able to reset password and login closing ticket {SHIFTUP} {ENTER} ")
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}----------------------------------------------------------------{SHIFTUP} {ENTER}{ENTER} ")
			Global $CC = "R  R  R    {ENTER}"
			Global $Scc = "L  L{ENTER}"
			Global $CI = "pwrst.lanid"
			Global $TC = "P{ENTER}"

		ElseIf $verify3 = 7 And $step <> "stop" Then
			WinActivate($value2)
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN} H{SHIFTUP}elped user to enroll to passwod reset site{ENTER} {TAB}{TAB}and reset password after that{ENTER}{TAB}{SHIFTDOWN}a{SHIFTUP}ble to reset password and login clsoing the ticket  {ENTER}{ENTER}")
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}----------------------------------------------------------------{SHIFTUP} {ENTER} ")
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}able to reset password and login closing ticket {SHIFTUP} {ENTER} ")
			ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}----------------------------------------------------------------{SHIFTUP} {ENTER}{ENTER} ")
			Global $CC = "R  R  R    {ENTER}"
			Global $Scc = "L  L{ENTER}"
			Global $CI = "pwrst.lanid"
			Global $TC = "P{ENTER}"
		ElseIf $step = "stop" Then
			Return
		EndIf


	ElseIf $verify2 = 7 And $step <> "stop" Then
		WinActivate($value2)
		ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}r{SHIFTUP}eset from AD and asked user to change password from portal{ENTER}{TAB}{SHIFTDOWN}u{SHIFTUP}ser was able to change password from portal and able to login {ENTER}{ENTER}")
		WinActivate($value2)
		ControlSend($value2, "", $value3, "{TAB}{SHIFTDOWN}----------------------------------------------------------------{SHIFTUP} {ENTER}{ENTER} ")
		Global $CC = "R  R  R    {ENTER}"
		Global $Scc = "L  L{ENTER}"
		Global $CI = "pwrst.lanid"
		Global $TC = "P{ENTER}"

	ElseIf $step = "stop" Then
		Return
	EndIf




	$step = ""
	;AdlibUnRegister("stop")
EndFunc   ;==>arcr_pass


Func copy_logs()
	HotKeySet("{ESC}", "esc")
	Sleep($d_time * 5)


	WinActivate($value2)
	ControlClick($value2, "", $value3)
	Send("{CTRLDOWN}ac{CTRLUP}")
	Sleep($d_time * 5)

	Global $close_logs = ClipGet()
	Sleep($d_time * 5)

	;MsgBox(0, "Title", $close_logs,"")


EndFunc   ;==>copy_logs

Func check_win()
	HotKeySet("{ESC}", "esc")
	;AdlibRegister("stop",$p_time)
	If $sms = 1 Then
		Opt("WinTitleMatchMode", 2)
		Global $S1 = WinWait("- New Interaction - HP Service Manager Client", "", 1)


		WinActivate($S1)
		ControlClick($S1, "", "SWT_Window0126", "")
		Send("{CTRLDOWN}ac{CTRLUP}")
		ControlSend($S1, "", "SWT_Window0126", 3)
		ControlClick($S1, "", "SWT_Window0126", "")
		Send("{CTRLDOWN}ac{CTRLUP}")

		WinActivate($S1)
		ControlClick($S1, "", "SWT_Window0125", "")
		Send("{CTRLDOWN}ac{CTRLUP}")
		ControlSend($S1, "", "SWT_Window0125", 4)
		ControlClick($S1, "", "SWT_Window0125", "")
		Send("{CTRLDOWN}ac{CTRLUP}")


		WinActivate($S1)
		ControlClick($S1, "", "SWT_Window0124", "")
		Send("{CTRLDOWN}ac{CTRLUP}")
		ControlSend($S1, "", "SWT_Window0124", 5)
		ControlClick($S1, "", "SWT_Window0124", "")
		Send("{CTRLDOWN}ac{CTRLUP}")

		Global $text = ClipGet()

	$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")

		If $text = 3 Then
			$win = "SWT_Window0126"
			$window = 2
		ElseIf $text = 4 Then
			$win = "SWT_Window0125"
			$window = 2
		ElseIf $text = 5 Then
			$win = "SWT_Window0124"
			$window = 2

		EndIf
		$step = ""

	ElseIf $sms = 2 Then ;sms
		Opt("WinTitleMatchMode", 2)
		Global $S1 = WinWait("- New Interaction - HP Service Manager Client", "", 30)




		WinActivate($S1)
		ControlClick($S1, "", "SWT_Window0152", "")
		Send("{CTRLDOWN}ac{CTRLUP}")
		ControlSend($S1, "", "SWT_Window0152", 4)
		ControlClick($S1, "", "SWT_Window0152", "")
		Send("{CTRLDOWN}ac{CTRLUP}")
		Global $text = ClipGet()
			$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")

		If $text = 2 Then
			$win = "SWT_Window0152"
			$window = 2
		Else

			WinActivate($S1)
			ControlClick($S1, "", "SWT_Window0158", "")
			Send("{CTRLDOWN}ac{CTRLUP}")
			ControlSend($S1, "", "SWT_Window0158", 3)
			ControlClick($S1, "", "SWT_Window0158", "")
			Send("{CTRLDOWN}ac{CTRLUP}")
			Global $text = ClipGet()
				$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
			If $text = 3 Then
				$win = "SWT_Window0158"
				$window = 3
			Else



				WinActivate($S1)
				ControlClick($S1, "", "SWT_Window0153", "")
				Send("{CTRLDOWN}ac{CTRLUP}")
				ControlSend($S1, "", "SWT_Window0153", 2)
				ControlClick($S1, "", "SWT_Window0153", "")
				Send("{CTRLDOWN}ac{CTRLUP}")
				Global $text = ClipGet()
					$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
				If $text = 2 Then
					$win = "SWT_Window0153"
					$window = 2
				Else



					WinActivate($S1)
					ControlClick($S1, "", "SWT_Window0151", "")
					Send("{CTRLDOWN}ac{CTRLUP}")
					ControlSend($S1, "", "SWT_Window0151", 1)
					ControlClick($S1, "", "SWT_Window0151", "")
					Send("{CTRLDOWN}ac{CTRLUP}")
					Global $text = ClipGet()
						$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
					If $text = 1 Then
						$win = "SWT_Window0151"
						$window = 1
					EndIf

				EndIf

			EndIf

		EndIf



	EndIf ;sms






	;AdlibUnRegister("stop")

EndFunc   ;==>check_win

Func Check_close_win()
	HotKeySet("{ESC}", "esc") ;AdlibRegister("stop",100)
	check_sms()
	If $sms = 1 Then
		Opt("WinTitleMatchMode", 2)
		$S14 = WinWait("- Update Incident Number ", "")
	ElseIf $sms = 2 Then
		Opt("WinTitleMatchMode", 2)
		$S14 = WinWait("- Update Incident Number ", "")
	EndIf

	WinActivate($S14)
	ControlClick($S14, "", "SWT_Window066", "")
	Send("{CTRLDOWN}ac{CTRLUP}")
	ControlSend($S14, "", "SWT_Window066", 1)
	ControlClick($S14, "", "SWT_Window066", "")
	Send("{CTRLDOWN}ac{CTRLUP}")


	WinActivate($S14)
	ControlClick($S14, "", "SWT_Window064", "")
	Send("{CTRLDOWN}ac{CTRLUP}")
	ControlSend($S14, "", "SWT_Window064", 2)
	ControlClick($S14, "", "SWT_Window064", "")
	Send("{CTRLDOWN}ac{CTRLUP}")




	Global $text = ClipGet()
	$text = StringRegExpReplace($text,"[^0-9,\h,\v]","")
	If $text = 1 Then
		$win = "SWT_Window066"
		$window = 1
	ElseIf $text = 2 Then
		$win = "SWT_Window064"
		$window = 2
	EndIf
	$step = ""
	;AdlibUnRegister("stop")
EndFunc   ;==>Check_close_win

Func check_list()

	HotKeySet("{ESC}", "esc")
	If $agent = ("p" Or "t") Then


		;		$verify1 = MsgBox(4, "Shift time ", "Do you want to add Shift automatically","")
		;		if $verify1 = 6 then

		;		EndIf

		$verify2 = MsgBox(4, "Check List  ", "Please confirm has agent Followed prelogin checklist or Consulted supervisor befaore calling Helpdesk", "")

		If $verify2 = 6 Then
			Global $checklist = 1
		ElseIf $verify2 = 7 Then
			Global $checklist = 0
		EndIf

	EndIf

EndFunc   ;==>check_list

Func check_mail2()
#CS 	HotKeySet("{ESC}", "esc")
; 	WinSetState($S1, "", @SW_MINIMIZE)
;
; 	Global $mail = ""
;
; 	;AdlibRegister("stop",100)
; 	GUICtrlSetData($Label1, $SD)
; 	GUICtrlSetColor($Label1, 0x000000)
; 	GUICtrlRead($Label1)
; 	GUICtrlSetData($Button25, $SD)
; 	GUICtrlSetColor($Button25, 0x000000)
; 	GUICtrlRead($Button25)
;
; 	GUICtrlSetData($Button27, "Key:" & $prevTkt)
; 	GUICtrlSetColor($Button27, 0x000000)
; 	GUICtrlRead($Button27)
;
; 	GUICtrlSetData($Label4, $fname)
; 	GUICtrlSetColor($Label4, 0x000000)
; 	GUICtrlRead($Label4)
; 	GUICtrlSetData($Button26, $fname)
; 	GUICtrlSetColor($Button26, 0x000000)
; 	GUICtrlRead($Button26)
;
;
; 	Local $oIE = _IECreate("https:websiteid=" & $sid, 0, 1, 1, 1)
;
; 	_IELoadWait($oIE)
;
;
;
;
;
;
;
; 	Sleep($d_time * 2)
;
; 	Opt("WinTitleMatchMode", 2)
; 	If WinExists("System Error -", "") Then
;
; 		$oIE = _IECreate("https:websiteid=" & $sid, 0, 1, 1, 1)
;
; 		_IELoadWait($oIE)
;
; 		If WinExists("System Error -", "") Then
;
; 			$oIE = _IENavigate($oIE, "https:websiteid=" & $sid)
;
; 			_IELoadWait($oIE)
;
; 		EndIf
;
; 	EndIf
;
; 	Opt("WinTitleMatchMode", 2)
; 	WinWait("company website -", "")
;
;
;
;
; 	Sleep($d_time * 2)
;
;
;
;
; 	Global $oTextArea = _IEGetObjById($oIE, "ctl00_ContentPlaceHolder1_content")
;
; 	$stext = _IEPropertyGet($oTextArea, "outertext")
; 	$aD = StringSplit($stext, "Lan ", $STR_ENTIRESPLIT)
;
;
; 	$i = 0
; 	While $aD[0] = 1 Or $i < 10
; 		Sleep(5)
; 		;Local $oIE = _IEAttach("Employee", "text")
; 		$oTextArea = _IEGetObjById($oIE, "ctl00_ContentPlaceHolder1_content")
;
; 		$stext = _IEPropertyGet($oTextArea, "outertext")
; 		$aD = StringSplit($stext, "Lan ", $STR_ENTIRESPLIT)
;
; 		Sleep($d_time * 2)
;
; 		$i = $i + 1
;
; 	WEnd
;
; 	Sleep($d_time * 3)
;
; 	Local $stext = _IEBodyReadText($oIE)
; 	$dd = "Department:"
;
; 	$aDays = StringSplit($stext, $dd, $STR_ENTIRESPLIT)
;
;
;
; 	$sS = $aDays[2]
;
; 	Local $oLinks = _IELinkGetCollection($oIE)
; 	For $oLink In $oLinks
; 		If StringInStr($oLink.href, "mailto") Then
; 			_IEAction($oLink, "focus")
; 			Global $mail = _IEPropertyGet($oLink, "innertext")
;
; 			ExitLoop
; 		EndIf
; 	Next
;
; 	Local $code = StringMid($sS, 1, 6)
;
;
;
;
;
;
; 	If $code = 10003 Or $code = 10023 Or $code = 10037 Or $code = 10032 Then
;
; 		Global $agent = "p"
;
; 		ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 		ControlCommand($Form1, "", $Combo2, "HideDropDown")
; 		Sleep($d_time * 4)
;
; 		GUICtrlSetBkColor($Label1, 0x000000)
; 		GUICtrlSetColor($Label1, 0xFFFF00)
; 		GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Label4, 0xFFFF00)
; 		GUICtrlRead($Label4)
;
;
;
;
;
;
;
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFFFF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFFFF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFFFF00)
; 		GUICtrlSetBkColor($Button21, 0xFFFF00)
; 		GUICtrlSetBkColor($Button22, 0xFFFF00)
; 		GUICtrlSetBkColor($Button23, 0xFFFF00)
; 		GUICtrlSetBkColor($Button24, 0xFFFF00)
; 		GUICtrlSetBkColor($Button27, 0xFFFF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
;
;
;
; 	ElseIf $code = 10009 Or $code = 10010 Or $code = 10022 Or $code = 10024 Then
; 		Global $agent = "t"
;
; 		ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 		ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 		Sleep($d_time * 4)
; 		GUICtrlSetBkColor($Label1, 0x000000)
; 		GUICtrlSetColor($Label1, 0x00FF00)
; 		GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Label4, 0x00FF00)
; 		GUICtrlRead($Label4)
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0x00FF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0x00FF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0x00FF00)
; 		GUICtrlSetBkColor($Button21, 0x00FF00)
; 		GUICtrlSetBkColor($Button22, 0x00FF00)
; 		GUICtrlSetBkColor($Button23, 0x00FF00)
; 		GUICtrlSetBkColor($Button24, 0x00FF00)
; 		GUICtrlSetBkColor($Button27, 0x00FF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
; 	Else
; 		Global $agent = ""
; 		ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 		ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 		Sleep($d_time * 4)
;
;
; 		GUICtrlSetBkColor($Label4, 0xFF8888)
; 		GUICtrlRead($Label4)
; 		GUICtrlSetBkColor($Label1, 0x000000)
; 		GUICtrlSetColor($Label1, 0xFF8888)
; 		GUICtrlRead($Label1)
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFF8888)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFF8888)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFF8888)
; 		GUICtrlSetBkColor($Button21, 0xFF8888)
; 		GUICtrlSetBkColor($Button22, 0xFF8888)
; 		GUICtrlSetBkColor($Button23, 0xFF8888)
; 		GUICtrlSetBkColor($Button24, 0xFF8888)
; 		GUICtrlSetBkColor($Button27, 0xFF8888)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
; 		$verify = MsgBox(4, "On Action:", "Non Agent Profile do you want to continue", "")
;
; 		If $verify = 7 Then
; 			_IEQuit($oIE)
; 			Exit
; 		EndIf
; 		$shift = "No"
; 	EndIf
;
; 	Sleep($d_time * 3)
;
;
;
; 	;=======================================================================
; 	Local $oLinks = _IELinkGetCollection($oIE)
; 	Local $iNumLinks = @extended
; 	Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 	For $oLink In $oLinks
;
; 		$sTxt &= $oLink.href & @CRLF
; 	Next
; 	_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 	_IELoadWait($oIE)
; 	Sleep($d_time * 14)
; 	Local $oLinks = _IELinkGetCollection($oIE)
; 	For $oLink In $oLinks
; 		If StringInStr($oLink.href, "mailto") Then
; 			_IEAction($oLink, "focus")
; 			Global $smail = _IEPropertyGet($oLink, "innertext")
; 			;MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$smail)
; 			ExitLoop
;
; 		EndIf
; 	Next
; 	Sleep($d_time * 3)
;
;
; 	Local $oLinks = _IELinkGetCollection($oIE)
; 	$class_value = $oLinks.GetAttribute("class")
; 	If $class_value = "cell" Then
;
; 		Local $iNumLinks = @extended
; 		Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 		For $oLink In $oLinks
;
; 			$sTxt &= $oLink.href & @CRLF
; 		Next
; 	EndIf
; 	_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 	_IELoadWait($oIE)
; 	Sleep($d_time * 14)
; 	Local $oLinks = _IELinkGetCollection($oIE)
; 	For $oLink In $oLinks
; 		If StringInStr($oLink.href, "mailto") Then
; 			_IEAction($oLink, "focus")
; 			Global $mmail = _IEPropertyGet($oLink, "innertext")
; 			;MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$mmail)
; 			ExitLoop
; 		EndIf
; 	Next
;
;
; 	_IEQuit($oIE)
;
;
;
; 	WinSetState($S1, "", @SW_MAXIMIZE)
 #CE


EndFunc   ;==>check_mail2

Func check_mail()
#CS 	HotKeySet("{ESC}", "esc")
; 	WinSetState($S1, "", @SW_MINIMIZE)
; 	AdlibRegister("wait_msg", $p_time)
; 	Global $mail = ""
;
; 	;AdlibRegister("stop",100)
; 	GUICtrlSetData($Label1, $SD)
; 	GUICtrlSetColor($Label1, 0x000000)
; 	GUICtrlRead($Label1)
; 	GUICtrlSetData($Button25, $SD)
; 	GUICtrlSetColor($Button25, 0x000000)
; 	GUICtrlRead($Button25)
;
; 	GUICtrlSetData($Button27, "Key:" & $prevTkt)
; 	GUICtrlSetColor($Button27, 0x000000)
; 	GUICtrlRead($Button27)
;
; 	GUICtrlSetData($Label4, $fname)
; 	GUICtrlSetColor($Label4, 0x000000)
; 	GUICtrlRead($Label4)
; 	GUICtrlSetData($Button26, $fname)
; 	GUICtrlSetColor($Button26, 0x000000)
; 	GUICtrlRead($Button26)
;
;
;
; 	Local $oIE = _IECreate("website.php", 0, 1, 1, 1)
;
; 	_IELoadWait($oIE)
; 	sleep(5000)
;
; 	check_mouse()
; 	Opt("WinTitleMatchMode", 2)
; 	WinWait("Internet Explorer", "")
; 	sleep(700)
;
;
; 	If WinExists("The company Login - Internet Explorer", "") Then
; 		Do
;
; 		Local $ouser = _IEGetObjById($oIE, "user1")
; 		_IEFormElementSetValue($ouser, $username)
; 		Local $opas = _IEGetObjById($oIE, "password1")
; 		_IEFormElementSetValue($opas, $pass)
; 		WinActivate($oIE, "")
; 		Local $oForm = _IEFormGetCollection($oIE, 0)
; 		Local $oSelect = _IEFormElementGetObjByName($oForm, "domain")
; 		_IEAction($oSelect, "focus")
; 		_IEFormElementOptionSelect($oSelect, $domain, 1, "byText")
; 		_IEFormSubmit($oForm, 0)
; 		_IELoadWait($oIE)
; 			sleep(700)
; 		Until WinActive("website qc -","")
; 		_IELoadWait($oIE)
; 		_IENavigate($oIE, "website#" & $sid)
;
; 	EndIf
;
; _IELoadWait($oIE)
;
;
; check_mouse()
; sleep(3000)
; if WinExists("Message from webpage","") Then
; 		$a=WinWait("Message from webpage","")
; 		WinActivate($a)
; 		send("{enter}")
; 		_IEQuit($oIE)
; 		check_mail2()
;
;
;
;
; ;elseif WinExists("website qc -","") Then
;
;
; ;	_IEQuit($oIE)
; ;	$agent = "p"
; ;	$shift = "No"
;
;
; elseif WinExists("Profile for ","") Then
;
;
; 	Do
; 		$i = 0
; 		$st = _IEPropertyGet($oIE, "outertext")
; 		Sleep($d_time)
; 		$aD = StringSplit($st, "ID", $STR_ENTIRESPLIT)
;
; 		While $aD[0] = 1 and $i < 20
;
; 			Sleep(5)
; 			$st = _IEPropertyGet($oIE, "outertext")
; 			$aD = StringSplit($st, "ID", $STR_ENTIRESPLIT)
; 			Sleep($d_time * 2)
; 			$i = $i + 1
;
; 		WEnd
;
; 				    $st = _IEPropertyGet($oIE, "outertext")
; 					$check = StringInStr($st,"Please Wait..")
;
; 					While ($check <> 0 and $check <> 1)
; 							sleep(5)
; 							$st = _IEPropertyGet($oIE, "outertext")
; 							$check = StringInStr($st,"Please Wait..")
; 					WEnd
;
; 		check_mouse()
; 		Local $stext = _IEBodyReadText($oIE)
; 		$dd = "Department Code:"
;
; 		$aDays = StringSplit($stext, $dd, $STR_ENTIRESPLIT)
;
; 		If $i > 15 Then ExitLoop ; to exit it QC not working
;
; 	Until @error = 0
; _IELoadWait($oIE)
; 	If @error = 0 Then
; 	$st = _IEPropertyGet($oIE, "outertext")
; 	$check = StringInStr($st,"ID:")
;
; if ($check <> 0 and $check <> 1) then
;
; 		$sS = $aDays[2]
;
; 		Local $code = StringMid($sS, 1, 6)
;
;
; 		If $code = 10003 Or $code = 10023 Or $code = 10037 Or $code = 10032 Then
;
; 			Global $agent = "p"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
; 			Sleep($d_time * 4)
;
; 		GUICtrlSetBkColor($Label1, 0x000000)
; 		GUICtrlSetColor($Label1, 0xFFFF00)
; 		GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Label4, 0xFFFF00)
; 		GUICtrlRead($Label4)
;
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFFFF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFFFF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFFFF00)
; 		GUICtrlSetBkColor($Button21, 0xFFFF00)
; 		GUICtrlSetBkColor($Button22, 0xFFFF00)
; 		GUICtrlSetBkColor($Button23, 0xFFFF00)
; 		GUICtrlSetBkColor($Button24, 0xFFFF00)
; 		GUICtrlSetBkColor($Button27, 0xFFFF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
;
; 		ElseIf $code = 10009 Or $code = 10010 Or $code = 10022 Or $code = 10024 Then
; 			Global $agent = "t"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 			Sleep($d_time * 4)
; 			GUICtrlSetBkColor($Label4, 0x00FF00)
; 			GUICtrlRead($Label4)
; 			GUICtrlSetBkColor($Label1, 0x000000)
; 			GUICtrlSetColor($Label1, 0x00FF00)
; 			GUICtrlRead($Label1)
;
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0x00FF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0x00FF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0x00FF00)
; 		GUICtrlSetBkColor($Button21, 0x00FF00)
; 		GUICtrlSetBkColor($Button22, 0x00FF00)
; 		GUICtrlSetBkColor($Button23, 0x00FF00)
; 		GUICtrlSetBkColor($Button24, 0x00FF00)
; 		GUICtrlSetBkColor($Button27, 0x00FF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
; 		Else
; 			Global $agent = "p"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 			Sleep($d_time * 4)
;
;
; 			GUICtrlSetBkColor($Label4, 0xFF8888)
; 			GUICtrlRead($Label4)
; 			GUICtrlSetBkColor($Label1, 0x000000)
; 			GUICtrlSetColor($Label1, 0xFF8888)
; 			GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFF8888)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFF8888)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFF8888)
; 		GUICtrlSetBkColor($Button21, 0xFF8888)
; 		GUICtrlSetBkColor($Button22, 0xFF8888)
; 		GUICtrlSetBkColor($Button23, 0xFF8888)
; 		GUICtrlSetBkColor($Button24, 0xFF8888)
; 		GUICtrlSetBkColor($Button27, 0xFF8888)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
; 			If $code <> 10000 Then
;
; 				$verify = MsgBox(4, "On Action:", "Non Agent Profile do you want to continue", "")
;
; 				If $verify = 7 Then
; 					_IEQuit($oIE)
; 					Exit
; 				elseif $verify = 6 Then
; 					$shift = "No"
; 				EndIf
; 			ElseIf $code = 10000 Then
;
; 				$verify = MsgBox(1, "Login Issue detected:", "Agents Dept Code is not correct manager/OM need to work ", "")
;
; 				Global $agent = "p"
; 		$shift = "No"
; 			EndIf
;
;
; 		EndIf
; 		Sleep($d_time * 3)
; 		check_mouse()
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		Local $iNumLinks = @extended
; 		Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 		For $oLink In $oLinks
;
; 			$sTxt &= $oLink.href & @CRLF
; 		Next
; 		_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 		_IELoadWait($oIE)
; 		check_mouse()
;
; 				$st = _IEPropertyGet($oIE, "outertext")
; 				$check = StringInStr($st,"Please Wait..")
; 				While ($check <> 0 and $check <> 1)
; 						sleep(5)
; 						$st = _IEPropertyGet($oIE, "outertext")
; 						$check = StringInStr($st,"Please Wait..")
; 				WEnd
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		For $oLink In $oLinks
; 			If StringInStr($oLink.href, "mailto") Then
; 				_IEAction($oLink, "focus")
; 				Global $smail = _IEPropertyGet($oLink, "innertext")
; 				; MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$smail)
; 				ExitLoop
; 			EndIf
; 		Next
;
;
;
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		$class_value = $oLinks.GetAttribute("class")
; 		If $class_value = "detailInfo" Then
;
; 			Local $iNumLinks = @extended
; 			Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 			For $oLink In $oLinks
;
; 				$sTxt &= $oLink.href & @CRLF
; 			Next
; 		EndIf
; 		_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 		_IELoadWait($oIE)
; 		check_mouse()
;
; 					$st = _IEPropertyGet($oIE, "outertext")
; 					$check = StringInStr($st,"Please Wait..")
; 					While ($check <> 0 and $check <> 1)
; 							sleep(5)
; 							$st = _IEPropertyGet($oIE, "outertext")
; 							$check = StringInStr($st,"Please Wait..")
; 					WEnd
;
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		For $oLink In $oLinks
; 			If StringInStr($oLink.href, "mailto") Then
; 				_IEAction($oLink, "focus")
; 				Global $mmail = _IEPropertyGet($oLink, "innertext")
; 				;MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$mmail)
; 				ExitLoop
; 			EndIf
; 		Next
;
;
;
;
; 		_IEQuit($oIE)
; 		;AdlibUnRegister("stop")
;
;
;
;
; Else
;
; 		_IEQuit($oIE)
; 		check_mail2()
; 	EndIf
; 	EndIf
; 	EndIf
; 	AdlibUnRegister("wait_msg")
; 	WinSetState($S1, "", @SW_MAXIMIZE)
 #CE
EndFunc   ;==>check_mail



Func check_mail1()
#CS 	HotKeySet("{ESC}", "esc")
; 	WinSetState($S1, "", @SW_MINIMIZE)
; 	AdlibRegister("wait_msg", $p_time)
; 	Global $mail = ""
;
; 	;AdlibRegister("stop",100)
; 	GUICtrlSetData($Label1, $SD)
; 	GUICtrlSetColor($Label1, 0x000000)
; 	GUICtrlRead($Label1)
; 	GUICtrlSetData($Button25, $SD)
; 	GUICtrlSetColor($Button25, 0x000000)
; 	GUICtrlRead($Button25)
;
; 	GUICtrlSetData($Button27, "Key:" & $prevTkt)
; 	GUICtrlSetColor($Button27, 0x000000)
; 	GUICtrlRead($Button27)
;
; 	GUICtrlSetData($Label4, $fname)
; 	GUICtrlSetColor($Label4, 0x000000)
; 	GUICtrlRead($Label4)
; 	GUICtrlSetData($Button26, $fname)
; 	GUICtrlSetColor($Button26, 0x000000)
; 	GUICtrlRead($Button26)
;
;
;
; 	Local $oIE = _IECreate("https://www.myworkday.com/company/" , 0, 1, 1, 1)
;
; 	_IELoadWait($oIE)
;
; 	check_mouse()
; 	Opt("WinTitleMatchMode", 2)
; 	WinWait("Internet Explorer", "")
; 	sleep(700)
;
;
; 	If WinExists("OpenAM (Login) -", "") Then
; 		Do
;
; 		Local $ouser = _IEGetObjById($oIE, "IDToken1")
; 		_IEFormElementSetValue($ouser, $username&"@"&$domain&".company.com" )
; 		Local $opas = _IEGetObjById($oIE, "IDToken2")
; 		_IEFormElementSetValue($opas, $pass)
; 		WinActivate($oIE, "")
; 		Local $oForm = _IEFormGetCollection($oIE, 0)
; 		_IEFormSubmit($oForm, 0)
; 		_IELoadWait($oIE)
;
; 		Until WinActive("Workday -","")
; 		_IELoadWait($oIE)
;
; 		Do
; 			_IELoadWait($oIE)
; 			sleep(1000)
; 		Until WinActive("Home -","")
;
; 	EndIf
;
; _IELoadWait($oIE)
;
; check_mouse()
; sleep(1700)
; if WinExists("Message from webpage","") Then
; 		$a=WinWait("Message from webpage","")
; 		WinActivate($a)
; 		send("{enter}")
; 		_IEQuit($oIE)
; 		check_mail2()
; 			Else
;
;
;
;
; 	Do
; 		$i = 0
; 		$st = _IEPropertyGet($oIE, "outertext")
; 		Sleep($d_time)
; 		$aD = StringSplit($st, "ID", $STR_ENTIRESPLIT)
;
; 		While $aD[0] = 1 and $i < 20
;
; 			Sleep(5)
; 			$st = _IEPropertyGet($oIE, "outertext")
; 			$aD = StringSplit($st, "ID", $STR_ENTIRESPLIT)
; 			Sleep($d_time * 2)
; 			$i = $i + 1
;
; 		WEnd
;
; 				    $st = _IEPropertyGet($oIE, "outertext")
; 					$check = StringInStr($st,"Please Wait..")
;
; 					While ($check <> 0 and $check <> 1)
; 							sleep(5)
; 							$st = _IEPropertyGet($oIE, "outertext")
; 							$check = StringInStr($st,"Please Wait..")
; 					WEnd
;
; 		check_mouse()
; 		Local $stext = _IEBodyReadText($oIE)
; 		$dd = "Department Code:"
;
; 		$aDays = StringSplit($stext, $dd, $STR_ENTIRESPLIT)
;
; 		If $i > 15 Then ExitLoop ; to exit it QC not working
;
; 	Until @error = 0
; _IELoadWait($oIE)
; 	If @error = 0 Then
; 	$st = _IEPropertyGet($oIE, "outertext")
; 	$check = StringInStr($st,"ID:")
;
; if ($check <> 0 and $check <> 1) then
;
; 		$sS = $aDays[2]
;
; 		Local $code = StringMid($sS, 1, 6)
;
;
; 		If $code = 10003 Or $code = 10023 Or $code = 10037 Or $code = 10032 Then
;
; 			Global $agent = "p"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
; 			Sleep($d_time * 4)
;
; 		GUICtrlSetBkColor($Label1, 0x000000)
; 		GUICtrlSetColor($Label1, 0xFFFF00)
; 		GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Label4, 0xFFFF00)
; 		GUICtrlRead($Label4)
;
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFFFF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFFFF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFFFF00)
; 		GUICtrlSetBkColor($Button21, 0xFFFF00)
; 		GUICtrlSetBkColor($Button22, 0xFFFF00)
; 		GUICtrlSetBkColor($Button23, 0xFFFF00)
; 		GUICtrlSetBkColor($Button24, 0xFFFF00)
; 		GUICtrlSetBkColor($Button27, 0xFFFF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
;
; 		ElseIf $code = 10009 Or $code = 10010 Or $code = 10022 Or $code = 10024 Then
; 			Global $agent = "t"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 			Sleep($d_time * 4)
; 			GUICtrlSetBkColor($Label4, 0x00FF00)
; 			GUICtrlRead($Label4)
; 			GUICtrlSetBkColor($Label1, 0x000000)
; 			GUICtrlSetColor($Label1, 0x00FF00)
; 			GUICtrlRead($Label1)
;
;
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0x00FF00)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0x00FF00)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0x00FF00)
; 		GUICtrlSetBkColor($Button21, 0x00FF00)
; 		GUICtrlSetBkColor($Button22, 0x00FF00)
; 		GUICtrlSetBkColor($Button23, 0x00FF00)
; 		GUICtrlSetBkColor($Button24, 0x00FF00)
; 		GUICtrlSetBkColor($Button27, 0x00FF00)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
;
;
;
; 		Else
; 			Global $agent = "p"
; 			ControlCommand($Form1, "", $Combo2, "ShowDropDown")
; 			ControlCommand($Form1, "", $Combo2, "HideDropDown")
;
; 			Sleep($d_time * 4)
;
;
; 			GUICtrlSetBkColor($Label4, 0xFF8888)
; 			GUICtrlRead($Label4)
; 			GUICtrlSetBkColor($Label1, 0x000000)
; 			GUICtrlSetColor($Label1, 0xFF8888)
; 			GUICtrlRead($Label1)
;
; 		GUICtrlSetBkColor($Button25, 0x000000)
; 		GUICtrlSetColor($Button25, 0xFF8888)
; 		GUICtrlRead($Button25)
;
; 		GUICtrlSetBkColor($Button26, 0x000000)
; 		GUICtrlSetColor($Button26, 0xFF8888)
; 		GUICtrlRead($Button26)
;
;
; 		GUICtrlSetBkColor($Button20, 0xFF8888)
; 		GUICtrlSetBkColor($Button21, 0xFF8888)
; 		GUICtrlSetBkColor($Button22, 0xFF8888)
; 		GUICtrlSetBkColor($Button23, 0xFF8888)
; 		GUICtrlSetBkColor($Button24, 0xFF8888)
; 		GUICtrlSetBkColor($Button27, 0xFF8888)
; 		GUICtrlRead($Button20)
; 		GUICtrlRead($Button21)
; 		GUICtrlRead($Button22)
; 		GUICtrlRead($Button23)
; 		GUICtrlRead($Button24)
; 		GUICtrlRead($Button27)
; 			If $code <> 10000 Then
;
; 				$verify = MsgBox(4, "On Action:", "Non Agent Profile do you want to continue", "")
;
; 				If $verify = 7 Then
; 					_IEQuit($oIE)
; 					Exit
; 				elseif $verify = 6 Then
; 					$shift = "No"
; 				EndIf
; 			ElseIf $code = 10000 Then
;
; 				$verify = MsgBox(1, "Login Issue detected:", "Agents Dept Code is not correct manager/OM need to work ", "")
;
; 				Global $agent = "p"
; 		$shift = "No"
; 			EndIf
;
;
; 		EndIf
; 		Sleep($d_time * 3)
; 		check_mouse()
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		Local $iNumLinks = @extended
; 		Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 		For $oLink In $oLinks
;
; 			$sTxt &= $oLink.href & @CRLF
; 		Next
; 		_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 		_IELoadWait($oIE)
; 		check_mouse()
;
; 				$st = _IEPropertyGet($oIE, "outertext")
; 				$check = StringInStr($st,"Please Wait..")
; 				While ($check <> 0 and $check <> 1)
; 						sleep(5)
; 						$st = _IEPropertyGet($oIE, "outertext")
; 						$check = StringInStr($st,"Please Wait..")
; 				WEnd
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		For $oLink In $oLinks
; 			If StringInStr($oLink.href, "mailto") Then
; 				_IEAction($oLink, "focus")
; 				Global $smail = _IEPropertyGet($oLink, "innertext")
; 				; MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$smail)
; 				ExitLoop
; 			EndIf
; 		Next
;
;
;
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		$class_value = $oLinks.GetAttribute("class")
; 		If $class_value = "detailInfo" Then
;
; 			Local $iNumLinks = @extended
; 			Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
; 			For $oLink In $oLinks
;
; 				$sTxt &= $oLink.href & @CRLF
; 			Next
; 		EndIf
; 		_IELinkClickByIndex($oIE, $iNumLinks - 1, 0)
; 		_IELoadWait($oIE)
; 		check_mouse()
;
; 					$st = _IEPropertyGet($oIE, "outertext")
; 					$check = StringInStr($st,"Please Wait..")
; 					While ($check <> 0 and $check <> 1)
; 							sleep(5)
; 							$st = _IEPropertyGet($oIE, "outertext")
; 							$check = StringInStr($st,"Please Wait..")
; 					WEnd
;
;
; 		Local $oLinks = _IELinkGetCollection($oIE)
; 		For $oLink In $oLinks
; 			If StringInStr($oLink.href, "mailto") Then
; 				_IEAction($oLink, "focus")
; 				Global $mmail = _IEPropertyGet($oLink, "innertext")
; 				;MsgBox($MB_SYSTEMMODAL, "","TL's Email "&$mmail)
; 				ExitLoop
; 			EndIf
; 		Next
;
;
;
;
; 		_IEQuit($oIE)
; 		;AdlibUnRegister("stop")
;
;
;
; Else
;
; 		_IEQuit($oIE)
; 		check_mail2()
; 	EndIf
; 	EndIf
; 	EndIf
; 	AdlibUnRegister("wait_msg")
; 	WinSetState($S1, "", @SW_MAXIMIZE)
 #CE
EndFunc   ;==>check_mail

Func check_sms()
;~ 	HotKeySet("{ESC}", "esc") ;AdlibRegister("stop",100)
;~ 	Global $S1 = WinWait("company Service Management Suite - New Interaction - HP Service Manager Client", "", 1)
;~ 	WinActivate($S1)
;~ 	Global $S13 = WinWait("company Service Management Suite - Update Incident Number - HP Service Manager Client", "", 1)
;~ 	WinActivate($S13)

;~ 	Local $smsw = WinGetTitle("[ACTIVE]")
;~ 	$a = StringSplit($smsw, "-")
;~ 	If $a[1] = "company Service Management Suite " Then
;~ 		Global $sms = 1
;~ 	Else
;~ 		Global $sms = 2
;~ 	EndIf
;~ 	;MsgBox(0, "", "SMS version :" & $sms , .5)
;~ 	;AdlibUnRegister("stop")
EndFunc   ;==>check_sms

Func start_setup()


			$usern = "ASHISH KUMAR SAKLANI"
			$domain = "company"
			$user = "ASHISH "

#CS
; 	HotKeySet("{ESC}", "esc")
;
;
; 	Switch $username
;
;
; 		Case "Myid"
; 			$usern = "ASHISH KUMAR SAKLANI"
; 			$user = "ASHISH "
;
; 		
;
; 		Case Else
; 			$usern = ""
;
;
; 			$step = ""
;
; 	EndSwitch
; 	;AdlibUnRegister("stop")
 #CE

EndFunc   ;==>start_setup

Func check_date()
	HotKeySet("{ESC}", "esc")

	Global $day = @MDAY
	$min = @MIN - 30
	$Hour = @HOUR - 10

	;MsgBox("","",$day)

	If $Hour < 0 Then $day = $day - 1
	If $Hour < 0 Then $Hour = 24 + $Hour
	If $min < 0 Then $Hour = $Hour + 1
	If $min < 0 Then $min = 60 - $min



	$da = "0" & $day

	Global $day = StringRight($da, 2)

	Global $date = @MON & "/" & $day & "/" & @YEAR

	;MsgBox("","",$day)
EndFunc   ;==>check_date

Func check_issue()
	HotKeySet("{ESC}", "esc")
	;===================================================================================
	; Local $sAnswer = InputBox("Question", "Where were you born?", "Planet Earth", "", _
	; 300,150, 300, 0)
	;===================================================================================



	If GUICtrlRead($Combo1) == "VPN" Then

	Global $cat1_tier1 = "VPN"
	Global $cat1_tier2 = "Cisco AnyConnect"
	Global $cat1_tier3 = "Unable to Connect"
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "User Productivity App"





				If GUICtrlRead($Combo2) =="Cisco" Then

					Global $value4 = "Cisco VPN"
					Global $logs = "Unable to login to Cisco any connect"
					Global $logs1 = "working fine after reset"
					Global $cat1_tier2 = "Cisco AnyConnect"


				ElseIf GUICtrlRead($Combo2) =="Global" Then

					Global $value4 = "Palo Alto Virtual Firewalls"
					Global $logs = "Global protect not working"
					Global $logs1 = "working fine after relaunch"
					Global $cat1_tier2 = "Other/ Client VPN"

				ElseIf GUICtrlRead($Combo2) =="Forti" Then
					Global $value4 = "Fortintet VPN"
					Global $logs = " unable to login to Forticlient VPN"
					Global $logs1 = "password reset"
					Global $cat1_tier2 = "FortiClient"
				EndIf

					Global $cat2_tier2 = $cat1_tier2

	ElseIf GUICtrlRead($Combo1) == "VDI" Then





	Global $cat1_tier1 = "Cloud Services"
	Global $cat1_tier2 = "Azure Cloud Platform"
	Global $cat1_tier3 = ""
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Operating System (Cloud Instance)"
	Global $cat1_tier6 = "Standard OS"




				If GUICtrlRead($Combo2) =="Session" Then

					Global $value4 = "Azure VDI"
					Global $logs = "RD /VDI not working"
					Global $logs1 = "working fine after RD reset"


				ElseIf GUICtrlRead($Combo2) =="Update" Then

					Global $value4 = "Azure VDI"
					Global $logs = "Update Neeed for Remote Desktop"
					Global $logs1 = "RD updated"

				ElseIf GUICtrlRead($Combo2) =="Display" Then

					Global $value4 = "Azure VDI"
					Global $logs = "VDI display settings correction"
					Global $logs1 = "corrected display settings on VDI"

				ElseIf GUICtrlRead($Combo2) =="Blank" Then

					Global $value4 = "Azure VDI"
					Global $logs = "Blank screen on VDI session"
					Global $logs1 = "Reset VDI session from console"

				ElseIf GUICtrlRead($Combo2) =="SRW" Then

					Global $value4 = "Remote Desktop Connection"
					Global $logs = "Secure Remote Worker Application not working"
					Global $logs1 = "reinstalled SRW"

				ElseIf GUICtrlRead($Combo2) =="Freezing" Then

					Global $value4 = "Data Access - WAH (User Provided)"
					Global $logs = "VDI running slow"
					Global $logs1 = "Reset VDI session"

				ElseIf GUICtrlRead($Combo2) =="Disconnecting" Then

					Global $value4 = "Data Access - WAH (User Provided)"
					Global $logs = "VDI disconnecting frequently"
					Global $logs1 = "Power cycle done, ran Network Test and MTR"

				ElseIf GUICtrlRead($Combo2) =="No Avaya" Then

					Global $value4 = "Data Access - WAH (User Provided)"
					Global $logs = "No Avaya in VDI desktop"
					Global $logs1 = "Session Reset Changed Host"

				ElseIf GUICtrlRead($Combo2) =="Bliss" Then

					Global $value4 = "Data Access - WAH (User Provided)"
					Global $logs = "Bliss Ext not working in VDI"
					Global $logs1 = "Session Reset changed VDI host"


				EndIf


	ElseIf GUICtrlRead($Combo1) == "Password" Then


	Global $cat1_tier1 = "Password & User ID Management"
	Global $cat1_tier2 = "company SSO ID /Login ID"
	Global $cat1_tier3 =  "Password Reset"
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Service"
	Global $cat1_tier5 = "Business Service"
	Global $cat1_tier6 = "Application"



		If GUICtrlRead($Combo2) =="Portal" Then

			Global $value4 = "Workday"
			Global $logs = "Unable to Login to Workday Portal "
			Global $logs1 = "Helped user to login"
			Global $cat1_tier2 = "company SSO ID /Login ID"


		ElseIf GUICtrlRead($Combo2) =="SSO"  Then

			Global $value4 = "SSO (Single Sign On) - Password Reset"
			Global $logs = "Unable to Login SSO Password reset Required "
			Global $logs1 = "helped user to reset password"
			Global $cat1_tier2 = "company SSO ID /Login ID"

		ElseIf GUICtrlRead($Combo2) =="LAN" Then

			Global $value4 = "Active Directory Lan ID - Password Reset"
			Global $logs = "Unable to Login LAN Password reset Required "
			Global $logs1 = "Helped user to reset password"
			Global $cat1_tier2 = "Lcompany LAN ID (Ncompany)"


		ElseIf GUICtrlRead($Combo2) =="Estart" Then


			Global $value4 = "eStart - Password Reset"
			Global $logs = "eStart Password not working"
			Global $logs1 = "asked user to contact hd"
			Global $cat1_tier2 = "eStart"

		ElseIf GUICtrlRead($Combo2) =="Nice" Then

			Global $value4 = "NICE - Lcompany"
			Global $logs = "Unable to Login to NICE Password reset Required "
			Global $logs1 = "asked user to contact HD"
			Global $cat1_tier2 = "NICE/IEX Agent"

		ElseIf GUICtrlRead($Combo2) =="Client" Then

			Global $value4 = "Active Directory Lan ID - Password Reset"
			Global $logs = "Unable to Login with Client ID and Password "
			Global $logs1 = "asked user to contact Client HD"
			Global $cat1_tier2 = "Any Other Domain Login ID"

		EndIf

		Global $cat2_tier2 = $cat1_tier2

	ElseIf GUICtrlRead($Combo1) == "MultiFactor" Then



	Global $cat1_tier1 = "DUO MFA"
	Global $cat1_tier2 = "DUO Authentication"
	Global $cat1_tier3 = "Send Activation Code"
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = "Problem was resolved by providing instructions to the user"
	Global $cat1_tier4 = "Service"
	Global $cat1_tier5 = "Business Service"
	Global $cat1_tier6 = "Application"




				If GUICtrlRead($Combo2) =="DUO act" Then

					Global $value4 = "DUO AuthSync - Lcompany"
					Global $logs = "Duo activation required"
					Global $logs1 = "Duo account activated for user"

				ElseIf GUICtrlRead($Combo2) =="DUO num" Then

					Global $value4 = "DUO AuthSync - Lcompany"
					Global $logs = "DUo number changed new number activated"
					Global $logs1 = "Helped user to update number for duo"
					Global $cat1_tier3 = "Update Phone Number"


				ElseIf GUICtrlRead($Combo2) =="Secure CX" Then

					Global $value4 = "SecureLogin ActivIdentity"
					Global $logs = "Secure CX / Camera not working"
					Global $logs1 = "Helped user with Secure CX working now"

					Global $cat1_tier1 = "SecureCX" = $cat2_tier1
					Global $cat1_tier2 = "SecureCX Application Issues" = $cat2_tier2
					Global $cat1_tier3 = ""	= $cat2_tier3

				ElseIf GUICtrlRead($Combo2) =="BIOnix" Then
					Global $value4 = "Bionix"
					Global $logs = "Bionix byass Required"
					Global $logs1 = "helped user with Bionix"

				EndIf

	ElseIf GUICtrlRead($Combo1) == "Sound" Then

	Global $cat1_tier1 = "Voice & Telephony"
	Global $cat1_tier2 = "Corporate Phone - Avaya Not Working"
	Global $cat1_tier3 = "Corporate Phone - Avaya Voice Quality Issue(Call Drop/Noise)"
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "Virtual Machine Software"



				If GUICtrlRead($Combo2) =="No sound" Then

					Global $value4 = "Headset"
					Global $logs = "NO sound on  calls"
					Global $logs1 = "corrected sound settings"

				ElseIf GUICtrlRead($Combo2) =="sound quality" Then

					Global $value4 = "Data Access - WAH (User Provided)"
					Global $logs = "Poor Sound quality on calls"
					Global $logs1 = "Network test run referred user to ISP"
				EndIf

	ElseIf GUICtrlRead($Combo1) == "Avaya" Then


	Global $cat1_tier1 = "Voice & Telephony"
	Global $cat1_tier2 = "Corporate Phone - Avaya Not Working"
	Global $cat1_tier3 = "Corporate Phone - Avaya Voice Quality Issue(Call Drop/Noise)"
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = $cat1_tier2
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "Virtual Machine Software"

				If GUICtrlRead($Combo2) =="login failed" Then
					Global $value4 = "Avaya one-X"
					Global $logs = " Unable to login to avaya phone"
					Global $logs1 = "reset avaya"

				ElseIf GUICtrlRead($Combo2) =="Desktop Mode" Then

					Global $value4 = "Avaya one-X"
					Global $logs = "Unable to login to avaya phone"
					Global $logs1 = "reset avaya working fine now"


				ElseIf GUICtrlRead($Combo2) =="Hardphone" Then

					Global $value4 = "HA Hard Phone"
					Global $logs = "Issue:issue with avaya hardphone"
					Global $logs1 = "reconfigured avaya hardphone"

				EndIf



	ElseIf GUICtrlRead($Combo1) == "Windows" Then


	Global $cat1_tier1 = "Hardware/ Desktop/ Laptop"
	Global $cat1_tier2 = "Desktop / Laptop"
	Global $cat1_tier3 = ""
	Global $cat2_tier1 = $cat1_tier1
	Global $cat2_tier2 = "Desktop/ Laptop"
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "Virtual Machine Software"





		If GUICtrlRead($Combo2) =="Resolution" Then

			Global $value4 = "Monitor"
			Global $logs = "Issue with Display  settings in Windows"
			Global $logs1 = "corrected display settings"
			Global $cat1_tier2 = "Monitor FT (Display)"


		ElseIf GUICtrlRead($Combo2) =="Login" Then

			Global $value4 = "Desktop Operating System"
			Global $logs = "unable to login to computer"
			Global $logs1 = "Helped user to login to computer"


		ElseIf GUICtrlRead($Combo2) =="Settings" Then

			Global $value4 = "Desktop Operating System"
			Global $logs = "settings in windows "
			Global $logs1 = "corrected settings in windows working now"


		ElseIf GUICtrlRead($Combo2) =="Disk Space" Then

			Global $value4 = "Desktop Operating System"
			Global $logs = "Low Disk Space in windows"
			Global $logs1 = "cleared disk space in windows working now"

		ElseIf GUICtrlRead($Combo2) =="Update" Then

			Global $value4 = "Windows Server Update Services (WSUS)"
			Global $logs = "Issue with  windows update "
			Global $logs1 = "Helped user with windows update"

		ElseIf GUICtrlRead($Combo2) =="Boot" Then

			Global $value4 = "Desktop Operating System"
			Global $logs = "no boot in windows"
			Global $logs1 = "helped user to login to windows"


		ElseIf GUICtrlRead($Combo2) =="Network Test" Then

			Global $value4 = "Data Access - WAH (User Provided)"
			Global $logs = "User need help in Network test"
			Global $logs1 = "Network test done"

		ElseIf GUICtrlRead($Combo2) =="No Display" Then

			Global $value4 = "Desktop Operating System"
			Global $logs = "No display after booting Windows  "
			Global $logs1 = "helped user with display settings"


		EndIf


	ElseIf GUICtrlRead($Combo1) == "Application" Then

	Global $cat1_tier1 = "Application / Software"
	Global $cat1_tier2 = "Client Application"
	Global $cat1_tier3
	Global $cat2_tier1 = "Client Applications"
	Global $cat2_tier2 = "Client Applications"
	Global $cat2_tier3 = $cat1_tier3
	Global $cat1_tier4 = "Software"
	Global $cat1_tier5 = "Software Application/System"
	Global $cat1_tier6 = "User Productivity App"


		If GUICtrlRead($Combo2) =="Client application" Then

			Global $value4 = "Client Application Service"
			Global $logs = "Client Application not working in VDI"
			Global $logs1 = "working fine on relaunch"

		ElseIf GUICtrlRead($Combo2) =="Estart" Then

			Global $cat1_tier2 = "eSTART Timekeeping"
			Global $cat2_tier2 = $cat1_tier2
			Global $value4 = "eStart - Password Reset"
			Global $logs = "eStart not working"
			Global $logs1 = "asked user to contact hd"


		ElseIf GUICtrlRead($Combo2) =="Secure CX" Then

					Global $cat1_tier1 = "SecureCX" = $cat2_tier1
					Global $cat1_tier2 = "SecureCX Application Issues" = $cat2_tier2
					Global $cat1_tier3 = ""	= $cat2_tier3

			Global $value4 = "SecureLogin ActivIdentity"
			Global $logs = "Secure CX / Camera not working"
			Global $logs1 = "Helped user with Secure CX working now"


		ElseIf GUICtrlRead($Combo2) =="Citrix" Then

			Global $cat1_tier2 = "Citrix Server"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "Citrix Admin"
			Global $logs = "Citrix app not working"
			Global $logs1 = "helped user with login"




		ElseIf GUICtrlRead($Combo2) =="Workday" Then

			Global $cat1_tier2 = "Workday Application"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "Workday"
			Global $logs = "Unable to login to Workday"
			Global $logs1 = "helped user with workday login"

		ElseIf GUICtrlRead($Combo2) =="Sharepoint" Then


			Global $cat1_tier2 = "Sharepoint"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "SharePoint company"
			Global $logs = "Sharepoint not working"
			Global $logs1 = "helped user with Sharepoint"

		ElseIf GUICtrlRead($Combo2) =="company University" Then


			Global $cat1_tier2 = "Non-ITHelpdesk Application"
			Global $cat2_tier2 = "company Workspace"
			Global $value4 = "company University"
			Global $logs = "University website not working"
			Global $logs1 = "helped user with companyU"

		ElseIf GUICtrlRead($Combo2) =="Quickconnect" Then

			Global $cat1_tier2 = "QuickConnect" = $cat2_tier2
			Global $value4 = "QuickConnect Chat Broadcast Mail"
			Global $logs = "QuickConnect Not working"
			Global $logs1 = "helped user with QC"

		ElseIf GUICtrlRead($Combo2) =="Textexpander" Then

			Global $value4 = "Client Application Service"
			Global $logs = "Textexpander not working in VDI"
			Global $logs1 = "Reset VDI session"

		ElseIf GUICtrlRead($Combo2) =="Okta" Then

			Global $value4 = "Client Application Service"
			Global $logs = "Client application not working in OKTA"
			Global $logs1 = "working fine on relaunch"

		ElseIf GUICtrlRead($Combo2) =="SSCM" Then


			Global $cat1_tier2 = "Non-ITHelpdesk Application"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "SCCM_Application - Lcompany"
			Global $logs = "SSCM application not working on company PC"
			Global $logs1 = "Guided user with SSCM"

		ElseIf GUICtrlRead($Combo2) =="Zimbra" Then


			Global $cat1_tier2 = "Zimbra"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "Zimbra Email"
			Global $logs = "Zimbra Email not working"
			Global $logs1 = "working fine on relaunch"

		ElseIf GUICtrlRead($Combo2) =="Zoom" Then

			Global $cat1_tier2 = "Non-ITHelpdesk Application"
			Global $cat2_tier2 = "company Workspace"

			Global $value4 = "Zoom"
			Global $logs = "Zoom not working"
			Global $logs1 = "working fine on relaunch"

		ElseIf GUICtrlRead($Combo2) =="Adobeconnect" Then

			Global $cat1_tier2 = "Non-ITHelpdesk Application"
			Global $cat2_tier2 = "company Workspace"


			Global $value4 = "Adobeconnect"
			Global $logs = " user finding issue with Adobeconnect"
			Global $logs1 = "Helped user with Adobe connect"

		ElseIf GUICtrlRead($Combo2) =="Office" Then

			Global $cat1_tier2 = "Office 365" = $cat2_tier2


			Global $value4 = "Microsoft Office"
			Global $logs = " user finding issue with Office application"
			Global $logs1 = "Helped user with OFFICE"


		EndIf


	ElseIf GUICtrlRead($Combo1) == "Omni Chat" Then


			Global $cat1_tier1 = "Query"
			Global $cat1_tier2 = "Request for help/Education/Advise"
			Global $cat1_tier3 = "Query related to Non IT Request / Issues"

			Global $cat2_tier1 = "Query"
			Global $cat2_tier2 = "Request for help/Education/Advise"

		If GUICtrlRead($Combo2) == "Disconnected" Then

			Global $value4 = "Client Application Service"
			Global $logs = "User disconnected chat without details"
			Global $logs1 = "User disconnected chat without details"


		ElseIf GUICtrlRead($Combo2) == "Misroute" Then

			Global $value4 = "Client Application Service"
			Global $logs = "User joined wrong support queue "
			Global $logs1 = "Misroute"


		ElseIf GUICtrlRead($Combo2) == "Test" Then

			Global $value4 = "Case (a.k.a Case Tracker) - Test Environment"
			Global $logs = "test user joined Omni chat for test purpose"
			Global $logs1 = "testing chat done"

		EndIf

	ElseIf GUICtrlRead($Combo1) == "Install" Then


		If GUICtrlRead($Combo2) =="Secure Cx" Then

			Global $value4 = "SecureLogin ActivIdentity"
			Global $logs = "Install Secure CX client"
			Global $logs1 = "INstalled Secure CX"

		ElseIf GUICtrlRead($Combo2) =="Remote Desktop" Then

			Global $value4 = "Azure VDI"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed RD	"

		ElseIf GUICtrlRead($Combo2) =="Global Protect" Then

			Global $value4 = "Palo Alto Virtual Firewalls"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)

		ElseIf GUICtrlRead($Combo2) =="Ms Office" Then

			Global $cat1_tier2 = "Office 365" = $cat2_tier2
			Global $value4 = "Microsoft Office"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)
		ElseIf GUICtrlRead($Combo2) =="Cisco" Then

			Global $cat1_tier1 = "VPN"
			Global $cat1_tier2 = "Cisco AnyConnect"
			Global $cat2_tier1 = "VPN"
			Global $cat2_tier2 = "Cisco AnyConnect"

			Global $value4 = "Cisco VPN"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)


		ElseIf GUICtrlRead($Combo2) =="Quickconnect" Then

			Global $cat1_tier2 = "QuickConnect" = $cat2_tier2
			Global $value4 = "QuickConnect Chat Broadcast Mail"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)

		ElseIf GUICtrlRead($Combo2) =="Zoom" Then

			Global $value4 = "Client Application Service"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)
		ElseIf GUICtrlRead($Combo2) =="Citrix" Then

			Global $value4 = "Client Application Service"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)


		Else

			Global $value4 = "Client Application Service"
			Global $logs = "Install " & GUICtrlRead($Combo2)
			Global $logs1 = "Installed " & GUICtrlRead($Combo2)

		EndIf
	ElseIf GUICtrlRead($Combo1) == "Replacement" Then


		If GUICtrlRead($Combo2) =="CPU" Then

			Global $value4 = "Desktop/Laptop Hardware and Peripherals"
			Global $logs = "CPU replacement needed for computer"
			Global $logs1 = "CPU replacement requested"

		ElseIf GUICtrlRead($Combo2) =="Headset" Then

			Global $value4 = "Headset"
			Global $logs = "Headset replacement needed"
			Global $logs1 = "replacement requested"

		ElseIf GUICtrlRead($Combo2) =="Monitor" Then

			Global $value4 = "Desktop/Laptop Hardware and Peripherals"
			Global $logs = "Monitor replacement needed"
			Global $logs1 = "replacement requested"

		ElseIf GUICtrlRead($Combo2) =="Cables" Then

			Global $value4 = "Desktop/Laptop Hardware and Peripherals"
			Global $logs = "Cable replacement needed"
			Global $logs1 = "replacement requested"
		EndIf

	EndIf

EndFunc   ;==>check_issue

Func check_id()

$sid  = StringRegExpReplace($sid,"[^0-9,\h,\v]","")

$len=  StringLen ($sid)

EndFunc

Func check_record()
	Opt("SendKeyDelay", 2)
	Opt("WinTitleMatchMode", 2)

	Global $S1 = WinWait("- New Interaction - HP Service Manager Client", "")
	WinActivate($S1)
	ControlClick($S1, "", "ToolbarWindow324")
	Send("S")
	check_mouse()
	Sleep($d_time * 2)
	$window = ControlCommand($S1, "Search Interaction Records", "SysTreeView321", "IsVisible")

	Sleep($d_time)
	If $window = 1 Then
		Send("Conn {RIGHT} ")
		check_mouse()
		Sleep($d_time * 2)
		Send("Menu {RIGHT} ")
		check_mouse()
		Sleep($d_time * 2)
		Send("Servi{RIGHT} ")
		check_mouse()
		Sleep($d_time * 2)
		Send("Search")
		check_mouse()
		Sleep($d_time * 2)
		Send("{ENTER}")
	EndIf




	$S15 = WinWait("- Display Which Service Desk Interactions? - ", "")
	WinActivate($S15)
	;MsgBox("","",$sid)
	check_mouse()
	Sleep($d_time * 5)

	ControlSetText($S15, "", "Edit26", $sid)

	If @MDAY > 3 Then

		ControlSetText($S15, "", "Edit3", @MON & "/" & @MDAY - 3 & "/" & @YEAR & " 00:00:00")

	ElseIf @MDAY <= 3 Then

		ControlSetText($S15, "", "Edit3", @MON - 1 & "/28" & "/" & @YEAR & " 00:00:00")

	EndIf

	ControlCommand($S15, "", "ToolbarWindow3211", "SendCommandID", 2)
	WinActivate($S15)
	Send("{F6}")


	check_mouse()


	If WinExists("- Interaction:", "") Then
		$S16 = WinWait("- Interaction:", "", 10)

		$item = ControlListView($S16, "", "SysListView321", "GetItemCount")
		$subitem = ControlListView($S16, "", "SysListView321", "GetSubItemCount")
		ControlListView($S16, "", "SysListView321", "Select", $item - 1)

		$data = ControlListView($S16, "", "SysListView321", "GetText", 0, 0)
			#cs
				;MsgBox("","",$data)
				If (StringLeft($data, 2) = "SD") Then
					ControlSend($S16, "", "SysListView321", "{Enter}")


							$text = ""
							$data = ""


							For $j = 0 To $item Step 1




								For $i = 0 To $subitem Step 1

									If $i = 2 Then
										$data &= ""
									ElseIf $i = 3 Then
										$data &= ""
									ElseIf $i = 4 Then
										$data &= ""
									ElseIf $i = 7 Then
										$data &= ""
									ElseIf $i = 8 Then
										$data &= ""
									Else

										$data &= ControlListView($S16, "", "SysListView321", "GetText", $j, $i) & " "

										If StringInStr($data, "IM", 1) = 1 Then
											ExitLoop (2)
										EndIf

									EndIf
								Next

								$data &= @CRLF

							Next


							$text = $data & @CRLF

							$text &= @CRLF


							$data = "     SD Number		Project		StartDate		EndDate				Subject			 " & @LF
							$data &= "																" & @LF


							If $text <> "" Then

								_ExtMsgBoxSet(1, 2, 0x004080, 0xFFFF00, 10, "Comic Sans MS")
								_ExtMsgBoxSet(-1, 1, -1, -1, -1, -1, @DesktopWidth - 70)
								$n = 0
								$n = _ExtMsgBox("", "CONFIRM", "Agents Record of past 48 hrs", $text)

								Do
									Sleep(20)
								Until $n = 1


							EndIf
							Sleep(2000)

						ControlCommand($S16, "", "ToolbarWindow321", "SendCommandID", 4)

		Else
			$S16 = WinWait("- Interaction:", "", 10)
			Local $sid2 = ControlGetText($S16, "", "Edit1") ;12
			Local $issue = ControlGetText($S16, "", "Edit41") ;12
			_ExtMsgBoxSet(1, 2, 0x004080, 0xFFFF00, 10, "Comic Sans MS")
			_ExtMsgBoxSet(-1, 1, -1, -1, -1, -1, @DesktopWidth - 70)
			_ExtMsgBox("", "CONFIRM", "Agents Record of past 48 hrs", $sid2 & "   " & $issue)
			Sleep($d_time * 5)
			WinActivate($S16)
			Send("{F3}")
			;ControlSend($S16, "", "ToolbarWindow3211", "{F3}")
			;================issue line==================================================================================issue line====================================================
			$S15 = WinWait("- Display Which Service Desk Interactions? - ", "")
			WinActivate($S15)
			Send("{F3}")
			;ControlSend($S15, "", "", "{F3}")

		EndIf


		#ce

	Else
		WinActivate($S15)
		Send("{F3}")
	EndIf



EndFunc   ;==>check_record


Func esc()
	;$step = "stop"

	if WinExists("     SCRIPT PAUSED","") then
		WinClose("     SCRIPT PAUSED","")
		SoundSetWaveVolume(100)
		send("{VOLUME_UP 100}")

		$chat = "ZOHO"
	Else
		_ExtMsgBoxSet(1, 2, 0xFF8888, 0x004080, 14, "Comic Sans MS")
		_ExtMsgBoxSet(-1, 1, -1, -1, -1, -1, @DesktopWidth - 70)
		$n = 0
		$n = _ExtMsgBox("", "RESET|Continue|Quit", "     SCRIPT PAUSED", "For QC chat sound Press Esc again", "")

		If $n = 1 Then

			RunWait("ast.exe", "C:\")
			If @error <> 0 Then reset()
		ElseIf $n = 3 Then

			If ProcessExists('chromedriver.exe') Then
				ProcessClose('chromedriver.exe')
			EndIf
			delete_temp_files()

			Exit
		EndIf
	EndIf
	forward()
	back()

EndFunc   ;==>esc

Func EmptyFolder($FolderToDelete)
	$Debug =1

    $AllFiles=_FileListToArray($FolderToDelete,"*",0)
    If $Debug Then ConsoleWrite("-->" & $FolderToDelete & @CRLF )
    If IsArray($AllFiles) Then
        If $Debug Then
          ;  _ArrayDisplay( $AllFiles,$FolderToDelete)
        EndIf
        For $i = 1 To $AllFiles[0]
            $crt = FileGetTime( $FolderToDelete & "\" & $AllFiles[$i], 1 )


			#cs
														If $crt[2] = @MDAY And $crt[0] = @YEAR And $crt[1] = @MON Then
															If $Debug Then
																ConsoleWrite( $FolderToDelete & "\" & $AllFiles[$i] & " --> Today's File, Skipping!" & @CRLF )
															EndIf
															ContinueLoop
														EndIF

			#ce

			$delete = FileDelete($FolderToDelete & "\" & $AllFiles[$i])
            If $Debug Then
                ConsoleWrite($FolderToDelete & "\" & $AllFiles[$i]& " =>"&$delete & @CRLF  )
            EndIf
            DirRemove($FolderToDelete & "\" & $AllFiles[$i], 1)
        Next
    EndIf
EndFunc

Func delete_temp_files()
EmptyFolder (@WindowsDir & "\Temp")
EmptyFolder (EnvGet('userprofile')& "\Temp")
EmptyFolder (EnvGet('userprofile')& "\Appdata\Local\Temp")
;EmptyFolder (EnvGet('userprofile')& "\Local Settings\Temporary Internet Files\Content.IE5")
;EmptyFolder (EnvGet('userprofile')& "\Local Settings\Temporary Internet Files")
;EmptyFolder (EnvGet('userprofile')& "\Cookies")
;EmptyFolder (EnvGet('userprofile')& "\Local Settings\History")
;EmptyFolder (EnvGet('userprofile')& "\Temp\Temporary Internet Files")
;EmptyFolder (EnvGet('userprofile')& "\Recent")
;EmptyFolder (EnvGet('userprofile')& "\Application Data\Microsoft\Office\Recent")
;EmptyFolder (EnvGet('userprofile')& "\Local Settings\Temp")

;Run(@ComSpec & " /c del "&@TempDir&"\*.*  /s /f /q" , "", @SW_HIDE);
ShellExecuteWait("RunDll32.exe"," InetCpl.cpl,ClearMyTracksByProcess 255")
EndFunc

Func sms_login()

	HotKeySet("{ESC}", "esc")
	If $username <> "" And $password <> "" Then

		If WinExists("Server not available...") Then

			$l3 = WinWait("Server not available...", "", .5)

			WinActivate($l3)
			ControlClick($l3, "", "Button2", "")

			Sleep(3000)


			;
			;		$l3 = WinWait("SOAP Fault","")
			;		WinActivate($l3)
			;		ControlClick($l3, "", "Button1", "")

			;		sleep(3000)

			Opt("WinTitleMatchMode", 2)
			$l4 = WinWait("HP Service Manager Client", "")
			WinActivate($l4)
			ControlSend($l4, "", "SWT_Window024", "{ALTDOWN}fc1{ALTUP}")
			;ControlSend($l4, "", "SWT_Window026", "{ALTDOWN}fc1{ALTUP}")

			Sleep(3000)


			Opt("WinTitleMatchMode", 2)
			If WinExists("Auto Login") Then
				$l1 = WinWait("Auto Login", "")
				WinActivate($l1)
				ControlClick($l1, "", "Button1", "", 2)
				ControlClick($l1, "", "Button2", "", 2)

				Sleep(3000)

				Opt("WinTitleMatchMode", 2)
				$l2 = WinWait("Connections", "")
				WinActivate($l2)
				ControlClick($l2, "", "Edit3", "", 2)
				start_setup()


				Sleep(3000)

				ControlSend($l2, "", "Edit3", $user & "@domain")
				ControlClick($l2, "", "Edit4", "", 2)
				ControlSend($l2, "", "Edit4", $pass)
				ControlClick($l2, "", "Button13", "", 2)
				ControlClick($l2, "", "Button7", "", 2)

			EndIf
			#cs
				ElseIf WinExists("SOAP Fault") Then


				$l3 = WinWait("SOAP Fault","")
				WinActivate($l3)
				ControlClick($l3, "", "Button1", "")

				Opt("WinTitleMatchMode", 2)
				$l4 = WinWait("HP Service Manager Client", "")
				WinActivate($l4)
				ControlSend($l4, "", "SWT_Window024", "{ALTDOWN}fc1{ALTUP}")
				ControlSend($l4, "", "SWT_Window026", "{ALTDOWN}fc1{ALTUP}")

				Opt("WinTitleMatchMode", 2)
				$l1 = WinWait("Auto Login", "")
				WinActivate($l1)
				ControlClick($l1, "", "Button1", "")

				Opt("WinTitleMatchMode", 2)
				$l2 = WinWait("Connections", "")
				WinActivate($l2)
				ControlClick($l2, "", "Edit3", "", 2)
				start_setup()

				ControlSend($l2, "", "Edit3", $user & "@india.company.com")
				ControlClick($l2, "", "Edit4", "", 2)
				ControlSend($l2, "", "Edit4", $pass)
				ControlClick($l2, "", "Button13", "", 2)
				ControlClick($l2, "", "Button7", "", 2)

				ElseIf WinExists("Auto Login") Then

				Opt("WinTitleMatchMode", 2)
				$l1 = WinWait("Auto Login", "")
				WinActivate($l1)
				ControlClick($l1, "", "Button1", "")

				Opt("WinTitleMatchMode", 2)
				$l2 = WinWait("Connections", "")
				WinActivate($l2)
				ControlClick($l2, "", "Edit3", "", 2)
				start_setup()

				ControlSend($l2, "", "Edit3", $user & "@india.company.com")
				ControlClick($l2, "", "Edit4", "", 2)
				ControlSend($l2, "", "Edit4", $pass)
				ControlClick($l2, "", "Button13", "", 2)
				ControlClick($l2, "", "Button7", "", 2)


				ElseIf WinExists("Connections") Then

				Opt("WinTitleMatchMode", 2)
				$l2 = WinWait("Connections", "")
				WinActivate($l2)
				ControlClick($l2, "", "Edit3", "", 2)
				start_setup()

				ControlSend($l2, "", "Edit3", $user & "@india.company.com")
				ControlClick($l2, "", "Edit4", "", 2)
				ControlSend($l2, "", "Edit4", $pass)
				ControlClick($l2, "", "Button13", "", 2)
				ControlClick($l2, "", "Button7", "", 2)
			#ce

		EndIf
	EndIf
EndFunc   ;==>sms_login

Func bomgar()
	If WinExists("Bomgar - Idle Logout") Then
		$a = WinWait("Bomgar - Idle Logout", "", 2)
		WinActivate($a)
		Send("{ENTER}")


	ElseIf WinExists("Bomgar - New Session") Then
		$a = WinWait("Bomgar - New Session", "", 1)
		WinActivate($a)
		Send("{UP}{UP}{UP}{UP}{ENTER}")
	EndIf
EndFunc   ;==>bomgar

Func day()
	Opt("SendKeyDelay", 10)
	Global $sid = ""
	Global $timet = ""
	;get_id()
	$step = ""
	Return
EndFunc   ;==>day


Func notepad()


	Global $file = "M:\" & $months & " Data\" & @MDAY & "-" & $months & "-" & @YEAR & ".txt"

	; Create a temporary file to write data to.


	If Not FileExists($file) Then
		_FileCreate($file)
	EndIf

	; Open the file for writing (append to the end of a file) and store the handle to a variable.
	Local $hFileOpen = FileOpen($file, $FO_APPEND)
	If $hFileOpen = -1 Then


		Global $file = "Z:\" & $months & " Data\" & @MDAY & "- " & $months & " -" & @YEAR & ".txt"

		; Create a temporary file to write data to.


		If Not FileExists($file) Then
			_FileCreate($file)
		EndIf

		; Open the file for writing (append to the end of a file) and store the handle to a variable.
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		If $hFileOpen = -1 Then
			Global $file = "Y:\" & $months & " Data\" & @MDAY & "- " & $months & " -" & @YEAR & ".txt"

			; Create a temporary file to write data to.


			If Not FileExists($file) Then
				_FileCreate($file)
			EndIf

			; Open the file for writing (append to the end of a file) and store the handle to a variable.
			Local $hFileOpen = FileOpen($file, $FO_APPEND)
			If $hFileOpen = -1 Then
				Global $file = @UserProfileDir & "\Desktop\Data\" & @MDAY & "- " & $months & " -" & @YEAR & ".txt"

				; Create a temporary file to write data to.


				If Not FileExists($file) Then
					_FileCreate($file)
				EndIf

				; Open the file for writing (append to the end of a file) and store the handle to a variable.
				Local $hFileOpen = FileOpen($file, $FO_APPEND)
				If $hFileOpen = -1 Then


					MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
					Return False
				EndIf


				Return False
			EndIf
		EndIf
	EndIf

	; Write data to the file using the handle returned by FileOpen.
	Global $log = $SD & " " & @HOUR & ":" & @MIN & "  " & @MDAY & "/" & $months & "/" & @YEAR & "  " & $sid & " " & $name & "    	" & GUICtrlRead($Combo1) & " " & GUICtrlRead($Combo2) & " Issue  "

	; FileWriteLine($hFileOpen, "____________________________________________________________________________________")
	FileWriteLine($hFileOpen, $log)


	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)

	; Display the contents of the file passing the filepath to FileRead instead of a handle returned by FileOpen.
	MsgBox($MB_SYSTEMMODAL, "Please find Daily Log in Shared Drive month folder", "M:\Month Data\date.txt" & @CRLF & FileRead($file), 1)
	; Delete the temporary file.
	; FileDelete($sFilePath)
	$step = ""

	;+=========================================================================================================


EndFunc   ;==>notepad

Func notepad1()

	Global $file = "M:\" & $month & " Data\" & $user & ".csv"

	; Create a temporary file to write data to.


	If Not FileExists($file) Then
		_FileCreate($file)
	EndIf

	; Open the file for writing (append to the end of a file) and store the handle to a variable.
	Local $hFileOpen = FileOpen($file, $FO_APPEND)
	If $hFileOpen = -1 Then


		Global $file = "Z:\" & $month & " Data\" & $user & ".csv"

		; Create a temporary file to write data to.


		If Not FileExists($file) Then
			_FileCreate($file)
		EndIf

		; Open the file for writing (append to the end of a file) and store the handle to a variable.
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		If $hFileOpen = -1 Then
			Global $file = "Y:\" & $month & " Data\" & $user & ".csv"

			; Create a temporary file to write data to.


			If Not FileExists($file) Then
				_FileCreate($file)
			EndIf

			; Open the file for writing (append to the end of a file) and store the handle to a variable.
			Local $hFileOpen = FileOpen($file, $FO_APPEND)
			If $hFileOpen = -1 Then


				MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
				Return False
			EndIf
		EndIf
	EndIf

	; Write data to the file using the handle returned by FileOpen.
	Global $log = $SD & "," & @MDAY & "/" & @MON & "/" & @YEAR & "," & @HOUR & ":" & @MIN & "," & $sid & "," & $name & "," & GUICtrlRead($Combo1) & GUICtrlRead($Combo2)

	;  FileWriteLine($hFileOpen, "____________________________________________________________________________________")
	FileWriteLine($hFileOpen, $log)


	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)

	; Display the contents of the file passing the filepath to FileRead instead of a handle returned by FileOpen.
	MsgBox($MB_SYSTEMMODAL, "Please find Daily Log in Shared Drive month folder", "M:\Month Data\date.txt" & @CRLF & FileRead($file), 1)
	; Delete the temporary file.
	; FileDelete($sFilePath)
	$step = ""

	;+=========================================================================================================


EndFunc   ;==>notepad1

Func notepad2()

	start_setup()


	Global $file = "\\comutername\as\as\as\" & $months & " Data.xls"

	; Create a temporary file to write data to.


	If Not FileExists($file) Then
		_FileCreate($file)
	EndIf

	; Open the file for writing (append to the end of a file) and store the handle to a variable.
	Local $hFileOpen = FileOpen($file, $FO_APPEND)
	If $hFileOpen = -1 Then
		; Open the file for writing (append to the end of a file) and store the handle to a variable.
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		If $hFileOpen = -1 Then


			MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
			Return False
		EndIf
	EndIf

	; Write data to the file using the handle returned by FileOpen.
	Global $log = $usern & "	" & $SD & "	" & @MON & "/" & @MDAY & "/" & @YEAR & "	" & @HOUR & ":" & @MIN & "	" & $sid & "	" & $name & "	" & GUICtrlRead($Combo1) & "	" & GUICtrlRead($Combo2)

	; FileWriteLine($hFileOpen, "____________________________________________________________________________________")
	FileWriteLine($hFileOpen, $log)


	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)

	; Display the contents of the file passing the filepath to FileRead instead of a handle returned by FileOpen.
	;  MsgBox($MB_SYSTEMMODAL, "Please find Daily Log in Shared Drive month folder", "M:\Month Data\date.txt" & @CRLF & FileRead($file),1)
	; Delete the temporary file.
	; FileDelete($sFilePath)
	$step = ""

EndFunc   ;==>notepad2

Func FileCreate($file, $sString)
	Local $bReturn = True ; Create a variable to store a boolean value.
	If FileExists($file) = 0 Then $bReturn = FileWrite($file, $sString) = 1 ; If FileWrite returned 1 this will be True otherwise False.
	Return $bReturn ; Return the boolean value of either True of False, depending on the return value of FileWrite.
EndFunc   ;==>FileCreate

Func sendmail()


	GUISetState(@SW_SHOW, $Form2)
	Global $st = 4

EndFunc   ;==>sendmail

Func thumbdrive()

	 #CS
; 	;=====
; 	; HTML
; 	;=====
;
; 	Global $oOutlook = _OL_Open()
; 	FileInstall("\\comutername\as\as\Thumb Drive.oft", @TempDir & "\Thumb Drive.oft", 1)
; 	Sleep($d_time)
; 	$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @TempDir & "\Thumb Drive.oft", "SentOnBehalfOfName=ha@company.com", "to=" & $mail, "cc=Home Agent Support;" & $smail & ";" & $mmail)
; 	Sleep($d_time)
; 	FileDelete(@TempDir & "\Thumb Drive.oft")
; 	$oItem.BodyFormat = $olFormatHTML
; 	Sleep($d_time)
; 	$oItem.GetInspector
; 	Sleep($d_time)
; 	$sBody = $oItem.HTMLBody
; 	Sleep($d_time)
; 	$oItem.HTMLBody = "" & $sBody
; 	Sleep($d_time)
; 	$oItem.Display
;
;
; 	Opt("WinTitleMatchMode", 2)
; 	Global $S1 = WinWait("IR#0000000", ""); s1 selected to match minmize option on tt1 ment for first case
; 	WinActivate($S1)
;
;
; 	start_setup()
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "XXXXXX", $fname & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000", " " & $SD & " ")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "<Engineer Name>", $usern & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
;
;
; 	check_issue()
; 	WinActivate($S1)
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Subject")
;
;
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000 - Closed", $SD & " Closing Ticket " & $logs)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Subject=" & $PropertyUpd)
 #CE



	#cs

		;===========
		; Plain Text
		;===========
		#cs
			$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "","\\comutername\as\as\as.oft","to=ashish.saklani@company.com","cc=ashish.saklani@company.com ;ashish.saklani@company.com", "Subject=Subject for BodyFormat=Plain Text")
			$oItem.BodyFormat = $olFormatPlain
			$oItem.GetInspector
			$sBody = $oItem.Body
			$oItem.Body = ""
			$oItem.Display

		#ce
		ControlClick($S1, "","Button2", "",2)
		ControlSend($S1, "","", "{UP}{DOWN}{ENTER}")
		sleep ($d_time *1 )
		$oMail = ObjGet("","Outlook.Application")
		If $oMail.ActiveExplorer.Selection.Count = 1 Then
		$text = $oMail.ActiveExplorer.Selection.Item(1).Body
		sleep(500)

		EndIf


		StringReplace($text,"XXXXXX","Mr ABC")
		MsgBox(0,"",$text)

		;==========
		; Rich Text
		;==========
		$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", "", "Subject=Subject for BodyFormat=RTF")
		$oItem.BodyFormat = $olFormatRichText
		$oItem.GetInspector
		$sBody = $oItem.HTMLBody
		$oItem.HTMLBody = "RTF: <b>Mail</b> Text" & $sBody
		$oItem.Display
	#ce
	$step = ""
EndFunc   ;==>thumbdrive

Func headset()
#CS
; 	;=====
; 	; HTML
; 	;=====
; 	Global $oOutlook = _OL_Open()
; 	FileInstall("\\comutername\as\as\Headset.oft", @TempDir & "\Headset.oft", 1)
; 	Sleep($d_time)
; 	$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @TempDir & "\Headset.oft", "SentOnBehalfOfName=ha@company.com", "to=" & $mail, "cc=Support;" & $smail & ";" & $mmail)
; 	Sleep($d_time)
; 	FileDelete(@TempDir & "\Headset.oft")
;
; 	$oItem.BodyFormat = $olFormatHTML
; 	Sleep($d_time)
; 	$oItem.GetInspector
; 	Sleep($d_time)
; 	$sBody = $oItem.HTMLBody
; 	Sleep($d_time)
; 	$oItem.HTMLBody = "" & $sBody
; 	Sleep($d_time)
; 	$oItem.Display
;
;
; 	Opt("WinTitleMatchMode", 2)
; 	Global $S1 = WinWait("IR#0000000", ""); s1 selected to match minmize option on tt1 ment for first case
; 	WinActivate($S1)
;
;
; 	start_setup()
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "XXXXXX", $fname & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000", " " & $SD & " ")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "<Engineer Name>", $usern & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
;
;
; 	check_issue()
; 	WinActivate($S1)
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Subject")
;
;
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000 - Closed", $SD & " Closing Ticket " & $logs)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Subject=" & $PropertyUpd)
;
 #CE


	$step = ""
EndFunc   ;==>headset

Func ISP()

	;=====
	; HTML
	;=====


	$step = ""
EndFunc   ;==>ISP

Func Hardware()

#CS 	;=====
; 	; HTML
; 	;=====
; 	Global $oOutlook = _OL_Open()
;
;
;
;
; 	FileInstall("\\comutername\as\as\Computer Issue.oft", @TempDir & "\Computer Issue.oft", 1)
; 	Sleep($d_time)
; 	$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @TempDir & "\Computer Issue.oft", "SentOnBehalfOfName=g.support@company.com", "to=" & $mail, "cc=Home Agent Support;" & $smail & ";" & $mmail)
; 	Sleep($d_time)
; 	FileDelete(@TempDir & "\Computer Issue.oft")
;
;
; 	$oItem.BodyFormat = $olFormatHTML
; 	Sleep($d_time)
; 	$oItem.GetInspector
; 	Sleep($d_time)
; 	$sBody = $oItem.HTMLBody
; 	Sleep($d_time)
; 	$oItem.HTMLBody = "" & $sBody
; 	Sleep($d_time)
; 	$oItem.Display
;
;
; 	Opt("WinTitleMatchMode", 2)
; 	Global $S1 = WinWait("IR#0000000", ""); s1 selected to match minmize option on tt1 ment for first case
; 	WinActivate($S1)
;
;
; 	start_setup()
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "XXXXXX", $fname & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000", " " & $SD & " ")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "<Engineer Name>", $usern & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
;
;
; 	check_issue()
; 	WinActivate($S1)
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Subject")
;
;
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000 - Closed", $SD & " Closing Ticket " & $logs)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Subject=" & $PropertyUpd)
;
 #CE


	$step = ""
EndFunc   ;==>Hardware

Func dept_code()
#CS 	;=====
; 	; HTML
; 	;=====
; 	Global $oOutlook = _OL_Open()
; 	FileInstall("\\comutername\as\as\dept code.oft", @TempDir & "\dept code.oft", 1)
; 	Sleep($d_time)
; 	$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @TempDir & "\dept code.oft", "SentOnBehalfOfName=gst.support@company.com", "to=" & $mail, "cc= Support;" & $smail & ";" & $mmail)
; 	Sleep($d_time)
; 	FileDelete(@TempDir & "\dept code.oft")
;
; 	$oItem.BodyFormat = $olFormatHTML
; 	Sleep($d_time)
; 	$oItem.GetInspector
; 	Sleep($d_time)
; 	$sBody = $oItem.HTMLBody
; 	Sleep($d_time)
; 	$oItem.HTMLBody = "" & $sBody
; 	Sleep($d_time)
; 	$oItem.Display
;
;
; 	Opt("WinTitleMatchMode", 2)
; 	Global $S1 = WinWait("IR#0000000", ""); s1 selected to match minmize option on tt1 ment for first case
; 	WinActivate($S1)
;
;
; 	start_setup()
; 	check_date()
; 	check_issue()
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "XXXXXX", $fname)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000", " " & $SD & " ")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "<Engineer Name>", $usern & " ,")
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
; 	$PropertyUpd = StringReplace($Property[1][1], "<issue>", $logs)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)
;
;
;
;
;
;
; 	check_issue()
; 	WinActivate($S1)
; 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Subject")
;
;
; 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000 -", $SD & " Closing Ticket " & $logs)
; 	_OL_ItemModify($oOutlook, $oItem, Default, "Subject=" & $PropertyUpd)
;
 #CE


	$step = ""
EndFunc   ;==>dept_code

Func quiz()


	Local Const $sFilePath = "\\comutername\as\as\as\Test.xls"

	Global $iCount = _FileCountLines($sFilePath) ; Retrieve the number of lines in the current script.


	Global $record1 = ""
	Global $Record = ""
	Global $i = 1

	Global $hFileOpen = FileOpen($sFilePath, $FO_READ)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
		Return False
	EndIf

	Global $iCount = _FileCountLines($sFilePath)
	Global $score = 0
	For $i = 1 To $iCount Step 1


		Global $hFileOpen = FileOpen($sFilePath, $FO_READ)
		Global $iCount = _FileCountLines($sFilePath)

		; Read the fist line of the file using the handle returned by FileOpen.
		Global $sFileRead = FileReadLine($hFileOpen, $i)
		Global $data = StringSplit($sFileRead, "	")
		FileClose($hFileOpen)







		$sMsg = "Move this window around the screen to see some of the child windows centre themselves upon it "
		$sMsg &= "when they display." & @CRLF & @CRLF
		$sMsg &= "If you place the window too close to the edge of the screen, the child windows will "
		$sMsg &= "adjust their position automatically to prevent being partially hidden."
		If @Compiled = 0 Then $sMsg &= @CRLF & @CRLF & "Look in the SciTE console to see the return values as buttons are pressed"

		; Set the message box right justification, not TOPMOST, disable closure, colours (orange text on green background) and change font
		_ExtMsgBoxSet(1 + 2 + 64, 0, 0x000, 0x00FF00, 12, "MS Sans Serif")




		$sMsg1 = " Q" & $data[1] & @CRLF & @CRLF
		$sMsg1 &= "1" & $data[2] & @CRLF
		$sMsg1 &= "2" & $data[3] & @CRLF
		$sMsg1 &= "3" & $data[4] & @CRLF
		$sMsg1 &= "4" & $data[5] & @CRLF

		$iRetValue = _ExtMsgBox(128, "1|2|3|&4", "Question " & $i, $sMsg1, 15, 1000, 200)
		If $iRetValue = $data[6] Then $score = $score + 1



	Next


	FileClose($hFileOpen)

	;_speak($Default, "Your score is   ," & $score & "Out of " & $iCount)

	;ConsoleWrite(@error & @CRLF)
	;ConsoleWrite("Test 6 returned: " & $iRetValue & @CRLF)

	; Reset to default
	_ExtMsgBoxSet(Default)



	$step = ""

EndFunc   ;==>quiz

Func check_mouse()
	Do
		Sleep(20)
	Until MouseGetCursor() <> 15
EndFunc   ;==>check_mouse

Func hide()

	GUISetState(@SW_HIDE, $Form1)




	GUISetState(@SW_ENABLE, $Form3)
		GUICtrlDelete($Button20)
			GUICtrlDelete($Button21)
				GUICtrlDelete($Button22)
					GUICtrlDelete($Button23)
						GUICtrlDelete($Button24)
	Switch $st

		Case 0
						$Button20 = GUICtrlCreateButton("Start Browser", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button20, "Step0")
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()



		Case 1
			$Button21 = GUICtrlCreateButton("Relate IM", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button21, "Step1")
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()



		Case 2


			$Button22 = GUICtrlCreateButton("Copy notes", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button22, "Step2")
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()

		Case 3

			$Button23 = GUICtrlCreateButton("Close ticket", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button23, "Step3")
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()

		Case 4

			$Button24 = GUICtrlCreateButton("Send Email", 0, 0, 90, 20)
			GUICtrlSetOnEvent($Button24, "Step4")
			GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
			set_color()



	EndSwitch

		GUISetState(@SW_SHOW, $Form3)


forward()
back()

			$step = ""
			Return

EndFunc   ;==>hide

Func show()

	GUISetState(@SW_HIDE, $Form3)



	GUISetState(@SW_ENABLE, $Form1)
	GUICtrlDelete($Button0)
	GUICtrlDelete($Button1)
	GUICtrlDelete($Button2)
	GUICtrlDelete($Button3)
	GUICtrlDelete($Button4)

	Switch $st
		Case 0

			$Button0 = GUICtrlCreateButton("Start Browser", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button0, "Step0")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")



		Case 1

			$Button1 = GUICtrlCreateButton("Relate IM", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button1, "Step1")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")




		Case 2
			$Button2 = GUICtrlCreateButton("Copy notes", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button2, "Step2")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")


		Case 3
			$Button3 = GUICtrlCreateButton("Close ticket", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button3, "Step3")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")


		Case 4


			$Button4 = GUICtrlCreateButton("Send Email", 20, 60, 135, 30)
			GUICtrlSetOnEvent($Button4, "Step4")
			GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")



	EndSwitch

	GUISetState(@SW_SHOW, $Form1)



forward()
back()



EndFunc   ;==>show

Func reset()
	GUIDelete($Form1)
	GUIDelete($Form2)
	GUIDelete($Form3)
	SMS()

EndFunc   ;==>reset

Func wait_msg()

;~ $element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div[4]/table/tbody/tr[1]/td/div")

;~ ;/html/body/div[2]/a
;~ ConsoleWrite($element)
;~  local $Currentx = MouseGetPos(0)+1
;~   local $Currenty = MouseGetPos(1)+1
;~ MouseMove($Currentx,$Currenty)

;~ 	Sleep(5)
EndFunc   ;==>wait_msg

Func open_chat()
	Opt("WinTitleMatchMode", 2)



	If $chat = 0 Then

		WinActivate("[CLASS:SunAwtFrame]")
	;SoundPlay(@WindowsDir & "\media\tada.wav", .5)
	;sleep(1000)

		SoundSetWaveVolume(100)
		send("{VOLUME_UP 100}")
		sleep(1000)

			SoundPlay(@WindowsDir & "\media\Festival\Windows Exclamation.wav", 1)
			SoundPlay(@WindowsDir & "\media\Festival\Windows Hardware Fail.wav", 1)
			sleep(1000)
			;_speak($Default, " YOU! GOT A CHAT!" )
		;if WinExists($chatwin) and ($state =31  or $state = 15) Then
		SoundSetWaveVolume(30)
		send("{VOLUME_DOWN 30}")


		MsgBox("", "", "You Have a chat ")
	EndIf
	$chat = 1
EndFunc   ;==>open_chat


Func test1()






Global $file = @ScriptDir & "\test.txt"
; Open the file for writing (append to the end of a file) and store the handle to a variable.
 			Local $hFileOpen = FileOpen($file, $FO_READ)
 			Local $sFileRead = FileReadLine($hFileOpen, 1)
 			Local $script = FileReadLine($hFileOpen, 2)
 			FileClose($hFileOpen)



			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$sFileRead)
			_WD_ExecuteScript($sSession,$script)

			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

;~ 			_WD_ElementAction($sSession,$element,"click")

;~ 			_WD_ExecuteScript($sSession,$script)

;~ 			_WD_ElementActionEx($sSession, $element, 'doubleclick')
			_WD_ElementAction($sSession,$element,'value', "AAA")

			;ConsoleWrite("|"&$sFileRead&"|")














	GUISetState(@SW_HIDE, $Form2)
	HotKeySet("{ESC}", "esc")
	$filename = "C:\Users\%userprofle%\Downloads\worksheet.xlsx"
	$filename = @ScriptDir & "\Worksheet.xlsx"
	$cellrange = "A1:J150"
	If NOT FileExists($filename) Then
		MsgBox(0, "Excel Data Test", "Error: Can't find file " & $filename)
		Exit
	EndIf
	$oexceldoc = ObjGet($filename)
	If (NOT @error) AND IsObj($oexceldoc) Then
		$odocument = $oexceldoc.worksheets(1)
		$aarray = $odocument.range($cellrange).value
		If IsArray($aarray) AND UBound($aarray, 0) > 0 Then
			MsgBox(0, "", "TEST program any issue please contact :Ashish Saklani ")
			MsgBox(0, "Excel Data Test", "Debugging information for retrieved cells:" & @CRLF & "Number of dimensions: " & UBound($aarray, 0) & @CRLF & "Size of first dimension: " & UBound($aarray, 1) & @CRLF & "Size of second dimension: " & UBound($aarray, 2))
			$string = ""
			For $x = 0 To UBound($aarray, 1) - 1
				For $y = 0 To UBound($aarray, 2) - 1
					$string = $string & "(" & $x & "," & $y & ")=" & $aarray[$x][$y] & @CRLF
				Next
			Next
			MsgBox(0, "Excel Data Test", "Debug information: Read Cell contents: " & @CRLF & $aarray[0][0])
			For $i = 0 To 149 Step 1
				Global $sid = $aarray[1][$i]
				Global $aim= $aarray[4][$i]
				Global $issue= $aarray[5][$i]
				Global $time= $aarray[6][$i]
				Global $ltncy= $aarray[9][$i]


				MsgBox("","", "id= "&$sid&"  IM ="&$aim&" issue ="&$issue&" time= "&$time&" Latency="&$ltncy)


			Global $value4 = "Master"
			Global $logs = "Linphone Error "&$time&" EST:Call Dropped"
			Global $logs1 = "Latency found: "&$ltncy

			if $aim == "" Then

					check_sms()
		;			id()
		;			slogs()


			EndIf











			Next
		Else
			MsgBox(0, "Excel Data Test", "Error: Could not retrieve data from cell range: " & $cellrange)
		EndIf
		$oexceldoc.saved = 1
	Else
		MsgBox(0, "Excel Data Test", "Error: Could not open " & $filename & " as an Excel Object.")
	EndIf
	$oexceldoc = 0

EndFunc


func first_followup()

$reason = "First Followup"
 check_mail_details($reason,$sSession)

EndFunc


func No_Details()

$reason = "Incomplete Details"
 check_mail_details($reason,$sSession)

EndFunc



func Client_Application()

$reason = "Client Application"
 check_mail_details($reason,$sSession)

EndFunc


func Replacement_Details()

$reason = "Replacement Details"
 check_mail_details($reason,$sSession)

EndFunc


func second_followup()

$reason = "Second Followup"
 check_mail_details($reason,$sSession)

EndFunc


func lost_session()

$reason = "Lost Session"
 check_mail_details($reason,$sSession)

EndFunc

Func expire_closer()


$reason = "Closer Notification Session Expired"
 check_mail_details($reason,$sSession)

EndFunc


func home_network()

$reason = "Home network - ISP"
 check_mail_details($reason,$sSession)

EndFunc


func no_response()

$reason = "No Response"
 check_mail_details($reason,$sSession)

EndFunc

#cs
	;==================================================Read only Input=============================================================
	$sid = GUICtrlCreateInput("", 44,0, 60, 20, BitOR($ES_AUTOHSCROLL,$ES_READONLY))
	;==================================================Read only Input=============================================================



	;==================================================Check record Function=============================================================



	;HotKeySet("{ENTER}", "_Continue")
	;Dim $binFlag = False
	; Create a constant variable in Local scope of the filepath that will be read/written to.
	Local Const $sFilePath = "\\comutername\as\as\as\"&$months&" Data.xls"

	Global $iCount= _FileCountLines($sFilePath) ; Retrieve the number of lines in the current script.


	Global $record1 = ""
	Global $Record =""
	Global $i=1


	for $i= 1 to $iCount step 1

	Global $iCount= _FileCountLines($sFilePath)

	; Open the file for reading and store the handle to a variable.

	Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
	If $hFileOpen = -1 Then
	MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
	Return False
	EndIf



	; Read the fist line of the file using the handle returned by FileOpen.
	Local $sFileRead = FileReadLine($hFileOpen,$i)


	Global $Record= StringInStr ($sFileRead,$sid)
	$month = @MON
	if $month < 10 then $month = StringRight($month,1)

	Global $twodays= StringInStr ($sFileRead,@MON &"/"&@MDAY&"/"& @YEAR )

	if $twodays = 0 or $twodays = 1 then
	$twodays = StringInStr ($sFileRead,$month &"/"&@MDAY-1&"/"& @YEAR )
	if $twodays = 0 or $twodays = 1 then
	$twodays = StringInStr ($sFileRead,$month &"/"&@MDAY-2&"/"& @YEAR )
	if $twodays = 0 or $twodays = 1 then
	$twodays = StringInStr ($sFileRead,$month&"/"&@MDAY-3&"/"& @YEAR )
	if $twodays = 0 or $twodays = 1 then
	$twodays =0
	else
	$twodays =2

	EndIf
	else
	$twodays =2

	EndIf
	else
	$twodays =2

	EndIf
	else
	$twodays =2

	EndIf


	if $Record <> 0 and $Record <> 1 and $twodays <> 0 Then $record1 &= $sFileRead &@LF


	; Close the handle returned by FileOpen.
	FileClose($hFileOpen)

	Next


	; Display the first line of the file.

	$data="Tech Name		TicketNo	Date		Time	Emp ID		Agent Name		 "&@LF
	$data&="																"&@LF




	if $record1 <> "" then

	_ExtMsgBoxSet(1, 2, 0x004080, 0xFFFF00, 10, "Comic Sans MS")
	_ExtMsgBoxSet(-1, -1, -1, -1, -1, -1, @DesktopWidth -100)
	_ExtMsgBox ("", "CONFIRM", "Agents Record of past 48 hrs",$data &$record1,80)

	EndIf

	;$iCount= StringSplit($record1,@LF)
	;Local $inum = UBound($iCount, $UBOUND_ROWS)
	;$inum=$inum-1
	;MsgBox("","",$inum)
	;	   SplashTextOn("Agents Record of past 48 hrs Press {ENTER} to continue",$data&$record1, 1000, $inum*50, -1, -1, -1, "", 10)
	;	Do
	;		Sleep(100)
	;	Until $binFlag = True

	; Delete the temporary file.
	;FileDelete($sFilePath)




#ce






Func outlook_mail($agent_sso,$tl_email,$reason,$SD,$issue,$Message)
		Global $oOutlook = _OL_Open()

							$d_time =300
							Sleep($d_time)
							$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @ScriptDir & "\_OL_ItemCreate.oft", "SentOnBehalfOfName=ha@company.com", "to=" & $agent_sso, "cc="& $tl_email&";homeagentsupport@company.com", "Subject="&$reason&":"&$SD&$issue,"Body="&$Message )
							;$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @ScriptDir & "\_OL_ItemCreate.oft, "SentOnBehalfOfName=ha@company.com", "to=" & $mail, "cc=Home Agent Support;" & $smail & ";" & $mmail)
							$oItem.BodyFormat = $olFormatHTML
							Sleep($d_time)
							$oItem.GetInspector
							Sleep($d_time)
							$sBody = $oItem.HTMLBody
							Sleep($d_time)
							;$oItem.HTMLBody = ""
							Sleep($d_time)
							$oItem.Display


EndFunc


Func launch_mail($name_a,$agent_sso,$tl_email,$SD,$issue,$reason,$loc_a)

local $a =""

_WD_Attach($sSession,'BMC Remedy (Search)')
sleep(555)

if $loc_a =="USA" Then
	$a = "Live chat Portal ( https://helpdesk.company.com/NA)"
ElseIf $loc_a =="CAN" Then
		$a = "Live chat Portal ( https://helpdesk.company.com/NA )"
ElseIf $loc_a =="PHL" Then
		$a = "Live chat Portal ( https://helpdesk.company.com/PH )"
ElseIf $loc_a =="IND" Then
		$a = "Live chat Portal ( https://helpdesk.company.com/IN )"

else
		$a="remote support process from website https://re.company.com "
EndIf

$name_a = StringSplit($name_a," ")
$name_a = $name_a[1]
	Switch $reason
		Case "First Followup"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is Followup email with ref to your ticket number "&$SD&" open with us for issue related to "&$issue&@CRLF&@CRLF&".we tried to reach you for troubleshooting on this issue, but we are not able to reach you with contact detail mentioned on ticket, if you are still facing issue please reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Second Followup"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is Second Followup email with ref to your ticket number "&$SD&" open with us for issue related to "&$issue&@CRLF&@CRLF&".we tried to reach you for troubleshooting on this issue, but we are not able to reach you with contact detail mentioned on ticket, if you are still facing issue please reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Lost Session"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is Followup email with ref to your ticket number "&$SD&" We Accepted your session for working on issue "&$issue&" and we lost session while working or you did not approve it from your end so our session waiting for connection approval from your end or its disconnected, if you are still facing issue please reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email."&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Home network - ISP"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" issue related to "&$issue&@CRLF&@CRLF&".While troubleshooting we observed network issue from your side snapshot of network test attached on ticket and email ,to get this resolved please contact your internet provider ,if you still need some assistance or details please reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "No Response"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" for issue related to "&$issue&@CRLF&@CRLF&".we tried to reach you for troubleshooting on the issue several times, but we are not able to reach you with contact detail mentioned on ticket, so we are closing ticket for now ,if you are still facing issue please reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Closer Notification Session Expired"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" for issue related to "&$issue&@CRLF&@CRLF&".we accepted your session for troubleshooting on the issue and we were waiting for 60 mins for you to join the session , but session was not approved from your side or it did not connect, after waiting for 60 mins and no response from your side we hope issue is resolved for you or you don't need support on this issue so we are closing ticket for now."&@CRLF&@CRLF&" if you are still facing issue and need assistance please reach us back with "&$a&". We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. waiting for you quick response"&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "user_PHIL"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" for issue related to "&$issue&@CRLF&@CRLF&"This is to inform correct support Process for you. We are closing ticket for now and We hope issue is resolved for you, "&@CRLF&@CRLF&" but if you are still facing issue please reach us through PHIL ZOOM WAR ROOM meeting ID 96559316715 and password PHIL063 or you can also call Service desk for support . We will try to provide resolution quickly,"&@CRLF&" please select Reply all option if you reply on this email."&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "user_IND"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" for issue related to "&$issue&@CRLF&@CRLF&"This is to inform correct support Process for you. We are closing ticket for now and We hope issue is resolved for you, "&@CRLF&@CRLF&" if you are still facing issue and need assistance please reach us back with "&$a&" or you can also call Service Desk. We will try to provide resolution quickly,"&@CRLF&"  please select Reply all option if you reply on this email."&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Incomplete Details"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is Followup email with ref to your ticket number "&$SD&" open with us for issue related to "&$issue&@CRLF&@CRLF&"Found Incomplete detail on your ticket,"&@CRLF&" Important details for us to followup on ticket are (1)Phone Number (2)Shift time with timezone (3) Week offs (4)company device or personal,(5)Issue details  or you can reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Client Application"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is email with ref to your ticket number "&$SD&" open with us for issue related to "&$issue&@CRLF&@CRLF&"After checking details on this issue, we want to inform you company IT does not manage this application this is managed from client IT team ,to get this issue resolved you may need to report this to client IT team with help from your TL /Floor support ,you can reach us back with "&$a&",If anything required from our side on this issue ,please reach us back with details whats required from our side we will be happy to help.We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "Replacement Details"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This is Followup email with ref to your ticket number "&$SD&" open with us for issue related to "&$issue&@CRLF&@CRLF&"to start processing your request some details are required,ticket can not be assigned to correct team for processing in absence of required detail on ticket,"&@CRLF&" Some details required to start replacement request are  (1)Phone Number (2)Shift time with timezone (3) Week offs (4)company device or personal, (5)Shipping Address (6)Device Make and Model,"&@CRLF&" you can reach us back with "&$a&",We will try to provide resolution quickly,"&@CRLF&@CRLF&" please select Reply all option if you reply on this email. "&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
		Case "WRONG_ASSIGNMENT"
			$Message ="Hi "&$name_a&","&@CRLF&@CRLF&@CRLF&"This email is with ref to your ticket number "&$SD&" for issue related to "&$issue&@CRLF&@CRLF&"This is to inform you ticket is wrongly assigned to our team . We are closing ticket for now and We hope issue is resolved for you, "&@CRLF&@CRLF&" if you are still facing issue Please verify promper support chanel for your team or call Global Service Desk."&@CRLF&"  please select Reply all option if you reply on this email."&@CRLF&@CRLF&@CRLF&"Thanks & regards"&@CRLF&(StringSplit($sso,"."))[1]&" "&(StringSplit($sso,"."))[2]&@CRLF&$sso&"@company.com"
	EndSwitch



ClipPut(" remote support process from website https://re.company.com and select VDI support")

if ((StringLeft($reason,4) <> "user")  or (GUICtrlRead($Combo2) <> "Wrong Assignment")) Then


;Global $mail_launch = False


	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&4&"]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&3&"]"
	EndIf



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[1]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[4]/div[10]/span")

_WD_ExecuteScript($sSession,"document.querySelector('#sub-301648200 > div.VNavLeaf.VNavLevel2.arfid1000005679.ardbnz2NI_EmailSystem').classList.add('VNavHover');")

_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

sleep(500)


$xpath = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[1]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[4]/div[10]/span"
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
				 highlight_item($xpath)
				_WD_ElementAction($sSession,$element,"click")


_WD_LoadWait($sSession)
_WD_Attach($sSession,'Email System')


;~ $element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[4]/fieldset[1]/div[7]/div[2]/div/div[2]/table/tbody/tr[2]/td[5]/nobr")
;~ _WD_LoadWait($sSession,$element)

_WD_LoadWait($sSession)
check_mouse()

;email address



		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[6]/textarea")
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'value',";"&$tl_email&";"&"homeagentsupport@company.com")


;email_subject



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[7]/textarea")
		_WD_ElementAction($sSession,$element,'value', "")
		_WD_ElementAction($sSession,$element,'clear')
		_WD_ElementAction($sSession,$element,'value', $reason&" :"&$SD&" :"&$issue)


;email_Body




$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[8]/textarea")
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'value',$Message)




#cs

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/a[5]/div")
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'click')

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/a[6]/div/div")
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'click')



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="WIN_0_301614900"]/div/div')
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'click')






sleep(3000)



 sleep(2000)


wait_for_value($sSession,"/html/body/div[1]/div[5]/div[8]/textarea")



_WD_LoadWait($sSession)
_WD_Attach($sSession,'Email System')


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/a[6]/div/div")
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'click')



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="WIN_0_301614900"]/div/div')
		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
		_WD_ElementAction($sSession,$element,'click')

#ce



EndIf




external_email(@CRLF&"#"&@CRLF&$agent_sso&"©"&$tl_email&";"&"homeagentsupport@company.com"&"©"& $reason&" :"&$SD&" :"&$issue&"©"& $Message)
outlook_mail($agent_sso,$tl_email,$reason,$SD,$issue,$Message)

	;ConsoleWrite("----i was here ----")

EndFunc



Func external_email($data)

	Global $file = @ScriptDir & "\email.txt"

	If Not FileExists($file) Then
			_FileCreate($file)
	EndIf
			Local $hFileOpen = FileOpen($file,$FO_APPEND)
			FileWriteLine($hFileOpen, $data)
	FileClose($hFileOpen)

EndFunc



Func test()

Global $file = @ScriptDir & "\test.txt"
; Open the file for writing (append to the end of a file) and store the handle to a variable.
 			Local $hFileOpen = FileOpen($file, $FO_READ)
 			Local $sFileRead = FileReadLine($hFileOpen, 1)
 			Local $script = FileReadLine($hFileOpen, 2)
 			FileClose($hFileOpen)



			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$sFileRead)
			_WD_ExecuteScript($sSession,$script)

			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

 			_WD_ElementAction($sSession,$element,"click")

;~ 			_WD_ExecuteScript($sSession,$script)

;~ 			_WD_ElementActionEx($sSession, $element, 'doubleclick')
;			_WD_ElementAction($sSession,$element,'value', "AAA")

			;ConsoleWrite("|"&$sFileRead&"|")


EndFunc


func check_mail_details($reason,$sSession)


_WD_Attach($sSession,'BMC Remedy (Search)')

Local $tl_email,$SD,$issue


$SD=""

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[1]/textarea")
$SD= _WD_ElementAction($sSession, $element,'value')

$issue=""
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")
$issue= _WD_ElementAction($sSession, $element,'value')



	;ConsoleWrite($SD&" "&$sid&""&$issue&@CRLF)

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="WIN_3_303498000"]/div')
_WD_ElementAction($sSession,$element,"click")

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/a[2]/div")
_WD_ElementAction($sSession,$element,"click")


sleep(300)

_WD_Attach($sSession,'People')

_WD_LoadWait($sSession)
check_mouse()
sleep(300)

$sid =""

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_0_1000000054"]')
Global $sid= _WD_ElementAction($sSession, $element,'value')


;close button
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'/html/body/div[1]/div[5]/a[5]/div/div')
_WD_ElementAction($sSession,$element,"click")

;ConsoleWrite("----i was here ----"&$sid&@CRLF)
sleep(300)


$tl_email = Search_tl_details($sSession)
$tl_email = StringSplit($tl_email,"©")

;if $mail_launch == True Then




launch_mail($tl_email[1],$tl_email[2],$tl_email[3],$SD,$issue,$reason,$tl_email[5])

;EndIf


EndFunc


Func check_signIn()
check_mouse()
_WD_LoadWait($sSession,"")
$title= _WD_Action($sSession, "TITLE")




	if $title == "Sign In" Then


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='userNameInput']")

				$xpath = "//input[@id='userNameInput']"
				ConsoleWrite('document.evaluate("'& $xpath & '", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.style.outline = "#f00 solid 2px"')
				 highlight_item($xpath)
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $sso&"@company.com")


				$xpath = "//input[@id='passwordInput']"
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)

				 highlight_item($xpath)
				sleep(555)
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $pass)


				$xpath = "//span[@id='submitButton']"
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $xpath)
				sleep(555)
				 highlight_item($xpath)
				_WD_ElementAction($sSession,$element,"click")



	ElseIf $title == "The company Login"	Then


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='user1']")
				_WD_LoadWait($sSession,500,$element)
				sleep(555)


				$xpath =  "//input[@id='user1']"
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
				sleep(555)
				 highlight_item($xpath)
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $sso&"@company.com")

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='password1']")
				sleep(555)
				 highlight_item($element)
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $pass)


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//select[@id='dominio1']")
				sleep(555)
				 highlight_item($element)
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', "company")



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//*[@id='dhtmlgoodies_tabView2']/div/form/table/tbody/tr[4]/td[3]/input")
				sleep(555)
				 highlight_item($element)
				_WD_ElementAction($sSession,$element,"click")

				_WD_LoadWait($sSession)

	EndIf

EndFunc

func login_chrome()


		Local Const $sFilePath = StringLeft(@WindowsDir,3)&'chrome\'

        ; If the directory exists the don't continue.
        If Not FileExists($sFilePath) Then
  				  DirCreate($sFilePath)
		EndIf




			;FileInstall('C:\Users\%userprofle%\Documents\chromedriver.exe', @TempDir & '\chromedriver.exe', 1)
			FileInstall('C:\Users\%userprofle%\Documents\chromedriver.exe', StringLeft(@WindowsDir,3)&'chrome\chromedriver.exe', 1)

			$_WD_DEBUG = false
			_WD_Option('Driver', StringLeft(@WindowsDir,3)&'chrome\chromedriver.exe')
			_WD_Option('Port', 9515)
			;_WD_Option('DriverParams', '--verbose --log-path="' & @ScriptDir & '\chrome.log"')
			$sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"goog:chromeOptions": {"w3c": true, "excludeSwitches": [ "enable-automation"]}}}}'




			;$sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"goog:chromeOptions": {"w3c": true, "args":["user-data-dir=userData", "--no-sandbox","-incognito"]}}}}'

EndFunc

func wait_for_element($sSession,$path)


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $path)
			while $element == ""
			_WD_Alert($sSession, 'accept')

				sleep(500)

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $path)
			ConsoleWrite($element&"___searched element____")

			;ConsoleWrite($path&@CRLF)
			WEnd

		_WD_HighlightElements($sSession,$element,3)
		sleep (500)
		ConsoleWrite($element&"___Found element____")

EndFunc

func wait_for_value($sSession,$path)
;101712309
	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $path)
	$loc_a= _WD_ElementAction($sSession, $element,'value')

		;$path= '"'&$path&'"'
ConsoleWrite(@CRLF&$element&"_______"&$path&@CRLF)
			while $loc_a == ""

				sleep(500)

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
				$loc_a= _WD_ElementAction($sSession, $element,'value')
				;ConsoleWrite("__value_____"&$loc_a&@CRLF)

				if $loc_a =="" Then
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
					$loc_a= _WD_ElementAction($sSession, $element,'text')

					;ConsoleWrite("__text_____"&$loc_a&@CRLF)
				EndIf

			WEnd

EndFunc

func click_on_value($sSession,$element,$path)

	wait_for_value($sSession,$path)
sleep(2555)
 highlight_item($path)
	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $path)
			_WD_ElementAction($sSession,$element,"click")

EndFunc

Func ajax_select($sSession,$path)
	sleep(1500)

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
		 highlight_item($path)
				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
			;	_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE007")
EndFunc

Func ajax_select_element($sSession,$element)
sleep(1500)

				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE007")
EndFunc

func Search_qc_details($sSession)

				;=============================================================New Tab Quick connect============================================

ConsoleWrite($firsttimeonly&"--first time only 5825-")

				sleep(555)

				if $firsttimeonly == true then
				_WD_NewTab($sSession)
				_WD_LoadWait($sSession)
				_WD_Attach($sSession,'about:blank')
				sleep(555)
				ConsoleWrite($firsttimeonly&"--first time only i was in-")
				EndIf
				_WD_Attach($sSession,('Profile for' or 'Profiler' or 'website Lookup'))


				_WD_Navigate($sSession, "websiteID="&$sid)
				_WD_LoadWait($sSession)




			check_mouse()
			check_signIn()
			check_mouse()





				sleep(1555)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/form/input[1]")
				_WD_LoadWait($sSession,2500,$element)

				Global $project_a,$email_a,$loc_a,$name_a





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[4]/center/b")
				$email_a= _WD_ElementAction($sSession, $element,'text')


				if $email_a == "" Then
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[5]/center/b")
					$email_a= _WD_ElementAction($sSession, $element,'text')
				EndIf


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[2]/center/b")
				$name_a= _WD_ElementAction($sSession, $element,'text')





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[4]/center/b")
				$project_a= _WD_ElementAction($sSession, $element,'text')





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[6]/center/b")
				$loc_a= _WD_ElementAction($sSession, $element,'text')




				;$loc_a=StringRight(StringLeft($loc_a, 4),3)
				$project_a=StringLeft($project_a, 8)
				$email_a= (StringSplit($email_a,"@"))[1]
				;$name_a = StringTrimLeft($name_a, 11)

	;			_WD_Window($sSession, 'close')

				if $loc_a =="" then
					_WD_Navigate($sSession, "websiteID="&$sid)

			check_mouse()
			check_signIn()
			check_mouse()

					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/right/center/table[1]/tbody/tr[2]/td[2]")
					_WD_LoadWait($sSession,500,$element)
					sleep(555)



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[1]")
				$project_a= _WD_ElementAction($sSession, $element,'text')





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[3]/tbody/tr[2]/td[4]")
				$loc_a= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[2]")
				$name_a= _WD_ElementAction($sSession, $element,'text')

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[3]")
				$email_a= _WD_ElementAction($sSession, $element,'text')



				EndIf

					if (StringInStr($email_a,"." )) then

						$email_a=StringLower($email_a)

						Local $email_b = StringSplit($email_a,".")

						if $loc_a ="PHL" Then

						$email_a = ""&StringLeft($email_b[1],14)&"."&StringLeft($email_b[2],10)&""
						Else
						$email_a = ""&StringLeft($email_b[1],10)&"."&StringLeft($email_b[2],10)&""
						EndIf

					Else
						MsgBox("","","USER EMAIL NOT FOUND COPY user email and click on OK")

						$email_a = ClipGet()

						$email_a=StringLower($email_a)

						$email_a = StringSplit($email_a,'@')
						$email_a = $email_a[1]


					EndIf

				;ConsoleWrite($email_a)

				If $chat == "OB" Then
					$loc_a = "ONBOARDING"
				EndIf

				$loc_a = $loc_a&"©"&$project_a&"©"&$email_a&"©"&$name_a

				;ConsoleWrite($loc_a&"___"&$project_a&"___"&$email_a&"_____")


					_WD_Attach($sSession,'BMC Remedy (Search)')
					sleep(555)

				Global $firsttimeonly = False

				Return $loc_a

;ConsoleWrite($firsttimeonly&"---")


EndFunc

func Search_tl_details($sSession)

				;=============================================================New Tab Quick connect============================================

				_WD_Attach($sSession,"")
				if $firsttimeonly = true then
				_WD_NewTab($sSession)
				_WD_LoadWait($sSession)
				;_WD_Attach($sSession,'about:blank')

				sleep(555)
				EndIf

					_WD_Attach($sSession,('Profile for' or 'Profiler' or 'website Lookup'))

				_WD_Navigate($sSession, "companyID="&$sid)
				_WD_LoadWait($sSession)




			check_mouse()
			check_signIn()
			check_mouse()





				sleep(1555)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/form/input[1]")
				_WD_LoadWait($sSession,2500,$element)

				Local $name_a,$agent_sso,$tl_email,$tl_sid,$agent_flag,$loc_a




				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[2]")
				$name_a= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[3]")
				$agent_flag= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[4]/center/b")
				$agent_sso= _WD_ElementAction($sSession, $element,'text')




				if $agent_sso == "" Then
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[5]/center/b")
					$agent_sso= _WD_ElementAction($sSession, $element,'text')
				EndIf

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[6]/center/b")
				$loc_a= _WD_ElementAction($sSession, $element,'text')







				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[3]/tbody/tr[2]/td[1]/center/b")
				$tl_sid= _WD_ElementAction($sSession, $element,'text')






				_WD_Navigate($sSession, "companyID="&$tl_sid)
				_WD_LoadWait($sSession)




			check_mouse()
			check_signIn()
			check_mouse()


			sleep(300)





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[4]/center/b")
				$tl_email= _WD_ElementAction($sSession, $element,'text')


				if $tl_email == "" Then
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[5]/center/b")
					$tl_email= _WD_ElementAction($sSession, $element,'text')
				EndIf


				;ConsoleWrite($name_a&"  "&$tl_sid&"  "&$tl_email&@CRLF)


				$tl_email=StringLower($tl_email)


				$name_a = $name_a&"©"&$agent_sso&"©"&$tl_email&"©"&$agent_flag&"©"&$loc_a&"©"

				;ConsoleWrite($loc_a&"___"&$project_a&"___"&$email_a&"_____")


					_WD_Attach($sSession,'BMC Remedy (Search)')
					sleep(555)


				Global $firsttimeonly = False


				Return $name_a




EndFunc

func Search_tl_details_orig($sSession)

				;=============================================================New Tab Quick connect============================================

				_WD_Attach($sSession,"")
				if $firsttimeonly = true then
				_WD_NewTab($sSession)
				_WD_LoadWait($sSession)
				;_WD_Attach($sSession,'about:blank')

				sleep(555)
				EndIf

					_WD_Attach($sSession,('Profile for' or 'Profiler' or 'website Lookup'))

			_WD_Navigate($sSession, "website#"&$sid)

			_WD_LoadWait($sSession)
			check_mouse()
			check_signIn()
			check_mouse()





					sleep(1555)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]")
				_WD_LoadWait($sSession,2500,$element)

				Local $name_a,$agent_sso,$tl_email,$tl_sid,$agent_flag,$loc_a






				$i = 1
				While $i < 14
								$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr["&$i&"]/td[2]")
								$email_b= _WD_ElementAction($sSession, $element,'text')

					 				$test_string = StringInStr($email_b ,"@company.com")
										if ($test_string <> 0) and ($test_string <> -1) then
											Global $email_a = $email_b
											Global $test_case =$i
											ConsoleWrite($i&" attempt "& $email_b&"=="&$test_case& $test_string)
										EndIf


								;ConsoleWrite( @CRLF & $i & ",   "& $email_b &@CRLF)
								$i = $i +1


				WEnd












				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/div")
				$name_a= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td[2]')
				$agent_flag= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[9]/td[2]/a')
				$agent_sso= _WD_ElementAction($sSession, $element,'text')




				if $agent_sso == "" Then

					;'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[9]/td[2]/a'
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[16]/td[2]')
					$agent_sso= _WD_ElementAction($sSession, $element,'text')
				EndIf

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[13]/td[2]')
				$loc_a= _WD_ElementAction($sSession, $element,'text')







				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[3]/tbody/tr[2]/td[1]/center/b")
				$tl_sid= _WD_ElementAction($sSession, $element,'text')




;'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[9]/td[2]/a'


				_WD_Navigate($sSession, "companyID="&$tl_sid)
				_WD_LoadWait($sSession)




			check_mouse()
			check_signIn()
			check_mouse()


			sleep(300)





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[4]/center/b")
				$tl_email= _WD_ElementAction($sSession, $element,'text')


				if $tl_email == "" Then
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[5]/center/b")
					$tl_email= _WD_ElementAction($sSession, $element,'text')
				EndIf


				;ConsoleWrite($name_a&"  "&$tl_sid&"  "&$tl_email&@CRLF)


				$tl_email=StringLower($tl_email)


				$name_a = $name_a&"©"&$agent_sso&"©"&$tl_email&"©"&$agent_flag&"©"&$loc_a&"©"

				;ConsoleWrite($loc_a&"___"&$project_a&"___"&$email_a&"_____")


					_WD_Attach($sSession,'BMC Remedy (Search)')
					sleep(555)


				Global $firsttimeonly = False


				Return $name_a




EndFunc

func Search_qc_details_orig($sSession)

				;=============================================================New Tab Quick connect============================================


				if $firsttimeonly = true then
					_WD_NewTab($sSession)
				EndIf

				_WD_LoadWait($sSession)
				;_WD_Attach($sSession,'about:blank')
				_WD_Attach($sSession,('Profile for' or 'Profiler' or 'website Lookup'))
				sleep(555)

				_WD_Navigate($sSession, "website#"&$sid)

				check_signIn()
				_WD_LoadWait($sSession)


				if $firsttimeonly = true then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='user1']")
				_WD_LoadWait($sSession,500,$element)
				sleep(555)



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='user1']")
				sleep(555)
				 highlight_item("//input[@id='user1']")
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $sso&"@company.com")

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//input[@id='password1']")
				sleep(555)
				 highlight_item("//input[@id='password1']")
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', $pass)


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//select[@id='dominio1']")
				sleep(555)
				 highlight_item("//select[@id='dominio1']")
				_WD_ElementAction($sSession,$element,"click")
				_WD_ElementAction($sSession,$element,'value', "company")



				$xpath = "//*[@id='dhtmlgoodies_tabView2']/div/form/table/tbody/tr[4]/td[3]/input"
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, $xpath)
				sleep(555)
				 highlight_item($xpath)
				_WD_ElementAction($sSession,$element,"click")

				_WD_LoadWait($sSession)

				_WD_Navigate($sSession, "website#"&$sid)

				EndIf

				sleep(1555)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]")
				_WD_LoadWait($sSession,2500,$element)

				Global $project_a,$email_a,$loc_a,$name_a,$email_b

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[8]/td[2]/a")
				$email_a= _WD_ElementAction($sSession, $element,'text')
				Global $test_case = 8
				$email_b = $email_a

		;	'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[10]/td[2]



				;ConsoleWrite("First attempt "& $email_a)
				$test_string = StringInStr($email_a ,"@company.com")

				if ($test_string == 0) or ($test_string == -1) then

				$i = 1
				While $i < 14
								$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr["&$i&"]/td[2]")
								$email_b= _WD_ElementAction($sSession, $element,'text')

					 				$test_string = StringInStr($email_b ,"@company.com")
										if ($test_string <> 0) and ($test_string <> -1) then
											Global $email_a = $email_b
											Global $test_case =$i
											ConsoleWrite($i&" attempt "& $email_b&"=="&$test_case& $test_string)
										EndIf


								;ConsoleWrite( @CRLF & $i & ",   "& $email_b &@CRLF)
								$i = $i +1


				WEnd

           EndIf

	ConsoleWrite(@CRLF&"second last attempt "& $email_a&"=="&$email_b&"--"&$test_case&"--"& $test_string)




if $test_case = 8 Then




$xpath = "//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]")
				$project_a= _WD_ElementAction($sSession, $element,'text')



$xpath = "//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]")
				$loc_a= _WD_ElementAction($sSession, $element,'text')



ElseIf $test_case = 9 Then


$xpath = "//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]")
				$project_a= _WD_ElementAction($sSession, $element,'text')




$xpath = "//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]")
				$loc_a= _WD_ElementAction($sSession, $element,'text')


ElseIf $test_case = 10 Then



$xpath = "//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td[2]"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='nameFieldContainer']/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td[2]")
				$project_a= _WD_ElementAction($sSession, $element,'text')




$xpath = '//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]'
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="nameFieldContainer"]/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]')
				$loc_a= _WD_ElementAction($sSession, $element,'text')




EndIf

$xpath = "/html/body/div/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/div"
				 highlight_item($xpath)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div/table/tbody/tr[2]/td/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/div")
				$name_a= _WD_ElementAction($sSession, $element,'text')

ConsoleWrite(@CRLF&"last attempt "& $email_a&"=="&$test_case& $test_string)

				$loc_a = StringRight(StringLeft($loc_a, 4),3)
				$project_a = StringLeft($project_a, 8)
				$email_a = (StringSplit( $email_a,"@"))[1]
						if (StringInStr($email_a,"." )) then

						$email_a=StringLower($email_a)

						Local $email_b = StringSplit($email_a,".")

						if $loc_a ="PHL" Then

						$email_a = ""&StringLeft($email_b[1],14)&"."&StringLeft($email_b[2],10)&""
						Else
						$email_a = ""&StringLeft($email_b[1],10)&"."&StringLeft($email_b[2],10)&""
						EndIf

					Else
						MsgBox("","","USER EMAIL NOT FOUND COPY user email and click on OK")

						$email_a = ClipGet()

						$email_a=StringLower($email_a)

						$email_a = StringSplit($email_a,'@')
						$email_a = $email_a[1]


					EndIf
				$name_a = StringTrimLeft($name_a, 11)

	;			_WD_Window($sSession, 'close')

	ConsoleWrite(@CRLF&"second last attempt "& $email_a&"=="&$test_case& $test_string)
;~ 	ConsoleWrite(@CRLF&"last attempt "& $email_a&"=="&$test_case& $test_string)

				if $email_a == "" then
					_WD_Navigate($sSession, "websiteID="&$sid)

					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/right/center/table[1]/tbody/tr[2]/td[2]")
					_WD_LoadWait($sSession,500,$element)
					sleep(555)



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[2]/tbody/tr[2]/td[1]")
				$project_a= _WD_ElementAction($sSession, $element,'text')





				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[3]/tbody/tr[2]/td[4]")
				$loc_a= _WD_ElementAction($sSession, $element,'text')


				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[2]")
				$name_a= _WD_ElementAction($sSession, $element,'text')

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/right/center/table[1]/tbody/tr[2]/td[3]")
				$email_a= _WD_ElementAction($sSession, $element,'text')


					if $email_a =="" then
						$email_a = ClipGet()
					EndIf
				EndIf



				$email_a=StringLower($email_a)
				$loc_a = $loc_a&"©"&$project_a&"©"&$email_a&"©"&$name_a

				;ConsoleWrite($loc_a&"___"&$project_a&"___"&$email_a&"_____")
				Global $firsttimeonly = False
				Return $loc_a


EndFunc

func login_remedy()



$sSession = _WD_CreateSession($sDesiredCapabilities)


				_WD_Navigate($sSession, "website")




				_WD_LoadWait($sSession)

				sleep(1555)

				;FileReadLine ( "filehandle/filename" [, line = 1] )


check_mouse()
check_signIn()
check_mouse()




				;=============================================================login started=============================================


				wait_for_element($sSession,"//img[@id='reg_img_304248660']")


				_WD_LoadWait($sSession,500)


				;ConsoleWrite($element&"_")

					sleep(1000)

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "//td[@class='trimJustc']")


					sleep(1000)
				;ConsoleWrite($element&"_")



				;_WD_Alert($sSession, 'accept')


					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div[2]/a")
				 highlight_item($element)
				_WD_ElementAction($sSession,$element,"click")

				;  id  PopupMsgBox   //*[@id="PopupMsgFooter"]/a     /html/body/div[2]/a



				;=============================================================login complete==============================================

EndFunc

Func system_type($str1)

        Switch $str1
			Case StringInStr ( $str1, "user" ) > 0
                $sOut = "BYOH"
				$sum_line[2] = "user"
			Case StringInStr ( $str1, "BASS" ) > 0
                $sOut = "BYOH"
				$sum_line[2] = "CABELAS"
			Case StringInStr ( $str1, "CHRY" ) > 0
                $sOut = "BYOD"
				 $sum_line[2] = "CHRYSTLER"
            Case StringInStr ( $str1, "EYEM" ) > 0
                $sOut = "BYOD"
				 $sum_line[2] = "EYEMED"
			Case StringInStr ( $str1, "BIO" ) > 0
				$sOut = "BYOD"
				 $sum_line[2] = "BIO_REF"
			Case StringInStr ( $str1, "SPRI" ) > 0
                $sOut = "BYOD"
				 $sum_line[2] = "SPRINT"
			Case StringInStr ( $str1, "WYND" ) > 0
                $sOut = "BYOD"
				$sum_line[2] = "WHYNDHAM"
            Case StringInStr ( $str1, "ANTH" ) > 0
                $sOut = "BYOH"
				$sum_line[2] = "ANTHEM"
			Case StringInStr ( $str1, "HRB" ) > 0
                $sOut = "BYOH"
				$sum_line[2] = "HRB"
            Case StringInStr ( $str1, "USBA" ) > 0
                $sOut = "BYOH"
				$sum_line[2] = "USBANK"
			Case StringInStr ( $str1, "United" ) > 0
				 $sOut = "BYOH"
				 $sum_line[2] = "UHG"
			Case StringInStr ( $str1, "Everest ORx" ) > 0
                $sOut = "BYOH"
				 $sum_line[2] = "UHG"
			Case StringInStr ( $str1, "LGE" ) > 0
                $sOut = "BYOD"
				 $sum_line[2] = "LGE"
			Case StringInStr ( $str1, "PAYP" ) > 0
                $sOut = "IGEL"
				 $sum_line[2] = "PAYPAL"
			Case StringInStr ( $str1, "ACADIA" ) > 0
                $sOut = "BCP"
				 $sum_line[2] = "GOLDMAN"
			Case StringInStr ( $str1, "GOLD" ) > 0
                $sOut = "BCP"
				 $sum_line[2] = "GOLDMAN"
			Case StringInStr ( $str1, "Trane Te" ) > 0
                $sOut = "SRW"
				 $sum_line[2] = "Ingersoll"

			Case Else
                $sOut = "BCP"
        EndSwitch

Return $sOut

EndFunc

Func get_agent_details_from_ticket()





$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="WIN_3_303498000"]/div')
_WD_ElementAction($sSession,$element,"click")

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/a[2]/div")
_WD_ElementAction($sSession,$element,"click")


sleep(300)

_WD_Attach($sSession,'People')

_WD_LoadWait($sSession)
check_mouse()
sleep(300)


$loc =""

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_0_260000001"]')
$loc= _WD_ElementAction($sSession, $element,'value')

$loc = StringLeft($loc,3)


Global $sid =""

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_0_1000000054"]')
Global $sid= _WD_ElementAction($sSession, $element,'value')




$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_0_1000000048"]')
Global $agent_sso= _WD_ElementAction($sSession, $element,'value')








;close button
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'/html/body/div[1]/div[5]/a[5]/div/div')
_WD_ElementAction($sSession,$element,"click")

 Return $loc


ConsoleWrite(@CRLF&$loc)


EndFunc




Func user_followup()



Search_old_ticket("SD")




if ((GUICtrlRead($Combo2) <> "Check") and (GUICtrlRead($Combo2) <> "Status")) Then






$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_3_1000000215"]')

$source = _WD_ElementAction($sSession, $element,'value')


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")

$status = _WD_ElementAction($sSession, $element,'value')

$loc =""

$loc = get_agent_details_from_ticket()



If ((GUICtrlRead($Combo2) == "Mail") or (GUICtrlRead($Combo2) == "Check") or (GUICtrlRead($Combo2) == "Wrong Assignment")) Then


$tl_email = Search_tl_details($sSession)
$tl_email = StringSplit($tl_email,"©")

EndIf


if ($tl_email[4] == "N") Then


_WD_Attach($sSession,'BMC Remedy (Search)')




	$element = _WD_GetElementById($sSession,"arid_WIN_3_1000000217")

	_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)
		_WD_ElementAction($sSession,$element,'value', "IT Services - Global - Virtual - T2")


sleep(555)

ajax_select_element($sSession,$element)
	sleep(500)


		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[3]/fieldset/div/div[2]/textarea")
			_WD_ElementAction($sSession,$element,'clear')

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[3]/fieldset/div/div[2]/textarea")
				_WD_ElementAction($sSession,$element,'value', "Agent Flag = N , Admin Staff"  );Comments



	sleep(555)





sleep(555)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue013")

sleep(555)

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[2]/td[1]")

			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

 			_WD_ElementAction($sSession,$element,"click")
;save button"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[1]/div/div"





;MsgBox("","","save")

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[1]/div/div")
_WD_ElementAction($sSession,$element,"click")







sleep(1000)


sleep(3555)










Else
	$n = 0
If (GUICtrlRead($Combo2) == "Wrong Assignment")  Then

	_ExtMsgBoxSet(1, 2, 0xFF8888, 0x004080, 14, "Comic Sans MS")
		_ExtMsgBoxSet(-1, 1, -1, -1, -1, -1, @DesktopWidth - 70)

		$n = _ExtMsgBox("", "Close|with Mail|Skip", "     SCRIPT PAUSED", "Do you want to Close ticket ,Close ticket with Email or Skip", "")
EndIf



if ($n <> 3) Then



If ((GUICtrlRead($Combo2) == "Status") or (GUICtrlRead($Combo2) == "Check")) Then





	$element = _WD_GetElementById($sSession, "arid_WIN_3_1000000218")
	$assigned = _WD_ElementAction($sSession, $element,'value')

		local $file = "list.txt"
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		$data = "--"& $status&"--"&$assigned&"--"&$SD

		FileWriteLine($hFileOpen, $data)
		; Close the handle returned by FileOpen.
		FileClose($hFileOpen)


else



if ($loc == "PHL") then

$reason = "user_PHIL"

$msg = "user needs to Joinmeeting on zoom"

Else



$reason = "user_IND"
$msg = " user needs to contact the Zoho Support via re.company.com"
EndIf


If (GUICtrlRead($Combo2) == "Wrong Assignment")  Then

$reason = "WRONG_ASSIGNMENT"
$msg = " Wrongly assigned to WAH team. Incase issue persists, please reach out to Global HD and get the tickets escalated to proper support groups"

EndIf


sleep(300)




if (($source == "Self Service")  or ($source == "Email") or ($loc == "PH") or (GUICtrlRead($Combo2) == "Wrong Assignment")) Then




if (($status <> "Resolved")  and ($status <> "Closed") ) Then








if (($status == "Assigned")  or ($status == "Pending")) Then


assign_ticket_to_me()

sleep(1000)


sleep(3555)






;-------------------------------search locatio-----------------------------------


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[1]/textarea")
$SD= _WD_ElementAction($sSession, $element,'value')

$issue=""
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")
$issue= _WD_ElementAction($sSession, $element,'value')




If ((GUICtrlRead($Combo2) == "Mail") or (GUICtrlRead($Combo2) == "Check") or (GUICtrlRead($Combo2) == "Wrong Assignment")) Then


;if $mail_launch == True Then








launch_mail($tl_email[1],$tl_email[2],$tl_email[3],$SD,$issue,$reason,$loc)
EndIf
sleep(1000)








;-------------------------------search locatio-----------------------------------





 $prevTkt = "NO"
close_im($msg)








		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/a[3]/div/div")
			_WD_ElementAction($sSession,$element,"click")


			check_mouse()
			_WD_LoadWait($sSession)

		sleep(5000)


		local $file = "list.txt"
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		$data = "--"& $status&"--"&"Closed"&"--"&GUICtrlRead($b)&"--"&$SD

		FileWriteLine($hFileOpen, $data)
		; Close the handle returned by FileOpen.
		FileClose($hFileOpen)
EndIf
EndIf
EndIf
EndIf

EndIf
EndIf


EndIf

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")

$status = _WD_ElementAction($sSession, $element,'value')


	$element = _WD_GetElementById($sSession, "arid_WIN_3_1000000218")
	$assigned = _WD_ElementAction($sSession, $element,'value')

		$element = _WD_GetElementById($sSession,"arid_WIN_3_1000000217")
		$assignment = _WD_ElementAction($sSession, $element,'value')

if ((GUICtrlRead($Combo2) <> "Check") and (GUICtrlRead($Combo2) <> "Status")) Then
		if (($tl_email[4] == "N") and ($status == "Assigned")) Then
			$status ="Reassigned"
		EndIf

EndIf

		local $file = "list.txt"
		Local $hFileOpen = FileOpen($file, $FO_APPEND)
		$data = "--"& $status&"--"&$assigned&"--"&$assignment&"--"&$SD

		FileWriteLine($hFileOpen, $data)
		; Close the handle returned by FileOpen.
		FileClose($hFileOpen)


EndFunc


Func assign_ticket_to_me()

_WD_Attach($sSession,'BMC Remedy (Search)')


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[1]/textarea"
$value = "Azure VDI"             ;"Azure VDI"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)










	$element = _WD_GetElementById($sSession,"arid_WIN_3_1000000217")

	_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)
		_WD_ElementAction($sSession,$element,'value', "Help Desk - Work at Home")


sleep(555)

ajax_select_element($sSession,$element)
	sleep(500)

	$element = _WD_GetElementById($sSession, "arid_WIN_3_1000000218")
		_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)


		_WD_ElementAction($sSession,$element,'value', StringLeft(StringSplit($sso,"@")[1],4))
	;	_WD_ElementAction($sSession,$element,'value', GUICtrlRead($b))
		;MsgBox("","",GUICtrlRead($b),2)
sleep(500)
		ajax_select_element($sSession,$element)





Category_select()






;status


sleep(555)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue013")

sleep(555)

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")

			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

 			_WD_ElementAction($sSession,$element,"click")
;save button"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[1]/div/div"





;MsgBox("","","save")

$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[1]/div/div")
_WD_ElementAction($sSession,$element,"click")

EndFunc


func Search_old_ticket($query = "SID")



if (GUICtrlRead($Combo1) <> "Followup") Then

				_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
				_WD_Navigate($sSession, "website")
				_WD_LoadWait($sSession)

EndIf

	_WD_LoadWait($sSession)
	check_mouse()
		_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
				check_mouse()




$path = '//*[@id="arid_WIN_1_80024"]'
 highlight_item($path)

$element =  _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
	$loc_a=_WD_ElementAction($sSession, $element, "Property", "outerHTML")
ConsoleWrite(@crlf&@crlf&"___Found element____"&StringStripWS($loc_a,8)&@crlf&@crlf)






$path = '//*[@id="Toolbar"]/table/tbody/tr/td[2]/a[1]/span'
 highlight_item($path)

	$element =  _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
	$loc_a=_WD_ElementAction($sSession, $element, "Property", "outerHTML")


sleep (500)

sleep (1500)
ConsoleWrite("___Found element____"&StringStripWS($loc_a,8))

if (StringStripWS($loc_a,8) == "<span>Newsearch</span>") and GUICtrlRead($Combo1) == "Followup"  Then
	$loc_a=""

	sleep(555)

$path = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[2]/a[1]/span"
 highlight_item($path)
			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
			_WD_ElementAction($sSession,$element,"focus");New search button
			_WD_ElementAction($sSession,$element,"click");New search button
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue006")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue007")
			sleep(555)


	sleep (1500)


Else



$loc_a=""



	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&4&"]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&3&"]"
	EndIf


		sleep(755)


		;search incident

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[5]/div/div[9]/div/div[3]/a")
		_WD_ElementAction($sSession,$element,"click")

		_WD_ExecuteScript($sSession,"CallARGHPD_58INC_58AppListEntryPoint_95Searchinmums_45remapplbEPFunc(false, this);")
	sleep(555)
	sleep(555)
;=================================================changed here======================================================V=================
		wait_for_element($sSession,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")

		sleep(555)



$element =  _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_1_80024"]')
	$loc_a=_WD_ElementAction($sSession, $element, "Property", "outerHTML")
ConsoleWrite(@crlf&@crlf&"___Found element____"&StringStripWS($loc_a,8)&@crlf&@crlf)

				;ConsoleWrite($sum_line)

						sleep(555)

;~ 				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[4]/textarea")
;~ 				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")

;~ 						sleep(755)
;~ 				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")
;~ 				_WD_ElementAction($sSession,$element,'value',$sum_line[3] );customer



sleep(2000)

EndIf

        if ($query == "SD") Then

		$xpath ="/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[1]/textarea"
		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)

		_WD_ElementAction($sSession,$element,'clear')
		 highlight_item($xpath)

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
		_WD_ElementAction($sSession,$element,'value',$SD );ticket







			sleep(555)
$xpath = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a/div"
 highlight_item($xpath)

			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementAction($sSession,$element,"focus");search button
			_WD_ElementAction($sSession,$element,"click");search button
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue006")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue007")
			sleep(555)


Else




$xpath = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea"
 highlight_item($xpath)

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)




$xpath = '//*[@id="arid_WIN_3_303530000"]'
 highlight_item($xpath)

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)


;check if input is >0 then its EID if ==0 then  SSO

If (number($sid) <> 0 ) Then
		_WD_ElementAction($sSession,$element,'value',$sum_line[3] );customer
Else

if $sid == "" Then
	MsgBox("","","USER EMAIL NOT FOUND COPY user email and click on OK")

	$sid = ClipGet()
EndIf


				 Global $agent_sso = $sid

				  $agent_sso = StringSplit($agent_sso,'@')
				  $agent_sso = $agent_sso[1]
ConsoleWrite("---------"&  $agent_sso&"-------"&@CRLF)
			ClipPut($agent_sso)
		 $agent_sso = StringSplit($agent_sso,".")




				  $agent_sso = ""&StringLeft($agent_sso[1],14)&"."&StringLeft($agent_sso[2],14)&""
					If @error then MsgBox(0,"Error",$agent_sso[1] & $agent_sso[2])
									  If @error then MsgBox(0,"Error","error has error")





		If $prevTkt == "NO" Then
			    $xpath ='//*[@id="arid_WIN_4_303530000"]'
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
				 highlight_item($xpath)
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
		_WD_ElementAction($sSession,$element,'value',$agent_sso );customer

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_3_303530000"]')
		_WD_ElementAction($sSession,$element,'value',$agent_sso );customer
		Else
		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
		_WD_ElementAction($sSession,$element,'value',$agent_sso );customer
		EndIf
		ConsoleWrite($prevTkt)



EndIf


		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/table[2]/tbody/tr/td/font")

		_WD_ElementActionEx($sSession, $element, "clickandhold")

	sleep(555)
			click_on_value($sSession,$element,"/html/body/table[2]/tbody/tr/td/font")




			ajax_select($sSession,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")




			wait_for_element($sSession,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a/div")

			sleep(555)
			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a/div")
			_WD_ElementAction($sSession,$element,"focus");search button
			_WD_ElementAction($sSession,$element,"click");search button
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a/div")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue006")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "/ue007")
			sleep(555)


EndIf
#CS 							_WD_Attach($sSession,'People Search')
;
; 							;search button  //*[@id='WIN_3_304248190']/div
; 									$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_3_304248190']/div")
; 									_WD_ElementAction($sSession,$element,"click")
;
; 									sleep(555)
;
;
;
; 									_WD_Attach($sSession,'People Search')
;
; 									;coporate id //*[@id="arid_WIN_0_1000000054"]
;
; 									sleep(555)
;
; 									$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='arid_WIN_0_1000000054']")
; 									_WD_ElementAction($sSession,$element,'value', $sid)
;
; 									;search button //*[@id="WIN_0_301867800"]/div
; 									$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_0_301867800']/div")
; 									ConsoleWrite($element&"_")
; 									_WD_ElementAction($sSession,$element,"click")
;
;
; 									sleep(2555)
;
; 									;_WD_Attach($sSession,'People Search')
; 									$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='T301394438']/tbody/tr[2]")
; 									 _WD_ExecuteScript($sSession,"(document.getElementsByTagName('tr')[8].classList.add('SelPrimary');")
;
; 									_WD_ElementAction($sSession,$element,"click")
;
;
;
; 									sleep(555)
 #CE







;latest removed

;~ 		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_0_304248710']/fieldset/div/dl/dd[3]/span[2]/a")
;~ 		_WD_ElementAction($sSession,$element,"click")


;~ 		sleep(555)


;latest removed








		;================================================search page for user ============================

;~ 		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_0_301912800']")
;~ 		_WD_ElementAction($sSession,$element,"click")



;~ 		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_0_301912800']")
;~ 		_WD_Attach($sSession,'BMC Remedy (Search)')
;~ 		_WD_ElementAction($sSession,$element,"click")

;~ 		sleep(555)
		;=============================================================New Tab============================================


;~ 		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[1]/a")
;~ 		_WD_ElementAction($sSession,$element,'click')
;~ 			sleep(555)

;~ 		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="TBsearchsavechanges"]')
;~ 		_WD_ElementAction($sSession,$element,'click')
;~ 		_WD_LinkClickByText($sSession,"Search")


EndFunc

func CCL($dummy,$actual,$value,$modi)
$write_val = StringTrimRight($value,1)

if $actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[1]/textarea" Then

	$local_time = 1000
Else
	$local_time =600
EndIf



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)
				$read_txt=_WD_ElementAction($sSession,$element,'text')
				$read_val=_WD_ElementAction($sSession,$element,'value')
while (($read_val <> $value) and ($read_txt <> $value))

			;~ 			$dummy=	FileReadLine ( "a.txt" ,1 )  dummy click
			;~ 			$actual=	FileReadLine ( "a.txt" ,2 )  actual Click
			;~ 			$value=	FileReadLine ( "a.txt" ,3 )  Value to put
			;~ 			$modi=	FileReadLine ( "a.txt" ,4 )  Click Modifier





			sleep(100)
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$dummy)

				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, $modi)   						 ;dummy click
			sleep(100)
					$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)
					_WD_ElementAction($sSession,$element,'clear')                   						            ;clear value
			sleep(100)
						$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)
						_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, $modi)					;modifier click
			sleep(100)
							$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)
							_WD_ElementAction($sSession,$element,'value',$write_val)										;put value
			sleep($local_time)
								$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/table[2]/tbody/tr/td/font")
								_WD_ElementAction($sSession,$element,"click")  											;Fix  value

			sleep(200)

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)
				$read_txt=_WD_ElementAction($sSession,$element,'text')
				$read_val=_WD_ElementAction($sSession,$element,'value')

				zoomin_out($actual)
;ConsoleWrite( $value & "____"& $read_val&"___"&$read_txt&@CRLF)


WEnd


EndFunc



Func assign_tickets()



	$element = _WD_GetElementById($sSession,"arid_WIN_3_303497300")
	_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)
		_WD_ElementAction($sSession,$element,'value', "wah")

sleep(500)
		ajax_select_element($sSession,$element)




	$element = _WD_GetElementById($sSession,"arid_WIN_3_1000000217")

	_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)
		_WD_ElementAction($sSession,$element,'value', "Help Desk - Work at Home")


sleep(555)

ajax_select_element($sSession,$element)
	sleep(500)

	$element = _WD_GetElementById($sSession, "arid_WIN_3_1000000218")
		_WD_ElementAction($sSession,$element,'clear')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\uE03D")
sleep(500)
		_WD_ElementAction($sSession,$element,'value', StringLeft($sso, 4))
sleep(500)
		ajax_select_element($sSession,$element)


		;arid_WIN_3_1000000151  , summary   arid_WIN_3_1000000000
		;arid_WIN_3_1000000217      =Help Desk - Work at Home
		;arid_WIN_3_1000000218    ===Ashish Saklani

EndFunc

Func zoomin_out($element)

		_WD_ExecuteScript($sSession,'document.body.style.zoom = "200%";')

		_WD_ExecuteScript($sSession, "arguments[0].scrollIntoView(true);", 'document.evaluate("'& $xpath & '", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue')



		sleep(500)

		_WD_ExecuteScript($sSession,'document.body.style.zoom = "100%";')



EndFunc

func highlight_item($element)

	_WD_ExecuteScript($sSession, 'document.evaluate("'& $xpath & '", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.style.backgroundColor = "#FDFF47";')
	_WD_ExecuteScript($sSession, 'document.evaluate("'& $xpath & '", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.style.outline = "#f00 solid 2px"')

EndFunc




Func New_ticket($sSession,$agent_sso="")

				_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)

	if  $prevTkt == "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&4&"]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&3&"]"
	EndIf


				;MsgBox("","",$agent_d)

				;---------------------------new ticket-----------------

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='WIN_0_304248710']/fieldset/div/dl/dd[3]/span[2]/a")
				_WD_ElementAction($sSession,$element,"click")


				;return to home



				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"//*[@id='FormContainer']/div[5]/div/div[9]/div/div[2]/a")
				_WD_ElementAction($sSession,$element,"click")
				_WD_ExecuteScript($sSession,"CallARGHPD_58INC_58AppListEntryPointinmums_45remapplbEPFunc(false, this);")
				;new ticket window


				sleep(500)

				;---------------------------new ticket-----------------


			wait_for_element($sSession,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")

		sleep(2555)

;-------------------------------------------------------------------------------------------------------

;~           	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")

;~ 			_WD_ElementAction($sSession,$element,'clear')
;~ 			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
;~ 			_WD_ElementAction($sSession,$element,'value',$sum_line[3] );customer

;~ 	sleep(555)


;~ 	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/a[1]/div")
;~ 	_WD_ElementAction($sSession,$element,'click') ; click on search to finalize fix customer name



;---------------------------------------------------------------------------------------





		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")
		;_WD_ExecuteScript($sSession, "arguments[0].scrollIntoView(true);", '{"' & $_WD_ELEMENT_ID & '":"' & $oElement2 & '"}')

		_WD_ElementAction($sSession,$element,'clear')


		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")



		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_4_303530000"]')
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_4_303530000"]')
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_4_303530000"]')


		$xpath = '//*[@id="arid_WIN_4_303530000"]'
		 highlight_item($xpath)

			if ($agent_sso <> "") Then
				_WD_ElementAction($sSession,$element,'value',$agent_sso );customer
			Else
				_WD_ElementAction($sSession,$element,'value',$sum_line[3] );customer
				$agent_sso = $sum_line[3]
			EndIf

			ClipPut($agent_sso)
			ConsoleWrite($agent_sso&@CRLF)


		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/table[2]/tbody/tr/td/font")

		_WD_ElementActionEx($sSession, $element, "clickandhold")

	sleep(555)
			click_on_value($sSession,$element,"/html/body/table[2]/tbody/tr/td/font")

			ajax_select($sSession,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")










;check if input is >0 then its EID if ==0 then  SSO



If ((number($sid) == 0 ) and $chat <> "OB")  Then


;~ 				  Global $agent_sso = $sid

;~ 				  $agent_sso = StringSplit($agent_sso,'@')
;~ 				  $agent_sso = $agent_sso[1]

;~ 				  $agent_sso = StringSplit($agent_sso,".")
;~ 				  $agent_sso = ""&StringLeft($agent_sso[1],14)&"."&StringLeft($agent_sso[2],10)&""


				ConsoleWrite($agent_sso&@CRLF)
;"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/a[2]/div"
;'//*[@id="WIN_4_303498000"]/div'

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="WIN_4_303498000"]/div')
				_WD_ElementAction($sSession,$element,"click")

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/a[2]/div")
				_WD_ElementAction($sSession,$element,"click")


				sleep(300)

				_WD_Attach($sSession,'People')

				_WD_LoadWait($sSession)
				check_mouse()
				sleep(400)

				$sid =""

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_0_1000000054"]')
				Global $sid= _WD_ElementAction($sSession, $element,'value')

				ConsoleWrite($sid)

				sleep(400)
				;close button
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'/html/body/div[1]/div[5]/a[5]/div/div')
				_WD_ElementAction($sSession,$element,"click")

				;ConsoleWrite("----i was here ----"&$sid&@CRLF)
				sleep(500)




if $qc = 1 then
	$sum_line = Search_qc_details_orig($sSession)
Else
	$sum_line = Search_qc_details($sSession)
EndIf

$sum_line = StringSplit($sum_line,"©")


EndIf





















;------------------------------------------------------------------------------------


EndFunc

func New_ticket_2($sSession)
	_WD_LoadWait($sSession)
	check_mouse()
	_WD_Attach($sSession,'BMC Remedy (Search)')
				sleep(555)
	_WD_LoadWait($sSession)
				check_mouse()
	if  $prevTkt == "YES" then
		$xpath = $path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
	Else
		$xpath = $path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
	EndIf


		 highlight_item($xpath)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$xpath)
		_WD_ElementAction($sSession,$element,'clear')
sleep(100)

		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
sleep(100)

	$system = system_type($sum_line[2])

	 check_issue()


	If $chat == "OB" Then

	_WD_ElementAction($sSession,$element,'value', $sum_line[1]&"__"&$class_name&"_"&$logs &@CRLF&@CRLF&"EmaiL      :  "&$sum_line[3]&"@company.com")

	Else

	_WD_ElementAction($sSession,$element,'value', $sum_line[1]&"__"&$sum_line[2]&"__"&$system&"_"&$logs &@CRLF&@CRLF&"Name       : "&$sum_line[4]&@CRLF&"Emp Id     :  "&$sid&@CRLF&"EmaiL      :  "&$sum_line[3]&"@company.com"&@CRLF&"Phone      : "&@CRLF&"Shift         : "&@CRLF&"WeekOff  : ")

	EndIf


	;notes  NOTES




				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[1]/textarea");customer details
				$agent_t= _WD_ElementAction($sSession, $element,'value')

				;ConsoleWrite($agent_t) ; AGENTS TICKET


	sleep(555)






		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")




				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")


	If $chat == "OB" Then

	_WD_ElementAction($sSession,$element,'value', $sum_line[1]&"__"&$class_name&"_"&$logs )

	Else

	_WD_ElementAction($sSession,$element,'value',  $sum_line[1]&"__"&$sum_line[2]&"__"&$system&"_"&$logs  );summary

	EndIf

	sleep(555)



$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[1]/textarea"
$value = $value4                 ;"Azure VDI"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)




	sleep(555)

If $prevTkt == "NO" Then


;impact

;$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[4]/div/input")
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_3_1000000163"]')



_WD_ElementAction($sSession,$element,"click")  ;Fix impackt
		sleep(555)
;impact select
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[4]/td[1]")

	_WD_ElementAction($sSession,$element,"click")  ;Fix impackt
			sleep(555)
Else



	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[4]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[4]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
	sleep(555)

				;Impact




;	if  $prevTkt == "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
;	Else
				;$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[3]/td[1]")
;				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[5]/td[1]")
;	EndIf

				_WD_ElementAction($sSession,$element,"click")  ;Fix impackt

	sleep(555)




EndIf

ConsoleWrite($prevTkt)




If $prevTkt == "NO" Then



;urgency


;$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,'//*[@id="arid_WIN_3_1000000162"]')
_WD_ElementAction($sSession,$element,"click")  ;Fix urgency
		sleep(555)
;urgency select
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[4]/td[1]")

	_WD_ElementAction($sSession,$element,"click")  ;Fix Urgency
			sleep(555)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea")

	_WD_ElementAction($sSession,$element,"click")  ;dummy



;urgency


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
_WD_ElementAction($sSession,$element,"click")  ;Fix urgency
		sleep(555)
;urgency select
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[4]/td[1]")

	_WD_ElementAction($sSession,$element,"click")  ;Fix Urgency
			sleep(555)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea")

	_WD_ElementAction($sSession,$element,"click")  ;dummy





Else


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
	sleep(555)
				;pripority




	;if  $prevTkt == "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
	;Else

	;			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[5]/td[1]")
	;EndIf
				_WD_ElementAction($sSession,$element,"click")  ;Fix iURGENCY




	sleep(555)
EndIf


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[1]/textarea"
$value ="Help Desk - Work at Home"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)

sleep(555)


ConsoleWrite($prevTkt)



If $prevTkt == "NO" Then




;status


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/a/img")
_WD_ElementAction($sSession,$element,"click")  ;Fix ;status
		sleep(555)
;status select


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[3]/td[1]")

	_WD_ElementAction($sSession,$element,"click")  ;Fix ;status select
			sleep(555)




;status

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

;status select


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[3]/td[1]")

	_WD_ElementAction($sSession,$element,"click")  ;Fix ;status select
			sleep(555)



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea")

	_WD_ElementAction($sSession,$element,"click")  ;dummy


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
	_WD_ElementAction($sSession,$element,"click")
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
	_WD_ElementAction($sSession,$element,"click")


Else










	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

	sleep(555)				;Status


;	if  $prevTkt == "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
;	Else

;				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
;	EndIf
					_WD_ElementAction($sSession,$element,"click")  ;Status

EndIf





$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[2]/textarea"
;~ $value ="Ashish Saklani"
;~ $modi ="\ue015"
;~ CCL($dummy,$actual,$value,$modi)
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$actual)

sleep(500)
		_WD_ElementAction($sSession,$element,'value', StringLeft($sso, 4))
sleep(500)
		ajax_select_element($sSession,$element)
sleep(555)
$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$dummy)
_WD_ElementAction($sSession,$element,"click")




If $prevTkt == "NO" Then



;reported source



$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[8]/div/input")
_WD_ElementAction($sSession,$element,"click")  ;Fix  source
		sleep(555)

;source select


						If $chat == "OMNI"  Then

						$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[4]/td[1]")

						Else

						$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[13]/td[1]")

						EndIf
							_WD_ElementAction($sSession,$element,"click")  ;Fix source select
									sleep(555)
Else





		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[8]/div/input")
			_WD_ElementAction($sSession,$element,"click")  ;Fix Reported source
	sleep(555)


				If ((GUICtrlRead($Combo1) == "Omni Chat") or ($chat == "OMNI" )) Then

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[4]/td[1]")

				Else

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[13]/td[1]")

				EndIf
				_WD_ElementAction($sSession,$element,"click")    ;Fix Reported source



	sleep(555)

EndIf








	sleep(555)


		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[3]/fieldset/div/div[2]/textarea")
			_WD_ElementAction($sSession,$element,'clear')

				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[1]/div[2]/div/div/div[3]/fieldset/div/div[2]/textarea")
				_WD_ElementAction($sSession,$element,'value', "User Came on Chat and reported issue " & @CRLF &$logs );Comments



	sleep(555)












Global $SD = $agent_t
	GUICtrlSetData($Label1, $SD)
 	GUICtrlRead($Label1)


EndFunc

Func no_responce()



;~ 				_WD_Attach($sSession,'BMC Remedy (Search)')
;~ 				sleep(555)

;~ 	if  $prevTkt = "YES" then
;~ 		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&4&"]"
;~ 	Else
;~ 		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&3&"]"
;~ 	EndIf





;~ $element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[3]/textarea")
;~ local $customer_name= _WD_ElementAction($sSession, $element,'value')

;~ 				;===============agent name=======================


;~ 	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[1]/fieldset/div[1]/textarea");customer details
;~ 				$agent_t= _WD_ElementAction($sSession, $element,'value')

;~ 				;===============agent ticket=======================


;~ 	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[3]/textarea")
;~ 		local $ticket_summary = _WD_ElementAction($sSession, $element,'value')

;~ MsgBox ("","customer name",$customer_name)
;~ MsgBox ("","customer ticket",$agent_t)
;~ MsgBox ("","ticket_summary",$ticket_summary)




;~ $element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[1]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[4]/div[10]")
;~ _WD_ElementAction($sSession,$element,"click")



;~ 	sleep(1505)



;~ 				_WD_LoadWait($sSession)
;~ 				;_WD_Attach($sSession,'Email System - Google Chrome')

;~ 				sleep(3555)



;~ 			$element = _WD_FindElement($sSession,$_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[8]/textarea")
;~ 				_WD_ElementAction($sSession,$element,"click")
;~ 					_WD_ElementAction($sSession,$element,'value',"hi this is ashish")



;~ 					_WD_FrameEnter($sSession,1)


;~ 			$element = _WD_FindElement($sSession,$_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[8]/textarea")
;~ 				_WD_ElementAction($sSession,$element,"click")
;~ 					_WD_ElementAction($sSession,$element,'value',"hi this is ashish1")



;~ 					_WD_FrameEnter($sSession,2)


;~ 			$element = _WD_FindElement($sSession,$_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[8]/textarea")
;~ 				_WD_ElementAction($sSession,$element,"click")
;~ 					_WD_ElementAction($sSession,$element,'value',"hi this is ashish2")













; 	; HTML
; 	;=====
 	Global $oOutlook = _OL_Open()
 	;FileInstall("\\comutername\as\as\first_followup.oft", @TempDir & "\first_followup.oft", 1)
 	FileInstall('C:\Users\%userprofle%\Documents\first_followup.oft', @TempDir & "\first_followup.oft", 1)
 	Sleep($d_time)
 	$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "", @TempDir & "\first_followup.oft", "SentOnBehalfOfName=ha@company.com", "to=" & $mail, "cc=Home Agent Support;" & $smail & ";" & $mmail)
 	Sleep($d_time)
 	FileDelete(@TempDir & "\first_followup.oft.oft")

 	$oItem.BodyFormat = $olFormatHTML
 	Sleep($d_time)
 	$oItem.GetInspector
 	Sleep($d_time)
 	$sBody = $oItem.HTMLBody
 	Sleep($d_time)
 	$oItem.HTMLBody = "" & $sBody
 	Sleep($d_time)
 	$oItem.Display


 	Opt("WinTitleMatchMode", 2)
 	Global $S1 = WinWait("IR#0000000", ""); s1 selected to match minmize option on tt1 ment for first case
 	WinActivate($S1)


 	start_setup()
 	check_date()
 	check_issue()

 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "XXXXXX", $fname & " ,")
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)

 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000", " " & $SD & " ")
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)

 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "<Engineer Name>", $usern & " ,")
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)

 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "<issue>", $logs)
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)


 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "000-000-0000", $stel & " ,")
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)

 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Body")
 	$PropertyUpd = StringReplace($Property[1][1], "00/00/2010", $date)
 	_OL_ItemModify($oOutlook, $oItem, Default, "Body=" & $PropertyUpd)




 	check_issue()
 	WinActivate($S1)
 	Global $Property = _OL_ItemGet($oOutlook, $oItem, "", "Subject")


 	$PropertyUpd = StringReplace($Property[1][1], "IR#0000000 -", $SD & " Closing Ticket " & $logs)
 	_OL_ItemModify($oOutlook, $oItem, Default, "Subject=" & $PropertyUpd)



$step = ""
EndFunc   ;==>no_responce


Func close_im($msg)
	HotKeySet("{ESC}", "esc")
	Opt("SendKeyDelay", 10)
	;hide()
	_WD_Attach($sSession,'BMC Remedy (Search)')

	if  $prevTkt == "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]"
	EndIf


;ConsoleWrite($prevTkt)


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&'/fieldset/div/div/div/div[4]/div[4]/div/div/div[1]/fieldset/div/div/div/div/div[3]/fieldset/div/div[10]/a')
	_WD_ElementAction($sSession,$element,"click")  ;Open menu


	sleep(1000)

	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[1]/td[2]")
	_WD_ElementAction($sSession,$element,"click")  ;Open menu

	sleep(300)


	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[5]/div[2]/table/tbody/tr[1]/td[1]")
	_WD_ElementAction($sSession,$element,"click")  ;Close Incident



sleep(1005)

_WD_Attach($sSession,'Incident Modification')

_WD_LoadWait($sSession)
sleep(555)


$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[5]/textarea")
_WD_ElementAction($sSession,$element,'clear')
_WD_ElementAction($sSession,$element,'value',$msg)




sleep(555)






	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[6]/a")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[1]/div[5]/div[6]/a")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
	sleep(555)
				;pripority
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[5]/td[1]")
				_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
				_WD_ElementAction($sSession,$element,"click")
ClipPut("Working fine after session reset")




	;show()
EndFunc   ;==>close_im


Func close_im_self_1()
	HotKeySet("{ESC}", "esc")
	Opt("SendKeyDelay", 10)
	hide()

	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]"
	EndIf




;"/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[2]/span[2]/a"


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[1]/textarea"
$value = "Client Application Service"                 ;"Azure VDI"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)




	sleep(555)













$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[4]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[4]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
	sleep(555)

				;Impact

	if  $prevTkt = "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
	Else
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[3]/div[2]/table/tbody/tr[4]/td[1]")
	EndIf

				_WD_ElementAction($sSession,$element,"click")  ;Fix impackt

	sleep(555)




	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[3]/fieldset/div[5]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")
	sleep(555)
				;pripority
	if  $prevTkt = "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
	Else
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[4]/td[1]")
	EndIf
				_WD_ElementAction($sSession,$element,"click")  ;Fix iURGENCY





	sleep(555)



$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[1]/textarea"
$value ="Help Desk - Work at Home"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)

sleep(555)





	$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementAction($sSession,$element,'clear')

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[5]/div/input")
			_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")

	sleep(555)				;Status


	if  $prevTkt = "YES" then
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[2]/td[1]")
	Else
				$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,"/html/body/div[4]/div[2]/table/tbody/tr[3]/td[1]")
	EndIf
					_WD_ElementAction($sSession,$element,"click")  ;Status




$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[4]/fieldset/div[2]/textarea"
$value ="Ashish Saklani"
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)

sleep(555)






	;show()
EndFunc   ;==>close_im_self



Func close_im_self()
	HotKeySet("{ESC}", "esc")
	Opt("SendKeyDelay", 10)
	hide()

	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]"
	EndIf

check_issue()

Category_select()

	;show()
EndFunc   ;==>close_im



Func DisplayMsg($text, $button,$msgwidth=300,$msgheight=80,$msgxpos=-1,$msgypos=-1)



;===========================================================================================
;
; Usage: _DisplayMsg($text,$button,[msgwidth,[msgheight,[msgxpos,[msgypos]]]])
;
; Where:
;     $text    =  The text to be displayed. Long lines will wrap. The text will also
;                    be centered.
;     $button    =  The text for the buttons. Seperate buttons with the
;                    pipe (|) character. To set a button as the default button,
;                    place the ampersand (&) character before the word. If you
;                    put more than 1 ampersand character, the function will
;                    fail and return -1, plus set @error to 1
;     msgwidth  =  Width of the displayed window. This value will be automatically
;                    increased if the resulting window will not be wide enough to
;                    accommodate the buttons requested. Default is 250
;     msgheight   =  Height of the displayed window. This is not adjusted. If the
;                    text is too big, it will not display it all. Default is 80
;     msgxpos    =  Where you want the window positioned horizontally. Default is centered.
;     msgypos    =  Where you want the window positioned vertically. Default is centered.
;
;     Success:  Returns the button pressed, starting at 1, counting from the LEFT.
;     Failure:  Returns 0 if more than 1 default button is set, or an error with the window
;                 occurs.
;
;===========================================================================================

    Local $buttonarray,$msgbutton[5]=[1,1,1,1,1],$buttoncount,$defbutton=0
    If StringInStr($button,"&",0,2)<>0 Then
        SetError(1)
        Return 0
    EndIf
    $buttonarray=StringSplit($button,"|")
    If $buttonarray[0]>5 Then $buttonarray[0]=5
    If 88*$buttonarray[0]+8>$msgwidth Then
        $msgwidth=88*$buttonarray[0]+8
    EndIf
    $msggui = GUICreate("", $msgwidth, $msgheight, $msgxpos, $msgypos, $WS_popup + $WS_DLGframe, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
    If $msggui=0 Then Return 0
    GUICtrlCreateLabel($text, 8, 8, $msgwidth-16, $msgheight-40, $SS_CENTER)
    $buttonxpos=(($msgwidth/$buttonarray[0])-80)/2
    For $buttoncount=0 To $buttonarray[0]-1
        $buttonwork=$buttonarray[$buttoncount+1]
        If StringLeft($buttonwork,1)="&" Then
            $defbutton=$BS_DEFPUSHBUTTON
            $buttonarray[$buttoncount+1]="[ " & StringTrimLeft($buttonwork,1) & " ]"
        EndIf
        $msgbutton[$buttoncount] = GUICtrlCreateButton($buttonarray[$buttoncount+1], $buttonxpos+($buttoncount*80)+($buttoncount*$buttonxpos*2), $msgheight-32, 80, 24,$defbutton)
        $defbutton=0
    Next
    GUISetState()
    While 1


	$Diff = TimerDiff($Timer)
    If $Diff >= 100000 Then
           $Timer = TimerInit()
		   mouse_move()
	EndIf

        $mmsg = GUIGetMsg()
        Select
            Case $mmsg = $msgbutton[0]
                GUIDelete($msggui)
                Return 1
            Case $mmsg = $msgbutton[1]
                GUIDelete($msggui)
                Return 2
            Case $mmsg = $msgbutton[2]
                GUIDelete($msggui)
                Return 3
            Case $mmsg = $msgbutton[3]
                GUIDelete($msggui)
                Return 4
            Case $mmsg = $msgbutton[4]
                GUIDelete($msggui)
                Return 5
        EndSelect
    WEnd
EndFunc



func Category_select()


_WD_Attach($sSession,'BMC Remedy (Search)')






		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[2]/span[2]/a")

		_WD_ElementActionEx($sSession, $element, "modifierclick", 0, 0,0, 2000, "\ue015")


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/div[3]/textarea"
$value =$cat1_tier1
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)

sleep(555)




$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/div[5]/textarea"
$value =$cat1_tier2
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)

sleep(555)

			$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[2]/fieldset/div/a[1]/div")

				_WD_ElementAction($sSession,$element,"click")  ;Next Page



;~  	HotKeySet("!d", "sfun1")
;~   			while 1

;~   					sleep(555)


;~  			WEnd



;~  	sleep(500)

$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[3]/textarea"
$value =$cat2_tier1
$modi ="\ue015"


CCL($dummy,$actual,$value,$modi)

sleep(300)

$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[5]/textarea"
$value =$cat2_tier2
$modi ="\ue015"

CCL($dummy,$actual,$value,$modi)

sleep(300)

$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[10]/textarea"
$value = $cat1_tier4
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)
sleep(300)


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[12]/textarea"
$value = $cat1_tier5
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)
sleep(300)


$dummy =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[2]/fieldset/div/div/fieldset/div[2]/fieldset/div[1]/textarea"
$actual =$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[2]/div[2]/div/div/div[3]/fieldset/div/div[14]/textarea"
$value = $cat1_tier6
$modi ="\ue015"
CCL($dummy,$actual,$value,$modi)
sleep(300)



		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[5]/span[2]/a")
		_WD_ElementAction($sSession,$element,"click")

		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/fieldset[5]/div[18]/fieldset/div/span/input")
		_WD_ElementAction($sSession,$element,"click")



		$element = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath,$path_for_tickt&"/fieldset/div/div/div/div[4]/div[4]/div/div/div[2]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div[2]/div[2]/div/dl/dd[1]/span[2]/a")
		_WD_ElementAction($sSession,$element,"click")






EndFunc


func mouse_move()


	;		while True

				MouseMove(Random(0,700,1),Random(0,700,1),500)
				;MouseDown($MOUSE_CLICK_LEFT)




	;			sleep(100000)


	;		WEnd
EndFunc



Func Error_Handle()

;~ 	if ((Hex($oErrors.number,8) <> '80020006' ) and (Hex($oErrors.number,8) <> '8001011B'))  then
;~     Msgbox(0, "Debug Information", "An error has occured!" & @CRLF & @CRLF & _
;~              "err.description is: " & @TAB & $oErrors.description   & @CRLF & _
;~              "err.windescription:"  & @TAB & $oErrors.windescription & @CRLF & _
;~              "err.number is: "  & @TAB & Hex($oErrors.number,8) & @CRLF & _
;~              "err.lastdllerror is: " & @TAB & $oErrors.lastdllerror & @CRLF & _
;~              "err.scriptline is: "  & @TAB & $oErrors.scriptline    & @CRLF & _
;~              "err.source is: "  & @TAB & $oErrors.source    & @CRLF & _
;~              "err.helpfile is: "    & @TAB & $oErrors.helpfile  & @CRLF & _
;~              "err.helpcontext is: " & @TAB & $oErrors.helpcontext   & @CRLF & @CRLF & @CRLF)
;~     Return SetError($oErrors.number)
;~ 	endif

#comments-start
Error Codes:
    Automation  Automation  Error Description
    Error       Error
    in Decimal  in Hex
    -2147467255 0x00084009  Unable to initialize class cache.
    -2147467254 0x0008400A  Unable to initialize RPC services.
    -2147483647 0x80000001  Not implemented.
    -2147483646 0x80000002  Ran out of memory.
    -2147483645 0x80000003  One or more arguments are invalid.
    -2147483644 0x80000004  No such interface supported.
    -2147483643 0x80000005  Invalid pointer.
    -2147483642 0x80000006  Invalid handle.
    -2147483641 0x80000007  Operation aborted.
    -2147483640 0x80000008  Unspecified error.
    -2147483639 0x80000009  General access denied error.
    -2147483638 0x8000000A  The data necessary to complete this operation not yet available.
    -2147467263 0x80004001  Not implemented.
    -2147467262 0x80004002  No such interface supported.
    -2147467261 0x80004003  Invalid pointer.
    -2147467260 0x80004004  Operation aborted.
    -2147467259 0x80004005  Unspecified error.
    -2147467258 0x80004006  Thread local storage failure.
    -2147467257 0x80004007  Get shared memory allocator failure.
    -2147467256 0x80004008  Get memory allocator failure.
    -2147467253 0x8000400B  Cannot set thread local storage channel control.
    -2147467252 0x8000400C  Could not allocate thread local storage channel control.
    -2147467251 0x8000400D  The user supplied memory allocator is unacceptable.
    -2147467250 0x8000400E  The OLE service mutex already exists.
    -2147467249 0x8000400F  The OLE service file mapping already exists.
    -2147467248 0x80004010  Unable to map view of file for OLE service.
    -2147467247 0x80004011  Failure attempting to launch OLE service.
    -2147467246 0x80004012  There was an attempt to call CoInitialize a second time while single threaded.
    -2147467245 0x80004013  A Remote activation was necessary but was not allowed.
    -2147467244 0x80004014  A Remote activation was necessary but the server name provided was invalid.
    -2147467243 0x80004015  The class is configured to run as a security id different from the caller.
    -2147467242 0x80004016  Use of Ole1 services requiring DDE windows is disabled.
    -2147467241 0x80004017  A RunAs specification must be A RunAs specification must be <domain name>\<user name> or simply <user name>.
    -2147467240 0x80004018  The server process could not be started. The pathname may be incorrect.
    -2147467239 0x80004019  The server process could not be started as the configured identity. The pathname may be incorrect or unavailable.
    -2147467238 0x8000401A  The server process could not be started because the configured identity is incorrect. Check the username and password.
    -2147467237 0x8000401B  The client is not allowed to launch this server.
    -2147467236 0x8000401C  The service providing this server could not be started.
    -2147467235 0x8000401D  This computer was unable to communicate with the computer providing the server.
    -2147467234 0x8000401E  The server did not respond after being launched.
    -2147467233 0x8000401F  The registration information for this server is inconsistent or incomplete.
    -2147467232 0x80004020  The registration information for this interface is inconsistent or incomplete.
    -2147467231 0x80004021  The operation attempted is not supported.
    -2147418113 0x8000FFFF  Catastrophic failure.
    -2147418111 0x80010001  Call was rejected by callee.
    -2147418110 0x80010002  Call was canceled by the message filter.
    -2147418109 0x80010003  The caller is dispatching an intertask SendMessage call and cannot call out via PostMessage.
    -2147418108 0x80010004  The caller is dispatching an asynchronous call and cannot make an outgoing call on behalf of this call.
    -2147418107 0x80010005  It is illegal to call out while inside message filter.
    -2147418106 0x80010006  The connection terminated or is in a bogus state and cannot be used any more. Other connections are still valid.
    -2147418105 0x80010007  The callee (server [not server application]) is not available and disappeared; all connections are invalid. The call may have executed.
    -2147418104 0x80010008  The caller (client) disappeared while the callee (server) was processing a call.
    -2147418103 0x80010009  The data packet with the marshalled parameter data is incorrect.
    -2147418102 0x8001000A  The call was not transmitted properly; the message queue was full and was not emptied after yielding.
    -2147418101 0x8001000B  The client (caller) cannot marshal the parameter data - low memory, etc.
    -2147418100 0x8001000C  The client (caller) cannot unmarshal the return data - low memory, etc.
    -2147418099 0x8001000D  The server (callee) cannot marshal the return data - low memory, etc.
    -2147418098 0x8001000E  The server (callee) cannot unmarshal the parameter data - low memory, etc.
    -2147418097 0x8001000F  Received data is invalid; could be server or client data.
    -2147418096 0x80010010  A particular parameter is invalid and cannot be (un)marshalled.
    -2147418095 0x80010011  There is no second outgoing call on same channel in DDE conversation.
    -2147418094 0x80010012  The callee (server [not server application]) is not available and disappeared; all connections are invalid. The call did not execute.
    -2147417856 0x80010100  System call failed.
    -2147417855 0x80010101  Could not allocate some required resource (memory, events, ...)
    -2147417854 0x80010102  Attempted to make calls on more than one thread in single threaded mode.
    -2147417853 0x80010103  The requested interface is not registered on the server object.
    -2147417852 0x80010104  RPC could not call the server or could not return the results of calling the server.
    -2147417851 0x80010105  The server threw an exception.
    -2147417850 0x80010106  Cannot change thread mode after it is set.
    -2147417849 0x80010107  The method called does not exist on the server.
    -2147417848 0x80010108  The object invoked has disconnected from its clients.
    -2147417847 0x80010109  The object invoked chose not to process the call now. Try again later.
    -2147417846 0x8001010A  The message filter indicated that the application is busy.
    -2147417845 0x8001010B  The message filter rejected the call.
    -2147417844 0x8001010C  A call control interfaces was called with invalid data.
    -2147417843 0x8001010D  An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    -2147417842 0x8001010E  The application called an interface that was marshalled for a different thread.
    -2147417841 0x8001010F  CoInitialize has not been called on the current thread.
    -2147417840 0x80010110  The version of OLE on the client and server machines does not match.
    -2147417839 0x80010111  OLE received a packet with an invalid header.
    -2147417838 0x80010112  OLE received a packet with an invalid extension.
    -2147417837 0x80010113  The requested object or interface does not exist.
    -2147417836 0x80010114  The requested object does not exist.
    -2147417835 0x80010115  OLE has sent a request and is waiting for a reply.
    -2147417834 0x80010116  OLE is waiting before retrying a request.
    -2147417833 0x80010117  Call context cannot be accessed after call completed.
    -2147417832 0x80010118  Impersonate on unsecured calls is not supported.
    -2147417831 0x80010119  Security must be initialized before any interfaces are marshalled or unmarshalled. It cannot be changed once initialized.
    -2147417830 0x8001011A  No security packages are installed on this machine or the user is not logged on or there are no compatible security packages between the client and server.
    -2147417829 0x8001011B  Access is denied.
    -2147417828 0x8001011C  Remote calls are not allowed for this process.
    -2147417827 0x8001011D  The marshalled interface data packet (oBJREF) has an invalid or unknown format.
    -2147352577 0x8001FFFF  An internal error occurred.
    -2147352575 0x80020001  Unknown interface.
    -2147352573 0x80020003  Member not found.
    -2147352572 0x80020004  Parameter not found.
    -2147352571 0x80020005  Type mismatch.
    -2147352570 0x80020006  Unknown name.
    -2147352569 0x80020007  No named arguments.
    -2147352568 0x80020008  Bad variable type.
    -2147352567 0x80020009  Exception occurred.
    -2147352566 0x8002000A  Out of present range.
    -2147352565 0x8002000B  Invalid index.
    -2147352564 0x8002000C  Unknown language.
    -2147352563 0x8002000D  Memory is locked.
    -2147352562 0x8002000E  Invalid number of parameters.
    -2147352561 0x8002000F  Parameter not optional.
    -2147352560 0x80020010  Invalid callee.
    -2147352559 0x80020011  Does not support a collection.
    -2147319786 0x80028016  Buffer too small.
    -2147319784 0x80028018  Old format or invalid type library.
    -2147319783 0x80028019  Old format or invalid type library.
    -2147319780 0x8002801C  Error accessing the OLE registry.
    -2147319779 0x8002801D  Library not registered.
    -2147319769 0x80028027  Bound to unknown type.
    -2147319768 0x80028028  Qualified name disallowed.
    -2147319767 0x80028029  Invalid forward reference, or reference to uncompiled type.
    -2147319766 0x8002802A  Type mismatch.
    -2147319765 0x8002802B  Element not found.
    -2147319764 0x8002802C  Ambiguous name.
    -2147319763 0x8002802D  Name already exists in the library.
    -2147319762 0x8002802E  Unknown LCID.
    -2147319761 0x8002802F  Function not defined in specified DLL.
    -2147317571 0x800288BD  Wrong module kind for the operation.
    -2147317563 0x800288C5  Size may not exceed 64K.
    -2147317562 0x800288C6  Duplicate ID in inheritance hierarchy.
    -2147317553 0x800288CF  Incorrect inheritance depth in standard OLE hmember.
    -2147316576 0x80028CA0  Type mismatch.
    -2147316575 0x80028CA1  Invalid number of arguments.
    -2147316574 0x80028CA2  I/O Error.
    -2147316573 0x80028CA3  Error creating unique tmp file.
    -2147312566 0x80029C4A  Error loading type library/DLL.
    -2147312509 0x80029C83  Inconsistent property functions.
    -2147312508 0x80029C84  Circular dependency between types/modules.
    -2147287039 0x80030001  Unable to perform requested operation.
    -2147287038 0x80030002  %1 could not be found.
    -2147287037 0x80030003  The path %1 could not be found.
    -2147287036 0x80030004  There are insufficient resources to open another file.
    -2147287035 0x80030005  Access Denied.
    -2147287034 0x80030006  Attempted an operation on an invalid object.
    -2147287032 0x80030008  There is insufficient memory available to complete operation.
    -2147287031 0x80030009  Invalid pointer error.
    -2147287022 0x80030012  There are no more entries to return.
    -2147287021 0x80030013  Disk is write-protected.
    -2147287015 0x80030019  An error occurred during a seek operation.
    -2147287011 0x8003001D  A disk error occurred during a write operation.
    -2147287010 0x8003001E  A disk error occurred during a read operation.
    -2147287008 0x80030020  A share violation has occurred.
    -2147287007 0x80030021  A lock violation has occurred.
    -2147286960 0x80030050  %1 already exists.
    -2147286953 0x80030057  Invalid parameter error.
    -2147286928 0x80030070  There is insufficient disk space to complete operation.
    -2147286800 0x800300F0  Illegal write of non-simple property to simple property set.
    -2147286790 0x800300FA  An API call exited abnormally.
    -2147286789 0x800300FB  The file %1 is not a valid compound file.
    -2147286788 0x800300FC  The name %1 is not valid.
    -2147286787 0x800300FD  An unexpected error occurred.
    -2147286786 0x800300FE  That function is not implemented.
    -2147286785 0x800300FF  Invalid flag error.
    -2147286784 0x80030100  Attempted to use an object that is busy.
    -2147286783 0x80030101  The storage has been changed since the last commit.
    -2147286782 0x80030102  Attempted to use an object that has ceased to exist.
    -2147286781 0x80030103  Can't save.
    -2147286780 0x80030104  The compound file %1 was produced with an incompatible version of storage.
    -2147286779 0x80030105  The compound file %1 was produced with a newer version of storage.
    -2147286778 0x80030106  Share.exe or equivalent is required for operation.
    -2147286777 0x80030107  Illegal operation called on non-file based storage.
    -2147286776 0x80030108  Illegal operation called on object with extant marshallings.
    -2147286775 0x80030109  The docfile has been corrupted.
    -2147286768 0x80030110  OLE32.DLL has been loaded at the wrong address.
    -2147286527 0x80030201  The file download was aborted abnormally. The file is incomplete.
    -2147286526 0x80030202  The file download has been terminated.
    -2147221504 0x80040000  Invalid OLEVERB structure.
    -2147221503 0x80040001  Invalid advise flags.
    -2147221502 0x80040002  Can't enumerate anymore, because the associated data is missing.
    -2147221501 0x80040003  This implementation doesn't take advises.
    -2147221500 0x80040004  There is no connection for this connection ID.
    -2147221499 0x80040005  Need to run the object to perform this operation.
    -2147221498 0x80040006  There is no cache to operate on.
    -2147221497 0x80040007  Uninitialized object.
    -2147221496 0x80040008  Linked object's source class has changed.
    -2147221495 0x80040009  Not able to get the moniker of the object.
    -2147221494 0x8004000A  Not able to bind to the source.
    -2147221493 0x8004000B  Object is static; operation not allowed.
    -2147221492 0x8004000C  User cancelled out of save dialog.
    -2147221491 0x8004000D  Invalid rectangle.
    -2147221490 0x8004000E  compobj.dll is too old for the ole2.dll initialized.
    -2147221489 0x8004000F  Invalid window handle.
    -2147221488 0x80040010  Object is not in any of the inplace active states.
    -2147221487 0x80040011  Not able to convert object.
    -2147221486 0x80040012  Not able to perform the operation because object is not given storage yet.
    -2147221404 0x80040064  Invalid FORMATETC structure.
    -2147221403 0x80040065  Invalid DVTARGETDEVICE structure.
    -2147221402 0x80040066  Invalid STDGMEDIUM structure.
    -2147221401 0x80040067  Invalid STATDATA structure.
    -2147221400 0x80040068  Invalid lindex.
    -2147221399 0x80040069  Invalid tymed.
    -2147221398 0x8004006A  Invalid clipboard format.
    -2147221397 0x8004006B  Invalid aspect(s).
    -2147221396 0x8004006C  tdSize parameter of the DVTARGETDEVICE structure is invalid.
    -2147221395 0x8004006D  Object doesn't support IViewObject interface.
    -2147221248 0x80040100  Trying to revoke a drop target that has not been registered.
    -2147221247 0x80040101  This window has already been registered as a drop target.
    -2147221246 0x80040102  Invalid window handle.
    -2147221232 0x80040110  Class does not support aggregation (or class object is remote).
    -2147221231 0x80040111  ClassFactory cannot supply requested class.
    -2147221184 0x80040140  Error drawing view.
    -2147221168 0x80040150  Could not read key from registry.
    -2147221167 0x80040151  Could not write key to registry.
    -2147221166 0x80040152  Could not find the key in the registry.
    -2147221165 0x80040153  Invalid value for registry.
    -2147221164 0x80040154  Class not registered.
    -2147221163 0x80040155  Interface not registered.
    -2147221136 0x80040170  Cache not updated.
    -2147221120 0x80040180  No verbs for OLE object.
    -2147221119 0x80040181  Invalid verb for OLE object.
    -2147221088 0x800401A0  Undo is not available.
    -2147221087 0x800401A1  Space for tools is not available.
    -2147221056 0x800401C0  OLESTREAM Get method failed.
    -2147221055 0x800401C1  OLESTREAM Put method failed.
    -2147221054 0x800401C2  Contents of the OLESTREAM not in correct format.
    -2147221053 0x800401C3  There was an error in a Windows GDI call while converting the bitmap to a DIB.
    -2147221052 0x800401C4  Contents of the IStorage not in correct format.
    -2147221051 0x800401C5  Contents of IStorage is missing one of the standard streams.
    -2147221050 0x800401C6  There was an error in a Windows GDI call while converting the DIB to a bitmap.
    -2147221040 0x800401D0  OpenClipboard Failed.
    -2147221039 0x800401D1  EmptyClipboard Failed.
    -2147221038 0x800401D2  SetClipboard Failed.
    -2147221037 0x800401D3  Data on clipboard is invalid.
    -2147221036 0x800401D4  CloseClipboard Failed.
    -2147221024 0x800401E0  Moniker needs to be connected manually.
    -2147221023 0x800401E1  Operation exceeded deadline.
    -2147221022 0x800401E2  Moniker needs to be generic.
    -2147221021 0x800401E3  Operation unavailable.
    -2147221020 0x800401E4  Invalid syntax.
    -2147221019 0x800401E5  No object for moniker.
    -2147221018 0x800401E6  Bad extension for file.
    -2147221017 0x800401E7  Intermediate operation failed.
    -2147221016 0x800401E8  Moniker is not bindable.
    -2147221015 0x800401E9  Moniker is not bound.
    -2147221014 0x800401EA  Moniker cannot open file.
    -2147221013 0x800401EB  User input required for operation to succeed.
    -2147221012 0x800401EC  Moniker class has no inverse.
    -2147221011 0x800401ED  Moniker does not refer to storage.
    -2147221010 0x800401EE  No common prefix.
    -2147221009 0x800401EF  Moniker could not be enumerated.
    -2147221008 0x800401F0  CoInitialize has not been called.
    -2147221007 0x800401F1  CoInitialize has already been called.
    -2147221006 0x800401F2  Class of object cannot be determined.
    -2147221005 0x800401F3  Invalid class string.
    -2147221004 0x800401F4  Invalid interface string.
    -2147221003 0x800401F5  Application not found.
    -2147221002 0x800401F6  Application cannot be run more than once.
    -2147221001 0x800401F7  Some error in application program.
    -2147221000 0x800401F8  DLL for class not found.
    -2147220999 0x800401F9  Error in the DLL.
    -2147220998 0x800401FA  Wrong OS or OS version for application.
    -2147220997 0x800401FB  Object is not registered.
    -2147220996 0x800401FC  Object is already registered.
    -2147220995 0x800401FD  Object is not connected to server.
    -2147220994 0x800401FE  Application was launched but it didn't register a class factory.
    -2147220993 0x800401FF  Object has been released.
                0x80041001  Failed
                0x80041002  Not Found
                0x80041003  Access Denied (user rights)
                0x80041004  Provider Failure
                0x80041005  Type Mismatch
                0x80041006  Out Of Memory
                0x80041007  Invalid Context
                0x80041008  Invalid Parameter
                0x80041009  Not Available
                0x8004100A  Critical Error
                0x8004100B  Invalid Stream
                0x8004100C  Not Supported
                0x8004100D  Invalid Superclass
                0x8004100E  Invalid Namespace
                0x8004100F  Invalid Object
                0x80041010  Invalid Class
                0x80041011  Provider Not Found
                0x80041012  Invalid Provider Registration
                0x80041013  Provider Load Failure
                0x80041014  Initialization Failure
                0x80041015  Transport Failure
                0x80041016  Invalid Operation
                0x80041017  Invalid Query
                0x80041018  Invalid Query Type
                0x80041019  Already Exists
                0x8004101A  Override Not Allowed
                0x8004101B  Propagated Qualifier
                0x8004101C  Propagated Property
                0x8004101D  Unexpected
                0x8004101E  Illegal Operation
                0x8004101F  Cannot Be Key
                0x80041020  Incomplete Class
                0x80041021  Invalid Syntax
                0x80041022  Nondecorated Object
                0x80041023  Read Only
                0x80041024  Provider Not Capable
                0x80041025  Class Has Children
                0x80041026  Class Has Instances
                0x80041027  Query Not Implemented
                0x80041028  Illegal Null
                0x80041029  Invalid Qualifier Type
                0x8004102A  Invalid Property Type
                0x8004102B  Value Out Of Range
                0x8004102C  Cannot Be Singleton
                0x8004102D  Invalid Cim Type
                0x8004102E  Invalid Method
                0x8004102F  Invalid Method Parameters
                0x80041030  System Property
                0x80041031  Invalid Property
                0x80041032  Call Cancelled
                0x80041033  Shutting Down
                0x80041034  Propagated Method
                0x80041035  Unsupported Parameter
                0x80041036  Missing Parameter Id
                0x80041037  Invalid Parameter Id
                0x80041038  Nonconsecutive Parameter Ids
                0x80041039  Parameter Id On Retval
                0x8004103A  Invalid Object Path
                0x8004103B  Out Of Disk Space
                0x8004103C  Buffer Too Small
                0x8004103D  Unsupported Put Extension
                0x8004103E  Unknown Object Type
                0x8004103F  Unknown Packet Type
                0x80041040  Marshal Version Mismatch
                0x80041041  Marshal Invalid Signature
                0x80041042  Invalid Qualifier
                0x80041043  Invalid Duplicate Parameter
                0x80041044  Too Much Data
                0x80041045  Server Too Busy
                0x80041046  Invalid Flavor
                0x80041047  Circular Reference
                0x80041048  Unsupported Class Update
                0x80041049  Cannot Change Key Inheritance
                0x80041050  Cannot Change Index Inheritance
                0x80041051  Too Many Properties
                0x80041052  Update Type Mismatch
                0x80041053  Update Override Not Allowed
                0x80041054  Update Propagated Method
                0x80041055  Method Not Implemented
                0x80041056  Method Disabled
                0x80041064  User credentials cannot be used for local connections
                0x8004106C  WMI is taking up too much memory
                0x80042001  Wbemess E Registration Too Broad
                0x80042002  Wbemess E Registration Too Precise
    -2147024891 0x80070005  General access denied error (incorrect login).
    -2147024890 0x80070006  Invalid handle.
    -2147942413 0x8007000D  The Data is invalid.
    -2147024882 0x8007000E  Ran out of memory.
    -2147024809 0x80070057  One or more arguments are invalid.
                0x800706BA  The RPC server is unavailable
    -2146959359 0x80080001  Attempt to create a class object failed.
    -2146959358 0x80080002  OLE service could not bind object.
    -2146959357 0x80080003  RPC communication failed with OLE service.
    -2146959356 0x80080004  Bad path to object.
    -2146959355 0x80080005  Internal execution failure in the WMI Service.
    -2146959354 0x80080006  OLE service could not communicate with the object server.
    -2146959353 0x80080007  Moniker path could not be normalized.
    -2146959352 0x80080008  Object server is stopping when OLE service contacts it.
    -2146959351 0x80080009  An invalid root block pointer was specified.
    -2146959344 0x80080010  An allocation chain contained an invalid link pointer.
    -2146959343 0x80080011  The requested allocation size was too large.
    -2146893823 0x80090001  Bad UID.
    -2146893822 0x80090002  Bad Hash.
    -2146893821 0x80090003  Bad Key.
    -2146893820 0x80090004  Bad Length.
    -2146893819 0x80090005  Bad Data.
    -2146893818 0x80090006  Invalid Signature.
    -2146893817 0x80090007  Bad Version of provider.
    -2146893816 0x80090008  Invalid algorithm specified.
    -2146893815 0x80090009  Invalid flags specified.
    -2146893814 0x8009000A  Invalid type specified.
    -2146893813 0x8009000B  Key not valid for use in specified state.
    -2146893812 0x8009000C  Hash not valid for use in specified state.
    -2146893811 0x8009000D  Key does not exist.
    -2146893810 0x8009000E  Insufficient memory available for the operation.
    -2146893809 0x8009000F  Object already exists.
    -2146893808 0x80090010  Access denied.
    -2146893807 0x80090011  Object was not found.
    -2146893806 0x80090012  Data already encrypted.
    -2146893805 0x80090013  Invalid provider specified.
    -2146893804 0x80090014  Invalid provider type specified.
    -2146893803 0x80090015  Provider's public key is invalid.
    -2146893802 0x80090016  Keyset does not exist.
    -2146893801 0x80090017  Provider type not defined.
    -2146893800 0x80090018  Provider type as registered is invalid.
    -2146893799 0x80090019  The keyset is not defined.
    -2146893798 0x8009001A  Keyset as registered is invalid.
    -2146893797 0x8009001B  Provider type does not match registered value.
    -2146893796 0x8009001C  The digital signature file is corrupt.
    -2146893795 0x8009001D  Provider DLL failed to initialize correctly.
    -2146893794 0x8009001E  Provider DLL could not be found.
    -2146893793 0x8009001F  The Keyset parameter is invalid.
    -2146893792 0x80090020  An internal error occurred.
    -2146893791 0x80090021  A base error occurred.
    -2146762751 0x800B0001  The specified trust provider is not known on this system.
    -2146762750 0x800B0002  The trust verification action specified is not supported by the specified trust provider.
    -2146762749 0x800B0003  The form specified for the subject is not one supported or known by the specified trust provider.
    -2146762748 0x800B0004  The subject is not trusted for the specified action.
    -2146762747 0x800B0005  Error due to problem in ASN.1 encoding process.
    -2146762746 0x800B0006  Error due to problem in ASN.1 decoding process.
    -2146762745 0x800B0007  Reading / writing Extensions where Attributes are appropriate, and visa versa.
    -2146762744 0x800B0008  Unspecified cryptographic failure.
    -2146762743 0x800B0009  The size of the data could not be determined.
    -2146762742 0x800B000A  The size of the indefinite-sized data could not be determined.
    -2146762741 0x800B000B  This object does not read and write self-sizing data.
    -2146762496 0x800B0100  No signature was present in the subject.
    -2146762495 0x800B0101  A required certificate is not within its validity period.
    -2146762494 0x800B0102  The validity periods of the certification chain do not nest correctly.
    -2146762493 0x800B0103  A certificate that can only be used as an end-entity is being used as a CA or visa versa.
    -2146762492 0x800B0104  A path length constraint in the certification chain has been violated.
    -2146762491 0x800B0105  An extension of unknown type that is labeled 'critical' is present in a certificate.
    -2146762490 0x800B0106  A certificate is being used for a purpose other than that for which it is permitted.
    -2146762489 0x800B0107  A parent of a given certificate in fact did not issue that child certificate.
    -2146762488 0x800B0108  A certificate is missing or has an empty value for an important field, such as a subject or issuer name.
    -2146762487 0x800B0109  A certification chain processed correctly, but terminated in a root certificate which isn't trusted by the trust provider.
    -2146762486 0x800B010A  A chain of certs didn't chain as they should in a certain application of chaining.

#comments-end

EndFunc


Func Old_tkt_status()
	GUICtrlDelete($a)

	if  $prevTkt = "YES" then
		Global $prevTkt = "NO"
	Else
		Global $prevTkt = "YES"
	EndIf

	$a= GUICtrlCreateButton($prevTkt, 120, 0, 53, 20);140, 0, 54, 20)
	GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
	GUICtrlSetOnEvent($a, "Old_tkt_status")
	GUIRegisterMsg(0x0111, "_LostFocus")

	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[4]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]"
	EndIf

EndFunc

Func Old_tkt_status_b()

	if  $prevTkt = "YES" then
		Global $prevTkt = "NO"
		Global $color = "0xFF0000"



	Else
		Global $prevTkt = "YES"
		Global $color = "0x00FF00"
	EndIf

	GUICtrlSetData($Button27, "Old Tkt:" & $prevTkt)
		GUICtrlSetBkColor($Button27, $color)
 	GUICtrlRead($Button27)
	if  $prevTkt = "YES" then
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&4&"]"
	Else
		Global $path_for_tickt = "/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div["&3&"]"
	EndIf

EndFunc