#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#Include <GuiToolBar.au3>
#include <Process.au3>
;#RequireAdmin

Global $fileLog
Global $ffmpegTitle = "record screen use ffmpeg"

ikChangeToEnglish()
;调整窗口大小为 x,y，使得窗口大小固定，界面中的相对位置不变。
Func ikWinSetSize( $title, $x, $y)
   Local $aPos = WinGetPos($title)
    ; Move the Notepad window back to the original position by using the array returned by WinGetPos.
    WinMove($title, "", $aPos[0], $aPos[1], $x, $y) ;
EndFunc

; 窗口最大化,使得窗口大小固定，界面中的相对位置不变。
Func ikWinMaxsize($title)
	WinSetState($title, "", @SW_MAXIMIZE)
EndFunc

;ikWinForceClose returns nothing
;WhatItDoes:强制关闭窗口
Func ikWinForceClose($title)
	WinKill ($title)
EndFunc

;WhatItDoes:鼠标相对位置点击左键
;title需要用Autoit Window Info获取
;x,y 为相对位置，需要将Autoit Window Info的Options->Coord Mode->Window选中。
Func ikMouseRelativeClick($title, $x, $y, $count = 1)
   WinActivate($title)
   WinWaitActive($title)
   $pos = WinGetPos($title)
   MouseClick("primary", $pos[0] + $x, $pos[1] + $y, $count)
EndFunc

; 切换为英文输入法
Func ikChangeToEnglish()
	$hWnd = WinGetHandle("[ACTIVE]");
	$ret = DllCall("user32.dll", "long", "LoadKeyboardLayout", "str", "08040804", "int", 1 + 0)
	DllCall("user32.dll", "ptr", "SendMessage", "hwnd", $hWnd, "int", 0x50, "int", 1, "int", $ret[0])
EndFunc

; 从directory目录找文件file
Func ikFileSearch($StartDir, $SearchFile)
    Local $Search, $RFString = "File not found"
    $Search = FileFindFirstFile($StartDir & "\*.*")
    If @error Then Return $RFString
    ;Search through all files and folders in directory
    While $RFString = "File not found"
        $Next = FileFindNextFile($Search)
        If @error Then ExitLoop
        ;If folder, recurse
        If StringInStr(FileGetAttrib($StartDir & "\" & $Next), "D") Then
            $RFString = ikFileSearch($SearchFile, $StartDir & "\" & $Next)
        Else
            If $Next = $SearchFile Then $RFString = $StartDir & "\" & $Next
        EndIf
    WEnd
    FileClose($Search)
    Return $RFString
EndFunc

; 可以跟绝对路径,也可以跟windows普通命令，也可用RunLink跟快捷方式
; 大小写敏感
Func ikRunLink($file)
	$filelink = $file & ".lnk"
	$fullName = ikFileSearch("C:\ProgramData\Microsoft\Windows\Start Menu\Programs", $filelink)
	if $fullName == "File not found" Then
		$fullName = ikFileSearch("C:\Users\" & @UserName & "\AppData\Roaming\Microsoft\Windows\Start Menu", $filelink)
	Endif
	if $fullName <> "File not found" Then
		ShellExecute($fullName)
	Else
		Run($file)
	Endif
EndFunc
Func ikWiresharkStart()
	ikRunLink("Wireshark")
	ikSleep(1)
	ikWinMaxsize("The Wireshark Network Analyzer"); 最大化Wireshark窗口，使得窗口大小固定，界面中的相对位置不变。
	ikSleep(10) ; 延迟10秒，等待Wireshark初始化完成
	 ; The Wireshark Network Analyzer需要用Autoit Window Info获取
	 ; 318,837位置为以太网卡的位置。 ，Autoit Window Info需要将Options->Coord Mode->Window选中。
	 ikMouseRelativeClick("The Wireshark Network Analyzer", 388, 369,2); Wireshark.exe
EndFunc

; wireshark 保存退出
Func ikWiresharkSaveQuit($file)
	WinActivate("Capturing from")
	WinClose("Capturing from")
	ikSendEnter()
	ikSleep(1)
	ikSend($file)
	ikSleep(1)
	ikSendEnter()
EndFunc

Func ikGetSystray()
	; Find systray handle
    $hWnd = ControlGetHandle('[Class:Shell_TrayWnd]', '', '[Class:ToolbarWindow32;Text:User Promoted Notification Area]')
    If @error Then
        MsgBox(16, "Error", "System tray not found")
        Exit
    EndIf
	Return $hWnd
EndFunc

; 获取系统托盘index
Func ikGetSystrayIndex($hWnd, $tip)
    ; Find systray handle
    $hWnd = ikGetSystray()
    ; Get systray item count
    Local $iSystray_ButCount = _GUICtrlToolbar_ButtonCount($hWnd)
    If $iSystray_ButCount == 0 Then
        Return -1
    EndIf

    ; Look for wanted tooltip
    For $index = 0 To $iSystray_ButCount - 1
        If StringInStr(_GUICtrlToolbar_GetButtonText($hWnd, $index), $tip) = 1 Then ExitLoop
    Next

    If $index = $iSystray_ButCount Then
        Return -1 ; Not found
    Else
        Return $index ; Found
    EndIf
EndFunc

; 点击系统托盘，提示为tip的图标
Func ikSystrayItemClick($tip)
	$hWnd = ikGetSystray()
	$index = ikGetSystrayIndex($hWnd, $tip)
	if $index == -1 Then
		return -1
	Else
		_GUICtrlToolbar_ClickButton($hWnd, $index, "left")
		return 0
	EndIf
EndFunc

; 点击系统托盘，提示为tip的图标
Func ikSystrayItemRightClick($tip)
	$hWnd = ikGetSystray()
	$index = ikGetSystrayIndex($hWnd, $tip)
	if $index == -1 Then
		return -1
	Else
		_GUICtrlToolbar_ClickButton($hWnd, $index, "right")
		ikSleep(1)
		return 0
	EndIf
EndFunc
; 系统托盘的图标存在
Func ikFindSystrayItem($tip)
	$hWnd = ikGetSystray()
	$index = ikGetSystrayIndex($hWnd, $tip)
	If $index == -1 Then
		return -1
	Endif
	return 0
EndFunc
; 窗口存在且没有隐藏
Func ikWinExist($title)
	Local $state = WinGetState($title)
	If $state == 0 Then
		Return 0
	EndIf
	If BitAND($state, $WIN_STATE_VISIBLE) Then
		Return 0
	Endif
EndFunc

Func ikSend($str, $hWnd="")
	If $hWnd <> "" Then
		ControlSend($hWnd, "", "", $str)
	Else
		Send($str)
	Endif
EndFunc

; ctrl+c 复制或者结束控制台程序
Func ikSendCtrlC($hWnd = "")
	If $hWnd <> "" Then
		ControlSend($hWnd, "", "", "^c")
	Else
		Send("^c")
	Endif
EndFunc

; 按Enter
Func ikSendEnter($hWnd = "")
	ikSend("{Enter}", $hWnd)
	Sleep(100)
EndFunc

; ctrl+v 粘贴
Func ikSendCtrlV()
	Send("^v")
EndFunc

; 按方向键上num次
Func ikSendUp($num=1)
	For $i=1 to $num
		Send("{up down}")
		Sleep(100)
		Send("{up up}")
		Sleep(100)
	Next
EndFunc

; 按方向键下num次
Func ikSendDown($num=1)
	For $i=1 to $num
		Send("{down down}")
		Sleep(100)
		Send("{down up}")
		Sleep(100)
	Next
EndFunc

; 按方向键左num次
Func ikSendLeft($num=1)
	For $i=1 to $num
		Send("{left down}")
		Sleep(100)
		Send("{left up}")
		Sleep(100)
	Next
EndFunc

; 按方向键右num次
Func ikSendRight($num=1)
	For $i=1 to $num
		Send("{right down}")
		Sleep(100)
		Send("{right up}")
		Sleep(100)
	Next
EndFunc

; 时间转成字符串
Func ikFormatTime($sep = "-")
	return @YEAR & $sep & @MON & $sep & @MDAY & $sep & @HOUR & $sep & @MIN & $sep & @SEC ;& $sep & @MSEC
EndFunc

; 开始录制
Func ikRecordStart($file)
	Run("cmd /k title " & $ffmpegTitle, "", @SW_SHOW)
	ikSleep(1)
	Local $hWnd = WinGetHandle($ffmpegTitle)
	ikSend("ffmpeg.exe -y -f gdigrab -i desktop -r 10 -vcodec libx264 -pix_fmt yuv420p " &$file & @CR, $hWnd)
EndFunc
; 接收录制
Func ikRecordStop()
	WinClose("record screen use ffmpeg")
EndFunc

; 开始抓包，开始录屏
Func ikStartTest ($name)
	ikChangeToEnglish()
	$filePrefix = ikFormatTime()
	DirCreate(@ScriptDir & "\screenrecordings\" & $name)
	$fileRadio = @ScriptDir & "\screenrecordings\" & $name & "\" & $filePrefix & "-" & $name & ".mp4"
	$filePcap = @ScriptDir & "\screenrecordings\" & $name & "\" & $filePrefix & $name & "-yy.pcap"
	$fileLog = @ScriptDir & "\screenrecordings\" & $name & "\" & $filePrefix & $name & "-yy.log"
	ikLog("StarTesting...")
	Local $hub[3]
	$hub[0] = $fileRadio
    $hub[1] = $filePcap
	$title = ikRecordStart($fileRadio)
	$hub[2] = $title
	ikSleep(2)
	ikWiresharkStart()
    Return $hub
EndFunc

; 保存抓包，退出录屏
Func ikStopTest ($hub)
	ikLog("StopTesting...")
	$fileRadio = $hub[0]
	$filePcap = $hub[1]
	$title = $hub[2]
	ikWiresharkSaveQuit($filePcap)
	ikSleep(5)
	ikRecordStop()
	ikSleep(10)
	ikLog("Stopped")
	if WinExists($ffmpegTitle) Then
		ikLogError("测试结束，但是ffmpeg没有正常退出");
	EndIf
EndFunc
Func ikLog($msg)
	FileWrite($fileLog, "[" & @HOUR & ":" & @MIN & ":"& @SEC & "]" & " Info: " & $msg & @CR)
EndFunc
Func ikLogError($msg)
	FileWrite($fileLog, "[" & @HOUR & ":" & @MIN & ":"& @SEC & "]" & " Error: " & $msg & @CR)
EndFunc

Func ikSleep($sec)
	Sleep($sec * 1000)
EndFunc
Func ikMsgBox($title, $msg, $timeout=2)
	If $timeout == 0 Then
		$timeout = 1
	EndIf
	If $title == "Info" Then
		MsgBox(64, $title, $msg, $timeout)
	ElseIf $title == "Error" Then
		MsgBox(16, $title, $msg, $timeout)
	Else
		MsgBox(0, $title, $msg, $timeout)
	Endif
EndFunc