; Send( "keys" [, flag] ) Command.
; ^ Ctrl    ! Alt    + Shift    # Win
#include "ikuailib.au3"
Main()
Func Main ()
	$hub = ikStartTest(@ScriptName)
	ikSleep(2)

	TestCaseYY()

	ikSleep(2)
	ikStopTest($hub)
EndFunc

Func TestCaseYY ()
	ikRunLink("YY语音")
	ikSleep(5)
	If ikWinExist("YY") Then
		MsgBox(0, "Info", "YY ikWinExist", 2)
	EndIf
	$hWnd = WinGetHandle("YY")
	$title = WinGetTitle($hWnd)

	If $title == "YY" Then
		ikMsgBox("Info", "YY 经典模式", 2)
		ikWinSetSize($title, 300, 800)
		ikLog("YY 窗口为经典模式")
	ElseIf $title == "YY8.43" Then
		ikMsgBox("Info", "YY 大屏模式",2)
		ikLog("YY 窗口为大屏模式")
		ikWinSetSize($title, 1200, 800)
	Else
		ikMsgBox("Error", "没找到 YY的窗口，请检查", 5)
		ikLogError("没找到YY语音窗口")
		Return
	EndIf
	; 左键点击系统图标
	If ikFindSystrayItem("YY号") == 0 Then
		ikSleep(2)
		ikSystrayItemClick("YY号")
	Endif
	; 右键键点击系统图标
	If ikFindSystrayItem("YY号") == 0 Then
		ikSleep(2)
		ikSystrayItemRightClick("YY号")
		ikSleep(2)
		ikSendUp(5)
		ikSendEnter()
		ikSleep(5)
	Endif
EndFunc
