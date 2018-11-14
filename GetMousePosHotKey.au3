#include <MsgBoxConstants.au3>
#include <Process.au3>
HotKeySet( "!{F1}", "MousePosToClipBoard")
MsgBox(64, "Important", "Click Alt+F1 to Get Mouse Relative Position")
Func MousePosToClipBoard ()
	$title = WinGetTitle("")
	$pid = WinGetProcess("")
	$processName = _ProcessGetName($pid)
	$wPos = WinGetPos("")
	$mPos = MouseGetPos()
	ConsoleWrite($wPos[0] & @CR)
	ConsoleWrite($mPos[0] & @CR)
	$x = $mPos[0] - $wPos[0]
	$y = $mPos[1] - $wPos[1]
	$str = 'ikMouseRelativeClick("' & $title & '"' & ', ' & $x & ', ' & $y & ' ,1); ' & $processName
	ClipPut($str)
	MsgBox( 64, "Succeed to get Mouse Relative Pos", "Use Ctrl+V to paste", 2)
EndFunc

While 1
   sleep(10)
Wend
