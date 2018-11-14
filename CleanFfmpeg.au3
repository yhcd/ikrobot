#include "ikuailib.au3"
if WinExists($ffmpegTitle) Then
	ikMsgBox("Error", "测试结束，但是ffmpeg没有正常退出", 5)
EndIf
ikRecordStop()
