taskkill /im ffmpeg.exe /f
taskkill /im wireshark.exe /f
yy.au3
rem 如果脚本挂了，需要清理 ffmpeg.exe 进程
CleanFfmpeg.au3
taskkill /im wireshark.exe /f
