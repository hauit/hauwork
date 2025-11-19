@echo off
setlocal

:: tên file exe cần tạo shortcut (đặt cùng folder với file .bat)
set TARGET=MOS360_EXCEL_TEST.exe

:: tên shortcut sẽ tạo
set SHORTCUT=EXCEL_TEST.lnk

:: lấy thư mục hiện tại (chứa .bat + app)
set CURRDIR=%~dp0

:: tạo đối tượng WScript.Shell để tạo shortcut trong chính folder này
(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo sLinkFile = "%CURRDIR%%SHORTCUT%"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%CURRDIR%%TARGET%"
echo oLink.WorkingDirectory = "%CURRDIR%"
echo oLink.IconLocation = "%CURRDIR%%TARGET%,0"
echo oLink.Save
) > "%TEMP%\createshortcut.vbs"

:: chạy script VBS
cscript //nologo "%TEMP%\createshortcut.vbs"

del "%TEMP%\createshortcut.vbs"
echo Shortcut da duoc tao trong folder app!
pause
