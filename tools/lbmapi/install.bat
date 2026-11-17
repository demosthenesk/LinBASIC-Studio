@echo off
rem LinBASIC Studio, phase 7.9: lbmapi.dll as the default mail program of Windows (Simple MAPI), for the VB6 tests.
rem Run it as administrator (right click > Run as administrator). uninstall.bat takes it away.
mkdir C:\LBMAPI 2>nul
mkdir C:\LBMAPI\store 2>nul
copy /Y "%~dp0lbmapi.dll" C:\LBMAPI\lbmapi.dll
if errorlevel 1 goto fail
reg add "HKLM\SOFTWARE\Clients\Mail\LinBASIC Test Mail" /ve /d "LinBASIC Test Mail" /f >nul
reg add "HKLM\SOFTWARE\Clients\Mail\LinBASIC Test Mail" /v DLLPath /d "C:\LBMAPI\lbmapi.dll" /f >nul
reg add "HKLM\SOFTWARE\Clients\Mail\LinBASIC Test Mail" /v DLLPathEx /d "C:\LBMAPI\lbmapi.dll" /f >nul
reg add "HKLM\SOFTWARE\Clients\Mail" /ve /d "LinBASIC Test Mail" /f >nul
reg add "HKCU\SOFTWARE\Clients\Mail" /ve /d "LinBASIC Test Mail" /f >nul
if errorlevel 1 goto fail
echo LinBASIC Test Mail is the mail program now (C:\LBMAPI\lbmapi.dll, store C:\LBMAPI\store).
pause
exit /b 0
:fail
echo FAILED: run it as administrator.
pause
exit /b 1
