@echo off
rem LinBASIC Studio, phase 7.9: takes lbmapi.dll away as the mail program (install.bat put it there). As administrator.
reg delete "HKCU\SOFTWARE\Clients\Mail" /ve /f >nul 2>nul
reg delete "HKLM\SOFTWARE\Clients\Mail" /ve /f >nul 2>nul
reg delete "HKLM\SOFTWARE\Clients\Mail\LinBASIC Test Mail" /f >nul 2>nul
del /Q C:\LBMAPI\lbmapi.dll 2>nul
echo LinBASIC Test Mail removed (the messages of C:\LBMAPI\store stay; delete the folder if you like).
pause
