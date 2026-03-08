@echo off
copy /b "DevConference2026.zip.01" + "DevConference2026.zip.02" "DevConference2026.zip.restored.tmp"
if errorlevel 1 goto :error

set i=3
:loop
if exist "DevConference2026.zip.0%i%" (
    copy /b "DevConference2026.zip.restored.tmp" + "DevConference2026.zip.0%i%" "DevConference2026.zip.restored2.tmp" >nul
    move /y "DevConference2026.zip.restored2.tmp" "DevConference2026.zip.restored.tmp" >nul
    set /a i+=1
    goto loop
)

if exist "DevConference2026.zip.%i%" (
    copy /b "DevConference2026.zip.restored.tmp" + "DevConference2026.zip.%i%" "DevConference2026.zip.restored2.tmp" >nul
    move /y "DevConference2026.zip.restored2.tmp" "DevConference2026.zip.restored.tmp" >nul
    set /a i+=1
    goto loop
)

move /y "DevConference2026.zip.restored.tmp" "DevConference2026.zip" >nul
echo 復允E��亁E DevConference2026.zip
goto :eof

:error
echo 復允E��失敗しました
exit /b 1
