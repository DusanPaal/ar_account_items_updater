@ECHO OFF

ECHO === AR Account Item Text Updater Setup ver. 1.0.20220902 ===
set /P choice="Do you wish to install the application? (Y/N): "

IF NOT %choice%==Y IF NOT %choice%==y (
  pause>nul|set/p=Invalid parameter entered. Press any key to exit setup ...
  EXIT /b 0
)

ECHO Installing application...

IF NOT EXIST "C:\Program Files (x86)\Accounting Document Updater" (
  ECHO Creating installation directory ...
  md  "C:\Program Files (x86)\Accounting Document Updater"
)

ECHO Copying files ...
xcopy /s \app.xlsm "C:\Program Files (x86)\Accounting Document Updater"
xcopy /s \LEDbot.ico "C:\Program Files (x86)\Accounting Document Updater"
xcopy /s \LEDbot.jpeg "C:\Program Files (x86)\Accounting Document Updater"
ECHO Creating desktop shortcut ...

pause>nul|set/p=Installation completed. Press any key to exit setup ...