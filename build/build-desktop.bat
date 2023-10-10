CD /D %~dp0
call npm install -g grunt-cli
call npm ci
call grunt --level=WHITESPACE_ONLY --desktop=true --formatting=PRETTY_PRINT
rem call grunt --level=ADVANCED --desktop=true

rmdir "..\..\desktop-apps\win-linux\build\debug\win_64\editors\sdkjs"
xcopy /s/e/k/c/y/q/i "..\deploy\sdkjs" "..\..\desktop-apps\win-linux\build\debug\win_64\editors\sdkjs"

pause