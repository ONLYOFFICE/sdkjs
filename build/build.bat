CD /D %~dp0
call npm install -g grunt-cli
call npm ci
rem call grunt --level=WHITESPACE_ONLY --desktop=false --formatting=PRETTY_PRINT
call grunt --level=ADVANCED

pause