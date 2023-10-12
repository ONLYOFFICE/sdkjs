CD /D %~dp0
call npm install -g grunt-cli
call npm ci

REM call grunt --level=ADVANCED --addon=sdkjs-forms --addon=sdkjs-ooxml  --desktop=true
REM call grunt --level=ADVANCED --addon=sdkjs-forms --addon=sdkjs-ooxml
REM call grunt --level=ADVANCED --addon=sdkjs-forms
call grunt --level=WHITESPACE_ONLY
call grunt develop

pause