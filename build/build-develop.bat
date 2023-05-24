CD /D %~dp0
call npm install -g grunt-cli
call npm install

REM call grunt --level=ADVANCED --addon=sdkjs-forms --addon=sdkjs-ooxml
REM call grunt --level=ADVANCED --addon=sdkjs-forms
call grunt --level=ADVANCED --addon=sdkjs-forms --addon=sdkjs-ooxml  --desktop=true
call grunt develop

pause