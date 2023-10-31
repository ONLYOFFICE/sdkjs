CD /D %~dp0
call npm install -g grunt-cli
call npm ci

call grunt --level=WHITESPACE_ONLY --addon=sdkjs-forms --addon=sdkjs-ooxml
call grunt develop --addon=sdkjs-forms --addon=sdkjs-ooxml

pause