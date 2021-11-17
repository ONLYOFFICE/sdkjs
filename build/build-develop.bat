CD /D %~dp0
call npm install -g grunt-cli
call npm install

call grunt --level=WHITESPACE_ONLY
call grunt develop

pause