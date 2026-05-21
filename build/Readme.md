This document describes the steps needed to build SDK.

1. Required software installation:
	- Python 3.

2. SDK build:
	- python build.py

   Useful flags:
	- --develop                     generate browser-debugging scripts.js
	- --desktop                     include desktop-specific files
	- --mobile                      include mobile-specific files
	- --product word                build only one product (repeatable)
	- --addon ../sdkjs-addon-name   merge addon configs (repeatable)
