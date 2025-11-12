#!/usr/bin/env python
import sys
sys.path.append('../../build_tools/scripts')
import base
import traceback

try:
    base.cmd_in_dir('.', "npm", ["install", "-g", "grunt-cli"])
    base.cmd_in_dir('.', "npm", ["ci"])

    base.cmd_in_dir('.', "grunt", ["--level=WHITESPACE_ONLY", "--addon=sdkjs-forms", "--addon=sdkjs-ooxml"])
    base.cmd_in_dir('.', "grunt", ["develop", "--addon=sdkjs-forms", "--addon=sdkjs-ooxml"])

    input("Press Enter to continue...")
    exit(0)
except SystemExit:
    input("Ignoring SystemExit. Press Enter to continue...")
    exit(0)
except KeyboardInterrupt:
    pass
except:
    input("Unexpected error. " + traceback.format_exc() + "Press Enter to continue...")
    exit(0)
