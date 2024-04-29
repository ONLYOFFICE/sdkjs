#!/usr/bin/env python
import sys
sys.path.append('../../build_tools/scripts')
import base
import traceback

try:
    base.cmd_in_dir('.', "npm", ["install", "-g", "grunt-cli"])
    base.cmd_in_dir('.', "npm", ["ci"])

    base.cmd_in_dir('.', "grunt", ["--level=WHITESPACE_ONLY", "--addon=sdkjs-forms", "--addon=sdkjs-ooxml"])
    base.cmd_in_dir('.', "grunt", ["develop", "--compiled", "--addon=sdkjs-forms", "--addon=sdkjs-ooxml"])

    raw_input("Press Enter to continue...")

except SystemExit:
    raw_input("Ignoring SystemExit. Press Enter to continue...")
    exit(0)
except KeyboardInterrupt:
    pass
except:
    raw_input("Unexpected error. " + traceback.format_exc() + "Press Enter to continue...")
