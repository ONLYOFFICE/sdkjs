import sys
sys.path.append("../../../../build_tools/scripts")
import base
import os

base.configure_common_apps()

# remove previous version
if base.is_dir("./../deploy"):
    base.delete_dir("./../deploy")
base.create_dir("./../deploy")

# command
base.cmd_in_dir("./../../../../core/DesktopEditor/graphics/pro/js", "python", ["make.py"])

# finalize
if base.is_exist("./../../../../core/DesktopEditor/graphics/pro/js/deploy/drawingfile.wasm"):
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/drawingfile.js", "./../deploy/drawingfile.js")
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/drawingfile.wasm", "./../deploy/drawingfile.wasm")
else:
    print("xps_djvu_make.py error")
    base.copy_dir("./../pdf", "./../deploy")
    base.copy_dir("./../xps_djvu", "./../deploy")

# write new version
