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
if base.is_exist("./../../../../core/DesktopEditor/graphics/pro/js/deploy/xps_djvu_pdf.wasm"):
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/xps.js", "./../deploy/xps.js")
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/djvu.js", "./../deploy/djvu.js")
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/pdf.js", "./../deploy/pdf.js")
    base.copy_file("./../../../../core/DesktopEditor/graphics/pro/js/deploy/xps_djvu_pdf.wasm", "./../deploy/xps_djvu_pdf.wasm")
else:
    print("xps_djvu_make.py error")
    base.copy_dir("./../pdf", "./../deploy")
    base.copy_dir("./../xps_djvu", "./../deploy")

# write new version
