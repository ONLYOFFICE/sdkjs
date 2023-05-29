#!/usr/bin/env python

import sys
sys.path.append('../../build_tools/scripts')
import base
import os
import glob

base_dir = base.get_script_dir() + "/../../sdkjs/common/Images/cursors"

content = "{\n"

for file in glob.glob(base_dir + "/*.svg"):
  basename = os.path.basename(file)[:-4]
  file_content = base.readFile(file)
  file_content = file_content.replace("\"", "'")
  file_content = file_content.replace("\n", "")
  content += "    \"" + basename + "\" : \"" + file_content + "\",\n"

content = content[:-2]
content += "\n}"

base.writeFile(base_dir + "/svg.json", content)
