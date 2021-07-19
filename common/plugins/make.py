#!/usr/bin/env python

import sys
sys.path.append('../../../build_tools/scripts')
import base
import os

#compilation_level = "WHITESPACE_ONLY"
compilation_level = "SIMPLE_OPTIMIZATIONS"
base.cmd("java", ["-jar", "../../build/node_modules/google-closure-compiler-java/compiler.jar", 
                  "--compilation_level", compilation_level,
                  "--js_output_file", "plugins.js",
                  "--js", "../device_scale.js", 
                  "--js", "plugin_base.js", 
                  "--js", "plugin_base_api.js"])

code_content = base.readFile("plugins.js")
dst_content = base.readFile("../plugins.js")

startCode = dst_content.find("/*<code>*/")
endCode = dst_content.find("/*</code>*/")

code_content = code_content.replace("\r", "")
code_content = code_content.replace("\n", "")
code_content = code_content.replace("\\", "\\\\")
code_content = code_content.replace("\"", "\\\"")

content = ""
content += dst_content[0:startCode]
content += "/*<code>*/\""
content += code_content
content += "\""
content += dst_content[endCode:]

base.delete_file("plugins.js")
base.delete_file("../plugins.js")

base.writeFile("../plugins.js", content)
