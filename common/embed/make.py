#!/usr/bin/env python

import sys
sys.path.append('../../../build_tools/scripts')
import base
import os

params = sys.argv[1:]

#compilation_level = "WHITESPACE_ONLY"
compilation_level = "SIMPLE_OPTIMIZATIONS"
base.cmd("python", ["./min.py"])

min_content = base.readFile("./embed.min.js")

if (1 != len(params)):
  exit(0)

api_file = params[0]
api_content = base.readFile(api_file)

pos_return_editor_embed = api_content.find("var onMouseUp")
pos_return_editor_obj = api_content.find("return {")

new_content = ""
new_content += "function _createEmbedWorker() { return AscEmbed.initWorker(iframe); }"
new_content += "\n\n"

new_content += "        return {"
new_content += "\n"
new_content += "            createEmbedWorker   : _createEmbedWorker,"

new_api_content = api_content[0:pos_return_editor_embed] + min_content + "\n        " + api_content[pos_return_editor_embed:pos_return_editor_obj] + new_content + api_content[pos_return_editor_obj + 8:]

base.delete_file(api_file)
base.writeFile(api_file, new_api_content)
