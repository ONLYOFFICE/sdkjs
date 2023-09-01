#!/usr/bin/env python

import sys
sys.path.append('../../../build_tools/scripts')
import base
import os

#compilation_level = "WHITESPACE_ONLY"
#compilation_level = "SIMPLE_OPTIMIZATIONS"
compilation_level = "ADVANCED"

base.writeFile("./begin.js", "window[\"AscCommon\"] = window[\"AscCommon\"] || {};\n\n")

scripts_code = [
  "./begin.js",
  "./../../common/errorCodes.js",
  "./../../common/device_scale.js",
  "./../../common/browser.js",
  "./../../common/stringserialize.js",
  "./../../common/skin.js",
  "./../../common/libfont/loader.js",
  "./../../common/libfont/map.js",
  "./../../common/libfont/character.js",
  "./../../common/SerializeCommonWordExcel.js",
  "./../../common/Drawings/Externals.js",
  "./../../common/GlobalLoaders.js",
  "./../../common/scroll.js",
  "./../../common/Drawings/WorkEvents.js",
  "./../../common/Overlay.js",
  "./../src/thumbnails.js",
  "./../src/viewer.js",
  "./../src/file.js",
  "./api.js"
]

externals = [
    "./../../common/externs/global.js",
    "./../../common/externs/jquery-3.2.js",
    "./../../common/externs/xregexp-3.0.0.js",
    "./../../common/externs/socket.io.js",
    "./../../common/externs/word.js",
    "./../../common/externs/cell.js",
    "./../../common/externs/slide.js"
]

build_params = []
build_params.append("-jar")
build_params.append("../../build/node_modules/google-closure-compiler-java/compiler.jar")
build_params.append("--compilation_level")
build_params.append(compilation_level)
build_params.append("--jscomp_off=checkVars")
build_params.append("--warning_level=QUIET")
build_params.append("--js_output_file")
build_params.append("./../src/engine/viewer.js")

for item in scripts_code:
  build_params.append("--js")
  build_params.append(item)

for item in externals:
  build_params.append("--externs=" + item)

base.cmd("java", build_params)

base.delete_file("./begin.js")

licence_text = base.readFile("./api.js")
licence_text_end = licence_text.find("(function(){")
licence_text = licence_text[:licence_text_end]

generate_file = base.readFile("./../src/engine/viewer.js")
base.writeFile("./../src/engine/viewer.js", licence_text + generate_file)
