#!/usr/bin/env python

import sys
sys.path.append('../../../build_tools/scripts')
import base
import os

#compilation_level = "WHITESPACE_ONLY"
compilation_level = "SIMPLE_OPTIMIZATIONS"
base.cmd("java", ["-jar", "../../build/node_modules/google-closure-compiler-java/compiler.jar", 
                  "--compilation_level", compilation_level,
                  "--js_output_file", "embed.min.js",
                  "--js", "embed.js"])

