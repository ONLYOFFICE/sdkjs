import subprocess
import os
import sys

def run_minification():
    input_file = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'common', 'Drawings', 'Format', 'path-boolean.js'))
    if not os.path.isfile(input_file):
        print(f'Source file not found: {input_file}')
        sys.exit(1)

    closure_compiler_path = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'build', 'node_modules', '.bin', 'google-closure-compiler'))
    if not os.path.isfile(closure_compiler_path):
        print(f'Closure Compiler file not found: {closure_compiler_path}')
        sys.exit(1)

    output_file = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'common', 'Drawings', 'Format', 'path-boolean-min.js'))
    command = [
        closure_compiler_path,
        '--language_out=ECMASCRIPT5',
        '--compilation_level=ADVANCED',
        '--warning_level=QUIET',
        f'--js={input_file}',
        f'--js_output_file={output_file}'
    ]

    result = subprocess.run(command, cwd=None, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f'Command execution error: {result.stderr}')
        sys.exit(1)
    
    print(f'Minified {os.path.basename(input_file)}')

run_minification()
