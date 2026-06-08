#!/usr/bin/env python3
"""Build OnlyOffice SDK JS — concatenation only, no compiler.

Usage:
    python build.py                           # build all (word, cell, slide, visio)
    python build.py --product word            # build only word
    python build.py --desktop                 # include desktop-specific files
    python build.py --mobile                  # include mobile-specific files
    python build.py --addon ../sdkjs-forms    # merge addon config (repeatable)
    python build.py --develop                 # scripts.js for browser debugging
    python build.py --develop --compiled      # scripts.js pointing at deploy

Environment variables (optional):
    APP_COPYRIGHT    — copyright string in license header
    PUBLISHER_URL    — company URL in license header
    PRODUCT_VERSION  — value substituted in license header (default: 0.0.0)
    BUILD_NUMBER     — value substituted in license header (default: 0)
"""
import os
import re
import sys
import json
import glob
import shutil
import argparse
from datetime import datetime

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.join(SCRIPT_DIR, '..')
DEPLOY_DIR  = os.path.join(ROOT_DIR, 'deploy', 'sdkjs')
DEVELOP_DIR = os.path.join(ROOT_DIR, 'develop', 'sdkjs')

CONFIG_NAMES = ['externs', 'word', 'cell', 'slide', 'visio']

OTHER_FILES = [
    {
        'cwd':  os.path.join(ROOT_DIR, 'common'),
        'src':  [
            'Drawings/Format/path-boolean-min.js',
            'Charts/ChartStyles.js',
            'SmartArts/SmartArtData/*',
            'SmartArts/SmartArtDrawing/*',
            'Images/*',
            'Images/placeholders/*',
            'Images/content_controls/*',
            'Images/cursors/*',
            'Images/reporter/*',
            'Images/icons/*',
            'Native/*.js',
            'libfont/engine/*',
            'spell/spell/*',
            'hash/hash/*',
            'zlib/engine/*',
            'serviceworker/*',
        ],
        'dest': os.path.join(DEPLOY_DIR, 'common'),
    },
    {
        'cwd':  os.path.join(ROOT_DIR, 'cell', 'css'),
        'src':  ['*.css'],
        'dest': os.path.join(DEPLOY_DIR, 'cell', 'css'),
    },
    {
        'cwd':  os.path.join(ROOT_DIR, 'slide', 'themes'),
        'src':  ['**/**'],
        'dest': os.path.join(DEPLOY_DIR, 'slide', 'themes'),
    },
    {
        'cwd':  os.path.join(ROOT_DIR, 'pdf'),
        'src':  ['src/engine/*', 'src/annotations/stamps/*.json'],
        'dest': os.path.join(DEPLOY_DIR, 'pdf'),
    },
]


# ---------------------------------------------------------------------------
# Config loading (with addon merge)
# ---------------------------------------------------------------------------

def load_json(path):
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return None


def fix_paths(obj, base_dir):
    """Recursively prepend base_dir to every string leaf (config paths are relative to sdkjs root)."""
    if isinstance(obj, list):
        return [os.path.join(base_dir, item) for item in obj]
    if isinstance(obj, dict):
        return {k: fix_paths(v, base_dir) for k, v in obj.items()}
    return obj


def merge_into(base, addon):
    for key, val in addon.items():
        if isinstance(val, list):
            base[key] = base.get(key, []) + val
        elif isinstance(val, dict):
            merge_into(base.setdefault(key, {}), val)
        else:
            base[key] = val


def load_configs(dirs):
    """Load configs from sdkjs root + addon directories. Returns a dict keyed by config name."""
    configs = {}
    for d in dirs:
        for name in CONFIG_NAMES:
            data = load_json(os.path.join(d, 'configs', name + '.json'))
            if data is None:
                continue
            data = fix_paths(data, d)
            if name not in configs:
                configs[name] = data
            else:
                merge_into(configs[name], data)
    return configs


def get_files(sdk, mobile=False, desktop=False):
    """Return (min_files, common_files) for one product SDK config."""
    min_files = []
    if mobile:
        min_files += sdk.get('mobile_banners', {}).get('min', [])
    min_files += sdk['min']
    if desktop:
        min_files += sdk.get('desktop', {}).get('min', [])

    common = list(sdk['common'])
    if mobile:
        common = sdk.get('mobile_banners', {}).get('common', []) + common
        common = [f for f in common if f not in set(sdk.get('exclude_mobile', []))]
        common += sdk.get('mobile', [])
    if desktop:
        common += sdk.get('desktop', {}).get('common', [])

    return min_files, common


# ---------------------------------------------------------------------------
# Build steps
# ---------------------------------------------------------------------------

def apply_defines(content, defines):
    """Rewrite `window.AscCommon.<name> = "..."` assignments to the given values.
    Tolerant of any whitespace around `=` (matches both `name = "x"` and `name="x"`).
    """
    for name, value in defines.items():
        pattern = re.compile(r'window\.AscCommon\.' + re.escape(name) + r'\s*=\s*"[^"]*"')
        replacement = f'window.AscCommon.{name} = "{value}"'
        content = pattern.sub(lambda m, r=replacement: r, content)
    return content


def concat_files(files, dest, header, defines=None, iife=False):
    """Concatenate JS source files into dest, prepending the license header.
    If iife=True, wrap the concatenated body in `(function(window, undefined){ ... })(window);`.
    If defines is given, apply window.AscCommon.* value replacements via regex.
    """
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    parts = [header, '\n']
    if iife:
        parts.append('(function(window, undefined) {\n')
    for path in files:
        with open(path, 'r', encoding='utf-8') as f:
            parts.append(f.read())
            parts.append('\n')
    if iife:
        parts.append('})(window);\n')

    content = ''.join(parts)
    if defines:
        content = apply_defines(content, defines)

    with open(dest, 'w', encoding='utf-8') as out:
        out.write(content)


def build_sdk(name, sdk, license_text, defines, args):
    """Concatenate one product into sdk-all-min.js + sdk-all.js."""
    out_dir = os.path.join(DEPLOY_DIR, name)
    min_files, common_files = get_files(sdk, mobile=args.mobile, desktop=args.desktop)
    print(f'Building {name}...')
    concat_files(min_files,    os.path.join(out_dir, 'sdk-all-min.js'), license_text, defines)
    concat_files(common_files, os.path.join(out_dir, 'sdk-all.js'),     license_text, defines, iife=True)


def deploy_assets():
    """Copy static assets and helper JS files into deploy/."""
    for item in OTHER_FILES:
        cwd = item['cwd']
        for pattern in item['src']:
            for src_abs in glob.glob(os.path.join(cwd, pattern), recursive=True):
                if not os.path.isfile(src_abs):
                    continue
                rel  = os.path.relpath(src_abs, cwd)
                dest = os.path.join(item['dest'], rel)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copy2(src_abs, dest)


def write_scripts_js(sdk, name, compiled=False):
    """Write develop/sdkjs/<name>/scripts.js listing source files for browser debugging."""
    out_dir = os.path.join(DEVELOP_DIR, name)
    os.makedirs(out_dir, exist_ok=True)

    def to_url(p):
        return '../../../../sdkjs/' + os.path.relpath(p, ROOT_DIR).replace(os.sep, '/')

    if compiled:
        files = [
            to_url(os.path.join(ROOT_DIR, 'vendor', 'polyfill.js')),
            to_url(os.path.join(ROOT_DIR, 'common', 'AllFonts.js')),
            to_url(os.path.join(DEPLOY_DIR, name, 'sdk-all-min.js')),
        ]
    else:
        min_files, common_files = get_files(sdk)
        files = [to_url(f) for f in (
            [os.path.join(ROOT_DIR, 'vendor', 'polyfill.js'),
             os.path.join(ROOT_DIR, 'common', 'AllFonts.js'),
             os.path.join(ROOT_DIR, 'common', 'applyDocumentChanges.js')]
            + min_files + common_files
        )]

    out_path = os.path.join(out_dir, 'scripts.js')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('var sdk_scripts = [\n\t"' + '",\n\t"'.join(files) + '"\n];')
    print(f'  Written {os.path.relpath(out_path, SCRIPT_DIR)}')


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='Build OnlyOffice SDK JS by concatenating sources (no compiler).',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('--src',     default=None, metavar='PATH',
                        help='sdkjs root directory (default: parent of build/)')
    parser.add_argument('--develop', action='store_true',
                        help='generate develop/sdkjs/<product>/scripts.js for browser debugging')
    parser.add_argument('--compiled', action='store_true',
                        help='(with --develop) point scripts.js at deploy/sdk-all-min.js')
    parser.add_argument('--addon',   action='append', default=[], metavar='PATH',
                        help='addon directory whose configs are merged in (repeatable)')
    parser.add_argument('--mobile',  action='store_true', help='include mobile-specific files')
    parser.add_argument('--desktop', action='store_true', help='include desktop-specific files')
    parser.add_argument('--beta',    default='false',     help='value for window.AscCommon.g_cIsBeta (default: false)')
    parser.add_argument('--map',     action='store_true', help='(no-op in concat build; accepted for compatibility)')
    parser.add_argument('--minimize', action=argparse.BooleanOptionalAction, default=None,
                        help='(no-op in concat build; accepted for compatibility)')
    parser.add_argument('--product', action='append',     default=[],
                        choices=['word', 'cell', 'slide', 'visio'], metavar='PRODUCT',
                        help='build only this product (default: all four, repeatable)')
    args = parser.parse_args()

    src_dir = os.path.abspath(args.src) if args.src else ROOT_DIR
    addon_dirs = [
        os.path.abspath(a) if (os.path.isabs(a) or os.path.isdir(a)) else os.path.join(ROOT_DIR, '..', a)
        for a in args.addon
    ]

    configs = load_configs([src_dir] + addon_dirs)
    if not all(configs.get(n) for n in CONFIG_NAMES):
        print('ERROR: Could not load all required configs. '
              'Make sure configs/ directory exists with word/cell/slide/visio/externs JSON files.')
        sys.exit(1)

    products = args.product or ['word', 'cell', 'slide', 'visio']

    if args.develop:
        print('Building develop scripts...')
        if os.path.exists(DEVELOP_DIR):
            shutil.rmtree(DEVELOP_DIR)
        for name in products:
            write_scripts_js(configs[name]['sdk'], name, compiled=args.compiled)
        print('Done.')
        return

    version      = os.environ.get('PRODUCT_VERSION', '0.0.0')
    build_number = os.environ.get('BUILD_NUMBER', '0')
    year         = datetime.now().year

    with open(os.path.join(SCRIPT_DIR, 'license.header'), encoding='utf-8') as f:
        license_text = (f.read()
            .replace('@@AppCopyright', os.environ.get('APP_COPYRIGHT',
                f'Copyright (C) Ascensio System SIA 2009-{year}. All rights reserved'))
            .replace('@@PublisherUrl',  os.environ.get('PUBLISHER_URL', 'https://www.onlyoffice.com/'))
            .replace('@@Version',       version)
            .replace('@@Build',         build_number))

    defines = {
        'g_cCompanyName':    os.environ.get('COMPANY_NAME', 'onlyoffice'),
        'g_cProductVersion': version,
        'g_cBuildNumber':    build_number,
        'g_cIsBeta':         args.beta,
    }

    print('Cleaning deploy directory...')
    if os.path.exists(DEPLOY_DIR):
        shutil.rmtree(DEPLOY_DIR)

    for name in products:
        build_sdk(name, configs[name]['sdk'], license_text, defines, args)

    print('Copying other files...')
    deploy_assets()
    print('Build complete.')


if __name__ == '__main__':
    main()
