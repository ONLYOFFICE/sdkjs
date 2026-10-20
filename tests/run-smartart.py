#!/usr/bin/env python3
"""Focused real-browser SmartArt checks. Requires Chromium and sibling web-apps.

Run: python tests/run-smartart.py [--desktop|--mobile] [word|cell|slide|pdf ...]
"""
import argparse
import concurrent.futures
import functools
import html
import http.server
import json
from pathlib import Path
import re
import subprocess
import sys
import threading

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
variant = parser.add_mutually_exclusive_group()
variant.add_argument("--desktop", action="store_true")
variant.add_argument("--mobile", action="store_true")
parser.add_argument("products", nargs="*", choices=("word", "cell", "slide", "pdf"))
args = parser.parse_args()
products = args.products or ["word", "cell", "slide", "pdf"]
build = [sys.executable, str(root / "build/build.py")]
if args.desktop:
    build.append("--desktop")
if args.mobile:
    build.append("--mobile")
for product in dict.fromkeys("word" if product == "pdf" else product for product in products):
    build.extend(["--product", product])
subprocess.run(build, cwd=root, check=True)


class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass


server = http.server.ThreadingHTTPServer(
    ("127.0.0.1", 0), functools.partial(Handler, directory=str(root.parent)))
threading.Thread(target=server.serve_forever, daemon=True).start()


def check(product):
    url = f"http://127.0.0.1:{server.server_port}/{root.name}/tests/smartart-editing.html?product={product}"
    result = subprocess.run([
        "chromium", "--headless", "--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage",
        "--dump-dom", "--virtual-time-budget=30000", url,
    ], capture_output=True, text=True, timeout=55, check=True)
    match = re.search(r'<pre id="result">(.*?)</pre>', result.stdout, re.S)
    if not match or match[1] == "RUNNING":
        raise RuntimeError(f"{product}: browser did not finish")
    data = json.loads(html.unescape(match[1]))
    failures = [item for item in data["results"] if "error" in item]
    additions = sum(item.get('operation') == 'add/remove' for item in data['results'])
    print(f"{product}: {len(data['results']) - len(failures)}/{len(data['results'])} passed ({additions} presets with added nodes)", flush=True)
    for failure in failures[:5]:
        print(f"  {failure['name']}: {failure['error']}", flush=True)
    if len(failures) > 5:
        print(f"  ... {len(failures) - 5} more failures", flush=True)
    return not failures and len(data["results"]) == 152


try:
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        passed = list(executor.map(check, products))
finally:
    server.shutdown()
sys.exit(0 if all(passed) else 1)
