/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Synthetic stand-in for ONLYOFFICE/DocumentServer#3844.
// Three Ink appearances: Multiply at /CA 1, Normal at /CA 1, Multiply at /CA 0.4.
// The shipped wasm reports blend mode 0, so the editor reads /BM from the
// appearance dictionaries. No customer PDF is required.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "base.js"), "utf8");
const start = source.indexOf('window["AscPDF"].indexAppearanceBlend = function');
const end = source.indexOf('window["AscPDF"].startBlendMode = function');
assert.ok(start > 0 && end > start, "indexAppearanceBlend was not found in base.js");

const sandbox = { TextDecoder: TextDecoder };
sandbox.window = sandbox;
sandbox.AscPDF = {};
vm.runInNewContext(source.slice(start, end), sandbox);

const indexAppearanceBlend = sandbox.AscPDF.indexAppearanceBlend;
const fixture = fs.readFileSync(path.join(__dirname, "ink-multiply.fixture.pdf"));
const byName = indexAppearanceBlend(fixture);

assert.strictEqual(byName["ink-9"], 1);
assert.strictEqual(byName["ink-11"], 1);
assert.strictEqual(byName["ink-10"], undefined);
assert.strictEqual(Object.keys(byName).length, 2);
assert.deepStrictEqual(Object.keys(indexAppearanceBlend(new Uint8Array())), []);

const text = fixture.toString("latin1");
assert.ok(text.indexOf("/BM /Multiply") > 0);
assert.ok(text.indexOf("/BM /Normal") > 0);
assert.ok(text.indexOf("/CA 1") > 0);
assert.ok(text.indexOf("/CA 0.4") > 0);

console.log("appearance-blend.test.js ok");
