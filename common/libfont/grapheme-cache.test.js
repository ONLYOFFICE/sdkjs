/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Synthetic stand-in for ONLYOFFICE/DocumentServer#3745.
// A symbolic subset font can put a letter on the same glyph id that a
// missing cmap also uses for U+0020. Editing any other text used to
// rewrite that letter as a space. No customer PDF is required.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const sandbox = {};
sandbox.window = sandbox;
sandbox.AscFonts = {};
vm.runInNewContext(
	fs.readFileSync(path.join(__dirname, "grapheme.js"), "utf8"),
	sandbox
);
const fonts = sandbox.AscFonts;

function codePoints(value) {
	return {
		getCount: function () { return 1; },
		get: function () { return value; }
	};
}

function graphemeFor(gid, codePoint) {
	fonts.InitGrapheme(1, 0);
	fonts.AddGlyphToGrapheme(gid, 100, 0, 0, 0);
	return fonts.GetGrapheme(codePoints(codePoint));
}

const space = graphemeFor(0x20, 0x20);
const letterV = graphemeFor(0x20, 0x56);
const letterK = graphemeFor(0x20, 0x6B);

function codes(graphemeId) {
	return Array.from(fonts.GetGraphemeCodePoints(graphemeId));
}

assert.notStrictEqual(space, letterV);
assert.notStrictEqual(space, letterK);
assert.notStrictEqual(letterV, letterK);
assert.deepStrictEqual(codes(space), [0x20]);
assert.deepStrictEqual(codes(letterV), [0x56]);
assert.deepStrictEqual(codes(letterK), [0x6B]);

assert.strictEqual(graphemeFor(0x20, 0x56), letterV);
assert.strictEqual(graphemeFor(0x20, 0x20), space);

const letterFirst = graphemeFor(0x11, 0x56);
const spaceSecond = graphemeFor(0x11, 0x20);
assert.notStrictEqual(letterFirst, spaceSecond);
assert.deepStrictEqual(codes(letterFirst), [0x56]);
assert.deepStrictEqual(codes(spaceSecond), [0x20]);

console.log("grapheme-cache.test.js ok");
