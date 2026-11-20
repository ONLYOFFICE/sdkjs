/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Arial Narrow is not installed on a stock DocumentServer. A PDF that names it
// without embedding it (ArialNarrow, ArialNarrow,Bold, ArialNarrow,BoldItalic)
// got regular Arial from the name penalty, because "Arial" is a substring of
// "ArialNarrow". Edit Text and save then drew the narrow text with Arial
// glyphs at Arial Narrow advances. Liberation Sans Narrow has the same metrics
// and is the substitute; an installed Arial Narrow still wins.
// Run: node common/libfont/narrow-like-fonts.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

let source = fs.readFileSync(path.join(__dirname, "map.js"), "utf8");
// the picker's list entries and request are internal to map.js
const hook = "var g_fontApplication = new CApplicationFonts();";
assert.ok(source.includes(hook));
source = source.replace(hook, hook + "\n\twindow.__test = {CFontSelect: CFontSelect, CFontSelectFormat: CFontSelectFormat};");

const sandbox = {AscFonts: {FT_Common: {}}, AscCommon: {}};
sandbox.window = sandbox;
vm.runInNewContext(source, sandbox);
const {CFontSelect, CFontSelectFormat} = sandbox.__test;
const app = sandbox.AscFonts.g_fontApplication;

function face(name, bold, italic, width) {
	const f = new CFontSelect();
	f.m_wsFontName = name;
	f.m_bBold = bold;
	f.m_bItalic = italic;
	f.m_usWidth = width;
	f.m_usWeigth = bold ? 700 : 400;
	return f;
}
function family(name, width) {
	return [face(name, false, false, width), face(name, true, false, width), face(name, false, true, width), face(name, true, true, width)];
}
function pick(list, name) {
	app.g_fontSelections.List = list;
	app.DefaultIndex = 0;
	app.FontPickerMap = {};
	return app.GetFontFileWeb(name).m_wsFontName;
}

// a stock server: Arial first (the default), Liberation Sans Narrow, an unrelated narrow face
const stock = [].concat(family("Arial", 5), family("Liberation Sans Narrow", 3), family("Mukti Narrow", 5), family("Times New Roman", 5));

for (const name of ["ArialNarrow", "Arial Narrow", "HelveticaNarrow", "Helvetica Narrow", "Nimbus Sans Narrow", "Liberation Sans Narrow"])
	assert.strictEqual(pick(stock, name), "Liberation Sans Narrow", name);

// regular Arial and its likes are unchanged
for (const name of ["Arial", "Helvetica", "Liberation Sans"])
	assert.strictEqual(pick(stock, name), "Arial", name);
assert.strictEqual(pick(stock, "TimesNewRoman"), "Times New Roman");

// the real face wins when it is installed, wherever it is in the list
const withArialNarrow = stock.concat(family("Arial Narrow", 3));
assert.strictEqual(pick(withArialNarrow, "ArialNarrow"), "Arial Narrow");
assert.strictEqual(pick(withArialNarrow, "Arial Narrow"), "Arial Narrow");

// the group is the one core's picker uses (DesktopEditor/fontengine/ApplicationFonts.h)
const dict = app.g_fontDictionary;
const group = dict.FD_Ascii_Font_Like_Names[dict.FD_Ascii_Font_Like_Main["ArialNarrow"]];
assert.deepStrictEqual(Array.from(group), ["Arial Narrow", "Helvetica Narrow", "Nimbus Sans Narrow", "Liberation Sans Narrow"]);
assert.strictEqual(dict.CheckLikeFonts("Arial", "ArialNarrow"), false);

// bold and italic are applied to the picked family by the style, not by the name
const request = new CFontSelectFormat();
request.wsName = "ArialNarrow";
request.bBold = true;
request.bItalic = true;
app.g_fontSelections.List = stock;
const boldItalic = dict.GetFontIndex(request, stock, 0, true);
assert.strictEqual(boldItalic.m_wsFontName, "Liberation Sans Narrow");
assert.strictEqual(boldItalic.m_bBold && boldItalic.m_bItalic, true);

console.log("narrow-like-fonts: ok");
