/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// A click in Edit PDF restores the page's saved shapes as drawings, drawn above the page image. A saved
// page background (a page-sized fill saved, for example, after Edit Text) would then cover the rest of the
// page, so a page that has one is not restored by a click (ONLYOFFICE/DocumentServer#3940).
// Usage: node pdf/src/saved-page-background.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

function slice(file, from, to) {
	const source = fs.readFileSync(path.join(__dirname, file), "utf8");
	const start = source.indexOf(from);
	const end = source.indexOf(to, start + from.length);
	assert.ok(start >= 0 && end > start, from + " was not found in " + file);
	return source.slice(start, end);
}

// Letter page in mm
const PAGE_W = 612 * 25.4 / 72, PAGE_H = 792 * 25.4 / 72;

const box = {
	AscCommon: {
		History: { TurnOff: function() {}, TurnOn: function() {} },
		g_oTableId: { IsOn: function() { return false; }, TurnOff: function() {}, TurnOn: function() {} }
	},
	AscFonts: { CFont: function(name) { this.name = name; } },
	CPDFDoc: function() {}
};
vm.runInNewContext(slice("document.js", "CPDFDoc.prototype.IsSavedPageBackground = function", "// fCallback is called after the fonts are loaded"), box);

const visibleLn = { isVisible: function() { return true; } };
const hiddenLn = { isVisible: function() { return false; } };
function drawing(x, o) {
	o = o || {};
	return { spPr: { xfrm: { offX: x[0], offY: x[1], extX: x[2], extY: x[3], rot: o.rot || 0 } },
		pen: o.pen || null, blipFill: o.blipFill || null,
		getDocContent: function() { return o.noContent ? null : { IsEmpty: function() { return !o.text; } }; } };
}
const BACKGROUND = drawing([2.9, 5, 210.1, 269.4]);          // white page fill saved after Edit Text
const TEXT_BOX = drawing([40, 60, 80, 10], { text: true });   // a text box inserted in the editor

function doc(saved) {
	const d = new box.CPDFDoc();
	const page = { restored: null, GetRestoredShapes: function() { return this.restored; }, SetRestoredShapes: function(a) { this.restored = a; },
		GetOriginIndex: function() { return 0; } };
	d.added = [];
	d.CanRestoreSavedShapes = function() { return true; };
	d.GetPageInfo = function() { return page; };
	d.GetPageWidthMM = function() { return PAGE_W; };
	d.GetPageHeightMM = function() { return PAGE_H; };
	d.Viewer = { file: { scanSavedShapes: function() { return saved.map(function(s, i) { return { mcid: i, data: String(i) }; }); } } };
	d.ReadSavedShape = function(data) { return saved[+data]; };
	d.AddDrawing = function(o) { d.added.push(o); };
	d.page = page;
	return d;
}

// --- IsSavedPageBackground --------------------------------------------------------------------------
const isBackground = (o) => doc([]).IsSavedPageBackground(0, o);
assert.strictEqual(isBackground(BACKGROUND), true);
assert.strictEqual(isBackground(drawing([0, 0, PAGE_W, PAGE_H])), true, "exactly the page");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { pen: hiddenLn })), true, "an invisible outline is not content");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { noContent: true })), true);
assert.strictEqual(isBackground(TEXT_BOX), false, "a text box");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { text: true })), false, "a page-sized text box");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { pen: visibleLn })), false, "a page border");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { blipFill: {} })), false, "a picture");
assert.strictEqual(isBackground(drawing([2.9, 5, 210.1, 269.4], { rot: 1 })), false);
assert.strictEqual(isBackground(drawing([6.7, 36.6, 202.5, 67.1])), false, "a large box is not the page");
assert.strictEqual(isBackground(drawing([0, 0, PAGE_W, PAGE_H * 0.85])), false);
assert.strictEqual(isBackground(null), false);

// --- HitSavedShape / RestoreSavedShapes -------------------------------------------------------------
// a saved text box alone: a click on it restores it, a click elsewhere does not
let d = doc([TEXT_BOX]);
assert.strictEqual(d.HitSavedShape(0, 50, 65), true);
assert.strictEqual(d.HitSavedShape(0, 150, 200), false);
let r = d.RestoreSavedShapes(0);
assert.ok(r && r.drawings.length === 1 && d.added.length === 1);
assert.deepStrictEqual(JSON.parse(JSON.stringify(d.page.restored)), [0]);

// a saved page background next to a text box: no click restores the page, not even on the text box
d = doc([BACKGROUND, TEXT_BOX]);
assert.strictEqual(d.HitSavedShape(0, 50, 65), false, "a page with a saved background is not restored by a click");
assert.strictEqual(d.HitSavedShape(0, 150, 200), false);
assert.strictEqual(d.RestoreSavedShapes(0), null, "RestoreSavedShapes keeps it as page content");
assert.strictEqual(d.added.length, 0);
assert.strictEqual(d.page.restored, null, "no hidden shapes, nothing to undo");

// order does not matter: a background saved after the text box blocks the page too
d = doc([TEXT_BOX, BACKGROUND]);
assert.strictEqual(d.HitSavedShape(0, 50, 65), false);
assert.strictEqual(d.RestoreSavedShapes(0), null);

console.log("saved-page-background.test.js: ok");
