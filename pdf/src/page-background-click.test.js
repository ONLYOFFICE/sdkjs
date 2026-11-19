/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Edit Text turns a fill that covers the page into a shape. A click on empty page space must go
// through it, so the click clears the selection instead of selecting a page-sized shape
// (ONLYOFFICE/DocumentServer#3940). Other shapes keep their hit tests.
// Usage: node pdf/src/page-background-click.test.js
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

// Letter page in mm, as CPDFDoc.GetPageWidthMM / GetPageHeightMM report it.
const PAGE_W = 612 * 25.4 / 72, PAGE_H = 792 * 25.4 / 72;

const baseCalls = [];
const box = {
	AscFormat: {
		CShape: function() {}
	}
};
box.AscFormat.CShape.prototype = {
	hitInInnerArea: function(x, y) { baseCalls.push(["inner", x, y]); return true; },
	hitInPath: function(x, y) { baseCalls.push(["path", x, y]); return true; },
	hitInTextRectWord: function(x, y) { baseCalls.push(["text", x, y]); return true; }
};
vm.runInNewContext("function CPdfShape() {}\nCPdfShape.prototype = Object.create(AscFormat.CShape.prototype);\n" +
	slice("drawings/shape.js", "CPdfShape.prototype.hitInTextRect = function", "CPdfShape.prototype.GetDocument = function") +
	"\nthis.CPdfShape = CPdfShape; this.IsPageBackgroundShape = IsPageBackgroundShape;", box);

const visibleLn = { isVisible: function() { return true; } };
const hiddenLn = { isVisible: function() { return false; } };
function content(empty) { return { IsEmpty: function() { return empty; } }; }

function shape(o) {
	const s = new box.CPdfShape();
	const x = o.xfrm || [2.9, 5, 210.1, 269.4]; // the sales order's white page fill
	s.spPr = { xfrm: { offX: x[0], offY: x[1], extX: x[2], extY: x[3], rot: o.rot || 0 } };
	s.pen = o.pen === undefined ? null : o.pen;
	s.blipFill = o.blipFill || null;
	const c = o.text ? content(false) : content(true);
	s.getDocContent = function() { return o.noContent ? null : c; };
	s.IsFromScan = function() { return o.fromScan !== false; };
	s.IsEditFieldShape = function() { return !!o.editField; };
	s.GetPage = function() { return 0; };
	s.GetDocument = function() {
		return o.noFile ? { Viewer: {} } : {
			Viewer: { file: { pages: [{ W: 612, H: 792 }] } },
			GetPageWidthMM: function() { return PAGE_W; },
			GetPageHeightMM: function() { return PAGE_H; }
		};
	};
	return s;
}

// --- IsPageBackgroundShape -----------------------------------------------------------------------
const isBackground = (o) => box.IsPageBackgroundShape(shape(o), PAGE_W, PAGE_H);
assert.strictEqual(isBackground({}), true, "a page-sized fill without text or outline is the background");
assert.strictEqual(isBackground({ pen: hiddenLn }), true, "an outline with no fill is not drawn");
assert.strictEqual(isBackground({ noContent: true }), true);
assert.strictEqual(isBackground({ xfrm: [0, 0, PAGE_W, PAGE_H] }), true, "exactly the page");
assert.strictEqual(isBackground({ text: true }), false, "a shape with text is content");
assert.strictEqual(isBackground({ pen: visibleLn }), false, "a page border is visible content");
assert.strictEqual(isBackground({ blipFill: {} }), false, "a picture is visible content");
assert.strictEqual(isBackground({ rot: 1 }), false);
// the bordered address box on the same page and a table row shading stay selectable
assert.strictEqual(isBackground({ xfrm: [6.7, 36.6, 202.5, 67.1] }), false);
assert.strictEqual(isBackground({ xfrm: [6.7, 153.7, 202.5, 4.4] }), false);
// a fill just under the coverage on one side is not the page background
assert.strictEqual(isBackground({ xfrm: [0, 0, PAGE_W, PAGE_H * 0.85] }), false);
assert.strictEqual(box.IsPageBackgroundShape(shape({}), 0, PAGE_H), false);
assert.strictEqual(box.IsPageBackgroundShape(null, PAGE_W, PAGE_H), false);

// --- CPdfShape hit tests ------------------------------------------------------------------------
function hits(o) {
	baseCalls.length = 0;
	const s = shape(o);
	return { inner: s.hitInInnerArea(10, 20), path: s.hitInPath(10, 20), text: s.hitInTextRect(10, 20), calls: baseCalls.slice() };
}
let r = hits({});
assert.deepStrictEqual([r.inner, r.path, r.text], [false, false, false], "a click goes through the page background");
assert.deepStrictEqual(r.calls, [], "and never reaches the shape geometry");

r = hits({ fromScan: false });
assert.deepStrictEqual([r.inner, r.path, r.text], [true, true, true], "a page-sized shape the user inserted stays selectable");
assert.deepStrictEqual(r.calls.map(c => c[0]), ["inner", "path", "text"]);
assert.deepStrictEqual(r.calls[0].slice(1), [10, 20], "the base hit test gets the click point");

r = hits({ text: true });
assert.deepStrictEqual([r.inner, r.path, r.text], [true, true, true], "reconstructed text stays clickable");

r = hits({ xfrm: [6.7, 36.6, 202.5, 67.1] });
assert.deepStrictEqual([r.inner, r.path, r.text], [true, true, true], "a large reconstructed box stays selectable");

r = hits({ noFile: true });
assert.deepStrictEqual([r.inner, r.path, r.text], [true, true, true], "without page sizes nothing changes");

r = hits({ editField: true });
assert.deepStrictEqual([r.inner, r.path, r.text], [true, true, false], "form edit shapes keep their own text rule");

console.log("page-background-click.test.js: ok");
