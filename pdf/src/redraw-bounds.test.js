/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// PDF text boxes keep vertOverflow = overflow, so glyphs extend past a short
// shape. A partial redraw used to clip to the geometry only, which hides the
// rest of the line and, after a move or resize, paints only the overlap.
// Usage: node pdf/src/redraw-bounds.test.js
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

const viewer = fs.readFileSync(path.join(__dirname, "viewer.js"), "utf8");
const prototype = fs.readFileSync(path.join(__dirname, "drawings/drawingPrototype.js"), "utf8");
assert.ok(viewer.indexOf("this.bounds.r !== 0 && this.bounds.r !== 0") < 0);
assert.ok(prototype.indexOf("this.bounds.r !== 0 && this.bounds.r !== 0") < 0);
assert.ok(prototype.indexOf("AscPDF.currentDrawingRedrawBounds(this)") > 0);
assert.ok(viewer.indexOf("currentDrawingRedrawBounds(d)") > 0);

const box = {};
vm.runInNewContext(
	slice("viewer.js", "function isBoundsIntersect", "function getTextMarkupBoundsMM") +
	"\nthis.isUnsetRedrawBounds = isUnsetRedrawBounds;" +
	"\nthis.unionRedrawBounds = unionRedrawBounds;" +
	"\nthis.textOverflowBounds = textOverflowBounds;" +
	"\nthis.currentDrawingRedrawBounds = currentDrawingRedrawBounds;" +
	"\nthis.expandRedrawBounds = expandRedrawBounds;" +
	"\nthis.collectDrawingsForPartialRedraw = collectDrawingsForPartialRedraw;",
	box
);

function textBox(geometry, text, xLimit, height) {
	const content = {
		XLimit: xLimit,
		GetSummaryHeight: function() { return height; },
		GetText: function() { return text; }
	};
	return {
		bounds: geometry,
		getBoundsByDrawing: function() { return geometry; },
		// The editor places the text body at the shape origin.
		transformText: {
			TransformPointX: function(x) { return geometry.l + x; },
			TransformPointY: function(x, y) { return geometry.t + y; }
		},
		GetDocContent: function() { return content; },
		content: content
	};
}

// A brand new shape has no box yet. That must not become a 0,0 clip.
assert.strictEqual(box.isUnsetRedrawBounds({l: 0, t: 0, r: 0, b: 0}), true);
assert.strictEqual(box.currentDrawingRedrawBounds({bounds: {l: 0, t: 0, r: 0, b: 0}, getBoundsByDrawing: function() { return {l: 0, t: 0, r: 0, b: 0}; }}), null);
assert.strictEqual(box.expandRedrawBounds(null, null, 5), null);

// A box on the left edge of the page is a real box.
const atLeft = {l: 0, t: 20, r: 40, b: 28};
assert.strictEqual(box.isUnsetRedrawBounds(atLeft), false);
assert.deepStrictEqual(JSON.parse(JSON.stringify(box.expandRedrawBounds(atLeft, null, 5))), {l: -5, t: 15, r: 45, b: 33});

// The line is taller than the shape. The clip has to cover the whole line,
// not the 3mm geometry, or the tops of the glyphs disappear.
const short = textBox({l: 30, t: 40, r: 70, b: 43}, "LINE 2 NOW PAY", 55, 8);
const redraw = box.currentDrawingRedrawBounds(short);
assert.ok(redraw.t < 40, "text above the short box is inside the clip");
assert.ok(redraw.b > 43, "text below the short box is inside the clip");
assert.ok(redraw.r - redraw.l >= 55, "the full line width is inside the clip");
assert.strictEqual(short.content.GetText(), "LINE 2 NOW PAY");

// Typing a longer line widens the text bounds. The characters stay on the content.
short.content.XLimit = 90;
const longer = box.currentDrawingRedrawBounds(short);
assert.ok(longer.r - longer.l >= 90);
assert.strictEqual(short.content.GetText(), "LINE 2 NOW PAY");

// Dragging: the previous painted box and the new box are both invalidated.
const moved = box.expandRedrawBounds(
	{l: 80, t: 60, r: 160, b: 68},
	{l: 30, t: 40, r: 70, b: 48},
	5
);
assert.ok(moved.l <= 25 && moved.t <= 35, "the old position is cleared");
assert.ok(moved.r >= 165 && moved.b >= 73, "the new position is painted");

// A neighbour that only shares the text overflow is still redrawn.
const neighbour = textBox({l: 200, t: 200, r: 220, b: 210}, "other", 20, 4);
const hit = textBox({l: 30, t: 40, r: 40, b: 43}, "LINE 2 NOW PAY", 55, 8);
const selected = box.collectDrawingsForPartialRedraw([neighbour, hit], redraw);
assert.strictEqual(selected.length, 1);
assert.strictEqual(selected[0], hit);

// Resize keeps the text that was painted in the old width.
const resized = box.expandRedrawBounds(
	box.currentDrawingRedrawBounds(textBox({l: 30, t: 40, r: 140, b: 43}, "LINE 2 NOW PAY", 100, 8)),
	{l: 30, t: 40, r: 50, b: 43},
	0
);
assert.ok(resized.l <= 30 && resized.r >= 130);

console.log("redraw-bounds.test.js ok");
