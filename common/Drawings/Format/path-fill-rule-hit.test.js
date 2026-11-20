/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Path.hitInInnerArea must test a point with the fill rule the path is painted with. An even-odd frame
// (outer and inner contour drawn the same way round) is painted as a border with an empty hole; a click
// in the hole is not a click on the shape, a click on the border is. Nonzero paths are unchanged.
// Usage: node common/Drawings/Format/path-fill-rule-hit.test.js
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

const box = { moveTo: 0, lineTo: 1, arcTo: 2, bezier3: 3, bezier4: 4, close: 5, ArcToOnCanvas: function() {}, Path: function() {} };
vm.runInNewContext(slice("Path.js", "Path.prototype.hitInInnerArea = function", "Path.prototype.hitInPath = function"), box);

// A canvas context that records the path and answers isPointInPath geometrically (winding / crossings).
function context() {
	const rings = [];
	let ring = null;
	return {
		calls: [],
		beginPath: function() { rings.length = 0; ring = null; },
		moveTo: function(x, y) { ring = [[x, y]]; rings.push(ring); },
		lineTo: function(x, y) { ring.push([x, y]); },
		closePath: function() {},
		isPointInPath: function(x, y, rule) {
			this.calls.push(rule);
			let winding = 0, crossings = 0;
			rings.forEach(function(r) {
				for (let i = 0; i < r.length; i++) {
					const a = r[i], b = r[(i + 1) % r.length];
					if ((a[1] <= y) !== (b[1] <= y)) {
						const xi = a[0] + (y - a[1]) * (b[0] - a[0]) / (b[1] - a[1]);
						if (xi > x) { crossings++; winding += b[1] > a[1] ? 1 : -1; }
					}
				}
			});
			return rule === "evenodd" ? crossings % 2 === 1 : winding !== 0;
		}
	};
}

// The bordered box from DocumentServer#3883 / #3940: outer 0..200 x 0..100, inner 2..198 x 2..98, same direction.
function frame(fill) {
	const p = new box.Path();
	p.fill = fill;
	p.ArrPathCommand = [
		{ id: 0, X: 0, Y: 0 }, { id: 1, X: 200, Y: 0 }, { id: 1, X: 200, Y: 100 }, { id: 1, X: 0, Y: 100 }, { id: 5 },
		{ id: 0, X: 2, Y: 2 }, { id: 1, X: 198, Y: 2 }, { id: 1, X: 198, Y: 98 }, { id: 1, X: 2, Y: 98 }, { id: 5 }
	];
	return p;
}
function hit(fill, x, y) {
	const ctx = context();
	const result = frame(fill).hitInInnerArea(ctx, x, y);
	return { result: result, rules: ctx.calls };
}

// even-odd: the hole is empty, the border is the shape
assert.deepStrictEqual(hit("evenOdd", 100, 50), { result: false, rules: ["evenodd"] }, "a click in the hole of an even-odd frame misses it");
assert.strictEqual(hit("evenOdd", 1, 50).result, true, "its left border is hit");
assert.strictEqual(hit("evenOdd", 100, 99).result, true, "its bottom border is hit");
assert.strictEqual(hit("evenOdd", 250, 50).result, false, "outside is outside");

// nonzero (and paths without a rule) paint the hole and keep hitting it, as before
assert.deepStrictEqual(hit("norm", 100, 50), { result: true, rules: ["nonzero"] });
assert.deepStrictEqual(hit(null, 100, 50), { result: true, rules: ["nonzero"] });
assert.deepStrictEqual(hit(undefined, 1, 50), { result: true, rules: ["nonzero"] });

// an unfilled path has no inner area and never asks the canvas
assert.deepStrictEqual(hit("none", 100, 50), { result: false, rules: [] });

// a single filled even-odd rectangle is still solid
const solid = new box.Path();
solid.fill = "evenOdd";
solid.ArrPathCommand = [{ id: 0, X: 0, Y: 0 }, { id: 1, X: 50, Y: 0 }, { id: 1, X: 50, Y: 50 }, { id: 1, X: 0, Y: 50 }, { id: 5 }];
assert.strictEqual(solid.hitInInnerArea(context(), 25, 25), true, "a solid even-odd fill is hit inside");

console.log("path-fill-rule-hit.test.js: ok");
