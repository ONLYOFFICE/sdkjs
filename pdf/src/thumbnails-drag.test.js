/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// PDF thumbnails: a click must select pages, and only a real pointer move may start a page drag.
// A trackpad force click fires mousemove while the pressure changes, with little or no movement.
// Runs the real thumbnails.js mouse handlers and the real CPDFDoc.MovePages with a stubbed layout.
// Usage: node pdf/src/thumbnails-drag.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const PAGE_COUNT = 11;
const PAGE_W = 100;
const PAGE_H = 100;
const PAGE_STEP = 120; // 20 px gap between thumbnails

function noopContext() {
	return new Proxy({}, {
		get: function(target, name) {
			if (name in target) return target[name];
			return name === "measureText" ? function() { return { width: 10 }; } : function() {};
		},
		set: function(target, name, value) { target[name] = value; return true; }
	});
}

function load() {
	const mouse = { X: 0, Y: 0, Button: 0, IsLocked: false, LockMouse: function() { this.IsLocked = true; } };
	const editor = { restricted: false, isRestrictionView: function() { return this.restricted; }, getPDFDoc: function() { return { UpdateInterface: function() {} }; } };
	const sandbox = {
		Asc: { editor: editor },
		AscCommon: {
			global_mouseEvent: mouse,
			AscBrowser: { retinaPixelRatio: 2, zoom: 1 },
			Cursors: { Grabbing: "grabbing" },
			GlobalSkin: {},
			stopEvent: function() {},
			check_MouseDownEvent: function(e) { mouse.X = e.pageX; mouse.Y = e.pageY; mouse.Button = e.button; },
			check_MouseMoveEvent: function(e) { mouse.X = e.pageX; mouse.Y = e.pageY; },
			check_MouseUpEvent: function(e) { mouse.X = e.pageX; mouse.Y = e.pageY; mouse.Button = e.button; mouse.IsLocked = false; }
		},
		AscDFH: {},
		console: console
	};
	sandbox.window = sandbox;
	sandbox.global_mouseEvent = mouse;
	vm.runInNewContext(fs.readFileSync(path.join(__dirname, "thumbnails.js"), "utf8"), sandbox);

	// Real MovePages on a page list, so a reorder is checked against the editor's own algorithm.
	const docSource = fs.readFileSync(path.join(__dirname, "document.js"), "utf8");
	const start = docSource.indexOf("CPDFDoc.prototype.MovePages = function");
	const end = docSource.indexOf("CPDFDoc.prototype.SetPageRotate");
	assert.ok(start > 0 && end > start, "MovePages was not found in document.js");
	const docSandbox = { CPDFDoc: function() {} };
	vm.runInNewContext(docSource.slice(start, end), docSandbox);

	return { sandbox: sandbox, mouse: mouse, editor: editor, CPDFDoc: docSandbox.CPDFDoc };
}

const env = load();
const CDocument = env.sandbox.AscCommon.ThumbnailsControl;
assert.strictEqual(typeof CDocument, "function", "ThumbnailsControl was not exported");

function createThumbnails() {
	const doc = new env.CPDFDoc();
	doc.order = [];
	for (let i = 0; i < PAGE_COUNT; i++) doc.order.push(i);
	doc.actions = 0;
	doc.GetPagesCount = function() { return this.order.length; };
	doc.RemovePage = function(i) { return this.order.splice(i, 1)[0]; };
	doc.AddPage = function(i, page) { this.order.splice(i, 0, page); };
	doc.DoAction = function(fn) { this.actions++; fn(); };
	doc.Viewer = { navigateToPage: function() {}, scrollMaxX: 0 };

	const th = Object.create(CDocument.prototype);
	th.viewer = { doc: doc, navigateToPage: function() {}, scrollMaxX: 0 };
	th.coordsOffset = { x: 0, y: 0 };
	th.scrollY = 0;
	th.betweenH = PAGE_STEP - PAGE_H;
	th.pages = new Array(PAGE_COUNT);
	th.selectedPages = [0];
	th.hoverPage = -1;
	th.canvas = { style: {} };
	th.dragCanvas = { style: { display: "none" } };
	th.dragCtx = noopContext();
	th.repaint = function() {};
	th._paint = function() {};
	th.getDrawingPage = function(num) {
		if (num == null || num < 0 || num >= PAGE_COUNT) return null;
		return { num: num, page: {}, pageRect: { x: 0, y: num * PAGE_STEP, w: PAGE_W, h: PAGE_H } };
	};
	th.getPageByCoords = function(x, y) {
		const num = Math.floor(y / PAGE_STEP);
		if (x < 0 || x > PAGE_W || num < 0 || num >= PAGE_COUNT || y - num * PAGE_STEP > PAGE_H) return null;
		return th.getDrawingPage(num);
	};
	return th;
}

function center(num) {
	return { x: PAGE_W / 2, y: num * PAGE_STEP + PAGE_H / 2 };
}

function event(x, y, mods, extra) {
	mods = mods || {};
	return Object.assign({ pageX: x, pageY: y, button: 0, buttons: 1, shiftKey: !!mods.shift, ctrlKey: !!mods.ctrl, metaKey: !!mods.meta, preventDefault: function() {} }, extra);
}

// Press on a page, move by each [dx, dy] offset, release at the last position.
function press(th, page, moves, mods) {
	const c = center(page);
	let x = c.x, y = c.y;
	const seen = { dragging: false, ghost: false };
	th.onMouseDown(event(x, y, mods));
	(moves || []).forEach(function(m) {
		x = c.x + m[0];
		y = c.y + m[1];
		th.onMouseMove(event(x, y, mods));
		seen.dragging = seen.dragging || !!th.isDragging;
		seen.ghost = seen.ghost || th.dragCanvas.style.display === "block";
	});
	th.onMouseUp(event(x, y, mods, { buttons: 0 }));
	return seen;
}

// Press on one page and release over another, moving in small steps.
function drag(th, from, to, mods) {
	const a = center(from), z = center(to);
	const moves = [];
	for (let i = 1; i <= 10; i++) moves.push([0, (z.y - a.y) * i / 10]);
	return press(th, from, moves, mods);
}

const ORIGINAL = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
const results = [];
function test(name, fn) {
	env.editor.restricted = false;
	fn(createThumbnails());
	results.push(name);
}
function assertIdle(th) {
	assert.strictEqual(!!th.isDragging, false, "still dragging");
	assert.strictEqual(!!th.pendingDrag, false, "press still pending");
	assert.notStrictEqual(th.dragCanvas.style.display, "block", "drag ghost still shown");
}
function assertUnchanged(th) {
	assert.deepStrictEqual(th.viewer.doc.order, ORIGINAL);
	assert.strictEqual(th.viewer.doc.actions, 0, "a reorder action was recorded");
}

test("1. click without movement selects the page", function(th) {
	const seen = press(th, 3, []);
	assert.deepStrictEqual(th.selectedPages, [3]);
	assert.strictEqual(seen.dragging, false);
	assertUnchanged(th);
	assertIdle(th);
});

test("2. click with 1-3 px movement selects the page and does not drag", function(th) {
	const seen = press(th, 3, [[1, 0], [2, 1], [3, 3]]);
	assert.deepStrictEqual(seen, { dragging: false, ghost: false });
	assert.deepStrictEqual(th.selectedPages, [3]);
	assertUnchanged(th);
	assertIdle(th);
});

test("3. force click (mousemove at the same or nearby point) selects the page", function(th) {
	const seen = press(th, 4, [[0, 0], [0, 0], [1, 0], [0, 1], [-1, 1], [2, -2], [0, 0]]);
	assert.deepStrictEqual(seen, { dragging: false, ghost: false });
	assert.deepStrictEqual(th.selectedPages, [4]);
	assertUnchanged(th);
	assertIdle(th);
});

test("threshold is the slide thumbnail distance: 10 px stays a click, 11 px drags", function(th) {
	assert.strictEqual(press(th, 2, [[10, 0], [0, 10], [-10, -10]]).dragging, false);
	assert.deepStrictEqual(th.selectedPages, [2]);
	assert.strictEqual(press(th, 2, [[0, 11]]).dragging, true);
});

test("4. Cmd-click toggles the page and does not drag, even with jitter", function(th) {
	press(th, 2, []);
	let seen = press(th, 5, [[1, 1], [3, 0]], { meta: true });
	assert.strictEqual(seen.dragging, false);
	assert.deepStrictEqual(th.selectedPages, [2, 5]);
	seen = press(th, 5, [[0, 2]], { meta: true });
	assert.strictEqual(seen.dragging, false);
	assert.deepStrictEqual(th.selectedPages, [2]);
	assertUnchanged(th);
	assertIdle(th);
});

test("Ctrl-click toggles the page like Cmd-click (Windows/Linux, and macOS button 0 + ctrlKey)", function(th) {
	press(th, 2, []);
	assert.strictEqual(press(th, 6, [[1, 0]], { ctrl: true }).dragging, false);
	assert.deepStrictEqual(th.selectedPages, [2, 6]);
	assertUnchanged(th);
});

test("5. Shift-click selects the range and does not drag, even past the threshold", function(th) {
	press(th, 1, []);
	const seen = press(th, 4, [[2, 2], [15, 15]], { shift: true });
	assert.deepStrictEqual(seen, { dragging: false, ghost: false });
	assert.deepStrictEqual(th.selectedPages, [1, 2, 3, 4]);
	assertUnchanged(th);
});

test("modifier press dragged across other pages does not reorder", function(th) {
	press(th, 1, []);
	assert.strictEqual(drag(th, 2, 6, { meta: true }).dragging, false);
	assert.strictEqual(drag(th, 3, 7, { shift: true }).dragging, false);
	assertUnchanged(th);
});

test("6. moving past the threshold starts the drag and shows the ghost", function(th) {
	const seen = press(th, 1, [[0, 5], [0, 12]]);
	assert.deepStrictEqual(seen, { dragging: true, ghost: true });
	assertIdle(th);
});

test("7. valid drop moves the page", function(th) {
	const seen = drag(th, 1, 3);
	assert.strictEqual(seen.dragging, true);
	assert.deepStrictEqual(th.viewer.doc.order, [0, 2, 3, 1, 4, 5, 6, 7, 8, 9, 10]);
	assert.strictEqual(th.viewer.doc.actions, 1);
	assertIdle(th);
});

test("8. invalid drop (off the pages, or back onto the dragged page) keeps the order", function(th) {
	const a = center(2);
	th.onMouseDown(event(a.x, a.y));
	th.onMouseMove(event(a.x, a.y + 30));
	th.onMouseMove(event(PAGE_W + 50, a.y + 30)); // off every thumbnail
	assert.strictEqual(th.isDragging, true);
	assert.strictEqual(th.dropInsertPosition, -1);
	th.onMouseUp(event(PAGE_W + 50, a.y + 30, null, { buttons: 0 }));
	assertUnchanged(th);
	assertIdle(th);

	press(th, 2, [[0, 40], [0, 0]]);
	assertUnchanged(th);
	assertIdle(th);
});

test("9. release after a minor move that leaves the page selects the pressed page", function(th) {
	// Page 5 bottom edge is 50 px below its center; start 45 px down and drift 9 px into the gap.
	const c = center(5);
	th.onMouseDown(event(c.x, c.y + 45));
	th.onMouseMove(event(c.x, c.y + 54));
	assert.strictEqual(th.getPageByCoords(c.x, c.y + 54), null, "release point should be in the gap");
	th.onMouseUp(event(c.x, c.y + 54, null, { buttons: 0 }));
	assert.deepStrictEqual(th.selectedPages, [5]);
	assertUnchanged(th);
	assertIdle(th);
});

test("10. restricted view never starts a drag", function(th) {
	env.editor.restricted = true;
	const seen = drag(th, 1, 4);
	assert.deepStrictEqual(seen, { dragging: false, ghost: false });
	assertUnchanged(th);
	assertIdle(th);
});

test("11. jittery multi-select keeps every page for Extract Selected Pages", function(th) {
	press(th, 7, [[1, 0]]);
	press(th, 8, [[0, 1], [1, 1]], { meta: true });
	press(th, 10, [[1, 0], [0, 0]], { meta: true });
	press(th, 9, [[2, 1]], { meta: true });
	// asc_extractSelectedPages() with no argument reads getSelectedPages() and sorts it.
	assert.deepStrictEqual(th.getSelectedPages().slice().sort(function(a, b) { return a - b; }), [7, 8, 9, 10]);
	assertUnchanged(th);
	const api = fs.readFileSync(path.join(__dirname, "../api.js"), "utf8");
	assert.ok(/asc_extractSelectedPages[\s\S]{0,400}getSelectedPages\(\)/.test(api), "extraction no longer reads the thumbnail selection");
});

test("12. selected group moves together on an intentional drag", function(th) {
	press(th, 1, []);
	press(th, 2, [], { meta: true });
	const seen = drag(th, 1, 3);
	assert.strictEqual(seen.dragging, true);
	assert.deepStrictEqual(th.viewer.doc.order, [0, 3, 1, 2, 4, 5, 6, 7, 8, 9, 10]);
	assert.strictEqual(th.viewer.doc.actions, 1);
});

test("lost mouseup: hover with no button held does not drag, next press starts clean", function(th) {
	const c = center(2);
	th.onMouseDown(event(c.x, c.y));
	// mouseup happened outside the frame; the pointer comes back with no button held.
	th.onMouseMove(event(c.x, c.y + 250, null, { buttons: 0 }));
	assertIdle(th);
	th.onMouseDown(event(c.x, c.y + 240));
	th.onMouseMove(event(c.x, c.y + 280));
	assert.strictEqual(th.isDragging, true);
	th.onMouseDown(event(center(6).x, center(6).y)); // mouseup lost again mid-drag
	assert.strictEqual(th.isDragging, false, "old drag carried into the new press");
	assert.strictEqual(th.dragCanvas.style.display, "none");
	assert.strictEqual(th.dropInsertPosition, -1);
	th.onMouseUp(event(center(6).x, center(6).y, null, { buttons: 0 }));
	assert.deepStrictEqual(th.selectedPages, [6]);
	assertUnchanged(th);
});

results.forEach(function(name) { console.log("ok - " + name); });
console.log("thumbnails-drag.test.js ok (" + results.length + " tests)");
