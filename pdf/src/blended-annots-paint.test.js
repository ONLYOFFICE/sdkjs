/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// ONLYOFFICE/DocumentServer#3923. A page with a Multiply ink annotation paints
// that annotation into the cached page bitmap. Each partial redraw left one
// saved state with its dirty clip on that bitmap. The clips stacked up, so text
// typed into a text box was painted only where it overlapped the first dirty
// rectangle, until a zoom recreated the bitmap.
// Usage: node pdf/src/blended-annots-paint.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const viewer = fs.readFileSync(path.join(__dirname, "viewer.js"), "utf8");
const start = viewer.indexOf("CHtmlPage.prototype._drawBlendedAnnotsOnCtx = function");
const end = viewer.indexOf("CHtmlPage.prototype._paintDrawings = function", start);
assert.ok(start > 0 && end > start, "_drawBlendedAnnotsOnCtx was not found in viewer.js");

// The stub below saves in init() because the real CGraphics.init() does.
const graphics = fs.readFileSync(path.join(__dirname, "../../word/Drawing/Graphics.js"), "utf8");
const initStart = graphics.indexOf("CGraphics.prototype.init = function");
const initBody = graphics.slice(initStart, graphics.indexOf("CGraphics.prototype.", initStart + 10));
assert.ok(initStart > 0 && initBody.indexOf("this.m_oContext.save();") > 0,
	"CGraphics.init no longer saves the context; update this test");

// A 2D context that keeps the state stack and the clips on it.
function FakeContext(width, height) {
	this.canvas = {width: width, height: height};
	this.stack = [];
	this.clips = [];
	this.pendingRect = null;
	this.globalAlpha = 1;
	this.paintedWithClips = [];
}
FakeContext.prototype.save = function() { this.stack.push(this.clips.slice()); };
FakeContext.prototype.restore = function() { if (this.stack.length) this.clips = this.stack.pop(); };
FakeContext.prototype.beginPath = function() { this.pendingRect = null; };
FakeContext.prototype.rect = function(x, y, w, h) { this.pendingRect = {x: x, y: y, w: w, h: h}; };
FakeContext.prototype.clip = function() { this.clips.push(this.pendingRect); };

const sandbox = {
	CHtmlPage: function() {},
	AscPDF: {
		PRINT_CONTENT_TYPES: {docAndMarkups: 0},
		CPDFGraphics: function() {
			this.Init = function() {};
			this.SetCurPage = function() {};
		}
	},
	AscCommon: {
		g_fontManager: {},
		CGraphics: function() {
			this.init = function(ctx) { ctx.save(); };
			this.setEndGlobalAlphaColor = function() {};
			this.transform = function() {};
		}
	}
};
vm.runInNewContext(viewer.slice(start, end), sandbox);

function page(annots) {
	const p = new sandbox.CHtmlPage();
	p._annotsCompositedOnPage = function() { return annots; };
	p.getPDFDoc = function() {
		return {GetPageWidthMM: function() { return 215.9; }, GetPageHeightMM: function() { return 279.4; }};
	};
	p.file = {getPageWidth: function() { return 816; }, getPageHeight: function() { return 1056; }};
	return p;
}
function inkAnnot(ctx) {
	return {DrawOnPage: function() { ctx.paintedWithClips.push(ctx.clips.length); }};
}

// Partial redraws, as typing does: one per keystroke, each with a new dirty box.
{
	const ctx = new FakeContext(816, 1056);
	const viewerPage = page([inkAnnot(ctx)]);
	const boxes = [
		{l: 33, t: 90, r: 75, b: 107},
		{l: 33, t: 90, r: 76, b: 111},
		{l: 33, t: 90, r: 76, b: 115}
	];
	boxes.forEach(function(box) {
		viewerPage._drawBlendedAnnotsOnCtx(0, ctx, box);
		assert.strictEqual(ctx.stack.length, 0, "the page bitmap keeps a saved state after a partial redraw");
		assert.strictEqual(ctx.clips.length, 0, "the dirty clip stays on the page bitmap");
	});
	// The annotation itself was drawn inside exactly one clip, the current one.
	assert.deepStrictEqual(ctx.paintedWithClips, [1, 1, 1]);
}

// A full redraw draws without a clip and also leaves the context as it was.
{
	const ctx = new FakeContext(816, 1056);
	page([inkAnnot(ctx)])._drawBlendedAnnotsOnCtx(0, ctx, null);
	assert.strictEqual(ctx.stack.length, 0);
	assert.strictEqual(ctx.clips.length, 0);
	assert.deepStrictEqual(ctx.paintedWithClips, [0]);
}

// No composited annotations: the context is not touched at all.
{
	const ctx = new FakeContext(816, 1056);
	page([])._drawBlendedAnnotsOnCtx(0, ctx, {l: 0, t: 0, r: 10, b: 10});
	assert.strictEqual(ctx.stack.length, 0);
	assert.strictEqual(ctx.clips.length, 0);
}

console.log("blended-annots-paint.test.js ok");
