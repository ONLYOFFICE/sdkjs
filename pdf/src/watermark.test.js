/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// PDF watermarks reuse the Word watermark name and the same diagonal sizing.
// A drawing with that name is the watermark. Anything else stays selectable.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "watermark.js"), "utf8");
const sandbox = {window: {AscPDF: {}}};
vm.runInNewContext(source, sandbox);
const api = sandbox.window.AscPDF;

assert.strictEqual(api.isWatermarkName("PowerPlusWaterMarkObject"), true);
assert.strictEqual(api.isWatermarkName("PowerPlusWaterMarkObject12"), true);
assert.strictEqual(api.isWatermarkName("Shape 1"), false);
assert.strictEqual(api.isWatermarkName(""), false);
assert.strictEqual(api.isWatermarkName(undefined), false);

assert.strictEqual(api.isWatermarkDrawing(null), false);
assert.strictEqual(api.isWatermarkDrawing({}), false);
assert.strictEqual(api.isWatermarkDrawing({
	getCNvProps: function () {
		return {name: "PowerPlusWaterMarkObject3"};
	}
}), true);
assert.strictEqual(api.isWatermarkDrawing({
	getCNvProps: function () {
		return {name: "Text Box 1"};
	}
}), false);

// Letter and a square page both center the box. A wide page and a tall page
// do not share a center.
const letter = api.watermarkOffset(215.9, 279.4, 100, 20);
assert.strictEqual(letter.x, (215.9 - 100) / 2);
assert.strictEqual(letter.y, (279.4 - 20) / 2);
const square = api.watermarkOffset(200, 200, 80, 40);
assert.strictEqual(square.x, 60);
assert.strictEqual(square.y, 80);
const wide = api.watermarkOffset(400, 200, 80, 40);
const tall = api.watermarkOffset(200, 400, 80, 40);
assert.notStrictEqual(wide.x, tall.x);
assert.notStrictEqual(wide.y, tall.y);

// Horizontal automatic text uses the page width. Diagonal text uses the
// same proportion Word uses for a 45 degree watermark.
assert.strictEqual(api.watermarkAutoWidth(200, 100, 50, 10, false), 200);
const diagonal = api.watermarkAutoWidth(200, 100, 50, 10, true);
assert.strictEqual(diagonal, Math.SQRT2 * 100 / (10 / 50 + 1));
assert.ok(diagonal < 200);
assert.strictEqual(api.watermarkAutoWidth(200, 100, 0, 10, true), 200);

assert.strictEqual(api.watermarkRotation(false), 0);
assert.strictEqual(api.watermarkRotation(true), Math.PI * (360 - 45) / 180);

assert.strictEqual(api.watermarkImageOpacity(null), 128);
assert.strictEqual(api.watermarkImageOpacity(undefined), 128);
assert.strictEqual(api.watermarkImageOpacity(255), 255);
assert.strictEqual(api.watermarkImageOpacity(64), 64);

// A half-red watermark shows on paper and leaves the artwork alone.
const page = new Uint8ClampedArray([255, 255, 255, 255, 20, 60, 40, 255]);
const mark = new Uint8ClampedArray([255, 0, 0, 128, 255, 0, 0, 128]);
api.compositeImageWatermarkBehind(page, mark);
assert.strictEqual(page[0], 255);
assert.strictEqual(page[1], 127);
assert.strictEqual(page[2], 127);
assert.strictEqual(page[4], 20);
assert.strictEqual(page[5], 60);
assert.strictEqual(page[6], 40);

// Hit testing and the text stream skip the watermark and keep searching.
function sourceHas(rel, needle) {
	const text = fs.readFileSync(path.join(__dirname, rel), "utf8");
	assert.ok(text.indexOf(needle) !== -1, rel + " is missing: " + needle);
}
sourceHas("file.js", "if (AscPDF.isWatermarkDrawing(drawing))");
sourceHas("GraphicObjects.js", "if (AscPDF.isWatermarkDrawing(object))");
sourceHas("../../common/Drawings/DrawingObjectsHandlers.js", "AscPDF.isWatermarkDrawing(drawing)");
sourceHas("watermark.js", "AscFormat.text_fit_No");
sourceHas("../../common/Drawings/TextDrawer.js", "oKeepFill.transparent");
