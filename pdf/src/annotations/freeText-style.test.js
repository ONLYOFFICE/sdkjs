/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// New text callouts and text comments inherit last-used formatting.
// The style object must not carry text, geometry, or annotation identity.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "freeText.js"), "utf8");
const start = source.indexOf("function isByteComponent(value)");
const end = source.indexOf('window["AscPDF"].normalizeFreeTextStyle');
assert.ok(start > 0 && end > start, "free text style helpers were not found");

const sandbox = {
	AscPDF: {
		FREE_TEXT_INTENT_TYPE: {freeText: 0, freeTextCallout: 1},
		ANNOT_COMPLEX_BORDER_TYPES: {solid: 0, dash1: 1, dash2: 2, dash3: 3, dash4: 4, dash5: 5, dash6: 6, cloud1: 7, cloud2: 8},
		ANNOT_BORDER_DASHED_VALUES: {dash1: [2, 2], dash2: [3, 3], dash3: [4, 4], dash4: [4, 3, 2, 3], dash5: [4, 3, 16, 3], dash6: [8, 4, 4, 4]},
		BORDER_TYPES: {solid: 0, dashed: 2},
		BORDER_EFFECT_STYLES: {none: 0, cloud: 1},
		getInternalAlignByPdfType: function(align) { return align + 10; }
	}
};
sandbox.window = sandbox;
vm.runInNewContext(source.slice(start, end), sandbox);

const normalize = sandbox.normalizeFreeTextStyle;
const apply = sandbox.applyFreeTextInitialStyle;
const CALLOUT = 1;
const COMMENT = 0;

function mockAnnot() {
	const calls = [];
	const annot = {
		calls: calls,
		contents: "",
		rect: [1, 2, 3, 4],
		callout: [9, 9, 9, 9, 9, 9]
	};
	["SetFillColor", "SetBorderColor", "SetBorderWidth", "SetBorderStyle", "SetDashPattern", "SetBorderEffectStyle", "SetBorderEffectIntensity", "private_UpdateRect", "SetOpacity", "SetLineEnd", "SetAlign", "SetFontFamily", "SetFontSize", "SetTextColor", "SetBold", "SetItalic", "SetContents", "SetRect", "SetCallout", "SetName", "SetAuthor", "SetSubject"].forEach(function(name) {
		annot[name] = function() {
			calls.push([name].concat([].slice.call(arguments)));
		};
	});
	annot.GetDocContent = function() {
		return {
			SetAlign: function(value) { calls.push(["content.SetAlign", value]); }
		};
	};
	return annot;
}
function names(annot) {
	return annot.calls.map(function(entry) { return entry[0]; });
}
function call(annot, name) {
	return annot.calls.filter(function(item) { return item[0] === name; });
}
function arg(annot, name, index) {
	const value = call(annot, name)[0][index == null ? 1 : index];
	return Array.isArray(value) ? Array.prototype.slice.call(value) : value;
}

// 1. No saved preferences: the helper does not replace engine defaults.
const untouched = mockAnnot();
assert.strictEqual(apply(untouched, null, CALLOUT), false);
assert.strictEqual(apply(untouched, undefined, CALLOUT), false);
assert.strictEqual(apply(untouched, "red", CALLOUT), false);
assert.strictEqual(untouched.calls.length, 0);

// 2. Custom red border and white background.
const colored = mockAnnot();
assert.strictEqual(apply(colored, {stroke: {r: 255, g: 0, b: 0}, fill: {r: 255, g: 255, b: 255}}, CALLOUT), true);
assert.deepStrictEqual(arg(colored, "SetBorderColor"), [1, 0, 0]);
assert.deepStrictEqual(arg(colored, "SetFillColor"), [1, 1, 1]);

// 3. Custom black text and 12 pt font.
const typed = mockAnnot();
assert.strictEqual(apply(typed, {textColor: {r: 0, g: 0, b: 0}, fontSize: 12, fontFamily: "Arial", bold: true, italic: false, align: 0}, CALLOUT), true);
assert.deepStrictEqual([arg(typed, "SetTextColor", 1), arg(typed, "SetTextColor", 2), arg(typed, "SetTextColor", 3)], [0, 0, 0]);
assert.strictEqual(arg(typed, "SetFontSize"), 12);
assert.strictEqual(arg(typed, "SetFontFamily"), "Arial");
assert.strictEqual(arg(typed, "SetBold"), true);
assert.strictEqual(arg(typed, "SetItalic"), false);
assert.strictEqual(arg(typed, "content.SetAlign"), 0);

// 4-7. The next annotation inherits the style. The previous one is not touched,
// text stays empty, and geometry setters are not used.
const previous = mockAnnot();
previous.contents = "1";
apply(previous, {stroke: {r: 255, g: 0, b: 0}, fontSize: 12}, CALLOUT);
const previousCalls = previous.calls.length;
const next = mockAnnot();
apply(next, {stroke: {r: 255, g: 0, b: 0}, fontSize: 12}, CALLOUT);
assert.strictEqual(previous.calls.length, previousCalls);
assert.strictEqual(previous.contents, "1");
assert.strictEqual(next.contents, "");
assert.deepStrictEqual(next.rect, [1, 2, 3, 4]);
assert.ok(names(next).indexOf("SetContents") < 0);
assert.ok(names(next).indexOf("SetRect") < 0);
assert.ok(names(next).indexOf("SetCallout") < 0);
assert.ok(names(next).indexOf("SetName") < 0);
assert.ok(names(next).indexOf("SetAuthor") < 0);
assert.deepStrictEqual(arg(next, "SetBorderColor"), [1, 0, 0]);
assert.strictEqual(arg(next, "SetFontSize"), 12);

// 8. Callout endpoint style does not apply to an ordinary text comment.
const comment = mockAnnot();
apply(comment, {lineEnd: 4, strokeWidth: 2, opacity: 40, strokeStyle: 1}, COMMENT);
assert.ok(names(comment).indexOf("SetLineEnd") < 0);
assert.strictEqual(arg(comment, "SetBorderWidth"), 2);
assert.strictEqual(arg(comment, "SetOpacity"), 0.4);
assert.deepStrictEqual(arg(comment, "SetDashPattern"), [2, 2]);
const callout = mockAnnot();
apply(callout, {lineEnd: 4}, CALLOUT);
assert.strictEqual(arg(callout, "SetLineEnd"), 4);

// 11. Invalid values are dropped. Forbidden payload fields cannot smuggle content or geometry.
assert.strictEqual(normalize(null), null);
assert.strictEqual(normalize("{"), null);
assert.strictEqual(normalize({fontSize: 0, strokeWidth: -1, opacity: 140, lineEnd: 99, align: 8, fontFamily: "  ", stroke: {r: 256, g: 0, b: 0}}), null);
assert.deepStrictEqual(JSON.parse(JSON.stringify(normalize({
	fontSize: 12,
	contents: "1",
	text: "1",
	rect: [0, 0, 1, 1],
	callout: [1, 2, 3, 4, 5, 6],
	name: "id",
	author: "someone",
	creationDate: 1,
	modDate: 2
}))), {fontSize: 12});
const poisoned = mockAnnot();
apply(poisoned, {contents: "copied", rect: [8, 8, 8, 8], fontSize: 12}, CALLOUT);
assert.strictEqual(poisoned.contents, "");
assert.deepStrictEqual(poisoned.rect, [1, 2, 3, 4]);
assert.ok(names(poisoned).indexOf("SetContents") < 0);
assert.ok(names(poisoned).indexOf("SetRect") < 0);

// Transparent fill is an intentional value, not a missing one.
const clear = mockAnnot();
apply(clear, {fill: null}, CALLOUT);
assert.strictEqual(call(clear, "SetFillColor")[0][1], undefined);

// 12-15. Creation, history, permissions, and unrelated annotation tools.
const documentSource = fs.readFileSync(path.join(__dirname, "../document.js"), "utf8");
const apiSource = fs.readFileSync(path.join(__dirname, "../../api.js"), "utf8");
assert.strictEqual(documentSource.split("AscPDF.applyFreeTextInitialStyle").length, 2);
assert.ok(documentSource.indexOf("oFreeText.SetFillColor([1, 1, 1])") > 0);
assert.ok(documentSource.indexOf("oFreeText.SetBorderColor([0, 0, 0])") > 0);
assert.ok(documentSource.indexOf("if (oStyle)") < documentSource.indexOf("oFreeText.SetInTextBox(true)"));
assert.ok(apiSource.indexOf("this.canModifyPdfAnnots() === false") > 0);
assert.ok(apiSource.indexOf("oDoc.AddFreeTextAnnot(nType, oDoc.Viewer.currentPage, oStyle)") > 0);
const addStart = apiSource.indexOf("PDFEditorApi.prototype.AddFreeTextAnnot = function");
const addEnd = apiSource.indexOf("PDFEditorApi.prototype.AddStampAnnot");
const addBody = apiSource.slice(addStart, addEnd);
assert.ok(addBody.indexOf("canModifyPdfAnnots() === false") < addBody.indexOf("oDoc.DoAction"));
assert.ok(addBody.indexOf("AscDFH.historydescription_Pdf_AddAnnot") > 0);
assert.ok(apiSource.indexOf('PDFEditorApi.prototype[\'AddFreeTextAnnot\']') > 0);
assert.ok(source.indexOf("SetContents") > 0);
assert.ok(documentSource.indexOf("contents:       ''") > 0);

console.log("freeText-style.test.js ok");
