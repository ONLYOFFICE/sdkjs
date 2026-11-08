/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// Shapes inserted in the PDF editor are saved as /OShapes marked content. After reopening, a click in
// Edit PDF restores them as drawings. This checks the pieces that do not need a browser:
// the engine result parsing, the history change used for undo and co-editing, and the save command
// that asks the writer to remove the old copy (layout read by core MetafileToRenderer, ctPageOShapesDrop).
// Usage: node pdf/src/restore-saved-shapes.test.js
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

// --- CFile.scanSavedShapes ------------------------------------------------------------------------
const fileBox = { CFile: function() {} };
vm.runInNewContext(slice("file.js", "CFile.prototype.scanSavedShapes = function", "CFile.prototype.addPage"), fileBox);
function scan(result, nMode) {
	const file = new fileBox.CFile();
	const calls = [];
	file.nativeFile = { "scanPage": function(page, mode) { calls.push([page, mode]); if (result instanceof Error) throw result; return result; } };
	return { shapes: file.scanSavedShapes(3, nMode), calls: calls };
}
let r = scan(["0", "AAA=", "2", "BBB="], 3);
assert.deepStrictEqual(r.calls, [[3, 3]]);
assert.deepStrictEqual(JSON.parse(JSON.stringify(r.shapes)), [{ mcid: 0, data: "AAA=" }, { mcid: 2, data: "BBB=" }]);
// an engine without the mode returns no shapes, never fails
assert.deepStrictEqual(scan([], 4).shapes.length, 0);
assert.deepStrictEqual(scan(null, 3).shapes.length, 0);
assert.deepStrictEqual(scan(new Error("unknown mode"), 3).shapes.length, 0);
assert.deepStrictEqual(scan(["x", "AAA=", "1"], 3).shapes.length, 0);
assert.deepStrictEqual(new fileBox.CFile().scanSavedShapes(0, 3).length, 0);

// --- CChangesPDFDocumentRestoreSavedShapes ---------------------------------------------------------
const changeBox = {
	AscDFH: {
		CChangesBaseProperty: function(Class, Old, New) { this.Class = Class; this.Old = Old; this.New = New; },
		historyitem_PDF_Document_RestoreSavedShapes: 0x1A
	}
};
changeBox.AscDFH.CChangesBaseProperty.prototype = {};
vm.runInNewContext(slice("history/documentChanges.js", "function CChangesPDFDocumentRestoreSavedShapes", "/**") +
	"\nthis.CChangesPDFDocumentRestoreSavedShapes = CChangesPDFDocumentRestoreSavedShapes;", changeBox);
const Change = changeBox.CChangesPDFDocumentRestoreSavedShapes;
function roundTrip(oldValue, newValue) {
	const longs = [];
	new Change(null, oldValue, newValue).WriteToBinary({ WriteLong: function(v) { longs.push(v); } });
	let i = 0;
	const read = new Change(null, undefined, undefined);
	read.ReadFromBinary({ GetLong: function() { return longs[i++]; } });
	assert.strictEqual(i, longs.length, "reader did not consume the whole change");
	return { old: read.Old, new: read.New };
}
assert.deepStrictEqual(JSON.parse(JSON.stringify(roundTrip(null, [0, 3, 7]))), { old: null, new: [0, 3, 7] });
assert.deepStrictEqual(JSON.parse(JSON.stringify(roundTrip([1], null))), { old: [1], new: null });
assert.deepStrictEqual(JSON.parse(JSON.stringify(roundTrip([], [2]))), { old: [], new: [2] });
const applied = [];
new Change({ private_SetRestoredShapes: function(v) { applied.push(v); } }, null, [4]).private_SetValue([4]);
assert.deepStrictEqual(applied, [[4]]);
assert.ok(/AscDFH\.changesFactory\[AscDFH\.historyitem_PDF_Document_RestoreSavedShapes\]\s*=\s*CChangesPDFDocumentRestoreSavedShapes/.test(
	fs.readFileSync(path.join(__dirname, "history/documentChanges.js"), "utf8")), "change is not registered for co-editing");

// --- save command ---------------------------------------------------------------------------------
const metafile = fs.readFileSync(path.join(__dirname, "../../common/Drawings/Metafile.js"), "utf8");
assert.ok(/this\.ctPageOShapesDrop\s*=\s*209;/.test(metafile), "ctPageOShapesDrop must stay 209 (core MetafileToRenderer.h)");
const saveBox = { AscCommon: { CommandType: { ctPageOShapesDrop: 209 } } };
vm.runInNewContext(slice("viewer.js", "function writeRestoredShapesDrop", "function CPageInfo") +
	"\nthis.writeRestoredShapesDrop = writeRestoredShapesDrop;", saveBox);
function written(aMCID) {
	const out = [];
	const memory = { WriteByte: function(v) { out.push(["byte", v]); }, WriteLong: function(v) { out.push(["long", v]); } };
	saveBox.writeRestoredShapesDrop(memory, { GetRestoredShapes: function() { return aMCID; } });
	return out;
}
// command, length counted from the length field (as core reads it), count, MCIDs
assert.deepStrictEqual(written([0, 5]), [["byte", 209], ["long", 16], ["long", 2], ["long", 0], ["long", 5]]);
assert.deepStrictEqual(written(null), []);
assert.deepStrictEqual(written([]), []);

console.log("restore-saved-shapes.test.js ok");
