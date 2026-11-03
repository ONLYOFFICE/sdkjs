/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// View-only PDF annotations must not move, resize, or delete.
// Comment mode and PDF edit still may. Form filling is a separate permission
// and does not grant annotation geometry changes.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "base.js"), "utf8");
const start = source.indexOf('window["AscPDF"].annotChangeAllowed = function');
const end = source.indexOf('window["AscPDF"].CAnnotationBase');
assert.ok(start > 0 && end > start, "annotChangeAllowed was not found in base.js");

const sandbox = {
	Asc: { editor: null },
	AscCommon: { History: { UndoRedoInProgress: false, Is_On: function() { return true; } } },
	editor: null
};
sandbox.window = sandbox;
sandbox.AscPDF = {};
vm.runInNewContext(source.slice(start, end), sandbox);

const annotChangeAllowed = sandbox.AscPDF.annotChangeAllowed;
const isAnnotChangeLocked = sandbox.AscPDF.isAnnotChangeLocked;

function allow(state) {
	assert.strictEqual(annotChangeAllowed(state), true, JSON.stringify(state));
}
function deny(state) {
	assert.strictEqual(annotChangeAllowed(state), false, JSON.stringify(state));
}

allow(null);
allow(undefined);
deny({ explicit: false, canEdit: true, edit: true, comment: true });
allow({ explicit: true, pdfViewer: true, canEdit: false, formsOnly: true, edit: false, comment: false });
deny({ pdfViewer: true });
allow({ canEdit: true });
deny({ formsOnly: true, canEdit: false, edit: true, comment: true });
deny({ edit: false, comment: false });
allow({ edit: false, comment: true });
allow({ edit: true, comment: false });
allow({});

const placed = { IsUseInDocument: function() { return true; } };
const unplaced = { IsUseInDocument: function() { return false; } };

assert.strictEqual(isAnnotChangeLocked(null), false);
assert.strictEqual(isAnnotChangeLocked(unplaced), false);

sandbox.AscCommon.History.UndoRedoInProgress = true;
assert.strictEqual(isAnnotChangeLocked(placed), false);
sandbox.AscCommon.History.UndoRedoInProgress = false;

sandbox.AscCommon.History.Is_On = function() { return false; };
assert.strictEqual(isAnnotChangeLocked(placed), false);
sandbox.AscCommon.History.Is_On = function() { return true; };

sandbox.Asc.editor = {};
assert.strictEqual(isAnnotChangeLocked(placed), false);

sandbox.Asc.editor = { canModifyPdfAnnots: function() { return false; } };
assert.strictEqual(isAnnotChangeLocked(placed), true);

sandbox.Asc.editor = { canModifyPdfAnnots: function() { return true; } };
assert.strictEqual(isAnnotChangeLocked(placed), false);

const apiSource = fs.readFileSync(path.join(__dirname, "../../api.js"), "utf8");
assert.ok(apiSource.indexOf("PDFEditorApi.prototype.canModifyPdfAnnots = function()") > 0);
assert.ok(apiSource.indexOf("PDFEditorApi.prototype.asc_setPdfCanModifyAnnots = function(canModify)") > 0);
assert.ok(apiSource.indexOf('PDFEditorApi.prototype[\'asc_setPdfCanModifyAnnots\']') > 0);

console.log("annot-change-allowed.test.js ok");
