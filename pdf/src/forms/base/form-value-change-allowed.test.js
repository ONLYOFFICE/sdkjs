/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";

// View and comment must not change AcroForm values. Form filling and PDF edit may.
// Annotation permission is a separate capability and is not reused here.
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "base.js"), "utf8");
const start = source.indexOf('window["AscPDF"].formValueChangeAllowed = function');
const end = source.indexOf("CBaseField.prototype.updateSelectionState = function");
assert.ok(start > 0 && end > start, "formValueChangeAllowed was not found in base.js");

const sandbox = {
	Asc: { editor: null },
	AscCommon: { History: { UndoRedoInProgress: false, Is_On: function() { return true; } } },
	editor: null
};
sandbox.window = sandbox;
sandbox.AscPDF = {};
vm.runInNewContext(source.slice(start, end), sandbox);

const formValueChangeAllowed = sandbox.AscPDF.formValueChangeAllowed;
const isFormValueChangeLocked = sandbox.AscPDF.isFormValueChangeLocked;
const pdfSessionCanFillForms = sandbox.AscPDF.pdfSessionCanFillForms;

function allow(state) {
	assert.strictEqual(formValueChangeAllowed(state), true, JSON.stringify(state));
}
function deny(state) {
	assert.strictEqual(formValueChangeAllowed(state), false, JSON.stringify(state));
}

deny(null);
deny(undefined);
deny({});
deny({ formsOnly: false, canEdit: false });
deny({ canEdit: false });
deny({ formsOnly: false });
deny({ formsOnly: false, canEdit: false, comment: true });
allow({ formsOnly: true });
allow({ canEdit: true });
allow({ formsOnly: true, canEdit: false });
allow({ formsOnly: true, canEdit: true });

function placedField() {
	return { IsUseInDocument: function() { return true; } };
}
function api(canFill, renderer) {
	return {
		canFillPdfForms: function() { return canFill; },
		getDocumentRenderer: function() { return renderer || null; }
	};
}

sandbox.Asc.editor = api(false);
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), true);
assert.strictEqual(pdfSessionCanFillForms(), false);

sandbox.Asc.editor = api(true);
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
assert.strictEqual(pdfSessionCanFillForms(), true);

sandbox.Asc.editor = { getDocumentRenderer: function() { return null; } };
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
assert.strictEqual(pdfSessionCanFillForms(), true);

sandbox.Asc.editor = api(false);
sandbox.editor = sandbox.Asc.editor;
sandbox.AscCommon.History.UndoRedoInProgress = true;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
sandbox.AscCommon.History.UndoRedoInProgress = false;
sandbox.AscCommon.History.Is_On = function() { return false; };
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
sandbox.AscCommon.History.Is_On = function() { return true; };
assert.strictEqual(isFormValueChangeLocked({ IsUseInDocument: function() { return false; } }), false);
sandbox.Asc.editor = api(false, { IsOpenFormsInProgress: true });
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
assert.strictEqual(pdfSessionCanFillForms(), false);

const root = path.join(__dirname, "..", "..", "..", "..");
function read(rel) {
	return fs.readFileSync(path.join(root, rel), "utf8");
}

const apiSrc = read("pdf/api.js");
const docSrc = read("pdf/src/document.js");
const textSrc = read("pdf/src/forms/text.js");
const comboSrc = read("pdf/src/forms/combobox.js");
const listSrc = read("pdf/src/forms/listbox.js");
const checkSrc = read("pdf/src/forms/base/basecheckbox.js");

function has(src, needle) {
	assert.ok(src.indexOf(needle) !== -1, needle);
}

has(apiSrc, "PDFEditorApi.prototype.canFillPdfForms = function");
has(apiSrc, "AscPDF.formValueChangeAllowed");
has(apiSrc, "isFormValueChangeLocked(doc.activeForm)");
has(apiSrc, "isFormValueChangeLocked(textController)");
has(apiSrc, "isFormValueChangeLocked(oActiveForm)");
has(apiSrc, "isFormValueChangeLocked(oField)");
has(apiSrc, "this.canFillPdfForms() !== true");
assert.strictEqual(docSrc.split("pdfSessionCanFillForms() !== true").length - 1, 2);
has(docSrc, "isFormValueChangeLocked(oObject)");
has(docSrc, "isFormValueChangeLocked(oField)");
has(textSrc, "isFormValueChangeLocked(this)");
has(comboSrc, "isFormValueChangeLocked(this)");
has(listSrc, "isFormValueChangeLocked(this)");
has(checkSrc, "isFormValueChangeLocked(this)");
has(read("pdf/apiBuilder.js"), "isFormValueChangeLocked(this.Field)");
assert.strictEqual(read("pdf/apiBuilder.js").split("isFormValueChangeLocked(this.Field)").length - 1, 3);
assert.ok(source.indexOf("isFormValueChangeLocked(this)") < source.indexOf("this.private_SetValue(defValue)"));
assert.ok(source.indexOf("canModifyPdfAnnots") === -1 || source.indexOf("formValueChangeAllowed") < source.indexOf("canModifyPdfAnnots"));

const annotPath = path.join(root, "pdf/src/annotations/base.js");
if (fs.existsSync(annotPath)) {
	const annot = fs.readFileSync(annotPath, "utf8");
	if (annot.indexOf("annotChangeAllowed") !== -1) {
		const annotStart = annot.indexOf('window["AscPDF"].annotChangeAllowed = function');
		const annotEnd = annot.indexOf('window["AscPDF"].isAnnotChangeLocked = function');
		const annotSandbox = { window: {}, AscPDF: {} };
		annotSandbox.window = annotSandbox;
		vm.runInNewContext(annot.slice(annotStart, annotEnd), annotSandbox);
		assert.strictEqual(annotSandbox.AscPDF.annotChangeAllowed({ formsOnly: true, canEdit: false }), false);
		assert.strictEqual(annotSandbox.AscPDF.annotChangeAllowed({ explicit: false }), false);
		assert.strictEqual(annotSandbox.AscPDF.annotChangeAllowed({ canEdit: true }), true);
	}
}

console.log("form-value-change-allowed.test.js ok");
