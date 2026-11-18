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
const editModePressEditsValue = sandbox.AscPDF.editModePressEditsValue;

// A release within the editor's 3 px click tolerance is a click. Any more is a drag that moved the widget.
sandbox.global_mouseEvent = { X: 100, Y: 100, LastX: 102, LastY: 97, KoefPixToMM: 1 };
assert.strictEqual(editModePressEditsValue(), true);
sandbox.global_mouseEvent = { X: 100, Y: 100, LastX: 104, LastY: 100, KoefPixToMM: 1 };
assert.strictEqual(editModePressEditsValue(), false);
sandbox.global_mouseEvent = { X: 100, Y: 100, LastX: 130, LastY: 100, KoefPixToMM: 1 };
assert.strictEqual(editModePressEditsValue(), false);
sandbox.global_mouseEvent = null;
assert.strictEqual(editModePressEditsValue(), false);

// Edit mode: the first click selects a widget. A click on the widget that is already the only
// selection fills its value. Typing text keeps the caret in the value.
const editModeFieldPress = sandbox.AscPDF.editModeFieldPress;
sandbox.AscPDF.FIELD_TYPES = { button: 27, radiobutton: 28, checkbox: 29, text: 30, combobox: 31, listbox: 32, signature: 33 };
const T = sandbox.AscPDF.FIELD_TYPES;
function press(over) {
	return editModeFieldPress(Object.assign({ type: T.text, wasSelected: true, selectedCount: 1, ctrl: false, shift: false, button: 0, inForm: false, locked: false }, over));
}
assert.strictEqual(press({ wasSelected: false }), "select", "first click only selects");
assert.strictEqual(press({}), "release", "click on the selected text field fills it on release");
assert.strictEqual(press({ inForm: true }), "press", "a press in the typed value places the caret");
assert.strictEqual(press({ type: T.combobox, inForm: true }), "press");
assert.strictEqual(press({ type: T.checkbox, inForm: true }), "release", "a toggled checkbox can still be dragged");
assert.strictEqual(press({ type: T.radiobutton }), "release");
assert.strictEqual(press({ type: T.listbox }), "release");
assert.strictEqual(press({ type: T.button }), "select", "push button actions do not run in Edit");
assert.strictEqual(press({ type: T.signature }), "select");
assert.strictEqual(press({ locked: true }), "select", "View and Comment never fill");
assert.strictEqual(press({ locked: undefined }), "select", "an unknown lock state does not fill");
assert.strictEqual(press({ selectedCount: 2 }), "select", "a multi selection stays a selection");
assert.strictEqual(press({ ctrl: true }), "select");
assert.strictEqual(press({ shift: true }), "select");
assert.strictEqual(press({ button: 2 }), "select", "right click opens the menu");
assert.strictEqual(editModeFieldPress(null), "select");
sandbox.AscPDF.FIELD_TYPES = undefined;
assert.strictEqual(press({}), "select");

function allow(state) {
	assert.strictEqual(formValueChangeAllowed(state), true, JSON.stringify(state));
}
function deny(state) {
	assert.strictEqual(formValueChangeAllowed(state), false, JSON.stringify(state));
}

deny(null);
deny(undefined);
deny({});
deny({ editSession: false });
deny({ viewSession: true });
deny({ viewSession: true, editSession: true });
deny({ view: true, formsOnly: true, canEdit: true });
allow({ editSession: true });
allow({ viewSession: false, editSession: true });


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

sandbox.Asc.editor = null;
sandbox.editor = null;
assert.strictEqual(isFormValueChangeLocked(placedField()), true);
assert.strictEqual(pdfSessionCanFillForms(), false);

sandbox.Asc.editor = { getDocumentRenderer: function() { return null; } };
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), true);
assert.strictEqual(pdfSessionCanFillForms(), false);

sandbox.Asc.editor = api(false);
sandbox.editor = sandbox.Asc.editor;
sandbox.AscCommon.History.UndoRedoInProgress = true;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
sandbox.AscCommon.History.UndoRedoInProgress = false;
sandbox.AscCommon.History.Is_On = function() { return false; };
assert.strictEqual(isFormValueChangeLocked(placedField()), true, "history off is not permission");
sandbox.AscCommon.History.Is_On = function() { return true; };
assert.strictEqual(isFormValueChangeLocked({ IsUseInDocument: function() { return false; } }), false);
sandbox.Asc.editor = api(false, { IsOpenFormsInProgress: true });
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(isFormValueChangeLocked(placedField()), false);
assert.strictEqual(pdfSessionCanFillForms(), false);

function widget(value) {
	return {
		value: value,
		IsUseInDocument: function() { return true; },
		GetDocument: function() { return this._doc || null; }
	};
}
function userSet(field, next) {
	if (isFormValueChangeLocked(field))
		return false;
	field.value = next;
	return true;
}
function userActivate(field) {
	return isFormValueChangeLocked(field) !== true;
}

const text = widget("hello");
const multiline = widget("line");
const checkbox = widget(false);
const radio = widget("a");
const combo = widget("one");
const list = widget(["one"]);
const date = widget("2026-01-01");
const signature = widget(null);
sandbox.Asc.editor = api(false);
sandbox.editor = sandbox.Asc.editor;
[text, multiline, checkbox, radio, combo, list, date, signature].forEach(function(field) {
	assert.strictEqual(userActivate(field), false);
	assert.strictEqual(userSet(field, "changed"), false);
	assert.notStrictEqual(field.value, "changed");
});
assert.strictEqual(pdfSessionCanFillForms(), false);

sandbox.Asc.editor = api(true);
sandbox.editor = sandbox.Asc.editor;
assert.strictEqual(userActivate(text), true);
assert.strictEqual(userSet(text, "typed"), true);
assert.strictEqual(text.value, "typed");
assert.strictEqual(userSet(checkbox, true), true);
assert.strictEqual(checkbox.value, true);
assert.strictEqual(userSet(combo, "two"), true);
assert.strictEqual(userSet(list, ["two"]), true);
assert.strictEqual(userSet(date, "2026-02-02"), true);
assert.strictEqual(userSet(signature, "signed"), true);
assert.strictEqual(userSet(multiline, "line\n2"), true);

const session = api(true);
sandbox.Asc.editor = session;
sandbox.editor = session;
const live = widget("before");
assert.strictEqual(userSet(live, "during edit"), true);
session.canFillPdfForms = function() { return false; };
assert.strictEqual(userActivate(live), false);
assert.strictEqual(userSet(live, "during view"), false);
assert.strictEqual(live.value, "during edit");
live._doc = { _keepFormValueOnRestriction: true };
assert.strictEqual(userSet(live, "committed"), true);
assert.strictEqual(live.value, "committed");
live._doc = null;
session.canFillPdfForms = function() { return true; };
assert.strictEqual(userActivate(live), true);
assert.strictEqual(userSet(live, "after edit"), true);
assert.strictEqual(live.value, "after edit");

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
assert.strictEqual(docSrc.split("pdfSessionCanFillForms() !== true").length - 1, 3);
has(docSrc, "formEditBlocked");
has(docSrc, "editModePressEditsValue");
has(docSrc, "AscPDF.editModeFieldPress(");
has(docSrc, "wasSelected:    oCurObject === oMouseDownField && !!isFloatSelected");
// The typed value is the controller's text target, so Backspace, select all, and the clipboard edit
// the text instead of deleting or copying the widget. Undo keeps that target for the typed field only.
for (const src of [textSrc, comboSrc]) {
	has(src, "oController.selection.textSelection = this;");
	has(src, "this.selectStartPage = this.GetPage();");
	has(src, "new AscFormat.TextAddState(oController, this");
}
has(docSrc, "CPDFDoc.prototype.IsEditModeValueText = function");
assert.strictEqual(docSrc.split("this.IsEditModeValueText(").length - 1, 2, "selection restore and co-editing reload");
has(textSrc, "enterValue !== true");
has(comboSrc, "enterValue !== true");
has(listSrc, "enterValue !== true");
{
	const hit = listSrc.indexOf("this.content.Internal_GetContentPosByXY");
	assert.ok(hit > 0, "list click hit test");
	assert.ok(listSrc.lastIndexOf("this.SetNeedRecalc(true)", hit) > hit - 600, "list options are laid out before the click hit test");
	const guard = listSrc.indexOf("if (!oPara)", hit);
	assert.ok(guard > hit && guard < listSrc.indexOf("oPara.Pr.Shd", hit), "a missed list row does not throw");
}
has(checkSrc, "enterValue !== true");
has(docSrc, "_keepFormValueOnRestriction");
has(apiSrc, "this.isPdfViewer === true");
function guardBeforeFocus(src, fnName) {
	const start = src.indexOf(fnName);
	const guard = src.indexOf("isFormValueChangeLocked(this)", start);
	const focus = src.indexOf("activeForm = this", start);
	assert.ok(start !== -1 && guard !== -1 && focus !== -1 && guard < focus, fnName);
}
guardBeforeFocus(textSrc, "CTextField.prototype.onMouseDown");
guardBeforeFocus(comboSrc, "CComboBoxField.prototype.onMouseDown");
guardBeforeFocus(listSrc, "CListBoxField.prototype.onMouseDown");
guardBeforeFocus(checkSrc, "CBaseCheckBoxField.prototype.onMouseDown");
guardBeforeFocus(read("pdf/src/forms/pushbutton.js"), "CPushButtonField.prototype.onMouseDown");
has(textSrc, "CTextField.prototype.SetValue");
has(textSrc, "CTextField.prototype.private_SetValue");
has(checkSrc, "CBaseCheckBoxField.prototype.SetChecked");
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

// canFillPdfForms reads the session mode, not the restriction: annotating in an editing session
// runs under a View restriction and still fills forms. A document opened for viewing never does.
{
	const apiStart = apiSrc.indexOf("PDFEditorApi.prototype.canFillPdfForms = function");
	const apiEnd = apiSrc.indexOf("};", apiStart) + 2;
	const apiBox = { AscPDF: sandbox.AscPDF };
	vm.runInNewContext("var PDFEditorApi = function() {};" + apiSrc.slice(apiStart, apiEnd) + ";this.P = PDFEditorApi;", apiBox);
	const session = function(fields) {
		const o = Object.assign(Object.create(apiBox.P.prototype), {
			isViewMode: false, isPdfViewer: false, restrictions: 0,
			isRestrictionView: function() { return this.restrictions === 128; },
			isRestrictionForms: function() { return this.restrictions === 1; },
			canEdit: function() { return this.restrictions === 0; }
		}, fields);
		return o.canFillPdfForms();
	};
	assert.strictEqual(session({ isPdfViewer: true, restrictions: 128 }), false, "view tab");
	assert.strictEqual(session({ isPdfViewer: true, restrictions: 1 }), false, "view tab with forms bit");
	assert.strictEqual(session({ isViewMode: true }), false, "view mode");
	assert.strictEqual(session({ restrictions: 128 }), true, "edit tab, annotating");
	assert.strictEqual(session({ restrictions: 1 }), true, "edit tab, form filling");
	assert.strictEqual(session({ restrictions: 0 }), true, "edit tab, Edit PDF");
}

console.log("form-value-change-allowed.test.js ok");
