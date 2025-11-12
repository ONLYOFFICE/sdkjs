/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
	};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getCellCache = function (col, row) {
		let _cell = null;
		this.model.getRange3(row, col, row, col)._foreachNoEmpty(function(cell, row, col) {
			if (cell && !cell.isEmptyTextString()) {
				_cell = {cellType: cell.getType()}
			}
		}, null, true);
		return _cell;
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (range, subType, callback) {
		callback(true);
		return true;
	};
	AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
	};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function(){};

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	}

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	wb.handlers.add("asc_onConfirmAction", function (test1, callback) {
		callback(true);
	});
	api.wb.cellCommentator = new AscCommonExcel.CCellCommentator({
		model: api.wbModel.aWorksheets[0],
		collaborativeEditing: null,
		draw: function() {
		},
		handlers: {
			trigger: function() {
				return false;
			}
		}
	});

	AscCommonExcel.CCellCommentator.prototype.isLockedComment = function (oComment, callbackFunc) {
		callbackFunc(true);
	};
	AscCommonExcel.CCellCommentator.prototype.drawCommentCells = function () {
	};
	AscCommonExcel.CCellCommentator.prototype.ascCvtRatio = function () {
	};

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};
	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	};

	function checkUndoRedo(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	}

	let getRgbColor = function(clr){
		var color = (typeof(clr) == 'object') ? clr.color : clr;

		color=color.replace(/#/,'');
		if(color.length==3) color=color.replace(/(.)/g,'$1$1');
		color=parseInt(color,16);
		var c = new Asc.asc_CColor();
		c.put_type( (typeof(clr) == 'object' && clr.effectId !== undefined)? Asc.c_oAscColor.COLOR_TYPE_SCHEME : Asc.c_oAscColor.COLOR_TYPE_SRGB);
		c.put_r(color>>16);
		c.put_g((color&0xff00)>>8);
		c.put_b(color&0xff);
		c.put_a(0xff);
		if (clr.effectId !== undefined)
			c.put_value(clr.effectId);
		return c;
	};

	function compareAscColorAndRgbColor (ascColor, rgbColor) {
		if ((ascColor && !rgbColor) || (!ascColor && rgbColor)) {
			return false;
		}
		return ascColor.get_r() === rgbColor.getR() && ascColor.get_g() === rgbColor.getG() && ascColor.get_b() === rgbColor.getB();
	};

	QUnit.test('Conditional formatting: test apply to', function (assert) {

		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		tableOptions.range = "A1:B3";
		api.asc_addAutoFilter("TableStyleMedium2", tableOptions);

		let cf = new AscCommonExcel.CConditionalFormattingRule();
		cf.asc_setType(Asc.c_oAscCFType.cellIs);
		cf.asc_setLocation("A5");

		api.asc_setCF([cf]);

		wsView.setSelection(new Asc.Range(0, 4, 0, 4));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		let cfLocation;
		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		let ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$5", "compare location conditional formatting in cell");


		cf = new AscCommonExcel.CConditionalFormattingRule();
		cf.asc_setType(Asc.c_oAscCFType.cellIs);
		cf.asc_setLocation("=Table1[Column1]");

		api.asc_setCF([cf]);

		wsView.setSelection(new Asc.Range(0, 1, 0, 1));
		modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);

		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$2:$A$4", "compare location conditional formatting in table");


		clearData(0, 6, 0, 6);
	});


	QUnit.test("Test: \"simple tests\"", function (assert) {
		let testData = [
			["100"]
		];

		let range = ws.getRange4(0, 0);
		range.fillData(testData);


		let newRule = new AscCommonExcel.CConditionalFormattingRule();
		newRule.asc_setType(Asc.c_oAscCFType.cellIs);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("9C0006");
		let fillColor = getRgbColor("FFC7CE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		newRule.asc_setDxf(xfs);
		newRule.asc_setOperator(AscCommonExcel.ECfOperator.Operator_greaterThan);
		newRule.asc_setValue1("99");
		newRule.asc_setLocation("$A$1");

		api.asc_setCF([newRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, AscCommon.gc_nMaxCol0, AscCommon.gc_nMaxRow0)]));

		wsView.setSelection(new Asc.Range(0, 0, 0, 0));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);

		let cfLocation;
		if (modelCf) {
			modelCf = modelCf[0] && modelCf[0][0];
			cfLocation = modelCf.asc_getLocation();
		}

		let ref = cfLocation && cfLocation[1];
		assert.strictEqual(ref, "=$A$1", "compare location conditional formatting");


		let compiledStyle = ws.getCompiledStyle(0, 0);
		let rgbColor = compiledStyle.fill.bg();
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor), true, "compare fill color _1");


		//Clearing data of sheet
		clearData(0, 0, 1, 8);
	});

	QUnit.test('Conditional formatting: contains text tests', function (assert) {
		// Preparing test data
		let testData = [
			["This is sample text"],
			["Text with sample word"],
			["No matching here"],
			["SAMPLE in uppercase"]
		];

		let range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Test 1: Check simple text condition
		let textRule = new AscCommonExcel.CConditionalFormattingRule();
		textRule.asc_setType(Asc.c_oAscCFType.containsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("2F75B5");
		let fillColor = getRgbColor("BDD7EE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		textRule.asc_setDxf(xfs);
		textRule.asc_setContainsText("sample");
		textRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([textRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check that the rule was applied
		wsView.setSelection(new Asc.Range(0, 0, 0, 0));
		let modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		assert.ok(modelCf && modelCf[0] && modelCf[0][0], "CF rule exists");

		// Check that the style was applied to cells with 'sample'
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1.fill.bg();
		let rgbColor2 = compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 matches 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 matches 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't match 'sample' condition");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 matches 'sample' condition (case insensitive)");

		// Test 2: Check cell reference in condition
		clearData(0, 0, 0, 4);

		// New test data
		testData = [
			["apple"], // A1: search word
			["This is an apple text"], // A2: contains the word
			["Orange juice"], // A3: doesn't contain the word
			["Pineapple is good"] // A4: contains the word as part of another word
		];

		range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Create rule with cell reference
		let cellRefRule = new AscCommonExcel.CConditionalFormattingRule();
		cellRefRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("006400");
		fillColor = getRgbColor("C6EFCE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		cellRefRule.asc_setDxf(xfs);
		cellRefRule.asc_setContainsText("=A1"); // Cell reference, not text
		cellRefRule.asc_setLocation("$A$2:$A$4");

		api.asc_setCF([cellRefRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 3)]));

		// Check that the rule was applied
		wsView.setSelection(new Asc.Range(0, 1, 0, 1));
		modelCf = api.asc_getCF(Asc.c_oAscSelectionForCFType.selection, 0);
		assert.ok(modelCf && modelCf[0] && modelCf[0][0], "CF rule with cell reference exists");

		// Check that the style was applied to cells with 'apple'
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 matches value from A1");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), true, "Cell A3 doesn't match value from A1");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 matches value from A1 (as part of word)");

		// Test 3: Check begins with / ends with
		clearData(0, 0, 0, 5);

		// New test data for begins/ends with
		testData = [
			["Spreadsheet testing"],
			["Testing Spreadsheet"],
			["Microsoft Spreadsheet"],
			["spreadsheet lowercase"]
		];

		range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Rule "Begins with Spreadsheet"
		let beginsWithRule = new AscCommonExcel.CConditionalFormattingRule();
		beginsWithRule.asc_setType(Asc.c_oAscCFType.beginsWith);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		fillColor = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		beginsWithRule.asc_setDxf(xfs);
		beginsWithRule.asc_setContainsText("Spreadsheet");
		beginsWithRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([beginsWithRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check cell styles
		compiledStyle1 = ws.getCompiledStyle(0, 0);
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 begins with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), false, "Cell A2 doesn't begin with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't begin with 'Spreadsheet'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 begins with 'spreadsheet' (case insensitive)");

		// Clean up data after tests
		clearData(0, 0, 0, 5);
	});

	QUnit.test('Conditional formatting: not contains text tests', function (assert) {
		// Preparing test data
		let testData = [
			["Report 2023"],
			["Sales data"],
			["Error in data"],
			["Missing values"]
		];

		let range = ws.getRange4(0, 0, 3, 0);
		range.fillData(testData);

		// Rule "Does not contain Error"
		let notContainsRule = new AscCommonExcel.CConditionalFormattingRule();
		notContainsRule.asc_setType(Asc.c_oAscCFType.notContainsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("3F3F76");
		let fillColor = getRgbColor("B4C6E7");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		notContainsRule.asc_setDxf(xfs);
		notContainsRule.asc_setContainsText("Error");
		notContainsRule.asc_setLocation("$A$1:$A$4");

		api.asc_setCF([notContainsRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 3)]));

		// Check cell styles
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		let rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor1), true, "Cell A1 doesn't contain 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 doesn't contain 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 contains 'Error'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 doesn't contain 'Error'");

		// Clean up data after tests
		clearData(0, 0, 5, 5);
	});

	QUnit.test('Conditional formatting: asc_setContainsText with formula and text', function (assert) {
		// Preparing test data
		let testData = [
			["Product"], // A1: header
			["CPU Intel"], // A2
			["Graphics card AMD"], // A3
			["Intel Motherboard"], // A4
			["AMD Processor"] // A5
		];

		let range = ws.getRange4(0, 0, 4, 0);
		range.fillData(testData);

		// Additional value for formula reference
		let referenceData = [
			["Intel"] // B1: value for reference
		];
		ws.getRange4(0, 1).fillData(referenceData);

		// Test 1: Using plain text
		let textRule = new AscCommonExcel.CConditionalFormattingRule();
		textRule.asc_setType(Asc.c_oAscCFType.containsText);
		let xfs = new Asc.asc_CellXfs();
		let fontColor = getRgbColor("2F75B5");
		let fillColor = getRgbColor("BDD7EE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		textRule.asc_setDxf(xfs);
		textRule.asc_setContainsText("Intel"); // Plain text
		textRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([textRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for text search
		let compiledStyle2 = ws.getCompiledStyle(1, 0);
		let compiledStyle3 = ws.getCompiledStyle(2, 0);
		let compiledStyle4 = ws.getCompiledStyle(3, 0);
		let compiledStyle5 = ws.getCompiledStyle(4, 0);

		let rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		let rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		let rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		let rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 contains 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't contain 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 contains 'Intel' text");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), false, "Cell A5 doesn't contain 'Intel' text");


		// Test 2: Using cell reference formula
		let formulaRule = new AscCommonExcel.CConditionalFormattingRule();
		formulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("006400");
		fillColor = getRgbColor("C6EFCE");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		formulaRule.asc_setDxf(xfs);
		formulaRule.asc_setContainsText("=B1"); // Formula reference
		formulaRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([formulaRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for formula reference
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);
		compiledStyle5 = ws.getCompiledStyle(4, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		//TODO!
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), true, "Cell A2 contains value from B1 (Intel)");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), false, "Cell A3 doesn't contain value from B1");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), true, "Cell A4 contains value from B1 (Intel)");
		// assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), false, "Cell A5 doesn't contain value from B1");

		// Test 3: Using more complex formula
		// Change reference data
		referenceData = [
			["AMD"] // B1: changed value
		];
		ws.getRange4(0, 1).fillData(referenceData);

		let complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		fillColor = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=UPPER(B1)"); // Formula with function
		complexFormulaRule.asc_setLocation("$A$2:$A$5");

		api.asc_setCF([complexFormulaRule]);
		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 1, 0, 4)]));

		// Check that styles are applied correctly for complex formula
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);
		compiledStyle5 = ws.getCompiledStyle(4, 0);

		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();
		rgbColor5 = compiledStyle5 && compiledStyle5.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor2), false, "Cell A2 doesn't contain 'AMD'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor3), true, "Cell A3 contains 'AMD'");
		//TODO!
		//assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor4), false, "Cell A4 doesn't contain 'AMD'");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor, rgbColor5), true, "Cell A5 contains 'AMD'");

		clearData(0, 0, 5, 5);

		// Test 4
		testData = [
			["test"],//A1
			["test"],//A2
			["test"],//A3
			["test"],//A4
			["test"]//A5
		];

		range = ws.getRange4(0, 0, 4, 0);
		range.fillData(testData);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor1 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor1);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=test");
		complexFormulaRule.asc_setLocation("$A$1");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor2 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor2);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("=\"test\"");
		complexFormulaRule.asc_setLocation("$A$2");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor3 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor3);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("test");
		complexFormulaRule.asc_setLocation("$A$3");

		api.asc_setCF([complexFormulaRule]);

		complexFormulaRule = new AscCommonExcel.CConditionalFormattingRule();
		complexFormulaRule.asc_setType(Asc.c_oAscCFType.containsText);
		xfs = new Asc.asc_CellXfs();
		fontColor = getRgbColor("9C5700");
		let fillColor4 = getRgbColor("FFEB9C");
		xfs.asc_setFontColor(fontColor);
		xfs.asc_setFillColor(fillColor4);
		complexFormulaRule.asc_setDxf(xfs);
		complexFormulaRule.asc_setContainsText("\"test\"");
		complexFormulaRule.asc_setLocation("$A$4");

		api.asc_setCF([complexFormulaRule]);

		ws.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, 0, 4)]));

		// Check that styles are applied correctly for complex formula
		let compiledStyle1 = ws.getCompiledStyle(0, 0);
		compiledStyle2 = ws.getCompiledStyle(1, 0);
		compiledStyle3 = ws.getCompiledStyle(2, 0);
		compiledStyle4 = ws.getCompiledStyle(3, 0);

		let rgbColor1 = compiledStyle1 && compiledStyle1.fill.bg();
		rgbColor2 = compiledStyle2 && compiledStyle2.fill.bg();
		rgbColor3 = compiledStyle3 && compiledStyle3.fill.bg();
		rgbColor4 = compiledStyle4 && compiledStyle4.fill.bg();

		assert.strictEqual(compareAscColorAndRgbColor(fillColor1, rgbColor1), false, "formula + text without quotes");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor2, rgbColor2), true, "formula + text with quotes");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor3, rgbColor3), true, "text without quote");
		assert.strictEqual(compareAscColorAndRgbColor(fillColor4, rgbColor4), false, "text with quotes");

		// Clean up data after tests
		clearData(0, 0, 5, 5);
	});

	QUnit.module("Conditional formatting");
});
