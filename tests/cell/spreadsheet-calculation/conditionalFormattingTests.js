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
		api._openDocument(AscCommon.getEmpty());
		api._openOnClient();
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
	api._openDocument(AscCommon.getEmpty());
	api._openOnClient();
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


	QUnit.module("Conditional formatting");
});
