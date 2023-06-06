/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
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
	const getRangeWithData = function (ws, data) {
		let range = ws.getRange4(0, 0);

		range.fillData(data);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];

		return ws;
	};

	function checkUndoRedo(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	}

	function compareData (assert, range, data, desc) {
		for (let i = range.r1; i <= range.r2; i++) {
			for (let j = range.c1; j <= range.c2; j++) {
				let rangeVal = ws.getCell3(i, j);
				let dataVal = data[i - range.r1][j - range.c1];
				assert.strictEqual(rangeVal.getValue(), dataVal, desc + " compare " + rangeVal.getName());
			}
		}
	}

	QUnit.test("Test: \"Move rows/cols\"", function (assert) {
		let testData = [
			["row1col1", "row1col2", "row1col3", "row1col4", "row1col5", "row1col6"],
			["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"],
			["row3col1", "row3col2", "row3col3", "row3col4", "row3col5", "row3col6"],
			["row4col1", "row4col2", "row4col3", "row4col4", "row4col5", "row4col6"],
			["row5col1", "row5col2", "row5col3", "row5col4", "row5col5", "row5col6"],
			["row6col1", "row6col2", "row6col3", "row6col4", "row6col5", "row6col6"],
			["row7col1", "row7col2", "row7col3", "row7col4", "row7col5", "row7col6"],
			["row8col1", "row8col2", "row8col3", "row8col4", "row8col5", "row8col6"],
			["row9col1", "row9col2", "row9col3", "row9col4", "row9col5", "row9col6"]
		];

		let range = ws.getRange4(0, 0);
		range.fillData(testData);

		//***COLS***
		//***move without ctrl***
		//***move without shift***
		//move from 1 to 3 cols
		wsView.activeMoveRange = getRange(3, 0, 3, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 1, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: false, shiftKey: false};
		wsView.applyMoveRangeHandle();

		let rangeCompare = getRange(0, 1, 5, 1);
		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "", "row2col3", "row2col2", "row2col5", "row2col6"]], desc);
		}, " move_col_1 ");

		//move from 1-2 to 3-4 cols
		wsView.activeMoveRange = getRange(3, 0, 4, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 2, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: false, shiftKey: false};
		wsView.applyMoveRangeHandle();

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "", "", "row2col2", "row2col3", "row2col6"]], desc);
		}, " move_col_2 ");


		//***move with ctrl***
		//***move without shift***
		//move from 1 to 3 cols
		wsView.activeMoveRange = getRange(3, 0, 3, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 1, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: true, shiftKey: false};
		wsView.applyMoveRangeHandle(true);

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col2", "row2col5", "row2col6"]], desc);
		}, " move_col_1 ");

		//move from 1-2 to 3-4 cols
		wsView.activeMoveRange = getRange(3, 0, 4, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 2, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: true, shiftKey: false};
		wsView.applyMoveRangeHandle(true);

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col2", "row2col3", "row2col6"]], desc);
		}, " move_col_2 ");


		//***move without ctrl***
		//***move with shift***
		wsView.activeMoveRange = getRange(3, 0, 3, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 1, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: false, shiftKey: true, colByX: 3};
		wsView.applyMoveRangeHandle();

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "", "row2col3", "row2col4", "row2col2", "row2col5"]], desc);
		}, " move_col_1 ");

		//***move with ctrl***
		//***move with shift***
		wsView.activeMoveRange = getRange(3, 0, 3, AscCommon.gc_nMaxRow);
		ws.selectionRange.ranges = [getRange(1, 0, 1, AscCommon.gc_nMaxRow)];
		wsView.startCellMoveRange = getRange(1, 0, 1, 0);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: true, shiftKey: true, colByX: 3};
		wsView.applyMoveRangeHandle(true);

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col5", "row2col6"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row2col1", "row2col2", "row2col3", "row2col4", "row2col2", "row2col5"]], desc);
		}, " move_col_1 ");


		//***ROWS***
		//***move without ctrl***
		//***move without shift***
		//move from 1 to 3 rows
		wsView.activeMoveRange = getRange(0, 3, AscCommon.gc_nMaxCol, 3);
		ws.selectionRange.ranges = [getRange(0, 1, AscCommon.gc_nMaxCol, 1)];
		wsView.startCellMoveRange = getRange(0, 1, 0, 1);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: false, shiftKey: false};
		wsView.applyMoveRangeHandle();

		rangeCompare = getRange(1, 0, 1, 8);
		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row4col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row1col2"], [""], ["row3col2"], ["row2col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, " move_row_1 ");

		//***move with ctrl***
		//***move without shift***
		//move from 1 to 3 rows
		wsView.activeMoveRange = getRange(0, 3, AscCommon.gc_nMaxCol, 3);
		ws.selectionRange.ranges = [getRange(0, 1, AscCommon.gc_nMaxCol, 1)];
		wsView.startCellMoveRange = getRange(0, 1, 0, 1);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: true, shiftKey: false};
		wsView.applyMoveRangeHandle(true);

		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row4col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row2col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, " move_row_2 ");

		//***move without ctrl***
		//***move with shift***
		//move from 1 to 3 rows
		wsView.activeMoveRange = getRange(0, 3, AscCommon.gc_nMaxCol, 3);
		ws.selectionRange.ranges = [getRange(0, 1, AscCommon.gc_nMaxCol, 1)];
		wsView.startCellMoveRange = getRange(0, 1, 0, 1);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: false, shiftKey: true, rowByY: 3};
		wsView.applyMoveRangeHandle();

		rangeCompare = getRange(1, 0, 1, 8);
		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row4col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row1col2"], [""], ["row3col2"], ["row4col2"], ["row2col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"]], desc);
		}, " move_row_3 ");


		//***move with ctrl***
		//***move with shift***
		//move from 1 to 3 rows
		wsView.activeMoveRange = getRange(0, 3, AscCommon.gc_nMaxCol, 3);
		ws.selectionRange.ranges = [getRange(0, 1, AscCommon.gc_nMaxCol, 1)];
		wsView.startCellMoveRange = getRange(0, 1, 0, 1);
		wsView.startCellMoveRange.colRowMoveProps = {ctrlKey: true, shiftKey: true, rowByY: 3};
		wsView.applyMoveRangeHandle(true);

		rangeCompare = getRange(1, 0, 1, 8);
		checkUndoRedo(function (desc) {
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row4col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"], ["row9col2"]], desc);
		}, function (desc){
			compareData(assert, rangeCompare, [["row1col2"], ["row2col2"], ["row3col2"], ["row4col2"], ["row2col2"], ["row5col2"], ["row6col2"], ["row7col2"], ["row8col2"]], desc);
		}, " move_row_4 ");
	});

	QUnit.module("Sheet structure");
});
