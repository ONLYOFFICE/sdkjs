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
	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	}

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
	function autofillData (assert, rangeTo, expectedData, description) {
		for (let i = rangeTo.r1; i <= rangeTo.r2; i++) {
			for (let j = rangeTo.c1; j <= rangeTo.c2; j++) {
				let rangeToVal = ws.getCell3(i, j);
				let dataVal = expectedData[i - rangeTo.r1][j - rangeTo.c1];
				assert.strictEqual(rangeToVal.getValue(), dataVal, `${description} Cell: ${rangeToVal.getName()}, Value: ${dataVal}`);
			}
		}
	}
	function reverseAutofillData (assert, rangeTo, expectedData, description) {
		for (let i = rangeTo.r1; i >= rangeTo.r2; i--) {
			for (let j = rangeTo.c1; j >= rangeTo.c2; j--) {
				let rangeToVal = ws.getCell3(i, j);
				let dataVal = expectedData[Math.abs(i - rangeTo.r1)][Math.abs(j - rangeTo.c1)];
				assert.strictEqual(rangeToVal.getValue(), dataVal, `${description} Cell: ${rangeToVal.getName()}, Value: ${dataVal}`);
			}
		}
	}
	function getAutoFillRange(wsView, c1To, r1To, c2To, r2To, nHandleDirection, nFillHandleArea) {
		wsView.fillHandleArea = nFillHandleArea;
		wsView.fillHandleDirection = nHandleDirection;
		wsView.activeFillHandle = getRange(c1To, r1To, c2To, r2To);
		wsView.applyFillHandle(0,0,false);

		return wsView;
	}
	function updateDataToUpCase (aExpectedData) {
		return aExpectedData.map (function (expectedData) {
			if (Array.isArray(expectedData)) {
				return [expectedData[0].toUpperCase()]
			}
			return expectedData.toUpperCase();
		});
	}
	function updateDataToLowCase (aExpectedData) {
		return aExpectedData.map (function (expectedData) {
			if (Array.isArray(expectedData)) {
				return [expectedData[0].toLowerCase()]
			}
			return expectedData.toLowerCase();
		});
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

		clearData(0, 0, AscCommon.gc_nMaxCol, AscCommon.gc_nMaxRow);
	});

	QUnit.test('Autofill - Asc horizontal sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday'],
			['SUNDAY'],
			['sunday'],
			['SuNdAy'],
			['SUnDaY'],
			['sUnDaY'],
			['suNDay'],
			['Sun'],
			['SUN'],
			['sun'],
			['SuN'],
			['SUn'],
			['suN'],
			['sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Monday', 'Tuesday', 'Wednesday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Mon', 'Tue', 'Wed'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0,0,0,0)];
		wsView = getAutoFillRange(wsView, 0, 0, 3, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(1, 0, 3, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(0,1,0,1)];
		wsView = getAutoFillRange(wsView, 0, 1, 3, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 1, 3, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(0,2,0,2)];
		wsView = getAutoFillRange(wsView, 0, 2, 3, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 2, 3, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(0,3,0,3)];
		wsView = getAutoFillRange(wsView, 0, 3, 3, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 3, 3, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(0,4,0,4)];
		wsView = getAutoFillRange(wsView, 0, 4, 3, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 4, 3, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(0,5,0,5)];
		wsView = getAutoFillRange(wsView, 0, 5, 3, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 5, 3, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(0,6,0,6)];
		wsView = getAutoFillRange(wsView, 0, 6, 3, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 6, 3, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(0,7,0,7)];
		wsView = getAutoFillRange(wsView, 0, 7, 3, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 7, 3, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(0,8,0,8)];
		wsView = getAutoFillRange(wsView, 0, 8, 3, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 8, 3, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(0,9,0,9)];
		wsView = getAutoFillRange(wsView, 0, 9, 3, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 9, 3, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(0,10,0,10)];
		wsView = getAutoFillRange(wsView, 0, 10, 3, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 10, 3, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(0,11,0,11)];
		wsView = getAutoFillRange(wsView, 0, 11, 3, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 11, 3, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(0,12,0,12)];
		wsView = getAutoFillRange(wsView, 0, 12, 3, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 12, 3, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(0,13,0,13)];
		wsView = getAutoFillRange(wsView, 0, 13, 3, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, 13, 3, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 3, 13);
	});
	QUnit.test('Autofill - Reverse horizontal sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday'],
			['SUNDAY'],
			['sunday'],
			['SuNdAy'],
			['SUnDaY'],
			['sUnDaY'],
			['suNDay'],
			['Sun'],
			['SUN'],
			['sun'],
			['SuN'],
			['SUn'],
			['suN'],
			['sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Saturday', 'Friday', 'Thursday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Sat', 'Fri', 'Thu'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(3,0,3,0)];
		wsView = getAutoFillRange(wsView, 3, 0, 0, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 0, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(3,1,3,1)];
		wsView = getAutoFillRange(wsView,3, 1, 0, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 0, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(3,2,3,2)];
		wsView = getAutoFillRange(wsView, 3, 2, 0, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 0, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3,3,3,3)];
		wsView = getAutoFillRange(wsView, 3, 3, 0, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 0, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(3,4,3,4)];
		wsView = getAutoFillRange(wsView, 3, 4, 0, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 0, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(3,5,3,5)];
		wsView = getAutoFillRange(wsView, 3, 5, 0, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 0, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(3,6,3,6)];
		wsView = getAutoFillRange(wsView, 3, 6, 0, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 0, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(3,7,3,7)];
		wsView = getAutoFillRange(wsView, 3, 7, 0, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 0, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(3,8,3,8)];
		wsView = getAutoFillRange(wsView, 3, 8, 0, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 0, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(3,9,3,9)];
		wsView = getAutoFillRange(wsView, 3, 9, 0, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 0, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(3,10,3,10)];
		wsView = getAutoFillRange(wsView, 3, 10, 0, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 0, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(3,11,3,11)];
		wsView = getAutoFillRange(wsView, 3, 11, 0, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 0, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(3,12,3,12)];
		wsView = getAutoFillRange(wsView, 3, 12, 0, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 0, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(3,13,3,13)];
		wsView = getAutoFillRange(wsView, 3, 13, 0, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 0, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 3, 13);
	});
	QUnit.test('Autofill - Asc horizontal even sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'Tuesday'],
			['SUNDAY', 'TUESDAY'],
			['sunday', 'tuesday'],
			['SuNdAy', 'TuEsDaY'],
			['SUnDaY', 'TUeSdAy'],
			['sUnDaY', 'tUeSdAy'],
			['suNDay', 'tuESdaY'],
			['Sun', 'Tue'],
			['SUN', 'TUE'],
			['sun', 'tue'],
			['SuN', 'TuE'],
			['SUn', 'TUe'],
			['suN', 'tuE'],
			['sUn', 'tUe']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Thursday', 'Saturday', 'Monday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Thu', 'Sat', 'Mon'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0,0,1,0)];
		wsView = getAutoFillRange(wsView, 0, 0, 4, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(0,1,1,1)];
		wsView = getAutoFillRange(wsView, 0, 1, 4, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 4, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(0,2,1,2)];
		wsView = getAutoFillRange(wsView, 0, 2, 4, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 4, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(0,3,1,3)];
		wsView = getAutoFillRange(wsView, 0, 3, 4, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 4, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(0,4,1,4)];
		wsView = getAutoFillRange(wsView, 0, 4, 4, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 4, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(0,5,1,5)];
		wsView = getAutoFillRange(wsView, 0, 5, 4, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 4, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(0,6,1,6)];
		wsView = getAutoFillRange(wsView, 0, 6, 4, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 4, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(0,7,1,7)];
		wsView = getAutoFillRange(wsView, 0, 7, 4, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 4, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(0,8,1,8)];
		wsView = getAutoFillRange(wsView, 0, 8, 4, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 4, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(0,9,1,9)];
		wsView = getAutoFillRange(wsView, 0, 9, 4, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 4, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(0,10,1,10)];
		wsView = getAutoFillRange(wsView, 0, 10, 4, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 4, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(0,11,1,11)];
		wsView = getAutoFillRange(wsView, 0, 11, 4, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 4, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(0,12,1,12)];
		wsView = getAutoFillRange(wsView, 0, 12, 4, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 4, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(0,13,1,13)];
		wsView = getAutoFillRange(wsView, 0, 13, 4, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 4, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Asc horizontal odd sequence: Days of the weeks', function (assert) {
		let testData = [
			['Monday', 'Wednesday'],
			['MONDAY', 'WEDNESDAY'],
			['monday', 'wednesday'],
			['MoNdAy', 'WeDnEsDaY'],
			['MOnDAy', 'WEdNEsDAy'],
			['mOnDaY', 'wEdNeSdAy'],
			['moNDay', 'weDNesDAy'],
			['Mon', 'Wed'],
			['MON', 'WED'],
			['mon', 'wed'],
			['MoN', 'WeD'],
			['MOn', 'WEd'],
			['moN', 'weD'],
			['mOn', 'wEd']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Friday', 'Sunday', 'Tuesday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Fri', 'Sun', 'Tue'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0,0,1,0)];
		wsView = getAutoFillRange(wsView, 0, 0, 4, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Monday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(0,1,1,1)];
		wsView = getAutoFillRange(wsView, 0, 1, 4, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 4, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Monday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(0,2,1,2)];
		wsView = getAutoFillRange(wsView, 0, 2, 4, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 4, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Monday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(0,3,1,3)];
		wsView = getAutoFillRange(wsView, 0, 3, 4, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 4, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Monday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(0,4,1,4)];
		wsView = getAutoFillRange(wsView, 0, 4, 4, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 4, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Monday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(0,5,1,5)];
		wsView = getAutoFillRange(wsView, 0, 5, 4, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 4, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Monday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(0,6,1,6)];
		wsView = getAutoFillRange(wsView, 0, 6, 4, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 4, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Monday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(0,7,1,7)];
		wsView = getAutoFillRange(wsView, 0, 7, 4, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 4, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Mon');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(0,8,1,8)];
		wsView = getAutoFillRange(wsView, 0, 8, 4, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 4, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Mon');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(0,9,1,9)];
		wsView = getAutoFillRange(wsView, 0, 9, 4, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 4, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Mon');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(0,10,1,10)];
		wsView = getAutoFillRange(wsView, 0, 10, 4, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 4, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Mon');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(0,11,1,11)];
		wsView = getAutoFillRange(wsView, 0, 11, 4, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 4, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Mon');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(0,12,1,12)];
		wsView = getAutoFillRange(wsView, 0, 12, 4, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 4, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Mon');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(0,13,1,13)];
		wsView = getAutoFillRange(wsView, 0, 13, 4, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 4, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Mon');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Reverse horizontal even sequence: Days of the weeks', function (assert) {
		let testData = [
			['Friday', 'Sunday'],
			['FRIDAY', 'SUNDAY'],
			['friday', 'sunday'],
			['FrIdAy', 'SuNdAy'],
			['FRiDaY', 'SUnDaY'],
			['fRiDaY', 'sUnDaY'],
			['frIDay', 'suNDay'],
			['Fri', 'Sun'],
			['FRI', 'SUN'],
			['fri', 'sun'],
			['FrI', 'SuN'],
			['FRi', 'SUn'],
			['frI', 'suN'],
			['fRi', 'sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Wednesday', 'Monday', 'Saturday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Wed', 'Mon', 'Sat'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(3,0,4,0)];
		wsView = getAutoFillRange(wsView, 4, 0, 0, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 0, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(3,1,4,1)];
		wsView = getAutoFillRange(wsView,4, 1, 0, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 0, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(3,2,4,2)];
		wsView = getAutoFillRange(wsView, 4, 2, 0, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 0, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3,3,4,3)];
		wsView = getAutoFillRange(wsView, 4, 3, 0, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 0, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(3,4,4,4)];
		wsView = getAutoFillRange(wsView, 4, 4, 0, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 0, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(3,5,4,5)];
		wsView = getAutoFillRange(wsView, 4, 5, 0, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 0, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(3,6,4,6)];
		wsView = getAutoFillRange(wsView, 4, 6, 0, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 0, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(3,7,4,7)];
		wsView = getAutoFillRange(wsView, 4, 7, 0, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 0, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(3,8,4,8)];
		wsView = getAutoFillRange(wsView, 4, 8, 0, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 0, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(3,9,4,9)];
		wsView = getAutoFillRange(wsView, 4, 9, 0, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 0, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(3,10,4,10)];
		wsView = getAutoFillRange(wsView, 4, 10, 0, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 0, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(3,11,4,11)];
		wsView = getAutoFillRange(wsView, 4, 11, 0, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 0, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(3,12,4,12)];
		wsView = getAutoFillRange(wsView, 4, 12, 0, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 0, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(3,13,4,13)];
		wsView = getAutoFillRange(wsView, 4, 13, 0, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 0, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Reverse horizontal odd sequence: Days of the weeks', function (assert) {
		let testData = [
			['Thursday', 'Saturday'],
			['THURSDAY', 'SATURDAY'],
			['thursday', 'saturday'],
			['ThUrSdAy', 'SaTuRdAy'],
			['THurSDay', 'SAtuRDay'],
			['tHuRsDaY', 'sAtUrDaY'],
			['thURsdAY', 'saTUrdAY'],
			['Thu', 'Sat'],
			['THU', 'SAT'],
			['thu', 'sat'],
			['ThU', 'SaT'],
			['THu', 'SAt'],
			['thU', 'saT'],
			['tHu', 'sAt']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Tuesday', 'Sunday', 'Friday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Tue', 'Sun', 'Fri'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(3,0,4,0)];
		wsView = getAutoFillRange(wsView, 4, 0, 0, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 0, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Saturday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(3,1,4,1)];
		wsView = getAutoFillRange(wsView,4, 1, 0, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 0, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Saturday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(3,2,4,2)];
		wsView = getAutoFillRange(wsView, 4, 2, 0, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 0, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Saturday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3,3,4,3)];
		wsView = getAutoFillRange(wsView, 4, 3, 0, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 0, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Saturday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(3,4,4,4)];
		wsView = getAutoFillRange(wsView, 4, 4, 0, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 0, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Saturday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(3,5,4,5)];
		wsView = getAutoFillRange(wsView, 4, 5, 0, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 0, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Saturday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(3,6,4,6)];
		wsView = getAutoFillRange(wsView, 4, 6, 0, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 0, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Saturday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(3,7,4,7)];
		wsView = getAutoFillRange(wsView, 4, 7, 0, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 0, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sat');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(3,8,4,8)];
		wsView = getAutoFillRange(wsView, 4, 8, 0, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 0, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sat');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(3,9,4,9)];
		wsView = getAutoFillRange(wsView, 4, 9, 0, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 0, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sat');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(3,10,4,10)];
		wsView = getAutoFillRange(wsView, 4, 10, 0, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 0, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sat');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(3,11,4,11)];
		wsView = getAutoFillRange(wsView, 4, 11, 0, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 0, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sat');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(3,12,4,12)];
		wsView = getAutoFillRange(wsView, 4, 12, 0, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 0, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sat');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(3,13,4,13)];
		wsView = getAutoFillRange(wsView, 4, 13, 0, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 0, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sat');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Asc horizontal sequence of full and short names: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'Sun'],
			['SUNDAY', 'SUN'],
			['sunday', 'sun'],
			['SuNdAy', 'SuN'],
			['SUnDaY', 'SUn'],
			['sUnDaY', 'sUn'],
			['suNDay', 'suN'],
			['Sun', 'Sunday'],
			['SUN', 'SUNDAY'],
			['sun', 'sunday'],
			['SuN', 'SuNdAy'],
			['SUn', 'SUnDaY'],
			['suN', 'suNDay'],
			['sUn', 'sUnDaY']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Monday', 'Mon', 'Tuesday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Mon', 'Monday', 'Tue'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0,0,1,0)];
		wsView = getAutoFillRange(wsView, 0, 0, 4, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(0,1,1,1)];
		wsView = getAutoFillRange(wsView, 0, 1, 4, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 4, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(0,2,1,2)];
		wsView = getAutoFillRange(wsView, 0, 2, 4, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 4, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(0,3,1,3)];
		wsView = getAutoFillRange(wsView, 0, 3, 4, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 4, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(0,4,1,4)];
		wsView = getAutoFillRange(wsView, 0, 4, 4, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 4, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(0,5,1,5)];
		wsView = getAutoFillRange(wsView, 0, 5, 4, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 4, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(0,6,1,6)];
		wsView = getAutoFillRange(wsView, 0, 6, 4, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 4, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(0,7,1,7)];
		wsView = getAutoFillRange(wsView, 0, 7, 4, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 4, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(0,8,1,8)];
		wsView = getAutoFillRange(wsView, 0, 8, 4, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 4, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(0,9,1,9)];
		wsView = getAutoFillRange(wsView, 0, 9, 4, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 4, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(0,10,1,10)];
		wsView = getAutoFillRange(wsView, 0, 10, 4, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 4, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(0,11,1,11)];
		wsView = getAutoFillRange(wsView, 0, 11, 4, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 4, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(0,12,1,12)];
		wsView = getAutoFillRange(wsView, 0, 12, 4, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 4, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(0,13,1,13)];
		wsView = getAutoFillRange(wsView, 0, 13, 4, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 4, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Reverse horizontal sequence of full and short names: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'Sun'],
			['SUNDAY', 'SUN'],
			['sunday', 'sun'],
			['SuNdAy', 'SuN'],
			['SUnDaY', 'SUn'],
			['sUnDaY', 'sUn'],
			['suNDay', 'suN'],
			['Sun', 'Sunday'],
			['SUN', 'SUNDAY'],
			['sun', 'sunday'],
			['SuN', 'SuNdAy'],
			['SUn', 'SUnDaY'],
			['suN', 'suNDay'],
			['sUn', 'sUnDaY']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Sat', 'Saturday', 'Fri'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Saturday', 'Sat', 'Friday'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.

		let range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(3,0,4,0)];
		wsView = getAutoFillRange(wsView, 4, 0, 0, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(2, 0, 0, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(3,1,4,1)];
		wsView = getAutoFillRange(wsView,4, 1, 0, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 1, 0, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(3,2,4,2)];
		wsView = getAutoFillRange(wsView, 4, 2, 0, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 2, 0, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3,3,4,3)];
		wsView = getAutoFillRange(wsView, 4, 3, 0, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 3, 0, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(3,4,4,4)];
		wsView = getAutoFillRange(wsView, 4, 4, 0, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 4, 0, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(3,5,4,5)];
		wsView = getAutoFillRange(wsView, 4, 5, 0, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 5, 0, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(3,6,4,6)];
		wsView = getAutoFillRange(wsView, 4, 6, 0, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 6, 0, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(3,7,4,7)];
		wsView = getAutoFillRange(wsView, 4, 7, 0, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 7, 0, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(3,8,4,8)];
		wsView = getAutoFillRange(wsView, 4, 8, 0, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 8, 0, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(3,9,4,9)];
		wsView = getAutoFillRange(wsView, 4, 9, 0, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 9, 0, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(3,10,4,10)];
		wsView = getAutoFillRange(wsView, 4, 10, 0, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 10, 0, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(3,11,4,11)];
		wsView = getAutoFillRange(wsView, 4, 11, 0, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 11, 0, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(3,12,4,12)];
		wsView = getAutoFillRange(wsView, 4, 12, 0, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 12, 0, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(3,13,4,13)];
		wsView = getAutoFillRange(wsView, 4, 13, 0, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, 13, 0, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill - Asc vertical sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Monday'], ['Tuesday'], ['Wednesday']];
		let expectedDataShortCapitalized = [['Mon'], ['Tue'], ['Wed']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 0;
		const r2To = 3;
		const r1 = 0;
		const r2 = 0;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 1, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 1, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 1, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 1, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 1, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 1, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 1, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 1, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 1, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 1, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 1, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 1, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 1, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 1, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 3);
	});
	QUnit.test('Autofill - Reverse vertical sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Saturday'], ['Friday'], ['Thursday']];
		let expectedDataShortCapitalized = [['Sat'], ['Fri'], ['Thu']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 3;
		const r2To = 0;
		const r1 = 3;
		const r2 = 3;
		let range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To - 1, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 1, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 1, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 1, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 1, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 1, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 1, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 1, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 1, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 1, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 1, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 1, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 1, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 1, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 3);
	});
	QUnit.test('Autofill - Asc vertical even sequence: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn'],
			['Tuesday', 'TUESDAY', 'tuesday', 'TuEsDaY', 'TUesDAy','tUeSdAy', 'tuESdaY', 'Tue', 'TUE', 'tue', 'TuE', 'TUe', 'tuE', 'tUe'],
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Thursday'], ['Saturday'], ['Monday']];
		let expectedDataShortCapitalized = [['Thu'], ['Sat'], ['Mon']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 0;
		const r2To = 4;
		const r1 = 0;
		const r2 = 1;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 2, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 2, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 2, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 2, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 2, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 2, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 2, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 2, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 2, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 2, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 2, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 2, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 2, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 2, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Asc vertical odd sequence: Days of the weeks', function (assert) {
		let testData = [
			['Monday', 'MONDAY', 'monday', 'MoNdaY', 'MOnDaY','mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn'],
			['Wednesday', 'WEDNESDAY', 'wednesday', 'WeDnEsDaY', 'WEdNEsDAy','wEdNeSdAy', 'weDNesDAy', 'Wed', 'WED', 'wed', 'WeD', 'WEd', 'weD', 'wEd'],
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Friday'], ['Sunday'], ['Tuesday']];
		let expectedDataShortCapitalized = [['Fri'], ['Sun'], ['Tue']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 0;
		const r2To = 4;
		const r1 = 0;
		const r2 = 1;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 2, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 2, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 2, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 2, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 2, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 2, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 2, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 2, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 2, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 2, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 2, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 2, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 2, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 2, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Asc vertical sequence of full and short names: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn'],
			['Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn', 'Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Monday'], ['Mon'], ['Tuesday']];
		let expectedDataShortCapitalized = [['Mon'], ['Monday'], ['Tue']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 0;
		const r2To = 4;
		const r1 = 0;
		const r2 = 1;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 2, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 2, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 2, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 2, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 2, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 2, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 2, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 2, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 2, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 2, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 2, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 2, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 2, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 2, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Reverse vertical sequence of full and short names: Days of the weeks', function (assert) {
		let testData = [
			['Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn', 'Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay'],
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Saturday'], ['Sat'], ['Friday']];
		let expectedDataShortCapitalized = [['Sat'], ['Saturday'], ['Fri']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 4;
		const r2To = 0;
		const r1 = 3;
		const r2 = 4;
		let range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To - 2, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 2, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 2, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 2, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 2, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 2, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 2, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 2, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 2, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 2, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 2, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 2, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 2, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 2, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Reverse vertical even sequence: Days of the weeks', function (assert) {
		let testData = [
			['Friday', 'FRIDAY', 'friday', 'FrIdAy', 'FRIdAy', 'fRIdAy', 'frIDaY', 'Fri', 'FRI', 'fri', 'FrI', 'FRi', 'frI', 'fRi'],
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Wednesday'], ['Monday'], ['Saturday']];
		let expectedDataShortCapitalized = [['Wed'], ['Mon'], ['Sat']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 4;
		const r2To = 0;
		const r1 = 3;
		const r2 = 4;
		let range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To - 2, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 2, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 2, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 2, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 2, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 2, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 2, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 2, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 2, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 2, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 2, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 2, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 2, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 2, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Reverse vertical odd sequence: Days of the weeks', function (assert) {
		let testData = [
			['Thursday', 'THURSDAY', 'thursday', 'ThUrSdAy', 'THUrSdAy', 'tHUrSdAy', 'thURsdAy', 'Thu', 'THU', 'thu', 'ThU', 'THu', 'thU', 'tHu'],
			['Saturday', 'SATURDAY', 'saturday', 'SaTuRdAy', 'SAtUrDaY', 'sAtUrDaY', 'saTUrdAy', 'Sat', 'SAT', 'sat', 'SaT', 'SAt', 'saT', 'sAt']
		];
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Tuesday'], ['Sunday'], ['Friday']];
		let expectedDataShortCapitalized = [['Tue'], ['Sun'], ['Fri']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		const nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		const r1To = 4;
		const r2To = 0;
		const r1 = 3;
		const r2 = 4;
		let range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To - 2, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 2, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 2, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 2, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 2, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 2, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 2, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 2, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 2, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 2, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 2, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 2, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 2, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 2, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 4);
	});
	QUnit.test('Autofill - Horizontal sequence. Range 9 cells : Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Monday'],
			['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY', 'MONDAY'],
			['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'monday'],
			['SuNdaY', 'MoNdaY', 'TuEsdaY', 'WeDnEsDaY', 'ThUrSdAy', 'FrIdAy', 'SaTuRdAy', 'SuNdaY', 'MoNdaY'],
			['SUnDAy', 'MOnDAy', 'TUEsDAy', 'WEDnEsDAy', 'THUrsDAy', 'FRIday', 'SATurday', 'SUnDAy', 'MOnDAy'],
			['sUnDaY', 'mOnDaY', 'tUeSdAy', 'wEdNeSdAy', 'tHuRsDaY', 'fRiDaY', 'sAtUrDaY', 'sUnDaY', 'mOnDaY'],
			['suNDay', 'moNDay', 'tuESday', 'weDNesday', 'thURsday', 'frIDay', 'saTUrday', 'suNDay', 'moNDay'],
			['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', 'Mon'],
			['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'MON'],
			['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'mon'],
			['SuN', 'MoN', 'TuE', 'WeD', 'ThU', 'FrI', 'SaT', 'SuN', 'MoN'],
			['SUn', 'MOn', 'TUe', 'WEd', 'THu', 'FRi', 'SAt', 'SUn', 'MOn'],
			['sUn', 'mOn', 'tUe', 'wEd', 'tHu', 'fRi', 'sAt', 'sUn', 'mOn'],
			['suN', 'moN', 'tuE', 'weD', 'thU', 'frI', 'saT', 'suN', 'moN']
		];

		//Asc sequence case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Tuesday', 'Wednesday', 'Thursday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Tue', 'Wed', 'Thu'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		let nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		let c1To = 8;
		let c2To = 11;
		let c1 = 0;
		let c2 = 8;

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(c1To + 1, 0, c2To, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 1, c2To, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 2, c2To, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 3, c2To, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 4, c2To, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 5, c2To, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 6, c2To, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 7, c2To, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 8, c2To, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 9, c2To, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 10, c2To, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 11, c2To, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 12, c2To, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 13, c2To, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2To, 13);
		// Reverse case
		expectedDataCapitalized = ['Saturday', 'Friday', 'Thursday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Sat', 'Fri', 'Thu'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		c1To = 3;
		c2To = 0;
		c1 = 3;
		c2 = 11;

		range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 0, c2To, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 1, c2To, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 2, c2To, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 3, c2To, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 4, c2To, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 5, c2To, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 6, c2To, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 7, c2To, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 8, c2To, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 9, c2To, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 10, c2To, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 11, c2To, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 12, c2To, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 13, c2To, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2, 13);
	});
	QUnit.test('Autofill - Horizontal Even sequence. Range 9 cells : Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'Tuesday', 'Thursday', 'Saturday', 'Monday', 'Wednesday', 'Friday', 'Sunday', 'Tuesday'],
			['SUNDAY', 'TUESDAY', 'THURSDAY', 'SATURDAY', 'MONDAY', 'WEDNESDAY', 'FRIDAY', 'SUNDAY', 'TUESDAY'],
			['sunday', 'tuesday', 'thursday', 'saturday', 'monday', 'wednesday', 'friday', 'sunday', 'tuesday'],
			['SuNdaY', 'TuEsdaY', 'ThUrSdAy', 'SaTuRdAy', 'MoNdaY', 'WeDnEsDaY', 'FrIdAy', 'SuNdaY', 'TuEsdaY'],
			['SUnDAy', 'TUEsDAy', 'THUrsDAy', 'SATurday', 'MOnDAy', 'WEDnEsDAy', 'FRiDAy', 'SUnDAy', 'TUEsDAy'],
			['sUnDaY', 'tUeSdAy', 'tHuRsDaY', 'sAtUrDaY', 'mOnDaY', 'wEdnEsDaY', 'fRiDaY',  'sUnDaY', 'tUeSdAy'],
			['suNDay', 'tuESday', 'thURsday', 'saTUrday', 'moNDay', 'weDNesday', 'frIDay', 'suNDay', 'tuESday'],
			['Sun', 'Tue', 'Thu', 'Sat', 'Mon', 'Wed', 'Fri', 'Sun', 'Tue'],
			['SUN', 'TUE', 'THU', 'SAT', 'MON', 'WED', 'FRI', 'SUN', 'TUE'],
			['sun', 'tue', 'thu', 'sat', 'mon', 'wed', 'fri', 'sun', 'tue'],
			['SuN', 'TuE', 'ThU', 'SaT', 'MoN', 'WeD', 'FrI', 'SuN', 'TuE'],
			['SUn', 'TUe', 'THu', 'SAt', 'MOn', 'WEd', 'FRi', 'SUn', 'TUe'],
			['sUn', 'tUe', 'tHu', 'sAt', 'mOn', 'wEd', 'fRi', 'sUn', 'tUe'],
			['suN', 'tuE', 'thU', 'saT', 'moN', 'weD', 'frI', 'suN', 'tuE']
		];

		//Asc sequence case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Thursday', 'Saturday', 'Monday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Thu', 'Sat', 'Mon'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		let nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		let c1To = 8;
		let c2To = 11;
		let c1 = 0;
		let c2 = 8;

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(c1To + 1, 0, c2To, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 1, c2To, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 2, c2To, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 3, c2To, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 4, c2To, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 5, c2To, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 6, c2To, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 7, c2To, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 8, c2To, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 9, c2To, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 10, c2To, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 11, c2To, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 12, c2To, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 13, c2To, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2To, 13);
		// Reverse case
		expectedDataCapitalized = ['Friday', 'Wednesday', 'Monday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Fri', 'Wed', 'Mon'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		c1To = 3;
		c2To = 0;
		c1 = 3;
		c2 = 11;

		range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 0, c2To, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 1, c2To, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 2, c2To, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 3, c2To, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 4, c2To, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 5, c2To, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 6, c2To, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 7, c2To, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 8, c2To, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 9, c2To, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 10, c2To, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 11, c2To, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 12, c2To, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 13, c2To, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2, 13);
	});
	QUnit.test('Autofill - Horizontal Odd sequence. Range 9 cells : Days of the weeks', function (assert) {
		let testData = [
			['Monday', 'Wednesday', 'Friday', 'Sunday', 'Tuesday', 'Thursday', 'Saturday', 'Monday', 'Wednesday'],
			['MONDAY', 'WEDNESDAY', 'FRIDAY', 'SUNDAY', 'TUESDAY', 'THURSDAY', 'SATURDAY', 'MONDAY', 'WEDNESDAY'],
			['monday', 'wednesday', 'friday', 'sunday', 'tuesday', 'thursday', 'saturday', 'monday', 'wednesday'],
			['MoNdAy', 'WeDnEsDaY', 'FrIdAy', 'SuNdAy', 'TuEsDaY', 'ThUrSdAy', 'SaTuRdAy', 'MoNdAy', 'WeDnEsDaY'],
			['MOnDAy', 'WEDnEsDAy', 'FRIday', 'SUnDAy', 'TUEsDAy', 'THUrsDAy', 'SATurday', 'MOnDAy', 'WEDnEsDAy'],
			['mOnDaY', 'wEdNeSdAy', 'fRiDaY', 'sUnDaY', 'tUeSdAy', 'tHuRsDaY', 'sAtUrDaY', 'mOnDaY', 'wEdNeSdAy'],
			['moNDay', 'weDNesday', 'frIDay', 'suNDay', 'tuESday','thURsday','saTUrday','moNDay','weDNesday'],
			['Mon', 'Wed', 'Fri', 'Sun', 'Tue', 'Thu', 'Sat', 'Mon', 'Wed'],
			['MON', 'WED', 'FRI', 'SUN', 'TUE', 'THU', 'SAT', 'MON', 'WED'],
			['mon', 'wed', 'fri', 'sun', 'tue', 'thu', 'sat', 'mon', 'wed'],
			['MoN', 'WeD', 'FrI', 'SuN', 'TuE', 'ThU', 'SaT', 'MoN', 'WeD'],
			['MOn', 'WEd','FRi','SUn','TUe','THu','SAt','MOn','WEd'],
			['mOn','wEd','fRi','sUn','tUe','tHu','sAt','mOn','wEd'],
			['moN','weD','frI','suN','tuE','thU','saT','moN','weD']
		];

		//Asc sequence case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['Friday', 'Sunday', 'Tuesday'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Fri', 'Sun', 'Tue'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		let nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		let c1To = 8;
		let c2To = 11;
		let c1 = 0;
		let c2 = 8;

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(c1To + 1, 0, c2To, 0);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 1, c2To, 1);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 2, c2To, 2);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 3, c2To, 3);
		autofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 4, c2To, 4);
		autofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 5, c2To, 5);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 6, c2To, 6);
		autofillData(assert, autoFillRange, [expectedDataLower], 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 7, c2To, 7);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 8, c2To, 8);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 9, c2To, 9);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 10, c2To, 10);
		autofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 11, c2To, 11);
		autofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 12, c2To, 12);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To + 1, 13, c2To, 13);
		autofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2To, 13);
		// Reverse case
		expectedDataCapitalized = ['Saturday', 'Thursday', 'Tuesday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Sat', 'Thu', 'Tue'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);
		//Fill handle direction and area
		nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		c1To = 3;
		c2To = 0;
		c1 = 3;
		c2 = 11;

		range = ws.getRange4(0,3);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(c1, 0, c2, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 0, c2To, 0);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence  With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 1, c2, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 1, c2To, 1);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1, 2, c2, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 2, c2To, 2);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1, 3, c2, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 3, c2To, 3);
		reverseAutofillData(assert, autoFillRange, [expectedDataCapitalized], 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1, 4, c2, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 4, c2To, 4);
		reverseAutofillData(assert, autoFillRange, [expectedDataUpper], 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1, 5, c2, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 5, c2To, 5);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1, 6, c2, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 6, c2To, 6);
		reverseAutofillData(assert, autoFillRange, [expectedDataLower], 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1, 7, c2, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 7, c2To, 7);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1, 8, c2,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 8, c2To, 8);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1,9,c2,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 9, c2To, 9);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1, 10, c2, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 10, c2To, 10);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortCapitalized], 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1, 11, c2, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 11, c2To, 11);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortUpper], 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1, 12, c2, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 12, c2To, 12);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1, 13, c2, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(c1To - 1, 13, c2To, 13);
		reverseAutofillData(assert, autoFillRange, [expectedDataShortLower], 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, c2, 13);
	});
	QUnit.test('Autofill - Vertical sequence. Range 9 cells: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn'],
			['Monday', 'MONDAY', 'monday', 'MoNdAy', 'MOnDaY','mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn'],
			['Tuesday', 'TUESDAY', 'tuesday', 'TuEsDaY', 'TUesDAy','tUeSdAy', 'tuESdaY', 'Tue', 'TUE', 'tue', 'TuE', 'TUe', 'tuE', 'tUe'],
			['Wednesday', 'WEDNESDAY', 'wednesday', 'WeDnesdAy', 'WEdnesDaY','wEdnesDaY', 'weDneSdAy', 'Wed', 'WED', 'wed', 'WeD', 'WEd', 'weD', 'wEd'],
			['Thursday', 'THURSDAY', 'thursday', 'ThUrsDaY', 'THursDaY','tHuRsDaY', 'thUrsDaY', 'Thu', 'THU', 'thu', 'ThU', 'THu', 'thU', 'tHu'],
			['Friday', 'FRIDAY', 'friday', 'FrIdAy', 'FRidAy','fRIdAy', 'frIdAy', 'Fri', 'FRI', 'fri', 'FrI', 'FRi', 'frI', 'fRI'],
			['Saturday', 'SATURDAY', 'saturday', 'SaTurDaY', 'SAturDaY','sAturDaY', 'saTurDaY', 'Sat', 'SAT', 'sat', 'SaT', 'SAt', 'saT', 'sAt'],
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn'],
			['Monday', 'MONDAY', 'monday', 'MoNdAy', 'MOnDaY','mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn']
		];

		//Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Tuesday'], ['Wednesday'], ['Thursday']];
		let expectedDataShortCapitalized = [['Tue'], ['Wed'], ['Thu']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		//Fill handle direction and area
		let nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 & r2 for autofill range and ranges
		let r1To = 8;
		let r2To = 11;
		let r1 = 0;
		let r2 = 8;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 1, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 1, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 1, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 1, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 1, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 1, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 1, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 1, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 1, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 1, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 1, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 1, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 1, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 1, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Saturday'], ['Friday'], ['Thursday']];
		expectedDataShortCapitalized = [['Sat'], ['Fri'], ['Thu']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		//Fill handle direction and area
		nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		r1To = 3;
		r2To = 0;
		r1 = 3;
		r2 = 11;

		range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(0, r1To - 1, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 1, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 1, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 1, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 1, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 1, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 1, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 1, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 1, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 1, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 1, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 1, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 1, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 1, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11);
	});
	QUnit.test('Autofill - Vertical even sequence. Range 9 cell: Days of the weeks', function (assert) {
		let testData = [
			['Sunday', 'SUNDAY', 'sunday', 'SuNdAy', 'SUnDaY','sUnDaY', 'suNDay', 'Sun', 'SUN', 'sun', 'SuN', 'SUn', 'suN', 'sUn'],
			['Tuesday', 'TUESDAY', 'tuesday', 'TuEsDaY', 'TUesDAy','tUeSdAy', 'tuESdaY', 'Tue', 'TUE', 'tue', 'TuE', 'TUe', 'tuE', 'tUe'],
			['Thursday', 'THURSDAY', 'thursday', 'ThUrSdAy', 'THuRsDaY','tHuRsDaY','thURsday','Thu','THU','thu','ThU','THu','thU','tHu',],
			['Saturday','SATURDAY','saturday','SaTuRdAy','SAtUrDaY','sAtUrDaY','saTUrday','Sat','SAT','sat','SaT','SAt','saT','sAt'],
			['Monday', 'MONDAY', 'monday', 'MoNdAy', 'MOnDaY', 'mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn'],
			['Wednesday', 'WEDNESDAY', 'wednesday', 'WeDnEsDaY', 'WEdNeSDaY','wEdNeSdAy', 'weDNesday', 'Wed', 'WED', 'wed', 'WeD', 'WEd', 'weD', 'wEd'],
			['Friday', 'FRIDAY', 'friday', 'FrIdAy', 'FRiDaY', 'fRIdaY', 'frIdAY', 'Fri', 'FRI', 'fri', 'FrI', 'FRi', 'fRI', 'frI'],
			['Sunday','SUNDAY','sunday','SuNdAy','SUnDaY','sUnDaY','suNDay','Sun','SUN','sun','SuN','SUn','suN','sUn'],
			['Tuesday', 'TUESDAY', 'tuesday', 'TuEsDaY', 'TUesDAy','tUeSdAy', 'tuESdaY', 'Tue', 'TUE', 'tue', 'TuE', 'TUe', 'tuE', 'tUe']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Thursday'], ['Saturday'], ['Monday']];
		let expectedDataShortCapitalized = [['Thu'], ['Sat'], ['Mon']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		let nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		let r1To = 8;
		let r2To = 11;
		let r1 = 0;
		let r2 = 8;

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 1, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 1, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 1, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 1, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 1, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 1, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 1, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 1, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 1, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 1, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 1, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 1, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 1, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 1, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Friday'], ['Wednesday'], ['Monday']];
		expectedDataShortCapitalized = [['Fri'], ['Wed'], ['Mon']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		//Fill handle direction and area
		nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		r1To = 3;
		r2To = 0;
		r1 = 3;
		r2 = 11;

		range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(0, r1To - 1, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 1, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 1, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 1, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 1, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 1, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 1, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 1, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 1, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 1, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 1, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 1, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 1, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 1, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11)
	});
	QUnit.test('Autofill - Vertical odd sequence. Range 9 cells: Days of the weeks', function (assert) {
		let testData = [
			['Monday', 'MONDAY', 'monday', 'MoNdaY', 'MOnDaY','mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn'],
			['Wednesday', 'WEDNESDAY', 'wednesday', 'WeDnEsDaY', 'WEdNEsDAy','wEdNeSdAy', 'weDNesDAy', 'Wed', 'WED', 'wed', 'WeD', 'WEd', 'weD', 'wEd'],
			['Friday','FRIDAY','friday','FrIdAy','FRiDaY','fRiDaY','frIDay','Fri','FRI','fri','FrI','FRi','frI','fRi'],
			['Sunday','SUNDAY','sunday','SuNdAy','SUnDaY','sUnDaY','suNDay','Sun','SUN','sun','SuN','SUn','suN','sUn'],
			['Tuesday', 'TUESDAY', 'tuesday', 'TuEsDaY', 'TUesDAy','tUeSdAy', 'tuESdaY', 'Tue', 'TUE', 'tue', 'TuE', 'TUe', 'tuE', 'tUe'],
			['Thursday', 'THURSDAY', 'thursday', 'ThUrSdAy', 'THuRsDaY','tHuRsDaY','thURsday','Thu','THU','thu','ThU','THu','thU','tHu'],
			['Saturday','SATURDAY','saturday','SaTuRdAy','SAtUrDaY','sAtUrDaY','saTUrday','Sat','SAT','sat','SaT','SAt','saT','sAt'],
			['Monday', 'MONDAY', 'monday', 'MoNdaY', 'MOnDaY','mOnDaY', 'moNDay', 'Mon', 'MON', 'mon', 'MoN', 'MOn', 'moN', 'mOn'],
			['Wednesday', 'WEDNESDAY', 'wednesday', 'WeDnEsDaY', 'WEdNEsDAy','wEdNeSdAy', 'weDNesDAy', 'Wed', 'WED', 'wed', 'WeD', 'WEd', 'weD', 'wEd']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['Friday'], ['Sunday'], ['Tuesday']];
		let expectedDataShortCapitalized = [['Fri'], ['Sun'], ['Tue']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);


		//Fill handle direction and area
		let nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		let nFillHandleArea = 3; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		let r1To = 8;
		let r2To = 11;
		let r1 = 0;
		let r2 = 8;
		let range = ws.getRange4(0,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, r1To + 1, 0, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To + 1, 1, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To + 1, 2, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To + 1, 3, r2To);
		autofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Asc sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To + 1, 4, r2To);
		autofillData(assert, autoFillRange, expectedDataUpper, 'Case: Asc sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To + 1, 5, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To + 1, 6, r2To);
		autofillData(assert, autoFillRange, expectedDataLower, 'Case: Asc sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To + 1, 7, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To + 1, 8, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To + 1, 9, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To + 1, 10, r2To);
		autofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Asc sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To + 1, 11, r2To);
		autofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Asc sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To + 1, 12, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To + 1, 13, r2To);
		autofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Asc sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Saturday'], ['Thursday'], ['Tuesday']];
		expectedDataShortCapitalized = [['Sat'], ['Thu'], ['Tue']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		//Fill handle direction and area
		nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		nFillHandleArea = 1; // 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// Fill r1 ? r2 for autofill range and ranges
		r1To = 3;
		r2To = 0;
		r1 = 3;
		r2 = 11;

		range = ws.getRange4(3,0);
		range.fillData(testData);

		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1, 0, r2)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(0, r1To - 1, 0, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence With capitalized start from Sunday');

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1, 1, r2)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, r1To - 1, 1, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Upper-registry start from Sunday');

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1,2, r2)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, r1To - 1, 2, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Lower-registry start from Sunday');

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1, 3, r2)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, r1To - 1, 3, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataCapitalized, 'Case: Reverse sequence Camel-registry - Su. Start from Sunday');

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1, 4, r2)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, r1To - 1, 4, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataUpper, 'Case: Reverse sequence Camel-registry - SU. Start from Sunday');

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1, 5, r2)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, r1To - 1, 5, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - sU. Start from Sunday');

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1, 6, r2)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, r1To - 1, 6, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataLower, 'Case: Reverse sequence Camel-registry - su. Start from Sunday');

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1, 7, r2)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, r1To - 1, 7, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name with capitalized start from Sun');

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1, 8, r2)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, r1To - 1, 8, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Upper-registry start from Sun');

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1, 9, r2)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, r1To - 1, 9, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Lower-registry start from Sun');

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1, 10, r2)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, r1To - 1, 10, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortCapitalized, 'Case: Reverse sequence Short name Camel-registry - Su. Start from Sun');

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1, 11, r2)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, r1To - 1, 11, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortUpper, 'Case: Reverse sequence Short name Camel-registry - SU. Start from Sun');

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1, 12, r2)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, r1To - 1, 12, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - sU. Start from Sun');

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1, 13 ,r2)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, r1To - 1, 13, r2To);
		reverseAutofillData(assert, autoFillRange, expectedDataShortLower, 'Case: Reverse sequence Short name Camel-registry - su. Start from Sun');

		clearData(0, 0, 13, 11)
	});

	QUnit.module("Sheet structure");
});
