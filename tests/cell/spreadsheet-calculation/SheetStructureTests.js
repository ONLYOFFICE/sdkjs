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
		ws.TableParts = [];
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
	function getHorizontalAutofillCases(c1From, c2From, c1To, c2To, assert, expectedData, nFillHandleArea) {
		const [
			expectedDataCapitalized,
			expectedDataUpper,
			expectedDataLower,
			expectedDataShortCapitalized,
			expectedDataShortUpper,
			expectedDataShortLower
		] = expectedData;

		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		let autofillC1 =  nFillHandleArea === 3 ? c2From + 1 : c1From - 1;
		const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;
		const descSequenceType = nFillHandleArea === 3 ? 'Asc sequence.' : 'Reverse sequence.';
		// With capitalized
		ws.selectionRange.ranges = [getRange(c1From, 0, c2From, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(autofillC1, 0, c2To, 0);
		autoFillAssert(assert, autoFillRange, [expectedDataCapitalized], `Case: ${descSequenceType} With capitalized`);

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1From, 1, c2From, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 1, c2To, 1);
		autoFillAssert(assert, autoFillRange, [expectedDataUpper], `Case: ${descSequenceType} Upper-registry`);

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1From, 2, c2From, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 2, c2To, 2);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Lower-registry`);

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1From, 3, c2From, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 3, c2To, 3);
		autoFillAssert(assert, autoFillRange, [expectedDataCapitalized], `Case: ${descSequenceType} Camel-registry - Su.`);

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1From, 4, c2From, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 4, c2To, 4);
		autoFillAssert(assert, autoFillRange, [expectedDataUpper], `Case: ${descSequenceType} Camel-registry - SU.`);

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1From, 5, c2From, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 5, c2To, 5);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Camel-registry - sU.`);

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1From, 6, c2From, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 6, c2To, 6);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Camel-registry - su.`);

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1From, 7, c2From, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 7, c2To, 7);
		autoFillAssert(assert, autoFillRange, [expectedDataShortCapitalized], `Case: ${descSequenceType} Short name with capitalized`);

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1From, 8, c2From,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 8, c2To, 8);
		autoFillAssert(assert, autoFillRange, [expectedDataShortUpper], `Case: ${descSequenceType} Short name Upper-registry start from Sun`);

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1From,9,c2From,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 9, c2To, 9);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Lower-registry`);

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1From, 10, c2From, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 10, c2To, 10);
		autoFillAssert(assert, autoFillRange, [expectedDataShortCapitalized], `Case: ${descSequenceType} Short name Camel-registry - Su.`);

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1From, 11, c2From, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 11, c2To, 11);
		autoFillAssert(assert, autoFillRange, [expectedDataShortUpper], `Case: ${descSequenceType} Short name Camel-registry - SU.`);

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1From, 12, c2From, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 12, c2To, 12);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Camel-registry - sU.`);

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1From, 13, c2From, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 13, c2To, 13);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Camel-registry - su.`);
	}

	function getVerticalAutofillCases (r1From, r2From, r1To, r2To, assert, expectedData, nFillHandleArea) {
		const [
			expectedDataCapitalized,
			expectedDataUpper,
			expectedDataLower,
			expectedDataShortCapitalized,
			expectedDataShortUpper,
			expectedDataShortLower
		] = expectedData;

		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		let autofillR1 =  nFillHandleArea === 3 ? r2From + 1 : r1From - 1;
		const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;
		const descSequenceType = nFillHandleArea === 3 ? 'Asc sequence.' : 'Reverse sequence.';
		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1From, 0, r2From)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, autofillR1, 0, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataCapitalized, `Case: ${descSequenceType} With capitalized`);

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1From, 1, r2From)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, autofillR1, 1, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataUpper, `Case: ${descSequenceType} Upper-registry`);

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1From, 2, r2From)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, autofillR1, 2, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Lower-registry`);

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1From, 3, r2From)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, autofillR1, 3, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataCapitalized, `Case: ${descSequenceType} Camel-registry - Su.`);

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1From, 4, r2From)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, autofillR1, 4, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataUpper, `Case: ${descSequenceType} Camel-registry - SU.`);

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1From, 5, r2From)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, autofillR1, 5, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Camel-registry - sU.`);

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1From, 6, r2From)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, autofillR1, 6, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Camel-registry - su.`);

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1From, 7, r2From)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, autofillR1, 7, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortCapitalized, `Case: ${descSequenceType} Short name with capitalized`);

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1From, 8, r2From)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, autofillR1, 8, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortUpper, `Case: ${descSequenceType} Short name Upper-registry`);

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1From, 9, r2From)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, autofillR1, 9, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Lower-registry`);

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1From, 10, r2From)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, autofillR1, 10, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortCapitalized, `Case: ${descSequenceType} Short name Camel-registry - Su.`);

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1From, 11, r2From)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, autofillR1, 11, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortUpper, `Case: ${descSequenceType} Short name Camel-registry - SU.`);

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1From, 12, r2From)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, autofillR1, 12, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Camel-registry - sU.`);

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1From, 13, r2From)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, autofillR1, 13, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Camel-registry - su.`);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 0, 0, 3, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 3, 3, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 0, 0, 3, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 3, 3, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);

		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

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

		let range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 8, 8, 11, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

		clearData(0, 0, 11, 13);

		// Reverse case
		expectedDataCapitalized = ['Saturday', 'Friday', 'Thursday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Sat', 'Fri', 'Thu'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

		clearData(0, 0, 11, 13);
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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 8, 0, 11, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

		clearData(0, 0, 11, 13);

		// Reverse case
		expectedDataCapitalized = ['Friday', 'Wednesday', 'Monday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Fri', 'Wed', 'Mon'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

		clearData(0, 0, 11, 13);
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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 8, 0, 11, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);


		clearData(0, 0, 11, 13);

		// Reverse case
		expectedDataCapitalized = ['Saturday', 'Thursday', 'Tuesday'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Sat', 'Thu', 'Tue'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0,3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

		clearData(0, 0, 11, 13);
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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 8, 0, 11, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Saturday'], ['Friday'], ['Thursday']];
		expectedDataShortCapitalized = [['Sat'], ['Fri'], ['Thu']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 8, 0, 11, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Friday'], ['Wednesday'], ['Monday']];
		expectedDataShortCapitalized = [['Fri'], ['Wed'], ['Mon']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

		clearData(0, 0, 13, 11);
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

		let range = ws.getRange4(0,0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 8, 0, 11, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);

		clearData(0, 0, 13, 11);

		//Reverse case
		expectedDataCapitalized = [['Saturday'], ['Thursday'], ['Tuesday']];
		expectedDataShortCapitalized = [['Sat'], ['Thu'], ['Tue']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3,0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 11, 11, 0, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);

		clearData(0, 0, 13, 11);
	});
	QUnit.test('Autofill - Horizontal sequence. Months', function (assert) {
		let testData = [
			['January'],
			['JANUARY'],
			['january'],
			['JaNuArY'],
			['JAnuARy'],
			['jAnUaRy'],
			['jaNUarY'],
			['Jan'],
			['JAN'],
			['jan'],
			['JaN'],
			['JAn'],
			['jaN'],
			['jAn']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['February', 'March', 'April'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Feb', 'Mar', 'Apr'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 0, 0, 3,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 3, 0);

		// Reverse case
		expectedDataCapitalized = ['December', 'November', 'October'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Dec', 'Nov', 'Oct'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 3, 3, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 3, 0);
	});
	QUnit.test('Autofill - Vertical sequence. Months', function (assert) {
		let testData = [
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['February'], ['March'], ['April']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Feb'], ['Mar'], ['Apr']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 0, 0, 3,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 3);

		// Reverse case
		expectedDataCapitalized = [['December'], ['November'], ['October']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Dec'], ['Nov'], ['Oct']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 3, 3, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 3);
	});
	QUnit.test('Autofill - Horizontal even sequence. Months', function (assert) {
		let testData = [
			['December', 'February'],
			['DECEMBER', 'FEBRUARY'],
			['december', 'february'],
			['DeCeMbEr', 'FeBrUaRy'],
			['DEcEMBeR', 'FEbRUaRY'],
			['dEcEMbEr', 'fEbRuArY'],
			['deCEmbER', 'feBRuaRY'],
			['Dec', 'Feb'],
			['DEC', 'FEB'],
			['dec', 'feb'],
			['DeC', 'FeB'],
			['DEc', 'FEb'],
			['deC', 'feB'],
			['dEc', 'fEb']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['April', 'June', 'August'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Apr', 'Jun', 'Aug'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 4, 0);

		// Reverse case
		expectedDataCapitalized = ['October', 'August', 'June'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Oct', 'Aug', 'Jun'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 4, 0);
	});
	QUnit.test('Autofill - Vertical even sequence.  Months', function (assert) {
		let testData = [
			['December', 'DECEMBER', 'december', 'DeCeMbEr', 'DEcEMBeR', 'dEcEMbEr', 'deCEmbER', 'Dec', 'DEC', 'dec', 'DeC', 'DEc', 'deC', 'dEc'],
			['February', 'FEBRUARY', 'february', 'FeBrUaRy', 'FEbRUARy', 'fEbRUaRy', 'feBRuaRY', 'Feb', 'FEB', 'feb', 'FeB', 'FEb', 'feB', 'fEb']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['April'], ['June'], ['August']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Apr'], ['Jun'], ['Aug']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 4);

		// Reverse case
		expectedDataCapitalized = [['October'], ['August'], ['June']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Oct'], ['Aug'], ['Jun']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Autofill - Horizontal odd sequence. Months', function (assert) {
		let testData = [
			['January', 'March'],
			['JANUARY', 'MARCH'],
			['january', 'march'],
			['JaNuArY', 'MaRCH'],
			['JAnuARy', 'MArCH'],
			['jAnUaRy', 'mArcH'],
			['jaNUarY', 'maRCh'],
			['Jan', 'Mar'],
			['JAN', 'MAR'],
			['jan', 'mar'],
			['JaN', 'MaR'],
			['JAn', 'MAr'],
			['jaN', 'maR'],
			['jAn', 'mAr']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['May', 'July', 'September'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['May', 'Jul', 'Sep'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 4, 0);

		// Reverse case
		expectedDataCapitalized = ['November', 'September', 'July'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Nov', 'Sep', 'Jul'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 4, 0);
	});
	QUnit.test('Autofill - Vertical odd sequence. Months', function (assert) {
		let testData = [
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
			['March', 'MARCH', 'march', 'MaRcH', 'MArCH', 'mArCh', 'maRcH', 'Mar', 'MAR', 'mar', 'MaR', 'MAr', 'maR', 'mAr']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['May'], ['July'], ['September']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['May'], ['Jul'], ['Sep']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 4);

		// Reverse case
		expectedDataCapitalized = [['November'], ['September'], ['July']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Nov'], ['Sep'], ['Jul']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Autofill - Horizontal sequence of full and short names. Months', function (assert) {
		let testData = [
			['January', 'Jan'],
			['JANUARY', 'JAN'],
			['january', 'jan'],
			['JaNuArY', 'JaN'],
			['JAnuARy', 'JAn'],
			['jAnUaRy', 'jAn'],
			['jaNUarY', 'jaN'],
			['Jan', 'January'],
			['JAN', 'JANUARY'],
			['jan', 'january'],
			['JaN', 'JaNuArY'],
			['JAn', 'JAnUaRy'],
			['jaN', 'jaNUarY'],
			['jAn', 'jAnUaRy']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['February', 'Feb', 'March'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Feb', 'February', 'Mar'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 4, 0);

		// Reverse case
		expectedDataCapitalized = ['Dec', 'December', 'Nov'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['December', 'Dec', 'November'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 4, 0);
	});
	QUnit.test('Autofill - Vertical sequence of full and short names. Months', function (assert) {
		let testData = [
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
			['Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jAn', 'jaN', 'January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jaNUarY', 'jAnUaRy']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['February'], ['Feb'], ['March']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Feb'], ['February'], ['Mar']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 4);

		// Reverse case
		expectedDataCapitalized = [['Dec'], ['December'], ['Nov']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['December'], ['Dec'], ['November']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 4);
	});

	QUnit.test('Autofill - Horizontal sequence: Range 14 cells. Months', function (assert) {
		let testData = [
			['December', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January'],
			['DECEMBER', 'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY'],
			['december', 'january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december', 'january'],
			['DeCeMbEr', 'JaNuArY', 'FeBrUaRy', 'MaRcH', 'ApRiL', 'MaY', 'JuNe', 'JuLy', 'AuGuSt', 'SePtEmBeR', 'OcToBeR', 'NoVemBeR', 'DeCeMbEr', 'JaNuArY'],
			['DEcEMBeR', 'JAnuARy', 'FEbrUAry', 'MArCH', 'APriL', 'MAy', 'JUne', 'JUly', 'AUguST', 'SEptEMbeR', 'OCtoBEr', 'NOveMBer', 'DEcEMBeR', 'JAnuARy'],
			['dEcEMbEr', 'jAnUaRy', 'fEbRuArY', 'mArCh', 'aPrIl', 'mAy', 'jUnE', 'jUlY', 'aUgUsT', 'sEpTeMbEr', 'oCtObEr', 'nOveMbEr', 'dEcEMbEr', 'jAnUaRy'],
			['deCEmbER', 'jaNUarY', 'feBRuaRY', 'maRCh', 'apRIl', 'maY', 'juNE', 'juLY', 'auGUst', 'sePTemBEr', 'ocTObeR', 'noVEmbER', 'deCEmbER', 'jaNUarY'],
			['Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan'],
			['DEC', 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC', 'JAN'],
			['dec', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec', 'jan'],
			['DeC', 'JaN', 'FeB', 'MaR', 'ApR', 'MaY', 'JuN', 'JuL', 'AuG', 'SeP', 'OcT', 'NoV', 'DeC', 'JaN'],
			['DEc', 'JAn', 'FEb', 'MAr', 'APr', 'MAy', 'JUn', 'JUl', 'AUg', 'SEp', 'OCt', 'NOv', 'DEc', 'JAn'],
			['deC', 'jaN', 'feB', 'maR', 'apR', 'maY', 'juN', 'juL', 'auG', 'seP', 'ocT', 'noV', 'deC', 'jaN'],
			['dEc', 'jAn', 'fEb', 'mAr', 'aPr', 'mAy', 'jUn', 'jUl', 'aUg', 'sEp', 'oCt', 'nOv', 'dEc', 'jAn']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['February', 'March', 'April'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Feb', 'Mar', 'Apr'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 13, 0, 16,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 16, 0);

		// Reverse case
		expectedDataCapitalized = ['November', 'October', 'September'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Nov', 'Oct', 'Sep'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 16, 16, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 16, 0);
	});
	QUnit.test('Autofill - Vertical sequence: Range 14 cells.  Months', function (assert) {
		let testData = [
			['December', 'DECEMBER', 'december', 'DeCeMbEr', 'DEcEMBeR', 'dEcEMbEr', 'deCEmbER', 'Dec', 'DEC', 'dec', 'DeC', 'DEc', 'deC', 'dEc'],
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
			['February', 'FEBRUARY', 'february', 'FeBrUaRy', 'FEbRUARy', 'fEbRUaRy', 'feBRuaRY', 'Feb', 'FEB', 'feb', 'FeB', 'FEb', 'feB', 'fEb'],
			['March', 'MARCH', 'march', 'MaRcH', 'MArCH', 'mArCh', 'maRcH', 'Mar', 'MAR', 'mar', 'MaR', 'MAr', 'maR', 'mAr'],
			['April', 'APRIL', 'april', 'ApRiL', 'APriL', 'aPrIL', 'apRIl', 'Apr', 'APR', 'apr', 'ApR', 'APr', 'apR', 'aPr'],
			['May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY', 'May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY'],
			['June', 'JUNE', 'june', 'JuNe', 'JUnE', 'jUnE', 'juNE', 'Jun', 'JUN', 'jun', 'JuN', 'JUn', 'juN', 'jUn'],
			['July', 'JULY', 'july', 'JuLy', 'JUly', 'jULY', 'juLY', 'Jul', 'JUL', 'jul', 'JuL', 'JUl', 'juL', 'jUl'],
			['August', 'AUGUST', 'august', 'AuGuSt', 'AUguST', 'aUGUSt', 'auGUST', 'Aug', 'AUG', 'aug', 'AuG', 'AUg', 'auG', 'aUg'],
			['September', 'SEPTEMBER', 'september', 'SePtEmBeR', 'SEptEMbeR', 'sEPTEMBer', 'sePTEMber', 'Sep', 'SEP', 'sep', 'SeP', 'SEp', 'seP', 'sEp'],
			['October', 'OCTOBER', 'october', 'OcToBeR', 'OCtoBEr', 'oCTOBEr', 'ocTOber', 'Oct', 'OCT', 'oct', 'OcT', 'OCt', 'ocT', 'oCt'],
			['November', 'NOVEMBER', 'november', 'NoVEmBeR', 'NOvEMbeR', 'nOvEMBeR', 'noVEMber', 'Nov', 'NOV', 'nov', 'NoV', 'NOv', 'noV', 'nOv'],
			['December', 'DECEMBER', 'december', 'DeCeMbEr', 'DEcEMBeR', 'dEcEMbEr', 'deCEmbER', 'Dec', 'DEC', 'dec', 'DeC', 'DEc', 'deC', 'dEc'],
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['February'], ['March'], ['April']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Feb'], ['Mar'], ['Apr']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 13, 0, 16, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 16);

		// Reverse case
		expectedDataCapitalized = [['November'], ['October'], ['September']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Nov'], ['Oct'], ['Sep']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 16, 16, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 16);
	});
	QUnit.test('Autofill - Horizontal even sequence: Range 10 cells. Months', function (assert) {
		let testData = [
			['December', 'February', 'April', 'June', 'August', 'October', 'December', 'February', 'April', 'June'],
			['DECEMBER', 'FEBRUARY', 'APRIL', 'JUNE', 'AUGUST', 'OCTOBER', 'DECEMBER', 'FEBRUARY', 'APRIL', 'JUNE'],
			['december', 'february', 'april', 'june', 'august', 'october', 'december', 'february', 'april', 'june'],
			['DeCeMbEr', 'FeBrUaRy', 'ApRiL', 'JuNe', 'AuGuSt', 'OcToBeR', 'DeCeMbEr', 'FeBrUaRy', 'ApRiL', 'JuNe'],
			['DEcEMBeR', 'FEbrUAry', 'APriL', 'JUne', 'AUguST', 'OCtoBEr', 'DEcEMBeR', 'FEbrUAry', 'APriL', 'JUne'],
			['dEcEMbEr', 'fEbRuArY', 'aPrIl', 'jUnE', 'aUgUsT', 'oCtObEr', 'dEcEMbEr', 'fEbRuArY', 'aPrIl', 'jUnE'],
			['deCEmbER', 'feBRuaRY', 'apRIl', 'juNE', 'auGUst', 'ocTObeR', 'deCEmbER', 'feBRuaRY', 'apRIl', 'juNE'],
			['Dec', 'Feb', 'Apr', 'Jun', 'Aug', 'Oct', 'Dec', 'Feb', 'Apr', 'Jun'],
			['DEC', 'FEB', 'APR', 'JUN', 'AUG', 'OCT','DEC', 'FEB', 'APR', 'JUN'],
			['dec','feb','apr','jun','aug','oct','dec', 'feb','apr','jun'],
			['DeC','FeB','ApR','JuN','AuG','OcT','DeC', 'FeB','ApR','JuN'],
			['DEc','FEb','APr','JUn','AUg','OCt','DEc', 'FEb','APr','JUn'],
			['deC','feB','apR','juN','auG','ocT','deC', 'feB','apR','juN'],
			['dEc','fEb','aPr','jUn','aUg','oCt','dEc', 'fEb','aPr','jUn']
		];


		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['August', 'October', 'December'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Aug', 'Oct', 'Dec'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 9, 0, 12, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 12, 0);

		// Reverse case
		expectedDataCapitalized = ['October', 'August', 'June'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Oct', 'Aug', 'Jun'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 12, 12, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 12, 0);
	});
	QUnit.test('Autofill - Vertical even sequence: Range 10 cells.  Months', function (assert) {
		let testData = [
			['December', 'DECEMBER', 'december', 'DeCeMbEr', 'DEcEMBeR', 'dEcEMbEr', 'deCEmbER', 'Dec', 'DEC', 'dec', 'DeC', 'DEc', 'deC', 'dEc'],
			['February', 'FEBRUARY', 'february', 'FeBrUaRy', 'FEbRUARy', 'fEbRUaRy', 'feBRuaRY', 'Feb', 'FEB', 'feb', 'FeB', 'FEb', 'feB', 'fEb'],
			['April', 'APRIL', 'april', 'ApRiL', 'APriL', 'aPrIL', 'apRIl', 'Apr', 'APR', 'apr', 'ApR', 'APr', 'apR', 'aPr'],
			['June', 'JUNE', 'june', 'JuNe', 'JUnE', 'jUnE', 'juNE', 'Jun', 'JUN', 'jun', 'JuN', 'JUn', 'juN', 'jUn'],
			['August', 'AUGUST', 'august', 'AuGuSt', 'AUguST', 'aUGUSt', 'auGUST', 'Aug', 'AUG', 'aug', 'AuG', 'AUg', 'auG', 'aUg'],
			['October', 'OCTOBER', 'october', 'OcToBeR', 'OCtoBEr', 'oCTOBEr', 'ocTOber', 'Oct', 'OCT', 'oct', 'OcT', 'OCt', 'ocT', 'oCt'],
			['December', 'DECEMBER', 'december', 'DeCeMbEr', 'DEcEMBeR', 'dEcEMbEr', 'deCEmbER', 'Dec', 'DEC', 'dec', 'DeC', 'DEc', 'deC', 'dEc'],
			['February', 'FEBRUARY', 'february', 'FeBrUaRy', 'FEbRUARy', 'fEbRUaRy', 'feBRuaRY', 'Feb', 'FEB', 'feb', 'FeB', 'FEb', 'feB', 'fEb'],
			['April', 'APRIL', 'april', 'ApRiL', 'APriL', 'aPrIL', 'apRIl', 'Apr', 'APR', 'apr', 'ApR', 'APr', 'apR', 'aPr'],
			['June', 'JUNE', 'june', 'JuNe', 'JUnE', 'jUnE', 'juNE', 'Jun', 'JUN', 'jun', 'JuN', 'JUn', 'juN', 'jUn']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['August'], ['October'], ['December']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Aug'], ['Oct'], ['Dec']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 9, 0, 12, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 12);

		// Reverse case
		expectedDataCapitalized = [['October'], ['August'], ['June']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Oct'], ['Aug'], ['Jun']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 12, 12, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 12);
	});
	QUnit.test('Autofill - Horizontal odd sequence: Range 10 cells. Months', function (assert) {
		let testData = [
			['January', 'March', 'May', 'July', 'September', 'November', 'January', 'March', 'May', 'July'],
			['JANUARY', 'MARCH', 'MAY', 'JULY', 'SEPTEMBER', 'NOVEMBER', 'JANUARY', 'MARCH', 'MAY', 'JULY'],
			['january', 'march', 'may', 'july', 'september', 'november', 'january', 'march', 'may', 'july'],
			['JaNuArY', 'MaRcH', 'MaY', 'JuLy', 'SePtEmBeR', 'NoVemBeR', 'JaNuArY', 'MaRcH', 'MaY', 'JuLy'],
			['JAnuARy', 'MArCH', 'MAy', 'JUly', 'SEptEMbeR', 'NOveMBer', 'JAnuARy', 'MArCH', 'MAy', 'JUly'],
			['jAnUaRy', 'mArCh', 'mAy', 'jUlY', 'sEpTeMbEr', 'nOveMbEr', 'jAnUaRy', 'mArCh', 'mAy', 'jUlY'],
			['jaNUarY', 'maRCh', 'maY', 'juLY', 'sePTemBEr', 'noVEmbER', 'jaNUarY', 'maRCh', 'maY', 'juLY'],
			['Jan', 'Mar', 'May', 'Jul', 'Sep', 'Nov', 'Jan', 'Mar', 'May', 'Jul'],
			['JAN', 'MAR', 'MAY', 'JUL', 'SEP', 'NOV', 'JAN', 'MAR', 'MAY', 'JUL'],
			['jan', 'mar', 'may', 'jul', 'sep', 'nov', 'jan', 'mar', 'may', 'jul'],
			['JaN', 'MaR', 'MaY', 'JuL', 'SeP', 'NoV', 'JaN', 'MaR', 'MaY', 'JuL'],
			['JAn', 'MAr', 'MAy', 'JUl', 'SEp', 'NOv', 'JAn', 'MAr', 'MAy', 'JUl'],
			['jAn', 'mAr', 'mAy', 'jUl', 'sEp', 'nOv', 'jAn', 'mAr', 'mAy', 'jUl'],
			['jaN', 'maR', 'maY', 'juL', 'seP', 'noV', 'jaN', 'maR', 'maY', 'juL']

		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['September', 'November', 'January'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Sep', 'Nov', 'Jan'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 9, 0, 12,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 12, 0);

		// Reverse case
		expectedDataCapitalized = ['November', 'September', 'July'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Nov', 'Sep', 'Jul'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 12, 12, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 12, 0);
	});
	QUnit.test('Autofill - Vertical odd sequence: Range 10 cells.  Months', function (assert) {
		let testData = [
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
			['March', 'MARCH', 'march', 'MaRcH', 'MArCH', 'mArCh', 'maRcH', 'Mar', 'MAR', 'mar', 'MaR', 'MAr', 'maR', 'mAr'],
			['May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY', 'May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY'],
			['July', 'JULY', 'july', 'JuLy', 'JUly', 'jULY', 'juLY', 'Jul', 'JUL', 'jul', 'JuL', 'JUl', 'juL', 'jUl'],
			['September', 'SEPTEMBER', 'september', 'SePtEmBeR', 'SEptEMbeR', 'sEPTEMBer', 'sePTEMber', 'Sep', 'SEP', 'sep', 'SeP', 'SEp', 'seP', 'sEp'],
			['November', 'NOVEMBER', 'november', 'NoVEmBeR', 'NOvEMbeR', 'nOvEMBeR', 'noVEMber', 'Nov', 'NOV', 'nov', 'NoV', 'NOv', 'noV', 'nOv'],
			['January', 'JANUARY', 'january', 'JaNuArY', 'JAnuARy', 'jAnUaRy', 'jaNUarY', 'Jan', 'JAN', 'jan', 'JaN', 'JAn', 'jaN', 'jAn'],
			['March', 'MARCH', 'march', 'MaRcH', 'MArCH', 'mArCh', 'maRcH', 'Mar', 'MAR', 'mar', 'MaR', 'MAr', 'maR', 'mAr'],
			['May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY', 'May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY'],
			['July', 'JULY', 'july', 'JuLy', 'JUly', 'jULY', 'juLY', 'Jul', 'JUL', 'jul', 'JuL', 'JUl', 'juL', 'jUl'],
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['September'], ['November'], ['January']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Sep'], ['Nov'], ['Jan']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 9, 0, 12, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 12);

		// Reverse case
		expectedDataCapitalized = [['November'], ['September'], ['July']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Nov'], ['Sep'], ['Jul']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 12, 12, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 12);
	});
	QUnit.test('Autofill - Horizontal sequence: May check previous cell in range. Months', function (assert) {
		let testData = [
			['March', 'May'],
			['MARCH', 'MAY'],
			['march', 'may'],
			['MaRcH', 'MaY'],
			['MArCH', 'MAy'],
			['mArCh', 'mAy'],
			['maRCh', 'maY'],
			['Mar', 'May'],
			['MAR', 'MAY'],
			['mar', 'may'],
			['MaR', 'MaY'],
			['MAr', 'MAy'],
			['mAr', 'mAy'],
			['maR', 'maY']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['July', 'September', 'November'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Jul', 'Sep', 'Nov'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 4, 0);

		// Reverse case
		expectedDataCapitalized = ['January', 'November', 'September'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Jan', 'Nov', 'Sep'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 4, 0);
	});
	QUnit.test('Autofill - Vertical sequence: May check previous cell in range. Months', function (assert) {
		let testData = [
			['March', 'MARCH', 'march', 'MaRcH', 'MArCH', 'mArCh', 'maRcH', 'Mar', 'MAR', 'mar', 'MaR', 'MAr', 'maR', 'mAr'],
			['May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY', 'May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['July'], ['September'], ['November']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Jul'], ['Sep'], ['Nov']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 4);

		// Reverse case
		expectedDataCapitalized = [['January'], ['November'], ['September']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Jan'], ['Nov'], ['Sep']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 4);
	});
	QUnit.test('Autofill - Horizontal sequence: May check next cell in range. Months', function (assert) {
		let testData = [
			['May', 'June'],
			['MAY', 'JUNE'],
			['may', 'june'],
			['MaY', 'JuNe'],
			['MAy', 'JUne'],
			['mAy', 'jUnE'],
			['maY', 'juNE'],
			['May', 'Jun'],
			['MAY', 'JUN'],
			['may', 'jun'],
			['MaY', 'JuN'],
			['MAy', 'JUn'],
			['maY', 'juN'],
			['mAy', 'jUn']
		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = ['July', 'August', 'September'];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = ['Jul', 'Aug', 'Sep'];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getHorizontalAutofillCases(0, 1, 0, 4,assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 4, 0);

		// Reverse case
		expectedDataCapitalized = ['April', 'March', 'February'];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = ['Apr', 'Mar', 'Feb'];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(0, 3);
		range.fillData(testData);
		getHorizontalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 4, 0);
	});
	QUnit.test('Autofill - Vertical sequence: May check next cell in range. Months', function (assert) {
		let testData = [
			['May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY', 'May', 'MAY', 'may', 'MaY', 'MAy', 'mAY', 'maY'],
			['June', 'JUNE', 'june', 'JuNe', 'JUnE', 'jUnE', 'juNE', 'Jun', 'JUN', 'jun', 'JuN', 'JUn', 'juN', 'jUn']

		];

		// Asc case
		// Add expected Data after Autofill
		let expectedDataCapitalized = [['July'], ['August'], ['September']];
		let expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		let expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		let expectedDataShortCapitalized = [['Jul'], ['Aug'], ['Sep']];
		let expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		let expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		//nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getVerticalAutofillCases(0, 1, 0, 4, assert, [expectedDataCapitalized,
			expectedDataUpper, expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 3);
		clearData(0, 0, 0, 4);

		// Reverse case
		expectedDataCapitalized = [['April'], ['March'], ['February']];
		expectedDataUpper = updateDataToUpCase(expectedDataCapitalized);
		expectedDataLower = updateDataToLowCase(expectedDataCapitalized);
		expectedDataShortCapitalized = [['Apr'], ['Mar'], ['Feb']];
		expectedDataShortUpper = updateDataToUpCase(expectedDataShortCapitalized);
		expectedDataShortLower = updateDataToLowCase(expectedDataShortCapitalized);

		range = ws.getRange4(3, 0);
		range.fillData(testData);
		getVerticalAutofillCases(3, 4, 4, 0, assert, [expectedDataCapitalized, expectedDataUpper,
			expectedDataLower, expectedDataShortCapitalized, expectedDataShortUpper, expectedDataShortLower], 1);
		clearData(0, 0, 0, 4);
	});
	/* TODO
	 * Not correct behavior for autofill Date for month and years compared to ms excel
	 * Context: If we try to fill 2 cells data e.g.  01.01.2000 and 01.02.2000 and try use autofill for this data.
	 * We'll get difference data compared to ms excel.
	 *
	 * Repro for month:
	 * 1. Fill data 01.01.2000 and 01.02.2000
	 * 2. Select range for filled data
	 * 3. Try to use autofill for 2 cells (asc sequence)
	 * Expected result:
	 *After used autofill we'll get 01.03.2000, 01.04.2000
	 * Actual result:
	 * After used autofill we'll get 03.03.2020, 03.04.2020
	 *
	 * Repro for year
	 * 1. Fill data 01.01.2000 and 01.01.2001
	 * 2. Select range for filled data
	 * 3. Try to use autofill for 2 cells (asc sequence)
	 * Expected result:
	 * After used autofill we'll get 01.01.2002, 01.01.2003
	 * Actual result:
	 * After used autofill we'll get 03.01.2003, 04.01.2004
	 */
	QUnit.test('Autofill - Horizontal sequence.', function (assert) {
		function getAutofillCase(aFrom, aTo, nFillHandleArea, sDescription, expectedData) {
			const [c1From, c2From, rFrom] = aFrom;
			const [c1To, c2To, rTo] = aTo;
			const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
			const autofillC1 =  nFillHandleArea === 3 ? c2From + 1 : c1From - 1;
			const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;

			ws.selectionRange.ranges = [getRange(c1From, rFrom, c2From, rFrom)];
			wsView = getAutoFillRange(wsView, c1To, rTo, c2To, rTo, nHandleDirection, nFillHandleArea);
			let autoFillRange = getRange(autofillC1, rTo, c2To, rTo);
			autoFillAssert(assert, autoFillRange, [expectedData], sDescription);
		}
		const testData = [
			['-1'],
			['-1', '0'],
			['1', '3'],
			['2', '4'],
			['Test'],
			['Test01'],
			['Test1'],
			['Test1', 'Test3'],
			['Test2', 'Test4'],
			['Test1', 'T1'],
			['01/01/2000'],
			['01/01/2000', '01/02/2000'],
			['01/02/2000', '01/04/2000'],
			['01/01/2000', '01/03/2000']
		];

		// Asc cases
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([0, 0, 0], [0, 1, 0], 3, 'Number. Asc sequence. Range 1 cell', ['-1']);
		getAutofillCase([0, 1, 1], [0, 4, 1], 3, 'Number. Asc sequence. Range 2 cell', ['1', '2', '3']);
		getAutofillCase([0, 1, 2], [0, 4, 2], 3, 'Number. Asc odd sequence. Range 2 cell', ['5', '7', '9']);
		getAutofillCase([0, 1, 3], [0, 4, 3], 3, 'Number. Asc even sequence. Range 2 cell', ['6', '8', '10']);
		getAutofillCase([0, 0, 4], [0, 1, 4], 3, 'Text. Asc sequence. Range 1 cell', ['Test']);
		getAutofillCase([0, 0, 5], [0, 3, 5], 3, 'Text with postfix 01. Asc sequence. Range 1 cell', ['Test02', 'Test03', 'Test04']);
		getAutofillCase([0, 0, 6], [0, 3, 6], 3, 'Text with postfix 1. Asc sequence. Range 1 cell', ['Test2', 'Test3', 'Test4']);
		getAutofillCase([0, 1, 7], [0, 4, 7], 3, 'Text with postfix. Asc odd sequence. Range 2 cell', ['Test5', 'Test7', 'Test9']);
		getAutofillCase([0, 1, 8], [0, 4, 8], 3, 'Text with postfix. Asc even sequence. Range 2 cell', ['Test6', 'Test8', 'Test10']);
		getAutofillCase([0, 1, 9], [0, 5, 9], 3, 'Text with postfix. Asc sequence of Test and T. Range 2 cell', ['Test2', 'T2', 'Test3', 'T3']);
		getAutofillCase([0, 0, 10], [0, 3, 10], 3, 'Date. Asc sequence. Range 1 cell', ['36527', '36528', '36529']); // 02.01.2000, 03.01.2000, 04.01.2000
		getAutofillCase([0, 1, 11], [0, 4, 11], 3, 'Date. Asc sequence. Range 2 cell', ['36528', '36529', '36530']); // 03.01.2000, 04.01.2000, 05.01.2000
		getAutofillCase([0, 1, 12], [0, 4, 12], 3, 'Date. Asc even sequence. Range 2 cell', ['36531', '36533', '36535']); // 06.01.2000, 08.01.2000, 10.01.2000
		getAutofillCase([0, 1, 13], [0, 4, 13], 3, 'Date. Asc odd sequence. Range 2 cell', ['36530', '36532', '36534']); // 05.01.2000, 07.01.2000, 09.01.2000

		clearData(0, 0, 5, 13);
		// Reverse cases
		range = ws.getRange4(0, 3);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([3, 3, 0], [2, 3, 0], 1, 'Number. Reverse sequence. Range 1 cell', ['-1']);
		getAutofillCase([3, 4, 1], [4, 0, 1], 1, 'Number. Reverse sequence. Range 2 cell', ['-2', '-3', '-4']);
		getAutofillCase([3, 4, 2], [4, 0, 2], 1, 'Number. Reverse odd sequence. Range 2 cell', ['-1', '-3', '-5']);
		getAutofillCase([3, 4, 3], [4, 0, 3], 1, 'Number. Reverse even sequence. Range 2 cell', ['0', '-2', '-4']);
		getAutofillCase([3, 3, 4], [2, 3, 4], 1, 'Text. Reverse sequence. Range 1 cell', ['Test']);
		getAutofillCase([3, 3, 5], [3, 0, 5], 1, 'Text with postfix 01. Reverse sequence. Range 1 cell', ['Test00', 'Test01', 'Test02']);
		getAutofillCase([3, 3, 6], [3, 0, 6], 1, 'Text with postfix 1. Reverse sequence. Range 1 cell', ['Test0', 'Test1', 'Test2']);
		getAutofillCase([3, 4, 7], [4, 0, 7], 1, 'Text with postfix. Reverse odd sequence. Range 2 cell', ['Test1', 'Test3', 'Test5']);
		getAutofillCase([3, 4, 8], [4, 0, 8], 1, 'Text with postfix. Reverse even sequence. Range 2 cell', ['Test0', 'Test2', 'Test4']);
		getAutofillCase([3, 4, 9], [4, 0, 9], 1, 'Text with postfix. Reverse sequence of Test and T. Range 2 cell', ['T0', 'Test0', 'T1']);
		getAutofillCase([3, 3, 10], [3, 0, 10], 1, 'Date. Reverse sequence. Range 1 cell', ['36525', '36524', '36523']); // 31.12.1999, 30.12.1999, 29.12.1999
		getAutofillCase([3, 4, 11], [4, 0, 11], 1, 'Date. Reverse sequence. Range 2 cell', ['36525', '36524', '36523']); // 31.12.1999, 30.12.1999, 29.12.1999
		getAutofillCase([3, 4, 12], [4, 0, 12], 1, 'Date. Reverse even sequence. Range 2 cell', ['36525', '36523', '36521']); // 30.12.1999, 28.12.1999, 26.12.1999
		getAutofillCase([3, 4, 13], [4, 0, 13], 1, 'Date. Reverse odd sequence. Range 2 cell', ['36524', '36522', '36520']); // 31.12.1999, 29.12.1999, 27.12.1999
		clearData(0, 0, 4, 13);

	});
	QUnit.test('Autofill - Vertical sequence.', function (assert) {
		function getAutofillCase(aFrom, aTo, nFillHandleArea, sDescription, expectedData) {
			const [r1From, r2From, cFrom] = aFrom;
			const [r1To, r2To, cTo] = aTo;
			const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
			const autofillR1 =  nFillHandleArea === 3 ? r2From + 1 : r1From - 1;
			const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;

			ws.selectionRange.ranges = [getRange(cFrom, r1From, cFrom, r2From)];
			wsView = getAutoFillRange(wsView, cTo, r1To, cTo, r2To, nHandleDirection, nFillHandleArea);
			let autoFillRange = getRange(cTo, autofillR1, cTo, r2To);
			autoFillAssert(assert, autoFillRange, expectedData, sDescription);
		}
		const testData = [
			['-1', '-1', '1', '2', 'Test', 'Test01', 'Test1', 'Test1', 'Test2', 'Test1', '01/01/2000', '01/01/2000', '01/02/2000', '01/01/2000'],
			['', '0', '3', '4', '', '', '', 'Test3', 'Test4', 'T1', '', '01/02/2000', '01/04/2000', '01/03/2000']
		];

		// Asc cases
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([0, 0, 0], [0, 1, 0], 3, 'Number. Asc sequence. Range 1 cell', [['-1']]);
		getAutofillCase([0, 1, 1], [0, 4, 1], 3, 'Number. Asc sequence. Range 2 cell', [['1'], ['2'], ['3']]);
		getAutofillCase([0, 1, 2], [0, 4, 2], 3, 'Number. Asc odd sequence. Range 2 cell', [['5'], ['7'], ['9']]);
		getAutofillCase([0, 1, 3], [0, 4, 3], 3, 'Number. Asc even sequence. Range 2 cell', [['6'], ['8'], ['10']]);
		getAutofillCase([0, 0, 4], [0, 1, 4], 3, 'Text. Asc sequence. Range 1 cell', [['Test']]);
		getAutofillCase([0, 0, 5], [0, 3, 5], 3, 'Text with postfix 01. Asc sequence. Range 1 cell', [['Test02'], ['Test03'], ['Test04']]);
		getAutofillCase([0, 0, 6], [0, 3, 6], 3, 'Text with postfix 1. Asc sequence. Range 1 cell', [['Test2'], ['Test3'], ['Test4']]);
		getAutofillCase([0, 1, 7], [0, 4, 7], 3, 'Text with postfix. Asc odd sequence. Range 2 cell', [['Test5'], ['Test7'], ['Test9']]);
		getAutofillCase([0, 1, 8], [0, 4, 8], 3, 'Text with postfix. Asc even sequence. Range 2 cell', [['Test6'], ['Test8'], ['Test10']]);
		getAutofillCase([0, 1, 9], [0, 5, 9], 3, 'Text with postfix. Asc sequence of Test and T. Range 2 cell', [['Test2'], ['T2'], ['Test3'], ['T3']]);
		getAutofillCase([0, 0, 10], [0, 3, 10], 3, 'Date. Asc sequence. Range 1 cell', [['36527'], ['36528'], ['36529']]); // 02.01.2000, 03.01.2000, 04.01.2000
		getAutofillCase([0, 1, 11], [0, 4, 11], 3, 'Date. Asc sequence. Range 2 cell', [['36528'], ['36529'], ['36530']]); // 03.01.2000, 04.01.2000, 05.01.2000
		getAutofillCase([0, 1, 12], [0, 4, 12], 3, 'Date. Asc even sequence. Range 2 cell', [['36531'], ['36533'], ['36535']]); // 06.01.2000, 08.01.2000, 10.01.2000
		getAutofillCase([0, 1, 13], [0, 4, 13], 3, 'Date. Asc odd sequence. Range 2 cell', [['36530'], ['36532'], ['36534']]); // 05.01.2000, 07.01.2000, 09.01.2000

		clearData(0, 0, 5, 13);
		// Reverse cases
		range = ws.getRange4(3, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([3, 3, 0], [2, 3, 0], 1, 'Number. Reverse sequence. Range 1 cell', [['-1']]);
		getAutofillCase([3, 4, 1], [4, 0, 1], 1, 'Number. Reverse sequence. Range 2 cell', [['-2'], ['-3'], ['-4']]);
		getAutofillCase([3, 4, 2], [4, 0, 2], 1, 'Number. Reverse odd sequence. Range 2 cell', [['-1'], ['-3'], ['-5']]);
		getAutofillCase([3, 4, 3], [4, 0, 3], 1, 'Number. Reverse even sequence. Range 2 cell', [['0'], ['-2'], ['-4']]);
		getAutofillCase([3, 3, 4], [2, 3, 4], 1, 'Text. Reverse sequence. Range 1 cell', [['Test']]);
		getAutofillCase([3, 3, 5], [3, 0, 5], 1, 'Text with postfix 01. Reverse sequence. Range 1 cell', [['Test00'], ['Test01'], ['Test02']]);
		getAutofillCase([3, 3, 6], [3, 0, 6], 1, 'Text with postfix 1. Reverse sequence. Range 1 cell', [['Test0'], ['Test1'], ['Test2']]);
		getAutofillCase([3, 4, 7], [4, 0, 7], 1, 'Text with postfix. Reverse odd sequence. Range 2 cell', [['Test1'], ['Test3'], ['Test5']]);
		getAutofillCase([3, 4, 8], [4, 0, 8], 1, 'Text with postfix. Reverse even sequence. Range 2 cell', [['Test0'], ['Test2'], ['Test4']]);
		getAutofillCase([3, 4, 9], [4, 0, 9], 1, 'Text with postfix. Reverse sequence of Test and T. Range 2 cell', [['T0'], ['Test0'], ['T1']]);
		getAutofillCase([3, 3, 10], [3, 0, 10], 1, 'Date. Reverse sequence. Range 1 cell', [['36525'], ['36524'], ['36523']]); // 31.12.1999, 30.12.1999, 29.12.1999
		getAutofillCase([3, 4, 11], [4, 0, 11], 1, 'Date. Reverse sequence. Range 2 cell', [['36525'], ['36524'], ['36523']]); // 31.12.1999, 30.12.1999, 29.12.1999
		getAutofillCase([3, 4, 12], [4, 0, 12], 1, 'Date. Reverse even sequence. Range 2 cell', [['36525'], ['36523'], ['36521']]); // 30.12.1999, 28.12.1999, 26.12.1999
		getAutofillCase([3, 4, 13], [4, 0, 13], 1, 'Date. Reverse odd sequence. Range 2 cell', [['36524'], ['36522'], ['36520']]); // 31.12.1999, 29.12.1999, 27.12.1999

		clearData(0, 0, 4, 13);
	});
	QUnit.test('Autofill: Days of week and months with spaces and "." - Horizontal sequence', function (assert) {
		function getAutofillCase(aFrom, aTo, nFillHandleArea, sDescription, expectedData) {
			const [c1From, c2From, rFrom] = aFrom;
			const [c1To, c2To, rTo] = aTo;
			const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
			const autofillC1 =  nFillHandleArea === 3 ? c2From + 1 : c1From - 1;
			const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;

			ws.selectionRange.ranges = [getRange(c1From, rFrom, c2From, rFrom)];
			wsView = getAutoFillRange(wsView, c1To, rTo, c2To, rTo, nHandleDirection, nFillHandleArea);
			let autoFillRange = getRange(autofillC1, rTo, c2To, rTo);
			autoFillAssert(assert, autoFillRange, [expectedData], sDescription);
		}
		const testData = [
			['monday '],
			['monday ', 'tuesday'],
			[' monday ', 'tuesday '],
			['mon.'],
			['mon.', 'tue'],
			['mon ', 'tue'],
			[' mon ', 'tue '],
			['january '],
			['january ', 'february'],
			[' january ', 'february'],
			['jan.'],
			['jan.', 'feb'],
			['mon.day']
		];
		// Asc cases
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([0, 0, 0], [0, 6, 0], 3, 'Day of week with space. Asc sequence. Range 1 cell', ['tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']);
		getAutofillCase([0, 1, 1], [0, 5, 1], 3, 'Day of week with space. Asc sequence. Range 2 cell', ['wednesday', 'thursday', 'friday', 'saturday', 'sunday']);
		getAutofillCase([0, 1, 2], [0, 5, 2], 3, 'Day of week with space. Asc sequence. Range 2 cell', ['wednesday', 'thursday', 'friday', 'saturday', 'sunday']);
		getAutofillCase([0, 0, 3], [0, 6, 3], 3, 'Day of week short with ".". Asc sequence. Range 1 cell', ['tue', 'wed', 'thu', 'fri', 'sat', 'sun']);
		getAutofillCase([0, 1, 4], [0, 5, 4], 3, 'Day of week short with ".". Asc sequence. Range 2 cell', ['wed', 'thu', 'fri', 'sat', 'sun']);
		getAutofillCase([0, 1, 5], [0, 5, 5], 3, 'Day of week short with space. Asc sequence. Range 2 cell', ['wed', 'thu', 'fri', 'sat', 'sun']);
		getAutofillCase([0, 1, 6], [0, 5, 6], 3, 'Day of week short with spaces. Asc sequence. Range 2 cell', ['wed', 'thu', 'fri', 'sat', 'sun']);
		getAutofillCase([0, 0, 7], [0, 6, 7], 3, 'Month with space. Asc sequence. Range 1 cell', ['february', 'march', 'april', 'may', 'june', 'july']);
		getAutofillCase([0, 1, 8], [0, 6, 8], 3, 'Month with space. Asc sequence. Range 2 cell', ['march', 'april', 'may', 'june', 'july', 'august']);
		getAutofillCase([0, 1, 9], [0, 6, 9], 3, 'Month with spaces. Asc sequence. Range 2 cell', ['march', 'april', 'may', 'june', 'july', 'august']);
		getAutofillCase([0, 0, 10], [0, 6, 10], 3, 'Month short with ".". Asc sequence. Range 1 cell', ['feb','mar', 'apr', 'may', 'jun', 'jul']);
		getAutofillCase([0, 1, 11], [0, 6, 11], 3, 'Month short with ".". Asc sequence. Range 2 cell', ['mar', 'apr', 'may', 'jun', 'jul', 'aug']);
		getAutofillCase([0, 0, 12], [0, 2, 12], 3, 'mon.day. Asc sequence. Range 1 cell', ['mon.day', 'mon.day']);
		clearData(0, 0, 6, 12);
		// Reverse cases
		range = ws.getRange4(0, 7);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([7, 7, 0], [6, 0, 0], 1, 'Day of week with space. Reverse sequence. Range 1 cell', ['sunday', 'saturday', 'friday', 'thursday', 'wednesday', 'tuesday', 'monday']);
		getAutofillCase([7, 8, 1], [6, 0, 1], 1, 'Day of week with space. Reverse sequence. Range 2 cell', ['sunday', 'saturday', 'friday', 'thursday', 'wednesday', 'tuesday', 'monday']);
		getAutofillCase([7, 8, 2], [6, 0, 2], 1, 'Day of week with space. Reverse sequence. Range 2 cell', ['sunday', 'saturday', 'friday', 'thursday', 'wednesday', 'tuesday', 'monday']);
		getAutofillCase([7, 7, 3], [6, 0, 3], 1, 'Day of week short with ".". Reverse sequence. Range 1 cell', ['sun', 'sat', 'fri', 'thu', 'wed', 'tue', 'mon']);
		getAutofillCase([7, 8, 4], [6, 0, 4], 1, 'Day of week short with ".". Reverse sequence. Range 2 cell', ['sun', 'sat', 'fri', 'thu', 'wed', 'tue', 'mon']);
		getAutofillCase([7, 8, 5], [6, 0, 5], 1, 'Day of week short with space. Reverse sequence. Range 2 cell', ['sun', 'sat', 'fri', 'thu', 'wed', 'tue', 'mon']);
		getAutofillCase([7, 8, 6], [6, 0, 6], 1, 'Day of week short with spaces. Reverse sequence. Range 2 cell', ['sun', 'sat', 'fri', 'thu', 'wed', 'tue', 'mon']);
		getAutofillCase([7, 7, 7], [6, 0, 7], 1, 'Month with space. Reverse sequence. Range 1 cell', ['december', 'november', 'october', 'september', 'august', 'july', 'june']);
		getAutofillCase([7, 8, 8], [6, 0, 8], 1, 'Month with space. Reverse sequence. Range 2 cell', ['december', 'november', 'october', 'september', 'august', 'july', 'june']);
		getAutofillCase([7, 8, 9], [6, 0, 9], 1, 'Month with spaces. Reverse sequence. Range 2 cell', ['december', 'november', 'october', 'september', 'august', 'july', 'june']);
		getAutofillCase([7, 7, 10], [6, 0, 10], 1, 'Month short with ".". Reverse sequence. Range 1 cell', ['dec', 'nov', 'oct', 'sep', 'aug', 'jul', 'jun']);
		getAutofillCase([7, 8, 11], [6, 0, 11], 1, 'Month short with ".". Reverse sequence. Range 2 cell', ['dec', 'nov', 'oct', 'sep', 'aug', 'jul', 'jun']);
		getAutofillCase([7, 7, 12], [6, 5, 12], 1, 'mon.day. Reverse sequence. Range 1 cell', ['mon.day', 'mon.day']);
		clearData(0, 0, 8, 12);
	});
	QUnit.test('Autofill: Days of week and months with spaces and "." - Vertical sequence.', function (assert) {
		function getAutofillCase(aFrom, aTo, nFillHandleArea, sDescription, expectedData) {
			const [r1From, r2From, cFrom] = aFrom;
			const [r1To, r2To, cTo] = aTo;
			const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
			const autofillR1 =  nFillHandleArea === 3 ? r2From + 1 : r1From - 1;
			const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;

			ws.selectionRange.ranges = [getRange(cFrom, r1From, cFrom, r2From)];
			wsView = getAutoFillRange(wsView, cTo, r1To, cTo, r2To, nHandleDirection, nFillHandleArea);
			let autoFillRange = getRange(cTo, autofillR1, cTo, r2To);
			autoFillAssert(assert, autoFillRange, expectedData, sDescription);
		}
		const testData = [
			['monday ', 'monday ', ' monday ', 'mon.', 'mon.', 'mon ', ' mon ', 'january ', 'january ', ' january ', 'jan.', 'jan.', 'mon.day'],
			['', 'tuesday', 'tuesday ', '', 'tue', 'tue', 'tue ', '', 'february', 'february', '', 'feb', '']
		];

		// Asc cases
		let range = ws.getRange4(0, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([0, 0, 0], [0, 6, 0], 3, 'Day of week with space. Asc sequence. Range 1 cell', [['tuesday'], ['wednesday'], ['thursday'], ['friday'], ['saturday'], ['sunday']]);
		getAutofillCase([0, 1, 1], [0, 5, 1], 3, 'Day of week with space. Asc sequence. Range 2 cell', [['wednesday'], ['thursday'], ['friday'], ['saturday'], ['sunday']]);
		getAutofillCase([0, 1, 2], [0, 5, 2], 3, 'Day of week with space. Asc sequence. Range 2 cell', [['wednesday'], ['thursday'], ['friday'], ['saturday'], ['sunday']]);
		getAutofillCase([0, 0, 3], [0, 6, 3], 3, 'Day of week short with ".". Asc sequence. Range 1 cell', [['tue'], ['wed'], ['thu'], ['fri'], ['sat'], ['sun']]);
		getAutofillCase([0, 1, 4], [0, 5, 4], 3, 'Day of week short with ".". Asc sequence. Range 2 cell', [['wed'], ['thu'], ['fri'], ['sat'], ['sun']]);
		getAutofillCase([0, 1, 5], [0, 5, 5], 3, 'Day of week short with space. Asc sequence. Range 2 cell', [['wed'], ['thu'], ['fri'], ['sat'], ['sun']]);
		getAutofillCase([0, 1, 6], [0, 5, 6], 3, 'Day of week short with spaces. Asc sequence. Range 2 cell', [['wed'], ['thu'], ['fri'], ['sat'], ['sun']]);
		getAutofillCase([0, 0, 7], [0, 6, 7], 3, 'Month with space. Asc sequence. Range 1 cell', [['february'], ['march'], ['april'], ['may'], ['june'], ['july']]);
		getAutofillCase([0, 1, 8], [0, 6, 8], 3, 'Month with space. Asc sequence. Range 2 cell', [['march'], ['april'], ['may'], ['june'], ['july'], ['august']]);
		getAutofillCase([0, 1, 9], [0, 6, 9], 3, 'Month with spaces. Asc sequence. Range 2 cell', [['march'], ['april'], ['may'], ['june'], ['july'], ['august']]);
		getAutofillCase([0, 0, 10], [0, 6, 10], 3, 'Month short with ".". Asc sequence. Range 1 cell', [['feb'],['mar'], ['apr'], ['may'], ['jun'], ['jul']]);
		getAutofillCase([0, 1, 11], [0, 6, 11], 3, 'Month short with ".". Asc sequence. Range 2 cell', [['mar'], ['apr'], ['may'], ['jun'], ['jul'], ['aug']]);
		getAutofillCase([0, 0, 12], [0, 2, 12], 3, 'mon.day. Asc sequence. Range 1 cell', [['mon.day'], ['mon.day']]);
		clearData(0, 0, 6, 11);
		// Reverse cases
		range = ws.getRange4(7, 0);
		range.fillData(testData);
		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		getAutofillCase([7, 7, 0], [6, 0, 0], 1, 'Day of week with space. Reverse sequence. Range 1 cell', [['sunday'], ['saturday'], ['friday'], ['thursday'], ['wednesday'], ['tuesday'], ['monday']]);
		getAutofillCase([7, 8, 1], [6, 0, 1], 1, 'Day of week with space. Reverse sequence. Range 2 cell', [['sunday'], ['saturday'], ['friday'], ['thursday'], ['wednesday'], ['tuesday'], ['monday']]);
		getAutofillCase([7, 8, 2], [6, 0, 2], 1, 'Day of week with space. Reverse sequence. Range 2 cell', [['sunday'], ['saturday'], ['friday'], ['thursday'], ['wednesday'], ['tuesday'], ['monday']]);
		getAutofillCase([7, 7, 3], [6, 0, 3], 1, 'Day of week short with ".". Reverse sequence. Range 1 cell', [['sun'], ['sat'], ['fri'], ['thu'], ['wed'], ['tue'], ['mon']]);
		getAutofillCase([7, 8, 4], [6, 0, 4], 1, 'Day of week short with ".". Reverse sequence. Range 2 cell', [['sun'], ['sat'], ['fri'], ['thu'], ['wed'], ['tue'], ['mon']]);
		getAutofillCase([7, 8, 5], [6, 0, 5], 1, 'Day of week short with space. Reverse sequence. Range 2 cell', [['sun'], ['sat'], ['fri'], ['thu'], ['wed'], ['tue'], ['mon']]);
		getAutofillCase([7, 8, 6], [6, 0, 6], 1, 'Day of week short with spaces. Reverse sequence. Range 2 cell', [['sun'], ['sat'], ['fri'], ['thu'], ['wed'], ['tue'], ['mon']]);
		getAutofillCase([7, 7, 7], [6, 0, 7], 1, 'Month with space. Reverse sequence. Range 1 cell', [['december'], ['november'], ['october'], ['september'], ['august'], ['july'], ['june']]);
		getAutofillCase([7, 8, 8], [6, 0, 8], 1, 'Month with space. Reverse sequence. Range 2 cell', [['december'], ['november'], ['october'], ['september'], ['august'], ['july'], ['june']]);
		getAutofillCase([7, 8, 9], [6, 0, 9], 1, 'Month with spaces. Reverse sequence. Range 2 cell', [['december'], ['november'], ['october'], ['september'], ['august'], ['july'], ['june']]);
		getAutofillCase([7, 7, 10], [6, 0, 10], 1, 'Month short with ".". Reverse sequence. Range 1 cell', [['dec'], ['nov'], ['oct'], ['sep'], ['aug'], ['jul'], ['jun']]);
		getAutofillCase([7, 8, 11], [6, 0, 11], 1, 'Month short with ".". Reverse sequence. Range 2 cell', [['dec'], ['nov'], ['oct'], ['sep'], ['aug'], ['jul'], ['jun']]);
		getAutofillCase([7, 7, 12], [6, 5, 12], 1, 'mon.day. Reverse sequence. Range 1 cell', [['mon.day'], ['mon.day']]);
		clearData(0, 0, 8, 11);
	});

	QUnit.test('Autofill: test toolbar down/up/left/right', function (assert) {
		clearData(0, 0, 6, 11);

		const testData = [
			['1', 'Test', 'Test1', '01/01/2000']
		];

		// Asc cases
		let range = ws.getRange4(0, 0);
		range.fillData(testData);

		let fillRange = new Asc.Range(0, 0, 0, 3);
		wsView.setSelection(fillRange);
		api.asc_FillCells(Asc.c_oAscFillType.fillDown);

		checkUndoRedo(function (_desc) {
			compareData(assert, fillRange, [["1"], [""], [""], [""]], _desc);
		}, function (_desc) {
			compareData(assert, fillRange, [["1"], ["1"], ["1"], ["1"]], _desc);
		}, "Autofill: down fill number");

		fillRange = new Asc.Range(1, 0, 1, 3);
		wsView.setSelection(fillRange);
		api.asc_FillCells(Asc.c_oAscFillType.fillDown);

		checkUndoRedo(function (_desc) {
			compareData(assert, fillRange, [["Test"], [""], [""], [""]], _desc);
		}, function (_desc) {
			compareData(assert, fillRange, [["Test"], ["Test"], ["Test"], ["Test"]], _desc);
		}, "Autofill: down fill text");

		fillRange = new Asc.Range(2, 0, 2, 3);
		wsView.setSelection(fillRange);
		api.asc_FillCells(Asc.c_oAscFillType.fillDown);

		checkUndoRedo(function (_desc) {
			compareData(assert, fillRange, [["Test1"], [""], [""], [""]], _desc);
		}, function (_desc) {
			compareData(assert, fillRange, [["Test1"], ["Test1"], ["Test1"], ["Test1"]], _desc);
		}, "Autofill: down fill text + number");


		fillRange = new Asc.Range(3, 0, 3, 3);
		wsView.setSelection(fillRange);
		api.asc_FillCells(Asc.c_oAscFillType.fillDown);

		checkUndoRedo(function (_desc) {
			compareData(assert, fillRange, [["36526"], [""], [""], [""]], _desc);
		}, function (_desc) {
			compareData(assert, fillRange, [["36526"], ["36526"], ["36526"], ["36526"]], _desc);
		}, "Autofill: down fill date");


		fillRange = new Asc.Range(2, 0, 4, 0);
		wsView.setSelection(fillRange);
		api.asc_FillCells(Asc.c_oAscFillType.fillRight);

		checkUndoRedo(function (_desc) {
			compareData(assert, fillRange, [["Test1", "36526", "", ""]], _desc);
		}, function (_desc) {
			compareData(assert, fillRange, [["Test1", "Test1", "Test1", "Test1"]], _desc);
		}, "Autofill: right fill text + number");


		clearData(0, 0, 6, 11);

	});

	QUnit.test('Table selection for formula', function (assert) {

		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		tableOptions.range = "A100:C103";
		api.asc_addAutoFilter("TableStyleMedium2", tableOptions);

		let tables = wsView.model.autoFilters.getTablesIntersectionRange(new Asc.Range(0, 100, 0, 100));
		assert.strictEqual(tables.length, 1, "compare tables length");

		let table = tables[0];
		let tableName = table.DisplayName;
		let activeCell = new AscCommon.CellBase(10, 10);
		let handleSelectionRange = new Asc.Range(0, 1, 0, 1);
		let sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, null, "check selection not table");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 0, 100);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, null, "check selection not table_2");


		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 0, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[Column1]", "check selection column1");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 1, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[Column1]:[Column2]]", "check selection table data from column1 to column2");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 2, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName, "check selection all table");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 2, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName, "check selection table data from column1 to column2");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 1, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#All],[Column1]:[Column2]]", "check selection table data from column1 to column2 + header");


		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 2, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[#All]", "check all selection table");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 2, 99);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[#Headers]", "check all selection table");

		assert.strictEqual(table.isTotalsRow(), false, "check total before total added");
		wsView.af_changeFormatTableInfo(tableName, Asc.c_oAscChangeTableStyleInfo.rowTotal, true);
		assert.strictEqual(table.isTotalsRow(), true, "check total added");

		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 2, 104);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#Data],[#Totals]]", "check data + totals selection table");

		//Table5[[#Data];[#Totals];[Column1]:[Column2]]
		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 100, 1, 104);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#Data],[#Totals],[Column1]:[Column2]]", "check data + totals selection table");

		//Table5[[#All];[Column1]]
		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 0, 104);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#All],[Column1]]", "check all column1 selection table");


		//Table5[[#All];[Column1]:[Column2]]
		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 1, 104);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#All],[Column1]:[Column2]]", "check all column1:column2 selection table");

		//Table5[[#Headers];[#Data];[Column1]:[Column2]]
		activeCell = new AscCommon.CellBase(10, 10);
		handleSelectionRange = new Asc.Range(0, 99, 1, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#Headers],[#Data],[Column1]:[Column2]]", "check headers + data + column1:column2 selection table");

		//@
		//Table5[@]
		activeCell = new AscCommon.CellBase(101, 4);
		handleSelectionRange = new Asc.Range(0, 101, 2, 101);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);
	
		assert.strictEqual(sTableData, tableName + "[@]", "check intersection all row");

		//Table5[@[Column1]:[Column2]]
		activeCell = new AscCommon.CellBase(101, 4);
		handleSelectionRange = new Asc.Range(0, 101, 1, 101);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);
		
		assert.strictEqual(sTableData, tableName + "[@[Column1]:[Column2]]", "check intersection column1:column2 row");

		//Table5[@Column1]
		activeCell = new AscCommon.CellBase(101, 4);
		handleSelectionRange = new Asc.Range(0, 101, 0, 101);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);
		
		assert.strictEqual(sTableData, tableName + "[@Column1]", "check intersection column1 row");

		//Table5[#Headers]
		activeCell = new AscCommon.CellBase(99, 4);
		handleSelectionRange = new Asc.Range(0, 99, 2, 99);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[#Headers]", "check selection Headers");

		// Table5[[#Headers],[#Data]]
		activeCell = new AscCommon.CellBase(99, 4);
		handleSelectionRange = new Asc.Range(0, 99, 2, 103);
		sTableData = table.getSelectionString(activeCell, handleSelectionRange);

		assert.strictEqual(sTableData, tableName + "[[#Headers],[#Data]]", "check selection Headers Data");


		clearData(0, 99, 0, 105);
	});

	QUnit.test('Table values/values for edit tests', function (assert) {
		/*This test checks whether the string is parsed and changed correctly when working with tables */
		let array;
		ws.getRange2("A100:C103").setValue("1");

		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		tableOptions.range = "A100:C103";
		api.asc_addAutoFilter("TableStyleMedium2", tableOptions);	// create table in A100:C103 range

		let tables = wsView.model.autoFilters.getTablesIntersectionRange(new Asc.Range(0, 100, 0, 100));
		assert.strictEqual(tables.length, 1, "compare tables length");

		let table = tables[0];
		let tableName = table.DisplayName;	// due to the fact that other tables are used in file, get the name of the one we need by this way
		wsView.af_changeFormatTableInfo(tableName, Asc.c_oAscChangeTableStyleInfo.rowTotal, true);

		// calc res check
		let cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 5);
		let oParser = new AscCommonExcel.parserFormula(tableName + "[@]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[@][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[@][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), 1, 'Result of Table[@][0,2]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 5);
		resCell.setValue("=" + tableName +"[@]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@]", "Value for edit in cell after Table[@] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#This Row]", "Formula in cell after Table[@] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 10);
		oParser = new AscCommonExcel.parserFormula(tableName + "[#This Row]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[#This Row][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[#This Row][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), 1, 'Result of Table[#This Row][0,2]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 10);
		resCell.setValue("=" + tableName +"[#This Row]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@]", "Value for edit in cell after Table[#This Row] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#This Row]", "Formula in cell after Table[#This Row] is typed");

		// =Table[[#This Row]] => =Table[@]
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 15);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[[#This Row]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[[#This Row]][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), 1, 'Result of Table[[#This Row]][0,2]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 15);
		resCell.setValue("=" + tableName +"[[#This Row]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@]", "Value for edit in cell after Table[@] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#This Row]", "Formula in cell after Table[@] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 20);
		oParser = new AscCommonExcel.parserFormula(tableName + "[@Column1]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Result of Table[@Column1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 20);
		resCell.setValue("=" + tableName +"[@Column1]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@Column1]", "Value for edit in cell after Table[@Column1] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]]", "Formula in cell after Table[@Column1] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 25);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row],[Column1]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Result of Table[[#This Row],[Column1]]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 25);
		resCell.setValue("=" + tableName +"[[#This Row],[Column1]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@Column1]", "Value for edit in cell after Table[[#This Row],[Column1]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]]", "Formula in cell after Table[[#This Row],[Column1]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 30);
		oParser = new AscCommonExcel.parserFormula(tableName + "[@[Column1]:[Column2]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[@[Column1]:[Column2]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[@[Column1]:[Column2]][0,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 30);
		resCell.setValue("=" + tableName +"[@[Column1]:[Column2]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@[Column1]:[Column2]]", "Value for edit in cell after Table[@[Column1]:[Column2]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]:[Column2]]", "Formula in cell after Table[@[Column1]:[Column2]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 35);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row],[Column1]:[Column2]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[[#This Row],[Column1]:[Column2]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[[#This Row],[Column1]:[Column2]][0,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 35);
		resCell.setValue("=" + tableName +"[[#This Row],[Column1]:[Column2]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@[Column1]:[Column2]]", "Value for edit in cell after Table[[#This Row],[Column1]:[Column2]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]:[Column2]]", "Formula in cell after Table[[#This Row],[Column1]:[Column2]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 40);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row],[Column1]:[Column2]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[[#This Row],[Column1]:[Column2]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[[#This Row],[Column1]:[Column2]][0,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 40);
		resCell.setValue("=" + tableName +"[[#This Row],[Column1]:[Column2]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@[Column1]:[Column2]]", "Value for edit in cell after Table[[#This Row],[Column1]:[Column2]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]:[Column2]]", "Formula in cell after Table[[#This Row],[Column1]:[Column2]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 45);
		oParser = new AscCommonExcel.parserFormula(tableName + "[@[Column1]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Result of Table[@[Column1]]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 45);
		resCell.setValue("=" + tableName +"[@[Column1]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[@Column1]", "Value for edit in cell after Table[@[Column1]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#This Row],[Column1]]", "Formula in cell after Table[@[Column1]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 50);
		oParser = new AscCommonExcel.parserFormula(tableName + "[#Headers]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[#Headers][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[#Headers][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), "Column3", 'Result of Table[#Headers][0,2]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 50);
		resCell.setValue("=" + tableName +"[#Headers]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#Headers]", "Value for edit in cell after Table[#Headers] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#Headers]", "Formula in cell after Table[#Headers] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 55);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Headers]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[[#Headers]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[[#Headers]][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), "Column3", 'Result of Table[[#Headers]][0,2]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 55);
		resCell.setValue("=" + tableName +"[[#Headers]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#Headers]", "Value for edit in cell after Table[[#Headers]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#Headers]", "Formula in cell after Table[[#Headers]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 60);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Headers],[Column2]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().getValue(), "Column2", 'Result of Table[[#Headers],[Column2]]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 60);
		resCell.setValue("=" + tableName +"[[#Headers],[Column2]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[[#Headers],[Column2]]", "Value for edit in cell after Table[[#Headers],[Column2]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#Headers],[Column2]]", "Formula in cell after Table[[#Headers],[Column2]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 65);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Headers],[Column2]:[Column3]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column2", 'Result of Table[[#Headers],[Column2]:[Column3]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column3", 'Result of Table[[#Headers],[Column2]:[Column3]][0,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 65);
		resCell.setValue("=" + tableName +"[[#Headers],[Column2]:[Column3]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[[#Headers],[Column2]:[Column3]]", "Value for edit in cell after Table[[#Headers],[Column2]:[Column3]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#Headers],[Column2]:[Column3]]", "Formula in cell after Table[[#Headers],[Column2]:[Column3]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 70);
		oParser = new AscCommonExcel.parserFormula(tableName + "[#All]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[#All][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[#All][0,1]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[#All][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[#All][1,1]');
		assert.strictEqual(array.getValueByRowCol(3, 0).getValue(), 1, 'Result of Table[#All][3,0]');
		assert.strictEqual(array.getValueByRowCol(3, 1).getValue(), 1, 'Result of Table[#All][3,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 70);
		resCell.setValue("=" + tableName +"[#All]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#All]", "Value for edit in cell after Table[#All] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#All]", "Formula in cell after Table[#All] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 75);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#All]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[[#All]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[[#All]][0,1]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[[#All]][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[[#All]][1,1]');
		assert.strictEqual(array.getValueByRowCol(3, 0).getValue(), 1, 'Result of Table[[#All]][3,0]');
		assert.strictEqual(array.getValueByRowCol(3, 1).getValue(), 1, 'Result of Table[[#All]][3,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 75);
		resCell.setValue("=" + tableName +"[[#All]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#All]", "Value for edit in cell after Table[[#All]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#All]", "Formula in cell after Table[[#All]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 80);
		oParser = new AscCommonExcel.parserFormula(tableName + "[#Data]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[#Data][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[#Data][0,1]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[#Data][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[#Data][1,1]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of Table[#Data][2,0]');
		assert.strictEqual(array.getValueByRowCol(2, 1).getValue(), 1, 'Result of Table[#Data][2,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 80);
		resCell.setValue("=" + tableName +"[#Data]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName +"[#Data]", "Value for edit in cell after Table[#Data] is typed");
		assert.strictEqual(resCell.getFormula(), tableName +"[#Data]", "Formula in cell after Table[#Data] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 85);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Data]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[[#Data]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[[#Data]][0,1]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[[#Data]][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[[#Data]][1,1]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of Table[[#Data]][2,0]');
		assert.strictEqual(array.getValueByRowCol(2, 1).getValue(), 1, 'Result of Table[[#Data]][2,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 85);
		resCell.setValue("=" + tableName +"[[#Data]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#Data]", "Value for edit in cell after Table[[#Data]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#Data]", "Formula in cell after Table[[#Data]] is typed");

		
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 90);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Totals]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Summary", 'Result of Table[[#Totals]][0,0]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 90);
		resCell.setValue("=" + tableName +"[[#Totals]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#Totals]", "Value for edit in cell after Table[#Totals]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#Totals]", "Formula in cell after Table[[#Totals]] is typed");


		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 95);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Data],[#Totals]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of Table[[#Data],[#Totals]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), 1, 'Result of Table[[#Data],[#Totals]][0,1]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[[#Data],[#Totals]][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[[#Data],[#Totals]][1,1]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of Table[[#Data],[#Totals]][2,0]');
		assert.strictEqual(array.getValueByRowCol(2, 1).getValue(), 1, 'Result of Table[[#Data],[#Totals]][2,1]');
		assert.strictEqual(array.getValueByRowCol(4, 0).getValue(), "Summary", 'Result of Table[[#Data],[#Totals]][4,0]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 95);
		resCell.setValue("=" + tableName +"[[#Data],[#Totals]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[[#Data],[#Totals]]", "Value for edit in cell after Table[[#Data],[#Totals]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#Data],[#Totals]]", "Formula in cell after Table[[#Data],[#Totals]] is typed");


		// =Table3[[#This Row],[#Data]]
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 100);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row],[#Data]]", cellWithFormula, ws);
		// assert.ok(oParser.parse()); // false
		array = oParser.calculate();
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'Result of Table[[#This Row],[#Data]]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 100);
		resCell.setValue("=" + tableName +"[[#This Row],[#Data]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "", "Value for edit in cell after Table[[#This Row],[#Data]] is typed");
		assert.strictEqual(resCell.getFormula(), "", "Formula in cell after Table[[#This Row],[#Data]] is typed");


		// =Table3[[#This Row],[#All]]
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 105);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#This Row],[#All]]", cellWithFormula, ws);
		// assert.ok(oParser.parse()); // false
		array = oParser.calculate();
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'Result of Table[[#This Row],[#All]]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 105);
		resCell.setValue("=" + tableName +"[[#This Row],[#All]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "", "Value for edit in cell after Table[[#This Row],[#All]] is typed");
		assert.strictEqual(resCell.getFormula(), "", "Formula in cell after Table[[#This Row],[#All]] is typed");


		// =Table3[[#Headers],[#Data]]
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 90);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Headers],[#Data]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[[#Headers],[#Data]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[[#Headers],[#Data]][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), "Column3", 'Result of Table[[#Headers],[#Data]][0,2]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[[#Headers],[#Data]][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[[#Headers],[#Data]][1,1]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of Table[[#Headers],[#Data]][2,0]');
		assert.strictEqual(array.getValueByRowCol(2, 1).getValue(), 1, 'Result of Table[[#Headers],[#Data]][2,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 90);
		resCell.setValue("=" + tableName +"[[#Headers],[#Data]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[[#Headers],[#Data]]", "Value for edit in cell after Table[[#Headers],[#Data]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#Headers],[#Data]]", "Formula in cell after Table[[#Headers],[#Data]] is typed");
		
		// todo
		// This entry is swapped in ms(Data is swapped with Headers) and is written as: =Table[[#Headers],[#Data]] and exactly the same is written to the file
		// Our editor does not have such functionality, but it does not affect the final result of the calculation
		// The only difference is that the string in the cell will differ from ms
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 95);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[#Data],[#Headers]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "Column1", 'Result of Table[[#Data],[#Headers]][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "Column2", 'Result of Table[[#Data],[#Headers]][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), "Column3", 'Result of Table[[#Data],[#Headers]][0,2]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[[#Data],[#Headers]][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[[#Data],[#Headers]][1,1]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of Table[[#Data],[#Headers]][2,0]');
		assert.strictEqual(array.getValueByRowCol(2, 1).getValue(), 1, 'Result of Table[[#Data],[#Headers]][2,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 95);
		resCell.setValue("=" + tableName +"[[#Data],[#Headers]]");
		
		//"=" + tableName + "[[#Headers],[#Data]]"
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[[#Data],[#Headers]]", "Value for edit in cell after Table[[#Data],[#Headers]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[[#Data],[#Headers]]", "Formula in cell after Table[[#Data],[#Headers]] is typed");

		// for bug 46174
		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 104, 0);
		oParser = new AscCommonExcel.parserFormula("[[Column1]]", cellWithFormula, ws);
		assert.ok(oParser.parse(true));
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Short notation. Result of [[Column1]][0,0] inside table');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Short notation. Result of [[Column1]][1,0] inside table');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Short notation. Result of [[Column1]][2,0] inside table');
		assert.strictEqual(array.getValueByRowCol(3, 0).getValue(), 1, 'Short notation. Result of [[Column1]][3,0] inside table');

		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 105, 0);
		oParser = new AscCommonExcel.parserFormula("[[Column1]]", cellWithFormula, ws);
		assert.ok(!oParser.parse(true));
		array = oParser.calculate();
		assert.strictEqual(array.type, AscCommonExcel.cElementType.error, 'Short notation. Result of [[Column1]] outside table');
		
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 20);
		oParser = new AscCommonExcel.parserFormula(tableName + "[[Column1]]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), 1, 'Result of [[Column1]][0,0]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of [[Column1]][1,0]');
		assert.strictEqual(array.getValueByRowCol(2, 0).getValue(), 1, 'Result of [[Column1]][2,0]');
		assert.strictEqual(array.getValueByRowCol(3, 0).getValue(), 1, 'Result of [[Column1]][3,0]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 20);
		resCell.setValue("=" + tableName +"[[Column1]]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[Column1]", "Value for edit in cell after Table[[Column1]] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[Column1]", "Formula in cell after Table[[Column1]] is typed");


		clearData(0, 99, 0, 105);
	});

	QUnit.test('Table special characters tests', function (assert) {
		/* This test checks for special characters inside a table, namely column names with escaped characters */
		let array;
		ws.getRange2("A100:C103").setValue("1");
		ws.getRange2("A100").setValue("With a single ' quote");
		ws.getRange2("B100").setValue("With a double '' quote");
		ws.getRange2("C100").setValue("With a special ' @ & ? * / | \ # '' [ ] characters");

		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		tableOptions.range = "A100:C103";
		tableOptions.isTitle = true;
		api.asc_addAutoFilter("TableStyleMedium2", tableOptions);	// create table in A100:C103 range

		let tables = wsView.model.autoFilters.getTablesIntersectionRange(new Asc.Range(0, 100, 0, 100));
		assert.strictEqual(tables.length, 1, "compare tables length");

		let table = tables[0];
		let tableName = table.DisplayName;	// due to the fact that other tables are used in file, get the name of the one we need by this way
		wsView.af_changeFormatTableInfo(tableName, Asc.c_oAscChangeTableStyleInfo.rowTotal, true);

		// calc res check
		cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 101, 70);
		oParser = new AscCommonExcel.parserFormula(tableName + "[#All]", cellWithFormula, ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValueByRowCol(0, 0).getValue(), "With a single ' quote", 'Result of Table[#All][0,0]');
		assert.strictEqual(array.getValueByRowCol(0, 1).getValue(), "With a double '' quote", 'Result of Table[#All][0,1]');
		assert.strictEqual(array.getValueByRowCol(0, 2).getValue(), "With a special ' @ & ? * / |  # '' [ ] characters", 'Result of Table[#All][0,2]');
		assert.strictEqual(array.getValueByRowCol(1, 0).getValue(), 1, 'Result of Table[#All][1,0]');
		assert.strictEqual(array.getValueByRowCol(1, 1).getValue(), 1, 'Result of Table[#All][1,1]');
		assert.strictEqual(array.getValueByRowCol(3, 0).getValue(), 1, 'Result of Table[#All][3,0]');
		assert.strictEqual(array.getValueByRowCol(3, 1).getValue(), 1, 'Result of Table[#All][3,1]');

		// value for edit and formula in cell check
		resCell = ws.getRange4(101, 70);
		resCell.setValue("=" + tableName +"[#All]");
		
		assert.strictEqual(resCell.getValueForEdit(), "=" + tableName + "[#All]", "Value for edit in cell after Table[#All] is typed");
		assert.strictEqual(resCell.getFormula(), tableName + "[#All]", "Formula in cell after Table[#All] is typed");


		/* column header check */
		// this pattern is looking for special characters in the string that need to be escaped
		let specialSymbolsPattern = /(['#@\[\]])/g;
		// set selection to the first column header
		let selectedRange = ws.getRange2("A100");
		wsView.setSelection(selectedRange.bbox);

		let ranges = wsView.model.selectionRange.ranges;
		let valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		let expectedValue = tableName + "[[#Headers],[" + ws.getRange2("A100").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Header value in first column. Simulation value for edit in cellEditMode");


		// set selection to the second column header
		selectedRange = ws.getRange2("B100");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[[#Headers],[" + ws.getRange2("B100").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Header value in second column. Simulation value for edit in cellEditMode");


		// set selection to the third column header
		selectedRange = ws.getRange2("C100");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[[#Headers],[" + ws.getRange2("C100").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Header value in third column. Simulation value for edit in cellEditMode");

		/* full column check */
		// first column select
		selectedRange = ws.getRange2("A101:A103");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[" + ws.getRange2("A100:A103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "First column without header select. Simulation value for edit in cellEditMode");

		// second column select
		selectedRange = ws.getRange2("B101:B103");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[" + ws.getRange2("B100:B103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Second column without header select. Simulation value for edit in cellEditMode");

		// third column select
		selectedRange = ws.getRange2("C101:C103");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[" + ws.getRange2("C100:C103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Third column without header select. Simulation value for edit in cellEditMode");

		/* two columns check */
		// first and second column
		selectedRange = ws.getRange2("A101:B103");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[" 
						+ "[" + ws.getRange2("A100:A103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]:" 
						+ "[" + ws.getRange2("B100:B103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]"
						+ "]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "First and second column without header select. Simulation value for edit in cellEditMode");

		// second and third column
		selectedRange = ws.getRange2("B101:C103");
		wsView.setSelection(selectedRange.bbox);

		ranges = wsView.model.selectionRange.ranges;
		valueInCellEditMode = tables[0].getSelectionString(wsView.model.getSelection().activeCell, ranges[0]);
		expectedValue = tableName + "[" 
						+ "[" + ws.getRange2("B100:B103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]:" 
						+ "[" + ws.getRange2("C100:C103").getValueForEdit().replace(specialSymbolsPattern, "'$1") + "]"
						+ "]";

		assert.strictEqual(valueInCellEditMode, expectedValue, "Second and third column without header select. Simulation value for edit in cellEditMode");


		clearData(0, 99, 0, 105);
	});

	/* for bug 61856 */
	QUnit.test('Array of arguments check after calling the wizard for the function', function (assert) { 
		let resArray;

		ws.getRange2("D1").setValue("11");
		ws.getRange2("D2").setValue("22");
		ws.getRange2("D3").setValue("33");

		api.wb.cellEditor =  {
			data : {},
			changeCellText: function () {},
			getText: function () {
				return "";
			}
		};

		api.wb.wsActive = ws.getIndex();
		api.wb.setCellEditMode(true);

		wb.dependencyFormulas.addDefName("def", "Sheet1!$A$100:$A$102");

		// api.asc_startWizard();
		let formula = "SUM(D1"
		let parser = new AscCommonExcel.parserFormula(formula, /*formulaParsed.parent*/null, ws);
		let _parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1", 'First argument in "SUM(D1"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlParenthesesCorrectCount, 'Error in parseResult for "SUM(D1" formula');


		formula = "SUM(D1,"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1", 'First argument in "SUM(D1,"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlOperandExpected, 'Error in parseResult for "SUM(D1," formula');


		formula = "SUM(D1,D2"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1", 'First argument in "SUM(D1,D2"');
		assert.strictEqual(resArray[1], "D2", 'Second argument in "SUM(D1,D2"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlParenthesesCorrectCount, 'Error in parseResult for "SUM(D1,D2" formula');


		formula = "SUM(D1,D2)"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1", 'First argument in "SUM(D1,D2)"');
		assert.strictEqual(resArray[1], "D2", 'Second argument in "SUM(D1,D2)"');
		assert.strictEqual(_parseResult.error, undefined, 'Error in parseResult for "SUM(D1,D2)" formula');

		
		formula = "SUM(D1:D2)"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1:D2", 'First argument in "SUM(D1:D2)"');
		assert.strictEqual(_parseResult.error, undefined, 'Error in parseResult for "SUM(D1:D2)" formula');


		formula = "SUM(D1:D2,)"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "D1:D2", 'First argument in "SUM(D1:D2,)"');
		assert.strictEqual(_parseResult.error, undefined, 'Error in parseResult for "SUM(D1:D2,)" formula');


		formula = "SUM("
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "", 'First argument in "SUM("');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlOperandExpected, 'Error in parseResult for "SUM(" formula');


		formula = "SUM(1"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "1", 'First argument in "SUM(1"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlParenthesesCorrectCount, 'Error in parseResult for "SUM(1" formula');

		formula = "SUM(1,D1"
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "1", 'First argument in "SUM(1,D1"');
		assert.strictEqual(resArray[1], "D1", 'Second argument in "SUM(1,D1"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlParenthesesCorrectCount, 'Error in parseResult for "SUM(1,D1" formula');


		formula = 'SUM(1,D1,"str",TRUE,def'
		parser = new AscCommonExcel.parserFormula(formula, null, ws);
		_parseResult = new AscCommonExcel.ParseResult([], []);

		parser.parse(true, true, _parseResult, true);

		resArray = _parseResult.getArgumentsValue(formula);
		assert.strictEqual(resArray[0], "1", 'First argument in "SUM(1,D1,"str",TRUE,def"');
		assert.strictEqual(resArray[1], "D1", 'Second argument in "SUM(1,D1,"str",TRUE,def"');
		assert.strictEqual(resArray[2], "\"str\"", 'Third argument in "SUM(1,D1,"str",TRUE,def"');
		assert.strictEqual(resArray[3], "TRUE", 'Fourth argument in "SUM(1,D1,"str",TRUE,def"');
		assert.strictEqual(resArray[4], "def", 'Fifth argument in "SUM(1,D1,"str",TRUE,def"');
		assert.strictEqual(_parseResult.error, c_oAscError.ID.FrmlParenthesesCorrectCount, 'Error in parseResult for "SUM(1,D1,"str",TRUE,def" formula');

		// remove all created earlier defNames
		wb.dependencyFormulas._foreachDefName(function(defName) {
			wb.dependencyFormulas.removeDefName(undefined, defName.name);
		});

		// for bug 61855
		let insertArgsRes = api.wb.insertArgumentsInFormula(["1"], 0, 0, "SUM", true/*bEndInsertArg*/);
		assert.strictEqual(insertArgsRes && insertArgsRes.functionResult, "1", "Calculation result for SUM function when insert first argument in formula");

		insertArgsRes = api.wb.insertArgumentsInFormula(["1,2"], 1, 0, "SUM", true);
		assert.strictEqual(insertArgsRes && insertArgsRes.functionResult, "3", "Calculation result for SUM function when insert second argument in formula");

		insertArgsRes = api.wb.insertArgumentsInFormula(["1"], 1, 0, "SUM", true);
		assert.strictEqual(insertArgsRes && insertArgsRes.functionResult, "1", "Calculation result for SUM function when delete second argument in formula");
		
	});

	QUnit.test('autoCompleteFormula', function (assert) {
		let resCell, range, fillRange, autoCompleteRes;

		ws.getRange2("A1:Z100").cleanAll();

		let testData = [
			// ['1', 'Test', 'Test1', '01/01/2000']
			['1'],
			['Test'],
			['Test1'],
			['01/01/2000']
		];

		// Asc cases
		range = ws.getRange2("A1");
		range.fillData(testData);
		// ws.getRange2("A1").setValue("1");
		// ws.getRange2("A2").setValue("Test");
		// ws.getRange2("A3").setValue("Test1");
		// ws.getRange2("A4").setValue("01/01/2000");

		// c1, r1, c2, r2
		// fillRange = new Asc.Range(0, 0, 0, 3);
		fillRange = ws.getRange2("A1:A4");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");
		
		resCell = ws.getRange2("A5");
		assert.strictEqual(resCell.getValueWithFormat(), "36527", "Value after A1:A4 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A1:A4)", "Formula after A1:A4 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A1:A5", "Selection after A1:A4 autosum");

		range = ws.getRange2("B1");
		ws.getRange2("B1").setValue("ds");
		ws.getRange2("B2").setValue("1");

		// fillRange = new Asc.Range(1, 0, 1, 1);
		fillRange = ws.getRange2("B1:B2");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("B3");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value after B1:B2 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula after B1:B2 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B1:B2", "Selection after B1:B2 autosum");


		ws.getRange2("C1").setValue("ds");
		ws.getRange2("C2").setValue("1");
		ws.getRange2("C3").setValue("1");

		// fillRange = new Asc.Range(2, 0, 2, 2);
		fillRange = ws.getRange2("C1:C3");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("C4");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after C1:C3 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(C2:C3)", "Formula after C1:C3 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "C2:C4", "Selection after C1:C3 autosum");


		ws.getRange2("D2").setValue("ds");
		ws.getRange2("D4").setValue("1");

		// fillRange = new Asc.Range(3, 0, 3, 3);
		fillRange = ws.getRange2("D1:D4");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("D5");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after D1:D4 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(D1:D4)", "Formula after D1:D4 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "D1:D5", "Selection after D1:D4 autosum");


		ws.getRange2("E1").setValue("ds");
		ws.getRange2("E3").setValue("1");
		ws.getRange2("E4").setValue("");

		// fillRange = new Asc.Range(4, 0, 4, 3);
		fillRange = ws.getRange2("E1:E4");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");


		resCell = ws.getRange2("E4");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after E1:E4 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(E3)", "Formula after E1:E4 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "E3:E4", "Selection after E1:E4 autosum");


		ws.getRange2("F1").setValue("ds");
		ws.getRange2("F3").setValue("1");
		ws.getRange2("F4").setValue("1");

		// fillRange = new Asc.Range(5, 0, 5, 3);
		fillRange = ws.getRange2("F1:F4");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("F5");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after F1:F4 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(F3:F4)", "Formula after F1:F4 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "F3:F5", "Selection after F1:F4 autosum");


		ws.getRange2("G2").setValue("ds");
		ws.getRange2("G4").setValue("1");
		ws.getRange2("G5").setValue("1");

		// fillRange = new Asc.Range(6, 0, 6, 4);
		fillRange = ws.getRange2("G1:G5");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("G6");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after G1:G5 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(G1:G5)", "Formula after G1:G5 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "G1:G6", "Selection after G1:G5 autosum");

		// col tests
		ws.getRange2("A10").setValue("ds");
		ws.getRange2("B10").setValue("1");
		ws.getRange2("C10").setValue("1");

		// fillRange = new Asc.Range(0, 9, 2, 9);
		fillRange = ws.getRange2("A10:C10");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("D10");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A10:C10 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B10:C10)", "Formula after A10:C10 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B10:D10", "Selection after A10:C10 autosum");


		ws.getRange2("A11").setValue("ds");
		ws.getRange2("B11").setValue("1");

		// fillRange = new Asc.Range(0, 10, 2, 10);
		fillRange = ws.getRange2("A11:C11");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("C11");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after A11:C11 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B11)", "Formula after A11:C11 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B11:C11", "Selection after A11:C11 autosum");
		

		ws.getRange2("B12").setValue("ds");
		ws.getRange2("D12").setValue("1");
		ws.getRange2("E12").setValue("1");

		// fillRange = new Asc.Range(0, 11, 4, 11);
		fillRange = ws.getRange2("A12:E12");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("F12");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A12:E12 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A12:E12)", "Formula after A12:E12 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A12:F12", "Selection after A12:E12 autosum");

		// row + col tests
		ws.getRange2("A20").setValue("ds");
		ws.getRange2("A21").setValue("1");
		ws.getRange2("A22").setValue("1");
		ws.getRange2("B20").setValue("ds");
		ws.getRange2("B21").setValue("1");
		ws.getRange2("B23").setValue("1");

		// fillRange = new Asc.Range(0, 19, 1, 22);
		fillRange = ws.getRange2("A20:B23");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A24");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A20:B23 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A21:A23)", "Formula after A20:B23 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A21:B24", "Selection after A20:B23 autosum");
		resCell = ws.getRange2("B24");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A20:B23 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B21:B23)", "Formula after A20:B23 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A21:B24", "Selection after A20:B23 autosum");


		ws.getRange2("K1").setValue("ds");
		ws.getRange2("L1").setValue("ds");
		ws.getRange2("K3").setValue("1");

		// fillRange = new Asc.Range(10, 0, 11, 2);
		fillRange = ws.getRange2("K1:L3");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("L3");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after K1:L3 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(K3)", "Formula after K1:L3 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "K3:L3", "Selection after K1:L3 autosum");


		ws.getRange2("K5:L10").cleanAll();
		ws.getRange2("K6").setValue("ds");
		ws.getRange2("K8").setValue("1");

		// fillRange = new Asc.Range(10, 4, 11, 8);
		fillRange = ws.getRange2("K5:L9");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("L8");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after K5:L9 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(K8)", "Formula after K5:L9 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "K8:L9", "Selection after K5:L9 autosum");
		resCell = ws.getRange2("K9");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after K5:L9 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(K8)", "Formula after K5:L9 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "K8:L9", "Selection after K5:L9 autosum");
		resCell = ws.getRange2("L9");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after K5:L9 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(K9)", "Formula after K5:L9 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "K8:L9", "Selection after K5:L9 autosum");


		ws.getRange2("A12:M14").cleanAll();
		ws.getRange2("K13").setValue("1");
		ws.getRange2("L13").setValue("1");

		// fillRange = new Asc.Range(9, 11, 12, 13);
		fillRange = ws.getRange2("A12:M14");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("M13");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A12:M14 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A13:L13)", "Formula after A12:M14 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A13:M14", "Selection after A12:M14 autosum");
		resCell = ws.getRange2("M14");
		assert.strictEqual(resCell.getValueWithFormat(), "2", "Value after A12:M14 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A14:L14)", "Formula after A12:M14 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A13:M14", "Selection after A12:M14 autosum");
		resCell = ws.getRange2("K14");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after A12:M14 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(K13)", "Formula after A12:M14 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A13:M14", "Selection after A12:M14 autosum");
		resCell = ws.getRange2("L14");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after A12:M14 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(L13)", "Formula after A12:M14 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A13:M14", "Selection after A12:M14 autosum");
		
		// for bug 32959
		ws.getRange2("B10:B14").cleanAll();
		ws.getRange2("B11").setValue("1");
		ws.getRange2("B12").setValue("ds");

		// B11:B12
		fillRange = ws.getRange2("B11:B12");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("B13");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B11:B12 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B11:B12)", "Formula after B11:B12 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B11:B13", "Selection after B11:B12 autosum");

		// B11:B13
		ws.getRange2("B13:B14").cleanAll();
		fillRange = ws.getRange2("B11:B13");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("B13");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B11:B13 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B11:B12)", "Formula after B11:B13 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B11:B13", "Selection after B11:B13 autosum");

		// B10:B12
		ws.getRange2("B13:B14").cleanAll();
		fillRange = ws.getRange2("B10:B12");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("B13");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B10:B12 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B10:B12)", "Formula after B10:B12 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B10:B13", "Selection after B10:B12 autosum");

		// B10:B13
		ws.getRange2("B13:B14").cleanAll();
		fillRange = ws.getRange2("B10:B13");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("B13");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B10:B13 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B10:B12)", "Formula after B10:B13 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B10:B13", "Selection after B10:B13 autosum");

		
		ws.getRange2("B20:F20").cleanAll();
		ws.getRange2("C20").setValue("1");
		ws.getRange2("D20").setValue("ds");
		// C20:D20
		ws.getRange2("E20:F20").cleanAll();
		fillRange = ws.getRange2("C20:D20");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("E20");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after C20:D20 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(C20:D20)", "Formula after C20:D20 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "C20:E20", "Selection after C20:D20 autosum");

		// C20:E20
		ws.getRange2("E20:F20").cleanAll();
		fillRange = ws.getRange2("C20:E20");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("E20");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after C20:E20 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(C20:D20)", "Formula after C20:E20 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "C20:E20", "Selection after C20:E20 autosum");

		// B20:D20
		ws.getRange2("E20:F20").cleanAll();
		fillRange = ws.getRange2("B20:D20");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("E20");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B20:D20 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B20:D20)", "Formula after B20:D20 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B20:E20", "Selection after B20:D20 autosum");

		// B20:E20
		ws.getRange2("E20:F20").cleanAll();
		fillRange = ws.getRange2("B20:E20");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("E20");
		assert.strictEqual(resCell.getValueWithFormat(), "1", "Value after B20:E20 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B20:D20)", "Formula after B20:E20 autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "B20:E20", "Selection after B20:E20 autosum");

		// for bug 37318
		ws.getRange2("A20:A22").cleanAll();
		ws.getRange2("A20:A22").setValue("1");
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("m/d/yyyy");		// change to the short date format
		// wsView.setSelectionInfo("format", "m/d/yyyy");

		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value after A20:A22(only dates in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula after A20:A22(only dates in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A22", "Selection after A20:A22(only dates in range) autosum");


		// number 
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("0.00"); 	// change to the number format

		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "3", "Value after A20:A22(only number in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A20:A22)", "Formula after A20:A22(only number in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A23", "Selection after A20:A22(only number in range) autosum");


		// fraction
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("# ?/?"); 	// change to the fraction format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "3", "Value after A20:A22(only fraction in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A20:A22)", "Formula after A20:A22(only fraction in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A23", "Selection after A20:A22(only fraction in range) autosum");


		// scientific
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("0.00E+00"); 	// change to the scientific format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "3", "Value after A20:A22(only scientific in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A20:A22)", "Formula after A20:A22(only scientific in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A23", "Selection after A20:A22(only scientific in range) autosum");


		// accounting
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("_([$$-409]* #,##0.00_);_([$$-409]* \\(#,##0.00\\);_([$$-409]* \"-\"??_);_(@_)"); 	// change to the accounting format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "3", "Value after A20:A22(only accounting in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A20:A22)", "Formula after A20:A22(only accounting in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A23", "Selection after A20:A22(only accounting in range) autosum");


		// percentage
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("0.00%"); 	// change to the percentage format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "3", "Value after A20:A22(only percents in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A20:A22)", "Formula after A20:A22(only percents in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A23", "Selection after A20:A22(only percents in range) autosum");


		// l.date
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("[$-F800]dddd\,\ mmmm\ d\,\ yyyy"); 	// change to the long date format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value after A20:A22(only dates in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula after A20:A22(only dates in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A22", "Selection after A20:A22(only dates in range) autosum");


		// text
		fillRange = ws.getRange2("A20:A22");
		fillRange.setNumFormat("@"); 	// change to the text format
		ws.getRange2("A23").cleanAll();
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A23");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value after A20:A22(only text in range) autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula after A20:A22(only text in range) autosum");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A20:A22", "Selection after A20:A22(only text in range) autosum");

		/* activeCell tests */
		/* if the data type does not allow the formula to be executed, the active cell is moved to the end of the select */
		let activeCell, supposedActiveCell;

		ws.getRange2("A1:Z100").cleanAll();
		ws.getRange2("B10").setValue("111");
		ws.getRange2("B20:B22").setValue("ds");
		fillRange = ws.getRange2("B20:B22");
		fillRange.setNumFormat("@");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		activeCell = ws.selectionRange.activeCell;
		supposedActiveCell = ws.getCell2("B22");
		assert.strictEqual(activeCell.col === supposedActiveCell.bbox.c1 && activeCell.row === supposedActiveCell.bbox.r1, true, "Active cell test. B20:B22(only text in range) autosum");


		ws.getRange2("B10").cleanAll();
		ws.getRange2("A20:A21").setValue("111");
		fillRange = ws.getRange2("B20:B22");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		activeCell = ws.selectionRange.activeCell;
		supposedActiveCell = ws.getCell2("B22");
		assert.strictEqual(activeCell.col === supposedActiveCell.bbox.c1 && activeCell.row === supposedActiveCell.bbox.r1, true, "Active cell test. B20:B22(only text in range) autosum");


		ws.getRange2("A22").setValue("111");
		fillRange = ws.getRange2("B20:B22");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		activeCell = ws.selectionRange.activeCell;
		supposedActiveCell = ws.getCell2("B22");
		assert.strictEqual(activeCell.col === supposedActiveCell.bbox.c1 && activeCell.row === supposedActiveCell.bbox.r1, true, "Active cell test. B20:B22(only text in range) autosum");

		/* autocomplete to merge cells */
		ws.getRange2("A30:B31").setValue("111");
		// C30:D31 merge & center
		ws.getRange2("C30:D31").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C30:D31").hasMerged(), 'Range C30:D31 is merged');

		fillRange = ws.getRange2("A30:D31");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("C30");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A30:D31 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A30:D31 autosum");
		resCell = ws.getRange2("C31");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A30:D31 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A30:D31 autosum");
		resCell = ws.getRange2("D30");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A30:D31 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A30:D31 autosum");
		resCell = ws.getRange2("D31");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A30:D31 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A30:D31 autosum");

		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A30:D31", "Selection after A30:D31 autosum");


		ws.getRange2("A40:B41").setValue("111");
		ws.getRange2("C40:C41").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C40:C41").hasMerged(), 'Range C40:C41 is merged');

		ws.getRange2("D40:D41").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("D40:C41").hasMerged(), 'Range D40:C41 is merged');

		fillRange = ws.getRange2("A40:D41");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("C40");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A40:D41 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A40:D41 autosum");
		resCell = ws.getRange2("C41");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A40:D41 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A40:D41 autosum");
		resCell = ws.getRange2("D40");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A40:D41 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A40:D41 autosum");
		resCell = ws.getRange2("D41");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A40:D41 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A40:D41 autosum");

		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A40:D41", "Selection after A40:D41 autosum");


		ws.getRange2("A50:A51").setValue("111");
		ws.getRange2("A52:A53").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A52:A53").hasMerged(), 'Range A52:A53 is merged');

		fillRange = ws.getRange2("A50:A53");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A52");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A50:A53 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A50:A53 autosum");
		resCell = ws.getRange2("A52");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A50:A53 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A50:A53 autosum");

		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A50:A53", "Selection after A50:A53 autosum");


		ws.getRange2("A60:B61").setValue("111");
		ws.getRange2("C60:D60").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C60:D60").hasMerged(), 'Range C60:D60 is merged');

		ws.getRange2("C62:D62").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C62:D62").hasMerged(), 'Range C62:D62 is merged');

		fillRange = ws.getRange2("A60:D62");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		resCell = ws.getRange2("A62");
		assert.strictEqual(resCell.getValueWithFormat(), "222", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A60:A61)", "Formula in merged cell after A60:D62 autosum");
		resCell = ws.getRange2("B62");
		assert.strictEqual(resCell.getValueWithFormat(), "222", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(B60:B61)", "Formula in merged cell after A60:D62 autosum");
		resCell = ws.getRange2("C62");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A60:D62 autosum");
		resCell = ws.getRange2("D62");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A60:D62 autosum");
		resCell = ws.getRange2("D60");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A60:D62 autosum");
		resCell = ws.getRange2("D61");
		assert.strictEqual(resCell.getValueWithFormat(), "222", "Value in merged cell after A60:D62 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "=SUM(A61:C61)", "Formula in merged cell after A60:D62 autosum");

		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A60:D62", "Selection after A60:D62 autosum");

		ws.getRange2("A1:D70").cleanAll();
		// autosum from merged cell
		ws.getRange2("A70:B70").setValue("111");
		ws.getRange2("C70:D70").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C70:D70").hasMerged(), 'Range C70:D70 is merged');

		fillRange = ws.getRange2("D70");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		assert.strictEqual(autoCompleteRes && autoCompleteRes.text, "A70:B70", "Formula after autosum from merged cell. C70:D70 autosum from right");
		resCell = ws.getRange2("C70");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after C70:D70 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after C70:D70 autosum");


		ws.getRange2("A80:B82").setValue("111");
		ws.getRange2("C80:H88").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("C80:H88").hasMerged(), 'Range C80:H88" is merged');

		fillRange = ws.getRange2("D82");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		assert.strictEqual(autoCompleteRes && autoCompleteRes.text, "A80:B80", "Formula after autosum from merged cell. C80:H88 autosum from right");
		resCell = ws.getRange2("C80");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after C80:H88 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after C80:H88 autosum");


		ws.getRange2("A90:B91").setValue("111");
		ws.getRange2("A92:D94").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A92:D94").hasMerged(), 'Range A92:D94 is merged');

		fillRange = ws.getRange2("D94");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		assert.strictEqual(autoCompleteRes && autoCompleteRes.text, "A90:A91", "Formula after autosum from merged cell. A92:D94 autosum from bottom");
		resCell = ws.getRange2("A92");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A92:D94 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A92:D94 autosum");


		ws.getRange2("A110:C111").setValue("111");
		ws.getRange2("C112").setValue("111");
		ws.getRange2("A113:C115").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A113:C115").hasMerged(), 'Range A113:C115 is merged');

		fillRange = ws.getRange2("C113");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		autoCompleteRes = wsView.autoCompleteFormula("SUM");

		assert.strictEqual(autoCompleteRes && autoCompleteRes.text, "A110:A112", "Formula after autosum from merged cell. A113:C115 autosum from bottom");
		resCell = ws.getRange2("A92");
		assert.strictEqual(resCell.getValueWithFormat(), "", "Value in merged cell after A113:C115 autosum");
		assert.strictEqual(resCell.getValueForEdit(), "", "Formula in merged cell after A113:C115 autosum");
		
		ws.getRange2("A1:Z100").cleanAll();
	});

	/* for bug 69708 */
	QUnit.test('Array formula ', function (assert) {
		let array, cellWithFormula, fillRange;

		ws.getRange2("A11").setNumFormat("@");
		ws.getRange2("B12").setNumFormat("@");

		ws.getRange2("A2").setValue("Jeans");
		ws.getRange2("A3").setValue("Sweater");
		ws.getRange2("A4").setValue("Shirt");

		// set flags for CSE formula call
		let flags = wsView._getCellFlags(9, 0);
		flags.ctrlKey = true;
		flags.shiftKey = true;

		// set selection A10:B13
		fillRange = ws.getRange2("A10:B12");
		wsView.setSelection(fillRange.bbox);
		wsView._initRowsCount();
		wsView._initColsCount();
		

		let fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText("=SIN({1,2})");

		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);		// calculate
		resCell = ws.getRange2("A10");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A10:B12", "Selection after =SIN({1,2}) cse formula call");
		assert.strictEqual(resCell.getValueWithFormat(), "0.841470985", "Value in A10 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in A10 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "General", "A10 cellFormat after =SIN({1,2}) calculate");

		resCell = ws.getRange2("A11");
		assert.strictEqual(resCell.getValueWithFormat(), "0.841470985", "Value in A11 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in A11 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "@", "A11 cellFormat after =SIN({1,2}) calculate");

		resCell = ws.getRange2("A12");
		assert.strictEqual(resCell.getValueWithFormat(), "0.841470985", "Value in A12 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in A12 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "General", "A12 cellFormat after =SIN({1,2}) calculate");

		resCell = ws.getRange2("B10");
		assert.strictEqual(resCell.getValueWithFormat(), "0.909297427", "Value in B10 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in B10 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "General", "B10 cellFormat after =SIN({1,2}) calculate");

		resCell = ws.getRange2("B11");
		assert.strictEqual(resCell.getValueWithFormat(), "0.909297427", "Value in B11 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in B11 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "General", "B11 cellFormat after =SIN({1,2}) calculate");

		resCell = ws.getRange2("B12");
		assert.strictEqual(resCell.getValueWithFormat(), "0.909297427", "Value in B12 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getValueForEdit(), "=SIN({1,2})", "Formula in B12 after =SIN({1,2}) calculate");
		assert.strictEqual(resCell.getNumFormatStr(), "@", "B12 cellFormat after =SIN({1,2}) calculate");


		fragment[0].setFragmentText('=HSTACK({"Red";"Blue";"Green"},A2:A4)');

		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);		// calculate
		resCell = ws.getRange2("A10");
		assert.strictEqual(wsView.model.selectionRange.getLast().getName(), "A10:B12", 'Selection after =HSTACK({"Red";"Blue";"Green"},A2:A4) cse formula call');
		assert.strictEqual(resCell.getValueWithFormat(), "Red", 'Value in A10 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in A10 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "General", 'A10 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');

		resCell = ws.getRange2("A11");
		assert.strictEqual(resCell.getValueWithFormat(), "Blue", 'Value in A11 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in A11 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "@", 'A11 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');

		resCell = ws.getRange2("A12");
		assert.strictEqual(resCell.getValueWithFormat(), "Green", 'Value in A12 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in A12 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "General", 'A12 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');

		resCell = ws.getRange2("B10");
		assert.strictEqual(resCell.getValueWithFormat(), "Jeans", 'Value in B10 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in B10 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "General", 'B10 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');

		resCell = ws.getRange2("B11");
		assert.strictEqual(resCell.getValueWithFormat(), "Sweater", 'Value in B11 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in B11 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "General", 'B11 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');

		resCell = ws.getRange2("B12");
		assert.strictEqual(resCell.getValueWithFormat(), "Shirt", 'Value in B12 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getValueForEdit(), '=HSTACK({\"Red\";\"Blue\";\"Green\"},A2:A4)', 'Formula in B12 after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');
		assert.strictEqual(resCell.getNumFormatStr(), "@", 'B12 cellFormat after =HSTACK({"Red";"Blue";"Green"},A2:A4) calculate');


		ws.getRange2("A1:Z100").cleanAll();
	});

	QUnit.test('sortRangeTest', function (assert) {
		let range, expectedRes;

		ws.getRange2("A1:Z100").cleanAll();

		let testData = [
			['a'],
			['h'],
			['f'],
			['é'],
			['e'],
			['d'],
			['c'],
			['b'],
			['á'],
			['g']
		];

		range = ws.getRange2("A1:A10");
		range.fillData(testData);

		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, [['a'],['á'],['b'],['c'],['d'],['e'],['é'],['f'],['g'],['h']], "check_sort_1");

		range = ws.getRange2("A1:A10");
		range.fillData(testData);

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, [['h'],['g'],['f'],['é'],['e'],['d'],['c'],['b'],['á'],['a']], "check_sort_2");

		testData = [
			['1'],
			['g'],
			['2'],
			['é'],
			['TEST3'],
			['á'],
			['c'],
			['Test2'],
			['test1'],
			['a']
		];

		range = ws.getRange2("A1:A10");
		range.fillData(testData);

		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, [['1'],['2'],['a'],['á'],['c'],['é'],['g'],['test1'],['Test2'],['TEST3']], "check_sort_3");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, [['TEST3'],['Test2'],['test1'],['g'],['é'],['c'],['á'],['a'],['2'],['1']], "check_sort_4");

		testData = [
			['-2'],
			['Test2'],
			['test1'],
			['g'],
			['é'],
			['12345'],
			['á'],
			['a'],
			['안세요'],
			['녕하'],
			['TEST0'],
			['하'],
			['TEST2'],
			['аА'],
			['é'],
			['1'],
			['2'],
			['АА'],
			['аа'],
			['-1']
		];

		range = ws.getRange2("A1:A20");
		range.fillData(testData);

		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, [['-2'],['-1'],['1'],['2'],['12345'],['a'],['á'],['é'],['é'],['g'],['TEST0'],['test1'],['Test2'],['TEST2'],['аА'],['АА'],['аа'],['녕하'],['안세요'],['하']], "check_sort_5");

		/* MS puts Arabic alphabet above Japanese, javascript sorts differently*/

		testData = [
			['أ'],	// arabic
			['А'],	// russian
			['あ'],	// japanese
		];

		range = ws.getRange2("A1:A3");
		range.fillData(testData);

		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, [['А'],['أ'],['あ']], "Asc check_sort_6");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, [['あ'],['أ'],['А']], "Desc check_sort_6");

		testData = [
			['A'],	// english
			['A'],	// spanish
			['A'],	// german
			['Ä'],	// german
			['A'],	// french
			['À'],	// french
			['A'],	// hungarian
			['Á'],	// hungarian
			['A'],	// turkish
			['أ'],	// arabic
			['А'],	// russian
			['あ'],	// japanese
		];

		range = ws.getRange2("A1:A12");
		range.fillData(testData);

		expectedRes = [['A'],['A'],['A'],['A'],['A'],['A'],['Á'],['À'],['Ä'],['А'],['أ'],['あ']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_7");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_7");
		
		// Tests with more of languages
		testData = [
			['City'],	// eng
			['城市'],	// cn
			['शहर'],	// hindi
			['Ciudad'],	// esp
			['Ville'],	// fr
			['مدينة'],	// arabic
			['শহর'],	// bengal
			['Город'],	// ru
			['Cidade'],	// pt-pt
			['Kota'],	// indonesian
			['Város'],	// Magyar
			['Stadt']	// Dutch
		];

		range = ws.getRange2("A1:A12");
		range.fillData(testData);

		expectedRes = [['Cidade'],['City'],['Ciudad'],['Kota'],['Stadt'],['Város'],['Ville'],['Город'],['مدينة'],['शहर'],['শহর'],['城市']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_8");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_8");
		
		testData = [
			['123City'],	// eng
			['123城市'],	// cn
			['123शहर'],		// hindi
			['123Ciudad'],	// esp
			['123Ville'],	// fr
			['123' + 'مدينة'],	// arabic
			['123শহর'],		// bengal
			['123Город'],	// ru
			['123Cidade'],	// pt-pt
			['123Kota'],	// indonesian
			['123Város'],	// Magyar
			['123Stadt']	// Dutch
		];

		range = ws.getRange2("A1:A12");
		range.fillData(testData);

		expectedRes = [['123Cidade'],['123City'],['123Ciudad'],['123Kota'],['123Stadt'],['123Város'],['123Ville'],['123Город'],['123مدينة'],['123शहर'],['123শহর'],['123城市']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_9");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_9");

		testData = [
			['2e2de2'],
			['zxc'],
			['f'],
			['_d'],
			['edasd'],
			['SiM`Gl23'],
			['dd23dd'],	
			['3e2de2'],
			['SDY'],
			['3e2de3'],
			['2e3de2'],
			['__z']
		];

		range = ws.getRange2("A1:A12");
		range.fillData(testData);

		expectedRes = [['__z'],['_d'],['2e2de2'],['2e3de2'],['3e2de2'],['3e2de3'],['dd23dd'],['edasd'],['f'],['SDY'],['SiM`Gl23'],['zxc']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_10");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_10");

		// sorting specials characters does not coincide with the sorting in ms (but the same as in gs and libre)
		/*
		testData = [
			['!a'],
			['#a'],
			['$a'],
			['^a'],
			['&a'],
			['*a'],
			['(a'],	
			[')a'],
			['_a'],
			['[a'],
			[']a'],
			['{a'],
			['}a'],
			['\a'],
			['/a'],
			['|a'],
			[';a'],
			[':a'],
			['a'],	
			[',a'],
			['.a'],
			['`a'],
			['>a'],
			['<a'],
			['?a'],
			['~a'],
		];

		range = ws.getRange2("A1:A26");
		range.fillData(testData);

		expectedRes = [['!a'],['#a'],['$a'],['&a'],['(a'],[')a'],['*a'],[',a'],['.a'],['/a'],[':a'],[';a'],['?a'],['[a'],['\a'],[']a'],['^a'],['_a'],['`a'],['{a'],['|a'],['}a'],['~a'],['<a'],['>a'],['a']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_11");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_11");

		// letters with diacritics sorting also does not coincide with the sorting in ms (but the same as in gs and libre)
		testData = [
			['á'],
			['Á'],
			['à'],
			['ă'],
			['â'],
			['Â'],
			['å'],
			['ä'],
			['ã'],
			['Ą'],
			['ā'],
			['Ā'],
		];
		
		range = ws.getRange2("A1:A4");
		range.fillData(testData);

		expectedRes = [['á'],['Á'],['à'],['â'], ['Â'],['ä'],['ă'],['ā'],['Ā'],['ã'],['å'],['Ą'],];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_12");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_12");
		*/

		// english
		testData = [
			['Test1'],
			['TEST1'],
			['tESt1'],
			['TesT1'],
		];

		range = ws.getRange2("A1:A4");
		range.fillData(testData);

		expectedRes = [['Test1'],['TEST1'],['tESt1'],['TesT1']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_13");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes, "Desc check_sort_13");

		// hungarian
		testData = [
			['Köszönöm1'],
			['KÖSZÖNÖM1'],
			['köSzÖnÖm1'],
			['KöszönöM1'],
		];

		range = ws.getRange2("A1:A4");
		range.fillData(testData);

		expectedRes = [['Köszönöm1'],['KÖSZÖNÖM1'],['köSzÖnÖm1'],['KöszönöM1']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_14");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes, "Desc check_sort_14");

		testData = [
			['ZZ'],['á'], ['é'],['í'], ['ó'], 
			['ö'], ['ő'], ['ú'], ['ü'], ['ű'],['a']
		];

		range = ws.getRange2("A1:A11");
		range.fillData(testData);

		expectedRes = [
			['a'],['á'],['é'], ['í'], ['ó'],
			['ö'], ['ő'], ['ú'], ['ü'], ['ű'],['ZZ']
		];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_15");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_15");


		// turkish
		testData = [
			['C'],['Ç'],['ç'],['Ğ'],['ğ'],
			['Ö'],['ö'],['Ş'],['ş'],['Ü'],['ü']
		];

		range = ws.getRange2("A1:A11");
		range.fillData(testData);

		expectedRes = [
			['C'],['Ç'],['ç'],['Ğ'],['ğ'],
			['Ö'],['ö'],['Ş'],['ş'],['Ü'],['ü']
		];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_16");

		expectedRes = [
			['Ü'],['ü'],['Ş'],['ş'],['Ö'],
			['ö'],['Ğ'],['ğ'],['Ç'],['ç'],['C'],
		];
		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes, "Desc check_sort_16");


		testData = [
			['ALİ'],
			['MURAT'],
			['İSMAİL']
		];

		range = ws.getRange2("A1:A3");
		range.fillData(testData);

		expectedRes = [['ALİ'],['İSMAİL'],['MURAT']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_17");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_17");


		// hungarian, portugese, deutsch, turkish
		testData = [
			['Évad'],
			['Óculos'],
			['Äpfel'],
			['Şehir']
		];

		range = ws.getRange2("A1:A4");
		range.fillData(testData);

		expectedRes = [['Äpfel'],['Évad'],['Óculos'],['Şehir']];
		range.sort(Asc.c_oAscSortOptions.Ascending, 0);
		compareData(assert, range.bbox, expectedRes, "Asc check_sort_18");

		range.sort(Asc.c_oAscSortOptions.Descending, 0);
		compareData(assert, range.bbox, expectedRes.reverse(), "Desc check_sort_18");

	});
	QUnit.test("Autofill - format Date, Date & Time and Time.", function (assert) {
		function getAutofillCase(aFrom, aTo, nFillHandleArea, sDescription, expectedData) {
			const [c1From, c2From, r1From, r2From] = aFrom;
			const [c1To, c2To, r1To, r2To] = aTo;
			const nHandleDirection = r1To === r2To ? 0 : 1; // 0 - Horizontal, 1 - Vertical
			const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;
			let  autoFillRange;

			if (nHandleDirection === 0) {
				let autofillC1 = nFillHandleArea === 3 ? c2From + 1 : c1From - 1;
				autoFillRange = getRange(autofillC1, r1To, c2To, r2To);
			} else {
				let autofillR1 = nFillHandleArea === 3 ? r2From + 1 : r1From - 1;
				autoFillRange = getRange(c1To, autofillR1, c2To, r2To);
			}
			ws.selectionRange.ranges = [getRange(c1From, r1From, c2From, r2From)];
			wsView = getAutoFillRange(wsView, c1To, r1To, c2To, r2To, nHandleDirection, nFillHandleArea);
			let undoRes = nHandleDirection === 0 ? [undoData] : undoData;
			let expectedRes = nHandleDirection === 0 ? [expectedData] : expectedData;
			checkUndoRedo(function (_desc) {
				autoFillAssert(assert, autoFillRange, undoRes, _desc);
			}, function (_desc) {
				autoFillAssert(assert, autoFillRange, expectedRes, _desc);
			}, sDescription);
		}

		ws.getRange2('A1:Z100').cleanAll();
		// Case #1: Format Date. Asc sequence. Vertical. Selected 2 cell
		let testData = [
			['12/12/2000', '12/12/2000'],
			['12/13/2000', '12/14/2000']
		];
		let range = ws.getRange4(0,0);
		range.fillData(testData);
		let undoData = [[''], [''], [''], ['']];

		// nFillHandleArea: 1 - Reverse, 3 - asc sequence, 2 - Reverse 1 elem.
		// aFrom, aTo: [c1,c2,r1,r2]
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, 'Format Date. Asc sequence. Vertical. Selected 2 cell', [['36874'], ['36875'], ['36876'], ['36877']]);
		getAutofillCase([1, 1, 0, 1], [1, 1, 2, 5], 3, 'Format Date. Asc sequence. Vertical. Selected 2 cell, even step.', [['36876'], ['36878'], ['36880'], ['36882']]);
		// Case #2: Format Date. Reverse sequence. Vertical. Selected 2 cell
		undoData = [[''],[''],['36873'], ['36872']];
		range = ws.getRange4(4, 0);
		range.fillData(testData);

		getAutofillCase([0, 0, 4, 5], [0, 0, 3, 0], 1, 'Format Date. Reverse sequence. Vertical. Selected 2 cell', [['36871'], ['36870'], ['36869'], ['36868']]);
		undoData = [[''],[''],['36874'], ['36872']];
		getAutofillCase([1, 1, 4, 5], [1, 1, 3, 0], 1, 'Format Date. Reverse sequence. Vertical. Selected 2 cell, even step', [['36870'], ['36868'], ['36866'], ['36864']]);
		// Case #3: Format Date. Asc sequence. Horizontal. Selected 1 cell
		clearData(0, 0, 1, 5);
		testData = [
			['12/12/2000']
		];
		undoData = ['', '', '', ''];
		range = ws.getRange4(0, 0);
		range.fillData(testData);

		getAutofillCase([0, 0, 0, 0], [1, 4, 0, 0], 3, 'Format Date. Asc sequence. Horizontal. Selected 1 cell', ['36873', '36874', '36875', '36876']);
		// Case #4: Format Date. Reverse sequence. Horizontal. Selected 1 cell
		undoData = ['', '', '', '36872'];
		range = ws.getRange4(0, 4);
		range.fillData(testData);

		getAutofillCase([4, 4, 0, 0], [3, 0, 0, 0], 1, 'Format Date. Reverse sequence. Horizontal. Selected 1 cell', ['36871', '36870', '36869', '36868']);
		// Case #5: Format Date. Asc sequence. Vertical. Selected 3 cells. Negative case - 12/12/2000, 12/13/2000, 12/12/2000.
		testData = [
			['12/12/2000'],
			['12/13/2000'],
			['12/12/2000']
		];
		undoData = [[''], [''], [''], ['']];
		range = ws.getRange4(0, 0);
		range.fillData(testData);

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Format Date. Asc sequence. Vertical. Selected 3 cells. Negative case - 12/12/2000, 12/13/2000, 12/12/2000.', [['36872'], ['36873'], ['36872'], ['36872']]);
		// Case #6: Format Date. Reverse sequence. Vertical. Selected 3 cells. Negative case - 12/12/2000, 12/13/2000, 12/12/2000.
		undoData = [[''], ['36872'], ['36873'], ['36872']];
		range = ws.getRange4(4, 0);
		range.fillData(testData);

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Format Date. Reverse sequence. Vertical. Selected 3 cells. Negative case - 12/12/2000, 12/13/2000, 12/12/2000.', [['36872'], ['36873'], ['36872'], ['36872']]);
		// Case #7: Format Date. Asc sequence. Horizontal. Selected 4 cells. Negative case - 12/12/2000, 12/13/2000, 12/14/2000, 12/16/2000.
		clearData(0, 0, 0, 6);
		testData = [
			['12/12/2000', '12/13/2000', '12/14/2000', '12/16/2000']
		];
		undoData = ['', '', '', '', ''];
		range = ws.getRange4(0, 0);
		range.fillData(testData);

		getAutofillCase([0, 3, 0, 0], [4, 8, 0, 0], 3, 'Format Date. Asc sequence. Horizontal. Selected 4 cells. Negative case - 12/12/2000, 12/13/2000, 12/14/2000, 12/16/2000.', ['36872', '36873', '36874', '36876', '36872']);
		// Case #8: Format Date. Reverse sequence. Horizontal. Selected 4 cells. Negative case - 12/12/2000, 12/13/2000, 12/14/2000, 12/16/2000.
		undoData = ['', '36876', '36874', '36873', '36872'];
		range = ws.getRange4(0, 5);
		range.fillData(testData);

		getAutofillCase([5, 8, 0, 0], [4, 0, 0, 0], 1, 'Format Date. Reverse sequence. Horizontal. Selected 4 cells. Negative case - 12/12/2000, 12/13/2000, 12/14/2000, 12/16/2000.', ['36876', '36874', '36873', '36872', '36876']);
		// Case #9: Format Date. Asc sequence. Vertical. Selected 1 cell. Non integer value in Date format.
		ws.getRange2('A1:I1').cleanAll();
		ws.getRange2('A1').setValue('12/12/2000');
		ws._getCell(0, 0,function(cell) {
			cell.setValueNumberInternal(36872.5);
		});
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 0], [0, 0, 1, 4], 3, 'Format Date. Asc sequence. Vertical. Selected 1 cell. Non integer value in Date format.', [['36873'], ['36874'], ['36875'], ['36876']]);
		// Case #10: Format Date. Reverse sequence. Horizontal. Selected 2 cells. Non integer value in Date format.
		ws.getRange2('E1').setValue('12/12/2000');
		ws.getRange2('F1').setValue('12/13/2000');
		ws._getCell(0, 4, function (cell) {
			cell.setValueNumberInternal(36872.5);
		});
		ws._getCell(0, 5, function (cell) {
			cell.setValueNumberInternal(36873.6);
		});
		undoData = ['', '', '', '36872.5'];

		getAutofillCase([4, 5, 0, 0], [3, 0, 0, 0], 1, 'Format Date. Reverse sequence. Horizontal. Selected 2 cells. Non integer value in Date format.', ['36871', '36870', '36869', '36868']);
		// Case #11: Format Date. Asc sequence. Vertical. Selected 2 cells. Step 0.
		ws.getRange2('A1:F9').cleanAll();
		testData = [
			['12/12/2000'],
			['12/12/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, 'Format Date. Asc sequence. Vertical. Selected 2 cells. Step 0.', [['36872'], ['36872'], ['36872'], ['36872']]);
		// Case #12: Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Step - time.
		ws.getRange2('A1:A6').cleanAll();
		testData = [
			['12/12/2000 12:00'],
			['12/12/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, 'Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Step - time.', [['36872.583333333336'], ['36872.62500000001'], ['36872.66666666668'], ['36872.70833333335']]);
		// Case #13: Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Step - time.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36872.541666666664'], ['36872.5']];

		getAutofillCase([0, 0, 4, 5], [0, 0, 3, 0], 1, 'Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Step - time.', [['36872.45833333334'], ['36872.416666666686'], ['36872.37500000003'], ['36872.33333333337']]);
		// Case #14: Format Date & Time. Asc sequence. Horizontal. Selected 1 cell.
		ws.getRange2('A1:A6').cleanAll();
		testData = [
			['12/12/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([0, 0, 0, 0], [1, 4, 0, 0], 3, 'Format Date & Time. Asc sequence. Horizontal. Selected 1 cell.', ['36873.541666666664', '36874.541666666664', '36875.541666666664', '36876.541666666664']);
		// Case #15: Format Date & Time. Reverse sequence. Horizontal. Selected 1 cell.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '', '', '36872.541666666664'];

		getAutofillCase([4, 4, 0, 0], [3, 0, 0, 0], 1, 'Format Date & Time. Reverse sequence. Horizontal. Selected 1 cell.', ['36871.541666666664', '36870.541666666664', '36869.541666666664', '36868.541666666664']);
		// Case #16: Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Diff days, same time.
		ws.getRange2('A1:F1').cleanAll();
		testData = [
			['12/12/2000 14:00'],
			['12/14/2000 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, 'Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Diff days, same time.', [['36876.583333333336'], ['36878.583333333336'], ['36880.583333333336'], ['36882.583333333336']]);
		// Case #17: Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Diff days, same time.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36874.583333333336'], ['36872.583333333336']];

		getAutofillCase([0, 0, 4, 5], [0, 0, 3, 0], 1, 'Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Diff days, same time.', [['36870.583333333336'], ['36868.583333333336'], ['36866.583333333336'], ['36864.583333333336']]);
		// Case #18: Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Step - time & day.
		testData = [
			['12/12/2000 12:00'],
			['12/14/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36872.583333333336'], ['36874.583333333336']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, 'Format Date & Time. Asc sequence. Vertical. Selected 2 cells. Step - time & day.', [['36876.583333333336'], ['36878.62500000001'], ['36880.66666666668'], ['36882.70833333335']]);
		// Case #19: Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Step - time & day.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36874.541666666664'], ['36872.5']];

		getAutofillCase([0, 0, 4, 5], [0, 0, 3, 0], 1, 'Format Date & Time. Reverse sequence. Vertical. Selected 2 cells. Step - time & day.', [['36870.45833333334'], ['36868.416666666686'], ['36866.37500000003'], ['36864.33333333337']]);
		// Case #20: Format Date & Time. Asc sequence. Horizontal. Negative case. Same day, time - 12:00, 13:00, 12:00.
		ws.getRange2('A1:A6').cleanAll();
		testData = [
			['12/12/2000 12:00', '12/12/2000 13:00', '12/12/2000 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([0, 2, 0, 0], [3, 6, 0, 0], 3, 'Format Date & Time. Asc sequence. Horizontal. Negative case: Same day, time - 12:00, 13:00, 12:00.', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #21: Format Date & Time. Reverse sequence. Horizontal. Negative case. Same day, time - 12:00, 13:00, 12:00.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '36872.5', '36872.541666666664', '36872.5'];

		getAutofillCase([4, 6, 0, 0], [3, 0, 0, 0], 1, 'Format Date & Time. Reverse sequence. Horizontal. Negative case: Same day, time - 12:00, 13:00, 12:00.', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #22: Format Date & Time. Asc sequence. Horizontal. Negative case. Same day, time - 13:00, 14:00, 16:00.
		testData = [
			['12/12/2000 13:00', '12/12/2000 14:00', '12/12/2000 16:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '36872.5', '36872.541666666664', '36872.5'];

		getAutofillCase([0, 2, 0, 0], [3, 6, 0, 0], 3, 'Format Date & Time. Asc sequence. Horizontal. Negative case: Same day, time - 13:00, 14:00, 16:00.', ['36872.541666666664', '36872.541666666664', '36872.541666666664', '36872.541666666664']);
		// Case #23: Format Date & Time. Reverse sequence. Horizontal. Negative case. Same day, time - 13:00, 14:00, 16:00.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '36872.666666666664', '36872.583333333336', '36872.541666666664'];

		getAutofillCase([4, 6, 0, 0], [3, 0, 0, 0], 1, 'Format Date & Time. Reverse sequence. Horizontal. Negative case: Same day, time - 13:00, 14:00, 16:00.', ['36872.541666666664', '36872.541666666664', '36872.541666666664', '36872.541666666664']);
		// Case #24: Format Date & Time. Asc sequence. Vertical. Negative case. Diff days, same time.
		ws.getRange2('A1:F1').cleanAll();
		testData = [
			['12/12/2000 12:00'],
			['12/13/2000 12:00'],
			['12/12/2000 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Format Date & Time. Asc sequence. Vertical. Negative case: Diff days, same time.', [['36872.5'], ['36873.5'], ['36872.5'], ['36872.5']]);
		// Case #25: Format Date & Time. Reverse sequence. Vertical. Negative case. Diff days, same time.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], ['36872.5'], ['36873.5'], ['36872.5']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Format Date & Time. Reverse sequence. Vertical. Negative case: Diff days, same time.', [['36872.5'], ['36873.5'], ['36872.5'], ['36872.5']]);
		// Case #26: Format Date & Time. Asc sequence. Vertical. Negative case. Diff days & time.
		testData = [
			['12/12/2000 12:00'],
			['12/13/2000 13:00'],
			['12/15/2000 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], ['36872.5'], ['36873.5'], ['36872.5']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Format Date & Time. Asc sequence. Vertical. Negative case: Diff days & time.', [['36872.5'], ['36873.541666666664'], ['36875.583333333336'], ['36872.5']]);
		// Case #27: Format Date & Time. Reverse sequence. Vertical. Negative case. Diff days & time.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], ['36875.583333333336'], ['36873.541666666664'], ['36872.5']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Format Date & Time. Reverse sequence. Vertical. Negative case: Diff days & time.', [['36875.583333333336'], ['36873.541666666664'], ['36872.5'], ['36875.583333333336']]);
		// Case #28: Mixed format. Asc sequence. Horizontal. Same day, diff time.
		ws.getRange2('A1:G7').cleanAll();
		testData = [
			['12/12/2000', '12/12/2000 12:00', '12/12/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([0, 2, 0, 0], [3, 6, 0, 0], 3, 'Mixed format. Asc sequence. Horizontal. Same day, diff time.', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #29: Mixed format. Reverse sequence. Horizontal. Same day, diff time. First cell with Date format
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '36872.541666666664', '36872.5', '36872'];

		getAutofillCase([4, 6, 0, 0], [3, 0, 0, 0], 1, 'Mixed format. Reverse sequence. Horizontal. Same day, diff time. First Date', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #30: Mixed format. Asc sequence. Horizontal. Same day, diff time. First cell with Date & Time format
		testData = [
			['12/12/2000 12:00', '12/12/2000', '12/12/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '36872', '36872.5', '36872.541666666664'];

		getAutofillCase([0, 2, 0, 0], [3, 6, 0, 0], 3, 'Mixed format. Asc sequence. Horizontal. Same day, diff time. First Date & Time', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #31: Mixed format. Reverse sequence. Horizontal. Same day, diff time. First cell with Date & Time format
		ws.getRange2('A1:F1').cleanAll();
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([4, 6, 0, 0], [3, 0, 0, 0], 1, 'Mixed format. Reverse sequence. Horizontal. Same day, diff time. First Date & Time', ['36872.5', '36872.5', '36872.5', '36872.5']);
		// Case #32: Mixed format.  Asc sequence. Vertical. Diff day, same time. First  Date format.
		ws.getRange2('A1:F1').cleanAll();
		testData = [
			['12/12/2000'],
			['12/13/2000 12:00'],
			['12/14/2000 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Mixed format. Asc sequence. Vertical. Diff day, same time. First Date', [['36875.5'], ['36876.5'], ['36877.5'], ['36878.5']]);
		// Case #33: Mixed format. Reverse sequence. Vertical. Diff day, same time. First  Date format.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], ['36874.5'], ['36873.5'], ['36872']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Mixed format. Reverse sequence. Vertical. Diff day, same time. First Date', [['36871.5'], ['36870.5'], ['36869.5'], ['36868.5']]);
		// Case #34: Mixed format.  Asc sequence. Vertical. Diff day, same time. First  Date & Time format.
		ws.getRange2('A1:A8').cleanAll();
		testData = [
			['12/12/2000 12:00'],
			['12/14/2000 12:00'],
			['12/16/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Mixed format. Asc sequence. Vertical. Diff day, same time. First Date & Time', [['36878.5'], ['36880.5'], ['36882.5'], ['36884.5']]);
		// Case #35: Mixed format. Reverse sequence. Vertical. Diff day, same time. First  Date & Time format.
		ws.getRange2('A1:A8').cleanAll();
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Mixed format. Reverse sequence. Vertical. Diff day, same time. First Date & Time', [['36870.5'], ['36868.5'], ['36866.5'], ['36864.5']]);
		// Case #36: Mixed format.  Asc sequence.  Vertical. Negative case: 12/12/2000, 12/13/2000, 12/12/2000 12:00. First Date format.
		ws.getRange2('A1:A8').cleanAll();
		testData = [
			['12/12/2000'],
			['12/13/2000 12:00'],
			['12/12/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Mixed format. Asc sequence. Vertical. Negative case. First Date', [['36872'], ['36873.5'], ['36872.541666666664'], ['36872']]);
		// Case #37: Mixed format. Reverse sequence.  Vertical. Negative case: 12/12/2000, 12/13/2000 12:00, 12/12/2000 13:00. First Date format.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], ['36872.541666666664'], ['36873.5'], ['36872']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Mixed format. Reverse sequence. Vertical. Negative case. First Date', [['36872.541666666664'], ['36873.5'], ['36872'], ['36872.541666666664']]);
		// Case #38: Mixed format.  Asc sequence.  Vertical. Negative case: 12/12/2000 12:00, 12/13/2000, 12/15/2000 13:00. First Date & Time format.
		ws.getRange2('A1:A8').cleanAll();
		testData = [
			['12/12/2000 12:00'],
			['12/13/2000'],
			['12/15/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 6], 3, 'Mixed format. Asc sequence. Vertical. Negative case. First Date & Time', [['36872.5'], ['36873'], ['36875.541666666664'], ['36872.5']]);
		// Case #39: Mixed format. Reverse sequence.  Vertical. Negative case: 12/12/2000 12:00, 12/13/2000, 12/15/2000 13:00. First Date & Time format.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], ['36875.541666666664'], ['36873'], ['36872.5']];

		getAutofillCase([0, 0, 4, 6], [0, 0, 3, 0], 1, 'Mixed format. Reverse sequence. Vertical. Negative case. First Date & Time', [['36875.541666666664'], ['36873'], ['36872.5'], ['36875.541666666664']]);
		// Case #40: 1900 year. Asc sequence. Horizontal. One selected cell. Date format.
		ws.getRange2('A1:A8').cleanAll();
		testData = [
			['01/01/1900']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([0, 0, 0, 0], [1, 4, 0, 0], 3, '1900 year. Asc sequence. Horizontal. One selected cell. Date format', ['2', '3', '4', '5']);
		// Case #41: 1900 year.  Reverse sequence. Horizontal. One selected cell. Date format.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '', '', '1'];

		getAutofillCase([4, 4, 0, 0], [3, 0, 0, 0], 1, '1900 year. Reverse sequence. Horizontal. One selected cell. Date format', ['0', '1', '1', '1']);
		// Case #42: 1900 year.  Asc sequence. Horizontal. Date format.
		testData = [
			['01/01/1900', '01/03/1900']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '1', ''];

		getAutofillCase([0, 1, 0, 0], [2, 5, 0, 0], 3, '1900 year. Asc sequence. Horizontal. Date format', ['5', '7', '9', '11']);
		// Case #43: 1900 year.  Reverse sequence. Horizontal. Date format.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '', '3', '1'];

		getAutofillCase([4, 5, 0, 0], [3, 0, 0, 0], 1, '1900 year. Reverse sequence. Horizontal. Date format', ['3', '1', '3', '1']);
		// Case #44: 1900 year.  Asc sequence. Vertical. One selected cell. Date & Time format.
		ws.getRange2('A1:F1').cleanAll();
		testData = [
			['01/01/1900 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 0], [0, 0, 1, 4], 3, '1900 year. Asc sequence. Vertical. One selected cell. Date & Time format', [['2.5'], ['3.5'], ['4.5'], ['5.5']]);
		// Case #45: 1900 year.  Reverse sequence. Vertical. One selected cell. Date & Time format.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['1.5']];

		getAutofillCase([0, 0, 4, 4], [0, 0, 3, 0], 1, '1900 year. Reverse sequence. Vertical. One selected cell. Date & Time format', [['0.5'], ['1.5'], ['1.5'], ['1.5']]);
		// Case #46: 1900 year.  Asc sequence. Vertical. Date & Time format.
		testData = [
			['01/01/1900 1:00'],
			['01/01/1900 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['1.5'], ['']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 5], 3, '1900 year. Asc sequence. Vertical. Date & Time format', [['1.958333333333333'], ['2.416666666666666'], ['2.874999999999999'], ['3.333333333333332']]);
		// Case #47: 1900 year.  Reverse sequence. Vertical. Date & Time format.
		range = ws.getRange4(4, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['1.5'], ['1.0416666666666667']];

		getAutofillCase([0, 0, 4, 5], [0, 0, 3, 0], 1, '1900 year. Reverse sequence. Vertical. Date & Time format', [['0.5833333333333333'], ['0.12499999999999978'], ['-0.3333333333333337'], ['-0.7916666666666672']]);
		// Case #48: 1900 year.  Asc sequence. Horizontal. Mixed format.
		ws.getRange2('A1:A8').cleanAll();
		testData = [
			['01/01/1900', '01/02/1900 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', ''];

		getAutofillCase([0, 1, 0, 0], [2, 5, 0, 0], 3, '1900 year. Asc sequence. Horizontal. Mixed format', ['3.5', '4.5', '5.5', '6.5']);
		// Case #49: 1900 year.  Reverse sequence. Horizontal. Mixed format.
		range = ws.getRange4(0, 4);
		range.fillData(testData);
		undoData = ['', '', '2.5', '1'];

		getAutofillCase([4, 5, 0, 0], [3, 0, 0, 0], 1, '1900 year. Reverse sequence. Horizontal. Mixed format', ['0.5', '1', '2.5', '1']);
		// Case #50: Time format. Asc sequence. Vertical. One selected cell.
		ws.getRange2('A1:F1').cleanAll();
		testData = [
			['12:00:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 0], [0, 0, 1, 4], 3, 'Time format. Asc sequence. Vertical. One selected cell', [['0.5416666666666666'], ['0.5833333333333334'], ['0.625'], ['0.6666666666666666']]);
		// Case #51: Time format.  Reverse sequence. Vertical. One selected cell.
		testData = [['0:00:00']]
		range = ws.getRange4(24, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''],
			[''], [''], [''], [''], [''], [''], [''], [''], ['0.5']];
		let expectedData = [['0.9583333333333334'], ['0.9166666666666666'], ['0.875'], ['0.8333333333333334'], ['0.7916666666666667'],
			['0.75'], ['0.7083333333333334'], ['0.6666666666666667'], ['0.625'], ['0.5833333333333334'], ['0.5416666666666667'],
			['0.5'], ['0.45833333333333337'], ['0.41666666666666674'], ['0.375'], ['0.33333333333333337'], ['0.29166666666666674'],
			['0.25'], ['0.20833333333333337'], ['0.16666666666666674'], ['0.125'], ['0.08333333333333337'], ['0.04166666666666674'], ['0']];

		getAutofillCase([0, 0, 24, 24], [0, 0, 23, 0], 1, 'Time format. Reverse sequence. Vertical. One selected cell', expectedData);
		// Case #52: Time format. Asc sequence. Vertical. Multiple selected cells.
		testData = [
			['12:00:00'],
			['14:00:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Time format. Asc sequence. Vertical. Multiple selected cells', [['0.6666666666666666'], ['0.7499999999999999'], ['0.8333333333333331'], ['0.9166666666666664'], ['0.9999999999999997'], ['1.083333333333333']]);
		// Case #53: Time format.  Reverse sequence. Vertical. Multiple selected cells.
		testData = [
			['23:00:00'],
			['0:00:00']
		];
		range = ws.getRange4(23, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''],
			[''], [''], [''], [''], ['0.5833333333333334'], ['0.5']];
		expectedData = [
			['1.9166666666666667'], ['2.875'], ['3.8333333333333335'], ['4.791666666666667'], ['5.75'],
			['6.708333333333334'], ['7.666666666666667'], ['8.625'], ['9.583333333333334'], ['10.541666666666668'], ['11.5'],
			['12.458333333333334'], ['13.416666666666668'], ['14.375'], ['15.333333333333334'], ['16.291666666666668'],
			['17.25'], ['18.208333333333336'], ['19.166666666666668'], ['20.125'], ['21.083333333333336'], ['22.041666666666668'], ['23']
		];

		getAutofillCase([0, 0, 23, 24], [0, 0, 22, 0], 1, 'Time format. Reverse sequence. Vertical. Multiple selected cells', expectedData);
		// Case #54: Time format.  Asc sequence. Horizontal. Multiple selected cells. Start from 1.
		ws.getRange2('A1').setValue('0:00:00');
		ws.getRange2('B1').setValue('1:00:00');
		ws._getCell(0, 0, function (cell) {
			cell.setValueNumberInternal(1);
		});
		ws._getCell(0, 1, function (cell) {
			cell.setValueNumberInternal(1.04166666666667);
		});
		undoData = ['', '', '', ''];

		getAutofillCase([0, 1, 0, 0], [2, 5, 0, 0], 3, 'Time format. Asc sequence. Horizontal. Multiple selected cells. Start from 1.', ['1.0833333333333401', '1.1250000000000102', '1.1666666666666803', '1.2083333333333504']);
		// Case #55: Time format.  Reverse sequence. Horizontal. Multiple selected cells. Start from 1.
		ws.getRange2('A1:Z30').cleanAll();
		ws.getRange2('Y1').setValue('23:00:00');
		ws.getRange2('Z1').setValue('00:00:00');
		ws._getCell(0, 24, function (cell) {
			cell.setValueNumberInternal(0.958333333333333);
		});
		ws._getCell(0, 25, function (cell) {
			cell.setValueNumberInternal(1);
		});
		undoData = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
		expectedData = ['0.9166666666666661', '0.8749999999999991', '0.8333333333333321', '0.7916666666666652', '0.7499999999999982', '0.7083333333333313', '0.6666666666666643',
			'0.6249999999999973', '0.5833333333333304', '0.5416666666666634', '0.49999999999999645', '0.4583333333333295', '0.4166666666666625', '0.37499999999999556', '0.3333333333333286',
			'0.29166666666666163', '0.24999999999999467', '0.2083333333333277', '0.16666666666666075', '0.12499999999999378', '0.08333333333332682', '0.04166666666665986', '1', '0.958333333333333'];

		getAutofillCase([24, 25, 0, 0], [23, 0, 0, 0], 1, 'Time format. Reverse sequence. Horizontal. Multiple selected cells. Start from 1.', expectedData);
		// Case #56: Time format.  Asc sequence. Vertical. Multiple selected cells. Start from 36872.5.
		ws.getRange2('A1:Z1').cleanAll();
		ws.getRange2('A1').setValue('12:00:00');
		ws.getRange2('A2').setValue('13:00:00');
		ws._getCell(0, 0, function (cell) {
			cell.setValueNumberInternal(36872.5);
		});
		ws._getCell(1, 0, function (cell) {
			cell.setValueNumberInternal(36873.5416666667);
		});
		undoData = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], ['']];
		expectedData = [['36872.5833333334'], ['36872.6250000001'], ['36872.6666666668'], ['36872.7083333335'], ['36872.750000000204'], ['36872.791666666904'], ['36872.833333333605'],
			['36872.875000000306'], ['36872.916666667006'], ['36872.95833333371'], ['36873.00000000041']];

		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 12], 3, 'Time format. Asc sequence. Vertical. Multiple selected cells. Start from 36872.5.', expectedData);
		// Case #57: Time format.  Reverse sequence. Vertical. Multiple selected cells. Start from 36872.5.
		ws.getRange2('A14').setValue('12:00:00');
		ws.getRange2('A15').setValue('13:00:00');
		ws._getCell(13, 0, function (cell) {
			cell.setValueNumberInternal(36872.5);
		});
		ws._getCell(14, 0, function (cell) {
			cell.setValueNumberInternal(36873.5416666667);
		});
		undoData = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], ['36873.5416666667'], ['36872.5']];
		expectedData = [['36872.4583333333'], ['36872.4166666666'], ['36872.3749999999'], ['36872.3333333332'], ['36872.2916666665'],
			['36872.249999999796'], ['36872.208333333096'], ['36872.166666666395'], ['36872.124999999694'], ['36872.083333332994'], ['36872.04166666629'],
			['36871.99999999959'], ['36871.95833333289']];

		getAutofillCase([0, 0, 13, 14], [0, 0, 12, 0], 1, 'Time format. Reverse sequence. Vertical. Multiple selected cells. Start from 36872.5.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #58: Date format. Asc sequence. Horizontal. Two selected cells. Step - month.
		testData = [
			['01/01/2000', '02/01/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['36586', '36617', '36647', '36678', '36708', '36739'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Date format. Asc sequence. Horizontal. Two selected cells. Step - month.', expectedData);
		// Case #59: Date format. Reverse sequence. Horizontal. Two selected cells. Step - month.
		range = ws.getRange4(0, 6);
		range.fillData(testData);
		undoData = ['', '', '', '', '36557', '36526'];
		expectedData = ['36495', '36465', '36434', '36404', '36373', '36342'];
		getAutofillCase([6, 7, 0, 0], [5, 0, 0, 0], 1, 'Date format. Reverse sequence. Horizontal. Two selected cells. Step - month.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #60: Date format. Asc sequence. Vertical. Three selected cells. Step - month.
		testData = [
			['01/01/2000'],
			['03/01/2000'],
			['05/01/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], [''], [''], ['']];
		expectedData = [['36708'], ['36770'], ['36831'], ['36892'], ['36951']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Date format. Asc sequence. Vertical. Three selected cells. Step - month.', expectedData);
		// Case #61: Date format. Reverse sequence. Vertical. Three selected cells. Step - month.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36647'], ['36586'], ['36526']];
		expectedData = [['36465'], ['36404'], ['36342'], ['36281'], ['36220']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Date format. Reverse sequence. Vertical. Three selected cells. Step - month.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #62: Date format. Asc sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.
		testData = [
			['01/01/2000', '03/01/2000', '04/01/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', ''];
		expectedData = ['36526', '36586', '36617', '36526', '36586'];
		getAutofillCase([0, 2, 0, 0], [3, 7, 0, 0], 3, 'Date format. Asc sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		// Case #63: Date format. Reverse sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.
		range = ws.getRange4(0, 5);
		range.fillData(testData);
		undoData = ['', '', '36617', '36586', '36526'];
		expectedData = ['36617', '36586', '36526', '36617', '36586'];
		getAutofillCase([5, 7, 0, 0], [4, 0, 0, 0], 1, 'Date format. Reverse sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #64: Date format. Asc sequence. Vertical. Two selected cells. Step - month. 1900 year.
		testData = [
			['01/01/1900'],
			['02/01/1900']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['61'], ['92'], ['122'], ['153'], ['183'], ['214']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Date format. Asc sequence. Vertical. Two selected cells. Step - month. 1900 year.', expectedData);
		// Case #65: Date format. Reverse sequence. Vertical. Two selected cells. Step - month. 1900 year.
		testData = [
			['04/01/1900'],
			['06/01/1900']
		];
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['32'], ['1']];
		expectedData = [['32'], ['92'], ['153'], ['92'], ['153'], ['92']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Date format. Reverse sequence. Vertical. Two selected cells. Step - month. 1900 year.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #66: Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - month.
		testData = [
			['01/01/2000 12:00', '02/01/2000 13:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', '', '', '', '', '', '', ''];
		expectedData = ['36586.583333333336', '36617.62500000001', '36647.66666666668', '36678.70833333335', '36708.75000000002', '36739.79166666669', '36770.833333333365', '36800.87500000004', '36831.91666666671', '36861.95833333338', '36893.00000000005', '36924.04166666672'];
		getAutofillCase([0, 1, 0, 0], [2, 13, 0, 0], 3, 'Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - month.', expectedData);
		// Case #67: Date & Time format. Reverse sequence. Horizontal. Two selected cells. Step - month.
		range = ws.getRange4(0, 12);
		range.fillData(testData);
		undoData = ['', '', '', '', '', '', '', '', '', '', '36557.541666666664', '36526.5'];
		expectedData = ['36495.45833333334', '36465.416666666686', '36434.37500000003', '36404.33333333337', '36373.291666666715', '36342.25000000006', '36312.2083333334', '36281.166666666744', '36251.12500000009', '36220.08333333343', '36192.04166666677', '36161.00000000012'];
		getAutofillCase([12, 13, 0, 0], [11, 0, 0, 0], 1, 'Date & Time format. Reverse sequence. Horizontal. Two selected cells. Step - month.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #68: Date & Time format. Asc sequence. Vertical. Three selected cells. Step - month.
		testData = [
			['02/01/2000 12:00'],
			['04/01/2000 14:00'],
			['06/01/2000 16:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['']];
		expectedData = [['36739.5'], ['36800.5'], ['36861.5'], ['36923.5'], ['36982.5']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Date & Time format. Asc sequence. Vertical. Three selected cells. Step - month.', expectedData);
		// Case #69: Date & Time format. Reverse sequence. Vertical. Three selected cells. Step - month.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36678.666666666664'], ['36617.583333333336'], ['36557.5']];
		expectedData = [['36495.5'], ['36434.5'], ['36373.5'], ['36312.5'], ['36251.5']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Date & Time format. Reverse sequence. Vertical. Three selected cells. Step - month.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #70: Date & Time format. Asc sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.
		testData = [
			['02/01/2000 12:00', '04/01/2000 14:00', '07/01/2000 16:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', ''];
		expectedData = ['36557.5', '36617.583333333336', '36708.666666666664', '36557.5', '36617.583333333336'];
		getAutofillCase([0, 2, 0, 0], [3, 7, 0, 0], 3, 'Date & Time format. Asc sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		// Case #71: Date & Time format. Reverse sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.
		range = ws.getRange4(0, 5);
		range.fillData(testData);
		undoData = ['', '', '36708.666666666664', '36617.583333333336', '36557.5'];
		expectedData = ['36708.666666666664', '36617.583333333336', '36557.5', '36708.666666666664', '36617.583333333336'];
		getAutofillCase([5, 7, 0, 0], [4, 0, 0, 0], 1, 'Date & Time format. Reverse sequence. Horizontal. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #72: Date & Time format. Asc sequence. Vertical. Two selected cells. Step - month. 1900 year.
		testData = [
			['02/01/1900 12:00'],
			['04/01/1900 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['153.66666666666666'], ['214.75'], ['275.8333333333333'], ['336.91666666666663'], ['398.99999999999994'], ['458.08333333333326']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Date & Time format. Asc sequence. Vertical. Two selected cells. Step - month. 1900 year.', expectedData);
		// Case #73: Date & Time format. Reverse sequence. Vertical. Two selected cells. Step - month. 1900 year.
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['92.58333333333333'], ['32.5']];
		expectedData = [['92.58333333333333'], ['32.5'], ['92.58333333333333'], ['32.5'], ['92.58333333333333'], ['32.5']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Date & Time format. Reverse sequence. Vertical. Two selected cells. Step - month. 1900 year.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #74: Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - month. Asc month seq, but time sequence is reverse.
		testData = [
			['02/01/2000 12:00', '04/01/2000 5:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['36677.916666666664', '36738.62499999999', '36799.33333333332', '36860.04166666665', '36921.74999999998', '36980.45833333331'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - month. Asc month seq, but time sequence is reverse.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #75: Mixed date format. Asc sequence. Vertical. Two selected cells. Step - month.
		testData = [
			['01/01/2000 12:00'],
			['02/01/2000']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['36586.5'], ['36617.5'], ['36647.5'], ['36678.5'], ['36708.5'], ['36739.5']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Mixed date format. Asc sequence. Vertical. Two selected cells. Step - month.', expectedData);
		// Case #76: Mixed date format. Reverse sequence. Vertical. Two selected cells. Step - month.
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['36557'], ['36526.5']];
		expectedData = [['36495.5'], ['36465.5'], ['36434.5'], ['36404.5'], ['36373.5'], ['36342.5']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Mixed date format. Reverse sequence. Vertical. Two selected cells. Step - month.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #77: Mixed date format. Asc sequence. Horizontal. Three selected cells. Step - month.
		testData = [
			['01/01/2000 12:00', '04/01/2000', '07/01/2000 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', ''];
		expectedData = ['36800.5', '36892.5', '36982.5', '37073.5', '37165.5'];
		getAutofillCase([0, 2, 0, 0], [3, 7, 0, 0], 3, 'Mixed date format. Asc sequence. Horizontal. Three selected cells. Step - month.', expectedData);
		// Case #78: Mixed date format. Reverse sequence. Horizontal. Three selected cells. Step - month.
		range = ws.getRange4(0, 5);
		range.fillData(testData);
		undoData = ['', '', '36708.583333333336', '36617', '36526.5'];
		expectedData = ['36434.5', '36342.5', '36251.5', '36161.5', '36069.5'];
		getAutofillCase([5, 7, 0, 0], [4, 0, 0, 0], 1, 'Mixed date format. Reverse sequence. Horizontal. Three selected cells. Step - month.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #79: Mixed date format. Asc sequence. Vertical. Three selected cells. Step - month. Negative case - incorrect sequence.
		testData = [
			['01/01/2000 12:00'],
			['03/01/2000'],
			['04/01/2000 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['']];
		expectedData = [['36526.5'], ['36586'], ['36617.583333333336'], ['36526.5'], ['36586']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Mixed date format. Asc sequence. Vertical. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		// Case #80: Mixed date format. Reverse sequence. Vertical. Three selected cells. Step - month. Negative case - incorrect sequence.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['36617.583333333336'], ['36586'], ['36526.5']];
		expectedData = [['36617.583333333336'], ['36586'], ['36526.5'], ['36617.583333333336'], ['36586']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Mixed date format. Reverse sequence. Vertical. Three selected cells. Step - month. Negative case - incorrect sequence.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #81: Mixed date format. Asc sequence. Horizontal. Two selected cells. Step - month. 1900 year.
		testData = [
			['01/01/1900', '02/01/1900 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['61.5', '92.5', '122.5', '153.5', '183.5', '214.5'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Mixed date format. Asc sequence. Horizontal. Two selected cells. Step - month. 1900 year.', expectedData);
		// Case #82: Mixed date format. Reverse sequence. Horizontal. Two selected cells. Step - month. 1900 year.
		range = ws.getRange4(0, 6);
		range.fillData(testData);
		undoData = ['', '', '', '', '32.5', '1'];
		expectedData = ['32.5', '1', '32.5', '1', '32.5', '1'];
		getAutofillCase([6, 7, 0, 0], [5, 0, 0, 0], 1, 'Mixed date format. Reverse sequence. Horizontal. Two selected cells. Step - month. 1900 year.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #83: Date format. Asc sequence. Vertical. Two selected cells. Step - year.
		testData = [
			['01/01/2000'],
			['01/01/2001']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['37257'], ['37622'], ['37987'], ['38353'], ['38718'], ['39083']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Date format. Asc sequence. Vertical. Two selected cells. Step - year.', expectedData);
		// Case #84: Date format. Reverse sequence. Vertical. Two selected cells. Step - year.
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['36892'], ['36526']];
		expectedData = [['36161'], ['35796'], ['35431'], ['35065'], ['34700'], ['34335']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Date format. Reverse sequence. Vertical. Two selected cells. Step - year.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #85: Date format. Asc sequence. Horizontal. Three selected cells. Step - year.
		testData = [
			['01/01/2000', '01/01/2025', '01/01/2050']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', ''];
		expectedData = ['63920', '73051', '82182', '91313', '100444'];
		getAutofillCase([0, 2, 0, 0], [3, 7, 0, 0], 3, 'Date format. Asc sequence. Horizontal. Three selected cells. Step - year.', expectedData);
		// Case #86: Date format. Reverse sequence. Horizontal. Three selected cells. Step - year.
		range = ws.getRange4(0, 5);
		range.fillData(testData);
		undoData = ['', '',  '54789', '45658', '36526'];
		expectedData = ['27395', '18264', '9133', '1', '45658'];
		getAutofillCase([5, 7, 0, 0], [4, 0, 0, 0], 1, 'Date format. Reverse sequence. Horizontal. Three selected cells. Step - year.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #87: Date format. Asc sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.
		testData = [
			['01/01/2000'],
			['01/01/2002'],
			['01/01/2005']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['']];
		expectedData = [['36526'], ['37257'], ['38353'], ['36526'], ['37257']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Date format. Asc sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.', expectedData);
		// Case #88: Date format. Reverse sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['38353'], ['37257'], ['36526']];
		expectedData = [['38353'], ['37257'], ['36526'], ['38353'], ['37257']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Date format. Reverse sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #89: Date format. Asc sequence. Horizontal. Two selected cells. Step - year. 1900 year.
		testData = [
			['01/01/1900', '01/01/1910']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['7306', '10959', '14611', '18264', '21916', '25569'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Date format. Asc sequence. Horizontal. Two selected cells. Step - year. 1900 year.', expectedData);
		// Case #90: Date format. Reverse sequence. Horizontal. Two selected cells. Step - year. 1900 year.
		range = ws.getRange4(0, 6);
		range.fillData(testData);
		undoData = ['', '', '', '', '3654', '1'];
		expectedData = ['3654', '1', '3654', '1', '3654', '1'];
		getAutofillCase([6, 7, 0, 0], [5, 0, 0, 0], 1, 'Date format. Reverse sequence. Horizontal. Two selected cells. Step - year. 1900 year.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #91: Date & Time format. Asc sequence. Vertical. Two selected cells. Step - year.
		testData = [
			['01/01/2000 12:00'],
			['01/01/2100 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['109575.66666666666'], ['146099.75'], ['182623.8333333333'], ['219148.91666666663'], ['255672.99999999994'], ['292197.08333333326']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Date & Time format. Asc sequence. Vertical. Two selected cells. Step - year.', expectedData);
		// Case #92: Date & Time format. Reverse sequence. Vertical. Two selected cells. Step - year.
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['73051.58333333333'], ['36526.5']];
		expectedData = [['1.4166666666715173'], ['36526.5'], ['73051.58333333333'], ['36526.5'], ['73051.58333333333'], ['36526.5']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Date & Time format. Reverse sequence. Vertical. Two selected cells. Step - year.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #93: Date & Time format. Asc sequence. Horizontal. Three selected cells. Step - year.
		testData = [
			['01/01/2000 12:00', '01/01/2005 13:00', '01/01/2010 14:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['42005.5', '43831.5', '45658.5', '47484.5', '49310.5'];
		getAutofillCase([0, 2, 0, 0], [3, 7, 0, 0], 3, 'Date & Time format. Asc sequence. Horizontal. Three selected cells. Step - year.', expectedData);
		// Case #94: Date & Time format. Reverse sequence. Horizontal. Three selected cells. Step - year.
		range = ws.getRange4(0, 5);
		range.fillData(testData);
		undoData = ['', '', '40179.583333333336', '38353.541666666664', '36526.5'];
		expectedData = ['34700.5', '32874.5', '31048.5', '29221.5', '27395.5'];
		getAutofillCase([5, 7, 0, 0], [4, 0, 0, 0], 1, 'Date & Time format. Reverse sequence. Horizontal. Three selected cells. Step - year.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #95: Date & Time format. Asc sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.
		testData = [
			['01/01/2000 12:00'],
			['01/01/2002 12:00'],
			['01/01/2005 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['']];
		expectedData =[['36526.5'], ['37257.5'], ['38353.5'], ['36526.5'], ['37257.5']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Date & Time format. Asc sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.', expectedData);
		// Case #96: Date & Time format. Reverse sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['38353.5'], ['37257.5'], ['36526.5']];
		expectedData = [['38353.5'], ['37257.5'], ['36526.5'], ['38353.5'], ['37257.5']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Date & Time format. Reverse sequence. Vertical. Three selected cells. Step - year. Negative case - incorrect sequence.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #97: Mixed date format. Asc sequence. Horizontal. Two selected cells. Step - year.
		testData = [
			['01/01/2000', '01/01/2001 12:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['37257.5', '37622.5', '37987.5', '38353.5', '38718.5', '39083.5'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Mixed date format. Asc sequence. Horizontal. Two selected cells. Step - year.', expectedData);
		// Case #98: Mixed date format. Reverse sequence. Horizontal. Two selected cells. Step - year.
		range = ws.getRange4(0, 6);
		range.fillData(testData);
		undoData = ['', '', '', '', '36892.5', '36526'];
		expectedData = ['36161.5', '35796.5', '35431.5', '35065.5', '34700.5', '34335.5'];
		getAutofillCase([6, 7, 0, 0], [5, 0, 0, 0], 1, 'Mixed date format. Reverse sequence. Horizontal. Two selected cells. Step - year.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #99: Mixed date format. Asc sequence. Vertical. Three selected cells. Step - year.
		testData = [
			['01/01/2000 12:00'],
			['01/01/2002 12:00'],
			['01/01/2004']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['']];
		expectedData = [['38718.5'], ['39448.5'], ['40179.5'], ['40909.5'], ['41640.5']];
		getAutofillCase([0, 0, 0, 2], [0, 0, 3, 7], 3, 'Mixed date format. Asc sequence. Vertical. Three selected cells. Step - year.', expectedData);
		// Case #100: Mixed date format. Reverse sequence. Vertical. Three selected cells. Step - year.
		range = ws.getRange4(5, 0);
		range.fillData(testData);
		undoData = [[''], [''], ['37987'], ['37257.5'], ['36526.5']];
		expectedData = [['35796.5'], ['35065.5'], ['34335.5'], ['33604.5'], ['32874.5']];
		getAutofillCase([0, 0, 5, 7], [0, 0, 4, 0], 1, 'Mixed date format. Reverse sequence. Vertical. Three selected cells. Step - year.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
		// Case #101: Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - year. Asc month seq, but time sequence is reverse.
		testData = [
			['01/01/2000 12:00', '01/01/2002 5:00']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = ['', '', '', '', '', ''];
		expectedData = ['37986.916666666664', '38717.62499999999', '39447.33333333332', '40178.04166666665', '40907.74999999998', '41638.45833333331'];
		getAutofillCase([0, 1, 0, 0], [2, 7, 0, 0], 3, 'Date & Time format. Asc sequence. Horizontal. Two selected cells. Step - year. Asc month seq, but time sequence is reverse.', expectedData);
		// Case #102: Date & Time format. Reverse sequence. Horizontal. Two selected cells. Step - year. Asc month seq, but time sequence is reverse.
		range = ws.getRange4(0, 6);
		range.fillData(testData);
		undoData = ['', '', '', '', '37257.208333333336', '36526.5'];
		expectedData = ['35796.79166666666', '35066.083333333314', '34336.37499999997', '33605.66666666663', '32875.958333333285', '32145.24999999994'];
		getAutofillCase([6, 7, 0, 0], [5, 0, 0, 0], 1, 'Date & Time format. Reverse sequence. Horizontal. Two selected cells. Step - year. Asc month seq, but time sequence is reverse.', expectedData);
		ws.getRange2('A1:Z20').cleanAll();
		// Case #103: Date format. Asc sequence. Vertical. Two selected cells. Step - month. Next month is February, and the day is more than the last day of the month.
		testData = [
			['12/30/2002'],
			['01/30/2003']
		];
		range = ws.getRange4(0, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], [''], ['']];
		expectedData = [['37680'], ['37710'], ['37741'], ['37771'], ['37802'], ['37832']];
		getAutofillCase([0, 0, 0, 1], [0, 0, 2, 7], 3, 'Date format. Asc sequence. Vertical. Two selected cells. Step - month. Next month is February, and the day is more than the last day of the month.', expectedData);
		// Case #104: Date format. Reverse sequence. Vertical. Two selected cells. Step - month. Next month is February, and the day is more than the last day of the month.
		testData = [
			['03/30/2003'],
			['04/30/2003']
		]
		range = ws.getRange4(6, 0);
		range.fillData(testData);
		undoData = [[''], [''], [''], [''], ['37651'], ['37620']];
		expectedData = [['37680'], ['37651'], ['37620'], ['37590'], ['37559'], ['37529']];
		getAutofillCase([0, 0, 6, 7], [0, 0, 5, 0], 1, 'Date format. Reverse sequence. Vertical. Two selected cells. Step - month. Next month is February, and the day is more than the last day of the month.', expectedData);
		ws.getRange2('A1:A20').cleanAll();
	});

	QUnit.test('Cells merge test', function (assert) {
		
		ws.getRange2("A1:Z100").cleanAll();
		ws.getRange2("A10").setValue("1");
		ws.getRange2("A11").setValue("2");
		ws.getRange2("B10").setValue("3");
		ws.getRange2("B11").setValue("4");

		// If the range is not merged, null is returned, otherwise a Range.bbox object
		assert.ok(ws.getRange2("A10:B11").hasMerged() === null, 'Range A10:B11 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 before merge in A10:B11');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "2", 'Value in A11 before merge in A10:B11');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "3", 'Value in B10 before merge in A10:B11');
		assert.strictEqual(ws.getRange2("B11").getValueWithoutFormat(), "4", 'Value in B11 before merge in A10:B11');

		// A10:B11 merge & center
		ws.getRange2("A10:B11").merge(Asc.c_oAscMergeOptions.MergeCenter);

		assert.ok(ws.getRange2("A10:B11").hasMerged(), 'Range A10:B11 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 after merge in A10:B11');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "", 'Value in A11 after merge in A10:B11');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after merge in A10:B11');
		assert.strictEqual(ws.getRange2("B11").getValueWithoutFormat(), "", 'Value in B11 after merge in A10:B11');

		// A10:B11 unmerge
		ws.getRange2("A10:B11").merge(Asc.c_oAscMergeOptions.MergeCenter);

		ws.getRange2("A10:B11").cleanAll();
		ws.getRange2("A10").setValue("1");
		ws.getRange2("A11").setValue("2");
		ws.getRange2("B10").setValue("3");
		ws.getRange2("B11").setValue("4");

		assert.ok(ws.getRange2("A9:B11").hasMerged() === null, 'Range A9:B11 is not merged');
		ws.getRange2("A9:B11").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A9:B11").hasMerged(), 'Range A9:B11 is merged');
		assert.strictEqual(ws.getRange2("A9").getValueWithoutFormat(), "1", 'Value in A9 after merge in A9:B11');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "", 'Value in A10 after merge in A9:B11');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "", 'Value in A11 after merge in A9:B11');
		assert.strictEqual(ws.getRange2("B9").getValueWithoutFormat(), "", 'Value in B9 after merge in A9:B11');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after merge in A9:B11');
		assert.strictEqual(ws.getRange2("B11").getValueWithoutFormat(), "", 'Value in B11 after merge in A9:B11');

		// A9:B11 unmerge
		ws.getRange2("A9:B11").merge(Asc.c_oAscMergeOptions.MergeCenter);
		ws.getRange2("A9:B11").cleanAll();

		let bbox = ws.getRange2("A11").bbox;
		ws.getRange2("A11").setValue("={2,4}", undefined, undefined, bbox);
		// cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		assert.ok(ws.getRange2("A10:A11").hasMerged() === null, 'Range A10:A11 is not merged');
		ws.getRange2("A10:A11").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A10:A11").hasMerged(), 'Range A10:A11 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "2", 'Value in A10 after merge in A10:A11 with one cell array');

		// A10:A11 unmerge
		ws.getRange2("A10:A11").merge(Asc.c_oAscMergeOptions.MergeCenter);
		ws.getRange2("A10:A12").cleanAll();

		bbox = ws.getRange2("A11:A12").bbox;
		ws.getRange2("A11").setValue("={2,4}", undefined, undefined, bbox);
		assert.ok(ws.getRange2("A10:A12").hasMerged() === null, 'Range A10:A12 is not merged');
		ws.getRange2("A10:A12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A10:A12").hasMerged(), 'Range A10:A12 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "", 'Value in A10 after merge in A10:A12 with two cells array');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "", 'Value in A11 after merge in A10:A12 with two cells array');
		assert.strictEqual(ws.getRange2("A12").getValueWithoutFormat(), "", 'Value in A12 after merge in A10:A12 with two cells array');

		// A10:A12 unmerge
		ws.getRange2("A10:A12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		ws.getRange2("A10:A12").cleanAll();

		ws.getRange2("A10").setValue("1");

		bbox = ws.getRange2("A11").bbox;
		ws.getRange2("A11").setValue("={2,4}", undefined, undefined, bbox);
		assert.ok(ws.getRange2("A10:A12").hasMerged() === null, 'Range A10:A12 is not merged');
		ws.getRange2("A10:A12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A10:A12").hasMerged(), 'Range A10:A12 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 after merge in A10:A12 with one cell value before one cell array');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "", 'Value in A11 after merge in A10:A12 with one cell value before one cell array');
		assert.strictEqual(ws.getRange2("A12").getValueWithoutFormat(), "", 'Value in A12 after merge in A10:A12 with one cell value before one cell array');

		// A10:A12 unmerge
		ws.getRange2("A10:A12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		ws.getRange2("A10:B12").cleanAll();

		ws.getRange2("A10").setValue("1");
		ws.getRange2("B10").setValue("1");

		bbox = ws.getRange2("A11:B11").bbox;
		ws.getRange2("A11:B11").setValue("={2,4}", undefined, undefined, bbox);
		assert.ok(ws.getRange2("A10:B12").hasMerged() === null, 'Range A10:B12 is not merged');
		ws.getRange2("A10:B12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		assert.ok(ws.getRange2("A10:B12").hasMerged(), 'Range A10:B12 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 after merge in A10:B12 with two cells value before two cells array');
		assert.strictEqual(ws.getRange2("A11").getValueWithoutFormat(), "", 'Value in A11 after merge in A10:B12 with two cells value before two cells array');
		assert.strictEqual(ws.getRange2("A12").getValueWithoutFormat(), "", 'Value in A12 after merge in A10:B12 with two cells value before two cells array');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after merge in A10:B12 with two cells value before two cells array');
		assert.strictEqual(ws.getRange2("B11").getValueWithoutFormat(), "", 'Value in B11 after merge in A10:B12 with two cells value before two cells array');
		assert.strictEqual(ws.getRange2("B12").getValueWithoutFormat(), "", 'Value in B12 after merge in A10:B12 with two cells value before two cells array');

		// A10:B12 unmerge
		ws.getRange2("A10:B12").merge(Asc.c_oAscMergeOptions.MergeCenter);
		ws.getRange2("A10:B12").cleanAll();

		/* merge across tests */
		let mergeRes;
		ws.getRange2("A1:Z100").cleanAll();
		ws.getRange2("A10").setValue("1");
		ws.getRange2("B10").setValue("2");
		ws.getRange2("C10").setValue("3");

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "2", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 before across merge in A10:C10');

		// A10:C10 merge across
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "1", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		ws.getRange2("A10").setValue("");
		ws.getRange2("B10").setValue("2");
		ws.getRange2("C10").setValue("3");

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "2", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 before across merge in A10:C10');

		// A10:C10 merge across
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "2", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		bbox = ws.getRange2("A10").bbox;
		ws.getRange2("A10").setValue("={4,6}", undefined, undefined, bbox);
		ws.getRange2("B10").setValue("2");
		ws.getRange2("C10").setValue("3");

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "4", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "2", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 before across merge in A10:C10');

		// A10:C10 merge across
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "4", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		bbox = ws.getRange2("B10").bbox;
		ws.getRange2("A10").setValue("2");
		ws.getRange2("B10").setValue("={4,6}", undefined, undefined, bbox);
		ws.getRange2("C10").setValue("3");

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "2", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "4", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 before across merge in A10:C10');

		// A10:C10 merge across
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "2", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		bbox = ws.getRange2("A10:B10").bbox;
		ws.getRange2("A10:B10").setValue("={4,6}", undefined, undefined, bbox);
		ws.getRange2("C10").setValue("3");

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "4", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "6", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 before across merge in A10:C10');

		// A10:C10 merge across with expected error
		mergeRes = ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.strictEqual(mergeRes && mergeRes.errorType, Asc.c_oAscError.ID.CannotChangeFormulaArray);
		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "4", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "6", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "3", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		bbox = ws.getRange2("B10:C10").bbox;
		ws.getRange2("A10").setValue("3");
		ws.getRange2("B10:C10").setValue("={4,6}", undefined, undefined, bbox);

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "3", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "4", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "6", 'Value in C10 before across merge in A10:C10');
		
		// A10:C10 merge across
		mergeRes = ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.strictEqual(mergeRes, undefined);
		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "3", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();


		bbox = ws.getRange2("B10:C10").bbox;
		ws.getRange2("A10").setValue("");
		ws.getRange2("B10:C10").setValue("={4,6}", undefined, undefined, bbox);

		assert.ok(ws.getRange2("A10:C10").hasMerged() === null, 'Range A10:C10 is not merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "", 'Value in A10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "4", 'Value in B10 before across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "6", 'Value in C10 before across merge in A10:C10');
		
		// A10:C10 merge across
		mergeRes = ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);

		assert.strictEqual(mergeRes, undefined);
		assert.ok(ws.getRange2("A10:C10").hasMerged(), 'Range A10:C10 is merged');
		assert.strictEqual(ws.getRange2("A10").getValueWithoutFormat(), "", 'Value in A10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("B10").getValueWithoutFormat(), "", 'Value in B10 after across merge in A10:C10');
		assert.strictEqual(ws.getRange2("C10").getValueWithoutFormat(), "", 'Value in C10 after across merge in A10:C10');

		// A10:C10 unmerge
		ws.getRange2("A10:C10").merge(Asc.c_oAscMergeOptions.MergeAcross);
		ws.getRange2("A10:C10").cleanAll();

	});

	QUnit.test('All selection test', function (assert) {

		ws.getRange2("A1:Z100").cleanAll();

		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("B1").setValue("3");
		ws.getRange2("B2").setValue("4");
		ws.getRange2("C2").setValue("5");

		let fillRange = new Asc.Range(0, 0, 0, 0);
		wsView.setSelection(fillRange);
		assert.strictEqual(ws.selectionRange.getLast().getName(), "A1");

		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getName(), "A1:C2");
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getType(), Asc.c_oAscSelectionType.RangeMax);

		fillRange = new Asc.Range(10, 10, 10, 10);
		wsView.setSelection(fillRange);
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getType(), Asc.c_oAscSelectionType.RangeMax);


		let tableOptions = new AscCommonExcel.AddFormatTableOptions();
		fillRange = new Asc.Range(0, 0, 2, 1);
		tableOptions.range = fillRange.getName();
		ws.autoFilters.addAutoFilter("style", fillRange, tableOptions);
		assert.strictEqual(ws.TableParts.length, 1);
		assert.strictEqual(ws.TableParts[0].Ref.getName(), "A1:C3");


		fillRange = new Asc.Range(1, 1, 1, 1);
		wsView.setSelection(fillRange);
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getName(), "A2:C3");
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getName(), "A1:C3");
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getType(), Asc.c_oAscSelectionType.RangeMax);

		fillRange = new Asc.Range(0, 0, 0, 0);
		wsView.setSelection(fillRange);
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getName(), "A1:C3");
		api.wb._onSelectAllByRange();
		assert.strictEqual(ws.selectionRange.getLast().getType(), Asc.c_oAscSelectionType.RangeMax);
	});

		QUnit.module("Sheet structure");
});
