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
		clearData(0, 0, 0, 12)

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
		clearData(0, 0, 0, 4)

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
		clearData(0, 0, 0, 4)

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


	QUnit.module("Sheet structure");
});
