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
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (oFromRange, subType, callback) {
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
	Asc.ReadDefTableStyles = function () {
	};

	let api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
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
	let wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	wb.handlers.add("asc_onConfirmAction", function (test1, callback) {
		callback(true);
	});
	let wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	// Initialize global variables and functions for tests
	let ws = api.wbModel.aWorksheets[0];
	const CSerial = window['AscCommonExcel'].CSerial;
	const parserFormula = AscCommonExcel.parserFormula;
	let oSeriesInType = Asc.c_oAscSeriesInType;
	let oSeriesType = Asc.c_oAscSeriesType;
	let oSeriesDateUnitType = Asc.c_oAscDateUnitType;
	let oRightClickOptions = Asc.c_oAscFillType;
	let settings, cSerial, autofillRange, oFromRange, expectedData, nType;
	
	const getRange = function(c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};
	const clearData = function(c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0, 0, 0, 0));
		ws.mergeManager.remove(getRange(c1, r1, c2, r2));
		wsView.activeFillHandle = null;
		wsView.fillHandleDirection = null;
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	};
	const autofillData = function(assert, fromRangeTo, expectedData, description) {
		for (let i = fromRangeTo.r1; i <= fromRangeTo.r2; i++) {
			for (let j = fromRangeTo.c1; j <= fromRangeTo.c2; j++) {
				let fromRangeToVal = ws.getCell3(i, j);
				let dataVal = expectedData[i - fromRangeTo.r1][j - fromRangeTo.c1];
				assert.strictEqual(fromRangeToVal.getValue(), dataVal, `${description} Cell: ${fromRangeToVal.getName()}, Value: ${dataVal}`);
			}
		}
	};
	const getFilledData = function(c1, r1, c2, r2, testData, oStartRange) {
		let [row, col] = oStartRange;
		let oFromRange = ws.getRange4(row, col);
		oFromRange.fillData(testData);
		oFromRange.worksheet.selectionRange.ranges = [getRange(c1, r1, c2, r2)];
		oFromRange.bbox = getRange(c1, r1, c2, r2);

		return oFromRange;
	};
	const initMergedCell = function(oRangeModel, oRange) {
		let oFromWs = oRangeModel.worksheet;
		oFromWs.mergeManager.add(oRange, 1);

		return oFromWs;
	};
	const checkUndoRedo = function(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	};

	QUnit.module('Series');
	QUnit.test('Autofill linear progression - one filled row/column', function (assert) {
		// Fill data
		let testData = [
		  ['1']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings = {
			'stepValue': 2,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.linear,
			'stopValue': null,
			'trend': false
		}
		// Run AutoFill -> Series
		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['3', '5', '7', '9', '11']], 'Autofill one Row');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange
		// Fill data
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['3'], ['5'], ['7'], ['9'], ['11']], 'Autofill one Column');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange with stopValue
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.stopValue = 7;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['3', '5', '7', '', '']], 'Autofill one Row with stopValue = 7');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange with stopValue
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.stopValue = 7;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['3'], ['5'], ['7'], [''], ['']], 'Autofill one Column with stopValue = 7');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange with trend step
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.trend = true;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['2', '3', '4', '5', '6']], 'Autofill one Row with trend step');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange with trend step
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['2'], ['3'], ['4'], ['5'], ['6']], 'Autofill one Column with trend step');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange with stop value = -10 and step value = -1. Bug #65704
		oFromRange = getFilledData(0, 0, 12, 0, testData, [0, 0]);
		let oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 12, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '', '', '', '', '', '', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '0', '-1', '-2', '-3', '-4', '-5', '-6', '-7', '-8', '-9', '-10', '']], _desc);
		}, "Autofill one Row with stop value = -10 and step value = -1. Bug #65704");
		clearData(0, 0, 12, 0);
		// Select vertical oFromRange with stop value = -10 and step value = -1. Bug #65704
		oFromRange = getFilledData(0, 0, 0, 12, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 12);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['0'], ['-1'], ['-2'], ['-3'], ['-4'], ['-5'], ['-6'], ['-7'], ['-8'], ['-9'], ['-10'], ['']], _desc);
		}, "Autofill one Column with stop value = -10 and step value = -1. Bug #65704");
		clearData(0, 0, 0, 12);
		// Select horizontal oFromRange with stop value = 10 and step value = -1. Bug #65705
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(10);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);
		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row with stop value = 10 and step value = -1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with stop value = 10 and step value = -1. Bug #65705
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(10);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column with stop value = 10 and step value = -1. Bug #65705");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange with stop value = -10 and step value = 1. Bug #65705
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1);
		oSeriesSettings.asc_setStopValue(-10);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row with stop value = -10 and step value = 1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select horizontal oFromRange with stop value = 1 and step value = 1. Bug #65705
		testData = [
			['1', '2']
		];
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStopValue(1);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		},function (_desc) {
			autofillData(assert, autofillRange, [['1', '2', '']], _desc);
		}, "Autofill one Row with stop value = 1 and step value = 1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with stop value = 1 and step value = 1. Bug #65705
		testData = [
			['1'],
			['2']
		];
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStopValue(1);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['2'], ['']], _desc);
		}, "Autofill one Column with stop value = 1 and step value = 1. Bug #65705");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange with stop value = 0 and step value = 1. Bug #65705
		testData = [
			['1', '2']
		];
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStopValue(0);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '2', '']], _desc);
		}, "Autofill one Row with stop value = 0 and step value = 1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with step value = -0.2. Bug #65871
		testData = [
			['1']
		];
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-0.2);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 5);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['0.8'], ['0.6'], ['0.4'], ['0.2'], ['0']], _desc);
		}, "Autofill one Column with step value = -0.2. Bug #65871");
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange with activeFillHandle. Step value = -1, stop value = 1 . Bug #65895
		testData = [
			['3']
		];
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 3, 0);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(1);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 3, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['3', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['3', '2', '1', '']], _desc);
		}, "Autofill one Row with activeFillHandle. Step value = -1, stop value = 1 . Bug #65895");
		clearData(0, 0, 3, 0);
		// Select vertical oFromRange with activeFillHandle. Step value = 1, stop value = 0 . Bug #65895
		testData = [
			['-2']
		];
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 3);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1);
		oSeriesSettings.asc_setStopValue(0);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['-2'], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['-2'], ['-1'], ['0'], ['']], _desc);
		}, "Autofill one Column with activeFillHandle. Step value = 1, stop value = 0 . Bug #65895");
		clearData(0, 0, 0, 3);
		// Select horizontal oFromRange. Step value = 1, stop value = -2 . Bug #65895
		testData = [
			['-4']
		];
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1);
		oSeriesSettings.asc_setStopValue(-2);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 3, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['-4', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['-4', '-3', '-2', '']], _desc);
		}, "Autofill one Row. Step value = 1, stop value = -2 . Bug #65895");
		clearData(0, 0, 3, 0);
		// Select vertical oFromRange with active fill handle, reverse direction. Bug #65877
		testData = [
			['1']
		];
		oFromRange = getFilledData(3, 3, 3, 3, testData, [3, 3]);
		wsView.activeFillHandle = getRange(3, 3, 3, 0);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = api.asc_GetSeriesSettings();
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(3, 0, 3, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['1']], _desc);
		}, "Autofill one Column with active fill handle, reverse direction. Bug #65877");
		clearData(3, 0, 3, 3);
		// Select horizontal oFromRange with active fill handle, reverse direction. Bug #65877
		oFromRange = getFilledData(3, 3, 3, 3, testData, [3, 3]);
		wsView.activeFillHandle = getRange(3, 3, 0, 3);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = api.asc_GetSeriesSettings();
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 3, 3, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['', '', '', '1']], _desc);
		}, "Autofill one Row with active fill handle, reverse direction. Bug #65877");
		clearData(0, 3, 3, 3);
	});
	QUnit.test('Autofill growth progression - one filled row/column', function (assert) {
		const testData = [
			['1']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings = {
			'stepValue': 2,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.growth,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['2', '4', '8', '16', '32']], 'Autofill one Row');
		clearData(0, 0, 5, 5);
		// Select vertical oFromRange
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		//oFromRange = ws.getRange4(0,0);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['2'], ['4'], ['8'], ['16'], ['32']], 'Autofill one Column');
		clearData(0, 0, 5, 5);
		// Select vertical oFromRange with stopValue
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.stopValue = 10;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['2'], ['4'], ['8'], [''], ['']], 'Autofill one Column with stopValue = 10');
		clearData(0, 0, 5, 5);
		// Select horizontal oFromRange with stopValue
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.stopValue = 10;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['2', '4', '8', '', '']], 'Autofill one Row with stopValue = 10');
		clearData(0, 0, 5, 0);
		// Select horizontal oFromRange with trend step
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.trend = true;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['1', '1', '1', '1', '1']], 'Autofill one Row with trend step');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange with trend step
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['1'], ['1'], ['1'], ['1'], ['1']], 'Autofill one Column with trend step');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange with step = -1 and stopValue = -10
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		let oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row with stop value = -10 and step value = -1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with stop value = 10 and step value = -2. Bug #65705
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-2);
		oSeriesSettings.asc_setStopValue(10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column with stop value = 10 and step value = -2. Bug #65705");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange with stop value = -10 and step value = 1. Bug #65705
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row with stop value = -10 and step value = 1. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with stop value = 10 and step value = 1. Bug #65705
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1);
		oSeriesSettings.asc_setStopValue(10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column with stop value = 10 and step value = 1. Bug #65705");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange with stop value = -10 and step value = 2. Bug #65705
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(2);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row with stop value = -10 and step value = 2. Bug #65705");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with stop value = 10 and step value = 0. Bug #65705
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0);
		oSeriesSettings.asc_setStopValue(10);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column with stop value = 10 and step value = 0. Bug #65705");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange with stop value = 0.25 and step value = 0.5. Bug #65897
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0.5);
		oSeriesSettings.asc_setStopValue(0.25);
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 3, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '0.5', '0.25', '']], _desc);
		}, "Autofill one Row with stop value = 0.25 and step value = 0.5. Bug #65897");
		clearData(0, 0, 2, 0);
		// Select vertical oFromRange with active fill handle, reverse direction. Bug #65877
		oFromRange = getFilledData(3, 3, 3, 3, testData, [3, 3]);
		wsView.activeFillHandle = getRange(3, 3, 3, 0);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(3, 0, 3, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['1']], _desc);
		}, "Autofill one Column with active fill handle, reverse direction. Bug #65877");
		clearData(3, 0, 3, 3);
		// Select horizontal oFromRange with active fill handle, reverse direction. Bug #65877
		oFromRange = getFilledData(3, 3, 3, 3, testData, [3, 3]);
		wsView.activeFillHandle = getRange(3, 3, 0, 3);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setType(oSeriesType.growth);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 3, 3, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['', '', '', '1']], _desc);
		}, "Autofill one Row with active fill handle, reverse direction. Bug #65877");
		clearData(0, 3, 3, 3);
	});
	QUnit.test('Autofill default mode', function (assert) {
		let testData = [
			['1', '2']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings = {
			'stepValue': null,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.autoFill,
			'stopValue': null,
			'trend': false
		}

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(2, 0, 5, 0);
		autofillData(assert, autofillRange, [['3', '4', '5', '6']], 'Autofill one Row');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange
		testData = [
			['1'],
			['2']
		];
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 2, 0, 5);
		autofillData(assert, autofillRange, [['3'], ['4'], ['5'], ['6']], 'Autofill one Column');
		clearData(0, 0, 0, 5);
		// Select vertical oFromRange, Series in: Rows. Bug #65724
		testData = [
			['1']
		];
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		let oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setSeriesIn(oSeriesInType.rows);
		oSeriesSettings.asc_setType(oSeriesType.autoFill);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column. Series in: Rows. Type: AutoFill. Bug #65724");
		clearData(0, 0, 0, 2);
		// Select horizontal oFromRange, Series in: Columns. Bug #65724
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setSeriesIn(oSeriesInType.columns);
		oSeriesSettings.asc_setType(oSeriesType.autoFill);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 2, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '']], _desc);
		}, "Autofill one Row. Series in: Columns. Type: AutoFill. Bug #65724");
		clearData(0, 0, 2, 0);
		// FillHandle Vertical. Series in: Rows. Type: AutoFill. Bug #65724
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 2);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setSeriesIn(oSeriesInType.rows);
		oSeriesSettings.asc_setType(oSeriesType.autoFill);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 2);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], ['']], _desc);
		}, "Autofill one Column fillHandle. Series in: Rows. Type: AutoFill. Bug #65724");
		clearData(0, 0, 0, 2);
	});
	QUnit.test('Negative cases', function (assert) {
		const testData = [
		   [''],
		   ['Test1']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings = {
			'stepValue': 1,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.linear,
			'stopValue': null,
			'trend': false
		}

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		let autoFillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill Linear progression - one Row');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autoFillRange, [['Test1'], [''], [''], [''], ['']], 'Autofill Linear progression - one Column');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange Growth
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.rows;
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill Growth progression - one Row');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange Growth
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autoFillRange, [['Test1'], [''], [''], [''], ['']], 'Autofill Growth progression - one Column');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange autofill default mode
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.rows;
		settings.type = oSeriesType.autoFill;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill default mode - one Row');
		clearData(0, 0, 5, 0);
		// Select vertical oFromRange autofill default mode
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.columns;
		settings.type = oSeriesType.autoFill;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autoFillRange, [['Test1'], [''], ['Test2'], [''], ['Test3']], 'Autofill default mode - one Column');
		clearData(0, 0, 0, 5);
		// Select horizontal oFromRange linear progression. String test
		oFromRange = getFilledData(0, 1, 5, 1, testData, [0, 0]);
		settings.seriesIn = oSeriesInType.rows;
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autoFillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autoFillRange, [['', '', '', '', '']], 'Autofill linear progression - one Row');
	});
	QUnit.test('Autofill horizontal progression - multiple filled cells', function (assert) {
		let testData = [
			['1'],
			['2'],
			['3'],
			['4'],
			['5'],
			['6']
		];
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings = {
			'stepValue': 1,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.linear,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['2', '3', '4', '5', '6'],
			['3', '4', '5', '6', '7'],
			['4', '5', '6', '7', '8'],
			['5', '6', '7', '8', '9'],
			['6', '7', '8', '9', '10'],
			['7', '8', '9', '10', '11']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression');
		clearData(0, 0, 5, 5);
		// Growth progression
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 2;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['2', '4', '8', '16', '32'],
			['4', '8', '16', '32', '64'],
			['6', '12', '24', '48', '96'],
			['8', '16', '32', '64', '128'],
			['10', '20', '40', '80', '160'],
			['12', '24', '48', '96', '192']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression');
		clearData(0, 0, 5, 5);
		// Linear progression with stop value
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.stopValue = 7;
		settings.stepValue = 2;
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['3', '5', '7', '', ''],
			['4', '6', '', '', ''],
			['5', '7', '', '', ''],
			['6', '', '', '', ''],
			['7', '', '', '', ''],
			['', '', '', '', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression with stop value');
		clearData(0, 0, 5, 5);
		// Growth progression with stop value
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.stopValue = 16;
		settings.type = oSeriesType.growth;
		settings.stepValue = 2;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['2', '4', '8', '16', ''],
			['4', '8', '16', '', ''],
			['6', '12', '', '', ''],
			['8', '16', '', '', ''],
			['10', '', '', '', ''],
			['12', '', '', '', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with stop value');
		clearData(0, 0, 5, 5);
		// Linear progression with trend
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.trend = true;
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['2', '3', '4', '5', '6'],
			['3', '4', '5', '6', '7'],
			['4', '5', '6', '7', '8'],
			['5', '6', '7', '8', '9'],
			['6', '7', '8', '9', '10'],
			['7', '8', '9', '10', '11']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression with trend');
		clearData(0, 0, 5, 5);
		// Growth progression with trend
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.trend = true;
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['1', '1', '1', '1', '1'],
			['2', '2', '2', '2', '2'],
			['2.9999999999999996', '2.9999999999999996', '2.9999999999999996', '2.9999999999999996', '2.9999999999999996'],
			['4', '4', '4', '4', '4'],
			['5', '5', '5', '5', '5'],
			['6', '6', '6', '6', '6']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with trend');
		clearData(0, 0, 5, 5);
		// Growth progression all cells filled. Step = 0.5
		testData = [
			['1', '2', '3', '4', '5', '6'],
			['2', '3', '4', '5', '6', '7'],
			['3', '4', '5', '6', '7', '8'],
			['4', '5', '6', '7', '8', '9'],
			['5', '6', '7', '8', '9', '10'],
			['6', '7', '8', '9', '10', '11']
		];
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 0.5;
		settings.trend = false;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 5);
		expectedData = [
			['0.5', '0.25', '0.125', '0.0625', '0.03125'],
			['1', '0.5', '0.25', '0.125', '0.0625'],
			['1.5', '0.75', '0.375', '0.1875', '0.09375'],
			['2', '1', '0.5', '0.25', '0.125'],
			['2.5', '1.25', '0.625', '0.3125', '0.15625'],
			['3', '1.5', '0.75', '0.375', '0.1875']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression all cells filled. Step = 0.5');
		clearData(0, 0, 5, 5);
		// Growth progression. Step = 0.5. With indentation row and column
		testData = [
			['1'],
			['2'],
			['3'],
			['4'],
			['5'],
			['6']
		];
		oFromRange = getFilledData(2, 1, 7, 6, testData, [1, 2]);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(3, 1, 7, 6);
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression. Step = 0.5. With indentation row and column');
		clearData(0, 0, 7, 6);
	});
	QUnit.test('Autofill vertical progression - multiple filled cells', function (assert) {
		let testData = [
			['1', '2', '3', '4', '5', '6']
		];
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings = {
			'stepValue': 1,
			'seriesIn': oSeriesInType.columns,
			'type': oSeriesType.linear,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['2', '3', '4', '5', '6', '7'],
			['3', '4', '5', '6', '7', '8'],
			['4', '5', '6', '7', '8', '9'],
			['5', '6', '7', '8', '9', '10'],
			['6', '7', '8', '9', '10', '11']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression');
		clearData(0, 0, 5, 5);
		// Growth progression
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 2;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['2', '4', '6', '8', '10', '12'],
			['4', '8', '12', '16', '20', '24'],
			['8', '16', '24', '32', '40', '48'],
			['16', '32', '48', '64', '80', '96'],
			['32', '64', '96', '128', '160', '192']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression');
		clearData(0, 0, 5, 5);
		// Linear progression with stop value
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.stopValue = 10;
		settings.stepValue = 2;
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['3', '4', '5', '6', '7', '8'],
			['5', '6', '7', '8', '9', '10'],
			['7', '8', '9', '10', '', ''],
			['9', '10', '', '', '', ''],
			['', '', '', '', '', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression with stop value');
		clearData(0, 0, 5, 5);
		// Growth progression with stop value
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.stopValue = 32;
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['2', '4', '6', '8', '10', '12'],
			['4', '8', '12', '16', '20', '24'],
			['8', '16', '24', '32', '', ''],
			['16', '32', '', '', '', ''],
			['32', '', '', '', '', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with stop value');
		clearData(0, 0, 5, 5);
		// Linear progression with trend
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.trend = true;
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['2', '3', '4', '5', '6', '7'],
			['3', '4', '5', '6', '7', '8'],
			['4', '5', '6', '7', '8', '9'],
			['5', '6', '7', '8', '9', '10'],
			['6', '7', '8', '9', '10', '11']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression with trend');
		clearData(0, 0, 5, 5);
		// Growth progression with trend
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.trend = true;
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['1', '2', '2.9999999999999996', '4', '5', '6'],
			['1', '2', '2.9999999999999996', '4', '5', '6'],
			['1', '2', '2.9999999999999996', '4', '5', '6'],
			['1', '2', '2.9999999999999996', '4', '5', '6'],
			['1', '2', '2.9999999999999996', '4', '5', '6']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with trend');
		clearData(0, 0, 5, 5);
		// Linear progression all cells filled. Step = -1.
		testData = [
			['1', '2', '3', '4', '5', '6'],
			['1', '2', '3', '4', '5', '6'],
			['1', '2', '3', '4', '5', '6'],
			['1', '2', '3', '4', '5', '6'],
			['1', '2', '3', '4', '5', '6'],
			['1', '2', '3', '4', '5', '6']
		];
		oFromRange = getFilledData(0, 0, 5, 5, testData, [0, 0]);
		settings.stepValue = -1;
		settings.type = oSeriesType.linear;
		settings.trend = false;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 5, 5);
		expectedData = [
			['0', '1', '2', '3', '4', '5'],
			['-1', '0', '1', '2', '3', '4'],
			['-2', '-1', '0', '1', '2', '3'],
			['-3', '-2', '-1', '0', '1', '2'],
			['-4', '-3', '-2', '-1', '0', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression all cells filled. Step = -1.');
		clearData(0, 0, 5, 5);
		// Linear. Step -1. With indentation row and column
		testData = [
			['1', '2', '3', '4', '5', '6']
		];
		oFromRange = getFilledData(1, 1, 6, 6, testData, [1, 1]);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 2, 6, 6);
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression. Step = -1. With indentation row and column');
		clearData(0, 0, 6, 6);
	});
	QUnit.test('Autofill Date type - one filled row/column', function (assert) {
		let testData = [
			['09/04/2023']
		];
		// Horizontal dateUnit - Day
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings = {
			'stepValue': 1,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.date,
			'dateUnit': oSeriesDateUnitType.day,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['45174', '45175', '45176', '45177', '45178']], 'Autofill Row. Date progression - Day');
		clearData(0, 0, 5, 0);
		// Horizontal dateUnit - Weekday
		oFromRange = getFilledData(0, 0, 7, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 7, 0);
		autofillData(assert, autofillRange, [['45174', '45175', '45176', '45177', '45180', '45181', '45182']], 'Autofill Row. Date progression - Weekday');
		clearData(0, 0, 7, 0);
		// Horizontal dateUnit - Month, Step - 2
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;
		settings.stepValue = 2;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['45234', '45295', '45355', '45416', '45477']], 'Autofill Row. Date progression - Month, Step - 2');
		clearData(0, 0, 5, 0);
		// Horizontal dateUnit - Year, Step - 1
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;
		settings.stepValue = 1;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['45539', '45904', '46269', '46634', '47000']], 'Autofill Row. Date progression - Year, Step - 1');
		clearData(0, 0, 5, 0);
		// Vertical dateUnit - Day
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.day;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], ['45177'], ['45178']], 'Autofill Column. Date progression - Day');
		clearData(0, 0, 0, 5);
		// Vertical dateUnit - Weekday
		oFromRange = getFilledData(0, 0, 0, 7, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 7);
		autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], ['45177'], ['45180'], ['45181'], ['45182']], 'Autofill Column. Date progression - Weekday');
		clearData(0, 0, 0, 7);
		// Vertical dateUnit - Month, Step - 2
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;
		settings.stepValue = 2;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['45234'], ['45295'], ['45355'], ['45416'], ['45477']], 'Autofill Column. Date progression - Month, Step - 2');
		clearData(0, 0, 0, 5);
		// Vertical dateUnit - Year, Step - 1
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;
		settings.stepValue = 1;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 5);
		autofillData(assert, autofillRange, [['45539'], ['45904'], ['46269'], ['46634'], ['47000']], 'Autofill Column. Date progression - Year, Step - 1');
		clearData(0, 0, 0, 5);
		// Horizontal dateUnit - Day, Stop value - 45176
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.day;
		settings.stopValue = 45176;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 0);
		autofillData(assert, autofillRange, [['45174', '45175', '45176', '', '']], 'Autofill Row. Date progression - Day, Stop value - 45176');
		clearData(0, 0, 5, 0);
		// Vertical dateUnit - Day, Stop value - 45176. With indentation row and column
		oFromRange = getFilledData(1, 1, 1, 6, testData, [1, 1]);
		settings.dateUnit = oSeriesDateUnitType.day;
		settings.stopValue = 45176;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 2, 1, 6);
		autofillData(assert, autofillRange, [['45174'], ['45175'], ['45176'], [''], ['']], 'Autofill Column. Date progression - Day, Stop value - 45176. With indentation row and column');
		// Case 01/01/1900 - 01/03/1900. Vertical dateUnit - Day, Step - 2. Bug #65559
		testData = [
			['01/01/1900']
		];
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.day;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 3);
		autofillData(assert, autofillRange, [['3'], ['5'], ['7']], 'Autofill Column. Date progression - Day, Step - 2. Bug #65559');
		clearData(0, 0, 0, 3);
		// Case 01/01/1900 - 01/03/1900. Horizontal dateUnit - Day, Step - 2. Bug #65559
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.day;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 3, 0);
		autofillData(assert, autofillRange, [['3', '5', '7']], 'Autofill Row. Date progression - Day, Step - 2. Bug #65559');
		clearData(0, 0, 3, 0);
		// Case 01/01/1900 - 01/03/1900. Vertical dateUnit - Weekday, Step - 2. Bug #65559
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 3);
		autofillData(assert, autofillRange, [['3'], ['5'], ['9']], 'Autofill Column. Date progression - Weekday, Step - 2. Bug #65559');
		clearData(0, 0, 0, 3);
		// Case 01/01/1900 - 01/03/1900. Horizontal dateUnit - Weekday, Step - 2. Bug #65559
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 3, 0);
		autofillData(assert, autofillRange, [['3', '5', '9']], 'Autofill Row. Date progression - Weekday, Step - 2. Bug #65559');
		clearData(0, 0, 3, 0);
		// Case 01/01/1900 - 01/03/1900. Vertical dateUnit - Month, Step - 2. Bug #65559
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 3);
		autofillData(assert, autofillRange, [['61'], ['122'], ['183']], 'Autofill Column. Date progression - Month, Step - 2. Bug #65559');
		clearData(0, 0, 0, 3);
		// Case 01/01/1900 - 01/03/1900. Horizontal dateUnit - Month, Step - 12. Bug #65559
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;
		settings.stepValue = 12;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 3, 0);
		autofillData(assert, autofillRange, [['367', '732', '1097']], 'Autofill Row. Date progression - Month, Step - 12. Bug #65559');
		clearData(0, 0, 3, 0);
		// Case 01/01/1900 - 01/03/1900. Vertical dateUnit - Year, Step - 2. Bug #65559
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;
		settings.stepValue = 2;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 0, 3);
		autofillData(assert, autofillRange, [['732'], ['1462'], ['2193']], 'Autofill Column. Date progression - Year, Step - 2. Bug #65559');
		clearData(0, 0, 0, 3);
		// Case 01/01/1900 - 01/03/1900. Horizontal dateUnit - Year, Step - 1. Bug #65559
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;
		settings.stepValue = 1;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 3, 0);
		autofillData(assert, autofillRange, [['367', '732', '1097']], 'Autofill Row. Date progression - Year, Step - 1. Bug #65559');
		clearData(1, 0, 3, 0);
		// Horizontal dateUnit - Day. Step - 0.2. Bug #65672
		testData = [
			['01/01/2000']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		let oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0.2);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 5, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['36526', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['36526', '36526.2', "36526.4", "36526.6", "36526.8", "36527"]], _desc);
		}, "Autofill Row: Date progression - Day, Step - 0.2. Bug #65672");
		clearData(0, 0, 5, 0);
		// Vertical dateUnit - Weekday. Step - 0.2. Case: 01/01/1900 - 01/03/1900/ Bug #65672
		testData = [
			['01/01/1900']
		];
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0.2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 5);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['2'], ['2'], ['2'], ['2'], ['3']], _desc);
		}, "Autofill Column: Date progression - Weekday, Step - 0.2. Case: 01/01/1900 - 01/03/1900. Bug #65672");
		clearData(0, 0, 0, 5);
		// Vertical dateUnit - Weekday. Step - 1. Bug #65900
		oFromRange = getFilledData(0, 0, 0, 6, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 6);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['2'], ['3'], ['4'], ['5'], ['6'], ['9']], _desc);
		}, "Autofill Column: Date progression - Weekday, Step - 1. Case: 01/01/1900 - 01/03/1900. Bug #65900");
		clearData(0, 0, 0, 6);
		// Horizontal dateUnit - Month. Step - -1. Bug #65899
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.month);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 3, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '#NUM!', '#NUM!', '#NUM!']], _desc);
		}, "Autofill Row: Date progression - Month, Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 3, 0);
		// Vertical dateUnit - Year. Step - -1. Bug #65899
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.year);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['#NUM!'], ['#NUM!'], ['#NUM!']], _desc);
		}, "Autofill Column: Date progression - Year, Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 0, 3);
		// Vertical dateUnit - Weekday, Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899
		oFromRange = getFilledData(0, 0, 0, 6, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 6);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['-1'], ['-2'], ['-3'], ['-4'], ['-5'], ['-8']], _desc);
		}, "Autofill Column: Date progression - Weekday, Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 0, 6);
		// Horizontal dateUnit - Day. Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.day);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 5, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '0', '-1', '-2', '-3', '-4']], _desc);
		}, "Autofill Row: Date progression - Day, Step - -1. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 5, 0);
		// Vertical dateUnit - Month. Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.month);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], ['']], _desc);
		}, "Autofill Column: Date progression - Month, Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 0, 3);
		// Vertical dateUnit - Year. Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.year);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], ['']], _desc);
		}, "Autofill Column: Date progression - Year, Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 0, 3);
		// Vertical dateUnit - Weekday, Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-1);
		oSeriesSettings.asc_setStopValue(-10);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [[''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], ['']], _desc);
		}, "Autofill Column: Date progression - Weekday, Step - -1, StopValue - -10. Case: 01/01/1900 - 01/03/1900. Bug #65899");
		clearData(0, 0, 0, 3);
		// Horizontal dateUnit - Day. Step - -0.5. Case: 01/01/1900 - 01/03/1900. Bug #65876
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-0.5);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 5, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '0.5', '0', '0.5', '0', '0.5']], _desc);
		}, "Autofill Row: Date progression - Day, Step - -0.5. Case: 01/01/1900 - 01/03/1900. Bug #65876");
		clearData(0, 0, 5, 0);
		// Horizontal dateUnit - Month. Step - 0.2. Bug #65672
		testData = [
			['01/01/2000']
		];
		oFromRange = getFilledData(0, 0, 5, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0.2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.month);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 5, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['36526', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['36526', '36526', "36526", "36526", "36526", "36557"]], _desc);
		}, "Autofill Row: Date progression - Month, Step - 0.2. Bug #65672");
		clearData(0, 0, 5, 0);
		// Vertical dateUnit - Year. Step - 0.2. Case: 01/01/1900 - 01/03/1900.  Bug #65672.
		testData = [
			['01/01/1900']
		];
		oFromRange = getFilledData(0, 0, 0, 5, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(0.2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.year);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 5);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['1'], ['1'], ['1'], ['1'], ['367']], _desc);
		}, "Autofill Column: Date progression - Year, Step - 0.2. Case: 01/01/1900 - 01/03/1900. Bug #65672");
		clearData(0, 0, 5, 0);
		// Horizontal dateUnit - Month. Step - 30.2. Case: 01/01/1900 - 01/03/1900. Bug #65796.
		testData = [
			['1', '32', '61', '92']
		];
		oFromRange = getFilledData(0, 0, 10, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setType(oSeriesType.date);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.month);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 10, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1', '32', '61', '92', '', '', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1', '913', '1828', '2739', '3654', '4597', '5511', '6423', '7337', '8249', '9192']], _desc);
		}, "Autofill Row: Date progression - Month, Step - 30.2. Case: 01/01/1900 - 01/03/1900. Bug #65796");
		clearData(0, 0, 10, 0);
		// Vertical dateUnit - Year. Step - 365.3. Case: 01/01/1900 - 01/03/1900. Bug #65796.
		testData = [
			['1'],
			['367'],
			['732'],
			['1097']
		];
		oFromRange = getFilledData(0, 0, 0, 4, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setType(oSeriesType.date);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.year);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 4);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['1'], ['367'], ['732'], ['1097'], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['1'],['133316'], ['266629'], ['399943'], ['533622']], _desc);
		}, "Autofill Column: Date progression - Year, Step - 365.3. Case: 01/01/1900 - 01/03/1900. Bug #65796");
		clearData(0, 0, 0, 4);
		// Horizontal dateUnit - Weekday. Step - 1.2. Bug #65796.
		testData = [
			['01/14/2024']
		];
		oFromRange = getFilledData(0, 0, 9, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1.2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 9, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['45305', '', '', '', '', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['45305', '45306', '45307', '45308', '45309', '45313', '45314', '45315', '45316', '45317']], _desc);
		}, "Autofill Row: Date progression - Weekday, Step - 1.2. Bug #65796");
		clearData(0, 0, 9, 0);
		// Vertical dateUnit - Weekday. Step - 1.2. Bug #65796.
		testData = [
			['01/13/2024']
		];
		oFromRange = getFilledData(0, 0, 0, 9, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(1.2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 9);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['45304'], [''], [''], [''], [''], [''], [''], [''], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['45304'], ['45306'], ['45307'], ['45308'], ['45309'], ['45313'], ['45314'], ['45315'], ['45316'], ['45317']], _desc);
		}, "Autofill Column: Date progression - Weekday, Step - 1.2. Bug #65796");
		clearData(0, 0, 0, 9);
		// Horizontal dateUnit - Weekday. Step - -2. Bug #65731.
		testData = [
			['01/01/2023']
		];
		oFromRange = getFilledData(0, 0, 10, 0, testData, [0, 0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		oSeriesSettings.asc_setStepValue(-2);
		oSeriesSettings.asc_setDateUnit(oSeriesDateUnitType.weekday);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 0, 10, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['44927', '', '', '', '', '', '', '', '', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['44927', '44924', '44922', '44918', '44916', '44914', '44910', '44908', '44904', '44902', '44900']], _desc);
		}, "Autofill Row: Date progression - Weekday, Step - -2. Bug #65731");
		clearData(0, 0, 10, 0);
	});
	QUnit.test('Autofill Date type - Horizontal multiple cells', function (assert) {
		const testData = [
			['01/01/2023'],
			['09/04/2023'],
			['01/12/2023'],
			['12/12/2023']
		];
		// DateUnit - Day. Step - 3
		oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
		settings = {
			'stepValue': 3,
			'seriesIn': oSeriesInType.rows,
			'type': oSeriesType.date,
			'dateUnit': oSeriesDateUnitType.day,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 3);
		expectedData = [
		   ['44930', '44933', '44936', '44939', '44942', '44945'],
		   ['45176', '45179', '45182', '45185', '45188', '45191'],
		   ['44941', '44944', '44947', '44950', '44953', '44956'],
		   ['45275', '45278', '45281', '45284', '45287', '45290']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Day, Step - 3');
		clearData(0, 0, 5, 3);
		// DateUnit - Weekday
		oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 3);
		expectedData = [
		   ['44930', '44935', '44938', '44943', '44946'],
		   ['45176', '45181', '45184', '45189', '45194'],
		   ['44943', '44946', '44951', '44956', '44959'],
		   ['45275', '45280', '45285', '45288', '45293']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Weekday, Step - 3');
		clearData(0, 0, 5, 3);
		// DateUnit - Month
		oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 3);
		expectedData = [
		   ['45017', '45108', '45200', '45292', '45383'],
		   ['45264', '45355', '45447', '45539', '45630'],
		   ['45028', '45119', '45211', '45303', '45394'],
		   ['45363', '45455', '45547', '45638', '45728']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Month, Step - 3');
		clearData(0, 0, 5, 3);
		// DateUnit - Year
		oFromRange = getFilledData(0, 0, 5, 3, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 5, 3);
		expectedData = [
			['46023', '47119', '48214', '49310', '50406'],
			['46269', '47365', '48461', '49556', '50652'],
			['46034', '47130', '48225', '49321', '50417'],
			['46368', '47464', '48560', '49655', '50751']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Year, Step - 3');
		clearData(0, 0, 5, 3);
	});
	QUnit.test('Autofill Date type - Vertical multiple cells', function (assert) {
	   const testData = [
		   ['01/01/2023', '09/04/2023', '01/12/2023', '12/12/2023']
	   ];
	   // DateUnit - Day. Step - 3
		oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
		settings = {
			'stepValue': 3,
			'seriesIn': oSeriesInType.columns,
			'type': oSeriesType.date,
			'dateUnit': oSeriesDateUnitType.day,
			'stopValue': null,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 3, 5);
		expectedData = [
			['44930', '45176', '44941', '45275'],
			['44933', '45179', '44944', '45278'],
			['44936', '45182', '44947', '45281'],
			['44939', '45185', '44950', '45284'],
			['44942', '45188', '44953', '45287'],
			['44945', '45191', '44956', '45290']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Day, Step - 3');
		clearData(0, 0, 3, 5);
		// DateUnit - Weekday
		oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.weekday;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 3, 5);
		expectedData = [
			['44930', '45176', '44943', '45275'],
			['44935', '45181', '44946', '45280'],
			['44938', '45184', '44951', '45285'],
			['44943', '45189', '44956', '45288'],
			['44946', '45194', '44959', '45293']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Weekday, Step - 3');
		clearData(0, 0, 3, 5);
		// DateUnit - Month
		oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.month;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 3, 5);
		expectedData = [
			['45017', '45264', '45028', '45363'],
			['45108', '45355', '45119', '45455'],
			['45200', '45447', '45211', '45547'],
			['45292', '45539', '45303', '45638'],
			['45383', '45630', '45394', '45728']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Month, Step - 3');
		clearData(0, 0, 3, 5);
		// DateUnit - Year
		oFromRange = getFilledData(0, 0, 3, 5, testData, [0, 0]);
		settings.dateUnit = oSeriesDateUnitType.year;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 3, 5);
		expectedData = [
			['46023', '46269', '46034', '46368'],
			['47119', '47365', '47130', '47464'],
			['48214', '48461', '48225', '48560'],
			['49310', '49556', '49321', '49655'],
			['50406', '50652', '50417', '50751']
		];
		autofillData(assert, autofillRange, expectedData, 'Date progression - Year, Step - 3');
		clearData(0, 0, 3, 5);
	});
	QUnit.test('Fill -> Series. Trend. Horizontal - Multiple cells', function (assert) {
		let testData = [
		  ['4', '2', '0']
		];
		// Linear type with trend mode
		oFromRange = getFilledData(0, 0, 6, 0, testData, [0, 0]);
		settings = {
			'type': oSeriesType.linear,
			'stepValue': null,
			'seriesIn': oSeriesInType.rows,
			'stopValue': null,
			'trend': true
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(3, 0, 6, 0);
		autofillData(assert, autofillRange, [['-2', '-4', '-6', '-8']], 'Autofill Rows. Linear type with trend mode');
		clearData(0,0,6,0);
		// Growth type with trend mode
		testData = [
			['16', '8', '4']
		];
		oFromRange = getFilledData(0, 0, 6, 0, testData, [0, 0]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(3, 0, 6, 0);
		// In UI results: 0.49999999999999994 and 0.25000000000000006 will be round to 0.5 and 0.25. Need's UI test it
		autofillData(assert, autofillRange, [['2', '1', '0.49999999999999994', '0.25000000000000006']], 'Autofill Rows. Growth type with trend mode');
		clearData(0,0,6,0);
		// Linear multiple lines with trend mode
		testData = [
			['4', '2', '0'],
			['16', '8', '4'],
			['2', '4', '8'],
			['1', '2'],
			['1']
		];
		oFromRange = getFilledData(0, 0, 6, 4, testData, [0, 0]);
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 6, 4);
		expectedData = [
			['4', '2', '0', '-2', '-4', '-6', '-8'],
			['15.333333333333334', '9.333333333333334', '3.333333333333334', '-2.666666666666666', '-8.666666666666666', '-14.666666666666666', '-20.666666666666664'],
			['1.666666666666667', '4.666666666666667', '7.666666666666667', '10.666666666666668', '13.666666666666668',	'16.666666666666668', '19.666666666666668'],
			['1', '2', '3', '4', '5', '6', '7'],
			['1', '2', '3', '4', '5', '6', '7']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear multiple lines with trend mode');
		clearData(0,0,6,4);
		// Growth multiple lines with trend mode
		oFromRange = getFilledData(0, 0, 6, 4, testData, [0, 0]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 6, 4);
		// In UI results: Numbers like 0.49999999999999994 and 0.25000000000000006 will be round to 0.5 and 0.25. Need's UI test it.
		expectedData = [
			['4', '2', '0', '', '', '', ''],
			['15.999999999999998', '7.999999999999998', '4', '2', '1', '0.49999999999999994', '0.25000000000000006'],
			['2', '4', '7.999999999999998', '15.999999999999991', '31.999999999999986', '63.99999999999998', '127.99999999999986'],
			['1', '2', '4', '7.999999999999998', '15.999999999999998', '32', '63.99999999999998'],
			['1', '1', '1', '1', '1', '1', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth multiple lines with trend mode');
		clearData(0, 0, 6, 4);
		// Linear multiple Rows with trend mode. With indentation
		oFromRange = getFilledData(1, 0, 7, 4, testData, [0, 1]);
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 7, 4);
		expectedData = [
			['4', '2', '0', '-2', '-4', '-6', '-8'],
			['15.333333333333336', '9.333333333333336', '3.3333333333333357', '-2.6666666666666643', '-8.666666666666664', '-14.666666666666664', '-20.666666666666664'],
			['1.666666666666667', '4.666666666666667', '7.666666666666667', '10.666666666666668', '13.666666666666668',	'16.666666666666668', '19.666666666666668'],
			['1', '2', '3', '4', '5', '6', '7'],
			['1', '2', '3', '4', '5', '6', '7']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear multiple Rows with trend mode. With indentation');
		clearData(0, 0, 7, 4);
		// Growth multiple Rows with trend mode. With indentation
		oFromRange = getFilledData(1, 0, 7, 4, testData, [0, 1]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 0, 7, 4);
		expectedData = [
			['4', '2', '0', '', '', '', ''],
			['15.999999999999991', '7.999999999999995', '3.999999999999999', '1.9999999999999993', '0.9999999999999996', '0.49999999999999994', '0.24999999999999994'],
			['2', '4', '7.999999999999995', '15.999999999999991', '31.999999999999986', '63.99999999999992', '127.99999999999986'],
			['1', '2', '4', '8.000000000000002', '16.000000000000007', '31.999999999999986', '63.99999999999998'],
			['1', '1', '1', '1', '1', '1', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth multiple Rows with trend mode. With indentation');
		clearData(0, 0, 7, 4);
		// Growth multiple Rows with trend mode. With indentation row and col
		oFromRange = getFilledData(1, 1, 7, 5, testData, [1, 1]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 1, 7, 5);
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth multiple Rows with trend mode. With indentation row and col');
		clearData(0, 0, 7, 5);
		// Growth progression with trend mode. Start with 0 value.
		testData = [
			['0', '1'],
			['0', '2']
		];
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(2, 0, 3, 1);
		expectedData = [
			['', ''],
			['', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with trend mode. Start with 0 value.');
		clearData(0, 0, 3, 1);
		// Growth progression with trend mode. Start with 0 value for first cell and use activeFillHandle
		oFromRange = getFilledData(0, 0, 1, 1, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 3, 1);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();

		autofillRange = getRange(2, 0, 3, 1);
		expectedData = [
			['0', '0'],
			['0', '0']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with trend mode. Start with 0 value for first cell and use activeFillHandle');
		clearData(0, 0, 3, 1);
		// Linear progression with trend mode. Start with empty cells
		testData = [
			['', '1', '2'],
			['', '', '3', '5'],
			['', '2']
		];
		oFromRange = getFilledData(0, 0, 5, 2, testData, [0, 0]);
		settings.type = oSeriesType.linear;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();

		autofillRange = getRange(0, 0, 5, 2);
		expectedData = [
			['0', '1', '2', '3', '4', '5'],
			['-1', '1', '3', '5', '7', '9'],
			['1', '2', '3', '4', '5', '6']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Linear progression with trend mode. Start with empty cells');
		clearData(0, 0, 5, 1);
		// Growth progression with trend mode. Start with empty cells
		testData = [
			['', '2', '4'],
			['', '4', '8'],
			['', '', '1']
		];
		oFromRange = getFilledData(0, 0, 5, 2, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();

		autofillRange = getRange(0, 0, 5, 2);
		expectedData = [
			['1', '2', '4', '7.999999999999998', '15.999999999999998', '32'],
			['2', '4', '7.999999999999998', '15.999999999999991', '31.999999999999986', '63.99999999999998'],
			['1', '1', '1', '1', '1', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Growth progression with trend mode. Start with empty cells');
		clearData(0, 0, 5, 1);
		// Growth progression with trend mode. With negative numbers as values in Range. Using activeFillHandle.
		testData = [
			['-1', '-2']
		];
		oFromRange = getFilledData(0, 0, 1, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 3, 0);
		wsView.fillHandleDirection = 0;
		let oSeriesSettings = api.asc_GetSeriesSettings();
		api.asc_FillCells(oRightClickOptions.growthTrend, oSeriesSettings);

		autofillRange = getRange(0, 0, 3, 0);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['-1', '-2', '', '']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['-1', '-2', '0', '0']], _desc);
		}, 'Autofill Rows. Growth progression with trend mode. With negative numbers as values in Range. Using activeFillHandle');
		clearData(0, 0, 3, 0);
	});
	QUnit.test('Fill -> Series. Trend. Vertical - Multiple cells', function (assert) {
		let testData = [
			['4'],
			['2'],
			['0']
		];
		// Linear type with trend mode
		oFromRange = getFilledData(0, 0, 0, 6, testData, [0, 0]);
		settings = {
			'type': oSeriesType.linear,
			'stepValue': null,
			'seriesIn': oSeriesInType.columns,
			'stopValue': null,
			'trend': true
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 3, 0, 6);
		autofillData(assert, autofillRange, [['-2'], ['-4'], ['-6'], ['-8']], 'Autofill Columns. Linear type with trend mode');
		clearData(0,0,0,6);
		//Growth type with trend mode
		testData = [
			['16'],
			['8'],
			['4']
		];
		oFromRange = getFilledData(0, 0, 0, 6, testData, [0, 0]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 3, 0, 6);
		// In UI results: 0.49999999999999994 and 0.25000000000000006 will be round to 0.5 and 0.25. Need's UI test it
		autofillData(assert, autofillRange, [['2'], ['1'], ['0.49999999999999994'], ['0.25000000000000006']], 'Autofill Columns. Growth type with trend mode');
		clearData(0,0,0,6);
		// Linear multiple lines with trend mode
		testData = [
			['4', '16', '2', '1', '1'],
			['2', '8', '4', '2'],
			['0', '4', '8']
		];
		oFromRange = getFilledData(0, 0, 4, 6, testData, [0, 0]);
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 4, 6);
		expectedData = [
			['4', '15.333333333333334', '1.666666666666667', '1', '1'],
			['2', '9.333333333333334', '4.666666666666667', '2', '2'],
			['0', '3.333333333333334', '7.666666666666667', '3', '3'],
			['-2', '-2.666666666666666', '10.666666666666668', '4', '4'],
			['-4', '-8.666666666666666', '13.666666666666668', '5', '5'],
			['-6', '-14.666666666666666', '16.666666666666668', '6', '6'],
			['-8', '-20.666666666666664', '19.666666666666668', '7', '7']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear multiple lines with trend mode');
		clearData(0, 0, 4, 6);
		// Growth multiple lines with trend mode
		oFromRange = getFilledData(0, 0, 4, 6, testData, [0, 0]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 4, 6);
		// In UI results: Numbers like 0.49999999999999994 and 0.25000000000000006 will be round to 0.5 and 0.25. Need's UI test it.
		expectedData = [
			['4', '15.999999999999998', '2', '1', '1'],
			['2', '7.999999999999998', '4', '2', '1'],
			['0', '4', '7.999999999999998', '4', '1'],
			['', '2', '15.999999999999991', '7.999999999999998', '1'],
			['', '1', '31.999999999999986', '15.999999999999998', '1'],
			['', '0.49999999999999994', '63.99999999999998', '32', '1'],
			['', '0.25000000000000006', '127.99999999999986', '63.99999999999998', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth multiple lines with trend mode');
		clearData(0, 0, 4, 6);
		// Linear multiple Rows with trend mode. With indentation
		oFromRange = getFilledData(0, 1, 4, 7, testData, [1, 0]);
		settings.type = oSeriesType.linear;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 4, 7);
		expectedData = [
			['4', '15.333333333333336', '1.666666666666667', '1', '1'],
			['2', '9.333333333333336', '4.666666666666667', '2', '2'],
			['0', '3.3333333333333357', '7.666666666666667', '3', '3'],
			['-2', '-2.6666666666666643', '10.666666666666668', '4', '4'],
			['-4', '-8.666666666666664', '13.666666666666668', '5', '5'],
			['-6', '-14.666666666666664', '16.666666666666668', '6', '6'],
			['-8', '-20.666666666666664', '19.666666666666668', '7', '7']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear multiple Rows with trend mode. With indentation');
		clearData(0, 0, 4, 7);
		// Growth multiple Rows with trend mode. With indentation
		oFromRange = getFilledData(0, 1, 4, 7, testData, [1, 0]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 1, 4, 7);
		expectedData = [
			['4', '15.999999999999991', '2', '1', '1'],
			['2', '7.999999999999995', '4', '2', '1'],
			['0', '3.999999999999999', '7.999999999999995', '4', '1'],
			['', '1.9999999999999993', '15.999999999999991', '8.000000000000002', '1'],
			['', '0.9999999999999996', '31.999999999999986', '16.000000000000007', '1'],
			['', '0.49999999999999994', '63.99999999999992', '31.999999999999986', '1'],
			['', '0.24999999999999994', '127.99999999999986', '63.99999999999998', '1']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth multiple Rows with trend mode. With indentation');
		clearData(0, 0, 4, 7);
		// Growth multiple Columns with trend mode. With indentation row and col
		oFromRange = getFilledData(1, 1, 5, 7, testData, [1, 1]);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(1, 1, 5, 7);
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth multiple Columns with trend mode. With indentation row and col');
		clearData(0, 0, 5, 7);
		// Growth progression with trend mode. Start with 0 value.
		testData = [
			['0', '0'],
			['1', '2']
		];
		oFromRange = getFilledData(0, 0, 1, 3, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 2, 1, 3);
		expectedData = [
			['', ''],
			['', '']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with trend mode. Start with 0 value.');
		clearData(0, 0, 1, 3);
		// Growth progression with trend mode. Start with 0 value for first cell and use activeFillHandle
		oFromRange = getFilledData(0, 0, 1, 1, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 1, 3);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();

		autofillRange = getRange(0, 2, 1, 3);
		expectedData = [
			['0', '0'],
			['0', '0']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with trend mode. Start with 0 value for first cell and use activeFillHandle');
		clearData(0, 0, 3, 1);
		// Linear progression with trend mode. Start with empty cells
		testData = [
			['', '', ''],
			['1', '', '2'],
			['2', '3', ''],
			['', '5', '']
		];
		oFromRange = getFilledData(0, 0, 2, 5, testData, [0, 0]);
		settings.type = oSeriesType.linear;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();

		autofillRange = getRange(0, 0, 2, 5);
		expectedData = [
			['0', '-1', '1'],
			['1', '1', '2'],
			['2', '3', '3'],
			['3', '5', '4'],
			['4', '7', '5'],
			['5', '9', '6']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Linear progression with trend mode. Start with empty cells');
		clearData(0, 0, 2, 5);
		// Growth progression with trend mode. Start with empty cells
		testData = [
			['', '', ''],
			['2','4', ''],
			['4', '8', '2']
		];
		oFromRange = getFilledData(0, 0, 2, 5, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.trend = true;
		settings.stopValue = null;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();

		autofillRange = getRange(0, 0, 2, 5);
		expectedData = [
			['1', '2', '2'],
			['2', '4', '2'],
			['4', '7.999999999999998', '2'],
			['7.999999999999998', '15.999999999999991', '2'],
			['15.999999999999998', '31.999999999999986', '2'],
			['32', '63.99999999999998', '2']
		]
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Growth progression with trend mode. Start with empty cells');
		clearData(0, 0, 1, 5);
		// Growth progression with trend mode. With negative numbers as values in Range. Using activeFillHandle.
		testData = [
			['-1'],
			['-2']
		];
		oFromRange = getFilledData(0, 0, 0, 1, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 3);
		wsView.fillHandleDirection = 1;
		let oSeriesSettings = api.asc_GetSeriesSettings();
		api.asc_FillCells(oRightClickOptions.growthTrend, oSeriesSettings);

		autofillRange = getRange(0, 0, 0, 3);
		checkUndoRedo(function (_desc) {
			autofillData(assert, autofillRange, [['-1'], ['-2'], [''], ['']], _desc);
		}, function (_desc) {
			autofillData(assert, autofillRange, [['-1'], ['-2'], ['0'], ['0']], _desc);
		}, 'Autofill Columns. Growth progression with trend mode. With negative numbers as values in Range. Using activeFillHandle');
		clearData(0, 0, 0, 3);
	});
	QUnit.test('Autofill Series. StopValue out of range', function (assert) {
		const testData = [
			['1']
		];
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		settings = {
			'type': oSeriesType.linear,
			'stepValue': 1,
			'seriesIn': oSeriesInType.rows,
			'stopValue': 10,
			'trend': false
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 10, 0);
		autofillData(assert, autofillRange, [['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '']], 'Autofill Rows. Linear. StopValue out of range. Step 1. StopValue 10');
		clearData(0, 0, 9, 0);
		// Growth. Step 2. StopValue 20
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 2;
		settings.stopValue = 20;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 9, 0);
		autofillData(assert, autofillRange, [['1', '2', '4', '8', '16', '', '', '', '', '']], 'Autofill Rows. Growth. Step 2. StopValue 20');
		clearData(0, 0, 9, 0);
		// Columns. Linear. Step 2. StopValue 15
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		settings.type = oSeriesType.linear;
		settings.stepValue = 2;
		settings.stopValue = 15;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 9);
		autofillData(assert, autofillRange, [['1'], ['3'], ['5'], ['7'], ['9'], ['11'], ['13'], ['15'], [''], ['']], 'Autofill Columns. Linear. Step 2. StopValue 15');
		clearData(0, 0, 0, 9);
		// Columns. Growth. Step 3. StopValue 100
		oFromRange = getFilledData(0, 0, 0, 0, testData, [0, 0]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 3;
		settings.stopValue = 100;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 9);
		autofillData(assert, autofillRange, [['1'], ['3'], ['9'], ['27'], ['81'], [''], [''], [''], [''], ['']], 'Autofill Columns. Growth. Step 3. StopValue 100');
		clearData(0, 0, 0, 9);
		// Rows. Linear. Step 0.5. StopValue 10. With indentation rows and columns
		oFromRange = getFilledData(5, 2, 5, 2, testData, [2, 5]);
		settings.type = oSeriesType.linear;
		settings.stepValue = 0.5;
		settings.stopValue = 6;
		settings.seriesIn = oSeriesInType.rows;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(5, 2, 16, 2);
		autofillData(assert, autofillRange, [['1', '1.5', '2', '2.5', '3', '3.5', '4', '4.5', '5', '5.5', '6', '']], 'Autofill Rows. Linear. Step 0.5. StopValue 10. With indentation rows and columns');
		clearData(5, 2, 14, 2);
		// Columns. Growth. Step 10. StopValue 10000. With indentation rows and columns
		oFromRange = getFilledData(5, 2, 5, 2, testData, [2, 5]);
		settings.type = oSeriesType.growth;
		settings.stepValue = 10;
		settings.stopValue = 10000;
		settings.seriesIn = oSeriesInType.columns;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.exec();
		autofillRange = getRange(5, 2, 5, 7);
		autofillData(assert, autofillRange, [['1'], ['10'], ['100'], ['1000'], ['10000'], ['']], 'Autofill Columns. Growth. Step 10. StopValue 10000. With indentation rows and columns');
		clearData(5, 2, 5, 7);
	});
	QUnit.test('Autofill Series. Context menu. Horizontal', function (assert) {
		let testData = [
			['4', '2', '0']
		];
		// Context menu - Linear Trend
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 5, 0);
		settings = {
			'type': oSeriesType.linear,
			'stepValue': null,
			'seriesIn': oSeriesInType.rows,
			'stopValue': null,
			'trend': true
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(3, 0, 5, 0);
		autofillData(assert, autofillRange, [['-2', '-4', '-6']], 'Autofill Rows. Context menu - Linear Trend');
		clearData(0,0,5,0);
		// Context menu, reverse - Linear Trend
		oFromRange = getFilledData(3, 0, 5, 0, testData, [0, 3]);
		wsView.activeFillHandle = getRange(5, 0, 0, 0);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 2, 0);
		autofillData(assert, autofillRange, [['10', '8', '6']], 'Autofill Rows. Context menu, reverse - Linear Trend');
		clearData(0,0,5,0);
		// Context menu - Growth Trend
		testData = [
			['2', '4', '8']
		];
		oFromRange = getFilledData(0, 0, 2, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 5, 0);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(3, 0, 5, 0);
		autofillData(assert, autofillRange, [['15.999999999999991', '31.999999999999986', '63.99999999999998']], 'Autofill Rows. Context menu - Growth Trend');
		clearData(0,0,5,0);
		// Context menu, reverse - Growth Trend
		oFromRange = getFilledData(3, 0, 5, 0, testData, [0, 3]);
		wsView.activeFillHandle = getRange(5, 0, 0, 0);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 2, 0);
		autofillData(assert, autofillRange, [['0.2500000000000001', '0.5000000000000002', '1.0000000000000002']], 'Autofill Rows. Context menu, reverse - Growth Trend');
		clearData(0,0,5,0);
		// Context menu - Growth Trend. Series prop. Start with empty cells
		testData = [
			['', '2', '4', '']
		];
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 5, 0);
		settings.type = oSeriesType.growth;
		settings.contextMenuChosenProperty = oRightClickOptions.series;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 5, 0);
		autofillData(assert, autofillRange, [['1', '2', '4', '7.999999999999998', '15.999999999999998', '32']], 'Autofill Rows. Context menu - Growth Trend. Series prop. Start with empty cells');
		clearData(0, 0, 5, 0);
		// Context menu - Growth Trend. Growth prop. Start with empty cells
		oFromRange = getFilledData(0, 0, 3, 0, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 5, 0);
		settings.type = oSeriesType.growth;
		settings.contextMenuChosenProperty = oRightClickOptions.growthTrend;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 5, 0);
		autofillData(assert, autofillRange, [['', '2', '4', '', '15.999999999999998', '32']], 'Autofill Rows. Context menu - Growth Trend. Growth prop. Start with empty cells');
		clearData(0, 0, 5, 0);
		// Context menu. Reverse - Growth Trend. Series prop.
		testData = [
			['', '2', '4', '8']
		];
		oFromRange = getFilledData(4, 0, 7, 0, testData, [0, 4]);
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		settings.type = oSeriesType.growth;
		settings.contextMenuChosenProperty = oRightClickOptions.series;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 7, 0);
		let expectedData = [
			['0.06250000000000008', '0.1250000000000002', '0.2500000000000003', '0.5000000000000004', '1.0000000000000009', '2.0000000000000018', '4.000000000000001', '8.000000000000002']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu. Reverse - Growth Trend. Series prop.');
		clearData(0, 0, 7, 0);
		// Context menu. Reverse - Growth Trend. Growth prop.
		oFromRange = getFilledData(4, 0, 7, 0, testData, [0, 4]);
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		settings.type = oSeriesType.growth;
		settings.contextMenuChosenProperty = oRightClickOptions.growthTrend;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 7, 0);
		expectedData = [
			['0.06250000000000008', '0.1250000000000002', '0.2500000000000003', '0.5000000000000004', '', '2', '4', '8']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu. Reverse - Growth Trend. Growth prop.');
		clearData(0, 0, 7, 0);
	});
	QUnit.test('Autofill Series. Context menu. Vertical', function (assert) {
		let testData = [
			['4'],
			['2'],
			['0']
		];
		// Context menu - Linear Trend
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 5);
		settings = {
			'type': oSeriesType.linear,
			'stepValue': null,
			'seriesIn': oSeriesInType.columns,
			'stopValue': null,
			'trend': true
		};

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 3, 0, 5);
		autofillData(assert, autofillRange, [['-2'], ['-4'], ['-6']], 'Autofill Columns. Context menu - Linear Trend');
		clearData(0,0,0,5);
		// Context menu, reverse - Linear Trend
		oFromRange = getFilledData(0, 3, 0, 5, testData, [3, 0]);
		wsView.activeFillHandle = getRange(0, 5, 0, 0);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 2);
		autofillData(assert, autofillRange, [['10'], ['8'], ['6']], 'Autofill Columns. Context menu, reverse - Linear Trend');
		clearData(0, 0, 0, 5);
		// Context menu - Growth Trend
		testData = [
			['2'],
			['4'],
			['8']
		];
		oFromRange = getFilledData(0, 0, 0, 2, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 5);
		settings.type = oSeriesType.growth;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 3, 0, 5);
		autofillData(assert, autofillRange, [['15.999999999999991'], ['31.999999999999986'], ['63.99999999999998']], 'Autofill Columns. Context menu - Growth Trend');
		clearData(0, 0, 0, 5);
		// Context menu, reverse - Growth Trend
		oFromRange = getFilledData(0, 3, 0, 5, testData, [3, 0]);
		wsView.activeFillHandle = getRange(0, 5, 0, 0);

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 2);
		autofillData(assert, autofillRange, [['0.2500000000000001'], ['0.5000000000000002'], ['1.0000000000000002']], 'Autofill Columns. Context menu, reverse - Growth Trend');
		clearData(0, 0, 0, 5);
		// Context menu - Linear Trend. Series prop. Start with empty cells
		testData = [
			[''],
			['1'],
			['2'],
			['']
		];
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 5);
		settings.type = oSeriesType.linear;
		settings.contextMenuChosenProperty = oRightClickOptions.series;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 5);
		let expectedData = [
			['0'],
			['1'],
			['2'],
			['3'],
			['4'],
			['5']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Linear Trend. Series prop. Start with empty cells');
		clearData(0, 0, 0, 5);
		// Context menu. Linear Trend. Linear prop. Start with empty cells
		oFromRange = getFilledData(0, 0, 0, 3, testData, [0, 0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 5);
		settings.type = oSeriesType.linear;
		settings.contextMenuChosenProperty = oRightClickOptions.linearTrend;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [
			[''],
			['1'],
			['2'],
			[''],
			['4'],
			['5']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu. Linear Trend. Linear prop. Start with empty cells');
		clearData(0, 0, 0, 5);
		// Context menu. Reverse - Linear Trend. Series prop.
		testData = [
			[''],
			['2'],
			['3'],
			['4']
		];
		oFromRange = getFilledData(0, 4, 0, 7, testData, [4, 0]);
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		settings.type = oSeriesType.linear;
		settings.contextMenuChosenProperty = oRightClickOptions.series;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 7);
		expectedData = [
			['-3'],
			['-2'],
			['-1'],
			['0'],
			['1'],
			['2'],
			['3'],
			['4']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu. Reverse - Linear Trend. Series prop. Start with empty cells');
		clearData(0, 0, 0, 7);
		// Context menu. Reverse - Linear Trend. Linear prop.
		oFromRange = getFilledData(0, 4, 0, 7, testData, [4, 0]);
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		settings.type = oSeriesType.linear;
		settings.contextMenuChosenProperty = oRightClickOptions.linearTrend;

		cSerial = new CSerial(settings);
		cSerial.setFromRange(oFromRange);
		cSerial.setActiveFillHandle(wsView.activeFillHandle);
		cSerial.exec();
		autofillRange = getRange(0, 0, 0, 7);
		expectedData = [
			['-3'],
			['-2'],
			['-1'],
			['0'],
			[''],
			['2'],
			['3'],
			['4']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu. Reverse - Linear Trend. Linear prop. Start with empty cells');
		clearData(0, 0, 0, 7);
	});
	QUnit.test('CSeriesSettings: method prepare for prepare data in UI', function (assert) {
		let cSeriesSettings = Asc.asc_CSeriesSettings;

		// Series settings contains two filled cells. Horizontal. Not a date. Toolbar
		let testData = [
			['1', '3']
		];
		let oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings who contains two filled cells is created used by toolbar. Rows');
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings.prepare(wsView);

		// Check series settings data
		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step" is detected as 2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		let oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 3, 0);
		// Series settings contains two filled cells. Vertical. Not a date. Toolbar
		testData = [
			['1', '2'],
			['-1', '-2']
		];
		getFilledData(0, 0, 1, 3, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings who contains two filled cells is created used by toolbar. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, -2, 'oSeriesSettings: "Step" is detected as -2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 1, 3);
		// Series settings contains two filled cells. Horizontal. Date. Toolbar
		testData = [
			['11/08/2023', '11/11/2023'],
		]
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings who contains two filled cells is created used by toolbar. Rows, date');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 3, 'oSeriesSettings: "Step" is detected as 3.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 3, 0);
		// Series settings contains one filled cell. Vertical. Date. Toolbar.
		getFilledData(0, 0, 1, 3, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings who contains one filled cells is created used by toolbar. Columns, date');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 1, 3);
		// Series settings with empty cells. Horizontal. Toolbar
		testData = [];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with empty cells is created used by toolbar. Rows');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Series settings with empty cells. Vertical. Toolbar
		getFilledData(0, 0, 0, 3, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with empty cells is created used by toolbar. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 0, 3);
		// Series settings with one filled cell string type. Vertical. Toolbar
		testData = [
			['Test']
		];
		getFilledData(0, 0, 0, 3, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with one filled cells string type is created used by toolbar. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 0, 3);
		// Series settings with two filled cell formula type. Horizontal. Toolbar
		let oParseFormula =  new parserFormula('1+1', "A1", ws);
		oParseFormula.parse();
		oFromRange = ws.getRange2("A1");
		oFromRange._foreach(function(oCell) {
			oCell.formulaParsed = oParseFormula;
			oCell.setIsDirty(true);
			oCell._checkDirty();
			oCell.setValueNumberInternal(oCell.formulaParsed.value.getValue());
		});
		ws.getRange2("B1").setValue("4");
		oFromRange.worksheet.selectionRange.ranges = [getRange(0, 0, 3, 0)];
		oFromRange.bbox = getRange(0, 0, 3, 0);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with two filled cells formula type is created used by toolbar. Rows');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Series settings with two filled cell string with number at the end. Horizontal. Toolbar
		testData = [
			['Test1', 'Test2']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with two filled cells string with number at the end is created used by toolbar. Rows');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Series settings with one selected filled cell with using activeFillHandle. Horizontal. Context menu
		testData = [
			['1']
		];
		getFilledData(0, 0, 0, 0, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 3);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with one selected filled cell with using activeFillHandle is created used by context menu. Rows');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 3);
		// Series settings with one selected filled cell with using activeFillHandle. Vertical. Context menu
		getFilledData(0, 0, 0, 0, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 3, 0);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with one selected filled cell with using activeFillHandle is created used by context menu. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is nulld as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 0);
		// Series settings with one selected filled cell with using activeFillHandle - String without number. Vertical. Context menu
		testData = [
			['Test']
		];
		getFilledData(0, 0, 0, 0, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 3, 0);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings with one selected filled cell with using activeFillHandle is created used by context menu. Columns. String type without number in the end.');
		oSeriesSettings.prepare(wsView);

		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], false, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "false".');
		clearData(0, 0, 3, 0);
		// "Series". Three filled cells and four cells is selected. Rows.
		testData = [
			['10', '100', '100']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells and four cells is selected. Rows.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 45, 'oSeriesSettings: "Step" is detected as 45.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 3, 0);
		// "Series". Three filled cells are selected. Rows.
		getFilledData(0, 0, 2, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells are selected. Rows.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 2, 0);
		// "Series". Three filled cells  and four cells is selected. Columns.
		testData = [
			['10'],
			['100'],
			['100']
		];
		getFilledData(0, 7, 0, 10, testData, [7, 0])
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells  and four cells is selected. Columns.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 45, 'oSeriesSettings: "Step" is detected as 45.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 7, 0, 10);
		// "Series". Three filled cells are selected. Columns.
		getFilledData(0, 7, 0, 9, testData, [7,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells are selected. Columns.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 7, 0, 9);
		// "Series". Three filled cells but between first and second filled cells is empty cell. Selected five cells. Rows.
		testData = [
			['10', '', '100', '100']
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells but between first and second filled cells is empty cell. Selected five cells. Rows.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 32.14285714285714, 'oSeriesSettings: "Step" is detected as 33.13253012048193.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 4, 0);
		// "Series". Three filled cells but between first and second filled cells is empty cell. Selected five cells. Columns.
		testData = [
			['10'],
			[''],
			['100'],
			['100']
		];
		getFilledData(0, 7, 0, 11, testData, [7,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, '"Series". Three filled cells but between first and second filled cells is empty cell. Selected five cells. Columns.');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 32.14285714285714, 'oSeriesSettings: "Step" is detected as 33.13253012048193.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 7, 0, 11);
		// Case with 0 value of the first cell. 3 selected cells. Row
		testData = [
			['0', '2']
		];
		getFilledData(0, 0, 2, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with 0 value of the first cell. 3 selected cells. Row');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step" is detected as 2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 3, 0);
		// Case with 0 value of the first cell. 2 selected cells. Row
		getFilledData(0, 0, 1, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with 0 value of the first cell. 2 selected cells. Row');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 2, 0);
		// Case with 0 value of the first cell. 3 selected cells. Columns
		testData = [
			['0'],
			['2']
		];
		getFilledData(0, 0, 0, 2, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with 0 value of the first cell. 3 selected cells. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step" is detected as 2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 2);
		// Case with empty first cell, but one another cell in range is filled. 3 selected cells. Row
		testData = [
			['', '1']
		];
		getFilledData(0, 0, 2, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with empty first cell, but one another cell in range is filled. 3 selected cells. Row');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 2, 0);
		// Case with empty first cell, but one another cell in range is filled. 3 selected cells. Columns
		testData = [
			[''],
			['1']
		];
		getFilledData(0, 0, 0, 2, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with empty first cell, but one another cell in range is filled. 3 selected cells. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 2);
		// Case with empty first cell, but another cells in range are filled. 3 selected cells. Rows
		testData = [
			['', '', '2', '4'],
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with empty first cell, but another cells in range are filled. 5 selected cells. Rows');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 4, 0);
		// Case with empty first cell, but another cells in range are filled. 3 selected cells. Columns
		testData = [
			[''],
			[''],
			['2'],
			['4']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Case with empty first cell, but another cells in range are filled. 5 selected cells. Columns');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], true, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 4);
		// Select cells A1:A2 and active fill handle A3:A4 with one filled cell
		testData = [
			['1']
		];
		getFilledData(0, 0, 0, 1, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 3);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Select cells A1:A2 and active fill handle A3:A4 with one filled cell');
		oSeriesSettings.prepare(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], false, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 3);
		// Date unit - Month. Horizontal. Bug #65671.
		testData = [
			['01/01/1900', '02/01/1900', '03/01/1900', '04/01/1900', '05/01/1900']
		];
		getFilledData(0, 0, 5, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Date unit - Month. Selected cells A1:F1. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.month, 'oSeriesSettings: "Date unit" is detected as "Month".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 5, 0);
		// Date unit - Year. Vertical. Bug #65671.
		testData = [
			['01/01/1900'], ['01/01/1901'], ['01/01/1902'], ['01/01/1903'], ['01/01/1904']
		];
		getFilledData(0, 0, 0, 5, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Date unit - Year. Selected cells A1:A6 Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.year, 'oSeriesSettings: "Date unit" is detected as "Year".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 0, 5);
		// Date unit - day. Horizontal. Step - 9 Bug #65671.
		testData = [
			['01/01/1900', '01/10/1900', '01/11/1900', '01/12/1900', '01/13/1900']
		];
		getFilledData(0, 0, 5, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Date unit - day. Selected cells: A1:F1. Step - 9 Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 9, 'oSeriesSettings: "Step" is detected as 9.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		// contextMenuAllowedProps
		oMenuAllowedProps = oSeriesSettings.contextMenuAllowedProps;
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.copyCells], true, 'oSeriesSettings - contextMenuAllowedProps: "Copy cells" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillSeries], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill series" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillFormattingOnly], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill formatting only" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWithoutFormatting], null, 'oSeriesSettings - contextMenuAllowedProps: "Fill without formatting" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillDays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill days" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillWeekdays], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill weekdays" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillMonths], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill months" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.fillYears], true, 'oSeriesSettings - contextMenuAllowedProps: "Fill years" is detected as "true".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.linearTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Linear trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.growthTrend], false, 'oSeriesSettings - contextMenuAllowedProps: "Growth trend" is detected as "false".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.flashFill], null, 'oSeriesSettings - contextMenuAllowedProps: "Flash fill" is detected as "null".');
		assert.strictEqual(oMenuAllowedProps[oRightClickOptions.series], true, 'oSeriesSettings - contextMenuAllowedProps: "Series" is detected as "true".');
		clearData(0, 0, 5, 0);
		// Date unit - day. Vertical. Step - 14 Bug #65671.
		testData = [
			['02/01/2000'],
			[''],
			['02/29/2000']
		];
		getFilledData(0, 0, 0, 3, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Date unit - day. Selected cells: A1:A4. Step - 14 Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 14, 'oSeriesSettings: "Step" is detected as 14.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 3);
		// Date unit - day. Horizontal. Step - 1. Bug #65671.
		testData = [
			['01/01/1900', '', '02/01/1900']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Date unit - day. Selected cells: A1:D1. Step - 1. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Type Date. Date unit - Day. Step - 60. Horizontal. Bug #65671.
		testData = [
			['01/01/2000', '03/01/2000', '01/01/2000']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:D1. Step - 60. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 60, 'oSeriesSettings: "Step" is detected as 60.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Type Date. Date unit - Day. Step - 0. Vertical. Bug #65671.
		testData = [
			['01/01/2000'],
			['01/01/2000'],
			['03/01/2000'],
			['05/01/2000']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:A5. Step - 0. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 0, 'oSeriesSettings: "Step" is detected as 0.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 4);
		// Type Date. Date unit - Year. Step - 2. Horizontal. Bug #65671.
		testData = [
			['01/01/1900', '01/01/1902', '01/01/1904', '01/01/1906']
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Year. Selected cells: A1:E1. Step - 2. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.year, 'oSeriesSettings: "Date unit" is detected as "Year".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step" is detected as 2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 4, 0);
		// Type Date. Date unit - Month. Step - 16. Vertical. Bug #65671.
		testData = [
			['01/01/1900'],
			['05/01/1901'],
			['01/01/1902'],
			['01/01/1903']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Month. Selected cells: A1:A5. Step - 16. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.month, 'oSeriesSettings: "Date unit" is detected as "Month".');
		assert.strictEqual(oSeriesSettings.stepValue, 16, 'oSeriesSettings: "Step" is detected as 16.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 4);
		// Type Date. Date unit - Year. Step - 1. Horizontal. Bug #65671.
		testData = [
			['02/01/1900', '02/01/1901', '06/01/1902', '08/01/1903']
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Year. Selected cells: A1:E1. Step - 1. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.year, 'oSeriesSettings: "Date unit" is detected as "Year".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 4, 0);
		// Type Date. Date unit - Day. Step - 370. Vertical. Bug #65671.
		testData = [
			['01/01/1900'],
			['01/05/1901'],
			['01/01/1902'],
			['01/01/1903']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:A5. Step - 370. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 370, 'oSeriesSettings: "Step" is detected as 370.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 4);
		// Type Date. Date unit - Day. Step - 365. Horizontal. Bug #65671.
		testData = [
			['10/10/2000', '10/10/2001', '10/05/2002', '10/10/2003']
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:E1. Step - 365. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 365, 'oSeriesSettings: "Step" is detected as 365.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 4, 0);
		// Type Date. Date unit - Year. Step - -1. Vertical. Bug #65671.
		testData = [
			['01/01/1903'],
			['01/01/1902'],
			['01/01/1901'],
			['01/01/1900']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Year. Selected cells: A1:A5. Step - -1. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.year, 'oSeriesSettings: "Date unit" is detected as "Year".');
		assert.strictEqual(oSeriesSettings.stepValue, -1, 'oSeriesSettings: "Step" is detected as -1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 4);
		// Type Date. Date unit - Day. Step - 0. Horizontal. Bug #65671.
		testData = [
			['10/10/2000', '10/10/2000', '10/05/2002', '10/10/2003']
		];
		getFilledData(0, 0, 4, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:E1. Step - 0. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 0, 'oSeriesSettings: "Step" is detected as 0.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 4, 0);
		// Type Date. Date unit - Day. Step - 365. Vertical. Bug #65671.
		testData = [
			['01/01/1905'],
			['01/01/1906'],
			['01/01/1905'],
			['01/01/1904']
		];
		getFilledData(0, 0, 0, 4, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Type Date. Date unit - Day. Selected cells: A1:A5. Step - 365. Bug #65671.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 365, 'oSeriesSettings: "Step" is detected as 365.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 4);
		// Vertical selected range, first cell in range has type Date, another cells General. Bug #65873
		testData = [
			['01/01/1900'],
			['2'],
			['3']
		];
		getFilledData(0, 0, 0, 3, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Vertical selected range, first cell in range has type Date, another cells General. Bug #65873');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 3);
		// Horizontal selected range, first cell in range has type General, another cells Data. Bug #65873
		testData = [
			['1', '01/02/1900', '01/03/1900']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Horizontal selected range, first cell in range has type General, another cells Data. Bug #65873');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step" is detected as 1.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 3, 0);
		// Vertical selected range. Bug #65898
		testData = [
			['1'],
			['-0.2']
		];
		getFilledData(0, 0, 0, 2, testData, [0,0]);
		oSeriesSettings = api.asc_GetSeriesSettings();
		assert.ok(oSeriesSettings, 'Vertical selected range. Bug #65898.');

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.day, 'oSeriesSettings: "Date unit" is detected as "Day".');
		assert.strictEqual(oSeriesSettings.stepValue, -1.2, 'oSeriesSettings: "Step" is detected as -1.2.');
		assert.strictEqual(oSeriesSettings.stopValue, null, 'oSeriesSettings: "Stop value" is detected as empty.');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		clearData(0, 0, 0, 2);
	});
	QUnit.test('CSeriesSettings: init method for update type and trend step by chosen menu prop', function(assert) {
		const cSeriesSettings = Asc.asc_CSeriesSettings;

		//  Context menu property is "Linear Trend"
		let testData = [
			['1', '2']
		];
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		let oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created with chosen property "Linear Trend".');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setContextMenuChosenProperty(oRightClickOptions.linearTrend);
		oSeriesSettings.init(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.linear, 'oSeriesSettings: "Type" is detected as "Linear".');
		assert.strictEqual(oSeriesSettings.trend, true, 'oSeriesSettings: "Trend" is detected as "true".');
		clearData(0, 0, 3, 0);
		//  Context menu property is "Growth Trend"
		getFilledData(0, 0, 3, 0, testData, [0,0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created with chosen property "Growth Trend".');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setContextMenuChosenProperty(oRightClickOptions.growthTrend);
		oSeriesSettings.init(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.growth, 'oSeriesSettings: "Type" is detected as "Growth".');
		assert.strictEqual(oSeriesSettings.trend, true, 'oSeriesSettings: "Trend" is detected as "true".');
		clearData(0, 0, 3, 0);
		// Context menu property is "fillMonths" - Horizontal direction
		testData = [
			['01/01/2000', '02/01/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created with chosen property "fillMonths".');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setContextMenuChosenProperty(oRightClickOptions.fillMonths);
		oSeriesSettings.init(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.month, 'oSeriesSettings: "Date unit" is detected as "Month".');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		assert.strictEqual(oSeriesSettings.stepValue, 1, 'oSeriesSettings: "Step value" is detected as "1".');
		clearData(0, 0, 4, 0);
		// Context menu property is "fillMonths" - Vertical direction
		testData = [
			['01/01/2000'],
			['03/01/2000']
		];
		getFilledData(0, 0, 0, 1, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 0, 4);
		wsView.fillHandleDirection = 1;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created with chosen property "fillMonths". Columns');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setContextMenuChosenProperty(oRightClickOptions.fillMonths);
		oSeriesSettings.init(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.columns, 'oSeriesSettings: "Series in" is detected as "Columns".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.month, 'oSeriesSettings: "Date unit" is detected as "Month".');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step value" is detected as "2".');
		clearData(0, 0, 0, 4);
		// Context menu property is "fillYears" - Horizontal direction
		testData = [
			['01/01/2000', '01/01/2002']
		];
		getFilledData(0, 0, 1, 0, testData, [0,0]);
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created with chosen property "fillYears".');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setContextMenuChosenProperty(oRightClickOptions.fillYears);
		oSeriesSettings.init(wsView);

		assert.strictEqual(oSeriesSettings.seriesIn, oSeriesInType.rows, 'oSeriesSettings: "Series in" is detected as "Rows".');
		assert.strictEqual(oSeriesSettings.type, oSeriesType.date, 'oSeriesSettings: "Type" is detected as "Date".');
		assert.strictEqual(oSeriesSettings.dateUnit, oSeriesDateUnitType.year, 'oSeriesSettings: "Date unit" is detected as "Year".');
		assert.strictEqual(oSeriesSettings.trend, false, 'oSeriesSettings: "Trend" is detected as "false".');
		assert.strictEqual(oSeriesSettings.stepValue, 2, 'oSeriesSettings: "Step value" is detected as "2".');
		clearData(0, 0, 4, 0);
	});
	QUnit.test('applySeriesSettings', function(assert) {
		const cSeriesSettings = Asc.asc_CSeriesSettings;
		let testData = [
			['2', '4']
		];
		//  Context menu property is "Linear Trend"
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		let oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		nType = oRightClickOptions.linearTrend;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		api.asc_FillCells(nType, oSeriesSettings);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['6', '8', '10']], 'Autofill Rows. Context menu - Linear Trend');
		clearData(0, 0, 4, 0);
		//Context menu property is "Growth Trend"
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		nType = oRightClickOptions.growthTrend;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		api.asc_FillCells(nType, oSeriesSettings);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['7.999999999999998', '15.999999999999991', '31.999999999999986']], 'Autofill Rows. Context menu - Growth Trend');
		clearData(0, 0, 4, 0);
		//Context menu property is "Copy cells"
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.copyCells;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['2', '4', '2']], 'Autofill Rows. Context menu - Copy cells');
		clearData(0, 0, 4, 0);
		// Context menu property is "Fill series"
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['6', '8', '10']], 'Autofill Rows. Context menu - Fill series');
		clearData(0, 0, 4, 0);
		// Toolbar "Series"
		getFilledData(0, 0, 4, 0, testData, [0, 0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		nType = oRightClickOptions.series;
		api.asc_FillCells(nType, oSeriesSettings);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['6', '8', '10']], 'Autofill Rows. Toolbar - Series - Linear');
		clearData(0, 0, 4, 0);
		// Context menu property is "Series". Rows.
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.series;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		api.asc_FillCells(nType, oSeriesSettings);

		autofillRange = getRange(2, 0, 4, 0);
		autofillData(assert, autofillRange, [['4', '5', '6']], 'Autofill Rows. Context menu - Series - Linear.');
		clearData(0, 0, 4, 0);
		// Context menu property is "Copy cell" with selected one cell number type data. Rows.
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.copyCells;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 4, 0);
		autofillData(assert, autofillRange, [['2', '2', '2', '2']], 'Autofill Rows. Context menu - Copy cells. Selected one cell with number type data.');
		clearData(0, 0, 4, 0);
		// Context menu property is "Fill series" with selected one cell number type data. Rows.
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 4, 0);
		autofillData(assert, autofillRange, [['3', '4', '5', '6']], 'Autofill Rows. Context menu - Fill series. Selected one cell with number type data.');
		clearData(0, 0, 4, 0);
		// Context menu property is "Series" with selected one cell number type data. Rows.
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.series;
		wsView.activeFillHandle = getRange(0, 0, 4, 0);
		wsView.fillHandleDirection = 0;
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		api.asc_FillCells(nType, oSeriesSettings);

		autofillRange = getRange(1, 0, 4, 0);
		autofillData(assert, autofillRange, [['3', '4', '5', '6']], 'Autofill Rows. Context menu - Series - Linear. Selected one cell with number type data.');
		clearData(0, 0, 4, 0);
		// Context menu property is "Copy cells" with selected one cell Date type data. Columns.
		testData = [
			['01/01/2000']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.copyCells;
		wsView.activeFillHandle = getRange(0, 0, 0, 4);
		wsView.fillHandleDirection = 1;
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 4);
		autofillData(assert, autofillRange, [['36526'], ['36526'], ['36526'], ['36526']], 'Autofill Columns. Context menu - Copy cells. Selected one cell with Date type data.');
		clearData(0, 0, 0, 4);
		// Context menu property is "Fill series" with selected one cell Date type data. Columns.
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 0, 4);
		wsView.fillHandleDirection = 1;
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 4);
		autofillData(assert, autofillRange, [['36527'], ['36528'], ['36529'], ['36530']], 'Autofill Columns. Context menu - Fill series. Selected one cell with Date type data.');
		clearData(0, 0, 0, 4);
		// The context menu property is "Fill series" with selected filled cells - A1:F1. The fill handle has a vertical direction. Case: bug #65405
		testData = [
			['1', '1', 'Test1', 'Test1', '01/01/2000', '01/01/2000']
		];
		getFilledData(0, 0, 5, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 5, 2);
		wsView.fillHandleDirection = 1;
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 5, 2);
		expectedData = [
			['2', '2', 'Test2', 'Test2', '36527', '36527'],
			['3', '3', 'Test3', 'Test3', '36528', '36528']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill series. Selected filled cells - A1:F1. The fill handle has a vertical direction. Case: bug #65405');
		clearData(0, 0, 5, 2);
		// The context menu property is "Copy cell" with selected filled cells - A1:F1. The fill handle has a vertical direction. Case: bug #65405
		getFilledData(0, 0, 5, 0, testData, [0, 0]);
		nType = oRightClickOptions.copyCells;
		wsView.activeFillHandle = getRange(0, 0, 5, 2);
		wsView.fillHandleDirection = 1;
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 5, 2);
		expectedData = [
			['1', '1', 'Test1', 'Test1', '36526', '36526'],
			['1', '1', 'Test1', 'Test1', '36526', '36526']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Copy cells. Selected filled cells - A1:F1. The fill handle has a vertical direction. Case: bug #65405');
		clearData(0, 0, 5, 2);
		// The context menu property is "Fill series" with selected filled cells - A1:A6. The fill handle has a horizontal direction. Case: bug #65405
		testData = [
			['1'],
			['1'],
			['Test1'],
			['Test1'],
			['01/01/2000'],
			['01/01/2000']
		];
		getFilledData(0, 0, 0, 5, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 2, 5);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 2, 5);
		expectedData = [
			['2', '3'],
			['2', '3'],
			['Test2', 'Test3'],
			['Test2', 'Test3'],
			['36527', '36528'],
			['36527', '36528']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill series. Selected filled cells - A1:A6. The fill handle has a horizontal direction. Case: bug #65405');
		clearData(0, 0, 2, 5);
		// The context menu property is "Copy cell" with selected filled cells - A1:A6. The fill handle has a horizontal direction. Case: bug #65405
		getFilledData(0, 0, 0, 5, testData, [0, 0]);
		nType = oRightClickOptions.copyCells;
		wsView.activeFillHandle = getRange(0, 0, 2, 5);
		wsView.fillHandleDirection = 0;
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 2, 5);
		expectedData = [
			['1', '1'],
			['1', '1'],
			['Test1', 'Test1'],
			['Test1', 'Test1'],
			['36526', '36526'],
			['36526', '36526']
		];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Copy cells. Selected filled cells - A1:A6. The fill handle has a horizontal direction. Case: bug #65405');
		clearData(0, 0, 2, 5);
		// Vertical selected range, but "Series in" - Rows. Bug #65551
		testData = [
			['1']
		];
		getFilledData(0, 0, 0, 3, testData, [0, 0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setSeriesIn(oSeriesInType.rows);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(1, 0, 1, 0);
		expectedData = [
			['']
		];
		autofillData(assert, autofillRange, expectedData, 'Vertical selected range, but "Series in" - Rows.');
		clearData(0, 0, 0, 3);
		// Horizontal selected range, but "Series in" - Columns. Bug #65551
		getFilledData(0, 0, 3, 0, testData, [0, 0]);
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setSeriesIn(oSeriesInType.columns);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(0, 1, 0, 1);
		expectedData = [
			['']
		];
		autofillData(assert, autofillRange, expectedData, 'Horizontal selected range, but "Series in" - Columns.');
		clearData(0, 0, 3, 0);
		// Case: The context menu property "Fill weekdays". Vertical. One selected cell. Asc sequence
		testData = [
			['01/01/2000']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['36528'], ['36529'], ['36530'], ['36531'], ['36532'], ['36535'], ['36536']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. One selected cell. Asc sequence');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Horizontal. One selected cell. Asc sequence. Type Date
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 7, 0);
		expectedData = [['36528', '36529', '36530', '36531', '36532', '36535', '36536']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. One selected cell. Asc sequence');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Vertical. One selected cell. Reverse sequence. Type Date
		getFilledData(0, 7, 0, 7, testData, [7, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['36517'], ['36518'], ['36521'], ['36522'], ['36523'], ['36524'], ['36525']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. One selected cell. Reverse sequence');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Horizontal. One selected cell. Reverse sequence. Type Date
		getFilledData(7, 0, 7, 0, testData, [0, 7]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 6, 0);
		expectedData = [['36517', '36518', '36521', '36522', '36523', '36524', '36525']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. One selected cell. Reverse sequence. Type Date');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Vertical. Two selected cells. Asc sequence. Type Date.
		testData = [
			['01/01/2000'],
			['01/03/2000']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 7);
		expectedData = [['36530'], ['36532'], ['36536'], ['36538'], ['36542'], ['36544']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. Two selected cells. Asc sequence. Type Date');
		clearData(0, 0, 0, 7);

		// Case: The context menu property "Fill weekdays". Vertical. Two selected cells. Reverse sequence. Type Date.
		getFilledData(0, 6, 0, 7, testData, [6, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [['36510'], ['36514'], ['36516'], ['36518'], ['36522'], ['36524']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. Two selected cells. Reverse sequence. Type Date');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Horizontal. Three selected cells. Asc sequence. Type Date.
		testData = [
			['01/02/2000', '01/04/2000', '01/06/2000']
		];
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 7, 0);
		expectedData = [['36535', '36537', '36539', '36543', '36545']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Three selected cells. Asc sequence. Type Date');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Horizontal. Three selected cells. Reverse sequence. Type Date
		getFilledData(5, 0, 7, 0, testData, [0, 5]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 4, 0);
		expectedData = [['36514', '36516', '36518', '36522', '36524']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Three selected cells. Reverse sequence. Type Date');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Vertical. One selected cell. Asc sequence. Type Date & Time
		testData = [
			['01/01/2000 12:00']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['36528'], ['36529'], ['36530'], ['36531'], ['36532'], ['36535'], ['36536']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. One selected cell. Asc sequence. Type Date & Time');
		clearData(0, 0, 0, 7);

		// Case: The context menu property "Fill weekdays". Vertical. One selected cell. Reverse sequence. Type Date & Time
		getFilledData(0, 7, 0, 7, testData, [7, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['36517'], ['36518'], ['36521'], ['36522'], ['36523'], ['36524'], ['36525']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. One selected cell. Reverse sequence. Type Date & Time');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekends". Horizontal.  Two selected cells. Asc sequence. Type Mixed date.
		testData = [
			['01/01/2000 12:00', '01/02/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 7, 0);
		expectedData = [['36528', '36529', '36530', '36531', '36532', '36535']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Two selected cells. Asc sequence. Type Mixed date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Horizontal. Two selected cells. Reverse sequence. Type Mixed date.
		testData = [
			['01/01/2000 12:00', '01/02/2000']
		];
		getFilledData(6, 0, 7, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['36518', '36521', '36522', '36523', '36524', '36525']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekends. Two selected cells. Reverse sequence. Type Mixed date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill weekdays". Vertical. One selected cell. Asc sequence. Type Date. 1900 year.
		testData = [
			['01/01/1900']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['2'], ['3'], ['4'], ['5'], ['6'], ['9'], ['10']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekends. One selected cell. Asc sequence. Type Date. 1900 year.');
		// Case: The context menu property "Fill weekdays". Horizontal. One selected cell. Reverse sequence. Type Date. 1900 year
		getFilledData(7, 0, 7, 0, testData, [0, 7]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 6, 0);
		expectedData = [['1', '1', '1', '1', '1', '1', '1']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. One selected cell. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Vertical. Two selected cells. Asc sequence. Type Date. 1900 year
		testData = [
			['01/01/1900'],
			['01/03/1900']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 7);
		expectedData = [['5'], ['9'], ['11'], ['13'], ['17'], ['19']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekends. Two selected cells. Asc sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Horizontal. Two selected cells. Reverse sequence. Type Date. 1900 year
		testData = [
			['01/04/1900', '01/06/1900']
		];
		getFilledData(6, 0, 7, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['4', '6', '4', '6', '4', '2']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Two selected cells. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);

		// Case: The context menu property "Fill weekdays". Vertical. Three selected cells. Asc sequence. Type Date. 1900 year
		testData = [
			['01/02/1900'],
			['01/04/1900'],
			['01/06/1900']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 7);
		expectedData = [['10'], ['12'], ['16'], ['18'], ['20']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. Three selected cells. Asc sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);

		// Case: The context menu property "Fill weekdays". Horizontal. Three selected cells. Reverse sequence. Type Date. 1900 year
		testData = [
			['01/09/1900', '01/11/1900', '01/13/1900']
		];
		getFilledData(5, 0, 7, 0, testData, [0, 5]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 4, 0);
		expectedData = [['11','13', '9', '3','5']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Three selected cells. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Vertical. Three selected cells. Negative case - incorrect sequence. Asc sequence. Type Date.
		testData = [
			['01/02/2000'],
			['01/04/2000'],
			['01/05/2000']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 7);
		expectedData = [['36527'], ['36529'], ['36530'], ['36527'], ['36529']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill weekdays. Three selected cells. Negative case - incorrect sequence. Asc sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill weekdays". Vertical. Three selected cells. Negative case - incorrect sequence. Reverse sequence. Type Date.
		getFilledData(0, 5, 0, 7, testData, [5, 0]);
		nType = oRightClickOptions.fillWeekdays;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 4);
		expectedData = [['36529'], ['36530'], ['36527'], ['36529'], ['36530']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill weekdays. Three selected cells. Negative case - incorrect sequence. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Vertical.  One selected cell. Asc sequence. Type Date.
		testData = [
			['01/01/2000']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['36557'], ['36586'], ['36617'], ['36647'], ['36678'], ['36708'], ['36739']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. One selected cell. Asc sequence. Type Date.');
		clearData(0, 0, 0, 7);

		// Case: The context menu property "Fill months". Vertical. One selected cell. Reverse sequence. Type Date.
		getFilledData(0, 7, 0, 7, testData, [7, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['36312'], ['36342'], ['36373'], ['36404'], ['36434'], ['36465'], ['36495']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. One selected cell. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Asc sequence. Type Date.
		testData = [
			['01/01/2000', '03/01/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 7, 0);
		expectedData = [['36647', '36708', '36770', '36831', '36892', '36951']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Asc sequence. Type Date.');
		clearData(0, 0, 7, 0);

		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Reverse sequence. Type Date.
		getFilledData(6, 0, 7, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['36161', '36220', '36281', '36342', '36404', '36465']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Asc sequence. Type Date.
		testData = [
			['02/01/2000'],
			['04/01/2000'],
			['06/01/2000']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 7);
		expectedData = [['36739'], ['36800'], ['36861'], ['36923'], ['36982']]
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Asc sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Reverse sequence. Type Date.
		getFilledData(0, 5, 0, 7, testData, [5, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);
		
		autofillRange = getRange(0, 0, 0, 4);
		expectedData = [['36251'], ['36312'], ['36373'], ['36434'], ['36495']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Three selected cells. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Asc sequence. Negative case - incorrect sequence. Type Date.
		testData = [
			['01/01/2000', '03/01/2000', '06/01/2000']
		];
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 7, 0);
		expectedData = [['36526', '36586', '36678', '36526', '36586']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Asc sequence. Negative case - incorrect sequence. Type Date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Reverse sequence. Negative case - incorrect sequence. Type Date.
		testData = [
			['01/01/2000', '03/01/2000', '06/01/2000']
		];
		getFilledData(5, 0, 7, 0, testData, [0, 5]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 4, 0);
		expectedData = [['36586', '36678', '36526', '36586', '36678']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Reverse sequence. Negative case - incorrect sequence. Type Date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Two selected cells. Vertical. Asc sequence. Type Date & Time.
		testData = [
			['01/01/2000 12:00'],
			['03/01/2000 13:00']
		];
		getFilledData(0, 0, 0, 1, testData, [0,0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 7);
		expectedData = [['36647'], ['36708'], ['36770'], ['36831'], ['36892'],['36951']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Vertical. Asc sequence. Type Date & Time.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Vertical. Two selected cells. Reverse sequence. Type Date & Time.
		getFilledData(0, 6, 0, 7, testData, [6, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [['36161'], ['36220'], ['36281'], ['36342'], ['36404'], ['36465']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date & Time.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Asc sequence. Type Mixed date.
		testData = [
			['02/01/2000 12:00', '04/01/2000', '06/01/2000 13:00']
		];
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 7, 0);
		expectedData = [['36739', '36800', '36861', '36923', '36982']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Three selected cells. Asc sequence. Type Mixed date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Reverse sequence. Type Mixed date.
		getFilledData(5, 0, 7, 0, testData, [0, 5]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 4, 0);
		expectedData = [['36251', '36312', '36373', '36434', '36495']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Three selected cells. Reverse sequence. Type Mixed date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Vertical. One selected cell. Asc sequence. Type Date. 1900 year.
		testData = [
			['01/01/1900']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['32'], ['61'], ['92'], ['122'], ['153'], ['183'], ['214']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. One selected cell. Asc sequence. Type Date. 1900 year');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Vertical. One selected cell. Reverse sequence. Type Date. 1900 year.
		getFilledData(0, 7, 0, 7, testData, [7, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['1'], ['1'], ['1'], ['1'], ['1'], ['1'], ['1']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. One selected cell. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Asc sequence. Type Date. 1900 year.
		testData = [
			['02/01/1900', '04/01/1900']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 7, 0);
		expectedData = [['153', '214', '275', '336', '398', '457']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Asc sequence. Type Date. 1900 year.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Reverse sequence. Type Date. 1900 year.
		testData = [
			['06/01/1900', '08/01/1900']
		];
		getFilledData(6, 0, 7, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['153', '214', '153', '214', '32', '92']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill years". Vertical. One selected cell. Asc sequence. Type Date.
		testData = [
			['01/01/2000']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 1, 0, 7);
		expectedData = [['36892'], ['37257'], ['37622'], ['37987'], ['38353'], ['38718'], ['39083']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. One selected cell. Asc sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years". Vertical. One selected cell. Reverse sequence. Type Date.
		getFilledData(0, 7, 0, 7, testData, [7, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['33970'], ['34335'], ['34700'], ['35065'], ['35431'], ['35796'], ['36161']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. One selected cell. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years". Horizontal. Two selected cells. Asc sequence. Type Date.
		testData = [
			['01/01/2000', '01/01/2002']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 7, 0);
		expectedData = [['37987', '38718', '39448', '40179', '40909', '41640']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Two selected cells. Asc sequence. Type Date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill years". Horizontal. Two selected cells. Reverse sequence. Type Date.
		getFilledData(6, 0, 7, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['32143', '32874', '33604', '34335', '35065', '35796']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Two selected cells. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years". Vertical. Three selected cells. Asc sequence. Type Date.
		testData = [
			['01/01/2000'],
			['01/01/2025'],
			['01/01/2050']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 7);
		expectedData = [['63920'], ['73051'], ['82182'], ['91313'], ['100444']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Three selected cells. Asc sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years". Vertical. Three selected cells. Reverse sequence. Type Date.
		getFilledData(0, 5, 0, 7, testData, [5, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 4);
		expectedData = [['45658'], ['1'], ['9133'], ['18264'], ['27395']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Three selected cells. Reverse sequence. Type Date.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years". Horizontal. One selected cell. Asc sequence. Type Date. 1900 year.
		testData = [
			['01/01/1900']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 7, 0);
		expectedData = [['367', '732', '1097', '1462', '1828', '2193', '2558']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. One selected cell. Asc sequence. Type Date. 1900 year.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill years". Horizontal. One selected cell. Reverse sequence. Type Date. 1900 year.
		getFilledData(7, 0, 7, 0, testData, [0, 7]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(7, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 6, 0);
		expectedData = [['1', '1', '1', '1', '1', '1', '1']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. One selected cell. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill years".  Vertical. Two selected cells. Asc sequence. Type Date. 1900 year
		testData = [
			['01/01/1900'],
			['01/01/1902']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 7);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 7);
		expectedData = [['1462'], ['2193'], ['2923'], ['3654'], ['4384'], ['5115']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Asc sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill years".  Vertical. Two selected cells. Asc sequence. Type Date. 1900 year
		getFilledData(0, 6, 0, 7, testData, [6, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 7, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5)
		expectedData = [['1'], ['732'], ['1'], ['732'], ['1'], ['732']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Reverse sequence. Type Date. 1900 year.');
		clearData(0, 0, 0, 7);
		// Case: The context menu property "Fill series". Horizontal. Two selected cells. Asc sequence. Type Date.
		testData = [
			['01/01/2000', '04/01/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillSeries;
		wsView.activeFillHandle = getRange(0, 0, 7, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 7, 0);
		expectedData = [['36708', '36800', '36892', '36982', '37073', '37165']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill series. Two selected cells. Asc sequence. Type Date.');
		clearData(0, 0, 7, 0);
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Asc sequence. Type Date. Diff days.
		testData = [
			['01/01/2000'],
			['01/02/2000'],
			['01/03/2000']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 8);
		expectedData = [['36557'], ['36558'], ['36559'], ['36586'], ['36587'], ['36588']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill month. Three selected cells. Asc sequence. Type Date. Diff days.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Reverse sequence. Type Date. Diff days.
		getFilledData(0, 6, 0, 8, testData, [6, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [['36465'], ['36466'], ['36467'], ['36495'], ['36496'], ['36497']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill month. Three selected cells. Reverse sequence. Type Date. Diff days.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill years". Horizontal. Three selected cells. Asc sequence. Type Date. Diff days.
		testData = [
			['01/01/2000', '01/02/2000', '01/03/2000']
		]
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 8, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 8, 0);
		expectedData = [['36892', '36893', '36894', '37257', '37258', '37259']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Three selected cells. Asc sequence. Type Date. Diff days.');
		// Case: The context menu property "Fill years". Horizontal. Three selected cells. Reverse sequence. Type Date. Diff days.
		getFilledData(6, 0, 8, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(8, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['35796', '35797', '35798', '36161', '36162', '36163']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Three selected cells. Reverse sequence. Type Date. Diff days.');
		clearData(0, 0, 8, 0);
		// Case: The context menu property "Fill months". Vertical. Two selected cells. Asc sequence. Type Date. Diff days.
		testData = [
			['01/01/2000'],
			['01/30/2000']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 8);
		expectedData = [['36557'], ['36585'], ['36586'], ['36615'], ['36617'], ['36646'], ['36647']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Asc sequence. Type Date. Diff days.');
		// Case: The context menu property "Fill months". Vertical. Two selected cells. Reverse sequence. Type Date. Diff days.
		getFilledData(0, 7, 0, 8, testData, [7, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['36433'], ['36434'], ['36463'], ['36465'], ['36494'], ['36495'], ['36524']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date. Diff days.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Asc sequence. Type Date. Diff days.
		testData = [
			['01/01/2000', '01/31/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 8, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 8, 0);
		expectedData = [['36557', '36585', '36586', '36616', '36617', '36646', '36647']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Asc sequence. Type Date. Diff days.');
		// Case: The context menu property "Fill months". Horizontal. Two selected cells. Reverse sequence. Type Date. Diff days.
		getFilledData(7, 0, 8, 0, testData, [0, 7]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(8, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 6, 0);
		expectedData = [['36433', '36434', '36464', '36465', '36494', '36495', '36525']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date. Diff days.');
		clearData(0, 0, 8, 0);
		// Case: The context menu property "Fill years". Vertical. Two selected cells. Asc sequence. Type Date. Diff days.
		testData = [
			['01/01/2000'],
			['01/31/2000']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 8);
		expectedData = [['36892'], ['36922'], ['37257'], ['37287'], ['37622'], ['37652'], ['37987']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Asc sequence. Type Date. Diff days.');
		// Case: The context menu property "Fill years". Vertical. Two selected cells. Reverse sequence. Type Date. Diff days.
		getFilledData(0, 7, 0, 8, testData, [7, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1 // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['35095'], ['35431'], ['35461'], ['35796'], ['35826'], ['36161'], ['36191']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Reverse sequence. Type Date. Diff days.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill years". Horizontal. Two selected cells. Asc sequence. Type Date. Diff months.
		testData = [
			['01/01/2000', '02/01/2000']
		];
		getFilledData(0, 0, 1, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 8, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(2, 0, 8, 0);
		expectedData = [['36892', '36923', '37257', '37288', '37622', '37653', '37987']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Two selected cells. Asc sequence. Type Date. Diff months.');
		// Case: The context menu property "Fill years". Horizontal. Two selected cells. Reverse sequence. Type Date. Diff months.
		getFilledData(7, 0, 8, 0, testData, [0, 7]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(8, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 6, 0);
		expectedData = [['35096', '35431', '35462', '35796', '35827', '36161', '36192']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Two selected cells. Reverse sequence. Type Date. Diff months.');
		clearData(0, 0, 8, 0);
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Asc sequence. Type Date. Negative case - incorrect sequence.
		testData = [
			['01/12/2000'],
			['01/13/2000'],
			['01/15/2000']
		];
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 8);
		expectedData = [['36537'], ['36538'], ['36540'], ['36537'], ['36538'], ['36540']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Asc sequence. Type Date. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill months". Vertical. Three selected cells. Reverse sequence. Type Date. Negative case - incorrect sequence.
		getFilledData(0, 6, 0, 8, testData, [6, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [['36537'], ['36538'], ['36540'], ['36537'], ['36538'], ['36540']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Reverse sequence. Type Date. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill years". Vertical. Three selected cells. Asc sequence. Type Date. Negative case - incorrect sequence.
		getFilledData(0, 0, 0, 2, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 3, 0, 8);
		expectedData = [['36537'], ['36538'], ['36540'], ['36537'], ['36538'], ['36540']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Three selected cells. Asc sequence. Type Date. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill years". Vertical. Three selected cells. Reverse sequence. Type Date. Negative case - incorrect sequence.
		getFilledData(0, 6, 0, 8, testData, [6, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 5);
		expectedData = [['36537'], ['36538'], ['36540'], ['36537'], ['36538'], ['36540']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Three selected cells. Reverse sequence. Type Date. Negative case - incorrect sequence.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Asc sequence. Type Date & Time. Negative case - incorrect sequence.
		testData = [
			['01/01/2000 12:00', '01/02/2000 13:00', '01/04/2000 14:00']
		];
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 8, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 8, 0);
		expectedData = [['36526.5', '36527.541666666664', '36529.583333333336', '36526.5', '36527.541666666664', '36529.583333333336']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Three selected cells. Asc sequence. Type Date & Time. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill months". Horizontal. Three selected cells. Reverse sequence. Type Date & Time. Negative case - incorrect sequence.
		getFilledData(6, 0, 8, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(8, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['36526.5', '36527.541666666664', '36529.583333333336', '36526.5', '36527.541666666664', '36529.583333333336']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill months. Three selected cells. Reverse sequence. Type Date & Time. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill years". Horizontal. Three selected cells. Asc sequence. Type Date & Time. Negative case - incorrect sequence.
		getFilledData(0, 0, 2, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 8, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(3, 0, 8, 0);
		expectedData = [['36526.5', '36527.541666666664', '36529.583333333336', '36526.5', '36527.541666666664', '36529.583333333336']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Three selected cells. Asc sequence. Type Date & Time. Negative case - incorrect sequence.');
		// Case: The context menu property "Fill years". Horizontal. Three selected cells. Reverse sequence. Type Date & Time. Negative case - incorrect sequence.
		getFilledData(6, 0, 8, 0, testData, [0, 6]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(8, 0, 0, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 5, 0);
		expectedData = [['36526.5', '36527.541666666664', '36529.583333333336', '36526.5', '36527.541666666664', '36529.583333333336']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Rows. Context menu - Fill years. Three selected cells. Reverse sequence. Type Date & Time. Negative case - incorrect sequence.');
		clearData(0, 0, 8, 0);
		// Case: The context menu property "Fill months". Vertical. Two selected cells. Asc sequence. Type Date & Time. Diff time.
		testData = [
			['01/01/2000 12:00'],
			['01/01/2000 13:00']
		];
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 8);
		expectedData = [['36557'], ['36557'], ['36586'], ['36586'], ['36617'], ['36617'], ['36647']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Asc sequence. Type Date & Time. Diff time.');
		// Case: The context menu property "Fill months". Vertical. Two selected cells. Reverse sequence. Type Date & Time. Diff time.
		getFilledData(0, 7, 0, 8, testData, [7, 0]);
		nType = oRightClickOptions.fillMonths;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['36404'], ['36434'], ['36434'], ['36465'], ['36465'], ['36495'], ['36495']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill months. Two selected cells. Reverse sequence. Type Date & Time. Diff time.');
		// Case: The context menu property "Fill years". Vertical. Two selected cells. Asc sequence. Type Date & Time. Diff time.
		getFilledData(0, 0, 0, 1, testData, [0, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 0, 0, 8);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 2, 0, 8);
		expectedData = [['36892'], ['36892'], ['37257'], ['37257'], ['37622'], ['37622'], ['37987']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Asc sequence. Type Date & Time. Diff time.');
		// Case: The context menu property "Fill years". Vertical. Two selected cells. Reverse sequence. Type Date & Time. Diff time.
		getFilledData(0, 7, 0, 8, testData, [7, 0]);
		nType = oRightClickOptions.fillYears;
		wsView.activeFillHandle = getRange(0, 8, 0, 0);
		wsView.fillHandleDirection = 1; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(0, 0, 0, 6);
		expectedData = [['35065'], ['35431'], ['35431'], ['35796'], ['35796'], ['36161'], ['36161']];
		autofillData(assert, autofillRange, expectedData, 'Autofill Columns. Context menu - Fill years. Two selected cells. Reverse sequence. Type Date & Time. Diff time.');
		clearData(0, 0, 0, 8);
		// Case: The context menu property "Fill months". Horizontal. One selected cells. Asc sequence. Type Date. The junction between February and March. Case 1900
		testData = [
			['01/31/1900']
		];
		getFilledData(0, 0, 0, 0, testData, [0, 0]);
		nType = oRightClickOptions.fillMonths
		wsView.activeFillHandle = getRange(0, 0, 3, 0);
		wsView.fillHandleDirection = 0; // 0 - Horizontal, 1 - Vertical.
		api.asc_FillCells(nType);

		autofillRange = getRange(1, 0, 3, 0);
		expectedData = [['59', '91', '121']];
		autofillData(assert, autofillRange, expectedData, "Autofill Columns. Context menu - Fill months. One selected cells. Asc sequence. Type Date. The junction between February and March. Case 1900");
		clearData(0, 0, 3, 0);
	});
	QUnit.test('Toolbar: Fill -> "Up/Down, Left/Right"', function(assert) {
		const testData = [
			['1']
		];
		// Fill -> Down with merged cell
		oFromRange = getFilledData(0, 0, 1, 3, testData, [0,0]);
		let oFromWs = initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		api.asc_FillCells(oRightClickOptions.fillDown);

		autofillRange = getRange(0, 2, 0, 2);
		autofillData(assert, autofillRange, [['1']], 'Fill -> Down with merged cell');
		let oToMergedCell = oFromWs.mergeManager.get(autofillRange);
		assert.strictEqual(oToMergedCell.all.length > 0, true, "Filled cell must be merged" );
		clearData(0, 0, 1, 3);
		// Fill -> Up with merged cell
		oFromRange = getFilledData(0, 0, 1, 3, testData, [3,0]);
		oFromWs = initMergedCell(oFromRange, getRange(0, 2, 1, 3));
		api.asc_FillCells(oRightClickOptions.fillUp);

		autofillRange = getRange(0, 1, 0, 1);
		autofillData(assert, autofillRange, [['1']], 'Fill -> Up with merged cell');
		oToMergedCell = oFromWs.mergeManager.get(autofillRange);
		assert.strictEqual(oToMergedCell.all.length > 0, true, "Filled cell must be merged" );
		clearData(0, 0, 1, 3);
		// Fill -> Right with merged cell
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		oFromWs = initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		api.asc_FillCells(oRightClickOptions.fillRight);

		autofillRange = getRange(2, 0, 2, 0);
		autofillData(assert, autofillRange, [['1']], 'Fill -> Right with merged cell');
		oToMergedCell = oFromWs.mergeManager.get(autofillRange);
		assert.strictEqual(oToMergedCell.all.length > 0, true, "Filled cell must be merged" );
		clearData(0, 0, 3, 1);
		// Fill -> Left with merged cell
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,3]);
		oFromWs = initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		api.asc_FillCells(oRightClickOptions.fillLeft);

		autofillRange = getRange(1, 0, 1, 0);
		autofillData(assert, autofillRange, [['1']], 'Fill -> Left with merged cell');
		oToMergedCell = oFromWs.mergeManager.get(autofillRange);
		assert.strictEqual(oToMergedCell.all.length > 0, true, "Filled cell must be merged" );
		clearData(0, 0, 3, 1);
	});
	QUnit.test('Series with merged cells', function(assert) {
		const cSeriesSettings = Asc.asc_CSeriesSettings;
		let testData = [
			['1']
		];
		// Toolbar and context menu - Series: Filled cells has merged cell, cells that need to be filled are not merged cells. NOT Trend mode.
		oFromRange = getFilledData(0, 0, 5, 1, testData, [0,0]);
		let oWsFrom = initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		let oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(2, 0, 5, 1);
		autofillData(assert, autofillRange, [['2','', '3', ''], ['', '', '', '']], 'Toolbar and context menu - Series: Filled cells has merged cell, cells that need to be filled are not merged cells. NOT Trend mode.');
		oWsFrom.mergeManager.get(autofillRange);
		assert.strictEqual(oWsFrom.mergeManager.get(autofillRange).all.length > 0, true, "ToRange must be merged" );
		clearData(0, 0, 5, 1);
		// Toolbar and context menu - Series: Filled cell has merged cell, cells that need to be filled are not merged cells. Trend mode.
		oFromRange = getFilledData(0, 0, 5, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Negative case: oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setTrend(true);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(2, 0, 5, 1);
		autofillData(assert, autofillRange, [['', '', '', ''], ['', '', '', '']], 'Toolbar and context menu - Series: Filled cell has merged cell, cells that need to be filled are not merged cells. Trend mode.');
		clearData(0, 0, 5, 1);
		// Toolbar and context menu - Series: Filled cell has merged cells, cells that need to be filled are merged cells. Trend mode.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Negative case: oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setTrend(true);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(2, 0, 3, 1);
		autofillData(assert, autofillRange, [['', ''], ['', '']], 'Toolbar and context menu - Series: Filled cell has merged cells, cells that need to be filled are merged cells. Trend mode.');
		clearData(0, 0, 3, 1);
		// Toolbar and context menu - Series: Filled cells has merged cell, cells that need to be filled are  merged cells. NOT Trend mode.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(2, 0, 3, 1);
		autofillData(assert, autofillRange, [['2', ''], ['', '']], 'Toolbar and context menu - Series: Filled cells has merged cell, cells that need to be filled are merged cells. NOT Trend mode.');
		clearData(0, 0, 3, 1);

		testData = [
			['1', '', '2', '']
		];
		// Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are not merged cells. NOT Trend mode.
		oFromRange = getFilledData(0, 0, 7, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Negative case: oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(4, 0, 7, 1);
		autofillData(assert, autofillRange, [['','', '', ''], ['', '', '', '']], 'Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are not merged cells. NOT Trend mode.');
		clearData(0, 0, 7, 1);
		// Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are not merged cells. Trend mode.
		oFromRange = getFilledData(0, 0, 7, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Negative case: oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setTrend(true);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(4, 0, 7, 1);
		autofillData(assert, autofillRange, [['', '', '', ''], ['', '', '', '']], 'Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are not merged cells. Trend mode.');
		clearData(0, 0, 7, 1);
		// Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are merged cells. Trend mode.
		oFromRange = getFilledData(0, 0, 5, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		initMergedCell(oFromRange, getRange(4, 0, 5, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'Negative case: oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setTrend(true);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(4, 0, 5, 1);
		autofillData(assert, autofillRange, [['', ''], ['', '']], 'Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are merged cells. Trend mode.');
		clearData(0, 0, 5, 1);
		// Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are merged cells. NOT Trend mode.
		oFromRange = getFilledData(0, 0, 5, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		initMergedCell(oFromRange, getRange(4, 0, 5, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		oSeriesSettings.asc_setStepValue(1);
		api.asc_FillCells(oRightClickOptions.series, oSeriesSettings);

		autofillRange = getRange(4, 0, 5, 1);
		autofillData(assert, autofillRange, [['3', ''], ['', '']], 'Toolbar and context menu - Series: Filled cells have merged cells, cells that need to be filled are merged cells. NOT Trend mode.');
		clearData(0, 0, 5, 1);
		// Context menu - Linear trend: Filled cells have merged cells, cells that need to be filled are not merged cells.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		oWsFrom = initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		wsView.activeFillHandle = getRange(0, 0, 7, 1);
		api.asc_FillCells(oRightClickOptions.linearTrend, oSeriesSettings);

		autofillRange = getRange(4, 0, 7, 0);
		autofillData(assert, autofillRange, [['3', '3.5', '4', '4.5']], 'Context menu - Linear trend: Filled cells have merged cells, cells that need to be filled are not merged cells.');
		oWsFrom.mergeManager.get(autofillRange);
		assert.strictEqual(oWsFrom.mergeManager.get(autofillRange).all.length > 0, true, "ToRange must be merged" );
		clearData(0,0,7,1);
		// Context menu - Growth trend: Filled cells have merged cells, cells that need to be filled are not merged cells.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		oWsFrom = initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		wsView.activeFillHandle = getRange(0, 0, 7, 1);
		api.asc_FillCells(oRightClickOptions.growthTrend, oSeriesSettings);

		autofillRange = getRange(4, 0, 7, 0);
		autofillData(assert, autofillRange, [['4', '5.65685424949238', '7.999999999999998', '11.31370849898476']], 'Context menu - Growth trend: Filled cells have merged cells, cells that need to be filled are not merged cells.');
		oWsFrom.mergeManager.get(autofillRange);
		assert.strictEqual(oWsFrom.mergeManager.get(autofillRange).all.length > 0, true, "ToRange must be merged" );
		clearData(0,0,7,1);
		// Context menu - Linear trend: Filled cells have merged cells, cells that need to be filled are merged cells.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		initMergedCell(oFromRange, getRange(4, 0, 5, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		wsView.activeFillHandle = getRange(0, 0, 5, 1);
		api.asc_FillCells(oRightClickOptions.linearTrend, oSeriesSettings);

		autofillRange = getRange(4, 0, 5, 0);
		autofillData(assert, autofillRange, [['3', '3.5']], 'Context menu - Linear trend: Filled cells have merged cells, cells that need to be filled are merged cells.');
		oWsFrom.mergeManager.get(autofillRange);
		assert.strictEqual(oWsFrom.mergeManager.get(autofillRange).all.length > 0, true, "ToRange must be merged" );
		clearData(0,0,5,1);
		// Context menu - Growth trend: Filled cells have merged cells, cells that need to be filled are merged cells.
		oFromRange = getFilledData(0, 0, 3, 1, testData, [0,0]);
		initMergedCell(oFromRange, getRange(0, 0, 1, 1));
		initMergedCell(oFromRange, getRange(2, 0, 3, 1));
		initMergedCell(oFromRange, getRange(4, 0, 5, 1));
		oSeriesSettings = new cSeriesSettings();
		assert.ok(oSeriesSettings, 'oSeriesSettings is created.');
		oSeriesSettings.prepare(wsView);
		wsView.activeFillHandle = getRange(0, 0, 5, 1);
		api.asc_FillCells(oRightClickOptions.growthTrend, oSeriesSettings);

		autofillRange = getRange(4, 0, 5, 0);
		autofillData(assert, autofillRange, [['4', '5.65685424949238']], 'Context menu - Growth trend: Filled cells have merged cells, cells that need to be filled are merged cells.');
		oWsFrom.mergeManager.get(autofillRange);
		assert.strictEqual(oWsFrom.mergeManager.get(autofillRange).all.length > 0, true, "ToRange must be merged" );
		clearData(0,0,5,1);
	});
});
