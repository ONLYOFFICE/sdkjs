/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
QUnit.config.autostart = false;
$(function() {

	Asc.spreadsheet_api.prototype._init = function() {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function(fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function(fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function() {
	};
	AscCommonExcel.WorkbookView.prototype._init = function() {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function() {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function() {
	};
	AscCommonExcel.WorksheetView.prototype._init = function() {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function() {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function() {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function() {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function() {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function() {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function() {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIY = function() {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIX = function() {
	};


	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {
			setTimeout(startTests, 0)
		}
	};
	window["Asc"]["editor"] = api;
	var wb, ws, wsData, wsView;

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());
		api._openOnClient();
		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function () {
			return null;
		});
		ws = api.wbModel.aWorksheets[0];
		api.asc_insertWorksheet(["Data"]);
		wsData = wb.getWorksheetByName(["Data"], 0);

		wsView = api.wb.getWorksheet(0);
		wsView.handlers = api.handlers;
		wsView.objectRender = new AscFormat.DrawingObjects();
		wsView.objectRender.init(wsView);
		//wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);
	}

	function fillData(ws, data, range) {
		range = ws.getRange4(range.r1, range.c1);
		range.fillData(data);
	}

	function testChartBaseTypes() {
		QUnit.test("Test: Base Charts Draw ", function(assert ) {
			let testData =  [
				["", 2014, 2015, 2016],
				["Projected Revenue", 200, 240, 280],
				["Estimated Costs", 250, 260]
			];
			let testDataRange = new Asc.Range(0, 0, testData[0].length - 1, testData.length - 1);
			fillData(wsData, testData, testDataRange);

			var oWorksheet = api.GetActiveSheet();
			var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
			oChart.Chart.worksheet = wsView.model;

			oChart.Chart.recalcInfo = {
				/*recalculateAxisLabels: true,
				recalculateAxisTickMark: true,
				recalculateAxisVal: true,
				recalculateBBoxRange: true,
				recalculateBounds: true,*/
				dataLbls: [],
				axisLabels: [],
				recalculateChart:true
				/*recalculateDLbls: true,
				recalculateFormulas: true,
				recalculateGridLines: true,
				recalculateHiLowLines: true,
				recalculateLegend: true,
				recalculatePen: true,
				recalculatePenBrush: true,
				recalculatePlotAreaBrush: true,
				recalculatePlotAreaPen: true,
				recalculateReferences: true,
				recalculateSeriesColors: true,
				recalculateTextPr: true,
				recalculateUpDownBars: true*/
			};

			oChart.Chart.recalculate();
			assert.strictEqual(1, 1, "Validations");
		});
	}


	QUnit.module("ChartsDraw");

	function startTests() {
		QUnit.start();

		testChartBaseTypes();
	}
});
