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
QUnit.config.autostart = false;
$(function() {


	Asc.spreadsheet_api.prototype._init = function() {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype.asyncImagesDocumentEndLoaded = function() {
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function(fonts, callback) {
		openDocument();
	};
	Asc.spreadsheet_api.prototype._onEndLoadSdk = function () {
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);
	};
	Asc.spreadsheet_api.prototype.initGlobalObjects = function(wbModel) {
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wbModel, -1);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function() {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function() {
	};
	AscCommonExcel.WorkbookView.prototype.getZoom = function() {
		return 1;
	};
	AscCommonExcel.WorkbookView.prototype.changeZoom = function() {
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
		this.objectRender = new AscFormat.DrawingObjects();
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function() {
		return 1;
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
		AscFonts.g_fontApplication.Init();

		this.FontLoader  = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);
		this.FontLoader.SetStandartFonts();
	};

	Asc.spreadsheet_api.prototype.fAfterLoad = function(fonts, callback) {
		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function() {
			return null;
		});
		wsView = api.wb.getWorksheet();
		ws = api.wbModel.aWorksheets[0];

		startTests();
	};

	AscCommon.CHistory.prototype.Create_NewPoint = function() {
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());
		api._openOnClient();
	}

	api.HtmlElement = document.createElement("div");
	var curElem = document.getElementById("editor_sdk");
	curElem.appendChild(api.HtmlElement);
	window["Asc"]["editor"] = api;

	function updateView () {
		wsView._cleanCache(new Asc.Range(0, 0, wsView.cols.length - 1, wsView.rows.length - 1));
		wsView.changeWorksheet("update", {reinitRanges: true});
	}

	let wb, wsView, ws;
	function testShowFormulasOption() {
		QUnit.test("Test: show formulas option ", function(assert ) {

			let testData = [
				["=1+1", "test1"],
				["=2+1", "test2"],
				["=3+1", "test3"],
				["=4+1", "test4"],
				["=5+1", "test5"],
				["=6+1", "test6"]
			];

			let range = ws.getRange4(0, 0);
			range.fillData(testData);

			updateView();

			api.asc_setShowFormulas(true);

			updateView();

			assert.strictEqual("1", "1");

		});
	}

	QUnit.module("Print");

	function startTests() {
		QUnit.start();
		testShowFormulasOption();
	}

});
