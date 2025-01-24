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

QUnit.config.autostart = false;
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

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
			setTimeout(startTests, 0)
		}
	};
	window["Asc"]["editor"] = api;

	var wb, ws, wsview;

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());
		api._openOnClient();
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);

		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function () {
			return null;
		});

		wsview = api.wb.getWorksheet();
		wsview.objectRender = {};
		wsview.objectRender.updateDrawingObject = function () {
		};
		wsview.objectRender.updateSizeDrawingObjects = function () {
		};
		wsview.objectRender.selectedGraphicObjectsExists = function () {
		};
		wsview.handlers = {};
		wsview.handlers.trigger = function () {
		};
		ws = api.wbModel.aWorksheets[0];
	}

	function testValidTitle() {
		QUnit.test("Test: check valid title name", function (assert) {
			let checkRes = api.asc_checkProtectedRangeName("test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_1");
			checkRes = api.asc_checkProtectedRangeName("test1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_2");
			checkRes = api.asc_checkProtectedRangeName("test_1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_3");
			checkRes = api.asc_checkProtectedRangeName("test_ 1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_4");
			checkRes = api.asc_checkProtectedRangeName("test _ 1");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_4");

			checkRes = api.asc_checkProtectedRangeName("test!");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.WrongName, "check valid name_5");
			checkRes = api.asc_checkProtectedRangeName("1test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.WrongName, "check valid name_6");

			checkRes = api.asc_checkProtectedRangeName("t test");
			assert.strictEqual(checkRes, Asc.c_oAscDefinedNameReason.OK, "check valid name_8");
		});
	}

	QUnit.module("ProtectTests");

	function startTests() {
		QUnit.start();

		testValidTitle();
	}
});
