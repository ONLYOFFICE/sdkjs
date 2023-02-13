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
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
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

	var wb, ws;

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
	}

	function create(ref, name) {
		let obj = new Asc.CUserProtectedRange(ws);
		obj.asc_setRef(ref);
		obj.asc_setName("test");
		//obj.asc_setUsers("");
		api.asc_addUserProtectedRange(obj);
		return obj;
	}

	function testCreate(init) {
		QUnit.test("Test: create", function (assert) {
			//ADD
			init("B2:B5", "test1");

			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "ref compare");

			init("D2:E5", "test2");
			assert.strictEqual(ws.userProtectedRanges.length, 2, "add test");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "ref compare");

			AscCommon.History.Undo();
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "redo add test");

			//DELETE
			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0]]);
			assert.strictEqual(ws.userProtectedRanges.length, 1, "delete_test_1");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", "ref compare");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_2");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 1, "delete_test_3");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_4");

			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0], ws.userProtectedRanges[1]]);
			assert.strictEqual(ws.userProtectedRanges.length, 0, "delete_test_5");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_6");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "delete_test_7");
			AscCommon.History.Undo();

		});
	}

	function testChange(init) {
		QUnit.test("Test: change", function (assert) {
			init("B2:B5", "test1");

			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "ref compare");

			init("D2:E5", "test2");
			assert.strictEqual(ws.userProtectedRanges.length, 2, "add test");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "ref compare");

			AscCommon.History.Undo();
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "redo add test");

			//DELETE
			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0]]);
			assert.strictEqual(ws.userProtectedRanges.length, 1, "delete_test_1");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", "ref compare");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_2");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 1, "delete_test_3");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_4");

			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0], ws.userProtectedRanges[1]]);
			assert.strictEqual(ws.userProtectedRanges.length, 0, "delete_test_5");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 2, "delete_test_6");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "delete_test_7");
			AscCommon.History.Undo();

		});
	}

	QUnit.module("UserProtectedRanges");

	function startTests() {
		QUnit.start();

		testCreate(create);
		testChange(create);
	}
});
