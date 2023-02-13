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
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
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
		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
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
		wsview.handlers = {};
		wsview.handlers.trigger = function () {
		};
		ws = api.wbModel.aWorksheets[0];
	}

	function create(ref, name) {
		let obj = new Asc.CUserProtectedRange(ws);
		obj.asc_setRef(ref);
		obj.asc_setName(name);
		//obj.asc_setUsers("");
		api.asc_addUserProtectedRange(obj);
		return obj;
	}

	function testCreate() {
		QUnit.test("Test: create", function (assert) {
			//ADD
			create("B2:B5", "test");

			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "ref compare");

			create("D2:E5", "test2");
			assert.strictEqual(ws.userProtectedRanges.length, 2, "add test");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getName(), "test2", "name compare");
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
		});
	}

	function testChange() {
		QUnit.test("Test: change", function (assert) {
			create("B2:B5", "test1");

			let obj = ws.userProtectedRanges[0].clone(ws);
			obj.asc_setRef("B2:B10");

			api.asc_changeUserProtectedRange(ws.userProtectedRanges[0], obj);
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$10", "change ref compare1");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "change ref compare2");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$10", "change ref compare3");

			obj = ws.userProtectedRanges[0].clone(ws);
			obj.asc_setName("test2");
			api.asc_changeUserProtectedRange(ws.userProtectedRanges[0], obj);
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test2", "change name compare1");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test1", "change name compare2");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test2", "change name compare3");

			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0]]);
			assert.strictEqual(ws.userProtectedRanges.length, 0, "delete_test_8");
		});
	}

	function checkUndoRedo(fBefore, fAfter) {
		fAfter();
		AscCommon.History.Undo();
		fBefore();
		AscCommon.History.Redo();
		fAfter();
		AscCommon.History.Undo();
	}

	function testManipulation() {
		QUnit.test("Test: change", function (assert) {
			create("B2:B5", "test1");
			create("D2:E5", "test2");

			let beforeFunc = function() {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "before_reference_val");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "before_reference_val_2");
			};

			wsview.setSelection(new Asc.Range(0, 0, 0, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertColumns);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$C$2:$C$5", "insert columns ref compare1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$E$2:$F$5", "insert columns ref compare2");
			});

			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertColumns);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "insert columns ref compare9");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$F$5", "insert columns ref compare10");
			});

			wsview.setSelection(new Asc.Range(0, 1, 2, 4));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftRight);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$E$2:$E$5", "insert columns ref compare17");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$G$2:$H$5", "insert columns ref compare18");
			});

			wsview.setSelection(new Asc.Range(0, 1, 2, 3));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftRight);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "insert columns ref compare19");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "insert columns ref compare20");
			});


			wsview.setSelection(new Asc.Range(0, 0, 3, 0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftDown);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$3:$B$6", "insert columns ref compare19");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "insert columns ref compare20");
			});

			//delete cells
			wsview.setSelection(new Asc.Range(1, 0, 3, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteColumns);
			checkUndoRedo(beforeFunc, function (){
				assert.strictEqual(ws.userProtectedRanges.length, 1, "delete columns ref compare1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "delete columns ref compare2");
			});


			/*wsview.setSelection(new Asc.Range(0, 1, 2, 4));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftLeft);
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$E$2:$E$5", "insert columns ref compare17");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$G$2:$H$5", "insert columns ref compare18");
			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "insert columns ref compare19");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", "insert columns ref compare20");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$E$2:$E$5", "insert columns ref compare21");
			assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$G$2:$H$5", "insert columns ref compare22");
			AscCommon.History.Undo();*/


		});
	}

	QUnit.module("UserProtectedRanges");

	function startTests() {
		QUnit.start();

		testCreate();
		testChange();
		testManipulation();
	}
});
