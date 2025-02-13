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

		api.DocInfo = new Asc.asc_CDocInfo();
		var userInfo = new Asc.asc_CUserInfo();
		userInfo.asc_putId("user3");
		api.DocInfo.put_UserInfo(userInfo);
	}

	function create(ref, name, users) {
		let oUsers;
		for (let i = 0; i < users.length; i++) {
			if (!oUsers) {
				oUsers = [];
			}
			let oUser = new Asc.CUserProtectedRangeUserInfo();
			oUser.id = users[i].id;
			oUsers.push(oUser);
		}

		let obj = new Asc.CUserProtectedRange(ws);
		obj.asc_setRef(ref);
		obj.asc_setName(name);
		if (oUsers) {
			obj.asc_setUsers(oUsers);
		}
		api.asc_addUserProtectedRange(obj);
		return obj;
	}

	function testCreate() {
		QUnit.test("Test: create", function (assert) {
			//ADD
			create("B2:B5", "test", [{id: "user3"}]);

			AscCommon.History.Undo();
			assert.strictEqual(ws.userProtectedRanges.length, 0, "undo add test");
			AscCommon.History.Redo();
			assert.strictEqual(ws.userProtectedRanges[0].asc_getName(), "test", "name compare");
			assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", "ref compare");

			create("D2:E5", "test2", [{id: "user3"}]);
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
			create("B2:B5", "test1", [{id: "user3"}]);

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

	function checkUndoRedo(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	}

	function testManipulationRange() {
		QUnit.test("Test: change", function (assert) {
			create("B2:B5", "test1", [{id: "user3"}]);
			create("D2:E5", "test2", [{id: "user3"}]);

			let beforeFunc = function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
			};

			wsview.setSelection(new Asc.Range(0, 0, 0, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertColumns);
			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$C$2:$C$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$E$2:$F$5", desc + "_val_2");
			}, "insert_1");

			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertColumns);
			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$F$5", desc + "_val_2");
			}, "insert_2");

			wsview.setSelection(new Asc.Range(0, 1, 2, 4));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftRight);
			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$E$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$G$2:$H$5", desc + "_val_2");
			}, "insert_3");

			wsview.setSelection(new Asc.Range(0, 1, 2, 3));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftRight);
			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
			}, "insert_4");


			wsview.setSelection(new Asc.Range(0, 0, 3, 0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftDown);
			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$3:$B$6", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
			}, "insert_5");

			//delete cells
			wsview.setSelection(new Asc.Range(1, 0, 3, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteColumns);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 1, desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, "delete_1");


			wsview.setSelection(new Asc.Range(0, 1, 2, 4));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftLeft);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 1, desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$A$2:$B$5", desc + "_val_2");
			}, "delete_2");

			wsview.setSelection(new Asc.Range(0, 0, 3, 0));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftTop);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$1:$B$4", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
			}, "delete_3");

			wsview.setSelection(new Asc.Range(0, 0, 6, 10));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftTop);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 0, "delete columns ref compare1", desc + "_val_1");
			}, "delete_4");

			wsview.setSelection(new Asc.Range(0, 0, 6, 10));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteColumns);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 0, desc + "_val_1");
			}, "delete_5");

			wsview.setSelection(new Asc.Range(0, 0, 6, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteColumns);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 0, desc + "_val_1");
			});

			wsview.moveRangeHandle(ws.getRange2("D2:E5").bbox, ws.getRange2("D10:E13").bbox);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$10:$E$13", desc + "_val_2");
			}, "move_1");

			wsview.moveRangeHandle(ws.getRange2("D2:E5").bbox, ws.getRange2("D10:E13").bbox, true);
			checkUndoRedo(function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_2");
			}, function (desc){
				assert.strictEqual(ws.userProtectedRanges.length, 3, desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
				assert.strictEqual(ws.userProtectedRanges[2].asc_getRef(), "=Sheet1!$D$10:$E$13", desc + "_val_3");
			}, "move_2");

			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0]]);
			api.asc_deleteUserProtectedRange([ws.userProtectedRanges[0]]);
		});
	}

	function testCheckProtect() {
		QUnit.test("Test: change_protect", function (assert) {
			ws.getRange2("A1").setValue("1");
			ws.getRange2("A2").setValue("3");
			ws.getRange2("E3").setValue("1");
			ws.getRange2("E4").setValue("3");
			create("B2:B5", "test1", ["user1"]);
			create("D2:E5", "test2", ["user2"]);

			let beforeFunc = function(desc) {
				assert.strictEqual(ws.userProtectedRanges[0].asc_getRef(), "=Sheet1!$B$2:$B$5", desc + "_val_1");
				assert.strictEqual(ws.userProtectedRanges[1].asc_getRef(), "=Sheet1!$D$2:$E$5", desc + "_val_2");
			};

			//try change. intersection with protected ranges
			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertColumns);
			beforeFunc("check_insert_1");

			wsview.setSelection(new Asc.Range(4, 0, 4, 5));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftRight);
			beforeFunc("check_insert_2");

			wsview.setSelection(new Asc.Range(4, 0, 4, 5));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertCellsAndShiftDown);
			beforeFunc("check_insert_3");

			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscInsertOptions.InsertRows);
			beforeFunc("check_insert_4");

			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteColumns);
			beforeFunc("check_delete_1");

			wsview.setSelection(new Asc.Range(4, 0, 4, 5));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftLeft);
			beforeFunc("check_delete_2");

			wsview.setSelection(new Asc.Range(4, 0, 4, 5));
			wsview.changeWorksheet("delCell", Asc.c_oAscDeleteOptions.DeleteCellsAndShiftTop);
			beforeFunc("check_delete_3");

			wsview.setSelection(new Asc.Range(4, 0, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("insCell", Asc.c_oAscDeleteOptions.DeleteRows);
			beforeFunc("check_delete_4");


			//next actions must be interrupted and  actions will not add into history
			//try change ws
			let historyPointsLength = History.Points.length;
			wsview.setSelection(new Asc.Range(4, 3, 4, AscCommon.gc_nMaxRow0));
			wsview.changeWorksheet("colWidth", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_1");
			wsview.changeWorksheet("showCols", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_2");
			wsview.changeWorksheet("hideCols", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_3");
			wsview.changeWorksheet("rowHeight", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_4");
			wsview.changeWorksheet("showRows", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_5");
			wsview.changeWorksheet("hideRows", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_6");
			wsview.changeWorksheet("groupRows", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_7");
			wsview.changeWorksheet("groupCols", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_8");
			wsview.changeWorksheet("clearOutline", 12);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_9");

			//try change cell value
			wsview.setSelection(ws.getRange2("E3:E4").bbox);
			api.asc_insertInCell("SUM", Asc.c_oAscPopUpSelectorType.Func, true);
			api.wb._checkStopCellEditorInFormulas();
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_11");

			//try change cell settings
			wsview.setSelection(ws.getRange2("E3:E4").bbox);
			api.asc_setCellBold(true);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_12");

			api.asc_setCellItalic(true);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_13");

			api.asc_setCellUnderline(true);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_14");

			api.asc_setCellStrikeout(true);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_15");

			api.asc_setCellSuperscript(true);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_16");

			api.asc_increaseFontSize();
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_16");

			//try add filter
			/*wsview.setSelection(ws.getRange2("A1:A2").bbox);
			api.asc_addAutoFilter();
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_16");*/

			//move
			wsview.moveRangeHandle(ws.getRange2("D2:E5").bbox, ws.getRange2("D10:E13").bbox);
			assert.strictEqual(historyPointsLength, History.Points.length, "history_test_16");

			AscCommon.History.Undo();
			AscCommon.History.Undo();
		});
	}

	QUnit.module("UserProtectedRanges");

	function startTests() {
		QUnit.start();

		testCreate();
		testChange();
		testManipulationRange();
		testCheckProtect();
	}
});
