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

$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};


	var sData = AscCommon.getEmpty();
	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api._openDocument(sData);
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

	QUnit.test("Test: \"simple tests\"", function (assert) {

		ws.getRange2("A1").setValue("-4");

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValue(), ws.getRange2("A1").getValue());

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B7").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B8").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B9").getValue(), "-4");
	});

	QUnit.test("Test: \"formula tests\"", function (assert) {
		var val = "=SIN(1)";

		ws.getRange2("A1").setValue(val);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValueForEdit(), ws.getRange2("A1").getValueForEdit());

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B6").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B7").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B8").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B9").getValueForEdit(), val);


		var val1 = "=SIN(A2)";
		var val2 = "=SIN(A3)";

		ws.getRange2("A1").setValue(val1);
		ws.getRange2("A2").setValue(val2);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 1)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(2, 1, 2, 6), getRange(3, 5, 3, 8), getRange(4, 5, 4, 7)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("C2").getValueForEdit(), "=SIN(C3)");
		assert.strictEqual(ws.getRange2("C3").getValueForEdit(), "=SIN(C4)");
		assert.strictEqual(ws.getRange2("C4").getValueForEdit(), "=SIN(C5)");
		assert.strictEqual(ws.getRange2("C5").getValueForEdit(), "=SIN(C6)");
		assert.strictEqual(ws.getRange2("C6").getValueForEdit(), "=SIN(C7)");
		assert.strictEqual(ws.getRange2("C7").getValueForEdit(), "=SIN(C8)");

		assert.strictEqual(ws.getRange2("D6").getValueForEdit(), "=SIN(D7)");
		assert.strictEqual(ws.getRange2("D7").getValueForEdit(), "=SIN(D8)");
		assert.strictEqual(ws.getRange2("D8").getValueForEdit(), "=SIN(D9)");
		assert.strictEqual(ws.getRange2("D9").getValueForEdit(), "=SIN(D10)");

		assert.strictEqual(ws.getRange2("E6").getValueForEdit(), "=SIN(E7)");
		assert.strictEqual(ws.getRange2("E7").getValueForEdit(), "=SIN(E8)");
		assert.strictEqual(ws.getRange2("E8").getValueForEdit(), "");
	});

	QUnit.test("Test: \"comment tests\"", function (assert) {

		ws.getRange2("E10").setValue("-4");
		var comment = new  window["Asc"].asc_CCommentData(null);
		comment.asc_putText("test");
		comment.bDocument = false;
		/*comment.asc_putTime(this.utcDateToString(new Date()));
		comment.asc_putOnlyOfficeTime(this.ooDateToString(new Date()));
		comment.asc_putUserId(this.currentUserId);
		comment.asc_putUserName(this.currentUserName);
		comment.asc_putSolved(false);*/
		api.asc_addComment(comment);
		comment.nCol = 4;
		comment.nRow = 9;
		comment.coords.nCol = 4;
		comment.coords.nRow = 9;

		ws.selectionRange.ranges = [getRange(4, 9, 4, 9)];

		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValue(), ws.getRange2("E10").getValue());
		assert.strictEqual(wsView.cellCommentator.getComment(4,9).nCol, 4);

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B7").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B8").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B9").getValue(), "-4");

		assert.strictEqual(wsView.cellCommentator.getComment(0,5).nRow, 5);
		assert.strictEqual(wsView.cellCommentator.getComment(1,5).nRow, 5);
		assert.strictEqual(wsView.cellCommentator.getComment(1,6).nRow, 6);
		assert.strictEqual(wsView.cellCommentator.getComment(1,7).nRow, 7);
		assert.strictEqual(wsView.cellCommentator.getComment(1,8).nRow, 8);

		assert.strictEqual(wsView.cellCommentator.getComment(1,9), null);
	});

	QUnit.test("Test: \"tables\"", function (assert) {
		ws.autoFilters.addAutoFilter("TableStyleMedium2", getRange(3, 5, 3, 8));

		ws.selectionRange.ranges = [getRange(3, 5, 3, 9)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(4, 10, 4, 10)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.c1, 4);

		ws.selectionRange.ranges = [getRange(5, 10, 5, 10), getRange(6, 10, 6, 10)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 2].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 2].Ref.c1, 5);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.c1, 6)
	});

	QUnit.module("CopyPaste");
});
