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
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
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
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});

	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};

	QUnit.test("Test: \"changeTextCase\"", function (assert) {

		let generateMultiText = function (arr) {
			let _res = [];
			for (let i = 0; i < arr.length; i++) {
				let multiElem1 = new AscCommonExcel.CMultiTextElem();
				multiElem1.text = arr[i].t;
				multiElem1.format = arr[i].f;
				_res.push(multiElem1);
			}
			return _res;
		};

		function checkUndoRedo(_before, _after, _desc) {
			assert.strictEqual(ws.getRange2("A1").getValue(), _after, _desc + " _before_undo");
			AscCommon.History.Undo();
			assert.strictEqual(ws.getRange2("A1").getValue(), _before, _desc + "_undo");
			AscCommon.History.Redo();
			assert.strictEqual(ws.getRange2("A1").getValue(), _after, _desc + "_redo");
			AscCommon.History.Undo();
			assert.strictEqual(ws.getRange2("A1").getValue(), _before, _desc + "_lastUndo");
		}

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];

		let aMultiText = generateMultiText([{t: 'te'}, {t: 'st TES'}, {t: 'T'}, {t: ' t'}, {t: 'Es'}, {t: 't  Te'}, {t: 'st\nt'},
			{t: 'Est te'}, {t: 's'}, {t: 't   Tee'}, {t: 'est '}, {t: 'tesT'}, {t: '\n', },
			{t: 'TEST te', }, {t: 'st Test', }]);


		let cellValue = new AscCommonExcel.CCellValue({multiText: aMultiText});
		ws.getRange2("A1").setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, cellValue));

		let trueResult = "test TEST tEst  Test\ntEst test   Teeest tesT\nTEST test Test";
		assert.strictEqual(ws.getRange2("A1").getValue(), trueResult);

		//LowerCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.LowerCase);
		let result = "test test test  test\n" +
			"test test   teeest test\n" +
			"test test test";
		checkUndoRedo(trueResult, result, "LowerCase");

		//UpperCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.UpperCase);
		result = "TEST TEST TEST  TEST\n" +
			"TEST TEST   TEEEST TEST\n" +
			"TEST TEST TEST";
		checkUndoRedo(trueResult, result, "UpperCase");

		//ToggleCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.ToggleCase);
		result = "TEST test TeST  tEST\n" +
			"TeST TEST   tEEEST TESt\n" +
			"test TEST tEST";
		checkUndoRedo(trueResult, result, "ToggleCase");

		//CapitalizeWords
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.CapitalizeWords);
		result = "Test TEST Test  Test\n" +
			"Test Test   Teeest Test\n" +
			"TEST Test Test";
		checkUndoRedo(trueResult, result, "CapitalizeWords");

		//SentenceCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.SentenceCase);
		result = "Test TEST test  Test\n" +
			"Test test   Teeest test\n" +
			"TEST test Test";
		checkUndoRedo(trueResult, result, "SentenceCase");

		aMultiText = generateMultiText([{t: 'te'}, {t: 'st TE'}, {t: 'ST tEst  T'}, {t: 'est\ntE'}, {t: 'st. test   TeE'}, {t: 'Est. te'}, {t: 'ST\nTEST te'},
			{t: 'st Test teEEst\ntes'}, {t: 't.test\ntest,test'}, {t: ';test,tEst\\tes'}, {t: 't\nteSt T'}, {t: 'est Test TESt TES'}, {t: 'T tesT'}]);


		cellValue = new AscCommonExcel.CCellValue({multiText: aMultiText});
		ws.getRange2("A1").setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, cellValue));

		trueResult = "test TEST tEst  Test\ntEst. test   TeEEst. teST\nTEST test Test teEEst\ntest.test\ntest,test;test,tEst\\test\nteSt Test Test TESt TEST tesT";

		assert.strictEqual(ws.getRange2("A1").getValue(), trueResult);

		//LowerCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.LowerCase);
		result = "test test test  test\n" +
			"test. test   teeest. test\n" +
			"test test test teeest\n" +
			"test.test\n" +
			"test,test;test,test\\test\n" +
			"test test test test test test";
		checkUndoRedo(trueResult, result, "LowerCase2");

		//UpperCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.UpperCase);
		result = "TEST TEST TEST  TEST\n" +
			"TEST. TEST   TEEEST. TEST\n" +
			"TEST TEST TEST TEEEST\n" +
			"TEST.TEST\n" +
			"TEST,TEST;TEST,TEST\\TEST\n" +
			"TEST TEST TEST TEST TEST TEST";
		checkUndoRedo(trueResult, result, "UpperCase2");

		//ToggleCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.ToggleCase);
		result = "TEST test TeST  tEST\n" +
			"TeST. TEST   tEeeST. TEst\n" +
			"test TEST tEST TEeeST\n" +
			"TEST.TEST\n" +
			"TEST,TEST;TEST,TeST\\TEST\n" +
			"TEsT tEST tEST tesT test TESt";
		checkUndoRedo(trueResult, result, "ToggleCase2");

		//CapitalizeWords
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.CapitalizeWords);
		result = "Test TEST Test  Test\n" +
			"Test. Test   Teeest. Test\n" +
			"TEST Test Test Teeest\n" +
			"Test.Test\n" +
			"Test,Test;Test,Test\\Test\n" +
			"Test Test Test Test TEST Test";
		checkUndoRedo(trueResult, result, "CapitalizeWords2");

		//SentenceCase
		api.asc_ChangeTextCase(Asc.c_oAscChangeTextCaseType.SentenceCase);
		result = "Test TEST test  Test\n" +
			"Test. Test   teeest. Test\n" +
			"TEST test Test teeest\n" +
			"Test.Test\n" + //.T -> .t
			"Test,test;test,test\\test\n" +
			"Test Test Test test TEST test";
		checkUndoRedo(trueResult, result, "SentenceCase2");

	});

	QUnit.module("CopyPaste");
});
