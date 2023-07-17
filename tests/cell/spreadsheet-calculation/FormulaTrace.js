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

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
	};

	let g_oIdCounter = AscCommon.g_oIdCounter;

	let wb, ws, ws2, sData = AscCommon.getEmpty(), api;
	if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {
		Asc.spreadsheet_api.prototype._init = function() {
		};
		
		api = new Asc.spreadsheet_api({
			'id-view': 'editor_sdk'
		});

		api.FontLoader = {
			LoadDocumentFonts: function() {
				setTimeout(startTests, 0)
			}
		};

		let docInfo = new Asc.asc_CDocInfo();
		docInfo.asc_putTitle("TeSt.xlsx");
		api.DocInfo = docInfo;

		window["Asc"]["editor"] = api;
		AscCommon.g_oTableId.init();
		if (this.User) {
			g_oIdCounter.Set_UserId(this.User.asc_getId());
		}
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(AscCommon.getEmpty());	// this func set api.wbModel
		// api._openOnClient();
		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;

		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
		g_oIdCounter.Set_Load(false);

		let oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read(sData, wb);
		// ws = wb.getWorksheet(wb.getActive());
		ws = api.wbModel.aWorksheets[0];
		ws2 = api.wbModel.createWorksheet(0, "Sheet2");
		AscCommonExcel.getFormulasInfo();

		api.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);

	}

	let wsView = api.wb.getWorksheet();
	let traceManager = wsView.traceDependentsManager;
	let parserFormula = AscCommonExcel.parserFormula, oParser;

	function traceTests() {
		QUnit.test("Test: \"Base dependents test\"", function (assert) {
			// set active cell
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			// create cells with dependencies
			ws.getRange2("A1").setValue("1");
			ws.getRange2("B101").setValue("=A1");
			ws.getRange2("C101").setValue("=B101");
	
			// "click" on the button trace dependents
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			// check the object with dependency cell numbers for compliance
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
				C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1);
			
			// check A1 -> B101
			assert.strictEqual(traceManager._getDependents(A1Index, B101Index), 1);

			// check B101 -> C101
			assert.strictEqual(traceManager._getDependents(B101Index, C101Index), 1);
			
			// clear traces from canvas
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"Dependents\"", function (assert) {
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			ws.getRange2("A1").setValue("1");
			ws.getRange2("C10").setValue("=A1");
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				C10Index = AscCommonExcel.getCellIndex(ws.getRange2("C10").bbox.r1, ws.getRange2("C10").bbox.c1);

			ws.getRange2("A10").setValue("=A1:A2");
			ws.getRange2("A11").setValue("=A1:A2");
			let A10Index = AscCommonExcel.getCellIndex(ws.getRange2("A10").bbox.r1, ws.getRange2("A10").bbox.c1),
				A11Index = AscCommonExcel.getCellIndex(ws.getRange2("A11").bbox.r1, ws.getRange2("A11").bbox.c1);

			ws.getRange2("B101").setValue("=SUM(A1:B2)+I3:J4+B2");
			ws.getRange2("B102").setValue("=SUM(A1:B2)+I3:J4+B2");
			ws.getRange2("C101").setValue("=SUM(A1:B2)+I3:J4+B2");
			ws.getRange2("C102").setValue("=SUM(A1:B2)+I3:J4+B2");
			let B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
				B102Index = AscCommonExcel.getCellIndex(ws.getRange2("B102").bbox.r1, ws.getRange2("B102").bbox.c1),
				C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1),
				C102Index = AscCommonExcel.getCellIndex(ws.getRange2("C102").bbox.r1, ws.getRange2("C102").bbox.c1);

			ws.getRange2("E200").setValue("=C101:C102");
			ws.getRange2("E201").setValue("=C101:C102");
			let E200Index = AscCommonExcel.getCellIndex(ws.getRange2("E200").bbox.r1, ws.getRange2("E200").bbox.c1),
				E201Index = AscCommonExcel.getCellIndex(ws.getRange2("E201").bbox.r1, ws.getRange2("E201").bbox.c1);

			ws.getRange2("H200").setValue("=E200:E201");
			ws.getRange2("H201").setValue("=E200:E201");
			let H200Index = AscCommonExcel.getCellIndex(ws.getRange2("H200").bbox.r1, ws.getRange2("H200").bbox.c1),
				H201Index = AscCommonExcel.getCellIndex(ws.getRange2("H201").bbox.r1, ws.getRange2("H201").bbox.c1);

			// first "click"
			api.asc_TraceDependents();

			// C10
			assert.strictEqual(traceManager._getDependents(A1Index, C10Index), 1);

			// A10:A11
			assert.strictEqual(traceManager._getDependents(A1Index, A10Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, A11Index), 1);

			// B101:C102
			assert.strictEqual(traceManager._getDependents(A1Index, B101Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, B102Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C101Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C102Index), 1);

			// E200:E201
			assert.strictEqual(traceManager._getDependents(C101Index, E200Index), undefined);
			assert.strictEqual(traceManager._getDependents(C101Index, E201Index), undefined);

			// H200:H201
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), undefined);
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), undefined);

			// second "click"
			api.asc_TraceDependents();

			// E200:E201
			assert.strictEqual(traceManager._getDependents(C101Index, E200Index), 1);
			assert.strictEqual(traceManager._getDependents(C101Index, E201Index), 1);

			// H200:H201
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), undefined);
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), undefined);

			// third "click"
			api.asc_TraceDependents();

			// H200:H201
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), 1);
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), 1);
			
			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"External dependencies\"", function (assert) {
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			// external references
			ws.getRange2("A1").setValue("1");
			ws.getRange2("B1").setValue("=A1");
			ws2.getRange2("A1").setValue("=Sheet1!A1");
			ws2.getRange2("B1").setValue("=Sheet1!B1");
	
			api.asc_TraceDependents();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				A1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A1").bbox.r1, ws2.getRange2("A1").bbox.c1) + ";0",
				B1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("B1").bbox.r1, ws2.getRange2("B1").bbox.c1) + ";0"; 
			
			assert.strictEqual(traceManager._getDependents(A1Index, B1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1);
			assert.strictEqual(traceManager._getDependents(B1Index, B1ExternalIndex), undefined);

			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(B1Index, B1ExternalIndex), 1);
			
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"Base precedents test\"", function (assert) {
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			// create cells with dependencies
			ws.getRange2("A1").setValue("=B101");
			ws.getRange2("B101").setValue("=C101");
			ws.getRange2("C101").setValue("1");
	
			// "click" on the button trace precedents
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			// check the object with dependency cell numbers for compliance
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
				C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1);
			
			// check A1 <- B101
			assert.strictEqual(traceManager._getPrecedents(A1Index, B101Index), 1);

			// check B101 <- C101
			assert.strictEqual(traceManager._getPrecedents(B101Index, C101Index), 1);

			// clear traces from canvas
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Precedents\"", function (assert) {
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			ws.getRange2("A1").setValue("=Sheet2!A10:A11+I5:J6+C1+A10:A11+Sheet2!C3");
			ws.getRange2("C1").setValue("=Sheet2!A10:A11+Sheet2!C3");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				I5Index = AscCommonExcel.getCellIndex(ws.getRange2("I5").bbox.r1, ws.getRange2("I5").bbox.c1),
				A10Index = AscCommonExcel.getCellIndex(ws.getRange2("A10").bbox.r1, ws.getRange2("A10").bbox.c1),

				A10ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A10").bbox.r1, ws2.getRange2("A10").bbox.c1) + ";0",
				C3ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("C3").bbox.r1, ws2.getRange2("C3").bbox.c1) + ";0";

			// first "click"
			api.asc_TracePrecedents();
			
			// A1
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10Index), 1);
			assert.strictEqual(traceManager._getPrecedents(A1Index, I5Index), 1);
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), 1);
			assert.strictEqual(traceManager._getPrecedents(A1Index, C3ExternalIndex), 1);

			// C1
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, C3ExternalIndex), undefined);

			// second "click"
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), 1);
			assert.strictEqual(traceManager._getPrecedents(C1Index, C3ExternalIndex), 1);
			
			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"DefName tests\"", function (assert) {
			// debugger
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			wb.dependencyFormulas.addDefName("a", "Sheet1!$C$1:$D$2");
			// wb.dependencyFormulas.defNames.wb[name]

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1);
			ws.getRange2("A1").setValue("=a");

			let A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1),
				B4Index = AscCommonExcel.getCellIndex(ws.getRange2("B4").bbox.r1, ws.getRange2("B4").bbox.c1);


			let C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				C2Index = AscCommonExcel.getCellIndex(ws.getRange2("C2").bbox.r1, ws.getRange2("C2").bbox.c1),
				D2Index = AscCommonExcel.getCellIndex(ws.getRange2("D2").bbox.r1, ws.getRange2("D2").bbox.c1);
			ws.getRange2("C1").setValue("=C2");
			ws.getRange2("C2").setValue("2");
			ws.getRange2("D1").setValue("1");
			ws.getRange2("D2").setValue("=F1");

			let F1Index = AscCommonExcel.getCellIndex(ws.getRange2("F1").bbox.r1, ws.getRange2("F1").bbox.c1);

			// first show precedents click on A1
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), undefined);

			// second show precedents click on A1
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), 1);
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), 1);

			// third show precedents click on A1
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), 1);
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), 1);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
	
			// change selection to A3
			ws.selectionRange.ranges = [ws.getRange2("A3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A3").getBBox0().r1, ws.getRange2("A3").getBBox0().c1);

			let bbox = ws.getRange2("A3:B4").bbox;
			let cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();
			oParser.calculate();

			ws.getRange2("A3:B4").setValue("=a", undefined, undefined, bbox);	// ???

			// show precedents click on A3
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined);

			// change selection to A4
			ws.selectionRange.ranges = [ws.getRange2("A4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A4").getBBox0().r1, ws.getRange2("A4").getBBox0().c1);

			ws.getRange2("A4").setValue("=a");
			bbox = ws.getRange2("A4").bbox;
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			// show precedents click on A4
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined);

			// change selection to B3
			// ws.selectionRange.ranges = [ws.getRange2("B3").getBBox0()];
			// ws.selectionRange.setActiveCell(ws.getRange2("B3").getBBox0().r1, ws.getRange2("B3").getBBox0().c1);

			// bbox = ws.getRange2("B3").bbox;
			// cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			// oParser = new parserFormula("a", cellWithFormula, ws);
			// oParser.setArrayFormulaRef(bbox);
			// oParser.parse();
			// oParser.calculate();

			// // show precedents click on B3
			// api.asc_TracePrecedents();
			// assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined);

			// // change selection to B4
			// ws.selectionRange.ranges = [ws.getRange2("B4").getBBox0()];
			// ws.selectionRange.setActiveCell(ws.getRange2("B4").getBBox0().r1, ws.getRange2("B4").getBBox0().c1);

			// bbox = ws.getRange2("B4").bbox;
			// cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			// oParser = new parserFormula("a", cellWithFormula, ws);
			// oParser.setArrayFormulaRef(bbox);
			// oParser.parse();
			// oParser.calculate();

			// // show precedents click on B4
			// api.asc_TracePrecedents();
			// assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), 1);
			// assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), 1);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		// QUnit.test("Test: \"Shared tests\"", function (assert) {
		// });
		// QUnit.test("Test: \"Tables tests\"", function (assert) {
		// });
		QUnit.test("Test: \"Deletes tests\"", function (assert) {
			// ------------------- base precedents ------------------- //
			ws.selectionRange.ranges = [ws.getRange2("I1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("I1").getBBox0().r1, ws.getRange2("I1").getBBox0().c1);

			ws.getRange2("A1:J10").cleanAll();

			ws.getRange2("I1").setValue("=G1");
			ws.getRange2("G1").setValue("=E1+G4");
			ws.getRange2("G4").setValue("=I4");
			ws.getRange2("I4").setValue("=I3");
			ws.getRange2("I3").setValue("=H3");
			ws.getRange2("H3").setValue("1");
			ws.getRange2("E1").setValue("=C1+C4");
			ws.getRange2("C1").setValue("1");
			ws.getRange2("C4").setValue("2");

			let I1Index = AscCommonExcel.getCellIndex(ws.getRange2("I1").bbox.r1, ws.getRange2("I1").bbox.c1),
				G1Index = AscCommonExcel.getCellIndex(ws.getRange2("G1").bbox.r1, ws.getRange2("G1").bbox.c1),
				G4Index = AscCommonExcel.getCellIndex(ws.getRange2("G4").bbox.r1, ws.getRange2("G4").bbox.c1),
				I4Index = AscCommonExcel.getCellIndex(ws.getRange2("I4").bbox.r1, ws.getRange2("I4").bbox.c1),
				I3Index = AscCommonExcel.getCellIndex(ws.getRange2("I3").bbox.r1, ws.getRange2("I3").bbox.c1),
				H3Index = AscCommonExcel.getCellIndex(ws.getRange2("H3").bbox.r1, ws.getRange2("H3").bbox.c1),
				E1Index = AscCommonExcel.getCellIndex(ws.getRange2("E1").bbox.r1, ws.getRange2("E1").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				C4Index = AscCommonExcel.getCellIndex(ws.getRange2("C4").bbox.r1, ws.getRange2("C4").bbox.c1);

			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1);

			// first clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1);

			// second clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1);

			// third clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined);

			// fourth clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined);

			// fifth clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined);

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base precedents with external -------------------//
			ws2.getRange2("B1").setValue("1");
			ws.getRange2("I1").setValue("=G1+Sheet2!B1");

			let B1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("B1").bbox.r1, ws2.getRange2("B1").bbox.c1) + ";" + ws2.getIndex();

			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I1Index, B1ExternalIndex), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1);

			// first clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(I1Index, B1ExternalIndex), 1);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined);

			// second clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(I1Index, B1ExternalIndex), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined);

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base dependents ---------------------------------//
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			ws.getRange2("A1:J10").cleanAll();

			ws.getRange2("A1").setValue("1");
			ws.getRange2("C1").setValue("=A1");
			ws.getRange2("C4").setValue("=A1");
			ws.getRange2("E1").setValue("=C1");
			ws.getRange2("E4").setValue("=C4");
			ws.getRange2("G1").setValue("=E1+E4");
			ws.getRange2("G4").setValue("=E4");
			ws.getRange2("F6").setValue("=G4");
			ws.getRange2("H6").setValue("=G4");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				E4Index = AscCommonExcel.getCellIndex(ws.getRange2("E4").bbox.r1, ws.getRange2("E4").bbox.c1),
				F6Index = AscCommonExcel.getCellIndex(ws.getRange2("F6").bbox.r1, ws.getRange2("F6").bbox.c1),
				H6Index = AscCommonExcel.getCellIndex(ws.getRange2("H6").bbox.r1, ws.getRange2("H6").bbox.c1);

			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1);
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), 1);
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), 1);
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), 1);
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), 1);

			// first clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1);
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), 1);
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), 1);
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined);

			// second clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1);
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined);
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined);

			// third clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined);
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined);
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined);

			// fourth clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), undefined);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined);
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined);
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined);
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined);

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base dependents with external -------------------//
			ws2.getRange2("A1").setValue("=Sheet1!A1");

			let A1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A1").bbox.r1, ws2.getRange2("A1").bbox.c1) + ";" + ws2.getIndex();

			api.asc_TraceDependents();
			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1);

			// first clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1);
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined);

			// second clear
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), undefined);
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), undefined);
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), undefined);
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined);
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined);


			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Mixed tests\"", function (assert) {
			// TODO check formulas
			ws.getRange2("A1").setValue("=Sheet2!A10+12");
			ws.getRange2("B1").setValue("=Sheet2!A10+A1");
			ws.getRange2("C1").setValue("=Sheet2!A10+B1");
			ws2.getRange2("A1").setValue("=Sheet1!C1");
			// ws.getRange2("A1").setValue("=Sheet2!A10:A11+I5:J6+C1+A10:A11+Sheet2!C3");
			// ws.getRange2("C1").setValue("=Sheet2!A10:A11+Sheet2!C3");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),

				A1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A1").bbox.r1, ws2.getRange2("A1").bbox.c1) + ";0",
				A10ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A10").bbox.r1, ws2.getRange2("A10").bbox.c1) + ";0";

			ws.selectionRange.ranges = [ws.getRange2("B1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B1").getBBox0().r1, ws.getRange2("B1").getBBox0().c1);

			// trace precedents
			api.asc_TracePrecedents();
			// A1
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), undefined);
			// B1
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(B1Index, A10ExternalIndex), 1);
			// C1
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A1ExternalIndex), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), undefined);

			// trace precedents
			api.asc_TracePrecedents();
			// A1
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), 1);
			// B1
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(B1Index, A10ExternalIndex), 1);
			// C1
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A1ExternalIndex), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), undefined);
			
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			// A1
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), 1);
			// B1
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(B1Index, A10ExternalIndex), 1);
			// C1
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A1ExternalIndex), undefined);
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), undefined);

			// trace dependents
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			// A1
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), 1);
			// B1
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1);
			assert.strictEqual(traceManager._getPrecedents(B1Index, A10ExternalIndex), 1);
			// C1
			assert.strictEqual(traceManager._getDependents(C1Index, A1ExternalIndex), 1);
			// assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), 1);		// 1
			assert.strictEqual(traceManager._getDependents(B1Index, C1Index), 1);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("C1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C1").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			api.asc_TracePrecedents();

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			api.asc_TraceDependents();
			api.asc_TraceDependents();

			// A1
			assert.strictEqual(traceManager._getDependents(A1Index, B1Index), 1);
			// C1
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), 1);
			// assert.strictEqual(traceManager._getDependents(C1Index, A1ExternalIndex), 1);	// ?
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), 1);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
	}

	QUnit.module("FormulaTrace");

	function startTests() {
		QUnit.start();
		traceTests();
	}

	startTests();
});
