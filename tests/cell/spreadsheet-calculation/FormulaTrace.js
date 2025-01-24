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
		var self = this;
		this.model.handlers.add("changeDocument", function(prop, arg1, arg2, wsId) {
			let ws = wsId && self.getWorksheetById(wsId, true);
			if (ws) {
				ws.traceDependentsManager.changeDocument(prop, arg1, arg2);
			}
		});
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
		api.initCollaborativeEditing({});
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

		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);

	}

	let wsView = api.wb.getWorksheet();
	let traceManager = wsView.traceDependentsManager;
	let parserFormula = AscCommonExcel.parserFormula, oParser;

	function traceTests() {
		// TODO dependent perfomance, dependents delete perfomance, simple precedents delete prefomance
		QUnit.test("Test: \"Base dependents test\"", function (assert) {
			ws.getRange2("A1:J20").cleanAll();

			// set active cell
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			// create cells with dependencies
			ws.getRange2("A1").setValue("1");
			ws.getRange2("B101").setValue("=A1");
			ws.getRange2("C101").setValue("=B101");
	
			// "click" on the button trace dependents
			assert.ok(1, "Trace dependents from A1, two times");
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			// check the object with dependency cell numbers for compliance
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
				C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1);
			
			assert.strictEqual(traceManager._getDependents(A1Index, B101Index), 1, "A1->B101");
			assert.strictEqual(traceManager._getDependents(B101Index, C101Index), 1, "B101->C101");
			
			// clear traces from canvas
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"Dependents\"", function (assert) {
			ws.getRange2("A1:J20").cleanAll();
			
			ws.getRange2("A1").setValue("1");
			ws.getRange2("C10").setValue("=A1");
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				C10Index = AscCommonExcel.getCellIndex(ws.getRange2("C10").bbox.r1, ws.getRange2("C10").bbox.c1);

			// ws.getRange2("D3:D5").setValue("=Table1", undefined, undefined, bbox);
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

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace dependents from A1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A1Index, C10Index), 1, "A1->C10");
			assert.strictEqual(traceManager._getDependents(A1Index, A10Index), 1, "A1->A10");
			assert.strictEqual(traceManager._getDependents(A1Index, A11Index), 1, "A1->A11");
			assert.strictEqual(traceManager._getDependents(A1Index, B101Index), 1, "A1->B101");
			assert.strictEqual(traceManager._getDependents(A1Index, B102Index), 1, "A1->B102");
			assert.strictEqual(traceManager._getDependents(A1Index, C101Index), 1, "A1->C101");
			assert.strictEqual(traceManager._getDependents(A1Index, C102Index), 1, "A1->C102");
			assert.strictEqual(traceManager._getDependents(C101Index, E200Index), undefined, "C101->E200 === undefined");
			assert.strictEqual(traceManager._getDependents(C101Index, E201Index), undefined, "C101->E201 === undefined");
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), undefined, "E200->H200 === undefined");
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), undefined, "E200->H201 === undefined");

			assert.ok(1, "Trace dependents from A1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C101Index, E200Index), 1, "C101->E200");
			assert.strictEqual(traceManager._getDependents(C101Index, E201Index), 1, "C101->E201");
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), undefined, "E200->H200 === undefined");
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), undefined, "E200->H201 === undefined");

			assert.ok(1, "Trace dependents from A1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(E200Index, H200Index), 1, "E200->H200");
			assert.strictEqual(traceManager._getDependents(E200Index, H201Index), 1, "E200->H201");
			
			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"External dependencies\"", function (assert) {
			ws.getRange2("A1:J20").cleanAll();

			ws.getRange2("A1").setValue("1");
			ws.getRange2("B1").setValue("=A1");
			// external references
			ws2.getRange2("A1").setValue("=Sheet1!A1");
			ws2.getRange2("B1").setValue("=Sheet1!B1");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				A1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A1").bbox.r1, ws2.getRange2("A1").bbox.c1) + ";0",
				B1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("B1").bbox.r1, ws2.getRange2("B1").bbox.c1) + ";0"; 
	
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace dependents from A1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A1Index, B1Index), 1, "A1->B1");
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1, "A1->OtherSheet!A1");
			assert.strictEqual(traceManager._getDependents(B1Index, B1ExternalIndex), undefined, "B1->OtherSheet!B1 === undefined");

			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(B1Index, B1ExternalIndex), 1, "B1->OtherSheet!B1");
			
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"Base precedents test\"", function (assert) {
			ws.getRange2("A1:Z200").cleanAll();

			// create cells with dependencies
			ws.getRange2("A1").setValue("=B101");
			ws.getRange2("B101").setValue("=C101");
			ws.getRange2("C101").setValue("1");

			// check the object with dependency cell numbers for compliance
			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
			B101Index = AscCommonExcel.getCellIndex(ws.getRange2("B101").bbox.r1, ws.getRange2("B101").bbox.c1),
			C101Index = AscCommonExcel.getCellIndex(ws.getRange2("C101").bbox.r1, ws.getRange2("C101").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace precedents from A1");
			// "click" on the button trace precedents
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, B101Index), 1, "A1<-B101");
			assert.strictEqual(traceManager._getPrecedents(B101Index, C101Index), 1, "B101<-C101");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Precedents\"", function (assert) {
			let wsName = ws.getName();
			ws.getRange2("A1:J100").cleanAll();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1);
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				I5Index = AscCommonExcel.getCellIndex(ws.getRange2("I5").bbox.r1, ws.getRange2("I5").bbox.c1),
				A10Index = AscCommonExcel.getCellIndex(ws.getRange2("A10").bbox.r1, ws.getRange2("A10").bbox.c1),
				A10ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A10").bbox.r1, ws2.getRange2("A10").bbox.c1) + ";0",
				C3ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("C3").bbox.r1, ws2.getRange2("C3").bbox.c1) + ";0";

			if (wsName) {
				ws.getRange2("A3").setValue("=SUM(" + wsName + "!A1," + wsName + "!A2)");
				
				ws.selectionRange.ranges = [ws.getRange2("A3").getBBox0()];
				ws.selectionRange.setActiveCell(ws.getRange2("A3").getBBox0().r1, ws.getRange2("A3").getBBox0().c1);
	
				assert.ok(1, "Trace precedents from A3");
				api.asc_TracePrecedents();
				assert.strictEqual(traceManager._getPrecedents(A3Index, A1Index), 1, "A3<-A1");
				assert.strictEqual(traceManager._getPrecedents(A3Index, A2Index), 1, "A3<-A2");
		
				// clear traces
				api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
			}


			ws.getRange2("A1").setValue("=Sheet2!A10:A11+I5:J6+C1+A10:A11+Sheet2!C3");
			ws.getRange2("C1").setValue("=Sheet2!A10:A11+Sheet2!C3");
			
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace precedents from A1");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), 1, "A1<-C1");
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10Index), "A10:A11", "A1<-A10");
			assert.strictEqual(traceManager._getPrecedents(A1Index, I5Index), "I5:J6", "A1<-I5");
			if (traceManager.checkPrecedentExternal(A1Index)) {
				assert.ok(traceManager.precedentsExternal[A1Index][A10ExternalIndex], "A1<-OtherSheet!A10");
				assert.ok(traceManager.precedentsExternal[A1Index][C3ExternalIndex], "A1<-OtherSheet!C3");

				assert.strictEqual(traceManager.checkPrecedentExternal(C1Index), undefined, "From C1 external precedent don't exist");
			}

			api.asc_TracePrecedents();
			if (traceManager.checkPrecedentExternal(C1Index)) {
				assert.ok(traceManager.precedentsExternal[C1Index][A10ExternalIndex], "C1<-OtherSheet!A10");
				assert.ok(traceManager.precedentsExternal[C1Index][C3ExternalIndex], "C1<-OtherSheet!C3");
			}
			
			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"DefName tests\"", function (assert) {
			ws.getRange2("A1:J20").cleanAll();

			wb.dependencyFormulas.addDefName("a", "Sheet1!$C$1:$D$2");
			// wb.dependencyFormulas.defNames.wb[name]

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1),
				B4Index = AscCommonExcel.getCellIndex(ws.getRange2("B4").bbox.r1, ws.getRange2("B4").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				C2Index = AscCommonExcel.getCellIndex(ws.getRange2("C2").bbox.r1, ws.getRange2("C2").bbox.c1),
				D1Index = AscCommonExcel.getCellIndex(ws.getRange2("D1").bbox.r1, ws.getRange2("D1").bbox.c1),
				D2Index = AscCommonExcel.getCellIndex(ws.getRange2("D2").bbox.r1, ws.getRange2("D2").bbox.c1),
				F1Index = AscCommonExcel.getCellIndex(ws.getRange2("F1").bbox.r1, ws.getRange2("F1").bbox.c1);

			ws.getRange2("A1").setValue("=a");
			ws.getRange2("C1").setValue("=C2");
			ws.getRange2("C2").setValue("2");
			ws.getRange2("D1").setValue("1");
			ws.getRange2("D2").setValue("=F1");

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace precedents from A1");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), "$C$1:$D$2", "A1<-C1");
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), undefined, "C1<-C2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), undefined, "D2<-F1 === undefined");

			assert.ok(1, "Trace precedents from A1");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), "$C$1:$D$2", "A1<-C1");
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), 1, "C1<-C2");
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), 1, "D2<-F1");

			assert.ok(1, "Trace precedents from A1");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), "$C$1:$D$2", "A1<-C1");
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), 1, "C1<-C2");
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), 1, "D2<-F1");

			// clear traces
			assert.ok(1, "Clear all traces");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
	
			// change selection to A3
			ws.selectionRange.ranges = [ws.getRange2("A3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A3").getBBox0().r1, ws.getRange2("A3").getBBox0().c1);
			let bbox = ws.getRange2("A3:B4").bbox;
			ws.getRange2("A3:B4").setValue("=a", undefined, undefined, bbox);

			assert.ok(1, "Trace precedents from A3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), "$C$1:$D$2", "A3<-C1");
			assert.strictEqual(traceManager._getPrecedents(A1Index, C1Index), undefined, "A1<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), undefined, "B3<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined, "B4<-C1 === undefined");

			// change selection to A4
			ws.selectionRange.ranges = [ws.getRange2("A4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A4").getBBox0().r1, ws.getRange2("A4").getBBox0().c1);

			ws.getRange2("A4").setValue("=a");
			bbox = ws.getRange2("A4").bbox;
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			assert.ok(1, "Trace precedents from A4");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), "$C$1:$D$2", "A3<-C1");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), "$C$1:$D$2", "A4<-C1");
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), undefined, "B3<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined, "B4<-C1 === undefined");

			// change selection to B3
			ws.selectionRange.ranges = [ws.getRange2("B3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B3").getBBox0().r1, ws.getRange2("B3").getBBox0().c1);

			assert.ok(1, "Trace precedents from B3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), "$C$1:$D$2", "A3<-C1");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), "$C$1:$D$2", "A4<-C1");
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), "$C$1:$D$2", "B3<-C1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), undefined, "B4<-C1 === undefined");

			// change selection to B4
			ws.selectionRange.ranges = [ws.getRange2("B4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B4").getBBox0().r1, ws.getRange2("B4").getBBox0().c1);

			assert.ok(1, "Trace precedents from B4");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), "$C$1:$D$2", "A3<-C1");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), "$C$1:$D$2", "A4<-C1");
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), "$C$1:$D$2", "B3<-C1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), "$C$1:$D$2", "B4<-C1");

			assert.ok(1, "Trace precedents from B4");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A3Index, C1Index), "$C$1:$D$2", "A3<-C1");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), "$C$1:$D$2", "A4<-C1");
			assert.strictEqual(traceManager._getPrecedents(B3Index, C1Index), "$C$1:$D$2", "B3<-C1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, C1Index), "$C$1:$D$2", "B4<-C1");
			assert.strictEqual(traceManager._getPrecedents(C1Index, C2Index), 1, "C1<-C2");
			assert.strictEqual(traceManager._getPrecedents(D2Index, F1Index), 1, "D2<-F1");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			/* dependents tests */

			// change selection to C1
			ws.selectionRange.ranges = [ws.getRange2("C1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C1").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			bbox = ws.getRange2("A3:B4").bbox;
			ws.getRange2("A3:B4").setValue("=a", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			let F7Index = AscCommonExcel.getCellIndex(ws.getRange2("F7").bbox.r1, ws.getRange2("F7").bbox.c1);
			bbox = ws.getRange2("F7").bbox;
			ws.getRange2("F7").setValue("=a", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			let F9Index = AscCommonExcel.getCellIndex(ws.getRange2("F9").bbox.r1, ws.getRange2("F9").bbox.c1),
				G9Index = AscCommonExcel.getCellIndex(ws.getRange2("G9").bbox.r1, ws.getRange2("G9").bbox.c1);
			bbox = ws.getRange2("F9:G9").bbox;
			ws.getRange2("F9:G9").setValue("=a", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("a", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			assert.ok(1, "Trace dependents from C1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C1Index, A3Index), 1, "C1->A3");
			assert.strictEqual(traceManager._getDependents(C1Index, A4Index), 1, "C1->A4");
			assert.strictEqual(traceManager._getDependents(C1Index, B3Index), 1, "C1->B3");
			assert.strictEqual(traceManager._getDependents(C1Index, B4Index), 1, "C1->B4");
			// assert.strictEqual(traceManager._getDependents(C1Index, A1Index), undefined, "C1->A1 === undefined");	// todo add back .cross() for dependencies SINGLE(A1:A10)|SINGLE(A1:D10)
			assert.strictEqual(traceManager._getDependents(C1Index, F7Index), 1, "C1->F7");
			assert.strictEqual(traceManager._getDependents(C1Index, F9Index), 1, "C1->F9");
			assert.strictEqual(traceManager._getDependents(C1Index, G9Index), 1, "C1->G9");

			// change selection to D1
			ws.selectionRange.ranges = [ws.getRange2("D1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("D1").getBBox0().r1, ws.getRange2("D1").getBBox0().c1);

			assert.ok(1, "Trace dependents from D1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(D1Index, A3Index), 1, "D1->A3");
			assert.strictEqual(traceManager._getDependents(D1Index, A4Index), 1, "D1->A4");
			assert.strictEqual(traceManager._getDependents(D1Index, B3Index), 1, "D1->B3");
			assert.strictEqual(traceManager._getDependents(D1Index, B4Index), 1, "D1->B4");
			// assert.strictEqual(traceManager._getDependents(D1Index, A1Index), undefined, "D1->A1 === undefined");	// todo add back .cross() for dependencies SINGLE(A1:A10)|SINGLE(A1:D10)
			assert.strictEqual(traceManager._getDependents(D1Index, F7Index), undefined, "D1->F7 === undefined");
			assert.strictEqual(traceManager._getDependents(D1Index, F9Index), 1, "D1->F9");
			assert.strictEqual(traceManager._getDependents(D1Index, G9Index), 1, "D1->G9");

			// change selection to D2
			ws.selectionRange.ranges = [ws.getRange2("D2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("D2").getBBox0().r1, ws.getRange2("D2").getBBox0().c1);

			assert.ok(1, "Trace dependents from D2");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(D2Index, A3Index), 1, "D1->A3");
			assert.strictEqual(traceManager._getDependents(D2Index, A4Index), 1, "D1->A4");
			assert.strictEqual(traceManager._getDependents(D2Index, B3Index), 1, "D1->B3");
			assert.strictEqual(traceManager._getDependents(D2Index, B4Index), 1, "D1->B4");
			// assert.strictEqual(traceManager._getDependents(D2Index, A1Index), undefined, "D1->A1 === undefined");	// todo add back .cross() for dependencies SINGLE(A1:A10)|SINGLE(A1:D10)
			assert.strictEqual(traceManager._getDependents(D2Index, F7Index), undefined, "D1->F7 === undefined");
			assert.strictEqual(traceManager._getDependents(D2Index, F9Index), undefined, "D1->F9 === undefined");
			assert.strictEqual(traceManager._getDependents(D2Index, G9Index), undefined, "D1->G9 === undefined");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// defname in formulas
			// todo in future versions add a separation between the range in the formula and the range in SINGLE
			wb.dependencyFormulas.addDefName("OneCell", "Sheet1!$C$1");
			wb.dependencyFormulas.addDefName("TwoCellInARow", "Sheet1!$C$1:$D$1");
			// wb.dependencyFormulas.defNames.wb[name]

			let I3Index = AscCommonExcel.getCellIndex(ws.getRange2("I3").bbox.r1, ws.getRange2("I3").bbox.c1),
				I4Index = AscCommonExcel.getCellIndex(ws.getRange2("I4").bbox.r1, ws.getRange2("I4").bbox.c1),
				I5Index = AscCommonExcel.getCellIndex(ws.getRange2("I5").bbox.r1, ws.getRange2("I5").bbox.c1),
				I6Index = AscCommonExcel.getCellIndex(ws.getRange2("I6").bbox.r1, ws.getRange2("I6").bbox.c1);
			bbox = ws.getRange2("I3").bbox;
			ws.getRange2("I3").setValue("=IF(1,OneCell,0)", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(1,OneCell,0)", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			ws.getRange2("I4").setValue("=IF(1,OneCell,0)");
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(1,OneCell,0)", cellWithFormula, ws);
			oParser.parse();

			bbox = ws.getRange2("I5").bbox;
			ws.getRange2("I5").setValue("=IF(0,OneCell,0)", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(0,OneCell,0)", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			ws.getRange2("I6").setValue("=IF(0,OneCell,0)");
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(0,OneCell,0)", cellWithFormula, ws);
			oParser.parse();

			// change selection to C1
			ws.selectionRange.ranges = [ws.getRange2("C1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C1").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			assert.ok(1, "Trace dependents from C1. Defname in formula test.");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C1Index, I3Index), 1, "C1->I3");
			assert.strictEqual(traceManager._getDependents(C1Index, I4Index), 1, "C1->I4");
			assert.strictEqual(traceManager._getDependents(C1Index, I5Index), 1, "C1->I5");
			assert.strictEqual(traceManager._getDependents(C1Index, I6Index), 1, "C1->I6");


			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			let K3Index = AscCommonExcel.getCellIndex(ws.getRange2("K3").bbox.r1, ws.getRange2("K3").bbox.c1),
				K4Index = AscCommonExcel.getCellIndex(ws.getRange2("K4").bbox.r1, ws.getRange2("K4").bbox.c1),
				K5Index = AscCommonExcel.getCellIndex(ws.getRange2("K5").bbox.r1, ws.getRange2("K5").bbox.c1),
				K6Index = AscCommonExcel.getCellIndex(ws.getRange2("K6").bbox.r1, ws.getRange2("K6").bbox.c1);
			bbox = ws.getRange2("K3").bbox;
			ws.getRange2("K3").setValue("=IF(1,TwoCellInARow,0)", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(1,TwoCellInARow,0)", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			ws.getRange2("K4").setValue("=IF(1,TwoCellInARow,0)");
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(1,TwoCellInARow,0)", cellWithFormula, ws);
			oParser.parse();

			bbox = ws.getRange2("K5").bbox;
			ws.getRange2("K5").setValue("=IF(0,TwoCellInARow,0)", undefined, undefined, bbox);
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(0,TwoCellInARow,0)", cellWithFormula, ws);
			oParser.setArrayFormulaRef(bbox);
			oParser.parse();

			ws.getRange2("K6").setValue("=IF(0,TwoCellInARow,0)");
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("IF(0,TwoCellInARow,0)", cellWithFormula, ws);
			oParser.parse();

			assert.ok(1, "Trace dependents from C1. Defname in formula test 2.");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C1Index, K3Index), 1, "C1->K3");
			assert.strictEqual(traceManager._getDependents(C1Index, K4Index), 1, "C1->K4");
			assert.strictEqual(traceManager._getDependents(C1Index, K5Index), 1, "C1->K5");
			assert.strictEqual(traceManager._getDependents(C1Index, K6Index), 1, "C1->K6");


			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Areas tests\"", function (assert) {
			let bbox;
			// -------------- precedents --------------
			ws.getRange2("A1:Z40").cleanAll();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				A5Index = AscCommonExcel.getCellIndex(ws.getRange2("A5").bbox.r1, ws.getRange2("A5").bbox.c1),
				A6Index = AscCommonExcel.getCellIndex(ws.getRange2("A6").bbox.r1, ws.getRange2("A6").bbox.c1),
				A12Index = AscCommonExcel.getCellIndex(ws.getRange2("A12").bbox.r1, ws.getRange2("A12").bbox.c1),
				A20Index = AscCommonExcel.getCellIndex(ws.getRange2("A20").bbox.r1, ws.getRange2("A20").bbox.c1),
				A21Index = AscCommonExcel.getCellIndex(ws.getRange2("A21").bbox.r1, ws.getRange2("A21").bbox.c1),
				A22Index = AscCommonExcel.getCellIndex(ws.getRange2("A22").bbox.r1, ws.getRange2("A22").bbox.c1);

			ws.getRange2("A1").setValue("1");
			ws.getRange2("A2").setValue("2");
			ws.getRange2("A3").setValue("3");
			ws.getRange2("A4").setValue("=A12");
			ws.getRange2("A5").setValue("5");
			ws.getRange2("A6").setValue("6");
			ws.getRange2("A12").setValue("=B12");
			ws.getRange2("A20").setValue("2");
			ws.getRange2("A21").setValue("=A22");
			ws.getRange2("A22").setValue("=B22");

			let B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1),
				B4Index = AscCommonExcel.getCellIndex(ws.getRange2("B4").bbox.r1, ws.getRange2("B4").bbox.c1),
				B5Index = AscCommonExcel.getCellIndex(ws.getRange2("B5").bbox.r1, ws.getRange2("B5").bbox.c1),
				B6Index = AscCommonExcel.getCellIndex(ws.getRange2("B6").bbox.r1, ws.getRange2("B6").bbox.c1),
				B7Index = AscCommonExcel.getCellIndex(ws.getRange2("B7").bbox.r1, ws.getRange2("B7").bbox.c1),
				B8Index = AscCommonExcel.getCellIndex(ws.getRange2("B8").bbox.r1, ws.getRange2("B8").bbox.c1),
				B9Index = AscCommonExcel.getCellIndex(ws.getRange2("B9").bbox.r1, ws.getRange2("B9").bbox.c1),
				B10Index = AscCommonExcel.getCellIndex(ws.getRange2("B10").bbox.r1, ws.getRange2("B10").bbox.c1),
				B12Index = AscCommonExcel.getCellIndex(ws.getRange2("B12").bbox.r1, ws.getRange2("B12").bbox.c1),
				B13Index = AscCommonExcel.getCellIndex(ws.getRange2("B13").bbox.r1, ws.getRange2("B13").bbox.c1),
				B14Index = AscCommonExcel.getCellIndex(ws.getRange2("B14").bbox.r1, ws.getRange2("B14").bbox.c1),
				B15Index = AscCommonExcel.getCellIndex(ws.getRange2("B15").bbox.r1, ws.getRange2("B15").bbox.c1),
				B16Index = AscCommonExcel.getCellIndex(ws.getRange2("B16").bbox.r1, ws.getRange2("B16").bbox.c1),
				B22Index = AscCommonExcel.getCellIndex(ws.getRange2("B22").bbox.r1, ws.getRange2("B22").bbox.c1);

			ws.getRange2("B1").setValue("=A4");
			bbox = ws.getRange2("B3:B9").bbox;
			ws.getRange2("B3:B9").setValue("=A1:A6", undefined, undefined, bbox);
			ws.getRange2("B10").setValue("=E6");
			ws.getRange2("B12").setValue("=B13");
			ws.getRange2("B13").setValue("=B14");
			ws.getRange2("B14").setValue("=B15");
			ws.getRange2("B15").setValue("=B16");
			ws.getRange2("B16").setValue("0");
			ws.getRange2("B22").setValue("=C23");

			let C3Index = AscCommonExcel.getCellIndex(ws.getRange2("C3").bbox.r1, ws.getRange2("C3").bbox.c1),
				C4Index = AscCommonExcel.getCellIndex(ws.getRange2("C4").bbox.r1, ws.getRange2("C4").bbox.c1),
				C5Index = AscCommonExcel.getCellIndex(ws.getRange2("C5").bbox.r1, ws.getRange2("C5").bbox.c1),
				C6Index = AscCommonExcel.getCellIndex(ws.getRange2("C6").bbox.r1, ws.getRange2("C6").bbox.c1),
				C14Index = AscCommonExcel.getCellIndex(ws.getRange2("C14").bbox.r1, ws.getRange2("C14").bbox.c1),
				C23Index = AscCommonExcel.getCellIndex(ws.getRange2("C23").bbox.r1, ws.getRange2("C23").bbox.c1);

			bbox = ws.getRange2("C3").bbox;
			ws.getRange2("C3").setValue("=A1:A6", undefined, undefined, bbox);
			ws.getRange2("C4").setValue("24");
			bbox = ws.getRange2("C5:C6").bbox;
			ws.getRange2("C5:C6").setValue("=A1:A6", undefined, undefined, bbox);
			bbox = ws.getRange2("C14").bbox;
			ws.getRange2("C14").setValue("=A1:A6", undefined, undefined, bbox);

			let D22Index = AscCommonExcel.getCellIndex(ws.getRange2("D22").bbox.r1, ws.getRange2("D22").bbox.c1);
			bbox = ws.getRange2("D22").bbox;
			ws.getRange2("D22").setValue("=A20:A22", undefined, undefined, bbox);

			let E1Index = AscCommonExcel.getCellIndex(ws.getRange2("E1").bbox.r1, ws.getRange2("E1").bbox.c1),
				E2Index = AscCommonExcel.getCellIndex(ws.getRange2("E2").bbox.r1, ws.getRange2("E2").bbox.c1),
				E3Index = AscCommonExcel.getCellIndex(ws.getRange2("E3").bbox.r1, ws.getRange2("E3").bbox.c1),
				E4Index = AscCommonExcel.getCellIndex(ws.getRange2("E4").bbox.r1, ws.getRange2("E4").bbox.c1),
				E5Index = AscCommonExcel.getCellIndex(ws.getRange2("E5").bbox.r1, ws.getRange2("E5").bbox.c1),
				E6Index = AscCommonExcel.getCellIndex(ws.getRange2("E6").bbox.r1, ws.getRange2("E6").bbox.c1);

			ws.getRange2("E1").setValue("25");
			bbox = ws.getRange2("E2:E6").bbox;
			ws.getRange2("E2:E6").setValue("=B3:B7", undefined, undefined, bbox);

			let	E9Index = AscCommonExcel.getCellIndex(ws.getRange2("E9").bbox.r1, ws.getRange2("E9").bbox.c1),
				E10Index = AscCommonExcel.getCellIndex(ws.getRange2("E10").bbox.r1, ws.getRange2("E10").bbox.c1),
				E11Index = AscCommonExcel.getCellIndex(ws.getRange2("E11").bbox.r1, ws.getRange2("E11").bbox.c1);

			bbox = ws.getRange2("E9:E11").bbox;
			ws.getRange2("E9:E11").setValue("=B3:B5", undefined, undefined, bbox);

			let	E13Index = AscCommonExcel.getCellIndex(ws.getRange2("E13").bbox.r1, ws.getRange2("E13").bbox.c1),
				E14Index = AscCommonExcel.getCellIndex(ws.getRange2("E14").bbox.r1, ws.getRange2("E14").bbox.c1);

			bbox = ws.getRange2("E13:E14").bbox;
			ws.getRange2("E13:E14").setValue("=B8:B10", undefined, undefined, bbox);

			let G1Index = AscCommonExcel.getCellIndex(ws.getRange2("G1").bbox.r1, ws.getRange2("G1").bbox.c1),
				G2Index = AscCommonExcel.getCellIndex(ws.getRange2("G2").bbox.r1, ws.getRange2("G2").bbox.c1),
				G3Index = AscCommonExcel.getCellIndex(ws.getRange2("G3").bbox.r1, ws.getRange2("G3").bbox.c1),
				G4Index = AscCommonExcel.getCellIndex(ws.getRange2("G4").bbox.r1, ws.getRange2("G4").bbox.c1),
				G5Index = AscCommonExcel.getCellIndex(ws.getRange2("G5").bbox.r1, ws.getRange2("G5").bbox.c1),
				G6Index = AscCommonExcel.getCellIndex(ws.getRange2("G6").bbox.r1, ws.getRange2("G6").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("D22").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("D22").getBBox0().r1, ws.getRange2("D22").getBBox0().c1);

			assert.ok(1, "Trace precedents from D22 to A20:A22 area");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D22Index, A20Index), "A20:A22", "D22<-A20");
			assert.strictEqual(traceManager._getPrecedents(A21Index, A22Index), undefined, "A21<-A22 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A22Index, B22Index), undefined, "A22<-B22 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B22Index, C23Index), undefined, "B22<-C23 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A20:A22"]), "object", "Area A20:A22 exist");

			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D22Index, A20Index), "A20:A22", "D22<-A20");
			assert.strictEqual(traceManager._getPrecedents(A21Index, A22Index), 1, "A21<-A22");
			assert.strictEqual(traceManager._getPrecedents(A22Index, B22Index), 1, "A22<-B22");
			assert.strictEqual(traceManager._getPrecedents(B22Index, C23Index), undefined, "B22<-C23 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A20:A22"]), "object", "Area A20:A22 exist");

			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D22Index, A20Index), "A20:A22", "D22<-A20");
			assert.strictEqual(traceManager._getPrecedents(A21Index, A22Index), 1, "A21<-A22");
			assert.strictEqual(traceManager._getPrecedents(A22Index, B22Index), 1, "A22<-B22");
			assert.strictEqual(traceManager._getPrecedents(B22Index, C23Index), 1, "B22<-C23");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A20:A22"]), "object", "Area A20:A22 exist");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("E13").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("E13").getBBox0().r1, ws.getRange2("E13").getBBox0().c1);
			
			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(first click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), undefined, "B8<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), undefined, "B9<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), undefined, "B10<-E6 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), undefined, "E6<-B3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), undefined, "A4<-A12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), undefined, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), undefined, "B3<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), undefined, "B4<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), undefined, "B5<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), undefined, "B6<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), undefined, "B7<-A1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "undefined", "Area A1:A6 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "undefined", "Area B3:B7 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(second click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), undefined, "E6<-B3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), undefined, "A4<-A12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), undefined, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), undefined, "B3<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), undefined, "B4<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), undefined, "B5<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), undefined, "B6<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), undefined, "B7<-A1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "undefined", "Area B3:B7 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(third click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), undefined, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), undefined, "B3<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), undefined, "B4<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), undefined, "B5<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), undefined, "B6<-A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), undefined, "B7<-A1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(fourth click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), undefined, "B12<-B13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), undefined, "B13<-B14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), undefined, "B14<-B15 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(fifth click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), 1, "B12<-B13");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), undefined, "B13<-B14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), undefined, "B14<-B15 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(sixth click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), 1, "B12<-B13");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), 1, "B13<-B14");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), undefined, "B14<-B15 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(seventh click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), 	"B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), 1, "B12<-B13");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), 1, "B13<-B14");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), 1, "B14<-B15");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Trace precedents from E13 to a chain of links with ranges(eighth click)");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), 1, "B12<-B13");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), 1, "B13<-B14");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), 1, "B14<-B15");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), 1, "B15<-B16");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Delete one last built precedents from from linking multiple ranges and cells. Click(1) on E13");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), 1, "B12<-B13");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), 1, "B13<-B14");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), 1, "B14<-B15");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Delete four more last built precedents from from linking multiple ranges and cells. Click(5) on E13");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(E13Index, B8Index), "B8:B10", "E13<-B8");
			assert.strictEqual(traceManager._getPrecedents(B8Index, A1Index), "A1:A6", "B8<-A1");
			assert.strictEqual(traceManager._getPrecedents(B9Index, A1Index), "A1:A6", "B9<-A1");
			assert.strictEqual(traceManager._getPrecedents(B10Index, E6Index), 1, "B10<-E6");
			assert.strictEqual(traceManager._getPrecedents(E6Index, B3Index), "B3:B7", "E6<-B3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), undefined, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B12Index, B13Index), undefined, "B12<-B13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B13Index, B14Index), undefined, "B13<-B14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B14Index, B15Index), undefined, "B14<-B15 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B15Index, B16Index), undefined, "B15<-B16 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A1:A6", "B3<-A1");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), "A1:A6", "B4<-A1");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A1:A6", "B5<-A1");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), "A1:A6", "B6<-A1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A1:A6", "B7<-A1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "object", "Area B3:B7 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Delete three more last built precedents from from linking multiple ranges and cells. Click(8) on E13");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "object", "Area A1:A6 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "undefined", "Area B3:B7 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Delete one more last built precedents from from linking multiple ranges and cells. Click(9) on E13");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "undefined", "Area A1:A6 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "undefined", "Area B3:B7 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "object", "Area B8:B10 exist");

			assert.ok(1, "Delete one more last built precedents from from linking multiple ranges and cells. Click(10) on E13");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(typeof(traceManager.precedentsAreas["A1:A6"]), "undefined", "Area A1:A6 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B3:B7"]), "undefined", "Area B3:B7 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["B8:B10"]), "undefined", "Area B8:B10 doesn't exist");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			let N14Index = AscCommonExcel.getCellIndex(ws.getRange2("N14").bbox.r1, ws.getRange2("N14").bbox.c1),
				N15Index = AscCommonExcel.getCellIndex(ws.getRange2("N15").bbox.r1, ws.getRange2("N15").bbox.c1),
				L14Index = AscCommonExcel.getCellIndex(ws.getRange2("L14").bbox.r1, ws.getRange2("L14").bbox.c1),
				L15Index = AscCommonExcel.getCellIndex(ws.getRange2("L15").bbox.r1, ws.getRange2("L15").bbox.c1),
				J13Index = AscCommonExcel.getCellIndex(ws.getRange2("J13").bbox.r1, ws.getRange2("J13").bbox.c1),
				J14Index = AscCommonExcel.getCellIndex(ws.getRange2("J14").bbox.r1, ws.getRange2("J14").bbox.c1),
				I12Index = AscCommonExcel.getCellIndex(ws.getRange2("I12").bbox.r1, ws.getRange2("I12").bbox.c1),
				I13Index = AscCommonExcel.getCellIndex(ws.getRange2("I13").bbox.r1, ws.getRange2("I13").bbox.c1);

			bbox = ws.getRange2("N14:N15").bbox;
			ws.getRange2("N14:N15").setValue("=SUM(J13:J14)+SIN(L14:L15)", undefined, undefined, bbox);
			ws.getRange2("L14").setValue("123");
			ws.getRange2("L15").setValue("321");
			ws.getRange2("J13").setValue("=I12+L14");
			ws.getRange2("J14").setValue("=I13+L15");
			ws.getRange2("I12").setValue("1");
			ws.getRange2("I13").setValue("2");

			ws.selectionRange.ranges = [ws.getRange2("N14").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("N14").getBBox0().r1, ws.getRange2("N14").getBBox0().c1);
			
			assert.ok(1, "Trace precedents to a multiple consecutive ranges. First click on N14");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(N14Index, L14Index), "L14:L15", "N14<-L14");
			assert.strictEqual(traceManager._getPrecedents(N14Index, J13Index), "J13:J14", "N14<-J13");
			assert.strictEqual(traceManager._getPrecedents(J13Index, I12Index), undefined, "J13<-I12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J13Index, L14Index), undefined, "J13<-L14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, I13Index), undefined, "J14<-I13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, L15Index), undefined, "J14<-L15 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["J13:J14"]), "object", "Area J13:J14 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["L14:L15"]), "object", "Area L14:L15 exist");

			assert.ok(1, "Trace precedents to a multiple consecutive ranges. Second click on N14");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(N14Index, L14Index), "L14:L15", "N14<-L14");
			assert.strictEqual(traceManager._getPrecedents(N14Index, J13Index), "J13:J14", "N14<-J13");
			assert.strictEqual(traceManager._getPrecedents(J13Index, I12Index), 1, "J13<-I12");
			assert.strictEqual(traceManager._getPrecedents(J13Index, L14Index), 1, "J13<-L14");
			assert.strictEqual(traceManager._getPrecedents(J14Index, I13Index), 1, "J14<-I13");
			assert.strictEqual(traceManager._getPrecedents(J14Index, L15Index), 1, "J14<-L15");
			assert.strictEqual(typeof(traceManager.precedentsAreas["J13:J14"]), "object", "Area J13:J14 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["L14:L15"]), "object", "Area L14:L15 exist");

			assert.ok(1, "Remove last built precedent traces from a multiple consecutive ranges. First click on N14");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(N14Index, L14Index), "L14:L15", "N14<-L14");
			assert.strictEqual(traceManager._getPrecedents(N14Index, J13Index), "J13:J14", "N14<-J13");
			assert.strictEqual(traceManager._getPrecedents(J13Index, I12Index), undefined, "J13<-I12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J13Index, L14Index), undefined, "J13<-L14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, I13Index), undefined, "J14<-I13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, L15Index), undefined, "J14<-L15 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["J13:J14"]), "object", "Area J13:J14 exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["L14:L15"]), "object", "Area L14:L15 exist");

			assert.ok(1, "Remove last built precedent traces from a multiple consecutive ranges. Second click on N14");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(N14Index, L14Index), undefined, "N14<-L14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(N14Index, J13Index), undefined, "N14<-J13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J13Index, I12Index), undefined, "J13<-I12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J13Index, L14Index), undefined, "J13<-L14 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, I13Index), undefined, "J14<-I13 === undefined");
			assert.strictEqual(traceManager._getPrecedents(J14Index, L15Index), undefined, "J14<-L15 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["J13:J14"]), "undefined", "Area J13:J14 doesn't exist");
			assert.strictEqual(typeof(traceManager.precedentsAreas["L14:L15"]), "undefined", "Area L14:L15 doesn't exist");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// -------------- dependents --------------
			ws.selectionRange.ranges = [ws.getRange2("B12").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B12").getBBox0().r1, ws.getRange2("B12").getBBox0().c1);

			assert.ok(1, "Trace dependents from B12");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), 1, "B12->A12");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), undefined, "A4<-A12 === undefined");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), undefined, "A12->A4 === undefined");

			assert.ok(1, "Trace dependents from B12");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), 1, "B12->A12");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), 1, "A12->A4");
			assert.strictEqual(traceManager._getDependents(A4Index, B1Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, B3Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, B8Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, B9Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, C3Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, C5Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, C6Index), undefined);
			assert.strictEqual(traceManager._getDependents(A4Index, C14Index), undefined);

			assert.ok(1, "Trace dependents from B12");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), 1, "B12->A12");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), 1, "A12->A4");
			assert.strictEqual(traceManager._getDependents(A4Index, B1Index), 1, "A4->B1");
			assert.strictEqual(traceManager._getDependents(A4Index, B3Index), 1, "A4->B3");
			assert.strictEqual(traceManager._getDependents(A4Index, B8Index), 1, "A4->B8");
			assert.strictEqual(traceManager._getDependents(A4Index, B9Index), 1, "A4->B9");
			assert.strictEqual(traceManager._getDependents(A4Index, C3Index), 1, "A4->C3");
			assert.strictEqual(traceManager._getDependents(A4Index, C5Index), 1, "A4->C5");
			assert.strictEqual(traceManager._getDependents(A4Index, C6Index), 1, "A4->C6");
			assert.strictEqual(traceManager._getDependents(A4Index, C14Index), 1, "A4->C14");

			// clear last built dependents
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), 1, "B12->A12");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), 1, "A4<-A12");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), 1, "A12->A4");
			assert.strictEqual(traceManager._getDependents(A4Index, B1Index), undefined, "A4->B1 === undefined");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), 1, "A12<-B12");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), 1, "B12->A12");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), undefined, "A4<-A12 === undefined");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), undefined, "A12->A4 === undefined");
			assert.strictEqual(traceManager._getDependents(A4Index, B1Index), undefined, "A4->B1 === undefined");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getPrecedents(A12Index, B12Index), undefined, "A12<-B12 === undefined");
			assert.strictEqual(traceManager._getDependents(B12Index, A12Index), undefined, "B12->A12 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A4Index, A12Index), undefined, "A4<-A12 === undefined");
			assert.strictEqual(traceManager._getDependents(A12Index, A4Index), undefined, "A12->A4 === undefined");
			assert.strictEqual(traceManager._getDependents(A4Index, B1Index), undefined, "A4->B1 === undefined");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

		});
		QUnit.test("Test: \"Shared tests\"", function (assert) {
			// ???
			let cellWithFormula, oParser, sharedRef, bbox;
			ws.getRange2("A1:J20").cleanAll();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1),
				C3Index = AscCommonExcel.getCellIndex(ws.getRange2("C3").bbox.r1, ws.getRange2("C3").bbox.c1),
				A5Index = AscCommonExcel.getCellIndex(ws.getRange2("A5").bbox.r1, ws.getRange2("A5").bbox.c1),
				B5Index = AscCommonExcel.getCellIndex(ws.getRange2("B5").bbox.r1, ws.getRange2("B5").bbox.c1),
				C5Index = AscCommonExcel.getCellIndex(ws.getRange2("C5").bbox.r1, ws.getRange2("C5").bbox.c1),
				D5Index = AscCommonExcel.getCellIndex(ws.getRange2("D5").bbox.r1, ws.getRange2("D5").bbox.c1),
				E5Index = AscCommonExcel.getCellIndex(ws.getRange2("E5").bbox.r1, ws.getRange2("E5").bbox.c1);
			
			ws.getRange2("A1").setValue("=A3");
			ws.getRange2("B1").setValue("=B3");
			ws.getRange2("C1").setValue("=C3");
			ws.getRange2("A3").setValue("1");
			ws.getRange2("B3").setValue("2");
			ws.getRange2("C3").setValue("3");
			ws.getRange2("A5").setValue("=A3");
			ws.getRange2("B5").setValue("=B3");
			ws.getRange2("C5").setValue("=C3");
			ws.getRange2("D5").setValue("=D3");
			ws.getRange2("E5").setValue("=E3");

			bbox = ws.getRange2("B1:C1").bbox;
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("B3", cellWithFormula, ws);
			sharedRef = bbox.clone();
			oParser.setShared(sharedRef, cellWithFormula);
			oParser.parse();
			oParser.calculate();

			// set parser formula to the cell
			ws.getRange2("B1:C1")._foreachNoEmpty(function (cell) {
				cell.formulaParsed = oParser;
				// cell.formulaParsed.setShared(sharedRef, cellWithFormula);
				// cell.formulaParsed.parse();
				// cell.formulaParsed.calculate();
				// let parsed = cell.getFormulaParsed();
				// parsed.buildDependencies();
			});
			ws.selectionRange.ranges = [ws.getRange2("B3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B3").getBBox0().r1, ws.getRange2("B3").getBBox0().c1);

			bbox = ws.getRange2("B5:E5").bbox;
			cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
			oParser = new parserFormula("B3", cellWithFormula, ws);
			sharedRef = bbox.clone();

			ws.getRange2("B5:E5")._foreachNoEmpty(function (cell) {
				// change parserFormula
			})

			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(B3Index, B1Index), 1);
			assert.strictEqual(traceManager._getDependents(B3Index, B5Index), 1);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Tables tests\"", function (assert) {
			let bbox;
			// -------------- precedents --------------
			ws.getRange2("A1:Z100").cleanAll();
			ws.selectionRange.ranges = [ws.getRange2("A2:A4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A2").getBBox0().r1, ws.getRange2("A2").getBBox0().c1);

			let tableProp = AscCommonExcel.AddFormatTableOptions();
			tableProp.asc_setRange("$A$2:$A$4");
			tableProp.asc_setIsTitle(false);
			
			ws.autoFilters.addAutoFilter("TableStyleLight1", ws.selectionRange.getLast().clone(), tableProp);

			let A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				A5Index = AscCommonExcel.getCellIndex(ws.getRange2("A5").bbox.r1, ws.getRange2("A5").bbox.c1),
				A8Index = AscCommonExcel.getCellIndex(ws.getRange2("A8").bbox.r1, ws.getRange2("A8").bbox.c1),
				B2Index = AscCommonExcel.getCellIndex(ws.getRange2("B2").bbox.r1, ws.getRange2("B2").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1),
				B4Index = AscCommonExcel.getCellIndex(ws.getRange2("B4").bbox.r1, ws.getRange2("B4").bbox.c1),
				B5Index = AscCommonExcel.getCellIndex(ws.getRange2("B5").bbox.r1, ws.getRange2("B5").bbox.c1),
				C1Index = AscCommonExcel.getCellIndex(ws.getRange2("C1").bbox.r1, ws.getRange2("C1").bbox.c1),
				D2Index = AscCommonExcel.getCellIndex(ws.getRange2("D2").bbox.r1, ws.getRange2("D2").bbox.c1),
				D3Index = AscCommonExcel.getCellIndex(ws.getRange2("D3").bbox.r1, ws.getRange2("D3").bbox.c1),
				D4Index = AscCommonExcel.getCellIndex(ws.getRange2("D4").bbox.r1, ws.getRange2("D4").bbox.c1),
				D5Index = AscCommonExcel.getCellIndex(ws.getRange2("D5").bbox.r1, ws.getRange2("D5").bbox.c1),
				E2Index = AscCommonExcel.getCellIndex(ws.getRange2("E2").bbox.r1, ws.getRange2("E2").bbox.c1),
				E3Index = AscCommonExcel.getCellIndex(ws.getRange2("E3").bbox.r1, ws.getRange2("E3").bbox.c1),
				E4Index = AscCommonExcel.getCellIndex(ws.getRange2("E4").bbox.r1, ws.getRange2("E4").bbox.c1),
				E5Index = AscCommonExcel.getCellIndex(ws.getRange2("E5").bbox.r1, ws.getRange2("E5").bbox.c1),
				E7Index = AscCommonExcel.getCellIndex(ws.getRange2("E7").bbox.r1, ws.getRange2("E7").bbox.c1),
				E8Index = AscCommonExcel.getCellIndex(ws.getRange2("E8").bbox.r1, ws.getRange2("E8").bbox.c1);

			ws.getRange2("A3").setValue("1");
			ws.getRange2("A4").setValue("=C1");
			ws.getRange2("A5").setValue("3");
			ws.getRange2("A8").setValue("=A2");
			ws.getRange2("B3").setValue("a");
			ws.getRange2("B4").setValue("b");
			ws.getRange2("B5").setValue("c");
			ws.getRange2("D2").setValue("=A2");

			bbox = ws.getRange2("D3:D5").bbox;
			ws.getRange2("D3:D5").setValue("=Table1", undefined, undefined, bbox);

			ws.getRange2("E2").setValue("=Table1");
			ws.getRange2("E3").setValue("=Table1");
			ws.getRange2("E4").setValue("=Table1");
			ws.getRange2("E5").setValue("=Table1");
			ws.getRange2("E7").setValue("=Table1");

			bbox = ws.getRange2("E8").bbox;
			ws.getRange2("E8").setValue("=Table1", undefined, undefined, bbox);

			ws.selectionRange.ranges = [ws.getRange2("A3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A3").getBBox0().r1, ws.getRange2("A3").getBBox0().c1);

			assert.ok(1, "Trace dependents from A3");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A3Index, D2Index), undefined, "A3->D2 === undefined");
			assert.strictEqual(traceManager._getDependents(A3Index, D3Index), 1, "A3->D3");
			assert.strictEqual(traceManager._getDependents(A3Index, D4Index), 1, "A3->D4");	// ?
			assert.strictEqual(traceManager._getDependents(A3Index, D5Index), 1, "A3->D5");	// ?
			assert.strictEqual(traceManager._getDependents(A3Index, A8Index), undefined, "A3->A8 === undefined");

			assert.ok(1, "Remove last dependents from A3");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A3Index, D2Index), undefined, "A3->D2 === undefined");
			assert.strictEqual(traceManager._getDependents(A3Index, D3Index), undefined, "A3->D3 === undefined");
			assert.strictEqual(traceManager._getDependents(A3Index, D4Index), undefined, "A3->D4 === undefined");
			assert.strictEqual(traceManager._getDependents(A3Index, D5Index), undefined, "A3->D5 === undefined");
			assert.strictEqual(traceManager._getDependents(A3Index, A8Index), undefined, "A3->A8 === undefined");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("A2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A2").getBBox0().r1, ws.getRange2("A2").getBBox0().c1);

			assert.ok(1, "Trace dependents from A2");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A2Index, D2Index), 1, "A2->D2");
			assert.strictEqual(traceManager._getDependents(A2Index, D3Index), undefined, "A2->D3 === undefined");
			assert.strictEqual(traceManager._getDependents(A2Index, D4Index), undefined, "A2->D4 === undefined");
			assert.strictEqual(traceManager._getDependents(A2Index, D5Index), undefined, "A3->D5 === undefined");
			assert.strictEqual(traceManager._getDependents(A2Index, A8Index), 1, "A2->A8");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// precedents
			ws.selectionRange.ranges = [ws.getRange2("D2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("D2").getBBox0().r1, ws.getRange2("D2").getBBox0().c1);

			assert.ok(1, "Trace precedents from D2");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D2Index, A2Index), 1, "D2<-A2");
			assert.ok(1, "Remove last precedents from D2");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(D2Index, A2Index), undefined, "D2<-A2 === undefined");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("D3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("D3").getBBox0().r1, ws.getRange2("D3").getBBox0().c1);

			assert.ok(1, "Trace precedents from D3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D3Index, A3Index), "$A$3:$A$5", "D3<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "object", "Area A3:A5 exist");

			assert.ok(1, "Trace precedents from D3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(D3Index, A3Index), "$A$3:$A$5", "D3<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), 1, "A4<-C1");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "object", "Area A3:A5 exist");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(D3Index, A3Index), "$A$3:$A$5", "D3<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "object", "Area A3:A5 exist");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(D3Index, A3Index), undefined, "D3<-A3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "undefined", "Area A3:A5 doesn't exist");

			ws.selectionRange.ranges = [ws.getRange2("E2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("E2").getBBox0().r1, ws.getRange2("E2").getBBox0().c1);

			assert.ok(1, "Trace precedents from E2");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E2Index, A3Index), "$A$3:$A$5", "E2<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "object", "Area A3:A5 exist");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("E3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("E3").getBBox0().r1, ws.getRange2("E3").getBBox0().c1);

			assert.ok(1, "Trace precedents from E3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E2Index, A3Index), undefined, "E2<-A3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E3Index, A3Index), 1, "E3<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(traceManager.precedentsAreas, null);

			ws.selectionRange.ranges = [ws.getRange2("E8").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("E8").getBBox0().r1, ws.getRange2("E8").getBBox0().c1);

			assert.ok(1, "Trace precedents from E8");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(E8Index, A3Index), "$A$3:$A$5", "E8<-A3");
			assert.strictEqual(traceManager._getPrecedents(A4Index, C1Index), undefined, "A4<-C1 === undefined");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$3:$A$5"]), "object", "Area A3:A5 exist");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Deletes tests\"", function (assert) {
			let bbox;
			// ------------------- base precedents ------------------- //
			ws.selectionRange.ranges = [ws.getRange2("I1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("I1").getBBox0().r1, ws.getRange2("I1").getBBox0().c1);

			ws.getRange2("A1:J20").cleanAll();

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

			assert.ok(1, "Trace precedents from I1, six times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1, "G1<-G4");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1, "G4<-I4");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), 1, "I4<-I3");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), 1, "I3<-H3");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1, "E1<-C1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1, "E1<-C4");

			// first clear
			assert.ok(1, "Remove Precedents Arrows from I1. First click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1, "G1<-G4");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1, "G4<-I4");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), 1, "I4<-I3");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined, "I3<-H3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1, "E1<-C1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1, "E1<-C4");

			// second clear
			assert.ok(1, "Remove Precedents Arrows from I1. Second click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1, "G1<-G4");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), 1, "G4<-I4");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined, "I4<-I3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined, "I3<-H3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), 1, "E1<-C1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), 1, "E1<-C4");

			// third clear
			assert.ok(1, "Remove Precedents Arrows from I1. Third click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1, "G1<-G4");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined, "G4<-I4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined, "I4<-I3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined, "I3<-H3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined, "E1<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined, "E1<-C4 === undefined");

			// fourth clear
			assert.ok(1, "Remove Precedents Arrows from I1. Fourth click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined, "G1<-G4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined, "G4<-I4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined, "I4<-I3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined, "I3<-H3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined, "G1<-E1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined, "E1<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined, "E1<-C4 === undefined");

			// fifth clear
			assert.ok(1, "Remove Precedents Arrows from I1. Fifth click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), undefined, "I1<-G1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined, "G1<-G4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G4Index, I4Index), undefined, "G4<-I4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I4Index, I3Index), undefined, "I4<-I3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I3Index, H3Index), undefined, "I3<-H3 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined, "G1<-E1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C1Index), undefined, "E1<-C1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E1Index, C4Index), undefined, "E1<-C4 === undefined");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base precedents with external -------------------//
			assert.ok(1, "Add dependency from another sheet for cell I1");
			ws2.getRange2("B1").setValue("1");
			ws.getRange2("I1").setValue("=G1+Sheet2!B1");

			let B1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("B1").bbox.r1, ws2.getRange2("B1").bbox.c1) + ";" + ws2.getIndex();

			assert.ok(1, "Trace precedents from I1, two times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.ok(traceManager.precedentsExternal[I1Index][B1ExternalIndex], "I1<-B1(External)");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), 1, "G1<-G4");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1");

			// first clear
			assert.ok(1, "Remove Precedents Arrows from I1. First click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), 1, "I1<-G1");
			assert.ok(traceManager.precedentsExternal[I1Index][B1ExternalIndex], "I1<-B1(External)");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined, "G1<-G4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined, "G1<-E1 === undefined");

			// second clear
			assert.ok(1, "Remove Precedents Arrows from I1. Second click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);
			assert.strictEqual(traceManager._getPrecedents(I1Index, G1Index), undefined, "I1<-G1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(I1Index, B1ExternalIndex), undefined, "I1<-B1(External) === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, G4Index), undefined, "G1<-G4 === undefined");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined, "G1<-E1 === undefined");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base dependents ---------------------------------//
			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			ws.getRange2("A1:J10").cleanAll();

			ws.getRange2("A1").setValue("1");
			ws.getRange2("A6").setValue("7");
			ws.getRange2("A7").setValue("7");
			ws.getRange2("A8").setValue("=A7");
			ws.getRange2("C1").setValue("=A1");
			ws.getRange2("C4").setValue("=A1");
			bbox = ws.getRange2("C7").bbox;
			ws.getRange2("C7").setValue("=A6:A8", undefined, undefined, bbox);
			ws.getRange2("E1").setValue("=C1");
			ws.getRange2("E4").setValue("=C4");
			ws.getRange2("G1").setValue("=E1+E4");
			ws.getRange2("G4").setValue("=E4");
			ws.getRange2("F6").setValue("=G4");
			ws.getRange2("H6").setValue("=G4");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A6Index = AscCommonExcel.getCellIndex(ws.getRange2("A6").bbox.r1, ws.getRange2("A6").bbox.c1),
				A7Index = AscCommonExcel.getCellIndex(ws.getRange2("A7").bbox.r1, ws.getRange2("A7").bbox.c1),
				A8Index = AscCommonExcel.getCellIndex(ws.getRange2("A8").bbox.r1, ws.getRange2("A8").bbox.c1),
				C7Index = AscCommonExcel.getCellIndex(ws.getRange2("C7").bbox.r1, ws.getRange2("C7").bbox.c1),
				E4Index = AscCommonExcel.getCellIndex(ws.getRange2("E4").bbox.r1, ws.getRange2("E4").bbox.c1),
				F6Index = AscCommonExcel.getCellIndex(ws.getRange2("F6").bbox.r1, ws.getRange2("F6").bbox.c1),
				H6Index = AscCommonExcel.getCellIndex(ws.getRange2("H6").bbox.r1, ws.getRange2("H6").bbox.c1);

			assert.ok(1, "Trace dependents from A1, six times");
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4");
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), 1, "E1->G1");
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), 1, "E4->G4");
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), 1, "G4->F6");
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), 1, "G4->H6");

			// first clear
			assert.ok(1, "Remove Dependents Arrows from A1. First click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4");
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), 1, "E1->G1");
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), 1, "E4->G4");
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined, "G4->F6 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined, "G4->H6 === undeifned");

			// second clear
			assert.ok(1, "Remove Dependents Arrows from A1. Second click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4");
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined, "E1->G1 === undefined");
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined, "E4->G4 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined, "G4->F6 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined, "G4->H6 === undefined");

			// third clear
			assert.ok(1, "Remove Dependents Arrows from A1. Third click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined, "C1->E1 === undefined");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined");
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined, "E1->G1 === undefined");
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined, "E4->G4 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined, "G4->F6 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined, "G4->H6 === undeifned");

			// fourth clear
			assert.ok(1, "Remove Dependents Arrows from A1. Fourth click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), undefined, "A1->C1 === undefined");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), undefined, "A1->C4 === undefined");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined, "C1->E1 === undefined");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined");
			assert.strictEqual(traceManager._getDependents(E1Index, G1Index), undefined, "E1->G1 === undefined");
			assert.strictEqual(traceManager._getDependents(E4Index, G4Index), undefined, "E4->G4 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, F6Index), undefined, "G4->F6 === undefined");
			assert.strictEqual(traceManager._getDependents(G4Index, H6Index), undefined, "G4->H6 === undeifned");

			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// ------------------- base dependents with external -------------------//
			assert.ok(1, "Add dependency from another sheet for cell A1");
			ws2.getRange2("A1").setValue("=Sheet1!A1");

			let A1ExternalIndex = AscCommonExcel.getCellIndex(ws2.getRange2("A1").bbox.r1, ws2.getRange2("A1").bbox.c1) + ";" + ws2.getIndex();

			assert.ok(1, "Trace dependents from A1, two times");
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1, "A1->A1(External)");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4");

			// first clear
			assert.ok(1, "Remove Dependents Arrows from A1. First click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), 1, "A1->C1");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), 1, "A1->C4");
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), 1, "A1->A1(External)");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined, "C1->E1 === undefined");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined");

			// second clear
			assert.ok(1, "Remove Dependents Arrows from A1. Second click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A1Index, C1Index), undefined, "A1->C1 === undefined");
			assert.strictEqual(traceManager._getDependents(A1Index, C4Index), undefined, "A1->C4 === undefined");
			assert.strictEqual(traceManager._getDependents(A1Index, A1ExternalIndex), undefined, "A1->A1(External) === undefined");
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), undefined, "C1->E1 === undefined");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("A7").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A7").getBBox0().r1, ws.getRange2("A7").getBBox0().c1);

			assert.ok(1, "Trace dependents from A7");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A7Index, A8Index), 1, "A7->A8");
			assert.strictEqual(traceManager._getDependents(A7Index, C7Index), 1, "A7->C7");
			assert.strictEqual(traceManager._getDependents(A8Index, C7Index), undefined, "A8->C7 === undefined");
			
			assert.ok(1, "Trace dependents from A7");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A7Index, A8Index), 1, "A7->A8");
			assert.strictEqual(traceManager._getDependents(A7Index, C7Index), 1, "A7->C7");
			assert.strictEqual(traceManager._getDependents(A8Index, C7Index), 1, "A8->C7");

			// first clear
			assert.ok(1, "Remove Dependents Arrows from A7. First click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A7Index, A8Index), 1, "A7->A8");
			assert.strictEqual(traceManager._getDependents(A7Index, C7Index), 1, "A7->C7");
			assert.strictEqual(traceManager._getDependents(A8Index, C7Index), undefined, "A8->C7 === undefined");

			// second clear
			assert.ok(1, "Remove Dependents Arrows from A7. Second click");
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);
			assert.strictEqual(traceManager._getDependents(A7Index, A8Index), undefined, "A7->A8 === undefined");
			assert.strictEqual(traceManager._getDependents(A7Index, C7Index), undefined, "A7->C7 === undefined");
			assert.strictEqual(traceManager._getDependents(A8Index, C7Index), undefined, "A8->C7 === undefined");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Merged cells tests\"", function (assert) {
			ws.getRange2("A1:Z100").cleanAll();
			let bbox;
			// ------------------- do merge when cell already have formula ------------------- //
			ws.getRange2("A1").setValue("=E1");
			ws.getRange2("A2").setValue("=E2");
			ws.getRange2("B1").setValue("=F1");
			ws.getRange2("B2").setValue("=F2");
			ws.getRange2("E1").setValue("=H1");
			ws.getRange2("E2").setValue("=H2");
			ws.getRange2("F1").setValue("=I1");
			ws.getRange2("F2").setValue("=I2");

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				B1Index = AscCommonExcel.getCellIndex(ws.getRange2("B1").bbox.r1, ws.getRange2("B1").bbox.c1),
				B2Index = AscCommonExcel.getCellIndex(ws.getRange2("B2").bbox.r1, ws.getRange2("B2").bbox.c1),
				E1Index = AscCommonExcel.getCellIndex(ws.getRange2("E1").bbox.r1, ws.getRange2("E1").bbox.c1),
				E2Index = AscCommonExcel.getCellIndex(ws.getRange2("E2").bbox.r1, ws.getRange2("E2").bbox.c1),
				F1Index = AscCommonExcel.getCellIndex(ws.getRange2("F1").bbox.r1, ws.getRange2("F1").bbox.c1),
				F2Index = AscCommonExcel.getCellIndex(ws.getRange2("F2").bbox.r1, ws.getRange2("F2").bbox.c1),
				H1Index = AscCommonExcel.getCellIndex(ws.getRange2("H1").bbox.r1, ws.getRange2("H1").bbox.c1),
				H2Index = AscCommonExcel.getCellIndex(ws.getRange2("H2").bbox.r1, ws.getRange2("H2").bbox.c1),
				I1Index = AscCommonExcel.getCellIndex(ws.getRange2("I1").bbox.r1, ws.getRange2("I1").bbox.c1),
				I2Index = AscCommonExcel.getCellIndex(ws.getRange2("I2").bbox.r1, ws.getRange2("I2").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);
			assert.ok(1, "Trace precedents from A1");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			ws.selectionRange.ranges = [ws.getRange2("A2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A2").getBBox0().r1, ws.getRange2("A2").getBBox0().c1);
			assert.ok(1, "Trace precedents from A2");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			ws.selectionRange.ranges = [ws.getRange2("B1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B1").getBBox0().r1, ws.getRange2("B1").getBBox0().c1);
			assert.ok(1, "Trace precedents from B1");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			ws.selectionRange.ranges = [ws.getRange2("B2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B2").getBBox0().r1, ws.getRange2("B2").getBBox0().c1);
			assert.ok(1, "Trace precedents from B2");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();

			assert.strictEqual(traceManager._getPrecedents(A1Index, E1Index), 1, "A1<-E1");
			assert.strictEqual(traceManager._getDependents(E1Index, A1Index), 1, "E1->A1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, H1Index), 1, "E1<-H1");
			assert.strictEqual(traceManager._getPrecedents(A2Index, E2Index), 1, "A2<-E2");
			assert.strictEqual(traceManager._getDependents(E2Index, A2Index), 1, "E2->A2");
			assert.strictEqual(traceManager._getPrecedents(E2Index, H2Index), 1, "E2<-H2");
			assert.strictEqual(traceManager._getPrecedents(B1Index, F1Index), 1, "B1<-F1");
			assert.strictEqual(traceManager._getDependents(F1Index, B1Index), 1, "F1->B1");
			assert.strictEqual(traceManager._getPrecedents(F1Index, I1Index), 1, "F1<-I1");
			assert.strictEqual(traceManager._getPrecedents(B2Index, F2Index), 1, "B2<-F2");
			assert.strictEqual(traceManager._getDependents(F2Index, B2Index), 1, "F2->B2");
			assert.strictEqual(traceManager._getPrecedents(F2Index, I2Index), 1, "F2<-I2");

			ws.getRange2("A1:B2").merge(2);		// center merge
			assert.strictEqual(traceManager._getPrecedents(A1Index, E1Index), 1, "A1<-E1");
			assert.strictEqual(traceManager._getDependents(E1Index, A1Index), 1, "E1->A1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, H1Index), 1, "E1<-H1");
			assert.strictEqual(traceManager._getPrecedents(A2Index, E2Index), undefined, "A2<-E2 === undefined");
			assert.strictEqual(traceManager._getDependents(E2Index, A2Index), undefined, "E2->A2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E2Index, H2Index), 1, "E2<-H2");
			assert.strictEqual(traceManager._getPrecedents(B1Index, F1Index), undefined, "B1<-F1 === undefined");
			assert.strictEqual(traceManager._getDependents(F1Index, B1Index), undefined, "F1->B1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(F1Index, I1Index), 1, "F1<-I1");
			assert.strictEqual(traceManager._getPrecedents(B2Index, F2Index), undefined, "B2<-F2 === undefined");
			assert.strictEqual(traceManager._getDependents(F2Index, B2Index), undefined, "F2->B2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(F2Index, I2Index), 1, "F2<-I2");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.getRange2("A6").setValue("=A1");
			ws.getRange2("A7").setValue("=A2");
			ws.getRange2("B6").setValue("=B1");
			ws.getRange2("B7").setValue("=B2");
			ws.getRange2("A9").setValue("=A6");
			ws.getRange2("B9").setValue("=B6");

			let A6Index = AscCommonExcel.getCellIndex(ws.getRange2("A6").bbox.r1, ws.getRange2("A6").bbox.c1),
				A7Index = AscCommonExcel.getCellIndex(ws.getRange2("A7").bbox.r1, ws.getRange2("A7").bbox.c1),
				A9Index = AscCommonExcel.getCellIndex(ws.getRange2("A9").bbox.r1, ws.getRange2("A9").bbox.c1),
				B6Index = AscCommonExcel.getCellIndex(ws.getRange2("B6").bbox.r1, ws.getRange2("B6").bbox.c1),
				B7Index = AscCommonExcel.getCellIndex(ws.getRange2("B7").bbox.r1, ws.getRange2("B7").bbox.c1),
				B9Index = AscCommonExcel.getCellIndex(ws.getRange2("B9").bbox.r1, ws.getRange2("B9").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("A9").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A9").getBBox0().r1, ws.getRange2("A9").getBBox0().c1);
			assert.ok(1, "Trace precedents from A9, six times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A9Index, A6Index), 1, "A9<-A6");
			assert.strictEqual(traceManager._getPrecedents(A6Index, A1Index), 1, "A6<-A1");
			assert.strictEqual(traceManager._getPrecedents(A1Index, E1Index), 1, "A1<-E1");
			assert.strictEqual(traceManager._getDependents(E1Index, A1Index), 1, "E1->A1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, H1Index), 1, "E1<-H1");
			assert.strictEqual(traceManager._getPrecedents(A2Index, E2Index), undefined, "A2<-E2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(E2Index, H2Index), undefined, "E2<-H2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B1Index, F1Index), undefined, "B1<-F1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(F1Index, I1Index), undefined, "F1<-I1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B2Index, F2Index), undefined, "B2<-F2 === undefined");
			assert.strictEqual(traceManager._getPrecedents(F2Index, I2Index), undefined, "F2<-I2 === undefined");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("B9").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B9").getBBox0().r1, ws.getRange2("B9").getBBox0().c1);
			assert.ok(1, "Trace precedents from B9, six times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B9Index, B6Index), 1, "B9<-B6");
			assert.strictEqual(traceManager._getPrecedents(B6Index, B1Index), 1, "B6<-B1");
			assert.strictEqual(traceManager._getPrecedents(B1Index, F1Index), undefined, "B1<-F1");
			assert.strictEqual(traceManager._getDependents(F1Index, B1Index), undefined, "F1->B1");
			
			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("A7").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A7").getBBox0().r1, ws.getRange2("A7").getBBox0().c1);
			assert.ok(1, "Trace precedents from A7, six times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A7Index, A2Index), 1, "A7<-A2");
			assert.strictEqual(traceManager._getPrecedents(A2Index, E2Index), undefined, "A2<-E2");

			ws.selectionRange.ranges = [ws.getRange2("B7").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B7").getBBox0().r1, ws.getRange2("B7").getBBox0().c1);
			assert.ok(1, "Trace precedents from B7, six times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B7Index, B2Index), 1, "B7<-B2");
			assert.strictEqual(traceManager._getPrecedents(B2Index, F2Index), undefined, "B2<-F2");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			// trace dependents/precedents from merged range
			ws.selectionRange.ranges = [ws.getRange2("B2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B2").getBBox0().r1, ws.getRange2("B2").getBBox0().c1);
			assert.ok(1, "Trace precedents from merged range(A1:B2), three times");
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.ok(1, "Trace dependents from merged range(A1:B2), three times");
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, E1Index), 1, "A1<-E1");
			assert.strictEqual(traceManager._getDependents(E1Index, A1Index), 1, "E1->A1");
			assert.strictEqual(traceManager._getPrecedents(E1Index, H1Index), 1, "E1<-H1");
			assert.strictEqual(traceManager._getPrecedents(A2Index, E2Index), undefined, "A2<-E2");
			assert.strictEqual(traceManager._getPrecedents(E2Index, H2Index), undefined, "E2<-H2");
			assert.strictEqual(traceManager._getPrecedents(B1Index, F1Index), undefined, "B1<-F1");
			assert.strictEqual(traceManager._getPrecedents(F1Index, I1Index), undefined, "F1<-I1");
			assert.strictEqual(traceManager._getPrecedents(B2Index, F2Index), undefined, "B2<-F2");
			assert.strictEqual(traceManager._getPrecedents(F2Index, I2Index), undefined, "F2<-I2");
			assert.strictEqual(traceManager._getPrecedents(A9Index, A6Index), 1, "A9<-A6");
			assert.strictEqual(traceManager._getDependents(A6Index, A9Index), 1, "A6->A9");
			assert.strictEqual(traceManager._getPrecedents(A6Index, A1Index), 1, "A6<-A1");
			assert.strictEqual(traceManager._getDependents(A1Index, A6Index), 1, "A1->A6");
			assert.strictEqual(traceManager._getPrecedents(A7Index, A2Index), undefined, "A7<-A2");
			assert.strictEqual(traceManager._getPrecedents(B6Index, B1Index), undefined, "B6<-B1");
			assert.strictEqual(traceManager._getPrecedents(B7Index, B2Index), undefined, "B7<-B2");

			// clear all
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Mixed tests\"", function (assert) {
			ws.getRange2("A1:J20").cleanAll();
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

			assert.ok(1, "Trace precedents from B1");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A1Index, A10ExternalIndex), undefined, "A1<-OtherSheet!A10 === undefined");
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1, "B1<-A1");
			if (traceManager.checkPrecedentExternal(B1Index)) {
				assert.ok(traceManager.precedentsExternal[B1Index][A10ExternalIndex], "B1<-OtherSheet!A10");

				assert.strictEqual(traceManager.checkPrecedentExternal(A1Index), undefined, "From A1 external precedent don't exist");
				assert.strictEqual(traceManager.checkPrecedentExternal(C1Index), undefined, "From C1 external precedent don't exist");
			}

			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined, "C1<-B1 === undefined");

			api.asc_TracePrecedents();
			assert.ok(traceManager.precedentsExternal[A1Index][A10ExternalIndex], "A1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1, "B1<-A1");
			assert.ok(traceManager.precedentsExternal[B1Index][A10ExternalIndex], "B1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined, "C1<-B1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(C1Index, A1ExternalIndex), undefined, "C1<-OtherSheet!A1 === undefined");
			assert.strictEqual(traceManager._getPrecedents(C1Index, A10ExternalIndex), undefined, "C1<-OtherSheet!A10 === undefined");
			
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			api.asc_TracePrecedents();
			assert.ok(traceManager.precedentsExternal[A1Index][A10ExternalIndex], "A1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1, "B1<-A1");
			assert.ok(traceManager.precedentsExternal[B1Index][A10ExternalIndex], "B1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), undefined, "C1<-B1 === undefined");
			assert.strictEqual(traceManager.checkPrecedentExternal(C1Index), undefined, "From C1 external precedent don't exist");

			assert.ok(1, "Trace dependents from B1");
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			api.asc_TraceDependents();
			assert.ok(traceManager.precedentsExternal[A1Index][A10ExternalIndex], "A1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getPrecedents(B1Index, A1Index), 1, "B1<-A1");
			assert.ok(traceManager.precedentsExternal[B1Index][A10ExternalIndex], "B1<-OtherSheet!A10");
			assert.ok(traceManager.dependentsExternal[C1Index][A1ExternalIndex], "C1<-OtherSheet!A1");
			// assert.ok(traceManager.precedentsExternal[C1Index][A10ExternalIndex], "C1<-OtherSheet!A10");
			assert.strictEqual(traceManager._getDependents(B1Index, C1Index), 1, "B1->C1");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			ws.selectionRange.ranges = [ws.getRange2("C1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C1").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			assert.ok(1, "Trace precedents from C1");
			api.asc_TracePrecedents();

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace dependents from A1");
			api.asc_TraceDependents();
			api.asc_TraceDependents();

			assert.strictEqual(traceManager._getDependents(A1Index, B1Index), 1, "A1->B1");
			assert.strictEqual(traceManager._getPrecedents(C1Index, B1Index), 1, "C1<-B1");
			assert.ok(traceManager.precedentsExternal[C1Index][A10ExternalIndex], "C1<-OtherSheet!A10");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
			ws.getRange2("A1:J20").cleanAll();
			
			ws.getRange2("A1").setValue("1");
			ws.getRange2("C1").setValue("=A1+A2");
			ws.getRange2("E1").setValue("=C1+E2");
			ws.getRange2("G1").setValue("=E1");
			ws.getRange2("A4").setValue("1");
			ws.getRange2("C4").setValue("=A4+A5");
			ws.getRange2("E4").setValue("=C4+E5");
			ws.getRange2("G4").setValue("=E4");

			let E1Index = AscCommonExcel.getCellIndex(ws.getRange2("E1").bbox.r1, ws.getRange2("E1").bbox.c1),
				E2Index = AscCommonExcel.getCellIndex(ws.getRange2("E2").bbox.r1, ws.getRange2("E2").bbox.c1),
				G1Index = AscCommonExcel.getCellIndex(ws.getRange2("G1").bbox.r1, ws.getRange2("G1").bbox.c1),
				A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				A5Index = AscCommonExcel.getCellIndex(ws.getRange2("A5").bbox.r1, ws.getRange2("A5").bbox.c1),
				C4Index = AscCommonExcel.getCellIndex(ws.getRange2("C4").bbox.r1, ws.getRange2("C4").bbox.c1),
				E4Index = AscCommonExcel.getCellIndex(ws.getRange2("E4").bbox.r1, ws.getRange2("E4").bbox.c1),
				E5Index = AscCommonExcel.getCellIndex(ws.getRange2("E5").bbox.r1, ws.getRange2("E5").bbox.c1),
				G4Index = AscCommonExcel.getCellIndex(ws.getRange2("G4").bbox.r1, ws.getRange2("G4").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("C1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C1").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			assert.ok(1, "Trace dependents from C1. Checking two independent lines");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), undefined, "G1<-E1 === undefined. First line");
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), undefined, "G4<-E4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), undefined, "E4<-E5 === undefined. Second line");

			ws.selectionRange.ranges = [ws.getRange2("G1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("G1").getBBox0().r1, ws.getRange2("G1").getBBox0().c1);

			assert.ok(1, "Trace dependents from G1. Checking two independent lines");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1. First line");
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), undefined, "C4->E4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), undefined, "G4<-E4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), undefined, "E4<-E5 === undefined. Second line");

			ws.selectionRange.ranges = [ws.getRange2("C4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("C4").getBBox0().r1, ws.getRange2("C4").getBBox0().c1);

			assert.ok(1, "Trace dependents from C4. Checking two independent lines");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1. First line");	
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), undefined, "G4<-E4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), undefined, "E4<-E5 === undefined. Second line");

			ws.selectionRange.ranges = [ws.getRange2("G4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("G4").getBBox0().r1, ws.getRange2("G4").getBBox0().c1);

			assert.ok(1, "Trace dependents from G4. Checking two independent lines");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1. First line");
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), 1, "G4<-E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), undefined, "E4<-E5 === undefined. Second line");

			assert.ok(1, "Trace dependents from G4. Checking two independent lines");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1. First line");
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), 1, "G4<-E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), 1, "E4<-E5. Second line");
			assert.strictEqual(traceManager._getPrecedents(C4Index, A4Index), undefined, "C4<-A4 === undefined. Second line");
			assert.strictEqual(traceManager._getPrecedents(C4Index, A5Index), undefined, "C4<-A5 === undefined. Second line");

			assert.ok(1, "Trace dependents from G4. Checking two independent lines");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getDependents(C1Index, E1Index), 1, "C1->E1. First line");
			assert.strictEqual(traceManager._getPrecedents(G1Index, E1Index), 1, "G1<-E1. First line");
			assert.strictEqual(traceManager._getPrecedents(E1Index, E2Index), undefined, "E1<-E2 === undefined. First line");
			assert.strictEqual(traceManager._getDependents(C4Index, E4Index), 1, "C4->E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(G4Index, E4Index), 1, "G4<-E4. Second line");
			assert.strictEqual(traceManager._getPrecedents(E4Index, E5Index), 1, "E4<-E5. Second line");
			assert.strictEqual(traceManager._getPrecedents(C4Index, A4Index), 1, "C4<-A4. Second line");
			assert.strictEqual(traceManager._getPrecedents(C4Index, A5Index), 1, "C4<-A5. Second line");
			
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		QUnit.test("Test: \"Dependents perfomance tests\"", function (assert) {
			let bbox;
			// -------------- precedents --------------
			ws.getRange2("A1:Z10000").cleanAll();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
				A50Index = AscCommonExcel.getCellIndex(ws.getRange2("A50").bbox.r1, ws.getRange2("A50").bbox.c1),
				AA500Index = AscCommonExcel.getCellIndex(ws.getRange2("AA500").bbox.r1, ws.getRange2("AA500").bbox.c1);

			ws.getRange2("A1").setValue("2");
			bbox = ws.getRange2("A2:AA500").bbox;
			ws.getRange2("A2:AA500").setValue("=A1:C1", undefined, undefined, bbox);

			ws.selectionRange.ranges = [ws.getRange2("A1").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A1").getBBox0().r1, ws.getRange2("A1").getBBox0().c1);

			assert.ok(1, "Trace 13500 dependents from A1");
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A1Index, A50Index), 1, "A1->A50");
			assert.strictEqual(traceManager._getDependents(A1Index, AA500Index), 1, "A1->AA500");

			assert.ok(1, "Trace 13500 dependents from A2:AA500");
			// api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.dependent);  // ~10000ms

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);	
		});
		QUnit.test("Test: \"Precedents areas perfomance tests\"", function (assert) {
			let bbox;
			// -------------- precedents --------------
			ws.getRange2("A1:Z10000").cleanAll();

			ws.selectionRange.ranges = [ws.getRange2("I2").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("I2").getBBox0().r1, ws.getRange2("I2").getBBox0().c1);

			let I2Index = AscCommonExcel.getCellIndex(ws.getRange2("I2").bbox.r1, ws.getRange2("I2").bbox.c1),
				A2Index = AscCommonExcel.getCellIndex(ws.getRange2("A2").bbox.r1, ws.getRange2("A2").bbox.c1),
				C2Index = AscCommonExcel.getCellIndex(ws.getRange2("C2").bbox.r1, ws.getRange2("C2").bbox.c1),
				F2Index = AscCommonExcel.getCellIndex(ws.getRange2("F2").bbox.r1, ws.getRange2("F2").bbox.c1),
				A102Index = AscCommonExcel.getCellIndex(ws.getRange2("A102").bbox.r1, ws.getRange2("A102").bbox.c1),
				C102Index = AscCommonExcel.getCellIndex(ws.getRange2("C102").bbox.r1, ws.getRange2("C102").bbox.c1),
				F102Index = AscCommonExcel.getCellIndex(ws.getRange2("F102").bbox.r1, ws.getRange2("F102").bbox.c1),
				A1002Index = AscCommonExcel.getCellIndex(ws.getRange2("A1002").bbox.r1, ws.getRange2("A1002").bbox.c1),
				C1002Index = AscCommonExcel.getCellIndex(ws.getRange2("C1002").bbox.r1, ws.getRange2("C1002").bbox.c1),
				F1002Index = AscCommonExcel.getCellIndex(ws.getRange2("F1002").bbox.r1, ws.getRange2("F1002").bbox.c1),
				A2002Index = AscCommonExcel.getCellIndex(ws.getRange2("A2002").bbox.r1, ws.getRange2("A2002").bbox.c1),
				C2002Index = AscCommonExcel.getCellIndex(ws.getRange2("C2002").bbox.r1, ws.getRange2("C2002").bbox.c1),
				F2002Index = AscCommonExcel.getCellIndex(ws.getRange2("F2002").bbox.r1, ws.getRange2("F2002").bbox.c1);

			ws.getRange2("F2").setValue("=SUM(C2:E2)");

			let ctrlPress = false, 
				nIndex = 5000;
				// nIndex = 10;

			let oCanPromote = ws.getRange2("F2").canPromote(/*bCtrl*/ctrlPress, /*bVertical*/1, /*fill index*/ nIndex);
			ws.getRange2("F2").promote(ctrlPress, 1, nIndex, oCanPromote);

			bbox = ws.getRange2("I2:I6").bbox;
			ws.getRange2("I2:I6").setValue("=SUM($F$2:$F$2004,$C$2:$C$2004)*($A$2:$A$2004)", undefined, undefined, bbox);

			assert.ok(1, "Trace precedents from I2");
			api.asc_TracePrecedents();
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$2:$A$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$C$2:$C$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$F$2:$F$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["C2:E2"]), "undefined");
			assert.ok(1, "Trace 2000*3 precedents from I2");
			api.asc_TracePrecedents();
			assert.strictEqual(typeof(traceManager.precedentsAreas["$A$2:$A$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$C$2:$C$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["$F$2:$F$2004"]), "object");
			assert.strictEqual(typeof(traceManager.precedentsAreas["C2:E2"]), "object");
			// assert.strictEqual(typeof(traceManager.precedentsAreas["C1000:E1000"]), "object");
			api.asc_TracePrecedents();	// old: ~2800ms, new: ~280ms
			// api.asc_TracePrecedents();	// old: ~5400ms, new: ~280ms
			// api.asc_TracePrecedents();	// old: ~8400ms, new: ~280ms
			// api.asc_TracePrecedents();	// old: ~10800ms, new: ~280ms
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.precedent);

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		// 	let bbox;
		// 	// -------------- precedents --------------
		// 	ws.getRange2("A1:Z10000").cleanAll();

		// 	let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1),
		// 		A50Index = AscCommonExcel.getCellIndex(ws.getRange2("A50").bbox.r1, ws.getRange2("A50").bbox.c1),
		// 		AA500Index = AscCommonExcel.getCellIndex(ws.getRange2("AA500").bbox.r1, ws.getRange2("AA500").bbox.c1);

		// // / / // TODO check all the formulas, write them in a string according to the number of arguments and call the formulas.
		// // Also, you can divide formulas into having a return type and not

		// 	// console.log(AscCommonExcel.cFormulaFunctionGroup);
		// 	// console.log(AscCommonExcel.cReturnFormulaType);

		// 	let valueType = [],						// 0
		// 		value_replace_areaType = [],		// 1
		// 		arrayType = [],						// 2
		// 		area_to_refType = [],				// 3
		// 		replace_only_arrayType = [],		// 4
		// 		setArrayRefAsArgType = [],			// 5
		// 		noRetType = [],
		// 		withArrIndexes = [],
		// 		withoutArrIndexes = [],
		// 		retType = [], str = "";

		// 	console.log(AscCommonExcel.cFormulaFunctionGroup);
		// 	// TODO test all type functions
		// 	for (let index in AscCommonExcel.cFormulaFunctionGroup) {
		// 		let array = AscCommonExcel.cFormulaFunctionGroup[index];
		// 		for (let i = 0; i < array.length; i++) {
		// 			if (array[i].prototype.returnValueType !== null) {
		// 				retType.push(array[i]);
		// 			} else {
		// 				noRetType.push(array[i]);
		// 			}
		// 		}
		// 	}

		// 	let fullArrrayWithF = [];
		// 	for (let index in AscCommonExcel.cFormulaFunctionGroup) {
		// 		let array = AscCommonExcel.cFormulaFunctionGroup[index];
		// 		for (let i = 0; i < array.length; i++) {
		// 			let formula = array[i];
		// 			if (formula.prototype && formula.prototype.name) {
		// 				let minArg = formula.prototype.argumentsMin ? formula.prototype.argumentsMin : 1, maxArg = formula.prototype.argumentsMax ? formula.prototype.argumentsMax : 1;
		// 				let formulas = [];
		// 				let arrayMinArg = new Array(minArg);

		// 				maxArg = maxArg > 10 ? 10 : maxArg;
		// 				arrayMinArg.fill("1", 0, minArg);

		// 				for (let i = 0; i < minArg; i++) {
		// 					let curFormula = arrayMinArg.slice();
		// 					curFormula[i] = "A:A";
		// 					formulas.push("=" + formula.prototype.name + "(" + curFormula.join(";") + ")" + "\t");
		// 					// formulas.push(curFormula);
		// 				}
		// 				for (let i = minArg; i < maxArg; i++) {
		// 					arrayMinArg.push("1");
		// 					let curFormula = arrayMinArg.slice();
		// 					curFormula[i] = "A:A";
		// 					formulas.push("=" + formula.prototype.name + "(" + curFormula.join(";") + ")" + "\t");
		// 					// console.log(formulas);
		// 				}
		// 				formulas.push("\n");
		// 				fullArrrayWithF.push(formulas.join(""));
			
		// 			}
		// 		}
		// 	}
		// 	console.log(fullArrrayWithF.join(""));
		// 	// console.log(noRetType);
		// 	// console.log(retType);
		// 	// console.log(withArrIndexes);
		// 	// console.log(withoutArrIndexes);

		// 	ws.getRange2("A1").setValue("2");

		// 	// clear traces
		// 	api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		// });
		QUnit.test("Test: \"Formulas tests\"", function (assert) {
			let bbox;
			ws.getRange2("A:A").cleanAll();

			let A1Index = AscCommonExcel.getCellIndex(ws.getRange2("A1").bbox.r1, ws.getRange2("A1").bbox.c1);
			let noRetType = [],
				withArrIndexes = [],
				withoutArrIndexes = [],
				haveRetType = [];

			// this formulas doesn't exist in excel 2016
			let versionExceptions = [
				"FORECAST", "PHONETIC", "TEXTJOIN", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT", "MAXIFS", 
				"MINIFS", "JIS", "RANDARRAY", "SEQUENCE", "CHOOSECOLS", "CHOOSEROWS", "DROP", "EXPAND", 
				"FILTER", "SORT", "TAKE", "UNIQUE", "XLOOKUP", "VSTACK", "HSTACK", "TOROW", "TOCOL", "WRAPROWS", "IFS", "SWITCH"
			];
			// this formulas different in behavior in arrayIndexes
			let behaviourExceptions = [
				"GROWTH", "LINEST", "LOGEST", "TREND", "CUBEMEMBER", "CUBESET", "CUBEVALUE",
				"AREAS", "GETPIVOTDATA", "NETWORKDAYS.INTL", "WORKDAY.INTL", "MUNIT", "CHOOSE",
				"MATCH", "TRANSPOSE", "TYPE", "IF", "IFERROR", "IFNA"
			];

			// for (let index in AscCommonExcel.cFormulaFunctionGroup) {
			// 	let array = AscCommonExcel.cFormulaFunctionGroup[index];
			// 	for (let i = 0; i < array.length; i++) {
			// 		if (versionExceptions.includes(array[i].prototype.name) || behaviourExceptions.includes(array[i].prototype.name)) {
			// 			continue
			// 		}
			// 		// fill row with current formula
			// 	}
			// }

			ws.getRange2("B3").setValue("=SUM(A:A)");
			let A3Index = AscCommonExcel.getCellIndex(ws.getRange2("A3").bbox.r1, ws.getRange2("A3").bbox.c1),
				B3Index = AscCommonExcel.getCellIndex(ws.getRange2("B3").bbox.r1, ws.getRange2("B3").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("B3").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B3").getBBox0().r1, ws.getRange2("B3").getBBox0().c1);

			assert.ok(1, "Trace dependents from B3");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B3Index, A1Index), "A:A", "B3<-A1. Line should be directed to the range header");
			assert.strictEqual(traceManager._getPrecedents(B3Index, A3Index), undefined, "B3<-A3. Line shouldn't be directed to the opposite cell in range A:A");
			assert.strictEqual(traceManager.precedentsAreas ? typeof(traceManager.precedentsAreas["A:A"]) : undefined, "object", "A:A should exist as a range");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);


			ws.getRange2("B4").setValue("=SIN(A:A)");
			let A4Index = AscCommonExcel.getCellIndex(ws.getRange2("A4").bbox.r1, ws.getRange2("A4").bbox.c1),
				B4Index = AscCommonExcel.getCellIndex(ws.getRange2("B4").bbox.r1, ws.getRange2("B4").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("B4").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B4").getBBox0().r1, ws.getRange2("B4").getBBox0().c1);
			
			assert.ok(1, "Trace dependents from B4");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B4Index, A1Index), undefined, "B4<-A1. Line shouldn't be directed to the range header");
			assert.strictEqual(traceManager._getPrecedents(B4Index, A4Index), 1, "B4<-A4. Line should be directed to the opposite cell in range A:A");
			assert.strictEqual(traceManager.precedentsAreas ? typeof(traceManager.precedentsAreas["A:A"]) : undefined, undefined, "A:A shouldn't exist as a range");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
			
			
			ws.getRange2("B5").setValue("=NPV(1;A:A)");
			let A5Index = AscCommonExcel.getCellIndex(ws.getRange2("A5").bbox.r1, ws.getRange2("A5").bbox.c1),
				B5Index = AscCommonExcel.getCellIndex(ws.getRange2("B5").bbox.r1, ws.getRange2("B5").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("B5").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B5").getBBox0().r1, ws.getRange2("B5").getBBox0().c1);
			
			assert.ok(1, "Trace dependents from B5");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B5Index, A1Index), "A:A", "B5<-A1. Line should be directed to the range header");
			assert.strictEqual(traceManager._getPrecedents(B5Index, A5Index), undefined, "B5<-A5. Line shouldn't be directed to the opposite cell in range A:A");
			assert.strictEqual(traceManager.precedentsAreas ? typeof(traceManager.precedentsAreas["A:A"]) : undefined, "object", "A:A should exist as a range");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);


			ws.getRange2("B6").setValue("=NPV(A:A;1)");
			let A6Index = AscCommonExcel.getCellIndex(ws.getRange2("A6").bbox.r1, ws.getRange2("A6").bbox.c1),
				B6Index = AscCommonExcel.getCellIndex(ws.getRange2("B6").bbox.r1, ws.getRange2("B6").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("B6").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B6").getBBox0().r1, ws.getRange2("B6").getBBox0().c1);
			
			assert.ok(1, "Trace dependents from B6");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B6Index, A1Index), undefined, "B6<-A1. Line shouldn't be directed to the range header");
			assert.strictEqual(traceManager._getPrecedents(B6Index, A6Index), 1, "B6<-A6. Line should be directed to the opposite cell in range A:A");
			assert.strictEqual(traceManager.precedentsAreas ? typeof(traceManager.precedentsAreas["A:A"]) : undefined, undefined, "A:A shouldn't exist as a range");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);

			
			ws.getRange2("B7").setValue("=NPV(A:A;A:A)");
			let A7Index = AscCommonExcel.getCellIndex(ws.getRange2("A7").bbox.r1, ws.getRange2("A7").bbox.c1),
				B7Index = AscCommonExcel.getCellIndex(ws.getRange2("B7").bbox.r1, ws.getRange2("B7").bbox.c1);

			ws.selectionRange.ranges = [ws.getRange2("B7").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("B7").getBBox0().r1, ws.getRange2("B7").getBBox0().c1);
			
			assert.ok(1, "Trace dependents from B7");
			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(B7Index, A1Index), "A:A", "B7<-A1. Line should be directed to both ways. Line to the range header");
			assert.strictEqual(traceManager._getPrecedents(B7Index, A7Index), 1, "B7<-A7. Line should be directed to both ways. Line to the opposite cell");
			assert.strictEqual(traceManager.precedentsAreas ? typeof(traceManager.precedentsAreas["A:A"]) : undefined, "object", "A:A should exist as a range");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
		});
		// QUnit.test("Test: \"Interface tests\"", function (assert) {
		// 	let bbox;
		// 	// trace dependents/precedents -> click on interface element -> check dependencies
		// 	// let a = new api.asc_CDefName();
		// 	// AddComment -> false
		// 	// asc_Paste -> false
		// 	// asc_PasteData -> true*
		// });
		QUnit.test("Test: \"Recursive formulas\"", function (assert) {
			// Case #1: Precedent trace for formula - =A100+1. A dot or line shouldn't be traced
			ws.getRange2("A100").setValue("=A100+1");
			let A100Index = AscCommonExcel.getCellIndex(ws.getRange2("A100").bbox.r1, ws.getRange2("A100").bbox.c1);
			ws.selectionRange.ranges = [ws.getRange2("A100").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A100").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A100Index, A100Index), undefined, "Precedent trace. Formula A100+1. A100. Dot shouldn't be show");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
			// Case #2: Dependent trace for formula - =A100+1. A dot or line shouldn't be traced
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A100Index, A100Index), undefined, "Dependent trace. Formula A100+1. A100. Dot shouldn't be show");
			// Case #3: Precedent trace for formula - A100: A100+B100, B100: B100+C100. A100<-B100.  Arrow without dot in A100 cell.
			ws.getRange2("A100").setValue("=A100+B100");
			ws.getRange2("B100").setValue("=B100+C100");
			A100Index = AscCommonExcel.getCellIndex(ws.getRange2("A100").bbox.r1, ws.getRange2("A100").bbox.c1);
			let B100Index = AscCommonExcel.getCellIndex(ws.getRange2("B100").bbox.r1, ws.getRange2("B100").bbox.c1);
			ws.selectionRange.ranges = [ws.getRange2("A100").getBBox0()];
			ws.selectionRange.setActiveCell(ws.getRange2("A100").getBBox0().r1, ws.getRange2("C1").getBBox0().c1);

			api.asc_TracePrecedents();
			assert.strictEqual(traceManager._getPrecedents(A100Index, B100Index), 1, "Precedent trace. Formula: A100: A100+B100, B100: B100+C100. A100<-B100. Arrow to A100 without dot in A100 cell.");
			assert.strictEqual(traceManager._getPrecedents(A100Index, A100Index), undefined, "Precedent trace. Formula: A100: A100+B100, B100: B100+C100. A100. Dot shouldn't be show");

			// clear traces
			api.asc_RemoveTraceArrows(Asc.c_oAscRemoveArrowsType.all);
			// Case #4: Dependent trace for formula - A100: A100+B100, B100: B100+C100. A100->B100.  Arrow without dot in B100 cell.
			api.asc_TraceDependents();
			assert.strictEqual(traceManager._getDependents(A100Index, A100Index), undefined, "Dependent trace. Formula: A100: A100+B100, B100: B100+C100. A100. Dot shouldn't be show");
		});
	}

	QUnit.module("FormulaTrace");

	function startTests() {
		QUnit.start();
		traceTests();
	}

	startTests();
});
