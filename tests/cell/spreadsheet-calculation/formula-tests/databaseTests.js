/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	// Mocks for API Testing
	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
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

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function () {
	};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (oFromRange, subType, callback) {
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

	// Init basic data
	const parserFormula = AscCommonExcel.parserFormula;
	const g_oIdCounter = AscCommon.g_oIdCounter;
	let oParser, wb, ws, sData = AscCommon.getEmpty();
	if (AscCommon.c_oSerFormat.Signature === sData.substring(0, AscCommon.c_oSerFormat.Signature.length)) {

		Asc.spreadsheet_api.prototype._init = function() {
			this.isLoadFullApi = true;
		};


		let api = new Asc.spreadsheet_api({
			'id-view': 'editor_sdk'
		});
		api.FontLoader = {
			LoadDocumentFonts: function () {
			}
		};

		let docInfo = new Asc.asc_CDocInfo();
		docInfo.asc_putTitle("TeSt.xlsx");
		api.DocInfo = docInfo;


		window["Asc"]["editor"] = api;
		AscCommon.g_oTableId.init(api);
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
		api.initCollaborativeEditing({});
		wb = new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);
		api.wbModel = wb;
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		AscCommon.History.init(wb);
		// There is no operation with history, disabled to avoid unnecessary serializations.
		AscCommon.History.TurnOff();
		wb.maxDigitWidth = 7;
		wb.paddingPlusBorder = 5;

		api.initCollaborativeEditing({});

		if (this.User) {
			g_oIdCounter.Set_UserId(this.User.asc_getId());
		}

		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
		g_oIdCounter.Set_Load(false);

		const oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read(sData, wb);
		ws = wb.getWorksheet(wb.getActive());
		AscCommonExcel.getFormulasInfo();
	}

	// Init basic functions
	/**
	 * Function creates table or edit existed table.
	 * * Creating happened when the function was called for the first time in the code totally.
	 * In that case, you should set range coordinates like you're creating table, select needed range.
	 * * For editing table you should consider that you have constant row with header and one data row as minimum required.
	 * In that case your minimum row coordinates must be equal 2 rows.
	 * For filling data use cells A601:L6**
	 * @param {number} r1
	 * @param {number} c1
	 * @param {number} r2
	 * @param {number} c2
	 * @returns {TablePart}
	 */
	function getTableType(r1, c1, r2, c2) {
		const range = new window["Asc"].Range(c1, r1, c2, r2);
		const  tableName = 'Table1';
		const tableParts = ws.TableParts;
		/** @type TablePart **/
		const foundedTable = tableParts.find(function(table) { return table.DisplayName === tableName});
		if (foundedTable) {
			// Change table
			ws.autoFilters.changeTableRange(tableName, range);
			return foundedTable;
		}
		// Create table
		const tableOptFormat = ws.autoFilters.getAddFormatTableOptions(range);
		const styleName = 'TableStyleMedium2';
		ws.autoFilters.addAutoFilter(styleName, range, tableOptFormat);

		return tableParts[tableParts.length - 1]
	}
	function getSecondSheet () {
		let ws2 = wb.getWorksheetByName('Sheet2');
		if (!ws2) {
			ws2 = wb.createWorksheet(null, 'Sheet2');
		}

		return ws2;
	}
	function initDefNames() {
		const defName = new Asc.asc_CDefName('TestName', ws.getName() + '!$A$201');
		const defName2 = new Asc.asc_CDefName('TestName1', ws.getName() + '!$A$202');
		const defName3 = new Asc.asc_CDefName('TestName2', ws.getName() + '!$A$203');
		const defName4 = new Asc.asc_CDefName('TestName3', ws.getName() + '!$A$204');
		const defName5 = new Asc.asc_CDefName('TestName4', ws.getName() + '!$A$205');
		const defNameArea = new Asc.asc_CDefName('TestNameArea', ws.getName() + '!$A$206:$A$207');
		const defNameArea2 = new Asc.asc_CDefName('TestNameArea2', ws.getName() + '!$A$208:$B$208');
		const ws2 = getSecondSheet();
		const defName3D = new Asc.asc_CDefName('TestName3D', ws2.getName() + '!$A$11');
		const defName3D2 = new Asc.asc_CDefName('TestName3D1', ws2.getName() + '!$A$12');
		const defName3D3 = new Asc.asc_CDefName('TestName3D2', ws2.getName() + '!$A$13');
		const defName3D4 = new Asc.asc_CDefName('TestName3D3', ws2.getName() + '!$A$14');
		const defName3D5 = new Asc.asc_CDefName('TestName3D4', ws2.getName() + '!$A$15');
		const defNameArea3D = new Asc.asc_CDefName('TestNameArea3D', ws2.getName() + '!$A$16:$A$17');
		const defNameArea3D2 = new Asc.asc_CDefName('TestNameArea3D2', ws2.getName() + '!$A$18:$B$18');

		wb.editDefinesNames(null, defName);
		wb.editDefinesNames(null, defName2);
		wb.editDefinesNames(null, defName3);
		wb.editDefinesNames(null, defName4);
		wb.editDefinesNames(null, defName5);
		wb.editDefinesNames(null, defNameArea);
		wb.editDefinesNames(null, defNameArea2);
		wb.editDefinesNames(null, defName3D);
		wb.editDefinesNames(null, defName3D2);
		wb.editDefinesNames(null, defName3D3);
		wb.editDefinesNames(null, defName3D4);
		wb.editDefinesNames(null, defName3D5);
		wb.editDefinesNames(null, defNameArea3D);
		wb.editDefinesNames(null, defNameArea3D2);
	}
	function putDataForDatabase() {
		ws.getRange2("A1").setValue("Tree");
		ws.getRange2("A2").setValue("Apple");
		ws.getRange2("A3").setValue("Pear");

		ws.getRange2("A4").setValue("Tree");

		ws.getRange2("A5").setValue("Apple");
		ws.getRange2("A6").setValue("Pear");
		ws.getRange2("A7").setValue("Cherry");
		ws.getRange2("A8").setValue("Apple");
		ws.getRange2("A9").setValue("Pear");
		ws.getRange2("A10").setValue("Apple");


		ws.getRange2("B1").setValue("Height");
		ws.getRange2("B2").setValue(">10");
		ws.getRange2("B3").setValue("");

		ws.getRange2("B4").setValue("Height");

		ws.getRange2("B5").setValue("18");
		ws.getRange2("B6").setValue("12");
		ws.getRange2("B7").setValue("13");
		ws.getRange2("B8").setValue("14");
		ws.getRange2("B9").setValue("9");
		ws.getRange2("B10").setValue("8");


		ws.getRange2("C1").setValue("Age");
		ws.getRange2("C2").setValue("");
		ws.getRange2("C3").setValue("");

		ws.getRange2("C4").setValue("Age");

		ws.getRange2("C5").setValue("20");
		ws.getRange2("C6").setValue("12");
		ws.getRange2("C7").setValue("14");
		ws.getRange2("C8").setValue("15");
		ws.getRange2("C9").setValue("8");
		ws.getRange2("C10").setValue("9");


		ws.getRange2("C1").setValue("Age");
		ws.getRange2("C2").setValue("");
		ws.getRange2("C3").setValue("");

		ws.getRange2("C4").setValue("Age");

		ws.getRange2("C5").setValue("20");
		ws.getRange2("C6").setValue("12");
		ws.getRange2("C7").setValue("14");
		ws.getRange2("C8").setValue("15");
		ws.getRange2("C9").setValue("8");
		ws.getRange2("C10").setValue("9");


		ws.getRange2("D1").setValue("Yield");
		ws.getRange2("D2").setValue("");
		ws.getRange2("D3").setValue("");

		ws.getRange2("D4").setValue("Yield");

		ws.getRange2("D5").setValue("14");
		ws.getRange2("D6").setValue("10");
		ws.getRange2("D7").setValue("9");
		ws.getRange2("D8").setValue("10");
		ws.getRange2("D9").setValue("8");
		ws.getRange2("D10").setValue("6");


		ws.getRange2("E1").setValue("Profit");
		ws.getRange2("E2").setValue("");
		ws.getRange2("E3").setValue("");

		ws.getRange2("E4").setValue("Profit");

		ws.getRange2("E5").setValue("105");
		ws.getRange2("E6").setValue("96");
		ws.getRange2("E7").setValue("105");
		ws.getRange2("E8").setValue("75");
		ws.getRange2("E9").setValue("76.8");
		ws.getRange2("E10").setValue("45");

		ws.getRange2("F1").setValue("Height");
		ws.getRange2("F2").setValue("<16");
		ws.getRange2("F3").setValue("");
	}

	wb.dependencyFormulas.lockRecal();
	getTableType(599, 0, 599, 0); // Init table
	initDefNames();

	QUnit.module('Database formulas');
	QUnit.test('Test: "DAVERAGE"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");

		ws.getRange2("G1").setValue("noname");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 12
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, String, Area. Return 12');
		// Case #2: Area, String, Area. Return 13
		oParser = new parserFormula('DAVERAGE(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Positive case: Area, String, Area. Return 13');
		// Case #3: Area, String, Area. Return 13
		oParser = new parserFormula('DAVERAGE(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Positive case: Area, String, Area. Return 13');
		// Case #4: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #5: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DAVERAGE(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #6: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #7: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #8: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #9: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '17.5', 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #10: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #11: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #12: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #13: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #14: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #15: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DAVERAGE(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #16: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #17: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #18: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #19: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #20: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #21: Formula. DAVERAGE nested inside SUM formula.
		oParser = new parserFormula('SUM(DAVERAGE(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DAVERAGE(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 22, 'Test: Positive case: Formula. DAVERAGE nested inside SUM formula.');
		// Case #22: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #23: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '10.5', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #24: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #25: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9.5, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');

		// Negative cases:

		ws.getRange2("G1").setValue("Profit");
		// Case #1: Area, String, Area. Return #DIV/0!
		oParser = new parserFormula('DAVERAGE(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');
		*/// Case #7: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DAVERAGE(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #8: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #12: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DAVERAGE(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #13: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #14: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DAVERAGE(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DAVERAGE(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #15: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.');*/
		// Case #16: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, A103, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, A103, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.');
		// Case #17: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #18: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #19: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #20: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #21: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DAVERAGE(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DAVERAGE(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #22: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.');
		// Case #23: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #24: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #25: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DAVERAGE(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DAVERAGE(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DAVERAGE(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DAVERAGE(A4:E10, "Yield", A:B)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAVERAGE(A4:E10, "Yield", A:B) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9.5, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #15, #17, #20
		// Different logic converting string type comparing with MS - Negative case #6
		// Parser must be false - Negative case #21, #14
	});

	QUnit.test('Test: "DCOUNT"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 1
		oParser = new parserFormula('DCOUNT(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area, String, Area. Return 1');
		// Case #2: Area, Empty, Area. Return 1
		oParser = new parserFormula('DCOUNT(A4:E10,, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10,, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area, Empty, Area. Return 1');
		// Case #3: Area, String, Area. Return 0
		oParser = new parserFormula('DCOUNT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');

		ws.getRange2("G1").setValue("noname");
		// Case #4: Area, String, Area. Return 6
		oParser = new parserFormula('DCOUNT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #5: Area, String, Area. Return 2
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Return 2');
		// Case #6: Area, String, Area. Return 6
		oParser = new parserFormula('DCOUNT(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #7: Area, String, Area. Return 6
		oParser = new parserFormula('DCOUNT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #8: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #9: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DCOUNT(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #10: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DCOUNT(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #11: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #12: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #13: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DCOUNT(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #14: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #15: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #16: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #17: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #18: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #19: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DCOUNT(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #20: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #21: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #22: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #23: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #24: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #25: Formula. DCOUNT nested inside SUM formula.
		oParser = new parserFormula('SUM(DCOUNT(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DCOUNT(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Formula. DCOUNT nested inside SUM formula.');
		// Case #26: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #27: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #28: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #29: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #30: Area, String, Area. Return 0
		ws.getRange2("G1").setValue("Profit");
		oParser = new parserFormula('DCOUNT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #31: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DCOUNT(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #32: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #33: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DCOUNT(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #34: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.');

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DCOUNT(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNT(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DCOUNT(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DCOUNT(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DCOUNT(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, A103, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, A103, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.');*/
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');*/
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNT(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DCOUNT(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DCOUNT(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DCOUNT(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DCOUNT(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DCOUNT(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DCOUNT(A4:E10, "Yield", A:B)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNT(A4:E10, "Yield", A:B) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #14, #15, #16, #18
		// Different logic converting type comparing with MS - Negative case #6, #17
		// Parser must be false - Negative case #19, #13.
	});

	QUnit.test('Test: "DCOUNTA"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 1
		oParser = new parserFormula('DCOUNTA(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area, String, Area. Return 1');
		// Case #2: Area, Empty, Area. Return 1
		oParser = new parserFormula('DCOUNTA(A4:E10,, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10,, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area, Empty, Area. Return 1');
		// Case #3: Area, String, Area. Return 0
		oParser = new parserFormula('DCOUNTA(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #4: Area, String, Area. Return 6
		ws.getRange2("G1").setValue("noname");
		oParser = new parserFormula('DCOUNTA(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #5: Area, String, Area. Return 2
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Return 2');
		// Case #6: Area, String, Area. Return 6
		oParser = new parserFormula('DCOUNTA(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #7: Area, String, Area. Return 6
		oParser = new parserFormula('DCOUNTA(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Return 6');
		// Case #8: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #9: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DCOUNTA(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #10: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #11: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #12: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #13: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #14: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #15: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #16: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #17: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #18: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #19: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DCOUNTA(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #20: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #21: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #22: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #23: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #24: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #25: Formula. DCOUNTA nested inside SUM formula.
		oParser = new parserFormula('SUM(DCOUNTA(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DCOUNTA(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Formula. DCOUNTA nested inside SUM formula.');
		// Case #26: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #27: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #28: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #29: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #30: Area, String, Area. Return 0
		ws.getRange2("G1").setValue("Profit");
		oParser = new parserFormula('DCOUNTA(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #31: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DCOUNTA(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #32: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #33: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DCOUNTA(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #34: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.');

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DCOUNTA(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DCOUNTA(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DCOUNTA(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DCOUNTA(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.
		/*oParser = new parserFormula('DCOUNTA(A4:E10, A103, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, A103, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.');*/
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');*/
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DCOUNTA(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DCOUNTA(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DCOUNTA(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DCOUNTA(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DCOUNTA(A4:E10, "Yield", A:B)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DCOUNTA(A4:E10, "Yield", A:B) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #14, #16, #18
		// Different logic converting type comparing with MS - Negative case #6, #15, #17
		// Parser must be false - Negative case #19, #13.
	});

	QUnit.test('Test: "DGET"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("text");
		ws.getRange2("A103").setValue("");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");

		// Table type. Use A601:L6**
		getTableType(599, 0, 601, 0);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("9"); // Column1
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("9");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("9"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("9"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Database, field as column name string, criteria range. Extract single yield value for Apple with Height>10 AND Height<16. Return 10.
		oParser = new parserFormula('DGET(A4:E10, "Yield", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Extract single yield value for Apple with Height>10 AND Height<16. Return 10.');
		// Case #2: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Extract using column index. Return 10.
		oParser = new parserFormula('DGET(A4:E10, 4, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 4, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Extract using column index. Return 10.');
		// Case #3: Area, String, Area. Database, field as different column name, criteria range. Extract single height value. Return 14.
		oParser = new parserFormula('DGET(A4:E10, "Height", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Height", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Extract single height value. Return 14.');
		// Case #4: Area, String, Area. Database, field as "Profit" column, criteria range. Return 75.
		oParser = new parserFormula('DGET(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range. Return 75.');
		// Case #5: Area, Number, Area. Database, field as column 2 (Height), criteria range. Extract using numeric index. Return 14.
		oParser = new parserFormula('DGET(A4:E10, 2, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 2, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range. Extract using numeric index. Return 14.');
		// Case #6: Area, String, Area. Database, field as "Age" column, criteria range. Extract single age value. Return 15.
		oParser = new parserFormula('DGET(A4:E10, "Age", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Age", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Extract single age value. Return 15.');
		// Case #7: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range. Return 10.
		oParser = new parserFormula('DGET(A4:E10, A100, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, A100, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range. Return 10.');
		// Case #8: Area, String, Name. Database, field as string, criteria as named range.
		oParser = new parserFormula('DGET(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as named range.');
		// Case #9: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		oParser = new parserFormula('DGET(A4:E10, 4, TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 4, TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');
		// Case #11: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DGET(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #12: Area, String, Table. Database, field as string, criteria as table reference.
		oParser = new parserFormula('DGET(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #13: Area, Formula, Area. Database, field from nested IF formula, criteria range. Return 10.
		oParser = new parserFormula('DGET(A4:E10, IF(TRUE,"Yield","Age"), A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, IF(TRUE,"Yield","Age"), A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range. Return 10.');
		// Case #14: Area, String, Formula. Database, field as string, criteria from nested IF formula. Return 10.
		oParser = new parserFormula('DGET(A4:E10, "Yield", IF(1=1,A1:F2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", IF(1=1,A1:F2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula. Return 10.');
		// Case #15: Formula. DGET nested inside SUM formula. Return 20 (10+10).
		oParser = new parserFormula('SUM(DGET(A4:E10,"Yield",A1:F2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DGET(A4:E10,"Yield",A1:F2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Formula. DGET nested inside SUM formula. Return 20 (10+10).');
		// Case #16: Area, Boolean, Area. Database, field as boolean TRUE (converted to 1 = Tree column), criteria range. Return "Apple".
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, TRUE, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, TRUE, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Apple', 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE (converted to 1 = Tree column), criteria range. Return "Apple".');*/
		// Case #17: Area, Array, Area. Database, field as single-element array, criteria range. Return 10.
		oParser = new parserFormula('DGET(A4:E10, {"Yield"}, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, {"Yield"}, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range. Return 10.');
		// Case #18: Area, String, Area. Database, field as "Tree" column (text value), criteria range. Extract text value. Return "Apple".
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Tree", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Tree", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Apple', 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text value), criteria range. Extract text value. Return "Apple".');*/

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DGET(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DGET(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DGET(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DGET(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DGET(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DGET(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DGET(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DGET(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DGET(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Yield", A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #NULL! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.
		oParser = new parserFormula('DGET(A4:E10, A103, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, A103, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #NULL! error.');
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
	/*	oParser = new parserFormula('DGET(A4:E10, "Yield", A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DGET(A4:E10, A102, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, A102, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DGET(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DGET(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DGET(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DGET(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #22: Area, String, Area. Return #NUM!
		oParser = new parserFormula('DGET(A4:E10, "Yield", A1:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A1:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Area, String, Area. Return #NUM!');
		// Case #23: Area, String, Area. Return #VALUE!
		oParser = new parserFormula('DGET(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Return #VALUE!');
		// Case #24: Area, String, Area. Return #NUM!
		ws.getRange2("G1").setValue("noname");
		oParser = new parserFormula('DGET(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Area, String, Area. Return #NUM!');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum column index (1 = Tree column). Return "Apple".
		// Different result with MS
		/*oParser = new parserFormula('DGET(A4:E10, 1, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 1, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Apple', 'Test: Bounded case: Area, Number, Area. Minimum column index (1 = Tree column). Return "Apple".');*/
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column). Return 75.
		oParser = new parserFormula('DGET(A4:E10, 5, A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, 5, A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column). Return 75.');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference. Return 10.
		oParser = new parserFormula('DGET(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DGET(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference. Return 10.');

		// TODO Need to fix

		// Different result comparing with MS - Positive case #16, #18, Bounded case #1
		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #6, #14, #16, #18
		// Parser must be false - Negative case #19, #13.
	});

	QUnit.test('Test: "DMAX"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D
		// Positive cases:

		// Case #1: Area, String, Area. Return 96
		oParser = new parserFormula('DMAX(A4:E10, "Profit", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Profit", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 96, 'Test: Positive case: Area, String, Area. Return 96');
		// Case #2: Area, String, Area. Return 0
		oParser = new parserFormula('DMAX(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #3: Area, String, Area. Return 20
		ws.getRange2("G1").setValue("noname");
		oParser = new parserFormula('DMAX(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Area. Return 20');
		// Case #4: Area, String, Area. Return 75
		oParser = new parserFormula('DMAX(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #5: Area, String, Area. Return 14
		oParser = new parserFormula('DMAX(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Return 14');
		// Case #6: Area, String, Area. Return 20
		oParser = new parserFormula('DMAX(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Area. Return 20');
		// Case #7: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #8: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DMAX(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #9: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DMAX(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #10: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DMAX(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #11: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DMAX(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #12: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DMAX(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #13: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DMAX(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #14: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DMAX(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #15: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DMAX(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #16: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DMAX(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #17: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DMAX(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #18: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DMAX(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #19: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DMAX(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #20: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #21: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #22: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DMAX(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #23: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #24: Formula. DMAX nested inside SUM formula.
		oParser = new parserFormula('SUM(DMAX(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DMAX(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Formula. DMAX nested inside SUM formula.');
		// Case #25: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DMAX(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #26: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #27: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DMAX(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #28: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DMAX(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #29: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DMAX(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #30: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DMAX(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #31: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #VALUE! error.');

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DMAX(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		/*oParser = new parserFormula('DMAX(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DMAX(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DMAX(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DMAX(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DMAX(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DMAX(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DMAX(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DMAX(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #22: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMAX(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DMAX(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DMAX(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 105, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DMAX(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMAX(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #14, #16, #18.
		// Different logic converting type comparing with MS - Negative case #6.
		// Parser must be false - Negative case #19, #13.

	});

	QUnit.test('Test: "DMIN"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 75
		oParser = new parserFormula('DMIN(A4:E10, "Profit", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Profit", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #2: Area, String, Area. Return 0
		oParser = new parserFormula('DMIN(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		ws.getRange2("G1").setValue("noname");
		// Case #3: Area, String, Area. Return 8
		oParser = new parserFormula('DMIN(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Return 8');
		// Case #4: Area, String, Area. Return 75
		oParser = new parserFormula('DMIN(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #5: Area, String, Area. Return 10
		oParser = new parserFormula('DMIN(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Area. Return 10');
		// Case #6: Area, String, Area. Return 8
		oParser = new parserFormula('DMIN(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Return 8');
		// Case #7: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #8: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DMIN(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #9: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DMIN(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #10: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #11: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DMIN(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #12: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DMIN(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #13: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DMIN(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #14: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DMIN(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #15: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DMIN(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #16: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DMIN(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #17: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DMIN(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #18: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DMIN(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #19: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DMIN(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #20: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #21: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #22: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #23: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #24: Formula. DMIN nested inside SUM formula.
		oParser = new parserFormula('SUM(DMIN(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DMIN(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Formula. DMIN nested inside SUM formula.');
		// Case #25: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #26: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #27: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #28: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DMIN(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #29: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #30: Area, String, Area. Database, field as "Tree" column (text values), criteria range.
		oParser = new parserFormula('DMIN(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range.');
		// Case #31: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.
		oParser = new parserFormula('DMIN(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.');

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DMIN(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DMIN(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DMIN(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DMIN(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DMIN(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DMIN(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DMIN(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DMIN(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #22: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DMIN(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DMIN(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DMIN(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DMIN(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DMIN(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #14, #16, #18.
		// Different logic converting type comparing with MS - Negative case #6.
		// Parser must be false - Negative case #19, #13.
	});

	QUnit.test('Test: "DPRODUCT"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 800
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 800, 'Test: Positive case: Area, String, Area. Return 800');
		// Case #2: Area, String, Area. Return 0
		oParser = new parserFormula('DPRODUCT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		ws.getRange2("G1").setValue("noname");
		// Case #3: Area, String, Area. Return 3628800
		oParser = new parserFormula('DPRODUCT(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3628800, 'Test: Positive case: Area, String, Area. Return 3628800');
		// Case #4: Area, String, Area. Return 75
		oParser = new parserFormula('DPRODUCT(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #5: Area, String, Area. Return 140
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, String, Area. Return 140');
		// Case #6: Area, String, Area. Return 3628800
		oParser = new parserFormula('DPRODUCT(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3628800, 'Test: Positive case: Area, String, Area. Return 3628800');
		// Case #7: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #8: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DPRODUCT(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #9: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 252, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #10: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #11: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 252, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #12: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #13: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #14: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #15: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #16: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #17: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #18: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DPRODUCT(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #19: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #20: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #21: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #22: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #23: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #24: Formula. DPRODUCT nested inside SUM formula.
		oParser = new parserFormula('SUM(DPRODUCT(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DPRODUCT(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 150, 'Test: Positive case: Formula. DPRODUCT nested inside SUM formula.');
		// Case #25: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #26: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 11200, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #27: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 140, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #28: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 604800, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #29: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #30: Area, String, Area. Database, field as "Tree" column (text values), criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range.');
		// Case #31: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.
		oParser = new parserFormula('DPRODUCT(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.');

		// Negative cases:

		// Case #1: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DPRODUCT(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #2: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #3: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #4: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #5: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #6: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #7: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #8: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #11: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DPRODUCT(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #12: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #13: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DPRODUCT(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DPRODUCT(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #14: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #15: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #16: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		/*oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #19: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DPRODUCT(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DPRODUCT(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #20: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #21: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #22: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DPRODUCT(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DPRODUCT(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 252, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DPRODUCT(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7875, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DPRODUCT(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DPRODUCT(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 604800, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #3, #4, #5, #14, #16, #18.
		// Different logic converting type comparing with MS - Negative case #6.
		// Parser must be false - Negative case #19, #13.
	});

	QUnit.test('Test: "DSTDEV"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("noname");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 1.1547
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '1.154700538', 'Test: Positive case: Area, String, Area. Return 1.1547');
		// Case #2: Area, String, Area. Return 4.381780460041329
		oParser = new parserFormula('DSTDEV(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '4.38178046', 'Test: Positive case: Area, String, Area. Return 4.381780460041329');
		// Case #3: Area, String, Area. Return 14
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, String, Area. Return 14');
		// Case #4: Area, String, Area. Return 20
		oParser = new parserFormula('DSTDEV(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '4.38178046', 'Test: Positive case: Area, String, Area. Return 20');
		// Case #5: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #6: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DSTDEV(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #7: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DSTDEV(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #8: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #9: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #10: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DSTDEV(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '3.535533906', 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #11: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #12: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #13: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #14: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #15: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '21.21320344', 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #16: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DSTDEV(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #17: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #18: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #19: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #20: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #21: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #22: Formula. DSTDEV nested inside SUM formula.
		oParser = new parserFormula('SUM(DSTDEV(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DSTDEV(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '12.82842712', 'Test: Positive case: Formula. DSTDEV nested inside SUM formula.');
		// Case #23: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #24: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.516611478', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #25: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.828427125', 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #26: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.664582519', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');

		// Negative cases:

		ws.getRange2("G1").setValue("Profit");
		// Case #1: Area, String, Area. Return #DIV/0!
		oParser = new parserFormula('DSTDEV(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');
		// Case #2: Area, String, Area. Return #DIV/0!
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');*/
		// Case #3: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DSTDEV(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #4: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #5: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #6: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #7: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #8: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #9: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #12: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #13: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DSTDEV(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #14: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #15: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSTDEV(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSTDEV(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #16: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #19: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #20: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEV(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #21: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSTDEV(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSTDEV(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #22: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEV(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #23: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #24: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEV(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #25: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DSTDEV(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #26: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DSTDEV(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #27: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEV(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DSTDEV(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2.8284271247461903, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DSTDEV(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 21.213203435596427, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DSTDEV(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEV(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2.6645825188948455, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #2, #4, #5, #6, #7, #16, #18, #20.
		// Different logic converting type comparing with MS - Negative case #8.
		// Parser must be false - Negative case #21, #15.

	});

	QUnit.test('Test: "DSTDEVP"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("noname");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 0.942809
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '0.942809042', 'Test: Positive case: Area, String, Area. Return 0.942809');
		// Case #2: Area, String, Area. Return 4.381780460041329
		oParser = new parserFormula('DSTDEVP(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Return 4.381780460041329');
		// Case #3: Area, String, Area. Return 0
		oParser = new parserFormula('DSTDEVP(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #4: Area, String, Area. Return 2
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Return 2');
		// Case #5: Area, String, Area. Return 4
		oParser = new parserFormula('DSTDEVP(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Return 4');
		// Case #6: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #7: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DSTDEVP(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #8: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #9: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #10: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #11: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '2.5', 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #12: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #13: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #14: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #15: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #16: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #17: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DSTDEVP(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #18: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #19: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #20: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #21: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #22: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #23: Formula. DSTDEVP nested inside SUM formula.
		oParser = new parserFormula('SUM(DSTDEVP(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DSTDEVP(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Formula. DSTDEVP nested inside SUM formula.');
		// Case #24: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #25: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '2.179449472', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #26: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #27: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '2.43241992', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');

		// Negative cases:

		ws.getRange2("G1").setValue("Profit");
		// Case #1: Area, String, Area. Return #DIV/0!
		oParser = new parserFormula('DSTDEVP(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');
		// Case #2: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DSTDEVP(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #3: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #4: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #5: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #6: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #7: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #8: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #12: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DSTDEVP(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #13: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #14: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSTDEVP(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSTDEVP(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #15: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #16: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #17: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #18: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #19: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #20: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSTDEVP(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSTDEVP(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #21: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #22: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #23: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #24: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DSTDEVP(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #25: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		oParser = new parserFormula('DSTDEVP(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #26: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DSTDEVP(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DSTDEVP(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DSTDEVP(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DSTDEVP(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSTDEVP(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2.4324199198877374, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #3, #4, #5, #6, #15, #17, #19.
		// Different logic converting type comparing with MS - Negative case #7.
		// Parser must be false - Negative case #20, #14.
	});

	QUnit.test('Test: "DSUM"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 225
		oParser = new parserFormula('DSUM(A4:E10,"Profit", A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10,"Profit", A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, String, Area. Return 225');
		// Case #2: Area, String, Area. Return 247.8
		oParser = new parserFormula('DSUM(A4:E10,"Profit", A1:F3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10,"Profit", A1:F3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '247.8', 'Test: Positive case: Area, String, Area. Return 247.8');
		// Case #3: Area, String, Area. Return 15
		oParser = new parserFormula('DSUM(A4:E10, "Age",A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age",A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Area, String, Area. Return 15');
		// Case #4: Area, String, Area. Return 0
		oParser = new parserFormula('DSUM(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 0');
		// Case #5: Area, String, Area. Return 78
		ws.getRange2("G1").setValue("noname");
		oParser = new parserFormula('DSUM(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 78, 'Test: Positive case: Area, String, Area. Return 78');
		// Case #6: Area, String, Area. Return 75
		oParser = new parserFormula('DSUM(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #7: Area, String, Area. Return 24
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, String, Area. Return 24');
		// Case #8: Area, String, Area. Return 78
		oParser = new parserFormula('DSUM(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 78, 'Test: Positive case: Area, String, Area. Return 78');
		// Case #9: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #10: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DSUM(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #11: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DSUM(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #12: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #13: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DSUM(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #14: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DSUM(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 35, 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #15: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DSUM(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #16: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DSUM(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #17: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DSUM(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #18: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DSUM(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #19: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DSUM(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #20: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DSUM(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #21: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DSUM(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #22: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #23: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #24: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #25: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #26: Formula. DSUM nested inside SUM formula.
		oParser = new parserFormula('SUM(DSUM(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DSUM(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 34, 'Test: Positive case: Formula. DSUM nested inside SUM formula.');
		// Case #27: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #28: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 42, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #29: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #30: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');
		// Case #31: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #32: Area, String, Area. Database, field as "Tree" column (text values), criteria range.
		oParser = new parserFormula('DSUM(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Database, field as "Tree" column (text values), criteria range.');
		// Case #33: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.
		oParser = new parserFormula('DSUM(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range.');

		// Negative cases:

		// Case #1: Area, String(2). Return #VALUE!
		oParser = new parserFormula('DSUM(A4:E10, "Age","test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age","test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String(2). Return #VALUE!');
		// Case #2: Area, String, Reference link. Return #VALUE!
		oParser = new parserFormula('DSUM(A4:E10, "Age",E2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Age",E2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Return #VALUE!');
		// Case #3: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DSUM(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #4: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');
		// Case #5: Area, String, Name. Database, field as string, criteria as named range.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');
		// Case #6: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		oParser = new parserFormula('DSUM(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');
		// Case #7: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');
		// Case #8: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DSUM(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #9: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #12: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #13: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DSUM(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #14: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #15: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSUM(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSUM(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #16: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');
		// Case #17: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');
		// Case #19: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #20: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');
		// Case #21: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DSUM(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DSUM(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #22: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DSUM(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #23: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error.');
		// Case #24: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DSUM(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DSUM(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DSUM(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 180, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DSUM(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DSUM(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different logic converting type comparing with MS - Negative case #8.
		// Parser must be false - Negative case #21, #15.
	});

	QUnit.test('Test: "DVAR"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("Profit");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 8.8
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A1:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A1:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '8.8', 'Test: Positive case: Area, String, Area. Return 8.8');
		// Case #2: Area, String, Area. Return 19.2
		ws.getRange2("G1").setValue("noname");
		oParser = new parserFormula('DVAR(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '19.2', 'Test: Positive case: Area, String, Area. Return 19.2');
		// Case #3: Area, String, Area. Return 8
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Return 8');
		// Case #4: Area, String, Area. Return 19.2
		oParser = new parserFormula('DVAR(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '19.2', 'Test: Positive case: Area, String, Area. Return 19.2');
		// Case #5: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #6: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DVAR(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #7: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DVAR(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #8: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DVAR(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #9: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DVAR(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #10: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DVAR(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '12.5', 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #11: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DVAR(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #12: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DVAR(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #13: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DVAR(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #14: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DVAR(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #15: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DVAR(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #16: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DVAR(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #17: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DVAR(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #18: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #19: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #20: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DVAR(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #21: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #22: Formula. DVAR nested inside SUM formula.
		oParser = new parserFormula('SUM(DVAR(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DVAR(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Formula. DVAR nested inside SUM formula.');
		// Case #23: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DVAR(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #24: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '6.333333333', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #25: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DVAR(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #26: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(1), '7.1', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');

		// Negative cases:

		// Case #1: Area, String, Area. Return #DIV/0!
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');*/
		// Case #2: Area, String, Area. Return #DIV/0!
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');*/
		// Case #3: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DVAR(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #4: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');*/
		// Case #5: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');*/
		// Case #6: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');*/
		// Case #7: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');*/
		// Case #8: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #9: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #12: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #13: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DVAR(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #14: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #15: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DVAR(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DVAR(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #16: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');*/
		// Case #17: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #18: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');*/
		// Case #19: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #20: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');*/
		// Case #21: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DVAR(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DVAR(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #22: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DVAR(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #23: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #24: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVAR(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #25: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DVAR(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #26: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		// Different result with MS
		oParser = new parserFormula('DVAR(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #27: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.
		// Different result with MS
		oParser = new parserFormula('DVAR(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.');*/

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DVAR(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DVAR(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 450, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DVAR(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVAR(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7.1, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #1, #2, #4, #5, #6, #7, #16, #18, #20, #25, #26, #27.
		// Different logic converting type comparing with MS - Negative case #8.
		// Parser must be false - Negative case #21, #15.
	});

	QUnit.test('Test: "DVARP"', function (assert) {
		putDataForDatabase();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Profit");
		ws.getRange2("A101").setValue("10");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("#NULL!");
		ws.getRange2("A105").setValue("text");

		ws.getRange2("G1").setValue("noname");
		ws.getRange2("G2").setValue("555");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("Yield"); // Column1
		ws.getRange2("A602").setValue("10"); // Column1
		ws.getRange2("B601").setValue("Profit"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("10");
		ws2.getRange2("A2").setValue("Profit");
		ws2.getRange2("A3").setValue("Yield");
		ws2.getRange2("A4").setValue("10");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("10"); // TestName
		ws.getRange2("A202").setValue("Profit"); // TestName1
		ws.getRange2("A206").setValue("Yield"); // TestNameArea
		ws.getRange2("A207").setValue("10"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("10") // TestName3D
		ws2.getRange2("A12").setValue("Profit") // TestName3D1
		ws2.getRange2("A16").setValue("Yield"); // TestNameArea3D
		ws2.getRange2("A17").setValue("10"); // TestNameArea3D

		// Positive cases:

		// Case #1: Area, String, Area. Return 7.04
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A1:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A1:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), '7.04', 'Test: Positive case: Area, String, Area. Return 7.04');
		// Case #2: Area, String, Area. Return 16
		oParser = new parserFormula('DVARP(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Area, String, Area. Return 16');
		// Case #3: Area, String, Area. Return 75
		oParser = new parserFormula('DVARP(A4:E10, "Profit", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Profit", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Area. Return 75');
		// Case #4: Area, String, Area. Return 14
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Return 14');
		// Case #5: Area, String, Area. Return 20
		oParser = new parserFormula('DVARP(A4:E10, 3, A4:E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 3, A4:E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Area, String, Area. Return 20');
		// Case #6: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Database, field as column name string, criteria range. Average yield for specified criteria.');
		// Case #7: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.
		oParser = new parserFormula('DVARP(A4:E10, 4, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 4, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, Number, Area. Database, field as column number (4 = Yield), criteria range. Average yield using column index.');
		// Case #8: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.
		oParser = new parserFormula('DVARP(A4:E10, "Height", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Height", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Area. Database, field as different column name, criteria range. Average height for criteria.');
		// Case #9: Area, String, Area. Database, field as "Profit" column, criteria range.
		oParser = new parserFormula('DVARP(A4:E10, "Profit", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Profit", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, String, Area. Database, field as "Profit" column, criteria range.');
		// Case #10: Area, Number, Area. Database, field as column 2 (Height), criteria range.
		oParser = new parserFormula('DVARP(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, Number, Area. Database, field as column 2 (Height), criteria range.');
		// Case #11: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.
		oParser = new parserFormula('DVARP(A4:E10, "Age", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Age", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), '6.25', 'Test: Positive case: Area, String, Area. Database, field as "Age" column, criteria range. Calculate average age.');
		// Case #12: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.
		oParser = new parserFormula('DVARP(A4:E10, A100, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, A100, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, Reference link, Area. Database, field as reference link to "Yield", criteria range.');
		// Case #13: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.
		oParser = new parserFormula('DVARP(A4:E10, Sheet2!A2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, Sheet2!A2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, Ref3D, Area. Database, field as ref3D to "Yield", criteria range.');
		// Case #14: Area, Name, Area. Database, field as Name to "Yield", criteria range.
		oParser = new parserFormula('DVARP(A4:E10, TestName1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, TestName1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, Name, Area. Database, field as Name to "Yield", criteria range.');
		// Case #15: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.
		oParser = new parserFormula('DVARP(A4:E10, TestName3D1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, TestName3D1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, Name3D, Area. Database, field as Name3D to "Yield", criteria range.');
		// Case #16: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.
		oParser = new parserFormula('DVARP(A4:E10, Table1[Column2], A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, Table1[Column2], A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Positive case: Area, Table, Area. Database, field as Table to column with "Yield", criteria range.');
		// Case #17: Area, Number, Area3D. Database, field as number, criteria as 3D area range.
		oParser = new parserFormula('DVARP(A4:E10, 4, Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 4, Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, Number, Area3D. Database, field as number, criteria as 3D area range.');
		// Case #18: Area, String, Table. Database, field as string, criteria as table reference.
		getTableType(599, 0, 601, 1);
		oParser = new parserFormula('DVARP(A4:E10, "Yield", Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Table. Database, field as string, criteria as table reference.');
		// Case #19: Area, String, Name. Database, field as string, criteria as name area reference.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name. Database, field as string, criteria as name area reference.');
		// Case #20: Area, String, Name3D. Database, field as string, criteria as name3D area reference.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area, String, Name3D. Database, field as string, criteria as name3D area reference.');
		// Case #21: Area, Formula, Area. Database, field from nested IF formula, criteria range.
		oParser = new parserFormula('DVARP(A4:E10, IF(TRUE,"Yield","Age"), A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, IF(TRUE,"Yield","Age"), A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, Formula, Area. Database, field from nested IF formula, criteria range.');
		// Case #22: Area, String, Formula. Database, field as string, criteria from nested IF formula.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", IF(1=1,A1:B2,A1:C2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, String, Formula. Database, field as string, criteria from nested IF formula.');
		// Case #23: Formula. DVARP nested inside SUM formula.
		oParser = new parserFormula('SUM(DVARP(A4:E10,"Yield",A1:B2),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DVARP(A4:E10,"Yield",A1:B2),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Formula. DVARP nested inside SUM formula.');
		// Case #24: Area, Array, Area. Database, field as single-element array, criteria range.
		oParser = new parserFormula('DVARP(A4:E10, {"Yield"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, {"Yield"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, Array, Area. Database, field as single-element array, criteria range.');
		// Case #25: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A1:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A1:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), '4.75', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with blank row below headers. Testing blank line behavior.');
		// Case #26: Area, Array, Area. Database, field as multi-element array, criteria range.
		oParser = new parserFormula('DVARP(A4:E10, {"Yield","Age"}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, {"Yield","Age"}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Area, Array, Area. Database, field as multi-element array, criteria range.');
		// Case #27: Area, String, Area. Database, field as string, criteria range with empty cells
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9), '5.916666667', 'Test: Positive case: Area, String, Area. Database, field as string, criteria range with empty cells');

		// Negative cases:

		ws.getRange2("G1").setValue("Profit");
		// Case #1: Area, String, Area. Return #DIV/0!
		oParser = new parserFormula('DVARP(A4:E10, "Age",G1:G2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Age",G1:G2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Return #DIV/0!');
		// Case #2: Area, String, Area. Empty string. Return #VALUE!
		oParser = new parserFormula('DVARP(A4:E10,"", A1:F2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10,"", A1:F2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Empty string. Return #VALUE!');
		// Case #3: Area, String, Reference link. Database, field as string, criteria as single-cell reference.
		/*// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as single-cell reference.');
		// Case #4: Area, String, Name. Database, field as string, criteria as named range.
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, "Yield", TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Name. Database, field as string, criteria as named range.');
		// Case #5: Area, Number, Name3D. Database, field as number, criteria as 3D named range.
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, 4, TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 4, TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Name3D. Database, field as number, criteria as 3D named range.');
		// Case #6: Area, String, Ref3D. Database, field as string, criteria as 3D reference.
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, "Yield", Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Ref3D. Database, field as string, criteria as 3D reference.');
		// Case #7: Area, String, Area. Database, field as numeric string "4", criteria range.
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, "4", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "4", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as numeric string "4", criteria range.');*/
		// Case #8: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, "InvalidColumn", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "InvalidColumn", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as non-existent column name, criteria range. Return #VALUE! error.');
		// Case #9: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, 0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 0 (invalid column index), criteria range. Return #VALUE! error.');
		// Case #10: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, 10, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 10, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as 10 (exceeds column count), criteria range. Return #VALUE! error.');
		// Case #11: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, -1, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, -1, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, Area. Database, field as negative number, criteria range. Return #VALUE! error.');
		// Case #12: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.
		oParser = new parserFormula('DVARP(A4:E10, #N/A, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, #N/A, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, Error, Area. Database, field as error value, criteria range. Return #N/A error.');
		// Case #13: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area, String, Error. Database, field as string, criteria as error value. Return #N/A error.');
		// Case #14: Error, String, Area. Database as error value, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DVARP(#N/A, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(#N/A, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DVARP(#N/A; "Yield"; S1:T2)', 'Test: Negative case: Error, String, Area. Database as error value, field as string, criteria range.');
		*/// Case #15: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVARP(A4:E10, "Yield", A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Area. Database, field as string, criteria as reference to error cell. Return #VALUE! error.');
		*/// Case #16: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, A104, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, A104, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to error cell, criteria range. Return #VALUE! error.');
		// Case #17: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVARP(A4:E10, "Yield", A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Reference link. Database, field as string, criteria as reference to text cell. Return #VALUE! error.');
		*/// Case #18: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, A105, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, A105, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Reference link, Area. Database, field as reference to text cell, criteria range. Return #VALUE! error.');
		// Case #19: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.
		// Different result with MS
		/*oParser = new parserFormula('DVARP(A4:E10, "Yield", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String, Empty. Database, field as string, criteria as empty. Return #VALUE! error.');
		*/// Case #20: Empty, String, Area. Database as empty, field as string, criteria range.
		// Different result with MS TODO Parser must be false need to fix!
		/*oParser = new parserFormula('DVARP(, "Yield", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(, "Yield", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '=DVARP(; "Yield"; S1:T2)', 'Test: Negative case: Empty, String, Area. Database as empty, field as string, criteria range.');
		*/// Case #21: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.
		oParser = new parserFormula('DVARP(A4:E10, 1/0, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 1/0, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Formula, Area. Database, field as formula resulting in #DIV/0! error, criteria range. Return #DIV/0! error.');
		// Case #22: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", 1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", 1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Formula. Database, field as string, criteria as formula resulting in #DIV/0! error. Return #DIV/0! error.');
		// Case #23: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.
		oParser = new parserFormula('DVARP(A4:E10, , A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, , A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Empty, Area. Database, field as empty, criteria range. Return #VALUE! error.');
		// Case #24: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.
		// Different result with MS
		/*oParser = new parserFormula('DVARP(A4:E10, TRUE, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, TRUE, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Boolean, Area. Database, field as boolean TRUE, criteria range.');
		// Case #25: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, "Tree", A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Tree", A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, String, Area. Database, field as "Tree" column (text values), criteria range. Return #DIV/0! error (no numeric values to average).');
		// Case #26: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.
		// Different result with MS
		oParser = new parserFormula('DVARP(A4:E10, {1,2}, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, {1,2}, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area, Array, Area. Database, field as multi-element array with numbers, criteria range. Return #DIV/0! error.');
		*/

		// Bounded cases:

		// Case #1: Area, Number, Area. Minimum accepted column index (2 = Height column).
		oParser = new parserFormula('DVARP(A4:E10, 2, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 2, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Bounded case: Area, Number, Area. Minimum accepted column index (2 = Height column).');
		// Case #2: Area, Number, Area. Maximum column index (5 = Profit column).
		oParser = new parserFormula('DVARP(A4:E10, 5, A1:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, 5, A1:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 225, 'Test: Bounded case: Area, Number, Area. Maximum column index (5 = Profit column).');
		// Case #3: Area, String, Area. Criteria range as entire columns. Testing whole column reference.
		oParser = new parserFormula('DVARP(A4:E10, "Yield", A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DVARP(A4:E10, "Yield", A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5.916666666666667, 'Test: Bounded case: Area, String, Area. Criteria range as entire columns. Testing whole column reference.');

		// TODO Need to fix

		// Different Error text comparing with MS - Negative case #3, #4, #5, #6, #15, #17, #19, #24, #25, #26.
		// Different logic converting type comparing with MS - Negative case #7.
		// Parser must be false - Negative case #20, #14.
	});

	wb.dependencyFormulas.unlockRecal();
});
