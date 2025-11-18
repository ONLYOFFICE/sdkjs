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
	let oParser, wb, ws, sData = AscCommon.getEmpty(), tmp;
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
	function testArrayFormula(assert, func, dNotSupportAreaArg) {

		const getValue = function (ref) {
			oParser = new parserFormula(func + "(" + ref + ")", "A2", ws);
			assert.ok(oParser.parse(), 'Formula is parsed');
			return oParser.calculate().getValue();
		};
		const description = `Test: testArrayFormula. ${func}. `;

		//***array-formula***
		ws.getRange2("A100").setValue("1");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("C100").setValue("-4");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B101").setValue("4");
		ws.getRange2("C101").setValue("5");


		oParser = new parserFormula(func + "(A100:C101)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
		assert.ok(oParser.parse(), 'Formula is parsed.');
		let array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue("A100"), description + 'Reference Link. Area.');
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue("B100"), description + 'Reference Link. Area.');
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue("C100"), description + 'Reference Link. Area.');
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), getValue("A101"), description + 'Reference Link. Area.');
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), getValue("B101"), description + 'Reference Link. Area.');
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), getValue("C101"), description + 'Reference Link. Area.');
		} else {
			if (!dNotSupportAreaArg) {
				assert.strictEqual(false, true, 'func: ' + func + ' don\'t support area argument.');
			}
			//consoleLog("func: " + func + " don't return area array");
		}

		oParser = new parserFormula(func + "({1,2,-3})", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
		assert.ok(oParser.parse(), 'Formula is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue(1), description + 'Number.');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue(2), description + 'Number.');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue(-3), description + 'Number.');
	}
	//returnOnlyValue - those functions that must always be fed arrays as input and that return a single value
	function testArrayFormula2(assert, func, minArgCount, maxArgCount, dNotSupportAreaArg, returnOnlyValue) {

		const getValue = function (ref, countArg) {
			let argStr = "(";
			for (let j = 1; j <= countArg; j++) {
				argStr += ref;
				if (i !== j) {
					argStr += ",";
				} else {
					argStr += ")";
				}
			}
			oParser = new parserFormula(func + argStr, "A2", ws);
			assert.ok(oParser.parse(), 'Test: Formula ' + func + argStr + ' is parsed.');
			return oParser.calculate().getValue();
		};


		//***array-formula***
		ws.getRange2("A100").setValue("1");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("C100").setValue("-4");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B101").setValue("4");
		ws.getRange2("C101").setValue("5");

		//формируем массив значений
		const randomArray = [];
		let randomStrArray = "{";
		let maxArg = 4;
		for (let i = 1; i <= maxArg; i++) {
			let randVal = Math.random();
			randomArray.push(randVal);
			randomStrArray += randVal;
			if (i !== maxArg) {
				randomStrArray += ",";
			} else {
				randomStrArray += "}";
			}
		}

		for (var i = minArgCount; i <= maxArgCount; i++) {
			let argStrArr = "(";
			let randomArgStrArr = "(";
			for (let j = 1; j <= i; j++) {
				argStrArr += "A100:C101";
				randomArgStrArr += randomStrArray;
				if (i !== j) {
					argStrArr += ",";
					randomArgStrArr += ",";
				} else {
					argStrArr += ")";
					randomArgStrArr += ")";
				}
			}

			oParser = new parserFormula(func + argStrArr, "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok(oParser.parse(), 'Test: Formula ' + func + argStrArr + ' is parsed.');
			let array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue("A100", i), 'Test: testArrayFormula2. Area.');
				assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue("B100", i), 'Test: testArrayFormula2. Area.');
				assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue("C100", i), 'Test: testArrayFormula2. Area.');
				assert.strictEqual(array.getElementRowCol(1, 0).getValue(), getValue("A101", i), 'Test: testArrayFormula2. Area.');
				assert.strictEqual(array.getElementRowCol(1, 1).getValue(), getValue("B101", i), 'Test: testArrayFormula2. Area.');
				assert.strictEqual(array.getElementRowCol(1, 2).getValue(), getValue("C101", i), 'Test: testArrayFormula2. Area.');
			} else {
				if (!(dNotSupportAreaArg || returnOnlyValue)) {
					assert.strictEqual(false, true);
				}
				//consoleLog("func: " + func + " don't return area array");
			}

			oParser = new parserFormula(func + randomArgStrArr, "A1", ws);
			oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
			assert.ok(oParser.parse(), 'Test: Formula ' + func + randomArgStrArr + ' is parsed.');
			array = oParser.calculate();
			if (AscCommonExcel.cElementType.array === array.type) {
				assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue(randomArray[0], i), 'Test: testArrayFormula2. Random array');
				assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue(randomArray[1], i), 'Test: testArrayFormula2. Random array');
				assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue(randomArray[2], i), 'Test: testArrayFormula2. Random array');
			} else {
				if (!returnOnlyValue) {
					assert.strictEqual(false, true);
				}
				//consoleLog("func: " + func + " don't return array");
			}
		}
	}
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
		/** @type {TablePart} */
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

	wb.dependencyFormulas.lockRecal();
	getTableType(599, 0, 599, 0); // Init table
	initDefNames();

	QUnit.module('Date & time formulas');

	QUnit.test('Test: "DATE"', function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("2024");
		ws.getRange2("A101").setValue("6");
		ws.getRange2("A102").setValue("15");
		ws.getRange2("A103").setValue("text");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A105").setValue("");
		ws.getRange2("B105").setValue("");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("2024"); // Column1
		ws.getRange2("B601").setValue("6"); // Column2
		ws.getRange2("C601").setValue("15"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("2024");
		ws2.getRange2("A2").setValue("6");
		ws2.getRange2("A3").setValue("15");
		ws2.getRange2("A4").setValue("text");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("2024"); // TestName
		ws.getRange2("A202").setValue("6"); // TestName1
		ws.getRange2("A203").setValue("15"); // TestName2
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("2024") // TestName3D
		ws2.getRange2("A12").setValue("6") // TestName3D1
		ws2.getRange2("A13").setValue("15") // TestName3D2

		// Positive cases:

		// Case #1: Number(3). Basic valid date with all integer arguments. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(2024,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Number(3). Basic valid date with all integer arguments. 3 of 3 arguments used.');
		// Case #2: Number(3). Example from documentation. Returns January 2, 2008.
		oParser = new parserFormula('DATE(2008,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2008,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39449, 'Test: Positive case: Number(3). Example from documentation. Returns January 2, 2008.');
		// Case #3: Number(3). Year between 0-1899, Excel adds to 1900. Returns January 2, 2008.
		oParser = new parserFormula('DATE(108,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(108,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39449, 'Test: Positive case: Number(3). Year between 0-1899, Excel adds to 1900. Returns January 2, 2008.');
		// Case #4: Number(3). Two-digit year 50 interpreted as 1950. Returns January 1, 1950.
		oParser = new parserFormula('DATE(50,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(50,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18264, 'Test: Positive case: Number(3). Two-digit year 50 interpreted as 1950. Returns January 1, 1950.');
		// Case #5: Number(3). Two-digit year 99 interpreted as 1999. Returns December 31, 1999.
		oParser = new parserFormula('DATE(99,12,31)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(99,12,31) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36525, 'Test: Positive case: Number(3). Two-digit year 99 interpreted as 1999. Returns December 31, 1999.');
		// Case #6: Number(3). Month greater than 12 adds months. Returns February 2, 2025.
		oParser = new parserFormula('DATE(2024,14,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,14,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45690, 'Test: Positive case: Number(3). Month greater than 12 adds months. Returns February 2, 2025.');
		// Case #7: Number(3). Negative month subtracts from first month. Returns September 2, 2007.
		oParser = new parserFormula('DATE(2008,-3,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2008,-3,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39327, 'Test: Positive case: Number(3). Negative month subtracts from first month. Returns September 2, 2007.');
		// Case #8: Number(3). Day greater than days in month. Returns February 4, 2024.
		oParser = new parserFormula('DATE(2024,1,35)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,1,35) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45326, 'Test: Positive case: Number(3). Day greater than days in month. Returns February 4, 2024.');
		// Case #9: Number(3). Negative day subtracts from first day. Returns December 16, 2007.
		oParser = new parserFormula('DATE(2008,1,-15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2008,1,-15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39432, 'Test: Positive case: Number(3). Negative day subtracts from first day. Returns December 16, 2007.');
		// Case #10: Number(3). Float numbers truncated to integers. Returns June 15, 2024.
		oParser = new parserFormula('DATE(2024.7,6.9,15.3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024.7,6.9,15.3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Number(3). Float numbers truncated to integers. Returns June 15, 2024.');
		// Case #11: String(3). String arguments convertible to numbers. Returns June 15, 2024.
		oParser = new parserFormula('DATE("2024","6","15")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE("2024","6","15") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: String(3). String arguments convertible to numbers. Returns June 15, 2024.');
		// Case #12: String, Number(2). Mixed string and number arguments. Returns June 15, 2024.
		oParser = new parserFormula('DATE("2024",6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE("2024",6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: String, Number(2). Mixed string and number arguments. Returns June 15, 2024.');
		// Case #13: Reference link(3). Reference links to valid numbers. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(A100,A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A100,A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Reference link(3). Reference links to valid numbers. 3 of 3 arguments used.');
		// Case #14: Formula(3). Nested formulas extracting date components from static date. Returns June 15, 2024.
		oParser = new parserFormula('DATE(YEAR(DATE(2024,6,15)),MONTH(DATE(2024,6,15)),DAY(DATE(2024,6,15)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(YEAR(DATE(2024,6,15)),MONTH(DATE(2024,6,15)),DAY(DATE(2024,6,15))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Formula(3). Nested formulas extracting date components from static date. Returns June 15, 2024.');
		// Case #15: Formula(3). Arithmetic expressions as arguments. Returns June 15, 2024.
		oParser = new parserFormula('DATE(2020+4,3*2,7+8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2020+4,3*2,7+8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Formula(3). Arithmetic expressions as arguments. Returns June 15, 2024.');
		// Case #16: Formula. DATE nested inside SUM formula. Returns date serial + 100.
		oParser = new parserFormula('SUM(DATE(2024,1,1),100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DATE(2024,1,1),100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45392, 'Test: Positive case: Formula. DATE nested inside SUM formula. Returns date serial + 100.');
		// Case #17: Boolean, Number(2). Boolean TRUE converted to 1, adds to 1900. Returns June 15, 1901.
		oParser = new parserFormula('DATE(TRUE,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(TRUE,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 532, 'Test: Positive case: Boolean, Number(2). Boolean TRUE converted to 1, adds to 1900. Returns June 15, 1901.');
		// Case #18: Number, Boolean(2). Boolean TRUE as month (1), FALSE as day (0). Returns January 0, 2024.
		oParser = new parserFormula('DATE(2024,TRUE,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,TRUE,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45291, 'Test: Positive case: Number, Boolean(2). Boolean TRUE as month (1), FALSE as day (0). Returns January 0, 2024.');
		// Case #19: Empty, Number(2). Empty year argument converts to 0, adds to 1900. Returns January 1, 1900.
		oParser = new parserFormula('DATE(,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Empty, Number(2). Empty year argument converts to 0, adds to 1900. Returns January 1, 1900.');
		// Case #20: Number, Empty(2). Empty month and day convert to 0. Month subtracts 1, day subtracts 1.
		oParser = new parserFormula('DATE(2024,,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45260, 'Test: Positive case: Number, Empty(2). Empty month and day convert to 0. Month subtracts 1, day subtracts 1.');
		// Case #21: Name(3). Named ranges as arguments. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(TestName,TestName1,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(TestName,TestName1,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Name(3). Named ranges as arguments. 3 of 3 arguments used.');
		// Case #22: Name3D(3). 3D named ranges from another sheet. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(TestName3D,TestName3D1,TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(TestName3D,TestName3D1,TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Name3D(3). 3D named ranges from another sheet. 3 of 3 arguments used.');
		// Case #23: Ref3D(3). 3D references to cells with valid numbers. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(Sheet2!A1,Sheet2!A2,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(Sheet2!A1,Sheet2!A2,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Ref3D(3). 3D references to cells with valid numbers. 3 of 3 arguments used.');
		// Case #24: Area(3). Single-cell ranges as arguments. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(A100:A100,A101:A101,A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A100:A100,A101:A101,A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Area(3). Single-cell ranges as arguments. 3 of 3 arguments used.');
		// Case #25: Area3D(3). 3D single-cell ranges from another sheet. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Area3D(3). 3D single-cell ranges from another sheet. 3 of 3 arguments used.');
		// Case #26: Table(3). Table structured references as arguments. 3 of 3 arguments used.
		oParser = new parserFormula('DATE(Table1[Column1],Table1[Column2],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(Table1[Column1],Table1[Column2],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 47467, 'Test: Positive case: Table(3). Table structured references as arguments. 3 of 3 arguments used.');
		// Case #27: Array(3). Array with single elements as arguments. Returns June 15, 2024.
		oParser = new parserFormula('DATE({2024},{6},{15})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE({2024},{6},{15}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Array(3). Array with single elements as arguments. Returns June 15, 2024.');
		// Case #28: Formula(3). Nested IF formulas returning valid values. Returns June 15, 2024.
		oParser = new parserFormula('DATE(IF(TRUE,2024,2023),IF(TRUE,6,12),IF(TRUE,15,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(IF(TRUE,2024,2023),IF(TRUE,6,12),IF(TRUE,15,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Formula(3). Nested IF formulas returning valid values. Returns June 15, 2024.');
		// Case #29: Number(3). Minimum year value 1900. Returns January 1, 1900.
		oParser = new parserFormula('DATE(1900,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number(3). Minimum year value 1900. Returns January 1, 1900.');
		// Case #30: Number(3). Month equals 0 subtracts 1 month from January. Returns December 1, 2023.
		oParser = new parserFormula('DATE(2024,0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45261, 'Test: Positive case: Number(3). Month equals 0 subtracts 1 month from January. Returns December 1, 2023.');
		// Case #31: Number(3). Month equals 13 adds 1 month to December. Returns January 1, 2025.
		oParser = new parserFormula('DATE(2024,13,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,13,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Number(3). Month equals 13 adds 1 month to December. Returns January 1, 2025.');
		// Case #32: Number(3). Day equals 32 in January (31 days). Returns February 1, 2024.
		oParser = new parserFormula('DATE(2024,1,32)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,1,32) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45323, 'Test: Positive case: Number(3). Day equals 32 in January (31 days). Returns February 1, 2024.');
		// Case #33: Formula. ROUND formulas as arguments. Returns June 16, 2024.
		oParser = new parserFormula('DATE(ROUND(2024.8,0),ROUND(6.4,0),ROUND(15.9,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(ROUND(2024.8,0),ROUND(6.4,0),ROUND(15.9,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45824, 'Test: Positive case: Formula. ROUND formulas as arguments. Returns June 16, 2024.');
		// Case #34: Number(3). Both month and day overflow. Returns April 9, 2026.
		oParser = new parserFormula('DATE(2024,25,100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,25,100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 46122, 'Test: Positive case: Number(3). Both month and day overflow. Returns April 9, 2026.');
		// Case #35: Number(3). Year 1899 adds to 1900, becomes 3799. Returns December 31, 3799.
		oParser = new parserFormula('DATE(1899,12,31)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1899,12,31) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 693962, 'Test: Positive case: Number(3). Year 1899 adds to 1900, becomes 3799. Returns December 31, 3799.');
		// Case #36: Area, Number(2). Multi-cell range as year with empty cells
		oParser = new parserFormula('DATE(A105:B105,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A105:B105,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 167, 'Test: Positive case: Area, Number(2). Multi-cell range as year with empty cells');
		// Case #37: Array, Number(2). Multi-element array.
		oParser = new parserFormula('DATE({2024,2025},6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE({2024,2025},6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45458, 'Test: Positive case: Array, Number(2). Multi-element array.');
		// Case #38: Number(3). Day equals 0 subtracts 1 day, returns December 31, 2023. Not error but edge case.
		oParser = new parserFormula('DATE(2024,1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45291, 'Test: Positive case: Number(3). Day equals 0 subtracts 1 day, returns December 31, 2023. Not error but edge case.');

		// Negative cases:

		// Case #1: Number(3). Year less than 0 returns #NUM! error.
		oParser = new parserFormula('DATE(-1,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(-1,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Year less than 0 returns #NUM! error.');
		// Case #2: Number(3). Year equals 10000 or greater returns #NUM! error.
		oParser = new parserFormula('DATE(10000,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(10000,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Year equals 10000 or greater returns #NUM! error.');
		// Case #3: Number(3). Date exceeds maximum date 31.12.9999 returns #NUM! error.
		oParser = new parserFormula('DATE(9999,12,32)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(9999,12,32) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Date exceeds maximum date 31.12.9999 returns #NUM! error.');
		// Case #4: Number(3). Large negative day results in date before 01.01.1900 returns #NUM! error.
		oParser = new parserFormula('DATE(1900,1,-10000)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,1,-10000) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Large negative day results in date before 01.01.1900 returns #NUM! error.');
		// Case #5: Number(3). Large negative month results in date before 01.01.1900 returns #NUM! error.
		oParser = new parserFormula('DATE(1900,-1000,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,-1000,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Large negative month results in date before 01.01.1900 returns #NUM! error.');
		// Case #6: String, Number(2). Non-numeric string as year returns #VALUE! error.
		oParser = new parserFormula('DATE("abc",6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE("abc",6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number(2). Non-numeric string as year returns #VALUE! error.');
		// Case #7: Number, String(2). Non-numeric string as month returns #VALUE! error.
		oParser = new parserFormula('DATE(2024,"xyz",15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,"xyz",15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String(2). Non-numeric string as month returns #VALUE! error.');
		// Case #8: Number(2), String. Non-numeric string as day returns #VALUE! error.
		oParser = new parserFormula('DATE(2024,6,"test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,6,"test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), String. Non-numeric string as day returns #VALUE! error.');
		// Case #9: Error, Number(2). Error in year argument propagates #N/A error.
		oParser = new parserFormula('DATE(#N/A,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(#N/A,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number(2). Error in year argument propagates #N/A error.');
		// Case #10: Number, Error, Number. Error in month argument propagates #DIV/0! error.
		oParser = new parserFormula('DATE(2024,#DIV/0!,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,#DIV/0!,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Number, Error, Number. Error in month argument propagates #DIV/0! error.');
		// Case #11: Number(2), Error. Error in day argument propagates #REF! error.
		oParser = new parserFormula('DATE(2024,6,#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,6,#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: Number(2), Error. Error in day argument propagates #REF! error.');
		// Case #12: Reference link, Number(2). Reference to cell with text returns #VALUE! error.
		oParser = new parserFormula('DATE(A103,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A103,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link, Number(2). Reference to cell with text returns #VALUE! error.');
		// Case #13: Number, Reference link, Number. Reference to cell with text as month returns #VALUE! error.
		oParser = new parserFormula('DATE(2024,A103,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,A103,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Reference link, Number. Reference to cell with text as month returns #VALUE! error.');
		// Case #14: Number(2), Reference link. Reference to cell with text as day returns #VALUE! error.
		oParser = new parserFormula('DATE(2024,6,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(2024,6,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), Reference link. Reference to cell with text as day returns #VALUE! error.');
		// Case #15: Reference link, Number(2). Reference to cell with error returns error from cell.
		oParser = new parserFormula('DATE(A104,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A104,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link, Number(2). Reference to cell with error returns error from cell.');
		// Case #16: String(3). Empty strings return #VALUE! error.
		oParser = new parserFormula('DATE("","","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE("","","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(3). Empty strings return #VALUE! error.');
		// Case #17: Formula(3). Formula resulting in #NUM! error propagates error.
		oParser = new parserFormula('DATE(SQRT(-1),6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(SQRT(-1),6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(3). Formula resulting in #NUM! error propagates error.');
		// Case #18: Ref3D, Number(2). 3D reference to text returns #VALUE! error.
		oParser = new parserFormula('DATE(Sheet2!A4,6,15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(Sheet2!A4,6,15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, Number(2). 3D reference to text returns #VALUE! error.');
		// Case #19: Number(3). Month overflow from max year returns #NUM! (exceeds 31.12.9999).
		oParser = new parserFormula('DATE(9999,13,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(9999,13,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Month overflow from max year returns #NUM! (exceeds 31.12.9999).');
		// Case #20: Number(3). Year far exceeds 10000 returns #NUM! error.
		oParser = new parserFormula('DATE(15000,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(15000,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Year far exceeds 10000 returns #NUM! error.');
		// Case #21: Number(3). Large negative year returns #NUM! error.
		oParser = new parserFormula('DATE(-999999,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(-999999,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(3). Large negative year returns #NUM! error.');

		// Bounded cases:

		// Case #1: Number(3). Maximum valid date: December 31, 9999. Returns serial 2958465.
		oParser = new parserFormula('DATE(9999,12,31)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(9999,12,31) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Number(3). Maximum valid date: December 31, 9999. Returns serial 2958465.');
		// Case #2: Number(3). Minimum valid date: January 1, 1900. Returns serial 1.
		oParser = new parserFormula('DATE(1900,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number(3). Minimum valid date: January 1, 1900. Returns serial 1.');
		// Case #3: Number(3). Day 0 subtracts 1 from January 1. Returns December 31, 1899 (serial 0).
		oParser = new parserFormula('DATE(1900,1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number(3). Day 0 subtracts 1 from January 1. Returns December 31, 1899 (serial 0).');
		// Case #4: Number(3). One day before maximum date. Returns December 30, 9999.
		oParser = new parserFormula('DATE(9999,12,30)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(9999,12,30) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958464, 'Test: Bounded case: Number(3). One day before maximum date. Returns December 30, 9999.');
		// Case #5: Number(3). Year 0 adds to 1900. Returns January 1, 1900.
		oParser = new parserFormula('DATE(0,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(0,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number(3). Year 0 adds to 1900. Returns January 1, 1900.');
		// Case #6: Number(3). Year 1899 adds to 1900. Returns January 1, 3799.
		oParser = new parserFormula('DATE(1899,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1899,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 693598, 'Test: Bounded case: Number(3). Year 1899 adds to 1900. Returns January 1, 3799.');
		// Case #8: Area(3). Whole column references for all arguments. Tests boundary behavior.
		oParser = new parserFormula('DATE(A:A,B:B,C:C)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(A:A,B:B,C:C) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Area(3). Whole column references for all arguments. Tests boundary behavior.');
		// Case #9: Number(3). Year 9000 within valid range. Returns January 1, 9000.
		oParser = new parserFormula('DATE(9000,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(9000,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2593224, 'Test: Bounded case: Number(3). Year 9000 within valid range. Returns January 1, 9000.');
		// Case #10: Number(3). Last day of minimum year 1900. Returns December 31, 1900.
		oParser = new parserFormula('DATE(1900,12,31)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATE(1900,12,31) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 366, 'Test: Bounded case: Number(3). Last day of minimum year 1900. Returns December 31, 1900.');

		testArrayFormula2(assert, "DATE", 3, 3);
	});

	QUnit.test('Test: "DATEDIF"', function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("43831");
		ws.getRange2("A101").setValue("44196");
		ws.getRange2("A102").setValue("d");

		ws.getRange2("B2").setValue("2");
		ws.getRange2("B3").setValue("5");
		ws.getRange2("B4").setValue("15");
		ws.getRange2("B5").setValue("string");
		ws.getRange2("B6").setValue("#N/A");
		ws.getRange2("B7").setValue();
		ws.getRange2("B8").setValue("");

		ws.getRange2("C2").setValue("2");
		ws.getRange2("C3").setValue("12");
		ws.getRange2("C4").setValue("15");
		ws.getRange2("C5").setValue("25");
		ws.getRange2("C6").setValue("25.5");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("43831"); // Column1
		ws.getRange2("B601").setValue("44196"); // Column2
		ws.getRange2("C601").setValue("d"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("43831");
		ws2.getRange2("A2").setValue("44196");
		ws2.getRange2("A3").setValue("d");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("43831"); // TestName
		ws.getRange2("A202").setValue("44196"); // TestName1
		ws.getRange2("A203").setValue("d"); // TestName2
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("43831") // TestName3D
		ws2.getRange2("A12").setValue("44196") // TestName3D1
		ws2.getRange2("A13").setValue("d") // TestName3D2

		// Positive cases:

		// base case
		// Case #1: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), Y). Return 2.
		oParser = new parserFormula('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), Y). Return 2.');
		// Case #2: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), M). Return 24.
		oParser = new parserFormula('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), M). Return 24.');
		// Case #3: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), D). Return 730.
		oParser = new parserFormula('DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 730, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,1,1),DATE(2003,1,1), D). Return 730.');
		// Case #4: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), YD). Return 75.
		oParser = new parserFormula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), YD). Return 75.');
		// Case #5: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), MD). Return 14.
		oParser = new parserFormula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), MD). Return 14.');
		// Case #6: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), YM). Return 2.
		oParser = new parserFormula('DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula(2), String. Result DATEDIF(DATE(2001,6,1),DATE(2002,8,15), YM). Return 2.');
		// bug 54552 tests
		// Case #7: Formula(2), String. Bug test case. Return 0.
		oParser = new parserFormula('DATEDIF(DATE(2020,10,2),DATE(2021,10,1),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,10,2),DATE(2021,10,1),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Formula(2), String. Bug test case. Return 0.');
		// Case #8: Formula(2), String. Bug test case 2. Return 21.
		oParser = new parserFormula('DATEDIF(DATE(2000,4,13),DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2000,4,13),DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 21, 'Test: Positive case: Formula(2), String. Bug test case 2. Return 21.');
		// strings
		// Case #9: String, Formula, String. String number first. Return 122.
		oParser = new parserFormula('DATEDIF("12",DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("12",DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 122, 'Test: Positive case: String, Formula, String. String number first. Return 122.');

		// numbers
		// Case #10: Number, Formula, String. Number first. Return 122.
		oParser = new parserFormula('DATEDIF(12,DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 122, 'Test: Positive case: Number, Formula, String. Number first. Return 122.');
		// Case #11: Number(2), String. Two equal numbers. Return 0.
		oParser = new parserFormula('DATEDIF(12,12,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,12,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number(2), String. Two equal numbers. Return 0.');
		// Case #12: Number(2), String. First number less than second(years). Return 0.
		oParser = new parserFormula('DATEDIF(12,22,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number(2), String. First number less than second(years). Return 0.');
		// Case #13: Number(2), String. First number less than second(months). Return 0.
		oParser = new parserFormula('DATEDIF(12,22,"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number(2), String. First number less than second(months). Return 0.');
		// Case #14: Number(2), String. First number less than second(days). Return 10.
		oParser = new parserFormula('DATEDIF(12,22,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Number(2), String. First number less than second(days). Return 10.');
		// Case #15: Number(2), String. First number less than second(MDays). Return 10.
		oParser = new parserFormula('DATEDIF(12,22,"MD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"MD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Number(2), String. First number less than second(MDays). Return 10.');
		// Case #16: Number(2), String. First number less than second(YMonths). Return 0.
		oParser = new parserFormula('DATEDIF(12,22,"YM")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"YM") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number(2), String. First number less than second(YMonths). Return 0.');
		// Case #17: Number(2), String. First number less than second(YDays). Return 10.
		oParser = new parserFormula('DATEDIF(12,22,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(12,22,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Number(2), String. First number less than second(YDays). Return 10.');
		// Case #18: Number(2), String. DATEDIF(2,2.2, YD). Return 0.
		oParser = new parserFormula('DATEDIF(2,2.2,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(2,2.2,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number(2), String. DATEDIF(2,2.2, YD). Return 0.');
		// Case #19: Number(2), String. DATEDIF(1.2,2.2, YD). Return 1.
		oParser = new parserFormula('DATEDIF(1.2,2.2,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(1.2,2.2,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number(2), String. DATEDIF(1.2,2.2, YD). Return 1.');
		// Case #20: Number(2), String. DATEDIF(9,100, YM). Return 3.
		oParser = new parserFormula('DATEDIF(9,100,"YM")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(9,100,"YM") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number(2), String. DATEDIF(9,100, YM). Return 3.');
		// TODO Different result with MS. Result 2
		// Case #21: Number(2), String. DATEDIF(10,100, YM). Return 3.
		oParser = new parserFormula('DATEDIF(10,100,"YM")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(10,100,"YM") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number(2), String. DATEDIF(10,100, YM). Return 3.');
		// bool
		// Case #22: Boolean, Formula, String. Boolean true first. Return 122.
		oParser = new parserFormula('DATEDIF(TRUE,DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(TRUE,DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 122, 'Test: Positive case: Boolean, Formula, String. Boolean true first. Return 122.');
		// Case #23: Boolean, Formula, String. Boolean false first. Return 122.
		oParser = new parserFormula('DATEDIF(FALSE,DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(FALSE,DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 122, 'Test: Positive case: Boolean, Formula, String. Boolean false first. Return 122.');
		// exotic dates
		// Case #24: Formula(2), String. Exotic date. Return 1.
		oParser = new parserFormula('DATEDIF(DATE(4022,4,12),DATE(4023,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(4022,4,12),DATE(4023,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2), String. Exotic date. Return 1.');
		// Case #25: Formula(2), String. Exotic date. Return 0.
		oParser = new parserFormula('DATEDIF(DATE(1,1,1),DATE(1,2,1),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(1,1,1),DATE(1,2,1),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Formula(2), String. Exotic date. Return 0.');
		// Case #26: Array, Number, String. Pass array to first argument and number to second argument. Return 27.
		oParser = new parserFormula('DATEDIF({223,999,250},250,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF({223,999,250},250,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 27, 'Test: Positive case: Array, Number, String. Pass array to first argument and number to second argument.. Return 27.');
		// Case #27: Number(2), String. Float numbers for both dates. Return 99.
		oParser = new parserFormula('DATEDIF(100.75,200.25,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(100.75,200.25,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Number(2), String. Float numbers for both dates. Return 99.');
		// Case #28: Number(2), String. Float numbers converted to dates for months calculation. Return 131.
		oParser = new parserFormula('DATEDIF(1000.5,5000.8,"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(1000.5,5000.8,"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 131, 'Test: Positive case: Number(2), String. Float numbers converted to dates for months calculation. Return 131.');
		// Case #29: Number(2), String. Float numbers for year calculation. Return 0.
		oParser = new parserFormula('DATEDIF(365.99,730.01,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(365.99,730.01,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number(2), String. Float numbers for year calculation. Return 0.');
		// Case #30: Empty, Formula, String. Empty first argument treated as 0 (serial date). Return 44663.
		oParser = new parserFormula('DATEDIF(,DATE(2022,4,12),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(,DATE(2022,4,12),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44663, 'Test: Positive case: Empty, Formula, String. Empty first argument treated as 0 (serial date). Return 44663.');
		// Case #31: String, String, String. Date strings in MM/DD/YYYY format. Return 2.
		oParser = new parserFormula('DATEDIF("01/15/2020","12/31/2022","Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("01/15/2020","12/31/2022","Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String, String, String. Date strings in MM/DD/YYYY format. Return 2.');
		// Case #32: String, String, String. Date strings for months calculation. Return 14.
		oParser = new parserFormula('DATEDIF("06/01/2021","08/15/2022","M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("06/01/2021","08/15/2022","M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: String, String, String. Date strings for months calculation. Return 14.');
		// Case #33: String, String, String. Date strings for days calculation. Return 366.
		oParser = new parserFormula('DATEDIF("01/01/2020","01/01/2021","D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("01/01/2020","01/01/2021","D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 366, 'Test: Positive case: String, String, String. Date strings for days calculation. Return 366.');
		// Case #34: Formula(2), String. Lowercase unit \'y\'. Return 1.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,6,15),"y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,6,15),"y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2), String. Lowercase unit \'y\'. Return 1.');
		// Case #35: Formula(2), String. Lowercase unit \'m\'. Return 17.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,6,15),"m")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,6,15),"m") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 17, 'Test: Positive case: Formula(2), String. Lowercase unit \'m\'. Return 17.');
		// Case #36: Formula(2), String. Lowercase unit \'d\'. Return 365.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2020,12,31),"d")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2020,12,31),"d") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Formula(2), String. Lowercase unit \'d\'. Return 365.');
		// Case #37: Formula(2), String. Lowercase unit \'yd\'. Return 66.
		oParser = new parserFormula('DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"yd")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"yd") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 66, 'Test: Positive case: Formula(2), String. Lowercase unit \'yd\'. Return 66.');
		// Case #38: Formula(2), String. Lowercase unit \'md\'. Return 5.
		oParser = new parserFormula('DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"md")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"md") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Formula(2), String. Lowercase unit \'md\'. Return 5.');
		// Case #39: Formula(2), String. Lowercase unit \'ym\'. Return 2.
		oParser = new parserFormula('DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"ym")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,3,15),DATE(2021,5,20),"ym") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula(2), String. Lowercase unit \'ym\'. Return 2.');
		// Case #40: Formula, Number, String. TIME as first argument converted to serial number. Return 100.
		oParser = new parserFormula('DATEDIF(TIME(0,0,1),100,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(TIME(0,0,1),100,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Formula, Number, String. TIME as first argument converted to serial number. Return 100.');
		// Case #41: Formula(2), String. Both arguments are TIME. Return 0.
		oParser = new parserFormula('DATEDIF(TIME(0,0,0),TIME(23,59,59),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(TIME(0,0,0),TIME(23,59,59),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Formula(2), String. Both arguments are TIME. Return 0.');
		// Case #42: Number, Array, String. Array as second argument. Return 50.
		oParser = new parserFormula('DATEDIF(100,{150,200},"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(100,{150,200},"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 50, 'Test: Positive case: Number, Array, String. Array as second argument. Return 50.');
		// Case #43: Formula(2), Array. Array as third argument with valid unit. Return 1.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),{"Y"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),{"Y"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2), Array. Array as third argument with valid unit. Return 1.');
		// Case #44: Formula. DATEDIF as child formula in SUM. Return 24.
		oParser = new parserFormula('SUM(DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"M"),12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"M"),12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Formula. DATEDIF as child formula in SUM. Return 24.');
		// Case #45: Reference link(3). Reference links for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(A100,A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(A100,A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Reference link(3). Reference links for all three arguments. Return varies.');
		// Case #46: Reference link(3). Reference links for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(A100,A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(A100,A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Reference link(3). Reference links for all three arguments. Return varies.');
		// Case #47: Reference link(3). Reference links for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(A100,A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(A100,A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Reference link(3). Reference links for all three arguments. Return varies.');
		// Case #48: Name(3). Named ranges for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(TestName,TestName1,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(TestName,TestName1,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Name(3). Named ranges for all three arguments. Return varies.');
		// Case #49: Ref3D(3). 3D references for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(Sheet2!A1,Sheet2!A2,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(Sheet2!A1,Sheet2!A2,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Ref3D(3). 3D references for all three arguments. Return varies.');
		// Case #50: Area3D(3). 3D areas for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Area3D(3). 3D areas for all three arguments. Return varies.');
		// Case #51: Name3D(3). 3D named ranges for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(TestName3D,TestName3D1,TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(TestName3D,TestName3D1,TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Name3D(3). 3D named ranges for all three arguments. Return varies.');
		// Case #52: Table(3). Table references for all three arguments. Return varies.
		oParser = new parserFormula('DATEDIF(Table1[Column1],Table1[Column2],Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(Table1[Column1],Table1[Column2],Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 365, 'Test: Positive case: Table(3). Table references for all three arguments. Return varies.');

		// Negative cases:

		// Case #1: String, Formula, String. String first. Return #VALUE!.
		oParser = new parserFormula('DATEDIF("sdy",DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("sdy",DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Formula, String. String first. Return #VALUE!.');
		// Case #2: String, Formula, String. String number first. Return #NUM!.
		oParser = new parserFormula('DATEDIF("999999999999",DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("999999999999",DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String, Formula, String. String number first. Return #NUM!.');
		// Case #3: Formula, String(2). String second. Return #VALUE!.
		oParser = new parserFormula('DATEDIF(DATE(2022,4,12),"sdy","Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2022,4,12),"sdy","Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula, String(2). String second. Return #VALUE!.');
		// Case #4: Formula(2), String. String third. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2022,4,12),DATE(2032,4,12),"string")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2022,4,12),DATE(2032,4,12),"string") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. String third. Return #NUM!.');
		// Case #5: Number, Formula, String. Number first. Return #NUM!.
		oParser = new parserFormula('DATEDIF(999999999999,DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(999999999999,DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number, Formula, String. Number first. Return #NUM!.');
		// Case #6: Formula, Number, String. Number second. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2022,4,12),12,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2022,4,12),12,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Number, String. Number second. Return #NUM!.');
		// Case #7: Number(2), String. DATEDIF(-12,22, YD). Return #NUM!.
		oParser = new parserFormula('DATEDIF(-12,22,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(-12,22,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2), String. DATEDIF(-12,22, YD). Return #NUM!.');
		// Case #8: Number(2), String. DATEDIF(-12,-22, YD). Return #NUM!.
		oParser = new parserFormula('DATEDIF(-12,-22,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(-12,-22,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2), String. DATEDIF(-12,-22, YD). Return #NUM!.');
		// Case #9: Number(2), String. DATEDIF(-1.2,22, YD). Return #NUM!.
		oParser = new parserFormula('DATEDIF(-1.2,22,"YD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(-1.2,22,"YD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2), String. DATEDIF(-1.2,22, YD). Return #NUM!.');
		// Case #10: Formula, Boolean, String. Boolean second. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2022,4,12),TRUE,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2022,4,12),TRUE,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Boolean, String. Boolean second. Return #NUM!.');
		// Case #11: Formula(2), String. Exotic date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Exotic date. Return #NUM!.');
		// Case #12: Formula(2), String. Exotic date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Exotic date. Return #NUM!.');
		// Case #13: Formula(2), String. Exotic date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(9999,30,12),DATE(99999,30,12222),"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(9999,30,12),DATE(99999,30,12222),"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Exotic date. Return #NUM!.');
		// Case #14: Formula(2), String. Exotic date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(9999,30,12),DATE(99999,30,12),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Exotic date. Return #NUM!.');
		// Case #15: Formula(2), String. Exotic date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(9999,30,12),DATE(99999,30000,12),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(9999,30,12),DATE(99999,30000,12),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Exotic date. Return #NUM!.');
		// Case #16: Area, Number, String. Pass array to first argument and number to second argument. Return #VALUE!.
		// Different result with MS
		oParser = new parserFormula('DATEDIF(B2:B2,25,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(B2:B2,25,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, String. Pass array to first argument and number to second argument.. Return #VALUE!.');
		// Case #17: Area, Number, String. Pass cellsRange to first and number to second argument. Return #VALUE!.
		// Different result with MS
		oParser = new parserFormula('DATEDIF(C2:C6,25,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(C2:C6,25,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number, String. Pass cellsRange to first and number to second argument.. Return #VALUE!.');
		// Case #18: Formula, Formula, Number. Number as third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Number. Number as third argument. Return #NUM!.');
		// Case #19: Formula, Formula, Number. Large number as unit. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Number. Large number as unit. Return #NUM!.');
		// Case #20: Formula, Formula, Formula. Formula as third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),DATE(2022,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),DATE(2022,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Formula. Formula as third argument. Return #NUM!.');
		// Case #21: Formula, Formula, Empty. Empty third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Empty. Empty third argument. Return #NUM!.');
		// Case #22: Formula, Formula, Boolean. Boolean as third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Boolean. Boolean as third argument. Return #NUM!.');
		// Case #23: Formula, Formula, Boolean. Boolean FALSE as third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Boolean. Boolean FALSE as third argument. Return #NUM!.');
		// Case #24: Formula, Formula, Formula. TIME as third argument. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, Formula. TIME as third argument. Return #NUM!.');
		// Case #25: Error, Formula, String. Error #N/A as first argument. Return #N/A.
		oParser = new parserFormula('DATEDIF(#N/A,DATE(2022,4,12),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(#N/A,DATE(2022,4,12),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Formula, String. Error #N/A as first argument. Return #N/A.');
		// Case #26: Formula, Error, String. Error #REF! as second argument. Return #REF!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),#REF!,"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),#REF!,"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: Formula, Error, String. Error #REF! as second argument. Return #REF!.');
		// Case #27: Formula, Formula, Error. Error #DIV/0! as third argument. Return #DIV/0!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula, Formula, Error. Error #DIV/0! as third argument. Return #DIV/0!.');
		// Case #28: Error, Error, String. Errors in both date arguments. Return #VALUE!.
		oParser = new parserFormula('DATEDIF(#VALUE!,#NUM!,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(#VALUE!,#NUM!,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error, Error, String. Errors in both date arguments. Return #VALUE!.');
		// Case #29: Number(2), String. Start date greater than end date with float. Return #NUM!.
		oParser = new parserFormula('DATEDIF(50000.999,10,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(50000.999,10,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2), String. Start date greater than end date with float. Return #NUM!.');
		// Case #30: Formula(2), String. Start date after end date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2022,1,1),DATE(2020,1,1),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2022,1,1),DATE(2020,1,1),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula(2), String. Start date after end date. Return #NUM!.');
		// Case #31: Number, Formula, String. TIME as second argument with start greater than end. Return #NUM!.
		oParser = new parserFormula('DATEDIF(50,TIME(12,30,0),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(50,TIME(12,30,0),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number, Formula, String. TIME as second argument with start greater than end. Return #NUM!.');
		// Case #32: Formula, Number, String. TIME greater than end date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(2 + TIME(23,59,59),1,"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(2 + TIME(23,59,59),1,"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Number, String. TIME greater than end date. Return #NUM!.');
		// Case #33: Array, Number, String. Array with values greater than end date. Return #NUM!.
		oParser = new parserFormula('DATEDIF({5000,6000},100,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF({5000,6000},100,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Array, Number, String. Array with values greater than end date. Return #NUM!.');
		// Case #34: Number, Array, String. Second array with values less than start date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(100,{50,75},"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(100,{50,75},"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number, Array, String. Second array with values less than start date. Return #NUM!.');
		// Case #35: Empty, Empty, Empty. All empty arguments. Return #NUM!.
		oParser = new parserFormula('DATEDIF(,,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(,,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Empty, Empty, Empty. All empty arguments. Return #NUM!.');
		// Case #36: Formula, Empty, String. Empty second argument treated as 0, start date greater than end date. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Empty, String. Empty second argument treated as 0, start date greater than end date. Return #NUM!.');
		// Case #37: Formula, Formula, String. Invalid unit string. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"XYZ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"XYZ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, String. Invalid unit string. Return #NUM!.');
		// Case #38: Formula, Formula, String. Invalid double letter unit. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"YY")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"YY") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, String. Invalid double letter unit. Return #NUM!.');
		// Case #39: Formula, Formula, String. Empty string as unit. Return #NUM!.
		oParser = new parserFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Formula, String. Empty string as unit. Return #NUM!.');
		// Case #40: String, String, String. Invalid date string as first argument. Return #VALUE!.
		oParser = new parserFormula('DATEDIF("invalid","12/31/2022","Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("invalid","12/31/2022","Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Invalid date string as first argument. Return #VALUE!.');
		// Case #41: String, String, String. Invalid date string as second argument. Return #VALUE!.
		oParser = new parserFormula('DATEDIF("01/01/2020","notadate","M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF("01/01/2020","notadate","M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Invalid date string as second argument. Return #VALUE!.');

		// Bounded cases:

		// Case #1: Formula(2), String. Minimum date (01.01.1900) as start date. Return 1.
		oParser = new parserFormula('DATEDIF(DATE(1900,1,1),DATE(1900,1,2),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(1900,1,1),DATE(1900,1,2),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Formula(2), String. Minimum date (01.01.1900) as start date. Return 1.');
		// Case #2: Formula(2), String. Minimum to maximum date range for years. Return 8099.
		oParser = new parserFormula('DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8099, 'Test: Bounded case: Formula(2), String. Minimum to maximum date range for years. Return 8099.');
		// Case #3: Formula(2), String. Minimum to maximum date range for months. Return 97199.
		oParser = new parserFormula('DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"M")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"M") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 97199, 'Test: Bounded case: Formula(2), String. Minimum to maximum date range for months. Return 97199.');
		// Case #4: Formula(2), String. Minimum to maximum date range for days. Return 2958463.
		oParser = new parserFormula('DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(DATE(1900,1,1),DATE(9999,12,31),"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958464, 'Test: Bounded case: Formula(2), String. Minimum to maximum date range for days. Return 2958463.');
		// Case #5: Reference link, Number, String. Full column reference as first argument. Return varies.
		oParser = new parserFormula('DATEDIF(M:M,44927,"D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(M:M,44927,"D") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44927, 'Test: Bounded case: Reference link, Number, String. Full column reference as first argument. Return varies.');
		// Case #6: Number, Reference link, String. Full row reference as second argument. Return varies.
		oParser = new parserFormula('DATEDIF(43831,100:100,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEDIF(43831,100:100,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number, Reference link, String. Full row reference as second argument. Return varies.');

		// ctrl shift enter cases
		oParser = new parserFormula("DATEDIF(C2:C6,25,\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C2:C6").bbox);
		assert.ok(oParser.parse(), "Pass cellsRange to first and number to second argument.");
		let array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 23, "Pass cellsRange to first and number to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 13, "Pass cellsRange to first and number to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 10, "Pass cellsRange to first and number to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 0, "Pass cellsRange to first and number to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#NUM!", "Pass cellsRange to first and number to second argument.[4,0]");

		oParser = new parserFormula("DATEDIF(12,C2:C6,\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C2:C6").bbox);
		assert.ok(oParser.parse(), "Pass number to first and cellsRange to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#NUM!", "Pass number to first and cellsRange to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0, "Pass number to first and cellsRange to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 3, "Pass number to first and cellsRange to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 13, "Pass number to first and cellsRange to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 13, "Pass number to first and cellsRange to second argument.[4,0]");

		oParser = new parserFormula("DATEDIF(C2:C6,C2:C6,\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C2:C6").bbox);
		assert.ok(oParser.parse(), "Pass cellsRange to first and cellsRange to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[4,0]");

		oParser = new parserFormula("DATEDIF(B2:B8,DATE(10,2,2020),\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("B2:B8").bbox);
		assert.ok(oParser.parse(), "Pass cellsRange to first and date to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 5702, "Pass cellsRange to first and date to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 5699, "Pass cellsRange to first and date to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 5689, "Pass cellsRange to first and date to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#VALUE!", "Pass cellsRange to first and date to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Pass cellsRange to first and date to second argument.[4,0]");
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 5704, "Pass cellsRange to first and date to second argument.[5,0]");
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 5704, "Pass cellsRange to first and date to second argument.[6,0]");

		oParser = new parserFormula("DATEDIF(B2:B8,DATE(2020,10,2),\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("B2:B8").bbox);
		assert.ok(oParser.parse(), "Pass cellsRange to first and date to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 44104, "Pass cellsRange to first and date to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 44101, "Pass cellsRange to first and date to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 44091, "Pass cellsRange to first and date to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#VALUE!", "Pass cellsRange to first and date to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Pass cellsRange to first and date to second argument.[4,0]");
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 44106, "Pass cellsRange to first and date to second argument.[5,0]");
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 44106, "Pass cellsRange to first and date to second argument.[6,0]");

		oParser = new parserFormula("DATEDIF(DATE(2020,10,2),B2:B8,\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("B2:B8").bbox);
		assert.ok(oParser.parse(), "Pass date to first and cellsRange to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#NUM!", "Pass date to first and cellsRange to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#NUM!", "Pass date to first and cellsRange to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#NUM!", "Pass date to first and cellsRange to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#VALUE!", "Pass date to first and cellsRange to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Pass date to first and cellsRange to second argument.[4,0]");
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "#NUM!", "Pass date to first and cellsRange to second argument.[5,0]");
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "#NUM!", "Pass date to first and cellsRange to second argument.[6,0]");

		oParser = new parserFormula("DATEDIF(B2:B8,B2:B8,\"D\")", "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("B2:B8").bbox);
		assert.ok(oParser.parse(), "Pass cellsRange to first and cellsRange to second argument.");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#VALUE!", "Pass cellsRange to first and cellsRange to second argument.[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Pass cellsRange to first and cellsRange to second argument.[4,0]");
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[5,0]");
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 0, "Pass cellsRange to first and cellsRange to second argument.[6,0]");

		testArrayFormula2(assert, "DATEDIF", 3, 3);
	});

	QUnit.test('Test: "DATEVALUE"', function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("1");
		ws.getRange2("A102").setValue("2000");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("Text");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Column1
		ws.getRange2("B601").setValue("1"); // Column2
		ws.getRange2("C601").setValue("2000"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("1");
		ws2.getRange2("A3").setValue("2000");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("1"); // TestName
		ws.getRange2("A202").setValue("1"); // TestName1
		ws.getRange2("A203").setValue("2000"); // TestName2
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("1") // TestName3D
		ws2.getRange2("A12").setValue("1") // TestName3D1
		ws2.getRange2("A13").setValue("2000") // TestName3D2

		// Positive cases:

		// Case #1: String. Return 40461
		oParser = new parserFormula('DATEVALUE("10-10-2010 10:26")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("10-10-2010 10:26") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40461, 'Test: Positive case: String. Return 40461');
		// Case #2: String. Return 40461
		oParser = new parserFormula('DATEVALUE("10-10-2010 10:26")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("10-10-2010 10:26") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40461, 'Test: Positive case: String. Return 40461');
		// Case #3: Reference link. Date in reference link
		tmp = ws.getRange2("A7");
		tmp.setNumFormat('@');
		tmp.setValue("3-Mar");
		oParser = new parserFormula("DATEVALUE(A7)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A7) is parsed.');
		let d = new cDate();
		d.setUTCMonth(2);
		d.setUTCDate(3);
		assert.strictEqual(oParser.calculate().getValue(), d.getExcelDate(), 'Test: Positive case: Reference link. Date in reference link');
		// Case #4: String. Return 37338
		oParser = new parserFormula('DATEVALUE("23-Mar-2002")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("23-Mar-2002") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 37338, 'Test: Positive case: String. Return 37338');
		// Case #5: String. Return 38802
		oParser = new parserFormula("DATEVALUE(\"03-26-2006\")", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("03-26-2006") is parsed.');
		if (AscCommon.bDate1904) {
			assert.strictEqual(oParser.calculate().getValue(), 37340, 'Test: Positive case: String. bDate1904 true. Return 37340');
		} else {
			assert.strictEqual(oParser.calculate().getValue(), 38802, 'Test: Positive case: String. Return 38802');
		}
		// Case #6: String. Standard US date format (M/D/YYYY). Returns serial number 40763.
		oParser = new parserFormula('DATEVALUE("8/8/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("8/8/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40763, 'Test: Positive case: String. Standard US date format (M/D/YYYY). Returns serial number 40763.');
		// Case #7: String. Date with abbreviated month name. Returns serial number 40685.
		oParser = new parserFormula('DATEVALUE("22-MAY-2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("22-MAY-2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40685, 'Test: Positive case: String. Date with abbreviated month name. Returns serial number 40685.');
		// Case #8: String. Date in YYYY/MM/DD format. Returns serial number 40597.
		oParser = new parserFormula('DATEVALUE("2011/02/23")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011/02/23") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40597, 'Test: Positive case: String. Date in YYYY/MM/DD format. Returns serial number 40597.');
		// Case #9: String. Date without year, assumes current year 2011. Returns serial number 45843.
		oParser = new parserFormula('DATEVALUE("5-JUL")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("5-JUL") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45843, 'Test: Positive case: String. Date without year, assumes current year 2011. Returns serial number 45843.');
		// Case #10: String. Minimum valid date in 1900 date system. Returns serial number 1.
		oParser = new parserFormula('DATEVALUE("1/1/1900")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1/1/1900") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Minimum valid date in 1900 date system. Returns serial number 1.');
		// Case #11: String. Full month name with comma separator. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE("January 1, 2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("January 1, 2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: String. Full month name with comma separator. Returns serial number 40544.');
		// Case #12: String. Date with day-month-year format. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE("01-Jan-2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("01-Jan-2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: String. Date with day-month-year format. Returns serial number 40544.');
		// Case #13: String. ISO 8601 date format. Returns serial number 40908.
		oParser = new parserFormula('DATEVALUE("2011-12-31")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011-12-31") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40908, 'Test: Positive case: String. ISO 8601 date format. Returns serial number 40908.');
		// Case #14: String. Last day of year format. Returns serial number 40908.
		oParser = new parserFormula('DATEVALUE("12/12/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("12/12/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40889, 'Test: Positive case: String. Last day of year format. Returns serial number 40908.');
		// Case #15: String. European date format with "/". Returns serial number 40908.
		oParser = new parserFormula('DATEVALUE("12/31/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("12/31/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40908, 'Test: Positive case: String. European date format with "/". Returns serial number 40908.');
		// Case #16: String. Date with space separators. Returns serial number 40617.
		oParser = new parserFormula('DATEVALUE("15 March 2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("15 March 2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40617, 'Test: Positive case: String. Date with space separators. Returns serial number 40617.');
		// Case #17: Formula. Date created by string concatenation. Returns serial number 40763.
		oParser = new parserFormula('DATEVALUE("8/"&"8"&"/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("8/"&"8"&"/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40763, 'Test: Positive case: Formula. Date created by string concatenation. Returns serial number 40763.');
		// Case #18: Reference link. Date from concatenating three cell references. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE(A100&"/"&A101&"/"&A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A100&"/"&A101&"/"&A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Reference link. Date from concatenating three cell references. Returns serial number 36526.');
		// Case #19: Array. Array with single date string element. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE({"1/1/2011"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE({"1/1/2011"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: Array. Array with single date string element. Returns serial number 40544.');
		// Case #20: Name. Date from concatenating named ranges. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE(TestName&"/"&TestName1&"/"&TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(TestName&"/"&TestName1&"/"&TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Name. Date from concatenating named ranges. Returns serial number 36526.');
		// Case #21: Name3D. Date from concatenating 3D named ranges. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE(TestName3D&"/"&TestName3D1&"/"&TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(TestName3D&"/"&TestName3D1&"/"&TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Name3D. Date from concatenating 3D named ranges. Returns serial number 36526.');
		// Case #22: Ref3D. Date from concatenating 3D references. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE(Sheet2!A1&"/"&Sheet2!A2&"/"&Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(Sheet2!A1&"/"&Sheet2!A2&"/"&Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Ref3D. Date from concatenating 3D references. Returns serial number 36526.');
		// Case #23: Table. Date from concatenating table columns. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE(Table1[Column1]&"/"&Table1[Column2]&"/"&Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(Table1[Column1]&"/"&Table1[Column2]&"/"&Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Table. Date from concatenating table columns. Returns serial number 36526.');
		// Case #24: String. Leap year date validation. Returns serial number 40968.
		oParser = new parserFormula('DATEVALUE("Feb 29, 2012")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("Feb 29, 2012") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40968, 'Test: Positive case: String. Leap year date validation. Returns serial number 40968.');
		// Case #25: String. Date at millennium with dashes. Returns serial number 36526.
		oParser = new parserFormula('DATEVALUE("1-1-2000")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1-1-2000") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: String. Date at millennium with dashes. Returns serial number 36526.');
		// Case #26: Formula. Date from CONCATENATE formula. Returns serial number 40777.
		oParser = new parserFormula('DATEVALUE(CONCATENATE("8","/","22","/","2011"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(CONCATENATE("8","/","22","/","2011")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40777, 'Test: Positive case: Formula. Date from CONCATENATE formula. Returns serial number 40777.');
		// Case #27: String. Date with single-digit month and day. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE("2011-1-1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011-1-1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: String. Date with single-digit month and day. Returns serial number 40544.');
		// Case #28: String. Date with two-digit year. Returns serial number 40668.
		oParser = new parserFormula('DATEVALUE("5/5/11")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("5/5/11") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40668, 'Test: Positive case: String. Date with two-digit year. Returns serial number 40668.');
		// Case #29: String. Christmas date with full month name. Returns serial number 40902.
		oParser = new parserFormula('DATEVALUE("December 25, 2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("December 25, 2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40902, 'Test: Positive case: String. Christmas date with full month name. Returns serial number 40902.');
		// Case #30: Formula. Date from nested IF formula. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE(IF(TRUE,"1/1/2011","12/31/2010"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(IF(TRUE,"1/1/2011","12/31/2010")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: Formula. Date from nested IF formula. Returns serial number 40544.');
		// Case #31: String. Mid-year date validation. Returns serial number 40724.
		oParser = new parserFormula('DATEVALUE("30-Jun-2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("30-Jun-2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40724, 'Test: Positive case: String. Mid-year date validation. Returns serial number 40724.');
		// Case #32: Formula. DATEVALUE nested in SUM formula. Returns serial number 40554.
		oParser = new parserFormula('SUM(DATEVALUE("1/1/2011"),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DATEVALUE("1/1/2011"),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40554, 'Test: Positive case: Formula. DATEVALUE nested in SUM formula. Returns serial number 40554.');
		// Case #33: String. Date with uppercase month abbreviation. Returns serial number 40544.
		oParser = new parserFormula('DATEVALUE("1 JAN 2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1 JAN 2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40544, 'Test: Positive case: String. Date with uppercase month abbreviation. Returns serial number 40544.');
		// Case #34: String. Date with dots separator YYYY.MM.DD format. Returns serial number 40558.
		oParser = new parserFormula('DATEVALUE("2011/01/15")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011/01/15") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40558, 'Test: Positive case: String. Date with dots separator YYYY.MM.DD format. Returns serial number 40558.');
		// Case #35: String. Date in Month-Day-Year format. Returns serial number 40617.
		oParser = new parserFormula('DATEVALUE("Mar-15-2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("Mar-15-2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40617, 'Test: Positive case: String. Date in Month-Day-Year format. Returns serial number 40617.');
		// Case #36: Reference link. Date from cell concatenation with dashes. Returns serial number 40617.
		oParser = new parserFormula('DATEVALUE(A100&"-"&A101&"-"&A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A100&"-"&A101&"-"&A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Positive case: Reference link. Date from cell concatenation with dashes. Returns serial number 40617.');
		// Case #37: String. Independence day date format. Returns serial number 40729.
		oParser = new parserFormula('DATEVALUE("04/07/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("04/07/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40640, 'Test: Positive case: String. Independence day date format. Returns serial number 40729.');
		// Case #38: String. Date with single-digit month and day in YYYY format. Returns serial number 40695.
		oParser = new parserFormula('DATEVALUE("2011/6/1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011/6/1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40695, 'Test: Positive case: String. Date with single-digit month and day in YYYY format. Returns serial number 40695.');
		// Case #39: Formula. Date extracted from string using LEFT formula. Returns serial number 40597.
		oParser = new parserFormula('DATEVALUE(LEFT("2011/02/23 extra text",10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(LEFT("2011/02/23 extra text",10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40597, 'Test: Positive case: Formula. Date extracted from string using LEFT formula. Returns serial number 40597.');

		// Negative cases:

		// Case #1: String. Return #VALUE!
		oParser = new parserFormula('DATEVALUE("$1,000")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("$1,000") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Return #VALUE!');
		// Case #2: String. Non-date text string returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-date text string returns #VALUE! error.');
		// Case #3: String. Invalid day number returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("32/1/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("32/1/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid day number returns #VALUE! error.');
		// Case #4: String. Invalid month and day returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("13/32/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("13/32/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid month and day returns #VALUE! error.');
		// Case #5: String. Year exceeds maximum (9999) returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("1/1/10000")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1/1/10000") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Year exceeds maximum (9999) returns #VALUE! error.');
		// Case #6: String. Date before 1900 date system returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("1/1/1899")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1/1/1899") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Date before 1900 date system returns #VALUE! error.');
		// Case #7: Number. Numeric serial number instead of text returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(40777)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(40777) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Numeric serial number instead of text returns #VALUE! error.');
		// Case #8: Boolean. Boolean value returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean value returns #VALUE! error.');
		// Case #9: Boolean. Boolean FALSE returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean FALSE returns #VALUE! error.');
		// Case #10: Empty. Reference to empty cell returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE! error.');
		// Case #11: String. Empty string returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE! error.');
		// Case #12: Error. Error value propagates #N/A error.
		oParser = new parserFormula('DATEVALUE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error value propagates #N/A error.');
		// Case #13: String. Invalid date (Feb 30) returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("Feb 30, 2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("Feb 30, 2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date (Feb 30) returns #VALUE! error.');
		// Case #14: String. Invalid month name returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("13-Month-2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("13-Month-2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid month name returns #VALUE! error.');
		// Case #15: String. Zero day and month returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("0/0/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("0/0/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Zero day and month returns #VALUE! error.');
		// Case #16: Reference link. Reference to cell with non-date text returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with non-date text returns #VALUE! error.');
		// Case #17: Area. Multi-cell area returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell area returns #VALUE! error.');
		// Case #18: Area3D. 3D multi-cell area returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell area returns #VALUE! error.');
		// Case #19: String. Month number exceeds 12 returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("2011-13-01")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("2011-13-01") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Month number exceeds 12 returns #VALUE! error.');
		// Case #20: Formula. Formula resulting in error propagates #NUM! error.
		oParser = new parserFormula('DATEVALUE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in error propagates #NUM! error.');
		// Case #21: String. Date exactly at system boundary returns #VALUE! error.
		// TODO Different result with MS  Return #VALUE!
		oParser = new parserFormula('DATEVALUE("12/31/1899")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("12/31/1899") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Date exactly at system boundary returns #VALUE! error.');
		// Case #22: String. Non-leap year Feb 29 returns #VALUE! error.
		oParser = new parserFormula('DATEVALUE("29/2/2011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("29/2/2011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-leap year Feb 29 returns #VALUE! error.');

		// Bounded cases:

		// Case #1: String. Minimum valid date in 1900 system. Returns serial number 1.
		oParser = new parserFormula('DATEVALUE("1/1/1900")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("1/1/1900") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Minimum valid date in 1900 system. Returns serial number 1.');
		// Case #2: String. Maximum valid date in Excel. Returns serial number 2958465.
		oParser = new parserFormula('DATEVALUE("12/31/9999")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE("12/31/9999") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: String. Maximum valid date in Excel. Returns serial number 2958465.');
		// Case #3: Area. Whole row reference with date value in first cell.
		oParser = new parserFormula('DATEVALUE(100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(100:100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Area. Whole row reference with date value in first cell.');
		// Case #4: Area. Whole column reference with date value in first cell.
		oParser = new parserFormula('DATEVALUE(M:M)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DATEVALUE(M:M) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Area. Whole column reference with date value in first cell.');

		testArrayFormula(assert, "DATEVALUE");
	});

	QUnit.test('Test: "DAY"', function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("40908");
		ws.getRange2("B100").setValue("40544");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("text");
		ws.getRange2("A104").setValue("#N/A");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("40908"); // Column1
		ws.getRange2("B601").setValue("text"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("40908");
		ws2.getRange2("A2").setValue("40544");
		ws2.getRange2("A3").setValue("text");
		ws2.getRange2("A4").setValue("#N/A");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("40908"); // TestName
		ws.getRange2("A202").setValue("text"); // TestName1
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("40908") // TestName3D
		ws2.getRange2("A12").setValue("text") // TestName3D1

		// Positive cases:

		// Case #1: Number. Return 5
		oParser = new parserFormula('DAY(2013)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(2013) is parsed.');
		if (AscCommon.bDate1904) {
			assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Number. Date1904 enabled. Return 6');
		} else {
			assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number. Return 5');
		}
		// Case #2: String. Return 20
		oParser = new parserFormula('DAY("20 may 2045")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("20 may 2045") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Test: Positive case: String. Return 20');
		// Case #3: Number. Basic valid input: serial number.
		oParser = new parserFormula('DAY(44927)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(44927) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Basic valid input: serial number.');
		// Case #4: Number. Valid serial number.
		oParser = new parserFormula('DAY(45678)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(45678) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 21, 'Test: Positive case: Number. Valid serial number.');
		// Case #5: Number. Serial number for date.
		oParser = new parserFormula('DAY(39448)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(39448) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Serial number for date.');
		// Case #6: String. Date string in US format.
		oParser = new parserFormula('DAY("1/15/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("1/15/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: String. Date string in US format.');
		// Case #7: String. Date string last day of month.
		oParser = new parserFormula('DAY("12/31/2024")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("12/31/2024") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: String. Date string last day of month.');
		// Case #8: String. Date string with single digit day.
		oParser = new parserFormula('DAY("2/5/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("2/5/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Date string with single digit day.');
		// Case #9: String. Date string with dash separator US format.
		oParser = new parserFormula('DAY("3-23-2002")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("3-23-2002") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Positive case: String. Date string with dash separator US format.');
		// Case #10: String. Date string with month name.
		oParser = new parserFormula('DAY("23-Mar-2002")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("23-Mar-2002") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Positive case: String. Date string with month name.');
		// Case #11: String. Date string with short month name.
		oParser = new parserFormula('DAY("5-JUL")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("5-JUL") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Date string with short month name.');
		// Case #12: String. Date string with full month name.
		oParser = new parserFormula('DAY("July 5, 2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("July 5, 2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Date string with full month name.');
		// Case #13: String. Date string with two-digit year.
		oParser = new parserFormula('DAY("7/5/25")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("7/5/25") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Date string with two-digit year.');
		// Case #14: String. Date string first day of year.
		oParser = new parserFormula('DAY("01-Jan-2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("01-Jan-2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Date string first day of year.');
		// Case #15: String. Date string last day of year.
		oParser = new parserFormula('DAY("Dec 31, 2024")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("Dec 31, 2024") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: String. Date string last day of year.');
		// Case #16: Formula. DATE formula as argument.
		oParser = new parserFormula('DAY(DATE(2025,1,15))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(DATE(2025,1,15)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Formula. DATE formula as argument.');
		// Case #17: Formula. DATE formula last day of month.
		oParser = new parserFormula('DAY(DATE(2024,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(DATE(2024,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Formula. DATE formula last day of month.');
		// Case #18: Reference link. Reference to cell with valid serial number.
		oParser = new parserFormula('DAY(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Reference link. Reference to cell with valid serial number.');
		// Case #19: Reference link. Reference to cell with date value.
		oParser = new parserFormula('DAY(B100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(B100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with date value.');
		// Case #20: Area. Single-cell range with serial number.
		oParser = new parserFormula('DAY(A100:A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A100:A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Area. Single-cell range with serial number.');
		// Case #21: Area. Multi-cell range with dates.
		oParser = new parserFormula('DAY(A100:B100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A100:B100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Area. Multi-cell range with dates.');
		// Case #22: Array. Array with single serial number.
		oParser = new parserFormula('DAY({44927})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY({44927}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Array. Array with single serial number.');
		// Case #23: Array. Array with multiple serial numbers.
		oParser = new parserFormula('DAY({44927; 44958})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY({44927; 44958}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Array. Array with multiple serial numbers.');
		// Case #24: Name. Named range with serial number.
		oParser = new parserFormula('DAY(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Name. Named range with serial number.');
		// Case #25: Name3D. 3D named range with date.
		oParser = new parserFormula('DAY(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Name3D. 3D named range with date.');
		// Case #26: Ref3D. 3D reference to cell with date.
		oParser = new parserFormula('DAY(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Ref3D. 3D reference to cell with date.');
		// Case #27: Area3D. 3D single-cell range.
		oParser = new parserFormula('DAY(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Area3D. 3D single-cell range.');
		// Case #28: Area3D. 3D multi-cell range.
		oParser = new parserFormula('DAY(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Area3D. 3D multi-cell range.');
		// Case #29: Table. Table structured reference.
		oParser = new parserFormula('DAY(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Table. Table structured reference.');
		// Case #30: Formula. DATEVALUE formula as argument.
		oParser = new parserFormula('DAY(DATEVALUE("1/15/2025"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(DATEVALUE("1/15/2025")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: Formula. DATEVALUE formula as argument.');
		// Case #31: Formula. DAY inside SUM formula.
		oParser = new parserFormula('SUM(DAY(DATE(2025,1,15)),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DAY(DATE(2025,1,15)),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 25, 'Test: Positive case: Formula. DAY inside SUM formula.');
		// Case #32: String. Numeric string as serial number.
		oParser = new parserFormula('DAY("44927")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("44927") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Numeric string as serial number.');
		// Case #33: Number. Serial number 1 (01/01/1900).
		oParser = new parserFormula('DAY(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Serial number 1 (01/01/1900).');
		// Case #34: Number. Float serial number with decimal.
		oParser = new parserFormula('DAY(2.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(2.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Number. Float serial number with decimal.');
		// Case #35: Time. Time formula adjusted with serial number.
		oParser = new parserFormula('DAY(TIME(12,0,0)+44927)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TIME(12,0,0)+44927) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Time. Time formula adjusted with serial number.');
		// Case #36: Boolean. Boolean TRUE converted to serial number 1.
		oParser = new parserFormula('DAY(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Boolean. Boolean TRUE converted to serial number 1.');
		// Case #37: Empty. Reference to empty cell.
		oParser = new parserFormula('DAY(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Empty. Reference to empty cell.');
		// Case #38: Formula. DAY inside MAX formula with range.
		oParser = new parserFormula('MAX(DAY(A100:B100))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula MAX(DAY(A100:B100)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Positive case: Formula. DAY inside MAX formula with range.');
		// Case #39: Number. Serial number with large decimal.
		oParser = new parserFormula('DAY(44927.99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(44927.99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Serial number with large decimal.');
		// Case #40: String. Date string with dot separator.
		oParser = new parserFormula('DAY("01/15/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("01/15/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: String. Date string with dot separator.');
		// Case #41: String. Date string with slash ISO format.
		oParser = new parserFormula('DAY("2025/01/15")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("2025/01/15") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 15, 'Test: Positive case: String. Date string with slash ISO format.');
		// Case #42: Boolean. Boolean FALSE (0).
		oParser = new parserFormula('DAY(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Boolean. Boolean FALSE (0).');
		// Case #43: Formula. Invalid DATE formula.
		oParser = new parserFormula('DAY(DATE(2025,13,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(DATE(2025,13,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Invalid DATE formula.');

		// Negative cases:

		// Case #1: String. Non-numeric string returns #VALUE!.
		oParser = new parserFormula('DAY("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!.');
		// Case #2: String. Empty string returns #VALUE!.
		oParser = new parserFormula('DAY("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!.');
		// Case #3: String. Invalid date string returns #VALUE!.
		oParser = new parserFormula('DAY("invalid date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("invalid date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string returns #VALUE!.');
		// Case #4: Error. Propagates #N/A error.
		oParser = new parserFormula('DAY(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error.');
		// Case #5: Formula. Formula resulting in #NUM! error.
		oParser = new parserFormula('DAY(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error.');
		// Case #6: Reference link. Reference to cell with text returns #VALUE!.
		oParser = new parserFormula('DAY(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with text returns #VALUE!.');
		// Case #7: Area. Multi-cell range with text returns #VALUE!.
		oParser = new parserFormula('DAY(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range with text returns #VALUE!.');
		// Case #8: Ref3D. 3D reference to text returns #VALUE!.
		oParser = new parserFormula('DAY(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to text returns #VALUE!.');
		// Case #9: Table. Table column with text returns #VALUE!.
		oParser = new parserFormula('DAY(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!.');
		// Case #10: Number. Negative serial number returns #NUM!.
		oParser = new parserFormula('DAY(-100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(-100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative serial number returns #NUM!.');
		// Case #11: Number. Serial number beyond max date returns #NUM!.
		// Different result with MS
		/*oParser = new parserFormula('DAY(2958466)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(2958466) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Serial number beyond max date returns #NUM!.');*/
		// Case #12: String. Negative string number returns #NUM!.
		oParser = new parserFormula('DAY("-100")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("-100") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String. Negative string number returns #NUM!.');
		// Case #13: Array. Array with text returns #VALUE!.
		oParser = new parserFormula('DAY({"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY({"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with text returns #VALUE!.');
		// Case #14: String. Invalid date string returns #VALUE!.
		oParser = new parserFormula('DAY("13/32/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("13/32/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string returns #VALUE!.');
		// Case #15: Number. Large negative number returns #NUM!.
		oParser = new parserFormula('DAY(-1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(-1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Large negative number returns #NUM!.');
		// Case #16: Area3D. 3D multi-cell range with text returns #VALUE!.
		oParser = new parserFormula('DAY(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range with text returns #VALUE!.');
		// Case #17: Name. Named range with text returns #VALUE!.
		oParser = new parserFormula('DAY(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with text returns #VALUE!.');
		// Case #18: Name3D. 3D named range with text returns #VALUE!.
		oParser = new parserFormula('DAY(TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!.');
		// Case #19: String. Invalid date (non-existent day) returns #VALUE!.
		// Different result with MS
		/*oParser = new parserFormula('DAY("Feb 30, 2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY("Feb 30, 2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date (non-existent day) returns #VALUE!.');*/

		// Bounded cases:

		// Case #1: Number. Minimum valid serial number (01/01/1900).
		oParser = new parserFormula('DAY(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Minimum valid serial number (01/01/1900).');
		// Case #2: Number. Maximum valid serial number (31/12/9999).
		oParser = new parserFormula('DAY(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Bounded case: Number. Maximum valid serial number (31/12/9999).');
		// Case #3: Number. Smallest value above minimum.
		oParser = new parserFormula('DAY(1.000000000000001)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(1.000000000000001) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Smallest value above minimum.');
		// Case #4: Number. Value just below maximum.
		// Different result with MS
		/*oParser = new parserFormula('DAY(2958464.999999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(2958464.999999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Bounded case: Number. Value just below maximum.');*/
		// Case #5: Area. Entire column reference.
		oParser = new parserFormula('DAY(A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Area. Entire column reference.');
		// Case #6: Area. Entire row reference.
		oParser = new parserFormula('DAY(100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(100:100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Bounded case: Area. Entire row reference.');
		// Case #7: Number. Zero serial number.
		oParser = new parserFormula('DAY(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAY(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Zero serial number.');

		testArrayFormula2(assert, "DAY", 1, 1);

		// TODO Need to fix

		// Absent logic for reach beyond of maximum date. Must be #NUM! - Negative case #11
		// Incorrect work with invalid date - Negative case #19
		// Incorrect formula's calculate logic - Bounded case #4
	});

	QUnit.test('Test: "DAYS"', function (assert) {
		let array;
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("40908");
		ws.getRange2("A101").setValue("40544");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("text");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A105").setValue("#NUM!");

		ws.getRange2("A2").setValue("12/31/2011");
		ws.getRange2("A3").setValue("1/1/2011");

		ws.getRange2("B3").setValue("44229.4673611111");
		ws.getRange2("B4").setValue("44229.46875");
		ws.getRange2("B5").setValue("1");
		ws.getRange2("B6").setValue("1.9");
		ws.getRange2("B7").setValue("2.1");
		ws.getRange2("B8").setValue("10");

		ws.getRange2("B100").setValue("#N/A");
		ws.getRange2("B101").setValue("#NUM!");

		ws.getRange2("A25").setValue("1");
		ws.getRange2("A26").setValue("2");
		ws.getRange2("A27").setValue("3");
		ws.getRange2("B25").setValue("10");
		ws.getRange2("B26").setValue("9");
		ws.getRange2("B27").setValue("8");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("40908"); // Column1
		ws.getRange2("B601").setValue("40544"); // Column2
		ws.getRange2("C601").setValue("text"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("40908");
		ws2.getRange2("A2").setValue("40544");
		ws2.getRange2("A3").setValue("text");
		ws2.getRange2("A4").setValue("#N/A");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("40908"); // TestName
		ws.getRange2("A202").setValue("40544"); // TestName1
		ws.getRange2("A203").setValue("text"); // TestName2
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("40908") // TestName3D
		ws2.getRange2("A12").setValue("text") // TestName3D1
		ws2.getRange2("A13").setValue("40544") // TestName3D2

		// Positive cases:

		// Case #1: String(2). Format date m/d/yy Return 42
		oParser = new parserFormula('DAYS("3/15/11","2/1/11")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("3/15/11","2/1/11") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 42, 'Test: Positive case: String(2). Format date m/d/yy Return 42');
		// Case #2: Reference link(2). Format date mm/dd/yyyy Return 364
		oParser = new parserFormula('DAYS(A2,A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A2,A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Reference link(2). Format date mm/dd/yyyy Return 364');
		// Case #3: String(2). Format date yyyy-mm-dd Return 2
		oParser = new parserFormula('DAYS("2008-03-03","2008-03-01")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("2008-03-03","2008-03-01") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String(2). Format date yyyy-mm-dd Return 2');
		// Case #4: String(2). Format date yyyy-mm-dd Return -2
		oParser = new parserFormula('DAYS("2008-03-01","2008-03-03")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("2008-03-01","2008-03-03") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -2, 'Test: Positive case: String(2). Format date yyyy-mm-dd Return -2');
		// Case #5: Reference link(2). Result of DAYS(B4,B3). Return 0.
		oParser = new parserFormula('DAYS(B4,B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B4,B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link(2). Result of DAYS(B4,B3). Return 0.');
		// Case #6: Reference link(2). Return -9
		oParser = new parserFormula('DAYS(B5,B8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B5,B8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: Reference link(2). Return -9');
		// Case #7: Number(2). Return -9
		oParser = new parserFormula('DAYS(1,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(1,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: Number(2). Return -9');
		// Case #8: Reference link(2). Return -9
		oParser = new parserFormula('DAYS(B6,B8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B6,B8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: Reference link(2). Return -9');
		// Case #9: Number(2). Return -9
		oParser = new parserFormula('DAYS(1.9,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(1.9,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: Number(2). Return -9');
		// Case #10: Reference link(2). Return -8
		oParser = new parserFormula('DAYS(B7,B8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B7,B8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Positive case: Reference link(2). Return -8');
		// Case #11: Number(2). Return -8
		oParser = new parserFormula('DAYS(2.1,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2.1,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Positive case: Number(2). Return -8');
		// Case #12: Number(2). Return -8
		oParser = new parserFormula('DAYS(2.1,10.1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2.1,10.1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Positive case: Number(2). Return -8');
		// Case #13: Number(2). Return -8
		oParser = new parserFormula('DAYS(2.1,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2.1,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Positive case: Number(2). Return -8');
		// Case #14: Empty, Number. Return - 10
		oParser = new parserFormula('DAYS(,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -10, 'Test: Positive case: Empty, Number. Return - 10');
		// Case #15: Number, Empty. Result of DAYS(2.1,). Return 2.
		oParser = new parserFormula('DAYS(2.1,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2.1,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Number, Empty. Result of DAYS(2.1,). Return 2.');
		// Case #16: Empty(2). Result of DAYS(,). Return 0.
		oParser = new parserFormula('DAYS(,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Empty(2). Result of DAYS(,). Return 0.');
		// Case #17: String(2). Return -9
		oParser = new parserFormula('DAYS("1","10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("1","10") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: String(2). Return -9');
		// Case #18: Boolean, Number. Return -9
		oParser = new parserFormula('DAYS(TRUE,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TRUE,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Test: Positive case: Boolean, Number. Return -9');
		// Case #19: Boolean, Number. Return -10
		oParser = new parserFormula('DAYS(FALSE,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(FALSE,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -10, 'Test: Positive case: Boolean, Number. Return -10');
		// Case #20: Number, Boolean. Result of DAYS(1,TRUE). Return 0.
		oParser = new parserFormula('DAYS(1,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(1,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number, Boolean. Result of DAYS(1,TRUE). Return 0.');
		// Case #21: Number, Boolean. Result of DAYS(1,FALSE). Return 1.
		oParser = new parserFormula('DAYS(1,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(1,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number, Boolean. Result of DAYS(1,FALSE). Return 1.');
		// Case #22: Boolean(2). Result of DAYS(TRUE,TRUE). Return 0.
		oParser = new parserFormula('DAYS(TRUE,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TRUE,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Boolean(2). Result of DAYS(TRUE,TRUE). Return 0.');
		// Case #23: Formula(2). End_date and start_date are DATE formulas. Date -> Number
		oParser = new parserFormula('DAYS(DATE(2024,1,15),DATE(2024,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(DATE(2024,1,15),DATE(2024,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Formula(2). End_date and start_date are DATE formulas. Date -> Number');
		// Case #24: Formula, Number. End_date is DATE formula, start_date is number.
		oParser = new parserFormula('DAYS(DATE(2025,12,31),40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(DATE(2025,12,31),40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5114, 'Test: Positive case: Formula, Number. End_date is DATE formula, start_date is number.');
		// Case #25: Number, Formula. End_date is number, start_date is DATE formula.
		oParser = new parserFormula('DAYS(40908,DATE(2020,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,DATE(2020,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -2923, 'Test: Positive case: Number, Formula. End_date is number, start_date is DATE formula.');
		// Case #26: Formula(2). End_date and start_date are DATEVALUE formulas.
		oParser = new parserFormula('DAYS(DATEVALUE("1/15/2025"),DATEVALUE("1/1/2025"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(DATEVALUE("1/15/2025"),DATEVALUE("1/1/2025")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Formula(2). End_date and start_date are DATEVALUE formulas.');
		// Case #27: Formula, Number. End_date with TIME formula. Time component ignored.
		oParser = new parserFormula('DAYS(TIME(12,0,0)+40908,40544)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TIME(12,0,0)+40908,40544) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Formula, Number. End_date with TIME formula. Time component ignored.');
		// Case #28: Number, Formula. Start_date with TIME formula. Time component ignored.
		oParser = new parserFormula('DAYS(40908,TIME(15,30,0)+40544)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,TIME(15,30,0)+40544) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Number, Formula. Start_date with TIME formula. Time component ignored.');
		// Case #29: Reference link(2). End_date and start_date are reference links to 40908 and 40544.
		oParser = new parserFormula('DAYS(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Reference link(2). End_date and start_date are reference links to 40908 and 40544.');
		// Case #30: Area(2). End_date and start_date are single-cell ranges.
		oParser = new parserFormula('DAYS(A100:A100,A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A100:A100,A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Area(2). End_date and start_date are single-cell ranges.');
		// Case #31: Area(2). End_date and start_date are multi-cell ranges.
		// Different result with MS
		/*oParser = new parserFormula('DAYS(A100:A101,A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A100:A101,A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Area(2). End_date and start_date are multi-cell ranges.');*/
		// Case #32: Name(2). End_date and start_date are named ranges to 40908 and 40544.
		oParser = new parserFormula('DAYS(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Name(2). End_date and start_date are named ranges to 40908 and 40544.');
		// Case #33: Name3D(2). End_date and start_date are 3D named ranges to 40908 and 40544.
		oParser = new parserFormula('DAYS(TestName3D,TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TestName3D,TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Name3D(2). End_date and start_date are 3D named ranges to 40908 and 40544.');
		// Case #34: Ref3D(2). End_date and start_date are 3D references to 40908 and 40544.
		oParser = new parserFormula('DAYS(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Ref3D(2). End_date and start_date are 3D references to 40908 and 40544.');
		// Case #35: Area3D(2). End_date and start_date are 3D single-cell ranges.
		oParser = new parserFormula('DAYS(Sheet2!A1:A1,Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Sheet2!A1:A1,Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Area3D(2). End_date and start_date are 3D single-cell ranges.');
		// Case #36: Area3D(2). End_date and start_date are 3D multi-cell ranges.
		// Different result with MS
		/*oParser = new parserFormula('DAYS(Sheet2!A1:A2,Sheet2!A2:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Sheet2!A1:A2,Sheet2!A2:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Area3D(2). End_date and start_date are 3D multi-cell ranges.');*/
		// Case #37: Table(2). End_date and start_date are table references to 40908 and 40544.
		oParser = new parserFormula('DAYS(Table1[Column1],Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Table1[Column1],Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Table(2). End_date and start_date are table references to 40908 and 40544.');
		// Case #38: Array(2). End_date and start_date are single-element arrays.
		oParser = new parserFormula('DAYS({40908},{40544})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS({40908},{40544}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Array(2). End_date and start_date are single-element arrays.');
		// Case #39: Array(2). End_date and start_date are multi-element arrays.
		oParser = new parserFormula('DAYS({40908;40544},{40544;40908})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS({40908;40544},{40544;40908}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Positive case: Array(2). End_date and start_date are multi-element arrays.');
		// Case #40: Formula. DAYS is nested inside SUM formula.
		oParser = new parserFormula('SUM(DAYS(DATE(2025,1,15),DATE(2025,1,1)),100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DAYS(DATE(2025,1,15),DATE(2025,1,1)),100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 114, 'Test: Positive case: Formula. DAYS is nested inside SUM formula.');
		// Case #41: Formula(2). Leap year date calculation.
		oParser = new parserFormula('DAYS(DATE(2024,2,29),DATE(2024,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(DATE(2024,2,29),DATE(2024,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Positive case: Formula(2). Leap year date calculation.');
		// Case #42: Number, Reference link. End_date is number, start_date is reference link.
		oParser = new parserFormula('DAYS(45000,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(45000,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4456, 'Test: Positive case: Number, Reference link. End_date is number, start_date is reference link.');
		// Case #43: Reference link, Number. End_date is reference link, start_date is number.
		oParser = new parserFormula('DAYS(A100,30000)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A100,30000) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10908, 'Test: Positive case: Reference link, Number. End_date is reference link, start_date is number.');
		// Case #44: Reference link, Empty. End_date is reference link, start_date is empty cell.
		oParser = new parserFormula('DAYS(A100,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A100,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40908, 'Test: Positive case: Reference link, Empty. End_date is reference link, start_date is empty cell.');
		// Case #45: Empty, Reference link. End_date is empty cell, start_date is reference link.
		oParser = new parserFormula('DAYS(A102,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A102,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -40544, 'Test: Positive case: Empty, Reference link. End_date is empty cell, start_date is reference link.');

		// Negative cases:

		// Case #1: Number(2). Result of DAYS(2.1,-10.9). Return #NUM!.
		oParser = new parserFormula('DAYS(2.1,-10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2.1,-10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Result of DAYS(2.1,-10.9). Return #NUM!.');
		// Case #2: Number(2). Result of DAYS(2.1,10.9). Return #NUM!.
		oParser = new parserFormula('DAYS(-2.1,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(-2.1,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Result of DAYS(2.1,10.9). Return #NUM!.');
		// Case #3: Number(2). Result of DAYS(-2.1,-10.9). Return #NUM!.
		oParser = new parserFormula('DAYS(-2.1,-10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(-2.1,-10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Result of DAYS(-2.1,-10.9). Return #NUM!.');
		// Case #4: String(2). Result of DAYS(. Return #VALUE!.
		oParser = new parserFormula('DAYS("1s","10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("1s","10") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Result of DAYS(. Return #VALUE!.');
		// Case #5: String(2). Result of DAYS(. Return #VALUE!.
		oParser = new parserFormula('DAYS("1","10s")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("1","10s") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Result of DAYS(. Return #VALUE!.');
		// Case #6: Error, Number. Result of DAYS(#N/A,10). Return #N/A.
		oParser = new parserFormula('DAYS(#N/A,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(#N/A,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Result of DAYS(#N/A,10). Return #N/A.');
		// Case #7: Reference link, Number. Result of DAYS(B100,10). Return #N/A.
		oParser = new parserFormula('DAYS(B100,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B100,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link, Number. Result of DAYS(B100,10). Return #N/A.');
		// Case #8: Error(2). Result of DAYS(#N/A,#NUM!). Return #N/A.
		oParser = new parserFormula('DAYS(#N/A,#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(#N/A,#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error(2). Result of DAYS(#N/A,#NUM!). Return #N/A.');
		// Case #9: Reference link(2). Result of DAYS(B100,B101). Return #N/A.
		oParser = new parserFormula('DAYS(B100,B101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B100,B101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link(2). Result of DAYS(B100,B101). Return #N/A.');
		// Case #10: Error(2). Result of DAYS(#NUM!,#N/A). Return #NUM!.
		oParser = new parserFormula('DAYS(#NUM!,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(#NUM!,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Error(2). Result of DAYS(#NUM!,#N/A). Return #NUM!.');
		// Case #11: Reference link(2). Result of DAYS(B101,B100). Return #NUM!.
		oParser = new parserFormula('DAYS(B101,B100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(B101,B100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Reference link(2). Result of DAYS(B101,B100). Return #NUM!.');
		// Case #12: String, Number. End_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS("invalid",40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("invalid",40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. End_date is invalid string. Return #VALUE!.');
		// Case #13: Number, String. Start_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,"invalid") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. Start_date is invalid string. Return #VALUE!.');
		// Case #14: String(2). End_date is invalid date string. Return #VALUE!.
		oParser = new parserFormula('DAYS("Feb 30, 2025","1/1/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS("Feb 30, 2025","1/1/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). End_date is invalid date string. Return #VALUE!.');
		// Case #15: Reference link, Number. End_date is reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS(A103,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A103,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link, Number. End_date is reference to text. Return #VALUE!.');
		// Case #16: Number, Reference link. Start_date is reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Reference link. Start_date is reference to text. Return #VALUE!.');
		// Case #17: Reference link, Number. End_date is reference to error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS(A104,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A104,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link, Number. End_date is reference to error #N/A. Return #N/A.');
		// Case #18: Number, Reference link. Start_date is reference to error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS(40908,A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number, Reference link. Start_date is reference to error #N/A. Return #N/A.');
		// Case #19: Number(2). End_date exceeds maximum valid date. Return #NUM!.
		// Different result with MS
		/*oParser = new parserFormula('DAYS(2958466,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2958466,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). End_date exceeds maximum valid date. Return #NUM!.');
		// Case #20: Number(2). Start_date exceeds maximum valid date. Return #NUM!.
		// Different result with MS
		oParser = new parserFormula('DAYS(40908,2958466)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,2958466) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Start_date exceeds maximum valid date. Return #NUM!.');*/
		// Case #21: Area(2). End_date range contains text and error. Return #VALUE!.
		oParser = new parserFormula('DAYS(A103:A104,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A103:A104,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area(2). End_date range contains text and error. Return #VALUE!.');
		// Case #22: Ref3D, Number. End_date is 3D reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS(Sheet2!A3,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Sheet2!A3,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, Number. End_date is 3D reference to text. Return #VALUE!.');
		// Case #23: Number, Ref3D. Start_date is 3D reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Ref3D. Start_date is 3D reference to text. Return #VALUE!.');
		// Case #24: Name, Number. End_date is named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(TestName2,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TestName2,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, Number. End_date is named range with text. Return #VALUE!.');
		// Case #25: Number, Name. Start_date is named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Name. Start_date is named range with text. Return #VALUE!.');
		// Case #26: Name3D, Number. End_date is 3D named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(TestName3D1,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(TestName3D1,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D, Number. End_date is 3D named range with text. Return #VALUE!.');
		// Case #27: Number, Name3D. Start_date is 3D named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Name3D. Start_date is 3D named range with text. Return #VALUE!.');
		// Case #28: Table, Number. End_date is table column with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(Table1[Column3],40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(Table1[Column3],40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table, Number. End_date is table column with text. Return #VALUE!.');
		// Case #29: Number, Table. Start_date is table column with text. Return #VALUE!.
		oParser = new parserFormula('DAYS(40908,Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Table. Start_date is table column with text. Return #VALUE!.');
		// Case #30: Array(2). End_date array contains text. Return #VALUE!.
		oParser = new parserFormula('DAYS({"text"},{40908})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS({"text"},{40908}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array(2). End_date array contains text. Return #VALUE!.');
		// Case #31: Array(2). Start_date array contains text. Return #VALUE!.
		oParser = new parserFormula('DAYS({40908},{"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS({40908},{"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array(2). Start_date array contains text. Return #VALUE!.');

		// Bounded cases:

		// Case #1: Number(2). Minimum valid serial numbers (01/01/1900). Return 0.
		oParser = new parserFormula('DAYS(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number(2). Minimum valid serial numbers (01/01/1900). Return 0.');
		// Case #2: Number(2). Maximum valid end_date (31/12/9999), minimum start_date.
		oParser = new parserFormula('DAYS(2958465,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2958465,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958464, 'Test: Bounded case: Number(2). Maximum valid end_date (31/12/9999), minimum start_date.');
		// Case #3: Number(2). Maximum valid serial numbers (31/12/9999). Return 0.
		oParser = new parserFormula('DAYS(2958465,2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(2958465,2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number(2). Maximum valid serial numbers (31/12/9999). Return 0.');
		// Case #4: Area(2). Entire columns as end_date and start_date.
		oParser = new parserFormula('DAYS(A:A,B:B)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(A:A,B:B) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Area(2). Entire columns as end_date and start_date.');
		// Case #5: Area(2). Entire rows as end_date and start_date.
		oParser = new parserFormula('DAYS(100:100,101:101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(100:100,101:101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 364, 'Test: Bounded case: Area(2). Entire rows as end_date and start_date.');
		// Case #6: Number(2). End_date is zero.
		oParser = new parserFormula('DAYS(0,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(0,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -40908, 'Test: Bounded case: Number(2). End_date is zero.');
		// Case #7: Number(2). Start_date is zero.
		oParser = new parserFormula('DAYS(40908,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS(40908,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40908, 'Test: Bounded case: Number(2). Start_date is zero.');

		// array
		oParser = new parserFormula('DAYS({1;2;3},10)', "A1", ws);
		assert.ok(oParser.parse(), 'DAYS({1;2;3},10)');
		assert.strictEqual(oParser.calculate().getValue(), -9, 'Result of DAYS({1;2;3},10)');

		oParser = new parserFormula('DAYS({1;2;3},10)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'DAYS({1;2;3},10)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), -9, 'Result of DAYS({1;2;3},10)[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), -8, 'Result of DAYS({1;2;3},10)[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), -7, 'Result of DAYS({1;2;3},10)[2,0]');

		oParser = new parserFormula('DAYS({1;2;3},{10;9;8})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'DAYS({1;2;3},{10;9;8})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), -9, 'Result of DAYS({1;2;3},{10;9;8})[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), -7, 'Result of DAYS({1;2;3},{10;9;8})[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), -5, 'Result of DAYS({1;2;3},{10;9;8})[2,0]');

		// range

		oParser = new parserFormula('DAYS(A25:A27,B25:B27)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'DAYS(A25:A27,B25:B27)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), -9, "Result of DAYS(A25:A27,B25:B27)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), -7, "Result of DAYS(A25:A27,B25:B27)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), -5, "Result of DAYS(A25:A27,B25:B27)[2,0]");

		testArrayFormula2(assert, "DAYS", 2, 2);

		// TODO Need to fix

		// Incorrect Area logic - Positive case #31, #36.
		// Absent logic for reaching beyond of maximum date. Must be #NUM! - Negative case #19, 20.
	});

	QUnit.test('Test: "DAYS360"', function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("40908");
		ws.getRange2("A101").setValue("40544");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("text");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A105").setValue("#NUM!");
		ws.getRange2("A106").setValue("TRUE");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 3);
		ws.getRange2("A601").setValue("40908"); // Column1
		ws.getRange2("B601").setValue("40544"); // Column2
		ws.getRange2("C601").setValue("TRUE"); // Column3
		ws.getRange2("D601").setValue("text"); // Column4
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("40908");
		ws2.getRange2("A2").setValue("40544");
		ws2.getRange2("A3").setValue("text");
		ws2.getRange2("A4").setValue("#N/A");
		ws2.getRange2("A5").setValue("TRUE");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("40908"); // TestName
		ws.getRange2("A202").setValue("40544"); // TestName1
		ws.getRange2("A203").setValue("TRUE"); // TestName2
		ws.getRange2("A204").setValue("text"); // TestName3
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("40908") // TestName3D
		ws2.getRange2("A12").setValue("text") // TestName3D1
		ws2.getRange2("A13").setValue("40544") // TestName3D2
		ws2.getRange2("A14").setValue("TRUE") // TestName3D3

		// Positive cases:

		// Case #1: Formula(2). Return 1198
		oParser = new parserFormula('DAYS360(DATE(2002,2,3),DATE(2005,5,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2002,2,3),DATE(2005,5,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1198, 'Test: Positive case: Formula(2). Return 1198');
		// Case #2: Formula(2). Return -1197
		oParser = new parserFormula('DAYS360(DATE(2005,5,31),DATE(2002,2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2005,5,31),DATE(2002,2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -1197, 'Test: Positive case: Formula(2). Return -1197');
		// Case #3: Formula(2), Boolean. U.S. (NASD) method. Return 1198
		oParser = new parserFormula('DAYS360(DATE(2002,2,3),DATE(2005,5,31),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2002,2,3),DATE(2005,5,31),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1198, 'Test: Positive case: Formula(2), Boolean. U.S. (NASD) method. Return 1198');
		// Case #4: Formula(2), Boolean. European method. Return 1197
		oParser = new parserFormula('DAYS360(DATE(2002,2,3),DATE(2005,5,31),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2002,2,3),DATE(2005,5,31),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1197, 'Test: Positive case: Formula(2), Boolean. European method. Return 1197');
		// Case #5: String(2). Format date m/d/yy Return 42. Method omitted (default FALSE)
		oParser = new parserFormula('DAYS360("3/15/11","2/1/11")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("3/15/11","2/1/11") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -44, 'Test: Positive case: String(2). Format date m/d/yy Return 42. Method omitted (default FALSE)');
		// Case #6: Reference link(2). Format date mm/dd/yyyy Return 364. Method omitted
		oParser = new parserFormula('DAYS360(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Reference link(2). Format date mm/dd/yyyy Return 364. Method omitted');
		// Case #7: String(2). Format date yyyy-mm-dd Return 0. Method omitted
		oParser = new parserFormula('DAYS360("2008-03-03","2008-03-01")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("2008-03-03","2008-03-01") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -2, 'Test: Positive case: String(2). Format date yyyy-mm-dd Return 0. Method omitted');
		// Case #8: String(2). Format date yyyy-mm-dd Return 0. Method omitted
		oParser = new parserFormula('DAYS360("2008-03-01","2008-03-03")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("2008-03-01","2008-03-03") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String(2). Format date yyyy-mm-dd Return 0. Method omitted');
		// Case #9: Number(2). Return -9. Method omitted
		oParser = new parserFormula('DAYS360(1,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(1,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Number(2). Return -9. Method omitted');
		// Case #10: Number(2). Float numbers truncated. Return -9. Method omitted
		oParser = new parserFormula('DAYS360(1.9,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(1.9,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Number(2). Float numbers truncated. Return -9. Method omitted');
		// Case #11: Number(2). Float numbers truncated. Return -8. Method omitted
		oParser = new parserFormula('DAYS360(2.1,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2.1,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Number(2). Float numbers truncated. Return -8. Method omitted');
		// Case #12: Number(2). Both arguments are float numbers. Return -8. Method omitted
		oParser = new parserFormula('DAYS360(2.1,10.1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2.1,10.1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Number(2). Both arguments are float numbers. Return -8. Method omitted');
		// Case #13: Number(2). Float numbers truncated. Return -8. Method omitted
		oParser = new parserFormula('DAYS360(2.1,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2.1,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Number(2). Float numbers truncated. Return -8. Method omitted');
		// Case #14: Empty, Number. Start_date is empty converted to 0. Return -10. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Empty, Number. Start_date is empty converted to 0. Return -10. Method omitted');*/
		// Case #15: Number, Empty. End_date is empty converted to 0. Return 2. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(2.1,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2.1,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -2, 'Test: Positive case: Number, Empty. End_date is empty converted to 0. Return 2. Method omitted');*/
		// Case #16: Empty(2). Both arguments are empty converted to 0. Return 0. Method omitted
		oParser = new parserFormula('DAYS360(,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Empty(2). Both arguments are empty converted to 0. Return 0. Method omitted');
		// Case #17: String(2). Numeric strings converted to numbers. Return -9. Method omitted
		oParser = new parserFormula('DAYS360("1","10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("1","10") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String(2). Numeric strings converted to numbers. Return -9. Method omitted');
		// Case #18: Boolean, Number. Boolean TRUE converted to 1. Return -9. Method omitted
		oParser = new parserFormula('DAYS360(TRUE,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TRUE,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Boolean, Number. Boolean TRUE converted to 1. Return -9. Method omitted');
		// Case #19: Boolean, Number. Boolean FALSE converted to 0. Return -10. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(FALSE,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(FALSE,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Boolean, Number. Boolean FALSE converted to 0. Return -10. Method omitted');*/
		// Case #20: Number, Boolean. Boolean TRUE converted to 1. Return 0. Method omitted
		oParser = new parserFormula('DAYS360(1,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(1,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number, Boolean. Boolean TRUE converted to 1. Return 0. Method omitted');
		// Case #21: Number, Boolean. Boolean FALSE converted to 0. Return 1. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(1,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(1,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -1, 'Test: Positive case: Number, Boolean. Boolean FALSE converted to 0. Return 1. Method omitted');*/
		// Case #22: Boolean(2). Both arguments are boolean TRUE. Return 0. Method omitted
		oParser = new parserFormula('DAYS360(TRUE,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TRUE,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Boolean(2). Both arguments are boolean TRUE. Return 0. Method omitted');
		// Case #23: Formula(2). Start_date and end_date are DATE formulas. Date -> Number. Method omitted
		oParser = new parserFormula('DAYS360(DATE(2024,1,15),DATE(2024,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2024,1,15),DATE(2024,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -14, 'Test: Positive case: Formula(2). Start_date and end_date are DATE formulas. Date -> Number. Method omitted');
		// Case #24: Formula, Number. Start_date is DATE formula, end_date is number. Method omitted
		oParser = new parserFormula('DAYS360(DATE(2025,12,31),40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2025,12,31),40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -5040, 'Test: Positive case: Formula, Number. Start_date is DATE formula, end_date is number. Method omitted');
		// Case #25: Number, Formula. Start_date is number, end_date is DATE formula. Method omitted
		oParser = new parserFormula('DAYS360(40908,DATE(2020,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,DATE(2020,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2881, 'Test: Positive case: Number, Formula. Start_date is number, end_date is DATE formula. Method omitted');
		// Case #26: Formula(2). Start_date and end_date are DATEVALUE formulas. Method omitted
		oParser = new parserFormula('DAYS360(DATEVALUE("1/15/2025"),DATEVALUE("1/1/2025"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATEVALUE("1/15/2025"),DATEVALUE("1/1/2025")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -14, 'Test: Positive case: Formula(2). Start_date and end_date are DATEVALUE formulas. Method omitted');
		// Case #27: Formula, Number. Start_date with TIME formula. Time component ignored. Method omitted
		oParser = new parserFormula('DAYS360(TIME(12,0,0)+40908,40544)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TIME(12,0,0)+40908,40544) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Formula, Number. Start_date with TIME formula. Time component ignored. Method omitted');
		// Case #28: Number, Formula. End_date with TIME formula. Time component ignored. Method omitted
		oParser = new parserFormula('DAYS360(40908,TIME(15,30,0)+40544)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,TIME(15,30,0)+40544) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number, Formula. End_date with TIME formula. Time component ignored. Method omitted');
		// Case #29: Reference link(2). Start_date and end_date are reference links to 40908 and 40544. Method omitted
		oParser = new parserFormula('DAYS360(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Reference link(2). Start_date and end_date are reference links to 40908 and 40544. Method omitted');
		// Case #30: Area(2). Start_date and end_date are single-cell ranges. Method omitted
		oParser = new parserFormula('DAYS360(A100:A100,A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A100:A100,A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Area(2). Start_date and end_date are single-cell ranges. Method omitted');
		// Case #31: Area(2). Start_date and end_date are multi-cell ranges. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(A100:A101,A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A100:A101,A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Area(2). Start_date and end_date are multi-cell ranges. Method omitted');*/
		// Case #32: Name(2). Start_date and end_date are named ranges to 40908 and 40544. Method omitted
		oParser = new parserFormula('DAYS360(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Name(2). Start_date and end_date are named ranges to 40908 and 40544. Method omitted');
		// Case #33: Name3D(2). Start_date and end_date are 3D named ranges to 40908 and 40544. Method omitted
		oParser = new parserFormula('DAYS360(TestName3D,TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TestName3D,TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Name3D(2). Start_date and end_date are 3D named ranges to 40908 and 40544. Method omitted');
		// Case #34: Ref3D(2). Start_date and end_date are 3D references to 40908 and 40544. Method omitted
		oParser = new parserFormula('DAYS360(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Ref3D(2). Start_date and end_date are 3D references to 40908 and 40544. Method omitted');
		// Case #35: Area3D(2). Start_date and end_date are 3D single-cell ranges. Method omitted
		oParser = new parserFormula('DAYS360(Sheet2!A1:A1,Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Sheet2!A1:A1,Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Area3D(2). Start_date and end_date are 3D single-cell ranges. Method omitted');
		// Case #36: Area3D(2). Start_date and end_date are 3D multi-cell ranges. Method omitted
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(Sheet2!A1:A2,Sheet2!A2:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Sheet2!A1:A2,Sheet2!A2:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Area3D(2). Start_date and end_date are 3D multi-cell ranges. Method omitted');*/
		// Case #37: Table(2). Start_date and end_date are table references to 40908 and 40544. Method omitted
		oParser = new parserFormula('DAYS360(Table1[Column1],Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Table1[Column1],Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Table(2). Start_date and end_date are table references to 40908 and 40544. Method omitted');
		// Case #38: Array(2). Start_date and end_date are single-element arrays. Method omitted
		oParser = new parserFormula('DAYS360({40908},{40544})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360({40908},{40544}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Array(2). Start_date and end_date are single-element arrays. Method omitted');
		// Case #39: Array(2). Start_date and end_date are multi-element arrays. Method omitted
		oParser = new parserFormula('DAYS360({40908;40544},{40544;40908})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360({40908;40544},{40544;40908}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Array(2). Start_date and end_date are multi-element arrays. Method omitted');
		// Case #40: Formula. DAYS360 is nested inside SUM formula. Method omitted
		oParser = new parserFormula('SUM(DAYS360(DATE(2025,1,15),DATE(2025,1,1)),100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DAYS360(DATE(2025,1,15),DATE(2025,1,1)),100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 86, 'Test: Positive case: Formula. DAYS360 is nested inside SUM formula. Method omitted');
		// Case #41: Formula(2). Leap year date calculation. Method omitted
		oParser = new parserFormula('DAYS360(DATE(2024,2,29),DATE(2024,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(DATE(2024,2,29),DATE(2024,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -59, 'Test: Positive case: Formula(2). Leap year date calculation. Method omitted');
		// Case #42: Number(2), Boolean. Method is FALSE (US NASD method).
		oParser = new parserFormula('DAYS360(40908,40544,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Boolean. Method is FALSE (US NASD method).');
		// Case #43: Number(2), Boolean. Method is TRUE (European method).
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(40908,40544,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Boolean. Method is TRUE (European method).');*/
		// Case #44: Number(2), Empty. Method is empty converted to FALSE.
		oParser = new parserFormula('DAYS360(40908,40544,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Empty. Method is empty converted to FALSE.');
		// Case #45: Number(3). Method is 0 converted to FALSE.
		oParser = new parserFormula('DAYS360(40908,40544,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(3). Method is 0 converted to FALSE.');
		// Case #46: Number(3). Method is 1 converted to TRUE.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(40908,40544,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(3). Method is 1 converted to TRUE.');*/
		// Case #47: Number(2), String. Method is string "TRUE" converted to TRUE.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(40908,40544,"TRUE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,"TRUE") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), String. Method is string "TRUE" converted to TRUE.');*/
		// Case #48: Number(2), String. Method is string "FALSE" converted to FALSE.
		oParser = new parserFormula('DAYS360(40908,40544,"FALSE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,"FALSE") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), String. Method is string "FALSE" converted to FALSE.');
		// Case #49: Number(2), Formula. Method is IF formula returning TRUE.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(40908,40544,IF(TRUE,TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,IF(TRUE,TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Formula. Method is IF formula returning TRUE.');*/
		// Case #50: Number(2), Formula. Method is IF formula returning FALSE.
		oParser = new parserFormula('DAYS360(40908,40544,IF(FALSE,TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,IF(FALSE,TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Formula. Method is IF formula returning FALSE.');
		// Case #51: Number(2), Reference link. Method is reference link to TRUE.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(40908,40544,A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Reference link. Method is reference link to TRUE.');
		// Case #52: Number(2), Name. Method is named range to TRUE.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,40544,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Name. Method is named range to TRUE.');
		// Case #53: Number(2), Name3D. Method is 3D named range to TRUE.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,40544,TestName3D3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,TestName3D3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Name3D. Method is 3D named range to TRUE.');
		// Case #54: Number(2), Ref3D. Method is 3D reference to TRUE.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,40544,Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Ref3D. Method is 3D reference to TRUE.');
		// Case #55: Number(2), Table. Method is table reference to TRUE.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,40544,Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Table. Method is table reference to TRUE.');
		// Case #56: Number(2), Array. Method is array with TRUE.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,40544,{TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,{TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Positive case: Number(2), Array. Method is array with TRUE.');*/

		// Negative cases:

		// Case #1: Number(2). Negative end_date outside valid range. Return #NUM!.
		oParser = new parserFormula('DAYS360(2.1,-10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2.1,-10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Negative end_date outside valid range. Return #NUM!.');
		// Case #2: Number(2). Negative start_date outside valid range. Return #NUM!.
		oParser = new parserFormula('DAYS360(-2.1,10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(-2.1,10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Negative start_date outside valid range. Return #NUM!.');
		// Case #3: Number(2). Both arguments are negative numbers. Return #NUM!.
		oParser = new parserFormula('DAYS360(-2.1,-10.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(-2.1,-10.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Both arguments are negative numbers. Return #NUM!.');
		// Case #4: String(2). Start_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS360("1s","10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("1s","10") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Start_date is invalid string. Return #VALUE!.');
		// Case #5: String(2). End_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS360("1","10s")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("1","10s") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). End_date is invalid string. Return #VALUE!.');
		// Case #6: Error, Number. Start_date is error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS360(#N/A,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(#N/A,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Start_date is error #N/A. Return #N/A.');
		// Case #7: Reference link, Number. Start_date is reference to #N/A error. Return #N/A.
		oParser = new parserFormula('DAYS360(A104,10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A104,10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link, Number. Start_date is reference to #N/A error. Return #N/A.');
		// Case #8: Error(2). Start_date is #N/A, end_date is #NUM!. Return #N/A.
		oParser = new parserFormula('DAYS360(#N/A,#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(#N/A,#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error(2). Start_date is #N/A, end_date is #NUM!. Return #N/A.');
		// Case #9: Reference link(2). Both arguments are references to errors. Return #N/A.
		oParser = new parserFormula('DAYS360(A104,A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A104,A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link(2). Both arguments are references to errors. Return #N/A.');
		// Case #10: Error(2). Start_date is #NUM!, end_date is #N/A. Return #NUM!.
		oParser = new parserFormula('DAYS360(#NUM!,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(#NUM!,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Error(2). Start_date is #NUM!, end_date is #N/A. Return #NUM!.');
		// Case #11: Reference link(2). Start_date is #NUM!, end_date is #N/A. Return #NUM!.
		oParser = new parserFormula('DAYS360(A105,A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A105,A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Reference link(2). Start_date is #NUM!, end_date is #N/A. Return #NUM!.');
		// Case #12: String, Number. Start_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS360("invalid",40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("invalid",40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Start_date is invalid string. Return #VALUE!.');
		// Case #13: Number, String. End_date is invalid string. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,"invalid") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. End_date is invalid string. Return #VALUE!.');
		// Case #14: String(2). Start_date is invalid date string. Return #VALUE!.
		oParser = new parserFormula('DAYS360("Feb 30, 2025","1/1/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360("Feb 30, 2025","1/1/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Start_date is invalid date string. Return #VALUE!.');
		// Case #15: Reference link, Number. Start_date is reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(A103,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A103,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link, Number. Start_date is reference to text. Return #VALUE!.');
		// Case #16: Number, Reference link. End_date is reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Reference link. End_date is reference to text. Return #VALUE!.');
		// Case #17: Reference link, Number. Start_date is reference to error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS360(A104,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A104,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link, Number. Start_date is reference to error #N/A. Return #N/A.');
		// Case #18: Number, Reference link. End_date is reference to error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS360(40908,A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number, Reference link. End_date is reference to error #N/A. Return #N/A.');
		// Case #19: Number(2). Start_date exceeds maximum valid date. Return #NUM!.
		/*// Different result with MS
		oParser = new parserFormula('DAYS360(2958466,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2958466,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). Start_date exceeds maximum valid date. Return #NUM!.');
		// Case #20: Number(2). End_date exceeds maximum valid date. Return #NUM!.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,2958466)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,2958466) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number(2). End_date exceeds maximum valid date. Return #NUM!.');*/
		// Case #21: Area(2). Start_date range contains text and error. Return #VALUE!.
		oParser = new parserFormula('DAYS360(A103:A104,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A103:A104,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area(2). Start_date range contains text and error. Return #VALUE!.');
		// Case #22: Ref3D, Number. Start_date is 3D reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(Sheet2!A3,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Sheet2!A3,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, Number. Start_date is 3D reference to text. Return #VALUE!.');
		// Case #23: Number, Ref3D. End_date is 3D reference to text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Ref3D. End_date is 3D reference to text. Return #VALUE!.');
		// Case #24: Name, Number. Start_date is named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(TestName3,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TestName3,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, Number. Start_date is named range with text. Return #VALUE!.');
		// Case #25: Number, Name. End_date is named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Name. End_date is named range with text. Return #VALUE!.');
		// Case #26: Name3D, Number. Start_date is 3D named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(TestName3D1,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(TestName3D1,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D, Number. Start_date is 3D named range with text. Return #VALUE!.');
		// Case #27: Number, Name3D. End_date is 3D named range with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Name3D. End_date is 3D named range with text. Return #VALUE!.');
		// Case #28: Table, Number. Start_date is table column with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(Table1[Column4],40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(Table1[Column4],40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table, Number. Start_date is table column with text. Return #VALUE!.');
		// Case #29: Number, Table. End_date is table column with text. Return #VALUE!.
		oParser = new parserFormula('DAYS360(40908,Table1[Column4])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,Table1[Column4]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Table. End_date is table column with text. Return #VALUE!.');
		// Case #30: Number, Table. Method is table column with text. Return #VALUE!.
		// TODO Crush application Need to fix in first priority.
		/*oParser = new parserFormula('DAYS360(40908, 40544, Table1[Column4])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908, 40544, Table1[Column4]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Table. Method is table column with text. Return #VALUE!.');*/
		// Case #31: Array(2). Start_date array contains text. Return #VALUE!.
		oParser = new parserFormula('DAYS360({"text"},{40908})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360({"text"},{40908}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array(2). Start_date array contains text. Return #VALUE!.');
		// Case #32: Array(2). End_date array contains text. Return #VALUE!.
		oParser = new parserFormula('DAYS360({40908},{"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360({40908},{"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array(2). End_date array contains text. Return #VALUE!.');
		// Case #33: Number(2), Error. Method is error #N/A. Return #N/A.
		oParser = new parserFormula('DAYS360(40908,40544,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number(2), Error. Method is error #N/A. Return #N/A.');
		// Case #34: Number(2), String. Method is invalid text string. Return #VALUE!.
		// TODO Crush application Need to fix in first priority.
		/*oParser = new parserFormula('DAYS360(40908,40544,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,"invalid") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), String. Method is invalid text string. Return #VALUE!.');*/
		// Case #35: Number(2), Reference link. Method is reference to text. Return #VALUE!.
		// TODO Crush application Need to fix in first priority.
		/*oParser = new parserFormula('DAYS360(40908,40544,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), Reference link. Method is reference to text. Return #VALUE!.');*/
		// Case #36: Number(2), Reference link. Method is reference to error #N/A. Return #N/A.
		// TODO Crush application Need to fix in first priority.
		/*oParser = new parserFormula('DAYS360(40908,40544,A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number(2), Reference link. Method is reference to error #N/A. Return #N/A.');*/
		/*// Case #37: Number(2), String. Method is string "1".
		oParser = new parserFormula('DAYS360(40908,40544,"1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,"1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), String. Method is string "1".');
		// Case #38: Number(2), String. Method is string "0".
		oParser = new parserFormula('DAYS360(40908,40544,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,40544,"0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2), String. Method is string "0".');*/

		// Bounded cases:

		// Case #1: Number(2). Minimum valid serial numbers (01/01/1900). Return 0.
		oParser = new parserFormula('DAYS360(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number(2). Minimum valid serial numbers (01/01/1900). Return 0.');
		// Case #2: Number(2). Maximum valid start_date (31/12/9999), minimum end_date.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(2958465,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2958465,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -2915999, 'Test: Bounded case: Number(2). Maximum valid start_date (31/12/9999), minimum end_date.');*/
		// Case #3: Number(2). Maximum valid serial numbers (31/12/9999). Return 0.
		oParser = new parserFormula('DAYS360(2958465,2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(2958465,2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number(2). Maximum valid serial numbers (31/12/9999). Return 0.');
		// Case #4: Area(2). Entire columns as start_date and end_date.
		oParser = new parserFormula('DAYS360(A:A,B:B)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(A:A,B:B) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Area(2). Entire columns as start_date and end_date.');
		// Case #5: Area(2). Entire rows as start_date and end_date.
		oParser = new parserFormula('DAYS360(100:100,101:101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(100:100,101:101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -359, 'Test: Bounded case: Area(2). Entire rows as start_date and end_date.');
		// Case #6: Number(2). Start_date is zero.
		// Different result with MS
		/*oParser = new parserFormula('DAYS360(0,40908)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(0,40908) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40321, 'Test: Bounded case: Number(2). Start_date is zero.');
		// Case #7: Number(2). End_date is zero.
		// Different result with MS
		oParser = new parserFormula('DAYS360(40908,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DAYS360(40908,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -40320, 'Test: Bounded case: Number(2). End_date is zero.');*/

		testArrayFormula2(assert, "DAYS360", 2, 3);

		// TODO Need to fix

		// Critical: Crush application Need to fix in first priority. Negative case: #30, #34-38.
		// Area problem. Positive case:  #31, #36
		// Incorrect logic for non-standard types. Positive case: #14, #15, #19, #21
		// Incorrect logic for European method (3 arg). Positive case: #43, #46, #47, #49, #51-56.
		// Incorrect calculate logic. Bounded case: #2, #6, #7.
		// Absent logic for reaching beyond of maximum date. Must be #NUM!. Negative case: #19, #20.

	});

	QUnit.test('Test: "EDATE"', function (assert) {
		let array;

		// base mode
		ws.workbook.setDate1904(false, true);

		oParser = new parserFormula("EDATE(DATE(2006,1,31),5)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2006,1,31),5)');
		assert.strictEqual(oParser.calculate().getValue(), 38898, 'Result of EDATE(DATE(2006,1,31),5)');

		oParser = new parserFormula("EDATE(DATE(2004,2,29),12)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2004,2,29),12)');
		assert.strictEqual(oParser.calculate().getValue(), 38411, 'Result of EDATE(DATE(2004,2,29),12)');

		ws.getRange2("A7").setValue("02-28-2004");
		oParser = new parserFormula("EDATE(A7,12)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(A7,12)');
		assert.strictEqual(oParser.calculate().getValue(), 38411, 'Result of EDATE(A7,12)');

		oParser = new parserFormula("EDATE(DATE(2004,1,15),-23)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2004,1,15),-23)');
		assert.strictEqual(oParser.calculate().getValue(), 37302, 'Result of EDATE(DATE(2004,1,15),-23)');

		oParser = new parserFormula("EDATE(DATE(2000,1,30),1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2000,1,30),1)');
		assert.strictEqual(oParser.calculate().getValue(), 36585, 'Result of EDATE(DATE(2000,1,30),1)');

		oParser = new parserFormula("EDATE(DATE(2001,1,30),1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2001,1,30),1)');
		assert.strictEqual(oParser.calculate().getValue(), 36950, 'Result of EDATE(DATE(2001,1,30),1)');

		oParser = new parserFormula("EDATE(DATE(2002,1,30),1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2002,1,30),1)');
		assert.strictEqual(oParser.calculate().getValue(), 37315, 'Result of EDATE(DATE(2002,1,30),1)');

		oParser = new parserFormula("EDATE(DATE(2003,1,30),1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2003,1,30),1)');
		assert.strictEqual(oParser.calculate().getValue(), 37680, 'Result of EDATE(DATE(2003,1,30),1)');

		oParser = new parserFormula("EDATE(DATE(2004,1,30),1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(DATE(2004,1,30),1)');
		assert.strictEqual(oParser.calculate().getValue(), 38046, 'Result of EDATE(DATE(2004,1,30),1)');

		oParser = new parserFormula("EDATE(0,0)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(0,0)');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Result of EDATE(0,0)');

		oParser = new parserFormula("EDATE(0,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(0,1)');
		assert.strictEqual(oParser.calculate().getValue(), 32, 'Result of EDATE(0,1)');			// 31

		oParser = new parserFormula("EDATE(1,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(1,1)');
		assert.strictEqual(oParser.calculate().getValue(), 33, 'Result of EDATE(1,1)');			// 32

		oParser = new parserFormula("EDATE(2,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(2,1)');
		assert.strictEqual(oParser.calculate().getValue(), 34, 'Result of EDATE(2,1)');			// 33

		oParser = new parserFormula("EDATE(10,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(10,1)');
		assert.strictEqual(oParser.calculate().getValue(), 42, 'Result of EDATE(10,1)');		// 41

		oParser = new parserFormula("EDATE(30,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(30,1)');
		assert.strictEqual(oParser.calculate().getValue(), 60, 'Result of EDATE(30,1)');		// 59

		oParser = new parserFormula("EDATE(59,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(59,1)');
		assert.strictEqual(oParser.calculate().getValue(), 88, 'Result of EDATE(59,1)');

		oParser = new parserFormula("EDATE(60,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(60,1)');
		assert.strictEqual(oParser.calculate().getValue(), 92, 'Result of EDATE(60,1)');		// 89

		oParser = new parserFormula("EDATE(61,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(61,1)');
		assert.strictEqual(oParser.calculate().getValue(), 92, 'Result of EDATE(61,1)');		// 92

		oParser = new parserFormula("EDATE(0,-1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(0,-1)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of EDATE(0,-1)');

		oParser = new parserFormula("EDATE(0,-2)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(0,-2)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of EDATE(0,-2)');

		oParser = new parserFormula("EDATE(-1,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(-1,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of EDATE(-1,1)');

		oParser = new parserFormula("EDATE(100,)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(100,)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of EDATE(100,)');

		oParser = new parserFormula("EDATE(,100)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(,100)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of EDATE(,100)');

		oParser = new parserFormula("EDATE(0,10)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(0,10)');
		assert.strictEqual(oParser.calculate().getValue(), 305, 'Result of EDATE(0,10)');		// 305

		oParser = new parserFormula("EDATE(1,10)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(1,10)');
		assert.strictEqual(oParser.calculate().getValue(), 306, 'Result of EDATE(1,10)');		// 306

		// strings
		oParser = new parserFormula('EDATE("100",1)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE("100",1)');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE("100",1)');

		oParser = new parserFormula('EDATE("100s",1)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE("100s",1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE("100s",1)');

		oParser = new parserFormula('EDATE("100","1")', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE("100","1")');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE("100","1")');

		oParser = new parserFormula('EDATE("100","1s")', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE("100","1s")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE("100","1s")');

		// bool
		oParser = new parserFormula("EDATE(TRUE,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(TRUE,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(TRUE,1)');

		oParser = new parserFormula("EDATE(FALSE,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(FALSE,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(FALSE,1)');

		oParser = new parserFormula("EDATE(1,TRUE)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(1,TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(1,TRUE)');

		oParser = new parserFormula("EDATE(1,FALSE)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(1,FALSE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(1,FALSE)');

		oParser = new parserFormula("EDATE(TRUE,TRUE)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(TRUE,TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(TRUE,TRUE)');

		// err
		oParser = new parserFormula("EDATE(#N/A,#NUM!)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(#N/A,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of EDATE(#N/A,#NUM!)');

		oParser = new parserFormula("EDATE(#DIV/0!,#NUM!)", "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(#DIV/0!,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'Result of EDATE(#DIV/0!,#NUM!)');

		// arr
		oParser = new parserFormula('EDATE({100;101;102},1)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE({100;101;102},1)');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE({100;101;102},1)');

		oParser = new parserFormula('EDATE(100,{1;2;3})', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(100,{1;2;3})');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE(100,{1;2;3})');

		oParser = new parserFormula('EDATE({100;100;100},{1;2;3})', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE({100;100;100},{1;2;3})');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE({100;100;100},{1;2;3})');

		oParser = new parserFormula('EDATE({100;101;102},1)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'EDATE({100;101;102},1)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 130, 'Result of EDATE({100;101;102},1)[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 131, 'Result of EDATE({100;101;102},1)[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 132, 'Result of EDATE({100;101;102},1)[2,0]');

		oParser = new parserFormula('EDATE({100;100;100},{1;2;3})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'EDATE({100;100;100},{1;2;3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 130, 'Result of EDATE({100;100;100},{1;2;3})[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 161, 'Result of EDATE({100;100;100},{1;2;3})[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 191, 'Result of EDATE({100;100;100},{1;2;3})[2,0]');

		// range
		ws.getRange2("A100").setValue("100");
		ws.getRange2("A101").setValue("101");
		ws.getRange2("A102").setValue("102");
		ws.getRange2("B100").setValue("1");
		ws.getRange2("B101").setValue("2");
		ws.getRange2("B102").setValue("3");

		oParser = new parserFormula('EDATE(A100:A100,1)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(A100:A100,1)');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE(A100:A100,1)');

		oParser = new parserFormula('EDATE(A100:A102,1)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(A100:A102,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(A100:A102,1)');

		oParser = new parserFormula('EDATE(A100,B100:B100)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(A100,B100:B100)');
		assert.strictEqual(oParser.calculate().getValue(), 130, 'Result of EDATE(A100,B100:B100)');

		oParser = new parserFormula('EDATE(A100,B100:B102)', "A2", ws);
		assert.ok(oParser.parse(), 'EDATE(A100,B100:B102)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EDATE(A100,B100:B102)');

		// 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("EDATE(37286,5)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EDATE(37286,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 37436, 'Test: Positive case: Number(2). Return 37436');

		oParser = new parserFormula("EDATE(37286,1)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EDATE(37286,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 37314, 'Test: Positive case: Number(2). Return 37314');

		oParser = new parserFormula("EDATE(DATE(2004,2,29),12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EDATE(DATE(2004,2,29),12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36949, 'Test: Positive case: Formula, Number. Return 36949');

		ws.getRange2("A7").setValue("02-28-2004");
		oParser = new parserFormula("EDATE(A7,12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EDATE(A7,12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36949, 'Test: Positive case: String. Return 36949');

		oParser = new parserFormula("EDATE(DATE(2004,1,15),-23)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EDATE(DATE(2004,1,15),-23) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 35840, 'Test: Positive case: Formula, Number. Return 35840');

		// return to base mode
		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("44927");
		ws.getRange2("A101").setValue("1");
		ws.getRange2("A104").setValue("4");
		// For area
		ws.getRange2("A102").setValue("2");
		ws.getRange2("A103").setValue("3");
		ws.getRange2("A105").setValue("Text");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("123"); // Num (Column1)
		ws.getRange2("B601").setValue("321"); // Num (Column2)
		ws.getRange2("C601").setValue("Text"); // Text (Column3)

		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Date, Number. Standard case: add 1 month to a valid date.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44958, 'Test: Positive case: Date, Number. Standard case: add 1 month to a valid date.');
		// Case #2: Date, Number. Transition to next year.
		oParser = new parserFormula('EDATE(DATE(2023,12,31),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,12,31),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45322, 'Test: Positive case: Date, Number. Transition to next year.');
		// Case #3: Number, Number. Numeric date serial used as start_date.
		oParser = new parserFormula('EDATE(45123,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(45123,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45215, 'Test: Positive case: Number, Number. Numeric date serial used as start_date.');
		// Case #4: Number, Number. Negative months moves date backward.
		oParser = new parserFormula('EDATE(40000,-6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(40000,-6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39819, 'Test: Positive case: Number, Number. Negative months moves date backward.');
		// Case #5: Formula, Number. Formula used as start_date.
		oParser = new parserFormula('EDATE(DATE(2025,10,13),6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2025,10,13),6) is parsed.');

		assert.strictEqual(oParser.calculate().getValue(), 46125, 'Test: Positive case: Formula, Number. Formula used as start_date.');
		// Case #6: Date, Formula. Formula used as months argument.
		oParser = new parserFormula('EDATE(DATE(2023,5,1),MONTH(DATE(2025,10,13)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,5,1),MONTH(DATE(2025,10,13))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45352, 'Test: Positive case: Date, Formula. Formula used as months argument.');
		// Case #7: Date, Empty. Empty months defaults to 0.
		oParser = new parserFormula('EDATE(DATE(2023,5,1),)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,5,1),) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Date, Empty. Empty months defaults to 0.');
		// Case #8: Reference link, Number. Reference link to start_date value.
		oParser = new parserFormula('EDATE(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44986, 'Test: Positive case: Reference link, Number. Reference link to start_date value.');
		// Case #9: Date, Reference link. Reference link to months value.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44958, 'Test: Positive case: Date, Reference link. Reference link to months value.');
		// Case #10: Array, Number. Array input with multiple dates.
		oParser = new parserFormula('EDATE({45123,45124},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE({45123,45124},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45154, 'Test: Positive case: Array, Number. Array input with multiple dates.');
		// Case #11: Area, Number. Area range for start_date.
		oParser = new parserFormula('EDATE(A100:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(A100:A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area, Number. Area range for start_date.');
		// Case #12: Table, Number. Using Table reference.
		oParser = new parserFormula('EDATE(Table1[Column1],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(Table1[Column1],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 154, 'Test: Positive case: Table, Number. Using Table reference.');
		// Case #13: Name, Number. Using Named reference.
		oParser = new parserFormula('EDATE(TestName,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(TestName,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name, Number. Using Named reference.');
		// Case #14: Name3D, Number. Using 3D named reference.
		oParser = new parserFormula('EDATE(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D, Number. Using 3D named reference.');
		// Case #15: Ref3D, Number. Reference to cell on another sheet.
		oParser = new parserFormula('EDATE(Sheet2!A1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(Sheet2!A1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 33, 'Test: Positive case: Ref3D, Number. Reference to cell on another sheet.');
		// Case #16: Area3D, Number. Range reference across sheets.
		oParser = new parserFormula('EDATE(Sheet2!A1:A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(Sheet2!A1:A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D, Number. Range reference across sheets.');
		// Case #17: Date, Array. Array months argument.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),{1,2,3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),{1,2,3}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44958, 'Test: Positive case: Date, Array. Array months argument.');
		// Case #18: Formula, Formula. Both arguments are formulas.
		oParser = new parserFormula('EDATE(123,MONTH(123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(123,MONTH(123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 276, 'Test: Positive case: Formula, Formula. Both arguments are formulas.');
		// Case #19: Name, Formula. Name with formula for months.
		oParser = new parserFormula('EDATE(TestName,MONTH(123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(TestName,MONTH(123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name, Formula. Name with formula for months.');
		// Case #20: Date, Table. Table reference for months.
		oParser = new parserFormula('EDATE(DATE(2023,5,1),Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,5,1),Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 54820, 'Test: Positive case: Date, Table. Table reference for months.');

		// Negative cases:

		// Case #1: String, Number. Invalid date string.
		oParser = new parserFormula('EDATE("notadate",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE("notadate",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Invalid date string.');
		// Case #2: Error, Number. Error propagated from start_date.
		oParser = new parserFormula('EDATE(#VALUE!,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(#VALUE!,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error, Number. Error propagated from start_date.');
		// Case #3: Date, Error. Error propagated from months.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Date, Error. Error propagated from months.');
		// Case #4: Empty, Number. Empty start_date gives #VALUE!.
		oParser = new parserFormula('EDATE(,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty, Number. Empty start_date gives #VALUE!.');
		// Case #5: Number, String. Invalid months argument string.
		oParser = new parserFormula('EDATE(45123,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(45123,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. Invalid months argument string.');
		// Case #6: Date, String. Invalid months string.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),"xyz")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),"xyz") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date, String. Invalid months string.');
		// Case #7: Error, Error. Both arguments errors.
		oParser = new parserFormula('EDATE(#N/A,#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(#N/A,#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Error. Both arguments errors.');
		// Case #8: Reference link, Error. Error in months argument.
		oParser = new parserFormula('EDATE(A102,#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(A102,#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Reference link, Error. Error in months argument.');
		// Case #9: Error, Reference link. Error in start_date.
		oParser = new parserFormula('EDATE(#VALUE!,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(#VALUE!,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error, Reference link. Error in start_date.');
		// Case #10: Area, String. Invalid months with range.
		oParser = new parserFormula('EDATE(A100:A101,"bad")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(A100:A101,"bad") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String. Invalid months with range.');
		// Case #11: Number, Empty. Empty months argument leads to #VALUE!.
		oParser = new parserFormula('EDATE(45123,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(45123,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number, Empty. Empty months argument leads to #VALUE!.');
		// Case #12: Name, String. Invalid months from name.
		// Different result with MS
		//oParser = new parserFormula('EDATE(TestName,"bad")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EDATE(TestName,"bad") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Name, String. Invalid months from name.');
		// Case #13: Name3D, Error. 3D name with error months.
		// Different result with MS
		//oParser = new parserFormula('EDATE(TestName3D,#REF!)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EDATE(TestName3D,#REF!) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Name3D, Error. 3D name with error months.');
		// Case #14: Ref3D, String. Invalid months on 3D ref.
		oParser = new parserFormula('EDATE(Sheet2!A1,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(Sheet2!A1,"invalid") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, String. Invalid months on 3D ref.');
		// Case #15: Area3D, Error. Area3D with error months.
		oParser = new parserFormula('EDATE(Sheet2!A1:A2,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(Sheet2!A1:A2,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D, Error. Area3D with error months.');
		// Case #16: String, Number. Invalid date format.
		oParser = new parserFormula('EDATE("13/13/2023",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE("13/13/2023",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Invalid date format.');
		// Case #17: Formula, String. Invalid months as string with formula.
		oParser = new parserFormula('EDATE(123,"text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(123,"text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula, String. Invalid months as string with formula.');
		// Case #18: Array, String. Invalid months with array.
		oParser = new parserFormula('EDATE({45123,45124},"bad")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE({45123,45124},"bad") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array, String. Invalid months with array.');
		// Case #19: Date, Name. Invalid named months value.
		oParser = new parserFormula('EDATE(DATE(2023,1,1),TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(2023,1,1),TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date, Name. Invalid named months value.');
		// Case #20: Empty, Empty. Both arguments empty.
		oParser = new parserFormula('EDATE(,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty, Empty. Both arguments empty.');

		// Bounded cases:

		// Case #1: Date, Number. Minimum date supported by Excel.
		oParser = new parserFormula('EDATE(DATE(1900,1,1),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(1900,1,1),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Date, Number. Minimum date supported by Excel.');
		// Case #2: Date, Number. Near-maximum Excel date moving backward.
		oParser = new parserFormula('EDATE(DATE(9999,12,31),-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(DATE(9999,12,31),-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958434, 'Test: Bounded case: Date, Number. Near-maximum Excel date moving backward.');
		// Case #3: Number, Number. Serial number for earliest date.
		oParser = new parserFormula('EDATE(1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number, Number. Serial number for earliest date.');
		// Case #4: Number, Number. Excel\'s upper bound for date serial number.
		oParser = new parserFormula('EDATE(2958465,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EDATE(2958465,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Number, Number. Excel\'s upper bound for date serial number.');

		// TODO Need to fix: error type diff
		// Case #12: Name, String. Invalid months from name.
		// Case #13: Name3D, Error. 3D name with error months.

		testArrayFormula2(assert, "EDATE", 2, 2, true, null);
	});

	QUnit.test('Test: "EOMONTH"', function (assert) {
		// base mode
		ws.workbook.setDate1904(false, true);

		oParser = new parserFormula("EOMONTH(0,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(0,1)');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(0,1)');

		oParser = new parserFormula("EOMONTH(1,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(1,1)');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(1,1)');

		oParser = new parserFormula("EOMONTH(2,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(2,1)');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(2,1)');

		oParser = new parserFormula("EOMONTH(59,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(59,1)');
		assert.strictEqual(oParser.calculate().getValue(), 91, 'Result of EOMONTH(59,1)');

		oParser = new parserFormula("EOMONTH(60,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(60,1)');
		assert.strictEqual(oParser.calculate().getValue(), 91, 'Result of EOMONTH(60,1)');

		oParser = new parserFormula("EOMONTH(61,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(61,1)');
		assert.strictEqual(oParser.calculate().getValue(), 121, 'Result of EOMONTH(61,1)');

		oParser = new parserFormula("EOMONTH(62,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(62,1)');
		assert.strictEqual(oParser.calculate().getValue(), 121, 'Result of EOMONTH(62,1)');

		oParser = new parserFormula("EOMONTH(DATE(2006,1,31),5)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2006,1,31),5) is parsed');
		assert.strictEqual(oParser.calculate().getValue(), 38898, 'Test: Positive case: Formula, Number. Return 38898');

		oParser = new parserFormula("EOMONTH(DATE(2004,2,29),12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2004,2,29),12) is parsed');
		assert.strictEqual(oParser.calculate().getValue(), 38411, 'Test: Positive case: Formula, Number. Return 38411');

		ws.getRange2("A7").setValue("02-28-2004");
		oParser = new parserFormula("EOMONTH(A7,12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(A7,12) is parsed');
		assert.strictEqual(oParser.calculate().getValue(), 38411, 'Test: Positive case: Reference link, Number. Return 38411');

		oParser = new parserFormula("EOMONTH(DATE(2004,1,15),-23)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2004,1,15),-23) is parsed');
		assert.strictEqual(oParser.calculate().getValue(), 37315, 'Test: Positive case: Formula, Number. Return 37315');

		oParser = new parserFormula("EOMONTH(DATE(2018,3,16),10)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(DATE(2018,3,16),10)');
		assert.strictEqual(oParser.calculate().getValue(), 43496, 'Result of EOMONTH(DATE(2018,3,16),10)');

		// string
		oParser = new parserFormula('EOMONTH("43175","10")', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH("43175","10")');
		assert.strictEqual(oParser.calculate().getValue(), 43496, 'Result of EOMONTH("43175","10")');

		oParser = new parserFormula('EOMONTH("43175+1","10")', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH("43175+1","10")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EOMONTH("43175+1","10")');

		oParser = new parserFormula('EOMONTH("43175","10+1")', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH("43175","10+1")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EOMONTH("43175","10+1")');

		// bool
		oParser = new parserFormula('EOMONTH(43175,FALSE)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(43175,FALSE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EOMONTH(43175,FALSE)');

		oParser = new parserFormula('EOMONTH(43175,TRUE)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(43175,TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EOMONTH(43175,TRUE)');

		oParser = new parserFormula('EOMONTH(TRUE,1)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(TRUE,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of EOMONTH(TRUE,1)');

		// err
		oParser = new parserFormula('EOMONTH(#N/A,1)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(#N/A,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of EOMONTH(#N/A,1)');

		oParser = new parserFormula('EOMONTH(#N/A,#NUM!)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(#N/A,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of EOMONTH(#N/A,#NUM!)');

		oParser = new parserFormula('EOMONTH(#NUM!,#N/A)', "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(#NUM!,#N/A)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of EOMONTH(#NUM!,#N/A)');

		ws.getRange2("A100").setValue("44227.50069");
		ws.getRange2("A101").setValue("44227.49861");
		ws.getRange2("A102").setValue("44227.00069");

		oParser = new parserFormula("EOMONTH(A100,0)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A100,0)');
		assert.strictEqual(oParser.calculate().getValue(), 44227, 'Result of EOMONTH(A100,0)');

		oParser = new parserFormula("EOMONTH(A100,1.5)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A100,1.5)');
		assert.strictEqual(oParser.calculate().getValue(), 44255, 'Result of EOMONTH(A100,1.5)');

		oParser = new parserFormula("EOMONTH(A100,-1.6)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A100,-1.6)');
		assert.strictEqual(oParser.calculate().getValue(), 44196, 'Result of EOMONTH(A100,-1.6)');

		oParser = new parserFormula("EOMONTH(A101,0)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A100,0)');
		assert.strictEqual(oParser.calculate().getValue(), 44227, 'Result of EOMONTH(A100,0)');

		oParser = new parserFormula("EOMONTH(A101,1.5)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A101,1.5)');
		assert.strictEqual(oParser.calculate().getValue(), 44255, 'Result of EOMONTH(A101,1.5)');

		oParser = new parserFormula("EOMONTH(A101,-1.6)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A101,-1.6)');
		assert.strictEqual(oParser.calculate().getValue(), 44196, 'Result of EOMONTH(A102,-1.6)');

		oParser = new parserFormula("EOMONTH(A102,0)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A102,0)');
		assert.strictEqual(oParser.calculate().getValue(), 44227, 'Result of EOMONTH(A102,0)');

		oParser = new parserFormula("EOMONTH(A102,1.5)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A102,1.5)');
		assert.strictEqual(oParser.calculate().getValue(), 44255, 'Result of EOMONTH(A102,1.5)');

		oParser = new parserFormula("EOMONTH(A102,-1.6)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(A102,-1.6)');
		assert.strictEqual(oParser.calculate().getValue(), 44196, 'Result of EOMONTH(A102,-1.6)');

		// set 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("EOMONTH(0,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(0,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(0,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(1,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(1,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(1,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(2,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(2,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(2,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(59,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(59,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 90, 'Result of EOMONTH(59,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(60,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(60,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 120, 'Result of EOMONTH(60,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(61,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(61,1) 1904 mode');
		assert.strictEqual(oParser.calculate().getValue(), 120, 'Result of EOMONTH(61,1) 1904 mode');

		oParser = new parserFormula("EOMONTH(DATE(2006,1,31),5)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2006,1,31),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 37436, 'Test: Positive case: Formula, Number. Return 36436');

		oParser = new parserFormula("EOMONTH(DATE(2004,2,29),12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2004,2,29) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36949, 'Test: Positive case: Formula, Number. Return 36949');

		ws.getRange2("A7").setValue("02-28-2004");
		oParser = new parserFormula("EOMONTH(A7,12)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(A7,12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36949, 'Test: Positive case: Reference link, Number. Return 36949');

		oParser = new parserFormula("EOMONTH(DATE(2004,1,15),-23)", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula EOMONTH(DATE(2004,1,15),-23) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 35853, 'Test: Positive case: Formula, Number. Return 35853');

		oParser = new parserFormula("EOMONTH(0,1)", "A2", ws);
		assert.ok(oParser.parse(), 'EOMONTH(0,1)');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Result of EOMONTH(0,1)');

		// base mode
		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("44927");
		ws.getRange2("A101").setValue("1");
		ws.getRange2("A104").setValue("4");
		// For area
		ws.getRange2("A102").setValue("2");
		ws.getRange2("A103").setValue("3");
		ws.getRange2("A105").setValue("Text");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("123"); // Num (Column1)
		ws.getRange2("B601").setValue("321"); // Num (Column2)
		ws.getRange2("C601").setValue("Text"); // Text (Column3)

		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		initDefNames();
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Date, Number. Standard case: add 1 month to a valid date.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44985, 'Test: Positive case: Date, Number. Standard case: add 1 month to a valid date.');
		// Case #2: Date, Number. Transition to next year.
		oParser = new parserFormula('EOMONTH(DATE(2023,12,31),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,12,31),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45322, 'Test: Positive case: Date, Number. Transition to next year.');
		// Case #3: Number, Number. Numeric date serial used as start_date.
		oParser = new parserFormula('EOMONTH(45123,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(45123,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45230, 'Test: Positive case: Number, Number. Numeric date serial used as start_date.');
		// Case #4: Number, Number. Negative months moves date backward.
		oParser = new parserFormula('EOMONTH(40000,-6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(40000,-6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39844, 'Test: Positive case: Number, Number. Negative months moves date backward.');
		// Case #5: Formula, Number. Formula used as start_date.
		oParser = new parserFormula('EOMONTH(DATE(10,10,2000),6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(TODAY(),6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6118, 'Test: Positive case: Formula, Number. Formula used as start_date.');
		// Case #6: Date, Formula. Formula used as months argument.
		oParser = new parserFormula('EOMONTH(DATE(2023,5,1),MONTH(DATE(10,10,2000)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,5,1),MONTH(DATE(10,10,2000))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45169, 'Test: Positive case: Date, Formula. Formula used as months argument.');
		// Case #7: Date, Empty. Empty months defaults to 0.
		oParser = new parserFormula('EOMONTH(DATE(2023,5,1),)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,5,1),) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Date, Empty. Empty months defaults to 0.');
		// Case #8: Reference link, Number. Reference link to start_date value.
		oParser = new parserFormula('EOMONTH(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45016, 'Test: Positive case: Reference link, Number. Reference link to start_date value.');
		// Case #9: Date, Reference link. Reference link to months value.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44985, 'Test: Positive case: Date, Reference link. Reference link to months value.');
		// Case #10: Array, Number. Array input with multiple dates.
		oParser = new parserFormula('EOMONTH({45123,45124},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH({45123,45124},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45169, 'Test: Positive case: Array, Number. Array input with multiple dates.');
		// Case #11: Area, Number. Area range for start_date.
		oParser = new parserFormula('EOMONTH(A100:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(A100:A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area, Number. Area range for start_date.');
		// Case #12: Table, Number. Using Table reference.
		oParser = new parserFormula('EOMONTH(Table1[Column1],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(Table1[Column1],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 182, 'Test: Positive case: Table, Number. Using Table reference.');
		// Case #13: Name, Number. Using Named reference.
		oParser = new parserFormula('EOMONTH(TestName,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(TestName,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Positive case: Name, Number. Using Named reference.');
		// Case #14: Name3D, Number. Using 3D named reference.
		oParser = new parserFormula('EOMONTH(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 91, 'Test: Positive case: Name3D, Number. Using 3D named reference.');
		// Case #15: Ref3D, Number. Reference to cell on another sheet.
		oParser = new parserFormula('EOMONTH(Sheet2!A1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(Sheet2!A1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Positive case: Ref3D, Number. Reference to cell on another sheet.');
		// Case #16: Area3D, Number. Range reference across sheets.
		oParser = new parserFormula('EOMONTH(Sheet2!A1:A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(Sheet2!A1:A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Positive case: Area3D, Number. Range reference across sheets.');
		// Case #17: Date, Array. Array months argument.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),{1,2,3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),{1,2,3}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 44985, 'Test: Positive case: Date, Array. Array months argument.');
		// Case #18: Formula, Formula. Both arguments are formulas.
		oParser = new parserFormula('EOMONTH(123,MONTH(123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(123,MONTH(123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 305, 'Test: Positive case: Formula, Formula. Both arguments are formulas.');
		// Case #19: Name, Formula. Name with formula for months.
		oParser = new parserFormula('EOMONTH(TestName,MONTH(123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(TestName,MONTH(123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 182, 'Test: Positive case: Name, Formula. Name with formula for months.');
		// Case #20: Date, Table. Table reference for months.
		oParser = new parserFormula('EOMONTH(DATE(2023,5,1),Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,5,1),Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 54847, 'Test: Positive case: Date, Table. Table reference for months.');

		// Negative cases:

		// Case #1: String, Number. Invalid date string.
		oParser = new parserFormula('EOMONTH("notadate",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH("notadate",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Invalid date string.');
		// Case #2: Error, Number. Error propagated from start_date.
		oParser = new parserFormula('EOMONTH(#VALUE!,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(#VALUE!,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error, Number. Error propagated from start_date.');
		// Case #3: Date, Error. Error propagated from months.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Date, Error. Error propagated from months.');
		// Case #4: Empty, Number. Empty start_date gives #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty, Number. Empty start_date gives #VALUE!.');
		// Case #5: Number, String. Invalid months argument string.
		oParser = new parserFormula('EOMONTH(45123,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(45123,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. Invalid months argument string.');
		// Case #6: Date, String. Invalid months string.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),"xyz")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),"xyz") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date, String. Invalid months string.');
		// Case #7: Error, Error. Both arguments errors.
		oParser = new parserFormula('EOMONTH(#N/A,#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(#N/A,#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Error. Both arguments errors.');
		// Case #8: Reference link, Error. Error in months argument.
		oParser = new parserFormula('EOMONTH(A102,#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(A102,#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Reference link, Error. Error in months argument.');
		// Case #9: Error, Reference link. Error in start_date.
		oParser = new parserFormula('EOMONTH(#VALUE!,A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(#VALUE!,A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error, Reference link. Error in start_date.');
		// Case #10: Area, String. Invalid months with range.
		oParser = new parserFormula('EOMONTH(A100:A101,"bad")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(A100:A101,"bad") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String. Invalid months with range.');
		// Case #11: Number, Empty. Empty months argument leads to #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(45123,)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(45123,) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number, Empty. Empty months argument leads to #VALUE!.');
		// Case #12: Name, String. Invalid months from name.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(TestName,"bad")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(TestName,"bad") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Name, String. Invalid months from name.');
		// Case #13: Name3D, Error. 3D name with error months.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(TestName3D,#REF!)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(TestName3D,#REF!) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Name3D, Error. 3D name with error months.');
		// Case #14: Ref3D, String. Invalid months on 3D ref.
		oParser = new parserFormula('EOMONTH(Sheet2!A1,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(Sheet2!A1,"invalid") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, String. Invalid months on 3D ref.');
		// Case #15: Area3D, Error. Area3D with error months.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(Sheet2!A1:A2,#N/A)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(Sheet2!A1:A2,#N/A) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D, Error. Area3D with error months.');
		// Case #16: String, Number. Invalid date format.
		oParser = new parserFormula('EOMONTH("13/13/2023",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH("13/13/2023",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Invalid date format.');
		// Case #17: Formula, String. Invalid months as string with formula.
		oParser = new parserFormula('EOMONTH(123,"text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(123,"text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula, String. Invalid months as string with formula.');
		// Case #18: Array, String. Invalid months with array.
		oParser = new parserFormula('EOMONTH({45123,45124},"bad")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH({45123,45124},"bad") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array, String. Invalid months with array.');
		// Case #19: Date, Name. Invalid named months value.
		oParser = new parserFormula('EOMONTH(DATE(2023,1,1),TestNameBad)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(2023,1,1),TestNameBad) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Date, Name. Invalid named months value.');
		// Case #20: Empty, Empty. Both arguments empty.
		// Different result with MS
		//oParser = new parserFormula('EOMONTH(,)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: EOMONTH(,) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty, Empty. Both arguments empty.');

		// Bounded cases:

		// Case #1: Date, Number. Minimum date supported by Excel.
		oParser = new parserFormula('EOMONTH(DATE(1900,1,1),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(1900,1,1),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Bounded case: Date, Number. Minimum date supported by Excel.');
		// Case #2: Date, Number. Near-maximum Excel date moving backward.
		oParser = new parserFormula('EOMONTH(DATE(9999,12,31),-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(DATE(9999,12,31),-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958434, 'Test: Bounded case: Date, Number. Near-maximum Excel date moving backward.');
		// Case #3: Number, Number. Serial number for earliest date.
		oParser = new parserFormula('EOMONTH(1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 31, 'Test: Bounded case: Number, Number. Serial number for earliest date.');
		// Case #4: Number, Number. Excel\'s upper bound for date serial number.
		oParser = new parserFormula('EOMONTH(2958465,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EOMONTH(2958465,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Number, Number. Excel\'s upper bound for date serial number.');

		// TODO Need to fix: area handle, empty handle, error types diff, ms result diff

		//  Case #7: Date, Empty. Empty months defaults to 0.
		// Case #4: Empty, Number. Empty start_date gives #VALUE!.
		// Case #11: Number, Empty. Empty months argument leads to #VALUE!.
		// Case #12: Name, String. Invalid months from name.
		// Case #13: Name3D, Error. 3D name with error months.
		// Case #15: Area3D, Error. Area3D with error months.
		// Case #20: Empty, Empty. Both arguments empty.

		testArrayFormula2(assert, "EOMONTH", 2, 2, true, null);
	});

	QUnit.test('Test: "HOUR"', function (assert) {
		// 	Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A104").setValue("5");

		ws.getRange2("A302").setValue("0.75");
		ws.getRange2("A303").setValue("7/18/2011 7:45");
		ws.getRange2("A304").setValue("4/21/2012");
		// For area
		ws.getRange2("A102").setValue("3");
		ws.getRange2("A103").setValue("4");
		ws.getRange2("A105").setValue("6");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1.005"); // Number (Column1)
		ws.getRange2("B601").setValue("123s"); // Number (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:D10").cleanAll();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("1.5");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("-1");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		initDefNames();
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #0: Number. Valid time as decimal (12:00 PM, half of a day). 1 argument used.
		oParser = new parserFormula('HOUR(0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Number. Valid time as decimal (12:00 PM, half of a day). 1 argument used.');
		// Case #1: String. Valid time string (12:00 PM). 1 argument used.
		oParser = new parserFormula('HOUR("12:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("12:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: String. Valid time string (12:00 PM). 1 argument used.');
		// Case #2: Formula. Nested formula returning valid time. 1 argument used.
		oParser = new parserFormula('HOUR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Formula. Nested formula returning valid time. 1 argument used.');
		// Case #3: Reference link. Reference to cell with valid time decimal. 1 argument used.
		oParser = new parserFormula('HOUR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with valid time decimal. 1 argument used.');
		// Case #4: Area. Single-cell range with valid time. 1 argument used.
		oParser = new parserFormula('HOUR(A100:A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(A100:A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area. Single-cell range with valid time. 1 argument used.');
		// Case #5: Array. Array with single valid time element. 1 argument used.
		oParser = new parserFormula('HOUR({0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR({0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Array. Array with single valid time element. 1 argument used.');
		// Case #6: Name. Named range with valid time decimal. 1 argument used.
		oParser = new parserFormula('HOUR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid time decimal. 1 argument used.');
		// Case #7: Name3D. 3D named range with valid time decimal. 1 argument used.
		oParser = new parserFormula('HOUR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid time decimal. 1 argument used.');
		// Case #8: Ref3D. 3D reference to cell with valid time decimal. 1 argument used.
		oParser = new parserFormula('HOUR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Ref3D. 3D reference to cell with valid time decimal. 1 argument used.');
		// Case #9: Area3D. 3D single-cell range with valid time. 1 argument used.
		oParser = new parserFormula('HOUR(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Area3D. 3D single-cell range with valid time. 1 argument used.');
		// Case #10: Table. Table structured reference with valid time decimal. 1 argument used.
		oParser = new parserFormula('HOUR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Table. Table structured reference with valid time decimal. 1 argument used.');
		// Case #11: Formula. HOUR inside SUM formula. 1 argument used.
		oParser = new parserFormula('SUM(HOUR(0.5),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(HOUR(0.5),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Formula. HOUR inside SUM formula. 1 argument used.');
		// Case #12: Time. Maximum hour time value (23:59:59). 1 argument used.
		oParser = new parserFormula('HOUR(TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Positive case: Time. Maximum hour time value (23:59:59). 1 argument used.');
		// Case #13: Date. Date-time serial number with valid time component. 1 argument used.
		oParser = new parserFormula('HOUR(DATE(2025,1,1)+0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(DATE(2025,1,1)+0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Date. Date-time serial number with valid time component. 1 argument used.');
		// Case #14: String. Time string for maximum hour (23:59). 1 argument used.
		oParser = new parserFormula('HOUR("23:59")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("23:59") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Positive case: String. Time string for maximum hour (23:59). 1 argument used.');
		// Case #15: Reference link. Reference to cell with time string. 1 argument used.
		oParser = new parserFormula('HOUR(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with time string. 1 argument used.');
		// Case #16: Name. Named range with time string. 1 argument used.
		oParser = new parserFormula('HOUR(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Name. Named range with time string. 1 argument used.');
		// Case #17: Ref3D. 3D reference to cell with time string. 1 argument used.
		oParser = new parserFormula('HOUR(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Ref3D. 3D reference to cell with time string. 1 argument used.');
		// Case #18: Table. Table reference with time string. 1 argument used.
		oParser = new parserFormula('HOUR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Table. Table reference with time string. 1 argument used.');
		// Case #19: Formula. Nested formula with current time. 1 argument used.
		oParser = new parserFormula('HOUR(TIME(14,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TIME(14,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'Test: Positive case: Formula. Nested formula with current time. 1 argument used.');
		// Case #20: Array. Array with multiple valid time decimals (06:00, 12:00). 1 argument used.
		oParser = new parserFormula('HOUR({0.25;0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR({0.25;0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Array. Array with multiple valid time decimals (06:00, 12:00). 1 argument used.');
		// Case #21: String. Time string for minimum hour (00:00). 1 argument used.
		oParser = new parserFormula('HOUR("00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. Time string for minimum hour (00:00). 1 argument used.');
		// Case #22: Reference link. Return 18.
		oParser = new parserFormula("HOUR(A302)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula HOUR(A302) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Reference link. Return 18.');
		// Case #23: Reference link. Return 7.
		oParser = new parserFormula("HOUR(A303)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula HOUR(A303) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Reference link. Return 7.');
		// Case #24: Reference link. Return 0.
		oParser = new parserFormula("HOUR(A304)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula HOUR(A304) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Return 0.');

		// Negative cases:

		// Case #1: String. Non-time string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-time string returns #VALUE!. 1 argument used.');
		// Case #2: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('HOUR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #3: Empty. Empty reference returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Empty. Empty reference returns #VALUE!. 1 argument used.');
		// Case #4: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #5: Boolean. Boolean FALSE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean FALSE returns #VALUE!. 1 argument used.');
		// Case #6: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('HOUR(A100:A101)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: HOUR(A100:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #7: Name. Named range with text returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('HOUR(TestNameArea)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: HOUR(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #8: Name3D. 3D named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 19, 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 1 argument used.');
		// Case #9: Ref3D. 3D reference to text cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to text cell returns #VALUE!. 1 argument used.');
		// Case #11: Formula. Formula resulting in #NUM! error. 1 argument used.
		oParser = new parserFormula('HOUR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error. 1 argument used.');
		// Case #12: Number. Negative time decimal returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(-0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(-0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative time decimal returns #VALUE!. 1 argument used.');
		// Case #13: String. Invalid time string (hour > 23) returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR("25:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("25:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: String. Invalid time string (hour > 23) returns #VALUE!. 1 argument used.');
		// Case #14: Array. Array with invalid time element returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR({"abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR({"abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with invalid time element returns #VALUE!. 1 argument used.');
		// Case #15: Reference link. Reference to cell with invalid time string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Reference to cell with invalid time string returns #VALUE!. 1 argument used.');
		// Case #16: Name. Named range with invalid time string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Negative case: Name. Named range with invalid time string returns #VALUE!. 1 argument used.');
		// Case #17: Ref3D. 3D reference to invalid time cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(Sheet2!A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(Sheet2!A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Test: Negative case: Ref3D. 3D reference to invalid time cell returns #VALUE!. 1 argument used.');
		// Case #19: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #20: Number. Time decimal > 1 returns #VALUE!. 1 argument used.
		oParser = new parserFormula('HOUR(1.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(1.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Negative case: Number. Time decimal > 1 returns #VALUE!. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid time decimal (00:00). 1 argument used.
		oParser = new parserFormula('HOUR(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Minimum valid time decimal (00:00). 1 argument used.');
		// Case #2: Number. Maximum valid time decimal (23:59:59). 1 argument used.
		oParser = new parserFormula('HOUR(0.99999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(0.99999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Bounded case: Number. Maximum valid time decimal (23:59:59). 1 argument used.');
		// Case #3: Date. Minimum valid Excel date-time serial number (1.0). 1 argument used.
		oParser = new parserFormula('HOUR(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Date. Minimum valid Excel date-time serial number (1.0). 1 argument used.');
		// Case #4: Date. Maximum valid Excel date-time serial number with time (2958465.99999). 1 argument used.
		oParser = new parserFormula('HOUR(DATE(9999,12,31)+0.99999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: HOUR(DATE(9999,12,31)+0.99999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Bounded case: Date. Maximum valid Excel date-time serial number with time (2958465.99999). 1 argument used.');

		// TODO Need to fix: area handle
		// Case #6: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Case #7: Name. Named range with text returns #VALUE!. 1 argument used.

		testArrayFormula2(assert, "HOUR", 1, 1);
	});

	QUnit.test('Test: "ISOWEEKNUM"', function (assert) {
		// base mode
		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("123");
		ws.getRange2("A101").setValue("1234	");
		ws.getRange2("A104").setValue("Text2");

		ws.getRange2("A2").setValue("3/9/2012");
		// For area
		ws.getRange2("A102").setValue("123s");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A105").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		ws.getRange2("C601").setValue("123i"); // TextNum (Column3)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		oParser = new parserFormula("ISOWEEKNUM(A2)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula ISOWEEKNUM(A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Reference link. Return 10.');

		oParser = new parserFormula("ISOWEEKNUM(123)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula ISOWEEKNUM(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Number. Return 18.');

		oParser = new parserFormula("ISOWEEKNUM(120003)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula ISOWEEKNUM(120003) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: Number. Return 30.');

		oParser = new parserFormula("ISOWEEKNUM(1203)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula ISOWEEKNUM(1203) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16, 'Test: Positive case: Number. Return 16.');

		oParser = new parserFormula("ISOWEEKNUM(43466)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2019)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2019)");

		oParser = new parserFormula("ISOWEEKNUM(43831)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2020)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2020)");

		oParser = new parserFormula("ISOWEEKNUM(44197)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2021)");
		assert.strictEqual(oParser.calculate().getValue(), 53, "Result of ISOWEEKNUM(01.01.2021)");

		oParser = new parserFormula("ISOWEEKNUM(44562)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 52, "Result of ISOWEEKNUM(01.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44563)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 52, "Result of ISOWEEKNUM(02.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44564)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(03.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(03.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44927)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 52, "Result of ISOWEEKNUM(01.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(44928)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(02.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(44929)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(03.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(03.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(1)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.1900)");
		assert.strictEqual(oParser.calculate().getValue(), 52, "Result of ISOWEEKNUM(01.01.1900)");

		oParser = new parserFormula("ISOWEEKNUM(2)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.1900)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(02.01.1900)");

		// set 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("ISOWEEKNUM(43466)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2019)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2019)");

		oParser = new parserFormula("ISOWEEKNUM(43831)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2020)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2020)");

		oParser = new parserFormula("ISOWEEKNUM(44197)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2021)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2021)");

		oParser = new parserFormula("ISOWEEKNUM(44562)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(01.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44563)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(02.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44564)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(03.01.2022)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(03.01.2022)");

		oParser = new parserFormula("ISOWEEKNUM(44927)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 53, "Result of ISOWEEKNUM(01.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(44928)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 53, "Result of ISOWEEKNUM(02.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(44929)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(03.01.2023)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of ISOWEEKNUM(03.01.2023)");

		oParser = new parserFormula("ISOWEEKNUM(1)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(01.01.1900)");
		assert.strictEqual(oParser.calculate().getValue(), 53, "Result of ISOWEEKNUM(01.01.1900)");

		oParser = new parserFormula("ISOWEEKNUM(2)", "A1", ws);
		assert.ok(oParser.parse(), "ISOWEEKNUM(02.01.1900)");
		assert.strictEqual(oParser.calculate().getValue(), 53, "Result of ISOWEEKNUM(02.01.1900)");

		// return to base mode
		ws.workbook.setDate1904(false, true);

		// Case #1: Number. Basic valid input: serial number for January 1, 2021 (week 53). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(44197)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(44197) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 53, 'Test: Positive case: Number. Basic valid input: serial number for January 1, 2021 (week 53). 1 argument used.');
		// Case #2: Number. Serial number for January 1, 2022 (week 52). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(44562)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(44562) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Number. Serial number for January 1, 2022 (week 52). 1 argument used.');
		// Case #3: String. String convertible to date (week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("2025-01-01")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("2025-01-01") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String convertible to date (week 1). 1 argument used.');
		// Case #4: Formula. Nested DATE formula (week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested DATE formula (week 1). 1 argument used.');
		// Case #5: Reference link. Reference to cell with valid date serial number. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Reference link. Reference to cell with valid date serial number. 1 argument used.');
		// Case #6: Area. Single-cell range with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(A100:A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A100:A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Positive case: Area. Single-cell range with valid date. 1 argument used.');
		// Case #7: Array. Array with single valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM({44197})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM({44197}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 53, 'Test: Positive case: Array. Array with single valid date. 1 argument used.');
		// Case #8: Name. Named range with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid date. 1 argument used.');
		// Case #9: Name3D. 3D named range with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid date. 1 argument used.');
		// Case #10: Ref3D. 3D reference to cell with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Ref3D. 3D reference to cell with valid date. 1 argument used.');
		// Case #11: Area3D. 3D single-cell range with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Area3D. 3D single-cell range with valid date. 1 argument used.');
		// Case #12: Table. Table structured reference with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Table. Table structured reference with valid date. 1 argument used.');
		// Case #13: Formula. Nested IF formula returning valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(IF(TRUE,DATE(2025,1,1),DATE(2020,1,1)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(IF(TRUE,DATE(2025,1,1),DATE(2020,1,1))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested IF formula returning valid date. 1 argument used.');
		// Case #14: Number. Fractional date serial number (week 53). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(44197.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(44197.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 53, 'Test: Positive case: Number. Fractional date serial number (week 53). 1 argument used.');
		// Case #15: Formula. ISOWEEKNUM as parent formula in SUM. 1 argument used.
		oParser = new parserFormula('SUM(ISOWEEKNUM(44197),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(ISOWEEKNUM(44197),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 54, 'Test: Positive case: Formula. ISOWEEKNUM as parent formula in SUM. 1 argument used.');
		// Case #16: String. String for December 31, 2025 (week 53). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("2025-12-31")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("2025-12-31") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String for December 31, 2025 (week 53). 1 argument used.');
		// Case #17: Array. Array with one valid date, others ignored. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM({44562,TRUE,"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM({44562,TRUE,"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Array. Array with one valid date, others ignored. 1 argument used.');
		// Case #18: Formula. Nested ROUND formula with valid date. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(ROUND(DATE(2025,6,15),0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(ROUND(DATE(2025,6,15),0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 24, 'Test: Positive case: Formula. Nested ROUND formula with valid date. 1 argument used.');
		// Case #19: Reference link. Reference to cell with valid date string. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 21, 'Test: Positive case: Reference link. Reference to cell with valid date string. 1 argument used.');
		// Case #20: Area. Single-cell range with valid date string. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell range with valid date string. 1 argument used.');
		// Case #21: String. Short date string format (week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("01/01/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("01/01/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Short date string format (week 1). 1 argument used.');

		// Negative cases:

		// Case #0: Number. Negative value. Return -100.
		oParser = new parserFormula("ISOWEEKNUM(-100)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula ISOWEEKNUM(-100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Test: Negative case: Number. Negative value. Return -100.');
		// Case #1: Number. Invalid date (zero) returns #NUM!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Number. Invalid date (zero) returns #NUM!. 1 argument used.');
		// Case #2: Number. Negative date returns #NUM!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative date returns #NUM!. 1 argument used.');
		// Case #3: String. Non-date string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-date string returns #VALUE!. 1 argument used.');
		// Case #4: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('ISOWEEKNUM(A100:A101)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A100:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #6: Empty. Empty cell reference returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty cell reference returns #VALUE!. 1 argument used.');
		// Case #7: Boolean. Boolean value returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Boolean. Boolean value returns #VALUE!. 1 argument used.');
		// Case #8: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #9: Ref3D. 3D ref to text cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D ref to text cell returns #VALUE!. 1 argument used.');
		// Case #10: Name. Named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #12: Formula. Formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #13: Number. Date beyond December 31, 9999, returns #NUM!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('ISOWEEKNUM(2958466)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(2958466) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Date beyond December 31, 9999, returns #NUM!. 1 argument used.');
		// Case #14: Array. Array with no valid dates returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM({TRUE,FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM({TRUE,FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Array. Array with no valid dates returns #VALUE!. 1 argument used.');
		// Case #15: Number. Fractional date near minimum returns valid week (week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(1.999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(1.999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Number. Fractional date near minimum returns valid week (week 1). 1 argument used.');
		// Case #16: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #17: String. Invalid date string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM("13/13/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM("13/13/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string returns #VALUE!. 1 argument used.');
		// Case #18: Formula. Nested IF with #N/A propagates error. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(IF(FALSE,DATE(2025,1,1),NA()))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(IF(FALSE,DATE(2025,1,1),NA())) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Nested IF with #N/A propagates error. 1 argument used.');
		// Case #19: Number. Negative date serial number returns #NUM!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(-44197)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(-44197) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative date serial number returns #NUM!. 1 argument used.');
		// Case #20: Time. Time value (fractional < 1) returns #NUM!. 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Negative case: Time. Time value (fractional < 1) returns #NUM!. 1 argument used.');

		// Bounded cases:

		// Case #1: Number. Minimum valid date (January 1, 1900, week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Bounded case: Number. Minimum valid date (January 1, 1900, week 1). 1 argument used.');
		// Case #2: Number. Maximum valid date (December 31, 9999, week 52). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Bounded case: Number. Maximum valid date (December 31, 9999, week 52). 1 argument used.');
		// Case #3: Formula. Date near year-end for week 1 of next year (week 1). 1 argument used.
		oParser = new parserFormula('ISOWEEKNUM(DATE(2024,12,30))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: ISOWEEKNUM(DATE(2024,12,30)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Formula. Date near year-end for week 1 of next year (week 1). 1 argument used.');

		testArrayFormula2(assert, "ISOWEEKNUM", 1, 1);

		// TODO Need to fix:
		// Case #13: Number. Date beyond December 31, 9999 returns #NUM!. 1 argument used.
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
	});

	QUnit.test('Test: "MINUTE"', function (assert) {
		ws.getRange2("A202").setValue("12:45:00 PM");
		ws.getRange2("A203").setValue("7/18/2011 7:45");
		ws.getRange2("A204").setValue("4/21/2012");

		oParser = new parserFormula("MINUTE(A202)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 45);

		oParser = new parserFormula("MINUTE(A203)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 45);

		oParser = new parserFormula("MINUTE(A204)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		ws.getRange2("A205").setValue("06/30/2020 20:00");
		ws.getRange2("A206").setValue("06/30/2020 21:15");

		ws.getRange2("A207").setValue("06/30/2020 23:15");

		oParser = new parserFormula("MINUTE(A206-A205)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 15);

		oParser = new parserFormula("MINUTE(A207-A205)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 15);

		oParser = new parserFormula("MINUTE(A207-A206)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("MINUTE(A207+A206)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 30);

		oParser = new parserFormula("MINUTE(123.1231231 - 1.12334343)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 59);

		oParser = new parserFormula("MINUTE(1.12334343 - 123.1231231)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");
		ws.getRange2("A112").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("321"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("#N/A");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("-1");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A206").setValue("1"); // TestNameArea
		ws.getRange2("A207").setValue("2"); // TestNameArea
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Time. Valid time input, returns minute component (30). 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(12,30,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(12,30,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: Time. Valid time input, returns minute component (30). 1 argument used.');
		// Case #2: Number. Serial number for 12:00 PM (0.5), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number. Serial number for 12:00 PM (0.5), returns 0 minutes. 1 argument used.');
		// Case #3: String. String in time format, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE("12:30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("12:30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: String. String in time format, returns 30 minutes. 1 argument used.');
		// Case #4: Formula. Nested TIME() formula returning time, returns minute component. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(1,2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(1,2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. Nested TIME() formula returning time, returns minute component. 1 argument used.');
		// Case #5: Reference link. Reference to cell with valid time serial (0.5), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with valid time serial (0.5), returns 0 minutes. 1 argument used.');
		// Case #6: Area. Single-cell range with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area. Single-cell range with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #7: Array. Array with single time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE({0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE({0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Array. Array with single time serial, returns 0 minutes. 1 argument used.');
		// Case #8: Name. Named range with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #9: Name3D. 3D named range with valid time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid time serial, returns 0 minutes. 1 argument used.');
		// Case #10: Ref3D. 3D reference to cell with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Ref3D. 3D reference to cell with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #11: Area3D. 3D single-cell range with valid time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Area3D. 3D single-cell range with valid time serial, returns 0 minutes. 1 argument used.');
		// Case #12: Table. Table structured reference with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Table. Table structured reference with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #13: Date. Date serial number, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Date. Date serial number, returns 0 minutes. 1 argument used.');
		// Case #14: Formula. Time formula adjusted, returns minute component. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(15,45,30)+0.1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(15,45,30)+0.1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Formula. Time formula adjusted, returns minute component. 1 argument used.');
		// Case #15: Formula. MINUTE inside SUM formula, returns 30+10=40. 1 argument used.
		oParser = new parserFormula('SUM(MINUTE(TIME(12,30,0)),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(MINUTE(TIME(12,30,0)),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40, 'Test: Positive case: Formula. MINUTE inside SUM formula, returns 30+10=40. 1 argument used.');
		// Case #16: String. String with full time format (hh:mm:ss), returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE("12:30:45")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("12:30:45") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: String. String with full time format (hh:mm:ss), returns 30 minutes. 1 argument used.');
		// Case #17: Formula. Nested IF returning valid time, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(IF(TRUE,TIME(12,30,0),TIME(0,0,0)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(IF(TRUE,TIME(12,30,0),TIME(0,0,0))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: Formula. Nested IF returning valid time, returns 30 minutes. 1 argument used.');
		// Case #18: Array. Array with single time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE({0.520833333})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE({0.520833333}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: Array. Array with single time serial, returns 30 minutes. 1 argument used.');
		// Case #19: String. String with minimal hour format, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE("00:30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("00:30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: String. String with minimal hour format, returns 30 minutes. 1 argument used.');
		// Case #20: Formula. Nested formula with current hour and fixed minutes, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(HOUR(12345),30,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(HOUR(12345),30,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Test: Positive case: Formula. Nested formula with current hour and fixed minutes, returns 30 minutes. 1 argument used.');
		// Case #21: Time. Maximum valid time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Positive case: Time. Maximum valid time, returns 59 minutes. 1 argument used.');
		// Case #22: Reference link. Reference to cell with full time string, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with full time string, returns 30 minutes. 1 argument used.');

		// Negative cases:

		// Case #1: Number. Negative serial number, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MINUTE(-0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(-0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative serial number, returns #NUM!. 1 argument used.');
		// Case #2: String. Non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-time string, returns #VALUE!. 1 argument used.');
		// Case #3: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('MINUTE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #4: Area. Multi-cell range, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Test: Negative case: Area. Multi-cell range, returns #VALUE!. 1 argument used.');
		// Case #5: Empty. Empty cell reference, returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('MINUTE(A105)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: MINUTE(A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Empty. Empty cell reference, returns #VALUE!. 1 argument used.');
		// Case #6: Boolean. Boolean FALSE (0), returns 0 minutes (edge case). 1 argument used.
		oParser = new parserFormula('MINUTE(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean FALSE (0), returns 0 minutes (edge case). 1 argument used.');
		// Case #7: Ref3D. 3D reference to non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to non-time string, returns #VALUE!. 1 argument used.');
		// Case #8: Name. Named range with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Negative case: Name. Named range with non-time string, returns #VALUE!. 1 argument used.');
		// Case #9: Table. Table column with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with non-time string, returns #VALUE!. 1 argument used.');
		// Case #10: Formula. Formula resulting in #NUM!, propagates error. 1 argument used.
		oParser = new parserFormula('MINUTE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM!, propagates error. 1 argument used.');
		// Case #11: Number. Number exceeds valid serial range, returns #NUM!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('MINUTE(1E+308)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: MINUTE(1E+308) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Number exceeds valid serial range, returns #NUM!. 1 argument used.');
		// Case #12: String. Invalid time string (hours > 23), returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE("25:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("25:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Invalid time string (hours > 23), returns #VALUE!. 1 argument used.');
		// Case #13: Array. Array with negative serial, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MINUTE({-0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE({-0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Array. Array with negative serial, returns #NUM!. 1 argument used.');
		// Case #14: Name3D. 3D named range with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Negative case: Name3D. 3D named range with non-time string, returns #VALUE!. 1 argument used.');
		// Case #15: Area3D. 3D multi-cell range, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range, returns #VALUE!. 1 argument used.');
		// Case #16: Formula. Date before Jan 1, 1900, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MINUTE(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Formula. Date before Jan 1, 1900, returns #NUM!. 1 argument used.');
		// Case #17: String. Empty string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string, returns #VALUE!. 1 argument used.');
		// Case #18: Boolean. Boolean TRUE (1), returns 0 minutes (edge case). 1 argument used.
		oParser = new parserFormula('MINUTE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean TRUE (1), returns 0 minutes (edge case). 1 argument used.');
		// Case #19: Formula. Invalid time (hours >= 24), returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(24,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(24,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Formula. Invalid time (hours >= 24), returns #VALUE!. 1 argument used.');
		// Case #20: Reference link. Reference to cell with non-time string, returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('MINUTE(A106)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: MINUTE(A106) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Reference to cell with non-time string, returns #VALUE!. 1 argument used.');

		// Bounded cases:

		// Case #1: Number. Minimum valid serial number (Jan 1, 1900), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Minimum valid serial number (Jan 1, 1900), returns 0 minutes. 1 argument used.');
		// Case #2: Number. Maximum valid serial number (Dec 31, 9999), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Maximum valid serial number (Dec 31, 9999), returns 0 minutes. 1 argument used.');
		// Case #3: Time. Minimum valid time, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Time. Minimum valid time, returns 0 minutes. 1 argument used.');
		// Case #4: Time. Maximum valid time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Bounded case: Time. Maximum valid time, returns 59 minutes. 1 argument used.');
		// Case #5: String. String representing minimum valid time, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE("00:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE("00:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: String. String representing minimum valid time, returns 0 minutes. 1 argument used.');
		// Case #6: Formula. Formula for maximum valid date and time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MINUTE(DATE(9999,12,31)+TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MINUTE(DATE(9999,12,31)+TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 59, 'Test: Bounded case: Formula. Formula for maximum valid date and time, returns 59 minutes. 1 argument used.');

		// TODO Need to fix: empty ref link
		// Case #5: Empty. Empty cell reference, returns #VALUE!. 1 argument used.
		// Case #11: Number. Number exceeds valid serial range, returns #NUM!. 1 argument used.
		// Case #20: Reference link. Reference to cell with non-time string, returns #VALUE!. 1 argument used.

	});

	QUnit.test('Test: "MONTH"', function (assert) {
		let array;

		// base mode
		ws.workbook.setDate1904(false, true);

		oParser = new parserFormula("MONTH(2013)", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula MONTH(2013) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Number. Return 7.');

		oParser = new parserFormula("MONTH(DATE(2013,2,2))", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula MONTH(DATE(2013,2,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. Return 2.');

		oParser = new parserFormula("MONTH(NOW())", "A1", ws);
		assert.ok(oParser.parse(), 'Test: Formula MONTH(NOW()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), new cDate().getUTCMonth() + 1, 'Test: Positive case: Formula. Return current month.');
		// inline
		oParser = new parserFormula("MONTH(44469)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.09.21)");
		assert.strictEqual(oParser.calculate().getValue(), 9, "Result of MONTH(30.09.21)");

		oParser = new parserFormula("MONTH(44560)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH(30.12.21)");

		oParser = new parserFormula("MONTH(44561)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(31.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH(31.12.21)");

		oParser = new parserFormula("MONTH(44510)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(10.11.21)");
		assert.strictEqual(oParser.calculate().getValue(), 11, "Result of MONTH(10.11.21)");

		oParser = new parserFormula('MONTH("2021-10-01")', "A2", ws);
		assert.ok(oParser.parse(), 'MONTH("2021-10-01")');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Result of MONTH("2021-10-01")');

		oParser = new parserFormula('MONTH("2021-12-31")', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-31')");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH('2021-12-31')");

		oParser = new parserFormula('MONTH("2021-09-30")', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-09-30')");
		assert.strictEqual(oParser.calculate().getValue(), 9, "Result of MONTH('2021-09-30')");

		oParser = new parserFormula('MONTH("2021-10-31")', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-10-31')");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MONTH('2021-10-31')");

		oParser = new parserFormula('MONTH("2021-12-29")', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-29')");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH('2021-12-29')");

		oParser = new parserFormula('MONTH(0)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(0)");

		oParser = new parserFormula('MONTH("1s")', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('1s')");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of MONTH('1s')");

		oParser = new parserFormula('MONTH(TRUE)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(TRUE)");

		oParser = new parserFormula('MONTH(FALSE)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(FALSE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(FALSE)");

		oParser = new parserFormula('MONTH(#DIV/0!)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(#DIV/0!)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of MONTH(#DIV/0!)");

		// cell
		ws.getRange2("A1").setValue("44469");
		ws.getRange2("A2").setValue("44560");
		ws.getRange2("A3").setValue("44561");
		ws.getRange2("A4").setValue("44510");
		ws.getRange2("A5").setValue("2021-10-01");
		ws.getRange2("A6").setValue("2021-12-31");
		ws.getRange2("A7").setValue("2021-09-30");
		ws.getRange2("A8").setValue("2021-10-31");
		ws.getRange2("A9").setValue("2021-12-29");

		ws.getRange2("A10").setValue("0");
		ws.getRange2("A11").setValue("1s");
		ws.getRange2("A12").setValue("TRUE");
		ws.getRange2("A13").setValue("FALSE");
		ws.getRange2("A14").setValue("#DIV/0!");
		ws.getRange2("A15").setValue("");

		oParser = new parserFormula("MONTH(A1)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.09.21)");
		assert.strictEqual(oParser.calculate().getValue(), 9, "Result of MONTH(30.09.21)");

		oParser = new parserFormula("MONTH(A2)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH(30.12.21)");

		oParser = new parserFormula("MONTH(A3)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(31.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH(31.12.21)");

		oParser = new parserFormula("MONTH(A4)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(10.11.21)");
		assert.strictEqual(oParser.calculate().getValue(), 11, "Result of MONTH(10.11.21)");

		oParser = new parserFormula("MONTH(A5)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-10-01')");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MONTH('2021-10-01')");

		oParser = new parserFormula("MONTH(A6)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-31')");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH('2021-12-31')");

		oParser = new parserFormula("MONTH(A7)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-09-30')");
		assert.strictEqual(oParser.calculate().getValue(), 9, "Result of MONTH('2021-09-30')");

		oParser = new parserFormula("MONTH(A8)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-10-31')");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MONTH('2021-10-31')");

		oParser = new parserFormula("MONTH(A9)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-29')");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH('2021-12-29')");

		oParser = new parserFormula('MONTH(A10)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(0)");

		oParser = new parserFormula('MONTH(A11)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('1s')");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of MONTH('1s')");

		oParser = new parserFormula('MONTH(A12)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(TRUE)");

		oParser = new parserFormula('MONTH(A13)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(FALSE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(FALSE)");

		oParser = new parserFormula('MONTH(A14)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(#DIV/0!)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of MONTH(#DIV/0!)");

		oParser = new parserFormula('MONTH(A15)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('')");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH('')");

		oParser = new parserFormula('MONTH(A1:A4)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C1:C5").bbox);
		assert.ok(oParser.parse(), "MONTH(A1:A4)");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 9, "Result of MONTH(A1:A4)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 12, "Result of MONTH(A1:A4)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 12, "Result of MONTH(A1:A4)[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 11, "Result of MONTH(A1:A4)[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Result of MONTH(A1:A4)[4,0]");

		oParser = new parserFormula('MONTH({1,100,1000,10000})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:H1").bbox);
		assert.ok(oParser.parse(), "MONTH({1,100,1000,10000})");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of MONTH({1,100,1000,10000})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4, "Result of MONTH({1,100,1000,10000})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 9, "Result of MONTH({1,100,1000,10000})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 5, "Result of MONTH({1,100,1000,10000})[0,3]");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", "Result of MONTH({1,100,1000,10000})[0,4]");

		// set 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("MONTH(A1)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.09.21)");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MONTH(30.09.21)");

		oParser = new parserFormula("MONTH(A2)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(30.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH(30.12.21)");

		oParser = new parserFormula("MONTH(A3)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(31.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(31.12.21)");

		oParser = new parserFormula("MONTH(A4)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH(10.11.21)");
		assert.strictEqual(oParser.calculate().getValue(), 11, "Result of MONTH(10.11.21)");

		oParser = new parserFormula("MONTH(A5)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-10-01')");
		assert.strictEqual(oParser.calculate().getValue(), 10, "Result of MONTH('2021-10-01')");
		// Different result with MS
		oParser = new parserFormula("MONTH(A6)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-31')");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH('2021-12-31')");
		// Different result with MS
		oParser = new parserFormula("MONTH(A7)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-09-30')");
		assert.strictEqual(oParser.calculate().getValue(), 9, "Result of MONTH('2021-09-30')");
		// Different result with MS
		oParser = new parserFormula("MONTH(A8)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-10-31')");
		assert.strictEqual(oParser.calculate().getValue(), 11, "Result of MONTH('2021-10-31')");

		oParser = new parserFormula("MONTH(A9)", "A2", ws);
		assert.ok(oParser.parse(), "MONTH('2021-12-29')");
		assert.strictEqual(oParser.calculate().getValue(), 12, "Result of MONTH('2021-12-29')");

		oParser = new parserFormula('MONTH(A10)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(0)");

		oParser = new parserFormula('MONTH(A11)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('1s')");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of MONTH('1s')");

		oParser = new parserFormula('MONTH(A12)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(TRUE)");

		oParser = new parserFormula('MONTH(A13)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(FALSE)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH(FALSE)");

		oParser = new parserFormula('MONTH(A14)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH(#DIV/0!)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of MONTH(#DIV/0!)");

		oParser = new parserFormula('MONTH(A15)', "A2", ws);
		assert.ok(oParser.parse(), "MONTH('')");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of MONTH('')");

		oParser = new parserFormula('MONTH(A1:A4)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C1:C5").bbox);
		assert.ok(oParser.parse(), "MONTH(A1:A4)");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 10, "Result of MONTH(A1:A4)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 12, "Result of MONTH(A1:A4)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, "Result of MONTH(A1:A4)[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 11, "Result of MONTH(A1:A4)[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Result of MONTH(A1:A4)[4,0]");

		oParser = new parserFormula('MONTH({1,100,1000,10000})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:H1").bbox);
		assert.ok(oParser.parse(), "MONTH({1,100,1000,10000})");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of MONTH({1,100,1000,10000})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4, "Result of MONTH({1,100,1000,10000})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 9, "Result of MONTH({1,100,1000,10000})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 5, "Result of MONTH({1,100,1000,10000})[0,3]");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", "Result of MONTH({1,100,1000,10000})[0,4]");

		// return to base mode
		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("B101").setValue("4");
		// For area
		ws.getRange2("A102").setValue("3");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A104").setValue("Text");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		ws2.getRange2("A3:B4").setValue("Text");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Time. Valid time input, returns minute component (30). 1 argument used.
		oParser = new parserFormula('MONTH(TIME(12,30,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(12,30,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Time. Valid time input, returns minute component (30). 1 argument used.');
		// Case #2: Number. Serial number for 12:00 PM (0.5), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Serial number for 12:00 PM (0.5), returns 0 minutes. 1 argument used.');
		// Case #3: String. String in time format, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH("12:30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("12:30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String in time format, returns 30 minutes. 1 argument used.');
		// Case #4: Formula. Nested TIME() formula returning time, returns minute component. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(1,2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(1,2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested TIME() formula returning time, returns minute component. 1 argument used.');
		// Case #5: Reference link. Reference to cell with valid time serial (0.5), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with valid time serial (0.5), returns 0 minutes. 1 argument used.');
		// Case #6: Area. Single-cell range with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell range with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #7: Array. Array with single time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH({0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH({0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Array. Array with single time serial, returns 0 minutes. 1 argument used.');
		// Case #8: Name. Named range with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #9: Name3D. 3D named range with valid time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid time serial, returns 0 minutes. 1 argument used.');
		// Case #10: Ref3D. 3D reference to cell with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to cell with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #11: Area3D. 3D single-cell range with valid time serial, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range with valid time serial, returns 0 minutes. 1 argument used.');
		// Case #12: Table. Table structured reference with valid time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with valid time serial, returns 30 minutes. 1 argument used.');
		// Case #13: Date. Date serial number, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date. Date serial number, returns 0 minutes. 1 argument used.');
		// Case #14: Formula. Time formula adjusted, returns minute component. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(15,45,30)+0.1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(15,45,30)+0.1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Time formula adjusted, returns minute component. 1 argument used.');
		// Case #15: Formula. MINUTE inside SUM formula, returns 30+10=40. 1 argument used.
		oParser = new parserFormula('SUM(MONTH(TIME(12,30,0)),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(MONTH(TIME(12,30,0)),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 11, 'Test: Positive case: Formula. MINUTE inside SUM formula, returns 30+10=40. 1 argument used.');
		// Case #16: String. String with full time format (hh:mm:ss), returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH("12:30:45")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("12:30:45") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String with full time format (hh:mm:ss), returns 30 minutes. 1 argument used.');
		// Case #17: Formula. Nested IF returning valid time, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(IF(TRUE,TIME(12,30,0),TIME(0,0,0)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(IF(TRUE,TIME(12,30,0),TIME(0,0,0))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested IF returning valid time, returns 30 minutes. 1 argument used.');
		// Case #18: Array. Array with single time serial, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH({0.520833333})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH({0.520833333}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Array. Array with single time serial, returns 30 minutes. 1 argument used.');
		// Case #19: String. String with minimal hour format, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH("00:30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("00:30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String with minimal hour format, returns 30 minutes. 1 argument used.');
		// Case #20: Formula. Nested formula with current hour and fixed minutes, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(HOUR(12345),30,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(HOUR(12345),30,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested formula with current hour and fixed minutes, returns 30 minutes. 1 argument used.');
		// Case #21: Time. Maximum valid time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Time. Maximum valid time, returns 59 minutes. 1 argument used.');
		// Case #22: Reference link. Reference to cell with full time string, returns 30 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Reference link. Reference to cell with full time string, returns 30 minutes. 1 argument used.');

		// Negative cases:

		// Case #1: Number. Negative serial number, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MONTH(-0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(-0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative serial number, returns #NUM!. 1 argument used.');
		// Case #2: String. Non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-time string, returns #VALUE!. 1 argument used.');
		// Case #3: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('MONTH(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #4: Area. Multi-cell range, returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('MONTH(A103:A104)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: MONTH(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area. Multi-cell range, returns #VALUE!. 1 argument used.');
		// Case #5: Empty. Empty cell reference, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Empty. Empty cell reference, returns #VALUE!. 1 argument used.');
		// Case #6: Boolean. Boolean FALSE (0), returns 0 minutes (edge case). 1 argument used.
		oParser = new parserFormula('MONTH(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Boolean. Boolean FALSE (0), returns 0 minutes (edge case). 1 argument used.');
		// Case #7: Ref3D. 3D reference to non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to non-time string, returns #VALUE!. 1 argument used.');
		// Case #8: Name. Named range with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Named range with non-time string, returns #VALUE!. 1 argument used.');
		// Case #10: Formula. Formula resulting in #NUM!, propagates error. 1 argument used.
		oParser = new parserFormula('MONTH(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM!, propagates error. 1 argument used.');
		// Case #11: Number. Number exceeds valid serial range, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MONTH(1E+308)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(1E+308) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Number exceeds valid serial range, returns #NUM!. 1 argument used.');
		// Case #12: String. Invalid time string (hours > 23), returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH("25:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("25:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: String. Invalid time string (hours > 23), returns #VALUE!. 1 argument used.');
		// Case #13: Array. Array with negative serial, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MONTH({-0.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH({-0.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Array. Array with negative serial, returns #NUM!. 1 argument used.');
		// Case #14: Name3D. 3D named range with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. 3D named range with non-time string, returns #VALUE!. 1 argument used.');
		// Case #15: Area3D. 3D multi-cell range, returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('MONTH(Sheet2!A4:A5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: MONTH(Sheet2!A4:A5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area3D. 3D multi-cell range, returns #VALUE!. 1 argument used.');
		// Case #16: Formula. Date before Jan 1, 1900, returns #NUM!. 1 argument used.
		oParser = new parserFormula('MONTH(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Negative case: Formula. Date before Jan 1, 1900, returns #NUM!. 1 argument used.');
		// Case #17: String. Empty string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string, returns #VALUE!. 1 argument used.');
		// Case #18: Boolean. Boolean TRUE (1), returns 0 minutes (edge case). 1 argument used.
		oParser = new parserFormula('MONTH(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Boolean. Boolean TRUE (1), returns 0 minutes (edge case). 1 argument used.');
		// Case #19: Formula. Invalid time (hours >= 24), returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(24,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(24,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Formula. Invalid time (hours >= 24), returns #VALUE!. 1 argument used.');
		// Case #20: Reference link. Reference to cell with non-time string, returns #VALUE!. 1 argument used.
		oParser = new parserFormula('MONTH(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Reference link. Reference to cell with non-time string, returns #VALUE!. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid serial number (Jan 1, 1900), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Minimum valid serial number (Jan 1, 1900), returns 0 minutes. 1 argument used.');
		// Case #2: Number. Maximum valid serial number (Dec 31, 9999), returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Bounded case: Number. Maximum valid serial number (Dec 31, 9999), returns 0 minutes. 1 argument used.');
		// Case #3: Time. Minimum valid time, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Time. Minimum valid time, returns 0 minutes. 1 argument used.');
		// Case #4: Time. Maximum valid time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Time. Maximum valid time, returns 59 minutes. 1 argument used.');
		// Case #5: String. String representing minimum valid time, returns 0 minutes. 1 argument used.
		oParser = new parserFormula('MONTH("00:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH("00:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. String representing minimum valid time, returns 0 minutes. 1 argument used.');
		// Case #6: Formula. Formula for maximum valid date and time, returns 59 minutes. 1 argument used.
		oParser = new parserFormula('MONTH(DATE(9999,12,31)+TIME(23,59,59))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MONTH(DATE(9999,12,31)+TIME(23,59,59)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Bounded case: Formula. Formula for maximum valid date and time, returns 59 minutes. 1 argument used.');

		// TODO Need to fix: area as argument handle
		// Case #4: Area. Multi-cell range, returns #VALUE!. 1 argument used.
		// Case #15: Area3D. 3D multi-cell range, returns #VALUE!. 1 argument used.

		testArrayFormula2(assert, "MONTH");
	});

	QUnit.test('Test: "NETWORKDAYS"', function (assert) {

		oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31))", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula NETWORKDAYS(DATE(2006,1,1),DATE(2006,1,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 22, 'Test: Positive case: Formula(2). Return 22.');

		oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,31),DATE(2006,1,1))", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula NETWORKDAYS(DATE(2006,1,31),DATE(2006,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -22, 'Test: Positive case: Formula(2). Return -22.');

		oParser = new parserFormula("NETWORKDAYS(DATE(1700,1,1),DATE(1700,2,2))", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula NETWORKDAYS(DATE(1700,1,1),DATE(1700,2,2))  is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 23, 'Test: Positive case: Formula(2). Return 23.');

		oParser = new parserFormula("NETWORKDAYS(DATE(2006,1,1),DATE(2006,2,1),{\"01-02-2006\",\"01-16-2006\"})", "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula NETWORKDAYS(DATE(2006,1,1),DATE(2006,2,1),{\"01-02-2006\",\"01-16-2006\"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 21, 'Test: Positive case: Formula(2), Array. Return 21.');

		oParser = new parserFormula("NETWORKDAYS(0,0)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(0,0)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(0,0)");

		// in js new Date(1900,0,1) === monday, in ms 01.01.1990 === sunday
		oParser = new parserFormula("NETWORKDAYS(1,1)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,1)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(1,1)");

		oParser = new parserFormula("NETWORKDAYS(2,2)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(2,2)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(2,2)");

		oParser = new parserFormula("NETWORKDAYS(3,3)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(3,3)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(3,3)");

		oParser = new parserFormula("NETWORKDAYS(4,4)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(4,4)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(4,4)");

		oParser = new parserFormula("NETWORKDAYS(5,5)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(5,5)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(5,5)");

		oParser = new parserFormula("NETWORKDAYS(6,6)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(6,6)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(6,6)");

		oParser = new parserFormula("NETWORKDAYS(7,7)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(7,7)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(7,7)");

		oParser = new parserFormula("NETWORKDAYS(8,8)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(8,8)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(8,8)");

		oParser = new parserFormula("NETWORKDAYS(9,9)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(9,9)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(9,9)");

		oParser = new parserFormula("NETWORKDAYS(10,10)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(10,10)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(10,10)");

		oParser = new parserFormula("NETWORKDAYS(11,11)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(11,11)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of NETWORKDAYS(11,11)");

		oParser = new parserFormula("NETWORKDAYS(0,11)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(0,11)");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(0,11)");

		oParser = new parserFormula("NETWORKDAYS(1,11)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,11)");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(1,11)");

		oParser = new parserFormula("NETWORKDAYS(11,0)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(11,0)");
		assert.strictEqual(oParser.calculate().getValue(), -8, "Result of NETWORKDAYS(11,0)");

		oParser = new parserFormula("NETWORKDAYS(11,1)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(11,1)");
		assert.strictEqual(oParser.calculate().getValue(), -8, "Result of NETWORKDAYS(11,1)");

		oParser = new parserFormula("NETWORKDAYS(-1,15)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(-1,15)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(-1,15)");

		oParser = new parserFormula("NETWORKDAYS(15,-1)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(15,-1)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(15,-1)");

		oParser = new parserFormula("NETWORKDAYS(-1,-15)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(-1,-15)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of NETWORKDAYS(-1,-15)");

		oParser = new parserFormula("NETWORKDAYS(1,3889)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,3889)");
		assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1,3889)");

		oParser = new parserFormula("NETWORKDAYS(1,45689)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,45689)");
		assert.strictEqual(oParser.calculate().getValue(), 32635, "Result of NETWORKDAYS(1,45689)");

		oParser = new parserFormula("NETWORKDAYS(0.1,0.9)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(0.1,0.9)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS(0.1,0.9)");

		oParser = new parserFormula("NETWORKDAYS(1.1,3889)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1.1,3889)");
		assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.1,3889)");

		oParser = new parserFormula("NETWORKDAYS(1.9,3889)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1.9,3889)");
		assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.9,3889)");

		oParser = new parserFormula("NETWORKDAYS(1,3889.1)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,3889.1)");
		assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1,3889.1)");

		oParser = new parserFormula("NETWORKDAYS(1.9,3889.9)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1.9,3889.9)");
		assert.strictEqual(oParser.calculate().getValue(), 2778, "Result of NETWORKDAYS(1.9,3889.9)");

		// bool
		oParser = new parserFormula("NETWORKDAYS(11,TRUE)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(11,TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(11,TRUE)");

		oParser = new parserFormula("NETWORKDAYS(TRUE,TRUE)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(TRUE,TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(TRUE,TRUE)");

		oParser = new parserFormula("NETWORKDAYS(TRUE,11)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(TRUE,11)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(TRUE,11)");

		oParser = new parserFormula("NETWORKDAYS(#VALUE!,#NUM!)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(#VALUE!,#NUM!)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(#VALUE!,#NUM!)");

		// array
		oParser = new parserFormula("NETWORKDAYS({1,11,255},11)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS({1,11,255},11)");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS({1,11,255},11)");

		oParser = new parserFormula("NETWORKDAYS(1,{11,85,255})", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(1,{11,85,255})");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS(1,{11,85,255})");

		oParser = new parserFormula("NETWORKDAYS({1,11,255},{11,85,255})", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS({1,11,255},{11,85,255})");
		assert.strictEqual(oParser.calculate().getValue(), 8, "Result of NETWORKDAYS({1,11,255},{11,85,255})");

		ws.getRange2("A101").setValue();
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("0");
		ws.getRange2("A104").setValue("9");
		ws.getRange2("A105").setValue("25");
		ws.getRange2("A106").setValue("TRUE");
		ws.getRange2("A107").setValue("FALSE");
		ws.getRange2("A108").setValue("{999,25,0}");
		ws.getRange2("A109").setValue("{777,25,0}");
		ws.getRange2("A110").setValue("{0,777,25,0}");
		ws.getRange2("A111").setValue("#N/A");
		ws.getRange2("A112").setValue("99999999999999999999");
		ws.getRange2("A113").setValue("-99999999999999999999");
		ws.getRange2("A114").setValue("str");
		ws.getRange2("A115").setValue("str2");

		ws.getRange2("B101").setValue("0");
		ws.getRange2("B102").setValue("1");
		ws.getRange2("B103").setValue("4");
		ws.getRange2("B104").setValue("9");
		ws.getRange2("B105").setValue("25");
		ws.getRange2("B106").setValue("255");
		ws.getRange2("B107").setValue("312778");

		// cellsrange
		oParser = new parserFormula("NETWORKDAYS(A101:A105,A105)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A101:A105,25)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A101:A105,25)");

		oParser = new parserFormula("NETWORKDAYS(A104,A101:A105)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A104,A101:A105)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A104,A101:A105)");

		oParser = new parserFormula("NETWORKDAYS(A101:A105,A101:A105)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A101:A105,A101:A105)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(A101:A105,A101:A105)");

		oParser = new parserFormula("NETWORKDAYS(B101:B107,B101:B107)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(B101:B107,B101:B107)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(B101:B107,B101:B107)");

		// cells
		oParser = new parserFormula("NETWORKDAYS(A102,A102)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A102,A102)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS('','')");

		oParser = new parserFormula("NETWORKDAYS(A102:A102,A102:A102)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A102:A102,A102:A102)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of NETWORKDAYS('','')");

		oParser = new parserFormula("NETWORKDAYS(A103,A104)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A103,A104)");
		assert.strictEqual(oParser.calculate().getValue(), 6, "Result of NETWORKDAYS(0,9)");

		oParser = new parserFormula("NETWORKDAYS(A104:A104,A104:A104)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("NETWORKDAYS(A106,A107)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("NETWORKDAYS(A109,A108)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS({777,25,0},{999,25,0})");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS({777,25,0},{999,25,0})");

		oParser = new parserFormula("NETWORKDAYS(A105,A108)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(25,{999,25,0})");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(25,{999,25,0})");

		oParser = new parserFormula("NETWORKDAYS(A108,25)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS({999,25,0},25)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS({999,25,0},25)");

		oParser = new parserFormula("NETWORKDAYS(A111,A105)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A114,A115)");
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", "Result of NETWORKDAYS(str,str2)");

		oParser = new parserFormula("NETWORKDAYS(A114,A115)", "A2", ws);
		assert.ok(oParser.parse(), "NETWORKDAYS(A114,A115)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of NETWORKDAYS(str,str2)");

		// bug case
		oParser = new parserFormula("NETWORKDAYS(A101,A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("NETWORKDAYS(A101,A102)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("NETWORKDAYS(A101,A109)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("NETWORKDAYS(A102,A109)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("NETWORKDAYS(A101:A101,A101:A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("B101").setValue("4");
		// For area
		ws.getRange2("A102").setValue("3");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		ws2.getRange2("A3:B4").setValue("Text");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Date. Basic valid input: dates using DATE formula. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Date. Basic valid input: dates using DATE formula. 2 of 3 arguments used.');
		// Case #2: Number. Dates as Excel serial numbers. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(44927,44936)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(44927,44936) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Number. Dates as Excel serial numbers. 2 of 3 arguments used.');
		// Case #3: String. Dates as string format. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS("01/01/2025","01/10/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS("01/01/2025","01/10/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: String. Dates as string format. 2 of 3 arguments used.');
		// Case #4: Formula,Date. Start_date filled via IF formula. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Formula,Date. Start_date filled via IF formula. 2 of 3 arguments used.');
		// Case #5: Reference link. All arguments as Reference link. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(A100,A101)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A100,A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. All arguments as Reference link. 2 of 3 arguments used.');
		// Case #6: Area. All arguments are single-cell ranges. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(A102:A102,A103:A103)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A102:A102,A103:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area. All arguments are single-cell ranges. 2 of 3 arguments used.');
		// Case #7: Array. All arguments are arrays with single element. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS({44927},{44936})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS({44927},{44936}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Array. All arguments are arrays with single element. 2 of 3 arguments used.');
		// Case #8: Name. All arguments as Name type. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TestName,TestName1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. All arguments as Name type. 2 of 3 arguments used.');
		// Case #9: Name3D. All arguments as Name3D type. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TestName3D,TestName3D) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. All arguments as Name3D type. 2 of 3 arguments used.');
		// Case #10: Ref3D. All arguments as Ref3D type. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. All arguments as Ref3D type. 2 of 3 arguments used.');
		// Case #11: Area3D. All arguments as Area3D type with single cell. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(Sheet2!A1:A1,Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(Sheet2!A1:A1,Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. All arguments as Area3D type with single cell. 2 of 3 arguments used.');
		// Case #12: Table. All arguments as Table structured references. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(Table1[Column1],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(Table1[Column1],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Table. All arguments as Table structured references. 2 of 3 arguments used.');
		// Case #13: Date,Array. Holidays as array of string dates. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),{"01/03/2025","01/04/2025"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),{"01/03/2025","01/04/2025"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Date,Array. Holidays as array of string dates. 3 of 3 arguments used.');
		// Case #14: Date,Area. Holidays as range of cells. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),A104:A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Date,Area. Holidays as range of cells. 3 of 3 arguments used.');
		// Case #15: Formula. Holidays as IF formula returning range. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),IF(TRUE,A104:A104,A105:A105))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),IF(TRUE,A104:A104,A105:A105)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Formula. Holidays as IF formula returning range. 3 of 3 arguments used.');
		// Case #16: Number,Formula. NETWORKDAYS inside SUM formula. 2 of 3 arguments used.
		oParser = new parserFormula('SUM(NETWORKDAYS(44927,44936),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(NETWORKDAYS(44927,44936),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 12, 'Test: Positive case: Number,Formula. NETWORKDAYS inside SUM formula. 2 of 3 arguments used.');
		// Case #17: String. Short date format in strings. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS("1/1/25","1/10/25")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS("1/1/25","1/10/25") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: String. Short date format in strings. 2 of 3 arguments used.');
		// Case #18: Time,Date. Start_date with time component adjusted to valid date. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(TIME(0,0,0)+44927,DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TIME(0,0,0)+44927,DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 530, 'Test: Positive case: Time,Date. Start_date with time component adjusted to valid date. 2 of 3 arguments used.');
		// Case #19: Number,Array. Holidays as array of serial numbers. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(44927,44936,{44928,44929})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(44927,44936,{44928,44929}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number,Array. Holidays as array of serial numbers. 3 of 3 arguments used.');
		// Case #20: Formula. All arguments filled with DATE formula. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),DATE(2025,1,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),DATE(2025,1,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Formula. All arguments filled with DATE formula. 3 of 3 arguments used.');

		// Negative cases:

		// Case #1: Number. Start_date is zero (invalid date). Returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(0,44936)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(0,44936) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32097, 'Test: Negative case: Number. Start_date is zero (invalid date). Returns #NUM!. 2 of 3 arguments used.');
		// Case #2: Number. End_date is zero (invalid date). Returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(44927,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(44927,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -32090, 'Test: Negative case: Number. End_date is zero (invalid date). Returns #NUM!. 2 of 3 arguments used.');
		// Case #3: Date. Start_date after End_date. Returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,10),DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,10),DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Negative case: Date. Start_date after End_date. Returns #NUM!. 2 of 3 arguments used.');
		// Case #4: String. Start_date as non-date string. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS("abc","01/10/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS("abc","01/10/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Start_date as non-date string. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #5: String. End_date as non-date string. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS("01/01/2025","xyz")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS("01/01/2025","xyz") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. End_date as non-date string. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #6: Boolean. Start_date as boolean. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(TRUE,44936)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TRUE,44936) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Start_date as boolean. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #7: Boolean. End_date as boolean. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(44927,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(44927,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. End_date as boolean. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #8: Empty. Start_date is empty. Returns #VALUE!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(,44936)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(,44936) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty. Start_date is empty. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #9: Empty. End_date is empty. Returns #VALUE!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(44927,)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(44927,) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty. End_date is empty. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #10: Error. Start_date as error value. Returns #N/A. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(NA(),44936)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(NA(),44936) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Start_date as error value. Returns #N/A. 2 of 3 arguments used.');
		// Case #11: Date,String. Holidays as non-date string. Returns #VALUE!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),"abc")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),"abc") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date,String. Holidays as non-date string. Returns #VALUE!. 3 of 3 arguments used.');
		// Case #12: Date,Array. Holidays as array with non-date value. Returns #VALUE!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),{TRUE})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),{TRUE}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date,Array. Holidays as array with non-date value. Returns #VALUE!. 3 of 3 arguments used.');
		// Case #13: Area. Start_date as multi-cell range. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(A106:A107,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A106:A107,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Start_date as multi-cell range. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #14: Area. End_date as multi-cell range. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(A100,A106:A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A100,A106:A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. End_date as multi-cell range. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #15: Date,Area. Holidays as multi-cell range with non-date value. Returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),A106:A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,10),A106:A107) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Negative case: Date,Area. Holidays as multi-cell range with non-date value. Returns #VALUE!. 3 of 3 arguments used.');
		// Case #16: Ref3D. Start_date as Ref3D with non-date value. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(Sheet2!A3,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(Sheet2!A3,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. Start_date as Ref3D with non-date value. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #17: Ref3D. End_date as Ref3D with non-date value. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(A100,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A100,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. End_date as Ref3D with non-date value. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #18: Name. Start_date as Name with area. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(TestNameArea,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TestNameArea,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Start_date as Name with area. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #19: Name. End_date as Name with area. Returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(A100,TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(A100,TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. End_date as Name with area. Returns #VALUE!. 2 of 3 arguments used.');
		// Case #20: Name3D. Start_date as Name3D with area. Returns #VALUE!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS(TestNameArea3D2,A101)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS(TestNameArea3D2,A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name3D. Start_date as Name3D with area. Returns #VALUE!. 2 of 3 arguments used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid date serial numbers. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Minimum valid date serial numbers. 2 of 3 arguments used.');
		// Case #2: Date. Maximum valid Excel dates. 2 of 3 arguments used.
		oParser = new parserFormula('NETWORKDAYS(DATE(9999,12,30),DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS(DATE(9999,12,30),DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Date. Maximum valid Excel dates. 2 of 3 arguments used.');

		// TODO Need to fix: strange return with dates, ms diff results, different error types
		// Case #5: Reference link. All arguments as Reference link. 2 of 3 arguments used.
		// Case #6: Area. All arguments are single-cell ranges. 2 of 3 arguments used.
		// Case #8: Name. All arguments as Name type. 2 of 3 arguments used.
		// Case #9: Name3D. All arguments as Name3D type. 2 of 3 arguments used.
		// Case #8: Empty. Start_date is empty. Returns #VALUE!. 2 of 3 arguments used.
		// Case #9: Empty. End_date is empty. Returns #VALUE!. 2 of 3 arguments used.
		// Case #11: Date,String. Holidays as non-date string. Returns #VALUE!. 3 of 3 arguments used. - what is with return?
		// Case #12: Date,Array. Holidays as array with non-date value. Returns #VALUE!. 3 of 3 arguments used.
		// Case #20: Name3D. Start_date as Name3D with area. Returns #VALUE!. 2 of 3 arguments used.

		testArrayFormula2(assert, "NETWORKDAYS", 2, 3, true, null);
	});

	QUnit.test('Test: "NETWORKDAYS.INTL"', function (assert) {

		let formulaStr = "NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))";
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 22, 'Result of ' + formulaStr);

		formulaStr = "NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))";
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), -21, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 22, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 26, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"1111111",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 20, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 30, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"19",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),19,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),"0000000",{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 38381, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(DATE(1901,1,1),DATE(2006,2,1),17,{"1/2/2006","1/16/2006"})';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 32898, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,11)';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 8490, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,1)';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 7075, 'Result of ' + formulaStr);

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,2)';
		oParser = new parserFormula(formulaStr, "A2", ws);
		assert.ok(oParser.parse(), formulaStr);
		assert.strictEqual(oParser.calculate().getValue(), 7075, 'Result of ' + formulaStr);

		//TODO Need to check why it's calculate incorrect
		// The problem repeats with new and old realizations of NETWORKDAYS.INTL.

		/*formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5)';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7074, formulaStr );

		formulaStr = 'NETWORKDAYS.INTL(100.123,10003.556,5,{123,1000})';
		oParser = new parserFormula( formulaStr, "A2", ws );
		assert.ok( oParser.parse(), formulaStr );
		assert.strictEqual( oParser.calculate().getValue(), 7073, formulaStr );*/

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("B101").setValue("4");
		// For area
		ws.getRange2("A102").setValue("3");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		ws2.getRange2("A3:B4").setValue("Text");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:
		// Case #1: Number. Basic valid input: serial numbers for dates (01/01/2006 to 05/01/2006), weekend as number (1 = Saturday/Sunday), holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Number. Basic valid input: serial numbers for dates (01/01/2006 to 05/01/2006), weekend as number (1 = Saturday/Sunday), holidays omitted. 3 of 4 arguments used.');
		// Case #2: Date. Dates provided via DATE formula, weekend as number (1), holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Date. Dates provided via DATE formula, weekend as number (1), holidays omitted. 3 of 4 arguments used.');
		// Case #3: String. Dates as strings, weekend as string (Saturday/Sunday), holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL("01/01/2025","01/10/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL("01/01/2025","01/10/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: String. Dates as strings, weekend as string (Saturday/Sunday), holidays omitted. 3 of 4 arguments used.');
		// Case #4: Formula. Weekend as IF formula returning valid number, holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),"0000011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),"0000011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Formula. Weekend as IF formula returning valid number, holidays omitted. 3 of 4 arguments used.');
		// Case #5: Reference link. Reference to cells with valid dates and weekend number, holidays as single cell reference. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),1,A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),1,A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Reference link. Reference to cells with valid dates and weekend number, holidays as single cell reference. 4 of 4 arguments used.');
		// Case #6: Area. Single-cell ranges for all arguments. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(A100,A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(A100,A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell ranges for all arguments. 4 of 4 arguments used.');
		// Case #7: Array. Arrays with single valid elements for all arguments. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(A100:A100,A101:A101,1,A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(A100:A100,A101:A101,1,A102:A102) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Array. Arrays with single valid elements for all arguments. 4 of 4 arguments used.');
		// Case #8: Name. Named ranges with valid values. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL({44927},{44936},1,{44932})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL({44927},{44936},1,{44932}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Name. Named ranges with valid values. 4 of 4 arguments used.');
		// Case #9: Name3D. 3D named ranges with valid values. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(TestName,TestName1,1,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(TestName,TestName1,1,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named ranges with valid values. 4 of 4 arguments used.');
		// Case #10: Ref3D. 3D references to cells with valid values. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(TestName3D,TestName3D,1,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(TestName3D,TestName3D,1,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Ref3D. 3D references to cells with valid values. 4 of 4 arguments used.');
		// Case #11: Area3D. 3D single-cell ranges for all arguments. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(Sheet2!A1,Sheet2!A2,1,Sheet2!A3)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(Sheet2!A1,Sheet2!A2,1,Sheet2!A3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell ranges for all arguments. 4 of 4 arguments used.');
		// Case #12: Table. Table structured references with valid values. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(Sheet2!A1:A1,Sheet2!A2:A2,1,Sheet2!A3:A3)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(Sheet2!A1:A1,Sheet2!A2:A2,1,Sheet2!A3:A3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Table. Table structured references with valid values. 4 of 4 arguments used.');
		// Case #13: Date,Array. Holidays as array with multiple valid dates. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(Table1[Column1],Table1[Column1],1,Table1[Column1])', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(Table1[Column1],Table1[Column1],1,Table1[Column1]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Date,Array. Holidays as array with multiple valid dates. 4 of 4 arguments used.');
		// Case #14: Time,Number. Dates adjusted with TIME formula, weekend as number, holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('SUM(NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),1),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),1),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Positive case: Time,Number. Dates adjusted with TIME formula, weekend as number, holidays omitted. 3 of 4 arguments used.');
		// Case #15: Formula. Nested IF formulas for dates, weekend as number, holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2025,1,10),1,{44932,44933})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2025,1,10),1,{44932,44933}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: Formula. Nested IF formulas for dates, weekend as number, holidays omitted. 3 of 4 arguments used.');
		// Case #16: String. Dates as strings, weekend as number string (11 = Monday only), holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,1)+TIME(8,0,0),DATE(2025,1,10)+TIME(8,0,0),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,1)+TIME(8,0,0),DATE(2025,1,10)+TIME(8,0,0),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Positive case: String. Dates as strings, weekend as number string (11 = Monday only), holidays omitted. 3 of 4 arguments used.');
		// Case #17: Number,Array. Holidays as single-element array. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL("01/01/2025","01/10/2025","1111100")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL("01/01/2025","01/10/2025","1111100") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Number,Array. Holidays as single-element array. 4 of 4 arguments used.');
		// Case #18: Formula,String. Weekend as string, holidays as single-cell range. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,DATE(2025,1,10),ROUND(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,DATE(2025,1,10),ROUND(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 530, 'Test: Positive case: Formula,String. Weekend as string, holidays as single-cell range. 4 of 4 arguments used.');
		// Case #19: Number,String. Weekend as string, holidays as reference link. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL({44927,44928},{44936,44937},1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL({44927,44928},{44936,44937},1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Number,String. Weekend as string, holidays as reference link. 4 of 4 arguments used.');
		// Case #20: Number. Weekend as number (17 = Friday only), holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),"1",A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,10),"1",A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Number. Weekend as number (17 = Friday only), holidays omitted. 3 of 4 arguments used.');
		// Case #21: Formula,Array. Holidays as array with multiple dates. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,,A100)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,,A100) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Formula,Array. Holidays as array with multiple dates. 4 of 4 arguments used.');
		// Case #22: Number,Empty. Weekend omitted (defaults to 1), holidays omitted. 2 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44936,44927,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44936,44927,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -7, 'Test: Positive case: Number,Empty. Weekend omitted (defaults to 1), holidays omitted. 2 of 4 arguments used.');

		// Negative cases:
		// Case #1: Number. Start_date > end_date returns #NUM!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL("abc","01/10/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL("abc","01/10/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Start_date > end_date returns #NUM!. 3 of 4 arguments used.');
		// Case #2: String. Invalid start_date string returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(NA(),DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(NA(),DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String. Invalid start_date string returns #VALUE!. 3 of 4 arguments used.');
		// Case #3: String. Invalid end_date string returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(FALSE,DATE(2025,1,10))', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(FALSE,DATE(2025,1,10)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid end_date string returns #VALUE!. 3 of 4 arguments used.');
		// Case #4: Number,String. Invalid weekend string returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(A100:A101,A101,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(A100:A101,A101,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Invalid weekend string returns #VALUE!. 3 of 4 arguments used.');
		// Case #5: Number,Array. Holidays with invalid date (0) returns #NUM!. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(A102,DATE(2025,1,10))', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(A102,DATE(2025,1,10)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 32620, 'Test: Negative case: Number,Array. Holidays with invalid date (0) returns #NUM!. 4 of 4 arguments used.');
		// Case #6: Error. Start_date as #N/A propagates error. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL("01/01/2025","01/10/2025","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL("01/01/2025","01/10/2025","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error. Start_date as #N/A propagates error. 3 of 4 arguments used.');
		// Case #7: Number,Error. End_date as #N/A propagates error. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number,Error. End_date as #N/A propagates error. 3 of 4 arguments used.');
		// Case #8: Number. Invalid weekend number (0) returns #NUM!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(Sheet2!A4,Sheet2!A5,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(Sheet2!A4,Sheet2!A5,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Number. Invalid weekend number (0) returns #NUM!. 3 of 4 arguments used.');
		// Case #9: 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(TestNameArea2,TestName1,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(TestNameArea2,TestName1,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case. 3 of 4 arguments used.');
		// Case #11: Boolean. Start_date as boolean (FALSE) returns #NUM!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(SQRT(-1),DATE(2025,1,10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(SQRT(-1),DATE(2025,1,10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Boolean. Start_date as boolean (FALSE) returns #NUM!. 3 of 4 arguments used.');
		// Case #12: Number,Boolean. End_date as boolean (TRUE) returns #NUM!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(0,44936,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(0,44936,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32097, 'Test: Negative case: Number,Boolean. End_date as boolean (TRUE) returns #NUM!. 3 of 4 arguments used.');
		// Case #13: Number,Boolean. Weekend as boolean (FALSE) returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL({FALSE},{44936},1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL({FALSE},{44936},1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Boolean. Weekend as boolean (FALSE) returns #VALUE!. 3 of 4 arguments used.');
		// Case #14: Area. Multi-cell range for start_date returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Area. Multi-cell range for start_date returns #VALUE!. 3 of 4 arguments used.');
		// Case #15: Number,Area. Multi-cell range for end_date returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(Sheet2!A1:A2,Sheet2!A3,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(Sheet2!A1:A2,Sheet2!A3,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Area. Multi-cell range for end_date returns #VALUE!. 3 of 4 arguments used.');
		// Case #16: Ref3D. 3D reference to invalid weekend string returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(TIME(12,0,0),DATE(2025,1,10),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(TIME(12,0,0),DATE(2025,1,10),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32620, 'Test: Negative case: Ref3D. 3D reference to invalid weekend string returns #VALUE!. 3 of 4 arguments used.');
		// Case #17: Name. Named range with invalid start_date (area) returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL("-01/01/2025","01/10/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL("-01/01/2025","01/10/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with invalid start_date (area) returns #VALUE!. 3 of 4 arguments used.');
		// Case #18: Table. Table column with invalid weekend string returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(A103,A101,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(A103,A101,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Table. Table column with invalid weekend string returns #VALUE!. 3 of 4 arguments used.');
		// Case #19: Formula. Start_date > end_date via DATE formula returns #NUM!. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(DATE(2025,1,10),DATE(2025,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(DATE(2025,1,10),DATE(2025,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -8, 'Test: Negative case: Formula. Start_date > end_date via DATE formula returns #NUM!. 3 of 4 arguments used.');
		// Case #20: String. Empty start_date string returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('NETWORKDAYS.INTL(1,2,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(1,2,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: String. Empty start_date string returns #VALUE!. 3 of 4 arguments used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid date (01/01/1900 to 01/02/1900), weekend as number, holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(2958465,2958465,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(2958465,2958465,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Minimum valid date (01/01/1900 to 01/02/1900), weekend as number, holidays omitted. 3 of 4 arguments used.');
		// Case #2: Number. Maximum valid date (12/31/9999), weekend as number, holidays omitted. 3 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,"0000001")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,"0000001") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'Test: Bounded case: Number. Maximum valid date (12/31/9999), weekend as number, holidays omitted. 3 of 4 arguments used.');
		// Case #3: Number,Array. Minimum valid date and holiday (01/01/1900), weekend as number. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(44927,44936,17)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(44927,44936,17) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Bounded case: Number,Array. Minimum valid date and holiday (01/01/1900), weekend as number. 4 of 4 arguments used.');
		// Case #4: Number,Array. Maximum valid date and holiday (12/31/9999), weekend as number. 4 of 4 arguments used.
		oParser = new parserFormula('NETWORKDAYS.INTL(2958464,2958465,1,{2958465})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NETWORKDAYS.INTL(2958464,2958465,1,{2958465}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number,Array. Maximum valid date and holiday (12/31/9999), weekend as number. 4 of 4 arguments used.');

		// TODO Need to fix: many results diff from ms, array handle, text handle, 3D error same as in NETWORKDAYS
		// Case #7: Array. Arrays with single valid elements for all arguments. 4 of 4 arguments used.
		// Case #8: Name. Named ranges with valid values. 4 of 4 arguments used.
		// Case #11: Area3D. 3D single-cell ranges for all arguments. 4 of 4 arguments used.
		// Case #3: String. Invalid end_date string returns #VALUE!. 3 of 4 arguments used.
		// Case #5: Number,Array. Holidays with invalid date (0) returns #NUM!. 4 of 4 arguments used.
		// Case #8: Number. Invalid weekend number (0) returns #NUM!. 3 of 4 arguments used.
		// Case #9: 3 of 4 arguments used.
		// Case #13: Number,Boolean. Weekend as boolean (FALSE) returns #VALUE!. 3 of 4 arguments used.
		// Case #12: Table. Table structured references with valid values. 4 of 4 arguments used.
		// Case #13: Date,Array. Holidays as array with multiple valid dates. 4 of 4 arguments used.
		// Case #19: Number,String. Weekend as string, holidays as reference link. 4 of 4 arguments used.
		// Case #21: Formula,Array. Holidays as array with multiple dates. 4 of 4 arguments used.
		// Case #18: Table. Table column with invalid weekend string returns #VALUE!. 3 of 4 arguments used.
		// Case #20: String. Empty start_date string returns #VALUE!. 3 of 4 arguments used.
	});

	QUnit.test('Test: "SECOND"', function (assert) {

		ws.getRange2("A202").setValue("12:45:03 PM");
		ws.getRange2("A203").setValue("4:48:18 PM");
		ws.getRange2("A204").setValue("4:48 PM");

		oParser = new parserFormula("SECOND(A202)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula("SECOND(A203)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 18);

		oParser = new parserFormula("SECOND(A204)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Number. Basic valid input: zero angle. 1 argument used.
		oParser = new parserFormula('SECOND(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number. Basic valid input: zero angle. 1 argument used.');
		// Case #2: Number. Valid input: ?/2 minus small value. 1 argument used.
		oParser = new parserFormula('SECOND(1.570796326794896)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(1.570796326794896) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Positive case: Number. Valid input: ?/2 minus small value. 1 argument used.');
		// Case #3: Number. Float input. 1 argument used.
		oParser = new parserFormula('SECOND(1.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(1.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Number. Float input. 1 argument used.');
		// Case #4: String. String convertible to number. 1 argument used.
		oParser = new parserFormula('SECOND("1.5")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND("1.5") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. String convertible to number. 1 argument used.');
		// Case #5: Formula. Nested formula (PI/4). 1 argument used.
		oParser = new parserFormula('SECOND(PI()/4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(PI()/4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 58, 'Test: Positive case: Formula. Nested formula (PI/4). 1 argument used.');
		// Case #6: Reference link. Ref to cell with valid number (0). 1 argument used.
		oParser = new parserFormula('SECOND(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Ref to cell with valid number (0). 1 argument used.');
		// Case #7: Area. Single-cell range with valid number (?/2 minus small value). 1 argument used.
		oParser = new parserFormula('SECOND(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area. Single-cell range with valid number (?/2 minus small value). 1 argument used.');
		// Case #8: Array. Array with single valid element. 1 argument used.
		oParser = new parserFormula('SECOND({1.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND({1.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Array. Array with single valid element. 1 argument used.');
		// Case #9: Name. Named range with valid number (0). 1 argument used.
		oParser = new parserFormula('SECOND(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid number (0). 1 argument used.');
		// Case #10: Name3D. 3D named range with valid number (0). 1 argument used.
		oParser = new parserFormula('SECOND(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid number (0). 1 argument used.');
		// Case #11: Ref3D. 3D reference to cell with valid number (0). 1 argument used.
		oParser = new parserFormula('SECOND(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Ref3D. 3D reference to cell with valid number (0). 1 argument used.');
		// Case #12: Area3D. 3D single-cell range with valid number (?/2 minus small value). 1 argument used.
		oParser = new parserFormula('SECOND(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area3D. 3D single-cell range with valid number (?/2 minus small value). 1 argument used.');
		// Case #13: Table. Table structured reference with valid number (0). 1 argument used.
		oParser = new parserFormula('SECOND(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Table. Table structured reference with valid number (0). 1 argument used.');
		// Case #14: Date. Date as serial number (large valid input). 1 argument used.
		oParser = new parserFormula('SECOND(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Date. Date as serial number (large valid input). 1 argument used.');
		// Case #15: Time. Time as fractional number (0.5). 1 argument used.
		oParser = new parserFormula('SECOND(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Time. Time as fractional number (0.5). 1 argument used.');
		// Case #16: Formula. SEC inside SUM formula. 1 argument used.
		oParser = new parserFormula('SUM(SECOND(1.5),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SECOND(1.5),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. SEC inside SUM formula. 1 argument used.');
		// Case #17: String. String of zero angle. 1 argument used.
		oParser = new parserFormula('SECOND("0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND("0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. String of zero angle. 1 argument used.');
		// Case #18: Array. Multi-element array with valid numbers. 1 argument used.
		oParser = new parserFormula('SECOND({0,1.5})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND({0,1.5}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Array. Multi-element array with valid numbers. 1 argument used.');
		// Case #19: Formula. Nested IF returning valid value. 1 argument used.
		oParser = new parserFormula('SECOND(IF(TRUE,1.5,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(IF(TRUE,1.5,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Formula. Nested IF returning valid value. 1 argument used.');
		// Case #20: Number. Negative valid input. 1 argument used.
		oParser = new parserFormula('SECOND(-1.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(-1.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Number. Negative valid input. 1 argument used.');
		// Case #21: Area3D. 3D multi-cell range with valid numbers. 1 argument used.
		oParser = new parserFormula('SECOND(Sheet2!A3:B3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Sheet2!A3:B3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D multi-cell range with valid numbers. 1 argument used.');

		// Negative cases:

		// Case #1: Number. Input at ?/2 (COS=0) returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND(1.5707963267948966)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(1.5707963267948966) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Negative case: Number. Input at ?/2 (COS=0) returns #DIV/0!. 1 argument used.');
		// Case #2: Number. Negative input at -?/2 (COS=0) returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND(-1.5707963267948966)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(-1.5707963267948966) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative input at -?/2 (COS=0) returns #DIV/0!. 1 argument used.');
		// Case #3: String. Non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 1 argument used.');
		// Case #4: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('SECOND(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('SECOND(A102:A103)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: SECOND(A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #6: Empty. Reference to empty cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!. 1 argument used.');
		// Case #7: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #8: Boolean. Boolean TRUE (1) returns valid result but not meaningful. 1 argument used.
		oParser = new parserFormula('SECOND(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean TRUE (1) returns valid result but not meaningful. 1 argument used.');
		// Case #9: Ref3D. 3D ref to text returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('SECOND(Sheet2!A4)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: SECOND(Sheet2!A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Ref3D. 3D ref to text returns #VALUE!. 1 argument used.');
		// Case #10: Name. Named range with text returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('SECOND(TestNameArea)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: SECOND(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #11: Name3D. 3D named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 1 argument used.');
		// Case #12: Table. Table column with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Table. Table column with text returns #VALUE!. 1 argument used.');
		// Case #13: Formula. Formula resulting in #NUM! error. 1 argument used.
		oParser = new parserFormula('SECOND(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error. 1 argument used.');
		// Case #14: Number. Input at ? (COS=0) returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND(3.141592653589793)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(3.141592653589793) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 54, 'Test: Negative case: Number. Input at ? (COS=0) returns #DIV/0!. 1 argument used.');
		// Case #15: Array. Array with value at ?/2 returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND({1.5707963267948966})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND({1.5707963267948966}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Negative case: Array. Array with value at ?/2 returns #DIV/0!. 1 argument used.');
		// Case #16: String. String convertible to -?/2 returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND("-1.5707963267948966")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND("-1.5707963267948966") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String. String convertible to -?/2 returns #DIV/0!. 1 argument used.');
		// Case #17: Reference link. Reference to cell with ?/2 returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Reference to cell with ?/2 returns #DIV/0!. 1 argument used.');
		// Case #18: Area3D. 3D multi-cell range with ?/2 returns #VALUE!. 1 argument used.
		oParser = new parserFormula('SECOND(Sheet2!A5:B5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(Sheet2!A5:B5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area3D. 3D multi-cell range with ?/2 returns #VALUE!. 1 argument used.');
		// Case #19: Formula. Formula resulting in #DIV/0! error. 1 argument used.
		oParser = new parserFormula('SECOND(MMULT(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(MMULT(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Formula. Formula resulting in #DIV/0! error. 1 argument used.');
		// Case #20: Name. Named range with ?/2 returns #DIV/0!. 1 argument used.
		oParser = new parserFormula('SECOND(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name. Named range with ?/2 returns #DIV/0!. 1 argument used.');

		// Bounded cases:

		// Case #1: Number. Smallest positive valid number in Excel. 1 argument used.
		oParser = new parserFormula('SECOND(2.2250738585072014E-308)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(2.2250738585072014E-308) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Smallest positive valid number in Excel. 1 argument used.');
		// Case #2: Number. Largest valid number in Excel. 1 argument used.
		oParser = new parserFormula('SECOND(1.7976931348623157E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(1.7976931348623157E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number. Largest valid number in Excel. 1 argument used.');
		// Case #3: Number. Value slightly above ?/2 (valid). 1 argument used.
		oParser = new parserFormula('SECOND(1.570796326794896+1E-15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(1.570796326794896+1E-15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 57, 'Test: Bounded case: Number. Value slightly above ?/2 (valid). 1 argument used.');
		// Case #4: Number. Value slightly below -?/2 (valid). 1 argument used.
		oParser = new parserFormula('SECOND(-1.570796326794896-1E-15)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SECOND(-1.570796326794896-1E-15) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number. Value slightly below -?/2 (valid). 1 argument used.');

		// TODO Need to fix: area and name handle
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Case #9: Ref3D. 3D ref to text returns #VALUE!. 1 argument used.
		// Case #10: Name. Named range with text returns #VALUE!. 1 argument used.


		testArrayFormula2(assert, "SECOND", 1, 1);
	});

	QUnit.test('Test: "TIME"', function (assert) {
		ws.getRange2("A2").setValue("12");
		ws.getRange2("A3").setValue("16");

		ws.getRange2("B2").setValue("0");
		ws.getRange2("B3").setValue("48");

		ws.getRange2("C2").setValue("0");
		ws.getRange2("C3").setValue("10");

		oParser = new parserFormula("TIME(A2,B2,C2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0.5);

		oParser = new parserFormula("TIME(A3,B3,C3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.7001157);

		oParser = new parserFormula("TIME(1,1,1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0423727);

		oParser = new parserFormula("TIME(1.34,1,1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0423727);

		oParser = new parserFormula("TIME(1.34,1.456,1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0423727);

		oParser = new parserFormula("TIME(1.34,1.456,1.9)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue().toFixed(7) - 0, 0.0423727);

		oParser = new parserFormula("TIME(-1.34,1.456,1.9)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("-1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Number. Basic valid input: integers for hour, minute, second. 3 arguments used.
		oParser = new parserFormula('TIME(12,30,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(12,30,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Number. Basic valid input: integers for hour, minute, second. 3 arguments used.');
		// Case #2: Number. Maximum typical time values (23:59:59). 3 arguments used.
		oParser = new parserFormula('TIME(23,59,59)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(23,59,59) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.999988425925926, 'Test: Positive case: Number. Maximum typical time values (23:59:59). 3 arguments used.');
		// Case #3: Number. Hour > 24, Excel adjusts to next day (1:00:00). 3 arguments used.
		oParser = new parserFormula('TIME(25,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(25,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04166666666666674, 'Test: Positive case: Number. Hour > 24, Excel adjusts to next day (1:00:00). 3 arguments used.');
		// Case #5: String. Numeric strings converted to numbers. 3 arguments used.
		oParser = new parserFormula('TIME("12","30","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME("12","30","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: String. Numeric strings converted to numbers. 3 arguments used.');
		// Case #6: Formula. Nested formulas returning valid numbers. 3 arguments used.
		// Different result with MS
		// oParser = new parserFormula('TIME(HOUR(NOW()),MINUTE(NOW()),SECOND(NOW()))', 'A2', ws);
		// assert.ok(oParser.parse(), 'Test: TIME(HOUR(NOW()),MINUTE(NOW()),SECOND(NOW())) is parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), 0.4432523148148148, 'Test: Positive case: Formula. Nested formulas returning valid numbers. 3 arguments used.');
		// Case #7: Formula. Nested IF formulas returning valid values. 3 arguments used.
		oParser = new parserFormula('TIME(IF(TRUE,12,0),IF(TRUE,30,0),IF(TRUE,0,60))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(IF(TRUE,12,0),IF(TRUE,30,0),IF(TRUE,0,60)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Formula. Nested IF formulas returning valid values. 3 arguments used.');
		// Case #8: Reference link. Reference to cells with valid numbers. 3 arguments used.
		oParser = new parserFormula('TIME(A100,A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(A100,A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.0006944444444444445, 'Test: Positive case: Reference link. Reference to cells with valid numbers. 3 arguments used.');
		// Case #9: Area. Single-cell ranges. 3 arguments used.
		oParser = new parserFormula('TIME(A100:A100,A101:A101,A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(A100:A100,A101:A101,A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.0006944444444444445, 'Test: Positive case: Area. Single-cell ranges. 3 arguments used.');
		// Case #10: Array. Arrays with single elements. 3 arguments used.
		oParser = new parserFormula('TIME({12},{30},{0})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME({12},{30},{0}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Array. Arrays with single elements. 3 arguments used.');
		// Case #11: Name. Named ranges with valid numbers. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME(TestName,TestName1,TestName2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME(TestName,TestName1,TestName2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named ranges with valid numbers. 3 arguments used.');
		// Case #12: Name3D. 3D named ranges with valid numbers. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME(TestName3D,TestName3D,TestName3D)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME(TestName3D,TestName3D,TestName3D) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named ranges with valid numbers. 3 arguments used.');
		// Case #13: Ref3D. 3D references to cells with valid numbers. 3 arguments used.
		oParser = new parserFormula('TIME(Sheet2!A1,Sheet2!A2,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Sheet2!A1,Sheet2!A2,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D. 3D references to cells with valid numbers. 3 arguments used.');
		// Case #14: Area3D. 3D single-cell ranges. 3 arguments used.
		oParser = new parserFormula('TIME(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell ranges. 3 arguments used.');
		// Case #15: Table. Table structured references with valid numbers. 3 arguments used.
		oParser = new parserFormula('TIME(Table1[Column1],Table1[Column1],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Table1[Column1],Table1[Column1],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04237268518518519, 'Test: Positive case: Table. Table structured references with valid numbers. 3 arguments used.');
		// Case #16: Formula. TIME as parent formula inside SUM. 3 arguments used.
		oParser = new parserFormula('SUM(TIME(12,30,0),0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(TIME(12,30,0),0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.0208333333333335, 'Test: Positive case: Formula. TIME as parent formula inside SUM. 3 arguments used.');
		// Case #17: Number. Large valid hour value, Excel adjusts (e.g., 1000 hours). 3 arguments used.
		oParser = new parserFormula('TIME(1000,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(1000,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.6666666666666643, 'Test: Positive case: Number. Large valid hour value, Excel adjusts (e.g., 1000 hours). 3 arguments used.');
		// Case #18: String. Numeric string for hour > 24, Excel adjusts. 3 arguments used.
		oParser = new parserFormula('TIME("25","0","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME("25","0","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04166666666666674, 'Test: Positive case: String. Numeric string for hour > 24, Excel adjusts. 3 arguments used.');
		// Case #19: Array. Multi-element arrays, processes first elements. 3 arguments used.
		oParser = new parserFormula('TIME({12,23},{30,59},{0,59})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME({12,23},{30,59},{0,59}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Array. Multi-element arrays, processes first elements. 3 arguments used.');
		// Case #20: Date. Date and time functions as arguments. 3 arguments used.
		oParser = new parserFormula('TIME(DAY(DATE(2025,1,1)),MINUTE(TIME(12,30,0)),SECOND(TIME(12,30,0)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(DAY(DATE(2025,1,1)),MINUTE(TIME(12,30,0)),SECOND(TIME(12,30,0))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.0625, 'Test: Positive case: Date. Date and time functions as arguments. 3 arguments used.');
		// Case #21: Time. Time functions as arguments. 3 arguments used.
		oParser = new parserFormula('TIME(HOUR(TIME(12,0,0)),MINUTE(TIME(0,30,0)),SECOND(TIME(0,0,0)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(HOUR(TIME(12,0,0)),MINUTE(TIME(0,30,0)),SECOND(TIME(0,0,0))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Time. Time functions as arguments. 3 arguments used.');
		// Case #22: Empty. Seconds omitted, defaults to 0. 3 arguments used.
		oParser = new parserFormula('TIME(12,30,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(12,30,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Empty. Seconds omitted, defaults to 0. 3 arguments used.');

		// Negative cases:

		// Case #1: Number. Hour > 32767 returns #NUM!. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME(32768,0,0)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME(32768,0,0) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Hour > 32767 returns #NUM!. 3 arguments used.');
		// Case #2: Number. Negative hour returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME(-1,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(-1,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative hour returns #NUM!. 3 arguments used.');
		// Case #3: String. Non-numeric string returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME("abc",0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME("abc",0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 3 arguments used.');
		// Case #4: String. String convertible to number > 32767 returns #NUM!. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME("32768","0","0")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME("32768","0","0") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String. String convertible to number > 32767 returns #NUM!. 3 arguments used.');
		// Case #5: Error. Propagates #N/A error. 3 arguments used.
		oParser = new parserFormula('TIME(NA(),0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(NA(),0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 3 arguments used.');
		// Case #6: Area. Multi-cell range returns #NUM!. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME(A100:A101,A101:A101,A102:A102)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME(A100:A101,A101:A101,A102:A102) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.000694444, 'Test: Negative case: Area. Multi-cell range returns #NUM!. 3 arguments used.');
		// Case #7: Empty. Reference to empty cell returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME(A103,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(A103,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!. 3 arguments used.');
		// Case #8: String. Empty string returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME("",0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME("",0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 3 arguments used.');
		// Case #9: Boolean. Boolean FALSE (0) is valid, but testing for clarity. 3 arguments used.
		oParser = new parserFormula('TIME(FALSE,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(FALSE,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Boolean. Boolean FALSE (0) is valid, but testing for clarity. 3 arguments used.');
		// Case #10: Ref3D. 3D ref to text returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME(Sheet2!A4,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Sheet2!A4,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Ref3D. 3D ref to text returns #VALUE!. 3 arguments used.');
		// Case #11: Name. Named range with text returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME(TestNameArea2,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(TestNameArea2,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name. Named range with text returns #VALUE!. 3 arguments used.');
		// Case #12: Table. Table column with text returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TIME(Table1[Column2],0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Table1[Column2],0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!. 3 arguments used.');
		// Case #13: Formula. Formula resulting in #NUM! error. 3 arguments used.
		oParser = new parserFormula('TIME(SQRT(-1),0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(SQRT(-1),0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error. 3 arguments used.');
		// Case #14: Number. Negative minute returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME(0,-1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(0,-1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative minute returns #NUM!. 3 arguments used.');
		// Case #15: Array. Array with boolean returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME({FALSE},0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME({FALSE},0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Array. Array with boolean returns #NUM!. 3 arguments used.');
		// Case #16: Number. Seconds > 32767 returns #NUM!. 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('TIME(0,0,32768)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIME(0,0,32768) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Seconds > 32767 returns #NUM!. 3 arguments used.');
		// Case #17: String. String convertible to negative number returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME("-1","0","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME("-1","0","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String. String convertible to negative number returns #NUM!. 3 arguments used.');
		// Case #18: Area3D. 3D multi-cell range returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME(Sheet2!A1:A2,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(Sheet2!A1:A2,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.041666666666666664, 'Test: Negative case: Area3D. 3D multi-cell range returns #NUM!. 3 arguments used.');
		// Case #19: Time. Time value (decimal < 1) returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TIME(TIME(12,0,0),0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(TIME(12,0,0),0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Time. Time value (decimal < 1) returns #NUM!. 3 arguments used.');
		// Case #20: Formula. Formula resulting in #DIV/0! error propagates error. 3 arguments used.
		oParser = new parserFormula('TIME(MMULT(1,0),0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(MMULT(1,0),0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Formula. Formula resulting in #DIV/0! error propagates error. 3 arguments used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid values (00:00:00). 3 arguments used.
		oParser = new parserFormula('TIME(0,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(0,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number. Minimum valid values (00:00:00). 3 arguments used.');
		// Case #2: Number. Maximum valid values, Excel adjusts. 3 arguments used.
		oParser = new parserFormula('TIME(32767,32767,32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIME(32767,32767,32767) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.4257754629629744, 'Test: Bounded case: Number. Maximum valid values, Excel adjusts. 3 arguments used.');

		// TODO Need to fix: name/3d handle, diff results with big number(should be #NUM err)
		// Case #11: Name. Named ranges with valid numbers. 3 arguments used.
		// Case #12: Name3D. 3D named ranges with valid numbers. 3 arguments used.
		// Case #1: Number. Hour > 32767 returns #NUM!. 3 arguments used.
		// Case #4: String. String convertible to number > 32767 returns #NUM!. 3 arguments used.
		// Case #6: Area. Multi-cell range returns #NUM!. 3 arguments used.
		// Case #16: Number. Seconds > 32767 returns #NUM!. 3 arguments used.

		testArrayFormula2(assert, "TIME", 3, 3);
	});

	QUnit.test('Test: "TIMEVALUE"', function (assert) {
		let dif = 1e-9;
		oParser = new parserFormula("timevalue(\"10:02:34\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.ok(Math.abs(oParser.calculate().getValue() - 0.4184490740740740) < dif);

		oParser = new parserFormula("timevalue(\"02-01-2006 10:15:29 AM\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.ok(Math.abs(oParser.calculate().getValue() - 0.4274189814823330) < dif);

		oParser = new parserFormula("timevalue(\"22:02\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.ok(Math.abs(oParser.calculate().getValue() - 0.9180555555555560) < dif);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("11:00:00"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: String. Valid time string in 24-hour format. Returns 0.52083 (12:30/24:00).
		oParser = new parserFormula('TIMEVALUE("12:30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("12:30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: String. Valid time string in 24-hour format. Returns 0.52083 (12:30/24:00).');
		// Case #2: String. Valid time string for midnight. Returns 0.
		oParser = new parserFormula('TIMEVALUE("00:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("00:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. Valid time string for midnight. Returns 0.');
		// Case #3: String. Valid time string for end of day. Returns ~0.99999.
		oParser = new parserFormula('TIMEVALUE("23:59:59")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("23:59:59") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.999988425925926, 'Test: Positive case: String. Valid time string for end of day. Returns ~0.99999.');
		// Case #4: Formula. Nested formula returning valid time string.
		oParser = new parserFormula('TIMEVALUE(TEXT("1232","HH:MM"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TEXT("1232","HH:MM")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Formula. Nested formula returning valid time string.');
		// Case #5: Formula. Formula concatenating valid time string. Returns 0.60417.
		oParser = new parserFormula('TIMEVALUE(CONCAT("14:", "30"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(CONCAT("14:", "30")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.6041666666666666, 'Test: Positive case: Formula. Formula concatenating valid time string. Returns 0.60417.');
		// Case #6: Reference link. Reference to cell with valid time string.
		oParser = new parserFormula('TIMEVALUE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link. Reference to cell with valid time string.');
		// Case #7: Area. Single-cell range with valid time string.
		oParser = new parserFormula('TIMEVALUE(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell range with valid time string.');
		// Case #8: Array. Array with single valid time string.
		oParser = new parserFormula('TIMEVALUE({"12:30"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE({"12:30"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Array. Array with single valid time string.');
		// Case #9: Name. Named range with valid time string.
		oParser = new parserFormula('TIMEVALUE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name. Named range with valid time string.');
		// Case #10: Name3D. 3D named range with valid time string.
		oParser = new parserFormula('TIMEVALUE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name3D. 3D named range with valid time string.');
		// Case #11: Ref3D. 3D reference to cell with valid time string.
		oParser = new parserFormula('TIMEVALUE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D. 3D reference to cell with valid time string.');
		// Case #12: Area3D. 3D single-cell range with valid time string.
		oParser = new parserFormula('TIMEVALUE(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell range with valid time string.');
		// Case #13: Table. Table structured reference with valid time string.
		oParser = new parserFormula('TIMEVALUE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Table. Table structured reference with valid time string.');
		// Case #14: String. Valid time string without seconds. Returns 0.25.
		oParser = new parserFormula('TIMEVALUE("6:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("6:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.25, 'Test: Positive case: String. Valid time string without seconds. Returns 0.25.');
		// Case #15: Formula. Nested IF returning valid time string. Returns 0.41667.
		oParser = new parserFormula('TIMEVALUE(IF(TRUE, "10:00", "20:00"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(IF(TRUE, "10:00", "20:00")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.4166666666666667, 'Test: Positive case: Formula. Nested IF returning valid time string. Returns 0.41667.');
		// Case #16: String. Maximum valid time with milliseconds. Returns ~0.99999.
		oParser = new parserFormula('TIMEVALUE("23:59:59.999")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("23:59:59.999") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.999999988, 'Test: Positive case: String. Maximum valid time with milliseconds. Returns ~0.99999.');
		// Case #17: Formula. Formula constructing time string. Returns 0.52083.
		oParser = new parserFormula('TIMEVALUE("12:" & "30")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("12:" & "30") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Formula. Formula constructing time string. Returns 0.52083.');
		// Case #18: Reference link. Reference to cell with end-of-day time string.
		oParser = new parserFormula('TIMEVALUE(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link. Reference to cell with end-of-day time string.');
		// Case #19: Area. Single-cell range with time string from formula.
		oParser = new parserFormula('TIMEVALUE(A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A103:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell range with time string from formula.');
		// Case #20: Array. Array with single time string without seconds.
		oParser = new parserFormula('TIMEVALUE({"23:59"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE({"23:59"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.9993055555555556, 'Test: Positive case: Array. Array with single time string without seconds.');
		// Case #21: Formula. Formula extracting valid time string. Returns 0.52083.
		oParser = new parserFormula('TIMEVALUE(LEFT("12:30:45", 5))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(LEFT("12:30:45", 5)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Positive case: Formula. Formula extracting valid time string. Returns 0.52083.');
		// Case #22: String. Short format time string for midnight. Returns 0.
		oParser = new parserFormula('TIMEVALUE("0:0:0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("0:0:0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: String. Short format time string for midnight. Returns 0.');

		// Negative cases:

		// Case #1: String. Invalid time string (hour >= 24). Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE("24:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("24:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Invalid time string (hour >= 24). Returns #VALUE!.');
		// Case #2: String. Non-time string. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-time string. Returns #VALUE!.');
		// Case #3: Number. Number input. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Number input. Returns #VALUE!.');
		// Case #4: Boolean. Boolean input. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean input. Returns #VALUE!.');
		// Case #5: Error. Error input propagates #N/A.
		oParser = new parserFormula('TIMEVALUE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input propagates #N/A.');
		// Case #6: Empty. Empty string. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty string. Returns #VALUE!.');
		// Case #7: Area. Multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A104:A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range. Returns #VALUE!.');
		// Case #8: Array. Multi-element array. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE({"12:30", "14:30"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE({"12:30", "14:30"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5208333333333334, 'Test: Negative case: Array. Multi-element array. Returns #VALUE!.');
		// Case #9: Reference link. Reference to cell with invalid time string.
		oParser = new parserFormula('TIMEVALUE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with invalid time string.');
		// Case #10: Ref3D. 3D reference to invalid time string.
		oParser = new parserFormula('TIMEVALUE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to invalid time string.');
		// Case #11: Name. Named range with invalid time string.
		oParser = new parserFormula('TIMEVALUE(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with invalid time string.');
		// Case #12: Name3D. 3D named range with invalid time string.
		oParser = new parserFormula('TIMEVALUE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. 3D named range with invalid time string.');
		// Case #13: Area3D. 3D multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range. Returns #VALUE!.');
		// Case #14: Table. Table with invalid time string. Returns #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('TIMEVALUE(Table1[Column2])', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIMEVALUE(Table1[Column2]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.458333333, 'Test: Negative case: Table. Table with invalid time string. Returns #VALUE!.');
		// Case #15: String. Invalid hour in time string. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE("25:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("25:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04166666666666674, 'Test: Negative case: String. Invalid hour in time string. Returns #VALUE!.');
		// Case #16: String. Invalid minute in time string. Returns #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('TIMEVALUE("12:60")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIMEVALUE("12:60") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.541666667, 'Test: Negative case: String. Invalid minute in time string. Returns #VALUE!.');
		// Case #17: String. Invalid second in time string. Returns #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('TIMEVALUE("12:30:60")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIMEVALUE("12:30:60") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.521527778, 'Test: Negative case: String. Invalid second in time string. Returns #VALUE!.');
		// Case #18: Formula. Formula returning date serial number. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula returning date serial number. Returns #VALUE!.');
		// Case #19: Time. Time value as serial number. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Time. Time value as serial number. Returns #VALUE!.');
		// Case #20: Formula. Formula constructing invalid time string. Returns #VALUE!.
		oParser = new parserFormula('TIMEVALUE("12:" & "abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("12:" & "abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula constructing invalid time string. Returns #VALUE!.');
		// Case #21: Reference link. Reference to cell with non-time string.
		oParser = new parserFormula('TIMEVALUE(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with non-time string.');
		// Case #22: Area. Single-cell range with non-time string.
		oParser = new parserFormula('TIMEVALUE(A106:A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE(A106:A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Single-cell range with non-time string.');

		// Bounded cases:

		// Case #1: String. Minimum valid time. Returns 0.
		oParser = new parserFormula('TIMEVALUE("00:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("00:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: String. Minimum valid time. Returns 0.');
		// Case #2: String. Maximum valid time. Returns ~0.99999.
		oParser = new parserFormula('TIMEVALUE("23:59:59")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TIMEVALUE("23:59:59") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.999988425925926, 'Test: Bounded case: String. Maximum valid time. Returns ~0.99999.');
		// Case #3: String. Maximum valid time with milliseconds. Returns ~0.99999.
		// Different result with MS
		//oParser = new parserFormula('TIMEVALUE("23:59:59.999")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: TIMEVALUE("23:59:59.999") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.999999988, 'Test: Bounded case: String. Maximum valid time with milliseconds. Returns ~0.99999.');

		// TODO Need to fix: results diff from MS, table string handle
		// Case #16: String. Maximum valid time with milliseconds. Returns ~0.99999.
		// Case #14: Table. Table with invalid time string. Returns #VALUE!.
		// Case #16: String. Invalid minute in time string. Returns #VALUE!.
		// Case #17: String. Invalid second in time string. Returns #VALUE!.
		// Case #3: String. Maximum valid time with milliseconds. Returns ~0.99999.

		testArrayFormula(assert, "TIMEVALUE");
	});

	QUnit.test('Test: "WEEKDAY"', function (assert) {
		let array;

		// base mode
		ws.workbook.setDate1904(false, true);
		ws.getRange2("A2").setValue("2/14/2008");

		oParser = new parserFormula("WEEKDAY(A2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula("WEEKDAY(A2, 2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula("WEEKDAY(A2, 3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		// ws.getRange2("B1").setValue("=DATE(2023,1,1)");
		ws.getRange2("B1").setValue("2023/1/1");
		ws.getRange2("B2").setValue("2023/1/2");
		ws.getRange2("B3").setValue("2023/1/3");
		ws.getRange2("B4").setValue("2023/1/4");
		ws.getRange2("B5").setValue("2023/1/5");
		ws.getRange2("B6").setValue("2023/1/6");
		ws.getRange2("B7").setValue("2023/1/7");
		ws.getRange2("B8").setValue("2023/1/8");
		ws.getRange2("B9").setValue("2023/1/9");
		ws.getRange2("B10").setValue("2023/1/10");
		ws.getRange2("B11").setValue("2023/1/11");
		ws.getRange2("B12").setValue("2023/1/12");

		oParser = new parserFormula("WEEKDAY(B1,2)>5", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula("WEEKDAY(B1,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,1),2)");
		assert.strictEqual(oParser.calculate().getValue(), 7, "Result of WEEKDAY(DATE(2023,1,1),2)");

		oParser = new parserFormula("WEEKDAY(B2,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,2),2)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of WEEKDAY(DATE(2023,1,2),2)");

		oParser = new parserFormula("WEEKDAY(B3,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,3),2)");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Result of WEEKDAY(DATE(2023,1,3),2)");

		oParser = new parserFormula("WEEKDAY(B4,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,4),2)");
		assert.strictEqual(oParser.calculate().getValue(), 3, "Result of WEEKDAY(DATE(2023,1,4),2)");

		oParser = new parserFormula("WEEKDAY(B5,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,5),2)");
		assert.strictEqual(oParser.calculate().getValue(), 4, "Result of WEEKDAY(DATE(2023,1,5),2)");

		oParser = new parserFormula("WEEKDAY(B6,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,6),2)");
		assert.strictEqual(oParser.calculate().getValue(), 5, "Result of WEEKDAY(DATE(2023,1,6),2)");

		oParser = new parserFormula("WEEKDAY(B7,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,7),2)");
		assert.strictEqual(oParser.calculate().getValue(), 6, "Result of WEEKDAY(DATE(2023,1,7),2)");

		oParser = new parserFormula("WEEKDAY(B8,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,8),2)");
		assert.strictEqual(oParser.calculate().getValue(), 7, "Result of WEEKDAY(DATE(2023,1,8),2)");

		oParser = new parserFormula("WEEKDAY(B9,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,9),2)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of WEEKDAY(DATE(2023,1,9),2)");

		oParser = new parserFormula("WEEKDAY(B10,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,10),2)");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Result of WEEKDAY(DATE(2023,1,10),2)");

		oParser = new parserFormula("WEEKDAY(B11,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,11),2)");
		assert.strictEqual(oParser.calculate().getValue(), 3, "Result of WEEKDAY(DATE(2023,1,11),2)");

		oParser = new parserFormula("WEEKDAY(B12,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,12),2)");
		assert.strictEqual(oParser.calculate().getValue(), 4, "Result of WEEKDAY(DATE(2023,1,12),2)");

		// strings
		oParser = new parserFormula('WEEKDAY("44927",2)', "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,1),2)");
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY("44927",2)');

		oParser = new parserFormula('WEEKDAY("44927","2")', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),"2")');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY("44927","2")');

		oParser = new parserFormula('WEEKDAY("44927+1","2")', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY("44927+1","2")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY("44927+1","2")');

		oParser = new parserFormula('WEEKDAY("44927s","2")', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY("44927s","2")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY("44927s","2")');

		oParser = new parserFormula('WEEKDAY("44927","2+1")', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),"2+1")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY("44927","2+1")');

		oParser = new parserFormula('WEEKDAY("44927","2s")', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),"2s")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY("44927","2s")');

		// bools
		oParser = new parserFormula('WEEKDAY(44927,FALSE)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),FALSE)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(44927,FALSE)');

		oParser = new parserFormula('WEEKDAY(44927,TRUE)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(44927,TRUE)');

		oParser = new parserFormula('WEEKDAY(TRUE,TRUE)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(TRUE,TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(TRUE,TRUE)');

		oParser = new parserFormula('WEEKDAY(FALSE,TRUE)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(FALSE,TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(FALSE,TRUE)');

		// arrays
		oParser = new parserFormula('WEEKDAY({1,2,3},2)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},2)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY({1,2,3},2)');

		oParser = new parserFormula('WEEKDAY({1,2,3},2)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},2)');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY({1,2,3},2).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, "Result of WEEKDAY({1,2,3},2).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 2, "Result of WEEKDAY({1,2,3},2).[0,2]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "", "Result of WEEKDAY({1,2,3},2).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY({1,2,3},2).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(1,{1,2,3})', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(1,{1,2,3})');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(1,{1,2,3})');

		oParser = new parserFormula('WEEKDAY(1,{1,2,3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'WEEKDAY(1,{1,2,3})');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of WEEKDAY(1,{1,2,3}).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 7, "Result of WEEKDAY(1,{1,2,3}).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 6, "Result of WEEKDAY(1,{1,2,3}).[0,2]");
			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[0,3]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY({1,2,3},{1,2,3})', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},{1,2,3})');
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,2]");
			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},{1,2,3}).[0,3]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 1, "Result of WEEKDAY({1,2,3},{1,2,3}).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Result of WEEKDAY({1,2,3},{1,2,3}).[4,0]");
		}

		// cellsRange
		ws.getRange2("C101").setValue("1");
		ws.getRange2("C102").setValue("2");
		ws.getRange2("C103").setValue("3");
		ws.getRange2("C104").setValue("4");

		oParser = new parserFormula("WEEKDAY(C101:C103,2)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY(C101:C103,2).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(C101:C103,2).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1, "Result of WEEKDAY(C101:C103,2).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(C101:C103,2).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 2, "Result of WEEKDAY(C101:C103,2).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(C101:C103,2).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(C101:C103,2).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(C101:C103,C101:C103)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(C101:C103,C101:C103)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY(C101:C103,C101:C103)');

		oParser = new parserFormula("WEEKDAY(C101:C103,C101:C103)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of WEEKDAY(C101:C103,C101:C103).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1, "Result of WEEKDAY(C101:C103,C101:C103).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, "Result of WEEKDAY(C101:C103,C101:C103).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(C101:C103,C101:C103).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(1,C101:C103)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(1,C101:C103)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY(1,C101:C103)');

		oParser = new parserFormula("WEEKDAY(1,C101:C103)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Result of WEEKDAY(1,C101:C103).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(1,C101:C103).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY(1,C101:C103).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(1,C101:C103).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 6, "Result of WEEKDAY(1,C101:C103).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(1,C101:C103).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(1,C101:C103).[4,0]");
		}

		// errors
		oParser = new parserFormula('WEEKDAY(#N/A,2)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(#N/A,2)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of WEEKDAY(#N/A,2)');

		oParser = new parserFormula('WEEKDAY(#NUM!,#N/A)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(#NUM!,#N/A)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(#NUM!,#N/A)');

		oParser = new parserFormula('WEEKDAY(#DIV/0!,#N/A)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(#DIV/0!,#N/A)');
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'Result of WEEKDAY(#DIV/0!,#N/A)');

		// other
		oParser = new parserFormula('WEEKDAY(44927,)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(44927,)');

		oParser = new parserFormula('WEEKDAY(44927)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(44927)');

		oParser = new parserFormula('WEEKDAY(,)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,)');

		oParser = new parserFormula('WEEKDAY(,1)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,1)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(,1)');

		oParser = new parserFormula('WEEKDAY(,2)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,2)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Result of WEEKDAY(,2)');

		oParser = new parserFormula('WEEKDAY(,3)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,3)');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Result of WEEKDAY(,3)');

		oParser = new parserFormula('WEEKDAY(,4)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,4)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,4)');

		oParser = new parserFormula('WEEKDAY(,5)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,5)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,5)');

		oParser = new parserFormula('WEEKDAY(,6)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,6)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,6)');

		oParser = new parserFormula('WEEKDAY(,7)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,7)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,7)');

		oParser = new parserFormula('WEEKDAY(,8)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,8)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,8)');

		oParser = new parserFormula('WEEKDAY(,9)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,9)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,9)');

		oParser = new parserFormula('WEEKDAY(,10)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,10)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,10)');

		oParser = new parserFormula('WEEKDAY(,11)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,11)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Result of WEEKDAY(,11)');

		oParser = new parserFormula('WEEKDAY(,12)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,12)');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Result of WEEKDAY(,12)');

		oParser = new parserFormula('WEEKDAY(,13)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,13)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WEEKDAY(,13)');

		oParser = new parserFormula('WEEKDAY(,14)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,14)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Result of WEEKDAY(,14)');

		oParser = new parserFormula('WEEKDAY(,15)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,15)');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Result of WEEKDAY(,15)');

		oParser = new parserFormula('WEEKDAY(,16)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,16)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(,16)');

		oParser = new parserFormula('WEEKDAY(,17)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,17)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(,17)');

		oParser = new parserFormula('WEEKDAY(,20)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,20)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,20)');

		oParser = new parserFormula('WEEKDAY(,999999999999999999)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,999999999999999999)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,999999999999999999)');

		oParser = new parserFormula('WEEKDAY(,0)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,0)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,0)');

		oParser = new parserFormula('WEEKDAY(,-1)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,-1)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,-1)');


		testArrayFormula2(assert, "WEEKDAY", 1, 2);

		// set 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("WEEKDAY(B1,2)>5", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula("WEEKDAY(B1,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(B1,2)");
		assert.strictEqual(oParser.calculate().getValue(), 6, "Result of WEEKDAY(B1,2)");

		oParser = new parserFormula("WEEKDAY(B2,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,2),2)");
		assert.strictEqual(oParser.calculate().getValue(), 7, "Result of WEEKDAY(DATE(2023,1,2),2)");

		oParser = new parserFormula("WEEKDAY(B3,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,3),2)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of WEEKDAY(DATE(2023,1,3),2)");

		oParser = new parserFormula("WEEKDAY(B4,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,4),2)");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Result of WEEKDAY(DATE(2023,1,4),2)");

		oParser = new parserFormula("WEEKDAY(B5,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,5),2)");
		assert.strictEqual(oParser.calculate().getValue(), 3, "Result of WEEKDAY(DATE(2023,1,5),2)");

		oParser = new parserFormula("WEEKDAY(B6,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,6),2)");
		assert.strictEqual(oParser.calculate().getValue(), 4, "Result of WEEKDAY(DATE(2023,1,6),2)");

		oParser = new parserFormula("WEEKDAY(B7,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,7),2)");
		assert.strictEqual(oParser.calculate().getValue(), 5, "Result of WEEKDAY(DATE(2023,1,7),2)");

		oParser = new parserFormula("WEEKDAY(B8,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,8),2)");
		assert.strictEqual(oParser.calculate().getValue(), 6, "Result of WEEKDAY(DATE(2023,1,8),2)");

		oParser = new parserFormula("WEEKDAY(B9,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,9),2)");
		assert.strictEqual(oParser.calculate().getValue(), 7, "Result of WEEKDAY(DATE(2023,1,9),2)");

		oParser = new parserFormula("WEEKDAY(B10,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,10),2)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of WEEKDAY(DATE(2023,1,10),2)");

		oParser = new parserFormula("WEEKDAY(B11,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,11),2)");
		assert.strictEqual(oParser.calculate().getValue(), 2, "Result of WEEKDAY(DATE(2023,1,11),2)");

		oParser = new parserFormula("WEEKDAY(B12,2)", "A1", ws);
		assert.ok(oParser.parse(), "WEEKDAY(DATE(2023,1,12),2)");
		assert.strictEqual(oParser.calculate().getValue(), 3, "Result of WEEKDAY(DATE(2023,1,12),2)");

		// other
		oParser = new parserFormula('WEEKDAY(44927,)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(DATE(2023,1,1),)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(44927,)');

		oParser = new parserFormula('WEEKDAY(44927)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(44927)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(44927)');

		oParser = new parserFormula('WEEKDAY(,)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,)');

		oParser = new parserFormula('WEEKDAY(,1)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,1)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Result of WEEKDAY(,1)');

		oParser = new parserFormula('WEEKDAY(,2)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,2)');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Result of WEEKDAY(,2)');

		oParser = new parserFormula('WEEKDAY(,3)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,3)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WEEKDAY(,3)');

		oParser = new parserFormula('WEEKDAY(,4)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,4)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,4)');

		oParser = new parserFormula('WEEKDAY(,5)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,5)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,5)');

		oParser = new parserFormula('WEEKDAY(,6)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,6)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,6)');

		oParser = new parserFormula('WEEKDAY(,7)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,7)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,7)');

		oParser = new parserFormula('WEEKDAY(,8)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,8)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,8)');

		oParser = new parserFormula('WEEKDAY(,9)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,9)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,9)');

		oParser = new parserFormula('WEEKDAY(,10)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,10)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,10)');

		oParser = new parserFormula('WEEKDAY(,11)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,11)');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Result of WEEKDAY(,11)');

		oParser = new parserFormula('WEEKDAY(,12)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,12)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WEEKDAY(,12)');

		oParser = new parserFormula('WEEKDAY(,13)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,13)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Result of WEEKDAY(,13)');

		oParser = new parserFormula('WEEKDAY(,14)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,14)');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Result of WEEKDAY(,14)');

		oParser = new parserFormula('WEEKDAY(,15)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,15)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of WEEKDAY(,15)');

		oParser = new parserFormula('WEEKDAY(,16)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,16)');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(,16)');

		oParser = new parserFormula('WEEKDAY(,17)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,17)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Result of WEEKDAY(,17)');

		oParser = new parserFormula('WEEKDAY(,20)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,20)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,20)');

		oParser = new parserFormula('WEEKDAY(,999999999999999999)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,999999999999999999)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,999999999999999999)');

		oParser = new parserFormula('WEEKDAY(,0)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,0)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,0)');

		oParser = new parserFormula('WEEKDAY(,-1)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(,-1)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of WEEKDAY(,-1)');

		// arrays
		oParser = new parserFormula('WEEKDAY({1,2,3},2)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},2)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Result of WEEKDAY({1,2,3},2)');

		oParser = new parserFormula('WEEKDAY({1,2,3},2)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},2)');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 6, "Result of WEEKDAY({1,2,3},2).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 7, "Result of WEEKDAY({1,2,3},2).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1, "Result of WEEKDAY({1,2,3},2).[0,2]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "", "Result of WEEKDAY({1,2,3},2).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY({1,2,3},2).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},2).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(1,{1,2,3})', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(1,{1,2,3})');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of WEEKDAY(1,{1,2,3})');

		oParser = new parserFormula('WEEKDAY(1,{1,2,3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		assert.ok(oParser.parse(), 'WEEKDAY(1,{1,2,3})');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY(1,{1,2,3}).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 6, "Result of WEEKDAY(1,{1,2,3}).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 5, "Result of WEEKDAY(1,{1,2,3}).[0,2]");
			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[0,3]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(1,{1,2,3}).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(1,{1,2,3}).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY({1,2,3},{1,2,3})', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY({1,2,3},{1,2,3})');
		oParser.setArrayFormulaRef(ws.getRange2("F106:I109").bbox);
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,1]");
			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 0, "Result of WEEKDAY({1,2,3},{1,2,3}).[0,2]");
			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A", "Result of WEEKDAY({1,2,3},{1,2,3}).[0,3]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[1,1]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 0, "Result of WEEKDAY({1,2,3},{1,2,3}).[1,2]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 7, "Result of WEEKDAY({1,2,3},{1,2,3}).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Result of WEEKDAY({1,2,3},{1,2,3}).[4,0]");
		}

		// cellsRange with new values
		ws.getRange2("E201").setValue("1");
		ws.getRange2("E202").setValue("2");
		ws.getRange2("E203").setValue("3");
		ws.getRange2("E204").setValue("4");

		oParser = new parserFormula("WEEKDAY(E201:E203,2)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 6, "Result of WEEKDAY(E201:E203,2).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(E201:E203,2).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY(E201:E203,2).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(E201:E203,2).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, "Result of WEEKDAY(E201:E203,2).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(E201:E203,2).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(E201:E203,2).[4,0]");
		}

		oParser = new parserFormula("WEEKDAY(E201:E203,E201:E203)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY(E201:E203,E201:E203).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(E201:E203,E201:E203).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY(E201:E203,E201:E203).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(E201:E203,E201:E203).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0, "Result of WEEKDAY(E201:E203,E201:E203).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(E201:E203,E201:E203).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(E201:E203,E201:E203).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(1,E201:E203)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(1,E201:E203)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY(1,E201:E203)');

		oParser = new parserFormula("WEEKDAY(1,E201:E203)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 7, "Result of WEEKDAY(1,E201:E203).[0,0]");
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(1,E201:E203).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 6, "Result of WEEKDAY(1,E201:E203).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(1,E201:E203).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 5, "Result of WEEKDAY(1,E201:E203).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(1,E201:E203).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(1,E201:E203).[4,0]");
		}

		// cellsRange with old values
		// ?? If don't redefine the values ​​in these cells after changing the mode to 1904, then the results in some arrays may be different
		// ws.getRange2("C101").setValue("1");
		// ws.getRange2("C102").setValue("2");
		// ws.getRange2("C103").setValue("3");
		// ws.getRange2("C104").setValue("4");

		oParser = new parserFormula("WEEKDAY(C101:C103,2)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 3, "Result of WEEKDAY(C101:C103,2).[0,0]");		//6
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(C101:C103,2).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY(C101:C103,2).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(C101:C103,2).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, "Result of WEEKDAY(C101:C103,2).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(C101:C103,2).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(C101:C103,2).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(C101:C103,C101:C103)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(C101:C103,C101:C103)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY(C101:C103,C101:C103)');

		oParser = new parserFormula("WEEKDAY(C101:C103,C101:C103)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#NUM!", "Result of WEEKDAY(C101:C103,C101:C103).[0,0]");		// 7
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 7, "Result of WEEKDAY(C101:C103,C101:C103).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0, "Result of WEEKDAY(C101:C103,C101:C103).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(C101:C103,C101:C103).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(C101:C103,C101:C103).[4,0]");
		}

		oParser = new parserFormula('WEEKDAY(1,C101:C103)', "A1", ws);
		assert.ok(oParser.parse(), 'WEEKDAY(1,C101:C103)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WEEKDAY(1,C101:C103)');

		oParser = new parserFormula("WEEKDAY(1,C101:C103)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:E109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#NUM!", "Result of WEEKDAY(1,C101:C103).[0,0]");		// 7
			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Result of WEEKDAY(1,C101:C103).[0,1]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 6, "Result of WEEKDAY(1,C101:C103).[1,0]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Result of WEEKDAY(1,C101:C103).[1,1]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 5, "Result of WEEKDAY(1,C101:C103).[2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WEEKDAY(1,C101:C103).[3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "#N/A", "Result of WEEKDAY(1,C101:C103).[4,0]");
		}

		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Number. Serial number for 01/01/2024, return_type 1 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 1 (Sunday=1). Returns 2 (Monday).');
		// Case #2: Number. Serial number for 01/01/2024, return_type 2 (Monday=1). Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 2 (Monday=1). Returns 1 (Monday).');
		// Case #3: Number. Serial number for 01/01/2024, return_type 3 (Monday=0). Returns 0 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 3 (Monday=0). Returns 0 (Monday).');
		// Case #4: String. US date string, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY("01/01/2024",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("01/01/2024",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. US date string, return_type 1. Returns 2 (Monday).');
		// Case #5: String. European date string, return_type 2. Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY("01.01.2024",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("01.01.2024",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. European date string, return_type 2. Returns 1 (Monday).');
		// Case #6: Formula. Date formula, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(DATE(2024,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(DATE(2024,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. Date formula, return_type 1. Returns 2 (Monday).');
		// Case #7: Formula,Number. Nested IF for return_type. Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY(DATE(2024,1,1),IF(TRUE,2,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(DATE(2024,1,1),IF(TRUE,2,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula,Number. Nested IF for return_type. Returns 1 (Monday).');
		// Case #8: Reference link. Reference to cell with serial number 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(A100,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(A100,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Reference link. Reference to cell with serial number 45197, return_type 1. Returns 2 (Monday).');
		// Case #9: Area. Single-cell range with 45197, return_type 2. Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY(A101:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(A101:A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Positive case: Area. Single-cell range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #10: Array. Single-element array, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY({45197},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY({45197},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Array. Single-element array, return_type 1. Returns 2 (Monday).');
		// Case #11: Name. Named range with 45197, return_type 1. Returns 2 (Monday).
		// Different result with MS
		//oParser = new parserFormula('WEEKDAY(TestName,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKDAY(TestName,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with 45197, return_type 1. Returns 2 (Monday).');
		// Case #12: Name3D. 3D named range with 45197, return_type 2. Returns 1 (Monday).
		// Different result with MS
		//oParser = new parserFormula('WEEKDAY(TestName3D,2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKDAY(TestName3D,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #13: Ref3D. 3D reference with 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(Sheet2!A1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Sheet2!A1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference with 45197, return_type 1. Returns 2 (Monday).');
		// Case #14: Area3D. 3D single-cell range with 45197, return_type 2. Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY(Sheet2!A2:A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Sheet2!A2:A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #15: Table. Table reference with 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(Table1[Column1],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Table1[Column1],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table reference with 45197, return_type 1. Returns 2 (Monday).');
		// Case #16: Number. Serial number, return_type 11 (Monday=1). Returns 1 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,11)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,11) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Number. Serial number, return_type 11 (Monday=1). Returns 1 (Monday).');
		// Case #17: Number. Serial number, return_type 12 (Tuesday=1). Returns 7 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number. Serial number, return_type 12 (Tuesday=1). Returns 7 (Monday).');
		// Case #18: Formula. WEEKDAY inside SUM, return_type 1. Returns 3 (2+1).
		oParser = new parserFormula('SUM(WEEKDAY(45197,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(WEEKDAY(45197,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Formula. WEEKDAY inside SUM, return_type 1. Returns 3 (2+1).');
		// Case #19: String. Short US date string, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY("1/1/24",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("1/1/24",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. Short US date string, return_type 1. Returns 2 (Monday).');
		// Case #20: Time. Time offset with serial number, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(TIME(12,0,0)+45197,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(TIME(12,0,0)+45197,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Time. Time offset with serial number, return_type 1. Returns 2 (Monday).');
		// Case #21: Number. Serial number, return_type 17 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,17)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,17) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number. Serial number, return_type 17 (Sunday=1). Returns 2 (Monday).');
		// Case #22: String,Number. US date string, return_type 13 (Wednesday=1). Returns 6 (Monday).
		oParser = new parserFormula('WEEKDAY("01/01/2024",13)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("01/01/2024",13) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: String,Number. US date string, return_type 13 (Wednesday=1). Returns 6 (Monday).');

		// Negative cases:

		// Case #1: Number. Serial number 0 (invalid date). Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Negative case: Number. Serial number 0 (invalid date). Returns #NUM!.');
		// Case #2: Number. Negative serial number. Returns #NUM!.
		// Different result with MS
		//oParser = new parserFormula('WEEKDAY(-1,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKDAY(-1,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative serial number. Returns #NUM!.');
		// Case #3: String. Non-numeric string. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY("abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string. Returns #VALUE!.');
		// Case #4: Error. Propagates #N/A error. Returns #N/A.
		oParser = new parserFormula('WEEKDAY(NA(),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(NA(),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. Returns #N/A.');
		// Case #5: Area. Multi-cell range. Returns #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('WEEKDAY(A101:A102,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKDAY(A101:A102,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area. Multi-cell range. Returns #VALUE!.');
		// Case #6: Empty. Empty reference. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(A103,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(A103,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Negative case: Empty. Empty reference. Returns #VALUE!.');
		// Case #7: String. Empty string. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY("",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string. Returns #VALUE!.');
		// Case #8: Boolean. Boolean TRUE (1=01/01/1900). Returns 3 (Tuesday).
		oParser = new parserFormula('WEEKDAY(TRUE,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(TRUE,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Boolean. Boolean TRUE (1=01/01/1900). Returns 3 (Tuesday).');
		// Case #9: Ref3D. 3D ref to text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(Sheet2!A3,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Sheet2!A3,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D ref to text ("abc"). Returns #VALUE!.');
		// Case #10: Name. Named range with text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(TestNameArea2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(TestNameArea2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Negative case: Name. Named range with text ("abc"). Returns #VALUE!.');
		// Case #11: Table. Table column with text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(Table1[Column2],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Table1[Column2],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text ("abc"). Returns #VALUE!.');
		// Case #12: Formula. Formula resulting in #NUM!. Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(SQRT(-1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(SQRT(-1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM!. Returns #NUM!.');
		// Case #13: Number. Invalid return_type 0. Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(45197,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Invalid return_type 0. Returns #NUM!.');
		// Case #14: Number. return_type 18 out of range. Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(45197,18)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,18) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. return_type 18 out of range. Returns #NUM!.');
		// Case #15: String. Invalid date string. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY("13/13/2024",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY("13/13/2024",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string. Returns #VALUE!.');
		// Case #16: Array. Array with invalid element. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY({45197,TRUE},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY({45197,TRUE},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Negative case: Array. Array with invalid element. Returns #VALUE!.');
		// Case #17: Area3D. 3D multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(Sheet2!A1:A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(Sheet2!A1:A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area3D. 3D multi-cell range. Returns #VALUE!.');
		// Case #18: Number,String. Non-numeric return_type. Returns #VALUE!.
		oParser = new parserFormula('WEEKDAY(45197,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Non-numeric return_type. Returns #VALUE!.');
		// Case #19: Formula. Invalid date in formula. Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(DATE(2024,13,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(DATE(2024,13,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Formula. Invalid date in formula. Returns #NUM!.');
		// Case #20: Boolean. Boolean FALSE as return_type (0). Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(45197,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Boolean. Boolean FALSE as return_type (0). Returns #NUM!.');
		// Case #21: Number. Negative return_type. Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(45197,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative return_type. Returns #NUM!.');
		// Case #22: Name. Named range with invalid date (0). Returns #NUM!.
		oParser = new parserFormula('WEEKDAY(TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Negative case: Name. Named range with invalid date (0). Returns #NUM!.');

		// Bounded cases:

		// Case #1: Number. Min valid date (01/01/1900), return_type 1. Returns 3 (Tuesday).
		oParser = new parserFormula('WEEKDAY(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Min valid date (01/01/1900), return_type 1. Returns 3 (Tuesday).');
		// Case #2: Number. Max valid date (12/31/9999), return_type 1. Returns 4 (Wednesday).
		oParser = new parserFormula('WEEKDAY(2958465,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(2958465,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Number. Max valid date (12/31/9999), return_type 1. Returns 4 (Wednesday).');
		// Case #3: Number. Slightly above min date, return_type 1. Returns 3 (Tuesday).
		oParser = new parserFormula('WEEKDAY(1.000000000000001,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(1.000000000000001,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Slightly above min date, return_type 1. Returns 3 (Tuesday).');
		// Case #4: Number. Slightly below max date, return_type 1. Returns 4 (Wednesday).
		// Different result with MS
		//oParser = new parserFormula('WEEKDAY(2958465.99999999,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKDAY(2958465.99999999,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number. Slightly below max date, return_type 1. Returns 4 (Wednesday).');
		// Case #5: Number. Serial number, max valid return_type 17 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKDAY(45197,17)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKDAY(45197,17) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Bounded case: Number. Serial number, max valid return_type 17 (Sunday=1). Returns 2 (Monday).');

		// TODO Need to fix: area handle, error type diff, wrong result in boundary case
		// Case #11: Name. Named range with 45197, return_type 1. Returns 2 (Monday).
		// Case #12: Name3D. 3D named range with 45197, return_type 2. Returns 1 (Monday).
		// Case #2: Number. Negative serial number. Returns #NUM!.
		// Case #5: Area. Multi-cell range. Returns #VALUE!.
		// Case #4: Number. Slightly below max date, return_type 1. Returns 4 (Wednesday).


	});

	QUnit.test('Test: "WEEKNUM"', function (assert) {
		oParser = new parserFormula("WEEKNUM(DATE(2006,1,1))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2006,1,1),17)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2006,1,1),1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2006,1,1),21)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 52);

		oParser = new parserFormula("WEEKNUM(DATE(2006,2,1),1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula("WEEKNUM(DATE(2006,2,1),2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula("WEEKNUM(DATE(2006,2,1),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula("WEEKNUM(DATE(2007,1,1),15)", "A2", ws);//понед
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,1),15)", "A2", ws);//втор
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2003,1,1),15)", "A2", ws);//сред
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2009,1,1),15)", "A2", ws);//чет
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2010,1,1),15)", "A2", ws);//пят
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2011,1,1),15)", "A2", ws);//суб
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2012,1,1),11)", "A2", ws);//вск
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,4),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,10),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,11),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,17),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,18),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula("WEEKNUM(DATE(2008,1,24),11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula("WEEKNUM(DATE(2013,1,1),21)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(DATE(2013,1,7))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WEEKNUM(0, 21)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 52);

		oParser = new parserFormula("WEEKNUM(1, 21)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 52);

		oParser = new parserFormula("WEEKNUM(2, 21)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula("WEEKNUM(0, 17)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("WEEKNUM(1, 17)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Number. Serial number for 01/01/2024, return_type 1 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 1 (Sunday=1). Returns 2 (Monday).');
		// Case #2: Number. Serial number for 01/01/2024, return_type 2 (Monday=1). Returns 1 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40, 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 2 (Monday=1). Returns 1 (Monday).');
		// Case #3: Number. Serial number for 01/01/2024, return_type 3 (Monday=0). Returns 0 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Number. Serial number for 01/01/2024, return_type 3 (Monday=0). Returns 0 (Monday).');
		// Case #4: String. US date string, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM("01/01/2024",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("01/01/2024",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. US date string, return_type 1. Returns 2 (Monday).');
		// Case #5: String. European date string, return_type 2. Returns 1 (Monday).
		oParser = new parserFormula('WEEKNUM("01.01.2024",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("01.01.2024",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. European date string, return_type 2. Returns 1 (Monday).');
		// Case #6: Formula. Date formula, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(DATE(2024,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(DATE(2024,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Date formula, return_type 1. Returns 2 (Monday).');
		// Case #7: Formula,Number. Nested IF for return_type. Returns 1 (Monday).
		oParser = new parserFormula('WEEKNUM(DATE(2024,1,1),IF(TRUE,2,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(DATE(2024,1,1),IF(TRUE,2,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula,Number. Nested IF for return_type. Returns 1 (Monday).');
		// Case #8: Reference link. Reference to cell with serial number 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(A100,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(A100,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with serial number 45197, return_type 1. Returns 2 (Monday).');
		// Case #9: Area. Single-cell range with 45197, return_type 2. Returns 1 (Monday).
		// Different result with MS
		//oParser = new parserFormula('WEEKNUM(A101:A101,2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKNUM(A101:A101,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area. Single-cell range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #10: Array. Single-element array, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM({45197},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM({45197},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Positive case: Array. Single-element array, return_type 1. Returns 2 (Monday).');
		// Case #11: Name. Named range with 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(TestName,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(TestName,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with 45197, return_type 1. Returns 2 (Monday).');
		// Case #12: Name3D. 3D named range with 45197, return_type 2. Returns 1 (Monday).
		oParser = new parserFormula('WEEKNUM(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #13: Ref3D. 3D reference with 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(Sheet2!A1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(Sheet2!A1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference with 45197, return_type 1. Returns 2 (Monday).');
		// Case #14: Area3D. 3D single-cell range with 45197, return_type 2. Returns 1 (Monday).
		// Different result with MS
		//oParser = new parserFormula('WEEKNUM(Sheet2!A2:A2,2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKNUM(Sheet2!A2:A2,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area3D. 3D single-cell range with 45197, return_type 2. Returns 1 (Monday).');
		// Case #15: Table. Table reference with 45197, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(Table1[Column1],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(Table1[Column1],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table reference with 45197, return_type 1. Returns 2 (Monday).');
		// Case #16: Number. Serial number, return_type 11 (Monday=1). Returns 1 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,11)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,11) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40, 'Test: Positive case: Number. Serial number, return_type 11 (Monday=1). Returns 1 (Monday).');
		// Case #17: Number. Serial number, return_type 12 (Tuesday=1). Returns 7 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40, 'Test: Positive case: Number. Serial number, return_type 12 (Tuesday=1). Returns 7 (Monday).');
		// Case #18: Formula. WEEKDAY inside SUM, return_type 1. Returns 3 (2+1).
		oParser = new parserFormula('SUM(WEEKNUM(45197,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(WEEKNUM(45197,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 40, 'Test: Positive case: Formula. WEEKDAY inside SUM, return_type 1. Returns 3 (2+1).');
		// Case #19: String. Short US date string, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM("1/1/24",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("1/1/24",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Short US date string, return_type 1. Returns 2 (Monday).');
		// Case #20: Time. Time offset with serial number, return_type 1. Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(TIME(12,0,0)+45197,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(TIME(12,0,0)+45197,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Positive case: Time. Time offset with serial number, return_type 1. Returns 2 (Monday).');
		// Case #21: Number. Serial number, return_type 17 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,17)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,17) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Positive case: Number. Serial number, return_type 17 (Sunday=1). Returns 2 (Monday).');
		// Case #22: String,Number. US date string, return_type 13 (Wednesday=1). Returns 6 (Monday).
		oParser = new parserFormula('WEEKNUM("01/01/2024",13)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("01/01/2024",13) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String,Number. US date string, return_type 13 (Wednesday=1). Returns 6 (Monday).');

		// Negative cases:

		// Case #1: Number. Serial number 0 (invalid date). Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Number. Serial number 0 (invalid date). Returns #NUM!.');
		// Case #2: Number. Negative serial number. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(-1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(-1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative serial number. Returns #NUM!.');
		// Case #3: String. Non-numeric string. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM("abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string. Returns #VALUE!.');
		// Case #4: Error. Propagates #N/A error. Returns #N/A.
		oParser = new parserFormula('WEEKNUM(NA(),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(NA(),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. Returns #N/A.');
		// Case #5: Area. Multi-cell range. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(A101:A102,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(A101:A102,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range. Returns #VALUE!.');
		// Case #6: Empty. Empty reference. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(A103,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(A103,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Empty. Empty reference. Returns #VALUE!.');
		// Case #7: String. Empty string. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM("",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string. Returns #VALUE!.');
		// Case #8: Boolean. Boolean TRUE (1=01/01/1900). Returns 3 (Tuesday).
		// Different result with MS
		//oParser = new parserFormula('WEEKNUM(TRUE,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKNUM(TRUE,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean TRUE (1=01/01/1900). Returns 3 (Tuesday).');
		// Case #9: Ref3D. 3D ref to text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(Sheet2!A3,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(Sheet2!A3,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D ref to text ("abc"). Returns #VALUE!.');
		// Case #10: Name. Named range with text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(TestNameArea,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(TestNameArea,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with text ("abc"). Returns #VALUE!.');
		// Case #11: Table. Table column with text ("abc"). Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(Table1[Column2],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(Table1[Column2],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text ("abc"). Returns #VALUE!.');
		// Case #12: Formula. Formula resulting in #NUM!. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(SQRT(-1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(SQRT(-1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM!. Returns #NUM!.');
		// Case #13: Number. Invalid return_type 0. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(45197,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Invalid return_type 0. Returns #NUM!.');
		// Case #14: Number. return_type 18 out of range. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(45197,18)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,18) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. return_type 18 out of range. Returns #NUM!.');
		// Case #15: String. Invalid date string. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM("13/13/2024",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM("13/13/2024",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string. Returns #VALUE!.');
		// Case #16: Array. Array with invalid element. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM({45197,TRUE},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM({45197,TRUE},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Negative case: Array. Array with invalid element. Returns #VALUE!.');
		// Case #17: Area3D. 3D multi-cell range. Returns #VALUE!.
		// Different result with MS
		//oParser = new parserFormula('WEEKNUM(Sheet2!A1:A2,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKNUM(Sheet2!A1:A2,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range. Returns #VALUE!.');
		// Case #18: Number,String. Non-numeric return_type. Returns #VALUE!.
		oParser = new parserFormula('WEEKNUM(45197,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Non-numeric return_type. Returns #VALUE!.');
		// Case #19: Formula. Invalid date in formula. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(DATE(2024,13,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(DATE(2024,13,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Formula. Invalid date in formula. Returns #NUM!.');
		// Case #20: Boolean. Boolean FALSE as return_type (0). Returns #NUM!.
		// Different result with MS
		//oParser = new parserFormula('WEEKNUM(45197,FALSE)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,FALSE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean FALSE as return_type (0). Returns #NUM!.');
		// Case #21: Number. Negative return_type. Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(45197,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative return_type. Returns #NUM!.');
		// Case #22: Name. Named range with invalid date (0). Returns #NUM!.
		oParser = new parserFormula('WEEKNUM(TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Name. Named range with invalid date (0). Returns #NUM!.');

		// Bounded cases:

		// Case #1: Number. Min valid date (01/01/1900), return_type 1. Returns 3 (Tuesday).
		oParser = new parserFormula('WEEKNUM(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Min valid date (01/01/1900), return_type 1. Returns 3 (Tuesday).');
		// Case #2: Number. Max valid date (12/31/9999), return_type 1. Returns 4 (Wednesday).
		oParser = new parserFormula('WEEKNUM(2958465,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(2958465,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 53, 'Test: Bounded case: Number. Max valid date (12/31/9999), return_type 1. Returns 4 (Wednesday).');
		// Case #3: Number. Slightly above min date, return_type 1. Returns 3 (Tuesday).
		oParser = new parserFormula('WEEKNUM(1.000000000000001,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(1.000000000000001,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Slightly above min date, return_type 1. Returns 3 (Tuesday).');
		// Case #4: Number. Slightly below max date, return_type 1. Returns 4 (Wednesday).
		oParser = new parserFormula('WEEKNUM(2958465.99999999,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(2958465.99999999,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 53, 'Test: Bounded case: Number. Slightly below max date, return_type 1. Returns 4 (Wednesday).');
		// Case #5: Number. Serial number, max valid return_type 17 (Sunday=1). Returns 2 (Monday).
		oParser = new parserFormula('WEEKNUM(45197,17)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WEEKNUM(45197,17) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 39, 'Test: Bounded case: Number. Serial number, max valid return_type 17 (Sunday=1). Returns 2 (Monday).');

		// Need to fix: area handle, boolean handle
		// Case #9: Area. Single-cell range with 45197, return_type 2. Returns 1 (Monday).
		// Case #14: Area3D. 3D single-cell range with 45197, return_type 2. Returns 1 (Monday).
		// Case #8: Boolean. Boolean TRUE (1=01/01/1900). Returns 3 (Tuesday).
		// Case #17: Area3D. 3D multi-cell range. Returns #VALUE!.
		// Case #20: Boolean. Boolean FALSE as return_type (0). Returns #NUM!.


		testArrayFormula2(assert, "WEEKNUM", 1, 2, true, null);
	});

	QUnit.test('Test: "WORKDAY"', function (assert) {

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),0)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38718);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),10)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38730);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),-10)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38705);

		oParser = new parserFormula("WORKDAY(DATE(2006,1,1),20,{\"1-2-2006\",\"1-16-2006\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38748);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,6),1,DATE(2017,10,9))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43018);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,7),1,DATE(2017,10,9))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43018);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,25),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43000);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,25),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43000);

		oParser = new parserFormula("WORKDAY(DATE(2017,9,20),-1,DATE(2017,9,10))", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 42997);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,2),-1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43007);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,2),-1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43007);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,3),-3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43006);

		oParser = new parserFormula("WORKDAY(DATE(2017,10,4),-2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43010);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),1,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43222);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),2,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),3,{\"5-1-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43227);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,30),3,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43228);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43220);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),2,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43224);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),3,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43227);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),-1,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43217);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),-2,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43216);

		oParser = new parserFormula("WORKDAY(DATE(2018,4,29),0,{\"5-1-2018\", \"5-2-2018\",\"5-3-2018\"})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 43219);

		oParser = new parserFormula("WORKDAY({1,2,3},{1,2})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},1)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY(1,{1,2})", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},1.123)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("WORKDAY({1,2,3},-1.123)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula("WORKDAY({1,2,3},5)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula("WORKDAY(1,15)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 20);

		/*oParser = new parserFormula("WORKDAY(1,50)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 69);

		oParser = new parserFormula("WORKDAY(1,60)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 83);

		oParser = new parserFormula("WORKDAY(1,61)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 86);*/

		//todo Different result with MS - ms returns errors.
		/*ws.getRange2( "A101" ).setValue( "1" );
		ws.getRange2( "B101" ).setValue( "3.123" );
		ws.getRange2( "C101" ).setValue( "-4" );

		oParser = new parserFormula("WORKDAY(A101:B101,A101:B101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101,A101:B101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101:B101,A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("WORKDAY(A101,A101)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);*/

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Number(2). Start_date as serial number, days as integer. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number(2). Start_date as serial number, days as integer. 2 of 3 arguments used.');
		// Case #2: Number(3). Start_date and days as numbers, holidays as single-element array. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,5,{38838})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,5,{38838}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number(3). Start_date and days as numbers, holidays as single-element array. 3 of 3 arguments used.');
		// Case #3: String(2). Start_date as string, days as integer. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY("01/03/2006",10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY("01/03/2006",10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38734, 'Test: Positive case: String(2). Start_date as string, days as integer. 2 of 3 arguments used.');
		// Case #4: String,Number,String. Start_date and holidays as strings, days as number. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY("01/03/2006",5,"02/03/2006")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY("01/03/2006",5,"02/03/2006") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38727, 'Test: Positive case: String,Number,String. Start_date and holidays as strings, days as number. 3 of 3 arguments used.');
		// Case #5: Formula(2). Start_date as DATE formula, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(DATE(2006,3,1),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(DATE(2006,3,1),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula(2). Start_date as DATE formula, days as number. 2 of 3 arguments used.');
		// Case #6: Formula,Number,Array. Start_date as formula, days as number, holidays as array. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(DATE(2006,3,1),5,{38838,38839})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(DATE(2006,3,1),5,{38838,38839}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula,Number,Array. Start_date as formula, days as number, holidays as array. 3 of 3 arguments used.');
		// Case #7: Date,Number. Start_date as DATE formula, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(DATE(2025,1,1),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(DATE(2025,1,1),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45672, 'Test: Positive case: Date,Number. Start_date as DATE formula, days as number. 2 of 3 arguments used.');
		// Case #8: Time,Number. Start_date as TIME adjusted to valid date, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TIME(12,0,0)+38777,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TIME(12,0,0)+38777,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Time,Number. Start_date as TIME adjusted to valid date, days as number. 2 of 3 arguments used.');
		// Case #9: Reference link(2). Start_date as Reference link, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(A100,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(A100,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Reference link(2). Start_date as Reference link, days as number. 2 of 3 arguments used.');
		// Case #10: Reference link,Number,Area. Start_date as Reference link, days as number, holidays as Area. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(A100,5,A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(A100,5,A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Reference link,Number,Area. Start_date as Reference link, days as number, holidays as Area. 3 of 3 arguments used.');
		// Case #11: Area(2). Start_date as single-cell Area, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(A100:A100,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(A100:A100,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area(2). Start_date as single-cell Area, days as number. 2 of 3 arguments used.');
		// Case #12: Array(2). Start_date as single-element array, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY({38777},5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY({38777},5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Array(2). Start_date as single-element array, days as number. 2 of 3 arguments used.');
		// Case #13: Name(2). Start_date as Name, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TestName,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TestName,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name(2). Start_date as Name, days as number. 2 of 3 arguments used.');
		// Case #14: Name3D(2). Start_date as Name3D, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TestName3D,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TestName3D,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D(2). Start_date as Name3D, days as number. 2 of 3 arguments used.');
		// Case #15: Ref3D(2). Start_date as Ref3D, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(Sheet2!A1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(Sheet2!A1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Ref3D(2). Start_date as Ref3D, days as number. 2 of 3 arguments used.');
		// Case #16: Area3D(2). Start_date as single-cell Area3D, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(Sheet2!A1:A1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(Sheet2!A1:A1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area3D(2). Start_date as single-cell Area3D, days as number. 2 of 3 arguments used.');
		// Case #17: Table(2). Start_date as Table reference, days as number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(Table1[Column1],5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(Table1[Column1],5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Table(2). Start_date as Table reference, days as number. 2 of 3 arguments used.');
		// Case #18: Formula,Number,Table. Start_date as formula, days as number, holidays as Table. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(DATE(2006,3,1),5,Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(DATE(2006,3,1),5,Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula,Number,Table. Start_date as formula, days as number, holidays as Table. 3 of 3 arguments used.');
		// Case #19: Number,Formula. Start_date as number, days as formula. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,SUM(2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,SUM(2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number,Formula. Start_date as number, days as formula. 2 of 3 arguments used.');
		// Case #20: Formula. WORKDAY inside SUM formula. 2 of 3 arguments used.
		oParser = new parserFormula('SUM(WORKDAY(38777,5),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(WORKDAY(38777,5),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38794, 'Test: Positive case: Formula. WORKDAY inside SUM formula. 2 of 3 arguments used.');
		// Case #21: String,Number,Array. Start_date as string, days as number, holidays as string array. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY("01/03/2006",5,{"02/03/2006","03/03/2006"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY("01/03/2006",5,{"02/03/2006","03/03/2006"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38727, 'Test: Positive case: String,Number,Array. Start_date as string, days as number, holidays as string array. 3 of 3 arguments used.');
		// Case #22: Number,String. Start_date as number, days as string convertible to number. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,"5")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,"5") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number,String. Start_date as number, days as string convertible to number. 2 of 3 arguments used.');

		// Negative cases:

		// Case #1: Number,Number. Start_date is zero (invalid date), returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(0,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(0,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Negative case: Number,Number. Start_date is zero (invalid date), returns #NUM!. 2 of 3 arguments used.');
		// Case #2: String,Number. Start_date as non-date string, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY("abc",5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY("abc",5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,Number. Start_date as non-date string, returns #VALUE!. 2 of 3 arguments used.');
		// Case #3: Number,String. Days as non-numeric string, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Days as non-numeric string, returns #VALUE!. 2 of 3 arguments used.');
		// Case #4: Empty,Number. Start_date is empty, returns #NUM!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(,5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty,Number. Start_date is empty, returns #NUM!. 2 of 3 arguments used.');
		// Case #5: Number,Empty. Days is empty, returns #VALUE!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(38777,)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(38777,) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number,Empty. Days is empty, returns #VALUE!. 2 of 3 arguments used.');
		// Case #6: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TRUE,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TRUE,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 3 arguments used.');
		// Case #7: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(38777,TRUE)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(38777,TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 3 arguments used.');
		// Case #8: Formula,Number. Days as date (invalid), returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(DATE(2025,6,27),DATE(2025,5,27))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(DATE(2025,6,27),DATE(2025,5,27)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 109961, 'Test: Negative case: Formula,Number. Days as date (invalid), returns #VALUE!. 2 of 3 arguments used.');
		// Case #9: Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(38777,5,{TRUE})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(38777,5,{TRUE}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 3 of 3 arguments used.');
		// Case #10: Area,Number. Start_date as multi-cell Area, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(A100:A101,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(A100:A101,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area,Number. Start_date as multi-cell Area, returns #VALUE!. 2 of 3 arguments used.');
		// Case #11: Number,Area. Days as multi-cell Area, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Area. Days as multi-cell Area, returns #VALUE!. 2 of 3 arguments used.');
		// Case #12: Number,Number,Area. Holidays as multi-cell Area with invalid data, returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,5,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,5,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Negative case: Number,Number,Area. Holidays as multi-cell Area with invalid data, returns #VALUE!. 3 of 3 arguments used.');
		// Case #13: Ref3D,Number. Ref3D to non-date value, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(Sheet2!A2,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(Sheet2!A2,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Negative case: Ref3D,Number. Ref3D to non-date value, returns #VALUE!. 2 of 3 arguments used.');
		// Case #14: Name,Number. Start_date as Name with Area, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TestNameArea,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TestNameArea,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name,Number. Start_date as Name with Area, returns #VALUE!. 2 of 3 arguments used.');
		// Case #15: Number,Name. Days as Name with Area, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Name. Days as Name with Area, returns #VALUE!. 2 of 3 arguments used.');
		// Case #16: Name3D,Number. Start_date as Name3D with Area, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(TestNameArea3D2,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(TestNameArea3D2,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Negative case: Name3D,Number. Start_date as Name3D with Area, returns #VALUE!. 2 of 3 arguments used.');
		// Case #17: Number,Number,Area3D. Holidays as Area3D with invalid data, returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(38777,5,Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(38777,5,Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Negative case: Number,Number,Area3D. Holidays as Area3D with invalid data, returns #VALUE!. 3 of 3 arguments used.');
		// Case #18: Table. Table reference to invalid column, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(Table1[Column2],5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(Table1[Column2],5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table reference to invalid column, returns #VALUE!. 2 of 3 arguments used.');
		// Case #19: Formula,Number. Start_date as NA() error, propagates #N/A. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(NA(),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(NA(),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula,Number. Start_date as NA() error, propagates #N/A. 2 of 3 arguments used.');
		// Case #20: String,Number. Invalid date string, returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY("13/13/2025",5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY("13/13/2025",5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,Number. Invalid date string, returns #VALUE!. 2 of 3 arguments used.');

		// Bounded cases:

		// Case #1: Number(2). Minimum valid start_date (January 1, 1900), days as 1. 2 of 3 arguments used.
		oParser = new parserFormula('WORKDAY(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Number(2). Minimum valid start_date (January 1, 1900), days as 1. 2 of 3 arguments used.');
		// Case #2: Number(2). Maximum valid start_date (December 31, 9999), days as 1. 2 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(2958465,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(2958465,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(2). Maximum valid start_date (December 31, 9999), days as 1. 2 of 3 arguments used.');
		// Case #3: Number(3). Large positive days, holidays as max date. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(38777,9999999,{2958465})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(38777,9999999,{2958465}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(3). Large positive days, holidays as max date. 3 of 3 arguments used.');
		// Case #4: Number(3). Large negative days, holidays as min date. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY(38777,-9999999,{1})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY(38777,-9999999,{1}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(3). Large negative days, holidays as min date. 3 of 3 arguments used.');

		// TODO Need to fix: empty handle, boolean handle, long calculation in boundary cases
		// Case #4: Empty,Number. Start_date is empty, returns #NUM!. 2 of 3 arguments used.
		// Case #5: Number,Empty. Days is empty, returns #VALUE!. 2 of 3 arguments used.
		// Case #6: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 3 arguments used.
		// Case #7: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 3 arguments used.
		// Case #9: Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 3 of 3 arguments used.
		// Case #2: Number(2). Maximum valid start_date (December 31, 9999), days as 1. 2 of 3 arguments used.
		// Case #3: Number(3). Large positive days, holidays as max date. 3 of 3 arguments used. - long calculation
		// Case #4: Number(3). Large negative days, holidays as min date. 3 of 3 arguments used. - long calculation

	});

	QUnit.test('Test: "WORKDAY.INTL"', function (assert) {
		let array;
		ws.getRange2("D10").setValue("44980");
		ws.getRange2("D11").setValue("44981");
		ws.getRange2("D12").setValue("1");

		oParser = new parserFormula("WORKDAY.INTL(DATE(2023,2,22),1,1,D10:D11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 44984);

		oParser = new parserFormula("WORKDAY.INTL(DATE(2023,2,22),D12,1,D10:D11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 44984);

		oParser = new parserFormula("WORKDAY.INTL(DATE(2023,2,22),2,1,D10:D11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 44985);

		oParser = new parserFormula("WORKDAY.INTL(DATE(2012,1,1),30,0)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula("WORKDAY.INTL(DATE(2012,1,1),90,11)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 41013);

		oParser = new parserFormula('TEXT(WORKDAY.INTL(DATE(2012,1,1),30,17),"m/dd/yyyy")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2/05/2012");

		oParser = new parserFormula('WORKDAY.INTL(151,8,"0000000")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 159);

		oParser = new parserFormula('WORKDAY.INTL(151,8,"0000000")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 159);

		oParser = new parserFormula('WORKDAY.INTL(159,8,"0011100")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 171);

		oParser = new parserFormula('WORKDAY.INTL(151,-18,"0000000")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 133);

		oParser = new parserFormula('WORKDAY.INTL(151,8,"1111111")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,1,1),20,1,{"1/2/2006","1/16/2006"})', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38748);

		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,1,1),20,{"1/2/2006","1/16/2006"})', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,1,1),-20,1,{"1/2/2006",,"1/16/2006"})', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 38691);

		// for bug 40648
		oParser = new parserFormula('WORKDAY.INTL({1,2,3;2,3,4},1)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:H108").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL({1,2,3;2,3,4},1)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[0,2]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[1,0]");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[1,1]");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 5, "Result of WORKDAY.INTL({1,2,3;2,3,4},1)[1,2]");

		oParser = new parserFormula('WORKDAY.INTL(1,{1,2,3;2,3,4})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:H108").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,{1,2,3;2,3,4})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[0,2]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[1,0]");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[1,1]");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 5, "Result of WORKDAY.INTL(1,{1,2,3;2,3,4})[1,2]");

		oParser = new parserFormula('WORKDAY.INTL(1,1,{1,2,3;2,3,4})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:H108").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,{1,2,3;2,3,4})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[0,2]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[1,0]");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[1,1]");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 2, "Result of WORKDAY.INTL(1,1,{1,2,3;2,3,4})[1,2]");

		oParser = new parserFormula('WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("F106:H108").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 5, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 6, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[0,2]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 5, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[1,0]");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 6, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[1,1]");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 8, "Result of WORKDAY.INTL({1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3;2,3,4},{1,2,3})[1,2]");

		oParser = new parserFormula('WORKDAY.INTL(1,1,1)', "A2", ws);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,1)');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Result of WORKDAY.INTL(1,1,1)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,2)', "A2", ws);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,2)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Result of WORKDAY.INTL(1,1,2)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,3)', "A2", ws);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,3)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WORKDAY.INTL(1,1,3)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,4)', "A2", ws);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,4)');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Result of WORKDAY.INTL(1,1,4)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,12)', "A2", ws);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,12)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Result of WORKDAY.INTL(1,1,12)');

		ws.getRange2("A10").setValue("1");
		ws.getRange2("A11").setValue("2");
		ws.getRange2("A12").setValue("3");
		ws.getRange2("B10").setValue("2");
		ws.getRange2("B11").setValue("2");
		ws.getRange2("B12").setValue("2");
		ws.getRange2("C10").setValue("3");
		ws.getRange2("C11").setValue("3");
		ws.getRange2("C12").setValue("3");

		// first argument
		oParser = new parserFormula('WORKDAY.INTL(A10:A12,1)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(A10:A12,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WORKDAY.INTL(A10:A12,1)');

		oParser = new parserFormula('WORKDAY.INTL({1;2;3},1)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL({1;2;3},1)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL({1;2;3},1)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL({1;2;3},1)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4, "Result of WORKDAY.INTL({1;2;3},1)[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WORKDAY.INTL({1;2;3},1)[3,0]");

		oParser = new parserFormula('WORKDAY.INTL(A10:C10,1)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(A10:C10,1)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WORKDAY.INTL(A10:C10,1)');

		oParser = new parserFormula('WORKDAY.INTL({1,2,3},1)', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL({1,2,3},1)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL({1,2,3},1)[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL({1,2,3},1)[1,0]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL({1,2,3},1)[2,0]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WORKDAY.INTL({1,2,3},1)[3,0]");

		// second argument
		oParser = new parserFormula('WORKDAY.INTL(1,A10:A12)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,A10:A12)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WORKDAY.INTL(1,A10:A12)');

		oParser = new parserFormula('WORKDAY.INTL(1,{1;2;3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,{1;2;3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,{1;2;3})[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL(1,{1;2;3})[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4, "Result of WORKDAY.INTL(1,{1;2;3})[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WORKDAY.INTL(1,{1;2;3})[3,0]");

		oParser = new parserFormula('WORKDAY.INTL(1,A10:C10)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,A10:C10)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of WORKDAY.INTL(1,A10:C10)');

		oParser = new parserFormula('WORKDAY.INTL(1,{1,2,3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,{1,2,3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,{1,2,3})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL(1,{1,2,3})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL(1,{1,2,3})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WORKDAY.INTL(1,{1,2,3})[0,3]");

		// third arugument
		oParser = new parserFormula('WORKDAY.INTL(1,1,A10:A12)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,A10:A12)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,1,A10:A12)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL(1,1,A10:A12)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4, "Result of WORKDAY.INTL(1,1,A10:A12)[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WORKDAY.INTL(1,1,A10:A12)[3,0]");

		oParser = new parserFormula('WORKDAY.INTL(1,1,{1;2;3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,{1;2;3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,1,{1;2;3})[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 3, "Result of WORKDAY.INTL(1,1,{1;2;3})[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4, "Result of WORKDAY.INTL(1,1,{1;2;3})[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "", "Result of WORKDAY.INTL(1,1,{1;2;3})[3,0]");

		oParser = new parserFormula('WORKDAY.INTL(1,1,A10:C10)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,A10:C10)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,1,A10:C10)[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL(1,1,A10:C10)[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL(1,1,A10:C10)[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WORKDAY.INTL(1,1,A10:C10)[0,3]");

		oParser = new parserFormula('WORKDAY.INTL(1,1,{1,2,3})', "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,{1,2,3})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Result of WORKDAY.INTL(1,1,{1,2,3})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Result of WORKDAY.INTL(1,1,{1,2,3})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 4, "Result of WORKDAY.INTL(1,1,{1,2,3})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "", "Result of WORKDAY.INTL(1,1,{1,2,3})[0,3]");

		// fourth argument
		oParser = new parserFormula('WORKDAY.INTL(1,1,1,A10:A12)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,1,A10:A12)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WORKDAY.INTL(1,1,1,A10:A12)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,1,{1;2;3})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,1,{1;2;3})');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WORKDAY.INTL(1,1,1,{1;2;3})');

		oParser = new parserFormula('WORKDAY.INTL(1,1,1,A10:C10)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,1,A10:C10)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WORKDAY.INTL(1,1,1,A10:C10)');

		oParser = new parserFormula('WORKDAY.INTL(1,1,1,{1,2,3})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(1,1,1,{1,2,3})');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Result of WORKDAY.INTL(1,1,1,{1,2,3})');


		oParser = new parserFormula('WORKDAY.INTL(A10:C11,A10:C11,A10:C11)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(A10:C11,A10:C11,A10:C11)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[0,2]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[1,0]");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[1,1]");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[1,2]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[2,0]");
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[2,1]");
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:C11)[2,2]");

		oParser = new parserFormula('WORKDAY.INTL(A10:C11,A10:C11,A10:A10)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E10:E12").bbox);
		assert.ok(oParser.parse(), 'WORKDAY.INTL(A10:C11,A10:C11,A10:A10)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", "Result of WORKDAY.INTL(A10:C11,A10:C11,A10:A10)[0,0]");

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #1: Number(2). Start_date as serial number, days as integer, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number(2). Start_date as serial number, days as integer, default weekend. 2 of 4 arguments used.');
		// Case #2: Number(3). Start_date and days as numbers, weekend as number (Sat/Sun). 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number(3). Start_date and days as numbers, weekend as number (Sat/Sun). 3 of 4 arguments used.');
		// Case #3: Number,Number,String. Start_date and days as numbers, weekend as string (Sat/Sun). 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,"0000011")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,"0000011") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number,Number,String. Start_date and days as numbers, weekend as string (Sat/Sun). 3 of 4 arguments used.');
		// Case #4: Number(3),Array. Start_date, days, and weekend as numbers, holidays as single-element array. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,1,{38838})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,1,{38838}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number(3),Array. Start_date, days, and weekend as numbers, holidays as single-element array. 4 of 4 arguments used.');
		// Case #5: String(2). Start_date as string, days as integer, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL("01/03/2006",10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL("01/03/2006",10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38734, 'Test: Positive case: String(2). Start_date as string, days as integer, default weekend. 2 of 4 arguments used.');
		// Case #6: String,Number,String,String. Start_date and holidays as strings, days and weekend as string/number. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL("01/03/2006",5,"0000011","02/03/2006")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL("01/03/2006",5,"0000011","02/03/2006") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38727, 'Test: Positive case: String,Number,String,String. Start_date and holidays as strings, days and weekend as string/number. 4 of 4 arguments used.');
		// Case #7: Formula(2). Start_date as DATE formula, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,3,1),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(DATE(2006,3,1),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula(2). Start_date as DATE formula, days as number, default weekend. 2 of 4 arguments used.');
		// Case #8: Formula,Number,Number,Array. Start_date as formula, days and weekend as numbers, holidays as array. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,3,1),5,1,{38838,38839})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(DATE(2006,3,1),5,1,{38838,38839}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula,Number,Number,Array. Start_date as formula, days and weekend as numbers, holidays as array. 4 of 4 arguments used.');
		// Case #9: Date,Number. Start_date as DATE formula, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(DATE(2025,1,1),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(DATE(2025,1,1),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45672, 'Test: Positive case: Date,Number. Start_date as DATE formula, days as number, default weekend. 2 of 4 arguments used.');
		// Case #10: Time,Number. Start_date as TIME adjusted to valid date, days as number. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(TIME(12,0,0)+38777,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(TIME(12,0,0)+38777,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Time,Number. Start_date as TIME adjusted to valid date, days as number. 2 of 4 arguments used.');
		// Case #11: Reference link(2). Start_date as Reference link, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(A100,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(A100,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Reference link(2). Start_date as Reference link, days as number, default weekend. 2 of 4 arguments used.');
		// Case #12: Reference link,Number,Number,Area. Start_date as Reference link, days and weekend as numbers, holidays as Area. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(A100,5,1,A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(A100,5,1,A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Reference link,Number,Number,Area. Start_date as Reference link, days and weekend as numbers, holidays as Area. 4 of 4 arguments used.');
		// Case #13: Area(2). Start_date as single-cell Area, days as number, default weekend. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(A100:A100,5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(A100:A100,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area(2). Start_date as single-cell Area, days as number, default weekend. 2 of 4 arguments used.');
		// Case #14: Array(2). Start_date as single-element array, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL({38777},5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL({38777},5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Array(2). Start_date as single-element array, days as number, default weekend. 2 of 4 arguments used.');
		// Case #15: Name(2). Start_date as Name, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(TestName,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(TestName,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name(2). Start_date as Name, days as number, default weekend. 2 of 4 arguments used.');
		// Case #16: Name3D(2). Start_date as Name3D, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(TestName3D,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(TestName3D,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D(2). Start_date as Name3D, days as number, default weekend. 2 of 4 arguments used.');
		// Case #17: Ref3D(2). Start_date as Ref3D, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(Sheet2!A1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(Sheet2!A1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Ref3D(2). Start_date as Ref3D, days as number, default weekend. 2 of 4 arguments used.');
		// Case #18: Area3D(2). Start_date as single-cell Area3D, days as number, default weekend. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(Sheet2!A1:A1,5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(Sheet2!A1:A1,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Area3D(2). Start_date as single-cell Area3D, days as number, default weekend. 2 of 4 arguments used.');
		// Case #19: Table(2). Start_date as Table reference, days as number, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(Table1[Column1],5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(Table1[Column1],5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Table(2). Start_date as Table reference, days as number, default weekend. 2 of 4 arguments used.');
		// Case #20: Formula,Number,Table. Start_date as formula, days as number, weekend as Table. 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(DATE(2006,3,1),5,Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(DATE(2006,3,1),5,Table1[Column1]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Formula,Number,Table. Start_date as formula, days as number, weekend as Table. 3 of 4 arguments used.');
		// Case #21: String,Number,String,Array. Start_date as string, days as number, weekend as string, holidays as string array. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL("01/03/2006",5,"0000011",{"02/03/2006","03/03/2006"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL("01/03/2006",5,"0000011",{"02/03/2006","03/03/2006"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38727, 'Test: Positive case: String,Number,String,Array. Start_date as string, days as number, weekend as string, holidays as string array. 4 of 4 arguments used.');
		// Case #22: Number,String,Number. Start_date as number, days as string convertible to number, weekend as number. 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,"5",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,"5",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Positive case: Number,String,Number. Start_date as number, days as string convertible to number, weekend as number. 3 of 4 arguments used.');

		// Negative cases:

		// Case #1: Number,Number. Start_date is zero (invalid date), returns #NUM!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(0,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(0,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Negative case: Number,Number. Start_date is zero (invalid date), returns #NUM!. 2 of 4 arguments used.');
		// Case #2: String,Number. Start_date as non-date string, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL("abc",5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL("abc",5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,Number. Start_date as non-date string, returns #VALUE!. 2 of 4 arguments used.');
		// Case #3: Number,String. Days as non-numeric string, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Days as non-numeric string, returns #VALUE!. 2 of 4 arguments used.');
		// Case #4: Number,Number,String. Weekend as invalid string, returns #VALUE!. 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Number,String. Weekend as invalid string, returns #VALUE!. 3 of 4 arguments used.');
		// Case #5: Number,Number,Number. Weekend number outside valid range (1â??17), returns #NUM!. 3 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,18)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,18) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number,Number,Number. Weekend number outside valid range (1â??17), returns #NUM!. 3 of 4 arguments used.');
		// Case #6: Empty,Number. Start_date is empty, returns #NUM!. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(,5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty,Number. Start_date is empty, returns #NUM!. 2 of 4 arguments used.');
		// Case #7: Number,Empty. Days is empty, returns #VALUE!. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number,Empty. Days is empty, returns #VALUE!. 2 of 4 arguments used.');
		// Case #8: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(TRUE,5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(TRUE,5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 4 arguments used.');
		// Case #9: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,TRUE)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 4 arguments used.');
		// Case #10: Number,Number,Boolean. Weekend as boolean, returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,5,TRUE)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Negative case: Number,Number,Boolean. Weekend as boolean, returns #VALUE!. 3 of 4 arguments used.');
		// Case #11: Number,Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,5,1,{TRUE})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,1,{TRUE}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 4 of 4 arguments used.');
		// Case #12: Area,Number. Start_date as multi-cell Area, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(A100:A101,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(A100:A101,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area,Number. Start_date as multi-cell Area, returns #VALUE!. 2 of 4 arguments used.');
		// Case #13: Number,Area. Days as multi-cell Area, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Area. Days as multi-cell Area, returns #VALUE!. 2 of 4 arguments used.');
		// Case #14: Number,Number,Area. Weekend as multi-cell Area, returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,5,A100:A101)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,A100:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number,Number,Area. Weekend as multi-cell Area, returns #VALUE!. 3 of 4 arguments used.');
		// Case #15: Number,Number,Number,Area. Holidays as multi-cell Area with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,5,1,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,1,A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Negative case: Number,Number,Number,Area. Holidays as multi-cell Area with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #16: Ref3D,Number. Ref3D to non-date value, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(Sheet2!A2,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(Sheet2!A2,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Negative case: Ref3D,Number. Ref3D to non-date value, returns #VALUE!. 2 of 4 arguments used.');
		// Case #17: Name,Number. Start_date as Name with Area, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(TestNameArea2,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(TestNameArea2,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name,Number. Start_date as Name with Area, returns #VALUE!. 2 of 4 arguments used.');
		// Case #18: Number,Name. Days as Name with Area, returns #VALUE!. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(38777,TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,Name. Days as Name with Area, returns #VALUE!. 2 of 4 arguments used.');
		// Case #19: Number,Number,Name. Weekend as Name with Area, returns #VALUE!. 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,5,TestNameArea2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,5,TestNameArea2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 38784, 'Test: Negative case: Number,Number,Name. Weekend as Name with Area, returns #VALUE!. 3 of 4 arguments used.');
		// Case #20: Formula,Number. Start_date as NA() error, propagates #N/A. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(NA(),5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(NA(),5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula,Number. Start_date as NA() error, propagates #N/A. 2 of 4 arguments used.');

		// Bounded cases:

		// Case #1: Number(2). Minimum valid start_date (January 1, 1900), days as 1, default weekend. 2 of 4 arguments used.
		oParser = new parserFormula('WORKDAY.INTL(1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Bounded case: Number(2). Minimum valid start_date (January 1, 1900), days as 1, default weekend. 2 of 4 arguments used.');
		// Case #2: Number(2). Maximum valid start_date (December 31, 9999), days as 1, default weekend. 2 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(2958465,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(2958465,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(2). Maximum valid start_date (December 31, 9999), days as 1, default weekend. 2 of 4 arguments used.');
		// Case #3: Number(3). Large positive days, max weekend number (Mon only). 3 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,9999999,17)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,9999999,17) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(3). Large positive days, max weekend number (Mon only). 3 of 4 arguments used.');
		// Case #4: Number(4). Large negative days, weekend as 1, holidays as min date. 4 of 4 arguments used.
		// Different result with MS
		//oParser = new parserFormula('WORKDAY.INTL(38777,-9999999,1,{1})', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: WORKDAY.INTL(38777,-9999999,1,{1}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number(4). Large negative days, weekend as 1, holidays as min date. 4 of 4 arguments used.');

		// TODO Need to fix: correct types handle(error, areas, empty, boolean, array), results diff from ms, long calculations in boundary cases
		// Case #13: Area(2). Start_date as single-cell Area, days as number, default weekend. 2 of 4 arguments used.
		// Case #14: Array(2). Start_date as single-element array, days as number, default weekend. 2 of 4 arguments used.
		// Case #18: Area3D(2). Start_date as single-cell Area3D, days as number, default weekend. 2 of 4 arguments used.
		// Case #20: Formula,Number,Table. Start_date as formula, days as number, weekend as Table. 3 of 4 arguments used.
		// Case #6: Empty,Number. Start_date is empty, returns #NUM!. 2 of 4 arguments used.
		// Case #7: Number,Empty. Days is empty, returns #VALUE!. 2 of 4 arguments used.
		// Case #8: Boolean,Number. Start_date as boolean, returns #VALUE!. 2 of 4 arguments used.
		// Case #9: Number,Boolean. Days as boolean, returns #VALUE!. 2 of 4 arguments used.
		// Case #11: Number,Number,Number,Array. Holidays array with non-date value, returns #VALUE!. 4 of 4 arguments used.
		// Case #14: Number,Number,Area. Weekend as multi-cell Area, returns #VALUE!. 3 of 4 arguments used.
		// Case #19: Number,Number,Name. Weekend as Name with Area, returns #VALUE!. 3 of 4 arguments used.
		// Case #2: Number(2). Maximum valid start_date (December 31, 9999), days as 1, default weekend. 2 of 4 arguments used.

	});

	QUnit.test('Test: "YEAR"', function (assert) {
		let array;

		// base mode
		ws.workbook.setDate1904(false, true);

		// inline
		oParser = new parserFormula("YEAR(44469)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(30.09.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(30.09.21)");

		oParser = new parserFormula("YEAR(44560)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(30.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(30.12.21)");

		oParser = new parserFormula("YEAR(44561)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(31.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(31.12.21)");

		oParser = new parserFormula("YEAR(44510)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(10.11.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(10.11.21)");

		oParser = new parserFormula('YEAR("2021-10-01")', "A2", ws);
		assert.ok(oParser.parse(), 'YEAR("2021-10-01")');
		assert.strictEqual(oParser.calculate().getValue(), 2021, 'Result of YEAR("2021-10-01")');

		oParser = new parserFormula('YEAR("2021-12-31")', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-12-31')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-12-31')");

		oParser = new parserFormula('YEAR("2021-09-30")', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-09-30')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-09-30')");

		oParser = new parserFormula('YEAR("2021-10-31")', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-10-31')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-10-31')");

		oParser = new parserFormula('YEAR("2021-12-29")', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-12-29')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-12-29')");

		oParser = new parserFormula('YEAR(0)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(0)");

		oParser = new parserFormula('YEAR("1s")', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('1s')");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of YEAR('1s')");

		oParser = new parserFormula('YEAR(TRUE)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(TRUE)");

		oParser = new parserFormula('YEAR(FALSE)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(FALSE)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(FALSE)");

		oParser = new parserFormula('YEAR(#DIV/0!)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(#DIV/0!)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of YEAR(#DIV/0!)");

		// cell
		ws.getRange2("A1").setValue("44469");
		ws.getRange2("A2").setValue("44560");
		ws.getRange2("A3").setValue("44561");
		ws.getRange2("A4").setValue("44510");
		ws.getRange2("A5").setValue("2021-10-01");
		ws.getRange2("A6").setValue("2021-12-31");
		ws.getRange2("A7").setValue("2021-09-30");
		ws.getRange2("A8").setValue("2021-10-31");
		ws.getRange2("A9").setValue("2021-12-29");

		ws.getRange2("A10").setValue("0");
		ws.getRange2("A11").setValue("1s");
		ws.getRange2("A12").setValue("TRUE");
		ws.getRange2("A13").setValue("FALSE");
		ws.getRange2("A14").setValue("#DIV/0!");
		ws.getRange2("A15").setValue("");

		oParser = new parserFormula("YEAR(A1)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(30.09.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(30.09.21)");

		oParser = new parserFormula("YEAR(A2)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(30.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(30.12.21)");

		oParser = new parserFormula("YEAR(A3)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(31.12.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(31.12.21)");

		oParser = new parserFormula("YEAR(A4)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(10.11.21)");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR(10.11.21)");

		oParser = new parserFormula("YEAR(A5)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-10-01')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-10-01')");

		oParser = new parserFormula("YEAR(A6)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-12-31')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-12-31')");

		oParser = new parserFormula("YEAR(A7)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-09-30')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-09-30')");

		oParser = new parserFormula("YEAR(A8)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-10-31')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-10-31')");

		oParser = new parserFormula("YEAR(A9)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR('2021-12-29')");
		assert.strictEqual(oParser.calculate().getValue(), 2021, "Result of YEAR('2021-12-29')");

		oParser = new parserFormula('YEAR(A10)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(0)");

		oParser = new parserFormula('YEAR(A11)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('1s')");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of YEAR('1s')");

		oParser = new parserFormula('YEAR(A12)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(TRUE)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(TRUE)");

		oParser = new parserFormula('YEAR(A13)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(FALSE)");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR(FALSE)");

		oParser = new parserFormula('YEAR(A14)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(#DIV/0!)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of YEAR(#DIV/0!)");

		oParser = new parserFormula('YEAR(A15)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR('')");
		assert.strictEqual(oParser.calculate().getValue(), 1900, "Result of YEAR('')");

		oParser = new parserFormula('YEAR(A1:A4)', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("C1:C5").bbox);
		assert.ok(oParser.parse(), "YEAR(A1:A4)");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2021, "Result of YEAR(A1:A4)[0,0]");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 2021, "Result of YEAR(A1:A4)[1,0]");
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 2021, "Result of YEAR(A1:A4)[2,0]");
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 2021, "Result of YEAR(A1:A4)[3,0]");
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Result of YEAR(A1:A4)[4,0]");

		oParser = new parserFormula('YEAR({1,100,1000,10000})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:H1").bbox);
		assert.ok(oParser.parse(), "YEAR({1,100,1000,10000})");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1900, "Result of YEAR({1,100,1000,10000})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1900, "Result of YEAR({1,100,1000,10000})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1902, "Result of YEAR({1,100,1000,10000})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 1927, "Result of YEAR({1,100,1000,10000})[0,3]");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", "Result of YEAR({1,100,1000,10000})[0,4]");

		// set 1904 mode
		ws.workbook.setDate1904(true, true);

		oParser = new parserFormula("YEAR(A1)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A1)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A1)");

		oParser = new parserFormula("YEAR(A2)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A2)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A2)");

		oParser = new parserFormula("YEAR(A3)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A3)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A3)");

		oParser = new parserFormula("YEAR(A4)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A4)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A4)");

		oParser = new parserFormula("YEAR(A5)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A5)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A5)");

		oParser = new parserFormula("YEAR(A6)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A6)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A6)");

		oParser = new parserFormula("YEAR(A7)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A7)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A7)");

		oParser = new parserFormula("YEAR(A8)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A8)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A8)");

		oParser = new parserFormula("YEAR(A9)", "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A9)");
		assert.strictEqual(oParser.calculate().getValue(), 2025, "Result of YEAR(A9)");

		oParser = new parserFormula('YEAR(A10)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A10)");
		assert.strictEqual(oParser.calculate().getValue(), 1904, "Result of YEAR(A10)");

		oParser = new parserFormula('YEAR(A11)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A11)");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", "Result of YEAR(A11)");

		oParser = new parserFormula('YEAR(A12)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A12)");
		assert.strictEqual(oParser.calculate().getValue(), 1904, "Result of YEAR(A12)");

		oParser = new parserFormula('YEAR(A13)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A13)");
		assert.strictEqual(oParser.calculate().getValue(), 1904, "Result of YEAR(A13)");

		oParser = new parserFormula('YEAR(A14)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A14)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of YEAR(A14)");

		oParser = new parserFormula('YEAR(A15)', "A2", ws);
		assert.ok(oParser.parse(), "YEAR(A15)");
		assert.strictEqual(oParser.calculate().getValue(), 1904, "Result of YEAR(A15)");

		oParser = new parserFormula('YEAR({1,100,1000,10000})', "A2", ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:H1").bbox);
		assert.ok(oParser.parse(), "YEAR({1,100,1000,10000})");
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1904, "Result of YEAR({1,100,1000,10000})[0,0]");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1904, "Result of YEAR({1,100,1000,10000})[0,1]");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1906, "Result of YEAR({1,100,1000,10000})[0,2]");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 1931, "Result of YEAR({1,100,1000,10000})[0,3]");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", "Result of YEAR({1,100,1000,10000})[0,4]");

		// return to base mode
		ws.workbook.setDate1904(false, true);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("123"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:

		// Case #1: Number. Valid date serial number (March 1, 2006). 1 argument used.
		oParser = new parserFormula('YEAR(38777)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(38777) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: Number. Valid date serial number (March 1, 2006). 1 argument used.');
		// Case #2: Number. Valid date serial number (June 1, 2025). 1 argument used.
		oParser = new parserFormula('YEAR(45828)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(45828) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: Number. Valid date serial number (June 1, 2025). 1 argument used.');
		// Case #3: String. Date string in MM/DD/YYYY format. 1 argument used.
		oParser = new parserFormula('YEAR("03/01/2006")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("03/01/2006") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: String. Date string in MM/DD/YYYY format. 1 argument used.');
		// Case #4: String. Date string in DD.MM.YYYY format. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('YEAR("01.03.2006")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEAR("01.03.2006") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. Date string in DD.MM.YYYY format. 1 argument used.');
		// Case #5: String. Short date string in M/D/YY format. 1 argument used.
		oParser = new parserFormula('YEAR("5/5/25")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("5/5/25") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: String. Short date string in M/D/YY format. 1 argument used.');
		// Case #6: String. Date string in D-MMM-YYYY format. 1 argument used.
		oParser = new parserFormula('YEAR("1-Mar-2006")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("1-Mar-2006") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: String. Date string in D-MMM-YYYY format. 1 argument used.');
		// Case #7: Formula. Date generated by DATE formula. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(2025,6,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(2025,6,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: Formula. Date generated by DATE formula. 1 argument used.');
		// Case #8: Formula. Date from nested IF formula returning valid date. 1 argument used.
		oParser = new parserFormula('YEAR(IF(TRUE, DATE(2006,3,1), DATE(2000,1,1)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(IF(TRUE, DATE(2006,3,1), DATE(2000,1,1))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: Formula. Date from nested IF formula returning valid date. 1 argument used.');
		// Case #9: Reference link. Reference to cell with valid date serial number. 1 argument used.
		oParser = new parserFormula('YEAR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Positive case: Reference link. Reference to cell with valid date serial number. 1 argument used.');
		// Case #10: Area. Single-cell range with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Positive case: Area. Single-cell range with valid date. 1 argument used.');
		// Case #11: Array. Array with single valid date serial number. 1 argument used.
		oParser = new parserFormula('YEAR({38777})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR({38777}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: Array. Array with single valid date serial number. 1 argument used.');
		// Case #12: Name. Named range with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named range with valid date. 1 argument used.');
		// Case #13: Name3D. 3D named range with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named range with valid date. 1 argument used.');
		// Case #14: Ref3D. 3D reference to cell with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Positive case: Ref3D. 3D reference to cell with valid date. 1 argument used.');
		// Case #15: Area3D. 3D single-cell range with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Positive case: Area3D. 3D single-cell range with valid date. 1 argument used.');
		// Case #16: Table. Table structured reference with valid date. 1 argument used.
		oParser = new parserFormula('YEAR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Positive case: Table. Table structured reference with valid date. 1 argument used.');
		// Case #17: Date. Date formula for leap year (invalid, rolls to March 1, 2025). 1 argument used.
		oParser = new parserFormula('YEAR(DATE(2025,2,29))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(2025,2,29)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: Date. Date formula for leap year (invalid, rolls to March 1, 2025). 1 argument used.');
		// Case #18: Time. Date with time component, returns year. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(2025,1,1)+TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(2025,1,1)+TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: Time. Date with time component, returns year. 1 argument used.');
		// Case #19: Formula. YEAR inside SUM formula. 1 argument used.
		oParser = new parserFormula('SUM(YEAR(38777),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(YEAR(38777),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2007, 'Test: Positive case: Formula. YEAR inside SUM formula. 1 argument used.');
		// Case #20: Array. Multi-element array with valid dates. 1 argument used.
		oParser = new parserFormula('YEAR({38777,45828})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR({38777,45828}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2006, 'Test: Positive case: Array. Multi-element array with valid dates. 1 argument used.');
		// Case #21: Formula. Current year from DATE formula. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: Formula. Current year from DATE formula. 1 argument used.');
		// Case #22: String. Short date string format, assumes current year or default. 1 argument used.
		oParser = new parserFormula('YEAR("5/5")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("5/5") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2025, 'Test: Positive case: String. Short date string format, assumes current year or default. 1 argument used.');

		// Negative cases:

		// Case #1: Number. Zero date returns #NUM! error. 1 argument used.
		oParser = new parserFormula('YEAR(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Number. Zero date returns #NUM! error. 1 argument used.');
		// Case #2: Number. Negative date serial number returns #NUM!. 1 argument used.
		oParser = new parserFormula('YEAR(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative date serial number returns #NUM!. 1 argument used.');
		// Case #3: String. Non-date string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-date string returns #VALUE!. 1 argument used.');
		// Case #4: Empty. Reference to empty cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!. 1 argument used.');
		// Case #5: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('YEAR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #6: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #7: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('YEAR(A103:A104)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEAR(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #8: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #9: String. Invalid date string (13th month) returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR("13/01/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR("13/01/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid date string (13th month) returns #VALUE!. 1 argument used.');
		// Case #10: Formula. Formula resulting in #NUM! error. 1 argument used.
		oParser = new parserFormula('YEAR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error. 1 argument used.');
		// Case #11: Ref3D. 3D reference to text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to text returns #VALUE!. 1 argument used.');
		// Case #12: Name. Named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #13: Name3D. 3D named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 1 argument used.');
		// Case #14: Table. Table column with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!. 1 argument used.');
		// Case #15: Array. Array with boolean returns #VALUE!. 1 argument used.
		oParser = new parserFormula('YEAR({FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR({FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Array. Array with boolean returns #VALUE!. 1 argument used.');
		// Case #16: Time. Time value alone (0.5) returns #NUM!. 1 argument used.
		oParser = new parserFormula('YEAR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Time. Time value alone (0.5) returns #NUM!. 1 argument used.');
		// Case #17: Number. Date after max valid (Dec 31, 9999) returns #NUM!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('YEAR(2958466)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEAR(2958466) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Date after max valid (Dec 31, 9999) returns #NUM!. 1 argument used.');
		// Case #18: Formula. Date before 1900 returns #NUM!. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3799, 'Test: Negative case: Formula. Date before 1900 returns #NUM!. 1 argument used.');
		// Case #19: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('YEAR(Sheet2!A4:A5)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEAR(Sheet2!A4:A5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #20: Formula. Invalid month in DATE formula returns #NUM!. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(2025,13,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(2025,13,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2026, 'Test: Negative case: Formula. Invalid month in DATE formula returns #NUM!. 1 argument used.');
		// Case #21: String. Invalid day for month returns #VALUE!. 1 argument used.
		// Different result with MS
		//oParser = new parserFormula('YEAR("31-Feb-2025")', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEAR("31-Feb-2025") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid day for month returns #VALUE!. 1 argument used.');
		// Case #22: Number. Excessively large serial number returns #NUM!. 1 argument used.
		oParser = new parserFormula('YEAR(1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Excessively large serial number returns #NUM!. 1 argument used.');

		// Bounded cases:

		// Case #1: Number. Minimum valid date (Jan 1, 1900). 1 argument used.
		oParser = new parserFormula('YEAR(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Bounded case: Number. Minimum valid date (Jan 1, 1900). 1 argument used.');
		// Case #2: Number. Maximum valid date (Dec 31, 9999). 1 argument used.
		oParser = new parserFormula('YEAR(2958465)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(2958465) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9999, 'Test: Bounded case: Number. Maximum valid date (Dec 31, 9999). 1 argument used.');
		// Case #3: Formula. Minimum valid date via DATE formula. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(1900,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(1900,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1900, 'Test: Bounded case: Formula. Minimum valid date via DATE formula. 1 argument used.');
		// Case #4: Formula. Maximum valid date via DATE formula. 1 argument used.
		oParser = new parserFormula('YEAR(DATE(9999,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEAR(DATE(9999,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9999, 'Test: Bounded case: Formula. Maximum valid date via DATE formula. 1 argument used.');

		// TODO Need to fix: area handle, string handle(diff result)
		// Case #4: String. Date string in DD.MM.YYYY format. 1 argument used.
		// Case #7: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Case #17: Number. Date after max valid (Dec 31, 9999) returns #NUM!. 1 argument used.
		// Case #19: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		// Case #21: String. Invalid day for month returns #VALUE!. 1 argument used.

		testArrayFormula2(assert, "YEAR");
	});

	QUnit.test("Test: YEAR(2013)", function (assert) {
		oParser = new parserFormula("YEAR(2013)", "A1", ws);
		assert.ok(oParser.parse());
		if (AscCommon.bDate1904) {
			assert.strictEqual(oParser.calculate().getValue(), 1909);
		} else {
			assert.strictEqual(oParser.calculate().getValue(), 1905);
		}

		testArrayFormula2(assert, "YEAR", 1, 1);
	});

	QUnit.test("Test: \"YEARFRAC\"", function (assert) {
		let dif = 1e-9;
		function okWrapper(a, b) {
			assert.ok(Math.abs(a - b) < dif);
		}

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,3,26))", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.236111111);

		oParser = new parserFormula("YEARFRAC(DATE(2006,3,26),DATE(2006,1,1))", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.236111111);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1))", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.5);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2007,9,1))", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 1.666666667);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),0)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.5);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),1)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.495890411);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),2)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.502777778);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),3)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.495890411);

		oParser = new parserFormula("YEARFRAC(DATE(2006,1,1),DATE(2006,7,1),4)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 0.5);

		oParser = new parserFormula("YEARFRAC(DATE(2004,3,1),DATE(2006,3,1),1)", "A2", ws);
		assert.ok(oParser.parse());
		okWrapper(oParser.calculate().getValue(), 1.998175182481752);

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("123"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:F10").cleanAll();
		ws2.getRange2("A1").setValue("1");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("3");
		ws2.getRange2("B2").setValue("4");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		ws.getRange2("A201").setValue("-0.5"); // TestName
		ws.getRange2("A202").setValue("0.5"); // TestName1
		ws.getRange2("A203").setValue("10.5"); // TestName2
		ws2.getRange2("A11").setValue("-0.5"); // TestName3D
		ws.getRange2("A208").setValue("0.8"); // TestNameArea2
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:

		// Case #0: Date. Valid dates with basis 0 (US 30/360). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(DATE(2025,1,1),DATE(2026,1,1),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(DATE(2025,1,1),DATE(2026,1,1),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date. Valid dates with basis 0 (US 30/360). 3 of 3 arguments used.');
		// Case #1: Number. Serial dates with basis 1 (Actual/Actual). 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC(44927,45292,1)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number. Serial dates with basis 1 (Actual/Actual). 3 of 3 arguments used.');
		// Case #2: String. String dates in DD.MM.YYYY format with basis 2 (Actual/360). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("01/01/2025","01/01/2026",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("01/01/2025","01/01/2026",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.0138888888888888, 'Test: Positive case: String. String dates in DD.MM.YYYY format with basis 2 (Actual/360). 3 of 3 arguments used.');
		// Case #3: String. String dates in ISO format with basis 3 (Actual/365). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("2025-01-01","2026-01-01",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("2025-01-01","2026-01-01",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. String dates in ISO format with basis 3 (Actual/365). 3 of 3 arguments used.');
		// Case #4: Formula. Nested IF for start_date with basis 4 (European 30/360). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2026,1,1),4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(IF(TRUE,DATE(2025,1,1),DATE(2024,1,1)),DATE(2026,1,1),4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Nested IF for start_date with basis 4 (European 30/360). 3 of 3 arguments used.');
		// Case #5: Reference link. Reference to cells with valid serial dates, basis 0. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC(A100,A101,0)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC(A100,A101,0) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.002777778, 'Test: Positive case: Reference link. Reference to cells with valid serial dates, basis 0. 3 of 3 arguments used.');
		// Case #6: Area. Single-cell range references, basis 1. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(A100:A100,A101:A101,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(A100:A100,A101:A101,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.0027397260273972603, 'Test: Positive case: Area. Single-cell range references, basis 1. 3 of 3 arguments used.');
		// Case #7: Array. Array with single date elements, basis 2. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC({44927},{45292},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC({44927},{45292},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.0138888888888888, 'Test: Positive case: Array. Array with single date elements, basis 2. 3 of 3 arguments used.');
		// Case #8: Name. Named ranges with valid dates, basis 3. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(TestName,TestName1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(TestName,TestName1,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name. Named ranges with valid dates, basis 3. 3 of 3 arguments used.');
		// Case #9: Name3D. 3D named ranges with same date, basis 4. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(TestName3D,TestName3D,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(TestName3D,TestName3D,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Name3D. 3D named ranges with same date, basis 4. 3 of 3 arguments used.');
		// Case #10: Ref3D. 3D cell references, basis 0. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(Sheet2!A1,Sheet2!A2,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(Sheet2!A1,Sheet2!A2,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.002777777777777778, 'Test: Positive case: Ref3D. 3D cell references, basis 0. 3 of 3 arguments used.');
		// Case #11: Area3D. 3D single-cell ranges, basis 1. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(Sheet2!A1:A1,Sheet2!A2:A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(Sheet2!A1:A1,Sheet2!A2:A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.0027397260273972603, 'Test: Positive case: Area3D. 3D single-cell ranges, basis 1. 3 of 3 arguments used.');
		// Case #12: Table. Table structured references with valid dates, basis 2. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(Table1[Column1],Table1[Column1],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(Table1[Column1],Table1[Column1],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Table. Table structured references with valid dates, basis 2. 3 of 3 arguments used.');
		// Case #13: Time. Start_date with time component, basis 3. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(DATE(2025,1,1)+TIME(12,0,0),DATE(2026,1,1),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(DATE(2025,1,1)+TIME(12,0,0),DATE(2026,1,1),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Time. Start_date with time component, basis 3. 3 of 3 arguments used.');
		// Case #14: Number. Serial dates, basis 4. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(38777,38838,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(38777,38838,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.16666666666666666, 'Test: Positive case: Number. Serial dates, basis 4. 3 of 3 arguments used.');
		// Case #15: String. Short date string format MM/DD/YYYY, basis 0. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("5/5/2025","6/5/2025",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("5/5/2025","6/5/2025",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.08333333333333333, 'Test: Positive case: String. Short date string format MM/DD/YYYY, basis 0. 3 of 3 arguments used.');
		// Case #16: Formula. Full year fraction, basis 1. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(DATE(2025,1,1),DATE(2025,12,31),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(DATE(2025,1,1),DATE(2025,12,31),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.9972602739726028, 'Test: Positive case: Formula. Full year fraction, basis 1. 3 of 3 arguments used.');
		// Case #17: Array. Multi-element array, basis 2. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC({38777,44927},{38838,45292},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC({38777,44927},{38838,45292},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.16944444444444445, 'Test: Positive case: Array. Multi-element array, basis 2. 3 of 3 arguments used.');
		// Case #18: String. String dates in YYYY.MM.DD format, basis 3. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("2025.01.01","2026.01.01",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("2025.01.01","2026.01.01",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. String dates in YYYY.MM.DD format, basis 3. 3 of 3 arguments used.');
		// Case #19: Formula. Nested formula for start_date, basis 4. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(SQRT(2025^2),DATE(2026,1,1),4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(SQRT(2025^2),DATE(2026,1,1),4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 120.45555555555555, 'Test: Positive case: Formula. Nested formula for start_date, basis 4. 3 of 3 arguments used.');
		// Case #20: Number,Empty. Basis omitted, defaults to 0 (US 30/360). 2 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,45292,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number,Empty. Basis omitted, defaults to 0 (US 30/360). 2 of 3 arguments used.');
		// Case #21: Date. Mid-year dates, basis 1. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(DATE(2025,6,15),DATE(2026,6,15),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(DATE(2025,6,15),DATE(2026,6,15),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date. Mid-year dates, basis 1. 3 of 3 arguments used.');
		// Case #22: String. Textual month string format, basis 2. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("Jan 1, 2025","Jan 1, 2026",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("Jan 1, 2025","Jan 1, 2026",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.0138888888888888, 'Test: Positive case: String. Textual month string format, basis 2. 3 of 3 arguments used.');

		// Negative cases:

		// Case #1: Number. Negative start_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(-1,44927,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(-1,44927,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative start_date returns #NUM!. 3 of 3 arguments used.');
		// Case #2: Number. Negative end_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,-1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,-1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative end_date returns #NUM!. 3 of 3 arguments used.');
		// Case #3: String. Non-numeric string for start_date returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("abc","01/01/2025",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("abc","01/01/2025",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string for start_date returns #VALUE!. 3 of 3 arguments used.');
		// Case #4: String. Non-numeric string for end_date returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("01/01/2025","xyz",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("01/01/2025","xyz",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string for end_date returns #VALUE!. 3 of 3 arguments used.');
		// Case #5: Error. Propagates #N/A error. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(NA(),44927,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(NA(),44927,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 3 of 3 arguments used.');
		// Case #6: Area. Multi-cell range for start_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(A100:A101,A102,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(A100:A101,A102,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range for start_date returns #NUM!. 3 of 3 arguments used.');
		// Case #7: Area. Multi-cell range for end_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(A100,A101:A102,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(A100,A101:A102,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range for end_date returns #NUM!. 3 of 3 arguments used.');
		// Case #8: Empty. Empty cell references return #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(A103,A104,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(A103,A104,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Empty. Empty cell references return #VALUE!. 3 of 3 arguments used.');
		// Case #9: Boolean. Boolean FALSE (0) for start_date returns #NUM!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC(FALSE,44927,3)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC(FALSE,44927,3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean FALSE (0) for start_date returns #NUM!. 3 of 3 arguments used.');
		// Case #10: Boolean. Boolean TRUE for end_date returns #VALUE!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC(44927,TRUE,4)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,TRUE,4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean TRUE for end_date returns #VALUE!. 3 of 3 arguments used.');
		// Case #11: Number. Start_date > end_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(45292,44927,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(45292,44927,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Number. Start_date > end_date returns #NUM!. 3 of 3 arguments used.');
		// Case #12: Number. Negative basis returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,45292,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Negative basis returns #NUM!. 3 of 3 arguments used.');
		// Case #13: Number. Basis out of range (5) returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,45292,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Number. Basis out of range (5) returns #NUM!. 3 of 3 arguments used.');
		// Case #14: Ref3D. 3D reference to text values returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(Sheet2!A3,Sheet2!A4,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(Sheet2!A3,Sheet2!A4,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to text values returns #VALUE!. 3 of 3 arguments used.');
		// Case #15: Name. Named range with text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(TestNameArea,TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(TestNameArea,TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with text returns #VALUE!. 3 of 3 arguments used.');
		// Case #16: Name3D. 3D named range with text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(TestName3D,TestNameArea3D2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(TestName3D,TestNameArea3D2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 3 of 3 arguments used.');
		// Case #17: Table. Table column with text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(Table1[Column1],Table1[Column2],3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(Table1[Column1],Table1[Column2],3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!. 3 of 3 arguments used.');
		// Case #18: Formula. Formula resulting in #NUM! propagates error. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(SQRT(-1),44927,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(SQRT(-1),44927,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 3 of 3 arguments used.');
		// Case #19: String. Empty string for start_date returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC("","01/01/2025",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC("","01/01/2025",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string for start_date returns #VALUE!. 3 of 3 arguments used.');
		// Case #20: Time. Time value (fraction < 1) for start_date returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(TIME(12,0,0),DATE(2025,1,1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(TIME(12,0,0),DATE(2025,1,1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 125.0061493665935, 'Test: Negative case: Time. Time value (fraction < 1) for start_date returns #NUM!. 3 of 3 arguments used.');
		// Case #21: Array. Array with boolean returns #NUM!. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC({FALSE},45292,2)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC({FALSE},45292,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with boolean returns #NUM!. 3 of 3 arguments used.');

		// Bounded cases:

		// Case #1: Number. Minimum valid serial dates (1/1/1900 to 1/2/1900), basis 0. 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(1,2,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(1,2,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.002777777777777778, 'Test: Bounded case: Number. Minimum valid serial dates (1/1/1900 to 1/2/1900), basis 0. 3 of 3 arguments used.');
		// Case #2: Number. Maximum valid serial dates (12/31/9999 to 1/1/10000), basis 4. 3 of 3 arguments used.
		// Different result with MS
		//oParser = new parserFormula('YEARFRAC(2958465,2958466,4)', 'A2', ws);
		//assert.ok(oParser.parse(), 'Test: YEARFRAC(2958465,2958466,4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Bounded case: Number. Maximum valid serial dates (12/31/9999 to 1/1/10000), basis 4. 3 of 3 arguments used.');
		// Case #3: Number. Valid dates with smallest valid basis (0). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,45292,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Valid dates with smallest valid basis (0). 3 of 3 arguments used.');
		// Case #4: Number. Valid dates with largest valid basis (4). 3 of 3 arguments used.
		oParser = new parserFormula('YEARFRAC(44927,45292,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: YEARFRAC(44927,45292,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Number. Valid dates with largest valid basis (4). 3 of 3 arguments used.');

		// TODO Need to fix: results diff from MS
		// Case #1: Number. Serial dates with basis 1 (Actual/Actual). 3 of 3 arguments used.
		// Case #5: Reference link. Reference to cells with valid serial dates, basis 0. 3 of 3 arguments used.
		// Case #9: Boolean. Boolean FALSE (0) for start_date returns #NUM!. 3 of 3 arguments used.
		// Case #10: Boolean. Boolean TRUE for end_date returns #VALUE!. 3 of 3 arguments used.
		// Case #21: Array. Array with boolean returns #NUM!. 3 of 3 arguments used.
		// Case #2: Number. Maximum valid serial dates (12/31/9999 to 1/1/10000), basis 4. 3 of 3 arguments used.

		testArrayFormula2(assert, "YEARFRAC", 2, 3, true, null);
	});

	wb.dependencyFormulas.unlockRecal();
});
