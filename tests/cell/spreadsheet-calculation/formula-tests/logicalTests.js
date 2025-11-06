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

    	function testArrayFormulaEqualsValues(assert, str, formula, isNotLowerCase) {
		//***array-formula***
		ws.getRange2("A1").setValue("1");
		ws.getRange2("B1").setValue("3.123");
		ws.getRange2("C1").setValue("-4");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("B2").setValue("4");
		ws.getRange2("C2").setValue("5");

		oParser = new parserFormula(formula, "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E6:H8").bbox);
		assert.ok(oParser.parse());
		var array = oParser.calculate();

		var splitStr = str.split(";");

		for (var i = 0; i < splitStr.length; i++) {
			var subSplitStr = splitStr[i].split(",");
			for (var j = 0; j < subSplitStr.length; j++) {
				var valMs = subSplitStr[j];
				var element;
				if (array.getElementRowCol) {
					var row = 1 === array.array.length ? 0 : i;
					var col = 1 === array.array[0].length ? 0 : j;
					if (array.array[row] && array.array[row][col]) {
						element = array.getElementRowCol(row, col);
					} else {
						element = new window['AscCommonExcel'].cError(window['AscCommonExcel'].cErrorType.not_available);
					}
				} else {
					element = array;
				}
				var ourVal = element && undefined != element.value ? element.value.toString() : "#N/A";
				if (!isNotLowerCase) {
					valMs = valMs.toLowerCase();
					ourVal = ourVal.toLowerCase();
				}
				assert.strictEqual(valMs, ourVal, "formula: " + formula + " i: " + i + " j: " + j)
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

	QUnit.module('Logical formulas');

    QUnit.test("Test: \"AND\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A2").setValue("50");
		ws.getRange2("A3").setValue("100");
		ws.getRange2("A100").setValue("TRUE");
		ws.getRange2("A101").setValue("TRUE");
		ws.getRange2("A102").setValue("FALSE");
		ws.getRange2("A103").setValue("Test");
		ws.getRange2("A104").setValue("Text");
		// Table type. Use A601:L6**
		getTableType(599, 0, 601, 1);
		ws.getRange2("A601").setValue("TRUE"); // Column1
		ws.getRange2("A602").setValue("FALSE"); // Column1
		ws.getRange2("B601").setValue("Test"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("TRUE");
		ws2.getRange2("A2").setValue("FALSE");
		ws2.getRange2("A3").setValue("Test");
		ws2.getRange2("A4").setValue("");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("TRUE"); // TestName
		ws.getRange2("A202").setValue("TRUE"); // TestName1
		ws.getRange2("A203").setValue("FALSE"); // TestName2
		ws.getRange2("A204").setValue("Test"); // TestName3
		ws.getRange2("A206").setValue("Test"); // TestNameArea
		ws.getRange2("A207").setValue("Text"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("TRUE") // TestName3D
		ws2.getRange2("A12").setValue("TRUE") // TestName3D1
		ws2.getRange2("A13").setValue("FALSE") // TestName3D2
		ws2.getRange2("A14").setValue("Test") // TestName3D3
		ws2.getRange2("A16").setValue("TRUE"); // TestNameArea3D
		ws2.getRange2("A17").setValue("TRUE"); // TestNameArea3D

		// Positive cases:

		// Case #1: Reference link(2). Checking that reference links > 1 and < 100
		oParser = new parserFormula('AND(A2>1,A2<100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A2>1,A2<100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link(2). Checking that reference links > 1 and < 100');
		// Case #2: Reference link(2). Checking that reference links > 100 and < 100
		oParser = new parserFormula('AND(A2<A3,A2<100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A2<A3,A2<100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link(2). Checking that reference links > 100 and < 100');
		// Case #3: Reference link(2). Checking that reference  links > 1 and < 100
		oParser = new parserFormula('AND(A3>1,A3<100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A3>1,A3<100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link(2). Checking that reference  links > 1 and < 100');
		// Case #4: Boolean(2). Checking logical TRUE values with basic AND operation
		oParser = new parserFormula('AND(TRUE,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean(2). Checking logical TRUE values with basic AND operation');
		// Case #5: Boolean(2). Checking that AND returns FALSE when one argument is FALSE
		oParser = new parserFormula('AND(TRUE,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean(2). Checking that AND returns FALSE when one argument is FALSE');
		// Case #6: Number(2). Checking that positive numbers are treated as TRUE in logical operations
		oParser = new parserFormula('AND(1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number(2). Checking that positive numbers are treated as TRUE in logical operations');
		// Case #7: Number(2). Checking that zero is treated as FALSE in logical operations
		oParser = new parserFormula('AND(1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number(2). Checking that zero is treated as FALSE in logical operations');
		// Case #8: Number(3). Checking that multiple non-zero numbers are all treated as TRUE
		oParser = new parserFormula('AND(1,2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1,2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number(3). Checking that multiple non-zero numbers are all treated as TRUE');
		// Case #9: Empty, Boolean. Checking that empty cell is treated as FALSE
		oParser = new parserFormula('AND(,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Empty, Boolean. Checking that empty cell is treated as FALSE');
		// Case #10: Reference link(2). Checking AND with multiple cell references containing TRUE values
		oParser = new parserFormula('AND(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link(2). Checking AND with multiple cell references containing TRUE values');
		// Case #11: Reference link(2). Checking AND with cell references containing mixed TRUE/FALSE values
		oParser = new parserFormula('AND(A101,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A101,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link(2). Checking AND with cell references containing mixed TRUE/FALSE values');
		// Case #12: Area. Checking AND with range containing all TRUE values
		oParser = new parserFormula('AND(A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Checking AND with range containing all TRUE values');
		// Case #13: Area. Checking AND with range containing mixed TRUE/FALSE values
		oParser = new parserFormula('AND(A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A101:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Checking AND with range containing mixed TRUE/FALSE values');
		// Case #14: Array. Checking AND with array of all TRUE values
		oParser = new parserFormula('AND({TRUE,TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND({TRUE,TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Checking AND with array of all TRUE values');
		// Case #15: Array. Checking AND with array of mixed TRUE/FALSE values
		oParser = new parserFormula('AND({TRUE,FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND({TRUE,FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Checking AND with array of mixed TRUE/FALSE values');
		// Case #16: Name. Checking AND with named range containing all TRUE values
		oParser = new parserFormula('AND(TestName, TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestName, TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Checking AND with named range containing all TRUE values');
		// Case #17: Name. Checking AND with named range containing mixed TRUE/FALSE values
		oParser = new parserFormula('AND(TestName1, TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestName1, TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Checking AND with named range containing mixed TRUE/FALSE values');
		// Case #18: Name3D. Checking AND with 3D named range containing all TRUE values
		oParser = new parserFormula('AND(TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. Checking AND with 3D named range containing all TRUE values');
		// Case #19: Ref3D. Checking AND with 3D reference to cell with TRUE value
		oParser = new parserFormula('AND(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. Checking AND with 3D reference to cell with TRUE value');
		// Case #20: Area3D. Checking AND with 3D range containing mixed TRUE/FALSE values
		oParser = new parserFormula('AND(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. Checking AND with 3D range containing mixed TRUE/FALSE values');
		// Case #21: Table. Checking AND with table column containing all TRUE values
		oParser = new parserFormula('AND(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Checking AND with table column containing all TRUE values');
		// Case #22: Formula(2). Checking AND with equality and greater than comparisons
		oParser = new parserFormula('AND(1=1,2>1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1=1,2>1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula(2). Checking AND with equality and greater than comparisons');
		// Case #23: Formula(2). Checking AND with not equal and greater than or equal comparisons
		oParser = new parserFormula('AND(1<>2,3>=2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1<>2,3>=2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula(2). Checking AND with not equal and greater than or equal comparisons');
		// Case #24: Formula(3). Checking AND with nested IF functions returning mixed values
		oParser = new parserFormula('AND(IF(TRUE,TRUE,FALSE),IF(TRUE,TRUE,FALSE),IF(FALSE,FALSE,TRUE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(IF(TRUE,TRUE,FALSE),IF(TRUE,TRUE,FALSE),IF(FALSE,FALSE,TRUE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula(3). Checking AND with nested IF functions returning mixed values');
		// Case #25: Formula(2). Checking AND with nested OR function
		oParser = new parserFormula('AND(TRUE,OR(TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE,OR(TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula(2). Checking AND with nested OR function');
		// Case #26: Formula(3). Checking AND with multiple comparison operators
		oParser = new parserFormula('AND(10<5,20<30,5=5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(10<5,20<30,5=5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula(3). Checking AND with multiple comparison operators');
		// Case #27: Date, Number. Checking that date values are treated as TRUE
		oParser = new parserFormula('AND(DATE(2023,1,1),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(DATE(2023,1,1),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Date, Number. Checking that date values are treated as TRUE');
		// Case #28: Time, Number. Checking that non-zero time values are treated as TRUE
		oParser = new parserFormula('AND(TIME(12,0,0),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TIME(12,0,0),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time, Number. Checking that non-zero time values are treated as TRUE');
		// Case #29: Formula. Checking AND used inside nested IF and SUM functions
		oParser = new parserFormula('SUM(IF(AND(TRUE,TRUE),1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(IF(AND(TRUE,TRUE),1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula. Checking AND used inside nested IF and SUM functions');
		// Case #30: Boolean(3). Checking AND with multiple TRUE arguments
		oParser = new parserFormula('AND(TRUE,TRUE,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE,TRUE,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean(3). Checking AND with multiple TRUE arguments');
		// Case #31: Formula(2). Checking AND with arithmetic expressions in comparisons
		oParser = new parserFormula('AND(2*3>5,4^2=16)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(2*3>5,4^2=16) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula(2). Checking AND with arithmetic expressions in comparisons');
		// Case #32: Formula(2). Checking AND with multiple FALSE information functions
		oParser = new parserFormula('AND(ISNUMBER("text"),ISTEXT(123))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(ISNUMBER("text"),ISTEXT(123)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula(2). Checking AND with multiple FALSE information functions');

		// Negative cases:

		// Case #1: String. Checking that non-numeric text causes #VALUE! error
		oParser = new parserFormula('AND("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Checking that non-numeric text causes #VALUE! error');
		// Case #2: String, Boolean. Checking that non-numeric text with boolean
		// Different result with MS TODO Need to fix. Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row #14
		/*oParser = new parserFormula('AND("text",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND("text",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String, Boolean. Checking that non-numeric text with boolean');*/
		// Case #3: Error. Checking that NA() error propagates through AND
		oParser = new parserFormula('AND(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Checking that NA() error propagates through AND');
		// Case #4: Boolean, Error. Checking that error values propagate even with valid arguments
		oParser = new parserFormula('AND(TRUE,NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE,NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Boolean, Error. Checking that error values propagate even with valid arguments');
		// Case #5: Formula. Checking that division by zero error propagates through AND
		oParser = new parserFormula('AND(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula. Checking that division by zero error propagates through AND');
		// Case #6: Reference link. Checking that reference to cell with error propagates error
		oParser = new parserFormula('AND(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Checking that reference to cell with error propagates error');
		// Case #7: Area. Checking that range containing error value propagates error
		oParser = new parserFormula('AND(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(A103:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Checking that range containing error value propagates error');
		// Case #8: Array. Checking that array with error value propagates error
		oParser = new parserFormula('AND({TRUE,#N/A})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND({TRUE,#N/A}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Array. Checking that array with error value propagates error');
		// Case #9: Name. Checking that named range with error value propagates error
		oParser = new parserFormula('AND(TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Checking that named range with error value propagates error');
		// Case #10: Name3D. Checking that 3D named range with error propagates error
		oParser = new parserFormula('AND(TestName3D3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestName3D3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. Checking that 3D named range with error propagates error');
		// Case #11: Ref3D. Checking that 3D reference to error value propagates error
		oParser = new parserFormula('AND(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. Checking that 3D reference to error value propagates error');
		// Case #12: Area3D. Checking that 3D range with error value propagates error
		oParser = new parserFormula('AND(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. Checking that 3D range with error value propagates error');
		// Case #13: Table. Checking that table column with error propagates error
		oParser = new parserFormula('AND(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Checking that table column with error propagates error');
		// Case #14: Formula. Checking that #NUM! error from SQRT propagates through AND
		oParser = new parserFormula('AND(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Checking that #NUM! error from SQRT propagates through AND');
		// Case #15: Error. Checking  #REF! error
		oParser = new parserFormula('AND(#REF!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(#REF!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#REF!', 'Test: Negative case: Error. Checking  #REF! error');
		// Case #16: Formula. Checking that #DIV/0! in comparison propagates through AND
		oParser = new parserFormula('AND(5/0=10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(5/0=10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula. Checking that #DIV/0! in comparison propagates through AND');
		// Case #17: Name. Checking that named range with area returns logical result of all cells
		oParser = new parserFormula('AND(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Checking that named range with area returns logical result of all cells');
		// Case #18: String(2). Checking that strings convertible to numbers Numbers: "1", "2"
		oParser = new parserFormula('AND("1","2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND("1","2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Checking that strings convertible to numbers Numbers: "1", "2"');
		// Case #19: String(2). Checking that strings convertible to numbers. Numbers: "1", "0"
		oParser = new parserFormula('AND("1","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND("1","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Checking that strings convertible to numbers. Numbers: "1", "0"');

		// Bounded cases:

		// Case #1: Boolean. Checking minimum arguments (1) with TRUE value
		oParser = new parserFormula('AND(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Boolean. Checking minimum arguments (1) with TRUE value');
		// Case #2: Boolean. Checking minimum arguments (1) with FALSE value
		oParser = new parserFormula('AND(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Boolean. Checking minimum arguments (1) with FALSE value');
		// Case #3: Number. Checking minimum arguments (1) with numeric TRUE
		oParser = new parserFormula('AND(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Checking minimum arguments (1) with numeric TRUE');
		// Case #4: Number. Checking minimum arguments (1) with numeric FALSE
		oParser = new parserFormula('AND(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Checking minimum arguments (1) with numeric FALSE');
		// Case #5: Number. Checking AND with maximum valid Excel number
		oParser = new parserFormula('AND(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Checking AND with maximum valid Excel number');
		// Case #6: Number. Checking AND with minimum valid Excel number
		oParser = new parserFormula('AND(-9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(-9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Checking AND with minimum valid Excel number');
		// Case #7: Formula(3). Checking AND with boundary value comparisons
		oParser = new parserFormula('AND(9.99999999999999E+307>0,0>-9.99999999999999E+307,0.00000000000001<>0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula AND(9.99999999999999E+307>0,0>-9.99999999999999E+307,0.00000000000001<>0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Formula(3). Checking AND with boundary value comparisons');

		testArrayFormula2(assert, "AND", 1, 8, null, true);
	});

	QUnit.test("Test: \"IF\"", function (assert) {

		oParser = new parserFormula('IF(1,"TEST")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TEST");

		oParser = new parserFormula('IF(0,"TEST")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2("A101").setValue("1");

		oParser = new parserFormula('IF(A101=1,"Yes","No")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Yes");

		oParser = new parserFormula('IF(A101=2,"Yes","No")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "No");

		oParser = new parserFormula('IF(1,#N/A,#NUM!)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(1,#N/A,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IF(0,#N/A,#NUM!)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(0,#N/A,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('IF(1,#NUM!,#N/A)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(1,#NUM!,#N/A)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!");

		oParser = new parserFormula('IF(0,#NUM!,#N/A)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(0,#NUM!,#N/A)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		ws.getRange2("A1").setValue("0");
		ws.getRange2("A2").setValue("1");
		ws.getRange2("A3").setValue("2");
		ws.getRange2("A4").setValue("3");

		oParser = new parserFormula('IF(A1:A3>0, "Positive", "Non-positive")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3>0, "Positive", "Non-positive")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Non-positive");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Positive");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Positive");

		oParser = new parserFormula('IF(A1:A3>=0, "Positive", "Non-positive")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3>=0, "Positive", "Non-positive")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Positive");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Positive");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Positive");

		oParser = new parserFormula('IF(A1:A3<0, "Positive", "Non-positive")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3<0, "Positive", "Non-positive")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Non-positive");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Non-positive");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Non-positive");

		oParser = new parserFormula('IF(A1:A3<=0, "Positive", "Non-positive")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3<=0, "Positive", "Non-positive")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Positive");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Non-positive");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Non-positive");

		oParser = new parserFormula('IF(A1:A3<>0, "Positive", "Non-positive")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3<>0, "Positive", "Non-positive")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Non-positive");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Positive");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Positive");

		oParser = new parserFormula('IF(A1:A3="", "Empty", "Filled")', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(A1:A3="", "Empty", "Filled")');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "Filled");
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), "Filled");
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "Filled");

		oParser = new parserFormula('IF(A1:A3="Yes", 1, 0)', "AA2", ws);
		assert.ok(oParser.parse(), 'IIF(A1:A3="Yes", 1, 0)');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), 0);
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), 0);
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), 0);

		oParser = new parserFormula('IF(TRUE, A1:A3, A1:A2)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(TRUE, A1:A3, A1:A2)');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), 0);
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), 1);
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), 2);

		oParser = new parserFormula('IF(FALSE, A1:A3, A1:A2)', "AA2", ws);
		assert.ok(oParser.parse(), 'IF(FALSE, A1:A3, A1:A2)');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), 0);
		assert.strictEqual(array.getElementRowCol(1,0).getValue(), 1);
		assert.strictEqual(array.getElementRowCol(2,0).getValue(), "#N/A");

		oParser = new parserFormula('IF({1,0,1}, {"A","B","C"}, {"X","Y","Z"})', "AA2", ws);
		assert.ok(oParser.parse(), 'IF({1,0,1}, {"A","B","C"}, {"X","Y","Z"})');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), "A");
		assert.strictEqual(array.getElementRowCol(0,1).getValue(), "Y");
		assert.strictEqual(array.getElementRowCol(0,2).getValue(), "C");

		oParser = new parserFormula('IF({TRUE,FALSE,TRUE}, {10,20,30}, {100,200,300})', "AA2", ws);
		assert.ok(oParser.parse(), 'IF({TRUE,FALSE,TRUE}, {10,20,30}, {100,200,300})');
		oParser.setArrayFormulaRef(ws.getRange2("C10:E12").bbox);
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0,0).getValue(), 10);
		assert.strictEqual(array.getElementRowCol(0,1).getValue(), 200);
		assert.strictEqual(array.getElementRowCol(0,2).getValue(), 30);

	});
    
	QUnit.test("Test: \"IFERROR\"", function (assert) {

		ws.getRange2("A2").setValue("210");
		ws.getRange2("A3").setValue("55");
		ws.getRange2("A4").setValue("");

		ws.getRange2("B2").setValue("35");
		ws.getRange2("B3").setValue("0");
		ws.getRange2("B4").setValue("23");


		oParser = new parserFormula('IFERROR(A2/B2,"Error in calculation")', "A22", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);

		oParser = new parserFormula('IFERROR(A3/B3,"Error in calculation")', "A22", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 'Error in calculation');

		oParser = new parserFormula('IFERROR(A4/B4,"Error in calculation")', "A22", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("#N/A");
		ws.getRange2("A101").setValue("#NUM!");
		ws.getRange2("A104").setValue("3");
		// For area
		ws.getRange2("A102").setValue("1");
		ws.getRange2("A103").setValue("2");
		ws.getRange2("A105").setValue("#DIV/0!");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
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
		// Case #1: Number,String. Basic valid input: number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(10, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(10, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Number,String. Basic valid input: number with string fallback. 2 of 2 arguments used.');
		// Case #2: Number,Number. Float input with numeric fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(1.5, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(1.5, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.5, 'Test: Positive case: Number,Number. Float input with numeric fallback. 2 of 2 arguments used.');
		// Case #3: String,String. String convertible to number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR("2", "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR("2", "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Positive case: String,String. String convertible to number with string fallback. 2 of 2 arguments used.');
		// Case #4: Formula,String. Nested formula with valid result and string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(SQRT(4), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(SQRT(4), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,String. Nested formula with valid result and string fallback. 2 of 2 arguments used.');
		// Case #5: Reference link,String. Reference to cell with valid number. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(A100, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(A100, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Positive case: Reference link,String. Reference to cell with valid number. 2 of 2 arguments used.');
		// Case #6: Area,String. Single-cell range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(A101:A101, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(A101:A101, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Positive case: Area,String. Single-cell range with string fallback. 2 of 2 arguments used.');
		// Case #7: Array,String. Array with single element and string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR({2}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR({2}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 2, 'Test: Positive case: Array,String. Array with single element and string fallback. 2 of 2 arguments used.');
		// Case #8: Name,String. Named range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TestName, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TestName, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), -0.5, 'Test: Positive case: Name,String. Named range with string fallback. 2 of 2 arguments used.');
		// Case #9: Name3D,String. 3D named range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TestName3D, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TestName3D, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), -0.5, 'Test: Positive case: Name3D,String. 3D named range with string fallback. 2 of 2 arguments used.');
		// Case #10: Ref3D,String. 3D reference to cell with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Sheet2!A1, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Sheet2!A1, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Test: Positive case: Ref3D,String. 3D reference to cell with string fallback. 2 of 2 arguments used.');
		// Case #11: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Sheet2!A2:A2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Sheet2!A2:A2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.');
		// Case #12: Table,String. Table structured reference with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Table1[Column1], "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Table1[Column1], "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Test: Positive case: Table,String. Table structured reference with string fallback. 2 of 2 arguments used.');
		// Case #13: Date,String. Date as serial number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(DATE(2025,1,1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(DATE(2025,1,1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Date,String. Date as serial number with string fallback. 2 of 2 arguments used.');
		// Case #14: Time,String. Time adjusted to valid number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TIME(12,0,0)+1, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TIME(12,0,0)+1, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.5, 'Test: Positive case: Time,String. Time adjusted to valid number with string fallback. 2 of 2 arguments used.');
		// Case #15: Formula,Formula. Nested formula in both arguments, both valid. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(SQRT(4), ABS(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(SQRT(4), ABS(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,Formula. Nested formula in both arguments, both valid. 2 of 2 arguments used.');
		// Case #16: Number,Formula. Number with nested IF formula as fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(1, IF(TRUE, 0, 1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(1, IF(TRUE, 0, 1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number,Formula. Number with nested IF formula as fallback. 2 of 2 arguments used.');
		// Case #17: String,Number. Short date string with numeric fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR("12/12", 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR("12/12", 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12/12', 'Test: Positive case: String,Number. Short date string with numeric fallback. 2 of 2 arguments used.');
		// Case #18: Array,Number. Multi-element array with numeric fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR({1, 2}, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR({1, 2}, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array,Number. Multi-element array with numeric fallback. 2 of 2 arguments used.');
		// Case #19: Formula,String. Nested IF returning valid value with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(IF(TRUE, 2, 0.5), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(IF(TRUE, 2, 0.5), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,String. Nested IF returning valid value with string fallback. 2 of 2 arguments used.');
		// Case #20: Error,String. Error input with string fallback. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(NA(), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(NA(), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Positive case: Error,String. Error input with string fallback. Returns "Error". 2 of 2 arguments used.');
		// Case #21: Formula,Reference link. Nested formula with reference link fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(SQRT(4), A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(SQRT(4), A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,Reference link. Nested formula with reference link fallback. 2 of 2 arguments used.');
		// Case #22: Error,Area. Error input with single-cell range fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(1/0, A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(1/0, A103:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Error,Area. Error input with single-cell range fallback. 2 of 2 arguments used.');

		// Negative cases:
		// Case #1: Error,Error. Error input with error fallback. Returns #N/A. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(NA(), NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(NA(), NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error,Error. Error input with error fallback. Returns #N/A. 2 of 2 arguments used.');
		// Case #2: Number,String. Division by zero returns #DIV/0!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(1/0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(1/0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Number,String. Division by zero returns #DIV/0!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #3: String,String. Non-numeric string returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR("abc", "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR("abc", "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String,String. Non-numeric string returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #4: Formula,String. Formula resulting in #NUM! with string fallback. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(SQRT(-1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(SQRT(-1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Formula,String. Formula resulting in #NUM! with string fallback. Returns "Error". 2 of 2 arguments used.');
		// Case #5: Reference link,String. Reference to cell with error value. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(A104, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(A104, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 3, 'Test: Negative case: Reference link,String. Reference to cell with error value. Returns "Error". 2 of 2 arguments used.');
		// Case #6: Area,String. Multi-cell range returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(A105:A106, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(A105:A106, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Area,String. Multi-cell range returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #7: Empty,String. Reference to empty cell returns 0, fallback not used. Returns 0. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(A107, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(A107, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", 'Test: Negative case: Empty,String. Reference to empty cell returns 0, fallback not used. Returns 0. 2 of 2 arguments used.');
		// Case #9: Boolean,String. Boolean FALSE (0) is valid, fallback not used. Returns 0. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(FALSE, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(FALSE, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean,String. Boolean FALSE (0) is valid, fallback not used. Returns 0. 2 of 2 arguments used.');
		// Case #10: Ref3D,String. 3D reference to text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Sheet2!A3, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Sheet2!A3, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 'Text', 'Test: Negative case: Ref3D,String. 3D reference to text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #11: Name,String. Named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TestNameArea2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TestNameArea2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), -0.8, 'Test: Negative case: Name,String. Named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #12: Table,String. Table column with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Table1[Column2], "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Table1[Column2], "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), '1s', 'Test: Negative case: Table,String. Table column with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #13: Formula,String. Formula resulting in #N/A with string fallback. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(VLOOKUP("x", A1:B1, 2, FALSE), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(VLOOKUP("x", A1:B1, 2, FALSE), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Formula,String. Formula resulting in #N/A with string fallback. Returns "Error". 2 of 2 arguments used.');
		// Case #14: Number,String. Zero division returns #DIV/0!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(0/0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(0/0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Number,String. Zero division returns #DIV/0!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #15: Array,String. Array with boolean returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR({FALSE}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR({FALSE}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'FALSE', 'Test: Negative case: Array,String. Array with boolean returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #16: Number,String. Number too large returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(1E+308, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(1E+308, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1e+307, 'Test: Negative case: Number,String. Number too large returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #17: Date,String. Invalid date returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(DATE(10000,1,1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(DATE(10000,1,1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Date,String. Invalid date returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #18: Time,String. Invalid time returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TIME(25,0,0), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TIME(25,0,0), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04166666666666674, 'Test: Negative case: Time,String. Invalid time returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #19: Area3D,String. 3D multi-cell range returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(Sheet2!A4:A5, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(Sheet2!A4:A5, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area3D,String. 3D multi-cell range returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');
		// Case #20: Name3D,String. 3D named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(TestNameArea3D2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(TestNameArea3D2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.8, 'Test: Negative case: Name3D,String. 3D named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.');

		// Bounded cases:
		// Case #1: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(9.99999999999999E+307, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(9.99999999999999E+307, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1e+308, 'Test: Bounded case: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.');
		// Case #2: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(2.2250738585072014E-308, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(2.2250738585072014E-308, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.');
		// Case #3: Date,String. Maximum valid Excel date with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(DATE(9999,12,31), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(DATE(9999,12,31), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Date,String. Maximum valid Excel date with string fallback. 2 of 2 arguments used.');
		// Case #4: Number,String. Zero value with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFERROR(0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFERROR(0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number,String. Zero value with string fallback. 2 of 2 arguments used.');

		// Need to fix: area handle, result can be ref or not?, diff result in boundary cases from MS
		// Case #11: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.
		// Case #22: Error,Area. Error input with single-cell range fallback. 2 of 2 arguments used.
		// Case #11: Name,String. Named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		// Case #16: Number,String. Number too large returns #NUM!, fallback to string. Returns "Error". 2 of 2 arguments used.
		// Case #19: Area3D,String. 3D multi-cell range returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		// Case #20: Name3D,String. 3D named range with text returns #VALUE!, fallback to string. Returns "Error". 2 of 2 arguments used.
		// Case #1: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.
		// Case #2: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.


		//testArrayFormula2(assert, "IFERROR", 2, 2);
	});

	QUnit.test("Test: \"IFNA\"", function (assert) {

		oParser = new parserFormula('IFNA(MATCH(30,B1:B5,0),"Not found")', "A2", ws);
		assert.ok(oParser.parse(), 'IFNA(MATCH(30,B1:B5,0),"Not found")');
		assert.strictEqual(oParser.calculate().getValue(), "Not found", 'IFNA(MATCH(30,B1:B5,0),"Not found")');

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("#N/A");
		ws.getRange2("A101").setValue("#NUM!");
		ws.getRange2("A104").setValue("3");
		// For area
		ws.getRange2("A102").setValue("1");
		ws.getRange2("A103").setValue("2");
		ws.getRange2("A105").setValue("#DIV/0!");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
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
		// Case #1: Number,String. Basic valid number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(10, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(10, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Number,String. Basic valid number with string fallback. 2 of 2 arguments used.');
		// Case #2: Number,Number. Float input with numeric fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(1.5, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(1.5, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.5, 'Test: Positive case: Number,Number. Float input with numeric fallback. 2 of 2 arguments used.');
		// Case #3: String,String. String convertible to number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA("NA", "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA("NA", "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NA', 'Test: Positive case: String,String. String convertible to number with string fallback. 2 of 2 arguments used.');
		// Case #4: Formula,String. Nested formula with valid result and string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(SQRT(4), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(SQRT(4), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,String. Nested formula with valid result and string fallback. 2 of 2 arguments used.');
		// Case #5: Reference link,String. Reference to cell with valid number. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A100, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A100, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Positive case: Reference link,String. Reference to cell with valid number. 2 of 2 arguments used.');
		// Case #6: Area,String. Single-cell range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A101:A101, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A101:A101, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Area,String. Single-cell range with string fallback. 2 of 2 arguments used.');
		// Case #7: Array,String. Array with single element and string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA({2}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA({2}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 2, 'Test: Positive case: Array,String. Array with single element and string fallback. 2 of 2 arguments used.');
		// Case #8: Name,String. Named range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TestName, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TestName, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), -0.5, 'Test: Positive case: Name,String. Named range with string fallback. 2 of 2 arguments used.');
		// Case #9: Name3D,String. 3D named range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TestName3D, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TestName3D, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), -0.5, 'Test: Positive case: Name3D,String. 3D named range with string fallback. 2 of 2 arguments used.');
		// Case #10: Ref3D,String. 3D reference to cell with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Sheet2!A1, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Sheet2!A1, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Test: Positive case: Ref3D,String. 3D reference to cell with string fallback. 2 of 2 arguments used.');
		// Case #11: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Sheet2!A2:A2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Sheet2!A2:A2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.');
		// Case #12: Table,String. Table structured reference with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Table1[Column1], "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Table1[Column1], "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Test: Positive case: Table,String. Table structured reference with string fallback. 2 of 2 arguments used.');
		// Case #13: Date,String. Date as serial number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(DATE(2025,1,1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(DATE(2025,1,1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Date,String. Date as serial number with string fallback. 2 of 2 arguments used.');
		// Case #14: Time,String. Time value with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TIME(12,0,0), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TIME(12,0,0), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Positive case: Time,String. Time value with string fallback. 2 of 2 arguments used.');
		// Case #15: Formula,Formula. Nested formula in both arguments, both valid. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(SQRT(4), ABS(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(SQRT(4), ABS(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,Formula. Nested formula in both arguments, both valid. 2 of 2 arguments used.');
		// Case #16: Number,Formula. Number with nested IF formula as fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(1, IF(TRUE, 0, 1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(1, IF(TRUE, 0, 1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number,Formula. Number with nested IF formula as fallback. 2 of 2 arguments used.');
		// Case #17: String,Number. Short date string with numeric fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA("12/12", 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA("12/12", 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12/12', 'Test: Positive case: String,Number. Short date string with numeric fallback. 2 of 2 arguments used.');
		// Case #18: Array,String. Multi-element array with string fallback. Returns array of results. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA({1, 2}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA({1, 2}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array,String. Multi-element array with string fallback. Returns array of results. 2 of 2 arguments used.');
		// Case #19: Formula,String. Nested IF returning valid value with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(IF(TRUE, 2, 0.5), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(IF(TRUE, 2, 0.5), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula,String. Nested IF returning valid value with string fallback. 2 of 2 arguments used.');
		// Case #20: Error,String. Error #N/A with string fallback. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(NA(), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(NA(), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Positive case: Error,String. Error #N/A with string fallback. Returns "Error". 2 of 2 arguments used.');
		// Case #21: Empty,String. Empty cell treated as empty string. Returns "". 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A102, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A102, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 1, 'Test: Positive case: Empty,String. Empty cell treated as empty string. Returns "". 2 of 2 arguments used.');
		// Case #22: Error,Empty. #N/A error with empty cell fallback. Returns "". 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(NA(), A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(NA(), A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 2, 'Test: Positive case: Error,Empty. #N/A error with empty cell fallback. Returns "". 2 of 2 arguments used.');

		// Negative cases:
		// Case #1: Error,Error. Error #N/A with error fallback. Returns #N/A. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(NA(), "n")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(NA(), "n") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'n', 'Test: Negative case: Error,Error. Error #N/A with error fallback. Returns #N/A. 2 of 2 arguments used.');
		// Case #2: Number,String. Division by zero returns #DIV/0!, not caught by IFNA. Returns #DIV/0!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(1/0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(1/0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Number,String. Division by zero returns #DIV/0!, not caught by IFNA. Returns #DIV/0!. 2 of 2 arguments used.');
		// Case #3: String,String. Non-numeric string returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA("abc", "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA("abc", "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String,String. Non-numeric string returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #4: Formula,String. Formula resulting in #NUM! not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(SQRT(-1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(SQRT(-1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula,String. Formula resulting in #NUM! not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.');
		// Case #5: Reference link,String. Reference to cell with #DIV/0! error. Returns #DIV/0!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A104, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A104, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 3, 'Test: Negative case: Reference link,String. Reference to cell with #DIV/0! error. Returns #DIV/0!. 2 of 2 arguments used.');
		// Case #6: Area,String. Multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A105:A106, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A105:A106, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Area,String. Multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #7: Reference link,String. Reference to cell with #VALUE! error. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(A107, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(A107, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", 'Test: Negative case: Reference link,String. Reference to cell with #VALUE! error. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #9: Boolean,String. Boolean FALSE (0) is valid. Returns 0. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(FALSE, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(FALSE, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean,String. Boolean FALSE (0) is valid. Returns 0. 2 of 2 arguments used.');
		// Case #10: Ref3D,String. 3D reference to text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Sheet2!A3, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Sheet2!A3, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), 'Text', 'Test: Negative case: Ref3D,String. 3D reference to text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #11: Name,String. Named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TestNameArea2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TestNameArea2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), -0.8, 'Test: Negative case: Name,String. Named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #12: Table,String. Table column with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Table1[Column2], "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Table1[Column2], "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().getValue(), '1s', 'Test: Negative case: Table,String. Table column with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #13: Formula,String. Formula resulting in #N/A with string fallback. Returns "Error". 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(VLOOKUP("x", A1:B1, 2, FALSE), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(VLOOKUP("x", A1:B1, 2, FALSE), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Error', 'Test: Negative case: Formula,String. Formula resulting in #N/A with string fallback. Returns "Error". 2 of 2 arguments used.');
		// Case #14: Number,String. Zero division returns #DIV/0!, not caught by IFNA. Returns #DIV/0!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(0/0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(0/0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Number,String. Zero division returns #DIV/0!, not caught by IFNA. Returns #DIV/0!. 2 of 2 arguments used.');
		// Case #15: Array,String. Array with boolean returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA({FALSE}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA({FALSE}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'FALSE', 'Test: Negative case: Array,String. Array with boolean returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #16: Number,String. Number too large returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(1E+307, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(1E+307, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1e+307, 'Test: Negative case: Number,String. Number too large returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.');
		// Case #17: Date,String. Invalid date returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(DATE(10000,1,1), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(DATE(10000,1,1), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Date,String. Invalid date returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.');
		// Case #18: Time,String. Invalid time returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TIME(25,0,0), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TIME(25,0,0), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.04166666666666674, 'Test: Negative case: Time,String. Invalid time returns #NUM!, not caught by IFNA. Returns #NUM!. 2 of 2 arguments used.');
		// Case #19: Area3D,String. 3D multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(Sheet2!A4:A5, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(Sheet2!A4:A5, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area3D,String. 3D multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');
		// Case #20: Name3D,String. 3D named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(TestNameArea3D2, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(TestNameArea3D2, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), -0.8, 'Test: Negative case: Name3D,String. 3D named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.');

		// Bounded cases:
		// Case #1: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(9.99999999999999E+307, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(9.99999999999999E+307, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1e+308, 'Test: Bounded case: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.');
		// Case #2: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(2.2250738585072014E-308, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(2.2250738585072014E-308, "Error") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.');
		// Case #3: Date,String. Maximum valid Excel date with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(DATE(9999,12,31), "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(DATE(9999,12,31), "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Date,String. Maximum valid Excel date with string fallback. 2 of 2 arguments used.');
		// Case #4: Number,String. Zero value with string fallback. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA(0, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA(0, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Number,String. Zero value with string fallback. 2 of 2 arguments used.');
		// Case #5: Array,String. Array with valid and #N/A elements. Returns {1, "Error"}. 2 of 2 arguments used.
		oParser = new parserFormula('IFNA({1, "NA"}, "Error")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IFNA({1, "NA"}, "Error") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Bounded case: Array,String. Array with valid and #N/A elements. Returns {1, "Error"}. 2 of 2 arguments used.');

		// Need to fix: area handle in tests, boundary cases diff from ms
		// Case #6: Area,String. Single-cell range with string fallback. 2 of 2 arguments used.
		// Case #11: Area3D,String. 3D single-cell range with string fallback. 2 of 2 arguments used.
		// Case #6: Area,String. Multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		// Case #11: Name,String. Named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		// Case #19: Area3D,String. 3D multi-cell range returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used. - empty array
		// Case #20: Name3D,String. 3D named range with text returns #VALUE!, not caught by IFNA. Returns #VALUE!. 2 of 2 arguments used.
		// Case #1: Number,String. Maximum valid Excel number with string fallback. 2 of 2 arguments used.
		// Case #2: Number,String. Minimum positive Excel number with string fallback. 2 of 2 arguments used.


	});

	QUnit.test("Test: \"IFS\"", function (assert) {

		oParser = new parserFormula('IFS(1,"TEST")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TEST");

		oParser = new parserFormula('IFS(0,"TEST",1,"TEST2")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TEST2");

		oParser = new parserFormula('IFS(2<1,">3")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IFS(2<1,">3",2>1)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IFS(2<1,"TEST",2<1,2,4>3,"TEST2")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TEST2");

		oParser = new parserFormula('IFS(1=1,"correct",#VALUE!,1)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "correct");

		oParser = new parserFormula('IFS(1<>1,"correct",#VALUE!,1)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('IFS(1<>1,"correct",#N/A,1)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IFS(1<>1,"correct",1=1, #N/A)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IFS(1<>1,"correct",1=1, #N/A, #VALUE!, #VALUE!)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula('IFS(1=1,"correct",1=2,1/0)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "correct");

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A", "IFS(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2,A1:C2)");
	});

	QUnit.test("Test: \"NOT\"", function (assert) {

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
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
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
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:
		// Case #1: Boolean. Basic valid input: Boolean TRUE returns FALSE.
		oParser = new parserFormula('NOT(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Basic valid input: Boolean TRUE returns FALSE.');
		// Case #2: Boolean. Basic valid input: Boolean FALSE returns TRUE.
		oParser = new parserFormula('NOT(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Basic valid input: Boolean FALSE returns TRUE.');
		// Case #3: Number. Number 0 is coerced to FALSE, returns TRUE.
		oParser = new parserFormula('NOT(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Number 0 is coerced to FALSE, returns TRUE.');
		// Case #4: Number. Number 1 is coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Number 1 is coerced to TRUE, returns FALSE.');
		// Case #5: Number. Non-zero number is coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number. Non-zero number is coerced to TRUE, returns FALSE.');
		// Case #6: String. String "TRUE" is coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT("TRUE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT("TRUE") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String. String "TRUE" is coerced to TRUE, returns FALSE.');
		// Case #7: String. String "FALSE" is coerced to FALSE, returns TRUE.
		oParser = new parserFormula('NOT("FALSE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT("FALSE") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. String "FALSE" is coerced to FALSE, returns TRUE.');
		// Case #8: Formula. Nested IF formula returns TRUE, NOT returns FALSE.
		oParser = new parserFormula('NOT(IF(TRUE,TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(IF(TRUE,TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Nested IF formula returns TRUE, NOT returns FALSE.');
		// Case #9: Formula. Nested AND formula returns FALSE, NOT returns TRUE.
		oParser = new parserFormula('NOT(AND(TRUE,FALSE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(AND(TRUE,FALSE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested AND formula returns FALSE, NOT returns TRUE.');
		// Case #10: Reference link. Reference to cell with TRUE, returns FALSE.
		oParser = new parserFormula('NOT(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with TRUE, returns FALSE.');
		// Case #11: Reference link. Reference to cell with FALSE, returns TRUE.
		oParser = new parserFormula('NOT(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cell with FALSE, returns TRUE.');
		// Case #12: Area. Single-cell range with TRUE, returns FALSE.
		oParser = new parserFormula('NOT(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell range with TRUE, returns FALSE.');
		// Case #13: Array. Array with single TRUE element, returns FALSE.
		oParser = new parserFormula('NOT({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with single TRUE element, returns FALSE.');
		// Case #14: Array. Array with single FALSE element, returns TRUE.
		oParser = new parserFormula('NOT({FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT({FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with single FALSE element, returns TRUE.');
		// Case #15: Name. Named range with TRUE, returns FALSE.
		oParser = new parserFormula('NOT(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named range with TRUE, returns FALSE.');
		// Case #16: Name3D. 3D named range with FALSE, returns TRUE.
		oParser = new parserFormula('NOT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name3D. 3D named range with FALSE, returns TRUE.');
		// Case #17: Ref3D. 3D reference to cell with TRUE, returns FALSE.
		oParser = new parserFormula('NOT(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D reference to cell with TRUE, returns FALSE.');
		// Case #18: Area3D. 3D single-cell range with FALSE, returns TRUE.
		oParser = new parserFormula('NOT(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell range with FALSE, returns TRUE.');
		// Case #19: Table. Table column with TRUE, returns FALSE.
		oParser = new parserFormula('NOT(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table column with TRUE, returns FALSE.');
		// Case #20: Formula. Formula evaluates to TRUE, NOT returns FALSE.
		oParser = new parserFormula('NOT(SUM(1,1)=2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(SUM(1,1)=2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. Formula evaluates to TRUE, NOT returns FALSE.');
		// Case #21: Date. Date serial number coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Date. Date serial number coerced to TRUE, returns FALSE.');
		// Case #22: Time. Time serial number coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Time. Time serial number coerced to TRUE, returns FALSE.');

		// Negative cases:
		// Case #1: Number. Non-zero negative number coerced to TRUE, returns FALSE (not an error but unexpected for non-logical input).
		oParser = new parserFormula('NOT(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Non-zero negative number coerced to TRUE, returns FALSE (not an error but unexpected for non-logical input).');
		// Case #2: String. Non-logical string returns #VALUE! error.
		oParser = new parserFormula('NOT("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-logical string returns #VALUE! error.');
		// Case #3: Empty. Reference to empty cell returns #VALUE! error.
		oParser = new parserFormula('NOT(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE! error.');
		// Case #4: Error. Error input propagates #N/A error.
		oParser = new parserFormula('NOT(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input propagates #N/A error.');
		// Case #5: Area. Multi-cell range returns #VALUE! error.
		oParser = new parserFormula('NOT(A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A104:A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Multi-cell range returns #VALUE! error.');
		// Case #6: Array. Multi-element array returns #VALUE! error.
		oParser = new parserFormula('NOT({TRUE,FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT({TRUE,FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Multi-element array returns #VALUE! error.');
		// Case #7: String. Empty string returns #VALUE! error.
		oParser = new parserFormula('NOT("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE! error.');
		// Case #8: Reference link. Reference to cell with text returns #VALUE! error.
		oParser = new parserFormula('NOT(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. Reference to cell with text returns #VALUE! error.');
		// Case #9: Area. Single-cell range with text returns #VALUE! error.
		oParser = new parserFormula('NOT(A105:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(A105:A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Single-cell range with text returns #VALUE! error.');
		// Case #10: Name. Named range with text returns #VALUE! error.
		oParser = new parserFormula('NOT(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with text returns #VALUE! error.');
		// Case #11: Name3D. 3D named range with text returns #VALUE! error.
		oParser = new parserFormula('NOT(TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TestNameArea3D) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with text returns #VALUE! error.');
		// Case #12: Ref3D. 3D reference to cell with text returns #VALUE! error.
		oParser = new parserFormula('NOT(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(Sheet2!A3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to cell with text returns #VALUE! error.');
		// Case #13: Area3D. 3D multi-cell range returns #VALUE! error.
		oParser = new parserFormula('NOT(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(Sheet2!A4:A5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE! error.');
		// Case #15: Number. Excessively large number coerced to TRUE, returns FALSE (not an error but unexpected).
		oParser = new parserFormula('NOT(1E+308)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(1E+308) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Excessively large number coerced to TRUE, returns FALSE (not an error but unexpected).');
		// Case #16: Formula. Formula resulting in #NUM! error propagates error.
		oParser = new parserFormula('NOT(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates error.');
		// Case #17: String. String convertible to number coerced to TRUE, returns FALSE (not an error but unexpected).
		oParser = new parserFormula('NOT("0.5")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT("0.5") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String convertible to number coerced to TRUE, returns FALSE (not an error but unexpected).');
		// Case #18: Date. Invalid date (before Excel\'s date system) may cause #NUM! or coerced behavior.
		oParser = new parserFormula('NOT(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(DATE(1899,12,31)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Date. Invalid date (before Excel\'s date system) may cause #NUM! or coerced behavior.');
		// Case #19: Time. Time value 0 coerced to FALSE, returns TRUE (not an error but unexpected).
		oParser = new parserFormula('NOT(TIME(0,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(TIME(0,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Time. Time value 0 coerced to FALSE, returns TRUE (not an error but unexpected).');
		// Case #20: Formula. Formula resulting in #DIV/0! error propagates error.
		oParser = new parserFormula('NOT(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula. Formula resulting in #DIV/0! error propagates error.');

		// Bounded cases:
		// Case #1: Number. Minimum valid number coerced to FALSE, returns TRUE.
		oParser = new parserFormula('NOT(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Minimum valid number coerced to FALSE, returns TRUE.');
		// Case #2: Number. Minimum non-zero number coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Minimum non-zero number coerced to TRUE, returns FALSE.');
		// Case #3: Number. Maximum valid Excel number coerced to TRUE, returns FALSE.
		oParser = new parserFormula('NOT(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NOT(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Maximum valid Excel number coerced to TRUE, returns FALSE.');

		// Need to fix: area/array returns, range single-cell
		// Case #5: Area. Multi-cell range returns #VALUE! error.
		// Case #9: Area. Single-cell range with text returns #VALUE! error.
		// Case #10: Name. Named range with text returns #VALUE! error.
		// Case #11: Name3D. 3D named range with text returns #VALUE! error.
		// Case #12: Ref3D. 3D reference to cell with text returns #VALUE! error.
		// Case #13: Area3D. 3D multi-cell range returns #VALUE! error.

		testArrayFormula2(assert, "NOT", 1, 1);
	});

	QUnit.test("Test: \"OR\"", function (assert) {

		ws.getRange2("A2").setValue("50");
		ws.getRange2("A3").setValue("100");

		oParser = new parserFormula("AND(A2>1,A2<100)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('AND(A2<A3,A2<100)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('AND(A3<1,A3>100)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");


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
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
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
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2

		// Positive cases:
		// Case #1: Number. Non-zero number as TRUE, zero as FALSE. 2 arguments used.
		oParser = new parserFormula('OR(1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Non-zero number as TRUE, zero as FALSE. 2 arguments used.');
		// Case #2: Number. Negative number as TRUE, zero as FALSE. 2 arguments used.
		oParser = new parserFormula('OR(-1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(-1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Negative number as TRUE, zero as FALSE. 2 arguments used.');
		// Case #3: String. String "TRUE" or "FALSE" as logical values. 2 arguments used.
		oParser = new parserFormula('OR("TRUE","FALSE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR("TRUE","FALSE") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. String "TRUE" or "FALSE" as logical values. 2 arguments used.');
		// Case #4: String. Numeric strings converted to numbers (TRUE/FALSE). 2 arguments used.
		oParser = new parserFormula('OR("1","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR("1","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. Numeric strings converted to numbers (TRUE/FALSE). 2 arguments used.');
		// Case #5: Formula. Logical expressions via formulas. 2 arguments used.
		oParser = new parserFormula('OR(1>0,2<1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(1>0,2<1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Logical expressions via formulas. 2 arguments used.');
		// Case #6: Formula. Nested IF formula returning number. 2 arguments used.
		oParser = new parserFormula('OR(IF(TRUE,1,0),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(IF(TRUE,1,0),0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF formula returning number. 2 arguments used.');
		// Case #7: Reference link. References to cells with logical values. 2 arguments used.
		oParser = new parserFormula('OR(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. References to cells with logical values. 2 arguments used.');
		// Case #8: Area. Single-cell ranges with numbers. 2 arguments used.
		oParser = new parserFormula('OR(A102:A102,A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(A102:A102,A103:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Single-cell ranges with numbers. 2 arguments used.');
		// Case #9: Array. Arrays with logical values. 2 arguments used.
		oParser = new parserFormula('OR({TRUE,FALSE},{FALSE,TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR({TRUE,FALSE},{FALSE,TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Arrays with logical values. 2 arguments used.');
		// Case #10: Name. Named ranges with logical values. 2 arguments used.
		oParser = new parserFormula('OR(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named ranges with logical values. 2 arguments used.');
		// Case #11: Name3D. 3D named ranges with logical values. 2 arguments used.
		oParser = new parserFormula('OR(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named ranges with logical values. 2 arguments used.');
		// Case #12: Ref3D. 3D references to cells with numbers. 2 arguments used.
		oParser = new parserFormula('OR(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D references to cells with numbers. 2 arguments used.');
		// Case #13: Area3D. 3D single-cell ranges with logical values. 2 arguments used.
		oParser = new parserFormula('OR(Sheet2!A3:A3,Sheet2!A4:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(Sheet2!A3:A3,Sheet2!A4:A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell ranges with logical values. 2 arguments used.');
		// Case #14: Table. Table structured references with logical values. 2 arguments used.
		oParser = new parserFormula('OR(Table1[Column1],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(Table1[Column1],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table structured references with logical values. 2 arguments used.');
		// Case #15: Date. Date comparison as logical value. 2 arguments used.
		oParser = new parserFormula('OR(DATE(2025,1,1)>DATE(2024,1,1),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(DATE(2025,1,1)>DATE(2024,1,1),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Date. Date comparison as logical value. 2 arguments used.');
		// Case #16: Time. Time comparison as logical value. 2 arguments used.
		oParser = new parserFormula('OR(TIME(12,0,0)>TIME(11,0,0),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TIME(12,0,0)>TIME(11,0,0),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Time comparison as logical value. 2 arguments used.');
		// Case #17: Formula. Multiple formula comparisons. 2 arguments used.
		oParser = new parserFormula('OR(SUM(1,1)=2,PRODUCT(2,2)=4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(SUM(1,1)=2,PRODUCT(2,2)=4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Multiple formula comparisons. 2 arguments used.');
		// Case #18: Empty. Empty cell reference treated as FALSE. 2 arguments used.
		oParser = new parserFormula('OR(A104,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(A104,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Empty. Empty cell reference treated as FALSE. 2 arguments used.');
		// Case #19: Array. Arrays with numbers treated as logical values. 2 arguments used.
		oParser = new parserFormula('OR({1,0},{0,1})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR({1,0},{0,1}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Arrays with numbers treated as logical values. 2 arguments used.');
		// Case #20: Formula. Nested AND formula as logical value. 2 arguments used.
		oParser = new parserFormula('OR(AND(1>0,2>1),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(AND(1>0,2>1),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested AND formula as logical value. 2 arguments used.');
		// Case #21: Number. Large number as TRUE. 2 arguments used.
		oParser = new parserFormula('OR(1E+307,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(1E+307,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Large number as TRUE. 2 arguments used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 2 arguments used.
		oParser = new parserFormula('OR(NA(),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(NA(),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 2 arguments used.');
		// Case #2: Error. Propagates #DIV/0! error. 2 arguments used.
		oParser = new parserFormula('OR(DIV0(),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(DIV0(),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Error. Propagates #DIV/0! error. 2 arguments used.');
		// Case #3: String. Non-logical string returns TRUE. 2 arguments used.
		oParser = new parserFormula('OR("abc",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR("abc",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. Non-logical string returns #VALUE!. 2 arguments used.');
		// Case #4: Formula. Formula returning #NUM! propagates error. 2 arguments used.
		oParser = new parserFormula('OR(SQRT(-1),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(SQRT(-1),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula returning #NUM! propagates error. 2 arguments used.');
		// Case #5: Area. Multi-cell range returns TRUE. 2 arguments used.
		oParser = new parserFormula('OR(A105:A106,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(A105:A106,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #6: Area3D. 3D multi-cell range returns TRUE. 2 arguments used.
		oParser = new parserFormula('OR(Sheet2!A5:A6,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(Sheet2!A5:A6,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #7: Name. Named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(TestNameArea,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TestNameArea,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #8: Name3D. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(TestNameArea3D2,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TestNameArea3D2,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name3D. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #10: Ref3D. 3D reference to text returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(Sheet2!A7,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(Sheet2!A7,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D reference to text returns #VALUE!. 2 arguments used.');
		// Case #11: String. Empty string returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR("",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR("",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String. Empty string returns #VALUE!. 2 arguments used.');
		// Case #12: Formula. Formula causing #DIV/0! propagates error. 2 arguments used.
		oParser = new parserFormula('OR(1/0,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(1/0,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula. Formula causing #DIV/0! propagates error. 2 arguments used.');
		// Case #13: Number. Number beyond Excel\'s limit returns #NUM!. 2 arguments used.
		oParser = new parserFormula('OR(-1E+307,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(-1E+307,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Number. Number beyond Excel\'s limit returns #NUM!. 2 arguments used.');
		// Case #14: Array. Array with error value propagates error. 2 arguments used.
		oParser = new parserFormula('OR({#VALUE!},TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR({#VALUE!},TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with error value propagates error. 2 arguments used.');
		// Case #15: Reference link. Reference to cell with text returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(A107,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(A107,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link. Reference to cell with text returns #VALUE!. 2 arguments used.');
		// Case #16: Formula. Formula returning error propagates error. 2 arguments used.
		oParser = new parserFormula('OR(ERROR.TYPE(1),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(ERROR.TYPE(1),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula returning error propagates error. 2 arguments used.');
		// Case #17: Date. Date before Excel\'s valid range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(DATE(1899,12,31),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(DATE(1899,12,31),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Date. Date before Excel\'s valid range returns #VALUE!. 2 arguments used.');
		// Case #18: Time. Invalid time returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR(TIME(25,0,0),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(TIME(25,0,0),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Time. Invalid time returns #VALUE!. 2 arguments used.');
		// Case #19: Array. Array with non-logical string returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('OR({TRUE,"abc"},TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR({TRUE,"abc"},TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array. Array with non-logical string returns #VALUE!. 2 arguments used.');
		// Case #20: Formula. Formula returning #N/A propagates error. 2 arguments used.
		oParser = new parserFormula('OR(VLOOKUP("x",A1:A2,1,FALSE),TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(VLOOKUP("x",A1:A2,1,FALSE),TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula returning #N/A propagates error. 2 arguments used.');

		// Bounded cases:
		// Case #1: Number. Smallest positive number as TRUE. 2 arguments used.
		oParser = new parserFormula('OR(1E-307,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(1E-307,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest positive number as TRUE. 2 arguments used.');
		// Case #2: Number. Largest valid Excel number as TRUE. 2 arguments used.
		oParser = new parserFormula('OR(9.99999999999999E+307,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(9.99999999999999E+307,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Largest valid Excel number as TRUE. 2 arguments used.');
		// Case #3: String. String of smallest positive number as TRUE. 2 arguments used.
		oParser = new parserFormula('OR("1E-307","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR("1E-307","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String. String of smallest positive number as TRUE. 2 arguments used.');
		// Case #4: Date. Earliest valid Excel date as TRUE. 2 arguments used.
		oParser = new parserFormula('OR(DATE(1900,1,1),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(DATE(1900,1,1),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Earliest valid Excel date as TRUE. 2 arguments used.');
		// Case #5: Date. Latest valid Excel date as TRUE. 2 arguments used.
		oParser = new parserFormula('OR(DATE(9999,12,31),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(DATE(9999,12,31),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Date. Latest valid Excel date as TRUE. 2 arguments used.');

		// Need to fix: string handle
		// Case #3: String. String "TRUE" or "FALSE" as logical values. 2 arguments used.
		// Case #8: Area. Single-cell ranges with numbers. 2 arguments used.
		// Case #13: Area3D. 3D single-cell ranges with logical values. 2 arguments used.
		// Case #3: String. Non-logical string returns #VALUE!. 2 arguments used.
		// Case #11: String. Empty string returns #VALUE!. 2 arguments used.
		// Case #14: Array. Array with error value propagates error. 2 arguments used.

		testArrayFormula2(assert, "OR", 1, 8, null, true);
	});    
	
	QUnit.test("Test: \"SWITCH\"", function (assert) {
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("99");
		ws.getRange2("A4").setValue("99");
		ws.getRange2("A5").setValue("2");
		ws.getRange2("A6").setValue("3");

		oParser = new parserFormula('SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")');
		assert.strictEqual(oParser.calculate().getValue(), "Monday", 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")');

		oParser = new parserFormula('SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")');

		oParser = new parserFormula('SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")');
		assert.strictEqual(oParser.calculate().getValue(), "No match", 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")');

		oParser = new parserFormula('SWITCH(A5,1,"Sunday",7,"Saturday","weekday")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")');
		assert.strictEqual(oParser.calculate().getValue(), "weekday", 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")');

		oParser = new parserFormula('SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")');
		assert.strictEqual(oParser.calculate().getValue(), "Tuesday", 'SWITCH(A6,1,"Sunday",2,"Monday",3,"Tuesday","No match")');

		oParser = new parserFormula('SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")');
		assert.strictEqual(oParser.calculate().getValue(), "No match", 'SWITCH(122,1,"Sunday",2,"Monday",3,"Tuesday","No match")');

		oParser = new parserFormula('SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'SWITCH({1,"2asd",3},{12,2,3},{"asd",2,3,4})');

		oParser = new parserFormula('SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")', "A1", ws);
		assert.ok(oParser.parse(), 'SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")');
		assert.strictEqual(oParser.calculate().getValue(), "sdf", 'SWITCH({"asd1","2asd",3},{"asd1",1,3},"sdf")');

		testArrayFormulaEqualsValues(assert, "1,3.123,-4,#N/A;2,4,5,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,1,1,#N/A;1,1,1,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:A1,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert, "1,1,1,#N/A;2,2,2,#N/A;#N/A,#N/A,#N/A,#N/A", "SWITCH(A1:C2,A1:C2,A1:A2,A1:C2,A1:A2,A1:C2)");

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

		// Positive cases
		// Case #1: Number. Expression and values as numbers, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Number. Expression and values as numbers, returns 100. 2 value-result pairs with default.');
		// Case #2: String. Expression and values as strings, returns "Apple". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH("A","A","Apple","B","Banana","Fruit")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("A","A","Apple","B","Banana","Fruit") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Apple', 'Test: Positive case: String. Expression and values as strings, returns "Apple". 2 value-result pairs with default.');
		// Case #3: Formula. Expression as formula evaluating to 2, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(SQRT(4),2,100,3,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(SQRT(4),2,100,3,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Formula. Expression as formula evaluating to 2, returns 100. 2 value-result pairs with default.');
		// Case #4: Reference link. Expression as reference link to cell with value 1, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(A100,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A100,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Reference link. Expression as reference link to cell with value 1, returns 100. 2 value-result pairs with default.');
		// Case #5: Area. Expression as single-cell range with value 2, returns 200. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(A101:A101,2,200,3,300,400)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A101:A101,2,200,3,300,400) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 400, 'Test: Positive case: Area. Expression as single-cell range with value 2, returns 200. 2 value-result pairs with default.');
		// Case #6: Array. Expression and values as single-element arrays, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH({1},{1},100,{2},200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH({1},{1},100,{2},200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Array. Expression and values as single-element arrays, returns 100. 2 value-result pairs with default.');
		// Case #7: Name. Expression as named range with value 1, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TestName,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TestName,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Name. Expression as named range with value 1, returns 100. 2 value-result pairs with default.');
		// Case #8: Name3D. Expression as 3D named range with value 1, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TestName3D,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TestName3D,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Name3D. Expression as 3D named range with value 1, returns 100. 2 value-result pairs with default.');
		// Case #9: Ref3D. Expression as 3D reference with value 1, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(Sheet2!A1,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Sheet2!A1,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Ref3D. Expression as 3D reference with value 1, returns 100. 2 value-result pairs with default.');
		// Case #10: Area3D. Expression as 3D single-cell range with value 2, returns 200. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(Sheet2!A2:A2,2,200,3,300,400)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Sheet2!A2:A2,2,200,3,300,400) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 200, 'Test: Positive case: Area3D. Expression as 3D single-cell range with value 2, returns 200. 2 value-result pairs with default.');
		// Case #11: Table. Expression as table reference with value 1, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(Table1[Column1],1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Table1[Column1],1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Table. Expression as table reference with value 1, returns 100. 2 value-result pairs with default.');
		// Case #12: Date. Expression as date (serial number 38777), returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(DATE(2025,1,1),38777,100,38838,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(DATE(2025,1,1),38777,100,38838,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Date. Expression as date (serial number 38777), returns 100. 2 value-result pairs with default.');
		// Case #13: Time. Expression as time (0.5), returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TIME(12,0,0),0.5,100,0.75,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TIME(12,0,0),0.5,100,0.75,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Time. Expression as time (0.5), returns 100. 2 value-result pairs with default.');
		// Case #14: Formula,String,Number. Expression as formula, values as numbers, results as strings, returns "One". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(IF(TRUE,1,2),1,"One",2,"Two","Default")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(IF(TRUE,1,2),1,"One",2,"Two","Default") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'One', 'Test: Positive case: Formula,String,Number. Expression as formula, values as numbers, results as strings, returns "One". 2 value-result pairs with default.');
		// Case #15: Formula. SWITCH inside SUM, returns 150 (100+50). 2 value-result pairs with default.
		oParser = new parserFormula('SUM(SWITCH(1,1,100,2,200,300),50)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SWITCH(1,1,100,2,200,300),50) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 150, 'Test: Positive case: Formula. SWITCH inside SUM, returns 150 (100+50). 2 value-result pairs with default.');
		// Case #16: Number,String. Expression and values as numbers, default as string, returns "NoMatch". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(3,1,100,2,200,"NoMatch")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(3,1,100,2,200,"NoMatch") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NoMatch', 'Test: Positive case: Number,String. Expression and values as numbers, default as string, returns "NoMatch". 2 value-result pairs with default.');
		// Case #17: Array. Expression and values as arrays, returns 200. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH({2},{1},100,{2},200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH({2},{1},100,{2},200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 200, 'Test: Positive case: Array. Expression and values as arrays, returns 200. 2 value-result pairs with default.');
		// Case #18: Number,String,Formula. Expression as number, results as formulas, returns date 38777. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,1,DATE(2025,1,1),2,DATE(2025,2,1),DATE(2025,3,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,1,DATE(2025,1,1),2,DATE(2025,2,1),DATE(2025,3,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Number,String,Formula. Expression as number, results as formulas, returns date 38777. 2 value-result pairs with default.');
		// Case #19: Empty. Empty expression treated as 0, returns 300 (default). 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(,,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(,,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Empty. Empty expression treated as 0, returns 300 (default). 2 value-result pairs with default.');
		// Case #20: Reference link,String. Expression as reference link to string "A", returns "Apple". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(A102,"A","Apple","B","Banana","Fruit")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A102,"A","Apple","B","Banana","Fruit") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Fruit', 'Test: Positive case: Reference link,String. Expression as reference link to string "A", returns "Apple". 2 value-result pairs with default.');
		// Case #21: Number. Single value-result pair, no default, returns 100.
		oParser = new parserFormula('SWITCH(1,1,100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,1,100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Positive case: Number. Single value-result pair, no default, returns 100.');
		// Case #22: String,Number. Expression and values as strings, results as numbers, returns 1. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH("X","X",1,"Y",2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("X","X",1,"Y",2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String,Number. Expression and values as strings, results as numbers, returns 1. 2 value-result pairs with default.');
		// Case #23: Number,String. Expression and values as numbers, one result as string, returns "Two". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(2,1,100,2,"Two",300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(2,1,100,2,"Two",300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Two', 'Test: Positive case: Number,String. Expression and values as numbers, one result as string, returns "Two". 2 value-result pairs with default.');
		// Case #24: Formula,Number,String. Expression as formula evaluating to 0, returns "Zero". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(MOD(4,2),0,"Zero",1,"One","Default")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(MOD(4,2),0,"Zero",1,"One","Default") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Zero', 'Test: Positive case: Formula,Number,String. Expression as formula evaluating to 0, returns "Zero". 2 value-result pairs with default.');
		// Case #25: Reference link,Number,String. Expression as reference link to value 3, returns "Three". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(A103,3,"Three",4,"Four","Default")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A103,3,"Three",4,"Four","Default") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Default', 'Test: Positive case: Reference link,Number,String. Expression as reference link to value 3, returns "Three". 2 value-result pairs with default.');

		// Negative cases:

		// Case #1: Error. Expression as error, returns #N/A. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(NA(),1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(NA(),1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Expression as error, returns #N/A. 2 value-result pairs with default.');
		// Case #2: Boolean. Expression as boolean, no match, returns #N/A (no default). 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH(TRUE,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TRUE,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Boolean. Expression as boolean, no match, returns #N/A (no default). 2 value-result pairs, no default.');
		// Case #3: String. No match and no default, returns #N/A. 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH("A","B","Banana","C","Cherry",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("A","B","Banana","C","Cherry",) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. No match and no default, returns #N/A. 2 value-result pairs, no default.');
		// Case #4: Number,String. Expression as number, values as strings, no match, returns "Default". 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,"1","One",2,"Two","Default")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,"1","One",2,"Two","Default") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Default', 'Test: Negative case: Number,String. Expression as number, values as strings, no match, returns "Default". 2 value-result pairs with default.');
		// Case #5: Area. Expression as multi-cell range, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(A104:A105,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A104:A105,1,100,2,200,300) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Area. Expression as multi-cell range, returns #VALUE!. 2 value-result pairs with default.');
		// Case #6: Area3D. Expression as 3D multi-cell range, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(Sheet2!A3:A4,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Sheet2!A3:A4,1,100,2,200,300) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Area3D. Expression as 3D multi-cell range, returns #VALUE!. 2 value-result pairs with default.');
		// Case #7: Name. Expression as named range with area, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TestNameArea2,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TestNameArea2,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Name. Expression as named range with area, returns #VALUE!. 2 value-result pairs with default.');
		// Case #8: Name3D. Expression as 3D named range with area, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TestNameArea3D2,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TestNameArea3D2,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Name3D. Expression as 3D named range with area, returns #VALUE!. 2 value-result pairs with default.');
		// Case #9: String. Empty string expression, no match, no default, returns #N/A. 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH("","A","Apple","B","Banana",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("","A","Apple","B","Banana",) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Empty string expression, no match, no default, returns #N/A. 2 value-result pairs, no default.');
		// Case #10: Formula. Expression as formula returning #NUM!, propagates error. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(SQRT(-1),1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(SQRT(-1),1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Expression as formula returning #NUM!, propagates error. 2 value-result pairs with default.');
		// Case #11: Number. Expression exceeds Excelâ??s number limit, returns #NUM!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1E+308,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1E+308,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Number. Expression exceeds Excelâ??s number limit, returns #NUM!. 2 value-result pairs with default.');
		// Case #12: Reference link. Reference link to empty cell, returns #N/A (no default). 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH(A106,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(A106,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Reference link. Reference link to empty cell, returns #N/A (no default). 2 value-result pairs, no default.');
		// Case #13: Ref3D. 3D reference to text string, returns #N/A (no default). 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH(Sheet2!A5,1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Sheet2!A5,1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Ref3D. 3D reference to text string, returns #N/A (no default). 2 value-result pairs, no default.');
		// Case #14: Table. Table reference to text string, returns #N/A (no default). 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH(Table1[Column2],1,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(Table1[Column2],1,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Table. Table reference to text string, returns #N/A (no default). 2 value-result pairs, no default.');
		// Case #15: Number,Boolean. Values as booleans, no match, returns 300 (default). 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,TRUE,100,FALSE,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,TRUE,100,FALSE,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Number,Boolean. Values as booleans, no match, returns 300 (default). 2 value-result pairs with default.');
		// Case #16: Array. Value as array with booleans, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,{TRUE,FALSE},100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,{TRUE,FALSE},100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Array. Value as array with booleans, returns #VALUE!. 2 value-result pairs with default.');
		// Case #17: Date. Invalid date in expression, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(DATE(2025,13,1),38777,100,38838,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(DATE(2025,13,1),38777,100,38838,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Date. Invalid date in expression, returns #VALUE!. 2 value-result pairs with default.');
		// Case #18: Time. Invalid time in expression, returns #VALUE!. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(TIME(25,0,0),0.5,100,0.75,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(TIME(25,0,0),0.5,100,0.75,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Time. Invalid time in expression, returns #VALUE!. 2 value-result pairs with default.');
		// Case #19: Number,String. Empty string as value, no match, returns 300 (default). 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,"",100,"2",200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,"",100,"2",200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Number,String. Empty string as value, no match, returns 300 (default). 2 value-result pairs with default.');
		// Case #20: Formula,Error. Value as error, returns #N/A. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,NA(),100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,NA(),100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula,Error. Value as error, returns #N/A. 2 value-result pairs with default.');
		// Case #21: String. Non-matching text string expression, no default, returns #N/A. 2 value-result pairs, no default.
		oParser = new parserFormula('SWITCH("Str","A","Apple","B","Banana",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("Str","A","Apple","B","Banana",) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Non-matching text string expression, no default, returns #N/A. 2 value-result pairs, no default.');
		// Case #22: Number,String. Result as error, no match, returns 300 (default). 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1,"1",NA(),2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1,"1",NA(),2,200,300) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Negative case: Number,String. Result as error, no match, returns 300 (default). 2 value-result pairs with default.');

		// Bounded cases:
		// Case #1: Number. Minimum positive number as expression and value, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(1E-307,1E-307,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(1E-307,1E-307,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Bounded case: Number. Minimum positive number as expression and value, returns 100. 2 value-result pairs with default.');
		// Case #2: Number. Maximum Excel number as expression and value, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH(9.99999999999999E+307,9.99999999999999E+307,100,2,200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH(9.99999999999999E+307,9.99999999999999E+307,100,2,200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Bounded case: Number. Maximum Excel number as expression and value, returns 100. 2 value-result pairs with default.');
		// Case #3: String. Empty string as expression and value, returns 100. 2 value-result pairs with default.
		oParser = new parserFormula('SWITCH("","",100,"A",200,300)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SWITCH("","",100,"A",200,300) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 100, 'Test: Bounded case: String. Empty string as expression and value, returns 100. 2 value-result pairs with default.');

		// Need to fix: error handle, area/3d handle, string handle
		// Case #3: String. No match and no default, returns #N/A. 2 value-result pairs, no default.
		// Case #5: Area. Expression as multi-cell range, returns #VALUE!. 2 value-result pairs with default.
		// Case #6: Area3D. Expression as 3D multi-cell range, returns #VALUE!. 2 value-result pairs with default.
		// Case #9: String. Empty string expression, no match, no default, returns #N/A. 2 value-result pairs, no default.
		// Case #21: String. Non-matching text string expression, no default, returns #N/A. 2 value-result pairs, no default.
		// Case #22: Number,String. Result as error, no match, returns 300 (default). 2 value-result pairs with default.


	});

	QUnit.test("Test: \"XOR\"", function (assert) {
		oParser = new parserFormula('XOR(3>0,2<9)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(3>0,2<9)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'XOR(3>0,2<9)');

		oParser = new parserFormula('XOR(3>12,4>6)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(3>12,4>6)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", 'XOR(3>12,4>6)');

		oParser = new parserFormula('XOR(3>12,4<6)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(3>12,4<6)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", 'XOR(3>12,4<6)');

		//area - specific for xor function
		//all empty - false result
		ws.getRange2("A101").setValue("5");
		ws.getRange2("A102").setValue("6");
		ws.getRange2("A103").setValue("test1");
		ws.getRange2("A104").setValue("");
		ws.getRange2("A105").setValue("false");

		ws.getRange2("B101").setValue("1");
		ws.getRange2("B102").setValue("1");
		ws.getRange2("B103").setValue("test2");
		ws.getRange2("B104").setValue("");
		ws.getRange2("B105").setValue("false");

		ws.getRange2("B106").setValue("#VALUE!");

		oParser = new parserFormula('XOR(A101:B102)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:B102)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('XOR(A101:B103)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:B103)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('XOR(A101:A103)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:A103)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('XOR(A101:A104)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:A104)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('XOR(A104:B104)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A104:B104)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('XOR(A101:B104)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:B104)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('XOR(A101:B105)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:B105)');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula('XOR(A101:A105)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(A101:A105)');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('XOR(B101:A106)', "A2", ws);
		assert.ok(oParser.parse(), 'XOR(B101:A106)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

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
		// Case #0: Number. Single numeric argument, coerces to TRUE. 1 argument used.
		oParser = new parserFormula('XOR(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Single numeric argument, coerces to TRUE. 1 argument used.');
		// Case #1: Number, Number. Two numeric arguments, 1=TRUE, 0=FALSE, odd TRUE count returns TRUE. 2 arguments used.
		oParser = new parserFormula('XOR(1, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(1, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number, Number. Two numeric arguments, 1=TRUE, 0=FALSE, odd TRUE count returns TRUE. 2 arguments used.');
		// Case #2: Number, Number, Number. Three numeric arguments, two TRUE (1,1), one FALSE (0), even TRUE count returns FALSE. 3 arguments used.
		oParser = new parserFormula('XOR(1, 0, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(1, 0, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Number, Number, Number. Three numeric arguments, two TRUE (1,1), one FALSE (0), even TRUE count returns FALSE. 3 arguments used.');
		// Case #3: Boolean. Single Boolean TRUE returns TRUE. 1 argument used.
		oParser = new parserFormula('XOR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Single Boolean TRUE returns TRUE. 1 argument used.');
		// Case #4: Boolean, Boolean. Two Boolean arguments, odd TRUE count returns TRUE. 2 arguments used.
		oParser = new parserFormula('XOR(TRUE, FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TRUE, FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean, Boolean. Two Boolean arguments, odd TRUE count returns TRUE. 2 arguments used.');
		// Case #5: Boolean, Boolean, Boolean. Three Boolean arguments, even TRUE count returns FALSE. 3 arguments used.
		oParser = new parserFormula('XOR(TRUE, FALSE, TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TRUE, FALSE, TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean, Boolean, Boolean. Three Boolean arguments, even TRUE count returns FALSE. 3 arguments used.');
		// Case #6: String. String "TRUE" coerces to Boolean TRUE. 1 argument used.
		oParser = new parserFormula('XOR("TRUE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR("TRUE") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. String "TRUE" coerces to Boolean TRUE. 1 argument used.');
		// Case #7: String, String. Numeric strings coerce to numbers (1=TRUE, 0=FALSE). 2 arguments used.
		oParser = new parserFormula('XOR("1", "0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR("1", "0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, String. Numeric strings coerce to numbers (1=TRUE, 0=FALSE). 2 arguments used.');
		// Case #8: Formula. Nested IF formula returns number (1=TRUE). 1 argument used.
		oParser = new parserFormula('XOR(IF(TRUE, 1, 0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(IF(TRUE, 1, 0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF formula returns number (1=TRUE). 1 argument used.');
		// Case #9: Formula, Formula. Nested IF formulas return numbers (1=TRUE, 0=FALSE). 2 arguments used.
		oParser = new parserFormula('XOR(IF(TRUE, 1, 0), IF(FALSE, 1, 0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(IF(TRUE, 1, 0), IF(FALSE, 1, 0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula, Formula. Nested IF formulas return numbers (1=TRUE, 0=FALSE). 2 arguments used.');
		// Case #10: Date. Date as serial number (>0, coerces to TRUE). 1 argument used.
		oParser = new parserFormula('XOR(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Date. Date as serial number (>0, coerces to TRUE). 1 argument used.');
		// Case #11: Time. Time as fractional number (>0, coerces to TRUE). 1 argument used.
		oParser = new parserFormula('XOR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Time as fractional number (>0, coerces to TRUE). 1 argument used.');
		// Case #12: Reference link. Reference to cell with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell with TRUE. 1 argument used.');
		// Case #13: Reference link, Reference link. References to cells with TRUE and FALSE, odd TRUE count. 2 arguments used.
		oParser = new parserFormula('XOR(A100, A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A100, A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link, Reference link. References to cells with TRUE and FALSE, odd TRUE count. 2 arguments used.');
		// Case #14: Area. Single-cell range with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(A100:A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A100:A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area. Single-cell range with TRUE. 1 argument used.');
		// Case #15: Array. Array with single TRUE element. 1 argument used.
		oParser = new parserFormula('XOR({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Array with single TRUE element. 1 argument used.');
		// Case #16: Array, Array. Arrays with single elements, odd TRUE count. 2 arguments used.
		oParser = new parserFormula('XOR({TRUE}, {FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR({TRUE}, {FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array, Array. Arrays with single elements, odd TRUE count. 2 arguments used.');
		// Case #17: Name. Named range with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name. Named range with TRUE. 1 argument used.');
		// Case #18: Name3D. 3D named range with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named range with TRUE. 1 argument used.');
		// Case #19: Ref3D. 3D reference to cell with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Ref3D. 3D reference to cell with TRUE. 1 argument used.');
		// Case #20: Area3D. 3D single-cell range with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D. 3D single-cell range with TRUE. 1 argument used.');
		// Case #21: Table. Table reference with TRUE. 1 argument used.
		oParser = new parserFormula('XOR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Table reference with TRUE. 1 argument used.');
		// Case #22: Formula. XOR with nested SUM formula (>0, coerces to TRUE). 1 argument used.
		oParser = new parserFormula('XOR(SUM(1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(SUM(1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. XOR with nested SUM formula (>0, coerces to TRUE). 1 argument used.');

		// Negative cases:
		// Case #1: String. Non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 1 argument used.');
		// Case #2: String, String. Non-numeric strings return #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR("abc", "def")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR("abc", "def") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Non-numeric strings return #VALUE!. 2 arguments used.');
		// Case #3: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('XOR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #4: Empty. Reference to empty cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!. 1 argument used.');
		// Case #5: Area. Multi-cell range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR(A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #6: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #7: Name. Named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #8: Name3D. 3D named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 1 argument used.');
		// Case #9: Ref3D. 3D reference to cell with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D reference to cell with text returns #VALUE!. 1 argument used.');
		// Case #10: Table. Table column with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Table1[Column2]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!. 1 argument used.');
		// Case #11: Number. Formula resulting in #DIV/0! error propagates error. 1 argument used.
		oParser = new parserFormula('XOR(0/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(0/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Number. Formula resulting in #DIV/0! error propagates error. 1 argument used.');
		// Case #12: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #13: Array. Array with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR({FALSE, "abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR({FALSE, "abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Array with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #14: Formula. Formula returning text causes #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR(IFERROR(1/0, "error"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(IFERROR(1/0, "error")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula returning text causes #VALUE!. 1 argument used.');
		// Case #15: Date, Date. Invalid string causes #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR(DATE(2025,1,1), "invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(DATE(2025,1,1), "invalid") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Date, Date. Invalid string causes #VALUE!. 2 arguments used.');
		// Case #16: Reference link, Reference link. Reference to valid and empty cells returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR(A100, A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(A100, A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Reference link, Reference link. Reference to valid and empty cells returns #VALUE!. 2 arguments used.');
		// Case #17: Name, Name. Named ranges with TRUE and text return #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR(TestName, TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(TestName, TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name, Name. Named ranges with TRUE and text return #VALUE!. 2 arguments used.');
		// Case #18: Ref3D, Ref3D. 3D references with TRUE and text return #VALUE!. 2 arguments used.
		oParser = new parserFormula('XOR(Sheet2!A1, Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(Sheet2!A1, Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D, Ref3D. 3D references with TRUE and text return #VALUE!. 2 arguments used.');
		// Case #19: Array. Multi-element array with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('XOR({TRUE, "abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR({TRUE, "abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array. Multi-element array with text returns #VALUE!. 1 argument used.');
		// Case #20: Formula. Formula resulting in #NUM! error propagates error. 1 argument used.
		oParser = new parserFormula('XOR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates error. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Zero (minimum numeric value) coerces to FALSE. 1 argument used.
		oParser = new parserFormula('XOR(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: Number. Zero (minimum numeric value) coerces to FALSE. 1 argument used.');
		// Case #2: Number. Maximum Excel number coerces to TRUE. 1 argument used.
		oParser = new parserFormula('XOR(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Maximum Excel number coerces to TRUE. 1 argument used.');
		// Case #3: Number. Smallest positive number coerces to TRUE. 1 argument used.
		oParser = new parserFormula('XOR(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number. Smallest positive number coerces to TRUE. 1 argument used.');
		// Case #4: Number, Number. Max number (TRUE) and zero (FALSE) return TRUE. 2 arguments used.
		oParser = new parserFormula('XOR(9.99999999999999E+307, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: XOR(9.99999999999999E+307, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: Number, Number. Max number (TRUE) and zero (FALSE) return TRUE. 2 arguments used.');

		// Need to fix: diff results from MS, string handle
		// Case #6: String. String "TRUE" coerces to Boolean TRUE. 1 argument used.
		// Case #10: Table. Table column with text returns #VALUE!. 1 argument used.
		// Case #15: Date, Date. Invalid string causes #VALUE!. 2 arguments used.

		testArrayFormula2(assert, "XOR", 1, 8, null, true);
	});

	wb.dependencyFormulas.unlockRecal();
});
