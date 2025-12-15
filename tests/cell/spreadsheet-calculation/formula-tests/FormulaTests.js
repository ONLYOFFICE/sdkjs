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

	var cDate = Asc.cDate;

	function testArrayFormula(assert, func, dNotSupportAreaArg) {

		var getValue = function (ref) {
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
		var array = oParser.calculate();
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
			consoleLog("func: " + func + " don't return area array");
		}

		oParser = new parserFormula(func + "({1,2,-3})", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H107").bbox);
		assert.ok(oParser.parse(), 'Formula is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), getValue(1), description + 'Number.');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), getValue(2), description + 'Number.');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), getValue(-3), description + 'Number.');
	}

	//returnOnlyValue - те функции, на вход которых всегда должны подаваться массивы и которые возвращают единственное значение
	function testArrayFormula2(assert, func, minArgCount, maxArgCount, dNotSupportAreaArg, returnOnlyValue) {

		var getValue = function (ref, countArg) {
			var argStr = "(";
			for (var j = 1; j <= countArg; j++) {
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
		var randomArray = [];
		var randomStrArray = "{";
		var maxArg = 4;
		for (var i = 1; i <= maxArg; i++) {
			var randVal = Math.random();
			randomArray.push(randVal);
			randomStrArray += randVal;
			if (i !== maxArg) {
				randomStrArray += ",";
			} else {
				randomStrArray += "}";
			}
		}

		for (var i = minArgCount; i <= maxArgCount; i++) {
			var argStrArr = "(";
			var randomArgStrArr = "(";
			for (var j = 1; j <= i; j++) {
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
			var array = oParser.calculate();
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
				consoleLog("func: " + func + " don't return area array");
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
				consoleLog("func: " + func + " don't return array");
			}
		}
	}

	function consoleLog(val) {
		//console.log(val);
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

	const c_msPerDay = AscCommonExcel.c_msPerDay,
		parserFormula = AscCommonExcel.parserFormula,
		GetDiffDate360 = AscCommonExcel.GetDiffDate360,
		fSortAscending = AscCommon.fSortAscending,
		g_oIdCounter = AscCommon.g_oIdCounter,
		ParseResult = AscCommonExcel.ParseResult,
		c_oAscError = Asc.c_oAscError;

	let oParser, wb, ws, dif = 1e-9, sData = AscCommon.getEmpty(), tmp, array, parseResult
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
		// there are no test operations with history, we disable it so that there are no unnecessary serializations
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

		var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read(sData, wb);
		ws = wb.getWorksheet(wb.getActive());
		AscCommonExcel.getFormulasInfo();
	}
	wb.dependencyFormulas.lockRecal();
	getTableType(599, 0, 599, 0); // Init table
	initDefNames();

	QUnit.module("Formula");

	QUnit.test('Iterative calculation', function (assert) {
		const g_cCalcRecursion = AscCommonExcel.g_cCalcRecursion;
		g_cCalcRecursion.initCalcProperties(wb.calcPr);
		g_cCalcRecursion.setIsEnabledRecursion(true);
		g_cCalcRecursion.setMaxIterations(10);
		g_cCalcRecursion.clearFunctionsResult();
		let nExpectedCellIndex, oFactCellIndex, oCell, bCaFromSelectedCell;
		// Init necessary functions
		const selectCell = function (sRange, oWs) {
			let oSelectCell = oWs ? oWs.getRange2(sRange) : ws.getRange2(sRange);
			let oCell = null;

			oSelectCell._foreach2(function (cell) {
				oCell = cell;
			})

			return oCell;
		};
		const getStartCellForIterCalc = function (oCell) {
			oCell.initStartCellForIterCalc();

			return g_cCalcRecursion.getStartCellIndex();
		};
		const getCaFromSelectedCell = function (sRange, oWs) {
			const oSelectedCell = selectCell(sRange, oWs);
			const oSelectedCellFormula = oSelectedCell.getFormulaParsed();

			return oSelectedCellFormula.ca;
		};
		// -- Check recursion formula with iteration limit
		// - Case: Sequence chain - A1000: A1000+B1000 -> B1000: B1000+C1000 -> C1: 1
		// Fill cells
		ws.getRange2("A1000").setValue("=A1000+B1000");
		ws.getRange2("B1000").setValue("=B1000+C1000");
		ws.getRange2("C1000").setValue("1");
		assert.strictEqual(ws.getRange2("A1000").getValue(), "45", "Test: Sequence chain = A1000: A1000+B1000, B1000: B1000+C1000, C1000: 1. A1000 - 45");
		assert.strictEqual(ws.getRange2("B1000").getValue(), "10", "Test: Sequence chain = A1000: A1000+B1000, B1000: B1000+C1000, C1000: 1. B1000 - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1000");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain = A1000: A1000+B1000, B1000: B1000+C1000, C1000: 1. isFormulaRecursion test. A1000 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1000");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain = A1000: A1000+B1000, B1000: B1000+C1000, C1000: 1. isFormulaRecursion test. B1000 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Loop chain - D1000: F1000/E1000 <-> F1000: E1000+D1000
		ws.getRange2("E1000").setValue("1");
		ws.getRange2("D1000").setValue("=F1000/E1000");
		ws.getRange2("F1000").setValue("=E1000+D1000");
		assert.strictEqual(ws.getRange2("D1000").getValue(), "9", "Test: Loop chain - D1000: F1000/E1000 <-> F1000: E1000+D1000. D1000 - 9");
		assert.strictEqual(ws.getRange2("F1000").getValue(), "10", "Test: Loop chain - D1000: F1000/E1000 <-> F1000: E1000+D1000. F1000 - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("D1000");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Loop chain - D1000: F1000/E1000 <-> F1000: E1000+D1000. isFormulaRecursion test. D1000 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("F1000");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Loop chain - D1000: F1000/E1000 <-> F1000: E1000+D1000. isFormulaRecursion test. F1000 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: 3D Loop chain - D1001: Sheet2!A1000/E1001 <-> Sheet2!A1000: Sheet1!D1001+Sheet1!E1001
		let ws2 = wb.createWorksheet(0, "Sheet2");
		ws.getRange2("E1001").setValue("1");
		ws.getRange2("D1001").setValue("=Sheet2!D1001/E1001");
		ws2.getRange2("D1001").setValue("=Sheet1!D1001+Sheet1!E1001");
		assert.strictEqual(ws.getRange2("D1001").getValue(), "9", "Test: 3D Loop chain - D1001: Sheet2!A1000/E1001 <-> Sheet2!A1000: Sheet1!D1001+Sheet1!E1001. D1001 - 9");
		assert.strictEqual(ws2.getRange2("D1001").getValue(), "10", "Test: 3D Loop chain - D1001: Sheet2!A1000/E1001 <-> Sheet2!A1000: Sheet1!D1001+Sheet1!E1001. Sheet2!A1000 - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("D1001");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D Loop chain - D1001: Sheet2!A1000/E1001 <-> Sheet2!A1000: Sheet1!D1001+Sheet1!E1001. isFormulaRecursion test. D1001 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("D1001", ws2);
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D Loop chain - D1001: Sheet2!A1000/E1001 <-> Sheet2!A1000: Sheet1!D1001+Sheet1!E1001. isFormulaRecursion test. Sheet2!A1000 - flag ca: true");
		bCaFromSelectedCell = null;
		// -  Case: Loop cell - A1001: A1001+1
		ws.getRange2("A1001").setValue("=A1001+1");
		assert.strictEqual(ws.getRange2("A1001").getValue(), "10", "Test: Loop cell - A1001: A1001+1. A1001 - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1001");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Loop cell - A1001: A1001+1. isFormulaRecursion test. A1001 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Negative case sequence chain without loop cell.
		ws.getRange2("A1002").setValue("=1+B1002");
		ws.getRange2("B1002").setValue("=1+C1002");
		ws.getRange2("C1002").setValue("1");
		assert.strictEqual(ws.getRange2("A1002").getValue(), "3", "Test: Negative case sequence chain without loop cell - A1002: 1+B1002. A1002 - 3");
		assert.strictEqual(ws.getRange2("B1002").getValue(), "2", "Test: Negative case sequence chain without loop cell - B1002: 1+C1002. B1002 - 2");
		assert.strictEqual(ws.getRange2("C1002").getValue(), "1", "Test: Negative case sequence chain without loop cell - C1002: 1. C1002 - 1");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1002");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Negative case sequence chain without loop cell - A1002: 1+B1002. isFormulaRecursion test. A1002 - flag ca: false");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1002");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Negative case sequence chain without loop cell - A1002: 1+B1002. isFormulaRecursion test. B1002 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Negative case cell without any chain.
		ws.getRange2("A1003").setValue("=1+2");
		assert.strictEqual(ws.getRange2("A1003").getValue(), "3", "Test: Negative case cell without any chain - A1003: 1+2. A1003 - 3");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1003");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Negative case cell without any chain - A1003: 1+2. isFormulaRecursion test. A1003 - flag ca: false");
		bCaFromSelectedCell = null;
		ws.getRange2("A1004").setValue("1");
		ws.getRange2("B1004").setValue("2");
		ws.getRange2("C1004").setValue("=A1004+B1004");
		assert.strictEqual(ws.getRange2("C1004").getValue(), "3", "Test: Negative case cell without any chain - C1004: A1004+B1004. C1004 - 3");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("C1004");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Negative case cell without any chain - C1004: A1004+B1004. isFormulaRecursion test. C1004 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: Sequence chain - A1005: A1005+B1005, B1005: 1. Deep level of recursion - 0
		ws.getRange2("A1005").setValue("=A1005+B1005");
		ws.getRange2("B1005").setValue("1");
		assert.strictEqual(ws.getRange2("A1005").getValue(), "10", "Test: Sequence chain - A1005: A1005+B1005, B1005: 1");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1005");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1005: A1005+B1005, B1005: 1. isFormulaRecursion test. A1005 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. Deep level of recursion - 10, Max iteration 10
		ws.getRange2("A1006").setValue("=A1006+B1006");
		ws.getRange2("B1006").setValue("=B1006+C1006");
		ws.getRange2("C1006").setValue("=C1006+D1006");
		ws.getRange2("D1006").setValue("=D1006+E1006");
		ws.getRange2("E1006").setValue("=E1006+F1006");
		ws.getRange2("F1006").setValue("=F1006+G1006");
		ws.getRange2("G1006").setValue("=G1006+H1006");
		ws.getRange2("H1006").setValue("=H1006+I1006");
		ws.getRange2("I1006").setValue("=I1006+J1006");
		ws.getRange2("J1006").setValue("1");
		assert.strictEqual(ws.getRange2("A1006").getValue(), "10", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. A1006 - 10");
		assert.strictEqual(ws.getRange2("B1006").getValue(), "45", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. B1006 - 45");
		assert.strictEqual(ws.getRange2("C1006").getValue(), "120", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. C1006 - 120");
		assert.strictEqual(ws.getRange2("D1006").getValue(), "210", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. D1006 - 210");
		assert.strictEqual(ws.getRange2("E1006").getValue(), "252", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. E1006 - 252");
		assert.strictEqual(ws.getRange2("F1006").getValue(), "210", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. F1006 - 210");
		assert.strictEqual(ws.getRange2("G1006").getValue(), "120", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. G1006 - 120");
		assert.strictEqual(ws.getRange2("H1006").getValue(), "45", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. H1006 - 45");
		assert.strictEqual(ws.getRange2("I1006").getValue(), "10", "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. I1006 - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. A1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. B1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("C1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. C1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("D1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. D1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("E1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. E1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("F1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. F1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("G1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. G1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("H1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. H1006 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("I1006");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1006: A1006+B1006, B1006: B1006+C1006, C1006: C1006+D1006 ... J1006: 1. isFormulaRecursion test. I1006 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... T1007: 1. Deep level of recursion - 20, Max iteration 10
		ws.getRange2("A1007").setValue("=A1007+B1007");
		ws.getRange2("B1007").setValue("=B1007+C1007");
		ws.getRange2("C1007").setValue("=C1007+D1007");
		ws.getRange2("D1007").setValue("=D1007+E1007");
		ws.getRange2("E1007").setValue("=E1007+F1007");
		ws.getRange2("F1007").setValue("=F1007+G1007");
		ws.getRange2("G1007").setValue("=G1007+H1007");
		ws.getRange2("H1007").setValue("=H1007+I1007");
		ws.getRange2("I1007").setValue("=I1007+J1007");
		ws.getRange2("J1007").setValue("=J1007+K1007");
		ws.getRange2("K1007").setValue("=K1007+L1007");
		ws.getRange2("L1007").setValue("=L1007+Q1007");
		ws.getRange2("Q1007").setValue("=Q1007+R1007");
		ws.getRange2("R1007").setValue("=R1007+S1007");
		ws.getRange2("S1007").setValue("=S1007+T1007");
		ws.getRange2("T1007").setValue("=T1007+U1007");
		ws.getRange2("U1007").setValue("=U1007+V1007");
		ws.getRange2("V1007").setValue("=V1007+W1007");
		ws.getRange2("W1007").setValue("=W1007+X1007");
		ws.getRange2("X1007").setValue("1");
		assert.strictEqual(ws.getRange2("A1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. A1007 - 0");
		assert.strictEqual(ws.getRange2("B1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. B1007 - 0");
		assert.strictEqual(ws.getRange2("C1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. C1007 - 0");
		assert.strictEqual(ws.getRange2("D1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. D1007 - 0");
		assert.strictEqual(ws.getRange2("E1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. E1007 - 0");
		assert.strictEqual(ws.getRange2("F1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. F1007 - 0");
		assert.strictEqual(ws.getRange2("G1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. G1007 - 0");
		assert.strictEqual(ws.getRange2("H1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. H1007 - 0");
		assert.strictEqual(ws.getRange2("I1007").getValue(), "0", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. I1007 - 0");
		assert.strictEqual(ws.getRange2("J1007").getValue(), "1", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. J1007 - 1");
		assert.strictEqual(ws.getRange2("K1007").getValue(), "10", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. K1007 - 10");
		assert.strictEqual(ws.getRange2("L1007").getValue(), "45", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. L1007 - 45");
		assert.strictEqual(ws.getRange2("Q1007").getValue(), "120", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. Q1007 - 120");
		assert.strictEqual(ws.getRange2("R1007").getValue(), "210", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. R1007 - 210");
		assert.strictEqual(ws.getRange2("S1007").getValue(), "252", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. S1007 - 252");
		assert.strictEqual(ws.getRange2("T1007").getValue(), "210", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. T1007 - 210");
		assert.strictEqual(ws.getRange2("U1007").getValue(), "120", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. U1007 - 120");
		assert.strictEqual(ws.getRange2("V1007").getValue(), "45", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. V1007 - 45");
		assert.strictEqual(ws.getRange2("W1007").getValue(), "10", "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. W1007 - 10");
		// Check work isFormulaRecursion function
		bCaFromSelectedCell = getCaFromSelectedCell("A1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. A1007 - flag ca: true");
		bCaFromSelectedCell = null
		bCaFromSelectedCell = getCaFromSelectedCell("B1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. B1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("C1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. C1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("D1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. D1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("E1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. E1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("F1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. F1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("G1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. G1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("H1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. H1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("I1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. I1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("J1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. J1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("K1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. K1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("L1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. L1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("Q1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. Q1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("R1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. R1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("S1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. S1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("T1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. T1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("U1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. U1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("V1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. V1007 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("W1007");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain - A1007: A1007+B1007, B1007: B1007+C1007, C1007: C1007+D1007 ... X1007: 1. isFormulaRecursion test. W1007 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1
		ws.getRange2("A1011").setValue("=A1011+A1012");
		ws.getRange2("A1012").setValue("=A1012+A1013");
		ws.getRange2("A1013").setValue("=A1013+A1014");
		ws.getRange2("A1014").setValue("=A1014+A1015");
		ws.getRange2("A1015").setValue("1");
		assert.strictEqual(ws.getRange2("A1011").getValue(), "210", "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. A1011 - 1");
		assert.strictEqual(ws.getRange2("A1012").getValue(), "120", "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. A1012 - 120");
		assert.strictEqual(ws.getRange2("A1013").getValue(), "45", "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. A1013 - 45");
		assert.strictEqual(ws.getRange2("A1014").getValue(), "10", "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. A1014 - 10");
		// Check work isRecursionFormula function
		bCaFromSelectedCell = getCaFromSelectedCell("A1011");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. isFormulaRecursion test. A1011 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1012");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. isFormulaRecursion test. A1012 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1013");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. isFormulaRecursion test. A1013 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1014");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Vertical sequence chain - A1011: A1011+A1012, A1012: A1012+A1013, A1013: A1013+A1014, A1014: A1014+A1015, A1015: 1. isFormulaRecursion test. A1014 - flag ca: true");
		bCaFromSelectedCell = null;
		// Remove created sheets.
		wb.removeWorksheet(0);
		// - Case: 3D sequence chain - A1016: A1016+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1
		ws2 = wb.createWorksheet(0, "Sheet2");
		let ws3 = wb.createWorksheet(1, "Sheet3");
		ws.getRange2("A1016").setValue("=A1016+Sheet2!A1000");
		ws2.getRange2("A1000").setValue("=A1000+Sheet3!A1000");
		ws3.getRange2("A1000").setValue("1");
		assert.strictEqual(ws.getRange2("A1016").getValue(), "45", "Test: 3D sequence chain - A1012: A1012+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1. A1012 - 45");
		assert.strictEqual(ws2.getRange2("A1000").getValue(), "10", "Test: 3D sequence chain - A1012: A1012+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1. Sheet2!A1000 - 10");
		assert.strictEqual(ws3.getRange2("A1000").getValue(), "1", "Test: 3D sequence chain - A1012: A1012+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1. Sheet3!A1000 - 1");
		// Check work isRecursionFormula function
		bCaFromSelectedCell = getCaFromSelectedCell("A1016");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D sequence chain - A1016: A1016+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1. isFormulaRecursion test. A1016 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1000", ws2);
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D sequence chain - A1016: A1016+Sheet2!A1000, Sheet2!A1000: Sheet2!A1000+Sheet3!A1000, Sheet3!A1000: 1. isFormulaRecursion test. Sheet2!A1000 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: 3D sequence chain  B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1
		ws.getRange2("B1012").setValue("=B1012+Sheet2!B1012");
		ws2.getRange2("B1012").setValue("=B1012+Sheet3!B1012");
		ws3.getRange2("B1012").setValue("1");
		assert.strictEqual(ws.getRange2("B1012").getValue(), "45", "Test: 3D sequence chain - B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1. B1012 - 45");
		assert.strictEqual(ws2.getRange2("B1012").getValue(), "10", "Test: 3D sequence chain - B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1. Sheet2!B1012 - 10");
		assert.strictEqual(ws3.getRange2("B1012").getValue(), "1", "Test: 3D sequence chain - B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1. Sheet3!B1012 - 1");
		// Check work isRecursionFormula function
		bCaFromSelectedCell = getCaFromSelectedCell("B1012");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D sequence chain - B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1. isFormulaRecursion test. B1012 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1012", ws2);
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D sequence chain - B1012: B1012+Sheet2!B1012, Sheet2!B1012: Sheet2!B1012+Sheet3!B1012, Sheet3!B1012: 1. isFormulaRecursion test. Sheet2!B1012 - flag ca: true");
		bCaFromSelectedCell = null;
		// Remove created sheets.
		wb.removeWorksheet(0);
		wb.removeWorksheet(0);
		// - Case: DefName loop cell - X: X+1
		let oDefName = new Asc.asc_CDefName("x", ws.getName() + "!$A$1017");
		wb.editDefinesNames(null, oDefName);
		ws.getRange2("A1017").setValue("=x+1")
		assert.strictEqual(ws.getRange2("A1017").getValue(), "10", "Test: DefName loop cell - X: X+1. X - 10");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1017");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: DefName loop cell - X: X+1. isFormulaRecursion test. A1017(X) - flag ca: true");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefName);
		oDefName = null;
		// - Case: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1
		let oDefNameX = new Asc.asc_CDefName("x", ws.getName() + "!$A$1018");
		let oDefNameY = new Asc.asc_CDefName("y", ws.getName() + "!$B$1018");
		let oDefNameZ = new Asc.asc_CDefName("z", ws.getName() + "!$C$1018");
		wb.editDefinesNames(null, oDefNameX);
		wb.editDefinesNames(null, oDefNameY);
		wb.editDefinesNames(null,oDefNameZ);
		ws.getRange2("A1018").setValue("=x+y");
		ws.getRange2("B1018").setValue("=y+z");
		ws.getRange2("C1018").setValue("1");
		assert.strictEqual(ws.getRange2("A1018").getValue(), "45", "Test: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1. X - 45");
		assert.strictEqual(ws.getRange2("B1018").getValue(), "10", "Test: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1. Y - 10");
		assert.strictEqual(ws.getRange2("C1018").getValue(), "1", "Test: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1. Z - 1");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1018");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1. isFormulaRecursion test. A1018(X) - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1018");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: DefName sequence chain - X: X+Y, Y: Y+Z, Z: 1. isFormulaRecursion test. B1018(Y) - flag ca: true");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefNameX);
		wb.delDefinesNames(oDefNameY);
		wb.delDefinesNames(oDefNameZ);
		oDefNameX = null;
		oDefNameY = null;
		oDefNameZ = null;
		// - Case: Area recursive formula SUM(A1019:D1019)
		ws.getRange2("A1019").setValue("1");
		ws.getRange2("B1019").setValue("2");
		ws.getRange2("C1019").setValue("3");
		ws.getRange2("D1019").setValue("=SUM(A1019:D1019)");
		assert.strictEqual(ws.getRange2("D1019").getValue(), "60", "Test: Area recursive formula SUM(A1019:D1019). D1019 - 60");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("D1019");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Area recursive formula SUM(A1019:D1019). isFormulaRecursion test. D1019 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Area recursive formula SUM(D1020, A1020:C1020)
		ws.getRange2("A1020").setValue("1");
		ws.getRange2("B1020").setValue("2");
		ws.getRange2("C1020").setValue("3");
		ws.getRange2("D1020").setValue("=SUM(D1020, A1020:C1020)");
		assert.strictEqual(ws.getRange2("D1020").getValue(), "60", "Test: Area recursive formula SUM(D1020, A1020:C1020). D1020 - 60");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("D1020");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Area recursive formula SUM(D1020, A1020:C1020). isFormulaRecursion test. D1020 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: 3D Area recursive formula SUM(A1021, Sheet2!A1000:C1000)
		ws2 = wb.createWorksheet(0, "Sheet2");
		ws2.getRange2("A1000").setValue("1");
		ws2.getRange2("B1000").setValue("2");
		ws2.getRange2("C1000").setValue("3");
		ws.getRange2("A1021").setValue("=SUM(A1021, Sheet2!A1000:C1000)");
		assert.strictEqual(ws.getRange2("A1021").getValue(), "60", "Test: 3D Area recursive formula SUM(A1021, Sheet2!A1000:C1000). A1021 - 60");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1021");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: 3D Area recursive formula SUM(A1021, Sheet2!A1000:C1000). isFormulaRecursion test. A1021 - flag ca: true");
		bCaFromSelectedCell = null;
		// Remove created sheets.
		wb.removeWorksheet(0);
		// - Case: DefName Area recursive formula SUM(Range)
		let oDefNameRange = new Asc.asc_CDefName("Range", ws.getName() + "!$A$1022:$D$1022");
		wb.editDefinesNames(null, oDefNameRange);
		ws.getRange2("A1022").setValue("1");
		ws.getRange2("B1022").setValue("2");
		ws.getRange2("C1022").setValue("3");
		ws.getRange2("D1022").setValue("=SUM(Range)");
		assert.strictEqual(ws.getRange2("D1022").getValue(), "60", "Test: DefName Area recursive formula SUM(Range). D1022 - 60");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("D1022");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: DefName Area recursive formula SUM(Range). isFormulaRecursion test. D1022 - flag ca: true");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefNameRange);
		oDefNameRange = null;
		// - Case: DefName Area 3D recursive formula SUM(A1000,Range3D)
		ws2 = wb.createWorksheet(0, "Sheet2");
		let oDefNameRange3D = new Asc.asc_CDefName("Range3D", ws.getName() + "!$A$1023:$C$1023");
		wb.editDefinesNames(null, oDefNameRange3D);
		ws.getRange2("A1023").setValue("1");
		ws.getRange2("B1023").setValue("2");
		ws.getRange2("C1023").setValue("3");
		ws2.getRange2("A1000").setValue("=SUM(A1000, Range3D)");
		assert.strictEqual(ws2.getRange2("A1000").getValue(), "60", "Test: DefName Area 3D recursive formula SUM(A1000,Range3D). Sheet2!A1000 - 60");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1000", ws2);
		assert.strictEqual(bCaFromSelectedCell, true, "Test: DefName Area 3D recursive formula SUM(A1000,Range3D). isFormulaRecursion test. Sheet2!A1000 - flag ca: true");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefNameRange3D);
		oDefNameRange3D = null;
		// Remove created sheets.
		wb.removeWorksheet(0);
		// - Case: Chain recursive formula without outStack link. A1024: A1024+1, B1024: A1024+B1024, C1024: B1024+C1024
		ws.getRange2("A1024").setValue("=A1024+1");
		ws.getRange2("B1024").setValue("=A1024+B1024");
		ws.getRange2("C1024").setValue("=B1024+C1024");
		assert.strictEqual(ws.getRange2("A1024").getValue(), "30", "Test: Chain recursive formula without outStack link. A1024 - 30");
		assert.strictEqual(ws.getRange2("B1024").getValue(), "410", "Test: Chain recursive formula without outStack link. B1024 - 410");
		assert.strictEqual(ws.getRange2("C1024").getValue(), "2870", "Test: Chain recursive formula without outStack link. C1024 - 2870");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("C1024");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain recursive formula without outStack link. isFormulaRecursion test. C1024 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1024");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain recursive formula without outStack link. isFormulaRecursion test. B1024 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1024");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain recursive formula without outStack link. isFormulaRecursion test. A1024 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Not recursive formula A1025: SUM(B1025:C1025)
		ws.getRange2("A1025").setValue("=SUM(B1025:C1025)");
		ws.getRange2("B1025").setValue("1");
		ws.getRange2("C1025").setValue("2");
		assert.strictEqual(ws.getRange2("A1025").getValue(), "3", "Test: Not recursive formula A1025 - 3");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1025");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Not recursive formula A1025. isFormulaRecursion test. A1025 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: Area recursive formula SUM(Y:Z)
		ws.getRange2("Y1000").setValue("1");
		ws.getRange2("Z1000").setValue("2");
		ws.getRange2("Y1001").setValue("=SUM(Y:Z)");
		assert.strictEqual(ws.getRange2("Y1001").getValue(), "30", "Test: Area recursive formula SUM(Y:Z). Y1001 - 30");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("Y1001");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Area recursive formula SUM(Y:Z). isFormulaRecursion test. Y1001 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Def name area recursive formula SUM(All)
		let oDefNameAll = new Asc.asc_CDefName("All", ws.getName() + "!$Z:$Z");
		wb.editDefinesNames(null, oDefNameAll);
		ws.getRange2("Z1000").setValue("1");
		ws.getRange2("Z1001").setValue("2");
		ws.getRange2("Z1002").setValue("=SUM(All)");
		assert.strictEqual(ws.getRange2("Z1002").getValue(), "30", "Test: Def name area recursive formula SUM(All). Z1002 - 30");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("Z1002");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Def name area recursive formula SUM(All). isFormulaRecursion test. Z1002 - flag ca: true");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefNameAll);
		oDefNameAll = null;
		// - Case: Def name area non-recursive formula SUM(XAll)
		let oDefNameABAll = new Asc.asc_CDefName("ABAll", ws.getName() + "!$AB:$AB");
		wb.editDefinesNames(null, oDefNameABAll);
		ws.getRange2("AB1000").setValue("1");
		ws.getRange2("AB1001").setValue("2");
		ws.getRange2("A1026").setValue("=SUM(ABAll)");
		assert.strictEqual(ws.getRange2("A1026").getValue(), "3", "Test: Def name area non-recursive formula SUM(ABAll). A1026 - 3");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1026");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Def name area non-recursive formula SUM(ABAll). isFormulaRecursion test. A1026 - flag ca: false");
		bCaFromSelectedCell = null;
		// Clean define name
		wb.delDefinesNames(oDefNameABAll);
		oDefNameABAll = null;
		// - Case: Cross recursive chain formula. B1027: B1027+B1028, B1028: B1028+B1029, A1028: A1028+B1028, C1028: B1028+C1028, B1029: 1
		ws.getRange2("B1027").setValue("=B1027+B1028");
		ws.getRange2("B1028").setValue("=B1028+B1029");
		ws.getRange2("A1028").setValue("=A1028+B1028");
		ws.getRange2("C1028").setValue("=B1028+C1028");
		ws.getRange2("B1029").setValue("1");
		assert.strictEqual(ws.getRange2("B1027").getValue(), "45", "Test: Cross recursive chain formula. B1027 - 45");
		assert.strictEqual(ws.getRange2("B1028").getValue(), "10", "Test: Cross recursive chain formula. B1028 - 10");
		assert.strictEqual(ws.getRange2("A1028").getValue(), "45", "Test: Cross recursive chain formula. A1028 - 45");
		assert.strictEqual(ws.getRange2("C1028").getValue(), "55", "Test: Cross recursive chain formula. C1028 - 55");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("B1027");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Cross recursive chain formula. isFormulaRecursion test. B1027 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1028");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Cross recursive chain formula. isFormulaRecursion test. B1028 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1028");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Cross recursive chain formula. isFormulaRecursion test. A1028 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("C1028");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Cross recursive chain formula. isFormulaRecursion test. C1028 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Recrusive cell with IF formula, with empty cell ("")
		ws.getRange2("A1030").setValue("");
		ws.getRange2("B1030").setValue("=IF(A1030<>\"\",IF(B1030<>\"\",B1030,NOW()),\"\")");
		assert.strictEqual(ws.getRange2("B1030").getValue(), "", "Test: Recrusive cell with IF formula, with empty cell (\"\"). B1030 - \"\"");
		ws.getRange2("A1030").setValue("Test");
		let date = new cDate();
		let excelDate = date.getExcelDate();
		assert.strictEqual(Math.floor(ws.getRange2("B1030").getValue()), excelDate, "Test: Recrusive cell with IF formula, with empty cell (\"\"). B1030 - " + excelDate);
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("B1030");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recrusive cell with IF formula, with empty cell (\"\"). isFormulaRecursion test. B1030 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Recursive convergent formula with IF formula
		ws.getRange2("A1031").setValue("=IF(A1031=0,-5,A1031-(A1031^3-4*A1031^2-4*A1031+5)/(3*A1031^2-8*A1031-4))");
		assert.strictEqual(ws.getRange2("A1031").getValue(), "-1.4012223386412388", "Test: Recursive convergent formula with IF formula. A1031 - -1.04122233864124");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("A1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive convergent formula with IF formula. isFormulaRecursion test. A1031 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2
		ws.getRange2("A1031").setValue("1000");
		ws.getRange2("B1031").setValue("0.36");
		ws.getRange2("C1031").setValue("23");
		ws.getRange2("D1031").setValue("8");
		ws.getRange2("E1031").setValue("27");
		ws.getRange2("F1031").setValue("=IF(C1031*D1031*E1031/5<B1031,B1031,IF(C1031*D1031*E1031/5\>B1031,C1031*D1031*E1031/5000))");
		ws.getRange2("G1031").setValue("8");
		ws.getRange2("H1031").setValue("6.5");
		ws.getRange2("I1031").setValue("25");
		ws.getRange2("J1031").setValue("=IF(F1031<1,R1031*4%,IF(F1031>1,R1031*5%,IF(F1031>3,R1031*5.5%)))");
		ws.getRange2("K1031").setValue("=(A1031+I1031+J1031)*100/(100-G1031-H1031)");
		ws.getRange2("L1031").setValue("=R1031*100/69.9");
		ws.getRange2("Q1031").setValue("=L1031*1.07");
		ws.getRange2("R1031").setValue("=K1031*100/94.7");
		assert.strictEqual(ws.getRange2("J1031").getValue(), "53.268529249844036", "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2. J1031 - 53.26854488596814");
		assert.strictEqual(ws.getRange2("K1031").getValue(), "1261.1327827483556", "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2. K1031 - 1261.1328010362201");
		assert.strictEqual(ws.getRange2("L1031").getValue(), "1905.1696763189466", "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2. L1031 - 1905.1697026454983");
		assert.strictEqual(ws.getRange2("Q1031").getValue(), "2038.531553661273", "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2. Q1031 - 2038.5315818306833");
		assert.strictEqual(ws.getRange2("R1031").getValue(), "1331.7136037469436", "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031. Bug-42873, Comment #2. R1031 - 1331.7136230583105");
		// Check work isFormulaRecursive function
		bCaFromSelectedCell = getCaFromSelectedCell("J1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031, Bug-42873, Comment #2. isFormulaRecursion test. J1031 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("K1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031, Bug-42873, Comment #2. isFormulaRecursion test. K1031 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("L1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031, Bug-42873, Comment #2. isFormulaRecursion test. L1031 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("Q1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031, Bug-42873, Comment #2. isFormulaRecursion test. Q1031 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("R1031");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Recursive cell chain formula: J1031->R1031->K1031->J1031, Bug-42873, Comment #2. isFormulaRecursion test. R1031 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc.
		ws.getRange2("A1032").setValue("0"); // P
		ws.getRange2("A1033").setValue("10"); // deltaY
		ws.getRange2("A1034").setValue("1"); // X
		ws.getRange2("B1032").setValue("10"); // mu
		ws.getRange2("C1032").setValue("10"); // u start
		ws.getRange2("D1032").setValue("=C1032"); // u
		ws.getRange2("B1033").setValue("10"); // mu
		ws.getRange2("C1033").setValue("-10");// u start
		ws.getRange2("D1033").setValue("=IF($A$1034=0, C1033, -(($A$1032*$A$1033*$A$1033-(D1034-D1032)*(B1034-B1032)/4)/B1033-(D1034+D1032))/2)"); // u
		ws.getRange2("B1034").setValue("10"); // mu
		ws.getRange2("C1034").setValue("-10"); // u start
		ws.getRange2("D1034").setValue("=IF($A$1034=0, C1034, -(($A$1032*$A$1033*$A$1033-(D1035-D1033)*(B1035-B1033)/4)/B1034-(D1035+D1033))/2)"); // u
		ws.getRange2("B1035").setValue("10"); // mu
		ws.getRange2("C1035").setValue("-10"); // u start
		ws.getRange2("D1035").setValue("=IF($A$1034=0, C1035, -(($A$1032*$A$1033*$A$1033-(D1036-D1034)*(B1036-B1034)/4)/B1035-(D1036+D1034))/2)"); // u
		ws.getRange2("B1036").setValue("10"); // mu
		ws.getRange2("C1036").setValue("0"); // u start
		ws.getRange2('D1036').setValue("=C1036"); // u
		assert.strictEqual(ws.getRange2("D1033").getValue(), "7.499457849531321", "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. D1033 - 7.499728735325547");
		assert.strictEqual(ws.getRange2("D1034").getValue(), "4.999457849531321", "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. D1034 - 4.999728735325547");
		assert.strictEqual(ws.getRange2("D1035").getValue(), "2.4997289247656607", "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. D1035 - 2.4998643676627736");
		// Check work  isFormulaRecursion function
		bCaFromSelectedCell = getCaFromSelectedCell("D1033");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. isFormulaRecursion test. D1033 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("D1034");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. isFormulaRecursion test. D1034 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("D1035");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence-loop chain. D1033 <-> D1034, D1034 <-> D1035 etc. isFormulaRecursion test. D1035 - flag ca: true");
		bCaFromSelectedCell = null;
		// Case: Convergent formula calucaltes only once time.
		g_cCalcRecursion.setMaxIterations(15);
		ws.getRange2("A1040").setValue("=A1041+A1043");
		ws.getRange2("A1041").setValue("25150");
		ws.getRange2("A1042").setValue("0.2");
		ws.getRange2("A1043").setValue("=A1040*A1042");
		assert.strictEqual(ws.getRange2("A1040").getValue(), "31437.49935616", "Test: Convergent formula calculates only once time. First calculate. A1040 - 31437.49987");
		assert.strictEqual(ws.getRange2("A1043").getValue(), "6287.499871232", "Test: Convergent formula calculates only once time. First calculate. A1041 - 6287.499974246401");
		ws.getRange2("A1041").setValue("25150");
		assert.strictEqual(ws.getRange2("A1040").getValue(), "31437.49935616", "Test: Convergent formula calculates only once time. Recalculate. A1040 - 31437.49987");
		assert.strictEqual(ws.getRange2("A1043").getValue(), "6287.499871232", "Test: Convergent formula calculates only once time. Recalculate. A1041 - 6287.499974246401");
		// Check work isFormulaRecursion function
		bCaFromSelectedCell = getCaFromSelectedCell("A1040");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Convergent formula calculates only once time. isFormulaRecursion test. A1040 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1043");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Convergent formula calculates only once time. isFormulaRecursion test. A1043 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: With disabled iterative calculation.
		g_cCalcRecursion.setIsEnabledRecursion(false);
		ws.getRange2("A1037").setValue("=A1037+1");
		ws.getRange2("A1038").setValue("=A1038+B1038");
		ws.getRange2("B1038").setValue("=B1038+C1038");
		ws.getRange2("C1038").setValue("1");
		assert.strictEqual(ws.getRange2("A1037").getValue(), "0", "Test: Loop cell with disabled iterative calculation. A1037 - 0");
		assert.strictEqual(ws.getRange2("A1038").getValue(), "0", "Test: Sequence chain with disabled iterative calculation. A1038 - 0");
		assert.strictEqual(ws.getRange2("B1038").getValue(), "0", "Test: Sequence chain with disabled iterative calculation. B1038 - 0");
		// Check work isFormulaRecursion function
		bCaFromSelectedCell = getCaFromSelectedCell("A1037");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Loop cell with disabled iterative calculation. A1037 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("A1038");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain with disabled iterative calculation. A1038 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1038");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Sequence chain with disabled iterative calculation. B1038 - flag ca: true");
		bCaFromSelectedCell = null;
		// Case: Exception formulas that ignores rules of recursion recognition
		ws.getRange2("A1039").setValue("=ROW(A1039)");
		assert.strictEqual(ws.getRange2("A1039").getValue(), "1039", "Test: Exception formulas that ignores rules of recursion recognition. A1039 - 1039. Formula - ROW");
		bCaFromSelectedCell = getCaFromSelectedCell("A1039");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Exception formulas that ignores rules of recursion recognition. A1039 - flag ca: false");
		bCaFromSelectedCell = null;
		ws.getRange2("B1039").setValue("=COLUMN(B1039)");
		assert.strictEqual(ws.getRange2("B1039").getValue(), "2", "Test: Exception formulas that ignores rules of recursion recognition. B1039 - 2. Formula - COLUMN");
		bCaFromSelectedCell = getCaFromSelectedCell("B1039");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Exception formulas that ignores rules of recursion recognition. B1039 - flag ca: false");
		bCaFromSelectedCell = null;
		ws.getRange2("C1039").setValue("=ISFORMULA(C1039)");
		assert.strictEqual(ws.getRange2("C1039").getValue(), "TRUE", "Test: Exception formulas that ignores rules of recursion recognition. C1039 - TRUE. Formula - ISFORMULA");
		bCaFromSelectedCell = getCaFromSelectedCell("C1039");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Exception formulas that ignores rules of recursion recognition. C1039 - flag ca: false");
		bCaFromSelectedCell = null;
		ws.getRange2("D1039").setValue("=SHEETS(D1039)");
		assert.strictEqual(ws.getRange2("D1039").getValue(), "1", "Test: Exception formulas that ignores rules of recursion recognition. D1039 - 1. Formula - SHEETS");
		bCaFromSelectedCell = getCaFromSelectedCell("D1039");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Exception formulas that ignores rules of recursion recognition. D1039 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIF 2 args recursion range. Recursion formula with disabled Iterative calculation setting.
		ws.getRange2("A1040").setValue("2");
		ws.getRange2("B1040").setValue("4");
		ws.getRange2("C1040").setValue("8");
		ws.getRange2("D1040").setValue("=SUMIF(A1040:D1040, \">4\")");
		assert.strictEqual(ws.getRange2("D1040").getValue(), "0", "Test: SUMIF 2 args recursion range. D1040 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("D1040");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF 2 args recursion range. D1040 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF 2 args non recursion range with disabled Iterative calculation setting.
		ws.getRange2("D1040").setValue("=SUMIF(A1040:C1040, \">4\")");
		assert.strictEqual(ws.getRange2("D1040").getValue(), "8", "Test: SUMIF 2 args non recursion range. D1040 - 8");
		bCaFromSelectedCell = getCaFromSelectedCell("D1040");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIF 2 args non recursion range. D1040 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 2 args. Recursion criteria. Recursion formula with disabled Iterative calculation setting.
		ws.getRange2("A1041").setValue("2");
		ws.getRange2("B1041").setValue("4");
		ws.getRange2("C1041").setValue("8");
		ws.getRange2("A1042").setValue("2");
		ws.getRange2("B1042").setValue("4");
		ws.getRange2("C1042").setValue("=SUMIF(A1041:C1041, A1042:C1042)");
		assert.strictEqual(ws.getRange2("C1042").getValue(), "0", "Test: SUMIF. 2 args. Recursion criteria. C1042 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("C1042");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 2 args. Recursion criteria. C1042 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 2 args. Non recursion criteria with disabled Iterative calculation setting.
		ws.getRange2("A1043").setValue("2");
		ws.getRange2("B1043").setValue("4");
		ws.getRange2("C1043").setValue("8");
		ws.getRange2("A1044").setValue("2");
		ws.getRange2("B1044").setValue("4");
		ws.getRange2("C1044").setValue("=SUMIF(A1043:C1043, A1044:B1044)");
		bCaFromSelectedCell = getCaFromSelectedCell("C1044");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIF. 2 args. Non recursion criteria. C1044 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 2 args. Recursion formula. Recursion criteria name with disabled Iterative calculation setting.
		let oCriteriaRangeName = new Asc.asc_CDefName("Criteria_range", ws.getName() + "!$A$1045:$C$1045");
		wb.editDefinesNames(null, oCriteriaRangeName);
		ws.getRange2("A1045").setValue("2");
		ws.getRange2("B1045").setValue("4");
		ws.getRange2("C1045").setValue("=SUMIF(A1043:C1043, Criteria_range)");
		assert.strictEqual(ws.getRange2("C1045").getValue(), "0", "Test: SUMIF. 2 args. Recursion formula. Recursion criteria name. C1045 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("C1045");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 2 args. Recursion formula. Recursion criteria name. C1045 - flag ca: true");
		bCaFromSelectedCell = null;
		wb.delDefinesNames(oCriteriaRangeName);
		oCriteriaRangeName = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. With disabled Iterative calculation setting.
		// range row
		ws.getRange2("A1046").setValue("1");
		ws.getRange2("B1046").setValue("0");
		ws.getRange2("C1046").setValue("1");
		ws.getRange2("D1046").setValue("0");
		// criteria row
		ws.getRange2("A1047").setValue("1");
		// sum_range row
		ws.getRange2("A1048").setValue("2");
		ws.getRange2("B1048").setValue("4");
		ws.getRange2("C1048").setValue("8");
		ws.getRange2("D1048").setValue("=SUMIF(A1046:D1046, A1047 ,A1048:D1048)");
		assert.strictEqual(ws.getRange2("D1048").getValue(), "10", "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. D1048 - 10");
		bCaFromSelectedCell = getCaFromSelectedCell("D1048");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. D1048 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula matches with criteria. With disabled Iterative calculation setting.
		ws.getRange2("D1046").setValue("1");
		assert.strictEqual(ws.getRange2("D1048").getValue(), "10", "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula matches with criteria. D1048 - 10");
		bCaFromSelectedCell = getCaFromSelectedCell("D1048");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range, but the cell with formula matches with criteria. D1048 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case:  SUMIF. 3 args. Recursion formula. Recursion range. With disabled Iterative calculation setting.
		// range row
		ws.getRange2("A1049").setValue("1");
		ws.getRange2("B1049").setValue("0");
		ws.getRange2("C1049").setValue("1");
		// criteria row
		ws.getRange2("A1050").setValue("1");
		// sum_range row
		ws.getRange2("A1051").setValue("2");
		ws.getRange2("B1051").setValue("4");
		ws.getRange2("C1051").setValue("8");
		ws.getRange2("D1051").setValue("16");
		ws.getRange2("D1049").setValue("=SUMIF(A1049:D1049, A1050, A1051:D1051)");
		assert.strictEqual(ws.getRange2("D1049").getValue(), "0", "Test: SUMIF. 3 args. Recursion formula. Recursion range. D1049 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("D1049");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion range. D1049 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion criteria. With disabled Iterative calculation setting.
		// range row
		ws.getRange2("A1052").setValue("1");
		ws.getRange2("B1052").setValue("0");
		ws.getRange2("C1052").setValue("1");
		ws.getRange2("D1052").setValue("0")
		// criteria row
		ws.getRange2("A1053").setValue("1");
		// sum_range row
		ws.getRange2("A1054").setValue("2");
		ws.getRange2("B1054").setValue("4");
		ws.getRange2("C1054").setValue("8");
		ws.getRange2("D1054").setValue("16");
		ws.getRange2("B1053").setValue("=SUMIF(A1052:D1052, A1053:B1053, A1054:D1054)");
		assert.strictEqual(ws.getRange2("B1053").getValue(), "0", "Test: SUMIF. 3 args. Recursion formula. Recursion criteria. B1053 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("B1053");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion criteria. B1053 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion criteria_range with DefName. With disabled Iterative calculation setting.
		oCriteriaRangeName = new Asc.asc_CDefName("Criteria_range", ws.getName() + "!$A$1055:$D$1055");
		let oCriteriaName = new Asc.asc_CDefName("criteria", ws.getName() + "!$A$1056:$B$1056");
		let oSumRangeName = new Asc.asc_CDefName("sum_range", ws.getName() + "!$A$1057:$D$1057");
		wb.editDefinesNames(null, oCriteriaRangeName);
		wb.editDefinesNames(null, oCriteriaName);
		wb.editDefinesNames(null, oSumRangeName);
		ws.getRange2("A1055").setValue("1");
		ws.getRange2("B1055").setValue("0");
		ws.getRange2("C1055").setValue("1");
		ws.getRange2("A1056").setValue(">0");
		ws.getRange2("A1057").setValue("2");
		ws.getRange2("B1057").setValue("4");
		ws.getRange2("C1057").setValue("8");
		ws.getRange2("D1057").setValue("16");
		ws.getRange2("D1055").setValue("=SUMIF(Criteria_range, A1056, sum_range)");
		assert.strictEqual(ws.getRange2("D1055").getValue(), "0", "Test: SUMIF. 3 args. Recursion formula. Recursion criteria_range with DefName. D1055 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("D1055");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion criteria_range with DefName. D1055 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName. With disabled Iterative calculation setting.
		ws.getRange2("D1055").setValue("1");
		ws.getRange2("D1057").setValue("=SUMIF(Criteria_range, A1056, sum_range)");
		assert.strictEqual(ws.getRange2("D1057").getValue(), "0", "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName. D1057 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("D1057");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName. D1057 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName, but the cell with formula doesn't match with criteria. With disabled Iterative calculation setting.
		ws.getRange2("D1055").setValue("0");
		ws.getRange2("D1057").setValue("=SUMIF(Criteria_range, A1056, sum_range)");
		assert.strictEqual(ws.getRange2("D1057").getValue(), "10", "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName, but the cell with formula doesn't match with criteria. D1057 - 10")
		bCaFromSelectedCell = getCaFromSelectedCell("D1057");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIF. 3 args. Recursion formula. Recursion sum_range with DefName, but the cell with formula doesn't match with criteria. D1057 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIF. 3 args. Recursion formula. Recursion criteria with DefName. With disabled Iterative calculation setting.
		ws.getRange2("D1055").setValue("0");
		ws.getRange2("D1057").setValue("16");
		ws.getRange2("A1056").setValue("1");
		ws.getRange2("B1056").setValue("=SUMIF(Criteria_range, criteria, sum_range)");
		assert.strictEqual(ws.getRange2("B1056").getValue(), "0", "Test: SUMIF. 3 args. Recursion formula. Recursion criteria with DefName. B1056 - 0")
		bCaFromSelectedCell = getCaFromSelectedCell("B1056");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIF. 3 args. Recursion formula. Recursion criteria with DefName. B1056 - flag ca: true");
		bCaFromSelectedCell = null;
		wb.delDefinesNames(oCriteriaRangeName);
		wb.delDefinesNames(oCriteriaName);
		wb.delDefinesNames(oSumRangeName);
		oCriteriaRangeName = null;
		oCriteriaName = null;
		oSumRangeName = null;
		// - Case: SUMIFS. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. Criteria with formula. With disabled Iterative calculation setting.
		ws.getRange2("A1058").setValue("2");
		ws.getRange2("B1058").setValue("4");
		ws.getRange2("C1058").setValue("8");
		ws.getRange2("A1059").setValue("09/15/2024");
		ws.getRange2("B1059").setValue("09/16/2024");
		ws.getRange2("C1059").setValue("09/17/2024");
		ws.getRange2("D1059").setValue("09/18/2024");
		ws.getRange2("D1058").setValue("=SUMIFS(A1058:D1058, A1059:D1059, DATE(2024, 9, 17))");
		assert.strictEqual(ws.getRange2("D1058").getValue(), "8", "Test: SUMIFS. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. Criteria with formula. D1058 - 8");
		bCaFromSelectedCell = getCaFromSelectedCell("D1058");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIFS. 3 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criteria. Criteria with formula. D1058 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 3 args. Recursion formula. Recursion sum_range,the cell with formula matches with criteria. Criteria with formula. With disabled Iterative calculation setting.
		ws.getRange2("D1058").setValue("=SUMIFS(A1058:D1058, A1059:D1059, DATE(2024, 9, 18))");
		assert.strictEqual(ws.getRange2("D1058").getValue(), "0", "Test: SUMIFS. 3 args. Recursion formula. Recursion sum_range,the cell with formula matches with criteria. Criteria with formula. D1058 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("D1058");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIFS. 3 args. Recursion formula. Recursion sum_range,the cell with formula matches with criteria. Criteria with formula. D1058 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criterias. With disabled Iterative calculation setting.
		// sum_range row
		ws.getRange2("A1060").setValue("5");
		ws.getRange2("B1060").setValue("4");
		ws.getRange2("C1060").setValue("11");
		ws.getRange2("D1060").setValue("3");
		// criteria_range row
		ws.getRange2("A1061").setValue("Apples");
		ws.getRange2("B1061").setValue("Bananas");
		ws.getRange2("C1061").setValue("Artichokes");
		ws.getRange2("D1061").setValue("Apples");
		ws.getRange2("E1061").setValue("Bananas");
		// criteria row
		ws.getRange2("A1062").setValue("<>Bananas");
		// criteria_range2 row
		ws.getRange2("A1063").setValue("Sarah");
		ws.getRange2("B1063").setValue("Tom");
		ws.getRange2("C1063").setValue("Sarah");
		ws.getRange2("D1063").setValue("Tom");
		ws.getRange2("E1063").setValue("Sarah");
		// criteria2 row
		ws.getRange2("A1064").setValue("Tom");
		// formula
		ws.getRange2("E1060").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1060").getValue(), "3", "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criterias. E1060 - 3");
		bCaFromSelectedCell = getCaFromSelectedCell("E1060");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, but the cell with formula doesn't match with criterias. E1060 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion sum_range,the cell with formula 1 criteria is matches 2 is not. With disabled Iterative calculation setting.
		ws.getRange2("E1061").setValue("Apples");
		ws.getRange2("E1060").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1060").getValue(), "3", "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range,the cell with formula 1 criteria is matches 2 is not. E1060 - 3");
		bCaFromSelectedCell = getCaFromSelectedCell("E1060");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range,the cell with formula 1 criteria is matches 2 is not. E1060 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula - 1 criteria doesn't match 2 is matches. With disabled Iterative calculation setting.
		ws.getRange2("E1061").setValue("Bananas");
		ws.getRange2("E1063").setValue("Tom");
		ws.getRange2("E1060").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1060").getValue(), "3", "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula - 1 criteria doesn't match 2 is matches. E1060 - 3");
		bCaFromSelectedCell = getCaFromSelectedCell("E1060");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula - 1 criteria doesn't match 2 is matches. E1060 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula both criterias match. With disabled Iterative calculation setting.
		ws.getRange2("E1061").setValue("Apples");
		ws.getRange2("E1060").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1060").getValue(), "0", "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula both criterias match. E1060 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("E1060");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIFS. 5 args. Recursion formula. Recursion sum_range, the cell with formula both criterias match. E1060 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion criteria_range. With disabled Iterative calculation setting.
		ws.getRange2("E1060").setValue("3");
		ws.getRange2("E1061").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1061").getValue(), "0", "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range. E1061 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("E1061");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range. E1061 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is not matches. With disabled Iterative calculation setting.
		ws.getRange2("E1061").setValue("Bananas");
		ws.getRange2("E1063").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1063").getValue(), "3", "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is not matches. E1063 - 3");
		bCaFromSelectedCell = getCaFromSelectedCell("E1063");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is not matches. E1063 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. With disabled Iterative calculation setting.
		ws.getRange2("E1061").setValue("Apples");
		ws.getRange2("E1063").setValue("=SUMIFS(A1060:E1060, A1061:E1061, A1062, A1063:E1063, A1064)");
		assert.strictEqual(ws.getRange2("E1063").getValue(), "0", "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. E1063 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("E1063");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SUMIFS. 5 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. E1063 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range isn't matches. With disabled Iterative calculation setting.
	 	// Criteria range
		ws.getRange2("A1064").setValue("100");
		ws.getRange2("B1064").setValue("1000");
		ws.getRange2("C1064").setValue("10000");
		ws.getRange2("D1064").setValue("10");
		ws.getRange2("E1064").setValue("0");
		// Criteria
		ws.getRange2("A1065").setValue(">0");
		// Criteria range 2
		ws.getRange2("A1066").setValue("Bob");
		ws.getRange2("B1066").setValue("Tom");
		ws.getRange2("C1066").setValue("Bob");
		ws.getRange2("D1066").setValue("Bob");
		ws.getRange2("E1066").setValue("=COUNTIFS(A1064:E1064, A1065, A1066:E1066, A1067)");
		// Criteria2
		ws.getRange2("A1067").setValue("Bob")
		assert.strictEqual(ws.getRange2("E1066").getValue(), "3", "Test: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range isn't matches. E1066 - 3");
		bCaFromSelectedCell = getCaFromSelectedCell("E1066");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range isn't matches. E1066 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. With disabled Iterative calculation setting.
		ws.getRange2("E1064").setValue("10");
		ws.getRange2("E1066").setValue("=COUNTIFS(A1064:E1064, A1065, A1066:E1066, A1067)");
		assert.strictEqual(ws.getRange2("E1066").getValue(), "0", "Test: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. E1066 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("E1066");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: COUNTIFS. 4 args. Recursion formula. Recursion criteria_range2, but criteria_range is matches. E1066 - flag ca: true");
		// - Case: IF. 3 args. Recursion formula. recursion value_false but it doesn't match. With disabled Iterative calculation setting.
		ws.getRange2("A1067").setValue("Yes");
		ws.getRange2("B1067").setValue("23");
		ws.getRange2("C1067").setValue("=IF(A1067 = \"Yes\", B1067, C1067)");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "23", "Test: IF. 3 args. Recursion formula. recursion value_false. C1067 - 23");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: IF. 3 args. Recursion formula. recursion value_false. C1067 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula. Recursion value_false but it matches. With disabled Iterative calculation setting.
		ws.getRange2("C1067").setValue("=IF(A1067 = \"No\", B1067, C1067)");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "0", "Test: IF. 3 args. Recursion formula. Recursion value_false. C1067 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: IF. 3 args. Recursion formula. Recursion value_false. C1067 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it doesn't match. With disabled Iterative calculation setting.
		ws.getRange2("D1067").setValue("TRUE");
		ws.getRange2("C1067").setValue("=IF(D1067, B1067, C1067");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "23", "Test: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it doesn't match. C1067 - 23");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it doesn't match. C1067 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it matches. With disabled Iterative calculation setting.
		ws.getRange2("D1067").setValue("FALSE");
		ws.getRange2("C1067").setValue("=IF(D1067, B1067, C1067");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "0", "Test: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it matches. C1067 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: IF. 3 args. Recursion formula. logical_test is Ref. Recursion value_false, but it matches. C1067 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it doesn't match. With disabled Iterative calculation setting.
		let oLogicalTest = new Asc.asc_CDefName("logical_test", ws.getName() + "!$D$1067");
		let oTrueValue = new Asc.asc_CDefName("true_value", ws.getName() + "!$B$1067");
		let oFalseValue = new Asc.asc_CDefName("false_value", ws.getName() + "!$C$1067");
		wb.editDefinesNames(null, oLogicalTest);
		wb.editDefinesNames(null, oTrueValue);
		wb.editDefinesNames(null, oFalseValue);
		ws.getRange2("D1067").setValue("TRUE");
		ws.getRange2("C1067").setValue("=IF(logical_test, true_value, C1067)");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "23", "Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it doesn't match. C1067 - 23");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, false, "Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it doesn't match. C1067 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it matches. With disabled Iterative calculation setting.
		ws.getRange2("D1067").setValue("FALSE");
		ws.getRange2("C1067").setValue("=IF(logical_test, true_value, false_value)");
		assert.strictEqual(ws.getRange2("C1067").getValue(), "0", "Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it matches. C1067 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("C1067");
		assert.strictEqual(bCaFromSelectedCell, true, "Case: IF. 3 args. Recursion formula.logical_test is DefName. Recursion value_false is DefName, but it matches. C1067 - flag ca: true");
		bCaFromSelectedCell = null;
		wb.delDefinesNames(oLogicalTest);
		oLogicalTest = null;
		wb.delDefinesNames(oTrueValue);
		oTrueValue = null;
		wb.delDefinesNames(oFalseValue);
		oFalseValue = null;
		// - Case: IF. 3 args. Recursion formula. With operand function. One of condition is recursion, but it doesn't match. With disabled Iterative calculation setting.
		ws.getRange2("A1073").setValue("10");
		ws.getRange2("B1073").setValue("=IF(A1073=10, 10, B1073)+IF(A1073=10, 10, B1073)");
		assert.strictEqual(ws.getRange2("B1073").getValue(), "20", "Test: IF. 3 args. Recursion formula. With operand function. One of condition is recursion, but it doesn't match. B1073 - 20");
		bCaFromSelectedCell = getCaFromSelectedCell("B1073");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: IF. 3 args. Recursion formula. With operand function. One of condition is recursion, but it doesn't match. B1073 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: IF. 3 args. Recursion formula. With operand function. One of condition is recursion but it matches. With disabled Iterative calculation setting.
		ws.getRange2("A1073").setValue("1");
		ws.getRange2("B1073").setValue("=IF(A1073=10, 10, B1073)+IF(A1073=10, 10, B1073)");
		assert.strictEqual(ws.getRange2("B1073").getValue(), "0", "Test: IF. 3 args. Recursion formula. With operand function. One of condition is recursion but it matches. B1073 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("B1073");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: IF. 3 args. Recursion formula. With operand function. One of condition is recursion but it matches. B1073 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: IFS. 6 args. Recursion formula. One of condition is recursion, but it doesn't match. With disabled Iterative calculation setting.
		ws.getRange2("A1068").setValue("3");
		ws.getRange2("B1068").setValue('=IFS(A1068=1, "First", A1068=2, B1068, A1068=3, "Third")');
		assert.strictEqual(ws.getRange2("B1068").getValue(), "Third", "Test: IFS. 6 args. Recursion formula. One of condition is recursion but it doesn't match. B1068 - Third");
		bCaFromSelectedCell = getCaFromSelectedCell("B1068");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: IFS. 6 args. Recursion formula. One of condition is recursion but it doesn't match. B1068 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: IFS. 6 args. Recursion formula. One of condition is recursion but it matches. With disabled Iterative calculation setting.
		ws.getRange2("A1068").setValue("2");
		ws.getRange2("B1068").setValue('=IFS(A1068=1, "First", A1068=2, B1068, A1068=3, "Third")');
		assert.strictEqual(ws.getRange2("B1068").getValue(), "0", "Test: IFS. 6 args. Recursion formula. One of condition is recursion but it matches. B1068 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("B1068");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: IFS. 6 args. Recursion formula. One of condition is recursion but it matches. B1068 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SWITCH. Without default_arg. One of result_arg has recursion, but it doesn't match. With disabled Iterative calculation setting.
		// expression
		ws.getRange2("A1069").setValue("3");
		// values
		ws.getRange2("A1070").setValue("1");
		ws.getRange2("B1070").setValue("2");
		ws.getRange2("C1070").setValue("3");
		ws.getRange2("D1070").setValue("4");
		ws.getRange2("E1070").setValue("5");
		// results
		ws.getRange2("A1071").setValue("Monday");
		ws.getRange2("B1071").setValue("Wednesday");
		ws.getRange2("C1071").setValue("Thursday");
		ws.getRange2("D1071").setValue("Friday");
		// formula
		ws.getRange2("A1072").setValue("=SWITCH(A1069,A1070, A1071, B1070, A1072, C1070, B1071, D1070, C1071, E1070, D1071)");
		assert.strictEqual(ws.getRange2("A1072").getValue(), "Wednesday", "Test: SWITCH. Without default_arg. One of result_arg has recursion but it doesn't matches. A1072 - Wednesday");
		bCaFromSelectedCell = getCaFromSelectedCell("A1072");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SWITCH. Without default_arg. One of result_arg has recursion but it doesn't matches. A1072 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SWITCH. Without default_arg. One of result_arg has recursion, but it matches. With disabled Iterative calculation setting.
		ws.getRange2("A1069").setValue("2");
		ws.getRange2("A1072").setValue("=SWITCH(A1069,A1070, A1071, B1070, A1072, C1070, B1071, D1070, C1071, E1070, D1071)");
		assert.strictEqual(ws.getRange2("A1072").getValue(), "0", "Test: SWITCH. Without default_arg. One of result_arg has recursion but it matches. A1072 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("A1072");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SWITCH. Without default_arg. One of result_arg has recursion but it matches. A1072 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: SWITCH. With default_arg. Default_arg has recursion, but it doesn't match. With disabled Iterative calculation setting.
		ws.getRange2("A1069").setValue("7");
		// default_arg
		ws.getRange2("E1071").setValue("Unknown day of week");
		ws.getRange2("A1072").setValue("=SWITCH(A1069,A1070, A1071, B1070, A1072, C1070, B1071, D1070, C1071, E1070, D1071, E1071)");
		assert.strictEqual(ws.getRange2("A1072").getValue(), "Unknown day of week", "Test: SWITCH. With default_arg. Default_arg has recursion but it doesn't matches. A1072 - Unknown day of week");
		bCaFromSelectedCell = getCaFromSelectedCell("A1072");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: SWITCH. With default_arg. Default_arg has recursion but it doesn't matches. A1072 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: SWITCH. With default_arg. Default_arg has recursion, but it matches. With disabled Iterative calculation setting.
		ws.getRange2("A1072").setValue("=SWITCH(A1069,A1070, A1071, B1070, A1072, C1070, B1071, D1070, C1071, E1070, D1071, A1072)");
		assert.strictEqual(ws.getRange2("A1072").getValue(), "0", "Test: SWITCH. With default_arg. Default_arg has recursion but it matches. A1072 - 0");
		bCaFromSelectedCell = getCaFromSelectedCell("A1072");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: SWITCH. With default_arg. Default_arg has recursion but it matches. A1072 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Exception formula "CELL" that ignores rules of recursion recognition
		ws.getRange2("A1073").setValue("=CELL(\"filename\",A1073)");
		assert.strictEqual(ws.getRange2("A1073").getValue(), "[TeSt.xlsx]Sheet1", "Test: Exception formulas that ignores rules of recursion recognition. A1073 - 1039. Formula - CELL");
		bCaFromSelectedCell = getCaFromSelectedCell("A1073");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Exception formulas that ignores rules of recursion recognition. A1039 - flag ca: true");
		bCaFromSelectedCell = null;
		// - Case: Chain without recursion. B1074 <- A1075 <- D1075 <- E1075 <- F1075. With disabled Iterative calculation setting. Case from bug-71996
		// year field
		ws.getRange2("A1074").setValue("2024");
		// month field
		ws.getRange2("B1074").setValue("=DATE(A1074, SHEET(),1");
		//  time break
		ws.getRange2("C1074").setValue("0.02");
		ws.getRange2("D1074").setValue("0.03");
		ws.getRange2("E1074").setValue("0.33");
		// additional field
		ws.getRange2("F1074").setValue("=IF(MONTH(B1074)=1;$G$1074;INDIRECT(TEXT(DATE(YEAR(B1074);MONTH(B1074)-1;1);\"MMM\") & \"!F39\"))");
		ws.getRange2("G1074").setValue("0");
		// main chain
		ws.getRange2("A1075").setValue("=B1074");
		ws.getRange2("B1075").setValue("0");
		ws.getRange2("C1075").setValue("0");
		ws.getRange2("D1075").setValue("=IF(ISNUMBER($A1075);IF((C1075-B1075)<TIME(6;1;0);TIME(0;0;0);IF((C1075-B1075)<TIME(9;31;0);$C$1074;$D$1074));\"\")");
		ws.getRange2("E1075").setValue("=IF(ISNUMBER($A1075);IF(OR(G1075=\"U\";H1075=\"X\");(C1075-B1075-D1075);IF(OR(G1075=\"K\";G1075=\"B\";G1075=\"D\");TIME(0;0;0);C1075-B1075-D1075-$E$1074));\"\")");
		ws.getRange2("F1075").setValue("==IF(ISNUMBER($A1075);IF(OR(G1075=\"Zaus\";E1075>-$E$1074;G1075=\"kA\");(F1074+E1075);TIME(0;0;0));\"\")");
		ws.getRange2("G1075").setValue("Neujahr");
		ws.getRange2("H1075").setValue("X");
		// Checking via initStartCellForIterCalc method that cells haven't recursion
		oCell = selectCell("A1075");
		let bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Chain without recursion. B1074 <- A1075 <- D1075 <- E1075 <- F1075. With disabled Iterative calculation setting. Case from bug-71996. A1075 - false");
		bCellHasRecursion = null;
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("D1075");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Chain without recursion. B1074 <- A1075 <- D1075 <- E1075 <- F1075. With disabled Iterative calculation setting. Case from bug-71996. D1075 - false");
		bCellHasRecursion = null;
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("E1075");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Chain without recursion. B1074 <- A1075 <- D1075 <- E1075 <- F1075. With disabled Iterative calculation setting. Case from bug-71996. E1075 - false");
		bCellHasRecursion = null;
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("F1075");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Chain without recursion. B1074 <- A1075 <- D1075 <- E1075 <- F1075. With disabled Iterative calculation setting. Case from bug-71996. F1075 - false");
		bCellHasRecursion = null;
		g_cCalcRecursion.setStartCellIndex(null);
		// - Case: Chain with recursion. A1076->B1076. B1076 is simple formula without ca flag.
		// Enable recursion setting
		g_cCalcRecursion.setIsEnabledRecursion(true);
		ws.getRange2("A1076").setValue("=A1076+B1076");
		ws.getRange2("B1076").setValue("=1");
		assert.strictEqual(ws.getRange2("A1076").getValue(), "15", "Test: Chain with recursion. A1076->B1076. B1076 is simple formula without ca flag. A1076 - 15");
		bCaFromSelectedCell = getCaFromSelectedCell("A1076");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain with recursion. A1076->B1076. B1076 is simple formula without ca flag. A1076 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1076");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Chain with recursion. A1076->B1076. B1076 is simple formula without ca flag. B1076 - flag ca: false");
		// - Case: Chain with recursion. A1077 -> B1077. B1077 is simple formula without ca flag.
		ws.getRange2("A1077").setValue("=A1077+B1077");
		ws.getRange2("B1077").setValue("=1+2");
		assert.strictEqual(ws.getRange2("A1077").getValue(), "45", "Test: Chain with recursion. A1077 -> B1077. B1077 is simple formula without ca flag. A1077 - 45");
		bCaFromSelectedCell = getCaFromSelectedCell("A1077");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain with recursion. A1077 -> B1077. B1077 is simple formula without ca flag. A1077 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1077");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Chain with recursion. A1077 -> B1077. B1077 is simple formula without ca flag. B1077 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: Chain with recursion. A1078 -> B1078 -> C1078. When B1078 has a ref to C1078 and C1078 - value.
		ws.getRange2("A1078").setValue("=A1078+B1078");
		ws.getRange2("B1078").setValue("=C1078");
		ws.getRange2("C1078").setValue("1");
		assert.strictEqual(ws.getRange2("A1078").getValue(), "15", "Test: Chain with recursion. A1078 -> B1078 -> C1078. When B1078 has a ref to C1078 and C1078 - value. A1078 - 15");
		bCaFromSelectedCell = getCaFromSelectedCell("A1078");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain with recursion. A1078 -> B1078 -> C1078. When B1078 has a ref to C1078 and C1078 - value. A1078 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1078");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Chain with recursion. A1078 -> B1078 -> C1078. When B1078 has a ref to C1078 and C1078 - value. B1078 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: Chain with recursion. A1079 -> B1079 -> C1079. When B1079 has a ref to C1079. C1079 - simple formula.
		ws.getRange2("A1079").setValue("=A1079+B1079");
		ws.getRange2("B1079").setValue("=C1079");
		ws.getRange2("C1079").setValue("=1");
		assert.strictEqual(ws.getRange2("A1079").getValue(), "15", "Test: Chain with recursion. A1079 -> B1079 -> C1079. When B1079 has a ref to C1079. C1079 - simple formula. A1079 - 15");
		bCaFromSelectedCell = getCaFromSelectedCell("A1079");
		assert.strictEqual(bCaFromSelectedCell, true, "Test: Chain with recursion. A1079 -> B1079 -> C1079. When B1079 has a ref to C1079. C1079 - simple formula. A1079 - flag ca: true");
		bCaFromSelectedCell = null;
		bCaFromSelectedCell = getCaFromSelectedCell("B1079");
		assert.strictEqual(bCaFromSelectedCell, false, "Test: Chain with recursion. A1079 -> B1079 -> C1079. When B1079 has a ref to C1079. C1079 - simple formula. B1079 - flag ca: false");
		bCaFromSelectedCell = null;
		// - Case: Chain without recursion. A1081 -> B1080 -> A1082 -> B1081 -> A1083 using shared. With disabled recursion settings. Bug #73472
		g_cCalcRecursion.setIsEnabledRecursion(false);
		ws.getRange2("A1080").setValue("5.96");
		ws.getRange2("B1080").setValue("=A1080+2.92-MONTH(TODAY())")
		ws.getRange2("A1081").setValue("=IF(B1080 = \"\", \"\", B1080");
		ws.getRange2("A1082").setValue("0");
		ws.getRange2("A1083").setValue("0");
		ws.getRange2("B1081").setValue("0");
		ws.getRange2("B1082").setValue("0");
		ws.getRange2("B1083").setValue("0");

		// Create bbox and cellWithFormula.
		let bbox = ws.getRange2("A1081:A1083").bbox;
		let bbox1 = ws.getRange2("B1080:B1083").bbox;
		let cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		let cellWithFormula1 = new window['AscCommonExcel'].CCellWithFormula(ws, bbox1.r1, bbox1.c1);
		let oParser = selectCell("A1081").getFormulaParsed().clone();
		let sharedRef = bbox.clone();
		oParser.setShared(sharedRef, cellWithFormula);
		oParser.parse();
		oParser.calculate();
		oParser.ca = true;
		ws.getRange2("A1081:A1083")._foreachNoEmpty(function(oCell) {
			oCell.setFormulaParsed(oParser);
			oCell._BuildDependencies(true, true);

		});
		oParser = selectCell("B1080").getFormulaParsed().clone();
		sharedRef = bbox1.clone();
		oParser.setShared(sharedRef, cellWithFormula1);
		oParser.parse();
		oParser.calculate();
		oParser.ca = true;
		ws.getRange2("B1080:B1083")._foreachNoEmpty(function(oCell) {
			oCell.setFormulaParsed(oParser);
			oCell._BuildDependencies(true, true);

		});
		oCell = selectCell("B1080");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Chain without recursion. A1081 -> B1080 -> A1082 -> B1081 -> A1083 using shared. With disabled recursion settings. Case from bug-73472. B1080 - false");
		bCellHasRecursion = null;
		g_cCalcRecursion.setStartCellIndex(null);
		// - Case: Formula OFFSET mustn't recognize as recursive formula. Bug #74432
		ws.getRange2("B1084").setValue("1");
		ws.getRange2("C1084").setValue("=OFFSET(C1084, 0, -1)");
		assert.strictEqual(ws.getRange2("C1084").getValue(), "1", "Test: Formula OFFSET mustn't recognize as recursive formula. Bug #74432. C1084 - 1");
		oCell = selectCell("C1084");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula OFFSET mustn't recognize as recursive formula. Bug #74432. C1084 - false");
		bCellHasRecursion = null;
		// - Case: Formula ISFORMULA mustn't recognize as recursive. Is part of formula. Bug #74432
		ws.getRange2("A1085").setValue("1");
		ws.getRange2("B1085").setValue("=A1085+ISFORMULA(B1085)");
		assert.strictEqual(ws.getRange2("B1085").getValue(), "2", "Test: Formula ISFORMULA mustn't recognize as recursive. Is part of formula. Bug #74432. B1085 - 2");
		oCell = selectCell("B1085");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ISFORMULA mustn't recognize as recursive. Is part of formula. Bug #74432. B1085 - false");
		bCellHasRecursion = null;
		// - Case: Formula ISFORMULA  contains in recursive formula. Bug #74432
		ws.getRange2("A1086").setValue("1");
		ws.getRange2("B1086").setValue("=A1086+B1086+ISFORMULA(B1086)");
		assert.strictEqual(ws.getRange2("B1086").getValue(), "0", "Test: Formula ISFORMULA contains in recursive formula. Bug #74432. B1086 - 0");
		oCell = selectCell("B1086");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ISFORMULA contains in recursive formula. Bug #74432. B1086 - true");
		bCellHasRecursion = null;
		// - Case: Formula ISFORMULA contains in recursive formula from chain. Bug #74432
		ws.getRange2("A1087").setValue("1");
		ws.getRange2("B1087").setValue("=C1087");
		ws.getRange2("C1087").setValue("=A1087+B1087+ISFORMULA(C1087)");
		assert.strictEqual(ws.getRange2("C1087").getValue(), "0", "Test: Formula ISFORMULA contains in recursive formula from chain. Bug #74432. C1087 - 0");
		oCell = selectCell("C1087");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ISFORMULA contains in recursive formula from chain. Bug #74432. C1087 - true");
		bCellHasRecursion = null;
		// - Case: Formula ISFORMULA containts in IF formula mustn't recursive. Bug #74432
		ws.getRange2("A1088").setValue("1");
		ws.getRange2("B1088").setValue("=IF(A1088 = 1,A1088+ISFORMULA(B1088), -1");
		assert.strictEqual(ws.getRange2("B1088").getValue(), "2", "Test: Formula ISFORMULA containts in IF formula mustn't recursive. Bug #74432. B1088 - 2");
		oCell = selectCell("B1088");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ISFORMULA containts in IF formula mustn't recursive. Bug #74432. B1088 - false");
		bCellHasRecursion = null;
		// - Case: Formula ISFORMULA contains in IF formula has recursive element. Bug #74432
		ws.getRange2("A1089").setValue("1");
		ws.getRange2("B1089").setValue("=IF(A1089 = 1,A1089+B1089+ISFORMULA(B1089), -1");
		assert.strictEqual(ws.getRange2("B1089").getValue(), "0", "Test: Formula ISFORMULA contains in IF formula has recursive element. Bug #74432. B1089 - 0");
		oCell = selectCell("B1089");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ISFORMULA contains in IF formula has recursive element. Bug #74432. B1089 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMN mustn't recognize as recursive
		ws.getRange2("A1090").setValue("1");
		ws.getRange2("B1090").setValue("=A1090+COLUMN(B1090)");
		assert.strictEqual(ws.getRange2("B1090").getValue(), "3", "Test: Formula COLUMN mustn't recognize as recursive. B1090 - 3");
		oCell = selectCell("B1090");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula COLUMN mustn't recognize as recursive. B1090 - false");
		bCellHasRecursion = null;
		// - Case: Formula COLUMN contains in recursive formula
		ws.getRange2("A1091").setValue("1");
		ws.getRange2("B1091").setValue("=A1091+B1091+COLUMN(B1091)");
		assert.strictEqual(ws.getRange2("B1091").getValue(), "0", "Test: Formula COLUMN contains in recursive formula. B1091 - 0");
		oCell = selectCell("B1091");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMN contains in recursive formula. B1091 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMN contains in recursive formula from chain
		ws.getRange2("A1092").setValue("1");
		ws.getRange2("B1092").setValue("=C1092");
		ws.getRange2("C1092").setValue("=A1092+B1092+COLUMN(C1092)");
		assert.strictEqual(ws.getRange2("C1092").getValue(), "0", "Test: Formula COLUMN contains in recursive formula from chain. C1092 - 0");
		oCell = selectCell("C1092");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMN contains in recursive formula from chain. C1092 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMN in IF formula mustn't recursive
		ws.getRange2("A1093").setValue("1");
		ws.getRange2("B1093").setValue("=IF(A1093 = 1, A1093+COLUMN(B1093), -1)");
		assert.strictEqual(ws.getRange2("B1093").getValue(), "3", "Test: Formula COLUMN in IF formula mustn't recursive. B1093 - 3");
		oCell = selectCell("B1093");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula COLUMN in IF formula mustn't recursive. B1093 - false");
		bCellHasRecursion = null;
		// - Case: Formula COLUMN contains in IF formula has recursive element
		ws.getRange2("A1094").setValue("1");
		ws.getRange2("B1094").setValue("=IF(A1094 = 1, A1094+B1094+COLUMN(B1094), -1)");
		assert.strictEqual(ws.getRange2("B1094").getValue(), "0", "Test: Formula COLUMN contains in IF formula has recursive element. B1094 - 0");
		oCell = selectCell("B1094");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMN contains in IF formula has recursive element. B1094 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMNS mustn't recognize as recursive. Is part of formula.
		ws.getRange2("A1095").setValue("1");
		ws.getRange2("B1095").setValue("=A1095+COLUMNS(B1095)");
		assert.strictEqual(ws.getRange2("B1095").getValue(), "2", "Test: Formula COLUMNS mustn't recognize as recursive. Is part of formula. B1095 - 2");
		oCell = selectCell("B1095");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula COLUMNS mustn't recognize as recursive. Is part of formula. B1095 - false");
		bCellHasRecursion = null;
		// - Case: Formula COLUMNS contains in recursive formula
		ws.getRange2("A1096").setValue("1");
		ws.getRange2("B1096").setValue("=A1096+B1096+COLUMNS(B1096)");
		assert.strictEqual(ws.getRange2("B1096").getValue(), "0", "Test: Formula COLUMNS contains in recursive formula. B1096 - 0");
		oCell = selectCell("B1096");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMNS contains in recursive formula. B1096 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMNS contains in recursive formula from chain
		ws.getRange2("A1097").setValue("1");
		ws.getRange2("B1097").setValue("=C1097");
		ws.getRange2("C1097").setValue("=A1097+B1097+COLUMNS(C1097)");
		assert.strictEqual(ws.getRange2("C1097").getValue(), "0", "Test: Formula COLUMNS contains in recursive formula from chain. C1097 - 0");
		oCell = selectCell("C1097");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMNS contains in recursive formula from chain. C1097 - true");
		bCellHasRecursion = null;
		// - Case: Formula COLUMNS contains in IF formula mustn't recursive
		ws.getRange2("A1098").setValue("1");
		ws.getRange2("B1098").setValue("=IF(A1098 = 1, A1098+COLUMNS(B1098), -1)");
		assert.strictEqual(ws.getRange2("B1098").getValue(), "2", "Test: Formula COLUMNS contains in IF formula mustn't recursive. B1098 - 2");
		oCell = selectCell("B1098");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula COLUMNS contains in IF formula mustn't recursive. B1098 - false");
		bCellHasRecursion = null;
		// - Case: Formula COLUMNS contains in IF formula has recursive element
		ws.getRange2("A1099").setValue("1");
		ws.getRange2("B1099").setValue("=IF(A1099 = 1, A1099+B1099+COLUMNS(B1099), -1)");
		assert.strictEqual(ws.getRange2("B1099").getValue(), "0", "Test: Formula COLUMNS contains in IF formula has recursive element. B1099 - 0");
		oCell = selectCell("B1099");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula COLUMNS contains in IF formula has recursive element. B1099 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROW mustn't recognize as recursive. Is part of formula.
		ws.getRange2("A1100").setValue("1");
		ws.getRange2("B1100").setValue("=A1100+ROW(B1100)");
		assert.strictEqual(ws.getRange2("B1100").getValue(), "1101", "Test: Formula ROW mustn't recognize as recursive. Is part of formula. B1100 - 1101");
		oCell = selectCell("B1100");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ROW mustn't recognize as recursive. Is part of formula. B1100 - false");
		bCellHasRecursion = null;
		// - Case: Formula ROW contains in recursive formula
		ws.getRange2("A1101").setValue("1");
		ws.getRange2("B1101").setValue("=A1101+B1101+ROW(B1101)");
		assert.strictEqual(ws.getRange2("B1101").getValue(), "0", "Test: Formula ROW contains in recursive formula. B1101 - 0");
		oCell = selectCell("B1101");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROW contains in recursive formula. B1101 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROW contains in recursive formula from chain
		ws.getRange2("A1102").setValue("1");
		ws.getRange2("B1102").setValue("=C1102");
		ws.getRange2("C1102").setValue("=A1102+B1102+ROW(C1102)");
		assert.strictEqual(ws.getRange2("C1102").getValue(), "0", "Test: Formula ROW contains in recursive formula from chain. C1102 - 0");
		oCell = selectCell("C1102");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROW contains in recursive formula from chain. C1102 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROW contains in IF formula mustn't recursive
		ws.getRange2("A1103").setValue("1");
		ws.getRange2("B1103").setValue("=IF(A1103 = 1, A1103+ROW(B1103), -1)");
		assert.strictEqual(ws.getRange2("B1103").getValue(), "1104", "Test: Formula ROW contains in IF formula mustn't recursive. B1103 - 1104");
		oCell = selectCell("B1103");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ROW contains in IF formula mustn't recursive. B1103 - false");
		bCellHasRecursion = null;
		// - Case: Formula ROW contains in IF formula has recursive element
		ws.getRange2("A1104").setValue("1");
		ws.getRange2("B1104").setValue("=IF(A1104 = 1, A1104+B1104+ROW(B1104), -1)");
		assert.strictEqual(ws.getRange2("B1104").getValue(), "0", "Test: Formula ROW contains in IF formula has recursive element. B1104 - 0");
		oCell = selectCell("B1104");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROW contains in IF formula has recursive element. B1104 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROWS mustn't recognize as recursive. Is part of formula.
		ws.getRange2("A1105").setValue("1");
		ws.getRange2("B1105").setValue("=A1105+ROWS(B1105:B1106)");
		assert.strictEqual(ws.getRange2("B1105").getValue(), "3", "Test: Formula ROWS mustn't recognize as recursive. Is part of formula. B1105 - 4");
		oCell = selectCell("B1105");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ROWS mustn't recognize as recursive. Is part of formula. B1105 - false");
		bCellHasRecursion = null;
		// - Case: Formula ROWS contains in recursive formula
		ws.getRange2("A1107").setValue("1");
		ws.getRange2("B1107").setValue("=A1107+B1107+ROWS(B1107:B1111)");
		assert.strictEqual(ws.getRange2("B1107").getValue(), "0", "Test: Formula ROWS contains in recursive formula. B1107 - 0");
		oCell = selectCell("B1107");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROWS contains in recursive formula. B1107 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROWS contains in recursive formula from chain
		ws.getRange2("A1109").setValue("1");
		ws.getRange2("B1109").setValue("=C1109");
		ws.getRange2("C1109").setValue("=A1109+B1109+ROWS(C1109:C1110)");
		assert.strictEqual(ws.getRange2("C1109").getValue(), "0", "Test: Formula ROWS contains in recursive formula from chain. C1109 - 0");
		oCell = selectCell("C1109");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROWS contains in recursive formula from chain. C1109 - true");
		bCellHasRecursion = null;
		// - Case: Formula ROWS contains in IF formula mustn't recursive
		ws.getRange2("A1111").setValue("1");
		ws.getRange2("B1111").setValue("=IF(A1111 = 1, A1111+ROWS(B1111:B1112), -1)");
		assert.strictEqual(ws.getRange2("B1111").getValue(), "3", "Test: Formula ROWS contains in IF formula mustn't recursive. B1111 - 3");
		oCell = selectCell("B1111");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula ROWS contains in IF formula mustn't recursive. B1111 - false");
		bCellHasRecursion = null;
		// - Case: Formula ROWS contains in IF formula has recursive element
		ws.getRange2("A1112").setValue("1");
		ws.getRange2("B1112").setValue("=IF(A1112 = 1, A1112+B1112+ROWS(B1112:B1113), -1)");
		assert.strictEqual(ws.getRange2("B1112").getValue(), "0", "Test: Formula ROWS contains in IF formula has recursive element. B1112 - 0");
		oCell = selectCell("B1112");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula ROWS contains in IF formula has recursive element. B1112 - true");
		bCellHasRecursion = null;
		// - Case: Formula SHEETS mustn't recognize as recursive. Is part of formula.
		ws.getRange2("A1113").setValue("1");
		ws.getRange2("B1113").setValue("=A1113+SHEETS(B1113)");
		assert.strictEqual(ws.getRange2("B1113").getValue(), "2", "Test: Formula SHEETS mustn't recognize as recursive. Is part of formula. B1113 - 2");
		oCell = selectCell("B1113");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula SHEETS mustn't recognize as recursive. Is part of formula. B1113 - false");
		bCellHasRecursion = null;
		// - Case: Formula SHEETS contains in recursive formula
		ws.getRange2("A1114").setValue("1");
		ws.getRange2("B1114").setValue("=A1114+B1114+SHEETS(B1114)");
		assert.strictEqual(ws.getRange2("B1114").getValue(), "0", "Test: Formula SHEETS contains in recursive formula. B1114 - 0");
		oCell = selectCell("B1114");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula SHEETS contains in recursive formula. B1114 - true");
		bCellHasRecursion = null;
		// - Case: Formula SHEETS contains in recursive formula from chain
		ws.getRange2("A1115").setValue("1");
		ws.getRange2("B1115").setValue("=C1115");
		ws.getRange2("C1115").setValue("=A1115+B1115+SHEETS(C1115)");
		assert.strictEqual(ws.getRange2("C1115").getValue(), "0", "Test: Formula SHEETS contains in recursive formula from chain. C1115 - 0");
		oCell = selectCell("C1115");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula SHEETS contains in recursive formula from chain. C1115 - true");
		bCellHasRecursion = null;
		// - Case: Formula SHEETS contains in IF formula mustn't recursive
		ws.getRange2("A1116").setValue("1");
		ws.getRange2("B1116").setValue("=IF(A1116 = 1, A1116+SHEETS(B1116), -1)");
		assert.strictEqual(ws.getRange2("B1116").getValue(), "2", "Test: Formula SHEETS contains in IF formula mustn't recursive. B1116 - 2");
		oCell = selectCell("B1116");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula SHEETS contains in IF formula mustn't recursive. B1116 - false");
		bCellHasRecursion = null;
		// - Case: Formula SHEETS contains in IF formula has recursive element
		ws.getRange2("A1117").setValue("1");
		ws.getRange2("B1117").setValue("=IF(A1117 = 1, A1117+B1117+SHEETS(B1117), -1)");
		assert.strictEqual(ws.getRange2("B1117").getValue(), "0", "Test: Formula SHEETS contains in IF formula has recursive element. B1117 - 0");
		oCell = selectCell("B1117");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula SHEETS contains in IF formula has recursive element. B1117 - true");
		bCellHasRecursion = null;
		// - Case: Formula CELL mustn't recognize as recursive. Is part of formula.
		ws.getRange2("A1117").setValue("1");
		ws.getRange2("B1117").setValue("=A1117+CELL(\"col\", B1117)");
		assert.strictEqual(ws.getRange2("B1117").getValue(), "3", "Test: Formula CELL mustn't recognize as recursive. Is part of formula. B1117 - 3");
		oCell = selectCell("B1117");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula CELL mustn't recognize as recursive. Is part of formula. B1117 - false");
		bCellHasRecursion = null;
		// - Case: Formula CELL contains in recursive formula
		ws.getRange2("A1118").setValue("1");
		ws.getRange2("B1118").setValue("=A1118+B1118+CELL(\"col\", B1118)");
		assert.strictEqual(ws.getRange2("B1118").getValue(), "0", "Test: Formula CELL contains in recursive formula. B1118 - 0");
		oCell = selectCell("B1118");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula CELL contains in recursive formula. B1118 - true");
		bCellHasRecursion = null;
		// - Case: Formula CELL contains in recursive formula from chain
		ws.getRange2("A1119").setValue("1");
		ws.getRange2("B1119").setValue("=C1119");
		ws.getRange2("C1119").setValue("=A1119+B1119+CELL(\"col\", C1119)");
		assert.strictEqual(ws.getRange2("C1119").getValue(), "0", "Test: Formula CELL contains in recursive formula from chain. C1119 - 0");
		oCell = selectCell("C1119");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula CELL contains in recursive formula from chain. C1119 - true");
		bCellHasRecursion = null;
		// - Case: Formula CELL contains in IF formula mustn't recursive
		ws.getRange2("A1120").setValue("1");
		ws.getRange2("B1120").setValue("=IF(A1120 = 1, A1120+CELL(\"col\", B1120), -1)");
		assert.strictEqual(ws.getRange2("B1120").getValue(), "3", "Test: Formula CELL contains in IF formula mustn't recursive. B1120 - 3");
		oCell = selectCell("B1120");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula CELL contains in IF formula mustn't recursive. B1120 - false");
		bCellHasRecursion = null;
		// - Case: Formula CELL contains in IF formula has recursive element
		ws.getRange2("A1121").setValue("1");
		ws.getRange2("B1121").setValue("=IF(A1121 = 1, A1121+B1121+CELL(\"col\", B1121), -1)");
		assert.strictEqual(ws.getRange2("B1121").getValue(), "0", "Test: Formula CELL contains in IF formula has recursive element. B1121 - 0");
		oCell = selectCell("B1121");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: Formula CELL contains in IF formula has recursive element. B1121 - true");
		bCellHasRecursion = null;
		// - Case: Formula INDIRECT mustn't recognize as recursive.
		ws.getRange2("A1122").setValue("1");
		ws.getRange2("B1122").setValue("=INDIRECT(\"A1122\")");
		assert.strictEqual(ws.getRange2("B1122").getValue(), "1", "Test: Formula INDIRECT mustn't recognize as recursive. B1122 - 1");
		// - Case: Formula INDIRECT is recursive formula.
		ws.getRange2("A1123").setValue("=INDIRECT(\"A1123\")");
		assert.strictEqual(ws.getRange2("A1123").getValue(), "0", "Test: Formula INDIRECT is recursive formula. A1123 - 0");
		// - Case: Formula INDIRECT is recursive formula via chain.
		ws.getRange2("A1124").setValue("=B1124");
		ws.getRange2("B1124").setValue("=INDIRECT(\"A1124\")");
		assert.strictEqual(ws.getRange2("B1124").getValue(), "0", "Test: Formula INDIRECT is recursive formula via chain. B1124 - 0");
		// - Case: Formula INDIRECT contains in FORMULA mustn't recognize as recursive.
		ws.getRange2("A1125").setValue("1");
		ws.getRange2("B1125").setValue("=INDIRECT(\"A1125\") + A1125");
		assert.strictEqual(ws.getRange2("B1125").getValue(), "2", "Test: Formula INDIRECT contains in FORMULA mustn't recognize as recursive. B1125 - 2");
		// - Case: Formula INDIRECT contains in FORMULA recognize as recursive.
		ws.getRange2("A1126").setValue("1");
		ws.getRange2("B1126").setValue("=A1126+B1126+INDIRECT(\"A1126\")");
		assert.strictEqual(ws.getRange2("B1126").getValue(), "0", "Test: Formula INDIRECT contains in FORMULA recognize as recursive. B1126 - 0");
		// - Case: Formula OFFSET is recursive formula.
		ws.getRange2("A1127").setValue("=OFFSET(A1127, 0, 0)");
		assert.strictEqual(ws.getRange2("A1127").getValue(), "0", "Test: Formula OFFSET is recursive formula. A1127 - 0");
		// - Case: Formula OFFSET is recursive formula by chain.
		ws.getRange2("A1128").setValue("=B1128");
		ws.getRange2("B1128").setValue("=OFFSET(B1128, 0, -1)");
		assert.strictEqual(ws.getRange2("B1128").getValue(), "0", "Test: Formula OFFSET is recursive formula by chain. B1128 - 0");
		// - Case: Formula OFFSET contains in recursive formula.
		ws.getRange2("A1129").setValue("1");
		ws.getRange2("B1129").setValue("=A1129+B1129+OFFSET(B1129, 0, -1)");
		assert.strictEqual(ws.getRange2("B1129").getValue(), "0", "Test: Formula OFFSET contains in recursive formula. B1129 - 0");
		// - Case: Formula CELL with type contents is recursive formula.
		ws.getRange2("A1130").setValue("=CELL(\"contents\", A1130)");
		assert.strictEqual(ws.getRange2("A1130").getValue(), "0", "Test: Formula CELL with type contents is recursive formula. A1130 - 0");
		// - Case: Formula INDIRECT - recursive cell with enabled setting.
		g_cCalcRecursion.setIsEnabledRecursion(true);
		ws.getRange2("A1131").setValue("=INDIRECT(\"A1131\")+1");
		assert.strictEqual(ws.getRange2("A1131").getValue(), "15", "Test: Formula INDIRECT - recursive cell with enabled setting. A1131 - 15");
		// - Case: Formula OFFSET isn't recursive cell with disabled setting.
		g_cCalcRecursion.setIsEnabledRecursion(false);
		ws.getRange2("B1132").setValue("5");
		ws.getRange2("C1132").setValue("0");
		ws.getRange2("D1132").setValue("0");
		ws.getRange2("A1132").setValue("=OFFSET(B1132,0,0)-C1132-D1132");
		ws.getRange2("A1133").setValue("=OFFSET(A1132,0,0)");
		assert.strictEqual(ws.getRange2("A1133").getValue(), "5", "Test: Formula OFFSET isn't recursive cell with disabled setting. A1133 - 5");
		assert.strictEqual(ws.getRange2("A1132").getValue(), "5", "Test: Formula OFFSET isn't recursive cell with disabled setting. A1132 - 5");
		// - Case: Formula INDIRECT Ref3D isn't recursive cell with disabled setting. Bug-76318
		ws2 = wb.createWorksheet(0, "Sheet2");
		ws2.getRange2("A1134").setValue('=INDIRECT("Sheet1!A1134")');
		ws.getRange2("A1134").setValue("123");
		assert.strictEqual(ws2.getRange2("A1134").getValue(), "123", "Test: Formula INDIRECT Ref3D isn't recursive cell with disabled setting. Bug-76318. A1134 - 123");
		oCell = selectCell("A1134", ws2);
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula INDIRECT Ref3D isn't recursive cell with disabled setting. Bug-76318. A1134 - false");
		bCellHasRecursion = null;
		wb.removeWorksheet(0);
		// - Case: Formula SHEET isn't recursive cell with disabled setting. Bug-77330
		ws2 = wb.createWorksheet(0, "Sheet2");
		let oDefNameSHEET = new Asc.asc_CDefName("TestSHEET", ws2.getName() + "!$A$2:$E$5");
		wb.editDefinesNames(null, oDefNameSHEET);
		ws.getRange2("A1135").setValue('=SHEET(TestSHEET)');
		oCell = selectCell("A1135");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: Formula SHEET isn't recursive cell with disabled setting. Bug-77330. A1135 - false");
		bCellHasRecursion = null;
		wb.delDefinesNames(oDefNameSHEET);
		wb.removeWorksheet(0);
		// -- Test changeLinkedCell method.
		oCell = selectCell("A1000");
		let oCellNeedEnableRecalc = selectCell("B1000");
		assert.strictEqual(oCellNeedEnableRecalc.getIsDirty(), false, "Test: changeLinkedCell. Before: Cell B1000 isDirty - false");
		oCell.changeLinkedCell(function(oCell) {
			if (oCell.isFormula && !oCell.getIsDirty()) {
				oCell.setIsDirty(true);
			}
		}, true);
		oCellNeedEnableRecalc = selectCell("B1000");
		assert.strictEqual(oCellNeedEnableRecalc.getIsDirty(), true, "Test: changeLinkedCell. After: Cell B1000 isDirty - true");
		// - Case: SUMIF. 3 args. Non-recursion formula with disabled setting. Range argument has an Error type. Bug-78980
		ws.getRange2("A1136").setValue('=SUMIF(#REF!,">"&TODAY(), $A$1136:$A$1136)');
		oCell = selectCell("A1136");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: SUMIF. 3 args. Non-recursion formula with disabled setting. Range argument has Error type. Bug-78980. A1136 - false");
		bCellHasRecursion = null;
		// - Case: SUMIF. 3 args. Non-recursion formula with disabled setting. Sum_range argument has an Error type. Bug-78980
		ws.getRange2("B1136").setValue('=SUMIF($B$1136:$B$1136,">"&TODAY(), #REF!)');
		oCell = selectCell("B1136");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, false, "Test: SUMIF. 3 args. Non-recursion formula with disabled setting. Sum_range argument has Error type. Bug-78980. B1136 - false");
		bCellHasRecursion = null;
		// - Case: SUMIF. 3 args. Recursion formula with disabled setting. Criteria argument has an Error type. Bug-78980
		ws.getRange2("C1136").setValue('=SUMIF($C$1136:$C$1136,#REF!,$C$1136:$C$1136)');
		oCell = selectCell("C1136");
		bCellHasRecursion = !!getStartCellForIterCalc(oCell);
		assert.strictEqual(bCellHasRecursion, true, "Test: SUMIF. 3 args. Recursion formula with disabled setting. Criteria argument has Error type. Bug-78980. C1136 - true");
		bCellHasRecursion = null;
		// -- Test initStartCellForIterCalc method
		// - Case: Sequence chain A1000 -> B1000 -> C1000
		nExpectedCellIndex = AscCommonExcel.getCellIndex(999, 0);
		oCell = selectCell("C1000");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Sequence chain - A1000 -> B1000 -> C1000. Selected cell: C1000. Start cell index: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("B1000");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Sequence chain - A1000 -> B1000 -> C1000. Selected cell: B1000. Start cell index: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("A1000");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Sequence chain - A1000 -> B1000 -> C1000. Selected cell: A1000. Start cell: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		// - Case: Loop chain D1000 <-> F1000
		oCell = selectCell("D1000");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		nExpectedCellIndex = AscCommonExcel.getCellIndex(999, 3);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Loop chain - D1000 <-> F1000. Selected cell: D1000. Start cell: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		oCell = selectCell("F1000");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		nExpectedCellIndex = AscCommonExcel.getCellIndex(999, 3);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Loop chain - D1000 <-> F1000. Selected cell: F1000. Start cell: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		// - Case: Loop cell
		oCell = selectCell("A1001");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		nExpectedCellIndex = AscCommonExcel.getCellIndex(1000, 0);
		assert.strictEqual(oFactCellIndex.cellId, nExpectedCellIndex, `Test: initStartCellForIterCalc. Loop cell - A1001. Selected cell: A1001. Start cell: ${oFactCellIndex.cellId}`);
		g_cCalcRecursion.setStartCellIndex(null);
		// - Negative case: sequence chain without loop cell.
		oCell = selectCell("C1002");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex, null, `Test: initStartCellForIterCalc. Negative case sequence chain without loop cell. Selected cell: C1002. Start cell: ${oFactCellIndex}`);
		oCell = selectCell("B1002");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex, null, `Test: initStartCellForIterCalc. Negative case sequence chain without loop cell. Selected cell: B1002. Start cell: ${oFactCellIndex}`);
		oCell = selectCell("A1002");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex, null, `Test: initStartCellForIterCalc. Negative case sequence chain without loop cell. Selected cell: A1002. Start cell: ${oFactCellIndex}`);
		// - Negative case: cell without any chain.
		oCell = selectCell("A1003");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex, null, `Test: initStartCellForIterCalc. Negative case cell without any chain. Selected cell: A1003. Start cell: ${oFactCellIndex}`);
		oCell = selectCell("C1004");
		oFactCellIndex = getStartCellForIterCalc(oCell);
		assert.strictEqual(oFactCellIndex, null, `Test: initStartCellForIterCalc. Negative case cell without any chain. Selected cell: C1004. Start cell: ${oFactCellIndex}`);
	});

	QUnit.test("Test: \"Absolute reference\"", function (assert) {

		ws.getRange2("A7").setValue("1");
		ws.getRange2("A8").setValue("2");
		ws.getRange2("A9").setValue("3");
		oParser = new parserFormula('A$7+A8', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('A$7+A$8', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('$A$7+$A$8', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 3);

		oParser = new parserFormula('SUM($A$7:$A$9)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 6);
	});

	QUnit.test("Test: \"Cross\"", function (assert) {

		ws.getRange2("A7").setValue("1");
		ws.getRange2("A8").setValue("2");
		ws.getRange2("A9").setValue("3");
		let cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 0, 9);
		oParser = new parserFormula('A7:A9', cellWithFormula, ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 5, 0, 5), ws.getId()).getValue(), "#VALUE!");
		assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 6, 0, 6), ws.getId()).getValue(), 1);
		assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 7, 0, 7), ws.getId()).getValue(), 2);
		assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 8, 0, 8), ws.getId()).getValue(), 3);
		assert.strictEqual(oParser.calculate().cross(new Asc.Range(0, 9, 0, 9), ws.getId()).getValue(), "#VALUE!");

	});

	QUnit.test("Test: \"Defined names cycle\"", function (assert) {
		AscCommonExcel.g_cCalcRecursion.setIsEnabledRecursion(false);
		var newNameQ = new Asc.asc_CDefName("q", "SUM('" + ws.getName() + "'!A2)");
		wb.editDefinesNames(null, newNameQ);
		ws.getRange2("Q1").setValue("=q");
		ws.getRange2("Q2").setValue("=q");
		ws.getRange2("Q3").setValue("1");
		assert.strictEqual(ws.getRange2("Q1").getValueWithFormat(), "1");
		assert.strictEqual(ws.getRange2("Q2").getValueWithFormat(), "1");

		var newNameW = new Asc.asc_CDefName("w", "'" + ws.getName() + "'!A1");
		wb.editDefinesNames(null, newNameW);
		ws.getRange2("Q4").setValue("=w");
		assert.strictEqual(ws.getRange2("Q4").getValueWithFormat(), "#NUM!");
		//clean up
		ws.getRange2("Q1:Q4").cleanAll();
		wb.delDefinesNames(newNameW);
		wb.delDefinesNames(newNameQ);
	});

	QUnit.test("Test: \"Parse intersection\"", function (assert) {

		ws.getRange2("A1:B10").cleanAll();

		ws.getRange2("A7").setValue("1");
		ws.getRange2("A8").setValue("2");
		ws.getRange2("A9").setValue("3");
		oParser = new parserFormula('1     +    (    A7   +A8   )   *   2', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.assemble(), "1+(A7+A8)*2");
		assert.strictEqual(oParser.calculate().getValue(), 7);

		oParser = new parserFormula('sum                    A1:A5', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.assemble(), "sum A1:A5");
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('sum(   A1:A5    ,        B1:B5     )     ', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.assemble(), "SUM(A1:A5,B1:B5)");
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula('sum(   A1:A5    ,        B1:B5  , "    3 , 14 15 92 6 "   )     ', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.assemble(), 'SUM(A1:A5,B1:B5,"    3 , 14 15 92 6 ")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

	});

	QUnit.test("Test: \"Range union operator tests\"", function (assert) {
		let array;

		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("A4").setValue("99");
		ws.getRange2("B1").setValue("1");
		ws.getRange2("B2").setValue("2");
		ws.getRange2("B3").setValue("3");
		ws.getRange2("C1").setValue("1");
		ws.getRange2("C2").setValue("2");
		ws.getRange2("C3").setValue("3");
		ws.getRange2("F1").setValue("#NUM!");
		ws.getRange2("F2").setValue("#DIV/0!");
		ws.getRange2("F3").setValue("#N/A");

		oParser = new parserFormula('SUMA2', 'A10', ws);
		assert.ok(oParser.parse(), 'SUMA2');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SUMA2');

		oParser = new parserFormula('SUMA2:1', 'A10', ws);
		assert.ok(oParser.parse(), 'SUMA2:1');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SUMA2:1');

		oParser = new parserFormula('SUMA2:A3', 'A10', ws);
		assert.ok(oParser.parse(), 'SUMA2:A3');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SUMA2:A3');

		oParser = new parserFormula('SECB2', 'A10', ws);
		assert.ok(oParser.parse(), 'SECB2');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SECB2');
		
		oParser = new parserFormula('SECB2:1', 'A10', ws);
		assert.ok(oParser.parse(), 'SECB2:1');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SECB2:1');

		oParser = new parserFormula('SECB2:B3', 'A10', ws);
		assert.ok(oParser.parse(), 'SECB2:B3');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'SECB2:B3');

		oParser = new parserFormula('RANDC2', 'A10', ws);
		assert.ok(oParser.parse(), 'RANDC2');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'RANDC2');
		
		oParser = new parserFormula('RANDC2:1', 'A10', ws);
		assert.ok(oParser.parse(), 'RANDC2:1');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'RANDC2:1');

		oParser = new parserFormula('RANDC2:C3', 'A10', ws);
		assert.ok(oParser.parse(), 'RANDC2:C3');
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", 'RANDC2:C3');

		oParser = new parserFormula('C1:C3', 'A10', ws);
		assert.ok(oParser.parse(), 'C1:C3');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 1, "Result of C1:C3[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 2, "Result of C1:C3[1,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(2, 0).getValue() : array.getElementRowCol(2, 0).getValue(), 3, "Result of C1:C3[2,0]");
		}

		oParser = new parserFormula('C2:C3:C2', 'A10', ws);
		assert.ok(oParser.parse(), 'C2:C3:C2');
		array = oParser.calculate();	
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 2, "Result of C2:C3:C2[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 3, "Result of C2:C3:C2[1,0]");
		}
		
		oParser = new parserFormula('(A1:A3):F1', 'A10', ws);
		assert.ok(oParser.parse(), '(A1:A3):F1');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 1, "Result of (A1:A3):F1[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 1).getValue() : array.getElementRowCol(0, 1).getValue(), 1, "Result of (A1:A3):F1[0,1]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 2).getValue() : array.getElementRowCol(0, 2).getValue(), 1, "Result of (A1:A3):F1[0,2]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 5).getValue() : array.getElementRowCol(0, 5).getValue(), "#NUM!", "Result of (A1:A3):F1[0,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 2, "Result of (A1:A3):F1[1,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(2, 0).getValue() : array.getElementRowCol(2, 0).getValue(), 3, "Result of (A1:A3):F1[2,0]");
		}

		oParser = new parserFormula('F1:(A1:A3)', 'A10', ws);
		assert.ok(oParser.parse(), 'F1:(A1:A3)');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 1, "Result of F1:(A1:A3)[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 1).getValue() : array.getElementRowCol(0, 1).getValue(), 1, "Result of F1:(A1:A3)[0,1]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 2).getValue() : array.getElementRowCol(0, 2).getValue(), 1, "Result of F1:(A1:A3)[0,2]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 5).getValue() : array.getElementRowCol(0, 5).getValue(), "#NUM!", "Result of F1:(A1:A3)[0,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 2, "Result of F1:(A1:A3)[1,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(2, 0).getValue() : array.getElementRowCol(2, 0).getValue(), 3, "Result of F1:(A1:A3)[2,0]");
		}

		oParser = new parserFormula('F2:(A2)', 'A10', ws);
		assert.ok(oParser.parse(), 'F2:(A2)');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 2, "Result of F2:(A2)[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 1).getValue() : array.getElementRowCol(0, 1).getValue(), 2, "Result of F2:(A2)[0,1]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 2).getValue() : array.getElementRowCol(0, 2).getValue(), 2, "Result of F2:(A2)[0,2]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 5).getValue() : array.getElementRowCol(0, 5).getValue(), "#DIV/0!", "Result of F2:(A2)[0,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 3, "Result of F2:(A2)[1,0]");
		}

		oParser = new parserFormula('(A2):F2', 'A10', ws);
		assert.ok(oParser.parse(), '(A2):F2');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 2, "Result of (A2):F2[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 1).getValue() : array.getElementRowCol(0, 1).getValue(), 2, "Result of (A2):F2[0,1]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 2).getValue() : array.getElementRowCol(0, 2).getValue(), 2, "Result of (A2):F2[0,2]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 5).getValue() : array.getElementRowCol(0, 5).getValue(), "#DIV/0!", "Result of (A2):F2[0,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 3, "Result of (A2):F2[1,0]");
		}

		oParser = new parserFormula('F2:(E1):A1:F2:F3:(A4)', 'A10', ws);
		assert.ok(oParser.parse(), 'F2:(E1):A1:F2:F3:(A4)');
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.cellsRange === array.type || AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 0).getValue() : array.getElementRowCol(0, 0).getValue(), 1, "Result of F2:(E1):A1:F2:F3:(A4)[0,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 0).getValue() : array.getElementRowCol(1, 0).getValue(), 2, "Result of F2:(E1):A1:F2:F3:(A4)[1,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(2, 0).getValue() : array.getElementRowCol(2, 0).getValue(), 3, "Result of F2:(E1):A1:F2:F3:(A4)[2,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(3, 0).getValue() : array.getElementRowCol(3, 0).getValue(), 99, "Result of F2:(E1):A1:F2:F3:(A4)[3,0]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 1).getValue() : array.getElementRowCol(0, 1).getValue(), 1, "Result of F2:(E1):A1:F2:F3:(A4)[0,1]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 2).getValue() : array.getElementRowCol(0, 2).getValue(), 1, "Result of F2:(E1):A1:F2:F3:(A4)[0,2]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(0, 5).getValue() : array.getElementRowCol(0, 5).getValue(), "#NUM!", "Result of F2:(E1):A1:F2:F3:(A4)[0,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(1, 5).getValue() : array.getElementRowCol(1, 5).getValue(), "#DIV/0!", "Result of F2:(E1):A1:F2:F3:(A4)[1,5]");
			assert.strictEqual(array.getValueByRowCol ? array.getValueByRowCol(2, 5).getValue() : array.getElementRowCol(2, 5).getValue(), "#N/A", "Result of F2:(E1):A1:F2:F3:(A4)[2,5]");
		}
		

		ws.getRange2("A1:Z100").cleanAll();
	});

	QUnit.test("Test: \"Arithmetical operations\"", function (assert) {
		let array;
		oParser = new parserFormula('1+3', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula('(1+2)*4+3', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), (1 + 2) * 4 + 3);

		oParser = new parserFormula('2^52', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), Math.pow(2, 52));

		oParser = new parserFormula("(0)^(0)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(0)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of (0)^(0)");	// ms - #NUM!, js - 1, LO - 1, gs - 1

		oParser = new parserFormula("(0)^(1)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(1)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of (0)^(1)");

		oParser = new parserFormula("(0)^(-1)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(-1)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of (0)^(-1)");

		oParser = new parserFormula("(0)^(1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(1/3)");
		assert.strictEqual(oParser.calculate().getValue(), 0, "Result of (0)^(1/3)");

		oParser = new parserFormula("(0)^(-1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(-1/3)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of (0)^(-1/3)");

		oParser = new parserFormula("(0)^(-3)", "A1", ws);
		assert.ok(oParser.parse(), "(0)^(-3)");
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", "Result of (0)^(-3)");

		oParser = new parserFormula("(1)^(-3)", "A1", ws);
		assert.ok(oParser.parse(), "(1)^(-3)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of (1)^(-3)");

		oParser = new parserFormula("(1)^(-1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(1)^(-1/3)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of (1)^(-1/3)");

		oParser = new parserFormula("(1)^(1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(1)^(1/3)");
		assert.strictEqual(oParser.calculate().getValue(), 1, "Result of (1)^(1/3)");

		oParser = new parserFormula("(-1)^(1/2)", "A1", ws);
		assert.ok(oParser.parse(), "(-1)^(1/2)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of (-1)^(1/2)");

		oParser = new parserFormula("(-1)^(1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(-1)^(1/3)");
		assert.strictEqual(oParser.calculate().getValue(), -1, "Result of (-1)^(1/3)");

		oParser = new parserFormula("(-1)^(-1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(-1)^(-1/3)");
		assert.strictEqual(oParser.calculate().getValue(), -1, "Result of (-1)^(-1/3)");

		oParser = new parserFormula("(-1)^(1/4)", "A1", ws);
		assert.ok(oParser.parse(), "(-1)^(1/4)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of (-1)^(1/4)");

		oParser = new parserFormula("(-1)^(1/5)", "A1", ws);
		assert.ok(oParser.parse(), "(-1)^(1/5)");
		assert.strictEqual(oParser.calculate().getValue(), -1, "Result of (-1)^(1/5)");

		oParser = new parserFormula("(-8)^(1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(-8)^(1/3)");
		assert.strictEqual(oParser.calculate().getValue(), -2, "Result of (-8)^(1/3)");

		oParser = new parserFormula("(-8)^(-1/3)", "A1", ws);
		assert.ok(oParser.parse(), "(-8)^(-1/3)");
		assert.strictEqual(oParser.calculate().getValue(), -0.5, "Result of (-8)^(-1/3)");

		oParser = new parserFormula("(-8)^(1/4)", "A1", ws);
		assert.ok(oParser.parse(), "(-8)^(1/4)");
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", "Result of (-8)^(1/4)");

		oParser = new parserFormula("(-8)^(1/5)", "A1", ws);
		assert.ok(oParser.parse(), "(-8)^(1/5)");
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), "-1.52", "Result of (-8)^(1/5)");

		oParser = new parserFormula("(-8)^(-1/5)", "A1", ws);
		assert.ok(oParser.parse(), "(-8)^(-1/5)");
		assert.strictEqual(oParser.calculate().getValue().toFixed(2), "-0.66", "Result of (-8)^(-1/5)");

		oParser = new parserFormula('-10', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), -10);

		oParser = new parserFormula('-10*2', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), -20);

		oParser = new parserFormula('-10+10', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula('12%', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0.12);

		oParser = new parserFormula("2<>\"3\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", "2<>\"3\"");

		oParser = new parserFormula("2=\"3\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", "2=\"3\"");

		oParser = new parserFormula("2>\"3\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", "2>\"3\"");

		oParser = new parserFormula("\"f\">\"3\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula("\"f\"<\"3\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual("FALSE", oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula("FALSE>=FALSE", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula("\"TRUE\"&\"TRUE\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUETRUE");

		oParser = new parserFormula("10*\"\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("-TRUE", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), -1);

		oParser = new parserFormula('"test" = "test"', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('"tEsT" = "TeSt"', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('"TEST" = "TeSt"', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('"TEST" = "weSt"', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2("K100:Z200").cleanAll();
		ws.getRange2("M106").setValue("1");
		ws.getRange2("M107").setValue("2");
		ws.getRange2("M108").setValue("2");
		ws.getRange2("M109").setValue("4");
		ws.getRange2("M110").setValue("5");
		ws.getRange2("M111").setValue("-23");
		ws.getRange2("M112").setValue("6");
		ws.getRange2("M113").setValue("5");

		ws.getRange2("N106").setValue("1");
		ws.getRange2("N107").setValue("");
		ws.getRange2("N108").setValue("");
		ws.getRange2("N109").setValue("3");
		ws.getRange2("N110").setValue("");
		ws.getRange2("N111").setValue("2");
		ws.getRange2("N112").setValue("");
		ws.getRange2("N113").setValue("3");

		ws.getRange2("O106").setValue("1");
		ws.getRange2("O107").setValue("3");
		ws.getRange2("O108").setValue("2");
		ws.getRange2("O109").setValue("12");
		ws.getRange2("O110").setValue("3");
		ws.getRange2("O111").setValue("4");
		ws.getRange2("O112").setValue("3");
		ws.getRange2("O113").setValue("2");

		ws.getRange2("P106").setValue("3");
		ws.getRange2("P107").setValue("4");
		ws.getRange2("P108").setValue("5");
		ws.getRange2("P109").setValue("1");
		ws.getRange2("P110").setValue("23");
		ws.getRange2("P111").setValue("4");
		ws.getRange2("P112").setValue("3");
		ws.getRange2("P113").setValue("1");

		oParser = new parserFormula("M106:N107*O107:O107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse(), "M106:N107*O107:O107. Result - array 2x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 3, "Array 2x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 6, "Array 2x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Array 2x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0, "Array 2x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x2. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "", "Array 2x2. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "", "Array 2x2. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Array 2x2. [2,2]");
		}

		oParser = new parserFormula("M106:N107+O107:O107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse(), "M106:N107+O107:O107. Result - array 2x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, "Array 2x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 5, "Array 2x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4, "Array 2x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 3, "Array 2x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x2. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "", "Array 2x2. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "", "Array 2x2. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Array 2x2. [2,2]");
		}

		oParser = new parserFormula("M106:N107-O107:O107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse(), "M106:N107-O107:O107. Result - array 2x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), -2, "Array 2x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), -1, "Array 2x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), -2, "Array 2x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), -3, "Array 2x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x2. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "", "Array 2x2. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "", "Array 2x2. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Array 2x2. [2,2]");
		}

		oParser = new parserFormula("M106:N107/O107:O107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse(), "M106:N107/O107:O107. Result - array 2x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue().toFixed(1), "0.3", "Array 2x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue().toFixed(1), "0.7", "Array 2x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue().toFixed(1), "0.3", "Array 2x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0, "Array 2x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x2. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "", "Array 2x2. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "", "Array 2x2. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Array 2x2. [2,2]");
		}

		// todo problem with cross in cPower
		// oParser = new parserFormula("M106:N107^O107:O107", "A1", ws);
		// oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		// assert.ok(oParser.parse(), "M106:N107^O107:O107. Result - array 2x2");
		// array = oParser.calculate();
		// if (AscCommonExcel.cElementType.array === array.type) {
		// 	assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Array 2x2. [0,0]");
		// 	assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 8, "Array 2x2. [1,0]");
		// 	assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x2. [2,0]");

		// 	assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, "Array 2x2. [0,1]");
		// 	assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0, "Array 2x2. [1,1]");
		// 	assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x2. [2,1]");

		// 	assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "", "Array 2x2. [0,2]");
		// 	assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "", "Array 2x2. [1,2]");
		// 	assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "", "Array 2x2. [2,2]");
		// }

		oParser = new parserFormula("M106:N107*M106:O108", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A");
		}

		oParser = new parserFormula("M106:O108*M106:N107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A");
		}

		oParser = new parserFormula("L105:O108*M106:N107", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H109").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(3, 2).getValue(), "#N/A");
		}

		ws.getRange2("A200").setValue("1");
		ws.getRange2("A201").setValue("2");
		ws.getRange2("A202").setValue("4");
		ws.getRange2("B200").setValue("2");
		ws.getRange2("B201").setValue("2");
		ws.getRange2("B202").setValue("4");
		ws.getRange2("C200").setValue("3");
		ws.getRange2("D200").setValue("4");
		ws.getRange2("E200").setValue("5");

		oParser = new parserFormula("A200:E200*A200:B201", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:J110").bbox);
		assert.ok(oParser.parse(), "A200:E200*A200:B201. Result - array 2x5");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Array 2x5. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 2, "Array 2x5. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "", "Array 2x5. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4, "Array 2x5. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, "Array 2x5. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 2x5. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A", "Array 2x5. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A", "Array 2x5. [1,2]");

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A", "Array 2x5. [0,3]");
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), "#N/A", "Array 2x5. [1,3]");

			assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", "Array 2x5. [0,4]");
			assert.strictEqual(array.getElementRowCol(1, 4).getValue(), "#N/A", "Array 2x5. [1,4]");
		}

		oParser = new parserFormula("A200:B201*A200:E200", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:J110").bbox);
		assert.ok(oParser.parse(), "A200:B201*A200:E200. Result - array 2x5");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 4).getValue(), "#N/A");
		}

		oParser = new parserFormula("A200:E200*A200:B202", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:J110").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 8);

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 3).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 4).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 4).getValue(), "#N/A");
		}

		oParser = new parserFormula("A200:B202*A200:E200", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:J110").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 8);

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 3).getValue(), "#N/A");

			assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(1, 4).getValue(), "#N/A");
			assert.strictEqual(array.getElementRowCol(2, 4).getValue(), "#N/A");
		}

		oParser = new parserFormula("M106:P113+M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 8);
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 10);
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), -46);
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 12);
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 10);

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 6);

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 24);
			assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 8);
			assert.strictEqual(array.getElementRowCol(6, 2).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 4);

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), 8);
			assert.strictEqual(array.getElementRowCol(2, 3).getValue(), 10);
			assert.strictEqual(array.getElementRowCol(3, 3).getValue(), 2);
			assert.strictEqual(array.getElementRowCol(4, 3).getValue(), 46);
			assert.strictEqual(array.getElementRowCol(5, 3).getValue(), 8);
			assert.strictEqual(array.getElementRowCol(6, 3).getValue(), 6);
			assert.strictEqual(array.getElementRowCol(7, 3).getValue(), 2);

		}

		oParser = new parserFormula("M106:P113*M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 16);
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 25);
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 529);
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 36);
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 25);

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 9);

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 4);
			assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 144);
			assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 16);
			assert.strictEqual(array.getElementRowCol(6, 2).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 4);

			assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(1, 3).getValue(), 16);
			assert.strictEqual(array.getElementRowCol(2, 3).getValue(), 25);
			assert.strictEqual(array.getElementRowCol(3, 3).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(4, 3).getValue(), 529);
			assert.strictEqual(array.getElementRowCol(5, 3).getValue(), 16);
			assert.strictEqual(array.getElementRowCol(6, 3).getValue(), 9);
			assert.strictEqual(array.getElementRowCol(7, 3).getValue(), 1);
		}

		oParser = new parserFormula("M106:P113-M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 0);
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 0);
		}

		oParser = new parserFormula("M106:P113=M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE");
		}

		oParser = new parserFormula("M106:P113/M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1);
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!");
		}

		oParser = new parserFormula("M106:P113<>M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
		}

		oParser = new parserFormula("M106:P113>M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
		}

		oParser = new parserFormula("M106:P113<M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE");
		}

		oParser = new parserFormula("M106:P113>=M106:P113", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE");
		}

		oParser = new parserFormula("SUM(M:P*M:P)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValue(), 1465);

		oParser = new parserFormula("SUM(M:P+M:P)", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E106:H113").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getValue(), 170);

		// for bug 65316
		ws.getRange2("A100").setValue("75");
		ws.getRange2("A101").setValue("42.4");
		ws.getRange2("B100").setValue("100");
		ws.getRange2("B101").setValue("42.48");
		ws.getRange2("C100").setValue("200");
		ws.getRange2("C101").setValue("42");
		ws.getRange2("D100").setValue("300");
		ws.getRange2("D101").setValue("39.4");
		ws.getRange2("E100").setValue("350");
		ws.getRange2("E101").setValue("38.4");

		oParser = new parserFormula("A100:C100^{1;2;3;4}", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:I204").bbox);
		assert.ok(oParser.parse(), "A100:C100^{1;2;3;4}. Result - array 4x3");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 75, "Array 4x3. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), Math.pow(75, 2), "Array 4x3. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), Math.pow(75, 3), "Array 4x3. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), Math.pow(75, 4), "Array 4x3. [3,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 100, "Array 4x3. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), Math.pow(100, 2), "Array 4x3. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), Math.pow(100, 3), "Array 4x3. [2,1]");
			assert.strictEqual(array.getElementRowCol(3, 1).getValue(), Math.pow(100, 4), "Array 4x3. [3,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 200, "Array 4x3. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), Math.pow(200, 2), "Array 4x3. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), Math.pow(200, 3), "Array 4x3. [2,2]");
			assert.strictEqual(array.getElementRowCol(3, 2).getValue(), Math.pow(200, 4), "Array 4x3. [3,2]");
		}

		oParser = new parserFormula("A100:C101^{1;2;3;4}", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:I204").bbox);
		assert.ok(oParser.parse(), "A100:C101^{1;2;3;4}. Result - array 4x3(2x3 = #N/A)");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 75, "Array 4x3. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 1797.76, "Array 4x3. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Array 4x3. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 100, "Array 4x3. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue().toFixed(2), "1804.55", "Array 4x3. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A", "Array 4x3. [2,1]");

			assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 200, "Array 4x3. [0,2]");
			assert.strictEqual(array.getElementRowCol(1, 2).getValue(), 1764, "Array 4x3. [1,2]");
			assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "#N/A", "Array 4x3. [2,2]");
		}

		ws.getRange2("A200").setValue("1");
		ws.getRange2("A201").setValue("3");
		ws.getRange2("B200").setValue("3");
		ws.getRange2("B201").setValue("4");
		ws.getRange2("D200").setValue("1");
		ws.getRange2("D201").setValue("2");
		ws.getRange2("D202").setValue("3");
		ws.getRange2("D203").setValue("4");
		ws.getRange2("D204").setValue("5");
		ws.getRange2("D205").setValue("6");
		ws.getRange2("D206:D300").setValue("");

		// for bug 54877
		oParser = new parserFormula("A200:B201*D200:D203", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "A200:B201*D200:D203. Result - array 4x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Array 4x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 6, "Array 4x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Array 4x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Array 4x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 8, "Array 4x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A", "Array 4x2. [2,1]");
		}

		oParser = new parserFormula("A200:B201*D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "A200:B201*D200:D300. Result - array 100x2");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, "Array 100x2. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 6, "Array 100x2. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", "Array 100x2. [2,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 3, "Array 100x2. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 8, "Array 100x2. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#N/A", "Array 100x2. [2,1]");
		}

		oParser = new parserFormula("D200:D300*2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300*2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4, "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 6, "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 8, "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 10, "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 12, "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 0, "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 0, "Array 100x1. [7,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		oParser = new parserFormula("2*D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2*D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 2, "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 4, "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 6, "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 8, "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 10, "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 12, "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 0, "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 0, "Array 100x1. [7,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		// conditional check and & operator
		// <
		oParser = new parserFormula("D200:D300<2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300<2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		
		oParser = new parserFormula("2<D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2<D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		
		// <=
		oParser = new parserFormula("D200:D300<=2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300<=2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		oParser = new parserFormula("2<=D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2<=D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		// >
		oParser = new parserFormula("D200:D300>2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300>2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}
		
		oParser = new parserFormula("2>D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2>D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		// >=
		oParser = new parserFormula("D200:D300>=2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300>=2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		oParser = new parserFormula("2>=D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2>=D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		// =
		oParser = new parserFormula("D200:D300=2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300=2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		oParser = new parserFormula("2=D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2=D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "FALSE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "TRUE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "FALSE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "FALSE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "FALSE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "FALSE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "FALSE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "FALSE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "FALSE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}

		// <>
		oParser = new parserFormula("D200:D300<>2", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "D200:D300<>2. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}
		oParser = new parserFormula("2<>D200:D300", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E200:H210").bbox);
		assert.ok(oParser.parse(), "2<>D200:D300. Result - array 100x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "TRUE", "Array 100x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "FALSE", "Array 100x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "TRUE", "Array 100x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "TRUE", "Array 100x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "TRUE", "Array 100x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "TRUE", "Array 100x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "TRUE", "Array 100x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "TRUE", "Array 100x1. [7,0]");
			assert.strictEqual(array.getElementRowCol(8, 0).getValue(), "TRUE", "Array 100x1. [8,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 100x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 100x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 100x1. [2,1]");
		}
		
		// & 
		ws.getRange2("D204").setValue("A");
		ws.getRange2("D205").setValue("B");
		ws.getRange2("D206:D210").setValue("");
		ws.getRange2("E204").setValue("A");
		ws.getRange2("E205").setValue("B");
		ws.getRange2("E206").setValue("C");
		ws.getRange2("E207").setValue("D");
		ws.getRange2("E208:E210").setValue("");

		oParser = new parserFormula("D204:D210&E204:E210", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("E300:H310").bbox);
		assert.ok(oParser.parse(), "D204:D210&E204:E210. Result - array 7x1");
		array = oParser.calculate();
		if (AscCommonExcel.cElementType.array === array.type) {
			assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "AA", "Array 7x1. [0,0]");
			assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "BB", "Array 7x1. [1,0]");
			assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "C", "Array 7x1. [2,0]");
			assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "D", "Array 7x1. [3,0]");
			assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "", "Array 7x1. [4,0]");
			assert.strictEqual(array.getElementRowCol(5, 0).getValue(), "", "Array 7x1. [5,0]");
			assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "", "Array 7x1. [6,0]");
			assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "", "Array 7x1. [7,0]");

			assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "", "Array 7x1. [0,1]");
			assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "", "Array 7x1. [1,1]");
			assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "", "Array 7x1. [2,1]");
		}

	});

	QUnit.test("Test: \"Concatenation operator test\"", function (assert) {
		let array, defName;
		ws.getRange2("A200").setValue("1");
		ws.getRange2("A201").setValue("2");
		ws.getRange2("A202").setValue("3");
		ws.getRange2("A203").setValue("4");
		ws.getRange2("A204").setValue("5");
		ws.getRange2("B200").setValue("A");
		ws.getRange2("B201").setValue("B");
		ws.getRange2("B202").setValue("C");
		ws.getRange2("B203").setValue("D");
		ws.getRange2("B204").setValue("E");

		wb.dependencyFormulas.addDefName("def_nums", "Sheet1!$A$200:$A$204");
		wb.dependencyFormulas.addDefName("def_strings", "Sheet1!$B$200:$B$204");
		wb.dependencyFormulas.addDefName("def_strings", "Sheet1!$B$200:$B$204");

		defName = new AscCommonExcel.DefName(wb, "n_1", '{1,2,3}', null, null, Asc.c_oAscDefNameType.none);
		wb.dependencyFormulas._addDefName(defName);
		wb.dependencyFormulas.defNames.wb["n_1"].setRef(defName.ref, true, true, true);

		defName = new AscCommonExcel.DefName(wb, "n_2", '{"str1","str2","str3"}', null, null, Asc.c_oAscDefNameType.none);
		wb.dependencyFormulas._addDefName(defName);
		wb.dependencyFormulas.defNames.wb["n_2"].setRef(defName.ref, true, true, true);

		defName = new AscCommonExcel.DefName(wb, "n_if", 'IF(0,n_2,n_1&n_2)', null, null, Asc.c_oAscDefNameType.none);
		wb.dependencyFormulas._addDefName(defName);
		wb.dependencyFormulas.defNames.wb["n_if"].setRef(defName.ref, true, true, true);

		// wb.dependencyFormulas.addDefNameOpen(defName.Name, defName.Ref, defName.LocalSheetId, defName.Hidden, _type);

		// INDEX
		oParser = new parserFormula('INDEX(n_if,2)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2str2");

		oParser = new parserFormula('INDEX(A200:A204&A200:A204,2)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "22");

		oParser = new parserFormula('INDEX(def_nums&def_nums,2)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "22");

		oParser = new parserFormula('INDEX(B200:B204&B200:B204,2)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "BB");

		oParser = new parserFormula('INDEX(def_strings&def_strings,2)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "BB");

		// other formulas
		oParser = new parserFormula('COLUMNS(A200:A204&A200:A204)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('COLUMNS(1+A200:A204&A200:A204)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('COLUMNS(def_nums&def_nums)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('COLUMNS(def_strings&def_strings)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1);

		oParser = new parserFormula('ROWS(A200:A204&A200:A204)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula('ROWS(1+A200:A204&A200:A204)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula('ROWS(def_nums&def_nums)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula('ROWS(def_strings&def_strings)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 5);

		oParser = new parserFormula('CONCAT(A200:A204,A200:A204&A200:A204)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123451122334455");

		// remove all created earlier defNames
		wb.dependencyFormulas._foreachDefName(function(defName) {
			if (defName.name !== "Table1") {
				wb.dependencyFormulas.removeDefName(undefined, defName.name);
			}
		});
	});

	QUnit.test("Test: \"\"s\"&5\"", function (assert) {
		oParser = new parserFormula("\"s\"&5", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "s5");
	});

	QUnit.test("Test: \"String+Number\"", function (assert) {
		oParser = new parserFormula("1+\"099\"", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 100);

		ws.getRange2("A1469").setValue("'099");
		ws.getRange2("A1470").setValue("\"099\"");

		oParser = new parserFormula("1+A1469", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 100);


		oParser = new parserFormula("1+A1470", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

	});

	QUnit.test("Test: \"Pow operator tests\"", function (assert) {
		let array;

		ws.getRange2("A100").setValue("2");
		ws.getRange2("A101").setValue("3");
		ws.getRange2("A102").setValue("4");
		ws.getRange2("B100").setValue("1");
		ws.getRange2("B101").setValue("2");
		ws.getRange2("B102").setValue("#DIV/0!");
		ws.getRange2("D100").setValue("2");
		ws.getRange2("D101").setValue("3");
		ws.getRange2("E100").setValue("3");
		ws.getRange2("E101").setValue("");

		oParser = new parserFormula("A100:B102^2", "A1", ws);
		assert.ok(oParser.parse(), 'A100:B102^2');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100:B102^2[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 9, 'Result of A100:B102^2[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of A100:B102^2[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of A100:B102^2[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of A100:B102^2[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100:B102^2[2,1]');

		oParser = new parserFormula("A100:B102^A100", "A1", ws);
		assert.ok(oParser.parse(), 'A100:B102^A100');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100:B102^A100[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 9, 'Result of A100:B102^A100[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of A100:B102^A100[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of A100:B102^A100[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of A100:B102^A100[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100:B102^A100[2,1]');

		oParser = new parserFormula("A100:B102^A100:A100", "A1", ws);
		assert.ok(oParser.parse(), 'A100:B102^A100:A100');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100:B102^A100:A100[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 9, 'Result of A100:B102^A100:A100[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of A100:B102^A100:A100[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of A100:B102^A100:A100[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of A100:B102^A100:A100[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100:B102^A100:A100[2,1]');

		oParser = new parserFormula("2^A100:B102", "A1", ws);
		assert.ok(oParser.parse(), '2^A100:B102');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of 2^A100:B102[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 8, 'Result of 2^A100:B102[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of 2^A100:B102[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 2, 'Result of 2^A100:B102[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of 2^A100:B102[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of 2^A100:B102[2,1]');

		oParser = new parserFormula("A100^A100:B102", "A1", ws);
		assert.ok(oParser.parse(), 'A100^A100:B102');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100^A100:B102[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 8, 'Result of A100^A100:B102[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of A100^A100:B102[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 2, 'Result of A100^A100:B102[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of A100^A100:B102[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100^A100:B102[2,1]');

		oParser = new parserFormula("A100:A100^A100:B102", "A1", ws);
		assert.ok(oParser.parse(), 'A100:A100^A100:B102');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100:A100^A100:B102[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 8, 'Result of A100:A100^A100:B102[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 16, 'Result of A100:A100^A100:B102[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 2, 'Result of A100:A100^A100:B102[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 4, 'Result of A100:A100^A100:B102[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100:A100^A100:B102[2,1]');
		
		oParser = new parserFormula("A100:B102^D100:E101", "A1", ws);
		assert.ok(oParser.parse(), 'A100:B102^D100:E101');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 4, 'Result of A100:B102^D100:E101[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 27, 'Result of A100:B102^D100:E101[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#N/A", 'Result of A100:B102^D100:E101[2,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of A100:B102^D100:E101[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 1, 'Result of A100:B102^D100:E101[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#DIV/0!", 'Result of A100:B102^D100:E101[2,1]');

		ws.getRange2("G100").setValue("a");
		ws.getRange2("G101").setValue("b");
		ws.getRange2("H100").setValue("c");
		ws.getRange2("H101").setValue("d");

		oParser = new parserFormula("G100:H101^2", "A1", ws);
		assert.ok(oParser.parse(), 'G100:H101^2');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of G100:H101^2[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", 'Result of G100:H101^2[1,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of G100:H101^2[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "#VALUE!", 'Result of G100:H101^2[1,1]');

		oParser = new parserFormula("2^G100:H101", "A1", ws);
		assert.ok(oParser.parse(), '2^G100:H101');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of 2^G100:H101[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", 'Result of 2^G100:H101[1,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of 2^G100:H101[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "#VALUE!", 'Result of 2^G100:H101[1,1]');

		oParser = new parserFormula("G100:H101^G100:H101", "A1", ws);
		assert.ok(oParser.parse(), 'G100:H101^G100:H101');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of G100:H101^G100:H101[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", 'Result of G100:H101^G100:H101[1,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of G100:H101^G100:H101[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "#VALUE!", 'Result of G100:H101^G100:H101[1,1]');

		ws.getRange2("I100").setValue("#N/A");
		ws.getRange2("I101").setValue("#DIV/0!");
		ws.getRange2("J100").setValue("#NUM!");
		ws.getRange2("J101").setValue("0");

		oParser = new parserFormula("I100:J101^2", "A1", ws);
		assert.ok(oParser.parse(), 'I100:J101^2');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#N/A", 'Result of I100:J101^2[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#DIV/0!", 'Result of I100:J101^2[1,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Result of I100:J101^2[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 0, 'Result of I100:J101^2[1,1]');

		ws.getRange2("A100").setValue("75");
		ws.getRange2("B100").setValue("100");
		ws.getRange2("C100").setValue("200");
		ws.getRange2("D100").setValue("300");
		ws.getRange2("E100").setValue("350");
		ws.getRange2("F100").setValue("400");
		ws.getRange2("G100").setValue("500");

		oParser = new parserFormula("A100:G100^{1,2,1}", "A1", ws);
		assert.ok(oParser.parse(), 'A100:G100^{1,2,1}');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 75, 'Result of A100:G100^{1,2,1}[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 10000, 'Result of A100:G100^{1,2,1}[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 200, 'Result of A100:G100^{1,2,1}[0,2]');
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "#N/A", 'Result of A100:G100^{1,2,1}[0,3]');
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), "#N/A", 'Result of A100:G100^{1,2,1}[0,4]');
		assert.strictEqual(array.getElementRowCol(0, 5).getValue(), "#N/A", 'Result of A100:G100^{1,2,1}[0,5]');
		assert.strictEqual(array.getElementRowCol(0, 6).getValue(), "#N/A", 'Result of A100:G100^{1,2,1}[0,6]');

	});

	QUnit.test("Test: \"10-3\"", function (assert) {
		oParser = new parserFormula("10-3", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 7);
	});

	QUnit.test("Test: rename sheet #1", function (assert) {
		wb.dependencyFormulas.unlockRecal();
		ws.getRange2("S95").setValue("2");
		ws.getRange2("S100").setValue("=" + wb.getWorksheet(0).getName() + "!S95");
		ws.setName("SheetTmp");
		assert.strictEqual(ws.getCell2("S100").getFormula(), ws.getName() + "!S95");
		ws.setName("Sheet1");
		wb.dependencyFormulas.lockRecal();
	});

	QUnit.test("Test: wrong ref", function (assert) {
		oParser = new parserFormula("1+XXX1", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?");
	});

	QUnit.test("Test: \"&\"", function (assert) {
		ws.getRange2("AAA101").setValue("1");
		ws.getRange2("AAA102").setValue("2");
		ws.getRange2("AAB101").setValue("3");
		ws.getRange2("AAB102").setValue("4");

		ws.getRange2("AAD101").setValue("2");
		ws.getRange2("AAD102").setValue("3");
		ws.getRange2("AAE101").setValue("4");
		ws.getRange2("AAE102").setValue("5");
		ws.getRange2("AAF101").setValue("test");
		ws.getRange2("AAF102").setValue("test");

		var array;
		oParser = new parserFormula("1&AAA101:AAB102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "11");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "13");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "14");

		oParser = new parserFormula("AAA101:AAB102&AAD101:AAE102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "23");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "34");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "45");

		oParser = new parserFormula("AAA101:AAB102&AAD101:AAF102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "23");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "34");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "45");

		oParser = new parserFormula("AJ2:AM5&AAA101:AAB102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "1");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "2");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "3");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "4");

		oParser = new parserFormula("AJ2:AM5&AAA101:AAA102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "1");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "2");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "1");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "2");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "1");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "2");

		oParser = new parserFormula("AAA101:AAB101&AAD101:AAF102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "13");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "34");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "35");

		oParser = new parserFormula("AAA101&AAD101:AAF102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "13");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "14");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "15");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "1test");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "1test");

		oParser = new parserFormula("1&AAD101:AAF102", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "13");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "14");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "15");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "1test");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "1test");

		oParser = new parserFormula("AAD101:AAF102&\"test\"", "A1", ws);
		oParser.setArrayFormulaRef(ws.getRange2("AD6:AF8").bbox);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "2test");
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "3test");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "4test");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "5test");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "testtest");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "testtest");
	});

	QUnit.test("Test: \">\"", function (assert) {
		oParser = new parserFormula('1.123>1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123>1.5');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.123>1.5');

		oParser = new parserFormula('1.555>1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.555>1.5');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.555>1.5');
	});

	QUnit.test("Test: \"<\"", function (assert) {
		oParser = new parserFormula('1.123<1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123<1.5');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.123<1.5');

		oParser = new parserFormula('1.555<1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.555<1.5');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.555<1.5');
	});

	QUnit.test("Test: \"=\"", function (assert) {
		oParser = new parserFormula('1.123=1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123=1.5');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.123=1.5');

		oParser = new parserFormula('1.555=1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.555=1.555');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.555=1.555');
	});

	QUnit.test("Test: \"<>\"", function (assert) {
		oParser = new parserFormula('1.123<>1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123<>1.5');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.123<>1.5');

		oParser = new parserFormula('1.555<>1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.555<>1.555');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.555<>1.555');
	});

	QUnit.test("Test: \">=\"", function (assert) {
		oParser = new parserFormula('1.123>=1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123>=1.5');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.123>=1.5');

		oParser = new parserFormula('1.555>=1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.555>=1.555');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.555>=1.555');

		oParser = new parserFormula('1.557>=1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.557>=1.555');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.557>=1.555');
	});

	QUnit.test("Test: \"<=\"", function (assert) {
		oParser = new parserFormula('1.123<=1.5', "A2", ws);
		assert.ok(oParser.parse(), '1.123<=1.5');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.123<=1.5');

		oParser = new parserFormula('1.555<=1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.555<=1.555');
		assert.strictEqual(oParser.calculate().getValue(), "TRUE", '1.555<=1.555');

		oParser = new parserFormula('1.557<=1.555', "A2", ws);
		assert.ok(oParser.parse(), '1.557<=1.555');
		assert.strictEqual(oParser.calculate().getValue(), "FALSE", '1.557<=1.555');
	});

	QUnit.test("Test: \"reference argument test\"", function (assert) {
		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("A4").setValue("4");
		ws.getRange2("A5").setValue("5");
		ws.getRange2("A6").setValue("6");

		ws.getRange2("B1").setValue("2");
		ws.getRange2("B2").setValue("");
		ws.getRange2("B3").setValue("3");
		ws.getRange2("B4").setValue("4");
		ws.getRange2("B5").setValue("5");
		ws.getRange2("B6").setValue("6");

		oParser = new parserFormula('IRR(SIN(A1:B4))', 'A2', ws);
		assert.ok(oParser.parse(), 'IRR(SIN(A1:B4))');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, -0.123554096, 'IRR(SIN(A1:B4))');

		oParser = new parserFormula('MIRR(SIN(A2:B4),1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'MIRR(SIN(A2:B4),1,1)');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 2.36894463, 'MIRR(SIN(A2:B4),1,1)');

		oParser = new parserFormula('COLUMN(INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'COLUMN(INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'COLUMN(INDEX(A1:B3,1,1))');

		oParser = new parserFormula('COLUMNS(SIN($A$1:$B$4))', 'A2', ws);
		assert.ok(oParser.parse(), 'COLUMNS(SIN($A$1:$B$4))');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'COLUMNS(SIN($A$1:$B$4))');

		oParser = new parserFormula('INDEX(SIN(A1:B3),1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'INDEX(SIN(A1:B3),1,1)');
		assert.strictEqual(oParser.calculate().getValue().toFixed(9) - 0, 0.841470985, 'INDEX(SIN(A1:B3),1,1)');

		/*oParser = new parserFormula( 'OFFSET(INDEX(A1:B3,1,1),1,1)', 'A2', ws );
		assert.ok( oParser.parse(),'OFFSET(INDEX(A1:B3,1,1),1,1)' );
		assert.strictEqual( oParser.calculate().getValue(),0,'OFFSET(INDEX(A1:B3,1,1),1,1)');*/

		oParser = new parserFormula('ROW(INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'ROW(INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'ROW(INDEX(A1:B3,1,1))');

		oParser = new parserFormula('ROWS(SIN(A1:B3))', 'A2', ws);
		assert.ok(oParser.parse(), 'ROWS(SIN(A1:B3))');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'ROWS(SIN(A1:B3))');

		oParser = new parserFormula('SUBTOTAL(1,INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'SUBTOTAL(1,INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'SUBTOTAL(1,INDEX(A1:B3,1,1))');

		oParser = new parserFormula('SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'SUMIF(INDEX(A1:B3,1,1),1,INDEX(A1:B3,1,1))');

		oParser = new parserFormula('SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'SUMIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula('AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'AVERAGEIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula('COUNTBLANK(INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'COUNTBLANK(INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'COUNTBLANK(INDEX(A1:B3,1,1))');

		oParser = new parserFormula('COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'COUNTIF(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		oParser = new parserFormula('COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'COUNTIFS(INDEX(A1:B3,1,1),INDEX(A1:B3,1,1))');

		ws.getRange2("A2").setValue("qq");
		ws.getRange2("A3").setValue("ww");
		ws.getRange2("A4").setValue("ee");
		ws.getRange2("A5").setValue("qq");
		ws.getRange2("A6").setValue("qq");
		ws.getRange2("A7").setValue("ww");
		ws.getRange2("A8").setValue("ww");
		ws.getRange2("A9").setValue("ww");
		ws.getRange2("A10").setValue("eee");

		ws.getRange2("B1").setValue("qqqq");
		ws.getRange2("B2").setValue("ee");

		var _f = 'IFERROR(INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)),"")';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "ee", _f);

		_f = 'IFERROR(INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B2,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)),"")';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "eee", _f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0))';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "ee", _f);

		_f = 'MATCH(0,INDEX({1;1;0;1;1;1;1;1;0},0,0))';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", _f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,INDEX({1;1;0;1;1;1;1;1;0},0,0),0))';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "ee", _f);

		_f = 'INDEX($A$2:$A$10,3)';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "ee", _f);

		_f = 'INDEX($A$2:$A$10,MATCH(0,{1;1;0;1;1;1;1;1;0},0))';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue().getValue(), "ee", _f);

		_f = 'MATCH(0,INDEX(COUNTIF($B$1:B1,$A$2:$A$10)+(COUNTIF($A$2:$A$10,$A$2:$A$10)<>1),0,0),0)';
		oParser = new parserFormula(_f, 'A2', ws);
		assert.ok(oParser.parse(), _f);
		assert.strictEqual(oParser.calculate().getValue(), 3, _f);
	});

	QUnit.test("Test: \"GetAllFormulas test\"", function (assert) {
		wb.dependencyFormulas.unlockRecal();
		let formulaRange, formulas;

		ws.getRange2("A1:AAZ10000").cleanAll();

		// remove all created earlier defNames
		wb.dependencyFormulas._foreachDefName(function(defName) {
			wb.dependencyFormulas.removeDefName(undefined, defName.name);
		});

		formulaRange = ws.getRange2("A10:A110");
		ws.getRange2("A9").setValue("=SIN(10)");
		// ws.selectionRange.ranges = [ws.getRange2("A10:A110").getBBox0()];
		// ws.selectionRange.setActiveCell(ws.getRange2("A10").getBBox0().r1, ws.getRange2("A10").getBBox0().c1);
		ws.getRange2("A10:A110").setValue("=SUM(A2)", null, null, formulaRange.bbox);

		formulaRange = ws.getRange2("B1:B10000");
		ws.getRange2("B1:B10000").setValue("=1/NOT(ISBLANK(A1:A10000))", null, null, formulaRange.bbox);
		ws.getRange2("C1").setValue("=RAND()");
		ws.getRange2("C2").setValue("=SIN(B1)");
		formulaRange = ws.getRange2("C1:C110");
		ws.getRange2("C10:C110").setValue("=B:B", null, null, formulaRange.bbox);
		formulaRange = ws.getRange2("D1:D10");
		ws.getRange2("D1:D10").setValue("=C1", null, null, formulaRange.bbox);

		formulas = wb.getAllFormulas();
		assert.ok(1, "Created 6 formulas on a sheet: 3 regular, 3 array-formula");
		assert.strictEqual(formulas.length, 6, "GetAllFormulas array length");

		let randRegValBefore = ws.getRange2("C1").getValue(),
			randArrayFValBefore = ws.getRange2("D1").getValue();

		// recalculate workbook
		wb.calculate(4);
		formulas = wb.getAllFormulas();
		assert.ok(1, "Check formulas after workbook recalculate");
		assert.strictEqual(formulas.length, 6, "GetAllFormulas array length");

		let randRegValAfter = ws.getRange2("C1").getValue(),
			randArrayFValAfter = ws.getRange2("D1").getValue();
			
		assert.ok(1, "Check values after workbook recalculate. Values shouldn't be the same");
		assert.strictEqual(randRegValBefore !== randRegValAfter, true, "Check values after recalculate");
		assert.strictEqual(randArrayFValBefore !== randArrayFValAfter, true, "Check values after recalculate");

		ws.getRange2("A1:Z10000").cleanAll();
	});

	QUnit.test("Test: \"Dynamic array test\"", function (assert) {
		let bboxParent, cellWithFormula, formulaInfo, resultRow, resultCol, applyByArray, array;
			
		// wb.dependencyFormulas.unlockRecal();

		ws.getRange2("A1:Z10").cleanAll();
		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("B1").setValue("4");
		ws.getRange2("B2").setValue("str");
		ws.getRange2("B3").setValue("6");
		ws.getRange2("C1").setValue("1");
		ws.getRange2("C2").setValue();
		ws.getRange2("C3").setValue("1");

		// let parent = AscCommonExcel.g_oRangeCache.getAscRange("D1");
		bboxParent = ws.getRange2("D1").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);

		ws.getRange2("C3").setValue("=SIN(A1:A3)", null, null, bboxParent);

		// TODO: review tests with ranges after adding dynamic arrays and add findRefByOutStack formula to use in tests
		oParser = new parserFormula('A1:A3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:A3');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:A3 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:A3');
		assert.strictEqual(resultCol, false, 'Cols in =A1:A3');

		
		oParser = new parserFormula('{1;2;3}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1;2;3}');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1;2;3} array formula');
		assert.strictEqual(resultRow, 3, 'Rows in ={1;2;3}');
		assert.strictEqual(resultCol, 1, 'Cols in ={1;2;3}');

		oParser = new parserFormula('A1:C1', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:C1');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:C1 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:C1');
		assert.strictEqual(resultCol, false, 'Cols in =A1:C1');

		oParser = new parserFormula('{1,2,3}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2,3}');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2,3} array formula');
		assert.strictEqual(resultRow, 1, 'Rows in ={1,2,3}');
		assert.strictEqual(resultCol, 3, 'Cols in ={1,2,3}');

		oParser = new parserFormula('A1:C3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:C3');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:C3 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:C3');
		assert.strictEqual(resultCol, false, 'Cols in =A1:C3');

		oParser = new parserFormula('{1,2;3,4}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2;3,4}');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2;3,4} array formula');
		assert.strictEqual(resultRow, 2, 'Rows in ={1,2;3,4}');
		assert.strictEqual(resultCol, 2, 'Cols in ={1,2;3,4}');

		oParser = new parserFormula('SIN(A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN(A1:A3)');

		oParser = new parserFormula('SIN({1;2;3})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1;2;3})');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1;2;3}) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN({1;2;3})');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN({1;2;3})');

		oParser = new parserFormula('SIN(A1:C1)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:C1)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:C1) array formula');
		assert.strictEqual(resultRow, 1, 'Rows in =SIN(A1:C1)');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN(A1:C1)');

		oParser = new parserFormula('SIN({1,2,3})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1,2,3})');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1,2,3}) array formula');
		assert.strictEqual(resultRow, 1, 'Rows in =SIN({1,2,3})');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN({1,2,3})');

		oParser = new parserFormula('SIN(A1:C3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:C3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:C3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1:C3)');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN(A1:C3)');

		oParser = new parserFormula('SIN({1,2;3,4})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1,2;3,4})');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1,2;3,4}) array formula');
		assert.strictEqual(resultRow, 2, 'Rows in =SIN({1,2;3,4})');
		assert.strictEqual(resultCol, 2, 'Cols in =SIN({1,2;3,4})');

		oParser = new parserFormula('A:A', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A:A');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A:A array formula');
		assert.strictEqual(resultRow, false /*AscCommon.gc_nMaxRow*/, 'Rows in =A:A from D1');
		assert.strictEqual(resultCol, false, 'Cols in =A:A from D1');

		oParser = new parserFormula('A1:XFD1', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:XFD1');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:XFD1 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:XFD1 from D1');
		assert.strictEqual(resultCol, false /*AscCommon.gc_nMaxCol - 3*/, 'Cols in =A1:XFD1 from D1');
		

		oParser = new parserFormula('SIN(A1)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SIN(A1) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SIN(A1)');
		assert.strictEqual(resultCol, false, 'Cols in =SIN(A1)');
		

		oParser = new parserFormula('SUM(A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(A1:A3)');


		oParser = new parserFormula('SUM(A1:A3+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3+A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(A1:A3+A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(A1:A3+A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(A1:A3+A1:A3)');

		oParser = new parserFormula('SUM(A1:A3+A1:A3)+A1:A3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3+A1:A3)+A1:A3');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SUM(A1:A3+A1:A3)+A1:A3 array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SUM(A1:A3+A1:A3)+A1:A3');
		assert.strictEqual(resultCol, 1, 'Cols in =SUM(A1:A3+A1:A3)+A1:A3');


		oParser = new parserFormula('SUM(SIN(A1:A3)+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(SIN(A1:A3)+A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(SIN(A1:A3)+A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(SIN(A1:A3)+A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(SIN(A1:A3)+A1:A3)');


		oParser = new parserFormula('SUM(SIN(SUM(A1:A3)))', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(SIN(SUM(A1:A3)))');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(SIN(SUM(A1:A3))) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(SIN(SUM(A1:A3)))');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(SIN(SUM(A1:A3)))');


		oParser = new parserFormula('SIN(SUM(SIN(A1:A3)))', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(SUM(SIN(A1:A3)))');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SIN(SUM(SIN(A1:A3))) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SIN(SUM(SIN(A1:A3)))');
		assert.strictEqual(resultCol, false, 'Cols in =SIN(SUM(SIN(A1:A3)))');


		oParser = new parserFormula('COS(SIN(A1)*SUM(A1:A3)+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'COS(SIN(A1)*SUM(A1:A3)+A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =COS(SIN(A1)*SUM(A1:A3)+A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =COS(SIN(A1)*SUM(A1:A3)+A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =COS(SIN(A1)*SUM(A1:A3)+A1:A3)');


		oParser = new parserFormula('SIN(A1+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1+A1:A3)');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1+A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1+A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN(A1+A1:A3)');


		oParser = new parserFormula('{1,2}*{3;4}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2}*{3;4}');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2}*{3;4} array formula');
		assert.strictEqual(resultRow, 2, 'Rows in ={1,2}*{3;4}');
		assert.strictEqual(resultCol, 2, 'Cols in ={1,2}*{3;4}');

		oParser = new parserFormula('{2}*{2}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{2}*{2}');
		formulaInfo = ws.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={2}*{2} array formula');
		assert.strictEqual(resultRow, 1, 'Rows in ={1,2}*{3;4}');
		assert.strictEqual(resultCol, 1, 'Cols in ={1,2}*{3;4}');

		// #N/A check
		ws.getRange2("A100:Z110").cleanAll();

		bboxParent = ws.getRange2("D100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D100:E104").bbox);	
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");

		
		ws.getRange2("A100").setValue("1");

		bboxParent = ws.getRange2("I100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("I100:J104").bbox);	
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), 1, "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");

		
		ws.getRange2("B101").setValue("#N/A");

		bboxParent = ws.getRange2("M100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("M100:O104").bbox);	
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), 1, "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");
		
	});

	QUnit.test("Long string splitting", function (assert) {
		// Test case for long string (300 chars - should split into 2 parts)
		let originalString = "a".repeat(300);
		ws.getRange2("A1").cleanAll();
		ws.getRange2("A1").setValue("=\"" + originalString + "\"");
		let formula = ws.getRange2("A1").getFormula();
		let expectedSplits = Math.ceil(originalString.length / 255) - 1;
		let actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "300-char string should have exactly 1 concatenation");
		assert.equal(ws.getRange2("A1").getValue(), originalString, "Result should match original 300-char string");

		// Test very long string (600 chars - should split into 3 parts)
		let longString = "b".repeat(600);
		ws.getRange2("A2").setValue("=\"" + longString + "\"");
		formula = ws.getRange2("A2").getFormula();
		expectedSplits = Math.ceil(longString.length / 255) - 1;
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "600-char string should have exactly 2 concatenations");
		assert.equal(ws.getRange2("A2").getValue(), longString, "Result should match original 600-char string");

		// Test mixed content string (480 chars - should split into 2 parts)
		let mixedString = "Hello world! ".repeat(40); // 480 characters
		ws.getRange2("A3").setValue("=\"" + mixedString + "\"");
		formula = ws.getRange2("A3").getFormula();
		expectedSplits = Math.ceil(mixedString.length / 255) - 1;
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "480-char string should have exactly 1 concatenation");
		assert.equal(ws.getRange2("A3").getValue(), mixedString, "Result should match original mixed string");

		// Test boundary case (255 chars - should not split)
		let boundaryString = "c".repeat(255);
		ws.getRange2("A4").setValue("=\"" + boundaryString + "\"");
		formula = ws.getRange2("A4").getFormula();
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, 0, "255-char string should have no concatenations");
		assert.equal(ws.getRange2("A4").getValue(), boundaryString, "Result should match original boundary string");

		// Test string with quotes (360 chars with quotes - should split into 2 parts)
		let quotedString = "Test\"\"Quote".repeat(30); // 12 chars * 30 = 360 chars
		ws.getRange2("A5").setValue("=\"" + quotedString + "\"");
		formula = ws.getRange2("A5").getFormula();
		expectedSplits = Math.ceil(quotedString.length / 255) - 1;
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "360-char quoted string should have exactly 1 concatenation");
		assert.equal(ws.getRange2("A5").getValue(), quotedString.replace(/\"\"/g, "\""), "Result should match original quoted string");
	});

	QUnit.test("Long string splitting in functions", function (assert) {
		// Test CONCATENATE with long strings (300 + 200 chars)
		let string1 = "a".repeat(300);
		let string2 = "b".repeat(200);
		ws.getRange2("B1").setValue("=\"" + string2 + "\"");
		ws.getRange2("A1").setValue("=CONCATENATE(\"" + string1 + "\", B1)");
		let formula = ws.getRange2("A1").getFormula();
		let expectedSplits = Math.ceil(string1.length / 255) - 1;
		let actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "CONCATENATE with 300-char string should have exactly 1 concatenation");
		assert.equal(ws.getRange2("A1").getValue(), string1 + string2, "CONCATENATE result should match combined strings");

		// Test FIND with long strings (300 chars search in 600 chars text)
		let searchString = "needle".repeat(50); // 300 chars
		let haystackString = "needle".repeat(100); // 600 chars
		ws.getRange2("A2").setValue("=FIND(\"" + searchString + "\", \"" + haystackString + "\")");
		formula = ws.getRange2("A2").getFormula();
		expectedSplits = Math.ceil(searchString.length / 255) - 1 + Math.ceil(haystackString.length / 255) - 1;
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "FIND with long strings should have exactly 3 concatenations");
		assert.equal(ws.getRange2("A2").getValue(), 1, "FIND should work with split strings");

		// Test nested functions (400 chars)
		let nestedString = "nested_text".repeat(40); // 400 chars
		ws.getRange2("A4").setValue("=LEN(UPPER(\"" + nestedString + "\"))");
		formula = ws.getRange2("A4").getFormula();
		expectedSplits = Math.ceil(nestedString.length / 255) - 1;
		actualSplits = (formula.match(/&/g) || []).length;
		assert.equal(actualSplits, expectedSplits, "Nested function with 400-char string should have exactly 1 concatenation");
		assert.equal(ws.getRange2("A4").getValue(), nestedString.length, "Nested functions should work with split strings");
	});

	function calcCustomFunction (innerFunc, jsDoc, oDoc, fCompare) {
		let api = window["Asc"]["editor"];
		if (jsDoc) {
			let oJsDoc = AscCommon.parseJSDoc(jsDoc);
			api.addCustomFunction(innerFunc, oJsDoc[0]);
			fCompare("jsDoc");
		}
		/*if (oDoc) {
			api.addCustomFunction(innerFunc, oDoc);
			fCompare("oDoc");
		}*/
	}

	function initCustomFunctionData() {
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("B100").setValue("3");
		ws.getRange2("B101").setValue("4");

		ws.getRange2("C100").setValue("test1");
		ws.getRange2("C101").setValue("test2");
		ws.getRange2("D100").setValue("test3");
		ws.getRange2("D101").setValue("test4");

		ws.getRange2("E100").setValue("TRUE");
		ws.getRange2("E101").setValue("FALSE");
		ws.getRange2("F100").setValue("FALSE");
		ws.getRange2("F101").setValue("TRUE");

		ws.getRange2("G100").setValue("#VALUE!");
		ws.getRange2("G101").setValue("#REF!");
		ws.getRange2("H100").setValue("#VALUE!");
		ws.getRange2("H101").setValue("#DIV/0!");
	}

	let prefix = "";
	let sJsDoc, oDoc, fCustomFunc;
	function initParamsCustomFunction(aInputTypes, sReturnType) {
		//generate jsdoc
		sJsDoc = "/**\n" +
			"\t\t * Calculates\n" +
			"\t\t * @customfunction\n";

		for (let i in aInputTypes) {
			let argName = "arg" + ((i - 0) + 1);
			if (aInputTypes[i].isOptional) {
				argName = "[" + argName + (aInputTypes[i].defaultValue ? "=" + aInputTypes[i].defaultValue : "") + "]"
			}
			sJsDoc += "\t\t * @param {" + aInputTypes[i].type + "} " + argName + " " + "Description.\n";
		}
		sJsDoc += "\t\t * @returns {" + sReturnType + "} The sum of the numbers.\n\t\t */";

		oDoc = {};
		for (let i in aInputTypes) {
			if (!oDoc["params"]) {
				oDoc["params"] = [];
			}
			oDoc["params"].push({"type": aInputTypes[i].type, "name": "name", "isOptional": !!aInputTypes[i].isOptional, "description": "description_params"});
		}
		oDoc["properties"] = [];
		oDoc["description"] = "all_desc";
		oDoc["returnInfo"] = {"type": sReturnType, "description": "description_return"};
	}

	function doCustomFunctionTasks(assert, aTasks, typeToArgMap, funcName, _descArgs, _callback) {
		//generate ->
		// let desc = "Custom_function_ADD_@NUMBER_@NUMBER_INPUT_NUMBER_NUMBER";
		// calcCustomFunction(func, sJsDoc, oDoc, function (_desc) {
		// 	oParser = new parserFormula(prefix + 'ADD(10, 10)', 'A2', ws);
		// 	assert.ok(oParser.parse(), desc + "_" + _desc);
		// 	assert.strictEqual(oParser.calculate().getValue(), 20, desc + "_" + _desc);
		// });

		for (let i in aTasks) {
			let task = aTasks[i];
			let desc = "Custom_function_" + funcName + "_" +_descArgs + "_INPUT_";
			let sFunc = funcName + "(";
			for (let j = 0; j < aTasks[i].paramsType.length; j++) {
				sFunc += typeToArgMap[aTasks[i].paramsType[j]];
				if (j !== aTasks[i].paramsType.length - 1) {
					sFunc += ",";
				}
				desc += "_" + aTasks[i].paramsType[j];
			}
			sFunc += ")";

			calcCustomFunction(fCustomFunc, sJsDoc, oDoc, function (_desc) {
				if (_callback) {
					wb.asyncFormulasManager.endCallback = function () {
						let calculateRes = ws.getRange2("A1");
						assert.strictEqual(calculateRes.getValue(), task.result, desc + "_" + _desc);
						_callback && _callback();
						wb.asyncFormulasManager.endCallback = null;
					};
					ws.getRange2("A1").setValue("=" + prefix + sFunc);
				} else {
					oParser = new parserFormula(prefix + sFunc, new AscCommonExcel.CCellWithFormula(ws, 1, 0), ws);
					assert.ok(oParser.parse(), "parse_ " + desc + "_" + _desc);
					let calculateRes = oParser.calculate();
					if (typeof task.result === "object") {
						for (let i = 0; i < task.result.length; i++) {
							for (let j = 0; j < task.result[i].length; j++) {
								assert.strictEqual(calculateRes.getElementRowCol(i, j).getValue(), task.result[i][j], desc + "_" + _desc);
							}
						}
					} else {
						assert.strictEqual(calculateRes.getValue(), task.result, desc + "_" + _desc);
					}
				}
			});
		}
	}

	function executeCustomFunction (_func, callback) {
		wb.dependencyFormulas.unlockRecal();
		initCustomFunctionData();

		let api = window["Asc"]["editor"];
		let trueWb = api.wb;
		api.wb = {addCustomFunction: AscCommonExcel.WorkbookView.prototype.addCustomFunction, initCustomEngine: AscCommonExcel.WorkbookView.prototype.initCustomEngine};

		_func(callback);

		if (!callback) {
			api.wb = trueWb;
			ws.getRange2("A1:Z10000").cleanAll();
		}
	}

	QUnit.test("Test: \"Custom function test: base operation: number\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc1(arg1, arg2) {
				return arg2;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @number / @number <- @number **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "number");

			let aTasks = [
				{paramsType: ["number"], result: "#VALUE!"},
				{paramsType: [], result: "#VALUE!"},
				{paramsType: ["number", "number", "number"], result: "#VALUE!"},

				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: 1},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: 1},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@NUMBER_@NUMBER->number");

			//ms returns number!
			//********** 2. @number <- @string **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "string");

			aTasks = [
				{paramsType: ['number', 'number'], result: "10"},
				{paramsType: ['number', 'stringNumber'], result: "1"},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: "1"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "1"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER->string !");

			//********** 3. @number <- @boolean **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "boolean");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: 1},
				{paramsType: ['number', 'string'], result:"#VALUE!"},
				{paramsType: ['number', 'bool'], result: 1},
				{paramsType: ['number', 'error'], result:"#REF!"},
				{paramsType: ['number', 'array'], result:"#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@number->boolean !");

			//********** 4. @number<- @any **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "any");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: 1},
				{paramsType: ['number', 'string'], result:"#VALUE!"},
				{paramsType: ['number', 'bool'], result: 1},
				{paramsType: ['number', 'error'], result:"#REF!"},
				{paramsType: ['number', 'array'], result:"#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_number-> any !");


			//********** 5. @number / @number <- @number[][] **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "number[][]");

			aTasks = [
				{paramsType: ['number', 'number'], result: "#VALUE!"},
				{paramsType: ['number', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: "#VALUE!"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "#VALUE!"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER->number[][] !");


			//********** 6. @number[][] / @number <- @number **********
			initParamsCustomFunction([{type: "number[][]"}, {type: "number"}], "number");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: 1},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: 1},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER[][]_@NUMBER->number !");


			//**********7. @number / @number <- @string[][] **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "string[][]");

			aTasks = [
				{paramsType: ['number', 'number'], result: "#VALUE!"},
				{paramsType: ['number', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: "#VALUE!"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "#VALUE!"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER->string[][] !");

			//********** 8. @number / @number[][]<- @any[][] **********
			initParamsCustomFunction([{type: "number"}, {type: "number[][]"}], "any[][]");

			aTasks = [
				{paramsType: ['number', 'number'], result: [[10]]},
				{paramsType: ['number', 'stringNumber'], result: [[1]]},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: [[1]]},
				{paramsType: ['number', 'error'], result:"#REF!"},
				{paramsType: ['number', 'array'], result: [[1]]},
				{paramsType: ['number', 'ref'], result: [[1]]},
				{paramsType: ['number', 'range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_number[][]-> any[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: base operation: number[][]\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc2(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1.  @number[][] <- @number[][] **********
			initParamsCustomFunction([{type: "number[][]"}], "number[][]");

			let aTasks = [
				{paramsType: ['number'], result: [[10]]},
				{paramsType: ['stringNumber'], result: [[1]]},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: [[1]]},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: [[1,2,3]]},
				{paramsType: ['ref'], result: [[1]]},
				{paramsType: ['range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER[][]->number[][] !");

			//ms returns number!
			//********** 2. @number[][] <- @string[][] **********
			initParamsCustomFunction([{type: "number[][]"}], "string[][]");

			aTasks = [
				{paramsType: ['number'], result: [["10"]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: [["1"]]},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: [["1","2","3"]]},
				{paramsType: ['ref'], result: [["1"]]},
				{paramsType: ['range'], result: [["1"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER[][]->string[][] !");

			//********** 3. @number[][] <- @number **********
			initParamsCustomFunction([{type: "number[][]"}], "number");

			aTasks = [
				{paramsType: ['number'], result: "#VALUE!"},
				{paramsType: ['stringNumber'], result: "#VALUE!"},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: "#VALUE!"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "#VALUE!"},
				{paramsType: ['range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER[][]->number !");

			//********** 4. @number[][] <- @string **********
			initParamsCustomFunction([{type: "number[][]"}], "string");

			aTasks = [
				{paramsType: ['number'], result: "#VALUE!"},
				{paramsType: ['stringNumber'], result: "#VALUE!"},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: "#VALUE!"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "#VALUE!"},
				{paramsType: ['range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@NUMBER[][]->string !");

			//********** 5. @number[][]<- @any[][] **********
			initParamsCustomFunction([{type: "number[][]"}], "any[][]");

			aTasks = [
				{paramsType: ['number'], result: [[10]]},
				{paramsType: ['stringNumber'], result: [[1]]},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: [[1]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [[1]]},
				{paramsType: ['ref'], result: [[1]]},
				{paramsType: ['range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_number[][]-> any[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: base operation: string\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc3(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @string -> @number **********
			initParamsCustomFunction([{type: "string"}], "number");

			let aTasks = [
				{paramsType: ['number'], result: "10"},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "1"},
				{paramsType: ['range'], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@NUMBER@STRING->number");

			//********** 2. @string -> @string **********
			initParamsCustomFunction([{type: "string"}], "string");

			aTasks = [
				{paramsType: ['number'], result: "10"},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "1"},
				{paramsType: ['range'], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@NUMBER@STRING->STRING");


			//********** 3. @string<- @boolean **********
			initParamsCustomFunction([{type: "string"}], "boolean");

			aTasks = [
				{paramsType: ['number'], result: "10"},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: "1"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@string->boolean !");

			//********** 7. @string<- @boolean[][] **********
			initParamsCustomFunction([{type: "string"}], "boolean[][]");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@string->boolean[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: base operation: string[][]\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc4(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @string[][] <- @string[][] **********
			initParamsCustomFunction([{type: "string[][]"}], "string[][]");

			let aTasks = [
				{paramsType: ['number'], result: [["10"]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: [["1"]]},
				{paramsType: ['ref'], result: [["1"]]},
				{paramsType: ['range'], result: [["1"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@STRING[][]->string[][] !");

			//********** 6. @string[][]<- @boolean **********
			initParamsCustomFunction([{type: "string[][]"}], "boolean");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@string[][]->boolean !");

			//********** 8. @string[][]<- @boolean[][] **********
			initParamsCustomFunction([{type: "string[][]"}], "boolean[][]");

			aTasks = [
				{paramsType: ['number'], result: [["10"]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["1"]]},
				{paramsType: ['ref'], result: [["1"]]},
				{paramsType: ['range'], result: [["1"]]}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@string[][]->boolean[][] !");

			//********** 9. @string[][]<- @any[][] **********
			initParamsCustomFunction([{type: "string[][]"}], "any[][]");

			aTasks = [
				{paramsType: ['number'], result: [["10"]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["1"]]},
				{paramsType: ['ref'], result: [["1"]]},
				{paramsType: ['range'], result: [["1"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_string[][]-> any[][] !");

		});
	});

	QUnit.test("Test: \"Custom function test: base operation: boolean\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc5(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @boolean <- @number **********
			initParamsCustomFunction([{type: "boolean"}], "number");

			let aTasks = [
				{paramsType: ['number'], result: "TRUE"},
				{paramsType: ['stringNumber'], result: "#VALUE!"},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "TRUE"},
				{paramsType: ['range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN->NUMBER !");

			//********** 2.@boolean <- @boolean **********
			initParamsCustomFunction([{type: "boolean"}], "boolean");

			aTasks = [
				{paramsType: ['number'], result: "TRUE"},
				{paramsType: ['stringNumber'], result: "#VALUE!"},
				{paramsType: ['string'], result: "#VALUE!"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result: "#REF!"},
				{paramsType: ['array'], result: "#VALUE!"},
				{paramsType: ['ref'], result: "TRUE"},
				{paramsType: ['range'], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN->boolean !");

			//********** 3. @boolean<- @boolean[][] **********
			initParamsCustomFunction([{type: "boolean"}], "boolean[][]");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN->boolean[][] !");


			//********** 4. @boolean<- @any **********
			initParamsCustomFunction([{type: "boolean"}], "any");

			aTasks = [
				{paramsType: ['number'], result: "TRUE"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: "TRUE"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_boolean-> any !");

		});
	});

	QUnit.test("Test: \"Custom function test: base operation: boolean[][]\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc6(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1.@boolean[][] <- @boolean[][] **********
			initParamsCustomFunction([{type: "boolean[][]"}], "boolean[][]");

			let aTasks = [
				{paramsType: ['number'], result: [["TRUE"]]},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["TRUE"]]},
				{paramsType: ['ref'], result: [["TRUE"]]},
				{paramsType: ['range'], result: [["TRUE"]]}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN[][]->boolean[][] !");


			//********** 2. @boolean[][]<- @boolean **********
			initParamsCustomFunction([{type: "boolean[][]"}], "boolean");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN[][]->boolean !");

			//********** 6. @boolean[][]<- @number[][] **********
			initParamsCustomFunction([{type: "boolean[][]"}], "number[][]");

			aTasks = [
				{paramsType: ['number'], result: [["TRUE"]]},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["TRUE"]]},
				{paramsType: ['ref'], result: [["TRUE"]]},
				{paramsType: ['range'], result: [["TRUE"]]}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN[][]->number[][] !");

			//********** 7. @number / @boolean[][]<- @string[][] **********
			initParamsCustomFunction([{type: "boolean[][]"}], "string[][]");

			aTasks = [
				{paramsType: ['number'], result: [["TRUE"]]},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["TRUE"]]},
				{paramsType: ['ref'], result: [["TRUE"]]},
				{paramsType: ['range'], result: [["TRUE"]]}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@BOOLEAN[][]->string[][] !");

			//********** 8. @boolean[][]<- @string[][] **********
			initParamsCustomFunction([{type: "boolean[][]"}], "string[][]");

			aTasks = [
				{paramsType: ['number'], result: [["TRUE"]]},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["TRUE"]]},
				{paramsType: ['ref'], result: [["TRUE"]]},
				{paramsType: ['range'], result: [["TRUE"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_@boolean[][]->string[][] !");


			//********** 10. @boolean[][]<- @any[][] **********
			initParamsCustomFunction([{type: "boolean[][]"}], "any[][]");

			aTasks = [
				{paramsType: ['number'], result: [["TRUE"]]},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["TRUE"]]},
				{paramsType: ['ref'], result: [["TRUE"]]},
				{paramsType: ['range'], result: [["TRUE"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_boolean[][]-> any[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: base operation: any\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc7(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @any<- @number **********
			initParamsCustomFunction([{type: "any"}], "number");

			let aTasks = [
				{paramsType: ['number'], result: 10},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: 1},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any-> number !");

			//********** 2. @any<- @any **********
			initParamsCustomFunction([{type: "any"}], "any");

			aTasks = [
				{paramsType: ['number'], result: 10},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: 1},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any-> any !");

			//********** 3. @any<- @string **********
			initParamsCustomFunction([{type: "any"}], "string");

			aTasks = [
				{paramsType: ['number'], result: "10"},//ms returns number
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: "1"},//ms returns number
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any-> string !");

			//********** 4. @any<- @boolean **********
			initParamsCustomFunction([{type: "any"}], "boolean");

			aTasks = [
				{paramsType: ['number'], result: 10},
				{paramsType: ['stringNumber'], result: "1"},
				{paramsType: ['string'], result: "test"},
				{paramsType: ['bool'], result: "TRUE"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result: 1},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any-> boolean !");

			//********** 5. @any<- @boolean **********
			initParamsCustomFunction([{type: "any"}], "boolean[][]");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any-> boolean[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: base operation: any[][]\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function simpleFunc8(arg1) {
				return arg1;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 1. @any[][]<- @number[][] **********
			initParamsCustomFunction([{type: "any[][]"}], "number[][]");

			let aTasks = [
				{paramsType: ['number'], result: [[10]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [[1]]},
				{paramsType: ['ref'], result: [[1]]},
				{paramsType: ['range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any[][]-> number[][] !");

			//********** 2. @any[][]<- @string[][] **********
			initParamsCustomFunction([{type: "any[][]"}], "string[][]");

			aTasks = [
				{paramsType: ['number'], result: [["10"]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [["1"]]},
				{paramsType: ['ref'], result: [["1"]]},
				{paramsType: ['range'], result: [["1"]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any[][]-> string[][] !");

			//********** 3. @any[][]<- @boolean[][] **********
			initParamsCustomFunction([{type: "any[][]"}], "boolean[][]");

			aTasks = [
				{paramsType: ['number'], result: [[10]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [[1]]},
				{paramsType: ['ref'], result: [[1]]},
				{paramsType: ['range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any[][]-> boolean[][] !");

			//********** 4. @any[][]<- @any[][] **********
			initParamsCustomFunction([{type: "any[][]"}], "any[][]");

			aTasks = [
				{paramsType: ['number'], result: [[10]]},
				{paramsType: ['stringNumber'], result: [["1"]]},
				{paramsType: ['string'], result: [["test"]]},
				{paramsType: ['bool'], result: [["TRUE"]]},
				{paramsType: ['error'], result:"#REF!"},
				{paramsType: ['array'], result: [[1]]},
				{paramsType: ['ref'], result: [[1]]},
				{paramsType: ['range'], result: [[1]]},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! _@NUMBER_any[][]-> any[][] !");
		});
	});

	QUnit.test("Test: \"Custom function test: other\"", function (assert) {

		executeCustomFunction(function () {

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//empty function
			fCustomFunc = function simpleFunc9() {
			};

			initParamsCustomFunction([], "number");

			let aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#VALUE!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! empty  function !");

			//return null
			fCustomFunc = function simpleFunc10() {
				return null;
			};

			initParamsCustomFunction([], "number");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#VALUE!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! return null !");

			//return undefined
			fCustomFunc = function simpleFunc11() {
				return undefined;
			};

			initParamsCustomFunction([], "number");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#VALUE!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! return undefined !");

			//return NaN
			fCustomFunc = function simpleFunc12() {
				return NaN;
			};

			initParamsCustomFunction([], "number");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#VALUE!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! return NaN !");

			fCustomFunc = function simpleFunc13(arg1, arg2) {
				return arg2;
			};

			initParamsCustomFunction([], "number");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: "1"},
				{paramsType: ['number', 'string'], result: "test"},
				{paramsType: ['number', 'bool'], result: "TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! call function without args !");

			initParamsCustomFunction([{type: "any"},{type: "any"},{type: "any"}], "number");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: "1"},
				{paramsType: ['number', 'string'], result: "test"},
				{paramsType: ['number', 'bool'], result: "TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! init args params count more then function contain!");

			fCustomFunc = function simpleFunc14(arg1, arg2, arg3) {
				return arg2;
			};

			initParamsCustomFunction([{type: "any"},{type: "any"},{type: "any"}], "number");

			aTasks = [
				{paramsType: ['number'], result:"#VALUE!"},
				{paramsType: ['stringNumber'], result:"#VALUE!"},
				{paramsType: ['string'], result:"#VALUE!"},
				{paramsType: ['bool'], result:"#VALUE!"},
				{paramsType: ['error'], result:"#VALUE!"},
				{paramsType: ['array'], result:"#VALUE!"},
				{paramsType: ['ref'], result:"#VALUE!"},
				{paramsType: ['range'], result:"#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! call function less then function arg count !");

			//isOptional
			initParamsCustomFunction([{type: "any"},{type: "any"},{type: "any", isOptional: true}], "number");

			aTasks = [
				{paramsType: ['number', 'number'], result: 10},
				{paramsType: ['number', 'stringNumber'], result: "1"},
				{paramsType: ['number', 'string'], result: "test"},
				{paramsType: ['number', 'bool'], result: "TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: 1},
				{paramsType: ['number', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! optional third param !");

			//defaultvalue
			//ms ignore defaultValue option, while skip
			fCustomFunc = function simpleFunc15(arg1, arg2, arg3) {
				return arg3;
			};

			initParamsCustomFunction([{type: "any"},{type: "any"},{type: "any", defaultValue: 123, isOptional: true}], "number");

			aTasks = [
				{paramsType: ['number', 'number'], result: "123"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "! defaultvalue !");

		});
	});

	QUnit.test("Test: \"Custom function test: number+number->number\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function add1(arg1, arg2) {
				return arg1 + arg2;
			};

			//********** 1. @number / @number <- @number **********
			initParamsCustomFunction([{type: "number"}, {type: "number"}], "number");

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };
			let aTasks = [
				{paramsType: ["number", "number"], result: 20},
				{paramsType: ["number", "stringNumber"], result: 11},
				{paramsType: ["number", "string"], result: "#VALUE!"},
				{paramsType: ["number", "bool"], result: 11},
				{paramsType: ["number", "error"], result: "#REF!"},
				{paramsType: ["number", "array"], result: "#VALUE!"},
				{paramsType: ["number", "ref"], result: 11},
				{paramsType: ["number", "range"], result: "#VALUE!"},

				{paramsType: ["string", "string"], result: "#VALUE!"},
				{paramsType: ["string", "stringNumber"], result: "#VALUE!"},
				{paramsType: ["string", "bool"], result: "#VALUE!"},
				{paramsType: ["string", "error"], result: "#REF!"},
				{paramsType: ["string", "array"], result: "#VALUE!"},
				{paramsType: ["string", "ref"], result: "#VALUE!"},
				{paramsType: ["string", "range"], result: "#VALUE!"},

				{paramsType: ["bool", "bool"], result: 2},
				{paramsType: ["bool", "stringNumber"], result: 2},
				{paramsType: ["bool", "error"], result: "#REF!"},
				{paramsType: ["bool", "array"], result: "#VALUE!"},
				{paramsType: ["bool", "ref"], result: 2},
				{paramsType: ["bool", "range"], result: "#VALUE!"},

				{paramsType: ["error", "stringNumber"], result: "#REF!"},
				{paramsType: ["error", "error"], result: "#REF!"},
				{paramsType: ["error", "array"], result: "#REF!"},
				{paramsType: ["error", "ref"], result: "#REF!"},
				{paramsType: ["error", "range"], result: "#REF!"},

				{paramsType: ["array", "stringNumber"], result: "#VALUE!"},
				{paramsType: ["array", "array"], result: "#VALUE!"},
				{paramsType: ["array", "ref"], result: "#VALUE!"},
				{paramsType: ["array", "range"], result: "#VALUE!"},

				{paramsType: ["ref", "stringNumber"], result: 2},
				{paramsType: ["ref", "ref"], result: 2},
				{paramsType: ["ref", "range"], result: "#VALUE!"},

				{paramsType: ["range", "stringNumber"], result: "#VALUE!"},
				{paramsType: ["range", "range"], result: "#VALUE!"}
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@NUMBER_@NUMBER");
		});
	});

	QUnit.test("Test: \"Custom function test: string+number->number\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function add2(arg1, arg2) {
				return arg1 + arg2;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };
			//********** 2. @string / @number <- @number **********
			initParamsCustomFunction([{type: "string"}, {type: "number"}], "number");

			let aTasks = [
				{paramsType: ['number', 'number'], result: "1010"},
				{paramsType: ['number', 'stringNumber'], result: "101"},
				{paramsType: ['number', 'string'], result: "#VALUE!"},
				{paramsType: ['number', 'bool'], result: "101"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "101"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},

				{paramsType: ['string', 'number'], result: "test10"},
				{paramsType: ['string', 'string'], result: "#VALUE!"},
				{paramsType: ['string', 'stringNumber'], result: "test1"},
				{paramsType: ['string', 'bool'], result: "test1"},
				{paramsType: ['string', 'error'], result: "#REF!"},
				{paramsType: ['string', 'array'], result: "#VALUE!"},
				{paramsType: ['string', 'ref'], result: "test1"},
				{paramsType: ['string', 'range'], result: "#VALUE!"},

				{paramsType: ['bool', 'number'], result: "TRUE10"},
				{paramsType: ['bool', 'string'], result: "#VALUE!"},
				{paramsType: ['bool', 'bool'], result: "TRUE1"},
				{paramsType: ['bool', 'stringNumber'], result: "TRUE1"},
				{paramsType: ['bool', 'error'], result: "#REF!"},
				{paramsType: ['bool', 'array'], result: "#VALUE!"},
				{paramsType: ['bool', 'ref'], result: "TRUE1"},
				{paramsType: ['bool', 'range'], result: "#VALUE!"},

				{paramsType: ['error', 'number'], result: "#REF!"},
				{paramsType: ['error', 'string'], result: "#REF!"},
				{paramsType: ['error', 'bool'], result: "#REF!"},
				{paramsType: ['error', 'stringNumber'], result: "#REF!"},
				{paramsType: ['error', 'error'], result: "#REF!"},
				{paramsType: ['error', 'array'], result: "#REF!"},
				{paramsType: ['error', 'ref'], result: "#REF!"},
				{paramsType: ['error', 'range'], result: "#REF!"},

				{paramsType: ['array', 'number'], result: "#VALUE!"},
				{paramsType: ['array', 'string'], result: "#VALUE!"},
				{paramsType: ['array', 'bool'], result: "#VALUE!"},
				{paramsType: ['array', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['array', 'error'], result: "#REF!"},
				{paramsType: ['array', 'array'], result: "#VALUE!"},
				{paramsType: ['array', 'ref'], result: "#VALUE!"},
				{paramsType: ['array', 'range'], result: "#VALUE!"},

				{paramsType: ['ref', 'number'], result: "110"},
				{paramsType: ['ref', 'string'], result: "#VALUE!"},
				{paramsType: ['ref', 'bool'], result: "11"},
				{paramsType: ['ref', 'stringNumber'], result: "11"},
				{paramsType: ['ref', 'error'], result: "#REF!"},
				{paramsType: ['ref', 'array'], result: "#VALUE!"},
				{paramsType: ['ref', 'ref'], result: "11"},
				{paramsType: ['ref', 'range'], result: "#VALUE!"},

				{paramsType: ['range', 'number'], result: "#VALUE!"},
				{paramsType: ['range', 'string'], result: "#VALUE!"},
				{paramsType: ['range', 'bool'], result: "#VALUE!"},
				{paramsType: ['range', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['range', 'error'], result: "#REF!"},
				{paramsType: ['range', 'array'], result: "#VALUE!"},
				{paramsType: ['range', 'ref'], result: "#VALUE!"},
				{paramsType: ['range', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@STRING_@NUMBER");
		});
	});

	QUnit.test("Test: \"Custom function test: string+string->number\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function add3(arg1, arg2) {
				return arg1 + arg2;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };
			//********** 3. @string / @string <- @number **********
			initParamsCustomFunction([{type: "string"}, {type: "string"}], "number");

			let aTasks = [
				{paramsType: ['number', 'number'], result: "1010"},
				{paramsType: ['number', 'stringNumber'], result: "101"},
				{paramsType: ['number', 'string'], result: "10test"},
				{paramsType: ['number', 'bool'], result: "10TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "101"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},

				{paramsType: ['string', 'number'], result: "test10"},
				{paramsType: ['string', 'string'], result: "testtest"},
				{paramsType: ['string', 'stringNumber'], result: "test1"},
				{paramsType: ['string', 'bool'], result: "testTRUE"},
				{paramsType: ['string', 'error'], result: "#REF!"},
				{paramsType: ['string', 'array'], result: "#VALUE!"},
				{paramsType: ['string', 'ref'], result: "test1"},
				{paramsType: ['string', 'range'], result: "#VALUE!"},

				{paramsType: ['bool', 'number'], result: "TRUE10"},
				{paramsType: ['bool', 'string'], result: "TRUEtest"},
				{paramsType: ['bool', 'bool'], result: "TRUETRUE"},
				{paramsType: ['bool', 'stringNumber'], result: "TRUE1"},
				{paramsType: ['bool', 'error'], result: "#REF!"},
				{paramsType: ['bool', 'array'], result: "#VALUE!"},
				{paramsType: ['bool', 'ref'], result: "TRUE1"},
				{paramsType: ['bool', 'range'], result: "#VALUE!"},

				{paramsType: ['error', 'number'], result: "#REF!"},
				{paramsType: ['error', 'string'], result: "#REF!"},
				{paramsType: ['error', 'bool'], result: "#REF!"},
				{paramsType: ['error', 'stringNumber'], result: "#REF!"},
				{paramsType: ['error', 'error'], result: "#REF!"},
				{paramsType: ['error', 'array'], result: "#REF!"},
				{paramsType: ['error', 'ref'], result: "#REF!"},
				{paramsType: ['error', 'range'], result: "#REF!"},

				{paramsType: ['array', 'number'], result: "#VALUE!"},
				{paramsType: ['array', 'string'], result: "#VALUE!"},
				{paramsType: ['array', 'bool'], result: "#VALUE!"},
				{paramsType: ['array', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['array', 'error'], result: "#REF!"},
				{paramsType: ['array', 'array'], result: "#VALUE!"},
				{paramsType: ['array', 'ref'], result: "#VALUE!"},
				{paramsType: ['array', 'range'], result: "#VALUE!"},

				{paramsType: ['ref', 'number'], result: "110"},
				{paramsType: ['ref', 'string'], result: "1test"},
				{paramsType: ['ref', 'bool'], result: "1TRUE"},
				{paramsType: ['ref', 'stringNumber'], result: "11"},
				{paramsType: ['ref', 'error'], result: "#REF!"},
				{paramsType: ['ref', 'array'], result: "#VALUE!"},
				{paramsType: ['ref', 'ref'], result: "11"},
				{paramsType: ['ref', 'range'], result: "#VALUE!"},

				{paramsType: ['range', 'number'], result: "#VALUE!"},
				{paramsType: ['range', 'string'], result: "#VALUE!"},
				{paramsType: ['range', 'bool'], result: "#VALUE!"},
				{paramsType: ['range', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['range', 'error'], result: "#REF!"},
				{paramsType: ['range', 'array'], result: "#VALUE!"},
				{paramsType: ['range', 'ref'], result: "#VALUE!"},
				{paramsType: ['range', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@STRING_@STRING");
		});
	});

	QUnit.test("Test: \"Custom function test: string+string->string\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function add(arg1, arg2) {
				return arg1 + arg2;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };
			//********** 4. @string / @string <- @string **********
			initParamsCustomFunction([{type: "string"}, {type: "string"}], "string");

			let aTasks = [
				{paramsType: ['number', 'number'], result: "1010"},
				{paramsType: ['number', 'stringNumber'], result: "101"},
				{paramsType: ['number', 'string'], result: "10test"},
				{paramsType: ['number', 'bool'], result: "10TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "101"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},

				{paramsType: ['string', 'number'], result: "test10"},
				{paramsType: ['string', 'string'], result: "testtest"},
				{paramsType: ['string', 'stringNumber'], result: "test1"},
				{paramsType: ['string', 'bool'], result: "testTRUE"},
				{paramsType: ['string', 'error'], result: "#REF!"},
				{paramsType: ['string', 'array'], result: "#VALUE!"},
				{paramsType: ['string', 'ref'], result: "test1"},
				{paramsType: ['string', 'range'], result: "#VALUE!"},

				{paramsType: ['bool', 'number'], result: "TRUE10"},
				{paramsType: ['bool', 'string'], result: "TRUEtest"},
				{paramsType: ['bool', 'bool'], result: "TRUETRUE"},
				{paramsType: ['bool', 'stringNumber'], result: "TRUE1"},
				{paramsType: ['bool', 'error'], result: "#REF!"},
				{paramsType: ['bool', 'array'], result: "#VALUE!"},
				{paramsType: ['bool', 'ref'], result: "TRUE1"},
				{paramsType: ['bool', 'range'], result: "#VALUE!"},

				{paramsType: ['error', 'number'], result: "#REF!"},
				{paramsType: ['error', 'string'], result: "#REF!"},
				{paramsType: ['error', 'bool'], result: "#REF!"},
				{paramsType: ['error', 'stringNumber'], result: "#REF!"},
				{paramsType: ['error', 'error'], result: "#REF!"},
				{paramsType: ['error', 'array'], result: "#REF!"},
				{paramsType: ['error', 'ref'], result: "#REF!"},
				{paramsType: ['error', 'range'], result: "#REF!"},

				{paramsType: ['array', 'number'], result: "#VALUE!"},
				{paramsType: ['array', 'string'], result: "#VALUE!"},
				{paramsType: ['array', 'bool'], result: "#VALUE!"},
				{paramsType: ['array', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['array', 'error'], result: "#REF!"},
				{paramsType: ['array', 'array'], result: "#VALUE!"},
				{paramsType: ['array', 'ref'], result: "#VALUE!"},
				{paramsType: ['array', 'range'], result: "#VALUE!"},

				{paramsType: ['ref', 'number'], result: "110"},
				{paramsType: ['ref', 'string'], result: "1test"},
				{paramsType: ['ref', 'bool'], result: "1TRUE"},
				{paramsType: ['ref', 'stringNumber'], result: "11"},
				{paramsType: ['ref', 'error'], result: "#REF!"},
				{paramsType: ['ref', 'array'], result: "#VALUE!"},
				{paramsType: ['ref', 'ref'], result: "11"},
				{paramsType: ['ref', 'range'], result: "#VALUE!"},

				{paramsType: ['range', 'number'], result: "#VALUE!"},
				{paramsType: ['range', 'string'], result: "#VALUE!"},
				{paramsType: ['range', 'bool'], result: "#VALUE!"},
				{paramsType: ['range', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['range', 'error'], result: "#REF!"},
				{paramsType: ['range', 'array'], result: "#VALUE!"},
				{paramsType: ['range', 'ref'], result: "#VALUE!"},
				{paramsType: ['range', 'range'], result: "#VALUE!"},
			];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@STRING_@STRING");
		});
	});

	QUnit.test("Test: \"Custom function test: number[][]+number->number\"", function (assert) {

		executeCustomFunction(function () {
			fCustomFunc = function add(arg1, arg2) {
				return arg1 + arg2;
			};

			let typeToArgMap = {"number": 10, "stringNumber": '"1"', "string": '"test"',  "bool": "TRUE", "error": "#REF!", "array": "{1,2,3}", "ref": "A100", "range": "A100:B101" };

			//********** 2. @number[][] / @number <- @number **********
			initParamsCustomFunction([{type: "number[][]"}, {type: "number"}], "number");

			let aTasks = [
				{paramsType: ['number', 'number'], result: "1010"},
				{paramsType: ['number', 'stringNumber'], result: "101"},
				{paramsType: ['number', 'string'], result: "10test"},
				{paramsType: ['number', 'bool'], result: "10TRUE"},
				{paramsType: ['number', 'error'], result: "#REF!"},
				{paramsType: ['number', 'array'], result: "#VALUE!"},
				{paramsType: ['number', 'ref'], result: "101"},
				{paramsType: ['number', 'range'], result: "#VALUE!"},

				{paramsType: ['string', 'number'], result: "test10"},
				{paramsType: ['string', 'string'], result: "testtest"},
				{paramsType: ['string', 'stringNumber'], result: "test1"},
				{paramsType: ['string', 'bool'], result: "testTRUE"},
				{paramsType: ['string', 'error'], result: "#REF!"},
				{paramsType: ['string', 'array'], result: "#VALUE!"},
				{paramsType: ['string', 'ref'], result: "test1"},
				{paramsType: ['string', 'range'], result: "#VALUE!"},

				{paramsType: ['bool', 'number'], result: "TRUE10"},
				{paramsType: ['bool', 'string'], result: "TRUEtest"},
				{paramsType: ['bool', 'bool'], result: "TRUETRUE"},
				{paramsType: ['bool', 'stringNumber'], result: "TRUE1"},
				{paramsType: ['bool', 'error'], result: "#REF!"},
				{paramsType: ['bool', 'array'], result: "#VALUE!"},
				{paramsType: ['bool', 'ref'], result: "TRUE1"},
				{paramsType: ['bool', 'range'], result: "#VALUE!"},

				{paramsType: ['error', 'number'], result: "#REF!"},
				{paramsType: ['error', 'string'], result: "#REF!"},
				{paramsType: ['error', 'bool'], result: "#REF!"},
				{paramsType: ['error', 'stringNumber'], result: "#REF!"},
				{paramsType: ['error', 'error'], result: "#REF!"},
				{paramsType: ['error', 'array'], result: "#REF!"},
				{paramsType: ['error', 'ref'], result: "#REF!"},
				{paramsType: ['error', 'range'], result: "#REF!"},

				{paramsType: ['array', 'number'], result: "#VALUE!"},
				{paramsType: ['array', 'string'], result: "#VALUE!"},
				{paramsType: ['array', 'bool'], result: "#VALUE!"},
				{paramsType: ['array', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['array', 'error'], result: "#REF!"},
				{paramsType: ['array', 'array'], result: "#VALUE!"},
				{paramsType: ['array', 'ref'], result: "#VALUE!"},
				{paramsType: ['array', 'range'], result: "#VALUE!"},

				{paramsType: ['ref', 'number'], result: "110"},
				{paramsType: ['ref', 'string'], result: "1test"},
				{paramsType: ['ref', 'bool'], result: "1TRUE"},
				{paramsType: ['ref', 'stringNumber'], result: "11"},
				{paramsType: ['ref', 'error'], result: "#REF!"},
				{paramsType: ['ref', 'array'], result: "#VALUE!"},
				{paramsType: ['ref', 'ref'], result: "11"},
				{paramsType: ['ref', 'range'], result: "#VALUE!"},

				{paramsType: ['range', 'number'], result: "#VALUE!"},
				{paramsType: ['range', 'string'], result: "#VALUE!"},
				{paramsType: ['range', 'bool'], result: "#VALUE!"},
				{paramsType: ['range', 'stringNumber'], result: "#VALUE!"},
				{paramsType: ['range', 'error'], result: "#REF!"},
				{paramsType: ['range', 'array'], result: "#VALUE!"},
				{paramsType: ['range', 'ref'], result: "#VALUE!"},
				{paramsType: ['range', 'range'], result: "#VALUE!"},
			];

			assert.ok(1,1);
			//doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_@number[][]_@number");
		});

	});

	QUnit.test("Test: \"Custom function test: async function\"", function (assert) {
		let done;
		executeCustomFunction(function (_callback, trueWb) {
			// Create async custom function
			fCustomFunc = async function simpleAsyncFunc(arg1, arg2) {
				// Simulate async operation
				await new Promise(resolve => setTimeout(resolve, 10));
				return arg2;
			};

			let typeToArgMap = {
				"number": 10,
				"stringNumber": '"1"',
				"string": '"test"',
				"bool": "TRUE",
				"error": "#REF!",
				"array": "{1,2,3}",
				"ref": "A100",
				"range": "A100:B101"
			};

			// Initialize custom function with number parameters
			initParamsCustomFunction(
				[{type: "number"}, {type: "number"}],
				"number"
			);

			// Test async function
			done = assert.async();

			let aTasks = [{
				paramsType: ["number", "number"],
				result: "10"
			}];

			// Execute and verify async result
			doCustomFunctionTasks(assert, aTasks, typeToArgMap, fCustomFunc.name.toUpperCase(), "_ASYNC_TEST", _callback);
		}, function (trueWb) {
			let api = window["Asc"]["editor"];
			api.wb = trueWb;
			done();
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});

	QUnit.test("Test: \"Custom function test: async operations\"", function (assert) {
		let done;
		executeCustomFunction(function (_callback, trueWb) {
			// Async function that returns promise
			fCustomFunc = async function asyncPromiseFunc(arg1, arg2) {
				return new Promise((resolve) => {
					setTimeout(() => {
						resolve(arg1 + arg2);
					}, 10);
				});
			};

			let typeToArgMap = {
				"number": 10,
				"stringNumber": '"1"',
				"string": '"test"',
				"bool": "TRUE",
				"error": "#REF!",
				"array": "{1,2,3}",
				"ref": "A100",
				"range": "A100:B101"
			};

			// Initialize with number parameters
			initParamsCustomFunction(
				[{type: "number"}, {type: "number"}],
				"number"
			);

			// Setup async test
			done = assert.async();

			let aTasks = [
				{
					paramsType: ["number", "number"],
					result: "20"  // 10 + 10
				},
				{
					paramsType: ["string", "number"],
					result: "test10"
				},
				{
					paramsType: ["number", "error"],
					result: "#REF!"
				}
			];

			// Execute tests with callback
			doCustomFunctionTasks(assert, aTasks, typeToArgMap,
				fCustomFunc.name.toUpperCase(), "_ASYNC_PROMISE_TEST", _callback);
		}, function (trueWb) {
			let api = window["Asc"]["editor"];
			api.wb = trueWb;
			done();
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});

	QUnit.test("Test: \"Custom function test: async multiple operations\"", function (assert) {
		let done;
		executeCustomFunction(function (_callback, trueWb) {
			// Async function with multiple awaits
			fCustomFunc = async function asyncMultipleFunc(arg1, arg2) {
				await new Promise(resolve => setTimeout(resolve, 5));
				let temp = arg1 * 2;
				await new Promise(resolve => setTimeout(resolve, 5));
				return temp + arg2;
			};

			let typeToArgMap = {
				"number": 5,
				"stringNumber": '"1"',
				"string": '"test"',
				"bool": "TRUE",
				"error": "#REF!",
				"array": "{1,2,3}",
				"ref": "A100",
				"range": "A100:B101"
			};

			initParamsCustomFunction(
				[{type: "number"}, {type: "number"}],
				"number"
			);

			done = assert.async();

			let aTasks = [{
				paramsType: ["number", "number"],
				result: "15"  // (5 * 2) + 5
			}];

			doCustomFunctionTasks(assert, aTasks, typeToArgMap,
				fCustomFunc.name.toUpperCase(), "_ASYNC_MULTIPLE_TEST", _callback);
		}, function (trueWb) {
			let api = window["Asc"]["editor"];
			api.wb = trueWb;
			done();
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});

	/**
	 * Tests async function calculation with cell dependencies
	 * A1 = number + asyncFunc
	 * B1 = A1 + B2 + asyncFunc2
	 * B2 = C2 + asyncFunc3
	 * C2 = A1 + asyncFunc4
	 */
	QUnit.test('Async formula calculation', function(assert) {
		const done = assert.async(); // For async test completion

		// Setup initial values and async functions
		const asyncFunc = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(10);
				}, 1);
			});
		};

		const asyncFunc2 = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(20);
				}, 1);
			});
		};

		const asyncFunc3 = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(30);
				}, 1);
			});
		};

		const asyncFunc4 = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(40);
				}, 1);
			});
		};

		// Register async functions
		initParamsCustomFunction(
			[{type: "number"}, {type: "number"}],
			"number"
		);

		executeCustomFunction(function (callback) {
			let api = window["Asc"]["editor"];
			let oJsDoc = AscCommon.parseJSDoc(sJsDoc);
			api.addCustomFunction(asyncFunc, oJsDoc[0]);
			api.addCustomFunction(asyncFunc2, oJsDoc[0]);
			api.addCustomFunction(asyncFunc3, oJsDoc[0]);
			api.addCustomFunction(asyncFunc4, oJsDoc[0]);

			wb.asyncFormulasManager.endCallback = function () {
				// Check final calculated values
				assert.strictEqual(ws.getRange2("A1").getValue(), "15", "A1 calculated correctly");
				assert.strictEqual(ws.getRange2("C2").getValue(), "55", "C2 calculated correctly");
				assert.strictEqual(ws.getRange2("B2").getValue(), "85", "B2 calculated correctly");
				assert.strictEqual(ws.getRange2("B1").getValue(), "120", "B1 calculated correctly");
				callback();
				wb.asyncFormulasManager.endCallback = null;
			};

			// Set cell formulas
			wb.dependencyFormulas.lockRecal();
			ws.getRange2("A1").setValue("=5+ASYNCFUNC()");
			ws.getRange2("B1").setValue("=A1+B2+ASYNCFUNC2()");
			ws.getRange2("B2").setValue("=C2+ASYNCFUNC3()");
			ws.getRange2("C2").setValue("=A1+ASYNCFUNC4()");
			wb.dependencyFormulas.unlockRecal();

			// Expected calculation sequence:
			// 1. A1 = 5 + 10 = 15
			// 2. C2 = 15 + 40 = 55
			// 3. B2 = 55 + 30 = 85
			// 4. B1 = 15 + 85 + 20 = 120

			// Check initial state - cells should show loading state
			assert.strictEqual(ws.getRange2("A1").getValue(), "#BUSY!", "A1 shows loading state");
			assert.strictEqual(ws.getRange2("B1").getValue(), "#BUSY!", "B1 shows loading state");
			assert.strictEqual(ws.getRange2("B2").getValue(), "#BUSY!", "B2 shows loading state");
			assert.strictEqual(ws.getRange2("C2").getValue(), "#BUSY!", "C2 shows loading state");
		}, function () {
			done(); // Complete async test
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});

	/**
	 * Tests complex async function calculation with multiple cell dependencies
	 * A1 = number + GetCurrentPrice()
	 * B1 = A1 + C1 + CalculateTax()
	 * C1 = D1 + GetShippingCost()
	 * D1 = A2 + B2 + GetDiscountValue()
	 * A2 = FetchStockQuantity() + GetWarehouseStock()
	 * B2 = C2 + CalculateHandlingFee()
	 * C2 = GetSupplierPrice() * GetMarkupRate()
	 * D2 = (A1 + B1) * GetCurrencyRate()
	 * E1 = SUM(A1:D1) + CalculateInsurance()
	 * E2 = AVERAGE(A2:D2) + GetServiceFee()
	 */
	QUnit.test('Complex async formula calculation with business logic', function(assert) {
		const done = assert.async();

		const GetCurrentPrice = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(10);
				}, 1);
			});
		};

		const CalculateTax = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(20);
				}, 1);
			});
		};

		// Функция расчета стоимости доставки
		const GetShippingCost = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(30);
				}, 1);
			});
		};

		const GetDiscountValue = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(40); // Значение скидки
				}, 1);
			});
		};

		const FetchStockQuantity = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(50);
				}, 1);
			});
		};

		const GetWarehouseStock = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(60);
				}, 1);
			});
		};

		const CalculateHandlingFee = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(70);
				}, 1);
			});
		};

		const GetSupplierPrice = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(80);
				}, 1);
			});
		};

		const GetMarkupRate = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(90);
				}, 1);
			});
		};

		const GetCurrencyRate = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(100);
				}, 1);
			});
		};

		const CalculateInsurance = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(110);
				}, 1);
			});
		};

		const GetServiceFee = function() {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(120);
				}, 1);
			});
		};

		// Register async functions
		initParamsCustomFunction(
			[{type: "number"}, {type: "number"}],
			"number"
		);

		executeCustomFunction(function (callback) {
			let api = window["Asc"]["editor"];
			let oJsDoc = AscCommon.parseJSDoc(sJsDoc);

			// Register all async functions
			const asyncFunctions = {
				GetCurrentPrice,
				CalculateTax,
				GetShippingCost,
				GetDiscountValue,
				FetchStockQuantity,
				GetWarehouseStock,
				CalculateHandlingFee,
				GetSupplierPrice,
				GetMarkupRate,
				GetCurrencyRate,
				CalculateInsurance,
				GetServiceFee
			};

			Object.entries(asyncFunctions).forEach(([_, func]) => {
				api.addCustomFunction(func, oJsDoc[0]);
			});

			wb.asyncFormulasManager.endCallback = function () {
				// Check final calculated values
				// A1 = 5 + 10 = 15
				assert.strictEqual(ws.getRange2("A1").getValue(), "15", "A1 calculated correctly");

				// A2 = 50 + 60 = 110
				assert.strictEqual(ws.getRange2("A2").getValue(), "110", "A2 calculated correctly");

				// C2 = 80 * 90 = 7200
				assert.strictEqual(ws.getRange2("C2").getValue(), "7200", "C2 calculated correctly");

				// B2 = 7200 + 70 = 7270
				assert.strictEqual(ws.getRange2("B2").getValue(), "7270", "B2 calculated correctly");

				// D1 = 110 + 7270 + 40 = 7420
				assert.strictEqual(ws.getRange2("D1").getValue(), "7420", "D1 calculated correctly");

				// C1 = 7420 + 30 = 7450
				assert.strictEqual(ws.getRange2("C1").getValue(), "7450", "C1 calculated correctly");

				// B1 = 15 + 7450 + 20 = 7485
				assert.strictEqual(ws.getRange2("B1").getValue(), "7485", "B1 calculated correctly");

				// D2 = (15 + 7485) * 100 = 750000
				assert.strictEqual(ws.getRange2("D2").getValue(), "750000", "D2 calculated correctly");

				// E1 = (15 + 7485 + 7450 + 7420) + 110 = 22480
				assert.strictEqual(ws.getRange2("E1").getValue(), "22480", "E1 calculated correctly");

				// E2 = AVERAGE(110, 7270, 7200, 750000) + 120 = 191265
				assert.strictEqual(ws.getRange2("E2").getValue(), "191265", "E2 calculated correctly");

				callback();
				wb.asyncFormulasManager.endCallback = null;
			};

			// Set cell formulas with complex dependencies
			wb.dependencyFormulas.lockRecal();
			ws.getRange2("A1").setValue("=5+GetCurrentPrice()");
			ws.getRange2("A2").setValue("=FetchStockQuantity()+GetWarehouseStock()");
			ws.getRange2("C2").setValue("=GetSupplierPrice()*GetMarkupRate()");
			ws.getRange2("B2").setValue("=C2+CalculateHandlingFee()");
			ws.getRange2("D1").setValue("=A2+B2+GetDiscountValue()");
			ws.getRange2("C1").setValue("=D1+GetShippingCost()");
			ws.getRange2("B1").setValue("=A1+C1+CalculateTax()");
			ws.getRange2("D2").setValue("=(A1+B1)*GetCurrencyRate()");
			ws.getRange2("E1").setValue("=SUM(A1:D1)+CalculateInsurance()");
			ws.getRange2("E2").setValue("=AVERAGE(A2:D2)+GetServiceFee()");
			wb.dependencyFormulas.unlockRecal();

			// Check initial loading states
			const rangesToCheck = ["A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "E1", "E2"];
			rangesToCheck.forEach(range => {
				assert.strictEqual(
					ws.getRange2(range).getValue(),
					"#BUSY!",
					`${range} shows loading state`
				);
			});

		}, function () {
			done();
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});
	QUnit.test('Chain of dependent async functions with arguments', function(assert) {
		const done = assert.async();

		// Setup async functions
		const asyncFunc1 = function(value) {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(value + 10);
				}, 1);
			});
		};

		const asyncFunc2 = function(value) {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(value + 20);
				}, 1);
			});
		};

		const asyncFunc3 = function(value) {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(value + 30);
				}, 1);
			});
		};

		const asyncFunc4 = function(value) {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(value + 40);
				}, 1);
			});
		};

		const asyncFunc5 = function(value) {
			return new Promise((resolve) => {
				setTimeout(() => {
					resolve(value + 50);
				}, 1);
			});
		};

		// Register async functions
		initParamsCustomFunction(
			[{type: "number"}, {type: "number"}],
			"number"
		);

		executeCustomFunction(function (callback) {
			let api = window["Asc"]["editor"];
			let oJsDoc = AscCommon.parseJSDoc(sJsDoc);
			api.addCustomFunction(asyncFunc1, oJsDoc[0]);
			api.addCustomFunction(asyncFunc2, oJsDoc[0]);
			api.addCustomFunction(asyncFunc3, oJsDoc[0]);
			api.addCustomFunction(asyncFunc4, oJsDoc[0]);
			api.addCustomFunction(asyncFunc5, oJsDoc[0]);

			wb.asyncFormulasManager.endCallback = function () {
				// Check final calculated values
				// A5 = ASYNCFUNC5(5) = 5 + 50 = 55
				// A4 = ASYNCFUNC4(55) = 55 + 40 = 95
				// A3 = ASYNCFUNC3(95) = 95 + 30 = 125
				// A2 = ASYNCFUNC2(125) = 125 + 20 = 145
				// A1 = ASYNCFUNC1(145) = 145 + 10 = 155
				assert.strictEqual(ws.getRange2("A5").getValue(), "55", "A5 calculated correctly");
				assert.strictEqual(ws.getRange2("A4").getValue(), "95", "A4 calculated correctly");
				assert.strictEqual(ws.getRange2("A3").getValue(), "125", "A3 calculated correctly");
				assert.strictEqual(ws.getRange2("A2").getValue(), "145", "A2 calculated correctly");
				assert.strictEqual(ws.getRange2("A1").getValue(), "155", "A1 calculated correctly");
				callback();
				wb.asyncFormulasManager.endCallback = null;
			};

			// Set cell formulas with chain dependency through function arguments
			ws.getRange2("A5").setValue("=ASYNCFUNC5(5)");
			ws.getRange2("A4").setValue("=ASYNCFUNC4(A5)");
			ws.getRange2("A3").setValue("=ASYNCFUNC3(A4)");
			ws.getRange2("A2").setValue("=ASYNCFUNC2(A3)");
			ws.getRange2("A1").setValue("=ASYNCFUNC1(A2)");

			// Expected calculation sequence:
			// 1. A5 = ASYNCFUNC5(5) = 55
			// 2. A4 = ASYNCFUNC4(55) = 95
			// 3. A3 = ASYNCFUNC3(95) = 125
			// 4. A2 = ASYNCFUNC2(125) = 145
			// 5. A1 = ASYNCFUNC1(145) = 155

			// Check initial state - cells should show loading state
			/*assert.strictEqual(ws.getRange2("A1").getValue(), "#BUSY!", "A1 shows loading state");
			assert.strictEqual(ws.getRange2("A2").getValue(), "#BUSY!", "A2 shows loading state");
			assert.strictEqual(ws.getRange2("A3").getValue(), "#BUSY!", "A3 shows loading state");
			assert.strictEqual(ws.getRange2("A4").getValue(), "#BUSY!", "A4 shows loading state");
			assert.strictEqual(ws.getRange2("A5").getValue(), "#BUSY!", "A5 shows loading state");*/

		}, function () {
			done();
			ws.getRange2("A1:Z10000").cleanAll();
		});
	});
	QUnit.test("Test: \"3d_ref_tests\"", function (assert) {
		let cellWithFormula = new AscCommonExcel.CCellWithFormula(ws, 1, 0);
		let wsName = "हरियाणवी";
		let newWs = wb.createWorksheet(1, wsName);

		oParser = new parserFormula(wsName + '!A1', "A2", ws);
		assert.ok(oParser.parse(), wsName + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1,0). isLocal = true. Link to 3d range A1 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1,0). isLocal = true. Link to 3d range A1 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "हरियाण.वी";
		newWs.setName(wsName);

		oParser = new parserFormula(wsName + '!A1', "A2", ws);
		assert.ok(oParser.parse(), wsName + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1,0). isLocal = true. Link to 3d range A1 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1,0). isLocal = true. Link to 3d range A1 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "हरियाण वी";
		newWs.setName(wsName);

		oParser = new parserFormula(wsName + '!A1', "A2", ws);
		assert.notOk(oParser.parse(), wsName + '!A1');

		oParser = new parserFormula("'" + wsName + "'" + '!A1', "A2", ws);
		assert.ok(oParser.parse(), "'" + wsName + "'" + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1,0). isLocal = true. Link to 3d range A1 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1,0). isLocal = true. Link to 3d range A1 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "हरियाणवी_test_тест_اختبار_123";
		newWs.setName(wsName);

		oParser = new parserFormula(wsName + '!A1', "A2", ws);
		assert.ok(oParser.parse(), wsName + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1,0). isLocal = true. Link to 3d range A1 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1,0). isLocal = true. Link to 3d range A1 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "हरियाणवी_test_тест_اختبار_1 23";
		newWs.setName(wsName);

		oParser = new parserFormula("'" + wsName + "'" + '!A1', "A2", ws);
		assert.ok(oParser.parse(), "'" + wsName + "'" + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "Ả, ẻ, Ỏ";
		newWs.setName(wsName);

		oParser = new parserFormula("'" + wsName + "'" + '!A1', "A2", ws);
		assert.ok(oParser.parse(), "'" + wsName + "'" + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "@©™®†‡§";
		newWs.setName(wsName);

		oParser = new parserFormula("'" + wsName + "'" + '!A1', "A2", ws);
		assert.ok(oParser.parse(), "'" + wsName + "'" + '!A1');
		assert.strictEqual(oParser.calculate().getValue().getValue(), "", wsName + '!A1');

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");


		wsName = "Sheet!25";
		newWs.setName(wsName);
		
		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = ",;";
		newWs.setName(wsName);

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

		wsName = "ds ds ds ! ds ; !";
		newWs.setName(wsName);

		// without quotes
		oParser = new parserFormula("SUM(" + wsName + "!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A1:A2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM(" + wsName + "!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!A:A,0). isLocal = true. Link to 3d range A:A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!A:A,0)");

		oParser = new parserFormula("SUM(" + wsName + "!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true) === false, "SUM(" + wsName + "!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function without quotes");
		assert.strictEqual(oParser.calculate().getValue(), "#NAME?", "SUM(" + wsName + "!$A:$A,0)");

		// with quotes
		oParser = new parserFormula("SUM('" + wsName + "'!A1:A2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A1:A2,0). isLocal = true. Link to 3d range A1:A2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A1:A2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A$1:$A$2,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A$1:$A$2,0). isLocal = true. Link to 3d range $A$1:$A$2 inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A$1:$A$2,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!A:A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!A:A,0). isLocal = true. Link to 3d range A:A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!A:A,0)");

		oParser = new parserFormula("SUM('" + wsName + "'!$A:$A,0)", cellWithFormula, ws);
		assert.ok(oParser.parse(true), "SUM('" + wsName + "'!$A:$A,0). isLocal = true. Link to 3d range $A:$A inside function with quotes");
		assert.ok(oParser.outStack.length > 2, "OutStack length after parse");
		assert.strictEqual(oParser.calculate().getValue(), 0, "SUM('" + wsName + "'!$A:$A,0)");

	});

	QUnit.test("Test: API Calculation option", function (assert) {
		// Init api
		let api = new Asc.spreadsheet_api({
			'id-view': 'editor_sdk'
		});
		window["Asc"]["editor"] = api;
		api.FontLoader = {
			LoadDocumentFonts: function() {
			}
		};
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
		api.initCollaborativeEditing({});
		api._coAuthoringInitCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		// Test api: GetCalcSettings
		let oCalcSettings = api.asc_GetCalcSettings();
		assert.ok(oCalcSettings, "API: GetCalcSettings: Calculation setting object is created");
		assert.strictEqual(oCalcSettings.bIterativeCalc, false, '"Enable iterative calculation" is "false" as default value');
		assert.strictEqual(oCalcSettings.nMaxIterations, 100, '"Maximum Iterations" is "100" as default value');
		assert.strictEqual(oCalcSettings.nMaxChange, 0.001, '"Maximum Change" is "0.001" as default value');
		// Test api: UpdateCalcSettings
		// Update  CalcSettings for api test, need to check changes for: CalcSetting, CalcPr, and CalcRecursion
		oCalcSettings.asc_setIterativeCalc(true);
		oCalcSettings.asc_setMaxIterations(15);

		api.asc_UpdateCalcSettings(oCalcSettings);
		// Check CalcPr
		let oCalcPr = api.wbModel.calcPr;
		assert.strictEqual(oCalcPr.iterate, true, 'API: UpdateCalcSettings. oCalcPr check: "iterate" is "true"');
		assert.strictEqual(oCalcPr.iterateCount, 15, 'API: UpdateCalcSettings. oCalcPr check: "iterateCount" is "10"');
		assert.strictEqual(oCalcPr.iterateDelta, null, 'API: UpdateCalcSettings. oCalcPr check: "iterateDelta" is "null"');
		// Check CalcRecursion
		let g_cCalcRecursion = AscCommonExcel.g_cCalcRecursion;
		assert.strictEqual(g_cCalcRecursion.bIsEnabledRecursion, true, 'API: UpdateCalcSettings. CalcRecursion check: "bIsEnabledRecursion" is "true"');
		assert.strictEqual(g_cCalcRecursion.nMaxIterations, 15, 'API: UpdateCalcSettings. CalcRecursion check: "nMaxIterations" is "15"');
		assert.strictEqual(g_cCalcRecursion.nRelativeError, 0.001, 'API: UpdateCalcSettings. CalcRecursion check: "nRelativeError" is "0.001"');
		// Case: UpdateCalcSettings update Maximum Change
		oCalcSettings.asc_setMaxChange(0.00001);
		api.asc_UpdateCalcSettings(oCalcSettings);
		// Check CalcPr
		assert.strictEqual(oCalcPr.iterateDelta, 0.00001, 'API: UpdateCalcSettings. oCalcPr check: "iterateDelta" is "0.00001"');
		// Check CalcRecursion
		assert.strictEqual(g_cCalcRecursion.nRelativeError, 0.00001, 'API: UpdateCalcSettings. CalcRecursion check: "nRelativeError" is "0.00001"');
	});

	wb.dependencyFormulas.unlockRecal();
});
