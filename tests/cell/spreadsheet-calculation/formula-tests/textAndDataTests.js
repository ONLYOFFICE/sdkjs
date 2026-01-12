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
    const c_msPerDay = AscCommonExcel.c_msPerDay;
	const GetDiffDate360 = AscCommonExcel.GetDiffDate360;
	const fSortAscending = AscCommon.fSortAscending;
	const ParseResult = AscCommonExcel.ParseResult;
	const c_oAscError = Asc.c_oAscError;

	let oParser, wb, ws, sData = AscCommon.getEmpty(), tmp, dif = 1e-9;

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

	QUnit.module('Text and data formulas');

    QUnit.test("Test: \"ARRAYTOTEXT\"", function (assert) {
		// Data for reference link.
		ws.getRange2("B10").setValue("");
		ws.getRange2("B11").setValue();
		ws.getRange2("C10").setValue("");
		ws.getRange2("C11").setValue("1");
		ws.getRange2("C110").setValue("12");
		ws.getRange2("C110").setNumFormat("@");
		ws.getRange2("D10").setValue("19");
		ws.getRange2("D11").setValue("#N/A");
		ws.getRange2("D12").setValue("0");
		ws.getRange2("E10").setValue("str");
		ws.getRange2("E11").setValue("TRUE");
		ws.getRange2("E12").setValue("1");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("19"); // Column1
		ws.getRange2("B601").setValue("1"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("19");
		ws2.getRange2("A2").setValue("0");
		ws2.getRange2("A3").setValue("str");
		ws2.getRange2("A4").setValue("TRUE");
		ws2.getRange2("A5").setValue("1");
		ws2.getRange2("A6").setValue("#N/A");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("19"); // TestName
		ws.getRange2("A202").setValue("0"); // TestName1
		ws.getRange2("A203").setValue("1"); // TestName2
		ws.getRange2("A204").setValue("#N/A"); // TestName3
		ws.getRange2("A206").setValue("12"); // TestNameArea
		ws.getRange2("A207").setValue("test"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("19") // TestName3D
		ws2.getRange2("A12").setValue("0") // TestName3D1
		ws2.getRange2("A13").setValue("1") // TestName3D2
		ws2.getRange2("A14").setValue("#N/A") // TestName3D3
		ws2.getRange2("A16").setValue("19"); // TestNameArea3D
		ws2.getRange2("A17").setValue("text"); // TestNameArea3D

		// Positive cases:

		// array|range && array|range|value
		// Case #1: Area, Number. Area and default format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(B10:C11,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(B10:C11,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ', , , 1', 'Test: Positive case: Area, Number. Area and default format. 2 of 2 arguments used.');
		// Case #2: Area, Number. Area and strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(B10:C11,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(B10:C11,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{,;,1}', 'Test: Positive case: Area, Number. Area and strict format. 2 of 2 arguments used.');
		// Case #3: Area, Number. Vertical only Area and default format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(C10:C11, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(C10:C11, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ', 1', 'Test: Positive case: Area, Number. Vertical only Area and default format. 2 of 2 arguments used.');
		// Case #4: Area, Number. Vertical only Area and strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(C10:C11, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(C10:C11, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{;1}', 'Test: Positive case: Area, Number. Vertical only Area and strict format. 2 of 2 arguments used.');
		// Case #5: Area, Number. Horizontal only area and format is emitted (default). 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(B10:C10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(B10:C10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ', ', 'Test: Positive case: Area, Number. Horizontal only area and format is emitted (default). 1 of 2 arguments used.');
		// Case #6: Area, Number. Horizontal only area and strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(B10:C10,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(B10:C10,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{,}', 'Test: Positive case: Area, Number. Horizontal only area and strict format. 2 of 2 arguments used.');
		// Case #7: Number, Area. Array is number, format is Area. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(12,B10:C11)', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(12,B10:C11) is parsed.');
		let array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12", 'Test: Positive case: Number, Area. Array is number, format is Area. 2 of 2 arguments used.[0,0][0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "12", 'Test: Positive case: Number, Area. Array is number, format is Area. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12", 'Test: Positive case: Number, Area. Array is number, format is Area. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{12}', 'Test: Positive case: Number, Area. Array is number, format is Area. 2 of 2 arguments used.[1,1]');
		// Case #8: String, Area. Array is string, format is Area. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("12",B10:C11)', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("12",B10:C11) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12", 'Test: Positive case: String, Area. Array is string, format is Area. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "12", 'Test: Positive case: String, Area. Array is string, format is Area. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12", 'Test: Positive case: String, Area. Array is string, format is Area. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{"12"}', 'Test: Positive case: String, Area. Array is string, format is Area. 2 of 2 arguments used.[1,1]');
		// Case #9: Reference link, Area. Array is single ref link, format is Area. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(C110,B10:C11)', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(C110,B10:C11) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "12", 'Test: Positive case: Reference link, Area. Array is single ref link, format is Area. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "12", 'Test: Positive case: Reference link, Area. Array is single ref link, format is Area. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12", 'Test: Positive case: Reference link, Area. Array is single ref link, format is Area. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{12}', 'Test: Positive case: Reference link, Area. Array is single ref link, format is Area. 2 of 2 arguments used.[1,1]');
		// Case #10: Reference link, Array. Array is single ref link, format is Array. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(C110,{1,0;FALSE,TRUE})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(C110,{1,0;FALSE,TRUE}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "{12}", 'Test: Positive case: Reference link, Array. Array is single ref link, format is Array. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "12", 'Test: Positive case: Reference link, Array. Array is single ref link, format is Array. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12", 'Test: Positive case: Reference link, Array. Array is single ref link, format is Array. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{12}', 'Test: Positive case: Reference link, Array. Array is single ref link, format is Array. 2 of 2 arguments used.[1,1]');
		// Case #11: Area(2). Array (,,, 1) and format (default, default, default, strict) are area. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(B10:C11,B10:C11)', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(B10:C11,B10:C11) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), ", , , 1", 'Test: Positive case: Area(2). Array (,,, 1) and format (default, default, default, strict) are area. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), ", , , 1", 'Test: Positive case: Area(2). Array (,,, 1) and format (default, default, default, strict) are area. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), ", , , 1", 'Test: Positive case: Area(2). Array (,,, 1) and format (default, default, default, strict) are area. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{,;,1}', 'Test: Positive case: Area(2). Array (,,, 1) and format (default, default, default, strict) are area. 2 of 2 arguments used.[1,1]');
		// Case #12: Area. Area with all primitive types and format is emitted. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '19, str, #N/A, TRUE', 'Test: Positive case: Area. Area with all primitive types and format is emitted. 1 of 2 arguments used.');
		// Case #13: Area, Number. Area with all primitive types and format is strict. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{19,\"str\";#N/A,TRUE}', 'Test: Positive case: Area, Number. Area with all primitive types and format is strict. 2 of 2 arguments used.');
		// Case #14: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11,{0,1,2})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11,{0,1,2}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "19, str, #N/A, TRUE", 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), '{19,"str";#N/A,TRUE}', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[0,2]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[1,1]');
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), '', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is vertical array with mostly correct types. 2 of 2 arguments used.[1,2]');
		// Case #15: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11,{0,1,2;FALSE,"1","TRUE"})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11,{0,1,2;FALSE,"1","TRUE"}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), '19, str, #N/A, TRUE', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), '{19,"str";#N/A,TRUE}', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), '#VALUE!', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[0,2]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '19, str, #N/A, TRUE', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{19,"str";#N/A,TRUE}', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[1,1]');
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), '#VALUE!', 'Test: Positive case: Area, Array. Array is area with all primitive types, format is array with mostly correct types. 2 of 2 arguments used.[1,2]');
		// value && array|range|value
		// Case #16: Number. Array is number. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Number. Array is number. 1 of 2 arguments used.');
		// Case #17: Number(2). Array is number, strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(12,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(12,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{12}', 'Test: Positive case: Number(2). Array is number, strict format. 2 of 2 arguments used.');
		// Case #18: String. Array is  number in string. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("12")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("12") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: String. Array is  number in string. 1 of 2 arguments used.');
		// Case #19: String(2). Array is number in string, strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("12",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("12",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"12"}', 'Test: Positive case: String(2). Array is number in string, strict format. 2 of 2 arguments used.');
		// Case #20: Boolean. Array is boolean (TRUE). 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Array is boolean (TRUE). 1 of 2 arguments used.');
		// Case #21: Boolean. Array is boolean (FALSE). 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Array is boolean (FALSE). 1 of 2 arguments used.');
		// Case #22: Boolean, Number. Array is boolean (TRUE), strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TRUE, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TRUE, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{TRUE}', 'Test: Positive case: Boolean, Number. Array is boolean (TRUE), strict format. 2 of 2 arguments used.');
		// Case #23: Boolean, Number. Array is boolean (FALSE), strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(FALSE,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(FALSE,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{FALSE}', 'Test: Positive case: Boolean, Number. Array is boolean (FALSE), strict format. 2 of 2 arguments used.');
		// Case #24: String. Array is text string. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: String. Array is text string. 1 of 2 arguments used.');
		// Case #25: String,Number. Array is text string, strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: String,Number. Array is text string, strict format. 2 of 2 arguments used.');
		// Case #26: String(2). Format number in string (default). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str", "0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str", "0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: String(2). Format number in string (default). 2 of 2 arguments used.');
		// Case #27: String(2). Format number in string (strict). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str", "1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str", "1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: String(2). Format number in string (strict). 2 of 2 arguments used.');
		// Case #28: String, Boolean. Format is boolean (FALSE). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str",FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str",FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: String, Boolean. Format is boolean (FALSE). 2 of 2 arguments used.');
		// Case #29: String, Boolean. Format is boolean (TRUE). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: String, Boolean. Format is boolean (TRUE). 2 of 2 arguments used.');
		// Case #30: Reference link. Array is ref link. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: Reference link. Array is ref link. 1 of 2 arguments used.');
		// Case #31: Reference link, String. Array is ref link,  format number is string (default). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10, "0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10, "0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: Reference link, String. Array is ref link,  format number is string (default). 2 of 2 arguments used.');
		// Case #32: Reference link, String. Array is ref link,  format number is string (strict). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10, "1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10, "1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: Reference link, String. Array is ref link,  format number is string (strict). 2 of 2 arguments used.');
		// Case #33: Reference link, Boolean. Array is ref link,  format is boolean (FALSE). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str', 'Test: Positive case: Reference link, Boolean. Array is ref link,  format is boolean (FALSE). 2 of 2 arguments used.');
		// Case #34: Reference link, Boolean. Array is ref link,  format is boolean (TRUE). 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10, TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10, TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: Reference link, Boolean. Array is ref link,  format is boolean (TRUE). 2 of 2 arguments used.');
		// Case #35: Reference link, Number. Array is ref link,  strict format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str"}', 'Test: Positive case: Reference link, Number. Array is ref link,  strict format. 2 of 2 arguments used.');
		// Case #36: Name(2). Array  (single cell) and format (default) are name. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestName, TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName, TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '19', 'Test: Positive case: Name(2). Array  (single cell) and format (default) are name. 2 of 2 arguments used.');
		// Case #37: Name(2). Array  (area) and format (strict) are name. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestNameArea , TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestNameArea , TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{12;"test"}', 'Test: Positive case: Name(2). Array  (area) and format (strict) are name. 2 of 2 arguments used.');
		// Case #38: Name3D(2). Array (single cell) and format (default) are name3D. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestName3D, TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName3D, TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '19', 'Test: Positive case: Name3D(2). Array (single cell) and format (default) are name3D. 2 of 2 arguments used.');
		// Case #39: Name3D(2). Array (area) and format (strict) are name3D. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestNameArea3D , TestName3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestNameArea3D , TestName3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{19;"text"}', 'Test: Positive case: Name3D(2). Array (area) and format (strict) are name3D. 2 of 2 arguments used.');
		// Case #40: Ref3D. Array (single cell) and format (default) are Ref3D. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(Sheet2!A1 , Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(Sheet2!A1 , Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "19", 'Test: Positive case: Ref3D. Array (single cell) and format (default) are Ref3D. 2 of 2 arguments used.');
		// Case #41: Area3D. Array (area) and format (strict) are Ref3D. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(Sheet2!A3:A4 , Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(Sheet2!A3:A4 , Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{"str";TRUE}', 'Test: Positive case: Area3D. Array (area) and format (strict) are Ref3D. 2 of 2 arguments used.');
		// Case #42: Table. Array (single cell) and format (default) are Table. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(Table1[Column1], Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(Table1[Column1], Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{19}', 'Test: Positive case: Table. Array (single cell) and format (default) are Table. 2 of 2 arguments used.');
		// Case #43: Formula. ARRAYTOTEXT part of another formula
		oParser = new parserFormula('UPPER(ARRAYTOTEXT({"test1","test2"}))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula UPPER(ARRAYTOTEXT({"test1","test2"})) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TEST1, TEST2', 'Test: Positive case: Formula. ARRAYTOTEXT part of another formula');
		// Case #44: Formula. Array is Date. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "45658", 'Test: Positive case: Formula. Array is Date. 1 of 2 arguments used.');
		// Case #45: Formula. Array is Time. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0.5", 'Test: Positive case: Formula. Array is Time. 1 of 2 arguments used.');

		// Negative cases:

		// Case #1: Area, Number. Area with correct data and incorrect format. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, Number. Area with correct data and incorrect format. 2 of 2 arguments used.');
		// Case #2: Number, Array. Array is number, format is array with incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(12,{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(12,{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Number, Array. Array is number, format is array with incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: Number, Array. Array is number, format is array with incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "12", 'Test: Negative case: Number, Array. Array is number, format is array with incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{12}', 'Test: Negative case: Number, Array. Array is number, format is array with incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #3: String, Array. Array is string, format is array with incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("12",{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("12",{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: String, Array. Array is string, format is array with incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: String, Array. Array is string, format is array with incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '12', 'Test: Negative case: String, Array. Array is string, format is array with incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{"12"}', 'Test: Negative case: String, Array. Array is string, format is array with incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #4: String, Array. Array is string (1s), format is array with incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("1s",{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("1s",{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: String, Array. Array is string (1s), format is array with incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: String, Array. Array is string (1s), format is array with incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '1s', 'Test: Negative case: String, Array. Array is string (1s), format is array with incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{"1s"}', 'Test: Negative case: String, Array. Array is string (1s), format is array with incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #5: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(FALSE,{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(FALSE,{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is  array with incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is  array with incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'FALSE', 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is  array with incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{FALSE}', 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is  array with incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #6: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TRUE,{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TRUE,{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'TRUE', 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{TRUE}', 'Test: Negative case: Boolean, Array. Array is boolean (FALSE), format is array with incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #7: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(D10:E11, D10:E12)', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(D10:E11, D10:E12) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#N/A", 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{19,"str";#N/A,TRUE}', 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[1,1]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "19, str, #N/A, TRUE", 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[2,0]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), '{19,"str";#N/A,TRUE}', 'Test: Negative case: Area(2). Array is area with all primitive types, format is array with mostly incorrect types. 2 of 2 arguments used.[2,1]');
		// Case #8: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT({"str",12,TRUE},{3,#NUM!;0,1})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT({"str",12,TRUE},{3,#NUM!;0,1}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "str, 12, TRUE", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{"str",12,TRUE}', 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[1,1]');
		// Case #9: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT({"str",12,TRUE},{3,#NUM!;"0",TRUE})', "A2", ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT({"str",12,TRUE},{3,#NUM!;"0",TRUE}) is parsed.');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#NUM!", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "str, 12, TRUE", 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[1,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '{"str",12,TRUE}', 'Test: Negative case: Array(2). Array - "str", 12, TRUE, format is array with mostly incorrect types. 2 of 2 arguments used.[1,1]');
		// value && array|range|value
		// Case #10: Empty, Number. Array argument is empty. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty, Number. Array argument is empty. 2 of 2 arguments used.');
		// Case #11: Empty, Reference link. Array argument is empty, format ref link with empty cell. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(,B10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(,B10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty, Reference link. Array argument is empty, format ref link with empty cell. 2 of 2 arguments used.');
		// Case #12: Empty, Reference link. Array argument is empty, format ref link with empty cell. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(,B11)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(,B11) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty, Reference link. Array argument is empty, format ref link with empty cell. 2 of 2 arguments used.');
		// Case #13: Error. Array argument with error #N/A. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Array argument with error #N/A. 1 of 2 arguments used.');
		// Case #14: String(2). Format  incorrect string. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str", "1s")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str", "1s") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Format  incorrect string. 2 of 2 arguments used.');
		// Case #15: Reference link, String. Format  incorrect string. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(E10, "1s")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(E10, "1s") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link, String. Format  incorrect string. 2 of 2 arguments used.');
		// Case #16: String, Number. Format incorrect number. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT("str", 12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT("str", 12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Format incorrect number. 2 of 2 arguments used.');
		// Case #17: Error(2). Array and format are error. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(#NUM!, #N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(#NUM!, #N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Error(2). Array and format are error. 2 of 2 arguments used.');
		// Case #18: Name. Array with error. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name. Array with error. 1 of 2 arguments used.');
		// Case #19: Name(2). Format with error. 2 of 2 arguments used.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#31
		/*oParser = new parserFormula('ARRAYTOTEXT(TestName, TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName, TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name(2). Format with error. 2 of 2 arguments used.');
		*/// Case #20: Name3D. Array with error. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(TestName3D3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName3D3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name3D. Array with error. 1 of 2 arguments used.');
		// Case #21: Name3D(2). Format with error. 2 of 2 arguments used.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#31
		/*oParser = new parserFormula('ARRAYTOTEXT(TestName3D, TestName3D3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(TestName3D, TestName3D3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name3D(2). Format with error. 2 of 2 arguments used.');
		*/// Case #22: Ref3D. Array with error. 1 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(Sheet2!A6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(Sheet2!A6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. Array with error. 1 of 2 arguments used.');
		// Case #23: Ref3D(2). Format with error. 2 of 2 arguments used.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#31
		/*oParser = new parserFormula('ARRAYTOTEXT(Sheet2!A1, Sheet2!A6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(Sheet2!A1, Sheet2!A6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D(2). Format with error. 2 of 2 arguments used.');
*/
		// Bounded cases:

		// Case #1: Number(2). Minimum accepted value. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT({1}, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT({1}, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "1", 'Test: Bounded case: Number(2). Minimum accepted value. 2 of 2 arguments used.');
		// Case #2: Number(2). Maximum accepted value. 2 of 2 arguments used.
		oParser = new parserFormula('ARRAYTOTEXT(SEQUENCE(6774),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ARRAYTOTEXT(SEQUENCE(6774),1) is parsed.');
		// assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: Number(2). Maximum accepted value. 2 of 2 arguments used.');
		
	});

	QUnit.test("Test: \"ASC\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("ＡＢＣＤ");
		ws.getRange2("A101").setValue("Ａ1２3");
		ws.getRange2("A102").setValue("１２３４５");
		ws.getRange2("B102").setValue("1");
		ws.getRange2("A104").setValue("TRUE");
		ws.getRange2("A105").setValue("12/12/2000");
		ws.getRange2("A106").setValue("#N/A");
		ws.getRange2("B106").setValue("#N/A");
		ws.getRange2("A108").setValue("");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("ＡＢＣＤABCD"); // Column1
		ws.getRange2("B601").setValue("#N/A"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("ｔｅＳｔ");
		ws2.getRange2("A2").setValue("#N/A");
		ws2.getRange2("B2").setValue("#NUM!");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("！＠＃＄％"); // TestName
		ws.getRange2("A208").setValue("#N/A"); // TestNameArea2
		ws.getRange2("B208").setValue("#NUM!"); // TestNameArea2
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("！＠＃＄％") // TestName3D
		ws2.getRange2("A12").setValue("#N/A") // TestName3D1

		// Positive cases:

		// Case #1: String. Default input
		oParser = new parserFormula('ASC("ｔｅＳｔ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ｔｅＳｔ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'teSt', 'Test: Positive case: String. Default input');
		// Case #2: String. String with japanese hieroglyphs
		oParser = new parserFormula('ASC("デジタル")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("デジタル") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'デジタル', 'Test: Positive case: String. String with japanese hieroglyphs');
		// Case #3: String. String with uknown symbol
		oParser = new parserFormula('ASC("￯")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("￯") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: String. String with uknown symbol');
		// Case #4: String. Basic test: full-width Latin characters to half-width. The ASC function converts full-width (double-byte) characters to half-width (single-byte) characters.
		oParser = new parserFormula('ASC("ＡＢＣＤ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣＤ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCD', 'Test: Positive case: String. Basic test: full-width Latin characters to half-width. The ASC function converts full-width (double-byte) characters to half-width (single-byte) characters.');
		// Case #5: String. Full-width Katakana to half-width.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#32
		/*oParser = new parserFormula('ASC("ｶﾀｶﾅ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ｶﾀｶﾅ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ｶﾀｶﾅ', 'Test: Positive case: String. Full-width Katakana to half-width.');*/
		// Case #6: String. Already half-width characters remain unchanged.
		oParser = new parserFormula('ASC("ABC")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ABC") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC', 'Test: Positive case: String. Already half-width characters remain unchanged.');
		// Case #7: String. Non-Latin/Katakana characters remain unchanged (Cyrillic).
		oParser = new parserFormula('ASC("Тест")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("Тест") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Тест', 'Test: Positive case: String. Non-Latin/Katakana characters remain unchanged (Cyrillic).');
		// Case #8: String. Mixed full-width and half-width alphanumeric characters.
		oParser = new parserFormula('ASC("Ａ1２3")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("Ａ1２3") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A123', 'Test: Positive case: String. Mixed full-width and half-width alphanumeric characters.');
		// Case #9: String. Full-width Latin and Japanese characters with full-width space.
		oParser = new parserFormula('ASC("ＡＢＣＤ　テスト")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣＤ　テスト") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCD　テスト', 'Test: Positive case: String. Full-width Latin and Japanese characters with full-width space.');
		// Case #10: String. Full-width numeric characters to half-width.
		oParser = new parserFormula('ASC("１２３４５")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("１２３４５") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12345', 'Test: Positive case: String. Full-width numeric characters to half-width.');
		// Case #11: String. Full-width special characters to half-width.
		oParser = new parserFormula('ASC("！＠＃＄％")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("！＠＃＄％") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#$%', 'Test: Positive case: String. Full-width special characters to half-width.');
		// Case #12: String. Mixed full-width and half-width Latin with hyphen.
		oParser = new parserFormula('ASC("ＡＢＣＤ-ABCD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣＤ-ABCD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCD-ABCD', 'Test: Positive case: String. Mixed full-width and half-width Latin with hyphen.');
		// Case #13: String. Empty string returns empty string.
		oParser = new parserFormula('ASC("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: String. Empty string returns empty string.');
		// Case #14: Number. Number automatically converted to string.
		oParser = new parserFormula('ASC(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: Number. Number automatically converted to string.');
		// Case #15: Formula. Nested formula returning full-width characters.
		oParser = new parserFormula('ASC(CONCATENATE("Ａ","Ｂ","Ｃ"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(CONCATENATE("Ａ","Ｂ","Ｃ")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC', 'Test: Positive case: Formula. Nested formula returning full-width characters.');
		// Case #16: Reference link. Reference to cell containing full-width characters.
		oParser = new parserFormula('ASC(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCD', 'Test: Positive case: Reference link. Reference to cell containing full-width characters.');
		// Case #17: Area. Single-cell range containing full-width characters.
		oParser = new parserFormula('ASC(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A123', 'Test: Positive case: Area. Single-cell range containing full-width characters.');
		// Case #18: Array. Array with single element containing full-width characters.
		oParser = new parserFormula('ASC({"ＡＢＣＤ"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC({"ＡＢＣＤ"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 'ABCD', 'Test: Positive case: Array. Array with single element containing full-width characters.');
		// Case #19: Name. Named range containing full-width characters.
		oParser = new parserFormula('ASC(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#$%', 'Test: Positive case: Name. Named range containing full-width characters.');
		// Case #20: Name3D. 3D named range containing full-width characters.
		oParser = new parserFormula('ASC(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#$%', 'Test: Positive case: Name3D. 3D named range containing full-width characters.');
		// Case #21: Ref3D. 3D reference to cell containing full-width characters.
		oParser = new parserFormula('ASC(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'teSt', 'Test: Positive case: Ref3D. 3D reference to cell containing full-width characters.');
		// Case #22: Area3D. 3D single-cell range containing full-width characters.
		oParser = new parserFormula('ASC(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'teSt', 'Test: Positive case: Area3D. 3D single-cell range containing full-width characters.');
		// Case #23: Table. Table structured reference to cell with full-width characters.
		oParser = new parserFormula('ASC(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCDABCD', 'Test: Positive case: Table. Table structured reference to cell with full-width characters.');
		// Case #24: String. Japanese characters mixed with half-width numerals.
		oParser = new parserFormula('ASC("テスト123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("テスト123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'テスト123', 'Test: Positive case: String. Japanese characters mixed with half-width numerals.');
		// Case #25: String. Full-width space character to half-width.
		oParser = new parserFormula('ASC("　")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("　") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '　', 'Test: Positive case: String. Full-width space character to half-width.');
		// Case #26: Formula. CHAR function to generate full-width character (я).
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#32
		/*oParser = new parserFormula('ASC(CHAR(255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(CHAR(255)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'я', 'Test: Positive case: Formula. CHAR function to generate full-width character (я).');*/
		// Case #27: Formula. ASC inside another formula (SUM and LEN).
		oParser = new parserFormula('SUM(LEN(ASC("ＡＢＣＤ")),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(LEN(ASC("ＡＢＣＤ")),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Formula. ASC inside another formula (SUM and LEN).');
		// Case #28: String. Concatenated full-width strings.
		oParser = new parserFormula('ASC("ＡＢＣ" & "ＤＥＦ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣ" & "ＤＥＦ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCDEF', 'Test: Positive case: String. Concatenated full-width strings.');
		// Case #29: String. String concatenated with CHAR function.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#32
		/*oParser = new parserFormula('ASC("ＡＢＣ" & CHAR(153))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣ" & CHAR(153)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC™', 'Test: Positive case: String. String concatenated with CHAR function.');*/
		// Case #30: String. Emoji with full-width character.
		oParser = new parserFormula('ASC("🀄︎Ａ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("🀄︎Ａ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '🀄︎A', 'Test: Positive case: String. Emoji with full-width character.');
		// Case #31: Boolean. Boolean TRUE converted to string "TRUE".
		oParser = new parserFormula('ASC(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean TRUE converted to string "TRUE".');
		// Case #32: Boolean. Boolean FALSE converted to string "FALSE".
		oParser = new parserFormula('ASC(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Boolean. Boolean FALSE converted to string "FALSE".');
		// Case #33: Date. Date value converted to string representation.
		oParser = new parserFormula('ASC(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658', 'Test: Positive case: Date. Date value converted to string representation.');
		// Case #34: Time. Time value converted to string representation.
		oParser = new parserFormula('ASC(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Time. Time value converted to string representation.');
		// Case #35: Formula. IF formula returning full-width characters.
		oParser = new parserFormula('ASC(IF(TRUE,"ＡＢＣＤ","1234"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(IF(TRUE,"ＡＢＣＤ","1234")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCD', 'Test: Positive case: Formula. IF formula returning full-width characters.');
		// Case #36: Area. Multi-cell range (creates array of results).
		oParser = new parserFormula('ASC(A102:B102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A102:B102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12345', 'Test: Positive case: Area. Multi-cell range (creates array of results).');
		// Case #37: Reference link. Reference to cell containing number (converts to string).
		oParser = new parserFormula('ASC(B102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(B102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Reference link. Reference to cell containing number (converts to string).');
		// Case #38: Reference link. Reference to cell containing boolean (converts to string).
		oParser = new parserFormula('ASC(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Reference link. Reference to cell containing boolean (converts to string).');
		// Case #39: Reference link. Reference to cell containing date (converts to string).
		oParser = new parserFormula('ASC(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '36872', 'Test: Positive case: Reference link. Reference to cell containing date (converts to string).');
		// Case #40: String. Long full-width string with spaces.
		oParser = new parserFormula('ASC("ＡＢＣ　ＤＥＦ　ＧＨＩ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣ　ＤＥＦ　ＧＨＩ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC　DEF　GHI', 'Test: Positive case: String. Long full-width string with spaces.');
		// Case #41: Array. Multi-element array with full-width characters.
		oParser = new parserFormula('ASC({"ＡＢＣ";"ＤＥＦ"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC({"ＡＢＣ";"ＤＥＦ"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 'ABC', 'Test: Positive case: Array. Multi-element array with full-width characters.');
		// Case #42: String. Full alphabet in full-width lowercase.
		oParser = new parserFormula('ASC("ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abcdefghijklmnopqrstuvwxyz', 'Test: Positive case: String. Full alphabet in full-width lowercase.');
		// Case #43: String. Full alphabet in full-width uppercase.
		oParser = new parserFormula('ASC("ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'Test: Positive case: String. Full alphabet in full-width uppercase.');

		// Negative cases:

		// Case #1: Error. Error value #N/A propagates through function.
		oParser = new parserFormula('ASC(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error value #N/A propagates through function.');
		// Case #2: Error. Error value #DIV/0! propagates through function.
		oParser = new parserFormula('ASC(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Error. Error value #DIV/0! propagates through function.');
		// Case #3: Error. Error value #VALUE! propagates through function.
		oParser = new parserFormula('ASC(#VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(#VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error. Error value #VALUE! propagates through function.');
		// Case #4: Reference link. Reference to cell containing error value.
		oParser = new parserFormula('ASC(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link. Reference to cell containing error value.');
		// Case #5: Area. Multi-cell range with error value.
		oParser = new parserFormula('ASC(A106:B106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A106:B106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area. Multi-cell range with error value.');
		// Case #6: Name. Named range with error value.
		oParser = new parserFormula('ASC(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name. Named range with error value.');
		// Case #7: Empty. Reference to empty cell returns empty string.
		oParser = new parserFormula('ASC(A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(A108) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Empty. Reference to empty cell returns empty string.');
		// Case #8: Formula. Formula resulting in #N/A error propagates error.
		oParser = new parserFormula('ASC(MATCH("xyz",{"a","b","c"},0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(MATCH("xyz",{"a","b","c"},0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Formula. Formula resulting in #N/A error propagates error.');
		// Case #9: String. Very long string (256 characters) may exceed limits.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#33
		/*oParser = new parserFormula('ASC(REPT("Ａ",256))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(REPT("Ａ",256)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Very long string (256 characters) may exceed limits.');*/
		// Case #10: Name3D. 3D named range containing error value.
		oParser = new parserFormula('ASC(TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name3D. 3D named range containing error value.');
		// Case #11: Ref3D. 3D reference to cell containing error value.
		oParser = new parserFormula('ASC(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. 3D reference to cell containing error value.');
		// Case #12: Area3D. 3D multi-cell range with error value.
		oParser = new parserFormula('ASC(Sheet2!A2:B2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Sheet2!A2:B2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area3D. 3D multi-cell range with error value.');
		// Case #13: Table. Table column with error value.
		oParser = new parserFormula('ASC(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Table. Table column with error value.');
		// Case #14: Formula. Nested formula resulting in error propagates error.
		oParser = new parserFormula('ASC(IF(FALSE,"Ａ",1/0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(IF(FALSE,"Ａ",1/0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Formula. Nested formula resulting in error propagates error.');
		// Case #15: String. String exceeding Excel\'s character limit.
		// Different result with MS TODO: Need to fix: Blocked by https://nct.onlyoffice.com/Products/Files/DocEditor.aspx?fileid=366936 Bugs Row#33
		/*oParser = new parserFormula('ASC(REPT("Ａ",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(REPT("Ａ",32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String exceeding Excel\'s character limit.');*/
		// Case #16: Array. Array containing error value.
		oParser = new parserFormula('ASC({#N/A})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC({#N/A}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#N/A', 'Test: Negative case: Array. Array containing error value.');
		// Case #17: Array. Array containing #DIV/0! error.
		oParser = new parserFormula('ASC({#DIV/0!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC({#DIV/0}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#DIV/0!', 'Test: Negative case: Array. Array containing #DIV/0! error.');
		// Case #18: Reference link. Reference to undefined name.
		oParser = new parserFormula('ASC(XYZ)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(XYZ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Reference link. Reference to undefined name.');

		// Bounded cases:

		// Case #1: String. Empty string - minimum valid input.
		oParser = new parserFormula('ASC("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: String. Empty string - minimum valid input.');
		// Case #2: String. Single full-width character.
		oParser = new parserFormula('ASC("Ａ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("Ａ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Bounded case: String. Single full-width character.');
		// Case #3: String. String with 255 characters (typical Excel limit).
		oParser = new parserFormula('ASC(REPT("Ａ",255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC(REPT("Ａ",255)) is parsed.');
		let str = "A"
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(255), 'Test: Bounded case: String. String with 255 characters (typical Excel limit).');
		// Case #4: String. Only full-width space character.
		oParser = new parserFormula('ASC("　")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("　") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '　', 'Test: Bounded case: String. Only full-width space character.');
		// Case #5: String. Only full-width special characters.
		oParser = new parserFormula('ASC("！＠＃＄％＾＆＊（）＿＋")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula ASC("！＠＃＄％＾＆＊（）＿＋") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#$%^&*()_+', 'Test: Bounded case: String. Only full-width special characters.');

	});

	QUnit.test("Test: \"CHAR\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("65");
		ws.getRange2("A101").setValue("33");
		ws.getRange2("A102").setValue("test");
		ws.getRange2("A103").setValue("256");
		ws.getRange2("A104").setValue("");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("33"); // Column1
		ws.getRange2("B601").setValue("test"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("65");
		ws2.getRange2("A2").setValue("test");
		ws2.getRange2("A3").setValue("0");
		ws2.getRange2("A4").setValue("-1");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("65"); // TestName
		ws.getRange2("A206").setValue("0"); // TestNameArea
		ws.getRange2("A207").setValue("-1"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("90") // TestName3D
		ws2.getRange2("A12").setValue("test") // TestName3D1

		// Positive cases:

		// Case #1: Number. Char 95 converts to a
		oParser = new parserFormula('CHAR(97)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(97) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Number. Char 95 converts to a');
		// Case #2: Formula. Nested formula CODE char converts to A
		oParser = new parserFormula('CHAR(CODE("A"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(CODE("A")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested formula CODE char converts to A');
		// Case #3: Number. Basic valid input: integer (65 = "A").
		oParser = new parserFormula('CHAR(65)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(65) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Number. Basic valid input: integer (65 = "A").');
		// Case #4: Number. Valid number input (97 = "a").
		oParser = new parserFormula('CHAR(97)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(97) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Number. Valid number input (97 = "a").');
		// Case #5: String. String convertible to number (66 = "B").
		oParser = new parserFormula('CHAR("66")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("66") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Positive case: String. String convertible to number (66 = "B").');
		// Case #6: String. String with decimal is truncated to 65 ("A").
		oParser = new parserFormula('CHAR("65.7")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("65.7") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: String. String with decimal is truncated to 65 ("A").');
		// Case #7: Formula. Nested formula (32+15 = 47 = "/").
		oParser = new parserFormula('CHAR(SUM(32,15))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(SUM(32,15)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '/', 'Test: Positive case: Formula. Nested formula (32+15 = 47 = "/").');
		// Case #8: Formula. Nested formula with explicit truncation.
		oParser = new parserFormula('CHAR(INT(65.7))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(INT(65.7)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested formula with explicit truncation.');
		// Case #9: Formula. Nested formula with rounding.
		oParser = new parserFormula('CHAR(ROUND(65.4,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(ROUND(65.4,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested formula with rounding.');
		// Case #10: Reference link. Reference to cell with valid number.
		oParser = new parserFormula('CHAR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Reference link. Reference to cell with valid number.');
		// Case #11: Area. Single-cell range with valid number.
		oParser = new parserFormula('CHAR(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!', 'Test: Positive case: Area. Single-cell range with valid number.');
		// Case #12: Array. Array with single element (77 = "M").
		oParser = new parserFormula('CHAR({77})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR({77}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 'M', 'Test: Positive case: Array. Array with single element (77 = "M").');
		// Case #13: Name. Named range with valid number.
		oParser = new parserFormula('CHAR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Name. Named range with valid number.');
		// Case #14: Name3D. 3D named range with valid number.
		oParser = new parserFormula('CHAR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Z', 'Test: Positive case: Name3D. 3D named range with valid number.');
		// Case #15: Ref3D. 3D reference to cell with valid number.
		oParser = new parserFormula('CHAR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Ref3D. 3D reference to cell with valid number.');
		// Case #16: Area3D. 3D single-cell range with valid number.
		oParser = new parserFormula('CHAR(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Area3D. 3D single-cell range with valid number.');
		// Case #17: Table. Table structured reference with valid number.
		oParser = new parserFormula('CHAR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!', 'Test: Positive case: Table. Table structured reference with valid number.');
		// Case #18: Formula. CHAR inside another formula (CODE returns 65).
		oParser = new parserFormula('SUM(CODE(CHAR(65)),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(CODE(CHAR(65)),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 66, 'Test: Positive case: Formula. CHAR inside another formula (CODE returns 65).');
		// Case #19: Formula. CHAR inside string concatenation.
		oParser = new parserFormula('CONCATENATE("Char: ",CHAR(64))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Char: ",CHAR(64)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Char: @', 'Test: Positive case: Formula. CHAR inside string concatenation.');
		// Case #20: Number. ; character (ASCII 59).
		oParser = new parserFormula('CHAR(59)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(59) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ';', 'Test: Positive case: Number. ; character (ASCII 59).');
		// Case #21: Number. < character (ASCII 60).
		oParser = new parserFormula('CHAR(60)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(60) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '<', 'Test: Positive case: Number. < character (ASCII 60).');
		// Case #22: Number. ? character (ASCII 63).
		oParser = new parserFormula('CHAR(63)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(63) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Positive case: Number. ? character (ASCII 63).');
		// Case #23: Number. > character (ASCII 62).
		oParser = new parserFormula('CHAR(62)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(62) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '>', 'Test: Positive case: Number. > character (ASCII 62).');
		// Case #24: Number. Minimum valid control character (ASCII 1).
		oParser = new parserFormula('CHAR(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Number. Minimum valid control character (ASCII 1).');
		// Case #25: Formula. Nested IF formula returning valid number.
		oParser = new parserFormula('CHAR(IF(TRUE,65,66))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(IF(TRUE,65,66)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested IF formula returning valid number.');
		// Case #26: Formula. Nested CHOOSE formula returning valid number.
		oParser = new parserFormula('CHAR(CHOOSE(2,64,65,66))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(CHOOSE(2,64,65,66)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested CHOOSE formula returning valid number.');
		// Case #27: Number. Digit character "0" (ASCII 48).
		oParser = new parserFormula('CHAR(48)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(48) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0", 'Test: Positive case: Number. Digit character "0" (ASCII 48).');
		// Case #28: Number. Tilde character "~" (ASCII 126).
		oParser = new parserFormula('CHAR(126)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(126) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '~', 'Test: Positive case: Number. Tilde character "~" (ASCII 126).');
		// Case #29: Number. Hash character "#" (ASCII 35).
		oParser = new parserFormula('CHAR(35)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(35) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#', 'Test: Positive case: Number. Hash character "#" (ASCII 35).');
		// Case #30: Boolean. Boolean TRUE converts to 1 (valid).
		oParser = new parserFormula('CHAR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Boolean. Boolean TRUE converts to 1 (valid).');
		// Case #31: Formula. Date function returning valid number (15).
		oParser = new parserFormula('CHAR(DAY("2023-01-15"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(DAY("2023-01-15")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Formula. Date function returning valid number (15).');
		// Case #32: Formula. Month from static date (returns 1).
		oParser = new parserFormula('CHAR(MONTH(DATE(2023,1,1)))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(MONTH(DATE(2023,1,1))) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Formula. Month from static date (returns 1).');
		// Case #33: Formula. String length as character code (5).
		oParser = new parserFormula('CHAR(LEN("ABCDE"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(LEN("ABCDE")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Formula. String length as character code (5).');
		// Case #34: Formula. Decimal number is truncated to 65 ("A").
		oParser = new parserFormula('CHAR(65.7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(65.7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Decimal number is truncated to 65 ("A").');
		// Case #35: Formula. Explicit truncation of decimal number.
		oParser = new parserFormula('CHAR(TRUNC(66.9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TRUNC(66.9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Positive case: Formula. Explicit truncation of decimal number.');
		// Case #36: Formula. ROW function returning valid number (1).
		oParser = new parserFormula('CHAR(ROW(A1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(ROW(A1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Formula. ROW function returning valid number (1).');
		// Case #37: Formula. COLUMN function returning valid number (5).
		oParser = new parserFormula('CHAR(COLUMN(E1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(COLUMN(E1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Formula. COLUMN function returning valid number (5).');
		// Case #38: Formula. CHAR inside MID function (returns "A").
		oParser = new parserFormula('MID(CHAR(65)&CHAR(66),1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula MID(CHAR(65)&CHAR(66),1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. CHAR inside MID function (returns "A").');
		// Case #39: Array. Array with two identical valid elements.
		oParser = new parserFormula('CHAR({65;65})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR({65;65}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 'A', 'Test: Positive case: Array. Array with two identical valid elements.');
		// Case #40: Formula. ABS function returning valid number (65).
		oParser = new parserFormula('CHAR(ABS(-65))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(ABS(-65)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. ABS function returning valid number (65).');
		// Case #41: Formula. MOD function bringing number into range (65).
		oParser = new parserFormula('CHAR(MOD(321,256))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(MOD(321,256)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. MOD function bringing number into range (65).');
		// Case #42: String. Maximum ASCII value as string.
		oParser = new parserFormula('CHAR("255")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("255") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ÿ', 'Test: Positive case: String. Maximum ASCII value as string.');

		// Negative cases:

		// Case #1: Number. Value below minimum (1) returns #VALUE!.
		oParser = new parserFormula('CHAR(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(0) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Value below minimum (1) returns #VALUE!.');
		// Case #2: Number. Value above maximum (255) returns #VALUE!.
		oParser = new parserFormula('CHAR(256)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(256) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Value above maximum (255) returns #VALUE!.');
		// Case #3: Number. Negative number returns #VALUE!.
		oParser = new parserFormula('CHAR(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(-1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Negative number returns #VALUE!.');
		// Case #4: String. Non-numeric string returns #VALUE!.
		oParser = new parserFormula('CHAR("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!.');
		// Case #5: String. String convertible to number below range returns #VALUE!.
		oParser = new parserFormula('CHAR("0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("0") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String convertible to number below range returns #VALUE!.');
		// Case #6: String. String convertible to number above range returns #VALUE!.
		oParser = new parserFormula('CHAR("256")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("256") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String convertible to number above range returns #VALUE!.');
		// Case #7: Error. Error input propagates #N/A error.
		oParser = new parserFormula('CHAR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input propagates #N/A error.');
		// Case #8: Error. Division by zero error propagates #DIV/0!.
		oParser = new parserFormula('CHAR(1/0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(1/0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Error. Division by zero error propagates #DIV/0!.');
		// Case #9: Area. Multi-cell range returns #VALUE! error.
		oParser = new parserFormula('CHAR(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range returns #VALUE! error.');
		// Case #10: Array. Array with invalid element returns #VALUE!.
		oParser = new parserFormula('CHAR({0;-1})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR({0;-1}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#VALUE!', 'Test: Negative case: Array. Array with invalid element returns #VALUE!.');
		// Case #11: Array. Array with invalid element returns #VALUE!.
		oParser = new parserFormula('CHAR({256;257})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR({256;257}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#VALUE!', 'Test: Negative case: Array. Array with invalid element returns #VALUE!.');
		// Case #12: Empty. Reference to empty cell returns #VALUE!.
		oParser = new parserFormula('CHAR(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!.');
		// Case #13: String. Empty string returns #VALUE!.
		oParser = new parserFormula('CHAR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!.');
		// Case #14: Reference link. Reference to cell with text returns #VALUE!.
		oParser = new parserFormula('CHAR(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with text returns #VALUE!.');
		// Case #15: Reference link. Reference to cell with number outside range returns #VALUE!.
		oParser = new parserFormula('CHAR(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to cell with number outside range returns #VALUE!.');
		// Case #16: Name. Named range with multiple cells returns #VALUE!.
		oParser = new parserFormula('CHAR(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with multiple cells returns #VALUE!.');
		// Case #17: Area3D. 3D range with multiple cells returns #VALUE!.
		oParser = new parserFormula('CHAR(Sheet2!A2:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Sheet2!A2:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D range with multiple cells returns #VALUE!.');
		// Case #18: Table. Table column with text returns #VALUE!.
		oParser = new parserFormula('CHAR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!.');
		// Case #19: Formula. Formula resulting in #NUM! error propagates error.
		oParser = new parserFormula('CHAR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates error.');
		// Case #20: Formula. Parent formula with CHAR error propagates #VALUE!.
		oParser = new parserFormula('SUM(CHAR(-1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(CHAR(-1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Parent formula with CHAR error propagates #VALUE!.');
		// Case #21: Boolean. Boolean FALSE (0) returns #VALUE!.
		oParser = new parserFormula('CHAR(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(FALSE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean FALSE (0) returns #VALUE!.');
		// Case #22: Formula. Formula resulting in number outside range returns #VALUE!.
		oParser = new parserFormula('CHAR(INT(1E+10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(INT(1E+10)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Formula resulting in number outside range returns #VALUE!.');
		// Case #23: Formula. Average in array exceeds range.
		oParser = new parserFormula('CHAR(AVERAGE({300,400,500}))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(AVERAGE({300,400,500})) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Average in array exceeds range.');
		// Case #24: Ref3D. Ref link from another sheet to cell with negative number returns #VALUE!.
		oParser = new parserFormula('CHAR(Sheet2!A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(Sheet2!A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. Ref link from another sheet to cell with negative number returns #VALUE!.');
		// Case #25: Name3D. Nemed Ref link from another sheet to cell with text returns #VALUE!.
		oParser = new parserFormula('CHAR(TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. Nemed Ref link from another sheet to cell with text returns #VALUE!.');

		// Bounded cases:

		// Case #1: Number. Minimum valid value (1).
		oParser = new parserFormula('CHAR(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: Number. Minimum valid value (1).');
		// Case #2: Number. Maximum valid value (255).
		oParser = new parserFormula('CHAR(255)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CHAR(255) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ÿ', 'Test: Bounded case: Number. Maximum valid value (255).');

        // Need to fix: diff results from MS(especially boreder and negative cases when we expect the error)

		testArrayFormula2(assert, "CHAR", 1, 1);
	});

	QUnit.test("Test: \"CLEAN\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)");
		ws.getRange2("A101").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)");
		ws.getRange2("A102").setValue("");
		ws.getRange2("A103").setValue("#N/A");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)"); // Column1
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)");
		ws2.getRange2("A2").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)");
		ws2.getRange2("A3").setValue("#N/A");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)"); // TestName
		ws.getRange2("A202").setValue("#N/A"); // TestName1
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("=CHAR(9)&\"Monthly report\"&CHAR(10)") // TestName3D
		ws2.getRange2("A12").setValue("#N/A") // TestName3D1

		// Positive cases:

		// Case #1: Reference link. Removes the nonprintable characters CHAR(9) and CHAR(10) from the text string in cell A100.
		oParser = new parserFormula('CLEAN(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Reference link. Removes the nonprintable characters CHAR(9) and CHAR(10) from the text string in cell A100.');
		// Case #2: String. Removes nonprintable characters (e.g., CHAR(9)).
		oParser = new parserFormula('CLEAN("Hello"&CHAR(9)&"World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN("Hello"&CHAR(9)&"World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HelloWorld', 'Test: Positive case: String. Removes nonprintable characters (e.g., CHAR(9)).');
		// Case #3: Number. Number input is treated as text. Returns the same number as text.
		oParser = new parserFormula('CLEAN(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: Number. Number input is treated as text. Returns the same number as text.');
		// Case #4: Empty. Empty string input. Returns an empty string.
		oParser = new parserFormula('CLEAN("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Empty. Empty string input. Returns an empty string.');
		// Case #5: Formula. Handles nested formulas with nonprintable characters.
		oParser = new parserFormula('CLEAN(TEXT(123,"0000")&CHAR(10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TEXT(123,"0000")&CHAR(10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0123', 'Test: Positive case: Formula. Handles nested formulas with nonprintable characters.');
		// Case #6: Reference link. Refers to a cell with nonprintable characters.
		oParser = new parserFormula('CLEAN(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Reference link. Refers to a cell with nonprintable characters.');
		// Case #7: Area. Single-cell range input.
		// Different result with MS
		oParser = new parserFormula('CLEAN(A100:A101)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A100:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Monthly report', 'Test: Positive case: Area. Single-cell range input.');
		// Case #8: Array. Array input with nonprintable characters.
		// Different result with MS
		oParser = new parserFormula('CLEAN({"Hello    World"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN({"Hello    World"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello    World', 'Test: Positive case: Array. Array input with nonprintable characters.');
		// Case #9: Name. Named range with nonprintable characters.
		oParser = new parserFormula('CLEAN(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Name. Named range with nonprintable characters.');
		// Case #10: Name3D. 3D named range with nonprintable characters.
		oParser = new parserFormula('CLEAN(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Name3D. 3D named range with nonprintable characters.');
		// Case #11: Ref3D. 3D reference input.
		oParser = new parserFormula('CLEAN(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Ref3D. 3D reference input.');
		// Case #12: Area3D. 3D multi-cell range input.
		oParser = new parserFormula('CLEAN(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Area3D. 3D multi-cell range input.');
		// Case #13: Table. Table structured reference with nonprintable characters.
		oParser = new parserFormula('CLEAN(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Table. Table structured reference with nonprintable characters.');
		// Case #14: String. Removes nonprintable characters but keeps spaces.
		oParser = new parserFormula('CLEAN(" Leading space")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(" Leading space") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ' Leading space', 'Test: Positive case: String. Removes nonprintable characters but keeps spaces.');
		// Case #15: Formula. Tests CLEAN within an IF formula.
		oParser = new parserFormula('CLEAN(IF(2>1,"Good"&CHAR(10),"Bad"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(IF(2>1,"Good"&CHAR(10),"Bad")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Good', 'Test: Positive case: Formula. Tests CLEAN within an IF formula.');
		// Case #16: Date. Date value concatenated with nonprintable character.
		oParser = new parserFormula('CLEAN(DATE(2022,1,1)&CHAR(9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(DATE(2022,1,1)&CHAR(9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '44562', 'Test: Positive case: Date. Date value concatenated with nonprintable character.');
		// Case #17: Time. Time value concatenated with nonprintable character.
		oParser = new parserFormula('CLEAN(TIME(12,0,0)&CHAR(10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TIME(12,0,0)&CHAR(10)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Time. Time value concatenated with nonprintable character.');
		// Case #18: String. Does not remove non-ASCII characters (e.g., CHAR(129)).
		// Different result with MS
		oParser = new parserFormula('CLEAN("abc"&CHAR(129))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN("abc"&CHAR(129)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String. Does not remove non-ASCII characters (e.g., CHAR(129)).');
		// Case #19: Number. Negative number treated as text. Returns the same number as text.
		oParser = new parserFormula('CLEAN(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1', 'Test: Positive case: Number. Negative number treated as text. Returns the same number as text.');
		// Case #20: Array. Array with Boolean.
		oParser = new parserFormula('CLEAN({FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN({FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Array. Array with Boolean.');
		// Case #21: Boolean. Boolean values.
		oParser = new parserFormula('CLEAN(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Boolean. Boolean values.');
		// Case #22: Empty. Empty cell reference returns an empty string.
		oParser = new parserFormula('CLEAN(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Empty. Empty cell reference returns an empty string.');
		// Case #23: Area. Whole column
		oParser = new parserFormula('CLEAN(A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area. Whole column');
		// Case #24: Area. Whole row
		oParser = new parserFormula('CLEAN(100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(100:100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Monthly report', 'Test: Positive case: Area. Whole row');

		// Negative cases:

		// Case #1: Error. Propagates #N/A error.
		oParser = new parserFormula('CLEAN(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error.');
		// Case #2: Ref3D. 3D ref to error value.
		oParser = new parserFormula('CLEAN(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. 3D ref to error value.');
		// Case #3: Reference link. Ref to error value.
		oParser = new parserFormula('CLEAN(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link. Ref to error value.');
		// Case #4: Name. Name to error value.
		oParser = new parserFormula('CLEAN(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name. Name to error value.');
		// Case #5: Name3D. Name3D to error value.
		oParser = new parserFormula('CLEAN(TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name3D. Name3D to error value.');

		// Bounded cases:

		// Case #1: String. Maximum string length input. Removes nonprintable characters.
		oParser = new parserFormula('CLEAN(REPT("A",32767)&CHAR(10))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(REPT("A",32767)&CHAR(10)) is parsed.');
		let longStr = "A".repeat(32767);
		assert.strictEqual(oParser.calculate().getValue(), longStr, 'Test: Bounded case: String. Maximum string length input. Removes nonprintable characters.');
		// Case #2: String. Single space input. Treated as printable character. Returns same value.
		oParser = new parserFormula('CLEAN(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ' ', 'Test: Bounded case: String. Single space input. Treated as printable character. Returns same value.');
		// Case #3: String. Empty string input. Returns an empty string.
		oParser = new parserFormula('CLEAN("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CLEAN("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: String. Empty string input. Returns an empty string.');


		testArrayFormula(assert, "CLEAN");
	});

	QUnit.test("Test: \"CODE\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("a");
		ws.getRange2("A101").setValue("A");
		ws.getRange2("A102").setValue("!");
		ws.getRange2("A103").setValue("O");
		ws.getRange2("A104").setValue("TRUE");
		ws.getRange2("A105").setValue("");
		ws.getRange2("A106").setValue("#N/A");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("A"); // Column1
		ws.getRange2("B601").setValue(""); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("A");
		ws2.getRange2("A2").setValue("a");
		ws2.getRange2("A3").setValue("!");
		ws2.getRange2("A4").setValue("");
		ws2.getRange2("A5").setValue("#N/A");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("A"); // TestName
		ws.getRange2("A206").setValue("A"); // TestNameArea
		ws.getRange2("A207").setValue("a"); // TestNameArea
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("A") // TestName3D
		ws2.getRange2("A16").setValue("A"); // TestNameArea3D
		ws2.getRange2("A17").setValue("a"); // TestNameArea3D

		// Positive cases:

		// Case #1: String. Convert sting abc to code.
		oParser = new parserFormula('CODE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 97, 'Test: Positive case: String. Convert sting abc to code.');
		// Case #2: Boolean. Convert  boolean - TRUE to code.
		oParser = new parserFormula('CODE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Positive case: Boolean. Convert  boolean - TRUE to code.');
		// Case #3: String. Basic test with an uppercase letter.
		oParser = new parserFormula('CODE("A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("A") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: String. Basic test with an uppercase letter.');
		// Case #4: String. Basic test with a lowercase letter.
		oParser = new parserFormula('CODE("a")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("a") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 97, 'Test: Positive case: String. Basic test with a lowercase letter.');
		// Case #5: String. Testing a number as a string. Should return the code for the character \'1\'.
		oParser = new parserFormula('CODE("1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: String. Testing a number as a string. Should return the code for the character \'1\'.');
		// Case #6: String. The function should only evaluate the first character of the string.
		oParser = new parserFormula('CODE("Hello")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("Hello") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 72, 'Test: Positive case: String. The function should only evaluate the first character of the string.');
		// Case #7: Number. Numeric input is converted to a text string "123". The code for "1" is returned.
		oParser = new parserFormula('CODE(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Number. Numeric input is converted to a text string "123". The code for "1" is returned.');
		// Case #8: String. Testing with a special character.
		oParser = new parserFormula('CODE("!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 33, 'Test: Positive case: String. Testing with a special character.');
		// Case #9: String. Testing a Cyrillic character.
		oParser = new parserFormula('CODE("Ю")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("Ю") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 63, 'Test: Positive case: String. Testing a Cyrillic character.');
		// Case #10: Reference link. Reference to a cell with a valid string.
		oParser = new parserFormula('CODE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 97, 'Test: Positive case: Reference link. Reference to a cell with a valid string.');
		// Case #11: Reference link. Reference to a cell with a lowercase letter.
		oParser = new parserFormula('CODE(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Reference link. Reference to a cell with a lowercase letter.');
		// Case #12: Formula. Nested formula where CHAR() returns "B".
		oParser = new parserFormula('CODE(CHAR(66))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(CHAR(66)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 66, 'Test: Positive case: Formula. Nested formula where CHAR() returns "B".');
		// Case #13: Formula. Testing where CODE is part of a parent formula SUM.
		oParser = new parserFormula('SUM(CODE("A"),CODE("B"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(CODE("A"),CODE("B")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 131, 'Test: Positive case: Formula. Testing where CODE is part of a parent formula SUM.');
		// Case #14: Name. Reference to a named range (single cell).
		oParser = new parserFormula('CODE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Name. Reference to a named range (single cell).');
		// Case #15: Ref3D. 3D Reference to a single cell on another sheet.
		oParser = new parserFormula('CODE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Ref3D. 3D Reference to a single cell on another sheet.');
		// Case #16: Name3D. 3D Reference to a named range on another sheet.
		oParser = new parserFormula('CODE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Name3D. 3D Reference to a named range on another sheet.');
		// Case #17: Area. Function spills results down for a multi-cell range.
		oParser = new parserFormula('CODE(A102:A103)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 33, 'Test: Positive case: Area. Function spills results down for a multi-cell range.');
		// Case #18: Area3D. Function spills results for a 3D multi-cell range.
		oParser = new parserFormula('CODE(Sheet2!A2:A3)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Sheet2!A2:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 97, 'Test: Positive case: Area3D. Function spills results for a 3D multi-cell range.');
		// Case #19: Name. Function spills results for a multi-cell named range.
		oParser = new parserFormula('CODE(TestNameArea)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TestNameArea) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 65, 'Test: Positive case: Name. Function spills results for a multi-cell named range.');
		// Case #20: Name3D. Function spills results for a 3D multi-cell named range.
		oParser = new parserFormula('CODE(TestNameArea3D)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 65, 'Test: Positive case: Name3D. Function spills results for a 3D multi-cell named range.');
		// Case #21: Array. Testing with a vertical array constant. Spills results.
		oParser = new parserFormula('CODE({"C";"D"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE({"C";"D"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), 67, 'Test: Positive case: Array. Testing with a vertical array constant. Spills results.');
		// Case #22: Table. Reference to a table column. Spills results.
		oParser = new parserFormula('CODE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Table. Reference to a table column. Spills results.');
		// Case #23: Area. Reference to a whole column.
		oParser = new parserFormula('CODE(A:A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A:A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Positive case: Area. Reference to a whole column.');
		// Case #24: Date. Date as a string is treated as text.
		oParser = new parserFormula('CODE("31/12/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("31/12/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 51, 'Test: Positive case: Date. Date as a string is treated as text.');
		// Case #25: Boolean. Boolean TRUE
		oParser = new parserFormula('CODE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Positive case: Boolean. Boolean TRUE');
		// Case #26: Boolean. Boolean FALSE
		oParser = new parserFormula('CODE(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 70, 'Test: Positive case: Boolean. Boolean FALSE');
		// Case #27: Reference link. Reference to a cell containing boolean TRUE, which is converted to string.
		oParser = new parserFormula('CODE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Positive case: Reference link. Reference to a cell containing boolean TRUE, which is converted to string.');

		// Negative cases:

		// Case #1: Empty. An empty string argument returns a #VALUE! error.
		oParser = new parserFormula('CODE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE("") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. An empty string argument returns a #VALUE! error.');
		// Case #2: Reference link. Reference to a completely empty cell returns a #VALUE! error.
		oParser = new parserFormula('CODE(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to a completely empty cell returns a #VALUE! error.');
		// Case #3: Error. Passing an error value directly to the function.
		oParser = new parserFormula('CODE(#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Passing an error value directly to the function.');
		// Case #4: Formula. A nested formula that results in an error.
		oParser = new parserFormula('CODE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. A nested formula that results in an error.');
		// Case #5: Reference link. Reference to a cell containing an error.
		oParser = new parserFormula('CODE(A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link. Reference to a cell containing an error.');
		// Case #6: Ref3D. 3D reference to an empty cell.
		oParser = new parserFormula('CODE(Sheet2!A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Sheet2!A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to an empty cell.');
		// Case #7: Area3D. 3D Area reference where the first cell is empty.
		oParser = new parserFormula('CODE(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D Area reference where the first cell is empty.');
		// Case #8: Table. Table column where the first cell is empty.
		oParser = new parserFormula('CODE(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Table1[Column2]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column where the first cell is empty.');
		// Case #9: Array. Array with an empty string returns an error.
		oParser = new parserFormula('CODE({""})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE({""}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#VALUE!', 'Test: Negative case: Array. Array with an empty string returns an error.');
		// Case #10: Array. Array with an error value returns an error.
		oParser = new parserFormula('CODE({#N/A})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE({#N/A}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), '#N/A', 'Test: Negative case: Array. Array with an error value returns an error.');
		// Case #11: Ref3D. 3D Reference to a cell containing an error.
		oParser = new parserFormula('CODE(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. 3D Reference to a cell containing an error.');

		// Bounded cases:

		// Case #1: String. Minimum bound: The space character, which is the first printable ASCII character (Code 32).
		oParser = new parserFormula('CODE(" ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(" ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32, 'Test: Bounded case: String. Minimum bound: The space character, which is the first printable ASCII character (Code 32).');
		// Case #2: Formula. Maximum bound (for Windows ANSI): Character with code 255 (\'ÿ\').
		oParser = new parserFormula('CODE(CHAR(255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CODE(CHAR(255)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 255, 'Test: Bounded case: Formula. Maximum bound (for Windows ANSI): Character with code 255 (\'ÿ\').');

        // Need to fix: diff results from MS, error types diff
        // Case #9: String. Testing a Cyrillic character.
        // Case #1: Empty. An empty string argument returns a #VALUE! error.
        // Case #2: Reference link. Reference to a completely empty cell returns a #VALUE! error.
        // Case #6: Ref3D. 3D reference to an empty cell.
        // Case #8: Table. Table column where the first cell is empty.
        // Case #9: Array. Array with an empty string returns an error.

		testArrayFormula2(assert, "CODE", 1, 1);
	});

	QUnit.test("Test: \"CONCATENATE\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("AA2").setValue("brook trout");
		ws.getRange2("AA3").setValue("species");
		ws.getRange2("AA4").setValue("32");

		ws.getRange2("AB2").setValue("Andreas");
		ws.getRange2("AB3").setValue("Fourth");

		ws.getRange2("AC2").setValue("Hauser");
		ws.getRange2("AC3").setValue("Pine");

		ws.getRange2("A100").setValue("Somebody ");
		ws.getRange2("A101").setValue("was told me");
		ws.getRange2("A102").setValue("the world is ");
		ws.getRange2("A103").setValue("gonna roll me");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A105").setValue("#REF!");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("Avangers "); // Column1
		ws.getRange2("B601").setValue("assemble"); // Column2
		ws.getRange2("C601").setValue("#NULL!"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("Hello");
		ws2.getRange2("A2").setValue("world");
		ws2.getRange2("A3").setValue("#DIV/0!");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("Life is "); // TestName
		ws.getRange2("A202").setValue("strong"); // TestName1
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("Hello") // TestName3D
		ws2.getRange2("A12").setValue("world") // TestName3D1

		// Positive cases:

		// Case #1: String(4), Reference link(4). Creates a sentence by joining the data in column AA with other text. 8 arguments were used.
		oParser = new parserFormula('CONCATENATE("Stream population for ", AA2, " ", AA3, " is ", AA4, "/mile.")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Stream population for ", AA2, " ", AA3, " is ", AA4, "/mile.") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Stream population for brook trout species is 32/mile.', 'Test: Positive case: String(4), Reference link(4). Creates a sentence by joining the data in column AA with other text. 8 arguments were used.');
		// Case #2: Reference link, String, Reference link. Joins three things: the string in cell AB2, a space character, and the value in cell AC2. 3 arguments were used.
		oParser = new parserFormula('CONCATENATE(AB2, " ", AC2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(AB2, " ", AC2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Andreas Hauser', 'Test: Positive case: Reference link, String, Reference link. Joins three things: the string in cell AB2, a space character, and the value in cell AC2. 3 arguments were used.');
		// Case #3: Reference link, String, Reference link. Joins three things: the string in cell AС2, a space character, and the value in cell AB2. 3 arguments were used.
		oParser = new parserFormula('CONCATENATE(AC2, ", ", AB2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(AC2, ", ", AB2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hauser, Andreas', 'Test: Positive case: Reference link, String, Reference link. Joins three things: the string in cell AС2, a space character, and the value in cell AB2. 3 arguments were used.');
		// Case #4: Reference link, String, Reference link. Joins three things: the string in cell AB3, a string consisting of a space with ampersand and another space, and the value in cell AC3. 3 arguments were used.
		oParser = new parserFormula('CONCATENATE(AB3, " & ", AC3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(AB3, " & ", AC3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Fourth & Pine', 'Test: Positive case: Reference link, String, Reference link. Joins three things: the string in cell AB3, a string consisting of a space with ampersand and another space, and the value in cell AC3. 3 arguments were used.');
		// Case #5: Boolean, String. Joins boolean and text string. 2 arguments were used.
		oParser = new parserFormula('CONCATENATE(TRUE,"test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(TRUE,"test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEtest', 'Test: Positive case: Boolean, String. Joins boolean and text string. 2 arguments were used.');
		// Case #6: Boolean, String. Concating boolean with string text. 2 arguments were used.
		oParser = new parserFormula('CONCATENATE(TRUE,"test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(TRUE,"test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEtest', 'Test: Positive case: Boolean, String. Concating boolean with string text. 2 arguments were used.');
		// Case #7: String(2). Basic string concatenation without delimiter.
		oParser = new parserFormula('CONCATENATE("Hello", "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Hello", "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HelloWorld', 'Test: Positive case: String(2). Basic string concatenation without delimiter.');
		// Case #8: String(3). Concatenation with space as separate string argument.
		oParser = new parserFormula('CONCATENATE("Hello", " ", "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Hello", " ", "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello World', 'Test: Positive case: String(3). Concatenation with space as separate string argument.');
		// Case #9: Number(2). Numbers are converted to text and concatenated.
		oParser = new parserFormula('CONCATENATE(1, 2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(1, 2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Number(2). Numbers are converted to text and concatenated.');
		// Case #10: String, Number. String and number combined.
		oParser = new parserFormula('CONCATENATE("Value: ", 42)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Value: ", 42) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Value: 42', 'Test: Positive case: String, Number. String and number combined.');
		// Case #11: Empty, String. Empty argument is treated as empty string.
		oParser = new parserFormula('CONCATENATE(, "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(, "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'World', 'Test: Positive case: Empty, String. Empty argument is treated as empty string.');
		// Case #12: String, Empty. Empty argument is treated as empty string.
		oParser = new parserFormula('CONCATENATE("Hello", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Hello", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: String, Empty. Empty argument is treated as empty string.');
		// Case #13: Reference link(2). Concatenation using cell references.
		oParser = new parserFormula('CONCATENATE(A100, A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(A100, A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Somebody was told me', 'Test: Positive case: Reference link(2). Concatenation using cell references.');
		// Case #14: Area. Single-column range concatenation.
		// Different result with MS
		oParser = new parserFormula('CONCATENATE(A102:A103)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'the world is ', 'Test: Positive case: Area. Single-column range concatenation.');
		// Case #15: Array. Array input with two elements.
		// Different result with MS
		oParser = new parserFormula('CONCATENATE({"Hello", "World"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE({"Hello", "World"}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: Array. Array input with two elements.');
		// Case #16: Name(2). Named ranges as arguments.
		oParser = new parserFormula('CONCATENATE(TestName, TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(TestName, TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Life is strong', 'Test: Positive case: Name(2). Named ranges as arguments.');
		// Case #17: Name3D(2). 3D named ranges.
		oParser = new parserFormula('CONCATENATE(TestName3D, TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(TestName3D, TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Helloworld', 'Test: Positive case: Name3D(2). 3D named ranges.');
		// Case #18: Ref3D(2). 3D cell references.
		oParser = new parserFormula('CONCATENATE(Sheet2!A1, Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(Sheet2!A1, Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Helloworld', 'Test: Positive case: Ref3D(2). 3D cell references.');
		// Case #19: Area3D. 3D range concatenation.
		oParser = new parserFormula('CONCATENATE(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: Area3D. 3D range concatenation.');
		// Case #20: Table(2). Table structured references.
		oParser = new parserFormula('CONCATENATE(Table1[Column1], Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(Table1[Column1], Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Avangers assemble', 'Test: Positive case: Table(2). Table structured references.');
		// Case #21: Formula, String. Date converted to serial number text.
		oParser = new parserFormula('CONCATENATE(DATE(2025,1,1), " is a date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(DATE(2025,1,1), " is a date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658 is a date', 'Test: Positive case: Formula, String. Date converted to serial number text.');
		// Case #22: Formula, String. Time converted to text "It\'s high noon".
		oParser = new parserFormula('CONCATENATE("It\'s ",TIME(12,0,0), " noon")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("It\'s ",TIME(12,0,0), " noon") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'It\'s 0.5 noon', 'Test: Positive case: Formula, String. Time converted to text "It\'s high noon".');
		// Case #23: Boolean(2). Booleans converted to "TRUE" and "FALSE".
		oParser = new parserFormula('CONCATENATE(TRUE, FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(TRUE, FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEFALSE', 'Test: Positive case: Boolean(2). Booleans converted to "TRUE" and "FALSE".');
		// Case #24: String, Formula. Nested function as argument.
		oParser = new parserFormula('CONCATENATE("Result: ", SQRT(4))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Result: ", SQRT(4)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Result: 2', 'Test: Positive case: String, Formula. Nested function as argument.');
		// Case #25: Formula. CONCATENATE inside another function.
		oParser = new parserFormula('SUM(LEN(CONCATENATE("A", "B")), 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(LEN(CONCATENATE("A", "B")), 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. CONCATENATE inside another function.');
		// Case #26: String(3). Formula returns "ABC".
		oParser = new parserFormula('CONCATENATE("A", "B", "C")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("A", "B", "C") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC', 'Test: Positive case: String(3). Formula returns "ABC".');
		// Case #27: Area. Full column reference.
		oParser = new parserFormula('CONCATENATE(AA:AA)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(AA:AA) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'a1', 'Test: Positive case: Area. Full column reference.');
		// Case #28: Area. Full row reference.
		oParser = new parserFormula('CONCATENATE(100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(100:100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Somebody ', 'Test: Positive case: Area. Full row reference.');
		// Case #29: String(3). Unicode and special characters handling.
		oParser = new parserFormula('CONCATENATE("Hello", "🌍", "!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Hello", "🌍", "!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello🌍!', 'Test: Positive case: String(3). Unicode and special characters handling.');

		// Negative cases:

		// Case #1: Error, String. Error value propagates.
		oParser = new parserFormula('CONCATENATE(#N/A, "Text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(#N/A, "Text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Error value propagates.');
		// Case #2: String, Error. Error value propagates.
		oParser = new parserFormula('CONCATENATE("Text", #VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Text", #VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Error. Error value propagates.');
		// Case #3: Formula. Result exceeds 32767 characters -> #VALUE! error.
		// Different result with MS
		oParser = new parserFormula('CONCATENATE(REPT("A", 32768))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(REPT("A", 32768)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Result exceeds 32767 characters -> #VALUE! error.');
		// Case #4: Reference link. Cell contains error -> propagates.
		oParser = new parserFormula('CONCATENATE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link. Cell contains error -> propagates.');
		// Case #5: Area. Range contains error -> propagates.
		// Different result with MS
		oParser = new parserFormula('CONCATENATE(A104:A105)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(A104:A105) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Area. Range contains error -> propagates.');
		// Case #6: Ref3D. 3D ref with error -> propagates.
		oParser = new parserFormula('CONCATENATE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Ref3D. 3D ref with error -> propagates.');
		// Case #7: Array. Array with error -> propagates.
		oParser = new parserFormula('CONCATENATE({#NUM!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE({#NUM!}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Array. Array with error -> propagates.');
		// Case #8: Table. Table column with error -> propagates.
		oParser = new parserFormula('CONCATENATE(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NULL!', 'Test: Negative case: Table. Table column with error -> propagates.');
		// Case #9: String, Formula. Nested function returns error -> propagates.
		oParser = new parserFormula('CONCATENATE("Text", SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("Text", SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String, Formula. Nested function returns error -> propagates.');

		// Bounded cases:

		// Case #1: String, Formula. Exceeds max length when combined. Exceeded char must be ignored
		// Different result with MS
		oParser = new parserFormula('CONCATENATE("A", REPT("B", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE("A", REPT("B", 32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A' + 'B'.repeat(32767), 'Test: Bounded case: String, Formula. Exceeds max length when combined. Exceeded char must be ignored');
		// Case #2: Formula. Max length without error (32767 chars).
		oParser = new parserFormula('CONCATENATE(REPT("A", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(REPT("A", 32767)) is parsed.');
		let expectedData = 'A'.repeat(32767)
		assert.strictEqual(oParser.calculate().getValue(), expectedData, 'Test: Bounded case: Formula. Max length without error (32767 chars).');
		// Case #3: Number(2). Min numeric values.
		oParser = new parserFormula('CONCATENATE(0, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(0, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '00', 'Test: Bounded case: Number(2). Min numeric values.');
		// Case #4: Number(2). Large numbers converted to text.
		oParser = new parserFormula('CONCATENATE(9.99999999999999E+307, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(9.99999999999999E+307, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '9.99999999999999e+3071', 'Test: Bounded case: Number(2). Large numbers converted to text.');
		// Case #5: Formula, String. Min date value (serial number 1).
		oParser = new parserFormula('CONCATENATE(DATE(1900,1,1), " date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(DATE(1900,1,1), " date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1 date', 'Test: Bounded case: Formula, String. Min date value (serial number 1).');
		// Case #6: Formula, String. Max date value.
		oParser = new parserFormula('CONCATENATE(DATE(9999,12,31), " date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(DATE(9999,12,31), " date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2958465 date', 'Test: Bounded case: Formula, String. Max date value.');
		// Case #7: Formula, String. Max length with empty string appended.
		// Different result with MS
		oParser = new parserFormula('CONCATENATE(REPT("A", 32767), "")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(REPT("A", 32767), "") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), expectedData + '', 'Test: Bounded case: Formula, String. Max length with empty string appended.');

        // Need to fix: array getValue?, too long string restriction

		testArrayFormula2(assert, "CONCATENATE", 1, 8);
	});

	QUnit.test("Test: \"CONCAT\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("Somebody ");
		ws.getRange2("A101").setValue("was told me");
		ws.getRange2("A102").setValue("the world is ");
		ws.getRange2("A103").setValue("gonna roll me");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A105").setValue("#REF!");

		ws.getRange2("AA:BB").cleanAll();
		ws.getRange2("AA1").setValue("a1");
		ws.getRange2("AA2").setValue("a2");
		ws.getRange2("AA4").setValue("a4");
		ws.getRange2("AA5").setValue("a5");
		ws.getRange2("AA6").setValue("a6");
		ws.getRange2("AA7").setValue("a7");

		ws.getRange2("BB:BB").cleanAll();
		ws.getRange2("BB1").setValue("b1");
		ws.getRange2("BB2").setValue("b2");
		ws.getRange2("BB4").setValue("b4");
		ws.getRange2("BB5").setValue("b5");
		ws.getRange2("BB6").setValue("b6");
		ws.getRange2("BB7").setValue("b7");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("Avangers "); // Column1
		ws.getRange2("B601").setValue("assemble"); // Column2
		ws.getRange2("C601").setValue("#NULL!"); // Column3
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("Hello");
		ws2.getRange2("A2").setValue("world");
		ws2.getRange2("A3").setValue("#DIV/0!");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("Life is "); // TestName
		ws.getRange2("A202").setValue("strong"); // TestName1
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("Hello") // TestName3D
		ws2.getRange2("A12").setValue("world") // TestName3D1


		// Positive cases:

		// Case #1: String(11). Concating partial strings. 11 arguments were used.
		oParser = new parserFormula('CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'The sun will come up tomorrow.', 'Test: Positive case: String(11). Concating partial strings. 11 arguments were used.');
		// Case #2: Area(2). Concating whole cols AA and BB. 2 arguments were used.
		oParser = new parserFormula('CONCAT(AA:AA, BB:BB)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(AA:AA, BB:BB) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a1a2a4a5a6a7b1b2b4b5b6b7', 'Test: Positive case: Area(2). Concating whole cols AA and BB. 2 arguments were used.');
		// Case #3: Area. Concating area with text. 1 argument was  used.
		oParser = new parserFormula('CONCAT(AA1:BB7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(AA1:BB7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a1b1a2b2a4b4a5b5a6b6a7b7', 'Test: Positive case: Area. Concating area with text. 1 argument was  used.');
		// Case #4: Boolean, String. Concating boolean with string text. 2 arguments were used.
		oParser = new parserFormula('CONCAT(TRUE,"test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(TRUE,"test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEtest', 'Test: Positive case: Boolean, String. Concating boolean with string text. 2 arguments were used.');
		// Case #5: String(2). Basic string concatenation without delimiter.
		oParser = new parserFormula('CONCAT("Hello", "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Hello", "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HelloWorld', 'Test: Positive case: String(2). Basic string concatenation without delimiter.');
		// Case #6: String(3). Concatenation with space as separate string argument.
		oParser = new parserFormula('CONCAT("Hello", " ", "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Hello", " ", "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello World', 'Test: Positive case: String(3). Concatenation with space as separate string argument.');
		// Case #7: Number(2). Numbers are converted to text and concatenated.
		oParser = new parserFormula('CONCAT(1, 2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(1, 2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Number(2). Numbers are converted to text and concatenated.');
		// Case #8: String, Number. String and number combined.
		oParser = new parserFormula('CONCAT("Value: ", 42)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Value: ", 42) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Value: 42', 'Test: Positive case: String, Number. String and number combined.');
		// Case #9: Empty, String. Empty argument is treated as empty string.
		oParser = new parserFormula('CONCAT(, "World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(, "World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'World', 'Test: Positive case: Empty, String. Empty argument is treated as empty string.');
		// Case #10: String, Empty. Empty argument is treated as empty string.
		oParser = new parserFormula('CONCAT("Hello", )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Hello", ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: String, Empty. Empty argument is treated as empty string.');
		// Case #11: Reference link(2). Concatenation using cell references.
		oParser = new parserFormula('CONCAT(A100, A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(A100, A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Somebody was told me', 'Test: Positive case: Reference link(2). Concatenation using cell references.');
		// Case #12: Area. Single-column range concatenation.
		oParser = new parserFormula('CONCAT(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'the world is gonna roll me', 'Test: Positive case: Area. Single-column range concatenation.');
		// Case #13: Array. Array input with two elements.
		oParser = new parserFormula('CONCAT({"Hello", "World"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT({"Hello", "World"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HelloWorld', 'Test: Positive case: Array. Array input with two elements.');
		// Case #14: Name(2). Named ranges as arguments.
		oParser = new parserFormula('CONCAT(TestName, TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(TestName, TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Life is strong', 'Test: Positive case: Name(2). Named ranges as arguments.');
		// Case #15: Name3D(2). 3D named ranges.
		oParser = new parserFormula('CONCAT(TestName3D, TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(TestName3D, TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Helloworld', 'Test: Positive case: Name3D(2). 3D named ranges.');
		// Case #16: Ref3D(2). 3D cell references.
		oParser = new parserFormula('CONCAT(Sheet2!A1, Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(Sheet2!A1, Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Helloworld', 'Test: Positive case: Ref3D(2). 3D cell references.');
		// Case #17: Area3D. 3D range concatenation.
		oParser = new parserFormula('CONCAT(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Helloworld', 'Test: Positive case: Area3D. 3D range concatenation.');
		// Case #18: Table(2). Table structured references.
		oParser = new parserFormula('CONCAT(Table1[Column1], Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(Table1[Column1], Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Avangers assemble', 'Test: Positive case: Table(2). Table structured references.');
		// Case #19: Formula, String. Date converted to serial number text.
		oParser = new parserFormula('CONCAT(DATE(2025,1,1), " is a date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(DATE(2025,1,1), " is a date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658 is a date', 'Test: Positive case: Formula, String. Date converted to serial number text.');
		// Case #20: Formula, String. Time converted to text "It\'s high noon".
		oParser = new parserFormula('CONCAT("It\'s ",TIME(12,0,0), " noon")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("It\'s ",TIME(12,0,0), " noon") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'It\'s 0.5 noon', 'Test: Positive case: Formula, String. Time converted to text "It\'s high noon".');
		// Case #21: Boolean(2). Booleans converted to "TRUE" and "FALSE".
		oParser = new parserFormula('CONCAT(TRUE, FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(TRUE, FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEFALSE', 'Test: Positive case: Boolean(2). Booleans converted to "TRUE" and "FALSE".');
		// Case #22: String, Formula. Nested function as argument.
		oParser = new parserFormula('CONCAT("Result: ", SQRT(4))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Result: ", SQRT(4)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Result: 2', 'Test: Positive case: String, Formula. Nested function as argument.');
		// Case #23: Formula. CONCAT inside another function.
		oParser = new parserFormula('SUM(LEN(CONCAT("A", "B")), 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(LEN(CONCAT("A", "B")), 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. CONCAT inside another function.');
		// Case #24: String(3). Formula returns "ABC".
		oParser = new parserFormula('CONCAT("A", "B", "C")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("A", "B", "C") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC', 'Test: Positive case: String(3). Formula returns "ABC".');
		// Case #25: Area. Full column reference.
		oParser = new parserFormula('CONCAT(AA:AA)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(AA:AA) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a1a2a4a5a6a7', 'Test: Positive case: Area. Full column reference.');
		// Case #26: Area. Full row reference.
		oParser = new parserFormula('CONCAT(100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(100:100) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Somebody 3-4CTRUE', 'Test: Positive case: Area. Full row reference.');
		// Case #27: String(3). Unicode and special characters handling.
		oParser = new parserFormula('CONCAT("Hello", "🌍", "!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Hello", "🌍", "!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello🌍!', 'Test: Positive case: String(3). Unicode and special characters handling.');

		// Negative cases:

		// Case #1: Error, String. Error value propagates.
		oParser = new parserFormula('CONCAT(#N/A, "Text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(#N/A, "Text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Error value propagates.');
		// Case #2: String, Error. Error value propagates.
		oParser = new parserFormula('CONCAT("Text", #VALUE!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Text", #VALUE!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Error. Error value propagates.');
		// Case #3: Formula. Result exceeds 32767 characters -> #VALUE! error.
		oParser = new parserFormula('CONCAT(REPT("A", 32768))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(REPT("A", 32768)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Result exceeds 32767 characters -> #VALUE! error.');
		// Case #4: Reference link. Cell contains error -> propagates.
		oParser = new parserFormula('CONCAT(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link. Cell contains error -> propagates.');
		// Case #5: Area. Range contains error -> propagates.
		// Different result with MS
		oParser = new parserFormula('CONCAT(A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(A104:A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area. Range contains error -> propagates.');
		// Case #6: Ref3D. 3D ref with error -> propagates.
		oParser = new parserFormula('CONCAT(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: Ref3D. 3D ref with error -> propagates.');
		// Case #7: Array. Array with error -> propagates.
		oParser = new parserFormula('CONCAT({#NUM!})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT({#NUM!}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Array. Array with error -> propagates.');
		// Case #8: Table. Table column with error -> propagates.
		oParser = new parserFormula('CONCAT(Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NULL!', 'Test: Negative case: Table. Table column with error -> propagates.');
		// Case #9: String, Formula. Nested function returns error -> propagates.
		oParser = new parserFormula('CONCAT("Text", SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("Text", SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String, Formula. Nested function returns error -> propagates.');
		// Case #10: String, Formula. Exceeds max length when combined -> #CALC! error.
		// Different result with MS
		oParser = new parserFormula('CONCAT("A", REPT("B", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT("A", REPT("B", 32767)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#CALC!', 'Test: Negative case: String, Formula. Exceeds max length when combined -> #VALUE! error.');

		// Bounded cases:

		// Case #1: Formula. Max length without error (32767 chars).
		oParser = new parserFormula('CONCAT(REPT("A", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(REPT("A", 32767)) is parsed.');
		let expectedData = 'A'.repeat(32767);
		assert.strictEqual(oParser.calculate().getValue(), expectedData, 'Test: Bounded case: Formula. Max length without error (32767 chars).');
		// Case #2: Number(2). Min numeric values.
		oParser = new parserFormula('CONCAT(0, 0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(0, 0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '00', 'Test: Bounded case: Number(2). Min numeric values.');
		// Case #3: Number(2). Large numbers converted to text.
		oParser = new parserFormula('CONCAT(9.99999999999999E+307, 1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(9.99999999999999E+307, 1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '9.99999999999999e+3071', 'Test: Bounded case: Number(2). Large numbers converted to text.');
		// Case #4: Formula, String. Min date value (serial number 1).
		oParser = new parserFormula('CONCAT(DATE(1900,1,1), " date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(DATE(1900,1,1), " date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1 date', 'Test: Bounded case: Formula, String. Min date value (serial number 1).');
		// Case #5: Formula, String. Max date value.
		oParser = new parserFormula('CONCAT(DATE(9999,12,31), " date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(DATE(9999,12,31), " date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2958465 date', 'Test: Bounded case: Formula, String. Max date value.');
		// Case #6: Formula, String. Max length with empty string appended.
		oParser = new parserFormula('CONCAT(REPT("A", 32767), "")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCAT(REPT("A", 32767), "") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), expectedData, 'Test: Bounded case: Formula, String. Max length with empty string appended.');

        // Need to fix: different results from MS, #VALUE error instead #CALC

		testArrayFormulaEqualsValues(assert,
			"13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245;13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245;13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245,13.123-424513.123-424513.123-4245",
			"CONCAT(A1:C2,A1:C2,A1:C2)")
	});

	QUnit.test("Test: \"DOLLAR\"", function (assert) {
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1234.567");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A102").setValue("2.5");
		ws.getRange2("A103").setValue("3.2");
		ws.getRange2("A104").setValue("");
		ws.getRange2("A105").setValue("text");
		ws.getRange2("A106").setValue("#NULL!");
		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1234.567"); // Column1
		ws.getRange2("B601").setValue("2"); // Column2
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("1234.567");
		ws2.getRange2("A2").setValue("2");
		ws2.getRange2("A3").setValue("2.5");
		ws2.getRange2("A4").setValue("3.2");
		// DefNames. Use A201-A208, B208
		ws.getRange2("A201").setValue("1234.567"); // TestName
		ws.getRange2("A202").setValue("2"); // TestName1
		// DefNames 3D. Use A11-A18, B18
		ws2.getRange2("A11").setValue("1234.567") // TestName3D
		ws2.getRange2("A12").setValue("2") // TestName3D1

		// Positive cases:

		// Case #1: Number. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number. Return $1,234.57');
		// Case #2: Number(2). Return $1,200
		oParser = new parserFormula('DOLLAR(1234.567,-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,200', 'Test: Positive case: Number(2). Return $1,200');
		// Case #3: Number(2). Return ($1,234.5670)
		oParser = new parserFormula('DOLLAR(-1234.567,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1234.567,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '($1,234.5670)', 'Test: Positive case: Number(2). Return ($1,234.5670)');
		// Case #4: Number(2). Return ($1,235)
		oParser = new parserFormula('DOLLAR(-1234.567,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1234.567,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '($1,235)', 'Test: Positive case: Number(2). Return ($1,235)');

		//set russia locale
		window["Asc"]["editor"].asc_setLocale(1049);

		// Case #5: Number. Return 1 234,57 ₽
		oParser = new parserFormula('DOLLAR(1234.567)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1 234,57 ₽', 'Test: Positive case: Number. Return 1 234,57 ₽');
		// Case #6: Number(2). Return 1 200 ₽
		oParser = new parserFormula('DOLLAR(1234.567,-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1 200 ₽', 'Test: Positive case: Number(2). Return 1 200 ₽');
		// Case #7: Number(2). Return 1 234,5679 ₽
		oParser = new parserFormula('DOLLAR(-1234.567,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1234.567,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1 234,5670 ₽', 'Test: Positive case: Number(2). Return 1 234,5679 ₽');
		// Case #8: Number(2). Return 1 235 ₽
		oParser = new parserFormula('DOLLAR(-1234.567,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1234.567,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1 235 ₽', 'Test: Positive case: Number(2). Return -1 235 ₽');

		//set default locale
		window["Asc"]["editor"].asc_setLocale(1033);
		// Case #9: Number(2). Displays number in currency format, 4 digits to the right of decimal point. Return $0.1230
		oParser = new parserFormula('DOLLAR(0.123,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(0.123,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.1230', 'Test: Positive case: Number(2). Displays number in currency format, 4 digits to the right of decimal point. Return $0.1230');
		// Case #10: Number. Decimals argument omitted, default value 2 used. Return $99.89
		oParser = new parserFormula('DOLLAR(99.888)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(99.888) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$99.89', 'Test: Positive case: Number. Decimals argument omitted, default value 2 used. Return $99.89');
		// Case #11: Number(2). Negative number displayed in parentheses with currency format. Return ($1,234.57)
		oParser = new parserFormula('DOLLAR(-1234.567,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1234.567,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '($1,234.57)', 'Test: Positive case: Number(2). Negative number displayed in parentheses with currency format. Return ($1,234.57)');
		// Case #12: Number(2). Zero decimals, number rounded to integer. Return $5,000
		oParser = new parserFormula('DOLLAR(5000,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(5000,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$5,000', 'Test: Positive case: Number(2). Zero decimals, number rounded to integer. Return $5,000');
		// Case #13: Number(2). Negative decimals -1, rounds to tens place. Return $1,230
		oParser = new parserFormula('DOLLAR(1234.567,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,230', 'Test: Positive case: Number(2). Negative decimals -1, rounds to tens place. Return $1,230');
		// Case #14: Number(2). Negative decimals -3, rounds to thousands place. Return $1,000
		oParser = new parserFormula('DOLLAR(1234.567,-3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,000', 'Test: Positive case: Number(2). Negative decimals -3, rounds to thousands place. Return $1,000');
		// Case #15: Number(2). Zero value with 2 decimals. Return $0.00
		oParser = new parserFormula('DOLLAR(0,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(0,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.00', 'Test: Positive case: Number(2). Zero value with 2 decimals. Return $0.00');
		// Case #16: Empty, Number. Empty first argument converted to 0. Return $0.00
		oParser = new parserFormula('DOLLAR(,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.00', 'Test: Positive case: Empty, Number. Empty first argument converted to 0. Return $0.00');
		// Case #17: Number, Empty. Empty decimals argument, default value 2 used. Return $100.00
		oParser = new parserFormula('DOLLAR(100,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(100,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$100', 'Test: Positive case: Number, Empty. Empty decimals argument, default value 2 used. Return $100.00');
		// Case #18: Empty(2). Both arguments empty, converted to 0 and default 2. Return $0.00
		oParser = new parserFormula('DOLLAR(,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0', 'Test: Positive case: Empty(2). Both arguments empty, converted to 0 and default 2. Return $0.00');
		// Case #19: String, Number. Numeric string converted to number. Return $1,234.56
		oParser = new parserFormula('DOLLAR("1234.56",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR("1234.56",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.56', 'Test: Positive case: String, Number. Numeric string converted to number. Return $1,234.56');
		// Case #20: Number, String. Decimals as numeric string converted to number. Return $1,234.56
		oParser = new parserFormula('DOLLAR(1234.56,"2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.56,"2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.56', 'Test: Positive case: Number, String. Decimals as numeric string converted to number. Return $1,234.56');
		// Case #21: String(2). Both arguments as numeric strings. Return $1,234.56
		oParser = new parserFormula('DOLLAR("1234.56","2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR("1234.56","2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.56', 'Test: Positive case: String(2). Both arguments as numeric strings. Return $1,234.56');
		// Case #22: Number, Boolean. Boolean TRUE converted to 1. Return $1,234.6
		oParser = new parserFormula('DOLLAR(1234.567,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.6', 'Test: Positive case: Number, Boolean. Boolean TRUE converted to 1. Return $1,234.6');
		// Case #23: Number, Boolean. Boolean FALSE converted to 0. Return $1,235
		oParser = new parserFormula('DOLLAR(1234.567,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,235', 'Test: Positive case: Number, Boolean. Boolean FALSE converted to 0. Return $1,235');
		// Case #24: Boolean, Number. Boolean TRUE converted to 1. Return $1.00
		oParser = new parserFormula('DOLLAR(TRUE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(TRUE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1.00', 'Test: Positive case: Boolean, Number. Boolean TRUE converted to 1. Return $1.00');
		// Case #25: Boolean, Number. Boolean FALSE converted to 0. Return $0.00
		oParser = new parserFormula('DOLLAR(FALSE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(FALSE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.00', 'Test: Positive case: Boolean, Number. Boolean FALSE converted to 0. Return $0.00');
		// Case #26: Number(2). Float decimals truncated to 2. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,2.9)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,2.9) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number(2). Float decimals truncated to 2. Return $1,234.57');
		// Case #27: Number(2). Float decimals truncated to 2. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,2.1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,2.1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number(2). Float decimals truncated to 2. Return $1,234.57');
		// Case #28: Formula, Number. Nested ROUND formula as first argument. Return $1,234.60
		oParser = new parserFormula('DOLLAR(ROUND(1234.567,1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(ROUND(1234.567,1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.60', 'Test: Positive case: Formula, Number. Nested ROUND formula as first argument. Return $1,234.60');
		// Case #29: Number, Formula. Nested ROUND formula as decimals argument. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,ROUND(2.8,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,ROUND(2.8,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.567', 'Test: Positive case: Number, Formula. Nested ROUND formula as decimals argument. Return $1,234.57');
		// Case #30: Formula(2). Both arguments using nested formulas. Return $1,234.57
		oParser = new parserFormula('DOLLAR(ABS(-1234.567),INT(2.9))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(ABS(-1234.567),INT(2.9)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Formula(2). Both arguments using nested formulas. Return $1,234.57');
		// Case #31: Formula. DOLLAR as part of SUM formula.
		oParser = new parserFormula('SUM(DOLLAR(100,2),DOLLAR(200,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula SUM(DOLLAR(100,2),DOLLAR(200,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 300, 'Test: Positive case: Formula. DOLLAR as part of SUM formula.');
		// Case #32: Formula. DOLLAR as part of CONCATENATE formula. Return $100.00 USD
		oParser = new parserFormula('CONCATENATE(DOLLAR(100,2)," USD")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula CONCATENATE(DOLLAR(100,2)," USD") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$100.00 USD', 'Test: Positive case: Formula. DOLLAR as part of CONCATENATE formula. Return $100.00 USD');
		// Case #33: Reference link, Number. First argument as reference link to 1234.567. Return $1,234.57
		oParser = new parserFormula('DOLLAR(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Reference link, Number. First argument as reference link to 1234.567. Return $1,234.57');
		// Case #34: Number, Reference link. Decimals argument as reference link to 2. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number, Reference link. Decimals argument as reference link to 2. Return $1,234.57');
		// Case #35: Reference link(2). Both arguments as reference links to 1234.567 and 2. Return $1,234.57
		oParser = new parserFormula('DOLLAR(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Reference link(2). Both arguments as reference links to 1234.567 and 2. Return $1,234.57');
		// Case #36: Reference link. First argument as reference link, decimals omitted. Return $1,234.57
		oParser = new parserFormula('DOLLAR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Reference link. First argument as reference link, decimals omitted. Return $1,234.57');
		// Case #37: Area, Number. First argument as area range. Return $2.50
		oParser = new parserFormula('DOLLAR(A102:A103,2)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A102:A103,2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '$2.50', 'Test: Positive case: Area, Number. First argument as area range. Return $2.50');
		// Case #38: Area, Number. First argument as area range. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,A102:A103)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '$1,234.57', 'Test: Positive case: Area, Number. First argument as area range. Return $1,234.57');
		// Case #39: Name, Number. First argument as defined name to 5000. Return $1,234.57
		oParser = new parserFormula('DOLLAR(TestName,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(TestName,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Name, Number. First argument as defined name to 5000. Return $1,234.57');
		// Case #40: Number, Name. Decimals argument as defined name to 3. Return $1,234.567
		oParser = new parserFormula('DOLLAR(1234.567,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number, Name. Decimals argument as defined name to 3. Return $1,234.567');
		// Case #41: Name(2). Both arguments as defined names to 5000 and 3. Return $1,234.57
		oParser = new parserFormula('DOLLAR(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Name(2). Both arguments as defined names to 5000 and 3. Return $1,234.57');
		// Case #42: Name3D, Number. First argument as 3D defined name to 7500. Return $1,234.57
		oParser = new parserFormula('DOLLAR(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Name3D, Number. First argument as 3D defined name to 7500. Return $1,234.57');
		// Case #43: Number, Name3D. Decimals argument as 3D defined name to 1. Return $1,234.57
		oParser = new parserFormula('DOLLAR(1234.567,TestName3D1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,TestName3D1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number, Name3D. Decimals argument as 3D defined name to 1. Return $1,234.57');
		// Case #44: Ref3D, Number. First argument as 3D reference to 2500. Return $2,500.00
		oParser = new parserFormula('DOLLAR(Sheet2!A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(Sheet2!A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Ref3D, Number. First argument as 3D reference to 2500. Return $2,500.00');
		// Case #45: Number, Ref3D. Decimals argument as 3D reference to 4. Return $1,234.5670
		oParser = new parserFormula('DOLLAR(1234.567,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number, Ref3D. Decimals argument as 3D reference to 4. Return $1,234.5670');
		// Case #46: Ref3D(2). Both arguments as 3D references to 2500 and 4. Return $2,500.0000
		oParser = new parserFormula('DOLLAR(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Ref3D(2). Both arguments as 3D references to 2500 and 4. Return $2,500.0000');
		// Case #47: Area3D, Number. First argument as 3D area to 3000. Return $2.50
		oParser = new parserFormula('DOLLAR(Sheet2!A3:A4,2)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(Sheet2!A3:A4,2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '$2.50', 'Test: Positive case: Area3D, Number. First argument as 3D area to 3000. Return $2.50');
		// Case #48: Area3D, Number. First argument as 3D area to 3000. Return $1,234.56
		oParser = new parserFormula('DOLLAR(1234.56,Sheet2!A3:A4)', 'A2', ws);
        oParser.setArrayFormulaRef(ws.getRange2("A2:A3").bbox);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.56,Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '$1,234.56', 'Test: Positive case: Area3D, Number. First argument as 3D area to 3000. Return $1,234.56');
		// Case #49: Table, Number. First argument as table reference to 1500. Return $1,500.00
		oParser = new parserFormula('DOLLAR(Table1[Column1],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(Table1[Column1],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Table, Number. First argument as table reference to 1500. Return $1,500.00');
		// Case #50: Number, Table. Decimals argument as table reference to 3. Return $1,234.567
		oParser = new parserFormula('DOLLAR(1234.567,Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.57', 'Test: Positive case: Number, Table. Decimals argument as table reference to 3. Return $1,234.567');
		// Case #51: Reference link. Reference to empty cell, converted to 0. Return $0.00
		oParser = new parserFormula('DOLLAR(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.00', 'Test: Positive case: Reference link. Reference to empty cell, converted to 0. Return $0.00');
		// Case #52: Number(2). Large number with thousands separators. Return $1,000,000.00
		oParser = new parserFormula('DOLLAR(1000000,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1000000,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,000,000.00', 'Test: Positive case: Number(2). Large number with thousands separators. Return $1,000,000.00');
		// Case #53: Number(2). Very small number with 5 decimals. Return $0.00100
		oParser = new parserFormula('DOLLAR(0.001,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(0.001,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0.00100', 'Test: Positive case: Number(2). Very small number with 5 decimals. Return $0.00100');
		// Case #54: Number(2). Decimals below minimum -127.
		oParser = new parserFormula('DOLLAR(1234.567,-128)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-128) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0', 'Test: Positive case: Number(2). Decimals below minimum -127.');
		// Case #55: Date, Number. Date formula.
		oParser = new parserFormula('DOLLAR(DATE(2000,13,1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(DATE(2000,13,1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$36,892.00', 'Test: Positive case: Date, Number. Date formula.');
		// Case #56: Boolean(2). Both arguments as TRUE, converts to DOLLAR(1,1). Return $1.0
		oParser = new parserFormula('DOLLAR(TRUE,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(TRUE,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1.0', 'Test: Positive case: Boolean(2). Both arguments as TRUE, converts to DOLLAR(1,1). Return $1.0');
		// Case #57: Array, Number. Array as first argument returns first element. Return $100.00
		oParser = new parserFormula('DOLLAR({100;200},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR({100;200},2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0), "100", 'Test: Positive case: Array, Number. Array as first argument returns first element. Return $100.00');

		// Negative cases:

		// Case #1: String, Number. Non-numeric string returns #VALUE! error
		oParser = new parserFormula('DOLLAR("text",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR("text",2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Non-numeric string returns #VALUE! error');
		// Case #2: Number, String. Non-numeric string in decimals returns #VALUE! error
		oParser = new parserFormula('DOLLAR(1234.567,"text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,"text") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. Non-numeric string in decimals returns #VALUE! error');
		// Case #3: String(2). Both arguments as non-numeric strings return #VALUE! error
		oParser = new parserFormula('DOLLAR("text","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR("text","abc") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2). Both arguments as non-numeric strings return #VALUE! error');
		// Case #4: Error, Number. First argument as error returns #N/A
		oParser = new parserFormula('DOLLAR(#N/A,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(#N/A,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. First argument as error returns #N/A');
		// Case #5: Number, Error. Decimals argument as error returns #N/A
		oParser = new parserFormula('DOLLAR(1234.567,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Number, Error. Decimals argument as error returns #N/A');
		// Case #6: Error(2). Both arguments as errors, first error returned #VALUE!
		oParser = new parserFormula('DOLLAR(#VALUE!,#NUM!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(#VALUE!,#NUM!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Error(2). Both arguments as errors, first error returned #VALUE!');
		// Case #7: Reference link, Number. Reference to cell with text returns #VALUE!
		oParser = new parserFormula('DOLLAR(A105,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A105,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link, Number. Reference to cell with text returns #VALUE!');
		// Case #8: Number, Reference link. Decimals reference to cell with text returns #VALUE!
		oParser = new parserFormula('DOLLAR(1234.567,A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, Reference link. Decimals reference to cell with text returns #VALUE!');
		// Case #9: Reference link, Number. Reference to cell with error returns #NULL!
		oParser = new parserFormula('DOLLAR(A106,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A106,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NULL!', 'Test: Negative case: Reference link, Number. Reference to cell with error returns #NULL!');
		// Case #10: Number, Reference link. Decimals reference to cell with error returns #NULL!
		oParser = new parserFormula('DOLLAR(1234.567,A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,A106) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NULL!', 'Test: Negative case: Number, Reference link. Decimals reference to cell with error returns #NULL!');
		// Case #11: Number(2). Decimals exceeds maximum 127, returns #NUM! error
		oParser = new parserFormula('DOLLAR(1234.567,128)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,128) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number(2). Decimals exceeds maximum 127, returns #NUM! error');
		// Case #12: String, Number. Invalid numeric string format returns #VALUE! error
		oParser = new parserFormula('DOLLAR("1234.56.78",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR("1234.56.78",2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Invalid numeric string format returns #VALUE! error');

		// Bounded cases:

		// Case #1: Number(2). Maximum possible number in Excel.
		oParser = new parserFormula('DOLLAR(1E+187,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1E+187,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$10,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000.00', 'Test: Bounded case: Number(2). Maximum possible number in Excel.');
		// Case #2: Number(2). Minimum possible number in Excel.
		oParser = new parserFormula('DOLLAR(-1E+186,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(-1E+186,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '($1,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000.00)', 'Test: Bounded case: Number(2). Minimum possible number in Excel.');
		// Case #3: Number(2). Maximum decimals value 127. Return $1,234.567000...
		oParser = new parserFormula('DOLLAR(1234.567,127)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,127) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$1,234.5670000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000', 'Test: Bounded case: Number(2). Maximum decimals value 127. Return $1,234.567000...');
		// Case #4: Number(2). Minimum decimals value -127. Return $0
		oParser = new parserFormula('DOLLAR(1234.567,-127)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(1234.567,-127) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '$0', 'Test: Bounded case: Number(2). Minimum decimals value -127. Return $0');
		// Case #6: Area(2). Whole column and row as arguments. Return value from intersection
		oParser = new parserFormula('DOLLAR(A:A,100:100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: Formula DOLLAR(A:A,100:100) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Area(2). Whole column and row as arguments. Return value from intersection');

        // Need to fix: error types diff, results diff from MS
        // Case #1: String, Number. Non-numeric string returns #VALUE! error
        // Case #2: Number, String. Non-numeric string in decimals returns #VALUE! error
        // Case #3: String(2). Both arguments as non-numeric strings return #VALUE! error
        // Case #11: Number(2). Decimals exceeds maximum 127, returns #NUM! error
        // Case #12: String, Number. Invalid numeric string format returns #VALUE! error
        // Case #6: Area(2). Whole column and row as arguments. Return value from intersection

		testArrayFormula2(assert, "DOLLAR", 2, 2);
	});

	QUnit.test("Test: \"EXACT\"", function (assert) {

		ws.getRange2("A2").setValue("word");
		ws.getRange2("A3").setValue("Word");
		ws.getRange2("A4").setValue("w ord");
		ws.getRange2("B2").setValue("word");
		ws.getRange2("B3").setValue("word");
		ws.getRange2("B4").setValue("word");

		oParser = new parserFormula("EXACT(A2,B2)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula("EXACT(A3,B3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula("EXACT(A4,B4)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		oParser = new parserFormula("EXACT(TRUE,TRUE)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('EXACT("TRUE",TRUE)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('EXACT("TRUE","TRUE")', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE");

		oParser = new parserFormula('EXACT("true",TRUE)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "FALSE");

		ws.getRange2("A100:D210").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("1s");
		ws.getRange2("A104").setValue("");
		// For area
		ws.getRange2("A102").setValue("Text");
		ws.getRange2("A103").setValue("Text2");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
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
		// Case #1: String. Basic valid input: identical strings, case-sensitive. Returns TRUE.
		oParser = new parserFormula('EXACT("Text","Text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("Text","Text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Basic valid input: identical strings, case-sensitive. Returns TRUE.');
		// Case #2: String. Identical lowercase strings. Returns TRUE.
		oParser = new parserFormula('EXACT("text","text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("text","text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Identical lowercase strings. Returns TRUE.');
		// Case #3: Number. Numbers converted to strings. Returns TRUE.
		oParser = new parserFormula('EXACT(123,123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(123,123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Number. Numbers converted to strings. Returns TRUE.');
		// Case #4: Formula. Nested CONCATENATE producing identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(CONCATENATE("Te","xt"),CONCATENATE("Te","xt"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(CONCATENATE("Te","xt"),CONCATENATE("Te","xt")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested CONCATENATE producing identical strings. Returns TRUE.');
		// Case #5: Reference link. Reference to cells with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link. Reference to cells with identical strings. Returns TRUE.');
		// Case #6: Area. Single-cell ranges with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(A102:A102,A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(A102:A102,A103:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area. Single-cell ranges with identical strings. Returns TRUE.');
		// Case #7: Array. Arrays with single identical string element. Returns TRUE.
		oParser = new parserFormula('EXACT({"Text"},{"Text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT({"Text"},{"Text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Arrays with single identical string element. Returns TRUE.');
		// Case #8: Name. Named ranges with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Name. Named ranges with identical strings. Returns TRUE.');
		// Case #9: Name3D. 3D named ranges with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D. 3D named ranges with identical strings. Returns TRUE.');
		// Case #10: Ref3D. 3D references to cells with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D. 3D references to cells with identical strings. Returns TRUE.');
		// Case #11: Area3D. 3D single-cell ranges with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(Sheet2!A3:A3,Sheet2!A4:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Sheet2!A3:A3,Sheet2!A4:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Area3D. 3D single-cell ranges with identical strings. Returns TRUE.');
		// Case #12: Table. Table references with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(Table1[Column1],Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Table1[Column1],Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Table. Table references with identical strings. Returns TRUE.');
		// Case #13: Date. Dates converted to identical serial numbers or strings. Returns TRUE.
		oParser = new parserFormula('EXACT(DATE(2025,1,1),DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(DATE(2025,1,1),DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Date. Dates converted to identical serial numbers or strings. Returns TRUE.');
		// Case #14: Time. Times converted to identical serial numbers or strings. Returns TRUE.
		oParser = new parserFormula('EXACT(TIME(12,0,0),TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TIME(12,0,0),TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Time. Times converted to identical serial numbers or strings. Returns TRUE.');
		// Case #15: Formula. Nested IF producing identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(IF(TRUE,"Text",""),IF(TRUE,"Text",""))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(IF(TRUE,"Text",""),IF(TRUE,"Text","")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested IF producing identical strings. Returns TRUE.');
		// Case #16: String. Identical uppercase strings. Returns TRUE.
		oParser = new parserFormula('EXACT("TEXT","TEXT")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("TEXT","TEXT") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Identical uppercase strings. Returns TRUE.');
		// Case #17: String. Identical strings with numbers. Returns TRUE.
		oParser = new parserFormula('EXACT("Text1","Text1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("Text1","Text1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Identical strings with numbers. Returns TRUE.');
		// Case #18: Array. Multi-element arrays with identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT({"Text1","Text2"},{"Text1","Text2"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT({"Text1","Text2"},{"Text1","Text2"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Array. Multi-element arrays with identical strings. Returns TRUE.');
		// Case #19: Formula. Nested TEXT function producing identical strings. Returns TRUE.
		oParser = new parserFormula('EXACT(TEXT(123,"0"),TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TEXT(123,"0"),TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula. Nested TEXT function producing identical strings. Returns TRUE.');
		// Case #20: String. Identical strings with single space. Returns TRUE.
		oParser = new parserFormula('EXACT(" "," ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(" "," ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Identical strings with single space. Returns TRUE.');
		// Case #21: String. Identical strings with special characters. Returns TRUE.
		oParser = new parserFormula('EXACT("!@#","!@#")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("!@#","!@#") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String. Identical strings with special characters. Returns TRUE.');

		// Negative cases:
		// Case #1: String. Case-sensitive strings differ. Returns FALSE.
		oParser = new parserFormula('EXACT("Text","text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("Text","text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Case-sensitive strings differ. Returns FALSE.');
		// Case #2: Number. Different numbers converted to strings. Returns FALSE.
		oParser = new parserFormula('EXACT(123,124)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(123,124) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Number. Different numbers converted to strings. Returns FALSE.');
		// Case #3: Error. Propagates #N/A error. Returns #N/A.
		oParser = new parserFormula('EXACT(NA(),NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(NA(),NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. Returns #N/A.');
		// Case #4: Empty. Empty cell references. Returns TRUE (empty strings are equal).
		oParser = new parserFormula('EXACT(A104,A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(A104,A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Empty. Empty cell references. Returns TRUE (empty strings are equal).');
		// Case #5: Boolean. Booleans converted to strings ("TRUE" vs "FALSE"). Returns FALSE.
		oParser = new parserFormula('EXACT(TRUE,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TRUE,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Booleans converted to strings ("TRUE" vs "FALSE"). Returns FALSE.');
		// Case #6: Area. Multi-cell ranges. Returns #VALUE! error.
		oParser = new parserFormula('EXACT(A105:A106,A107:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(A105:A106,A107:A108) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area. Multi-cell ranges. Returns #VALUE! error.');
		// Case #7: String. Strings differ by trailing space. Returns FALSE.
		oParser = new parserFormula('EXACT("Text","Text ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("Text","Text ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String. Strings differ by trailing space. Returns FALSE.');
		// Case #8: Reference link. References to different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(A100,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(A100,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Reference link. References to different strings. Returns FALSE.');
		// Case #9: Name. Named ranges with different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(TestName,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TestName,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name. Named ranges with different strings. Returns FALSE.');
		// Case #10: Ref3D. 3D references to different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(Sheet2!A1,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Sheet2!A1,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Ref3D. 3D references to different strings. Returns FALSE.');
		// Case #11: Area3D. 3D multi-cell ranges. Returns #VALUE! error.
		oParser = new parserFormula('EXACT(Sheet2!A5:A6,Sheet2!A7:A8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Sheet2!A5:A6,Sheet2!A7:A8) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Area3D. 3D multi-cell ranges. Returns #VALUE! error.');
		// Case #12: Table. Table references with different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(Table1[Column1],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Table1[Column1],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Table. Table references with different strings. Returns FALSE.');
		// Case #13: Formula. Nested CONCATENATE producing different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(CONCATENATE("Te","xt"),CONCATENATE("te","xt"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(CONCATENATE("Te","xt"),CONCATENATE("te","xt")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested CONCATENATE producing different strings. Returns FALSE.');
		// Case #14: Date. Different dates converted to strings. Returns FALSE.
		oParser = new parserFormula('EXACT(DATE(2025,1,1),DATE(2025,1,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(DATE(2025,1,1),DATE(2025,1,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Date. Different dates converted to strings. Returns FALSE.');
		// Case #15: Time. Different times converted to strings. Returns FALSE.
		oParser = new parserFormula('EXACT(TIME(12,0,0),TIME(13,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TIME(12,0,0),TIME(13,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Time. Different times converted to strings. Returns FALSE.');
		// Case #16: Formula. Nested IF producing different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(IF(TRUE,"Text",""),IF(FALSE,"","text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(IF(TRUE,"Text",""),IF(FALSE,"","text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Formula. Nested IF producing different strings. Returns FALSE.');
		// Case #17: Name3D. 3D named ranges with different strings. Returns FALSE.
		oParser = new parserFormula('EXACT(TestName3D,TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TestName3D,TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Name3D. 3D named ranges with different strings. Returns FALSE.');
		// Case #18: Array. Arrays with different strings. Returns FALSE.
		oParser = new parserFormula('EXACT({"Text"},{"text"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT({"Text"},{"text"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Array. Arrays with different strings. Returns FALSE.');
		// Case #19: Name. Named ranges with multi-cell data. Returns #VALUE! error.
		oParser = new parserFormula('EXACT(TestNameArea2,TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(TestNameArea2,TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Name. Named ranges with multi-cell data. Returns #VALUE! error.');
		// Case #20: Ref3D. 3D references to cells with error values. Returns #N/A.
		oParser = new parserFormula('EXACT(Sheet2!A9,Sheet2!A10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(Sheet2!A9,Sheet2!A10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Ref3D. 3D references to cells with error values. Returns #N/A.');

		// Bounded cases:
		// Case #1: String. Empty strings (minimum valid input). Returns TRUE.
		oParser = new parserFormula('EXACT("","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String. Empty strings (minimum valid input). Returns TRUE.');
		// Case #2: String. Maximum string length (32,767 characters). Returns TRUE.
		oParser = new parserFormula('EXACT("A"&REPT("Z",32766),"A"&REPT("Z",32766))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT("A"&REPT("Z",32766),"A"&REPT("Z",32766)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String. Maximum string length (32,767 characters). Returns TRUE.');
		// Case #3: String. Non-printable low ASCII character. Returns TRUE.
		oParser = new parserFormula('EXACT(CHAR(1),CHAR(1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(CHAR(1),CHAR(1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String. Non-printable low ASCII character. Returns TRUE.');
		// Case #4: String. High ASCII character. Returns TRUE.
		oParser = new parserFormula('EXACT(CHAR(255),CHAR(255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: EXACT(CHAR(255),CHAR(255)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String. High ASCII character. Returns TRUE.');

		// Need to fix: area handle
		// Case #6: Area. Multi-cell ranges. Returns #VALUE! error.
		// Case #11: Area3D. 3D multi-cell ranges. Returns #VALUE! error.

		testArrayFormula2(assert, "EXACT", 2, 2);
	});

	QUnit.test("Test: \"FIND\"", function (assert) {
		ws.getRange2("A101:F101").cleanAll();
		ws.getRange2("A101").setValue("Miriam McGovern");

		oParser = new parserFormula('FIND("M",A101)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND("M",A101)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'FIND("M",A101)');

		oParser = new parserFormula('FIND("m",A101)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND("m",A101)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'FIND("m",A101)');

		oParser = new parserFormula('FIND("M",A101,3)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND("M",A101,3)');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'FIND("M",A101,3)');

		oParser = new parserFormula('FIND("U",TRUE)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND("T",TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'FIND("T",TRUE)');

		// for bug 68343
		ws.getRange2("B101").setValue("31° 57' 14.6\" S BT 3 18° 54' 20.3\" E");
		oParser = new parserFormula('FIND(""" S",B101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(""" S",B101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'FIND(""" S",B101,1)');

		oParser = new parserFormula('FIND(" S",B101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(" S",B101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 14, 'FIND(" S",B101,1)');

		ws.getRange2("C101").setValue("6\" S");
		oParser = new parserFormula('FIND(""" S",C101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(""" S",C101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'FIND(""" S",C101,1)');

		oParser = new parserFormula('FIND(" S",C101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(" S",C101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'FIND(" S",C101,1)');

		ws.getRange2("D101").setValue("testtest \" String\"abcdString");
		oParser = new parserFormula('FIND(""" String",D101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(""" String",D101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'FIND(""" String",D101,1)');

		oParser = new parserFormula('FIND(" String",D101,1)', "A2", ws);
		assert.ok(oParser.parse(), 'FIND(" String",D101,1)');
		assert.strictEqual(oParser.calculate().getValue(), 11, 'FIND(" String",D101,1)');

		ws.getRange2("A100:D210").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A104").setValue("1.5");
		// For area
		ws.getRange2("A102").setValue("TRUE");
		ws.getRange2("A103").setValue("FALSE");
		ws.getRange2("A105").setValue("Text");
		ws.getRange2("A110").setValue("10");
		ws.getRange2("A111").setValue("20");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		ws.getRange2("C601").setValue("Text"); // Text (Column3)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
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
		// Case #1: String. Case-sensitive match, start_num skips first 7 characters. Returns 9.
		oParser = new parserFormula('FIND("Y","AYF0093.YoungMensApparel",8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("Y","AYF0093.YoungMensApparel",8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, start_num skips first 7 characters. Returns 9.');
		// Case #2: String. Case-sensitive match, start_num 1. Returns 9.
		oParser = new parserFormula('FIND("text","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, start_num 1. Returns 9.');
		// Case #3: Formula. Nested CONCATENATE for find_text, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(CONCATENATE("t","ext"),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(CONCATENATE("t","ext"),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Formula. Nested CONCATENATE for find_text, case-sensitive match. Returns 9.');
		// Case #4: Reference link. References to cells with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(A100,A101,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(A100,A101,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link. References to cells with strings, case-sensitive match. Returns 9.');
		// Case #5: Area. Single-cell ranges, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(A102:A102,A103:A103,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(A102:A102,A103:A103,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell ranges, case-sensitive match. Returns 9.');
		// Case #6: Array. Arrays with single string elements, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND({"text"},{"This is text"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND({"text"},{"This is text"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Array. Arrays with single string elements, case-sensitive match. Returns 9.');
		// Case #7: Name. Named ranges with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(TestName,TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TestName,TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name. Named ranges with strings, case-sensitive match. Returns 9.');
		// Case #8: Name3D. 3D named ranges with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(TestName3D,TestName3D,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TestName3D,TestName3D,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D. 3D named ranges with strings, case-sensitive match. Returns 9.');
		// Case #9: Ref3D. 3D references to cells with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(Sheet2!A1,Sheet2!A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Sheet2!A1,Sheet2!A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D. 3D references to cells with strings, case-sensitive match. Returns 9.');
		// Case #10: Area3D. 3D single-cell ranges, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(Sheet2!A3:A3,Sheet2!A4:A4,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Sheet2!A3:A3,Sheet2!A4:A4,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell ranges, case-sensitive match. Returns 9.');
		// Case #11: Table. Table references with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(Table1[Column1],Table1[Column2],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Table1[Column1],Table1[Column2],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table references with strings, case-sensitive match. Returns 9.');
		// Case #12: String. Empty find_text, matches first character at start_num. Returns 1.
		oParser = new parserFormula('FIND("","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Empty find_text, matches first character at start_num. Returns 1.');
		// Case #13: Formula. Nested IF producing find_text, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(IF(TRUE,"text",""),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(IF(TRUE,"text",""),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Formula. Nested IF producing find_text, case-sensitive match. Returns 9.');
		// Case #14: Number. Number converted to string for find_text, case-sensitive match. Returns 5.
		oParser = new parserFormula('FIND(123,"Text123",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(123,"Text123",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number. Number converted to string for find_text, case-sensitive match. Returns 5.');
		// Case #15: Formula. FIND inside SUM, case-sensitive match. Returns 10.
		oParser = new parserFormula('SUM(FIND("text","This is text",1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(FIND("text","This is text",1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Formula. FIND inside SUM, case-sensitive match. Returns 10.');
		// Case #16: String. Case-sensitive match, different case. Returns 9.
		oParser = new parserFormula('FIND("Text","This is Text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("Text","This is Text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, different case. Returns 9.');
		// Case #17: Date. Date converted to string, case-sensitive match. Returns 7.
		oParser = new parserFormula('FIND(DATE(2025,1,1),"Date: 2025-01-01",7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(DATE(2025,1,1),"Date: 2025-01-01",7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Date. Date converted to string, case-sensitive match. Returns 7.');
		// Case #18: Time. Time converted to string, case-sensitive match. Returns 7.
		oParser = new parserFormula('FIND(TIME(12,0,0),"Time: 12:00:00",7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TIME(12,0,0),"Time: 12:00:00",7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Time. Time converted to string, case-sensitive match. Returns 7.');
		// Case #19: String. Tilde as find_text, case-sensitive match. Returns 5.
		oParser = new parserFormula('FIND("~","Text~",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("~","Text~",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Tilde as find_text, case-sensitive match. Returns 5.');
		// Case #20: String. Case-sensitive match, start_num 1. Returns 2.
		oParser = new parserFormula('FIND("Y","AYF0093.YoungMensApparel",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("Y","AYF0093.YoungMensApparel",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. Case-sensitive match, start_num 1. Returns 2.');
		// Case #21: String. Case-sensitive match, partial string. Returns #VALUE!.
		oParser = new parserFormula('FIND("mens","YoungMensApparel",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("mens","YoungMensApparel",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. Case-sensitive match, partial string. Returns 10.');
		// Case #22: String. Special symbol
		oParser = new parserFormula('FIND("[","[Planning général light (1).xlsx]paramètres")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("[","[Planning général light (1).xlsx]paramètres") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Special symbol.');
		// Case #23: String. Special symbol
		oParser = new parserFormula('FIND("[","[Planning général light (1).xlsx]paramètres",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("[","[Planning général light (1).xlsx]paramètres",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Special symbol.');
		
		// Negative cases:
		// Case #1: String. Case-sensitive mismatch. Returns #VALUE! error.
		oParser = new parserFormula('FIND("text","This is Text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text","This is Text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Case-sensitive mismatch. Returns #VALUE! error.');
		// Case #2: String. find_text not in within_text. Returns #VALUE! error.
		oParser = new parserFormula('FIND("xyz","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("xyz","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. find_text not in within_text. Returns #VALUE! error.');
		// Case #3: Number. find_text not found in within_text. Returns #VALUE! error.
		oParser = new parserFormula('FIND(123,"Text456",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(123,"Text456",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. find_text not found in within_text. Returns #VALUE! error.');
		// Case #4: Number. start_num not greater than 0. Returns #VALUE! error.
		oParser = new parserFormula('FIND("text","This is text",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text","This is text",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. start_num not greater than 0. Returns #VALUE! error.');
		// Case #5: Number. start_num greater than within_text length (11). Returns #VALUE! error.
		oParser = new parserFormula('FIND("text","This is text",12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text","This is text",12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. start_num greater than within_text length (11). Returns #VALUE! error.');
		// Case #6: Error. find_text is #N/A. Returns #N/A.
		oParser = new parserFormula('FIND(NA(),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(NA(),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. find_text is #N/A. Returns #N/A.');
		// Case #7: Error. within_text is #N/A. Returns #N/A.
		oParser = new parserFormula('FIND("text",NA(),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text",NA(),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. within_text is #N/A. Returns #N/A.');
		// Case #8: Empty. Empty find_text cell, matches first character at start_num. Returns 1.
		oParser = new parserFormula('FIND(A104,A105,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(A104,A105,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty find_text cell, matches first character at start_num. Returns 1.');
		// Case #9: Empty. Empty within_text cell. Returns #VALUE! error.
		oParser = new parserFormula('FIND("text",A106,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text",A106,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty within_text cell. Returns #VALUE! error.');
		// Case #10: Boolean. Boolean find_text converted to string, case-sensitive match. Returns 9.
		oParser = new parserFormula('FIND(TRUE,"This is TRUE",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TRUE,"This is TRUE",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Negative case: Boolean. Boolean find_text converted to string, case-sensitive match. Returns 9.');
		// Case #11: Area. Multi-cell ranges for find_text and within_text. Returns #VALUE! error.
		oParser = new parserFormula('FIND(A107:A108,A109:A110,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(A107:A108,A109:A110,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell ranges for find_text and within_text. Returns #VALUE! error.');
		// Case #12: Name. Multi-cell named ranges. Returns #VALUE! error.
		oParser = new parserFormula('FIND(TestNameArea2,TestNameArea2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TestNameArea2,TestNameArea2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Multi-cell named ranges. Returns #VALUE! error.');
		// Case #13: Name3D. Multi-cell 3D named range for find_text. Returns #VALUE! error.
		oParser = new parserFormula('FIND(TestNameArea3D2,TestName3D,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(TestNameArea3D2,TestName3D,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. Multi-cell 3D named range for find_text. Returns #VALUE! error.');
		// Case #14: Ref3D. 3D references with non-matching strings. Returns #VALUE! error.
		oParser = new parserFormula('FIND(Sheet2!A5,Sheet2!A6,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Sheet2!A5,Sheet2!A6,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Ref3D. 3D references with non-matching strings. Returns #VALUE! error.');
		// Case #15: Area3D. Multi-cell 3D ranges. Returns #VALUE! error.
		oParser = new parserFormula('FIND(Sheet2!A7:A8,Sheet2!A9:A10,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Sheet2!A7:A8,Sheet2!A9:A10,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. Multi-cell 3D ranges. Returns #VALUE! error.');
		// Case #16: Table. Table references with non-matching strings. Returns #VALUE! error.
		oParser = new parserFormula('FIND(Table1[Column1],Table1[Column3],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(Table1[Column1],Table1[Column3],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table references with non-matching strings. Returns #VALUE! error.');
		// Case #17: Formula. Nested formula producing #NUM! error. Returns #NUM! error.
		oParser = new parserFormula('FIND(SQRT(-1),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND(SQRT(-1),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula producing #NUM! error. Returns #NUM! error.');
		// Case #18: String. Wildcard (*) not allowed, treated as literal. Returns 12.
		oParser = new parserFormula('FIND("*","This is text*",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("*","This is text*",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Negative case: String. Wildcard (*) not allowed, treated as literal. Returns 12.');
		// Case #19: String. Wildcard (?) not allowed, treated as literal. Returns 12.
		oParser = new parserFormula('FIND("?","This is text?",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("?","This is text?",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Negative case: String. Wildcard (?) not allowed, treated as literal. Returns 12.');
		// Case #20: Number. Negative start_num. Returns #VALUE! error.
		oParser = new parserFormula('FIND("text","This is text",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("text","This is text",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Negative start_num. Returns #VALUE! error.');

		// Bounded cases:
		// Case #1: String. Empty find_text, minimum valid input. Returns 1.
		oParser = new parserFormula('FIND("","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Empty find_text, minimum valid input. Returns 1.');
		// Case #2: String. Maximum within_text length (32,767 characters). Returns 1.
		oParser = new parserFormula('FIND("A","A"&REPT("Z",32766),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("A","A"&REPT("Z",32766),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Maximum within_text length (32,767 characters). Returns 1.');
		// Case #3: Number. Maximum valid start_num for within_text length (11). Returns #VALUE! error.
		oParser = new parserFormula('FIND("A","This is text",32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("A","This is text",32767) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Number. Maximum valid start_num for within_text length (11). Returns #VALUE! error.');
		// Case #4: String. Maximum find_text and within_text length (32,767 characters). Returns 1.
		oParser = new parserFormula('FIND("A"&REPT("Z",32766),"A"&REPT("Z",32766),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIND("A"&REPT("Z",32766),"A"&REPT("Z",32766),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Maximum find_text and within_text length (32,767 characters). Returns 1.');


		testArrayFormula2(assert, "FIND", 2, 3);
	});

	QUnit.test("Test: \"FINDB\"", function (assert) {
		ws.getRange2("A101").setValue("Miriam McGovern");

		oParser = new parserFormula('FINDB("M",A101)', "A2", ws);
		assert.ok(oParser.parse(), 'FINDB("M",A101)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'FINDB("M",A101)');

		oParser = new parserFormula('FINDB("m",A101)', "A2", ws);
		assert.ok(oParser.parse(), 'FINDB("m",A101)');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'FINDB("m",A101)');

		oParser = new parserFormula('FINDB("M",A101,3)', "A2", ws);
		assert.ok(oParser.parse(), 'FINDB("M",A101,3)');
		assert.strictEqual(oParser.calculate().getValue(), 8, 'FINDB("M",A101,3)');

		oParser = new parserFormula('FINDB("U",TRUE)', "A2", ws);
		assert.ok(oParser.parse(), 'FINDB("T",TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'FINDB("T",TRUE)');

		ws.getRange2("A100:D210").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("1");
		ws.getRange2("A101").setValue("2");
		ws.getRange2("A104").setValue("1.5");
		// For area
		ws.getRange2("A102").setValue("TRUE");
		ws.getRange2("A103").setValue("FALSE");
		ws.getRange2("A105").setValue("Text");
		ws.getRange2("A106").setValue("10");
		ws.getRange2("A107").setValue("20");
		ws.getRange2("A108").setValue("30");
		ws.getRange2("A109").setValue("40");
		ws.getRange2("A110").setValue("50");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		ws.getRange2("C601").setValue("Text"); // Text (Column3)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:C10").cleanAll();
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
		// Case #1: String. Case-sensitive match, start_num skips first 7 characters. Returns 9.
		oParser = new parserFormula('FINDB("Y","AYF0093.YoungMensApparel",8)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("Y","AYF0093.YoungMensApparel",8) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, start_num skips first 7 characters. Returns 9.');
		// Case #2: String. Case-sensitive match, start_num 1. Returns 9.
		oParser = new parserFormula('FINDB("text","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, start_num 1. Returns 9.');
		// Case #3: Formula. Nested CONCATENATE for find_text, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(CONCATENATE("t","ext"),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(CONCATENATE("t","ext"),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Formula. Nested CONCATENATE for find_text, case-sensitive match. Returns 9.');
		// Case #4: Reference link. References to cells with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(A100,A101,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(A100,A101,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link. References to cells with strings, case-sensitive match. Returns 9.');
		// Case #5: Area. Single-cell ranges, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(A102:A102,A103:A103,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(A102:A102,A103:A103,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area. Single-cell ranges, case-sensitive match. Returns 9.');
		// Case #6: Array. Arrays with single string elements, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB({"text"},{"This is text"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB({"text"},{"This is text"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Array. Arrays with single string elements, case-sensitive match. Returns 9.');
		// Case #7: Name. Named ranges with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(TestName,TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TestName,TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name. Named ranges with strings, case-sensitive match. Returns 9.');
		// Case #8: Name3D. 3D named ranges with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(TestName3D,TestName3D,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TestName3D,TestName3D,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D. 3D named ranges with strings, case-sensitive match. Returns 9.');
		// Case #9: Ref3D. 3D references to cells with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(Sheet2!A1,Sheet2!A2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Sheet2!A1,Sheet2!A2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D. 3D references to cells with strings, case-sensitive match. Returns 9.');
		// Case #10: Area3D. 3D single-cell ranges, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(Sheet2!A3:A3,Sheet2!A4:A4,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Sheet2!A3:A3,Sheet2!A4:A4,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D. 3D single-cell ranges, case-sensitive match. Returns 9.');
		// Case #11: Table. Table references with strings, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(Table1[Column1],Table1[Column2],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Table1[Column1],Table1[Column2],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table references with strings, case-sensitive match. Returns 9.');
		// Case #12: String. Empty find_text, matches first character at start_num. Returns 1.
		oParser = new parserFormula('FINDB("","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Empty find_text, matches first character at start_num. Returns 1.');
		// Case #13: Formula. Nested IF producing find_text, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(IF(TRUE,"text",""),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(IF(TRUE,"text",""),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: Formula. Nested IF producing find_text, case-sensitive match. Returns 9.');
		// Case #14: Number. Number converted to string for find_text, case-sensitive match. Returns 5.
		oParser = new parserFormula('FINDB(123,"Text123",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(123,"Text123",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Number. Number converted to string for find_text, case-sensitive match. Returns 5.');
		// Case #15: Formula. FIND inside SUM, case-sensitive match. Returns 10.
		oParser = new parserFormula('SUM(FINDB("text","This is text",1),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(FINDB("text","This is text",1),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 10, 'Test: Positive case: Formula. FIND inside SUM, case-sensitive match. Returns 10.');
		// Case #16: String. Case-sensitive match, different case. Returns 9.
		oParser = new parserFormula('FINDB("Text","This is Text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("Text","This is Text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Positive case: String. Case-sensitive match, different case. Returns 9.');
		// Case #17: Date. Date converted to string, case-sensitive match. Returns 7.
		oParser = new parserFormula('FINDB(DATE(2025,1,1),"Date: 2025-01-01",7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(DATE(2025,1,1),"Date: 2025-01-01",7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Date. Date converted to string, case-sensitive match. Returns 7.');
		// Case #18: Time. Time converted to string, case-sensitive match. Returns 7.
		oParser = new parserFormula('FINDB(TIME(12,0,0),"Time: 12:00:00",7)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TIME(12,0,0),"Time: 12:00:00",7) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Time. Time converted to string, case-sensitive match. Returns 7.');
		// Case #19: String. Tilde as find_text, case-sensitive match. Returns 5.
		oParser = new parserFormula('FINDB("~","Text~",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("~","Text~",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Tilde as find_text, case-sensitive match. Returns 5.');
		// Case #20: String. Case-sensitive match, start_num 1. Returns 2.
		oParser = new parserFormula('FINDB("Y","AYF0093.YoungMensApparel",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("Y","AYF0093.YoungMensApparel",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. Case-sensitive match, start_num 1. Returns 2.');
		// Case #21: String. Case-sensitive match, partial string. Returns 10.
		oParser = new parserFormula('FINDB("mens","YoungMensApparel",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("mens","YoungMensApparel",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String. Case-sensitive match, partial string. Returns 10.');
		// Case #22: String. Special symbol
		oParser = new parserFormula('FINDB("[","[Planning général light (1).xlsx]paramètres")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("[","[Planning général light (1).xlsx]paramètres") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Special symbol.');
		// Case #23: String. Special symbol
		oParser = new parserFormula('FINDB("[","[Planning général light (1).xlsx]paramètres",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("[","[Planning général light (1).xlsx]paramètres",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Special symbol.');

		// Negative cases:
		// Case #1: String. Case-sensitive mismatch. Returns #VALUE! error.
		oParser = new parserFormula('FINDB("text","This is Text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text","This is Text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Case-sensitive mismatch. Returns #VALUE! error.');
		// Case #2: String. find_text not in within_text. Returns #VALUE! error.
		oParser = new parserFormula('FINDB("xyz","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("xyz","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. find_text not in within_text. Returns #VALUE! error.');
		// Case #3: Number. find_text not found in within_text. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(123,"Text456",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(123,"Text456",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. find_text not found in within_text. Returns #VALUE! error.');
		// Case #4: Number. start_num not greater than 0. Returns #VALUE! error.
		oParser = new parserFormula('FINDB("text","This is text",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text","This is text",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. start_num not greater than 0. Returns #VALUE! error.');
		// Case #5: Number. start_num greater than within_text length (11). Returns #VALUE! error.
		oParser = new parserFormula('FINDB("text","This is text",12)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text","This is text",12) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. start_num greater than within_text length (11). Returns #VALUE! error.');
		// Case #6: Error. find_text is #N/A. Returns #N/A.
		oParser = new parserFormula('FINDB(NA(),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(NA(),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. find_text is #N/A. Returns #N/A.');
		// Case #7: Error. within_text is #N/A. Returns #N/A.
		oParser = new parserFormula('FINDB("text",NA(),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text",NA(),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. within_text is #N/A. Returns #N/A.');
		// Case #8: Empty. Empty find_text cell, matches first character at start_num. Returns 1.
		oParser = new parserFormula('FINDB(A104,A105,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(A104,A105,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty find_text cell, matches first character at start_num. Returns 1.');
		// Case #9: Empty. Empty within_text cell. Returns #VALUE! error.
		oParser = new parserFormula('FINDB("text",A106,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text",A106,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty within_text cell. Returns #VALUE! error.');
		// Case #10: Boolean. Boolean find_text converted to string, case-sensitive match. Returns 9.
		oParser = new parserFormula('FINDB(TRUE,"This is TRUE",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TRUE,"This is TRUE",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 9, 'Test: Negative case: Boolean. Boolean find_text converted to string, case-sensitive match. Returns 9.');
		// Case #11: Area. Multi-cell ranges for find_text and within_text. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(A107:A108,A109:A110,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(A107:A108,A109:A110,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell ranges for find_text and within_text. Returns #VALUE! error.');
		// Case #12: Name. Multi-cell named ranges. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(TestNameArea2,TestNameArea2,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TestNameArea2,TestNameArea2,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Multi-cell named ranges. Returns #VALUE! error.');
		// Case #13: Name3D. Multi-cell 3D named range for find_text. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(TestNameArea3D2,TestName3D,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(TestNameArea3D2,TestName3D,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. Multi-cell 3D named range for find_text. Returns #VALUE! error.');
		// Case #14: Ref3D. 3D references with non-matching strings. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(Sheet2!A5,Sheet2!A6,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Sheet2!A5,Sheet2!A6,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Ref3D. 3D references with non-matching strings. Returns #VALUE! error.');
		// Case #15: Area3D. Multi-cell 3D ranges. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(Sheet2!A7:A8,Sheet2!A9:A10,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Sheet2!A7:A8,Sheet2!A9:A10,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Test: Negative case: Area3D. Multi-cell 3D ranges. Returns #VALUE! error.');
		// Case #16: Table. Table references with non-matching strings. Returns #VALUE! error.
		oParser = new parserFormula('FINDB(Table1[Column1],Table1[Column3],1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(Table1[Column1],Table1[Column3],1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table references with non-matching strings. Returns #VALUE! error.');
		// Case #17: Formula. Nested formula producing #NUM! error. Returns #NUM! error.
		oParser = new parserFormula('FINDB(SQRT(-1),"This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB(SQRT(-1),"This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula producing #NUM! error. Returns #NUM! error.');
		// Case #18: String. Wildcard (*) not allowed, treated as literal. Returns 12.
		oParser = new parserFormula('FINDB("*","This is text*",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("*","This is text*",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Negative case: String. Wildcard (*) not allowed, treated as literal. Returns 12.');
		// Case #19: String. Wildcard (?) not allowed, treated as literal. Returns 12.
		oParser = new parserFormula('FINDB("?","This is text?",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("?","This is text?",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 13, 'Test: Negative case: String. Wildcard (?) not allowed, treated as literal. Returns 12.');
		// Case #20: Number. Negative start_num. Returns #VALUE! error.
		oParser = new parserFormula('FINDB("text","This is text",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("text","This is text",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Negative start_num. Returns #VALUE! error.');

		// Bounded cases:
		// Case #1: String. Empty find_text, minimum valid input. Returns 1.
		oParser = new parserFormula('FINDB("","This is text",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("","This is text",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Empty find_text, minimum valid input. Returns 1.');
		// Case #2: String. Maximum within_text length (32,767 characters). Returns 1.
		oParser = new parserFormula('FINDB("A","A"&REPT("Z",32766),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("A","A"&REPT("Z",32766),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Maximum within_text length (32,767 characters). Returns 1.');
		// Case #3: Number. Maximum valid start_num for within_text length (11). Returns #VALUE! error.
		oParser = new parserFormula('FINDB("A","This is text",32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("A","This is text",32767) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Number. Maximum valid start_num for within_text length (11). Returns #VALUE! error.');
		// Case #4: String. Maximum find_text and within_text length (32,767 characters). Returns 1.
		oParser = new parserFormula('FINDB("A"&REPT("Z",32766),"A"&REPT("Z",32766),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FINDB("A"&REPT("Z",32766),"A"&REPT("Z",32766),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Maximum find_text and within_text length (32,767 characters). Returns 1.');


	});

	QUnit.test("Test: \"FIXED\"", function (assert) {
		oParser = new parserFormula("FIXED(1234567,-3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1,235,000");
		oParser = new parserFormula("FIXED(.555555,10)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "0.5555550000");
		oParser = new parserFormula("FIXED(1234567.555555,4,TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1234567.5556");
		oParser = new parserFormula("FIXED(1234567)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1,234,567.00");

		// 	Data for reference link. Use A100-A111
		ws.getRange2("A1:C214").cleanAll();
		ws.getRange2("A100").setValue("0.1");
		ws.getRange2("A101").setValue("0.2");
		ws.getRange2("A104").setValue("0.5");
		// For area
		ws.getRange2("A102").setValue("0.3");
		ws.getRange2("A103").setValue("0.4");
		ws.getRange2("A105").setValue("0.6");
		ws.getRange2("A106").setValue("0.7");
		ws.getRange2("A107").setValue("0.8");
		ws.getRange2("A108").setValue("0.9");
		ws.getRange2("A109").setValue("1");
		ws.getRange2("A110").setValue("2");

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
		ws.getRange2("B208").setValue("-0.8"); // TestNameArea2
		ws2.getRange2("A18").setValue("0.8"); // TestNameArea3D2
		ws2.getRange2("B18").setValue("-0.8"); // TestNameArea3D2


		// Positive cases:
		// Case #1: Number. Basic valid input: number with 2 decimals, commas included. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Positive case: Number. Basic valid input: number with 2 decimals, commas included. 2 of 3 arguments used.');
		// Case #2: Number. Valid input with no commas. 3 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,2,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,2,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Positive case: Number. Valid input with no commas. 3 of 3 arguments used.');
		// Case #3: Number. Negative number with 2 decimals. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(-123.456,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(-123.456,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-123.46', 'Test: Positive case: Number. Negative number with 2 decimals. 2 of 3 arguments used.');
		// Case #4: String. String convertible to valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED("123.456",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED("123.456",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Positive case: String. String convertible to valid number. 2 of 3 arguments used.');
		// Case #5: Formula. Nested formula resolving to valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(SQRT(10000),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(SQRT(10000),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '100.0', 'Test: Positive case: Formula. Nested formula resolving to valid number. 2 of 3 arguments used.');
		// Case #6: Formula. Nested IF formula for number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(IF(TRUE,123.456,-123.456),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(IF(TRUE,123.456,-123.456),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Positive case: Formula. Nested IF formula for number. 2 of 3 arguments used.');
		// Case #7: Reference link. Reference link to valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.10', 'Test: Positive case: Reference link. Reference link to valid number. 2 of 3 arguments used.');
		// Case #8: Reference link. Reference links for number and decimals, no commas. 3 of 3 arguments used.
		oParser = new parserFormula('FIXED(A100,A101,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A100,A101,TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.0', 'Test: Positive case: Reference link. Reference links for number and decimals, no commas. 3 of 3 arguments used.');
		// Case #9: Area. Single-cell range for number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(A100:A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A100:A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.10', 'Test: Positive case: Area. Single-cell range for number. 2 of 3 arguments used.');
		// Case #10: Array. Array with single valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED({123.456},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED({123.456},2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Positive case: Array. Array with single valid number. 2 of 3 arguments used.');
		// Case #11: Name. Named range with valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TestName,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TestName,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.50', 'Test: Positive case: Name. Named range with valid number. 2 of 3 arguments used.');
		// Case #12: Name3D. 3D named range with valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.50', 'Test: Positive case: Name3D. 3D named range with valid number. 2 of 3 arguments used.');
		// Case #13: Ref3D. 3D reference to valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(Sheet2!A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(Sheet2!A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50', 'Test: Positive case: Ref3D. 3D reference to valid number. 2 of 3 arguments used.');
		// Case #14: Area3D. 3D single-cell range for number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(Sheet2!A1:A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(Sheet2!A1:A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50', 'Test: Positive case: Area3D. 3D single-cell range for number. 2 of 3 arguments used.');
		// Case #15: Table. Table structured reference with valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(Table1[Column1],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(Table1[Column1],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.01', 'Test: Positive case: Table. Table structured reference with valid number. 2 of 3 arguments used.');
		// Case #16: Date. Date as serial number with 0 decimals. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(DATE(2025,1,1),0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(DATE(2025,1,1),0) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '45.658', 'Test: Positive case: Date. Date as serial number with 0 decimals. 2 of 3 arguments used.');
		// Case #17: Time. Time adjusted to valid number. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TIME(12,0,0)*1000,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TIME(12,0,0)*1000,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '500.0', 'Test: Positive case: Time. Time adjusted to valid number. 2 of 3 arguments used.');
		// Case #18: Formula. FIXED inside SUM formula. 2 of 3 arguments used.
		oParser = new parserFormula('SUM(FIXED(123.456,2),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(FIXED(123.456,2),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 124.46, 'Test: Positive case: Formula. FIXED inside SUM formula. 2 of 3 arguments used.');
		// Case #19: Number. Negative decimals, rounds to left of decimal point. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '100', 'Test: Positive case: Number. Negative decimals, rounds to left of decimal point. 2 of 3 arguments used.');
		// Case #20: Number. Maximum decimals (127). 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,127)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,127) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123.456', 'Test: Positive case: Number. Maximum decimals (127). 2 of 3 arguments used.');
		// Case #21: Reference link. Reference links with commas included. 3 of 3 arguments used.
		oParser = new parserFormula('FIXED(A100,A101,FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A100,A101,FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.0', 'Test: Positive case: Reference link. Reference links with commas included. 3 of 3 arguments used.');

		// Negative cases:
		// Case #1: String. Nonnumeric string returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED("abc",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED("abc",2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Nonnumeric string returns #VALUE!. 2 of 3 arguments used.');
		// Case #2: Error. Propagates #N/A error. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(NA(),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(NA(),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 2 of 3 arguments used.');
		// Case #3: Empty. Empty reference link returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(A102,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A102,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.30', 'Test: Negative case: Empty. Empty reference link returns #VALUE!. 2 of 3 arguments used.');
		// Case #4: String. Empty string returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED("",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED("",2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 2 of 3 arguments used.');
		// Case #5: Boolean. Boolean TRUE returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TRUE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TRUE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.00', 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 2 of 3 arguments used.');
		// Case #6: Area. Multi-cell range returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(A100:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(A100:A101,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.1', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 2 of 3 arguments used.');
		// Case #7: Array. Multi-element array returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED({123.456,456.789},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED({123.456,456.789},2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Negative case: Array. Multi-element array returns #VALUE!. 2 of 3 arguments used.');
		// Case #8: Ref3D. 3D reference to nonnumeric value returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(Sheet2!A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(Sheet2!A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.50', 'Test: Negative case: Ref3D. 3D reference to nonnumeric value returns #VALUE!. 2 of 3 arguments used.');
		// Case #9: Name. Named range with nonnumeric value returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TestNameArea2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TestNameArea2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.80', 'Test: Negative case: Name. Named range with nonnumeric value returns #VALUE!. 2 of 3 arguments used.');
		// Case #10: Name3D. 3D named range with nonnumeric value returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(TestNameArea3D2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(TestNameArea3D2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.80', 'Test: Negative case: Name3D. 3D named range with nonnumeric value returns #VALUE!. 2 of 3 arguments used.');
		// Case #11: Table. Table column with nonnumeric value returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(Table1[Column2],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(Table1[Column2],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with nonnumeric value returns #VALUE!. 2 of 3 arguments used.');
		// Case #12: Formula. Formula resulting in #NUM! error. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(SQRT(-1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(SQRT(-1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error. 2 of 3 arguments used.');
		// Case #13: String. Nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,"abc") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.');
		// Case #14: Number,String. Non-boolean no_commas returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,2,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,2,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number,String. Non-boolean no_commas returns #VALUE!. 3 of 3 arguments used.');
		// Case #15: Empty. Empty decimals treated as 2, valid output. 3 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,,TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Negative case: Empty. Empty decimals treated as 2, valid output. 3 of 3 arguments used.');
		// Case #16: Area. Multi-cell range for decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,A100:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,A100:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Negative case: Area. Multi-cell range for decimals returns #VALUE!. 2 of 3 arguments used.');
		// Case #17: Array. Multi-element array for decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,{2,3})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,{2,3}) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123.46', 'Test: Negative case: Array. Multi-element array for decimals returns #VALUE!. 2 of 3 arguments used.');
		// Case #18: Ref3D. 3D reference to nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.50', 'Test: Negative case: Ref3D. 3D reference to nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.');
		// Case #19: Name. Named range with nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.0', 'Test: Negative case: Name. Named range with nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.');
		// Case #20: Name3D. 3D named range with nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.0', 'Test: Negative case: Name3D. 3D named range with nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.');

		// Bounded cases:
		// Case #1: Number. Minimum positive number with 2 decimals. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(1E-307,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(1E-307,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.00', 'Test: Bounded case: Number. Minimum positive number with 2 decimals. 2 of 3 arguments used.');
		// Case #2: Number. Minimum negative number with 2 decimals. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(-1E-307,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(-1E-307,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.00', 'Test: Bounded case: Number. Minimum negative number with 2 decimals. 2 of 3 arguments used.');
		// Case #3: Number. Maximum 15 significant digits. 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123456789012345,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123456789012345,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '123,456,789,012,345.00', 'Test: Bounded case: Number. Maximum 15 significant digits. 2 of 3 arguments used.');
		// Case #4: Number. Maximum decimals (127). 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123.456,127)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123.456,127) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 123.456, 'Test: Bounded case: Number. Maximum decimals (127). 2 of 3 arguments used.');
		// Case #5: Number. Minimum decimals (-127). 2 of 3 arguments used.
		oParser = new parserFormula('FIXED(123456789012345,-127)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: FIXED(123456789012345,-127) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0', 'Test: Bounded case: Number. Minimum decimals (-127). 2 of 3 arguments used.');

		// TODO проблема с округлением чисел
		// Need to fix: error type diff, MS result diff
		// Case #10: Array. Array with single valid number. 2 of 3 arguments used.
		// Case #16: Date. Date as serial number with 0 decimals. 2 of 3 arguments used.
		// Case #20: Number. Maximum decimals (127). 2 of 3 arguments used.
		// Case #1: String. Nonnumeric string returns #VALUE!. 2 of 3 arguments used.
		// Case #4: String. Empty string returns #VALUE!. 2 of 3 arguments used.
		// Case #6: Area. Multi-cell range returns #VALUE!. 2 of 3 arguments used.
		// Case #7: Array. Multi-element array returns #VALUE!. 2 of 3 arguments used.
		// Case #15: Empty. Empty decimals treated as 2, valid output. 3 of 3 arguments used.
		// Case #16: Area. Multi-cell range for decimals returns #VALUE!. 2 of 3 arguments used.
		// Case #13: String. Nonnumeric decimals returns #VALUE!. 2 of 3 arguments used.
		// Case #17: Array. Multi-element array for decimals returns #VALUE!. 2 of 3 arguments used.
		// Case #3: Number. Maximum 15 significant digits. 2 of 3 arguments used. - diff res
		// Case #4: Number. Maximum decimals (127). 2 of 3 arguments used.


		testArrayFormula2(assert, "FIXED", 2, 3);
	});

	QUnit.test("Test: \"LEFT\"", function (assert) {

		ws.getRange2("A2").setValue("Sale Price");
		ws.getRange2("A3").setValue("Sweden");


		oParser = new parserFormula("LEFT(A2,4)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Sale");

		oParser = new parserFormula("LEFT(A3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "S");

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("Text");
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
		// Case #1: String, Number. Basic usage: extract first 2 characters from string.
		oParser = new parserFormula('LEFT("Excel",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Excel",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Ex', 'Test: Positive case: String, Number. Basic usage: extract first 2 characters from string.');
		// Case #3: Number, Number. Number auto-converted to string.
		oParser = new parserFormula('LEFT(12345,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(12345,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "123", 'Test: Positive case: Number, Number. Number auto-converted to string.');
		// Case #4: String, Formula. Dynamic num_chars via formula.
		oParser = new parserFormula('LEFT("Test",LEN("Test")-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Test",LEN("Test")-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Tes', 'Test: Positive case: String, Formula. Dynamic num_chars via formula.');
		// Case #5: Reference link, Number. Text input via reference (A100="Data").
		oParser = new parserFormula('LEFT(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Positive case: Reference link, Number. Text input via reference (A100="Data").');
		// Case #6: Area, Number. Single-cell area (A101="Hello").
		oParser = new parserFormula('LEFT(A100:A100,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(A100:A100,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0.5", 'Test: Positive case: Area, Number. Single-cell area (A101="Hello").');
		// Case #7: Array, Number. Array input.
		oParser = new parserFormula('LEFT({"Array"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT({"Array"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Array, Number. Array input.');
		// Case #8: Name, Number. Named range (TestName="Alpha").
		oParser = new parserFormula('LEFT(TestName,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TestName,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name, Number. Named range (TestName="Alpha").');
		// Case #9: Name3D, Number. 3D named range (TestName1="Beta").
		oParser = new parserFormula('LEFT(TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0', 'Test: Positive case: Name3D, Number. 3D named range (TestName1="Beta").');
		// Case #10: Ref3D, Number. 3D reference.
		oParser = new parserFormula('LEFT(Sheet2!A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(Sheet2!A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Positive case: Ref3D, Number. 3D reference.');
		// Case #11: Area3D, Number. 3D area.
		oParser = new parserFormula('LEFT(Sheet2!A1:A1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(Sheet2!A1:A1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Area3D, Number. 3D area.');
		// Case #12: Table, Number. Table column input.
		oParser = new parserFormula('LEFT(Table1[Column1],3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(Table1[Column1],3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table, Number. Table column input.');
		// Case #14: Formula, Number. Nested function.
		oParser = new parserFormula('LEFT(UPPER("case"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(UPPER("case"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'CA', 'Test: Positive case: Formula, Number. Nested function.');
		// Case #15: Time, Number. Time converted to text.
		oParser = new parserFormula('LEFT(TEXT(TIME(12,30,0),"hh:mm"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TEXT(TIME(12,30,0),"hh:mm"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "12", 'Test: Positive case: Time, Number. Time converted to text.');
		// Case #16: Date, Number. Date converted to text.
		oParser = new parserFormula('LEFT(TEXT(DATE(2025,1,1),"yyyy"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TEXT(DATE(2025,1,1),"yyyy"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "20", 'Test: Positive case: Date, Number. Date converted to text.');
		// Case #20: String, Boolean. Boolean num_chars coerced to 1.
		oParser = new parserFormula('LEFT("Boolean",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Boolean",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Positive case: String, Boolean. Boolean num_chars coerced to 1.');

		// Negative cases:

		// Case #2: Error, Number. Error input propagates.
		oParser = new parserFormula('LEFT(#N/A,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(#N/A,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Error input propagates.');
		// Case #5: Boolean, Number. Boolean text input returns "TRUE".
		oParser = new parserFormula('LEFT(TRUE,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TRUE,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'T', 'Test: Negative case: Boolean, Number. Boolean text input returns "TRUE".');
		// Case #6: String, Array. Array num_chars returns #VALUE!.
		oParser = new parserFormula('LEFT("Text",{1,2})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Text",{1,2}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'T', 'Test: Negative case: String, Array. Array num_chars returns #VALUE!.');
		// Case #8: Name3D, Error. Error in 3D named range.
		oParser = new parserFormula('LEFT(TestName2,#N/A)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(TestName2,#N/A) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name3D, Error. Error in 3D named range.');
		// Case #10: Table, Boolean. Boolean num_chars in table.
		oParser = new parserFormula('LEFT(Table1[Column1],TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(Table1[Column1],TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Table, Boolean. Boolean num_chars in table.');

		// Bounded cases:
		// Case #1: String, Max Number
		oParser = new parserFormula('LEFT("Max",2^20)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Max",2^20) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Max', 'Test: Bounded case: String, Max Number.');
		// Case #2: String, Min Number
		oParser = new parserFormula('LEFT("Min",1E-100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT("Min",1E-100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: String, Min Number.');

		let longStr = "X";
		longStr = longStr.repeat(32767);
		// Case #3: Large String, Number
		oParser = new parserFormula('LEFT(REPT("X",32767),32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFT(REPT("X",32767),32767) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), longStr, 'Test: Bounded case: Large String, Number.');


		testArrayFormula2(assert, "LEFT", 1, 2);
	});

	QUnit.test("Test: \"LEFTB\"", function (assert) {
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
		// Case #1: String, Number. Basic string with valid byte count. 2 arguments used.
		oParser = new parserFormula('LEFTB("abcdef",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abcdef",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String, Number. Basic string with valid byte count. 2 arguments used.');
		// Case #2: String. String with omitted num_bytes (defaults to 1). 1 argument used.
		oParser = new parserFormula('LEFTB("abcdef")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abcdef") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: String. String with omitted num_bytes (defaults to 1). 1 argument used.');
		// Case #3: Number, Number. Numeric string converted to text. 2 arguments used.
		oParser = new parserFormula('LEFTB("12345",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("12345",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Number, Number. Numeric string converted to text. 2 arguments used.');
		// Case #4: Formula, Number. Nested formula resolving to string. 2 arguments used.
		oParser = new parserFormula('LEFTB(CONCAT("ab","cd"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(CONCAT("ab","cd"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ab', 'Test: Positive case: Formula, Number. Nested formula resolving to string. 2 arguments used.');
		// Case #5: Reference link, Number. Reference to cell with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(A100,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(A100,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link, Number. Reference to cell with valid string. 2 arguments used.');
		// Case #6: Area, Number. Single-cell range with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(A101:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(A101:A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.', 'Test: Positive case: Area, Number. Single-cell range with valid string. 2 arguments used.');
		// Case #7: Array, Number. Array with single string element. 2 arguments used.
		oParser = new parserFormula('LEFTB({"abc"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB({"abc"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Array, Number. Array with single string element. 2 arguments used.');
		// Case #8: Name, Number. Named range with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(TestName,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(TestName,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0', 'Test: Positive case: Name, Number. Named range with valid string. 2 arguments used.');
		// Case #9: Name3D, Number. 3D named range with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0', 'Test: Positive case: Name3D, Number. 3D named range with valid string. 2 arguments used.');
		// Case #10: Ref3D, Number. 3D reference to cell with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(Sheet2!A1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Sheet2!A1,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D, Number. 3D reference to cell with valid string. 2 arguments used.');
		// Case #11: Area3D, Number. 3D single-cell range with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(Sheet2!A1:A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Sheet2!A1:A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Area3D, Number. 3D single-cell range with valid string. 2 arguments used.');
		// Case #12: Table, Number. Table structured reference with valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(Table1[Column1],3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Table1[Column1],3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table, Number. Table structured reference with valid string. 2 arguments used.');
		// Case #13: Date, Number. Date serial number converted to string. 2 arguments used.
		oParser = new parserFormula('LEFTB(DATE(2025,1,1),4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(DATE(2025,1,1),4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '4565', 'Test: Positive case: Date, Number. Date serial number converted to string. 2 arguments used.');
		// Case #14: Time, Number. Time serial number converted to string. 2 arguments used.
		oParser = new parserFormula('LEFTB(TIME(12,0,0),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(TIME(12,0,0),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Positive case: Time, Number. Time serial number converted to string. 2 arguments used.');
		// Case #15: Formula, Number. LEFTB inside CONCAT formula. 2 arguments used.
		oParser = new parserFormula('CONCAT(LEFTB("xyz",2),"!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT(LEFTB("xyz",2),"!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xy!', 'Test: Positive case: Formula, Number. LEFTB inside CONCAT formula. 2 arguments used.');
		// Case #16: String, Number. Multi-byte Unicode string (Chinese, extracts 1 character = 2 bytes). 2 arguments used.
		oParser = new parserFormula('LEFTB("??",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("??",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??', 'Test: Positive case: String, Number. Multi-byte Unicode string (Chinese, extracts 1 character = 2 bytes). 2 arguments used.');
		// Case #17: String, Number. Emoji (4 bytes per character). 2 arguments used.
		oParser = new parserFormula('LEFTB("?",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("?",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Positive case: String, Number. Emoji (4 bytes per character). 2 arguments used.');
		// Case #18: Array, Number. Multi-element array of strings. 2 arguments used.
		oParser = new parserFormula('LEFTB({"abc","def"},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB({"abc","def"},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ab', 'Test: Positive case: Array, Number. Multi-element array of strings. 2 arguments used.');
		// Case #19: Formula, Number. Nested IF returning valid string. 2 arguments used.
		oParser = new parserFormula('LEFTB(IF(TRUE,"test","fail"),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(IF(TRUE,"test","fail"),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'tes', 'Test: Positive case: Formula, Number. Nested IF returning valid string. 2 arguments used.');
		// Case #20: String, Formula. num_bytes as formula resolving to valid number. 2 arguments used.
		oParser = new parserFormula('LEFTB("abcdef",LEN("abc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abcdef",LEN("abc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String, Formula. num_bytes as formula resolving to valid number. 2 arguments used.');
		// Case #21: Reference link. Reference to cell with multi-byte string, num_bytes omitted. 1 argument used.
		oParser = new parserFormula('LEFTB(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Reference link. Reference to cell with multi-byte string, num_bytes omitted. 1 argument used.');
		// Case #22: String, Number. String with num_bytes equal to string length. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc123",6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc123",6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc123', 'Test: Positive case: String, Number. String with num_bytes equal to string length. 2 arguments used.');

		// Negative cases:
		// Case #1: String, Number. Negative num_bytes returns #NUM!. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Negative num_bytes returns #NUM!. 2 arguments used.');
		// Case #3: Error, Number. Error input propagates #N/A. 2 arguments used.
		oParser = new parserFormula('LEFTB(NA(),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(NA(),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Error input propagates #N/A. 2 arguments used.');
		// Case #4: String, String. Non-numeric num_bytes returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc","def")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc","def") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Non-numeric num_bytes returns #VALUE!. 2 arguments used.');
		// Case #5: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(A102:A103,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(A102:A103,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Negative case: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #6: Reference link, Number. Reference to cell with invalid value (xyz) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(A102,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(A102,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Negative case: Reference link, Number. Reference to cell with invalid value (xyz) returns #VALUE!. 2 arguments used.');
		// Case #7: Boolean, Number. Boolean input returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(FALSE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(FALSE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FA', 'Test: Negative case: Boolean, Number. Boolean input returns #VALUE!. 2 arguments used.');
		// Case #8: Ref3D, Number. 3D reference to cell with non-string value returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(Sheet2!A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Sheet2!A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Negative case: Ref3D, Number. 3D reference to cell with non-string value returns #VALUE!. 2 arguments used.');
		// Case #9: Name, Number. Named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(TestNameArea2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(TestNameArea2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Negative case: Name, Number. Named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #10: Name3D, Number. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(TestNameArea3D2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(TestNameArea3D2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Negative case: Name3D, Number. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #11: Table, Number. Table column with non-string value returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(Table1[Column2],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Table1[Column2],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1s', 'Test: Negative case: Table, Number. Table column with non-string value returns #VALUE!. 2 arguments used.');
		// Case #12: Formula, Number. Formula resulting in #NUM! returns #NUM!. 2 arguments used.
		oParser = new parserFormula('LEFTB(SQRT(-1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(SQRT(-1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Number. Formula resulting in #NUM! returns #NUM!. 2 arguments used.');
		// Case #14: String, Number. num_bytes greater than string length returns entire string. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String, Number. num_bytes greater than string length returns entire string. 2 arguments used.');
		// Case #15: Array, Number. Array with boolean element returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB({FALSE},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB({FALSE},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FA', 'Test: Negative case: Array, Number. Array with boolean element returns #VALUE!. 2 arguments used.');
		// Case #16: Area3D, Number. 3D multi-cell range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB(Sheet2!A1:A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB(Sheet2!A1:A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Area3D, Number. 3D multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #17: String, Number. Odd num_bytes for multi-byte string (partial character) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB("??",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("??",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??', 'Test: Negative case: String, Number. Odd num_bytes for multi-byte string (partial character) returns #VALUE!. 2 arguments used.');
		// Case #18: String, Boolean. Boolean num_bytes returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Negative case: String, Boolean. Boolean num_bytes returns #VALUE!. 2 arguments used.');
		// Case #20: String, Number. num_bytes splitting emoji (4 bytes) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('LEFTB("?",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("?",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Negative case: String, Number. num_bytes splitting emoji (4 bytes) returns #VALUE!. 2 arguments used.');

		// Bounded cases:
		// Case #2: String, Number. Minimum non-zero num_bytes (1). 2 arguments used.
		oParser = new parserFormula('LEFTB("abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Bounded case: String, Number. Minimum non-zero num_bytes (1). 2 arguments used.');
		// Case #3: String, Number. Maximum valid num_bytes (Excel’s number limit) returns entire string. 2 arguments used.
		oParser = new parserFormula('LEFTB("abc",2.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEFTB("abc",2.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Bounded case: String, Number. Maximum valid num_bytes (Excel’s number limit) returns entire string. 2 arguments used.');

		// Need to fix:
		// Case #5: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.

	});

	QUnit.test("Test: \"LEN\"", function (assert) {

		ws.getRange2("A201").setValue("Phoenix, AZ");
		ws.getRange2("A202").setValue("");
		ws.getRange2("A203").setValue("     One   ");

		oParser = new parserFormula("LEN(A201)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 11);

		oParser = new parserFormula("LEN(A202)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 0);

		oParser = new parserFormula("LEN(A203)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 11);

		oParser = new parserFormula('LEN(TRUE)', "A2", ws);
		assert.ok(oParser.parse(), 'LEN(TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'LEN(TRUE)');

		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A104").setValue("#N/A");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
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
		// Case #1: String. Basic string input. 1 argument used.
		oParser = new parserFormula('LEN("Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN("Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: String. Basic string input. 1 argument used.');
		// Case #2: Number. Number converted to string. 1 argument used.
		oParser = new parserFormula('LEN(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number. Number converted to string. 1 argument used.');
		// Case #3: Empty. Empty string returns 0. 1 argument used.
		oParser = new parserFormula('LEN("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Empty. Empty string returns 0. 1 argument used.');
		// Case #4: Reference link. Ref to cell with text. 1 argument used.
		oParser = new parserFormula('LEN(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Reference link. Ref to cell with text. 1 argument used.');
		// Case #5: Area. Single-cell range. 1 argument used.
		oParser = new parserFormula('LEN(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(A101:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Area. Single-cell range. 1 argument used.');
		// Case #6: Array. Array with single element. 1 argument used.
		oParser = new parserFormula('LEN({"Excel"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN({"Excel"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Array. Array with single element. 1 argument used.');
		// Case #7: Name. Named range with text. 1 argument used.
		oParser = new parserFormula('LEN(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name. Named range with text. 1 argument used.');
		// Case #8: Name3D. 3D named range. 1 argument used.
		oParser = new parserFormula('LEN(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name3D. 3D named range. 1 argument used.');
		// Case #9: Ref3D. 3D reference to cell. 1 argument used.
		oParser = new parserFormula('LEN(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Ref3D. 3D reference to cell. 1 argument used.');
		// Case #10: Area3D. 3D single-cell range. 1 argument used.
		oParser = new parserFormula('LEN(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Area3D. 3D single-cell range. 1 argument used.');
		// Case #11: Table. Table structured reference. 1 argument used.
		oParser = new parserFormula('LEN(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference. 1 argument used.');
		// Case #12: Date. Date as serial number. 1 argument used.
		oParser = new parserFormula('LEN(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Date. Date as serial number. 1 argument used.');
		// Case #13: Time. Time formula. 1 argument used.
		oParser = new parserFormula('LEN(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Time. Time formula. 1 argument used.');
		// Case #14: Formula. Nested formula. 1 argument used.
		oParser = new parserFormula('LEN(UPPER("text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(UPPER("text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula. Nested formula. 1 argument used.');
		// Case #15: String. String with spaces. 1 argument used.
		oParser = new parserFormula('LEN("   ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN("   ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: String. String with spaces. 1 argument used.');
		// Case #16: String. Numeric string. 1 argument used.
		oParser = new parserFormula('LEN("123.45")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN("123.45") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: String. Numeric string. 1 argument used.');
		// Case #17: Boolean. Boolean converted to string. 1 argument used.
		oParser = new parserFormula('LEN(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Boolean. Boolean converted to string. 1 argument used.');
		// Case #18: Error. Error value. 1 argument used.
		oParser = new parserFormula('LEN(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Error. Error value. 1 argument used.');
		// Case #19: Array. Multi-element array. 1 argument used.
		oParser = new parserFormula('LEN({"A","B"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN({"A","B"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Array. Multi-element array. 1 argument used.');
		// Case #20: Formula. LEN inside SUM formula. 1 argument used.
		oParser = new parserFormula('SUM(LEN("A"),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(LEN("A"),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. LEN inside SUM formula. 1 argument used.');

		// Negative cases:
		// Case #1: Reference link. Ref to empty cell returns 0.
		oParser = new parserFormula('LEN(A122)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(A122) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Ref to empty cell returns 0.');
		// Case #2: Reference link. Ref to error cell returns error.
		oParser = new parserFormula('LEN(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Test: Negative case: Reference link. Ref to error cell returns error.');
		// Case #3: Area. Multi-cell range returns error.
		oParser = new parserFormula('LEN(A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(A104:A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Test: Negative case: Area. Multi-cell range returns error.');
		// Case #4: String. String with error name. 1 argument used.
		oParser = new parserFormula('LEN("#VALUE!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN("#VALUE!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Test: Negative case: String. String with error name. 1 argument used.');
		// Case #5: Empty. Missing argument returns error.
		oParser = new parserFormula('LEN()', 'A2', ws);
		assert.ok(oParser.parse() === false, 'Test: LEN() is not parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Empty. Missing argument returns error.');

		// Bounded cases:
		// Case #1: String. Max string length in Excel. 1 argument used.
		oParser = new parserFormula('LEN(REPT("A",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(REPT("A",32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32767, 'Test: Bounded case: String. Max string length in Excel. 1 argument used.');
		// Case #2: String. Min non-zero string length. 1 argument used.
		oParser = new parserFormula('LEN(REPT("A",1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(REPT("A",1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String. Min non-zero string length. 1 argument used.');
		// Case #3: Number. Max Excel number converted to string. 1 argument used.
		oParser = new parserFormula('LEN(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LEN(9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Number. Max Excel number converted to string. 1 argument used.');

		// Need to fix: areas cross check should be changed to array returns
		// Different result with MS
		// Case #5: Area. Single-cell range
		// Case #3: Area. Multi-cell range returns error
		// Case #3: Number. Max Excel number converted to string - different result


		testArrayFormula2(assert, "LEN", 1, 1);
	});

	QUnit.test("Test: \"LENB\"", function (assert) {

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
		// Case #1: String. Basic valid input: ASCII string, returns byte count (3 bytes in SBCS, 3 bytes in DBCS). 1 argument used.
		oParser = new parserFormula('LENB("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: String. Basic valid input: ASCII string, returns byte count (3 bytes in SBCS, 3 bytes in DBCS). 1 argument used.');
		// Case #2: String. Double-byte string (Chinese), returns byte count (4 bytes in DBCS, 2 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB("??")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("??") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String. Double-byte string (Chinese), returns byte count (4 bytes in DBCS, 2 bytes in SBCS). 1 argument used.');
		// Case #3: Number. Number converted to string, returns byte count (3 bytes in SBCS/DBCS). 1 argument used.
		oParser = new parserFormula('LENB(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Number. Number converted to string, returns byte count (3 bytes in SBCS/DBCS). 1 argument used.');
		// Case #4: Formula. Nested formula CONCAT creates string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(CONCAT("ab","c"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(CONCAT("ab","c")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested formula CONCAT creates string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #5: Formula. Nested IF returns string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(IF(TRUE,"abc","def"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(IF(TRUE,"abc","def")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested IF returns string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #6: Reference link. Reference to cell with string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Reference link. Reference to cell with string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #7: Area. Single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.
		oParser = new parserFormula('LENB(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(A101:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Area. Single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.');
		// Case #8: Array. Array with single string element, returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB({"abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB({"abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Array. Array with single string element, returns byte count (3 bytes). 1 argument used.');
		// Case #9: Name. Named range with string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name. Named range with string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #10: Name3D. 3D named range with string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Name3D. 3D named range with string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #11: Ref3D. 3D reference to cell with string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Sheet2!A1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to cell with string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #12: Area3D. 3D single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.
		oParser = new parserFormula('LENB(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Sheet2!A2:A2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area3D. 3D single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.');
		// Case #13: Table. Table structured reference with string "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with string "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #14: Date. Date converted to string (e.g., "1/1/2025"), returns byte count (e.g., 9 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: Date. Date converted to string (e.g., "1/1/2025"), returns byte count (e.g., 9 bytes in SBCS). 1 argument used.');
		// Case #15: Time. Time converted to string (e.g., "12:00:00"), returns byte count (8 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Time. Time converted to string (e.g., "12:00:00"), returns byte count (8 bytes in SBCS). 1 argument used.');
		// Case #16: Formula. Nested LEFT returns "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(LEFT("abcde",3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(LEFT("abcde",3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested LEFT returns "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #17: String. Single double-byte character (Japanese), returns byte count (2 bytes in DBCS). 1 argument used.
		oParser = new parserFormula('LENB("?")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("?") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String. Single double-byte character (Japanese), returns byte count (2 bytes in DBCS). 1 argument used.');
		// Case #18: Array. Array with multiple string elements, processes first element "abc" (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB({"abc","??"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB({"abc","??"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Array. Array with multiple string elements, processes first element "abc" (3 bytes). 1 argument used.');
		// Case #19: Formula. Nested RIGHT returns "cde", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(RIGHT("abcde",3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(RIGHT("abcde",3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested RIGHT returns "cde", returns byte count (3 bytes). 1 argument used.');
		// Case #20: Empty. Empty string, returns byte count (0 bytes). 1 argument used.
		oParser = new parserFormula('LENB("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Empty. Empty string, returns byte count (0 bytes). 1 argument used.');
		// Case #21: String. Short date-like string, returns byte count (5 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB("12/12")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("12/12") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 5, 'Test: Positive case: String. Short date-like string, returns byte count (5 bytes in SBCS). 1 argument used.');
		// Case #22: Formula. Nested MID returns "bcd", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(MID("abcde",2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(MID("abcde",2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Positive case: Formula. Nested MID returns "bcd", returns byte count (3 bytes). 1 argument used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('LENB(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #2: Boolean. Boolean TRUE, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Boolean. Boolean TRUE, returns #VALUE! error. 1 argument used.');
		// Case #3: Area. Multi-cell range, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Negative case: Area. Multi-cell range, returns #VALUE! error. 1 argument used.');
		// Case #4: Reference link. Reference to empty cell, returns byte count (0 bytes). 1 argument used.
		oParser = new parserFormula('LENB(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Reference link. Reference to empty cell, returns byte count (0 bytes). 1 argument used.');
		// Case #5: String. Non-numeric string, treated as valid input, returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB("Str")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("Str") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 3, 'Test: Negative case: String. Non-numeric string, treated as valid input, returns byte count (3 bytes). 1 argument used.');
		// Case #6: Formula. Formula resulting in #NUM! error, propagates #NUM!. 1 argument used.
		oParser = new parserFormula('LENB(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error, propagates #NUM!. 1 argument used.');
		// Case #7: Name. Named range with multi-cell area, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TestNameArea2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Named range with multi-cell area, returns #VALUE! error. 1 argument used.');
		// Case #8: Name3D. 3D named range with multi-cell area, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TestNameArea3D2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name3D. 3D named range with multi-cell area, returns #VALUE! error. 1 argument used.');
		// Case #9: Ref3D. 3D reference to text "Str", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(Sheet2!A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Sheet2!A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Ref3D. 3D reference to text "Str", returns byte count (3 bytes). 1 argument used.');
		// Case #10: Area3D. 3D multi-cell range, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(Sheet2!A5:A6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Sheet2!A5:A6) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area3D. 3D multi-cell range, returns #VALUE! error. 1 argument used.');
		// Case #11: Table. Table column with text "abc", returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Table. Table column with text "abc", returns byte count (3 bytes). 1 argument used.');
		// Case #12: Array. Array with boolean, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Array. Array with boolean, returns #VALUE! error. 1 argument used.');
		// Case #13: Formula. Nested IF propagates #N/A error. 1 argument used.
		oParser = new parserFormula('LENB(IF(FALSE,NA(),TRUE))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(IF(FALSE,NA(),TRUE)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Formula. Nested IF propagates #N/A error. 1 argument used.');
		// Case #14: Number. Negative number converted to string, returns byte count (2 bytes). 1 argument used.
		oParser = new parserFormula('LENB(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Number. Negative number converted to string, returns byte count (2 bytes). 1 argument used.');
		// Case #15: Formula. Formula resulting in #DIV/0! error, propagates #DIV/0!. 1 argument used.
		oParser = new parserFormula('LENB(MMULT(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(MMULT(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Formula. Formula resulting in #DIV/0! error, propagates #DIV/0!. 1 argument used.');
		// Case #16: Time. Invalid time value, propagates #NUM! error. 1 argument used.
		oParser = new parserFormula('LENB(TIME(25,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TIME(25,0,0)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 18, 'Test: Negative case: Time. Invalid time value, propagates #NUM! error. 1 argument used.');
		// Case #17: Date. Invalid date (beyond Excel limit), propagates #NUM! error. 1 argument used.
		oParser = new parserFormula('LENB(DATE(10000,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(DATE(10000,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Date. Invalid date (beyond Excel limit), propagates #NUM! error. 1 argument used.');
		// Case #18: Area3D. 3D multi-cell range, returns #VALUE! error. 1 argument used.
		oParser = new parserFormula('LENB(Sheet2!A1:B1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(Sheet2!A1:B1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area3D. 3D multi-cell range, returns #VALUE! error. 1 argument used.');
		// Case #19: Name. Named range with number 123, returns byte count (3 bytes). 1 argument used.
		oParser = new parserFormula('LENB(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Negative case: Name. Named range with number 123, returns byte count (3 bytes). 1 argument used.');
		// Case #20: Array. Array with empty string as first element, returns byte count (0 bytes). 1 argument used.
		oParser = new parserFormula('LENB({"","abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB({"","abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Array. Array with empty string as first element, returns byte count (0 bytes). 1 argument used.');

		// Bounded cases:
		// Case #1: String. Maximum string length (32,767 ASCII characters), returns byte count (32,767 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB(REPT("A",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(REPT("A",32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 32767, 'Test: Bounded case: String. Maximum string length (32,767 ASCII characters), returns byte count (32,767 bytes in SBCS). 1 argument used.');
		// Case #2: String. Maximum double-byte string length (16,383 characters in DBCS), returns byte count (32,766 bytes in DBCS). 1 argument used.
		oParser = new parserFormula('LENB(REPT("?",16383))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(REPT("?",16383)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 16383, 'Test: Bounded case: String. Maximum double-byte string length (16,383 characters in DBCS), returns byte count (32,766 bytes in DBCS). 1 argument used.');
		// Case #3: Empty. Minimum valid input (empty string), returns byte count (0 bytes). 1 argument used.
		oParser = new parserFormula('LENB("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Empty. Minimum valid input (empty string), returns byte count (0 bytes). 1 argument used.');
		// Case #4: Number. Maximum Excel number converted to string, returns byte count (e.g., 22 bytes in SBCS). 1 argument used.
		oParser = new parserFormula('LENB(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LENB(9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Bounded case: Number. Maximum Excel number converted to string, returns byte count (e.g., 22 bytes in SBCS). 1 argument used.');

		// Need to fix: areas cross check should be changed to array returns, diff results from MS
		// Case #7: Area. Single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.
		// Case #11: Ref3D. 3D reference to cell with string "abc", returns byte count (3 bytes). 1 argument used.
		// Case #12: Area3D. 3D single-cell range with string "??", returns byte count (4 bytes in DBCS). 1 argument used.
		// Case #3: Area. Multi-cell range, returns #VALUE! error. 1 argument used.
		// Case #4: Reference link. Reference to empty cell, returns byte count (0 bytes). 1 argument used.
		// Case #7: Name. Named range with multi-cell area, returns #VALUE! error. 1 argument used.
		// Case #8: Name3D. 3D named range with multi-cell area, returns #VALUE! error. 1 argument used.
		// Case #16: Time. Invalid time value, propagates #NUM! error. 1 argument used.
		// Case #4: Number. Maximum Excel number converted to string, returns byte count (e.g., 22 bytes in SBCS). 1 argument used.


		testArrayFormula2(assert, "LEN", 1, 1);
	});

	QUnit.test("Test: \"LOWER\"", function (assert) {
		ws.getRange2("A2").setValue("E. E. Cummings");
		ws.getRange2("A3").setValue("Apt. 2B");

		oParser = new parserFormula("LOWER(A2)", "A1", ws);
		assert.ok(oParser.parse(), "LOWER(A2)");
		assert.strictEqual(oParser.calculate().getValue(), "e. e. cummings", "LOWER(A2)");

		oParser = new parserFormula("LOWER(A3)", "A1", ws);
		assert.ok(oParser.parse(), "LOWER(A3)");
		assert.strictEqual(oParser.calculate().getValue(), "apt. 2b", "LOWER(A3)");

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
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("Col1Text"); // Text (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("#N/A");
		ws2.getRange2("B1").setValue("-1");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		initDefNames();
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
		// Case #1: String. Basic valid input: uppercase string. 1 argument used.
		oParser = new parserFormula('LOWER("HELLO")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER("HELLO") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'hello', 'Test: Positive case: String. Basic valid input: uppercase string. 1 argument used.');
		// Case #2: String. Alphanumeric string. 1 argument used.
		oParser = new parserFormula('LOWER("123ABC")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER("123ABC") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123abc', 'Test: Positive case: String. Alphanumeric string. 1 argument used.');
		// Case #3: String. Cyrillic string. 1 argument used.
		oParser = new parserFormula('LOWER("??????")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER("??????") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??????', 'Test: Positive case: String. Cyrillic string. 1 argument used.');
		// Case #5: Number. Number implicitly converted to string. 1 argument used.
		oParser = new parserFormula('LOWER(100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "100", 'Test: Positive case: Number. Number implicitly converted to string. 1 argument used.');
		// Case #6: Formula. Nested formula. 1 argument used.
		oParser = new parserFormula('LOWER(UPPER("test"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(UPPER("test")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'test', 'Test: Positive case: Formula. Nested formula. 1 argument used.');
		// Case #7: Reference link. Ref to cell with text. 1 argument used.
		oParser = new parserFormula('LOWER(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link. Ref to cell with text. 1 argument used.');
		// Case #8: Area. Single-cell range. 1 argument used.
		oParser = new parserFormula('LOWER(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area. Single-cell range. 1 argument used.');
		// Case #9: Array. Array with single element. 1 argument used.
		oParser = new parserFormula('LOWER({"HELLO"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER({"HELLO"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'hello', 'Test: Positive case: Array. Array with single element. 1 argument used.');
		// Case #10: Name. Named range. 1 argument used.
		oParser = new parserFormula('LOWER(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "-0.5", 'Test: Positive case: Name. Named range. 1 argument used.');
		// Case #11: Name3D. 3D named range. 1 argument used.
		oParser = new parserFormula('LOWER(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "-0.5", 'Test: Positive case: Name3D. 3D named range. 1 argument used.');
		// Case #12: Ref3D. 3D reference to cell. 1 argument used.
		oParser = new parserFormula('LOWER(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Ref3D. 3D reference to cell. 1 argument used.');
		// Case #13: Area3D. 3D single-cell range. 1 argument used.
		oParser = new parserFormula('LOWER(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Area3D. 3D single-cell range. 1 argument used.');
		// Case #14: Table. Table structured reference. 1 argument used.
		oParser = new parserFormula('LOWER(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "col1text", 'Test: Positive case: Table. Table structured reference. 1 argument used.');
		// Case #15: Date. Date converted to string. 1 argument used.
		oParser = new parserFormula('LOWER(TEXT(DATE(2025,1,1),"YYYY"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TEXT(DATE(2025,1,1),"YYYY")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "2025", 'Test: Positive case: Date. Date converted to string. 1 argument used.');
		// Case #16: Time. Time converted to string. 1 argument used.
		oParser = new parserFormula('LOWER(TEXT(TIME(12,0,0),"HH:MM"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TEXT(TIME(12,0,0),"HH:MM")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12:00', 'Test: Positive case: Time. Time converted to string. 1 argument used.');
		// Case #17: Formula. LOWER inside CONCAT. 1 argument used.
		oParser = new parserFormula('CONCAT(LOWER("A"), "B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT(LOWER("A"), "B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'aB', 'Test: Positive case: Formula. LOWER inside CONCAT. 1 argument used.');
		// Case #18: String. Special characters. 1 argument used.
		oParser = new parserFormula('LOWER("!@#")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER("!@#") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#', 'Test: Positive case: String. Special characters. 1 argument used.');
		// Case #19: Array. Multi-element array. 1 argument used.
		oParser = new parserFormula('LOWER({"A", "B"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER({"A", "B"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Array. Multi-element array. 1 argument used.');
		// Case #20: Formula. Nested IF returning valid value. 1 argument used.
		oParser = new parserFormula('LOWER(IF(TRUE, "YES", "NO"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(IF(TRUE, "YES", "NO")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'yes', 'Test: Positive case: Formula. Nested IF returning valid value. 1 argument used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('LOWER(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #2: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('LOWER(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'true', 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #3: Empty. Reference link is empty. 1 argument used.
		oParser = new parserFormula('LOWER(A112)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(A112) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Empty. Reference link is empty. 1 argument used.');
		// Case #4: Area. Multi-cell range returns arr. 1 argument used.
		oParser = new parserFormula('LOWER(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Area. Multi-cell range returns arr. 1 argument used.');
		// Case #5: Ref3D. 3D ref to error returns #VALUE!. 1 argument used.
		oParser = new parserFormula('LOWER(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. 3D ref to error returns #VALUE!. 1 argument used.');
		// Case #6: Name. Named range with val. 1 argument used.
		oParser = new parserFormula('LOWER(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), "1", 'Test: Negative case: Name. Named range with val. 1 argument used.');
		// Case #8: Number. Extremely large number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('LOWER(1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "1e+307", 'Test: Negative case: Number. Extremely large number returns #VALUE!. 1 argument used.');

		let longStr = "a";
		longStr = longStr.repeat(32767);
		// Bounded cases:
		// Case #1: String. Max string length in Excel. 1 argument used.
		oParser = new parserFormula('LOWER(REPT("A", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(REPT("A", 32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), longStr, 'Test: Bounded case: String. Max string length in Excel. 1 argument used.');
		// Case #2: String. Min non-empty string. 1 argument used.
		oParser = new parserFormula('LOWER("A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER("A") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Bounded case: String. Min non-empty string. 1 argument used.');
		// Case #3: Number. Zero converted to string. 1 argument used.
		oParser = new parserFormula('LOWER(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: LOWER(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0", 'Test: Bounded case: Number. Zero converted to string. 1 argument used.');

		// Need to fix:
		// Should return array when encounter the area
		// Case #4: Area. Multi-cell range returns arr.
		// Case #6: Name. Named range with val

		testArrayFormula2(assert, "LOWER", 1, 1);
	});

	QUnit.test("Test: \"MID\"", function (assert) {
		ws.getRange2("A101").setValue("Fluid Flow");

		oParser = new parserFormula("MID(A101,1,5)", "A2", ws);
		assert.ok(oParser.parse(), "MID(A101,1,5)");
		assert.strictEqual(oParser.calculate().getValue(), "Fluid", "MID(A101,1,5)");

		oParser = new parserFormula("MID(A101,7,20)", "A2", ws);
		assert.ok(oParser.parse(), "MID(A101,7,20)");
		assert.strictEqual(oParser.calculate().getValue(), "Flow", "MID(A101,7,20)");

		oParser = new parserFormula("MID(A101,20,5)", "A2", ws);
		assert.ok(oParser.parse(), "MID(A101,20,5)");
		assert.strictEqual(oParser.calculate().getValue(), "", "MID(A101,20,5))");

		oParser = new parserFormula("MID(TRUE,2,5)", "A2", ws);
		assert.ok(oParser.parse(), "MID(TRUE,2,5)");
		assert.strictEqual(oParser.calculate().getValue(), "RUE", "MID(TRUE,2,5)");

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
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("321"); // Number (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("#N/A");
		ws2.getRange2("A3").setValue("Text");
		ws2.getRange2("B1").setValue("-1");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		initDefNames();
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
		// Case #1: String, Number, Number. Basic valid input. All 3 arguments used.
		oParser = new parserFormula('MID("Excel",2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID("Excel",2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xce', 'Test: Positive case: String, Number, Number. Basic valid input. All 3 arguments used.');
		// Case #2: Reference link, Number, Number. Text as reference link. 3 arguments used.
		oParser = new parserFormula('MID(A100,1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(A100,1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link, Number, Number. Text as reference link. 3 arguments used.');
		// Case #3: Formula, Number, Number. Nested formula in text. 3 arguments used.
		oParser = new parserFormula('MID(UPPER("text"),2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(UPPER("text"),2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'EX', 'Test: Positive case: Formula, Number, Number. Nested formula in text. 3 arguments used.');
		// Case #4: Number, Number, Number. Number as text. 3 arguments used.
		oParser = new parserFormula('MID(12345,2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(12345,2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "234", 'Test: Positive case: Number, Number, Number. Number as text. 3 arguments used.');
		// Case #5: String, Formula, Number. Position as formula. 3 arguments used.
		oParser = new parserFormula('MID("Data",IF(TRUE,2,1),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID("Data",IF(TRUE,2,1),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ata', 'Test: Positive case: String, Formula, Number. Position as formula. 3 arguments used.');
		// Case #6: String, Number, Formula. Length as formula. 3 arguments used.
		oParser = new parserFormula('MID("Test",1,LEN("abc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID("Test",1,LEN("abc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Tes', 'Test: Positive case: String, Number, Formula. Length as formula. 3 arguments used.');
		// Case #7: Area, Number, Number. Text as single-cell area. 3 arguments used.
		oParser = new parserFormula('MID(A101:A101,2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(A101:A101,2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area, Number, Number. Text as single-cell area. 3 arguments used.');
		// Case #8: Array, Number, Number. Text as array. 3 arguments used.
		oParser = new parserFormula('MID({"Array"},1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID({"Array"},1,3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Arr', 'Test: Positive case: Array, Number, Number. Text as array. 3 arguments used.');
		// Case #9: Name, Number, Number. Text as named range. 3 arguments used.
		oParser = new parserFormula('MID(TestName,1,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(TestName,1,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "-0.5", 'Test: Positive case: Name, Number, Number. Text as named range. 3 arguments used.');
		// Case #10: Ref3D, Number, Number. Text as 3D reference. 3 arguments used.
		oParser = new parserFormula('MID(Sheet2!A1,1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(Sheet2!A1,1,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0.5", 'Test: Positive case: Ref3D, Number, Number. Text as 3D reference. 3 arguments used.');
		// Case #11: Table, Number, Number. Text from table column. 3 arguments used.
		oParser = new parserFormula('MID(Table1[Column1],2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(Table1[Column1],2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '21', 'Test: Positive case: Table, Number, Number. Text from table column. 3 arguments used.');
		// Case #13: String, Number, Empty. Empty length argument. 3 arguments used.
		oParser = new parserFormula('MID("Text",2,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID("Text",2,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: String, Number, Empty. Empty length argument. 3 arguments used.');
		// Case #14: String, Empty, Number. Empty start position. 3 arguments used.
		oParser = new parserFormula('MID("Excel",,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID("Excel",,3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, Empty, Number. Empty start position. 3 arguments used.');
		// Case #16: Time, Number, Number. Time as text. 3 arguments used.
		oParser = new parserFormula('MID(TEXT(TIME(12,0,0),"hh:mm"),1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(TEXT(TIME(12,0,0),"hh:mm"),1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "12", 'Test: Positive case: Time, Number, Number. Time as text. 3 arguments used.');
		// Case #17: Date, Number, Number. Date as text. 3 arguments used.
		oParser = new parserFormula('MID(TEXT(DATE(2025,1,1),"yyyy"),3,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(TEXT(DATE(2025,1,1),"yyyy"),3,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "25", 'Test: Positive case: Date, Number, Number. Date as text. 3 arguments used.');
		// Case #18: Name3D, Number, Number. Text as 3D named range. 3 arguments used.
		oParser = new parserFormula('MID(TestName3D,2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(TestName3D,2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0.5", 'Test: Positive case: Name3D, Number, Number. Text as 3D named range. 3 arguments used.');
		// Case #19: Area3D, Number, Number. Text as 3D area. 3 arguments used.
		oParser = new parserFormula('MID(Sheet2!A1:A1,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(Sheet2!A1:A1,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.', 'Test: Positive case: Area3D, Number, Number. Text as 3D area. 3 arguments used.');

		// Negative cases:
		// Case #5: Error, Number, Number. Error value as text. 3 arguments used.
		oParser = new parserFormula('MID(NA(),1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(NA(),1,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number, Number. Error value as text. 3 arguments used.');
		// Case #6: Empty, Number, Number. Empty text argument. 3 arguments used.
		oParser = new parserFormula('MID(,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Empty, Number, Number. Empty text argument. 3 arguments used.');
		// Case #8: Boolean, Number, Number. Boolean as text. 3 arguments used.
		oParser = new parserFormula('MID(TRUE,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MID(TRUE,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TR', 'Test: Negative case: Boolean, Number, Number. Boolean as text. 3 arguments used.');


		// Need to fix: empty, array and error check
		// Case #8: Array, Number, Number. Text as array
		// Case #14: String, Empty, Number. Empty start position
		// Case #5: Error, Number, Number - should return error

		testArrayFormula2(assert, "MID", 3, 3);
	});

	QUnit.test("Test: \"MIDB\"", function (assert) {
		ws.getRange2("A101").setValue("Fluid Flow");

		oParser = new parserFormula("MIDB(A101,1,5)", "A2", ws);
		assert.ok(oParser.parse(), "MIDB(A101,1,5)");
		assert.strictEqual(oParser.calculate().getValue(), "Fluid", "MIDB(A101,1,5)");

		oParser = new parserFormula("MIDB(A101,7,20)", "A2", ws);
		assert.ok(oParser.parse(), "MIDB(A101,7,20)");
		assert.strictEqual(oParser.calculate().getValue(), "Flow", "MIDB(A101,7,20)");

		oParser = new parserFormula("MIDB(A101,20,5)", "A2", ws);
		assert.ok(oParser.parse(), "MIDB(A101,20,5)");
		assert.strictEqual(oParser.calculate().getValue(), "", "MIDB(A101,20,5))");

		oParser = new parserFormula("MIDB(TRUE,2,5)", "A2", ws);
		assert.ok(oParser.parse(), "MIDB(TRUE,2,5)");
		assert.strictEqual(oParser.calculate().getValue(), "RUE", "MIDB(TRUE,2,5)");

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
		// Case #1: String, Number, Number. Basic valid input. All 3 arguments used.
		oParser = new parserFormula('MIDB("Excel",2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Excel",2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xce', 'Test: Positive case: String, Number, Number. Basic valid input. All 3 arguments used.');
		// Case #2: Reference link, Number, Number. Text as reference link. 3 arguments used.
		oParser = new parserFormula('MIDB(A100,1,5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(A100,1,5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link, Number, Number. Text as reference link. 3 arguments used.');
		// Case #3: Formula, Number, Number. Nested formula in text. 3 arguments used.
		oParser = new parserFormula('MIDB(UPPER("text"),2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(UPPER("text"),2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'EX', 'Test: Positive case: Formula, Number, Number. Nested formula in text. 3 arguments used.');
		// Case #4: Number, Number, Number. Number as text. 3 arguments used.
		oParser = new parserFormula('MIDB(12345,2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(12345,2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '234', 'Test: Positive case: Number, Number, Number. Number as text. 3 arguments used.');
		// Case #5: String, Formula, Number. Position as formula. 3 arguments used.
		oParser = new parserFormula('MIDB("Data",IF(TRUE,2,1),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Data",IF(TRUE,2,1),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ata', 'Test: Positive case: String, Formula, Number. Position as formula. 3 arguments used.');
		// Case #6: String, Number, Formula. Length as formula. 3 arguments used.
		oParser = new parserFormula('MIDB("Test",1,LEN("abc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Test",1,LEN("abc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Tes', 'Test: Positive case: String, Number, Formula. Length as formula. 3 arguments used.');
		// Case #7: Area, Number, Number. Text as single-cell area. 3 arguments used.
		oParser = new parserFormula('MIDB(A101:A101,2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(A101:A101,2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Area, Number, Number. Text as single-cell area. 3 arguments used.');
		// Case #8: Array, Number, Number. Text as array. 3 arguments used.
		oParser = new parserFormula('MIDB({"Array"},1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB({"Array"},1,3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Arr', 'Test: Positive case: Array, Number, Number. Text as array. 3 arguments used.');
		// Case #9: Name, Number, Number. Text as named range. 3 arguments used.
		oParser = new parserFormula('MIDB(TestName,1,4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(TestName,1,4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name, Number, Number. Text as named range. 3 arguments used.');
		// Case #10: Ref3D, Number, Number. Text as 3D reference. 3 arguments used.
		oParser = new parserFormula('MIDB(Sheet2!A1,1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(Sheet2!A1,1,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D, Number, Number. Text as 3D reference. 3 arguments used.');
		// Case #14: String, Empty, Number. Empty start position. 3 arguments used.
		oParser = new parserFormula('MIDB("Excel",,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Excel",,3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, Empty, Number. Empty start position. 3 arguments used.');
		// Case #16: Time, Number, Number. Time as text. 3 arguments used.
		oParser = new parserFormula('MIDB(TEXT(TIME(12,0,0),"hh:mm"),1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(TEXT(TIME(12,0,0),"hh:mm"),1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Time, Number, Number. Time as text. 3 arguments used.');
		// Case #17: Date, Number, Number. Date as text. 3 arguments used.
		oParser = new parserFormula('MIDB(TEXT(DATE(2025,1,1),"yyyy"),3,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(TEXT(DATE(2025,1,1),"yyyy"),3,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '25', 'Test: Positive case: Date, Number, Number. Date as text. 3 arguments used.');
		// Case #18: Name3D, Number, Number. Text as 3D named range. 3 arguments used.
		oParser = new parserFormula('MIDB(TestName3D,2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(TestName3D,2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Name3D, Number, Number. Text as 3D named range. 3 arguments used.');
		// Case #19: Area3D, Number, Number. Text as 3D area. 3 arguments used.
		oParser = new parserFormula('MIDB(Sheet2!A1:A1,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(Sheet2!A1:A1,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Area3D, Number, Number. Text as 3D area. 3 arguments used.');

		// Negative cases:
		// Case #1: String, Negative, Number
		oParser = new parserFormula('MIDB("Text",-1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Text",-1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Negative, Number.');
		// Case #2: String, Number, Negative
		oParser = new parserFormula('MIDB("Text",2,-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Text",2,-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number, Negative.');
		// Case #3: String, Text, Number
		oParser = new parserFormula('MIDB("Text","a",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Text","a",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Text, Number.');
		// Case #4: String, Number, Text
		oParser = new parserFormula('MIDB("Text",2,"b")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Text",2,"b") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number, Text.');
		// Case #5: Error, Number, Number. Error value as text. 3 arguments used.
		oParser = new parserFormula('MIDB(NA(),1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(NA(),1,1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number, Number. Error value as text. 3 arguments used.');
		// Case #6: Empty, Number, Number
		oParser = new parserFormula('MIDB(,1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(,1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Empty, Number, Number.');
		// Case #7: String, Large number, Number
		oParser = new parserFormula('MIDB("Text",999,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("Text",999,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: String, Large number, Number.');
		// Case #8: Boolean, Number, Number. Boolean as text. 3 arguments used.
		oParser = new parserFormula('MIDB(TRUE,1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(TRUE,1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TR', 'Test: Negative case: Boolean, Number, Number. Boolean as text. 3 arguments used.');

		let str = "X";
		// Bounded cases:
		// Case #1
		oParser = new parserFormula('MIDB("A",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("A",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Bounded case 1');
		// Case #2
		oParser = new parserFormula('MIDB(REPT("X",32767),1,32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB(REPT("X",32767),1,32767) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(32767), 'Test: Bounded case 2');
		// Case #3
		oParser = new parserFormula('MIDB("X",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: MIDB("X",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X', 'Test: Bounded case 3');

		// Need to fix:
		// Case #8: Array, Number, Number. Text as array. 3 arguments used.
		// Case #14: String, Empty, Number. Empty start position. 3 arguments used.
		// Case #5: Error, Number, Number. Error value as text. 3 arguments used.

	});

	QUnit.test("Test: \"NUMBERVALUE\"", function (assert) {
		oParser = new parserFormula('NUMBERVALUE("2.500,27",",",".")', "A1", ws);
		assert.ok(oParser.parse(), 'NUMBERVALUE("2.500,27",",",".")');
		assert.strictEqual(oParser.calculate().getValue(), 2500.27, 'NUMBERVALUE("2.500,27",",",".")');

		oParser = new parserFormula('NUMBERVALUE("3.5%")', "A1", ws);
		assert.ok(oParser.parse(), 'NUMBERVALUE("3.5%")');
		assert.strictEqual(oParser.calculate().getValue(), 0.035, 'NUMBERVALUE("3.5%")');

		oParser = new parserFormula('NUMBERVALUE("3.5%%%")', "A1", ws);
		assert.ok(oParser.parse(), 'NUMBERVALUE("3.5%%%")');
		assert.strictEqual(oParser.calculate().getValue(), 0.0000035, 'NUMBERVALUE("3.5%%%")');

		oParser = new parserFormula('NUMBERVALUE(123123,6,6)', "A1", ws);
		assert.ok(oParser.parse(), 'NUMBERVALUE(123123,6,6)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'NUMBERVALUE(123123,6,6)');

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
		// Case #0: String. Basic text string with dot decimal separator. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1234.56")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1234.56") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String. Basic text string with dot decimal separator. 1 of 3 arguments used.');
		// Case #1: String, String. Text with comma decimal separator specified. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.234,56",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.234,56",",") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String, String. Text with comma decimal separator specified. 2 of 3 arguments used.');
		// Case #2: String, String, String. Text with comma decimal and dot group separator. 3 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.234,56",",",".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.234,56",",",".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String, String, String. Text with comma decimal and dot group separator. 3 of 3 arguments used.');
		// Case #3: Number. Number input converted to text internally. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(1234.56)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(1234.56) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: Number. Number input converted to text internally. 1 of 3 arguments used.');
		// Case #4: Formula. Nested CONCAT formula creating valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(CONCAT("12","34.56"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(CONCAT("12","34.56")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: Formula. Nested CONCAT formula creating valid number string. 1 of 3 arguments used.');
		// Case #5: Reference link. Reference to cell with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Positive case: Reference link. Reference to cell with valid number string. 1 of 3 arguments used.');
		// Case #6: Area. Single-cell range with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.5, 'Test: Positive case: Area. Single-cell range with valid number string. 1 of 3 arguments used.');
		// Case #7: Array. Array with single valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE({"1234.56"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE({"1234.56"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: Array. Array with single valid number string. 1 of 3 arguments used.');
		// Case #8: Name. Named range with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name. Named range with valid number string. 1 of 3 arguments used.');
		// Case #9: Name3D. 3D named range with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name3D. 3D named range with valid number string. 1 of 3 arguments used.');
		// Case #10: Ref3D. 3D reference to cell with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Positive case: Ref3D. 3D reference to cell with valid number string. 1 of 3 arguments used.');
		// Case #11: Area3D. 3D single-cell range with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Positive case: Area3D. 3D single-cell range with valid number string. 1 of 3 arguments used.');
		// Case #12: Table. Table structured reference with valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table structured reference with valid number string. 1 of 3 arguments used.');
		// Case #13: Date. Date string in short format. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("12/12/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("12/12/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 46003, 'Test: Positive case: Date. Date string in short format. 1 of 3 arguments used.');
		// Case #14: Time. Time string converted to decimal. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("12:30:45")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("12:30:45") is parsed.');
		assert.strictEqual(oParser.calculate().getValue().toFixed(8), '0.52135417', 'Test: Positive case: Time. Time string converted to decimal. 1 of 3 arguments used.');
		// Case #15: Formula. Nested IF returning valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(IF(TRUE,"1234.56","0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(IF(TRUE,"1234.56","0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: Formula. Nested IF returning valid number string. 1 of 3 arguments used.');
		// Case #16: String, String. Text with dot decimal and comma group separator. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1,234.56",".",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1,234.56",".",",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String, String. Text with dot decimal and comma group separator. 2 of 3 arguments used.');
		// Case #17: String. Scientific notation string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.23456E+10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.23456E+10") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 12345600000, 'Test: Positive case: String. Scientific notation string. 1 of 3 arguments used.');
		// Case #19: Array, String, String. Array with valid number string and separators. 3 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE({"1.234,56"},",",".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE({"1.234,56"},",",".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: Array, String, String. Array with valid number string and separators. 3 of 3 arguments used.');
		// Case #20: String, Formula. Decimal separator from nested IF. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.234,56",IF(TRUE,",","."))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.234,56",IF(TRUE,",",".")) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String, Formula. Decimal separator from nested IF. 2 of 3 arguments used.');
		// Case #21: String, Empty. Empty decimal separator. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1234.56",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1234.56",) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, Empty. Empty decimal separator. 2 of 3 arguments used.');
		// Case #22: String, String, Empty. Empty group separator. 3 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1,234.56",".",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1,234.56",".",) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, String, Empty. Empty group separator. 3 of 3 arguments used.');

		// Negative cases:
		// Case #0: String. Non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		// Case #1: Error. Propagates #N/A error. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 of 3 arguments used.');
		// Case #2: Area. Multi-cell range returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 of 3 arguments used.');
		// Case #3: Empty. Empty cell reference returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -1, 'Test: Negative case: Empty. Empty cell reference returns #VALUE!. 1 of 3 arguments used.');
		// Case #4: String. Empty string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: String. Empty string returns #VALUE!. 1 of 3 arguments used.');
		// Case #5: Reference link. Reference to non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Negative case: Reference link. Reference to non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		// Case #6: Name. Named range with non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Negative case: Name. Named range with non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		// Case #7: Name3D. 3D named range with non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Negative case: Name3D. 3D named range with non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		// Case #8: Ref3D. 3D reference to non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.5, 'Test: Negative case: Ref3D. 3D reference to non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		/// Case #10: String, String. Same decimal and group separator returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1,234.56",",",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1,234.56",",",",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Same decimal and group separator returns #VALUE!. 3 of 3 arguments used.');
		// Case #11: String, String. Invalid separator combination returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.234,56",".",".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.234,56",".",".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Invalid separator combination returns #VALUE!. 2 of 3 arguments used.');
		// Case #12: Formula. Nested formula returning #NUM! propagates error. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula returning #NUM! propagates error. 1 of 3 arguments used.');
		// Case #13: Array. Array with non-numeric string returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE({"abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE({"abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array. Array with non-numeric string returns #VALUE!. 1 of 3 arguments used.');
		// Case #14: Area3D. 3D multi-cell range returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(Sheet2!A2:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(Sheet2!A2:A3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 of 3 arguments used.');
		// Case #15: String, String. Invalid decimal separator returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1,234.56","ab")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1,234.56","ab") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Negative case: String, String. Invalid decimal separator returns #VALUE!. 2 of 3 arguments used.');
		// Case #16: String, String, String. Invalid group separator returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1,234.56",".","ab")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1,234.56",".","ab") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Negative case: String, String, String. Invalid group separator returns #VALUE!. 3 of 3 arguments used.');
		// Case #17: Formula, String. Nested formula creating non-numeric string returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(CONCAT("abc","123"),".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(CONCAT("abc","123"),".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula, String. Nested formula creating non-numeric string returns #VALUE!. 2 of 3 arguments used.');
		// Case #18: Name. Named range with multi-cell area returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.8, 'Test: Negative case: Name. Named range with multi-cell area returns #VALUE!. 1 of 3 arguments used.');
		// Case #19: Name3D. 3D named range with multi-cell area returns #VALUE!. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.8, 'Test: Negative case: Name3D. 3D named range with multi-cell area returns #VALUE!. 1 of 3 arguments used.');
		// Case #20: Time. Time value alone returns #VALUE! without proper formatting. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Negative case: Time. Time value alone returns #VALUE! without proper formatting. 1 of 3 arguments used.');

		// Bounded cases:
		// Case #0: String. Maximum valid Excel number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.7976931348623157E+308")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.7976931348623157E+308") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String. Maximum valid Excel number string. 1 of 3 arguments used.');
		// Case #1: String. Minimum valid Excel number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("-1.7976931348623157E+308")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("-1.7976931348623157E+308") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String. Minimum valid Excel number string. 1 of 3 arguments used.');
		// Case #2: String, String. Maximum valid number with comma decimal separator. 2 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE("1.7976931348623157E+308",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE("1.7976931348623157E+308",",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String, String. Maximum valid number with comma decimal separator. 2 of 3 arguments used.');
		// Case #3: Reference link. Reference to maximum valid number string. 1 of 3 arguments used.
		oParser = new parserFormula('NUMBERVALUE(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: NUMBERVALUE(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Bounded case: Reference link. Reference to maximum valid number string. 1 of 3 arguments used.');


		// Need to fix: area/area3D handle, String handle, error types diff, empty should be zero
		// Case #1: String, String. Text with comma decimal separator specified. 2 of 3 arguments used.
		// Case #17: String. Scientific notation string. 1 of 3 arguments used.
		// Case #20: String, Formula. Decimal separator from nested IF. 2 of 3 arguments used.
		// Case #21: String, Empty. Empty decimal separator. 2 of 3 arguments used.
		// Case #22: String, String, Empty. Empty group separator. 3 of 3 arguments used.
		// Case #2: Area. Multi-cell range returns #VALUE!. 1 of 3 arguments used.
		// Case #4: String. Empty string returns #VALUE!. 1 of 3 arguments used.
		// Case #14: Area3D. 3D multi-cell range returns #VALUE!. 1 of 3 arguments used.
		// Case #15: String, String. Invalid decimal separator returns #VALUE!. 2 of 3 arguments used.
		// Case #3: Reference link. Reference to maximum valid number string. 1 of 3 arguments used.

		testArrayFormula2(assert, "NUMBERVALUE", 1, 3);
	});

	QUnit.test("Test: \"PROPER\"", function (assert) {

		oParser = new parserFormula("PROPER(\"2-cent's worth\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2-Cent'S Worth");

		oParser = new parserFormula("PROPER(\"76BudGet\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "76Budget");

		oParser = new parserFormula("PROPER(\"this is a TITLE\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "This Is A Title");

		oParser = new parserFormula('PROPER(TRUE)', "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "True");

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
		// Case #1: String. Basic string input with mixed case. 1 argument used.
		oParser = new parserFormula('PROPER("hello world")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("hello world") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello World', 'Test: Positive case: String. Basic string input with mixed case. 1 argument used.');
		// Case #2: String. String with all uppercase letters. 1 argument used.
		oParser = new parserFormula('PROPER("JOHN DOE")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("JOHN DOE") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'John Doe', 'Test: Positive case: String. String with all uppercase letters. 1 argument used.');
		// Case #3: String. String with all lowercase letters. 1 argument used.
		oParser = new parserFormula('PROPER("john doe")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("john doe") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'John Doe', 'Test: Positive case: String. String with all lowercase letters. 1 argument used.');
		// Case #4: Number. Number converted to string. 1 argument used.
		oParser = new parserFormula('PROPER(12345)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(12345) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12345', 'Test: Positive case: Number. Number converted to string. 1 argument used.');
		// Case #5: Formula. Nested CONCAT formula. 1 argument used.
		oParser = new parserFormula('PROPER(CONCAT("heLLo ","wOrLd"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(CONCAT("heLLo ","wOrLd")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello World', 'Test: Positive case: Formula. Nested CONCAT formula. 1 argument used.');
		// Case #6: Formula. Nested IF formula returning valid string. 1 argument used.
		oParser = new parserFormula('PROPER(IF(TRUE,"tEsT cAsE","error"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(IF(TRUE,"tEsT cAsE","error")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test Case', 'Test: Positive case: Formula. Nested IF formula returning valid string. 1 argument used.');
		// Case #7: Reference link. Reference to cell with string. 1 argument used.
		oParser = new parserFormula('PROPER(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link. Reference to cell with string. 1 argument used.');
		// Case #8: Area. Single-cell range with string. 1 argument used.
		oParser = new parserFormula('PROPER(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.5', 'Test: Positive case: Area. Single-cell range with string. 1 argument used.');
		// Case #9: Array. Array with single string element. 1 argument used.
		oParser = new parserFormula('PROPER({"test case"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER({"test case"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Test Case', 'Test: Positive case: Array. Array with single string element. 1 argument used.');
		// Case #10: Name. Named range with string. 1 argument used.
		oParser = new parserFormula('PROPER(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name. Named range with string. 1 argument used.');
		// Case #11: Name3D. 3D named range with string. 1 argument used.
		oParser = new parserFormula('PROPER(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name3D. 3D named range with string. 1 argument used.');
		// Case #12: Ref3D. 3D reference to cell with string. 1 argument used.
		oParser = new parserFormula('PROPER(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(Sheet2!A1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D. 3D reference to cell with string. 1 argument used.');
		// Case #13: Area3D. 3D single-cell range with string. 1 argument used.
		oParser = new parserFormula('PROPER(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(Sheet2!A1:A1) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Area3D. 3D single-cell range with string. 1 argument used.');
		// Case #14: Table. Table structured reference with string. 1 argument used.
		oParser = new parserFormula('PROPER(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table. Table structured reference with string. 1 argument used.');
		// Case #15: String. String with special characters and numbers. 1 argument used.
		oParser = new parserFormula('PROPER("hello@world! 123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("hello@world! 123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello@World! 123', 'Test: Positive case: String. String with special characters and numbers. 1 argument used.');
		// Case #16: String. String with multiple spaces. 1 argument used.
		oParser = new parserFormula('PROPER("  multiple   spaces  ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("  multiple   spaces  ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '  Multiple   Spaces  ', 'Test: Positive case: String. String with multiple spaces. 1 argument used.');
		// Case #17: Date. Date serial number converted to string. 1 argument used.
		oParser = new parserFormula('PROPER(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658', 'Test: Positive case: Date. Date serial number converted to string. 1 argument used.');
		// Case #18: Time. Time serial number converted to string. 1 argument used.
		oParser = new parserFormula('PROPER(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Time. Time serial number converted to string. 1 argument used.');
		// Case #19: Formula. Nested TEXT and DATE formula. 1 argument used.
		oParser = new parserFormula('PROPER(TEXT(DATE(2025,1,1),"mmmm d, yyyy"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TEXT(DATE(2025,1,1),"mmmm d, yyyy")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'January 1, 2025', 'Test: Positive case: Formula. Nested TEXT and DATE formula. 1 argument used.');
		// Case #20: String. Single character string. 1 argument used.
		oParser = new parserFormula('PROPER("a")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("a") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: String. Single character string. 1 argument used.');

		// Negative cases:
		// Case #2: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('PROPER(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #3: Array. Multi-element array returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER({"test","case"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER({"test","case"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Test', 'Test: Negative case: Array. Multi-element array returns #VALUE!. 1 argument used.');
		// Case #4: Area. Multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #5: Boolean. Boolean TRUE returns "True". 1 argument used.
		oParser = new parserFormula('PROPER(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'True', 'Test: Negative case: Boolean. Boolean TRUE returns "True". 1 argument used.');
		// Case #6: Boolean. Boolean FALSE returns "False". 1 argument used.
		oParser = new parserFormula('PROPER(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'False', 'Test: Negative case: Boolean. Boolean FALSE returns "False". 1 argument used.');
		// Case #7: Ref3D. 3D reference to cell with non-string value returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.5', 'Test: Negative case: Ref3D. 3D reference to cell with non-string value returns #VALUE!. 1 argument used.');
		// Case #8: Name. Named range with number returns stringified number. 1 argument used.
		oParser = new parserFormula('PROPER(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Name. Named range with number returns stringified number. 1 argument used.');
		// Case #9: Name3D. 3D named range with number returns stringified number. 1 argument used.
		oParser = new parserFormula('PROPER(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Negative case: Name3D. 3D named range with number returns stringified number. 1 argument used.');
		// Case #11: Reference link. Reference to cell with error returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Reference link. Reference to cell with error returns #VALUE!. 1 argument used.');
		// Case #12: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(Sheet2!A3:A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #13: Name. Named range with area returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Name. Named range with area returns #VALUE!. 1 argument used.');
		// Case #14: Name3D. 3D named range with area returns #VALUE!. 1 argument used.
		oParser = new parserFormula('PROPER(TestNameArea3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(TestNameArea3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. 3D named range with area returns #VALUE!. 1 argument used.');
		// Case #15: Formula. Formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('PROPER(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #16: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.
		oParser = new parserFormula('PROPER(DIVIDE(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(DIVIDE(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.');
		// Case #17: String. String with only spaces returns spaces. 1 argument used.
		oParser = new parserFormula('PROPER("   ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER("   ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '   ', 'Test: Negative case: String. String with only spaces returns spaces. 1 argument used.');
		// Case #19: Area. Single-cell range with empty cell returns empty string. 1 argument used.
		oParser = new parserFormula('PROPER(A104:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(A104:A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1', 'Test: Negative case: Area. Single-cell range with empty cell returns empty string. 1 argument used.');

		// Bounded cases:
		let longStr = "";
		// Case #1: String. Maximum string length (32,767 characters). 1 argument used.
		oParser = new parserFormula('PROPER(REPT("a",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(REPT("a",32767)) is parsed.');
		longStr = "a";
		assert.strictEqual(oParser.calculate().getValue(), 'A' + longStr.repeat(32766), 'Test: Bounded case: String. Maximum string length (32,767 characters). 1 argument used.');
		// Case #3: Number. Maximum valid Excel number converted to string. 1 argument used.
		oParser = new parserFormula('PROPER(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: PROPER(9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '9.99999999999999E+307', 'Test: Bounded case: Number. Maximum valid Excel number converted to string. 1 argument used.');

		// Need to fix: area handle, error types diff
		// Case #12: Ref3D. 3D reference to cell with string. 1 argument used.
		// Case #13: Area3D. 3D single-cell range with string. 1 argument used.
		// Case #4: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Case #12: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		// Case #13: Name. Named range with area returns #VALUE!. 1 argument used.


		testArrayFormula2(assert, "PROPER", 1, 1);
	});

	QUnit.test("Test: \"REPLACE\"", function (assert) {

		oParser = new parserFormula("REPLACE(\"abcdefghijk\",3,4,\"XY\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "abXYghijk");

		oParser = new parserFormula("REPLACE(\"abcdefghijk\",3,1,\"12345\")", "B2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "ab12345defghijk");

		oParser = new parserFormula("REPLACE(\"abcdefghijk\",15,4,\"XY\")", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "abcdefghijkXY");

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
		// Case #1: String, Number(2), String. Basic valid input: string, positive integers, and replacement string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Hello World",2,5,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Hello World",2,5,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HTestWorld', 'Test: Positive case: String, Number(2), String. Basic valid input: string, positive integers, and replacement string. 4 of 4 arguments used.');
		// Case #2: String, Number(2), String. Valid input: replacing first 3 characters. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Excel",1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Excel",1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newel', 'Test: Positive case: String, Number(2), String. Valid input: replacing first 3 characters. 4 of 4 arguments used.');
		// Case #3: Formula, Number(2), String. Old_text from CONCAT formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(CONCAT("He","llo"),3,2,"p")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(CONCAT("He","llo"),3,2,"p") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hepo', 'Test: Positive case: Formula, Number(2), String. Old_text from CONCAT formula. 4 of 4 arguments used.');
		// Case #4: String, Formula, Number, String. Start_num from INT formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Data",INT(2.7),2,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Data",INT(2.7),2,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'DXa', 'Test: Positive case: String, Formula, Number, String. Start_num from INT formula. 4 of 4 arguments used.');
		// Case #5: String, Number, Formula, String. Num_chars from LEN formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Text",2,LEN("Hi"),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Text",2,LEN("Hi"),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TYt', 'Test: Positive case: String, Number, Formula, String. Num_chars from LEN formula. 4 of 4 arguments used.');
		// Case #6: Reference link, Number(2), String. Old_text as Reference link to valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(A100,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(A100,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Reference link, Number(2), String. Old_text as Reference link to valid string. 4 of 4 arguments used.');
		// Case #7: Area, Number(2), String. Old_text as single-cell range. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(A101:A101,2,2,"Z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(A101:A101,2,2,"Z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1Z', 'Test: Positive case: Area, Number(2), String. Old_text as single-cell range. 4 of 4 arguments used.');
		// Case #8: Array, Number(2), String. Old_text as single-element array. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE({"Hello"},1,2,"Hi")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE({"Hello"},1,2,"Hi") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hillo', 'Test: Positive case: Array, Number(2), String. Old_text as single-element array. 4 of 4 arguments used.');
		// Case #9: Name, Number(2), String. Old_text as Name with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(TestName,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(TestName,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New5', 'Test: Positive case: Name, Number(2), String. Old_text as Name with valid string. 4 of 4 arguments used.');
		// Case #10: Name3D, Number(2), String. Old_text as Name3D with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(TestName3D,2,2,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(TestName3D,2,2,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-X5', 'Test: Positive case: Name3D, Number(2), String. Old_text as Name3D with valid string. 4 of 4 arguments used.');
		// Case #11: Ref3D, Number(2), String. Old_text as Ref3D with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(Sheet2!A1,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(Sheet2!A1,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Ref3D, Number(2), String. Old_text as Ref3D with valid string. 4 of 4 arguments used.');
		// Case #12: Area3D, Number(2), String. Old_text as Area3D with single cell. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(Sheet2!A2:A2,2,2,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(Sheet2!A2:A2,2,2,"Y") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '2Y', 'Test: Positive case: Area3D, Number(2), String. Old_text as Area3D with single cell. 4 of 4 arguments used.');
		// Case #13: Table, Number(2), String. Old_text as Table reference with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(Table1[Column1],1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(Table1[Column1],1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Table, Number(2), String. Old_text as Table reference with valid string. 4 of 4 arguments used.');
		// Case #14: String, Number, Date, String. New_text as Date (converts to serial number). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("DateTest",2,3,DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("DateTest",2,3,DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'D45658Test', 'Test: Positive case: String, Number, Date, String. New_text as Date (converts to serial number). 4 of 4 arguments used.');
		// Case #15: String, Number, Time, String. New_text as Time (converts to decimal). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("TimeTest",2,3,TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("TimeTest",2,3,TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'T0.5Test', 'Test: Positive case: String, Number, Time, String. New_text as Time (converts to decimal). 4 of 4 arguments used.');
		// Case #16: Number, Number(2), String. Old_text as Number (converts to string). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(12345,2,2,"XX")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(12345,2,2,"XX") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1XX45', 'Test: Positive case: Number, Number(2), String. Old_text as Number (converts to string). 4 of 4 arguments used.');
		// Case #17: String, Number(2), Number. New_text as Number (converts to string). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Number",2,2,123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Number",2,2,123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'N123ber', 'Test: Positive case: String, Number(2), Number. New_text as Number (converts to string). 4 of 4 arguments used.');
		// Case #18: Array, Number(2), Array. Old_text and new_text as arrays with single element. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE({"Test"},2,2,{"New"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE({"Test"},2,2,{"New"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TNewt', 'Test: Positive case: Array, Number(2), Array. Old_text and new_text as arrays with single element. 4 of 4 arguments used.');
		// Case #19: Formula, Number(2), Formula. Old_text and new_text as formulas. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(CONCAT("A","B"),2,1,TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(CONCAT("A","B"),2,1,TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A123', 'Test: Positive case: Formula, Number(2), Formula. Old_text and new_text as formulas. 4 of 4 arguments used.');
		// Case #20: String, Number(2), Empty. New_text as empty string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("EmptyTest",2,2,"")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("EmptyTest",2,2,"") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'EtyTest', 'Test: Positive case: String, Number(2), Empty. New_text as empty string. 4 of 4 arguments used.');

		// Negative cases:

		// Case #1: Empty, Number(2), String. Old_text is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Empty, Number(2), String. Old_text is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #2: Error, Number(2), String. Old_text is error, propagates #N/A. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(NA(),1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(NA(),1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number(2), String. Old_text is error, propagates #N/A. 4 of 4 arguments used.');
		// Case #3: String, Empty, Number, String. Start_num is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Empty, Number, String. Start_num is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #4: String, Number, Empty, String. Num_chars is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewTest', 'Test: Negative case: String, Number, Empty, String. Num_chars is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #5: String, Number(2), Error. New_text is error, propagates #N/A. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,3,NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,3,NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Number(2), Error. New_text is error, propagates #N/A. 4 of 4 arguments used.');
		// Case #9: Boolean, Number(2), String. Old_text is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(TRUE,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(TRUE,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewE', 'Test: Negative case: Boolean, Number(2), String. Old_text is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #10: String, Boolean, Number, String. Start_num is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",TRUE,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",TRUE,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newt', 'Test: Negative case: String, Boolean, Number, String. Start_num is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #11: String, Number, Boolean, String. Num_chars is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,TRUE,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,TRUE,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newest', 'Test: Negative case: String, Number, Boolean, String. Num_chars is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #12: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(A102:A103,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(A102:A103,1,3,"New") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #13: Array, Number(2), String. Old_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE({TRUE,FALSE},1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE({TRUE,FALSE},1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewE', 'Test: Negative case: Array, Number(2), String. Old_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #14: Ref3D, Number(2), String. Ref3D to invalid data (e.g., text "abc"), returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(Sheet2!A3,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(Sheet2!A3,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newt', 'Test: Negative case: Ref3D, Number(2), String. Ref3D to invalid data (e.g., text "abc"), returns #VALUE!. 4 of 4 arguments used.');
		// Case #15: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(TestNameArea,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(TestNameArea,1,3,"New") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #17: String, Number(2), Array. New_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,3,{TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,3,{TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEt', 'Test: Negative case: String, Number(2), Array. New_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #18: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,3,A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,3,A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.5t', 'Test: Negative case: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #19: String, Number(2), Ref3D. New_text as Ref3D with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,3,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,3,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Textt', 'Test: Negative case: String, Number(2), Ref3D. New_text as Ref3D with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #20: String, Number(2), Name. New_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",1,3,TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",1,3,TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1t', 'Test: Negative case: String, Number(2), Name. New_text as Name with range, returns #VALUE!. 4 of 4 arguments used.');

		// Bounded cases:
		let longStr = "";
		// Case #1: String, Number(2), String. Old_text is max string length (32767), valid replacement. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(REPT("A",32767),1,1,"B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(REPT("A",32767),1,1,"B") is parsed.');
		longStr = "A";
		assert.strictEqual(oParser.calculate().getValue(), 'B' + longStr.repeat(32767 - 1), 'Test: Bounded case: String, Number(2), String. Old_text is max string length (32767), valid replacement. 4 of 4 arguments used.');
		// Case #2: String, Number(2), String. New_text is max string length (32767), valid replacement. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("A",1,1,REPT("B",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("A",1,1,REPT("B",32767)) is parsed.');
		longStr = "B";
		assert.strictEqual(oParser.calculate().getValue(), longStr.repeat(32767), 'Test: Bounded case: String, Number(2), String. New_text is max string length (32767), valid replacement. 4 of 4 arguments used.');
		// Case #3: String, Number(2), String. Start_num is max valid position for short string, returns empty string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE("Test",32767,1,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE("Test",32767,1,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TestX', 'Test: Bounded case: String, Number(2), String. Start_num is max valid position for short string, returns empty string. 4 of 4 arguments used.');
		// Case #4: String, Number(2), String. Num_chars is max valid length, replaces entire string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACE(REPT("A",32767),1,32767,"B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACE(REPT("A",32767),1,32767,"B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Bounded case: String, Number(2), String. Num_chars is max valid length, replaces entire string. 4 of 4 arguments used.');

		// Need to fix: area handle, Numbers round diff with ms
		// Case #12: Area3D, Number(2), String. Old_text as Area3D with single cell. 4 of 4 arguments used.
		// Case #12: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		// Case #15: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		// Case #18: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		// Case #20: String, Number(2), Name. New_text as Name with range, returns #VALUE!. 4 of 4 arguments used.

		testArrayFormula2(assert, "REPLACE", 4, 4);
	});

	QUnit.test("Test: \"REPLACEB\"", function (assert) {
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
		// Case #1: String, Number(2), String. Basic valid input: string, positive integers, and replacement string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Hello World",2,5,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Hello World",2,5,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HTestWorld', 'Test: Positive case: String, Number(2), String. Basic valid input: string, positive integers, and replacement string. 4 of 4 arguments used.');
		// Case #2: String, Number(2), String. Valid input: replacing first 3 characters. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Excel",1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Excel",1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newel', 'Test: Positive case: String, Number(2), String. Valid input: replacing first 3 characters. 4 of 4 arguments used.');
		// Case #3: Formula, Number(2), String. Old_text from CONCAT formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(CONCAT("He","llo"),3,2,"p")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(CONCAT("He","llo"),3,2,"p") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hepo', 'Test: Positive case: Formula, Number(2), String. Old_text from CONCAT formula. 4 of 4 arguments used.');
		// Case #4: String, Formula, Number, String. Start_num from INT formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Data",INT(2.7),2,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Data",INT(2.7),2,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'DXa', 'Test: Positive case: String, Formula, Number, String. Start_num from INT formula. 4 of 4 arguments used.');
		// Case #5: String, Number, Formula, String. Num_chars from LEN formula. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Text",2,LEN("Hi"),"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Text",2,LEN("Hi"),"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TYt', 'Test: Positive case: String, Number, Formula, String. Num_chars from LEN formula. 4 of 4 arguments used.');
		// Case #6: Reference link, Number(2), String. Old_text as Reference link to valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(A100,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(A100,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Reference link, Number(2), String. Old_text as Reference link to valid string. 4 of 4 arguments used.');
		// Case #7: Area, Number(2), String. Old_text as single-cell range. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(A101:A101,2,2,"Z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(A101:A101,2,2,"Z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1Z', 'Test: Positive case: Area, Number(2), String. Old_text as single-cell range. 4 of 4 arguments used.');
		// Case #8: Array, Number(2), String. Old_text as single-element array. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB({"Hello"},1,2,"Hi")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB({"Hello"},1,2,"Hi") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hillo', 'Test: Positive case: Array, Number(2), String. Old_text as single-element array. 4 of 4 arguments used.');
		// Case #9: Name, Number(2), String. Old_text as Name with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(TestName,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(TestName,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New5', 'Test: Positive case: Name, Number(2), String. Old_text as Name with valid string. 4 of 4 arguments used.');
		// Case #10: Name3D, Number(2), String. Old_text as Name3D with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(TestName3D,2,2,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(TestName3D,2,2,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-X5', 'Test: Positive case: Name3D, Number(2), String. Old_text as Name3D with valid string. 4 of 4 arguments used.');
		// Case #11: Ref3D, Number(2), String. Old_text as Ref3D with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(Sheet2!A1,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(Sheet2!A1,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Ref3D, Number(2), String. Old_text as Ref3D with valid string. 4 of 4 arguments used.');
		// Case #12: Area3D, Number(2), String. Old_text as Area3D with single cell. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(Sheet2!A2:A2,2,2,"Y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(Sheet2!A2:A2,2,2,"Y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2Y', 'Test: Positive case: Area3D, Number(2), String. Old_text as Area3D with single cell. 4 of 4 arguments used.');
		// Case #13: Table, Number(2), String. Old_text as Table reference with valid string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(Table1[Column1],1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(Table1[Column1],1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Positive case: Table, Number(2), String. Old_text as Table reference with valid string. 4 of 4 arguments used.');
		// Case #14: String, Number, Date, String. New_text as Date (converts to serial number). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("DateTest",2,3,DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("DateTest",2,3,DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'D45658Test', 'Test: Positive case: String, Number, Date, String. New_text as Date (converts to serial number). 4 of 4 arguments used.');
		// Case #15: String, Number, Time, String. New_text as Time (converts to decimal). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("TimeTest",2,3,TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("TimeTest",2,3,TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'T0.5Test', 'Test: Positive case: String, Number, Time, String. New_text as Time (converts to decimal). 4 of 4 arguments used.');
		// Case #16: Number, Number(2), String. Old_text as Number (converts to string). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(12345,2,2,"XX")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(12345,2,2,"XX") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1XX45', 'Test: Positive case: Number, Number(2), String. Old_text as Number (converts to string). 4 of 4 arguments used.');
		// Case #17: String, Number(2), Number. New_text as Number (converts to string). 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Number",2,2,123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Number",2,2,123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'N123ber', 'Test: Positive case: String, Number(2), Number. New_text as Number (converts to string). 4 of 4 arguments used.');
		// Case #18: Array, Number(2), Array. Old_text and new_text as arrays with single element. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB({"Test"},2,2,{"New"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB({"Test"},2,2,{"New"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TNewt', 'Test: Positive case: Array, Number(2), Array. Old_text and new_text as arrays with single element. 4 of 4 arguments used.');
		// Case #19: Formula, Number(2), Formula. Old_text and new_text as formulas. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(CONCAT("A","B"),2,1,TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(CONCAT("A","B"),2,1,TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A123', 'Test: Positive case: Formula, Number(2), Formula. Old_text and new_text as formulas. 4 of 4 arguments used.');
		// Case #20: String, Number(2), Empty. New_text as empty string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("EmptyTest",2,2,"")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("EmptyTest",2,2,"") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'EtyTest', 'Test: Positive case: String, Number(2), Empty. New_text as empty string. 4 of 4 arguments used.');

		// Negative cases:
		// Case #1: Empty, Number(2), String. Old_text is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Empty, Number(2), String. Old_text is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #2: Error, Number(2), String. Old_text is error, propagates #N/A. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(NA(),1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(NA(),1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number(2), String. Old_text is error, propagates #N/A. 4 of 4 arguments used.');
		// Case #3: String, Empty, Number, String. Start_num is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Empty, Number, String. Start_num is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #4: String, Number, Empty, String. Num_chars is empty, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewTest', 'Test: Negative case: String, Number, Empty, String. Num_chars is empty, returns #VALUE!. 4 of 4 arguments used.');
		// Case #5: String, Number(2), Error. New_text is error, propagates #N/A. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,3,NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,3,NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Number(2), Error. New_text is error, propagates #N/A. 4 of 4 arguments used.');
		// Case #9: Boolean, Number(2), String. Old_text is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(TRUE,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(TRUE,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewE', 'Test: Negative case: Boolean, Number(2), String. Old_text is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #10: String, Boolean, Number, String. Start_num is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",TRUE,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",TRUE,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newt', 'Test: Negative case: String, Boolean, Number, String. Start_num is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #11: String, Number, Boolean, String. Num_chars is Boolean, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,TRUE,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,TRUE,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newest', 'Test: Negative case: String, Number, Boolean, String. Num_chars is Boolean, returns #VALUE!. 4 of 4 arguments used.');
		// Case #12: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(A102:A103,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(A102:A103,1,3,"New") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #13: Array, Number(2), String. Old_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB({TRUE,FALSE},1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB({TRUE,FALSE},1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NewE', 'Test: Negative case: Array, Number(2), String. Old_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #14: Ref3D, Number(2), String. Ref3D to invalid data (e.g., text "abc"), returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(Sheet2!A3,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(Sheet2!A3,1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Newt', 'Test: Negative case: Ref3D, Number(2), String. Ref3D to invalid data (e.g., text "abc"), returns #VALUE!. 4 of 4 arguments used.');
		// Case #15: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(TestNameArea,1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(TestNameArea,1,3,"New") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #16: Table, Number(2), String. Table with invalid data (e.g., "invalid"), returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(Table1[Column2],1,3,"New")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(Table1[Column2],1,3,"New") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'New', 'Test: Negative case: Table, Number(2), String. Table with invalid data (e.g., "invalid"), returns #VALUE!. 4 of 4 arguments used.');
		// Case #17: String, Number(2), Array. New_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,3,{TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,3,{TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUEt', 'Test: Negative case: String, Number(2), Array. New_text as array with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #18: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,3,A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,3,A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.5t', 'Test: Negative case: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.');
		// Case #19: String, Number(2), Ref3D. New_text as Ref3D with invalid data, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,3,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,3,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Textt', 'Test: Negative case: String, Number(2), Ref3D. New_text as Ref3D with invalid data, returns #VALUE!. 4 of 4 arguments used.');
		// Case #20: String, Number(2), Name. New_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",1,3,TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",1,3,TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.8t', 'Test: Negative case: String, Number(2), Name. New_text as Name with range, returns #VALUE!. 4 of 4 arguments used.');

		// Bounded cases:
		let str = "A";
		// Case #1: String, Number(2), String. Old_text is max string length (32767), valid replacement. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(REPT("A",32767),1,1,"B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(REPT("A",32767),1,1,"B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "B" + str.repeat(32766), 'Test: Bounded case: String, Number(2), String. Old_text is max string length (32767), valid replacement. 4 of 4 arguments used.');
		// Case #2: String, Number(2), String. New_text is max string length (32767), valid replacement. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("A",1,1,REPT("B",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("A",1,1,REPT("B",32767)) is parsed.');
		str = "B";
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(32767), 'Test: Bounded case: String, Number(2), String. New_text is max string length (32767), valid replacement. 4 of 4 arguments used.');
		// Case #3: String, Number(2), String. Start_num is max valid position for short string, returns empty string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB("Test",32767,1,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB("Test",32767,1,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TestX', 'Test: Bounded case: String, Number(2), String. Start_num is max valid position for short string, returns empty string. 4 of 4 arguments used.');
		// Case #4: String, Number(2), String. Num_chars is max valid length, replaces entire string. 4 of 4 arguments used.
		oParser = new parserFormula('REPLACEB(REPT("A",32767),1,32767,"B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPLACEB(REPT("A",32767),1,32767,"B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Bounded case: String, Number(2), String. Num_chars is max valid length, replaces entire string. 4 of 4 arguments used.');

		// Need to fix:
		// Case #12: Area, Number(2), String. Old_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.
		// Case #15: Name, Number(2), String. Old_text as Name with range, returns #VALUE!. 4 of 4 arguments used.
		// Case #18: String, Number(2), Area. New_text as multi-cell range, returns #VALUE!. 4 of 4 arguments used.

	});

	QUnit.test("Test: \"REPT\"", function (assert) {

		oParser = new parserFormula('REPT("*-", 3)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "*-*-*-");

		oParser = new parserFormula('REPT("-",10)', "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "----------");

		let str = "";
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
		ws.getRange2("B601").setValue("Text"); // Text (Column2)
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
		// Case #1: String, Number. Basic valid input: string and integer. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'aa', 'Test: Positive case: String, Number. Basic valid input: string and integer. 2 of 2 arguments used.');
		// Case #2: String, Number. String with multiple characters, positive integer. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("text",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("text",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'texttexttext', 'Test: Positive case: String, Number. String with multiple characters, positive integer. 2 of 2 arguments used.');
		// Case #3: Formula, Number. Text argument from nested formula CONCAT. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(CONCAT("a","b"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(CONCAT("a","b"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abab', 'Test: Positive case: Formula, Number. Text argument from nested formula CONCAT. 2 of 2 arguments used.');
		// Case #4: String, Formula. Number_times from ROUND formula, rounds to 3. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("x",ROUND(2.7,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("x",ROUND(2.7,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xxx', 'Test: Positive case: String, Formula. Number_times from ROUND formula, rounds to 3. 2 of 2 arguments used.');
		// Case #5: Reference link, Number. Text from single cell reference. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(A100,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(A100,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50.5', 'Test: Positive case: Reference link, Number. Text from single cell reference. 2 of 2 arguments used.');
		// Case #6: Area, Number. Text from single-cell range. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(A101:A101,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(A101:A101,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.51.51.5', 'Test: Positive case: Area, Number. Text from single-cell range. 2 of 2 arguments used.');
		// Case #7: Array, Number. Array with single text element. 2 of 2 arguments used.
		oParser = new parserFormula('REPT({"a"},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT({"a"},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'aa', 'Test: Positive case: Array, Number. Array with single text element. 2 of 2 arguments used.');
		// Case #8: Name, Number. Text from named range. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestName,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestName,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5-0.5', 'Test: Positive case: Name, Number. Text from named range. 2 of 2 arguments used.');
		// Case #9: Name3D, Number. Text from 3D named range. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestName3D,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestName3D,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5-0.5-0.5', 'Test: Positive case: Name3D, Number. Text from 3D named range. 2 of 2 arguments used.');
		// Case #10: Ref3D, Number. Text from 3D cell reference. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Sheet2!A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Sheet2!A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '11', 'Test: Positive case: Ref3D, Number. Text from 3D cell reference. 2 of 2 arguments used.');
		// Case #11: Area3D, Number. Text from 3D single-cell range. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Sheet2!A2:A2,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Sheet2!A2:A2,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '222', 'Test: Positive case: Area3D, Number. Text from 3D single-cell range. 2 of 2 arguments used.');
		// Case #12: Table, Number. Text from table structured reference. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Table1[Column1],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Table1[Column1],2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '11', 'Test: Positive case: Table, Number. Text from table structured reference. 2 of 2 arguments used.');
		// Case #13: String, Number. Special Unicode character in text. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("?",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("?",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??', 'Test: Positive case: String, Number. Special Unicode character in text. 2 of 2 arguments used.');
		// Case #14: String, Number. Escaped special character (newline). 2 of 2 arguments used.
		oParser = new parserFormula('REPT("\n",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("\n",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '\n\n', 'Test: Positive case: String, Number. Escaped special character (newline). 2 of 2 arguments used.');
		// Case #15: Formula, Formula. Both arguments from formulas (LEFT, ABS). 2 of 2 arguments used.
		oParser = new parserFormula('REPT(LEFT("abc",1),ABS(-2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(LEFT("abc",1),ABS(-2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'aa', 'Test: Positive case: Formula, Formula. Both arguments from formulas (LEFT, ABS). 2 of 2 arguments used.');
		// Case #16: Formula, Number. REPT inside SUM formula, non-standard test. 2 of 2 arguments used.
		oParser = new parserFormula('SUM(REPT("a",2),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(REPT("a",2),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Formula, Number. REPT inside SUM formula, non-standard test. 2 of 2 arguments used.');
		// Case #17: String, Number. Space character as text. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(" .",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(" .",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), ' . . .', 'Test: Positive case: String, Number. Space character as text. 2 of 2 arguments used.');
		// Case #18: Array, Number. Array with multiple text elements. 2 of 2 arguments used.
		oParser = new parserFormula('REPT({"a","b"},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT({"a","b"},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'aa', 'Test: Positive case: Array, Number. Array with multiple text elements. 2 of 2 arguments used.');
		// Case #19: String, Formula. Number_times from nested IF formula. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("z",IF(TRUE,2,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("z",IF(TRUE,2,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'zz', 'Test: Positive case: String, Formula. Number_times from nested IF formula. 2 of 2 arguments used.');
		// Case #20: String, Number. Zero repetitions, returns empty string. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: String, Number. Zero repetitions, returns empty string. 2 of 2 arguments used.');
		// Case #21: Date, Number. Date serial number converted to text. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(DATE(2025,1,1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(DATE(2025,1,1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '4565845658', 'Test: Positive case: Date, Number. Date serial number converted to text. 2 of 2 arguments used.');
		// Case #22: Time, Number. Time serial number converted to text. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TIME(12,0,0),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TIME(12,0,0),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50.5', 'Test: Positive case: Time, Number. Time serial number converted to text. 2 of 2 arguments used.');

		// Negative cases:
		// Case #1: Number, Number. Number as text argument, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '11', 'Test: Negative case: Number, Number. Number as text argument, returns #VALUE!. 2 of 2 arguments used.');
		// Case #2: Boolean, Number. Boolean as text argument, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TRUE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TRUE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUETRUE', 'Test: Negative case: Boolean, Number. Boolean as text argument, returns #VALUE!. 2 of 2 arguments used.');
		// Case #3: String, Number. Negative number_times, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Negative number_times, returns #VALUE!. 2 of 2 arguments used.');
		// Case #4: Error, Number. Error as text argument, propagates #N/A. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(NA(),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(NA(),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Error as text argument, propagates #N/A. 2 of 2 arguments used.');
		// Case #6: String, Empty. Empty cell for number_times, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Empty. Empty cell for number_times, returns #VALUE!. 2 of 2 arguments used.');
		// Case #7: Area, Number. Multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(A105:A106,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(A105:A106,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '11', 'Test: Negative case: Area, Number. Multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.');
		// Case #8: Area3D, Number. 3D multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Sheet2!A3:A4,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Sheet2!A3:A4,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TextText', 'Test: Negative case: Area3D, Number. 3D multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.');
		// Case #9: Name, Number. Named range with area, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestNameArea,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestNameArea,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '11', 'Test: Negative case: Name, Number. Named range with area, returns #VALUE!. 2 of 2 arguments used.');
		// Case #10: Name3D, Number. 3D named range with area, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestNameArea3D2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestNameArea3D2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.80.8', 'Test: Negative case: Name3D, Number. 3D named range with area, returns #VALUE!. 2 of 2 arguments used.');
		// Case #11: Ref3D, Number. 3D reference to non-text value, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Sheet2!A5+"?",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Sheet2!A5+"?",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, Number. 3D reference to non-text value, returns #VALUE!. 2 of 2 arguments used.');
		// Case #12: Table, Number. Table column with non-text value, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(Table1[Column2],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(Table1[Column2],2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'TextText', 'Test: Negative case: Table, Number. Table column with non-text value, returns #VALUE!. 2 of 2 arguments used.');
		// Case #13: String, Number. Empty string as text, returns empty string but valid. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("-",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("-",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '--', 'Test: Negative case: String, Number. Empty string as text, returns empty string but valid. 2 of 2 arguments used.');
		// Case #14: Formula, Number. Formula resulting in #NUM!, propagates error. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(SQRT(-1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(SQRT(-1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Number. Formula resulting in #NUM!, propagates error. 2 of 2 arguments used.');
		// Case #15: String, Number. Non-integer number_times, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",1.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",1.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Negative case: String, Number. Non-integer number_times, returns #VALUE!. 2 of 2 arguments used.');
		// Case #16: Array, Number. Array with boolean element, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT({TRUE},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT({TRUE},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUETRUE', 'Test: Negative case: Array, Number. Array with boolean element, returns #VALUE!. 2 of 2 arguments used.');
		// Case #17: String, Formula. Number_times from formula resulting in #NUM!, propagates error. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String, Formula. Number_times from formula resulting in #NUM!, propagates error. 2 of 2 arguments used.');
		// Case #18: String, Number. Excessively large number_times, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Excessively large number_times, returns #VALUE!. 2 of 2 arguments used.');
		// Case #19: Name, Number. Named range with non-text value, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestName1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestName1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50.5', 'Test: Negative case: Name, Number. Named range with non-text value, returns #VALUE!. 2 of 2 arguments used.');
		// Case #20: Name3D, Number. 3D named range with non-text value, returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('REPT(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5-0.5', 'Test: Negative case: Name3D, Number. 3D named range with non-text value, returns #VALUE!. 2 of 2 arguments used.');

		// Bounded cases:
		// Case #1: String, Number. Maximum valid number_times (32,767). 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",32767) is parsed.');
		str = "a";
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(32767), 'Test: 32767 of "a" Repeat');
		// Case #2: String, Number. Minimum valid number_times (0), returns empty string. 2 of 2 arguments used.
		oParser = new parserFormula('REPT("a",-9999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("a",-9999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String, Number. Minimum valid number_times (0), returns empty string. 2 of 2 arguments used.');
		// Case #3: String, Number. Large number_times with Unicode character, near cell character limit (32,767). 2 of 2 arguments used.
		oParser = new parserFormula('REPT("?",16383)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REPT("?",16383) is parsed.');
		str = "?";
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(16383), 'Test: 16383 of "?" Repeat');

		// TODO: critical problem repeat number should be equal or less than 2^28
		// Need to fix: ms results diff
		// Case #12: Table, Number. Text from table structured reference. 2 of 2 arguments used.
		// Case #7: Area, Number. Multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.
		// Case #8: Area3D, Number. 3D multi-cell range as text, returns #VALUE!. 2 of 2 arguments used.
		// Case #9: Name, Number. Named range with area, returns #VALUE!. 2 of 2 arguments used.
		// Case #12: Table, Number. Table column with non-text value, returns #VALUE!. 2 of 2 arguments used.
		// Case #18: String, Number. Excessively large number_times, returns #VALUE!. 2 of 2 arguments used.


		testArrayFormula2(assert, "REPT", 2, 2);
	});

	QUnit.test("Test: \"RIGHT\"", function (assert) {

		ws.getRange2("A2").setValue("Sale Price");
		ws.getRange2("A3").setValue("Stock Number");

		oParser = new parserFormula("RIGHT(A2,5)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Price");

		oParser = new parserFormula("RIGHT(A3)", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "r");

		let str = "";
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
		ws.getRange2("B601").setValue("Text"); // Text (Column2)
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
		// Case #1: String, Number. Basic valid input: string and integer num_chars. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xt', 'Test: Positive case: String, Number. Basic valid input: string and integer num_chars. 2 of 2 arguments used.');
		// Case #2: String. String with omitted num_chars (defaults to 1). 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT("example")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("example") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'e', 'Test: Positive case: String. String with omitted num_chars (defaults to 1). 1 of 2 arguments used.');
		// Case #3: String, Number. num_chars equals string length. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("abc",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("abc",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String, Number. num_chars equals string length. 2 of 2 arguments used.');
		// Case #4: String, Number. num_chars exceeds string length, returns entire string. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("abc",5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("abc",5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String, Number. num_chars exceeds string length, returns entire string. 2 of 2 arguments used.');
		// Case #5: Formula, Number. Text from formula with integer num_chars. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(CONCAT("te","xt"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(CONCAT("te","xt"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xt', 'Test: Positive case: Formula, Number. Text from formula with integer num_chars. 2 of 2 arguments used.');
		// Case #6: Formula. Formula with omitted num_chars (defaults to 1). 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(CONCAT("ex","ample"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(CONCAT("ex","ample")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'e', 'Test: Positive case: Formula. Formula with omitted num_chars (defaults to 1). 1 of 2 arguments used.');
		// Case #7: Number, Number. Number converted to text. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(12345,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(12345,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45', 'Test: Positive case: Number, Number. Number converted to text. 2 of 2 arguments used.');
		// Case #8: Date, Number. Date serial number converted to text. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(DATE(2025,1,1),4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(DATE(2025,1,1),4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5658', 'Test: Positive case: Date, Number. Date serial number converted to text. 2 of 2 arguments used.');
		// Case #9: Time, Number. Time value converted to text. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(TIME(12,0,0),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(TIME(12,0,0),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Time, Number. Time value converted to text. 2 of 2 arguments used.');
		// Case #10: Reference link. Reference to cell with valid string, num_chars omitted. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5', 'Test: Positive case: Reference link. Reference to cell with valid string, num_chars omitted. 1 of 2 arguments used.');
		// Case #11: Reference link, Number. Reference to cell with valid string and num_chars. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Reference link, Number. Reference to cell with valid string and num_chars. 2 of 2 arguments used.');
		// Case #12: Area. Single-cell range, num_chars omitted. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5', 'Test: Positive case: Area. Single-cell range, num_chars omitted. 1 of 2 arguments used.');
		// Case #14: Array. Array with single string element. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT({"text"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT({"text"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 't', 'Test: Positive case: Array. Array with single string element. 1 of 2 arguments used.');
		// Case #15: Array, Number. Array with multiple strings, returns first element’s result. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT({"example","abc"},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT({"example","abc"},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'le', 'Test: Positive case: Array, Number. Array with multiple strings, returns first element’s result. 2 of 2 arguments used.');
		// Case #16: Name. Named range with valid string. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5', 'Test: Positive case: Name. Named range with valid string. 1 of 2 arguments used.');
		// Case #17: Name3D. 3D named range with valid string. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5', 'Test: Positive case: Name3D. 3D named range with valid string. 1 of 2 arguments used.');
		// Case #18: Ref3D. 3D reference to cell with valid string. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D. 3D reference to cell with valid string. 1 of 2 arguments used.');
		// Case #19: Area3D. 3D single-cell range. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Positive case: Area3D. 3D single-cell range. 1 of 2 arguments used.');
		// Case #20: Table. Table structured reference with valid string. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Table1[Column1]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table. Table structured reference with valid string. 1 of 2 arguments used.');
		// Case #21: Formula, Formula. Nested IF and LEN formulas. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(IF(TRUE,"text","abc"),LEN("xt"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(IF(TRUE,"text","abc"),LEN("xt")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'xt', 'Test: Positive case: Formula, Formula. Nested IF and LEN formulas. 2 of 2 arguments used.');
		// Case #22: Formula. RIGHT inside SUM formula. 2 of 2 arguments used.
		oParser = new parserFormula('SUM(RIGHT("text",2),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(RIGHT("text",2),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Formula. RIGHT inside SUM formula. 2 of 2 arguments used.');

		// Negative cases:
		// Case #1: String, Number. Negative num_chars returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Negative num_chars returns #VALUE!. 2 of 2 arguments used.');
		// Case #2: String, String. Non-numeric num_chars returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Non-numeric num_chars returns #VALUE!. 2 of 2 arguments used.');
		// Case #3: Error. Propagates #N/A error. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 of 2 arguments used.');
		// Case #4: Empty. Reference to empty cell returns empty string or #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Empty. Reference to empty cell returns empty string or #VALUE!. 1 of 2 arguments used.');
		// Case #6: Boolean. Boolean FALSE returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'E', 'Test: Negative case: Boolean. Boolean FALSE returns #VALUE!. 1 of 2 arguments used.');
		// Case #7: Area. Multi-cell range returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(A105:A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(A105:A106) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 of 2 arguments used.');
		// Case #8: String. Empty string returns empty string. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.', 'Test: Negative case: String. Empty string returns empty string. 1 of 2 arguments used.');
		// Case #9: Ref3D. 3D reference to cell with non-string (e.g., boolean) returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Sheet2!A9+"str")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Sheet2!A9+"str") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to cell with non-string (e.g., boolean) returns #VALUE!. 1 of 2 arguments used.');
		// Case #10: Name. Named range with non-string (e.g., number) returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Name. Named range with non-string (e.g., number) returns #VALUE!. 1 of 2 arguments used.');
		// Case #11: Formula. Formula resulting in #NUM! propagates error. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 of 2 arguments used.');
		// Case #12: String, Boolean. Boolean num_chars returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 't', 'Test: Negative case: String, Boolean. Boolean num_chars returns #VALUE!. 2 of 2 arguments used.');
		// Case #13: Array. Array with boolean returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT({FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT({FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'E', 'Test: Negative case: Array. Array with boolean returns #VALUE!. 1 of 2 arguments used.');
		// Case #14: Area3D. 3D multi-cell range returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Sheet2!A3:A4) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 't', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 of 2 arguments used.');
		// Case #15: Name3D. 3D named range with non-string (e.g., number) returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '8', 'Test: Negative case: Name3D. 3D named range with non-string (e.g., number) returns #VALUE!. 1 of 2 arguments used.');
		// Case #16: Table. Table column with non-string (e.g., error) returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('RIGHT(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(Table1[Column2]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Table. Table column with non-string (e.g., error) returns #VALUE!. 1 of 2 arguments used.');
		// Case #17: String, Number. Excessively large num_chars returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text",1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text",1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Negative case: String, Number. Excessively large num_chars returns #VALUE!. 2 of 2 arguments used.');
		// Case #18: Formula, Number. Invalid date (beyond Excel limit) returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(DATE(9999,12,31)+1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(DATE(9999,12,31)+1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '66', 'Test: Negative case: Formula, Number. Invalid date (beyond Excel limit) returns #VALUE!. 2 of 2 arguments used.');
		// Case #19: String, String. Empty string as num_chars returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Empty string as num_chars returns #VALUE!. 2 of 2 arguments used.');
		// Case #20: Array, Number. Array with non-string element returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT({TRUE,"text"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT({TRUE,"text"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'E', 'Test: Negative case: Array, Number. Array with non-string element returns #VALUE!. 2 of 2 arguments used.');

		// Bounded cases:
		// Case #1: String, Number. Empty string with num_chars=0 returns empty string. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(".",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(".",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.', 'Test: Bounded case: String, Number. Empty string with num_chars=0 returns empty string. 2 of 2 arguments used.');
		// Case #2: String, Number. Smallest valid num_chars above 1. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("x",1.000000000000001)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("x",1.000000000000001) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'x', 'Test: Bounded case: String, Number. Smallest valid num_chars above 1. 2 of 2 arguments used.');
		// Case #3: String, Number. Maximum string length (32,767) with matching num_chars. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT(REPT("x",32767),32767)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT(REPT("x",32767),32767) is parsed.');
		str = "x";
 		assert.strictEqual(oParser.calculate().getValue(), str.repeat(32767), 'Test: Bounded case: String, Number. Maximum string length (32,767) with matching num_chars. 2 of 2 arguments used.');

		// Case #4: String, Number. Maximum valid num_chars, returns entire string. 2 of 2 arguments used.
		oParser = new parserFormula('RIGHT("text",9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHT("text",9.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Bounded case: String, Number. Maximum valid num_chars, returns entire string. 2 of 2 arguments used.');

		// Need to fix: area/3d, table handle
		// Case #7: Area. Multi-cell range returns #VALUE!. 1 of 2 arguments used.
		// Case #10: Name. Named range with non-string (e.g., number) returns #VALUE!. 1 of 2 arguments used.
		// Case #14: Area3D. 3D multi-cell range returns #VALUE!. 1 of 2 arguments used.
		// Case #16: Table. Table column with non-string (e.g., error) returns #VALUE!. 1 of 2 arguments used.
		// Case #20: Table. Table structured reference with valid string. 1 of 2 arguments used.

		testArrayFormula2(assert, "RIGHT", 1, 2);
	});

	QUnit.test("Test: \"RIGHTB\"", function (assert) {

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
		ws.getRange2("B601").setValue("1ssssss2"); // Text (Column2)
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
		// Case #1: String, Number. Basic string with valid byte count. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abcdef",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abcdef",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'def', 'Test: Positive case: String, Number. Basic string with valid byte count. 2 arguments used.');
		// Case #2: String. String with omitted num_bytes (defaults to 1). 1 argument used.
		oParser = new parserFormula('RIGHTB("abcdef")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abcdef") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'f', 'Test: Positive case: String. String with omitted num_bytes (defaults to 1). 1 argument used.');
		// Case #3: Number, Number. Numeric string converted to text. 2 arguments used.
		oParser = new parserFormula('RIGHTB("12345",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("12345",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45', 'Test: Positive case: Number, Number. Numeric string converted to text. 2 arguments used.');
		// Case #4: Formula, Number. Nested formula resolving to string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(CONCAT("ab","cd"),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(CONCAT("ab","cd"),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'cd', 'Test: Positive case: Formula, Number. Nested formula resolving to string. 2 arguments used.');
		// Case #5: Reference link, Number. Reference to cell with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(A100,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(A100,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link, Number. Reference to cell with valid string. 2 arguments used.');
		// Case #6: Area, Number. Single-cell range with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(A101:A101,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(A101:A101,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Area, Number. Single-cell range with valid string. 2 arguments used.');
		// Case #7: Array, Number. Array with single string element. 2 arguments used.
		oParser = new parserFormula('RIGHTB({"abc"},1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB({"abc"},1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'c', 'Test: Positive case: Array, Number. Array with single string element. 2 arguments used.');
		// Case #8: Name, Number. Named range with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(TestName,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(TestName,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Name, Number. Named range with valid string. 2 arguments used.');
		// Case #9: Name3D, Number. 3D named range with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(TestName3D,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(TestName3D,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Name3D, Number. 3D named range with valid string. 2 arguments used.');
		// Case #10: Ref3D, Number. 3D reference to cell with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Sheet2!A1,3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Sheet2!A1,3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D, Number. 3D reference to cell with valid string. 2 arguments used.');
		// Case #11: Area3D, Number. 3D single-cell range with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Sheet2!A1:A1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Sheet2!A1:A1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Area3D, Number. 3D single-cell range with valid string. 2 arguments used.');
		// Case #12: Table, Number. Table structured reference with valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Table1[Column1],3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Table1[Column1],3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table, Number. Table structured reference with valid string. 2 arguments used.');
		// Case #13: Date, Number. Date serial number converted to string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(DATE(2025,1,1),4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(DATE(2025,1,1),4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5658', 'Test: Positive case: Date, Number. Date serial number converted to string. 2 arguments used.');
		// Case #14: Time, Number. Time serial number converted to string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(TIME(12,0,0),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(TIME(12,0,0),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Positive case: Time, Number. Time serial number converted to string. 2 arguments used.');
		// Case #15: Formula, Number. LEFTB inside CONCAT formula. 2 arguments used.
		oParser = new parserFormula('CONCAT(RIGHTB("xyz",2),"!")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT(RIGHTB("xyz",2),"!") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'yz!', 'Test: Positive case: Formula, Number. LEFTB inside CONCAT formula. 2 arguments used.');
		// Case #16: String, Number. Multi-byte Unicode string (Chinese, extracts 1 character = 2 bytes). 2 arguments used.
		oParser = new parserFormula('RIGHTB("??",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("??",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??', 'Test: Positive case: String, Number. Multi-byte Unicode string (Chinese, extracts 1 character = 2 bytes). 2 arguments used.');
		// Case #17: String, Number. Emoji (4 bytes per character). 2 arguments used.
		oParser = new parserFormula('RIGHTB("?",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("?",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Positive case: String, Number. Emoji (4 bytes per character). 2 arguments used.');
		// Case #18: Array, Number. Multi-element array of strings. 2 arguments used.
		oParser = new parserFormula('RIGHTB({"abc","def"},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB({"abc","def"},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'bc', 'Test: Positive case: Array, Number. Multi-element array of strings. 2 arguments used.');
		// Case #19: Formula, Number. Nested IF returning valid string. 2 arguments used.
		oParser = new parserFormula('RIGHTB(IF(TRUE,"test","fail"),3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(IF(TRUE,"test","fail"),3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'est', 'Test: Positive case: Formula, Number. Nested IF returning valid string. 2 arguments used.');
		// Case #20: String, Formula. num_bytes as formula resolving to valid number. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abcdef",LEN("abc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abcdef",LEN("abc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'def', 'Test: Positive case: String, Formula. num_bytes as formula resolving to valid number. 2 arguments used.');
		// Case #21: Reference link. Reference to cell with multi-byte string, num_bytes omitted. 1 argument used.
		oParser = new parserFormula('RIGHTB(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '5', 'Test: Positive case: Reference link. Reference to cell with multi-byte string, num_bytes omitted. 1 argument used.');
		// Case #22: String, Number. String with num_bytes equal to string length. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc123",6)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc123",6) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc123', 'Test: Positive case: String, Number. String with num_bytes equal to string length. 2 arguments used.');

		// Negative cases:
		// Case #1: String, Number. Negative num_bytes returns #NUM!. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Number. Negative num_bytes returns #NUM!. 2 arguments used.');
		// Case #3: Error, Number. Error input propagates #N/A. 2 arguments used.
		oParser = new parserFormula('RIGHTB(NA(),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(NA(),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Number. Error input propagates #N/A. 2 arguments used.');
		// Case #4: String, String. Non-numeric num_bytes returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc","def")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc","def") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Non-numeric num_bytes returns #VALUE!. 2 arguments used.');
		// Case #5: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(A102:A103,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(A102:A103,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #6: Reference link, Number. Reference to cell with invalid value (xyz) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(A102,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(A102,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.5', 'Test: Negative case: Reference link, Number. Reference to cell with invalid value (xyz) returns #VALUE!. 2 arguments used.');
		// Case #7: Boolean, Number. Boolean input returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(FALSE,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(FALSE,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'SE', 'Test: Negative case: Boolean, Number. Boolean input returns #VALUE!. 2 arguments used.');
		// Case #8: Ref3D, Number. 3D reference to cell with non-string value returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Sheet2!A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Sheet2!A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Negative case: Ref3D, Number. 3D reference to cell with non-string value returns #VALUE!. 2 arguments used.');
		// Case #9: Name, Number. Named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(TestNameArea2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(TestNameArea2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '.8', 'Test: Negative case: Name, Number. Named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #10: Name3D, Number. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(TestNameArea3D2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(TestNameArea3D2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(),'.8', 'Test: Negative case: Name3D, Number. 3D named range with multi-cell area returns #VALUE!. 2 arguments used.');
		// Case #11: Table, Number. Table column with non-string value returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Table1[Column2],2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Table1[Column2],2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 's2', 'Test: Negative case: Table, Number. Table column with non-string value returns #VALUE!. 2 arguments used.');
		// Case #12: Formula, Number. Formula resulting in #NUM! returns #NUM!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(SQRT(-1),2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(SQRT(-1),2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Number. Formula resulting in #NUM! returns #NUM!. 2 arguments used.');
		// Case #14: String, Number. num_bytes greater than string length returns entire string. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String, Number. num_bytes greater than string length returns entire string. 2 arguments used.');
		// Case #15: Array, Number. Array with boolean element returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB({FALSE},2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB({FALSE},2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'SE', 'Test: Negative case: Array, Number. Array with boolean element returns #VALUE!. 2 arguments used.');
		// Case #16: Area3D, Number. 3D multi-cell range returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB(Sheet2!A1:A2,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB(Sheet2!A1:A2,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Area3D, Number. 3D multi-cell range returns #VALUE!. 2 arguments used.');
		// Case #17: String, Number. Odd num_bytes for multi-byte string (partial character) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB("??",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("??",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??', 'Test: Negative case: String, Number. Odd num_bytes for multi-byte string (partial character) returns #VALUE!. 2 arguments used.');
		// Case #18: String, Boolean. Boolean num_bytes returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'c', 'Test: Negative case: String, Boolean. Boolean num_bytes returns #VALUE!. 2 arguments used.');
		// Case #20: String, Number. num_bytes splitting emoji (4 bytes) returns #VALUE!. 2 arguments used.
		oParser = new parserFormula('RIGHTB("?",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("?",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Negative case: String, Number. num_bytes splitting emoji (4 bytes) returns #VALUE!. 2 arguments used.');

		// Bounded cases:
		// Case #2: String, Number. Minimum non-zero num_bytes (1). 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'c', 'Test: Bounded case: String, Number. Minimum non-zero num_bytes (1). 2 arguments used.');
		// Case #3: String, Number. Maximum valid num_bytes (Excel’s number limit) returns entire string. 2 arguments used.
		oParser = new parserFormula('RIGHTB("abc",2.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: RIGHTB("abc",2.99999999999999E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Bounded case: String, Number. Maximum valid num_bytes (Excel’s number limit) returns entire string. 2 arguments used.');

		// Need to fix:
		// Case #5: Area, Number. Multi-cell range returns #VALUE!. 2 arguments used.


		testArrayFormula2(assert, "LEN", 1, 1);
	});

	QUnit.test("Test: \"REGEXTEST\"", function (assert) {

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
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1ssssss2"); // Text (Column2)
		ws.getRange2("C601").setValue("[A-z]"); // Text (Column2)
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
		// Case #0: String, String. Basic match, case-sensitive by default (0)
		oParser = new parserFormula('REGEXTEST("Hello World","World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello World","World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String. Basic match, case-sensitive by default (0)');
		// Case #1: String, String, Number. Case-insensitive match (1)
		oParser = new parserFormula('REGEXTEST("Hello World","world",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello World","world",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String, Number. Case-insensitive match (1)');
		// Case #2: String, String, Number. Upper case text, insensitive mode
		oParser = new parserFormula('REGEXTEST("Hello World","WORLD",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello World","WORLD",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String, Number. Upper case text, insensitive mode');
		// Case #3: String, String. Digits token \\d+ (case-sensitive default)
		oParser = new parserFormula('REGEXTEST("abc123","\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc123","\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String, String. Digits token \\d+ (case-sensitive default)');
		// Case #4: String, String. Full email regex, anchors ^ and $
		oParser = new parserFormula('REGEXTEST("test@example.com","^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\\\.[a-zA-Z]{2,}$")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("test@example.com","^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String, String. Full email regex, anchors ^ and $');
		// Case #5: String, String, Number. Unicode (é), case-insensitive
		oParser = new parserFormula('REGEXTEST("Café","cafe",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Café","cafe",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String, String, Number. Unicode (é), case-insensitive');
		// Case #6: String, String. Digits inside longer text
		oParser = new parserFormula('REGEXTEST("price: 100$","\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("price: 100$","\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String, String. Digits inside longer text');
		// Case #7: String, String, Number. Character class [a-z] case-sensitive
		oParser = new parserFormula('REGEXTEST("ABCdef","[a-z]+",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("ABCdef","[a-z]+",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String, Number. Character class [a-z] case-sensitive');
		// Case #8: String, String. Zero or more quantifier *
		oParser = new parserFormula('REGEXTEST("aaa","a*")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("aaa","a*") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String. Zero or more quantifier *');
		// Case #9: String, String. One or more quantifier +
		oParser = new parserFormula('REGEXTEST("aaa","a+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("aaa","a+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String. One or more quantifier +');
		// Case #10: String, String. Dot matches any character
		oParser = new parserFormula('REGEXTEST("any.char.here",".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("any.char.here",".") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String. Dot matches any character');
		// Case #11: Formula, Formula. Nested formulas
		oParser = new parserFormula('REGEXTEST(CONCAT("Test","123"),"[0-9]+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(CONCAT("Test","123"),"[0-9]+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Formula, Formula. Nested formulas');
		// Case #12: Reference link, Reference link, Number. All arguments via cell references
		oParser = new parserFormula('REGEXTEST(A100,A101,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(A100,A101,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Reference link, Reference link, Number. All arguments via cell references');
		// Case #13: Area, Area, Number. Single-cell ranges
		oParser = new parserFormula('REGEXTEST(A102:A102,A103:A103,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(A102:A102,A103:A103,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area, Area, Number. Single-cell ranges');
		// Case #14: Array, Array. Array input returns array result
		oParser = new parserFormula('REGEXTEST({"Hello","Bye"},"^H")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST({"Hello","Bye"},"^H") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'TRUE', 'Test: Positive case: Array, Array. Array input returns array result');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), 'FALSE', 'Test: Positive case: Array, Array. Array input returns array result');
		// Case #15: Name, Name, Number. Named ranges
		oParser = new parserFormula('REGEXTEST(TestName,TestName1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(TestName,TestName1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name, Name, Number. Named ranges');
		// Case #16: Name3D, Name3D. 3D named ranges
		oParser = new parserFormula('REGEXTEST(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Name3D, Name3D. 3D named ranges');
		// Case #17: Ref3D, Ref3D, Number. 3D references
		oParser = new parserFormula('REGEXTEST(Sheet2!A1,Sheet2!A2,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(Sheet2!A1,Sheet2!A2,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Ref3D, Ref3D, Number. 3D references');
		// Case #18: Area3D, Area3D. 3D single-cell ranges
		oParser = new parserFormula('REGEXTEST(Sheet2!A3:A3,Sheet2!A4:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(Sheet2!A3:A3,Sheet2!A4:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Area3D, Area3D. 3D single-cell ranges');
		// Case #19: Table. Structured table references
		oParser = new parserFormula('REGEXTEST(Table1[Column2],Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(Table1[Column2],Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Table. Structured table references');
		// Case #20: Formula. REGEXTEST inside another formula (OR)
		oParser = new parserFormula('OR(REGEXTEST("abc123","\\\\d+"),FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: OR(REGEXTEST("abc123","\\\\d+"),FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: Formula. REGEXTEST inside another formula (OR)');
		// Case #21: String, String, Boolean. case_sensitivity as boolean TRUE ? 1
		oParser = new parserFormula('REGEXTEST("Test","test",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Test","test",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String, Boolean. case_sensitivity as boolean TRUE ? 1');
		// Case #22: String, String, String. case_sensitivity as text "1"
		oParser = new parserFormula('REGEXTEST("Test","test","1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Test","test","1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: String, String, String. case_sensitivity as text "1"');
		// Case #23: String, String, Empty. Optional argument omitted ? default 0 (case-sensitive)
		oParser = new parserFormula('REGEXTEST("Hello World","world",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello World","world",) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Positive case: String, String, Empty. Optional argument omitted ? default 0 (case-sensitive)');
		// Case #24: Empty, Empty. Optional argument omitted ? default 0 (case-sensitive)
		oParser = new parserFormula('REGEXTEST(,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Positive case: Empty, Empty. Optional argument omitted ? default 0 (case-sensitive)');

		// Negative cases:
		// Case #1: String, String. No match, case-sensitive ? FALSE (not error)
		oParser = new parserFormula('REGEXTEST("Hello","world")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello","world") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String, String. No match, case-sensitive ? FALSE (not error)');
		// Case #2: String, String. Invalid regex pattern (unclosed bracket) ? #VALUE!
		oParser = new parserFormula('REGEXTEST("Hello World","[")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello World","[") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Invalid regex pattern (unclosed bracket) ? #VALUE!');
		// Case #3: String, String. Unclosed parenthesis ? #VALUE!
		oParser = new parserFormula('REGEXTEST("Hello","(unclosed")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello","(unclosed") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String. Unclosed parenthesis ? #VALUE!');
		// Case #4: String, Empty. Empty pattern ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String, Empty. Empty pattern ? #VALUE!');
		// Case #5: String, String, Number. case_sensitivity not 0/1 ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","def",99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","def",99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. case_sensitivity not 0/1 ? #VALUE!');
		// Case #6: String, String, String. case_sensitivity non-numeric string ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","def","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","def","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. case_sensitivity non-numeric string ? #VALUE!');
		// Case #7: Error, String. Error in text ? propagates #N/A
		oParser = new parserFormula('REGEXTEST(NA(),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(NA(),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Error in text ? propagates #N/A');
		// Case #8: String, Error. Error in pattern ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc",NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc",NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Error. Error in pattern ? #VALUE!');
		// Case #9: Area, String. Multi-cell range ? #VALUE!
		oParser = new parserFormula('REGEXTEST(A104:A105,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(A104:A105,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'FALSE', 'Test: Negative case: Area, String. Multi-cell range ? #VALUE!');
		// Case #10: String, String, Area. case_sensitivity as multi-cell range ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","abc",A106:A107)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","abc",A106:A107) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'TRUE', 'Test: Negative case: String, String, Area. case_sensitivity as multi-cell range ? #VALUE!');
		// Case #11: Area3D, String. 3D multi-cell range ? #VALUE!
		oParser = new parserFormula('REGEXTEST(Sheet2!A5:B5,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(Sheet2!A5:B5,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'FALSE', 'Test: Negative case: Area3D, String. 3D multi-cell range ? #VALUE!');
		// Case #12: Table. Table with multiple rows/columns ? #VALUE!
		oParser = new parserFormula('REGEXTEST(Table1[Column2],Table1[Column3],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(Table1[Column2],Table1[Column3],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Table. Table with multiple rows/columns ? #VALUE!');
		// Case #13: String, String, Formula. case_sensitivity formula returns error ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","abc",IFERROR(1/0,2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","abc",IFERROR(1/0,2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Formula. case_sensitivity formula returns error ? #VALUE!');
		// Case #14: String, String. Multiline with ^$ ? FALSE (no /m flag)
		oParser = new parserFormula('REGEXTEST("line1\nline2","^line2$",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("line1\nline2","^line2$",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String, String. Multiline with ^$ ? FALSE (no /m flag)');
		// Case #15: String, String, Number. Inline flag (?i) ignored when case_sensitivity=0 ? FALSE
		oParser = new parserFormula('REGEXTEST("Hello","(?i)hello",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("Hello","(?i)hello",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String, String, Number. Inline flag (?i) ignored when case_sensitivity=0 ? FALSE');
		// Case #16: Empty, String. Empty text, non-empty pattern ? FALSE (no error)
		oParser = new parserFormula('REGEXTEST("","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Empty, String. Empty text, non-empty pattern ? FALSE (no error)');
		// Case #17: String, String, Number. quantifier more than 65535 cause #VALUE!
		// 65535 - max 16 bit unsigned int
		oParser = new parserFormula('REGEXTEST("abc","[a-z]{100000}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","[a-z]{100000}") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. uantifier more than 65535 cause #VALUE!');
		// Case #18: Name, Name. Named range is multi-cell ? #VALUE!
		oParser = new parserFormula('REGEXTEST(TestNameArea2,TestNamePattern)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(TestNameArea2,TestNamePattern) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Name, Name. Named range is multi-cell ? #VALUE!');
		// Case #19: Name3D. 3D named range is multi-cell ? #VALUE!
		oParser = new parserFormula('REGEXTEST(TestNameArea3D2,TestNamePattern3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(TestNameArea3D2,TestNamePattern3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Name3D. 3D named range is multi-cell ? #VALUE!');
		// Case #20: String, String, Number. Invalid quantifier without preceding token ? #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","*")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","*") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Invalid quantifier without preceding token ? #VALUE!');
		// Case #21: String, String, Number. Quantifier equal 65535 cause result
		oParser = new parserFormula('REGEXTEST("abc","[a-z]{65535}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","[a-z]{65535}") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: String, String, Number. Quantifier equal 65535 cause result');
		// Case #22: String, String, Number. Invalid quantifier when m > n cause #VALUE!
		oParser = new parserFormula('REGEXTEST("abc","[a-z]{2,1}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","[a-z]{2,1}") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Invalid quantifier when m > n cause #VALUE!');
		// Case #23: String, String, Number. Quantifier when m is empty cause result
		oParser = new parserFormula('REGEXTEST("abc","[a-z]{2,}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("abc","[a-z]{2,}") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: String, String, Number. Quantifier when m is empty cause result');


		// Bounded cases:
		// Case #1: String, String. Very long string and exact pattern (near Excel limit)
		oParser = new parserFormula('REGEXTEST(REPT("a",10000),"^a{10000}$")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST(REPT("a",10000),"^a{10000}$") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String, String. Very long string and exact pattern (near Excel limit)');
		// Case #2: String, String. Empty string with empty-only pattern ? TRUE
		oParser = new parserFormula('REGEXTEST("","^$")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("","^$") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String, String. Empty string with empty-only pattern ? TRUE');
		// Case #3: String, String, Number. Minimal valid case-insensitive match
		oParser = new parserFormula('REGEXTEST("A","a",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("A","a",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Bounded case: String, String, Number. Minimal valid case-insensitive match');
		// Case #4: String, String. Unicode escape sequence (PCRE2 supports \u00A9)
		oParser = new parserFormula('REGEXTEST("© Excel 2025","\\\\u00A9")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXTEST("© Excel 2025","\\\\u00A9") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Bounded case: String, String. Unicode escape sequence (PCRE2 supports \u00A9)');



		// testArrayFormula2(assert, "REGEXTEST", 1, 1);
	});

	QUnit.test("Test: \"REGEXEXTRACT\"", function (assert) {

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A107").setValue("[a-z]");
		ws.getRange2("A108").setValue("[a-z]{2}");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1ssssss2"); // Text (Column2)
		ws.getRange2("C601").setValue("[A-z]"); // Text (Column3)
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
		// Case #0: String, String. Basic literal digits extraction, return_mode default 0, case default 0
		oParser = new parserFormula('REGEXEXTRACT("Price: 199 USD","[0-9]+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Price: 199 USD","[0-9]+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '199', 'Test: Positive case: String, String. Basic literal digits extraction, return_mode default 0, case default 0');
		// Case #1: String, String, Number. return_mode = 1 ? returns all matches as array
		oParser = new parserFormula('REGEXEXTRACT("hello WORLD hello","hello",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("hello WORLD hello","hello",1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'hello', 'Test: Positive case: String, String, Number. return_mode = 1 ? returns all matches as array');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), 'hello', 'Test: Positive case: String, String, Number. return_mode = 1 ? returns all matches as array');
		// Case #2: String, String, Number, Number. Case insensitive, return all matches
		oParser = new parserFormula('REGEXEXTRACT("Cat cat CAT","cat",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Cat cat CAT","cat",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Cat', 'Test: Positive case: String, String, Number, Number. Case insensitive, return all matches');
		// Case #3: String, String, Number. return_mode 2 ? returns capturing groups from first match
		oParser = new parserFormula('REGEXEXTRACT("John Doe, 25 years, ID: A12345","(\\w+) (\\w+).+ID: (\\w+)",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("John Doe, 25 years, ID: A12345","(\\w+) (\\w+).+ID: (\\w+)",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'John', 'Test: Positive case: String, String, Number. return_mode 2 ? returns capturing groups from first match');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), 'Doe', 'Test: Positive case: String, String, Number. return_mode 2 ? returns capturing groups from first match');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,2).getValue(), 'A12345', 'Test: Positive case: String, String, Number. return_mode 2 ? returns capturing groups from first match');
		// Case #4: String, String. Realistic email extraction
		oParser = new parserFormula('REGEXEXTRACT("email: user@example.com","[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\\\.[A-Za-z]{2,}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("email: user@example.com","[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\\\.[A-Za-z]{2,}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: String, String. Realistic email extraction');
		// Case #5: String, String. Date in text
		oParser = new parserFormula('REGEXEXTRACT("Order #2025-12-10-ABC","\\d{4}-\\d{2}-\\d{2}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Order #2025-12-10-ABC","\\d{4}-\\d{2}-\\d{2}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '2025-12-10', 'Test: Positive case: String, String. Date in text');
		// Case #6: String, String, Number, Number. Case insensitive + digits
		oParser = new parserFormula('REGEXEXTRACT("TeSt123 test456","test\\d+",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("TeSt123 test456","test\\d+",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'TeSt123', 'Test: Positive case: String, String, Number, Number. Case insensitive + digits');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), 'test456', 'Test: Positive case: String, String, Number, Number. Case insensitive + digits');
		// Case #7: Formula, String. text from CONCAT formula
		oParser = new parserFormula('REGEXEXTRACT(CONCAT("Price: ",499," EUR"),"[0-9]+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(CONCAT("Price: ",499," EUR"),"[0-9]+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '499', 'Test: Positive case: Formula, String. text from CONCAT formula');
		// Case #8: String, Formula. pattern built with formula
		oParser = new parserFormula('REGEXEXTRACT("abc123",CONCAT("[a-z]+","[0-9]+"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc123",CONCAT("[a-z]+","[0-9]+")) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'abc123', 'Test: Positive case: String, Formula. pattern built with formula');
		// Case #11: Array, String. Array in text argument, spills correctly
		oParser = new parserFormula('REGEXEXTRACT({"Test1";"Test2";"ABC3"},"[A-Z]+[0-9]")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT({"Test1";"Test2";"ABC3"},"[A-Z]+[0-9]") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: Array, String. Array in text argument, spills correctly');
		// Case #12: String, String, Empty. return_mode omitted ? defaults to 0
		oParser = new parserFormula('REGEXEXTRACT("Find me","Find me",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Find me","Find me",) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Find me', 'Test: Positive case: String, String, Empty. return_mode omitted ? defaults to 0');
		// Case #13: String, String, Number, Empty. case_sensitivity omitted ? case sensitive ? only one match
		oParser = new parserFormula('REGEXEXTRACT("ABC abc","abc",1,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("ABC abc","abc",1,) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'abc', 'Test: Positive case: String, String, Number, Empty. case_sensitivity omitted ? case sensitive ? only one match');
		// Case #14: Name, Name. Named ranges
		oParser = new parserFormula('REGEXEXTRACT(TestName,TestNamePattern)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(TestName,TestNamePattern) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Positive case: Name, Name. Named ranges');
		// Case #15: Name3D, Name3D. 3D named ranges
		oParser = new parserFormula('REGEXEXTRACT(TestName3D,TestName3DPattern)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(TestName3D,TestName3DPattern) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Positive case: Name3D, Name3D. 3D named ranges');
		// Case #16: Ref3D, Ref3D. 3D references
		oParser = new parserFormula('REGEXEXTRACT(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: Ref3D, Ref3D. 3D references');
		// Case #17: Area3D. Area3D single cell, pattern in A104
		oParser = new parserFormula('REGEXEXTRACT(Sheet2!A3:A3,Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(Sheet2!A3:A3,Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Text', 'Test: Positive case: Area3D. Area3D single cell, pattern in A104');
		// Case #18: Table. Structured table references
		oParser = new parserFormula('REGEXEXTRACT(Table1[Column2],Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(Table1[Column2],Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 's', 'Test: Positive case: Table. Structured table references');
		// Case #19: Date. Date serial converts to text automatically
		oParser = new parserFormula('REGEXEXTRACT(DATE(2025,6,15),"\\\\d{4}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(DATE(2025,6,15),"\\\\d{4}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: Date. Date serial converts to text automatically');
		// Case #20: Formula. Simple positive case with escape
		oParser = new parserFormula('REGEXEXTRACT("Code: X99Y88","X\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Code: X99Y88","X\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: Formula. Simple positive case with escape');
		// Case #21: String, String, Number. No match but valid regex ? returns first (empty? wait — actually #N/A, but we will move to Negative)
		oParser = new parserFormula('REGEXEXTRACT("No match here","\\d+",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("No match here","\\d+",0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: String, String, Number. No match but valid regex ? returns first (empty? wait — actually #N/A, but we will move to Negative)');
		// Case #22: String, String, Number. return_mode 2 with multiple groups ? array of groups
		oParser = new parserFormula('REGEXEXTRACT("First Second Third","(\\\\w+)",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("First Second Third","(\\\\w+)",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: String, String, Number. return_mode 2 with multiple groups ? array of groups');
		// Case #23: String, String. Using capture group to trim
		oParser = new parserFormula('REGEXEXTRACT("  trim me  ","\\\\s*(.+)\\\\s*")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("  trim me  ","\\\\s*(.+)\\\\s*") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Positive case: String, String. Using capture group to trim');
		// Case #24: String, String. Empty string as text
		oParser = new parserFormula('REGEXEXTRACT("",".*")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("",".*") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '', 'Test: Positive case: String, String. Empty string as text');
		// Case #25: String, String. Empty string as pattern
		oParser = new parserFormula('REGEXEXTRACT("anything","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("anything","") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '', 'Test: Positive case: String, String. Empty string as pattern');
		// Case #26: Formula. REGEXEXTRACT inside IF
		oParser = new parserFormula('IF(TRUE,REGEXEXTRACT("yes123","\\\\d+"),"")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: IF(TRUE,REGEXEXTRACT("yes123","\\\\d+"),"") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Formula. REGEXEXTRACT inside IF');
		// Case #27: Formula. REGEXEXTRACT as part of another formula
		oParser = new parserFormula('CONCAT("Result: ",REGEXEXTRACT("Cost: 500","\\\\d+"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT("Result: ",REGEXEXTRACT("Cost: 500","\\\\d+")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Formula. REGEXEXTRACT as part of another formula');
		// Case #28: String, String, Number, Number. Final positive with all args
		oParser = new parserFormula('REGEXEXTRACT("DATA data Data","data",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("DATA data Data","data",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'DATA', 'Test: Positive case: String, String, Number, Number. Final positive with all args');

		// Negative cases:
		// Case #1: Empty, String. text is empty cell ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(,"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(,"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Empty, String. text is empty cell ? #VALUE!');
		// Case #3: String, String, Number. No match + return_mode 0 ? #N/A
		oParser = new parserFormula('REGEXEXTRACT("no digits","\\\\d+",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("no digits","\\\\d+",0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String, Number. No match + return_mode 0 ? #N/A');
		// Case #4: String, String, Number. Unclosed parenthesis ? #VALUE! (syntax error in regex)
		oParser = new parserFormula('REGEXEXTRACT("abc","(unclosed",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","(unclosed",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Unclosed parenthesis ? #VALUE! (syntax error in regex)');
		// Case #5: String, String. Unclosed bracket ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","[a-z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","[a-z") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: String, String. Unclosed bracket ? #VALUE!');
		// Case #6: Number, String. Number auto-converted to text ? works, but we mark as Negative for strictness ? actually works in Excel 365 ? move to Positive? No, keep here for discussion
		oParser = new parserFormula('REGEXEXTRACT(123,"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(123,"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Number, String. Number auto-converted to text ? works, but we mark as Negative for strictness ? actually works in Excel 365 ? move to Positive? No, keep here for discussion');
		// Case #6: String, Number. pattern as number ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc123",123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc123",123) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '123', 'Test: Negative case: String, Number. pattern as number ? #VALUE!');
		// Case #7: String, String, String. return_mode as text ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("text","pattern","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("text","pattern","0") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String, String. return_mode as text ? #VALUE!');
		// Case #8: String, String, Number, String. case_sensitivity as text ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("AbC","abc",1,"1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("AbC","abc",1,"1") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'AbC', 'Test: Negative case: String, String, Number, String. case_sensitivity as text ? #VALUE!');
		// Case #9: String, String, Number. No match + return_mode 1 ? #N/A (spill error)
		oParser = new parserFormula('REGEXEXTRACT("match","nomatch",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("match","nomatch",1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String, Number. No match + return_mode 1 ? #N/A (spill error)');
		// Case #10: Error. Error in text ? propagates #N/A
		oParser = new parserFormula('REGEXEXTRACT(NA(),"\\\\w+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(NA(),"\\\\w+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error in text ? propagates #N/A');
		// Case #11: Error. Error in pattern ? #N/A
		oParser = new parserFormula('REGEXEXTRACT("text",NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("text",NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error in pattern ? #N/A');
		// Case #12: Area. Multi-cell range in text ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(A105:A106,"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(A105:A106,"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Area. Multi-cell range in text ? #VALUE!');
		// Case #13: String, Area. Multi-cell range in pattern
		oParser = new parserFormula('REGEXEXTRACT("text",A107:A108)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("text",A107:A108) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 't', 'Test: Negative case: String, Area. Multi-cell range in pattern');
		assert.strictEqual(oParser.calculate().getElementRowCol(1,0).getValue(), 'te', 'Test: Negative case: String, Area. Multi-cell range in pattern');
		// Case #14: Array. Array in text + return_mode default ? #N/A on most cells
		oParser = new parserFormula('REGEXEXTRACT({"a";"b";"c"},"[0-9]")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT({"a";"b";"c"},"[0-9]") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Array. Array in text + return_mode default ? #N/A on most cells');
		// Case #15: String, String, Number. Invalid return_mode ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","\\\\d+",99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","\\\\d+",99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Invalid return_mode ? #VALUE!');
		// Case #16: String, String, Number, Number. Invalid case_sensitivity ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","\\\\d+",0,99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","\\\\d+",0,99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, Number, Number. Invalid case_sensitivity ? #VALUE!');
		// Case #17: String, Boolean. pattern as boolean ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("true",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("true",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, Boolean. pattern as boolean ? #VALUE!');
		// Case #18: Boolean, String. text as boolean ? works (converts to "TRUE"), but strict ? #VALUE in some contexts ? actually works ? keep as Negative for caution
		oParser = new parserFormula('REGEXEXTRACT(TRUE,"\\\\w+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(TRUE,"\\\\w+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Boolean, String. text as boolean ? works (converts to "TRUE"), but strict ? #VALUE in some contexts ? actually works ? keep as Negative for caution');
		// Case #19: Area3D. Multi-cell Area3D in text ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(Sheet2!A4:A5,"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(Sheet2!A4:A5,"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Area3D. Multi-cell Area3D in text ? #VALUE!');
		// Case #20: Name. Named range refers to 2 cells ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(TestNameArea2,"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(TestNameArea2,"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Name. Named range refers to 2 cells ? #VALUE!');
		// Case #21: Table. Table column with >1 row ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(Table1[Column2],"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(Table1[Column2],"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: Table. Table column with >1 row ? #VALUE!');
		// Case #22: String, String, Number. Unclosed capturing group ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","(abc",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","(abc",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Unclosed capturing group ? #VALUE!');
		// Case #23: String, String, Number. Invalid regex syntax ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc",")abc(",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc",")abc(",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: String, String, Number. Invalid regex syntax ? #VALUE!');
		// Case #24: String, String, Number, Number. Case insensitive but no match ? #N/A
		oParser = new parserFormula('REGEXEXTRACT("ABC","abc",0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("ABC","abc",0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'ABC', 'Test: Negative case: String, String, Number, Number. Case insensitive but no match ? #N/A');
		// Case #25: String, String. Empty text + requiring digit ? #N/A
		oParser = new parserFormula('REGEXEXTRACT("","\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("","\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String. Empty text + requiring digit ? #N/A');
		// Case #27: String, String, Number. Possessive quantifiers
		// oParser = new parserFormula('REGEXEXTRACT("test",".*+",0)', 'A2', ws);
		// assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("test",".*+",0) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'test', 'Test: Negative case: String, String, Number. Possessive quantifier not supported ? #VALUE!');
		// Case #28: String, String, Number. Unicode property not supported in PCRE2 Excel ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","\\\\p{L}",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","\\\\p{L}",0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String, Number. Unicode property not supported in PCRE2 Excel ? #VALUE!');
		// Case #29: String, String, Number. Pattern too complex / too long match attempt ? may cause timeout or #VALUE!
		oParser = new parserFormula('REGEXEXTRACT(REPT("a",10000),"a{10001}",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(REPT("a",10000),"a{10001}",0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Negative case: String, String, Number. Pattern too complex / too long match attempt ? may cause timeout or #VALUE!');
		// Case #30: String, String, Number. Named capture groups not supported ? #VALUE!
		oParser = new parserFormula('REGEXEXTRACT("abc","(?<name>abc)",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc","(?<name>abc)",2) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'abc', 'Test: Negative case: String, String, Number. Named capture groups not supported ? #VALUE!');

		// Bounded cases:
		let res = "x";
		// Case #1: String, String. Maximum text length in Excel (32767 chars)
		oParser = new parserFormula('REGEXEXTRACT(REPT("x",32767),"x{32767}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(REPT("x",32767),"x{32767}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), res.repeat(32767), 'Test: Bounded case: String, String. Maximum text length in Excel (32767 chars)');
		// Case #2: String, String. Maximum quantifier value
		oParser = new parserFormula('REGEXEXTRACT("a","a{1,32767}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("a","a{1,32767}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'a', 'Test: Bounded case: String, String. Maximum quantifier value');
		// Case #3: String, String, Number. Largest possible number as text
		oParser = new parserFormula('REGEXEXTRACT("Start 999999999999999 End","\\\\d+",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("Start 999999999999999 End","\\\\d+",0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Bounded case: String, String, Number. Largest possible number as text');
		// Case #4: String, String, Number. Smallest scientific notation
		oParser = new parserFormula('REGEXEXTRACT("1E-307","[0-9E.+-]+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("1E-307","[0-9E.+-]+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), "1E-307", 'Test: Bounded case: String, String, Number. Smallest scientific notation');
		// Case #5: String, String, Number. Maximum range quantifier
		oParser = new parserFormula('REGEXEXTRACT("abc",".{0,32767}")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("abc",".{0,32767}") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'abc', 'Test: Bounded case: String, String, Number. Maximum range quantifier');
		// Case #6: String, String, Number, Number. return_mode=1 on very long possible output (limited by Excel)
		oParser = new parserFormula('REGEXEXTRACT("ABCabc","abc",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("ABCabc","abc",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'ABC', 'Test: Bounded case: String, String, Number, Number. return_mode=1 on very long possible output (limited by Excel)');
		// Case #7: String, String, Number. Long text before match
		oParser = new parserFormula('REGEXEXTRACT(REPT("a",1000)&"999", "\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT(REPT("a",1000)&"999", "\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Bounded case: String, String, Number. Long text before match');
		// Case #8: String, String, Number. Long text after match
		oParser = new parserFormula('REGEXEXTRACT("999"&REPT("a",1000),"\\\\d+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("999"&REPT("a",1000),"\\\\d+") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#N/A', 'Test: Bounded case: String, String, Number. Long text after match');
		// Case #9: String, String, Number, Number. return_mode 2 with single char group ? many groups possible, bounded by cell limit
		oParser = new parserFormula('REGEXEXTRACT("AaBbCc","(.)",2,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("AaBbCc","(.)",2,0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Bounded case: String, String, Number, Number. return_mode 2 with single char group ? many groups possible, bounded by cell limit');
		// Case #10: String, String. Two empty strings as text and pattern 
		oParser = new parserFormula('REGEXEXTRACT("","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXEXTRACT("","") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '', 'Test: Bounded case: String, String. Two empty strings as text and pattern ');
		

		// Need to fix: pcre2 quantificators errror
		// Case #27: String, String, Number. Possessive quantifier


		// testArrayFormula2(assert, "REGEXEXTRACT", 1, 1);
	});

	QUnit.test("Test: \"REGEXREPLACE\"", function (assert) {

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("");
		ws.getRange2("A105").setValue("sd123");
		ws.getRange2("A106").setValue("");
		ws.getRange2("A107").setValue("[a-z]");
		ws.getRange2("A108").setValue("");
		ws.getRange2("A109").setValue("");
		ws.getRange2("A110").setValue("TRUE");
		ws.getRange2("A111").setValue("FALSE");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 2);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1ssssss2"); // Text (Column2)
		ws.getRange2("C601").setValue("[A-z]"); // Text (Column3)
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
		// Case #0: String, String, String. Basic replace all (occurrence default 0), case sensitive
		oParser = new parserFormula('REGEXREPLACE("Hello world hello","hello","hi")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("Hello world hello","hello","hi") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello world hi', 'Test: Positive case: String, String, String. Basic replace all (occurrence default 0), case sensitive');
		// Case #1: String, String, String, Number. Explicit occurrence 0 (all), case sensitive
		oParser = new parserFormula('REGEXREPLACE("cat CAT cat","cat","dog",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("cat CAT cat","cat","dog",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'dog CAT dog', 'Test: Positive case: String, String, String, Number. Explicit occurrence 0 (all), case sensitive');
		// Case #2: String, String, String, Number, Number. Case insensitive replace all
		oParser = new parserFormula('REGEXREPLACE("Cat cat CAT","cat","dog",0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("Cat cat CAT","cat","dog",0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'dog dog dog', 'Test: Positive case: String, String, String, Number, Number. Case insensitive replace all');
		// Case #3: String, String, String, Number. Replace first occurrence only
		oParser = new parserFormula('REGEXREPLACE("one two three two four","two","2",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two three two four","two","2",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one 2 three two four', 'Test: Positive case: String, String, String, Number. Replace first occurrence only');
		// Case #4: String, String, String, Number. Replace second occurrence
		oParser = new parserFormula('REGEXREPLACE("one two three two four","two","2",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two three two four","two","2",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one two three 2 four', 'Test: Positive case: String, String, String, Number. Replace second occurrence');
		// Case #5: String, String, String, Number. Negative occurrence: replace last (from end)
		oParser = new parserFormula('REGEXREPLACE("one two three two four","two","2",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two three two four","two","2",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one two three 2 four', 'Test: Positive case: String, String, String, Number. Negative occurrence: replace last (from end)');
		// Case #6: String, String, String, Number. Negative occurrence: second from end
		oParser = new parserFormula('REGEXREPLACE("one two three two four","two","2",-2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two three two four","two","2",-2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one 2 three two four', 'Test: Positive case: String, String, String, Number. Negative occurrence: second from end');
		// Case #7: String, String, String. Using capturing groups in replacement
		oParser = new parserFormula('REGEXREPLACE("John Doe","(\\w+) (\\w+)","$2, $1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("John Doe","(\\w+) (\\w+)","$2, $1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Doe, John', 'Test: Positive case: String, String, String. Using capturing groups in replacement');
		// Case #8: String, String, String. Mask digits
		oParser = new parserFormula('REGEXREPLACE("Price: 100 USD","\\d+","***")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("Price: 100 USD","\\d+","***") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Price: *** USD', 'Test: Positive case: String, String, String. Mask digits');
		// Case #9: String, String, String. Normalize spaces
		oParser = new parserFormula('REGEXREPLACE(" hello world ","\\s+","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(" hello world ","\\s+","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'helloworld', 'Test: Positive case: String, String, String. Normalize spaces');
		// Case #10: Formula, String, String. Text from formula
		oParser = new parserFormula('REGEXREPLACE(CONCAT("Test ","abc123"),"\\\\d+","XYZ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(CONCAT("Test ","abc123"),"\\\\d+","XYZ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test abc123', 'Test: Positive case: Formula, String, String. Text from formula');
		// Case #11: String, Formula, String. Pattern from formula
		oParser = new parserFormula('REGEXREPLACE("abc123",CONCAT("\\\\d","+"),"XYZ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc123",CONCAT("\\\\d","+"),"XYZ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc123', 'Test: Positive case: String, Formula, String. Pattern from formula');
		// Case #12: String, String, Formula. Replacement from formula
		oParser = new parserFormula('REGEXREPLACE("hello", "hello", UPPER("hi"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("hello", "hello", UPPER("hi")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HI', 'Test: Positive case: String, String, Formula. Replacement from formula');
		// Case #15: Array, String, String. Array in text, replaces in each
		oParser = new parserFormula('REGEXREPLACE({"text1";"text2"},"\\d","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE({"text1";"text2"},"\\d","0") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'text0', 'Test: Positive case: Array, String, String. Array in text, replaces in each');
		assert.strictEqual(oParser.calculate().getElementRowCol(1,0).getValue(), 'text0', 'Test: Positive case: Array, String, String. Array in text, replaces in each');
		// Case #16: Name, Name, Name. Named ranges
		oParser = new parserFormula('REGEXREPLACE(TestName,TestName1,TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(TestName,TestName1,TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-10.5', 'Test: Positive case: Name, Name, Name. Named ranges');
		// Case #17: Name3D, Name3D, Name3D. 3D named ranges
		oParser = new parserFormula('REGEXREPLACE(TestName3D,TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(TestName3D,TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name3D, Name3D, Name3D. 3D named ranges');
		// Case #18: Ref3D, Ref3D, Ref3D. 3D references
		oParser = new parserFormula('REGEXREPLACE(Sheet2!A1,Sheet2!A2,Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(Sheet2!A1,Sheet2!A2,Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D, Ref3D, Ref3D. 3D references');
		// Case #19: Area3D. Area3D single cell, pattern/repl in cells
		oParser = new parserFormula('REGEXREPLACE(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(Sheet2!A1:A1,Sheet2!A2:A2,Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Area3D. Area3D single cell, pattern/repl in cells');
		// Case #20: Table. Structured table references
		oParser = new parserFormula('REGEXREPLACE(Table1[Column1],Table1[Column2],Table1[Column3])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(Table1[Column1],Table1[Column2],Table1[Column3]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table. Structured table references');
		// Case #21: Date. Date converted to text
		oParser = new parserFormula('REGEXREPLACE(TEXT(DATE(2025,12,15),"yyyy-mm-dd"),"-","/")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(TEXT(DATE(2025,12,15),"yyyy-mm-dd"),"-","/") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2025/12/15', 'Test: Positive case: Date. Date converted to text');
		// Case #22: String, String, String, Empty. occurrence omitted ? all
		oParser = new parserFormula('REGEXREPLACE("hello hello","hello","hi",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("hello hello","hello","hi",) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'hi hi', 'Test: Positive case: String, String, String, Empty. occurrence omitted ? all');
		// Case #23: String, String, String, Number, Empty. case_sensitivity omitted ? case sensitive
		oParser = new parserFormula('REGEXREPLACE("Hello hello","hello","hi",0,)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("Hello hello","hello","hi",0,) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello hi', 'Test: Positive case: String, String, String, Number, Empty. case_sensitivity omitted ? case sensitive');
		// Case #24: String, String, String. No match ? returns original text
		oParser = new parserFormula('REGEXREPLACE("no match","\\d+","XXX")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("no match","\\d+","XXX") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'no match', 'Test: Positive case: String, String, String. No match ? returns original text');
		// Case #26: Formula. Basic with backreference-like replacement
		oParser = new parserFormula('REGEXREPLACE("data123","\\d+","[numbers]")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("data123","\\d+","[numbers]") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'data[numbers]', 'Test: Positive case: Formula. Basic with backreference-like replacement');
		// Case #27: Formula. Nested in another formula
		oParser = new parserFormula('CONCAT("Result: ",REGEXREPLACE("Cost: 500","\\\\d+","***"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT("Result: ",REGEXREPLACE("Cost: 500","\\\\d+","***")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Result: Cost: 500', 'Test: Positive case: Formula. Nested in another formula');
		// Case #28: String, String, String, Number, Number. All arguments used
		oParser = new parserFormula('REGEXREPLACE("Abc abc ABC","abc","XYZ",0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("Abc abc ABC","abc","XYZ",0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XYZ XYZ XYZ', 'Test: Positive case: String, String, String, Number, Number. All arguments used');

		// Negative cases:
		// Case #2: String, Empty, String. pattern empty ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text","","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XtXeXxXtX', 'Test: Negative case: String, Empty, String. pattern empty ? #VALUE!');
		// Case #3: String, String, Empty. replacement empty ? replaces with nothing (valid, but mark negative if strict)
		oParser = new parserFormula('REGEXREPLACE("text","\\\\w+","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","\\\\w+","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Negative case: String, String, Empty. replacement empty ? replaces with nothing (valid, but mark negative if strict)');
		// Case #4: String, String, String. Invalid regex syntax ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc","(unclosed","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","(unclosed","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Invalid regex syntax ? #VALUE!');
		// Case #5: String, String, String. Unclosed bracket ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc","[a-z","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","[a-z","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Unclosed bracket ? #VALUE!');
		// Case #6: Number, String, String. Number auto-converted to text ? works, but strict type ? negative
		oParser = new parserFormula('REGEXREPLACE(12345,"\\d","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(12345,"\\d","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '00000', 'Test: Negative case: Number, String, String. Number auto-converted to text ? works, but strict type ? negative');
		// Case #7: String, Number, String. pattern as number ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc123",123,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc123",123,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abcX', 'Test: Negative case: String, Number, String. pattern as number ? #VALUE!');
		// Case #8: String, String, Number. replacement as number ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text","pattern",123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","pattern",123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Negative case: String, String, Number. replacement as number ? #VALUE!');
		// Case #9: String, String, String, String. occurrence as text ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text","pattern","X","1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","pattern","X","1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Negative case: String, String, String, String. occurrence as text ? #VALUE!');
		// Case #10: String, String, String, Number, String. case_sensitivity as text ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text","pattern","X",0,"1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","pattern","X",0,"1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Negative case: String, String, String, Number, String. case_sensitivity as text ? #VALUE!');
		// Case #11: Error. Error in text ? propagates #N/A
		oParser = new parserFormula('REGEXREPLACE(NA(),"\\w+","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(NA(),"\\w+","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error in text ? propagates #N/A');
		// Case #12: Error. Error in pattern ? #VALUE! or #N/A
		oParser = new parserFormula('REGEXREPLACE("text",NA(),"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text",NA(),"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error in pattern ? #VALUE! or #N/A');
		// Case #13: Area. Multi-cell range in text
		oParser = new parserFormula('REGEXREPLACE(A105:A106,"\\d+","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(A105:A106,"\\d+","X") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'sdX', 'Test: Negative case: Area. Multi-cell range in text');
		assert.strictEqual(oParser.calculate().getElementRowCol(1,0).getValue(), '', 'Test: Negative case: Area. Multi-cell range in text');
		// Case #14: String, Area, String. Multi-cell in pattern ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text",A108:A109,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text",A108:A109,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'XtXeXxXtX', 'Test: Negative case: String, Area, String. Multi-cell in pattern ? #VALUE!');
		// Case #15: String, String, Area. Multi-cell in replacement ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("text","pattern",A110:A111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("text","pattern",A110:A111) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'text', 'Test: Negative case: String, String, Area. Multi-cell in replacement ? #VALUE!');
		// Case #16: String, String, String, Number. occurrence too large ? original text or #VALUE!
		oParser = new parserFormula('REGEXREPLACE("one two","two","X",99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two","two","X",99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one two', 'Test: Negative case: String, String, String, Number. occurrence too large ? original text or #VALUE!');
		// Case #17: String, String, String, Number. negative occurrence too large ? original or #VALUE!
		oParser = new parserFormula('REGEXREPLACE("one two","two","X",-99)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two","two","X",-99) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'one two', 'Test: Negative case: String, String, String, Number. negative occurrence too large ? original or #VALUE!');
		// Case #18: Boolean, String, String. Boolean text ? works, but strict ? negative
		oParser = new parserFormula('REGEXREPLACE(TRUE,"TRUE","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(TRUE,"TRUE","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X', 'Test: Negative case: Boolean, String, String. Boolean text ? works, but strict ? negative');
		// Case #19: Area3D. Multi-cell Area3D
		oParser = new parserFormula('REGEXREPLACE(Sheet2!A5:A6,"","WW")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(Sheet2!A5:A6,"","WW") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'WW', 'Test: Negative case: Area3D. Multi-cell Area3D');
		assert.strictEqual(oParser.calculate().getElementRowCol(1,0).getValue(), 'WW', 'Test: Negative case: Area3D. Multi-cell Area3D');
		// Case #20: Name. Named area multi-cell
		oParser = new parserFormula('REGEXREPLACE(TestNameArea2,"\\d+","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(TestNameArea2,"\\d+","X") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'X.X', 'Test: Negative case: Name. Named area multi-cell');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,1).getValue(), '-X.X', 'Test: Negative case: Name. Named area multi-cell');
		// Case #21: Table. Table with multi-row 
		oParser = new parserFormula('REGEXREPLACE(Table1[Column2],Table1[Column3],Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(Table1[Column2],Table1[Column3],Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '11111112', 'Test: Negative case: Table. Table with multi-row');
		// Case #22: String, String, String. Unicode property not supported ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc","\\p{L}","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","\\p{L}","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String, String, String. Unicode property not supported ? #VALUE!');
		// Case #23: String, String, String. Named groups in pattern ? #VALUE! (if not supported)
		oParser = new parserFormula('REGEXREPLACE("abc","(?<name>abc)","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","(?<name>abc)","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X', 'Test: Negative case: String, String, String. Named groups in pattern ? #VALUE! (if not supported)');
		// Case #24: String, String, String. Possessive quantifier issue
		oParser = new parserFormula('REGEXREPLACE("test",".*+","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("test",".*+","X") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'XX', 'Test: Negative case: String, String, String. Possessive quantifier issue ? #VALUE!');
		// Case #25: String, String, String, Number. occurrence fractional ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc abc","abc","X",0.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc abc","abc","X",0.5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X X', 'Test: Negative case: String, String, String, Number. occurrence fractional ? #VALUE!');
		// Case #26: String, String, String, Number, Number. Invalid case_sensitivity ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("abc","abc","X",1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","abc","X",1,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String, Number, Number. Invalid case_sensitivity ? #VALUE!');
		// Case #27: String, String, String. Text too long ? may error
		oParser = new parserFormula('REGEXREPLACE(REPT("a",32768),"a","b")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(REPT("a",32768),"a","b") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Text too long ? may error');
		// Case #28: Formula. Pattern empty via formula ? #VALUE!
		oParser = new parserFormula('REGEXREPLACE("data",IF(TRUE,"","invalid"),"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("data",IF(TRUE,"","invalid"),"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XdXaXtXaX', 'Test: Negative case: Formula. Pattern empty via formula ? #VALUE!');
		// Case #29: String, String, String. Invalid backreference ? treats as literal or error
		oParser = new parserFormula('REGEXREPLACE("abc","abc","$10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("abc","abc","$10") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Invalid backreference ? treats as literal or error');
		// Case #30: String, String, String. Multiline with ^/$ issues if any ? but works
		oParser = new parserFormula('REGEXREPLACE("line1\nline2","line1","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("line1\nline2","line1","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X\nline2', 'Test: Negative case: String, String, String. Multiline with ^/$ issues if any ? but works');

		let str = "";
		// Bounded cases:
		// Case #1: String, String, String. Max text length replace
		oParser = new parserFormula('REGEXREPLACE(REPT("a",32767),"a{32767}","b")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(REPT("a",32767),"a{32767}","b") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'b', 'Test: Bounded case: String, String, String. Max text length replace');
		// Case #2: String, String, String. Long digit match
		oParser = new parserFormula('REGEXREPLACE("123456789012345","\\d{15}","XXX")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("123456789012345","\\d{15}","XXX") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XXX', 'Test: Bounded case: String, String, String. Long digit match');
		// Case #3: String, String, String, Number. Long text before match
		str = "x ";
		oParser = new parserFormula('REGEXREPLACE(REPT("x ",10000)&"match","match","Y",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(REPT("x ",10000)&"match","match","Y",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(10000) + "Y", 'String, String, String, Number. Long text before match');
		// Case #4: String, String, String, Number. Long text after match
		str = " x";
		oParser = new parserFormula('REGEXREPLACE("match"&REPT(" x",10000),"match","Y",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("match"&REPT(" x",10000),"match","Y",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "Y" + str.repeat(10000), 'String, String, String, Number. Long text after match');
		// Case #5: String, String, String. Large group number ? may error or literal
		oParser = new parserFormula('REGEXREPLACE("a","(a){1000}","$1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("a","(a){1000}","$1") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Bounded case: String, String, String. Large group number ? may error or literal');
		// Case #6: String, String, String. Empty text full match
		oParser = new parserFormula('REGEXREPLACE("",".*","empty")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("",".*","empty") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'empty', 'Test: Bounded case: String, String, String. Empty text full match');
		// Case #7: String, String, String, Number. Negative occurrence bounded by matches
		oParser = new parserFormula('REGEXREPLACE("one two three","\\w+", "X", -3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("one two three","\\w+", "X", -3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X two three', 'Test: Bounded case: String, String, String, Number. Negative occurrence bounded by matches');
		// Case #8: String, String, String. Large number as text replace
		oParser = new parserFormula('REGEXREPLACE("999999999999999","9{15}","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("999999999999999","9{15}","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0", 'Test: Bounded case: String, String, String. Large number as text replace');
		// Case #9: String, String, String. Large quantifier on short text ? no replace
		oParser = new parserFormula('REGEXREPLACE("test",".{32767}","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE("test",".{32767}","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'test', 'Test: Bounded case: String, String, String. Large quantifier on short text ? no replace');
		// Case #10: String, String, String. Near max captures/backrefs
		oParser = new parserFormula('REGEXREPLACE(REPT("a",16383)&"c","(a){16384}c","$1")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: REGEXREPLACE(REPT("a",16383)&"c","(a){16384}c","$1") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String, String, String. Near max captures/backrefs');

		// Need to fix:
		// Case #24: String, String, String. Possessive quantifier issue - PCRE2 only pattern
		// Case #29: String, String, String. Invalid backreference ? treats as literal or error - excel special symbols
		// Case #10: String, String, String. Near max captures/backrefs


		// testArrayFormula2(assert, "REGEXREPLACE", 1, 1);
	});

	QUnit.test("Test: \"SEARCH\"", function (assert) {
		let array;

		oParser = new parserFormula("SEARCH(\"~*\",\"abc*dEF\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula("SEARCH(\"~\",\"abc~dEF\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula("SEARCH(\"de\",\"abcdEF\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 4);

		oParser = new parserFormula("SEARCH(\"?c*e\",\"abcdEF\")", "B2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 2);

		oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",3)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 7);

		oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",30)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("SEARCH(\"pe\",\"dEFabcdEF\",2)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",2)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 7);

		oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",0)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("SEARCH(\"de\",\"dEFabcdEF\",-2)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		// bool
		oParser = new parserFormula('SEARCH("a", "abcde", FALSE)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("a", "abcde", FALSE)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of SEARCH("a", "abcde", FALSE)');

		oParser = new parserFormula('SEARCH("a", "abcde", TRUE)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("a", "abcde", TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of SEARCH("a", "abcde", TRUE)');

		oParser = new parserFormula('SEARCH(FALSE,"abc10dTRUEFALSE")', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH(FALSE,"abc10dTRUEFALSE")');
		assert.strictEqual(oParser.calculate().getValue(), 11, 'Result of SEARCH(FALSE,"abc10dTRUEFALSE")');

		oParser = new parserFormula('SEARCH(TRUE,"abc10dTRUEFALSE")', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH(TRUE,"abc10dTRUEFALSE")');
		assert.strictEqual(oParser.calculate().getValue(), 7, 'Result of SEARCH(TRUE,"abc10dTRUEFALSE")');

		oParser = new parserFormula('SEARCH("T",TRUE)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("T",TRUE)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of SEARCH("T",TRUE)');

		oParser = new parserFormula('SEARCH("F",FALSE)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("F",FALSE)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of SEARCH("F",FALSE)');

		// err
		oParser = new parserFormula('SEARCH(#N/A, "abcde", 1)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH(#N/A, "abcde", 1)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of SEARCH(#N/A, "abcde", 1)');

		oParser = new parserFormula('SEARCH("abcde", #DIV/0!, 1)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("abcde", #DIV/0!, 1)');
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'Result of SEARCH("abcde", #DIV/0!, 1)');

		oParser = new parserFormula('SEARCH("abcde","a",#NUM!)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("abcde","a",#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#NUM!", 'Result of SEARCH("abcde","a",#NUM!)');

		oParser = new parserFormula('SEARCH(#N/A,"a",#NUM!)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH(#N/A,"a",#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Result of SEARCH(#N/A,"a",#NUM!)');

		oParser = new parserFormula('SEARCH("abcde",#DIV/0!,#NUM!)', "C2", ws);
		assert.ok(oParser.parse(), 'SEARCH("abcde",#DIV/0!,#NUM!)');
		assert.strictEqual(oParser.calculate().getValue(), "#DIV/0!", 'Result of SEARCH("abcde",#DIV/0!,#NUM!)');

		oParser = new parserFormula('SEARCH({5;6;7;1;2;3;4},123)', "K4", ws);
		oParser.setArrayFormulaRef(ws.getRange2("K4:K10").bbox);
		assert.ok(oParser.parse(), 'SEARCH({5;6;7;1;2;3;4},123)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({5;6;7;1;2;3;4},123)[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", 'Result of SEARCH({5;6;7;1;2;3;4},123)[1,0]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "#VALUE!", 'Result of SEARCH({5;6;7;1;2;3;4},123)[2,0]');
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 1, 'Result of SEARCH({5;6;7;1;2;3;4},123)[3,0]');
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 2, 'Result of SEARCH({5;6;7;1;2;3;4},123)[4,0]');
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), 3, 'Result of SEARCH({5;6;7;1;2;3;4},123)[5,0]');
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), "#VALUE!", 'Result of SEARCH({5;6;7;1;2;3;4},123)[6,0]');
		
		ws.getRange2("K1").setValue("hcd*prd");
		ws.getRange2("L1").setValue("*vmwprd*");
		ws.getRange2("M1").setValue("*vmcprd*");

		oParser = new parserFormula('SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M2)', "K4", ws);
		oParser.setArrayFormulaRef(ws.getRange2("K4:M4").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M2)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M2)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M2)[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M2)[0,2]');

		oParser = new parserFormula('SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})', "K4", ws);
		oParser.setArrayFormulaRef(ws.getRange2("K4:M7").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,2]');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,0]');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,1]');
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,2]');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,0]');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,1]');
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,2]');
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,0]');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,1]');
		assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},{"hcdpmhatlvmctst71a";"dsdsvmwprd";"hcdpmhatlvmcprd71a";"hcd*prd*vmwprd**vmcprd*"})[0,2]');

		ws.getRange2("A1").setValue("hcdpmhatlvmctst71a");
		ws.getRange2("A2").setValue("pmhwebvmwtst01");
		ws.getRange2("B1").setValue("sd");
		ws.getRange2("B2").setValue("dd");

		// array|area
		let bbox = ws.getRange2("K4").bbox;
		let cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		oParser = new parserFormula('SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M1)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("K4:M4").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M1)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 1, 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M1)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M1)[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", 'Result of SEARCH({"hcd*prd","*vmwprd*","*vmcprd*"},K1:M1)[0,2]');

		bbox = ws.getRange2("D1:F2").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		oParser = new parserFormula('SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F2").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1, 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,2]');

		bbox = ws.getRange2("E1").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		oParser = new parserFormula('SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F2").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 1, 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 1, 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:A2)[0,2]');


		bbox = ws.getRange2("D1:F2").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		oParser = new parserFormula('SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:B2)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F2").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:B2)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "#VALUE!", 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:B2)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "#VALUE!", 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:B2)[0,1]');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#VALUE!", 'Result of SEARCH({"*vmwtst*","hcd*tst","*vmctst*"}, A1:B2)[0,2]');

		ws.getRange2("A3:A5").cleanAll();

		bbox = ws.getRange2("D1:F5").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bbox.r1, bbox.c1);
		oParser = new parserFormula('SEARCH({"mhat","vmc"}, A1:A5)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F5").bbox);
		assert.ok(oParser.parse(), 'SEARCH({"mhat","vmc"}, A1:A5)');
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 5, 'Result of SEARCH({"mhat","vmc"}, A1:A5)[0,0]');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 10, 'Result of SEARCH({"mhat","vmc"}, A1:A5)[0,1]');

		ws.getRange2("S1").setValue("hcd*tst");
		ws.getRange2("S2").setValue("2");
		ws.getRange2("T1").setValue("#DIV/0!");
		ws.getRange2("T2").setValue("*vmwtst*");
		ws.getRange2("U1").setValue("1");
		ws.getRange2("U2").setValue("*vmctst*");

		oParser = new parserFormula('SEARCH(S1:U2,A1:A2)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F5").bbox);
		assert.ok(oParser.parse(), 'SEARCH(S1:U2,A1:A2)');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of SEARCH(S1:U2,A1:A2)');

		oParser = new parserFormula('SEARCH(S1:U2,"hcdpmh71")', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F5").bbox);
		assert.ok(oParser.parse(), 'SEARCH(S1:U2,"hcdpmh71")');
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!", 'Result of SEARCH(S1:U2,"hcdpmh71")');

		oParser = new parserFormula('SEARCH("hcd",A1:A2)', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D1:F5").bbox);
		assert.ok(oParser.parse(), 'SEARCH("hcd",A1:A2)');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Result of SEARCH("hcd",A1:A2)[0,0]');

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
		// Case #1: String(2). Basic string input, finds "a" in "abc". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2). Basic string input, finds "a" in "abc". Returns 1. 2 of 3 arguments used.');
		// Case #2: String(2), Number. String input with valid start_num. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2), Number. String input with valid start_num. Returns 1. 3 of 3 arguments used.');
		// Case #3: Formula(2). Nested formula for find_text. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(LEFT("abc",1),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(LEFT("abc",1),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2). Nested formula for find_text. Returns 1. 2 of 3 arguments used.');
		// Case #4: String, Formula. Formula for within_text. Returns 2. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("b",CONCAT("a","bc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("b",CONCAT("a","bc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String, Formula. Formula for within_text. Returns 2. 2 of 3 arguments used.');
		// Case #5: Reference link(2). Reference link for both arguments. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link(2). Reference link for both arguments. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.');
		// Case #6: Area(2). Single-cell range for both arguments. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(A102:A102,A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(A102:A102,A103:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area(2). Single-cell range for both arguments. Returns 1. 2 of 3 arguments used.');
		// Case #7: Array(2). Array with single element. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH({"a"},{"abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH({"a"},{"abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array(2). Array with single element. Returns 1. 2 of 3 arguments used.');
		// Case #8: Name(2). Named ranges. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name(2). Named ranges. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.');
		// Case #9: Name3D(2). 3D named ranges. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D(2). 3D named ranges. Returns 1. 2 of 3 arguments used.');
		// Case #10: Ref3D(2). 3D references. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D(2). 3D references. Returns 1. 2 of 3 arguments used.');
		// Case #11: Area3D(2). 3D single-cell ranges. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Sheet2!A1:A1,Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Sheet2!A1:A1,Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D(2). 3D single-cell ranges. Returns 1. 2 of 3 arguments used.');
		// Case #12: Table(2). Table structured references. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Table1[Column1],Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Table1[Column1],Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table(2). Table structured references. Returns 1. 2 of 3 arguments used.');
		// Case #13: Number(2). Number inputs, finds 39 in 39539. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(39,39539)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(39,39539) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number(2). Number inputs, finds 39 in 39539. Returns 1. 2 of 3 arguments used.');
		// Case #14: Date(2). Date serial numbers. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(DATE(2025,1,1),DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(DATE(2025,1,1),DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date(2). Date serial numbers. Returns 1. 2 of 3 arguments used.');
		// Case #15: Time, String. Time as find_text, string as within_text. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(TIME(12,0,0),"12:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(TIME(12,0,0),"12:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Time, String. Time as find_text, string as within_text. Returns 1. 2 of 3 arguments used.');
		// Case #16: Formula. SEARCH inside SUM formula. Returns 2. 2 of 3 arguments used.
		oParser = new parserFormula('SUM(SEARCH("a","abc"),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SEARCH("a","abc"),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. SEARCH inside SUM formula. Returns 2. 2 of 3 arguments used.');
		// Case #17: String, Number. String and number with start_num. Returns 2. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("39",39539,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("39",39539,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: String, Number. String and number with start_num. Returns 2. 3 of 3 arguments used.');
		// Case #18: Array(2). Multi-element arrays. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH({"a","b"},{"abc","bcd"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH({"a","b"},{"abc","bcd"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array(2). Multi-element arrays. Returns 1. 2 of 3 arguments used.');
		// Case #19: Formula(2), Number. Nested IF for find_text. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH(IF(TRUE,"a","b"),"abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(IF(TRUE,"a","b"),"abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2), Number. Nested IF for find_text. Returns 1. 3 of 3 arguments used.');
		// Case #20: String(3). All string inputs with start_num. Returns 2. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("b","abc",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("b","abc",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String(3). All string inputs with start_num. Returns 2. 3 of 3 arguments used.');
		// Case #21: String(2). Unicode characters. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("?","???")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("?","???") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2). Unicode characters. Returns 1. 2 of 3 arguments used.');

		// Negative cases:
		// Case #1: Empty, String. Empty find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Empty, String. Empty find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #2: String, Empty. Empty within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Empty. Empty within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #3: String(2), Number. start_num <= 0 returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Number. start_num <= 0 returns #VALUE!. 3 of 3 arguments used.');
		// Case #4: String(2), Number. start_num > length of within_text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Number. start_num > length of within_text returns #VALUE!. 3 of 3 arguments used.');
		// Case #5: Error, String. Propagates #N/A error. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(NA(),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(NA(),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Propagates #N/A error. 2 of 3 arguments used.');
		// Case #6: String, Error. Propagates #N/A error. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Error. Propagates #N/A error. 2 of 3 arguments used.');
		// Case #7: Area, String. Multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(A102:A103,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(A102:A103,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area, String. Multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #8: String, Area. Multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Area. Multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #9: Boolean, String. Boolean find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(FALSE,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(FALSE,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean, String. Boolean find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #10: String, Boolean. Boolean within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: String, Boolean. Boolean within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #11: Ref3D, String. 3D ref to text ("abc") returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Sheet2!A3,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Sheet2!A3,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, String. 3D ref to text ("abc") returns #VALUE!. 2 of 3 arguments used.');
		// Case #12: Name, String. Named range with text ("invalid") returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(TestNameArea,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(TestNameArea,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, String. Named range with text ("invalid") returns #VALUE!. 2 of 3 arguments used.');
		// Case #13: Table, String. Table column with text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Table1[Column2],"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Table1[Column2],"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table, String. Table column with text returns #VALUE!. 2 of 3 arguments used.');
		// Case #14: Formula, String. Formula resulting in #NUM! returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(SQRT(-1),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(SQRT(-1),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, String. Formula resulting in #NUM! returns #NUM!. 2 of 3 arguments used.');
		// Case #15: Array, String. Array with boolean returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH({FALSE},"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH({FALSE},"abc") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Array, String. Array with boolean returns #VALUE!. 2 of 3 arguments used.');
		// Case #16: String(2), Formula. start_num as #NUM! returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String(2), Formula. start_num as #NUM! returns #NUM!. 3 of 3 arguments used.');
		// Case #17: String(2), Boolean. Boolean start_num returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: String(2), Boolean. Boolean start_num returns #VALUE!. 3 of 3 arguments used.');
		// Case #18: Area3D, String. 3D multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(Sheet2!A1:B1,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(Sheet2!A1:B1,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D, String. 3D multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #19: String, Area3D. 3D multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",Sheet2!A1:B1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",Sheet2!A1:B1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Area3D. 3D multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #20: String(2), Ref3D. 3D ref to text as start_num returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Ref3D. 3D ref to text as start_num returns #VALUE!. 3 of 3 arguments used.');
		// Case #21: Name, String, Number. Named range with text as find_text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH(TestNameArea,"abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(TestNameArea,"abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, String, Number. Named range with text as find_text returns #VALUE!. 3 of 3 arguments used.');

		// Bounded cases:
		// Case #1: String(2), Number. Max string length (32,767) for within_text. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a",REPT("a",32767),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a",REPT("a",32767),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String(2), Number. Max string length (32,767) for within_text. Returns 1. 3 of 3 arguments used.');
		// Case #2: String(2), Number. Max valid start_num (length of within_text). Returns 3. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String(2), Number. Max valid start_num (length of within_text). Returns 3. 3 of 3 arguments used.');
		// Case #3: String(2), Number. Min valid start_num (1). Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCH("a","abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH("a","abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String(2), Number. Min valid start_num (1). Returns 1. 3 of 3 arguments used.');
		// Case #4: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCH(REPT("a",32767),"a"&REPT("a",32766))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCH(REPT("a",32767),"a"&REPT("a",32766)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.');

		// Need to fix: diff errors
		// Case #15: Array, String. Array with boolean returns #VALUE!. 2 of 3 arguments used.
		// Case #4: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.

		// testArrayFormula2(assert, "SEARCH", 2, 3);
	});

	QUnit.test("Test: \"SEARCHB\"", function (assert) {
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
		// Case #1: String(2). Basic string input, finds "a" in "abc". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2). Basic string input, finds "a" in "abc". Returns 1. 2 of 3 arguments used.');
		// Case #2: String(2), Number. String input with valid start_num. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2), Number. String input with valid start_num. Returns 1. 3 of 3 arguments used.');
		// Case #3: Formula(2). Nested formula for find_text. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(LEFT("abc",1),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(LEFT("abc",1),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2). Nested formula for find_text. Returns 1. 2 of 3 arguments used.');
		// Case #4: String, Formula. Formula for within_text. Returns 2. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("b",CONCAT("a","bc"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("b",CONCAT("a","bc")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String, Formula. Formula for within_text. Returns 2. 2 of 3 arguments used.');
		// Case #5: Reference link(2). Reference link for both arguments. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(A100,A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(A100,A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link(2). Reference link for both arguments. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.');
		// Case #6: Area(2). Single-cell range for both arguments. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(A102:A102,A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(A102:A102,A103:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Area(2). Single-cell range for both arguments. Returns 1. 2 of 3 arguments used.');
		// Case #7: Array(2). Array with single element. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB({"a"},{"abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB({"a"},{"abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array(2). Array with single element. Returns 1. 2 of 3 arguments used.');
		// Case #8: Name(2). Named ranges. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(TestName,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(TestName,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name(2). Named ranges. Finds "39" in "39539". Returns 1. 2 of 3 arguments used.');
		// Case #9: Name3D(2). 3D named ranges. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(TestName3D,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(TestName3D,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Name3D(2). 3D named ranges. Returns 1. 2 of 3 arguments used.');
		// Case #10: Ref3D(2). 3D references. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Sheet2!A1,Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Sheet2!A1,Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Ref3D(2). 3D references. Returns 1. 2 of 3 arguments used.');
		// Case #11: Area3D(2). 3D single-cell ranges. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Sheet2!A1:A1,Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Sheet2!A1:A1,Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Area3D(2). 3D single-cell ranges. Returns 1. 2 of 3 arguments used.');
		// Case #12: Table(2). Table structured references. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Table1[Column1],Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Table1[Column1],Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table(2). Table structured references. Returns 1. 2 of 3 arguments used.');
		// Case #13: Number(2). Number inputs, finds 39 in 39539. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(39,39539)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(39,39539) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Number(2). Number inputs, finds 39 in 39539. Returns 1. 2 of 3 arguments used.');
		// Case #14: Date(2). Date serial numbers. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(DATE(2025,1,1),DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(DATE(2025,1,1),DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Date(2). Date serial numbers. Returns 1. 2 of 3 arguments used.');
		// Case #15: Time, String. Time as find_text, string as within_text. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(TIME(12,0,0),"12:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(TIME(12,0,0),"12:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Time, String. Time as find_text, string as within_text. Returns 1. 2 of 3 arguments used.');
		// Case #16: Formula. SEARCH inside SUM formula. Returns 2. 2 of 3 arguments used.
		oParser = new parserFormula('SUM(SEARCHB("a","abc"),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(SEARCHB("a","abc"),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Formula. SEARCH inside SUM formula. Returns 2. 2 of 3 arguments used.');
		// Case #17: String, Number. String and number with start_num. Returns 2. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("39",39539,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("39",39539,2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: String, Number. String and number with start_num. Returns 2. 3 of 3 arguments used.');
		// Case #18: Array(2). Multi-element arrays. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB({"a","b"},{"abc","bcd"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB({"a","b"},{"abc","bcd"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 1, 'Test: Positive case: Array(2). Multi-element arrays. Returns 1. 2 of 3 arguments used.');
		// Case #19: Formula(2), Number. Nested IF for find_text. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(IF(TRUE,"a","b"),"abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(IF(TRUE,"a","b"),"abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Formula(2), Number. Nested IF for find_text. Returns 1. 3 of 3 arguments used.');
		// Case #20: String(3). All string inputs with start_num. Returns 2. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("b","abc",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("b","abc",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: String(3). All string inputs with start_num. Returns 2. 3 of 3 arguments used.');
		// Case #21: String(2). Unicode characters. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("?","???")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("?","???") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: String(2). Unicode characters. Returns 1. 2 of 3 arguments used.');

		// Negative cases:
		// Case #1: Empty, String. Empty find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Empty, String. Empty find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #2: String, Empty. Empty within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Empty. Empty within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #3: String(2), Number. start_num <= 0 returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Number. start_num <= 0 returns #VALUE!. 3 of 3 arguments used.');
		// Case #4: String(2), Number. start_num > length of within_text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Number. start_num > length of within_text returns #VALUE!. 3 of 3 arguments used.');
		// Case #5: Error, String. Propagates #N/A error. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(NA(),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(NA(),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Propagates #N/A error. 2 of 3 arguments used.');
		// Case #6: String, Error. Propagates #N/A error. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String, Error. Propagates #N/A error. 2 of 3 arguments used.');
		// Case #7: Area, String. Multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(A102:A103,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(A102:A103,"abc") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Area, String. Multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #8: String, Area. Multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Area. Multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #9: Boolean, String. Boolean find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(FALSE,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(FALSE,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean, String. Boolean find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #10: String, Boolean. Boolean within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Negative case: String, Boolean. Boolean within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #11: Ref3D, String. 3D ref to text ("abc") returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Sheet2!A3,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Sheet2!A3,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, String. 3D ref to text ("abc") returns #VALUE!. 2 of 3 arguments used.');
		// Case #12: Name, String. Named range with text ("invalid") returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(TestNameArea,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(TestNameArea,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, String. Named range with text ("invalid") returns #VALUE!. 2 of 3 arguments used.');
		// Case #13: Table, String. Table column with text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Table1[Column2],"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Table1[Column2],"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table, String. Table column with text returns #VALUE!. 2 of 3 arguments used.');
		// Case #14: Formula, String. Formula resulting in #NUM! returns #NUM!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(SQRT(-1),"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(SQRT(-1),"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, String. Formula resulting in #NUM! returns #NUM!. 2 of 3 arguments used.');
		// Case #15: Array, String. Array with boolean returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB({FALSE},"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB({FALSE},"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: Array, String. Array with boolean returns #VALUE!. 2 of 3 arguments used.');
		// Case #16: String(2), Formula. start_num as #NUM! returns #NUM!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: String(2), Formula. start_num as #NUM! returns #NUM!. 3 of 3 arguments used.');
		// Case #17: String(2), Boolean. Boolean start_num returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: String(2), Boolean. Boolean start_num returns #VALUE!. 3 of 3 arguments used.');
		// Case #18: Area3D, String. 3D multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(Sheet2!A1:B1,"abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(Sheet2!A1:B1,"abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D, String. 3D multi-cell range for find_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #19: String, Area3D. 3D multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",Sheet2!A1:B1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",Sheet2!A1:B1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, Area3D. 3D multi-cell range for within_text returns #VALUE!. 2 of 3 arguments used.');
		// Case #20: String(2), Ref3D. 3D ref to text as start_num returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String(2), Ref3D. 3D ref to text as start_num returns #VALUE!. 3 of 3 arguments used.');
		// Case #21: Name, String, Number. Named range with text as find_text returns #VALUE!. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(TestNameArea,"abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(TestNameArea,"abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name, String, Number. Named range with text as find_text returns #VALUE!. 3 of 3 arguments used.');

		// Bounded cases:
		// Case #1: String(2), Number. Max string length (32,767) for within_text. Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a",REPT("a",32767),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a",REPT("a",32767),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String(2), Number. Max string length (32,767) for within_text. Returns 1. 3 of 3 arguments used.');
		// Case #2: String(2), Number. Max valid start_num (length of within_text). Returns 3. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String(2), Number. Max valid start_num (length of within_text). Returns 3. 3 of 3 arguments used.');
		// Case #3: String(2), Number. Min valid start_num (1). Returns 1. 3 of 3 arguments used.
		oParser = new parserFormula('SEARCHB("a","abc",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB("a","abc",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: String(2), Number. Min valid start_num (1). Returns 1. 3 of 3 arguments used.');
		// Case #4: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.
		oParser = new parserFormula('SEARCHB(REPT("a",32767),"a"&REPT("a",32766))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SEARCHB(REPT("a",32767),"a"&REPT("a",32766)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.');

		// Need to fix: area handle, result diff in boundary case
		// Case #6: Area(2). Single-cell range for both arguments. Returns 1. 2 of 3 arguments used.
		// Case #4: String(2). Max string length for both arguments. Returns 1. 2 of 3 arguments used.


	});

	QUnit.test("Test: \"SUBSTITUTE\"", function (assert) {

		oParser = new parserFormula("SUBSTITUTE(\"abcaAabca\",\"a\",\"xx\")", "A2", ws);//?
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xxbcxxAxxbcxx");

		oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\")", "B2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xxbcxxxxxxbcxx");

		oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"\",10)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "abcaaabca");

		oParser = new parserFormula("SUBSTITUTE(\"abcaaabca\",\"a\",\"xx\",3)", "C2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "abcaxxabca");

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

		// Positive cases:
		// Case #0: String. Basic string replacement, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello World","World","Earth")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello World","World","Earth") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello Earth', 'Test: Positive case: String. Basic string replacement, 3 arguments used.');
		// Case #1: String,Number. String replacement with instance_num, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello World","l", "L", 2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello World","l", "L", 2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HelLo World', 'Test: Positive case: String,Number. String replacement with instance_num, 4 arguments used.');
		// Case #2: Formula,String. Nested formula in text argument, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(CONCAT("He","llo"),"ll","rr")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(CONCAT("He","llo"),"ll","rr") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Herro', 'Test: Positive case: Formula,String. Nested formula in text argument, 3 arguments used.');
		// Case #3: Number,String. Number converted to string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(12345,"2","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(12345,"2","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1X345', 'Test: Positive case: Number,String. Number converted to string, 3 arguments used.');
		// Case #4: Reference link,String. Reference link to cell with valid string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(A100,"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(A100,"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link,String. Reference link to cell with valid string, 3 arguments used.');
		// Case #5: Area,String. Single-cell range, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(A101:A101,"l","L")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(A101:A101,"l","L") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.5', 'Test: Positive case: Area,String. Single-cell range, 3 arguments used.');
		// Case #6: Array,String. Array with single string element, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE({"Hello"},"l","L")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE({"Hello"},"l","L") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HeLLo', 'Test: Positive case: Array,String. Array with single string element, 3 arguments used.');
		// Case #7: Name,String. Named range with valid string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestName,"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestName,"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name,String. Named range with valid string, 3 arguments used.');
		// Case #8: Name3D,String. 3D named range with valid string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestName3D,"l","L")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestName3D,"l","L") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name3D,String. 3D named range with valid string, 3 arguments used.');
		// Case #9: Ref3D,String. 3D reference to cell with valid string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(Sheet2!A1,"e","3")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(Sheet2!A1,"e","3") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D,String. 3D reference to cell with valid string, 3 arguments used.');
		// Case #10: Area3D,String. 3D single-cell range, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(Sheet2!A2:A2,"l","L")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(Sheet2!A2:A2,"l","L") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Positive case: Area3D,String. 3D single-cell range, 3 arguments used.');
		// Case #11: Table,String. Table structured reference with valid string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(Table1[Column1],"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(Table1[Column1],"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table,String. Table structured reference with valid string, 3 arguments used.');
		// Case #12: Date,String. Date as serial number converted to string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(DATE(2025,1,1),"2","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(DATE(2025,1,1),"2","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658', 'Test: Positive case: Date,String. Date as serial number converted to string, 3 arguments used.');
		// Case #13: Time,String. Time as decimal converted to string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TIME(12,0,0),"0","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TIME(12,0,0),"0","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'X.5', 'Test: Positive case: Time,String. Time as decimal converted to string, 3 arguments used.');
		// Case #14: Formula,String,Number. Nested IF formula returning valid string, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(IF(TRUE,"Hello","Hi"),"l","L",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(IF(TRUE,"Hello","Hi"),"l","L",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HeLlo', 'Test: Positive case: Formula,String,Number. Nested IF formula returning valid string, 4 arguments used.');
		// Case #15: Formula,String. SUBSTITUTE inside formula with string concatenation, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(SUM(1,1)&" times","1","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(SUM(1,1)&" times","1","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2 times', 'Test: Positive case: Formula,String. SUBSTITUTE inside formula with string concatenation, 3 arguments used.');
		// Case #16: String,Number. Replace first instance only, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("aaa","a","b",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("aaa","a","b",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'baa', 'Test: Positive case: String,Number. Replace first instance only, 4 arguments used.');
		// Case #17: Array,String,Number. Array with multiple strings, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE({"Hello","World"},"o","0",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE({"Hello","World"},"o","0",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hell0', 'Test: Positive case: Array,String,Number. Array with multiple strings, 4 arguments used.');
		// Case #18: Area,String,Number. Multi-cell range (2 cells), 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(A101:A102,"l","L",1)', 'A2', ws);
		oParser.setArrayFormulaRef(ws.getRange2("H1:J2").bbox);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(A101:A102,"l","L",1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '1.5', 'Test: Positive case: Area,String,Number. Multi-cell range (2 cells), 4 arguments used.');
		// Case #19: Area3D,String,Number. 3D multi-cell range (2 cells), 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(Sheet2!A2:A3,"o","0",1)', 'A2', ws);
		oParser.setArrayFormulaRef(ws.getRange2("H1:J2").bbox);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(Sheet2!A2:A3,"o","0",1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '2', 'Test: Positive case: Area3D,String,Number. 3D multi-cell range (2 cells), 4 arguments used.');
		assert.strictEqual(oParser.calculate().getElementRowCol(1,0).getValue(), 'Text', 'Test: Positive case: Area3D,String,Number. 3D multi-cell range (2 cells), 4 arguments used.');
		// Case #20: String,Empty,String. Empty old_text treated as empty string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello World","","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello World","","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello World', 'Test: Positive case: String,Empty,String. Empty old_text treated as empty string, 3 arguments used.');
		// Case #21: Name,String,Number. Named range with instance_num, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestName1,"o","0",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestName1,"o","0",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Name,String,Number. Named range with instance_num, 4 arguments used.');
		// Case #22: Name3D,String,Number. 3D named range with instance_num, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestName3D,"w","W",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestName3D,"w","W",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name3D,String,Number. 3D named range with instance_num, 4 arguments used.');

		// Negative cases:
		// Case #2: Error,String. Error in text propagates #N/A, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(NA(),"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(NA(),"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error,String. Error in text propagates #N/A, 3 arguments used.');
		// Case #3: String,Empty,String. Empty old_text returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello",,"X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello",,"X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Negative case: String,Empty,String. Empty old_text returns #VALUE!, 3 arguments used.');
		// Case #4: String,String,Empty. Empty new_text returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","o",)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","o",) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hell', 'Test: Negative case: String,String,Empty. Empty new_text returns #VALUE!, 3 arguments used.');
		// Case #5: String,String,Number. instance_num = 0 returns #VALUE!, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","o","0",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","o","0",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number. instance_num = 0 returns #VALUE!, 4 arguments used.');
		// Case #6: String,String,Number. Negative instance_num returns #VALUE!, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","o","0",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","o","0",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number. Negative instance_num returns #VALUE!, 4 arguments used.');
		// Case #7: Boolean,String. Boolean text returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TRUE,"T","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TRUE,"T","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XRUE', 'Test: Negative case: Boolean,String. Boolean text returns #VALUE!, 3 arguments used.');
		// Case #10: Name,String. Named range with area returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestNameArea2,"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestNameArea2,"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name,String. Named range with area returns #VALUE!, 3 arguments used.');
		// Case #11: Name3D,String. 3D named range with area returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(TestNameArea3D2,"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(TestNameArea3D2,"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name3D,String. 3D named range with area returns #VALUE!, 3 arguments used.');
		// Case #12: Table,String. Table with invalid data returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(Table1[Column2],"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(Table1[Column2],"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1s', 'Test: Negative case: Table,String. Table with invalid data returns #VALUE!, 3 arguments used.');
		// Case #13: Formula,String. Formula resulting in #NUM! propagates error, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(SQRT(-1),"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(SQRT(-1),"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula,String. Formula resulting in #NUM! propagates error, 3 arguments used.');
		// Case #14: String,String,String. old_text not in text returns original string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","x","y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","x","y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Negative case: String,String,String. old_text not in text returns original string, 3 arguments used.');
		// Case #15: Array,String. Array with boolean returns #VALUE!, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE({TRUE},"T","X")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE({TRUE},"T","X") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'XRUE', 'Test: Negative case: Array,String. Array with boolean returns #VALUE!, 3 arguments used.');
		// Case #16: Number,String,Number. Non-integer instance_num returns #VALUE!, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE(123,"2","X",1.5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(123,"2","X",1.5) is parsed.');//?
		assert.strictEqual(oParser.calculate().getValue(), '1X3', 'Test: Negative case: Number,String,Number. Non-integer instance_num returns #VALUE!, 4 arguments used.');
		// Case #17: Reference link,String. Reference to empty cell returns empty string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE(A102,"o","0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE(A102,"o","0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Reference link,String. Reference to empty cell returns empty string, 3 arguments used.');
		// Case #19: String,String,Number. instance_num greater than occurrences returns original string, 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","o","0",10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","o","0",10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Negative case: String,String,Number. instance_num greater than occurrences returns original string, 4 arguments used.');
		// Case #20: String,String,String. Case-sensitive mismatch returns original string, 3 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","hello","hi")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","hello","hi") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Negative case: String,String,String. Case-sensitive mismatch returns original string, 3 arguments used.');

		let longStr = "";
		// Bounded cases:
		// Case #1: String,String,String. Maximum string length (~32,767 chars), 3 arguments used.
		longStr = "C";
		oParser = new parserFormula('SUBSTITUTE("A"&REPT("B",32766),"B","C")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("A"&REPT("B",32766),"B","C") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "A" + longStr.repeat(32766), 'Test: Bounded case: String,String,String. Maximum string length (~32,767 chars), 3 arguments used.');
		// Case #2: String,String,Number. Maximum integer instance_num (2^31-1), 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","l","L",2147483647)', 'A2', ws);
		// oParser.setArrayFormulaRef(ws.getRange2("H1:J2").bbox);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","l","L",2147483647) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String,String,Number. Maximum integer instance_num (2^31-1), 4 arguments used.');
		// Case #3: String,String,Number. Minimum instance_num (1), 4 arguments used.
		oParser = new parserFormula('SUBSTITUTE("Hello","l","L",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUBSTITUTE("Hello","l","L",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HeLlo', 'Test: Bounded case: String,String,Number. Minimum instance_num (1), 4 arguments used.');

		testArrayFormula2(assert, "SUBSTITUTE", 3, 4);
	});

	QUnit.test("Test: \"T\"", function (assert) {

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("str0.5");
		ws.getRange2("A101").setValue("str1.5");
		ws.getRange2("A104").setValue("str-1");
		// For area
		ws.getRange2("A102").setValue("str2.5");
		ws.getRange2("A103").setValue("s");
		ws.getRange2("A105").setValue("str1");
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

		// Positive cases:
		// Case #1: String. Basic string input returns the string.
		oParser = new parserFormula('T("Hello")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("Hello") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: String. Basic string input returns the string.');
		// Case #2: String. Numeric string returns the string as text.
		oParser = new parserFormula('T("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: String. Numeric string returns the string as text.');
		// Case #3: String. String with spaces returns the string.
		oParser = new parserFormula('T("Test String")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("Test String") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test String', 'Test: Positive case: String. String with spaces returns the string.');
		// Case #4: String. Reference to cell with string returns the string.
		oParser = new parserFormula('T(A100)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(A100)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str0.51', 'Test: Reference to cell with string returns the string.');
		// Case #5: String. Single-cell range with string returns the string.
		oParser = new parserFormula('T(A101:A101)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(A101:A101)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str1.51', 'Test: Single-cell range with string returns the string.');
		// Case #6: Array. Array with single string element returns the string.
		oParser = new parserFormula('T({"Hello"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T({"Hello"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Hello', 'Test: Positive case: Array. Array with single string element returns the string.');

		// Case #7: Named range with string returns the string.
		oParser = new parserFormula('T(TestName)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TestName)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Named range with string returns the string.');
		// Case #8: 3D named range with string returns the string
		oParser = new parserFormula('T(TestName3D)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TestName3D)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: 3D named range with string returns the string.');
		// Case #9: 3D reference to cell with string returns the string.
		oParser = new parserFormula('T(Sheet2!A1)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(Sheet2!A1)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: 3D reference to cell with string returns the string.');
		// Case #10: 3D single-cell range with string returns the string.
		oParser = new parserFormula('T(Sheet2!A2:A2)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(Sheet2!A2:A2)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: 3D single-cell range with string returns the string.');
		// Case #11: Table structured reference with string returns the string.
		oParser = new parserFormula('T(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(Table1[Column2]) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '-1s', 'Test: Table structured reference with string returns the string.');
		// Case #12: String. String with special characters returns the string.
		oParser = new parserFormula('T("Special@#$%")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("Special@#$%") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Special@#$%', 'Test: Positive case: String. String with special characters returns the string.');
		// Case #13: Formula. Nested CONCAT formula returning string returns the string.
		oParser = new parserFormula('T(CONCAT("He","llo"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(CONCAT("He","llo")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello', 'Test: Positive case: Formula. Nested CONCAT formula returning string returns the string.');
		// Case #14: Formula. Nested IF formula returning string returns the string.
		oParser = new parserFormula('T(IF(TRUE,"Yes","No"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(IF(TRUE,"Yes","No")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Yes', 'Test: Positive case: Formula. Nested IF formula returning string returns the string.');
		// Case #15: Formula. TEXT formula returning string returns the string.
		oParser = new parserFormula('T(TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: Formula. TEXT formula returning string returns the string.');
		// Case #16: Formula. T function inside SUM with string input returns string length + 1.
		oParser = new parserFormula('SUM(LEN(T("Count")),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(LEN(T("Count")),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Positive case: Formula. T function inside SUM with string input returns string length + 1.');
		// Case #17: String. Short date-like string returns the string.
		oParser = new parserFormula('T("12/12")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("12/12") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12/12', 'Test: Positive case: String. Short date-like string returns the string.');
		// Case #18: Array. Array with multiple string elements returns first string.
		oParser = new parserFormula('T({"Test","Data"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T({"Test","Data"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Test', 'Test: Positive case: Array. Array with multiple string elements returns first string.');
		// Case #20: String. Long string input returns the string.
		oParser = new parserFormula('T("Very Long String 1234567890")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("Very Long String 1234567890") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Very Long String 1234567890', 'Test: Positive case: String. Long string input returns the string.');

		// Negative cases:
		// Case #1: Date input (serial number) returns empty string ("").
		oParser = new parserFormula('T(123)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(123)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Number input returns empty string ("").');
		// Case #2: Date input (serial number) returns empty string ("").
		oParser = new parserFormula('T(1.5)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(1.5)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Float number input returns empty string ("").');
		// Case #3: Date input (serial number) returns empty string ("").
		oParser = new parserFormula('T(DATE(2025,1,1))&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(DATE(2025,1,1))&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Date input (serial number) returns empty string ("").');
		// Case #4: Time input (fraction) returns empty string
		oParser = new parserFormula('T(TIME(12,0,0))&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TIME(12,0,0))&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Time input (fraction) returns empty string ("").');
		// Case #5: Error. Error input propagates #N/A error.
		oParser = new parserFormula('T(TRUE)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TRUE)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Error. Boolean input returns empty string ("").');
		// Case #6: Error. Error input propagates #N/A error.
		oParser = new parserFormula('T(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error input propagates #N/A error.');
		// Case #13: Ref3D. 3D reference to cell with number returns empty string ("").
		oParser = new parserFormula('T(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(Sheet2!A3) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Ref3D. 3D reference to cell with number returns empty string ("").');
		// Case #16: Name3D. 3D named range with number returns empty string ("").
		oParser = new parserFormula('T(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Name3D. 3D named range with number returns empty string ("").');
		// Case #17: Table. Table column with number returns empty string ("").
		oParser = new parserFormula('T("1s")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("1s") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1s', 'Test: Negative case: Table. Table column with number returns empty string ("").');
		// Case #18: Formula. Formula resulting in #NUM! error propagates #NUM! error.
		oParser = new parserFormula('T({123})&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T({123})&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates #NUM! error.');
		// Case #19: Formula. Formula resulting in #NUM! error propagates #NUM! error.
		oParser = new parserFormula('T(IF(TRUE,123,456))&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(IF(TRUE,123,456))&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates #NUM! error.');
		// Case #20: Formula. Formula resulting in #NUM! error propagates #NUM! error.
		oParser = new parserFormula('T(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates #NUM! error.');

		// Bounded cases:
		let longStr = "";
		// Case #1: String. Empty string (minimum valid string) returns empty string ("").
		oParser = new parserFormula('T("")&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("")&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Bounded case: String. Empty string (minimum valid string) returns empty string ("").');
		// Case #2: String. Single-character string returns the string.
		oParser = new parserFormula('T("A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T("A") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Bounded case: String. Single-character string returns the string.');
		// Case #3: String. Maximum string length (255 characters) returns the string.
		oParser = new parserFormula('T(REPT("A",255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(REPT("A",255)) is parsed.');
		longStr = "A";
		assert.strictEqual(oParser.calculate().getValue(), longStr.repeat(255), 'Test: Bounded case: String. Maximum string length (255 characters) returns the string.');
		// Case #4: String. Reference to cell with single-character string returns the string.
		oParser = new parserFormula('T(A105)&"1"', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: T(A105)&"1" is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'str11', 'Test: Bounded case: String. Reference to cell with single-character string returns the string.');

		// Need to fix: 3D ref handle as result(double getValue)
		// Case #13: Ref3D. 3D reference to cell with number returns empty string ("").
		// Case #11: Table structured reference with string returns the string.
	});

	QUnit.test("Test: \"T(123)\"", function (assert) {
		oParser = new parserFormula("T(123)", "A1", ws);
		assert.ok(oParser.parse());
		assert.ok(!oParser.calculate().getValue(), "123");
	});

	QUnit.test("Test: \"T(\"HELLO\")\"", function (assert) {
		oParser = new parserFormula("T(\"HELLO\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "HELLO");
	});

	QUnit.test("Test: \"TEXT\"", function (assert) {
		var culturelciddefault = AscCommon.g_oDefaultCultureInfo.LCID;

		oParser = new parserFormula("TEXT(1234.567,\"$0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "$1234.57");

		oParser = new parserFormula("TEXT(0.125,\"0.0%\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12.5%");

		oParser = new parserFormula("TEXT(123123,\"hh:mmm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:Feb:00");

		oParser = new parserFormula("TEXT(123123,\"hh:mmmm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:February:00");

		oParser = new parserFormula("TEXT(123123,\"hh:mmmmm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:F:00");

		testArrayFormula2(assert, "TEXT", 2, 2);

		//____________________________________en_____________________________________________
		AscCommon.setCurrentCultureInfo(1033);
		oParser = new parserFormula("TEXT(123,\"yy-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"yy-MM-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"hh:MM:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"general\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.00");

		oParser = new parserFormula("TEXT(123123123,\"#,###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,123");

		oParser = new parserFormula("TEXT(123123123,\"#,##0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,123.00");

		//todo
		// oParser = new parserFormula( "TEXT(123123123,\"###0,.00\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123123.12" );
		//
		// oParser = new parserFormula( "TEXT(123123123,\"###0,.00,\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123.12" );

		oParser = new parserFormula("TEXT(123123123,\"###0.,0,0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123123.00");

		oParser = new parserFormula("TEXT(123123123,\"#,,,#,#,0.,,0,0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,123.00");

		oParser = new parserFormula("TEXT(123123123,\"0,0,, q,w,,,w,,0,,0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,1 q,w,w,23");

		oParser = new parserFormula("TEXT(123123123,\"00,,q,,,q00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1231231q,q23");

		oParser = new parserFormula("TEXT(123123123,\"0,q,w,,\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123q,w,");

		oParser = new parserFormula("TEXT(123123123,\"#,,\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123123123,\",,,#\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), ",123123123");

		oParser = new parserFormula("TEXT(123123123,\"###0,.\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123.");

		oParser = new parserFormula("TEXT(123123123,\"###0.,\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123.");

		oParser = new parserFormula("TEXT(123123123,\"###0.,q\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123.q");

		oParser = new parserFormula("TEXT(123123123,\"###0,.\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123.");

		oParser = new parserFormula("TEXT(123123123,\"###0 ,\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123123 ,");

		oParser = new parserFormula("TEXT(123,\"hh qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"dd hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa am/pm\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Wed-Wednesday AM");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//__________________________________fi________________________________________________
		AscCommon.setCurrentCultureInfo(1035);
		oParser = new parserFormula("TEXT(123,\"vv-kk-pp\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"vv-mm-pp\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"tt.mm.ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00.00.00");

		oParser = new parserFormula("TEXT(123,\"tt.MM.ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00.00.00");

		oParser = new parserFormula("TEXT(123,\"tt.kk.ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00.05.00");

		oParser = new parserFormula("TEXT(125,\"yleinen\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125");

		oParser = new parserFormula("TEXT(125,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125,00");

		oParser = new parserFormula("TEXT(125125,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125 125");

		oParser = new parserFormula("TEXT(123,\"pp tt\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"p t\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2 0");

		oParser = new parserFormula("TEXT(123,\"tt qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "ke-keskiviikko");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//__________________________________fy________________________________________________
		AscCommon.setCurrentCultureInfo(1043);
		oParser = new parserFormula("TEXT(123,\"jj-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"uu:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(126,\"standaard\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "126");

		oParser = new parserFormula("TEXT(126,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "126,00");

		oParser = new parserFormula("TEXT(126126,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "126.126");

		oParser = new parserFormula("TEXT(123,\"dd uu\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"d u\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2 0");

		oParser = new parserFormula("TEXT(123,\"uu qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "wo-woensdag");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//__________________________________es________________________________________________
		AscCommon.setCurrentCultureInfo(3082);
		oParser = new parserFormula("TEXT(123,\"aa-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"estándar\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.123.123");

		oParser = new parserFormula("TEXT(123,\"dd hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"hh qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"ooo-oooo\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "mi.-miércoles");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//___________________________________ru______________________________________________
		AscCommon.setCurrentCultureInfo(1049);
		oParser = new parserFormula("TEXT(123,\"гг-ММ-дд\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"чч:ММ:сс\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"чч:мм:сс\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"гг-мм-дд\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"основной\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123123123;\"# ##0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123,00");

		//todo
		// oParser = new parserFormula( "TEXT(123123123,\"###0 ,00\")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123123,12" );
		//
		// oParser = new parserFormula( "TEXT(123123123,\"###0 ,00 \")", "A2", ws );
		// assert.ok( oParser.parse() );
		// assert.strictEqual( oParser.calculate().getValue(), "123,12" );

		oParser = new parserFormula("TEXT(123123123,\"###0, 0 0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123123,00");

		oParser = new parserFormula("TEXT(123123123,\"#   # # 0,  0 0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123,00");

		oParser = new parserFormula("TEXT(123123123,\"0 0   q w   w  0  0\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 1q w w 23");

		oParser = new parserFormula("TEXT(123123123,\"00  q   q00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1231231q q23");

		oParser = new parserFormula("TEXT(123123123,\"0 q w  \")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123q w ");

		oParser = new parserFormula("TEXT(123123123,\"#  \")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123123123,\"   #\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), " 123123123");

		oParser = new parserFormula("TEXT(123123123,\"###0 ,\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123,");

		oParser = new parserFormula("TEXT(123123123,\"###0, \")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123123,");

		oParser = new parserFormula("TEXT(123,\"дд чч\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"чч qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"ММ мм\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "05 00");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Ср-среда");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//__________________________________fr________________________________________________
		AscCommon.setCurrentCultureInfo(1036);
		oParser = new parserFormula("TEXT(123,\"aa-mm-jj\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"standard\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123,\"jj hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"j h\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2 0");

		oParser = new parserFormula("TEXT(123,\"hh qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"ooo-oooo\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "mer.-mercredi");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//_______________________________de___________________________________________________
		AscCommon.setCurrentCultureInfo(1031);
		oParser = new parserFormula("TEXT(123,\"jj-MM-tt\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"hh:MM:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"jj-mm-tt\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"standard\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.123.123");

		oParser = new parserFormula("TEXT(123,\"t h\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "2 0");

		oParser = new parserFormula("TEXT(123,\"h qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "0 qq");

		oParser = new parserFormula("TEXT(123,\"s t\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "0 2");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Mi-Mittwoch");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//_______________________________it___________________________________________________
		AscCommon.setCurrentCultureInfo(1040);
		oParser = new parserFormula("TEXT(123,\"aa-MM-gg\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"aa-mm-gg\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"hh:MM:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"standard\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.123.123");

		oParser = new parserFormula("TEXT(123,\"gg hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"hh qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"ooo-oooo\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "mer-mercoledì");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//_______________________________da____________________________________________________
		AscCommon.setCurrentCultureInfo(1053);
		oParser = new parserFormula("TEXT(123,\"åå-MM-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"åå-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"tt:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"tt:MM:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"standard\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123,\"dd tt\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"tt qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "ons-onsdag");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________ch________________________________________________
		AscCommon.setCurrentCultureInfo(1028);
		oParser = new parserFormula("TEXT(123,\"g/通用格式\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.00");

		oParser = new parserFormula("TEXT(123,\"dd hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "週三-星期三");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________el________________________________________________
		AscCommon.setCurrentCultureInfo(1032);
		oParser = new parserFormula("TEXT(123,\"εε-μμ-ηη\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"ωω:λλ:δδ\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"γενικός τύπος\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.123.123");

		oParser = new parserFormula("TEXT(123,\"ηη ωω\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"ωω qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"ηη qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 qq");

		oParser = new parserFormula("TEXT(123,\"δδ ηη\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 02");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Τετ-Τετάρτη");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

		//_____________________________hu________________________________________________
		AscCommon.setCurrentCultureInfo(1038);
		oParser = new parserFormula("TEXT(123,\"éé-hh-nn\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"óó:pp:mm\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"éé-pp-nn\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"óó:hh:mm\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:05:00");

		oParser = new parserFormula("TEXT(125,\"normál\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125");

		oParser = new parserFormula("TEXT(125,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Sze-szerda");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________tr________________________________________________
		AscCommon.setCurrentCultureInfo(1055);
		oParser = new parserFormula("TEXT(123,\"yy-aa-gg\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"ss:dd:nn\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"yy-dd-gg\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-00-02");

		oParser = new parserFormula("TEXT(123,\"ss:aa:nn\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:05:00");

		oParser = new parserFormula("TEXT(125,\"genel\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125");

		oParser = new parserFormula("TEXT(125,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125,00");

		oParser = new parserFormula("TEXT(123123123,\"#.###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.123.123");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "May-Mayıs");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________pl________________________________________________
		AscCommon.setCurrentCultureInfo(1045);
		oParser = new parserFormula("TEXT(123,\"rr-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"gg:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(125,\"standardowy\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125");

		oParser = new parserFormula("TEXT(125,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "śr.-środa");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________cs________________________________________________
		AscCommon.setCurrentCultureInfo(1029);
		oParser = new parserFormula("TEXT(123,\"rr-mm-dd\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00-05-02");

		oParser = new parserFormula("TEXT(123,\"hh:mm:ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00:00:00");

		oParser = new parserFormula("TEXT(123,\"vęeobecný\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123");

		oParser = new parserFormula("TEXT(123,\"0,00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,00");

		oParser = new parserFormula("TEXT(123123123,\"# ###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123 123 123");

		oParser = new parserFormula("TEXT(123,\"rr ss\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 00");

		oParser = new parserFormula("TEXT(123,\"dd hh\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "02 00");

		oParser = new parserFormula("TEXT(123,\"hh qq\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "00 qq");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "st-středa");

		AscCommon.setCurrentCultureInfo(culturelciddefault);
		//_____________________________ja________________________________________________
		AscCommon.setCurrentCultureInfo(1041);
		oParser = new parserFormula("TEXT(124,\"G/標準\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "124");

		oParser = new parserFormula("TEXT(123,\"0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.00");

		oParser = new parserFormula("TEXT(123123123,\"#,###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,123");
		//_____________________________ko________________________________________________
		AscCommon.setCurrentCultureInfo(1042);
		oParser = new parserFormula("TEXT(124,\"g/표준\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "124");

		oParser = new parserFormula("TEXT(123,\"0.00\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123.00");

		oParser = new parserFormula("TEXT(123123123,\"#,###\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "123,123,123");

		oParser = new parserFormula("TEXT(123,\"aaa-aaaa\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "수-수요일");

		AscCommon.setCurrentCultureInfo(culturelciddefault);

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


		// Positive cases:
		// Case #1: Number, String. Basic valid input: number and format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.45', 'Test: Positive case: Number, String. Basic valid input: number and format string. 2 of 2 arguments used.');
		// Case #2: Date, String. Date as serial number with date format. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(44197,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(44197,"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '01/01/2021', 'Test: Positive case: Date, String. Date as serial number with date format. 2 of 2 arguments used.');
		// Case #3: Time, String. Time as fraction with time format. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(0.5,"h:mm AM/PM")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(0.5,"h:mm AM/PM") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12:00 PM', 'Test: Positive case: Time, String. Time as fraction with time format. 2 of 2 arguments used.');
		// Case #4: String, String. String convertible to number with format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT("123.45","0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT("123.45","0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.45', 'Test: Positive case: String, String. String convertible to number with format string. 2 of 2 arguments used.');
		// Case #5: Formula, String. Nested formula producing number with format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(SQRT(16),"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(SQRT(16),"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '4.00', 'Test: Positive case: Formula, String. Nested formula producing number with format string. 2 of 2 arguments used.');
		// Case #6: Number, Formula. Number with formula producing format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,CONCAT("0.","REPT(""0",2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,CONCAT("0.","REPT(""0",2)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Number, Formula. Number with formula producing format string. 2 of 2 arguments used.');
		// Case #7: Reference link, String. Reference to cell with valid number. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(A100,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(A100,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.50', 'Test: Positive case: Reference link, String. Reference to cell with valid number. 2 of 2 arguments used.');
		// Case #8: Area, String. Single-cell range with date format. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(A101:A101,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(A101:A101,"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '01/01/1900', 'Test: Positive case: Area, String. Single-cell range with date format. 2 of 2 arguments used.');
		// Case #9: Array, String. Array with single number element. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT({123.45},"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT({123.45},"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.45', 'Test: Positive case: Array, String. Array with single number element. 2 of 2 arguments used.');
		// Case #10: Name, String. Named range with valid number. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TestName,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TestName,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.50', 'Test: Positive case: Name, String. Named range with valid number. 2 of 2 arguments used.');
		// Case #11: Name3D, String. 3D named range with valid number. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TestName3D,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TestName3D,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.50', 'Test: Positive case: Name3D, String. 3D named range with valid number. 2 of 2 arguments used.');
		// Case #12: Ref3D, String. 3D reference to cell with valid date. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Sheet2!A1,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Sheet2!A1,"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '01/01/1900', 'Test: Positive case: Ref3D, String. 3D reference to cell with valid date. 2 of 2 arguments used.');
		// Case #13: Area3D, String. 3D single-cell range with number. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Sheet2!A2:A2,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Sheet2!A2:A2,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2.00', 'Test: Positive case: Area3D, String. 3D single-cell range with number. 2 of 2 arguments used.');
		// Case #14: Table, String. Table structured reference with number. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Table1[Column1],"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Table1[Column1],"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.00', 'Test: Positive case: Table, String. Table structured reference with number. 2 of 2 arguments used.');
		// Case #15: Number, Reference link. Number with reference to format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.5', 'Test: Positive case: Number, Reference link. Number with reference to format string. 2 of 2 arguments used.');
		// Case #17: Formula, Formula. Both arguments as formulas producing valid inputs. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(IF(TRUE,123.45,0),CONCAT("dd/mm/","yyyy"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(IF(TRUE,123.45,0),CONCAT("dd/mm/","yyyy")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '02/05/1900', 'Test: Positive case: Formula, Formula. Both arguments as formulas producing valid inputs. 2 of 2 arguments used.');
		// Case #18: Date, String. Date with custom format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(44197,"mmmm yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(44197,"mmmm yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'January 2021', 'Test: Positive case: Date, String. Date with custom format string. 2 of 2 arguments used.');
		// Case #19: Array, String. Multi-element array for value with date format. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT({44197,123.45},"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT({44197,123.45},"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '01/01/2021', 'Test: Positive case: Array, String. Multi-element array for value with date format. 2 of 2 arguments used.');
		// Case #20: Time, String. Time formula with time format. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TIME(12,0,0),"h:mm:ss")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TIME(12,0,0),"h:mm:ss") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12:00:00', 'Test: Positive case: Time, String. Time formula with time format. 2 of 2 arguments used.');
		// Case #21: Number, Name. Number with named range format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123.5', 'Test: Positive case: Number, Name. Number with named range format string. 2 of 2 arguments used.');
		// Case #22: Date, Name3D. Date with 3D named range format string. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(44197,TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(44197,TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-44197.5', 'Test: Positive case: Date, Name3D. Date with 3D named range format string. 2 of 2 arguments used.');

		// Negative cases:
		// Case #1: Empty, String. Empty value returns #VALUE!. 1 of 2 arguments used.
		oParser = new parserFormula('TEXT(,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.00', 'Test: Negative case: Empty, String. Empty value returns #VALUE!. 1 of 2 arguments used.');
		// Case #3: String, String. Non-numeric string value returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT("abc","0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT("abc","0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Negative case: String, String. Non-numeric string value returns #VALUE!. 2 of 2 arguments used.');
		// Case #4: Boolean, String. Boolean value returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TRUE,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TRUE,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Boolean, String. Boolean value returns #VALUE!. 2 of 2 arguments used.');
		// Case #5: Error, String. Error value propagates #N/A. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(NA(),"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(NA(),"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String. Error value propagates #N/A. 2 of 2 arguments used.');
		// Case #6: Area, String. Multi-cell range for value returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(A100:A101,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(A100:A101,"0.00") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Negative case: Area, String. Multi-cell range for value returns #VALUE!. 2 of 2 arguments used.');
		// Case #7: Number, Number. Non-string format_text returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Negative case: Number, Number. Non-string format_text returns #VALUE!. 2 of 2 arguments used.');
		// Case #8: Date, Boolean. Boolean format_text returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(44197,TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(44197,TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date, Boolean. Boolean format_text returns #VALUE!. 2 of 2 arguments used.');
		// Case #9: Date, String. Negative date returns #NUM!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(-44197,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(-44197,"dd/mm/yyyy") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date, String. Negative date returns #NUM!. 2 of 2 arguments used.');
		// Case #10: Reference link, String. Reference to empty cell returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(A104,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(A104,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1.00', 'Test: Negative case: Reference link, String. Reference to empty cell returns #VALUE!. 2 of 2 arguments used.');
		// Case #11: Name, String. Named range with non-numeric data returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TestNameArea2,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TestNameArea2,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.80', 'Test: Negative case: Name, String. Named range with non-numeric data returns #VALUE!. 2 of 2 arguments used.');
		// Case #12: Name3D, String. 3D named range with non-numeric data returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(TestNameArea3D2,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(TestNameArea3D2,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.80', 'Test: Negative case: Name3D, String. 3D named range with non-numeric data returns #VALUE!. 2 of 2 arguments used.');
		// Case #13: Ref3D, String. 3D reference to non-numeric data returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Sheet2!A3,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Sheet2!A3,"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Ref3D, String. 3D reference to non-numeric data returns #VALUE!. 2 of 2 arguments used.');
		// Case #14: Area3D, String. 3D multi-cell range for value returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Sheet2!A4:A5,"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Sheet2!A4:A5,"0.00") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0', 'Test: Negative case: Area3D, String. 3D multi-cell range for value returns #VALUE!. 2 of 2 arguments used.');
		// Case #15: Table, String. Table column with non-numeric data returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(Table1[Column2],"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(Table1[Column2],"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1s', 'Test: Negative case: Table, String. Table column with non-numeric data returns #VALUE!. 2 of 2 arguments used.');
		// Case #16: Array, String. Array with boolean returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT({TRUE},"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT({TRUE},"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array, String. Array with boolean returns #VALUE!. 2 of 2 arguments used.');
		// Case #17: Number, String. Invalid format string returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,"invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,"invalid") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number, String. Invalid format string returns #VALUE!. 2 of 2 arguments used.');
		// Case #18: Formula, String. Formula resulting in #NUM! propagates error. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(SQRT(-1),"0.00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(SQRT(-1),"0.00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, String. Formula resulting in #NUM! propagates error. 2 of 2 arguments used.');
		// Case #20: Number, Area. Multi-cell range for format_text returns #VALUE!. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(123.45,A105:A106)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(123.45,A105:A106) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Number, Area. Multi-cell range for format_text returns #VALUE!. 2 of 2 arguments used.');

		// Bounded cases:
		// Case #1: Number, String. Minimum valid number for value. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(0,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(0,"0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0', 'Test: Bounded case: Number, String. Minimum valid number for value. 2 of 2 arguments used.');
		// Case #2: Number, String. Maximum valid Excel number for value. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(9.99999999999999E+307,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(9.99999999999999E+307,"0") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Number, String. Maximum valid Excel number for value. 2 of 2 arguments used.');
		// Case #3: Date, String. Minimum valid date (01/01/1900). 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(1,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(1,"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '01/01/1900', 'Test: Bounded case: Date, String. Minimum valid date (01/01/1900). 2 of 2 arguments used.');
		// Case #4: Date, String. Maximum valid date (12/31/9999). 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(2958465,"dd/mm/yyyy")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(2958465,"dd/mm/yyyy") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '31/12/9999', 'Test: Bounded case: Date, String. Maximum valid date (12/31/9999). 2 of 2 arguments used.');
		// Case #5: Time, String. Minimum valid time (00:00:00). 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(0,"h:mm:ss")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(0,"h:mm:ss") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0:00:00', 'Test: Bounded case: Time, String. Minimum valid time (00:00:00). 2 of 2 arguments used.');
		// Case #6: Number, String. Minimum valid negative Excel number for value. 2 of 2 arguments used.
		oParser = new parserFormula('TEXT(-9.99999999999999E+307,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXT(-9.99999999999999E+307,"0") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: Number, String. Minimum valid negative Excel number for value. 2 of 2 arguments used.');

		// Need to fix: many different results from MS, check round problem and correct text handle
		// Case #6: Number, Formula. Number with formula producing format string. 2 of 2 arguments used.
		// Case #6: Area, String. Multi-cell range for value returns #VALUE!. 2 of 2 arguments used.
		// Case #8: Date, Boolean. Boolean format_text returns #VALUE!. 2 of 2 arguments used.
		// Case #9: Date, String. Negative date returns #NUM!. 2 of 2 arguments used.
		// Case #14: Area3D, String. 3D multi-cell range for value returns #VALUE!. 2 of 2 arguments used.
		// Case #17: Number, String. Invalid format string returns #VALUE!. 2 of 2 arguments used.
		// Case #20: Number, Area. Multi-cell range for format_text returns #VALUE!. 2 of 2 arguments used.
		// Case #2: Number, String. Maximum valid Excel number for value. 2 of 2 arguments used.
		// Case #6: Number, String. Minimum valid negative Excel number for value. 2 of 2 arguments used.

	});

	QUnit.test("Test: \"TEXTJOIN\"", function (assert) {

		ws.getRange2("A2").setValue("Tulsa");
		ws.getRange2("A3").setValue("Seattle");
		ws.getRange2("A4").setValue("Iselin");
		ws.getRange2("A5").setValue("Fort Lauderdale");
		ws.getRange2("A6").setValue("Tempe");
		ws.getRange2("A7").setValue("end");

		ws.getRange2("B2").setValue("OK");
		ws.getRange2("B3").setValue("WA");
		ws.getRange2("B4").setValue("NJ");
		ws.getRange2("B5").setValue("FL");
		ws.getRange2("B6").setValue("AZ");
		ws.getRange2("B7").setValue("");

		ws.getRange2("C2").setValue("74133");
		ws.getRange2("C3").setValue("98109");
		ws.getRange2("C4").setValue("8830");
		ws.getRange2("C5").setValue("33309");
		ws.getRange2("C6").setValue("85285");
		ws.getRange2("C7").setValue("");

		ws.getRange2("D2").setValue("US");
		ws.getRange2("D3").setValue("US");
		ws.getRange2("D4").setValue("US");
		ws.getRange2("D5").setValue("US");
		ws.getRange2("D6").setValue("US");
		ws.getRange2("D7").setValue("");

		ws.getRange2("A9").setValue(",");
		ws.getRange2("B9").setValue(",");
		ws.getRange2("C9").setValue(",");
		ws.getRange2("D9").setValue(";");


		oParser = new parserFormula("TEXTJOIN(A9:D9, TRUE, A2:D7)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Tulsa,OK,74133,US;Seattle,WA,98109,US;Iselin,NJ,8830,US;Fort Lauderdale,FL,33309,US;Tempe,AZ,85285,US;end");

		oParser = new parserFormula("TEXTJOIN(A9:D9, FALSE, A2:D7)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Tulsa,OK,74133,US;Seattle,WA,98109,US;Iselin,NJ,8830,US;Fort Lauderdale,FL,33309,US;Tempe,AZ,85285,US;end,,,");

		oParser = new parserFormula("TEXTJOIN(A2:D5, 1, B6:D6)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "AZTulsa85285OKUS");

		testArrayFormulaEqualsValues(assert,
			"113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,#N/A;113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,#N/A;#N/A,#N/A,#N/A,#N/A",
			"TEXTJOIN(A1:C2,A1:C2,A1:C2,A1:C2, A1:C2)");
		testArrayFormulaEqualsValues(assert,
			"113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,#N/A;113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,113.1232-41224152113.1232-4122415,#N/A;#N/A,#N/A,#N/A,#N/A",
			"TEXTJOIN(A1:A2,A1:C2,A1:C2,A1:C2)");
		testArrayFormulaEqualsValues(assert,
			"113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445;113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445,113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-4224455113.1233.123-4-422445;#N/A,#N/A,#N/A,#N/A",
			"TEXTJOIN(A1:C2,A1:A2,A1:C2,A1:C2,A1:C2,A1:C2)");


		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A106").setValue("2");
		ws.getRange2("A107").setValue("3");
		ws.getRange2("A109").setValue("4");
		ws.getRange2("A109").setValue("5");
		ws.getRange2("A110").setValue("6");
		ws.getRange2("A111").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:D10").cleanAll();
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
		// Case #1: String, Boolean, String. Basic string concatenation with comma delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,"Hello","World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,"Hello","World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Hello,World', 'Test: Positive case: String, Boolean, String. Basic string concatenation with comma delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #2: Number, Boolean, Number. Numbers converted to text, semicolon delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(";",TRUE,123,456)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(";",TRUE,123,456) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123;456', 'Test: Positive case: Number, Boolean, Number. Numbers converted to text, semicolon delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #3: Formula, Boolean, String. Nested formula (UPPER), space delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(" ",TRUE,UPPER("hello"),"World")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(" ",TRUE,UPPER("hello"),"World") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HELLO World', 'Test: Positive case: Formula, Boolean, String. Nested formula (UPPER), space delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #4: Reference link, Boolean, String. Reference link to cell with string, dash delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN("-",TRUE,A100,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN("-",TRUE,A100,"Test") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Positive case: Reference link, Boolean, String. Reference link to cell with string, dash delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #5: Area, Boolean, String. Single-cell range, comma delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,A101:A101,"Data")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,A101:A101,"Data") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Data', 'Test: Positive case: Area, Boolean, String. Single-cell range, comma delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #6: Array, Boolean, String. Array with multiple strings, colon delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(":",TRUE,{"A","B"},"C")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(":",TRUE,{"A","B"},"C") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A:B:C', 'Test: Positive case: Array, Boolean, String. Array with multiple strings, colon delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #7: Name, Boolean, String. Named range with string, comma delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,TestName,"Name")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,TestName,"Name") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5,Name', 'Test: Positive case: Name, Boolean, String. Named range with string, comma delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #8: Name3D, Boolean, String. 3D named range, semicolon delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(";",TRUE,TestName3D2,"3D")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(";",TRUE,TestName3D2,"3D") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.8;3D', 'Test: Positive case: Name3D, Boolean, String. 3D named range, semicolon delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #9: Ref3D, Boolean, String. 3D reference to cell with string, space delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(" ",TRUE,Sheet2!A1,"Ref")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(" ",TRUE,Sheet2!A1,"Ref") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1 Ref', 'Test: Positive case: Ref3D, Boolean, String. 3D reference to cell with string, space delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #10: Area3D, Boolean, String. 3D single-cell range, dash delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN("-",TRUE,Sheet2!A2:A2,"Area")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN("-",TRUE,Sheet2!A2:A2,"Area") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2-Area', 'Test: Positive case: Area3D, Boolean, String. 3D single-cell range, dash delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #11: Table, Boolean, String. Table structured reference with string, comma delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,Table1[Column1],"Table")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,Table1[Column1],"Table") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1,Table', 'Test: Positive case: Table, Boolean, String. Table structured reference with string, comma delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #12: Date, Boolean, String. Date as serial number converted to text, semicolon delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(";",TRUE,DATE(2025,1,1),"Date")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(";",TRUE,DATE(2025,1,1),"Date") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658;Date', 'Test: Positive case: Date, Boolean, String. Date as serial number converted to text, semicolon delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #13: Time, Boolean, String. Time converted to text, space delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(" ",TRUE,TIME(12,0,0),"Time")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(" ",TRUE,TIME(12,0,0),"Time") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5 Time', 'Test: Positive case: Time, Boolean, String. Time converted to text, space delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #14: Formula, Boolean, String. TEXTJOIN inside SUM formula, comma delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('SUM(LEN(TEXTJOIN(",",TRUE,"A","B")),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(LEN(TEXTJOIN(",",TRUE,"A","B")),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula, Boolean, String. TEXTJOIN inside SUM formula, comma delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #15: String, Boolean, Empty. Empty string included, comma delimiter, ignore_empty TRUE skips empty. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,"Text","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,"Text","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Positive case: String, Boolean, Empty. Empty string included, comma delimiter, ignore_empty TRUE skips empty. 3 arguments used.');
		// Case #16: String, Boolean, String. Empty string included, comma delimiter, ignore_empty FALSE includes empty. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",FALSE,"Text","")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",FALSE,"Text","") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Text,', 'Test: Positive case: String, Boolean, String. Empty string included, comma delimiter, ignore_empty FALSE includes empty. 3 arguments used.');
		// Case #17: Array, Boolean, String. Array with empty element, semicolon delimiter, ignore_empty TRUE skips empty. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(";",TRUE,{"","A"},"B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(";",TRUE,{"","A"},"B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A;B', 'Test: Positive case: Array, Boolean, String. Array with empty element, semicolon delimiter, ignore_empty TRUE skips empty. 3 arguments used.');
		// Case #18: Formula, Boolean, String. Nested IF formula returning string, space delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(" ",TRUE,IF(TRUE,"Yes","No"),"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(" ",TRUE,IF(TRUE,"Yes","No"),"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Yes Test', 'Test: Positive case: Formula, Boolean, String. Nested IF formula returning string, space delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #19: Number, Boolean, Array. Number and array, dash delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN("-",TRUE,1000,{"A","B"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN("-",TRUE,1000,{"A","B"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1000-A-B', 'Test: Positive case: Number, Boolean, Array. Number and array, dash delimiter, ignore_empty TRUE. 3 arguments used.');
		// Case #20: String, Boolean, Reference link. String and reference link, colon delimiter, ignore_empty TRUE. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(":",TRUE,"Prefix",A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(":",TRUE,"Prefix",A102) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'Prefix', 'Test: Positive case: String, Boolean, Reference link. String and reference link, colon delimiter, ignore_empty TRUE. 3 arguments used.');

		// Negative cases:
		// Case #1: Error, Boolean, String. Error input (NA) propagates #N/A error. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,NA(),"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,NA(),"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, Boolean, String. Error input (NA) propagates #N/A error. 3 arguments used.');
		// Case #2: Empty, Boolean, String. Empty delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(,,TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(,,TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUETest', 'Test: Negative case: Empty, Boolean, String. Empty delimiter returns #VALUE!. 3 arguments used.');
		// Case #3: Boolean, Boolean, String. Boolean delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(TRUE,TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(TRUE,TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Boolean, Boolean, String. Boolean delimiter returns #VALUE!. 3 arguments used.');
		// Case #4: Formula, Boolean, String. Formula resulting in #NUM! for delimiter returns #NUM!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(SQRT(-1),TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(SQRT(-1),TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, Boolean, String. Formula resulting in #NUM! for delimiter returns #NUM!. 3 arguments used.');
		// Case #5: String, Empty, String. Empty ignore_empty parameter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: String, Empty, String. Empty ignore_empty parameter returns #VALUE!. 3 arguments used.');
		// Case #6: String, String, String. Non-boolean string for ignore_empty returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",","Invalid","Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",","Invalid","Test") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String. Non-boolean string for ignore_empty returns #VALUE!. 3 arguments used.');
		// Case #8: Area3D, Boolean, String. 3D multi-cell range returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,Sheet2!A3:A4)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,Sheet2!A3:A4) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Area3D, Boolean, String. 3D multi-cell range returns #VALUE!. 3 arguments used.');
		// Case #9: Name, Boolean, String. Named range with area returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,TestNameArea2,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,TestNameArea2,"Test") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '1,2,Test', 'Test: Negative case: Name, Boolean, String. Named range with area returns #VALUE!. 3 arguments used.');
		// Case #10: Name3D, Boolean, String. 3D named range with area returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,TestNameArea3D2,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,TestNameArea3D2,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0.8,-0.8,Test", 'Test: Negative case: Name3D, Boolean, String. 3D named range with area returns #VALUE!. 3 arguments used.');
		// Case #11: Ref3D, Boolean, String. 3D reference to cell with non-string value returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,Sheet2!A5,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,Sheet2!A5,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Ref3D, Boolean, String. 3D reference to cell with non-string value returns #VALUE!. 3 arguments used.');
		// Case #12: Table, Boolean, String. Table column with non-string data returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,Table1[Column2],"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,Table1[Column2],"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1s,Test', 'Test: Negative case: Table, Boolean, String. Table column with non-string data returns #VALUE!. 3 arguments used.');
		// Case #13: Array, Boolean, String. Array with boolean returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,{TRUE},"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,{TRUE},"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE,Test', 'Test: Negative case: Array, Boolean, String. Array with boolean returns #VALUE!. 3 arguments used.');
		// Case #14: Number, Boolean, String. Numeric delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(123,TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(123,TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Number, Boolean, String. Numeric delimiter returns #VALUE!. 3 arguments used.');
		// Case #15: Date, Boolean, String. Date as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(DATE(2025,1,1),TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(DATE(2025,1,1),TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Date, Boolean, String. Date as delimiter returns #VALUE!. 3 arguments used.');
		// Case #16: Time, Boolean, String. Time as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(TIME(12,0,0),TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(TIME(12,0,0),TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Time, Boolean, String. Time as delimiter returns #VALUE!. 3 arguments used.');
		// Case #17: Formula, Boolean, String. Formula returning number as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(IF(FALSE,"",1),TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(IF(FALSE,"",1),TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Formula, Boolean, String. Formula returning number as delimiter returns #VALUE!. 3 arguments used.');
		// Case #18: Reference link, Boolean, String. Reference link to cell with number as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(A105,TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(A105,TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Reference link, Boolean, String. Reference link to cell with number as delimiter returns #VALUE!. 3 arguments used.');
		// Case #19: Area, Boolean, String. Single-cell range with number as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(A106:A106,TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(A106:A106,TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Area, Boolean, String. Single-cell range with number as delimiter returns #VALUE!. 3 arguments used.');
		// Case #20: Array, Boolean, String. Array as delimiter returns #VALUE!. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN({"-"},TRUE,"Test")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN({"-"},TRUE,"Test") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Test', 'Test: Negative case: Array, Boolean, String. Array as delimiter returns #VALUE!. 3 arguments used.');

		// Bounded cases:
		let longStr = "";
		// Case #1: String, Boolean, String. Maximum string length (32,767 characters) for text argument. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,REPT("A",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,REPT("A",32767)) is parsed.');
		longStr = "A";
		assert.strictEqual(oParser.calculate().getValue(), longStr.repeat(32767), 'Test: Bounded case: String, Boolean, String. Maximum string length (32,767 characters) for text argument. 3 arguments used.');
		// Case #2: String, Boolean, Array. Maximum number of text arguments (252) in array. 3 arguments used.
		longStr = "{" + Array.from({length:252}, (v, index) => '"A' + (index + 1) + '"').join(",") + "}";
		oParser = new parserFormula('TEXTJOIN(",",TRUE,'+longStr+')', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,' + longStr + ') is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), longStr.replace(/[{}"]/g,""), 'Test: Bounded case: String, Boolean, Array. Maximum number of text arguments (252) in array. 3 arguments used.');
		// Case #3: Number, Boolean, Number. Maximum numeric value converted to text. 3 arguments used.
		oParser = new parserFormula('TEXTJOIN(",",TRUE,1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTJOIN(",",TRUE,1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1e+307', 'Test: Bounded case: Number, Boolean, Number. Maximum numeric value converted to text. 3 arguments used.');

		// TODO critical Case #6: String, String, String. Non-boolean string for ignore_empty returns #VALUE!. 3 arguments used
		// Need to fix: ms result difference, critical problem in case 6
		// Case #4: Reference link, Boolean, String. Reference link to cell with string, dash delimiter, ignore_empty TRUE. 3 arguments used.
		// Case #5: Area, Boolean, String. Single-cell range, comma delimiter, ignore_empty TRUE. 3 arguments used.
		// Case #8: Name3D, Boolean, String. 3D named range, semicolon delimiter, ignore_empty TRUE. 3 arguments used.
		// Case #20: String, Boolean, Reference link. String and reference link, colon delimiter, ignore_empty TRUE. 3 arguments used.
		// Case #6: String, String, String. Non-boolean string for ignore_empty returns #VALUE!. 3 arguments used. - critical
		// Case #9: Name, Boolean, String. Named range with area returns #VALUE!. 3 arguments used.


	});

	QUnit.test("Test: \"TEXTBEFORE\"", function (assert) {
		ws.getRange2("B2").setValue("TesttestTeesttestTESTttetstetest");
		ws.getRange2("B3").setValue("test");
		ws.getRange2("B4").setValue("test2");

		oParser = new parserFormula("TEXTBEFORE(B2;B3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Test");

		oParser = new parserFormula("TEXTBEFORE(B2;B4)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Test");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttestTESTttetste");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;4)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");


		oParser = new parserFormula("TEXTBEFORE(B2;B3;1;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;2;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Test");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;3;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;4;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;5;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttestTESTttetste");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;6, TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;0, TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");


		oParser = new parserFormula("TEXTBEFORE(B2;B3;-1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttestTESTttetste");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Test");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-4)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-1;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttestTESTttetste");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-2;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeesttest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-3;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TesttestTeest");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-4;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "Test");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-5;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");


		ws.getRange2("B2").setValue("12test434TESTtest233");
		ws.getRange2("B3").setValue("TEST");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;1;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;2;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12test434");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;3;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12test434TEST");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;4;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12test434TESTtest233");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;5;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;5;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");


		oParser = new parserFormula("TEXTBEFORE(B2;B3;-1;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12test434TEST");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-2;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12test434");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-3;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-4;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTBEFORE(B2;B3;-5;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");


		oParser = new parserFormula("TEXTBEFORE(\"12tessdadsadtestt434TESTtest233sd\";\"TEST\";2;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12tessdadsadtestt434");

		oParser = new parserFormula("TEXTBEFORE(\"12tessdadsadtestt434TESTtest233sd\";\"TEST\";;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12tessdadsad");

		oParser = new parserFormula("TEXTBEFORE(\"12tessdadsadtestt434TESTtest233sd\";\"TEST\";;;;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12tessdadsadtestt434");

		oParser = new parserFormula("TEXTBEFORE(\"12tessdadsadtestt434TESTtest233sd\";;;;;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTBEFORE(;;;;;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTBEFORE(\"\",\",\",\"test\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTBEFORE(\"\",\",\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTBEFORE(B1,\",\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		ws.getRange2("A1").setValue("TRUE1Brown,Jim");

		oParser = new parserFormula("TEXTBEFORE(A1,{\"i\",\"n\",\"b\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE1Brow");

		oParser = new parserFormula("TEXTBEFORE(A1,{\"i\",\"n\",TRUE})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTBEFORE(A1,{\"i\",\"n\",#VALUE!})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTBEFORE(A1,{\"i\",\"J\",\"o\",\"O\",\"m\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TRUE1Br");


		ws.getRange2("A1").setValue("test125test2test3test4FALSEtest5");
		ws.getRange2("B4").setValue("w");
		ws.getRange2("C4").setValue("FALSE");
		ws.getRange2("B5").setValue("tE");
		ws.getRange2("C5").setValue("125te");

		oParser = new parserFormula("TEXTBEFORE(A1,B4:C5)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test");

		oParser = new parserFormula("TEXTBEFORE(A1,B4:C5, 0)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTBEFORE(A1,B4:C5, 1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test");

		oParser = new parserFormula("TEXTBEFORE(A1,B4:C5, 2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test125test2test3test4");

		oParser = new parserFormula("TEXTBEFORE(A1,B4:C5, 3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A106").setValue("2");
		ws.getRange2("A107").setValue("3");
		ws.getRange2("A109").setValue("4");
		ws.getRange2("A109").setValue("5");
		ws.getRange2("A110").setValue("6");
		ws.getRange2("A111").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:D10").cleanAll();
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
		// Case #1: String,String. Basic split: returns \'apple\'
		oParser = new parserFormula('TEXTBEFORE("apple,banana,pear",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("apple,banana,pear",",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'apple', 'Test: Positive case: String,String. Basic split: returns \'apple\'');
		// Case #2: String,String,Number. Second instance delimiter: returns \'apple,banana\'
		oParser = new parserFormula('TEXTBEFORE("apple,banana,pear",",",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("apple,banana,pear",",",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'apple,banana', 'Test: Positive case: String,String,Number. Second instance delimiter: returns \'apple,banana\'');
		// Case #3: String,String,Number. Explicit first instance same as default
		oParser = new parserFormula('TEXTBEFORE("apple,banana,pear",",",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("apple,banana,pear",",",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'apple', 'Test: Positive case: String,String,Number. Explicit first instance same as default');
		// Case #4: String,String,Number,Number. Case-insensitive match returns \'Apple,\'
		oParser = new parserFormula('TEXTBEFORE("Apple,Banana","banana",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("Apple,Banana","banana",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Apple,', 'Test: Positive case: String,String,Number,Number. Case-insensitive match returns \'Apple,\'');
		// Case #5: String,String,Number,Number. Match_end=1 allows matching at end
		oParser = new parserFormula('TEXTBEFORE("abcXYZ","XYZ",1,0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abcXYZ","XYZ",1,0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'abc', 'Test: Positive case: String,String,Number,Number. Match_end=1 allows matching at end');
		// Case #6: String,String,Number,Number,Number. Delimiter not found, if_not_found provided returns \'NotFound\'
		oParser = new parserFormula('TEXTBEFORE("zzz","a",1,0,0,"NotFound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("zzz","a",1,0,0,"NotFound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NotFound', 'Test: Positive case: String,String,Number,Number,Number. Delimiter not found, if_not_found provided returns \'NotFound\'');
		// Case #7: Number,String. Numeric input coerced to string, returns \'12\'
		oParser = new parserFormula('TEXTBEFORE(12345,"3")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(12345,"3") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Number,String. Numeric input coerced to string, returns \'12\'');
		// Case #8: Date,String. Date string split by \'-\' returns \'2025\'
		oParser = new parserFormula('TEXTBEFORE("2025-08-28","-")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("2025-08-28","-") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2025', 'Test: Positive case: Date,String. Date string split by \'-\' returns \'2025\'');
		// Case #9: Time,String. Time string split by \':\' returns \'12\'
		oParser = new parserFormula('TEXTBEFORE("12:30:45",":")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("12:30:45",":") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12', 'Test: Positive case: Time,String. Time string split by \':\' returns \'12\'');
		// Case #10: Formula,String. Formula result as text
		oParser = new parserFormula('TEXTBEFORE(UPPER("abc,def"),",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(UPPER("abc,def"),",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'ABC', 'Test: Positive case: Formula,String. Formula result as text');
		// Case #11: Array,String. Array constant returns array {a;c}
		oParser = new parserFormula('TEXTBEFORE({"a,b","c,d"},",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE({"a,b","c,d"},",") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Array,String. Array constant returns array {a;c}');
		// Case #12: Reference link,String. Reference link test
		oParser = new parserFormula('TEXTBEFORE(A100,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(A100,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Reference link,String. Reference link test');
		// Case #13: Area,String. Area (2-cell) test
		oParser = new parserFormula('TEXTBEFORE(A101:A102,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(A101:A102,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Area,String. Area (2-cell) test');
		// Case #14: Table,String. Table reference test
		oParser = new parserFormula('TEXTBEFORE(Table1[Column1],",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Table1[Column1],",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Table,String. Table reference test');
		// Case #15: Name,String. Named range test
		oParser = new parserFormula('TEXTBEFORE(TestName,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(TestName,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Name,String. Named range test');
		// Case #16: Name3D,String. Name3D test
		oParser = new parserFormula('TEXTBEFORE(TestName3D,".")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(TestName3D,".") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Name3D,String. Name3D test');
		// Case #17: Ref3D,String. Ref3D reference test
		oParser = new parserFormula('TEXTBEFORE(Sheet2!A1,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Sheet2!A1,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Ref3D,String. Ref3D reference test');
		// Case #18: Area3D,String. Area3D test
		oParser = new parserFormula('TEXTBEFORE(Sheet2!A2:A3,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Sheet2!A2:A3,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Area3D,String. Area3D test');

		// Negative cases:
		// Case #1: String,String,Number. instance_num=0 ? invalid returns #VALUE!
		oParser = new parserFormula('TEXTBEFORE("apple,banana",",",0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("apple,banana",",",0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number. instance_num=0 ? invalid returns #VALUE!');
		// Case #2: String,String,Number. Negative instance_num invalid
		oParser = new parserFormula('TEXTBEFORE("apple,banana",",",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("apple,banana",",",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'apple', 'Test: Negative case: String,String,Number. Negative instance_num invalid');
		// Case #3: Empty,String. Empty string returns empty text
		oParser = new parserFormula('TEXTBEFORE("",";")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("",";") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty,String. Empty string returns empty text');
		// Case #4: Error,String. Propagates error
		oParser = new parserFormula('TEXTBEFORE(#N/A,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(#N/A,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error,String. Propagates error');
		// Case #6: String,Error. Delimiter error propagates
		oParser = new parserFormula('TEXTBEFORE("abc",#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc",#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: String,Error. Delimiter error propagates');
		// Case #8: String,String,Number. Too large instance_num ? #N/A
		oParser = new parserFormula('TEXTBEFORE("abc,def",",",100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc,def",",",100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number. Too large instance_num ? #N/A');
		// Case #9: String,String,Number,Number. Case sensitive no match ? #N/A
		oParser = new parserFormula('TEXTBEFORE("abc,def","DEF",1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc,def","DEF",1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String,Number,Number. Case sensitive no match ? #N/A');
		// Case #10: String,String,Number,Number,Number. Delimiter not found returns #N/A
		oParser = new parserFormula('TEXTBEFORE("abc","z",1,0,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc","z",1,0,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String,Number,Number,Number. Delimiter not found returns #N/A');
		// Case #11: Reference link,String. Reference no match ? #N/A
		oParser = new parserFormula('TEXTBEFORE(A102,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(A102,"z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Reference link,String. Reference no match ? #N/A');
		// Case #12: Area,String. Area returns #N/A for missing delimiter
		oParser = new parserFormula('TEXTBEFORE(A103:A104,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(A103:A104,"z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area,String. Area returns #N/A for missing delimiter');
		// Case #13: Table,String. Table column no delimiter
		oParser = new parserFormula('TEXTBEFORE(Table1[Column1],"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Table1[Column1],"z") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Negative case: Table,String. Table column no delimiter');
		// Case #14: Name,String. Name ref no delimiter
		oParser = new parserFormula('TEXTBEFORE(TestName,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(TestName,"z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Name,String. Name ref no delimiter');
		// Case #15: Name3D,String. Name3D ref no delimiter
		oParser = new parserFormula('TEXTBEFORE(TestName3D,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(TestName3D,"z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "#N/A", 'Test: Negative case: Name3D,String. Name3D ref no delimiter');
		// Case #16: Ref3D,String. Ref3D no delimiter
		oParser = new parserFormula('TEXTBEFORE(Sheet2!A4,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Sheet2!A4,"z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D,String. Ref3D no delimiter');
		// Case #17: Area3D,String. Area3D no delimiter
		oParser = new parserFormula('TEXTBEFORE(Sheet2!A5:A6,"z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE(Sheet2!A5:A6,"z") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Area3D,String. Area3D no delimiter');
		// Case #20: String,String,Number. Delimiter overlapping returns empty or error
		oParser = new parserFormula('TEXTBEFORE("text","t",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("text","t",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'tex', 'Test: Negative case: String,String,Number. Delimiter overlapping returns empty or error');

		// Bounded cases:
		// Case #2: String,String,Number. Max valid instance within text
		oParser = new parserFormula('TEXTBEFORE("abc,def,ghi",",",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc,def,ghi",",",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String,String,Number. Max valid instance within text');
		// Case #3: String,String,Number. Lower bound instance_num=1
		oParser = new parserFormula('TEXTBEFORE("x,y,z",",",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("x,y,z",",",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'x', 'Test: Bounded case: String,String,Number. Lower bound instance_num=1');
		// Case #4: String,String,Number. Upper bound valid instance
		oParser = new parserFormula('TEXTBEFORE("abc,def",",",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc,def",",",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String,String,Number. Upper bound valid instance');
		// Case #6: String,String,Number,Number. match_end=1 lower bound
		oParser = new parserFormula('TEXTBEFORE("EndTest","Test",1,0,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("EndTest","Test",1,0,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'End', 'Test: Bounded case: String,String,Number,Number. match_end=1 lower bound');
		// Case #7: String,String,Number,Number,Number. if_not_found minimal case
		oParser = new parserFormula('TEXTBEFORE("abc","d",1,0,0,"NF")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("abc","d",1,0,0,"NF") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'NF', 'Test: Bounded case: String,String,Number,Number,Number. if_not_found minimal case');
		// Case #11: String,String. Space delimiter minimal use
		oParser = new parserFormula('TEXTBEFORE("start middle end"," ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("start middle end"," ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'start', 'Test: Bounded case: String,String. Space delimiter minimal use');
		// Case #12: String,String. Semicolon delimiter
		oParser = new parserFormula('TEXTBEFORE("multi;delim;test",";")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("multi;delim;test",";") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'multi', 'Test: Bounded case: String,String. Semicolon delimiter');
		// Case #13: String,String,Number. Semicolon second instance
		oParser = new parserFormula('TEXTBEFORE("multi;delim;test",";",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("multi;delim;test",";",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'multi;delim', 'Test: Bounded case: String,String,Number. Semicolon second instance');
		// Case #14: String,String. Early char delimiter
		oParser = new parserFormula('TEXTBEFORE("short","h")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("short","h") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 's', 'Test: Bounded case: String,String. Early char delimiter');
		// Case #15: String,String. Last char delimiter
		oParser = new parserFormula('TEXTBEFORE("boundary","y")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("boundary","y") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'boundar', 'Test: Bounded case: String,String. Last char delimiter');
		// Case #16: String,String. Number as delimiter inside string
		oParser = new parserFormula('TEXTBEFORE("Excel2025","2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("Excel2025","2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Excel', 'Test: Bounded case: String,String. Number as delimiter inside string');
		// Case #17: String,String. Underscore delimiter
		oParser = new parserFormula('TEXTBEFORE("a_b_c","_")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("a_b_c","_") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Bounded case: String,String. Underscore delimiter');
		// Case #18: String,String. Dot delimiter
		oParser = new parserFormula('TEXTBEFORE("dot.test","." )', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("dot.test","." ) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'dot', 'Test: Bounded case: String,String. Dot delimiter');
		// Case #19: String,String. Newline delimiter
		oParser = new parserFormula('TEXTBEFORE("line-break\nhere","\n")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("line-break\nhere","\n") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'line-break', 'Test: Bounded case: String,String. Newline delimiter');
		// Case #20: String,String. Tab character delimiter
		oParser = new parserFormula('TEXTBEFORE("tab\there","\t")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTBEFORE("tab\there","\t") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'tab', 'Test: Bounded case: String,String. Tab character delimiter');

		// Need to fix: array to string problem
		// Case #11: Array,String. Array constant returns array {a;c}
		// Case #16: Name3D,String. Name3D test
		// Case #13: Table,String. Table column no delimiter
		// Case #17: Area3D,String. Area3D no delimiter

	});

	QUnit.test("Test: \"TEXTAFTER\"", function (assert) {
		ws.getRange2("C3").setValue("txttextTeXttextText234text stext text");
		ws.getRange2("C6").setValue("te");

		oParser = new parserFormula("TEXTAFTER(C3;C6;1;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xtTeXttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;2;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "XttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;4;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;6;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;7;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt");

		oParser = new parserFormula("TEXTAFTER(C3;C6;8;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTAFTER(C3;C6;8;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");

		oParser = new parserFormula("TEXTAFTER(C3;C6;9;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");


		oParser = new parserFormula("TEXTAFTER(C3;C6;-1;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-2;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-3;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-4;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xt234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-5;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xtText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-6;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "XttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-7;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xtTeXttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-7;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "xtTeXttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-8;TRUE;FALSE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "error");

		oParser = new parserFormula("TEXTAFTER(C3;C6;-8;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "txttextTeXttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3:D4;C6;-8;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "txttextTeXttextText234text stext text");

		oParser = new parserFormula("TEXTAFTER(C3:D4;C6:D7;-8;TRUE;TRUE;\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "");

		oParser = new parserFormula("TEXTAFTER(C3;\"asdasd\";-8;TRUE;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTAFTER(12333;123;1;TRUE;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "33");

		oParser = new parserFormula("TEXTAFTER(12333;123;-1;TRUE;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "33");

		oParser = new parserFormula("TEXTAFTER(12333;123;-2;TRUE;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "12333");

		oParser = new parserFormula("TEXTAFTER(12333;123;3;TRUE;TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTAFTER(\"\",\",\",\"test\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTAFTER(\"\",\",\",\"test\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTAFTER(\"\",\",\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		oParser = new parserFormula("TEXTAFTER(B1,\",\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		ws.getRange2("A1").setValue("TRUE1Brown,Jim");

		oParser = new parserFormula("TEXTAFTER(A1,{\"i\",\"n\",\"b\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), ",Jim");

		oParser = new parserFormula("TEXTAFTER(A1,{\"i\",\"n\",TRUE})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1Brown,Jim");

		oParser = new parserFormula("TEXTAFTER(A1,{\"i\",\"n\",#VALUE!})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTAFTER(A1,{\"i\",\"J\",\"o\",\"O\",\"m\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "wn,Jim");

		ws.getRange2("A1").setValue("test125test2test3test4FALSEtest5");
		ws.getRange2("B4").setValue("w");
		ws.getRange2("B5").setValue("test");
		ws.getRange2("C4").setValue("FALSE");
		ws.getRange2("C5").setValue("125");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "125test2test3test4FALSEtest5");

		ws.getRange2("B5").setValue("tE");
		ws.getRange2("C5").setValue("125te");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "st2test3test4FALSEtest5");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5, 0)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5, 1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "st2test3test4FALSEtest5");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5, 2)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "test5");

		oParser = new parserFormula("TEXTAFTER(A1,B4:C5, 3)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A106").setValue("2");
		ws.getRange2("A107").setValue("3");
		ws.getRange2("A109").setValue("4");
		ws.getRange2("A109").setValue("5");
		ws.getRange2("A110").setValue("6");
		ws.getRange2("A111").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("1"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:D10").cleanAll();
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
		// Case #1: String,String. Basic string split, default instance_num=1 ? returns \'banana,pear\'
		oParser = new parserFormula('TEXTAFTER("apple,banana,pear",",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana,pear",",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'banana,pear', 'Test: Positive case: String,String. Basic string split, default instance_num=1 ? returns \'banana,pear\'');
		// Case #2: String,String,Number. Extract substring after 2nd comma ? returns \'pear\'
		oParser = new parserFormula('TEXTAFTER("apple,banana,pear",",",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana,pear",",",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'pear', 'Test: Positive case: String,String,Number. Extract substring after 2nd comma ? returns \'pear\'');
		// Case #3: String,String,Number. Returns substring after 3rd dash ? \'d\'
		oParser = new parserFormula('TEXTAFTER("a-b-c-d","-",3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("a-b-c-d","-",3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'd', 'Test: Positive case: String,String,Number. Returns substring after 3rd dash ? \'d\'');
		// Case #4: String,String,Number. Case-sensitive search, 2nd \'t\' returns \'\' (empty)
		oParser = new parserFormula('TEXTAFTER("ExcelTest","t",2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("ExcelTest","t",2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: String,String,Number. Case-sensitive search, 2nd \'t\' returns \'\' (empty)');
		// Case #5: String,String,Number,Number. Case-insensitive match_mode=1 returns substring after \'T\' ? \'\'
		oParser = new parserFormula('TEXTAFTER("ExcelTest","T",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("ExcelTest","T",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'est', 'Test: Positive case: String,String,Number,Number. Case-insensitive match_mode=1 returns substring after \'T\' ? \'\'');
		// Case #6: String,String. Extract after first space ? \'World\'
		oParser = new parserFormula('TEXTAFTER("Hello World"," ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("Hello World"," ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'World', 'Test: Positive case: String,String. Extract after first space ? \'World\'');
		// Case #7: Number,String. Numeric input coerced to text \'12345\', returns \'45\'
		oParser = new parserFormula('TEXTAFTER(12345,"3")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(12345,"3") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45', 'Test: Positive case: Number,String. Numeric input coerced to text \'12345\', returns \'45\'');
		// Case #8: Date,String. Date coerced to text \'1/15/2023\', returns \'15/2023\'
		oParser = new parserFormula('TEXTAFTER(DATE(2023,1,15),"/")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(DATE(2023,1,15),"/") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Date,String. Date coerced to text \'1/15/2023\', returns \'15/2023\'');
		// Case #9: Time,String. Time coerced to text \'12:30:00\', returns \'30:00\'
		oParser = new parserFormula('TEXTAFTER(TIME(12,30,0),":")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(TIME(12,30,0),":") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Time,String. Time coerced to text \'12:30:00\', returns \'30:00\'');
		// Case #10: Formula,String. Nested formula as text input ? returns \'EL\'
		oParser = new parserFormula('TEXTAFTER(UPPER("excel"),"C")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(UPPER("excel"),"C") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'EL', 'Test: Positive case: Formula,String. Nested formula as text input ? returns \'EL\'');
		// Case #11: Reference link,String. Reference input, A100 points to text \'apple,orange\'
		oParser = new parserFormula('TEXTAFTER(A100,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(A100,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Reference link,String. Reference input, A100 points to text \'apple,orange\'');
		// Case #12: Area,String. Area input (2 cells), processes first cell only
		oParser = new parserFormula('TEXTAFTER(A101:A102,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(A101:A102,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Area,String. Area input (2 cells), processes first cell only');
		// Case #13: Table,String. Table reference as text input
		oParser = new parserFormula('TEXTAFTER(Table1[Column1],",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(Table1[Column1],",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Table,String. Table reference as text input');
		// Case #14: Name,String. Using named range TestName
		oParser = new parserFormula('TEXTAFTER(TestName,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(TestName,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Name,String. Using named range TestName');
		// Case #15: Name3D,String. Using 3D named range TestName3D2
		oParser = new parserFormula('TEXTAFTER(TestName3D2,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(TestName3D2,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Name3D,String. Using 3D named range TestName3D2');
		// Case #16: Area3D,String. Using 3D area from Sheet2
		oParser = new parserFormula('TEXTAFTER(Sheet2!A1:A2,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(Sheet2!A1:A2,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Area3D,String. Using 3D area from Sheet2');
		// Case #17: Ref3D,String. Using 3D reference from Sheet2
		oParser = new parserFormula('TEXTAFTER(Sheet2!A1,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(Sheet2!A1,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Positive case: Ref3D,String. Using 3D reference from Sheet2');
		// Case #20: Array,String. Array constant input ? returns {\'ne\',\'wo\'}
		oParser = new parserFormula('TEXTAFTER({"one","two"},"o")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER({"one","two"},"o") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 'ne', 'Test: Positive case: Array,String. Array constant input ? returns {\'ne\',\'wo\'}');

		// Negative cases:
		// Case #1: String,Number. Invalid delimiter type (number) ? #VALUE!
		oParser = new parserFormula('TEXTAFTER("apple,banana",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,Number. Invalid delimiter type (number) ? #VALUE!');
		// Case #2: Empty,String. Empty text, delimiter \'a\' ? returns #N/A
		oParser = new parserFormula('TEXTAFTER("","a")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("","a") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Empty,String. Empty text, delimiter \'a\' ? returns #N/A');
		// Case #3: String,String,Number. Instance_num too large ? #N/A
		oParser = new parserFormula('TEXTAFTER("apple,banana",",",5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",",",5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String,Number. Instance_num too large ? #N/A');
		// Case #4: String,String,Number. Negative instance_num not valid ? #VALUE!
		oParser = new parserFormula('TEXTAFTER("apple,banana",",",-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",",",-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'banana', 'Test: Negative case: String,String,Number. Negative instance_num not valid ? #VALUE!');
		// Case #5: Error,String. Input error propagates
		oParser = new parserFormula('TEXTAFTER(#N/A,",")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(#N/A,",") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error,String. Input error propagates');
		// Case #6: String,String,Error. Error as instance_num propagates ? #DIV/0!
		oParser = new parserFormula('TEXTAFTER("apple,banana",",",#DIV/0!)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",",",#DIV/0!) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#DIV/0!', 'Test: Negative case: String,String,Error. Error as instance_num propagates ? #DIV/0!');
		// Case #7: String,String,Number,Number. Invalid match_mode argument ? #VALUE!
		oParser = new parserFormula('TEXTAFTER("apple,banana",",",1,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",",",1,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number,Number. Invalid match_mode argument ? #VALUE!');
		// Case #8: String,String,Number,Number,Number. Invalid match_end argument ? #VALUE!
		oParser = new parserFormula('TEXTAFTER("apple,banana",",",1,0,2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple,banana",",",1,0,2) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String,String,Number,Number,Number. Invalid match_end argument ? #VALUE!');
		// Case #9: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("apple banana","z")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("apple banana","z") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #11: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test11","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test11","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #12: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test12","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test12","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #13: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test13","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test13","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #14: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test14","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test14","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #15: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test15","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test15","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #16: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test16","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test16","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #17: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test17","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test17","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #18: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test18","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test18","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #19: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test19","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test19","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');
		// Case #20: String,String. Delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("test20","notfound")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("test20","notfound") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: String,String. Delimiter not found ? #N/A');

		// Bounded cases:
		// Case #3: String,String,Number. Delimiter at beginning ? returns \'bc\'
		oParser = new parserFormula('TEXTAFTER("abc","a",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("abc","a",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'bc', 'Test: Bounded case: String,String,Number. Delimiter at beginning ? returns \'bc\'');
		// Case #5: String,String,Number. Delimiter in middle ? returns \'c\'
		oParser = new parserFormula('TEXTAFTER("abc","b",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("abc","b",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'c', 'Test: Bounded case: String,String,Number. Delimiter in middle ? returns \'c\'');
		// Case #6: String,String,Number. Instance_num exceeds length ? #N/A
		oParser = new parserFormula('TEXTAFTER("abc","a",9999)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("abc","a",9999) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String,String,Number. Instance_num exceeds length ? #N/A');
		// Case #7: String,String,Number,Number. Case-insensitive bounded test
		oParser = new parserFormula('TEXTAFTER("ABC","a",1,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("ABC","a",1,1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'BC', 'Test: Bounded case: String,String,Number,Number. Case-insensitive bounded test');
		// Case #8: String,String,Number,Number. Case-sensitive bounded test, delimiter not found ? #N/A
		oParser = new parserFormula('TEXTAFTER("ABC","a",1,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("ABC","a",1,0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String,String,Number,Number. Case-sensitive bounded test, delimiter not found ? #N/A');
		// Case #10: Date,String,Number. Earliest date supported by Excel ? returns \'1/1900\'
		oParser = new parserFormula('TEXTAFTER(DATE(1900,1,1),"/",1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(DATE(1900,1,1),"/",1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: Date,String,Number. Earliest date supported by Excel ? returns \'1/1900\'');
		// Case #11: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER("limit11","i")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER("limit11","i") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'mit11', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #12: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E+307,"i")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E+307,"i") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #13: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E-307,"i")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E-307,"i") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #14: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E+307,"+")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E+307,"+") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '307', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #15: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E-307,"-")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E-307,"-") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '307', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #16: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E+307,"E")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E+307,"E") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '+307', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #17: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E-307,"E")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E-307,"E") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '-307', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #18: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E+307,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E+307,"0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '7', 'Test: Bounded case: String,String. Bounded variation edge test');
		// Case #19: String,String. Bounded variation edge test
		oParser = new parserFormula('TEXTAFTER(1E-307,"0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTAFTER(1E-307,"0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '7', 'Test: Bounded case: String,String. Bounded variation edge test');

		// TODO problems with array
		// Need to fix: results diff form ms, boundary cases diff
		// Case #20: Array,String. Array constant input ? returns {\'ne\',\'wo\'}
		// Case #7: String,String,Number,Number. Invalid match_mode argument ? #VALUE!
		// Case #8: String,String,Number,Number,Number. Invalid match_end argument ? #VALUE!
		// Case #16: String,String. Bounded variation edge test
		// Case #17: String,String. Bounded variation edge test


	});

	QUnit.test("Test: \"TEXTSPLIT\"", function (assert) {
		ws.getRange2("C3").setValue(
			"test1del1TEst2Del2#NUM!DEl2 test3 tedel3del3st3del1del1del1 del2del2del2 testdel1\n" + "test3 del1del2del1 test123testdel3testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4");

		ws.getRange2("C4").setValue("tedel1asd");
		ws.getRange2("D3").setValue("fffdel1sdf");
		ws.getRange2("D4").setValue("rflde1");

		ws.getRange2("C5").setValue("del2");
		ws.getRange2("D5").setValue("teST");
		ws.getRange2("C6").setValue("del3");
		ws.getRange2("D6").setValue("del1");

		ws.getRange2("C11").setValue("error");

		let array;
		oParser = new parserFormula("TEXTSPLIT(C3,\"del1\")", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "test1");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "TEst2Del2#NUM!DEl2 test3 tedel3del3st3");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), " del2del2del2 test");
		assert.strictEqual(array.getElementRowCol(0, 5).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(0, 6).getValue(), "del2");
		assert.strictEqual(array.getElementRowCol(0, 7).getValue(), " test123testdel3testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4");

		oParser = new parserFormula("TEXTSPLIT(C3,{\"del1\",\"del2\"})", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "test1");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "TEst2Del2#NUM!DEl2 test3 tedel3del3st3");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), " ");
		assert.strictEqual(array.getElementRowCol(0, 5).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 6).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 7).getValue(), " test");
		assert.strictEqual(array.getElementRowCol(0, 8).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(0, 9).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 10).getValue(), "");
		assert.strictEqual(array.getElementRowCol(0, 11).getValue(), " test123testdel3testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4");


		oParser = new parserFormula("TEXTSPLIT(C3,{\"del1\",\"del2\"},\"del3\")", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();
		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), "test1");
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), "TEst2Del2#NUM!DEl2 test3 te");
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), "#N/A");

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), "");
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), "#N/A");
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), "#N/A");

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), "st3");
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 3).getValue(), " ");
		assert.strictEqual(array.getElementRowCol(2, 4).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 5).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 6).getValue(), " test");
		assert.strictEqual(array.getElementRowCol(2, 7).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(2, 8).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 9).getValue(), "");
		assert.strictEqual(array.getElementRowCol(2, 10).getValue(), " test123test");

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), "testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4");
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), "#N/A");
		assert.strictEqual(array.getElementRowCol(3, 2).getValue(), "#N/A");

		oParser = new parserFormula("TEXTSPLIT(C3,{\"del1\",\"del2\"},\"del3\",,TRUE)", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'test2');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), '#num!');
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), ' test3 te');
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), '#N/A');

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '#N/A');

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 'st3');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 3).getValue(), ' ');
		assert.strictEqual(array.getElementRowCol(2, 4).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 5).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 6).getValue(), ' test');
		assert.strictEqual(array.getElementRowCol(2, 7).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(2, 8).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 9).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 10).getValue(), ' test123test');

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 'testw');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), '#N/A');

		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), 'test2del4jjjdel4rrr');
		assert.strictEqual(array.getElementRowCol(4, 1).getValue(), '23rrrdel4');
		assert.strictEqual(array.getElementRowCol(4, 2).getValue(), '#N/A');


		oParser = new parserFormula("TEXTSPLIT(C3,{\"del2\"},{\"del3\",\"del1\"},FALSE,TRUE,\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(0, 3).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(0, 4).getValue(), '');

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'test2');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '#num!');
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), ' test3 te');
		assert.strictEqual(array.getElementRowCol(1, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), '');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(2, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), 'st3');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(3, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), '');
		assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(4, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), '');
		assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(5, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), ' ');
		assert.strictEqual(array.getElementRowCol(6, 1).getValue(), '');
		assert.strictEqual(array.getElementRowCol(6, 2).getValue(), '');
		assert.strictEqual(array.getElementRowCol(6, 3).getValue(), ' test');

		assert.strictEqual(array.getElementRowCol(7, 0).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(7, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(8, 0).getValue(), '');
		assert.strictEqual(array.getElementRowCol(8, 1).getValue(), '');
		assert.strictEqual(array.getElementRowCol(8, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(8, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(9, 0).getValue(), ' test123test');
		assert.strictEqual(array.getElementRowCol(9, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(9, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(9, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(10, 0).getValue(), 'testw');
		assert.strictEqual(array.getElementRowCol(10, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(10, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(10, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(11, 0).getValue(), 'test2del4jjjdel4rrr');
		assert.strictEqual(array.getElementRowCol(11, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(11, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(11, 3).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(12, 0).getValue(), '23rrrdel4');
		assert.strictEqual(array.getElementRowCol(12, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(12, 2).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(12, 3).getValue(), 'error');


		oParser = new parserFormula("TEXTSPLIT(C3,{\"del2\"},{\"del3\",\"del1\"},TRUE,TRUE,\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(0, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'test2');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), '#num!');
		assert.strictEqual(array.getElementRowCol(1, 2).getValue(), ' test3 te');

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 'st3');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(2, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), ' ');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), ' test');
		assert.strictEqual(array.getElementRowCol(3, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(4, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), ' test123test');
		assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(5, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 'testw');
		assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(6, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(7, 0).getValue(), 'test2del4jjjdel4rrr');
		assert.strictEqual(array.getElementRowCol(7, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(7, 2).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(8, 0).getValue(), '23rrrdel4');
		assert.strictEqual(array.getElementRowCol(8, 1).getValue(), 'error');
		assert.strictEqual(array.getElementRowCol(8, 2).getValue(), 'error');

		oParser = new parserFormula("TEXTSPLIT(C3,{\"del2\",\"teST\"},{\"del3\",\"del1\"},TRUE,,\"error\")", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'TEst2Del2#NUM!DEl2 test3 te');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 'st3');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), ' ');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), ' test');

		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), ' test123test');
		assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 'testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4');
		assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 'error');


		oParser = new parserFormula("TEXTSPLIT(C3,,\"del\",,,)", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), '1TEst2Del2#NUM!DEl2 test3 te');
		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), '3');
		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), '3st3');
		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), '1');
		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), '1');
		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), '1 ');
		assert.strictEqual(array.getElementRowCol(7, 0).getValue(), '2');
		assert.strictEqual(array.getElementRowCol(8, 0).getValue(), '2');
		assert.strictEqual(array.getElementRowCol(9, 0).getValue(), '2 test');
		assert.strictEqual(array.getElementRowCol(10, 0).getValue(), "1\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(11, 0).getValue(), '1');
		assert.strictEqual(array.getElementRowCol(12, 0).getValue(), '2');
		assert.strictEqual(array.getElementRowCol(13, 0).getValue(), '1 test123test');
		assert.strictEqual(array.getElementRowCol(14, 0).getValue(), '3testwDEL3test2DeL4jjj');
		assert.strictEqual(array.getElementRowCol(15, 0).getValue(), '4rrrDEL123rrrdEl4');

		oParser = new parserFormula("TEXTSPLIT(C3,,,,,)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(C3,\"test\",\"\",,,)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(C3,,,,,)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(C2,\"\",\"test\",,,)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(C3:D4,\"del1\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test1");


		oParser = new parserFormula("TEXTSPLIT(C3,C5:D5,C6:D6,TRUE,,C11)", "A1", ws);
		assert.ok(oParser.parse());
		array = oParser.calculate();

		assert.strictEqual(array.getElementRowCol(0, 0).getValue(), 'test1');
		assert.strictEqual(array.getElementRowCol(0, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(1, 0).getValue(), 'TEst2Del2#NUM!DEl2 test3 te');
		assert.strictEqual(array.getElementRowCol(1, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(2, 0).getValue(), 'st3');
		assert.strictEqual(array.getElementRowCol(2, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(3, 0).getValue(), ' ');
		assert.strictEqual(array.getElementRowCol(3, 1).getValue(), ' test');

		assert.strictEqual(array.getElementRowCol(4, 0).getValue(), "\n" + "test3 ");
		assert.strictEqual(array.getElementRowCol(4, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(5, 0).getValue(), ' test123test');
		assert.strictEqual(array.getElementRowCol(5, 1).getValue(), 'error');

		assert.strictEqual(array.getElementRowCol(6, 0).getValue(), 'testwDEL3test2DeL4jjjdel4rrrDEL123rrrdEl4');
		assert.strictEqual(array.getElementRowCol(6, 1).getValue(), 'error');

		oParser = new parserFormula("TEXTSPLIT(C3,C5:D5,C6:D6,TRUE,,C11:D11)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(C3,C5:D5,C6:D6,{TRUE,FALSE},,C11)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test1");

		/*oParser = new parserFormula("TEXTSPLIT(C3,C5:D5,C6:D6,C59:D59,{TRUE,FALSE},C11)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "test1");*/

		ws.getRange2("A2").setValue("Do. Or do not. There is no try. -Anonymous");
		oParser = new parserFormula("TEXTSPLIT(A2,,\".\")", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "Do");
		assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), " Or do not");
		assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), " There is no try");
		assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), " -Anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,{\".\";\"-\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "Do");
		assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), " Or do not");
		assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), " There is no try");
		assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), " ");
		assert.strictEqual(oParser.calculate().getElementRowCol(4, 0).getValue(), "Anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,{\".\";\"d\"})", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "Do");
		assert.strictEqual(oParser.calculate().getElementRowCol(1, 0).getValue(), " Or ");
		assert.strictEqual(oParser.calculate().getElementRowCol(2, 0).getValue(), "o not");
		assert.strictEqual(oParser.calculate().getElementRowCol(3, 0).getValue(), " There is no try");
		assert.strictEqual(oParser.calculate().getElementRowCol(4, 0).getValue(), " -Anonymous");

		ws.getRange2("C1").setValue("1");
		oParser = new parserFormula("TEXTSPLIT(A2,C1,C1,C1,C1,C1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "do. or do not. there is no try. -anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,C1,C1,C1,C1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "do. or do not. there is no try. -anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,,C1,C1,C1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula("TEXTSPLIT(A2,,C1,,C1,C1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "do. or do not. there is no try. -anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,C1,,,C1)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "Do. Or do not. There is no try. -Anonymous");

		oParser = new parserFormula("TEXTSPLIT(A2,,C1,,,)", "A1", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getElementRowCol(0, 0).getValue(), "Do. Or do not. There is no try. -Anonymous");

		ws.getRange2("A1:C214").cleanAll();
		// Data for reference link. Use A100-A111
		ws.getRange2("A100").setValue("0.5");
		ws.getRange2("A101").setValue("1.5");
		ws.getRange2("A104").setValue("-1");
		// For area
		ws.getRange2("A102").setValue("0.5");
		ws.getRange2("A103").setValue("Text");
		ws.getRange2("A105").setValue("1");
		ws.getRange2("A106").setValue("2");
		ws.getRange2("A107").setValue("3");
		ws.getRange2("A109").setValue("4");
		ws.getRange2("A109").setValue("5");
		ws.getRange2("A110").setValue("6");
		ws.getRange2("A111").setValue("");

		// Table type. Use A601:L6**
		getTableType(599, 0, 600, 1);
		ws.getRange2("A601").setValue("123"); // Number (Column1)
		ws.getRange2("B601").setValue("1s"); // Text (Column2)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1:D10").cleanAll();
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
		// Case #1: String, String, String, Boolean. Basic string splitting with comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: String, String, String, Boolean. Basic string splitting with comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #2: Number, String, String, Boolean. Number converted to text, colon column delimiter, comma row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(12345,":",",",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(12345,":",",",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '12345', 'Test: Positive case: Number, String, String, Boolean. Number converted to text, colon column delimiter, comma row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #3: Formula, String, String, Boolean. Nested formula (CONCAT), comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(CONCAT("X","Y"),",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(CONCAT("X","Y"),",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'XY', 'Test: Positive case: Formula, String, String, Boolean. Nested formula (CONCAT), comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #4: Reference link, String, String, Boolean. Reference link to cell with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(A100,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(A100,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Positive case: Reference link, String, String, Boolean. Reference link to cell with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #5: Area, String, String, Boolean. Single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(A101:A101,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(A101:A101,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Positive case: Area, String, String, Boolean. Single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #6: Array, String, String, Boolean. Array input, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT({"A,B","C,D"},",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT({"A,B","C,D"},",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: Array, String, String, Boolean. Array input, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #7: Name, String, String, Boolean. Named range with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TestName,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TestName,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '-0.5', 'Test: Positive case: Name, String, String, Boolean. Named range with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #8: Name3D, String, String, Boolean. 3D named range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TestName3D,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TestName3D,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '-0.5', 'Test: Positive case: Name3D, String, String, Boolean. 3D named range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #9: Ref3D, String, String, Boolean. 3D reference to cell with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Sheet2!A1,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Sheet2!A1,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '1', 'Test: Positive case: Ref3D, String, String, Boolean. 3D reference to cell with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #10: Area3D, String, String, Boolean. 3D single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Sheet2!A2:A2,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Sheet2!A2:A2,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area3D, String, String, Boolean. 3D single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #11: Table, String, String, Boolean. Table structured reference with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Table1[Column1],",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Table1[Column1],",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '123', 'Test: Positive case: Table, String, String, Boolean. Table structured reference with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #12: Date, String, String, Boolean. Date converted to text, slash column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(DATE(2025,1,1),"/",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(DATE(2025,1,1),"/",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '45658', 'Test: Positive case: Date, String, String, Boolean. Date converted to text, slash column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #13: Time, String, String, Boolean. Time converted to text, colon column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TIME(12,0,0),":",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TIME(12,0,0),":",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '0.5', 'Test: Positive case: Time, String, String, Boolean. Time converted to text, colon column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #14: Formula, String, String, Boolean. TEXTSPLIT inside SUM formula, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('SUM(LEN(TEXTSPLIT("A,B,C",",",";",TRUE)),1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(LEN(TEXTSPLIT("A,B,C",",",";",TRUE)),1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula, String, String, Boolean. TEXTSPLIT inside SUM formula, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #15: String, String, String, Boolean. Consecutive delimiters, comma column delimiter, semicolon row delimiter, ignore_empty TRUE skips empty. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,,B",",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,,B",",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: String, String, String, Boolean. Consecutive delimiters, comma column delimiter, semicolon row delimiter, ignore_empty TRUE skips empty. 4 arguments used.');
		// Case #16: String, String, String, Boolean. Consecutive delimiters, comma column delimiter, semicolon row delimiter, ignore_empty FALSE includes empty. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,,B",",",";",FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,,B",",",";",FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: String, String, String, Boolean. Consecutive delimiters, comma column delimiter, semicolon row delimiter, ignore_empty FALSE includes empty. 4 arguments used.');
		// Case #17: String, String, String, Boolean, Number. Case-insensitive match, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 5 arguments used.
		oParser = new parserFormula('TEXTSPLIT("a,B,C",",",";",TRUE,1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("a,B,C",",",";",TRUE,1) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'a', 'Test: Positive case: String, String, String, Boolean, Number. Case-insensitive match, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 5 arguments used.');
		// Case #18: Formula, String, String, Boolean, Number. Nested IF formula, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 5 arguments used.
		oParser = new parserFormula('TEXTSPLIT(IF(TRUE,"X,Y","Z"),",",";",TRUE,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(IF(TRUE,"X,Y","Z"),",",";",TRUE,0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'X', 'Test: Positive case: Formula, String, String, Boolean, Number. Nested IF formula, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 5 arguments used.');
		// Case #19: String, String, String, Boolean, Number, String. Custom pad_with value, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 6 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",",",";",TRUE,0,"N/A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",",",";",TRUE,0,"N/A") is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: String, String, String, Boolean, Number, String. Custom pad_with value, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 6 arguments used.');
		// Case #20: Reference link, String, String, Boolean, Number, String. Reference link, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive, empty pad_with. 6 arguments used.
		oParser = new parserFormula('TEXTSPLIT(A102,",",";",TRUE,0,"")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(A102,",",";",TRUE,0,"") is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Positive case: Reference link, String, String, Boolean, Number, String. Reference link, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive, empty pad_with. 6 arguments used.');

		// Negative cases:
		// Case #1: Error, String, String, Boolean. Error input (NA) propagates #N/A error. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(NA(),",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(NA(),",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error, String, String, Boolean. Error input (NA) propagates #N/A error. 4 arguments used.');
		// Case #2: Empty, String, String, Boolean. Empty text input returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty, String, String, Boolean. Empty text input returns #VALUE!. 4 arguments used.');
		// Case #3: Boolean, String, String, Boolean. Boolean text input returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TRUE,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TRUE,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'TRUE', 'Test: Negative case: Boolean, String, String, Boolean. Boolean text input returns #VALUE!. 4 arguments used.');
		// Case #4: Formula, String, String, Boolean. Formula resulting in #NUM! for text returns #NUM!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(SQRT(-1),",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(SQRT(-1),",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula, String, String, Boolean. Formula resulting in #NUM! for text returns #NUM!. 4 arguments used.');
		// Case #5: String, Empty, String, Boolean. Empty col_delimiter returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",,";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",,";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A,B,C', 'Test: Negative case: String, Empty, String, Boolean. Empty col_delimiter returns #VALUE!. 4 arguments used.');
		// Case #6: String, Boolean, String, Boolean. Boolean col_delimiter returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",TRUE,";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",TRUE,";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A,B,C', 'Test: Negative case: String, Boolean, String, Boolean. Boolean col_delimiter returns #VALUE!. 4 arguments used.');
		// Case #7: Area, String, String, Boolean. Multi-cell range returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(A103:A104,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(A103:A104,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'Text', 'Test: Negative case: Area, String, String, Boolean. Multi-cell range returns #VALUE!. 4 arguments used.');
		// Case #8: Area3D, String, String, Boolean. 3D multi-cell range returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Sheet2!A3:A4,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Sheet2!A3:A4,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '?', 'Test: Negative case: Area3D, String, String, Boolean. 3D multi-cell range returns #VALUE!. 4 arguments used.');
		// Case #9: Name, String, String, Boolean. Named range with area returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TestNameArea2,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TestNameArea2,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name, String, String, Boolean. Named range with area returns #VALUE!. 4 arguments used.');
		// Case #10: Name3D, String, String, Boolean. 3D named range with area returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(TestNameArea3D2,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(TestNameArea3D2,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name3D, String, String, Boolean. 3D named range with area returns #VALUE!. 4 arguments used.');
		// Case #11: Ref3D, String, String, Boolean. 3D reference to cell with non-string value returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Sheet2!A5,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Sheet2!A5,",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D, String, String, Boolean. 3D reference to cell with non-string value returns #VALUE!. 4 arguments used.');
		// Case #12: Table, String, String, Boolean. Table column with non-string data returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(Table1[Column2],",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(Table1[Column2],",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '1s', 'Test: Negative case: Table, String, String, Boolean. Table column with non-string data returns #VALUE!. 4 arguments used.');
		// Case #13: Array, String, String, Boolean. Array with boolean returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT({TRUE},",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT({TRUE},",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'TRUE', 'Test: Negative case: Array, String, String, Boolean. Array with boolean returns #VALUE!. 4 arguments used.');
		// Case #14: String, Number, String, Boolean. Numeric col_delimiter returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",123,";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",123,";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A,B,C', 'Test: Negative case: String, Number, String, Boolean. Numeric col_delimiter returns #VALUE!. 4 arguments used.');
		// Case #15: String, Date, String, Boolean. Date as col_delimiter returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",DATE(2025,1,1),";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",DATE(2025,1,1),";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A,B,C', 'Test: Negative case: String, Date, String, Boolean. Date as col_delimiter returns #VALUE!. 4 arguments used.');
		// Case #16: String, Time, String, Boolean. Time as col_delimiter returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",TIME(12,0,0),";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",TIME(12,0,0),";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A,B,C', 'Test: Negative case: String, Time, String, Boolean. Time as col_delimiter returns #VALUE!. 4 arguments used.');
		// Case #17: String, String, String, String. Non-boolean ignore_empty returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",",",";","Invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",",",";","Invalid") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String, String. Non-boolean ignore_empty returns #VALUE!. 4 arguments used.');
		// Case #18: String, String, String, Boolean, String. Non-numeric match_mode returns #VALUE!. 5 arguments used.
		oParser = new parserFormula('TEXTSPLIT("A,B,C",",",";",TRUE,"Invalid")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT("A,B,C",",",";",TRUE,"Invalid") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String, String, String, Boolean, String. Non-numeric match_mode returns #VALUE!. 5 arguments used.');
		// Case #19: Reference link, String, String, Boolean. Reference link to cell with number returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(A106,",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(A106,",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '2', 'Test: Negative case: Reference link, String, String, Boolean. Reference link to cell with number returns #VALUE!. 4 arguments used.');
		// Case #20: Array, String, String, Boolean. Array with mixed data (including boolean) returns #VALUE!. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT({"A,B","C",TRUE},",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT({"A,B","C",TRUE},",",";",TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Negative case: Array, String, String, Boolean. Array with mixed data (including boolean) returns #VALUE!. 4 arguments used.');

		// Bounded cases:
		// Case #1: String, String, String, Boolean. Maximum string length for splitting (32,767 characters, split into ~16,384 elements), comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		oParser = new parserFormula('TEXTSPLIT(REPT("A,",16384),",",";",TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(REPT("A,",16384),",",";",TRUE) is parsed.');
		//? assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Bounded case: String, String, String, Boolean. Maximum string length for splitting (32,767 characters, split into ~16,384 elements), comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.');
		// Case #3: Number, String, String, Boolean, Number. Maximum numeric value converted to text, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 5 arguments used.
		oParser = new parserFormula('TEXTSPLIT(1E+307,",",";",TRUE,0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TEXTSPLIT(1E+307,",",";",TRUE,0) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '1e+307', 'Test: Bounded case: Number, String, String, Boolean, Number. Maximum numeric value converted to text, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive match. 5 arguments used.');

		// Need to fix:
		// Case #17: String, String, String, String. Non-boolean ignore_empty returns #VALUE!. 4 arguments used.
		// Case #18: String, String, String, Boolean, String. Non-numeric match_mode returns #VALUE!. 5 arguments used.
		// Case #4: Reference link, String, String, Boolean. Reference link to cell with string, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		// Case #5: Area, String, String, Boolean. Single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		// Case #10: Area3D, String, String, Boolean. 3D single-cell range, comma column delimiter, semicolon row delimiter, ignore_empty TRUE. 4 arguments used.
		// Case #20: Reference link, String, String, Boolean, Number, String. Reference link, comma column delimiter, semicolon row delimiter, ignore_empty TRUE, case-sensitive, empty pad_with. 6 arguments used.
		// Case #2: Empty, String, String, Boolean. Empty text input returns #VALUE!. 4 arguments used.
		// Case #11: Ref3D, String, String, Boolean. 3D reference to cell with non-string value returns #VALUE!. 4 arguments used.

	});

	QUnit.test("Test: \"TRIM\"", function (assert) {

		oParser = new parserFormula("TRIM(\"     abc         def      \")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "abc def");

		oParser = new parserFormula("TRIM(\" First Quarter Earnings \")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "First Quarter Earnings");


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
		// Case #1: String. String with leading/trailing spaces. 1 argument used.
		oParser = new parserFormula('TRIM("  text  ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM("  text  ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: String. String with leading/trailing spaces. 1 argument used.');
		// Case #2: String. String without extra spaces. 1 argument used.
		oParser = new parserFormula('TRIM("text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM("text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: String. String without extra spaces. 1 argument used.');
		// Case #3: Number. Number converted to text. 1 argument used.
		oParser = new parserFormula('TRIM(123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: Number. Number converted to text. 1 argument used.');
		// Case #4: Formula. Nested CONCAT formula producing text with spaces. 1 argument used.
		oParser = new parserFormula('TRIM(CONCAT("  te","xt  "))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(CONCAT("  te","xt  ")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: Formula. Nested CONCAT formula producing text with spaces. 1 argument used.');
		// Case #5: Reference link. Reference to cell with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link. Reference to cell with string containing spaces. 1 argument used.');
		// Case #6: Area. Single-cell range with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.5', 'Test: Positive case: Area. Single-cell range with string containing spaces. 1 argument used.');
		// Case #7: Array. Array with single text element. 1 argument used.
		oParser = new parserFormula('TRIM({"  text  "})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM({"  text  "}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: Array. Array with single text element. 1 argument used.');
		// Case #8: Name. Named range with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name. Named range with string containing spaces. 1 argument used.');
		// Case #9: Name3D. 3D named range with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-0.5', 'Test: Positive case: Name3D. 3D named range with string containing spaces. 1 argument used.');
		// Case #10: Ref3D. 3D reference to cell with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Ref3D. 3D reference to cell with string containing spaces. 1 argument used.');
		// Case #11: Area3D. 3D single-cell range with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '2', 'Test: Positive case: Area3D. 3D single-cell range with string containing spaces. 1 argument used.');
		// Case #12: Table. Table structured reference with string containing spaces. 1 argument used.
		oParser = new parserFormula('TRIM(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Positive case: Table. Table structured reference with string containing spaces. 1 argument used.');
		// Case #13: Date. Date converted to text serial number. 1 argument used.
		oParser = new parserFormula('TRIM(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '45658', 'Test: Positive case: Date. Date converted to text serial number. 1 argument used.');
		// Case #14: Time. Time converted to text decimal. 1 argument used.
		oParser = new parserFormula('TRIM(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Time. Time converted to text decimal. 1 argument used.');
		// Case #15: Formula. Nested IF returning string with spaces. 1 argument used.
		oParser = new parserFormula('TRIM(IF(TRUE,"  text  ","text"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(IF(TRUE,"  text  ","text")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: Formula. Nested IF returning string with spaces. 1 argument used.');
		// Case #16: String. String with multiple internal spaces. 1 argument used.
		oParser = new parserFormula('TRIM("   multiple   spaces   ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM("   multiple   spaces   ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'multiple spaces', 'Test: Positive case: String. String with multiple internal spaces. 1 argument used.');
		// Case #17: Array. Multi-element array with text. Returns first element trimmed. 1 argument used.
		oParser = new parserFormula('TRIM({"text","  text  "})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM({"text","  text  "}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text', 'Test: Positive case: Array. Multi-element array with text. Returns first element trimmed. 1 argument used.');
		// Case #18: Reference link. Reference to cell with text without extra spaces. 1 argument used.
		oParser = new parserFormula('TRIM(A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1.5', 'Test: Positive case: Reference link. Reference to cell with text without extra spaces. 1 argument used.');
		// Case #19: Formula. Nested TEXT formula converting number to string. 1 argument used.
		oParser = new parserFormula('TRIM(TEXT(123,"0"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TEXT(123,"0")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123', 'Test: Positive case: Formula. Nested TEXT formula converting number to string. 1 argument used.');
		// Case #20: Name. Named range with text without extra spaces. 1 argument used.
		oParser = new parserFormula('TRIM(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Name. Named range with text without extra spaces. 1 argument used.');
		// Case #21: String. String with multiple spaces between words. 1 argument used.
		oParser = new parserFormula('TRIM("text   text")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM("text   text") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'text text', 'Test: Positive case: String. String with multiple spaces between words. 1 argument used.');
		// Case #22: Area. Single-cell range with multiple spaces. 1 argument used.
		oParser = new parserFormula('TRIM(A102:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A102:A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Area. Single-cell range with multiple spaces. 1 argument used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('TRIM(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #3: Boolean. Boolean FALSE converted to text "FALSE". 1 argument used.
		oParser = new parserFormula('TRIM(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'FALSE', 'Test: Negative case: Boolean. Boolean FALSE converted to text "FALSE". 1 argument used.');
		// Case #4: Area. Multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('TRIM(A104:A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A104:A105) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '-1', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #6: Ref3D. 3D reference to cell with non-text value returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Text', 'Test: Negative case: Ref3D. 3D reference to cell with non-text value returns text conversion. 1 argument used.');
		// Case #7: Name. Named range with multi-cell area returns #VALUE!. 1 argument used.
		oParser = new parserFormula('TRIM(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name. Named range with multi-cell area returns #VALUE!. 1 argument used.');
		// Case #8: Name3D. 3D named range with multi-cell area returns #VALUE!. 1 argument used.
		oParser = new parserFormula('TRIM(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.8', 'Test: Negative case: Name3D. 3D named range with multi-cell area returns #VALUE!. 1 argument used.');
		// Case #9: Table. Table column with non-text value returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1s', 'Test: Negative case: Table. Table column with non-text value returns text conversion. 1 argument used.');
		// Case #10: Formula. Formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('TRIM(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #12: Reference link. Reference to cell with number returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-1', 'Test: Negative case: Reference link. Reference to cell with number returns text conversion. 1 argument used.');
		// Case #14: Array. Array with boolean returns text conversion of first element. 1 argument used.
		oParser = new parserFormula('TRIM({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Array. Array with boolean returns text conversion of first element. 1 argument used.');
		// Case #15: Number. Negative number converted to text. 1 argument used.
		oParser = new parserFormula('TRIM(-123)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(-123) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '-123', 'Test: Negative case: Number. Negative number converted to text. 1 argument used.');
		// Case #16: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.
		oParser = new parserFormula('TRIM(DIVIDE(1,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(DIVIDE(1,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NAME?', 'Test: Negative case: Formula. Formula resulting in #DIV/0! propagates error. 1 argument used.');
		// Case #17: Reference link. Reference to cell with boolean returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(A105)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(A105) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '1', 'Test: Negative case: Reference link. Reference to cell with boolean returns text conversion. 1 argument used.');
		// Case #18: Name. Named range with number returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(TestName2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestName2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '10.5', 'Test: Negative case: Name. Named range with number returns text conversion. 1 argument used.');
		// Case #19: Name3D. 3D named range with number returns text conversion. 1 argument used.
		oParser = new parserFormula('TRIM(TestName3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(TestName3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Name3D. 3D named range with number returns text conversion. 1 argument used.');

		// Bounded cases:
		let str = "C";
		// Case #2: String. String at Excel\'s cell character limit (32,767 chars). 1 argument used.
		oParser = new parserFormula('TRIM(REPT("C",32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM(REPT("C",32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), str.repeat(32767), 'Test: Bounded case: String. String at Excel\'s cell character limit (32,767 chars). 1 argument used.');
		// Case #3: String. String at Excel\'s cell character limit with extra spaces. 1 argument used.
		oParser = new parserFormula('TRIM("   ...32767-char-string...   ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: TRIM("   ...32767-char-string...   ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '...32767-char-string...', 'Test: Bounded case: String. String at Excel\'s cell character limit with extra spaces. 1 argument used.');

		// Need to fix: area handle
		// Case #4: Area. Multi-cell range returns #VALUE!. 1 argument used.

		testArrayFormula2(assert, "TRIM", 1, 1);
	});

	QUnit.test("Test: \"UNICHAR\"", function (assert) {

		oParser = new parserFormula('UNICHAR(66)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "B");

		oParser = new parserFormula('UNICHAR(32)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), " ");

		oParser = new parserFormula('UNICHAR(0)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#VALUE!");

		oParser = new parserFormula('UNICHAR(48)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "0");

		oParser = new parserFormula('UNICHAR(49)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "1");

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
		// Case #1: Number. Basic valid input: integer code point for \'A\'. 1 argument used.
		oParser = new parserFormula('UNICHAR(65)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(65) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Number. Basic valid input: integer code point for \'A\'. 1 argument used.');
		// Case #2: Number. Basic valid input: integer code point for \'a\'. 1 argument used.
		oParser = new parserFormula('UNICHAR(97)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(97) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'a', 'Test: Positive case: Number. Basic valid input: integer code point for \'a\'. 1 argument used.');
		// Case #3: Number. Valid input: Unicode code point for emoji ?. 1 argument used.
		oParser = new parserFormula('UNICHAR(128512)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(128512) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '😀', 'Test: Positive case: Number. Valid input: Unicode code point for emoji ?. 1 argument used.');
		// Case #5: String. String convertible to valid number. 1 argument used.
		oParser = new parserFormula('UNICHAR("65")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR("65") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: String. String convertible to valid number. 1 argument used.');
		// Case #6: Formula. Nested formula resolving to valid integer (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(ROUND(65.7,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(ROUND(65.7,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'B', 'Test: Positive case: Formula. Nested formula resolving to valid integer (65). 1 argument used.');
		// Case #7: Formula. Nested formula resolving to 65. 1 argument used.
		oParser = new parserFormula('UNICHAR(SQRT(4225))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(SQRT(4225)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested formula resolving to 65. 1 argument used.');
		// Case #8: Reference link. Ref to cell with valid number (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Reference link. Ref to cell with valid number (65). 1 argument used.');
		// Case #9: Area. Single-cell range with valid number (97). 1 argument used.
		oParser = new parserFormula('UNICHAR(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area. Single-cell range with valid number (97). 1 argument used.');
		// Case #10: Array. Array with single valid element. 1 argument used.
		oParser = new parserFormula('UNICHAR({65})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR({65}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: Array. Array with single valid element. 1 argument used.');
		// Case #11: Name. Named range with valid number (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name. Named range with valid number (65). 1 argument used.');
		// Case #12: Name3D. 3D named range with valid number (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Positive case: Name3D. 3D named range with valid number (65). 1 argument used.');
		// Case #13: Ref3D. 3D reference to cell with valid number (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Ref3D. 3D reference to cell with valid number (65). 1 argument used.');
		// Case #14: Area3D. 3D single-cell range with valid number (97). 1 argument used.
		oParser = new parserFormula('UNICHAR(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area3D. 3D single-cell range with valid number (97). 1 argument used.');
		// Case #15: Table. Table structured reference with valid number (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '{', 'Test: Positive case: Table. Table structured reference with valid number (65). 1 argument used.');
		// Case #16: Formula. Nested IF returning valid code point (65). 1 argument used.
		oParser = new parserFormula('UNICHAR(IF(TRUE, 65, 97))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(IF(TRUE, 65, 97)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Formula. Nested IF returning valid code point (65). 1 argument used.');
		// Case #17: Date. Date as serial number (45654). 1 argument used.
		oParser = new parserFormula('UNICHAR(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '뉚', 'Test: Positive case: Date. Date as serial number (45654). 1 argument used.');
		// Case #18: Formula. Nested formula resolving to 195. 1 argument used.
		oParser = new parserFormula('UNICHAR(ABS(-65)+130)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(ABS(-65)+130) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'Ã', 'Test: Positive case: Formula. Nested formula resolving to 195. 1 argument used.');
		// Case #19: Array. Multi-element array, processes first valid element (65). 1 argument used.
		oParser = new parserFormula('UNICHAR({65,97})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR({65,97}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 'A', 'Test: Positive case: Array. Multi-element array, processes first valid element (65). 1 argument used.');
		// Case #20: Number. Valid input: Unicode code point for € symbol. 1 argument used.
		oParser = new parserFormula('UNICHAR(8364)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(8364) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '€', 'Test: Positive case: Number. Valid input: Unicode code point for € symbol. 1 argument used.');
		// Case #21: Formula. Nested formula resolving to 128512 (emoji ?). 1 argument used.
		oParser = new parserFormula('UNICHAR(FLOOR(128512.9,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(FLOOR(128512.9,1)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '😀', 'Test: Positive case: Formula. Nested formula resolving to 128512 (emoji ?). 1 argument used.');
		// Case #22: String. String convertible to emoji code point (128512). 1 argument used.
		oParser = new parserFormula('UNICHAR("128512")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR("128512") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '😀', 'Test: Positive case: String. String convertible to emoji code point (128512). 1 argument used.');

		// Negative cases:
		// Case #1: Number. Number < 1 returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Number < 1 returns #VALUE!. 1 argument used.');
		// Case #2: Number. Negative number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Negative number returns #VALUE!. 1 argument used.');
		// Case #3: String. Non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 1 argument used.');
		// Case #4: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #5: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('UNICHAR(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #6: Area. Multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(A102:A103) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #7: Empty. Reference to empty cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Reference to empty cell returns #VALUE!. 1 argument used.');
		// Case #8: String. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Empty string returns #VALUE!. 1 argument used.');
		// Case #9: Number. Code point > 1114111 returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(1114112)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(1114112) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Code point > 1114111 returns #VALUE!. 1 argument used.');
		// Case #10: Ref3D. 3D ref to text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D ref to text returns #VALUE!. 1 argument used.');
		// Case #11: Name. Named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name. Named range with text returns #VALUE!. 1 argument used.');
		// Case #12: Name3D. 3D named range with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Name3D. 3D named range with text returns #VALUE!. 1 argument used.');
		// Case #13: Table. Table column with text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with text returns #VALUE!. 1 argument used.');
		// Case #14: Formula. Formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('UNICHAR(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #15: Array. Array with boolean returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR({FALSE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR({FALSE}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: Array. Array with boolean returns #VALUE!. 1 argument used.');
		// Case #16: Number. Excessively large number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(1E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Number. Excessively large number returns #VALUE!. 1 argument used.');
		// Case #17: Time. Time value (0.5) returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Time. Time value (0.5) returns #VALUE!. 1 argument used.');
		// Case #18: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(Sheet2!A4:A5) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #19: String. String convertible to negative number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR("-65")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR("-65") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String convertible to negative number returns #VALUE!. 1 argument used.');
		// Case #20: Formula. Date resolving to 0 returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(DATE(1899,12,31))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(DATE(1899,12,31)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '򩛊', 'Test: Negative case: Formula. Date resolving to 0 returns #VALUE!. 1 argument used.');
		// Case #21: Formula. Nested formula resolving to negative number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR(-ABS(65))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(-ABS(65)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Nested formula resolving to negative number returns #VALUE!. 1 argument used.');
		// Case #22: Array. Array with invalid first element returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICHAR({0,65})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR({0,65}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), '#VALUE!', 'Test: Negative case: Array. Array with invalid first element returns #VALUE!. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Minimum valid code point (U+0001). 1 argument used.
		oParser = new parserFormula('UNICHAR(1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Bounded case: Number. Minimum valid code point (U+0001). 1 argument used.');
		// Case #2: Number. Maximum valid code point (U+10FFFF). 1 argument used.
		oParser = new parserFormula('UNICHAR(1114111)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICHAR(1114111) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: Number. Maximum valid code point (U+10FFFF). 1 argument used.');

		// Need to fix: different results from MS especially with big numbers
		// Case #3: Number. Valid input: Unicode code point for emoji ?. 1 argument used.
		// Case #21: Formula. Nested formula resolving to 128512 (emoji ?). 1 argument used.
		// Case #22: String. String convertible to emoji code point (128512). 1 argument used.
		// Case #16: Number. Excessively large number returns #VALUE!. 1 argument used.
		// Case #20: Formula. Date resolving to 0 returns #VALUE!. 1 argument used.
		// Case #2: Number. Maximum valid code point (U+10FFFF). 1 argument used.


	});

	QUnit.test("Test: \"UNICODE\"", function (assert) {

		oParser = new parserFormula('UNICODE(" ")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 32);

		oParser = new parserFormula('UNICODE("B")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 66);

		oParser = new parserFormula('UNICODE(0)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 48);

		oParser = new parserFormula('UNICODE(1)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 49);

		oParser = new parserFormula('UNICODE("true")', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 116);

		oParser = new parserFormula('UNICODE(#N/A)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "#N/A");

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
		// Case #0: String. Single character string. 1 argument used.
		oParser = new parserFormula('UNICODE("A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("A") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: String. Single character string. 1 argument used.');
		// Case #1: String. Multi-character string, returns first character’s code. 1 argument used.
		oParser = new parserFormula('UNICODE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 97, 'Test: Positive case: String. Multi-character string, returns first character’s code. 1 argument used.');
		// Case #2: Number. Number converted to string, returns code for "6". 1 argument used.
		oParser = new parserFormula('UNICODE(65)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(65) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 54, 'Test: Positive case: Number. Number converted to string, returns code for "6". 1 argument used.');
		// Case #3: Formula. Nested IF formula returning valid string. 1 argument used.
		oParser = new parserFormula('UNICODE(IF(TRUE,"A","B"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(IF(TRUE,"A","B")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 65, 'Test: Positive case: Formula. Nested IF formula returning valid string. 1 argument used.');
		// Case #4: Formula. Nested CHAR formula returning "B". 1 argument used.
		oParser = new parserFormula('UNICODE(CHAR(66))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(CHAR(66)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 66, 'Test: Positive case: Formula. Nested CHAR formula returning "B". 1 argument used.');
		// Case #5: Reference link. Reference to cell with single character. 1 argument used.
		oParser = new parserFormula('UNICODE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Positive case: Reference link. Reference to cell with single character. 1 argument used.');
		// Case #6: Area. Single-cell range with string. 1 argument used.
		oParser = new parserFormula('UNICODE(A100:A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A100:A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Positive case: Area. Single-cell range with string. 1 argument used.');
		// Case #7: Area. Multi-cell range, returns first cell’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(A101:A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A101:A102) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Area. Multi-cell range, returns first cell’s code. 1 argument used.');
		// Case #8: Array. Array with single string element. 1 argument used.
		oParser = new parserFormula('UNICODE({"A"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE({"A"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 65, 'Test: Positive case: Array. Array with single string element. 1 argument used.');
		// Case #9: Array. Multi-element array, returns first element’s code. 1 argument used.
		oParser = new parserFormula('UNICODE({"A","B"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE({"A","B"}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 65, 'Test: Positive case: Array. Multi-element array, returns first element’s code. 1 argument used.');
		// Case #10: Name. Named range with single character. 1 argument used.
		oParser = new parserFormula('UNICODE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45, 'Test: Positive case: Name. Named range with single character. 1 argument used.');
		// Case #11: Name3D. 3D named range with single character. 1 argument used.
		oParser = new parserFormula('UNICODE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45, 'Test: Positive case: Name3D. 3D named range with single character. 1 argument used.');
		// Case #12: Ref3D. 3D reference to cell with single character. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Ref3D. 3D reference to cell with single character. 1 argument used.');
		// Case #13: Area3D. 3D single-cell range. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Area3D. 3D single-cell range. 1 argument used.');
		// Case #14: Area3D. 3D multi-cell range, returns first cell’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A1:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A1:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Area3D. 3D multi-cell range, returns first cell’s code. 1 argument used.');
		// Case #15: Table. Table structured reference with single character. 1 argument used.
		oParser = new parserFormula('UNICODE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Table. Table structured reference with single character. 1 argument used.');
		// Case #16: Date. Date serial number converted to string. 1 argument used.
		oParser = new parserFormula('UNICODE(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 52, 'Test: Positive case: Date. Date serial number converted to string. 1 argument used.');
		// Case #17: Time. Time value converted to string. 1 argument used.
		oParser = new parserFormula('UNICODE(TIME(12,0,0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TIME(12,0,0)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Positive case: Time. Time value converted to string. 1 argument used.');
		// Case #18: Formula. Formula returning number converted to string. 1 argument used.
		oParser = new parserFormula('UNICODE(SUM(65,35))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(SUM(65,35)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Positive case: Formula. Formula returning number converted to string. 1 argument used.');
		// Case #19: String. Emoji (multi-byte Unicode character). 1 argument used.
		oParser = new parserFormula('UNICODE("?")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("?") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 63, 'Test: Positive case: String. Emoji (multi-byte Unicode character). 1 argument used.');
		// Case #20: String. Non-Latin script (Chinese character). 1 argument used.
		oParser = new parserFormula('UNICODE("??")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("??") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 63, 'Test: Positive case: String. Non-Latin script (Chinese character). 1 argument used.');
		// Case #21: Formula. UNICODE inside SUM formula. 1 argument used.
		oParser = new parserFormula('SUM(UNICODE("A"),10)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: SUM(UNICODE("A"),10) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 75, 'Test: Positive case: Formula. UNICODE inside SUM formula. 1 argument used.');

		// Negative cases:
		// Case #1: Empty. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty string returns #VALUE!. 1 argument used.');
		// Case #2: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('UNICODE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #3: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #4: Boolean. Boolean FALSE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE(FALSE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(FALSE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 70, 'Test: Negative case: Boolean. Boolean FALSE returns #VALUE!. 1 argument used.');
		// Case #5: String. Numeric string, returns code for "0". 1 argument used.
		oParser = new parserFormula('UNICODE("0.5")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("0.5") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: String. Numeric string, returns code for "0". 1 argument used.');
		// Case #6: Reference link. Reference to cell with error returns #N/A. 1 argument used.
		oParser = new parserFormula('UNICODE(A102)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A102) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: Reference link. Reference to cell with error returns #N/A. 1 argument used.');
		// Case #7: Area. Single-cell range with error returns #N/A. 1 argument used.
		oParser = new parserFormula('UNICODE(A103:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A103:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Area. Single-cell range with error returns #N/A. 1 argument used.');
		// Case #8: Array. Array with boolean returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE({TRUE})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE({TRUE}) is parsed.');
		assert.strictEqual(oParser.calculate().getElementRowCol(0,0).getValue(), 84, 'Test: Negative case: Array. Array with boolean returns #VALUE!. 1 argument used.');
		// Case #9: Name. Named range with multi-cell array, returns first element’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(TestNameArea2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TestNameArea2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: Name. Named range with multi-cell array, returns first element’s code. 1 argument used.');
		// Case #10: Name3D. 3D named range with multi-cell array, returns first element’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: Name3D. 3D named range with multi-cell array, returns first element’s code. 1 argument used.');
		// Case #11: Ref3D. 3D reference to cell with text string, returns first character’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 50, 'Test: Negative case: Ref3D. 3D reference to cell with text string, returns first character’s code. 1 argument used.');
		// Case #12: Area3D. 3D single-cell range with error returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A3:A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A3:A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Negative case: Area3D. 3D single-cell range with error returns #VALUE!. 1 argument used.');
		// Case #13: Table. Table column with multi-character string, returns first character’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Negative case: Table. Table column with multi-character string, returns first character’s code. 1 argument used.');
		// Case #14: Formula. Formula resulting in #NUM! error propagates error. 1 argument used.
		oParser = new parserFormula('UNICODE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Formula resulting in #NUM! error propagates error. 1 argument used.');
		// Case #15: Number. Negative number converted to string, returns code for "-". 1 argument used.
		oParser = new parserFormula('UNICODE(-1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(-1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45, 'Test: Negative case: Number. Negative number converted to string, returns code for "-". 1 argument used.');
		// Case #16: Number. Zero converted to string, returns code for "0". 1 argument used.
		oParser = new parserFormula('UNICODE(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: Number. Zero converted to string, returns code for "0". 1 argument used.');
		// Case #17: String. Date string, returns code for "1". 1 argument used.
		oParser = new parserFormula('UNICODE("12/12/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE("12/12/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Negative case: String. Date string, returns code for "1". 1 argument used.');
		// Case #18: Reference link. Reference to empty cell returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE(A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Reference link. Reference to empty cell returns #VALUE!. 1 argument used.');
		// Case #19: Name. Named range with multi-character string, returns first character’s code. 1 argument used.
		oParser = new parserFormula('UNICODE(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 48, 'Test: Negative case: Name. Named range with multi-character string, returns first character’s code. 1 argument used.');
		// Case #20: Ref3D. 3D reference to cell with error returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UNICODE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 84, 'Test: Negative case: Ref3D. 3D reference to cell with error returns #VALUE!. 1 argument used.');

		// Bounded cases:
		// Case #1: String. Minimum Unicode code point (U+0000, null character). 1 argument used.
		oParser = new parserFormula('UNICODE(CHAR(0))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(CHAR(0)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Bounded case: String. Minimum Unicode code point (U+0000, null character). 1 argument used.');
		// Case #2: String. Maximum Unicode code point (U+10FFFF). 1 argument used.
		oParser = new parserFormula('UNICODE(UNICHAR(1114111))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(UNICHAR(1114111)) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Bounded case: String. Maximum Unicode code point (U+10FFFF). 1 argument used.');
		// Case #3: Number. Maximum Excel number converted to string, returns code for "9". 1 argument used.
		oParser = new parserFormula('UNICODE(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Bounded case: Number. Maximum Excel number converted to string, returns code for "9". 1 argument used.');
		// Case #4: Number. Minimum positive Excel number converted to string, returns code for "1". 1 argument used.
		oParser = new parserFormula('UNICODE(1E-307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UNICODE(1E-307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 49, 'Test: Bounded case: Number. Minimum positive Excel number converted to string, returns code for "1". 1 argument used.');

		// Need to fix: area handle, diff error types, results diff from MS
		// Case #7: Area. Multi-cell range, returns first cell’s code. 1 argument used.
		// Case #1: Empty. Empty string returns #VALUE!. 1 argument used.
		// Case #7: Area. Single-cell range with error returns #N/A. 1 argument used.
		// Case #18: Reference link. Reference to empty cell returns #VALUE!. 1 argument used.
		// Case #1: String. Minimum Unicode code point (U+0000, null character). 1 argument used.
		// Case #2: String. Maximum Unicode code point (U+10FFFF). 1 argument used.
		// Case #3: Number. Maximum Excel number converted to string, returns code for "9". 1 argument used.

	});

	QUnit.test("Test: \"UPPER\"", function (assert) {
		ws.getRange2("A2").setValue("total");
		ws.getRange2("A3").setValue("Yield");

		oParser = new parserFormula('UPPER(A2)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "TOTAL");

		oParser = new parserFormula('UPPER(A3)', "AA2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), "YIELD");

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
		getTableType(599, 0, 600, 0);
		ws.getRange2("A601").setValue("Col1Text"); // Text (Column1)
		// 3D links. Use A1:Z10
		let ws2 = getSecondSheet();
		ws2.getRange2("A1").setValue("0.5");
		ws2.getRange2("A2").setValue("#N/A");
		ws2.getRange2("B1").setValue("-1");
		ws2.getRange2("C1").setValue("1");
		// DefNames.
		initDefNames();
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
		// Case #1: String. Basic valid input: uppercase string. 1 argument used.
		oParser = new parserFormula('UPPER("HELLO")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER("HELLO") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HELLO', 'Test: Positive case: String. Basic valid input: uppercase string. 1 argument used.');
		// Case #2: String. Alphanumeric string. 1 argument used.
		oParser = new parserFormula('UPPER("123ABC")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER("123ABC") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '123ABC', 'Test: Positive case: String. Alphanumeric string. 1 argument used.');
		// Case #3: String. Cyrillic string. 1 argument used.
		oParser = new parserFormula('UPPER("??????")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER("??????") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '??????', 'Test: Positive case: String. Cyrillic string. 1 argument used.');
		// Case #5: Number. Number implicitly converted to string. 1 argument used.
		oParser = new parserFormula('UPPER(100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "100", 'Test: Positive case: Number. Number implicitly converted to string. 1 argument used.');
		// Case #6: Formula. Nested formula. 1 argument used.
		oParser = new parserFormula('UPPER(LOWER("test"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(LOWER("test")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TEST', 'Test: Positive case: Formula. Nested formula. 1 argument used.');
		// Case #7: Reference link. Ref to cell with text. 1 argument used.
		oParser = new parserFormula('UPPER(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(A100) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Reference link. Ref to cell with text. 1 argument used.');
		// Case #8: Area. Single-cell range. 1 argument used.
		oParser = new parserFormula('UPPER(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(A101:A101) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Positive case: Area. Single-cell range. 1 argument used.');
		// Case #9: Array. Array with single element. 1 argument used.
		oParser = new parserFormula('UPPER({"HELLO"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER({"HELLO"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'HELLO', 'Test: Positive case: Array. Array with single element. 1 argument used.');
		// Case #10: Name. Named range. 1 argument used.
		oParser = new parserFormula('UPPER(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "-0.5", 'Test: Positive case: Name. Named range. 1 argument used.');
		// Case #11: Name3D. 3D named range. 1 argument used.
		oParser = new parserFormula('UPPER(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "-0.5", 'Test: Positive case: Name3D. 3D named range. 1 argument used.');
		// Case #12: Ref3D. 3D reference to cell. 1 argument used.
		oParser = new parserFormula('UPPER(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Ref3D. 3D reference to cell. 1 argument used.');
		// Case #13: Area3D. 3D single-cell range. 1 argument used.
		oParser = new parserFormula('UPPER(Sheet2!A1:A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(Sheet2!A1:A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '0.5', 'Test: Positive case: Area3D. 3D single-cell range. 1 argument used.');
		// Case #14: Table. Table structured reference. 1 argument used.
		oParser = new parserFormula('UPPER(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "COL1TEXT", 'Test: Positive case: Table. Table structured reference. 1 argument used.');
		// Case #15: Date. Date converted to string. 1 argument used.
		oParser = new parserFormula('UPPER(TEXT(DATE(2025,1,1),"YYYY"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TEXT(DATE(2025,1,1),"YYYY")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "2025", 'Test: Positive case: Date. Date converted to string. 1 argument used.');
		// Case #16: Time. Time converted to string. 1 argument used.
		oParser = new parserFormula('UPPER(TEXT(TIME(12,0,0),"HH:MM"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TEXT(TIME(12,0,0),"HH:MM")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '12:00', 'Test: Positive case: Time. Time converted to string. 1 argument used.');
		// Case #17: Formula. UPPER inside CONCAT. 1 argument used.
		oParser = new parserFormula('CONCAT(UPPER("A"), "B")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: CONCAT(UPPER("A"), "B") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'AB', 'Test: Positive case: Formula. UPPER inside CONCAT. 1 argument used.');
		// Case #18: String. Special characters. 1 argument used.
		oParser = new parserFormula('UPPER("!@#")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER("!@#") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '!@#', 'Test: Positive case: String. Special characters. 1 argument used.');
		// Case #19: Array. Multi-element array. 1 argument used.
		oParser = new parserFormula('UPPER({"A", "B"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER({"A", "B"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Positive case: Array. Multi-element array. 1 argument used.');
		// Case #20: Formula. Nested IF returning valid value. 1 argument used.
		oParser = new parserFormula('UPPER(IF(TRUE, "yes", "no"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(IF(TRUE, "yes", "no")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'YES', 'Test: Positive case: Formula. Nested IF returning valid value. 1 argument used.');

		// Negative cases:
		// Case #1: Error. Propagates #N/A error. 1 argument used.
		oParser = new parserFormula('UPPER(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Propagates #N/A error. 1 argument used.');
		// Case #2: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UPPER(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'TRUE', 'Test: Negative case: Boolean. Boolean TRUE returns #VALUE!. 1 argument used.');
		// Case #3: Empty. Reference link is empty. 1 argument used.
		oParser = new parserFormula('UPPER(A112)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(A112) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Empty. Reference link is empty. 1 argument used.');
		// Case #4: Area. Multi-cell range returns arr. 1 argument used.
		oParser = new parserFormula('UPPER(A103:A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(A103:A104) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), '', 'Test: Negative case: Area. Multi-cell range returns arr. 1 argument used.');
		// Case #5: Ref3D. 3D ref to error returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UPPER(Sheet2!A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(Sheet2!A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Ref3D. 3D ref to error returns #VALUE!. 1 argument used.');
		// Case #6: Name. Named range with val. 1 argument used.
		oParser = new parserFormula('UPPER(TestNameArea)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(TestNameArea) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), "1", 'Test: Negative case: Name. Named range with val. 1 argument used.');
		// Case #8: Number. Extremely large number returns #VALUE!. 1 argument used.
		oParser = new parserFormula('UPPER(1E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(1E+307) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "1E+307", 'Test: Negative case: Number. Extremely large number returns #VALUE!. 1 argument used.');

		let longStr = "A";
		longStr = longStr.repeat(32767);
		// Bounded cases:
		// Case #1: String. Max string length in Excel. 1 argument used.
		oParser = new parserFormula('UPPER(REPT("A", 32767))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(REPT("A", 32767)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), longStr, 'Test: Bounded case: String. Max string length in Excel. 1 argument used.');
		// Case #2: String. Min non-empty string. 1 argument used.
		oParser = new parserFormula('UPPER("A")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER("A") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 'A', 'Test: Bounded case: String. Min non-empty string. 1 argument used.');
		// Case #3: Number. Zero converted to string. 1 argument used.
		oParser = new parserFormula('UPPER(0)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: UPPER(0) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), "0", 'Test: Bounded case: Number. Zero converted to string. 1 argument used.');


		// Need to fix:
		// Should return array when encounter the area
		// Case #4: Area. Multi-cell range returns arr.
		// Case #6: Name. Named range with val

		testArrayFormula2(assert, "UPPER", 1, 1);
	});

	QUnit.test("Test: \"VALUE\"", function (assert) {

		oParser = new parserFormula("VALUE(\"123.456\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 123.456);

		oParser = new parserFormula("VALUE(\"$1,000\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 1000);

		oParser = new parserFormula("VALUE(\"23-Mar-2002\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), 37338);

		oParser = new parserFormula("VALUE(\"03-26-2006\")", "A2", ws);
		assert.ok(oParser.parse());

		if (AscCommon.bDate1904) {
			assert.strictEqual(oParser.calculate().getValue(), 37340);
		} else {
			assert.strictEqual(oParser.calculate().getValue(), 38802);
		}

		oParser = new parserFormula("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")", "A2", ws);
		assert.ok(oParser.parse());
		assert.strictEqual(oParser.calculate().getValue(), AscCommon.g_oFormatParser.parse("16:48:00").value - AscCommon.g_oFormatParser.parse("12:17:12").value);

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
		// Case #1: String. Text string representing a number converted to number. 1 argument used.
		oParser = new parserFormula('VALUE("123")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("123") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123, 'Test: Positive case: String. Text string representing a number converted to number. 1 argument used.');
		// Case #2: Number. Number input remains unchanged. 1 argument used.
		oParser = new parserFormula('VALUE(123.45)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(123.45) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123.45, 'Test: Positive case: Number. Number input remains unchanged. 1 argument used.');
		// Case #3: Formula. Nested SQRT formula evaluating to a number. 1 argument used.
		oParser = new parserFormula('VALUE(SQRT(16))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(SQRT(16)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 4, 'Test: Positive case: Formula. Nested SQRT formula evaluating to a number. 1 argument used.');
		// Case #4: Formula. Nested IF formula returning a numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(IF(TRUE, "123", "456"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(IF(TRUE, "123", "456")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123, 'Test: Positive case: Formula. Nested IF formula returning a numeric string. 1 argument used.');
		// Case #5: Date. Date string converted to serial number. 1 argument used.
		oParser = new parserFormula('VALUE("01/01/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("01/01/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Date. Date string converted to serial number. 1 argument used.');
		// Case #6: Time. Time string converted to decimal number. 1 argument used.
		oParser = new parserFormula('VALUE("12:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("12:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Positive case: Time. Time string converted to decimal number. 1 argument used.');
		// Case #7: Reference link. Reference to cell with valid numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(A100)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(A100) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Reference link. Reference to cell with valid numeric string. 1 argument used.');
		// Case #8: Area. Single-cell range with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(A101:A101)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(A101:A101) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Positive case: Area. Single-cell range with numeric string. 1 argument used.');
		// Case #9: Array. Array with single numeric string element. 1 argument used.
		oParser = new parserFormula('VALUE({"123"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE({"123"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123, 'Test: Positive case: Array. Array with single numeric string element. 1 argument used.');
		// Case #10: Name. Named range with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(TestName)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TestName) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name. Named range with numeric string. 1 argument used.');
		// Case #11: Name3D. 3D named range with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(TestName3D)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TestName3D) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -0.5, 'Test: Positive case: Name3D. 3D named range with numeric string. 1 argument used.');
		// Case #12: Ref3D. 3D reference to cell with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(Sheet2!A1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Sheet2!A1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Ref3D. 3D reference to cell with numeric string. 1 argument used.');
		// Case #13: Area3D. 3D single-cell range with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(Sheet2!A2:A2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Sheet2!A2:A2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2, 'Test: Positive case: Area3D. 3D single-cell range with numeric string. 1 argument used.');
		// Case #14: Table. Table column with numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(Table1[Column1])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Table1[Column1]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Positive case: Table. Table column with numeric string. 1 argument used.');
		// Case #15: Formula. Nested ROUND formula evaluating to a number. 1 argument used.
		oParser = new parserFormula('VALUE(ROUND(123.456, 2))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(ROUND(123.456, 2)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123.46, 'Test: Positive case: Formula. Nested ROUND formula evaluating to a number. 1 argument used.');
		// Case #16: String. Scientific notation string converted to number. 1 argument used.
		oParser = new parserFormula('VALUE("1.23E+10")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("1.23E+10") is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 12300000000, 'Test: Positive case: String. Scientific notation string converted to number. 1 argument used.');
		// Case #17: String. String with thousands separator converted to number. 1 argument used.
		oParser = new parserFormula('VALUE("1,234.56")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("1,234.56") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1234.56, 'Test: Positive case: String. String with thousands separator converted to number. 1 argument used.');
		// Case #18: Formula. Date formula evaluated to serial number. 1 argument used.
		oParser = new parserFormula('VALUE(DATE(2025,1,1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(DATE(2025,1,1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 45658, 'Test: Positive case: Formula. Date formula evaluated to serial number. 1 argument used.');
		// Case #19: String. String with extra spaces converted to number. 1 argument used.
		oParser = new parserFormula('VALUE(" 123 ")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(" 123 ") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123, 'Test: Positive case: String. String with extra spaces converted to number. 1 argument used.');
		// Case #20: Formula. Nested ABS formula evaluating to a number. 1 argument used.
		oParser = new parserFormula('VALUE(ABS(-123.45))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(ABS(-123.45)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123.45, 'Test: Positive case: Formula. Nested ABS formula evaluating to a number. 1 argument used.');
		// Case #21: String. Currency formatted string converted to number. 1 argument used.
		oParser = new parserFormula('VALUE("$123.45")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("$123.45") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123.45, 'Test: Positive case: String. Currency formatted string converted to number. 1 argument used.');
		// Case #22: Formula. Nested TEXT formula returning numeric string. 1 argument used.
		oParser = new parserFormula('VALUE(TEXT(123.45, "0.00"))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TEXT(123.45, "0.00")) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123.45, 'Test: Positive case: Formula. Nested TEXT formula returning numeric string. 1 argument used.');

		// Negative cases:
		// Case #1: String. Non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Non-numeric string returns #VALUE!. 1 argument used.');
		// Case #2: Boolean. Boolean value returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(TRUE)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TRUE) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Boolean. Boolean value returns #VALUE!. 1 argument used.');
		// Case #3: Empty. Empty string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Empty. Empty string returns #VALUE!. 1 argument used.');
		// Case #4: Error. Error value propagates #N/A. 1 argument used.
		oParser = new parserFormula('VALUE(NA())', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(NA()) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#N/A', 'Test: Negative case: Error. Error value propagates #N/A. 1 argument used.');
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(A102:A103)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(A102:A103) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area. Multi-cell range returns #VALUE!. 1 argument used.');
		// Case #6: Array. Multi-element array returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE({"123", "abc"})', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE({"123", "abc"}) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 123, 'Test: Negative case: Array. Multi-element array returns #VALUE!. 1 argument used.');
		// Case #7: Reference link. Reference to cell with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(A104)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(A104) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), -1, 'Test: Negative case: Reference link. Reference to cell with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #8: Name. Named range with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(TestName1)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TestName1) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.5, 'Test: Negative case: Name. Named range with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #9: Name3D. 3D named range with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(TestNameArea3D2)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(TestNameArea3D2) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 0.8, 'Test: Negative case: Name3D. 3D named range with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #10: Ref3D. 3D reference to cell with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(Sheet2!A3)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Sheet2!A3) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Ref3D. 3D reference to cell with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #11: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(Sheet2!A4:A5)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Sheet2!A4:A5) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 0, 'Test: Negative case: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.');
		// Case #12: Table. Table column with non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(Table1[Column2])', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(Table1[Column2]) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Table. Table column with non-numeric string returns #VALUE!. 1 argument used.');
		// Case #13: String. String representing division by zero returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("1/0")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("1/0") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 36526, 'Test: Negative case: String. String representing division by zero returns #VALUE!. 1 argument used.');
		// Case #14: Formula. Nested formula resulting in #NUM! propagates error. 1 argument used.
		oParser = new parserFormula('VALUE(SQRT(-1))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(SQRT(-1)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#NUM!', 'Test: Negative case: Formula. Nested formula resulting in #NUM! propagates error. 1 argument used.');
		// Case #15: String. Invalid number format string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("1..2")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("1..2") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. Invalid number format string returns #VALUE!. 1 argument used.');
		// Case #16: String. String with mixed numeric and text returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("123abc")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("123abc") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: String. String with mixed numeric and text returns #VALUE!. 1 argument used.');
		// Case #17: Formula. Nested formula resulting in #DIV/0! propagates error. 1 argument used.
		oParser = new parserFormula('VALUE(MMULT(2,3))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(MMULT(2,3)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 6, 'Test: Negative case: Formula. Nested formula resulting in #DIV/0! propagates error. 1 argument used.');
		// Case #18: Date. Invalid date string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("13/13/2025")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("13/13/2025") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Date. Invalid date string returns #VALUE!. 1 argument used.');
		// Case #19: Time. Invalid time string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE("25:00:00")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("25:00:00") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1.0416666666666667, 'Test: Negative case: Time. Invalid time string returns #VALUE!. 1 argument used.');
		// Case #20: Formula. Nested CHAR formula returning non-numeric string returns #VALUE!. 1 argument used.
		oParser = new parserFormula('VALUE(CHAR(255))', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(CHAR(255)) is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), '#VALUE!', 'Test: Negative case: Formula. Nested CHAR formula returning non-numeric string returns #VALUE!. 1 argument used.');

		// Bounded cases:
		// Case #1: Number. Maximum valid Excel number. 1 argument used.
		oParser = new parserFormula('VALUE(9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), 1e+308, 'Test: Bounded case: Number. Maximum valid Excel number. 1 argument used.');
		// Case #2: Number. Minimum valid Excel number. 1 argument used.
		oParser = new parserFormula('VALUE(-9.99999999999999E+307)', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE(-9.99999999999999E+307) is parsed.');
		//? assert.strictEqual(oParser.calculate().getValue(), -1e+308, 'Test: Bounded case: Number. Minimum valid Excel number. 1 argument used.');
		// Case #3: Date. Minimum valid Excel date serial number (1). 1 argument used.
		oParser = new parserFormula('VALUE("01/01/1900")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("01/01/1900") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1, 'Test: Bounded case: Date. Minimum valid Excel date serial number (1). 1 argument used.');
		// Case #4: Date. Maximum valid Excel date serial number (2958465). 1 argument used.
		oParser = new parserFormula('VALUE("12/31/9999")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("12/31/9999") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 2958465, 'Test: Bounded case: Date. Maximum valid Excel date serial number (2958465). 1 argument used.');
		// Case #5: String. Smallest positive number string convertible to number. 1 argument used.
		oParser = new parserFormula('VALUE("0.000000000000001")', 'A2', ws);
		assert.ok(oParser.parse(), 'Test: VALUE("0.000000000000001") is parsed.');
		assert.strictEqual(oParser.calculate().getValue(), 1e-15, 'Test: Bounded case: String. Smallest positive number string convertible to number. 1 argument used.');

		// Need to fix: link handle, results diff from ms
		// Case #7: Reference link. Reference to cell with valid numeric string. 1 argument used.
		// Case #8: Area. Single-cell range with numeric string. 1 argument used.
		// Case #16: String. Scientific notation string converted to number. 1 argument used.
		// Case #5: Area. Multi-cell range returns #VALUE!. 1 argument used.
		// Case #11: Area3D. 3D multi-cell range returns #VALUE!. 1 argument used.
		// Case #1: Number. Maximum valid Excel number. 1 argument used.
		// Case #2: Number. Minimum valid Excel number. 1 argument used.


		testArrayFormula2(assert, "value", 1, 1);
	});

	wb.dependencyFormulas.unlockRecal();
});
